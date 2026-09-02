# =============================================================================
#  Cardiochirurgie AZ Delta
#  Descriptive analysis: baseline characteristics + perioperative outcomes
# -----------------------------------------------------------------------------
#  Data   : data-raw/iz_cardio_versie_28082026.xlsx
#           one row = one ICU admission after cardiac surgery
#           (1892 records / 1869 unique patients, surgery 2020-12-29 .. 2026-08-19)
#
#  Output : output/  -> Table1_baseline.csv, Table2_outcomes.csv,
#                       verrichting_classification_AUDIT.csv, data_quality.txt,
#                       Cardiochirurgie_results.xlsx, analysis_dataset.rds
#
#  Requires: R >= 4.1 and the package 'readxl'.  'writexl' is optional
#            (Excel export; the script falls back to CSV without it).
#            No tidyverse dependency - everything else is base R.
#
#  HOW TO USE
#    1. Open this file in RStudio, set the working directory to the project
#       root (Session > Set Working Directory > To Source File Location).
#    2. Check the CONFIGURATION block below - especially ELECTIVE_DEFINITION.
#    3. Source the whole file (Ctrl+Shift+S).  Everything is printed to the
#       console and written to output/.
#    4. Review output/verrichting_classification_AUDIT.csv: it lists every one
#       of the 213 distinct 'verrichting' strings with the group it was
#       assigned to.  You know the practice - correct anything that is wrong
#       by editing the keyword lists in section 3.4.
# =============================================================================


# =============================================================================
# 0.  CONFIGURATION -- read this block before running
# =============================================================================

DATA_FILE <- "data-raw/iz_cardio_versie_28082026.xlsx"
SHEET     <- 1
OUT_DIR   <- "output"

## --- Analysis population -----------------------------------------------------
## The file holds 1892 records for 1869 patients: 23 patients have >1 record
## (re-operations / a second ICU episode, e.g. "Revisie", "Thoraxdrain",
## "Pericardfenestratie").  Choose the unit of analysis:
##   "all_records"       - every ICU admission (n = 1892)   [as delivered]
##   "first_per_patient" - index procedure only (n = 1869)  [cleaner Table 1]
ANALYSIS_UNIT <- "all_records"

## --- Elective vs non-elective ------------------------------------------------
## 'urgentie_graad' is blank in 1181/1892 (62.4%).  Observed outcome gradient:
##
##     grade      n     ICU death   any death   mech. support   median ICU LOS
##     (blank)  1181       0.8%        2.3%         0.2%            2.97 d
##     E          33       0.0%        6.1%         0.0%            3.28 d
##     S3        510       2.9%        6.1%         1.0%            3.07 d
##     S2         84       2.4%        4.8%         2.4%            3.97 d
##     S1         81      12.3%       18.5%         4.9%            4.02 d
##
## Blank and "E" behave like the lowest-risk cases and S1 like the sickest,
## i.e. urgency looks like it is only coded when the case is NOT elective.
##   "blank_is_elective" - blank + E = elective; S1/S2/S3 = non-elective  [default]
##   "blank_is_missing"  - blank = missing; E = elective; S1/S2/S3 = non-elective
## >>> Please confirm the local coding of urgentie_graad and set this. <<<
ELECTIVE_DEFINITION <- "blank_is_elective"

## --- Surgical group ----------------------------------------------------------
## Primary grouping (as requested): CABG only / CABG + valve / valve without
## CABG / Other.  Two judgement calls, both switchable:
##   TAVI_AS_VALVE  - TAVI is a transcatheter, not a surgical, valve procedure.
##                    FALSE keeps the 31 TAVI records out of "Valve only".
##   DAVID_AS_VALVE - David / Tirone-David is a valve-SPARING root replacement.
##                    FALSE treats it as aortic (root) surgery, not valve surgery.
TAVI_AS_VALVE  <- FALSE
DAVID_AS_VALVE <- FALSE

## --- Diabetes ----------------------------------------------------------------
## CM_diabetes is missing in 467/1892 (24.7%).  FALSE keeps those as missing;
## TRUE assumes "not recorded" = "no diabetes" (sensitivity analysis).
DM_MISSING_AS_NONE <- FALSE

## --- AKI ---------------------------------------------------------------------
## Stage from max_creatinine vs creatinine_dag_voor_operatie (as requested):
##   stage 1 : rise >= AKI_ABS_RISE  OR  ratio 1.5-1.9      (and ratio < 2)
##   stage 2 : ratio 2.0 - 2.9
##   stage 3 : ratio >= 3.0
## Creatinine unit is auto-detected; AKI_ABS_RISE is set to 0.3 mg/dL or
## 26.5 umol/L accordingly.  Set AKI_ABS_RISE manually to override.
AKI_ABS_RISE <- NA   # NA = auto

## KDIGO also requires the rise to happen inside a time window.  Only the DATE
## of max_creatinine is available, so this is a day-resolution sensitivity
## analysis: absolute rise within 2 days, ratio criteria within 7 days.
AKI_WINDOW_ABS_DAYS   <- 2
AKI_WINDOW_RATIO_DAYS <- 7

## --- Ventilation -------------------------------------------------------------
VENT_THRESHOLD_MIN <- 6 * 60      # "prolonged ventilation" cut-off (360 min)

## --- Descriptive statistics --------------------------------------------------
## Numeric variables are summarised as mean (SD) when they look normal and as
## median [IQR] otherwise.  Rule: Shapiro-Wilk p > SHAPIRO_P and |skew| < MAX_SKEW.
## Both summaries are always written to Table_S1_all_summaries.csv.
SHAPIRO_P <- 0.05
MAX_SKEW  <- 1.0
## Force a specific summary for named variables, e.g.
##   FORCE_SUMMARY <- c(icu_los_days = "median_iqr", age_years = "mean_sd")
FORCE_SUMMARY <- c()

## Optional: stratify Table 1 by a derived variable (NULL = overall only).
## e.g. STRATIFY_BY <- "surg_group"
STRATIFY_BY <- NULL

set.seed(20260828)
options(width = 200)          # keep the printed tables on one line


# =============================================================================
# 1.  PACKAGES AND GENERIC HELPERS
# =============================================================================

if (!requireNamespace("readxl", quietly = TRUE))
  stop("Package 'readxl' is required.  install.packages('readxl')")
HAVE_WRITEXL <- requireNamespace("writexl", quietly = TRUE)

dir.create(OUT_DIR, showWarnings = FALSE, recursive = TRUE)

`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || all(is.na(a))) b else a

## banner / section headers in the console log
hdr <- function(txt, ch = "=") {
  cat("\n", strrep(ch, 78), "\n", txt, "\n", strrep(ch, 78), "\n", sep = "")
}
sub_hdr <- function(txt) cat("\n--- ", txt, " ", strrep("-", max(0, 68 - nchar(txt))), "\n", sep = "")

## Resolve a logical variable name to an actual column of the raw file.
## Matching is by regular expression and case-insensitive, so the script keeps
## working if the next export renames/re-accents a column slightly.
resolve_col <- function(dat, pattern, required = TRUE, label = pattern) {
  hit <- grep(pattern, names(dat), ignore.case = TRUE, perl = TRUE)
  if (length(hit) == 0) {
    if (required) stop(sprintf("Required column not found for '%s' (pattern: %s)", label, pattern))
    return(NA_character_)
  }
  if (length(hit) > 1) {
    ## prefer an exact (case-insensitive) match when the pattern is ambiguous
    exact <- which(tolower(names(dat)) == tolower(gsub("[\\^\\$]", "", pattern)))
    hit <- if (length(exact) == 1) exact else hit[1]
  }
  names(dat)[hit]
}

## numeric helpers ------------------------------------------------------------
skewness <- function(x) {
  x <- x[is.finite(x)]; n <- length(x)
  if (n < 3) return(NA_real_)
  m <- mean(x); s <- sqrt(sum((x - m)^2) / n)
  if (s == 0) return(0)
  sum((x - m)^3) / (n * s^3)
}

fmt <- function(x, d = 1) formatC(round(x, d), format = "f", digits = d, big.mark = "")

pval_fmt <- function(p) {
  if (is.na(p)) return("")
  if (p < 0.001) "<0.001" else formatC(round(p, 3), format = "f", digits = 3)
}

## Decide mean(SD) vs median[IQR] for one numeric vector.
choose_summary <- function(x, varname = "") {
  if (!is.null(FORCE_SUMMARY[varname]) && !is.na(FORCE_SUMMARY[varname]))
    return(unname(FORCE_SUMMARY[varname]))
  x <- x[is.finite(x)]
  if (length(x) < 3 || length(unique(x)) < 3) return("median_iqr")
  sw <- try(shapiro.test(if (length(x) > 5000) sample(x, 5000) else x)$p.value, silent = TRUE)
  if (inherits(sw, "try-error")) sw <- 0
  sk <- skewness(x)
  if (!is.na(sw) && sw > SHAPIRO_P && !is.na(sk) && abs(sk) < MAX_SKEW) "mean_sd" else "median_iqr"
}

## Summarise one numeric vector -> one-row data.frame.
describe_num <- function(x, label, varname = label, digits = 1, force = NULL) {
  x  <- suppressWarnings(as.numeric(x))
  ok <- is.finite(x); v <- x[ok]
  n_miss <- sum(!ok)
  meth <- force %||% choose_summary(v, varname)
  if (length(v) == 0) {
    s <- "not available"
  } else if (meth == "mean_sd") {
    s <- sprintf("%s (%s)", fmt(mean(v), digits), fmt(sd(v), digits))
  } else {
    q <- quantile(v, c(.25, .5, .75), na.rm = TRUE, type = 7)
    s <- sprintf("%s [%s-%s]", fmt(q[2], digits), fmt(q[1], digits), fmt(q[3], digits))
  }
  data.frame(
    variable   = label, level = "", n = length(v), summary = s,
    method     = if (meth == "mean_sd") "mean (SD)" else "median [IQR]",
    missing_n  = n_miss,
    missing_pct= sprintf("%s", fmt(100 * n_miss / length(x), 1)),
    stringsAsFactors = FALSE
  )
}

## Summarise one categorical vector -> header row + one row per level.
## Percentages are of the non-missing total; missing is reported separately.
describe_cat <- function(x, label, drop_levels = NULL, na_level = NULL) {
  if (!is.factor(x)) x <- factor(x)
  if (!is.null(drop_levels)) x <- droplevels(x[!x %in% drop_levels])
  ## na_level: show missing as an explicit level with this name
  if (!is.null(na_level)) {
    x <- addNA(x, ifany = TRUE)
    levels(x)[is.na(levels(x))] <- na_level
  }
  n_miss <- sum(is.na(x)); n_ok <- sum(!is.na(x))
  tb <- table(x, useNA = "no")
  head_row <- data.frame(
    variable = label, level = "", n = n_ok, summary = "", method = "n (%)",
    missing_n = n_miss, missing_pct = fmt(100 * n_miss / length(x), 1),
    stringsAsFactors = FALSE)
  lvl_rows <- data.frame(
    variable    = "",
    level       = names(tb),
    n           = as.integer(tb),
    summary     = sprintf("%d (%s)", as.integer(tb), fmt(100 * as.integer(tb) / max(1, n_ok), 1)),
    method      = "n (%)",
    missing_n   = NA_integer_,
    missing_pct = "",
    stringsAsFactors = FALSE)
  rbind(head_row, lvl_rows)
}

## Binary yes/no variable -> single row "n (%)" for the "yes" level.
describe_bin <- function(x, label, yes = TRUE) {
  n_miss <- sum(is.na(x)); n_ok <- sum(!is.na(x))
  k <- sum(x == yes, na.rm = TRUE)
  data.frame(variable = label, level = "", n = n_ok,
             summary = sprintf("%d (%s)", k, fmt(100 * k / max(1, n_ok), 1)),
             method = "n (%)", missing_n = n_miss,
             missing_pct = fmt(100 * n_miss / length(x), 1),
             stringsAsFactors = FALSE)
}

## Pretty console printing of a table built by describe_* ----------------------
print_table <- function(tb, title = NULL) {
  if (!is.null(title)) sub_hdr(title)
  d <- tb
  d$Characteristic <- ifelse(d$level == "", d$variable, paste0("   ", d$level))
  keep <- data.frame(Characteristic = d$Characteristic, N = d$n, Summary = d$summary,
                     Method = d$method, `Missing n` = d$missing_n,
                     check.names = FALSE, stringsAsFactors = FALSE)
  keep$`Missing n`[is.na(keep$`Missing n`)] <- ""
  print(keep, row.names = FALSE, right = FALSE)
  invisible(tb)
}


# =============================================================================
# 2.  IMPORT AND COLUMN MAPPING
# =============================================================================

hdr("1. IMPORT")

if (!file.exists(DATA_FILE))
  stop("Data file not found: ", normalizePath(DATA_FILE, mustWork = FALSE),
       "\n     Set DATA_FILE in the configuration block.")

## A large guess_max is essential: several columns (e.g. actieve_endocarditis)
## are blank for the first ~1400 rows, so with readxl's default of 1000 guessed
## rows they are typed as logical and every later "Yes"/"No" is read as NA
## without any error.
raw <- as.data.frame(readxl::read_excel(DATA_FILE, sheet = SHEET,
                                        guess_max = 1e6,
                                        .name_repair = "minimal"),
                     stringsAsFactors = FALSE)
cat(sprintf("File   : %s\nRecords: %d    Columns: %d\n", DATA_FILE, nrow(raw), ncol(raw)))

## Logical name -> regex for the actual column.  Editing this list is the only
## thing needed if a future export renames a column.
COL_PATTERNS <- c(
  pat_id        = "^Pati.ntennummer$",
  adm_id        = "^Opnamenummer$",
  sex           = "^Geslacht$",
  dob           = "^Geboortedatum$",
  icu_in        = "^Opname.?Datetime$",
  icu_out       = "^Ontslag.?Datetime$",
  surg_start    = "^surgery_start_datetime$",
  surg_end      = "^surgery_end_datetime$",
  urgency       = "^urgentie_graad$",
  procedure     = "^verrichting$",
  approach      = "^benadering$",
  readmission   = "^Heropname$",
  los_raw       = "^Ligduur$",
  los_caldays   = "^Ligduur in ligdagen$",
  weight        = "^gewicht$",
  height        = "^lengte$",
  bmi           = "^bmi$",
  smoking       = "^rook_status$",
  diabetes      = "^CM_diabetes$",
  hypertension  = "^CM_hypertensie$",
  lung_disease  = "^VG_longlijden$",
  cva_tia       = "^VG_CVA_TIA$",
  endocarditis  = "^actieve_endocarditis$",
  extub_status  = "^on_table_extubation$",
  vent_min      = "^duur_invbead_minuten$",
  vent_event    = "^duur_invbead_event$",
  crea_pre      = "^creatinine_dag_voor_operatie$",
  crea_max      = "^max_creatinine$",
  crea_max_date = "^datum_max_creatinine$",
  mortality     = "^mortaliteit$",
  time_to_chair = "^tijd_tot_zetel_minuten$",
  vas24         = "^mean_VAS_24h$",
  pcia          = "^PCIA$",
  cpb_min       = "^bypasstijd_minuten$",
  xclamp_min    = "^aortaklemtijd_minuten$",
  ecmo_days     = "^ECMO \\(dagen\\)$",
  impella_days  = "^Impella \\(dagen\\)$",
  iabp_days     = "^IABP \\(dagen\\)$"
)
## Optional columns: absent = feature silently skipped, script still runs.
OPTIONAL <- c("los_caldays", "vent_event", "cpb_min", "xclamp_min", "approach",
              "endocarditis", "ecmo_days", "impella_days", "iabp_days")

COLMAP <- vapply(names(COL_PATTERNS), function(k)
  resolve_col(raw, COL_PATTERNS[[k]], required = !(k %in% OPTIONAL), label = k),
  character(1))

sub_hdr("Column mapping (analysis name -> column in the file)")
print(data.frame(analysis_name = names(COLMAP),
                 file_column   = ifelse(is.na(COLMAP), "<< NOT FOUND >>", COLMAP),
                 row.names = NULL), right = FALSE)

## g("x") returns the raw column for logical name x, or NULL if absent.
g <- function(k) if (is.na(COLMAP[[k]])) NULL else raw[[COLMAP[[k]]]]


# =============================================================================
# 3.  DERIVED VARIABLES
# =============================================================================

hdr("2. DERIVED VARIABLES")

d <- data.frame(row_id = seq_len(nrow(raw)))

## small coding helpers -------------------------------------------------------
## "ja"/"nee" -> TRUE/FALSE, anything else -> NA
ja_nee <- function(x) {
  s <- tolower(trimws(as.character(x)))
  out <- rep(NA, length(s))
  out[s %in% c("ja", "yes", "j", "y", "1", "true")]  <- TRUE
  out[s %in% c("nee", "neen", "no", "n", "0", "false")] <- FALSE
  out
}
## "presence-only" columns: the value is only filled in when the event happened,
## so a blank means "did not happen" (used for PCIA, mortality, readmission...).
present <- function(x) !is.na(x) & trimws(as.character(x)) != ""

## --- 3.0  identifiers, dates, unit of analysis -------------------------------
d$pat_id     <- g("pat_id")
d$adm_id     <- g("adm_id")
d$icu_in     <- as.POSIXct(g("icu_in"))
d$icu_out    <- as.POSIXct(g("icu_out"))
d$surg_start <- as.POSIXct(g("surg_start"))
d$surg_end   <- as.POSIXct(g("surg_end"))
d$surg_year  <- as.integer(format(d$surg_start, "%Y"))
d$surg_dur_min <- as.numeric(difftime(d$surg_end, d$surg_start, units = "mins"))

## first record per patient (index procedure)
ord <- order(d$pat_id, d$surg_start)
d$is_first_episode <- FALSE
d$is_first_episode[ord[!duplicated(d$pat_id[ord])]] <- TRUE

## --- 3.1  Demographics -------------------------------------------------------
d$age_years <- as.numeric(difftime(d$icu_in, as.POSIXct(g("dob")), units = "days")) / 365.25
d$sex <- factor(toupper(trimws(as.character(g("sex")))),
                levels = c("M", "F"), labels = c("Male", "Female"))
d$weight_kg <- as.numeric(g("weight"))
d$height_cm <- as.numeric(g("height"))
d$bmi       <- as.numeric(g("bmi"))
d$bmi_recomputed <- d$weight_kg / (d$height_cm / 100)^2
d$bmi_discrepant <- is.finite(d$bmi) & is.finite(d$bmi_recomputed) &
                    abs(d$bmi - d$bmi_recomputed) > 2

## --- 3.2  Smoking ------------------------------------------------------------
## The export mixes two vocabularies; harmonised to Never / Former / Current.
sm <- tolower(trimws(as.character(g("smoking"))))
d$smoking <- factor(NA, levels = c("Never smoker", "Former smoker", "Current smoker"))
d$smoking[grepl("^never", sm)]   <- "Never smoker"
d$smoking[grepl("^former", sm)]  <- "Former smoker"
d$smoking[grepl("^current smoker", sm)] <- "Current smoker"
## "Passive smoker" (n=9) and "Currently doesn't use tobacco..." (n=2) are not
## active smokers -> grouped with Never; the raw wording is kept below.
d$smoking[grepl("^passive", sm)] <- "Never smoker"
d$smoking[grepl("^currently doesn", sm)] <- "Never smoker"
d$smoking_raw     <- factor(as.character(g("smoking")))
d$current_smoker  <- d$smoking == "Current smoker"

## --- 3.3  Comorbidity --------------------------------------------------------
dm <- trimws(as.character(g("diabetes")))
d$diabetes3 <- factor(NA, levels = c("No diabetes", "NIDDM", "IDDM"))
d$diabetes3[grepl("geen", dm, ignore.case = TRUE)] <- "No diabetes"
d$diabetes3[grepl("^NIDDM", dm, ignore.case = TRUE)] <- "NIDDM"
d$diabetes3[grepl("^IDDM",  dm, ignore.case = TRUE)] <- "IDDM"
if (DM_MISSING_AS_NONE) d$diabetes3[is.na(d$diabetes3)] <- "No diabetes"
d$diabetes_any <- d$diabetes3 %in% c("NIDDM", "IDDM")
d$diabetes_any[is.na(d$diabetes3)] <- NA

d$prior_cva_tia  <- as.integer(g("cva_tia")) == 1
d$lung_disease   <- as.integer(g("lung_disease")) == 1
d$hypertension   <- as.integer(g("hypertension")) == 1
## actieve_endocarditis has explicit "No" values, so a blank is truly missing
## here (unlike the presence-only columns further down).
if (!is.na(COLMAP[["endocarditis"]])) d$endocarditis <- ja_nee(g("endocarditis"))

## --- 3.4  Urgency ------------------------------------------------------------
urg <- toupper(trimws(as.character(g("urgency"))))
urg[urg %in% c("", "NA")] <- NA
d$urgency_grade <- factor(urg, levels = c("E", "S3", "S2", "S1", "NG"))
d$elective <- NA
d$elective[urg %in% c("S1", "S2", "S3")] <- FALSE
d$elective[urg == "E"] <- TRUE
if (ELECTIVE_DEFINITION == "blank_is_elective") d$elective[is.na(urg)] <- TRUE
d$elective_f <- factor(ifelse(is.na(d$elective), NA, ifelse(d$elective, "Elective", "Non-elective")),
                       levels = c("Elective", "Non-elective"))

## --- 3.5  Type of surgery ----------------------------------------------------
## 'verrichting' is free text (213 distinct strings, several procedures joined
## by ";").  Classification is keyword based and fully auditable: every distinct
## string with its assigned group is written to
## output/verrichting_classification_AUDIT.csv - please review it.
##
##   CABG      : any "CAB" (CABG, OPCAB, MIDCAB, Robot geassisteerde Midcab, ...)
##               No other procedure name in this file contains "cab", so a plain
##               substring match is safe (Carotis / Carotisendarterectomie do not
##               match).
##   Valve     : surgical valve procedures only.  Bentall (composite valve
##               conduit) and Ross count as valve; "klepinterventie" does not,
##               because in this file it only ever accompanies TAVI.
RX_CABG  <- "cab"
RX_VALVE <- paste(
  "klep\\(plastie\\)", "aortaklep", "mitralisklep", "mitralissklep",
  "tricuspidklep", "tricuspiedklep", "tricus\\s*pied", "tricuspidplastie",
  "mitralisplastie", "mitralisklepplastie", "aortaklepvervanging",
  "tricuspidklepvervanging", "\\bavr\\b", "\\bmvr\\b", "\\btvr\\b",
  "\\bavp\\b", "\\bmvp\\b", "\\btvp\\b", "bentall", "ross", sep = "|")
RX_TAVI  <- "tavi"
RX_DAVID <- "david|tirone"                       # valve-sparing root replacement
RX_AORTA <- "ascendens|aortadissectie|endoprothese|bentall|david|tirone"

proc <- tolower(trimws(as.character(g("procedure"))))
d$procedure_raw  <- as.character(g("procedure"))
d$has_cabg       <- grepl(RX_CABG,  proc)
d$has_surg_valve <- grepl(RX_VALVE, proc)
d$has_tavi       <- grepl(RX_TAVI,  proc)
d$has_root_spar  <- grepl(RX_DAVID, proc) & !grepl("bentall", proc)
d$has_aorta      <- grepl(RX_AORTA, proc)

valve_for_group <- d$has_surg_valve |
                   (TAVI_AS_VALVE  & d$has_tavi) |
                   (DAVID_AS_VALVE & grepl(RX_DAVID, proc))

d$surg_group <- factor(
  ifelse( d$has_cabg &  valve_for_group, "CABG + valve",
  ifelse( d$has_cabg & !valve_for_group, "CABG only",
  ifelse(!d$has_cabg &  valve_for_group, "Valve only (no CABG)", "Other"))),
  levels = c("CABG only", "CABG + valve", "Valve only (no CABG)", "Other"))

## What is inside "Other"?  Reported so nothing is hidden in a residual box.
d$other_detail <- factor(
  ifelse(d$surg_group != "Other", NA,
  ifelse(d$has_tavi,      "TAVI (transcatheter)",
  ifelse(d$has_root_spar, "Valve-sparing root (David/Tirone-David)",
  ifelse(d$has_aorta,     "Aortic surgery (no valve, no CABG)",
                          "Other cardiac / non-cardiac")))),
  levels = c("TAVI (transcatheter)", "Valve-sparing root (David/Tirone-David)",
             "Aortic surgery (no valve, no CABG)", "Other cardiac / non-cardiac"))

## records where no keyword matched at all -> flagged for manual review
d$proc_unclassified <- !(d$has_cabg | d$has_surg_valve | d$has_tavi |
                         grepl(RX_DAVID, proc) | d$has_aorta)

if (!is.na(COLMAP[["approach"]]))
  d$approach <- factor(gsub("_", " ", as.character(g("approach"))),
                       levels = c("full sternotomie", "minimaal invasief"),
                       labels = c("Full sternotomy", "Minimally invasive"))
if (!is.na(COLMAP[["cpb_min"]]))    d$cpb_min    <- as.numeric(g("cpb_min"))
if (!is.na(COLMAP[["xclamp_min"]])) d$xclamp_min <- as.numeric(g("xclamp_min"))

## --- 3.6  Mortality ----------------------------------------------------------
## 'mortaliteit' holds ONE mutually exclusive label per patient; blank = alive.
##   overleden_IZ      -> died on the ICU
##   30D mortaliteit   -> died <=30 d, after ICU discharge  (i.e. NOT on the ICU)
##   180D mortaliteit  -> died 31-180 d
mo <- trimws(as.character(g("mortality")))
d$mort_cat <- factor("Alive at 180 days",
  levels = c("Alive at 180 days", "ICU death",
             "Death <=30 days (after ICU discharge)", "Death 31-180 days"))
d$mort_cat[grepl("overleden_?IZ", mo, ignore.case = TRUE)] <- "ICU death"
d$mort_cat[grepl("^30D",  mo, ignore.case = TRUE)] <- "Death <=30 days (after ICU discharge)"
d$mort_cat[grepl("^180D", mo, ignore.case = TRUE)] <- "Death 31-180 days"

d$death_icu  <- d$mort_cat == "ICU death"
## Cumulative figures assume every ICU death occurred within 30 days of surgery
## (true unless a patient died on the ICU after day 30 - check with QC below).
d$death_30d  <- d$mort_cat %in% c("ICU death", "Death <=30 days (after ICU discharge)")
d$death_180d <- d$mort_cat != "Alive at 180 days"

## --- 3.7  Acute kidney injury ------------------------------------------------
d$crea_pre <- as.numeric(g("crea_pre"))
d$crea_max <- as.numeric(g("crea_max"))
CREA_UNIT <- if (median(d$crea_max, na.rm = TRUE) > 30) "umol/L" else "mg/dL"
ABS_RISE  <- if (is.finite(AKI_ABS_RISE)) AKI_ABS_RISE else
             if (CREA_UNIT == "umol/L") 26.5 else 0.3
cat(sprintf("\nCreatinine unit detected: %s  ->  absolute-rise threshold = %.4g %s\n",
            CREA_UNIT, ABS_RISE, CREA_UNIT))

d$crea_ratio <- d$crea_max / d$crea_pre
d$crea_delta <- d$crea_max - d$crea_pre
ok <- is.finite(d$crea_ratio) & is.finite(d$crea_delta)

stage <- rep(NA_integer_, nrow(d))
stage[ok] <- 0L
stage[ok & (d$crea_ratio >= 1.5 | d$crea_delta >= ABS_RISE)] <- 1L
stage[ok &  d$crea_ratio >= 2   & d$crea_ratio < 3]          <- 2L
stage[ok &  d$crea_ratio >= 3]                               <- 3L
d$aki_stage <- factor(stage, levels = 0:3,
                      labels = c("No AKI", "AKI stage 1", "AKI stage 2", "AKI stage 3"))
d$aki_any <- ifelse(is.na(stage), NA, stage > 0)

## Time-windowed sensitivity analysis (only the DATE of the peak is available).
d$crea_max_lag_days <- as.numeric(as.Date(g("crea_max_date")) - as.Date(d$surg_end))
in_abs <- is.finite(d$crea_max_lag_days) & d$crea_max_lag_days <= AKI_WINDOW_ABS_DAYS
in_rat <- is.finite(d$crea_max_lag_days) & d$crea_max_lag_days <= AKI_WINDOW_RATIO_DAYS
stage_w <- rep(NA_integer_, nrow(d))
stage_w[ok] <- 0L
stage_w[ok & ((in_rat & d$crea_ratio >= 1.5) | (in_abs & d$crea_delta >= ABS_RISE))] <- 1L
stage_w[ok & in_rat & d$crea_ratio >= 2 & d$crea_ratio < 3] <- 2L
stage_w[ok & in_rat & d$crea_ratio >= 3]                    <- 3L
d$aki_stage_windowed <- factor(stage_w, levels = 0:3,
  labels = c("No AKI", "AKI stage 1", "AKI stage 2", "AKI stage 3"))

## --- 3.8  Invasive ventilation ----------------------------------------------
ext <- trimws(as.character(g("extub_status")))
d$extub_status <- factor(ext,
  levels = c("on-table extubation", "inv beademing op IZ", "geen inv beademing op OK"),
  labels = c("Extubated on table", "Ventilated on ICU", "No invasive ventilation in OR"))
d$vent_min_recorded <- as.numeric(g("vent_min"))
if (!is.na(COLMAP[["vent_event"]])) d$vent_event_min <- as.numeric(g("vent_event"))

## Whole-cohort ventilation time: extubated on table (and no invasive
## ventilation in theatre) contribute 0 minutes rather than a missing value.
d$vent_min_all <- d$vent_min_recorded
d$vent_min_all[d$extub_status == "Extubated on table"] <- 0
d$vent_min_all[d$extub_status == "No invasive ventilation in OR" &
               is.na(d$vent_min_recorded)] <- 0

## Time from END OF SURGERY to extubation.  The ventilation clock in this export
## runs from ICU admission, which is a median of 7 min after surgery end
## (IQR -6 to +18), so that offset is added back.
## In 602 records the ICU admission is registered a few minutes BEFORE the
## recorded end of surgery, so the sum is clamped at 0 (a negative time to
## extubation is not meaningful).
d$icu_offset_min    <- as.numeric(difftime(d$icu_in, d$surg_end, units = "mins"))
d$time_to_extub_min <- pmax(0, d$icu_offset_min + d$vent_min_all)
d$time_to_extub_min[d$extub_status == "Extubated on table"] <- 0
d$time_to_extub_min[d$vent_min_all == 0 & !is.na(d$vent_min_all)] <- 0

d$vent_gt6h        <- d$time_to_extub_min  > VENT_THRESHOLD_MIN   # primary
d$vent_gt6h_simple <- d$vent_min_all       > VENT_THRESHOLD_MIN   # ignores offset

## --- 3.9  ICU readmission ----------------------------------------------------
## 'Heropname' is a presence-only column already banded as requested;
## blank = no readmission.
rd <- trimws(as.character(g("readmission")))
d$icu_readmission <- present(rd)
d$icu_readm_cat <- factor(
  ifelse(!d$icu_readmission, "No ICU readmission",
  ifelse(grepl("<\\s*24", rd), "Readmission < 24 h",
  ifelse(grepl("24.*48", rd), "Readmission 24-48 h",
  ifelse(grepl(">\\s*48", rd), "Readmission > 48 h", NA)))),
  levels = c("No ICU readmission", "Readmission < 24 h",
             "Readmission 24-48 h", "Readmission > 48 h"))

## --- 3.10  Pain / PCIA -------------------------------------------------------
d$mean_vas_24h <- as.numeric(g("vas24"))
d$pcia <- factor(ifelse(present(g("pcia")), "PCIA", "No PCIA"),
                 levels = c("No PCIA", "PCIA"))

## --- 3.11  Length of stay and mobilisation -----------------------------------
## 'Ligduur' equals Ontslag - Opname exactly, and Opname falls a median of 7 min
## after the end of surgery => this is the ICU / post-operative unit stay.
## A HOSPITAL length of stay is NOT present in this export (see QC report).
d$icu_los_days  <- as.numeric(g("los_raw"))
d$icu_los_check <- as.numeric(difftime(d$icu_out, d$icu_in, units = "days"))
if (!is.na(COLMAP[["los_caldays"]])) d$icu_los_caldays <- as.numeric(g("los_caldays"))
d$hosp_los_days <- NA_real_          # <- fill in if a hospital LOS is added

d$time_to_chair_min <- as.numeric(g("time_to_chair"))
d$time_to_chair_h   <- d$time_to_chair_min / 60
d$chair_within_24h  <- d$time_to_chair_h <= 24

## --- 3.12  Mechanical circulatory support (for QC / context) -----------------
mcs_cols <- c("ecmo_days", "impella_days", "iabp_days")
mcs_cols <- mcs_cols[!is.na(COLMAP[mcs_cols])]
if (length(mcs_cols))
  d$mcs_any <- Reduce(`|`, lapply(mcs_cols, function(k) as.numeric(g(k)) > 0))

## --- 3.13  Analysis population ----------------------------------------------
d_all <- d
if (identical(ANALYSIS_UNIT, "first_per_patient")) d <- d[d$is_first_episode, ]
cat(sprintf("\nAnalysis unit: %s  ->  n = %d records (%d unique patients)\n",
            ANALYSIS_UNIT, nrow(d), length(unique(d$pat_id))))


# =============================================================================
# 4.  DATA QUALITY REPORT
# =============================================================================

hdr("3. DATA QUALITY")

qc <- function(check, value, comment = "")
  data.frame(check = check, value = as.character(value), comment = comment,
             stringsAsFactors = FALSE)

QC <- rbind(
  qc("Records in file", nrow(d_all)),
  qc("Unique patients", length(unique(d_all$pat_id)),
     "patients with >1 record contribute a re-operation / second ICU episode"),
  qc("Patients with >1 record", sum(table(d_all$pat_id) > 1)),
  qc("Duplicated Opnamenummer (same hospital admission, 2 procedures)",
     sum(duplicated(d_all$adm_id))),
  qc("Surgery date range", paste(format(range(d_all$surg_start), "%Y-%m-%d"), collapse = " to ")),
  qc("Analysis unit / n analysed", sprintf("%s / %d", ANALYSIS_UNIT, nrow(d))),

  qc("--- Timing consistency ---", ""),
  qc("ICU admission before end of surgery",
     sprintf("%d (%.1f%%)", sum(d$icu_offset_min < 0, na.rm = TRUE),
             100 * mean(d$icu_offset_min < 0, na.rm = TRUE)),
     sprintf("median offset ICU-in minus surgery-end = %.0f min (IQR %.0f to %.0f)",
             median(d$icu_offset_min, na.rm = TRUE),
             quantile(d$icu_offset_min, .25, na.rm = TRUE),
             quantile(d$icu_offset_min, .75, na.rm = TRUE))),
  qc("Ligduur equals Ontslag - Opname",
     ifelse(max(abs(d$icu_los_days - d$icu_los_check), na.rm = TRUE) < 1e-6, "yes", "NO"),
     "=> Ligduur is the ICU / post-operative unit stay, not the hospital stay"),
  qc("Hospital length of stay available", "NO",
     "no hospital admission/discharge column in this export - see notes"),
  qc("Ventilation time longer than ICU stay",
     sum(d$vent_min_all > as.numeric(difftime(d$icu_out, d$icu_in, units = "mins")), na.rm = TRUE)),
  qc("Time to chair longer than ICU stay",
     sum(d$time_to_chair_min > as.numeric(difftime(d$icu_out, d$icu_in, units = "mins")), na.rm = TRUE)),
  qc("ICU stay < 60 min", sum(d$icu_los_days * 1440 < 60, na.rm = TRUE)),

  qc("--- Mortality ---", ""),
  qc("ICU deaths with ICU stay > 30 days", sum(d$death_icu & d$icu_los_days > 30, na.rm = TRUE),
     "if > 0 the cumulative 30-day figure is an over-estimate"),

  qc("--- Creatinine / AKI ---", ""),
  qc("Creatinine unit detected", CREA_UNIT, sprintf("absolute-rise threshold %.4g", ABS_RISE)),
  qc("Pre-operative creatinine missing",
     sprintf("%d (%.1f%%)", sum(is.na(d$crea_pre)), 100 * mean(is.na(d$crea_pre))),
     "AKI cannot be staged in these records"),
  qc("Peak creatinine missing", sum(is.na(d$crea_max))),
  qc("Peak creatinine BELOW pre-operative value",
     sum(d$crea_ratio < 1, na.rm = TRUE), "post-operative haemodilution - staged as no AKI"),
  qc("Peak creatinine recorded > 7 days after surgery",
     sum(d$crea_max_lag_days > 7, na.rm = TRUE),
     "excluded from the KDIGO time-windowed sensitivity analysis"),

  qc("--- Anthropometry ---", ""),
  qc("Recorded BMI differs from weight/height^2 by > 2 kg/m2",
     sum(d$bmi_discrepant, na.rm = TRUE),
     sprintf("max difference %.1f kg/m2 - recorded 'bmi' is used",
             max(abs(d$bmi - d$bmi_recomputed), na.rm = TRUE))),

  qc("--- Procedure coding ---", ""),
  qc("Distinct 'verrichting' strings", length(unique(d$procedure_raw))),
  qc("Records where no procedure keyword matched",
     sprintf("%d (%.1f%%)", sum(d$proc_unclassified), 100 * mean(d$proc_unclassified)),
     "classified as 'Other' - listed in the audit file"),
  qc("TAVI records", sum(d$has_tavi),
     sprintf("TAVI_AS_VALVE = %s", TAVI_AS_VALVE)),
  qc("Valve-sparing root (David/Tirone) records", sum(d$has_root_spar),
     sprintf("DAVID_AS_VALVE = %s", DAVID_AS_VALVE)),

  qc("--- Missing data in analysis variables ---", "")
)

miss_vars <- c(age_years = "Age", sex = "Sex", elective_f = "Elective status",
               bmi = "BMI", smoking = "Smoking status", prior_cva_tia = "Previous CVA/TIA",
               lung_disease = "Chronic lung disease", surg_group = "Type of surgery",
               diabetes3 = "Diabetes", mort_cat = "Mortality status",
               aki_stage = "AKI stage", vent_min_all = "Ventilation time",
               icu_readm_cat = "ICU readmission", mean_vas_24h = "Mean VAS 24 h",
               icu_los_days = "ICU length of stay", time_to_chair_min = "Time to chair")
for (v in names(miss_vars))
  QC <- rbind(QC, qc(paste0("   ", miss_vars[[v]], " (", v, ")"),
                     sprintf("%d (%.1f%%)", sum(is.na(d[[v]])), 100 * mean(is.na(d[[v]])))))

print(QC, row.names = FALSE, right = FALSE)
writeLines(capture.output(print(QC, row.names = FALSE, right = FALSE)),
           file.path(OUT_DIR, "data_quality.txt"))

## Audit file: every distinct procedure string with its assigned group ---------
audit <- aggregate(list(n = d$row_id),
                   by = list(verrichting = d$procedure_raw, group = d$surg_group,
                             detail = addNA(d$other_detail),
                             CABG = d$has_cabg, surgical_valve = d$has_surg_valve,
                             TAVI = d$has_tavi, root_sparing = d$has_root_spar,
                             no_keyword_matched = d$proc_unclassified),
                   FUN = length)
audit <- audit[order(-audit$n), c("n", "group", "detail", "verrichting", "CABG",
                                  "surgical_valve", "TAVI", "root_sparing",
                                  "no_keyword_matched")]
write.csv(audit, file.path(OUT_DIR, "verrichting_classification_AUDIT.csv"),
          row.names = FALSE, na = "")
cat(sprintf("\nProcedure audit written: %s (%d distinct strings)\n",
            file.path(OUT_DIR, "verrichting_classification_AUDIT.csv"), nrow(audit)))


# =============================================================================
# 5.  TABLE ENGINE (overall + optional stratified)
# =============================================================================

## A table is defined once as a list of specs and can then be rendered either
## overall or stratified by a grouping variable.
##   var    - column in d
##   label  - row label
##   type   - "num" | "cat" | "bin"
##   digits - decimals for numeric summaries
##   force  - "mean_sd" / "median_iqr" to override the automatic choice
sp <- function(var, label, type = "num", digits = 1, force = NULL, yes = TRUE)
  list(var = var, label = label, type = type, digits = digits, force = force, yes = yes)

render_overall <- function(spec, data) {
  do.call(rbind, lapply(spec, function(s) {
    x <- data[[s$var]]
    if (is.null(x)) return(NULL)
    switch(s$type,
      num = describe_num(x, s$label, s$var, s$digits, s$force),
      cat = describe_cat(x, s$label),
      bin = describe_bin(x, s$label, s$yes))
  }))
}

## p-value for one variable across a grouping factor
group_p <- function(x, by, type) {
  by <- droplevels(factor(by))
  if (nlevels(by) < 2) return(NA_real_)
  out <- try({
    if (type == "num") {
      v <- suppressWarnings(as.numeric(x))
      if (choose_summary(v[is.finite(v)]) == "mean_sd")
        summary(aov(v ~ by))[[1]][["Pr(>F)"]][1]
      else kruskal.test(v ~ by)$p.value
    } else {
      tb <- table(factor(x), by)
      tb <- tb[rowSums(tb) > 0, , drop = FALSE]
      if (nrow(tb) < 2) return(NA_real_)
      e <- suppressWarnings(chisq.test(tb)$expected)
      if (any(e < 5)) fisher.test(tb, simulate.p.value = TRUE, B = 20000)$p.value
      else chisq.test(tb)$p.value
    }
  }, silent = TRUE)
  if (inherits(out, "try-error")) NA_real_ else as.numeric(out)
}

render_by <- function(spec, data, by_var) {
  by <- droplevels(factor(data[[by_var]]))
  ## Keep only variables that exist, so spec rows and table rows stay aligned.
  spec <- Filter(function(s) !is.null(data[[s$var]]), spec)
  ## Fix the summary method on the WHOLE cohort and force it in every stratum:
  ## a small stratum would otherwise pass Shapiro-Wilk and flip that one cell to
  ## mean (SD) while the rest of the row stays median [IQR].
  spec <- lapply(spec, function(s) {
    if (s$type == "num" && is.null(s$force)) {
      v <- suppressWarnings(as.numeric(data[[s$var]]))
      s$force <- choose_summary(v[is.finite(v)], s$var)
    }
    s
  })

  parts <- lapply(levels(by), function(L) render_overall(spec, data[which(by == L), , drop = FALSE]))
  base  <- render_overall(spec, data)
  out   <- base[, c("variable", "level", "method")]
  out$Overall <- ifelse(base$summary == "", sprintf("n=%d", base$n), base$summary)
  for (i in seq_along(levels(by))) {
    p <- parts[[i]]
    out[[levels(by)[i]]] <- ifelse(p$summary == "", sprintf("n=%d", p$n), p$summary)
  }
  ## one p-value per variable, printed on that variable's header row
  out$p_value <- ""
  hdr_rows <- which(out$variable != "")
  for (k in seq_along(spec)) {
    r <- hdr_rows[k]
    if (is.na(r)) next
    out$p_value[r] <- pval_fmt(group_p(data[[spec[[k]]$var]], data[[by_var]],
                                       if (spec[[k]]$type == "num") "num" else "cat"))
  }
  out
}


## small builders for hand-made rows in the outcome table ---------------------
row_hdr <- function(label)
  data.frame(variable = label, level = "", n = NA_integer_, summary = "",
             method = "", missing_n = NA_integer_, missing_pct = "",
             stringsAsFactors = FALSE)
row_pct <- function(label, k, n, note = "n (%)")
  data.frame(variable = label, level = "", n = n,
             summary = sprintf("%d (%s)", k, fmt(100 * k / n, 1)),
             method = note, missing_n = NA_integer_, missing_pct = "",
             stringsAsFactors = FALSE)


# =============================================================================
# 6.  TABLE 1 -- BASELINE CHARACTERISTICS
# =============================================================================

hdr("4. TABLE 1 - BASELINE CHARACTERISTICS")

TABLE1_SPEC <- list(
  sp("age_years",    "Age at ICU admission, years",              "num", 1),
  sp("sex",          "Sex",                                      "cat"),
  sp("elective_f",   "Urgency of surgery",                       "cat"),
  sp("bmi",          "Body mass index, kg/m2",                   "num", 1),
  sp("smoking",      "Smoking status",                           "cat"),
  sp("prior_cva_tia","Previous CVA / TIA (VG_CVA_TIA)",          "bin"),
  sp("lung_disease", "Chronic lung disease (VG_longlijden)",     "bin"),
  sp("surg_group",   "Type of surgery",                          "cat"),
  sp("diabetes3",    "Diabetes mellitus",                        "cat"),
  sp("diabetes_any", "Diabetes mellitus, any (NIDDM or IDDM)",   "bin")
)

table1 <- render_overall(TABLE1_SPEC, d)
print_table(table1, sprintf("Table 1. Baseline characteristics (n = %d)", nrow(d)))

cat("\nNotes on Table 1:\n")
cat(" - Percentages are of NON-MISSING observations; the 'Missing n' column\n",
    "   gives the number not recorded for that variable.\n", sep = "")
cat(" - Urgency: ELECTIVE_DEFINITION = '", ELECTIVE_DEFINITION, "'.\n", sep = "")
cat(" - Diabetes: DM_MISSING_AS_NONE = ", DM_MISSING_AS_NONE,
    " (", sum(is.na(d$diabetes3)), " records not recorded).\n", sep = "")
cat(" - Type of surgery: TAVI_AS_VALVE = ", TAVI_AS_VALVE,
    ", DAVID_AS_VALVE = ", DAVID_AS_VALVE, ".\n", sep = "")

## supporting detail that belongs with Table 1 but is not part of it ----------
table1_supp <- rbind(
  describe_cat(d$urgency_grade, "Urgency code as recorded (urgentie_graad)",
               na_level = "(blank - not coded)"),
  describe_cat(droplevels(d$other_detail[d$surg_group == "Other"]),
               "Composition of the 'Other' surgical group"),
  describe_cat(d$smoking_raw, "Smoking categories as recorded (rook_status)")
)
if (!is.null(d$approach))
  table1_supp <- rbind(table1_supp,
                       describe_cat(d$approach, "Surgical approach (benadering)"))
if (!is.null(d$cpb_min))
  table1_supp <- rbind(table1_supp,
                       describe_num(d$cpb_min,    "Cardiopulmonary bypass time, min (on-pump cases)", "cpb_min", 0),
                       describe_num(d$xclamp_min, "Aortic cross-clamp time, min (clamped cases)",     "xclamp_min", 0))
print_table(table1_supp, "Table 1b. Supporting detail")
cat("\nBypass and cross-clamp times are blank for off-pump (OPCAB / MIDCAB) and\n",
    "transcatheter cases: that is 'not applicable', not missing data.\n", sep = "")

table1_by <- NULL
if (!is.null(STRATIFY_BY)) {
  table1_by <- render_by(TABLE1_SPEC, d, STRATIFY_BY)
  sub_hdr(paste("Table 1, stratified by", STRATIFY_BY))
  print(table1_by, row.names = FALSE, right = FALSE)
}


# =============================================================================
# 7.  TABLE 2 -- OUTCOMES
# =============================================================================

hdr("5. TABLE 2 - OUTCOMES")

n_all <- nrow(d)

## --- 7.1 Mortality -----------------------------------------------------------
mort_tab <- rbind(
  row_hdr("MORTALITY - mutually exclusive categories (as coded)"),
  describe_cat(d$mort_cat, "Vital status"),
  row_hdr("MORTALITY - cumulative"),
  row_pct("ICU mortality",                      sum(d$death_icu),  n_all),
  row_pct("30-day mortality (ICU + post-ICU)",  sum(d$death_30d),  n_all),
  row_pct("180-day mortality (all deaths)",     sum(d$death_180d), n_all)
)

## --- 7.2 Acute kidney injury -------------------------------------------------
n_aki <- sum(!is.na(d$aki_stage))
aki_tab <- rbind(
  row_hdr(sprintf("ACUTE KIDNEY INJURY (assessable n = %d of %d)", n_aki, n_all)),
  describe_cat(d$aki_stage, "AKI stage (peak vs pre-operative creatinine)"),
  row_pct("AKI of any stage", sum(d$aki_any, na.rm = TRUE), n_aki),
  describe_num(d$crea_pre,   sprintf("Pre-operative creatinine, %s", CREA_UNIT), "crea_pre", 2),
  describe_num(d$crea_max,   sprintf("Peak creatinine, %s", CREA_UNIT),          "crea_max", 2),
  describe_num(d$crea_delta, sprintf("Rise in creatinine, %s", CREA_UNIT),       "crea_delta", 2),
  describe_num(d$crea_max_lag_days, "Days from surgery to peak creatinine", "crea_max_lag_days", 0),
  row_hdr(sprintf("Sensitivity: KDIGO time windows (rise <=%d d, ratio <=%d d)",
                  AKI_WINDOW_ABS_DAYS, AKI_WINDOW_RATIO_DAYS)),
  describe_cat(d$aki_stage_windowed, "AKI stage, time-windowed")
)

## --- 7.3 Invasive ventilation ------------------------------------------------
n_vent_icu <- sum(d$extub_status == "Ventilated on ICU", na.rm = TRUE)
vent_icu   <- d$vent_min_recorded[d$extub_status == "Ventilated on ICU"]
n_v_all    <- sum(!is.na(d$vent_gt6h))
n_v_icu    <- sum(!is.na(d$vent_gt6h) & d$extub_status == "Ventilated on ICU")
vent_tab <- rbind(
  row_hdr("INVASIVE VENTILATION"),
  describe_cat(d$extub_status, "Extubation setting"),
  describe_num(vent_icu, "Ventilation time, min - patients ventilated on ICU",
               "vent_min_recorded", 0),
  describe_num(vent_icu / 60, "Ventilation time, hours - patients ventilated on ICU",
               "vent_hours", 1),
  describe_num(d$vent_min_all,
               "Ventilation time, min - whole cohort (on-table extubation = 0)",
               "vent_min_all", 0),
  describe_num(d$time_to_extub_min,
               "Time from end of surgery to extubation, min", "time_to_extub_min", 0),
  row_hdr(sprintf("Ventilation beyond %g h after end of surgery", VENT_THRESHOLD_MIN / 60)),
  row_pct("   whole cohort", sum(d$vent_gt6h, na.rm = TRUE), n_v_all),
  row_pct("   among patients ventilated on ICU",
          sum(d$vent_gt6h & d$extub_status == "Ventilated on ICU", na.rm = TRUE), n_v_icu),
  row_pct("   whole cohort, ignoring the ICU-admission offset",
          sum(d$vent_gt6h_simple, na.rm = TRUE), sum(!is.na(d$vent_gt6h_simple)))
)
if (!is.null(d$vent_event_min))
  vent_tab <- rbind(vent_tab, describe_num(d$vent_event_min,
    "Ventilation time per event, min (duur_invbead_event)", "vent_event_min", 0))

## --- 7.4 ICU readmission -----------------------------------------------------
readm_tab <- rbind(
  row_hdr("ICU READMISSION"),
  row_pct("Readmitted to ICU", sum(d$icu_readmission), n_all),
  describe_cat(d$icu_readm_cat, "Time from ICU discharge to readmission"),
  describe_cat(droplevels(d$icu_readm_cat[d$icu_readmission]),
               "Distribution among readmitted patients only")
)

## --- 7.5 Pain ----------------------------------------------------------------
## The two PCIA sub-groups must be summarised the SAME way (n = 14 alone would
## pass Shapiro-Wilk and be reported as a mean), so the method chosen for the
## whole variable is forced on both rows.
vas_method <- choose_summary(d$mean_vas_24h[is.finite(d$mean_vas_24h)], "mean_vas_24h")
vas_tab <- rbind(
  row_hdr("POST-OPERATIVE PAIN"),
  describe_num(d$mean_vas_24h, "Mean VAS over the first 24 h", "mean_vas_24h", 2),
  describe_cat(d$pcia, "Patient-controlled IV analgesia"),
  describe_num(d$mean_vas_24h[d$pcia == "No PCIA"], "   mean VAS 24 h - no PCIA",
               "vas_nopcia", 2, force = vas_method),
  describe_num(d$mean_vas_24h[d$pcia == "PCIA"],    "   mean VAS 24 h - PCIA",
               "vas_pcia", 2, force = vas_method)
)

## --- 7.6 Length of stay and mobilisation -------------------------------------
los_tab <- rbind(
  row_hdr("LENGTH OF STAY AND MOBILISATION"),
  describe_num(d$icu_los_days, "ICU length of stay, days", "icu_los_days", 2),
  describe_num(d$icu_los_days * 24, "ICU length of stay, hours", "icu_los_hours", 1),
  describe_num(d$hosp_los_days, "Hospital length of stay, days - NOT IN THIS EXPORT",
               "hosp_los_days", 1),
  describe_num(d$time_to_chair_min, "Time to first sitting in chair, min", "time_to_chair_min", 0),
  describe_num(d$time_to_chair_h,   "Time to first sitting in chair, hours", "time_to_chair_h", 1),
  row_pct("Mobilised to chair within 24 h", sum(d$chair_within_24h, na.rm = TRUE),
          sum(!is.na(d$chair_within_24h)))
)
if (!is.null(d$icu_los_caldays))
  los_tab <- rbind(los_tab, describe_num(d$icu_los_caldays,
    "ICU stay in calendar days (Ligduur in ligdagen)", "icu_los_caldays", 0))

table2 <- rbind(mort_tab, aki_tab, vent_tab, readm_tab, vas_tab, los_tab)
print_table(table2, sprintf("Table 2. Outcomes (n = %d)", n_all))

## --- footnotes that a reader (or reviewer) will ask for ----------------------
n_icu_late <- sum(d$death_icu & d$icu_los_days > 30, na.rm = TRUE)
cat("\nNotes on Table 2:\n")
cat(" 1. 'mortaliteit' codes ONE mutually exclusive outcome per patient, so\n",
    "    'Death <=30 days' means died within 30 days AFTER ICU discharge and\n",
    "    'Death 31-180 days' means died later; the cumulative rows add them up.\n", sep = "")
if (n_icu_late > 0)
  cat("    ", n_icu_late, " ICU death(s) occurred after an ICU stay of >30 days, so the\n",
      "    cumulative 30-day figure is at most that many deaths too high.\n", sep = "")
cat(" 2. AKI is staged from peak vs pre-operative creatinine only. Urine output\n",
    "    and renal replacement therapy are not in this export, so stage 3 by RRT\n",
    "    and any oliguric AKI are not captured - these rates are a lower bound.\n", sep = "")
cat(" 3. Ventilation time is measured from ICU admission; the row 'time from end\n",
    "    of surgery to extubation' adds the ICU-admission offset back on. Patients\n",
    "    extubated in theatre count as 0 min, not as missing.\n", sep = "")
cat(" 4. ICU length of stay is Ontslag - Opname. A HOSPITAL length of stay is not\n",
    "    available in this export (see the data quality report).\n", sep = "")


# =============================================================================
# 8.  COMPARISON: MEAN VAS OVER 24 h, PCIA vs NO PCIA
# =============================================================================

hdr("6. MEAN VAS 24 h - PCIA vs NO PCIA")

## Two-group comparison of a numeric outcome.  Reports both a rank-based
## (Wilcoxon rank-sum, Hodges-Lehmann shift + 95% CI) and a parametric
## (Welch t) analysis, so the conclusion can be checked against either.
compare_two <- function(x, grp, ref, alt, outcome_label, digits = 2) {
  x <- as.numeric(x); grp <- as.character(grp)
  a <- x[grp == alt & is.finite(x)]      # group of interest
  b <- x[grp == ref & is.finite(x)]      # reference group
  w  <- suppressWarnings(try(wilcox.test(a, b, conf.int = TRUE, exact = FALSE), silent = TRUE))
  tt <- suppressWarnings(try(t.test(a, b), silent = TRUE))
  q <- function(v) quantile(v, c(.25, .5, .75), na.rm = TRUE, type = 7)
  data.frame(
    Outcome = outcome_label,
    Group   = c(ref, alt, "Difference (Wilcoxon)", "Difference (Welch t)"),
    n       = c(length(b), length(a), NA, NA),
    `Median [IQR]` = c(sprintf("%s [%s-%s]", fmt(q(b)[2], digits), fmt(q(b)[1], digits), fmt(q(b)[3], digits)),
                       sprintf("%s [%s-%s]", fmt(q(a)[2], digits), fmt(q(a)[1], digits), fmt(q(a)[3], digits)),
                       "", ""),
    `Mean (SD)` = c(sprintf("%s (%s)", fmt(mean(b), digits), fmt(sd(b), digits)),
                    sprintf("%s (%s)", fmt(mean(a), digits), fmt(sd(a), digits)),
                    "", ""),
    Estimate = c("", "",
      if (inherits(w, "try-error")) "" else sprintf("%s (Hodges-Lehmann shift)", fmt(unname(w$estimate), digits)),
      if (inherits(tt, "try-error")) "" else sprintf("%s (difference in means)", fmt(unname(diff(rev(tt$estimate))), digits))),
    `95% CI` = c("", "",
      if (inherits(w, "try-error")) "" else sprintf("%s to %s", fmt(w$conf.int[1], digits), fmt(w$conf.int[2], digits)),
      if (inherits(tt, "try-error")) "" else sprintf("%s to %s", fmt(tt$conf.int[1], digits), fmt(tt$conf.int[2], digits))),
    p = c("", "",
      if (inherits(w, "try-error")) "" else pval_fmt(w$p.value),
      if (inherits(tt, "try-error")) "" else pval_fmt(tt$p.value)),
    check.names = FALSE, stringsAsFactors = FALSE)
}

vas_pcia <- compare_two(d$mean_vas_24h, d$pcia, ref = "No PCIA", alt = "PCIA",
                        outcome_label = "Mean VAS, first 24 h")
print(vas_pcia, row.names = FALSE, right = FALSE)

n_pcia <- sum(d$pcia == "PCIA" & is.finite(d$mean_vas_24h))
cat("\nINTERPRET WITH CARE:\n")
cat(" - Only ", n_pcia, " patients with PCIA have a VAS score, against ",
    sum(d$pcia == "No PCIA" & is.finite(d$mean_vas_24h)), " without.\n", sep = "")
cat(" - PCIA is not randomised: it is prescribed to patients who are expected to\n",
    "   have, or already have, more pain. A higher VAS in the PCIA group is\n",
    "   confounding by indication, not evidence that PCIA works less well.\n", sep = "")
cat(" - 'PCIA' is a presence-only column (16 'ja', the rest blank), so 'No PCIA'\n",
    "   means 'not recorded as having had PCIA'.\n", sep = "")

## Same comparison for the other pain endpoints, if present in the file -------
extra_pain <- c(mean_VAS_12h = "Mean VAS, first 12 h",
                max_VAS_12h  = "Maximum VAS, first 12 h",
                mean_VAS_na_extubatie = "Mean VAS after extubation")
for (v in names(extra_pain)) {
  if (v %in% names(raw))
    vas_pcia <- rbind(vas_pcia,
      compare_two(raw[[v]][match(d$row_id, d_all$row_id)], d$pcia,
                  "No PCIA", "PCIA", extra_pain[[v]]))
}


# =============================================================================
# 9.  SUPPLEMENTARY: BOTH SUMMARIES FOR EVERY NUMERIC VARIABLE
# =============================================================================

hdr("7. SUPPLEMENTARY - mean (SD) AND median [IQR] side by side")

NUM_VARS <- c(age_years = "Age, years", bmi = "BMI, kg/m2",
              weight_kg = "Weight, kg", height_cm = "Height, cm",
              surg_dur_min = "Duration of surgery, min",
              crea_pre = "Pre-operative creatinine", crea_max = "Peak creatinine",
              crea_delta = "Rise in creatinine", crea_ratio = "Peak/pre-op creatinine ratio",
              vent_min_recorded = "Ventilation time (recorded), min",
              vent_min_all = "Ventilation time (whole cohort), min",
              time_to_extub_min = "Time surgery end to extubation, min",
              mean_vas_24h = "Mean VAS 24 h",
              icu_los_days = "ICU length of stay, days",
              time_to_chair_min = "Time to chair, min",
              time_to_chair_h = "Time to chair, hours")
NUM_VARS <- NUM_VARS[names(NUM_VARS) %in% names(d)]

tableS1 <- do.call(rbind, lapply(names(NUM_VARS), function(v) {
  x <- suppressWarnings(as.numeric(d[[v]])); x <- x[is.finite(x)]
  if (!length(x)) return(NULL)
  q <- quantile(x, c(.25, .5, .75), type = 7)
  sw <- try(shapiro.test(if (length(x) > 5000) sample(x, 5000) else x)$p.value, silent = TRUE)
  data.frame(
    Variable    = NUM_VARS[[v]],
    n           = length(x),
    missing     = sum(is.na(suppressWarnings(as.numeric(d[[v]])))),
    `Mean (SD)` = sprintf("%s (%s)", fmt(mean(x), 2), fmt(sd(x), 2)),
    `Median [IQR]` = sprintf("%s [%s-%s]", fmt(q[2], 2), fmt(q[1], 2), fmt(q[3], 2)),
    Min = fmt(min(x), 2), Max = fmt(max(x), 2),
    Skewness = fmt(skewness(x), 2),
    `Shapiro p` = if (inherits(sw, "try-error")) "" else pval_fmt(as.numeric(sw)),
    `Reported as` = if (choose_summary(x, v) == "mean_sd") "mean (SD)" else "median [IQR]",
    check.names = FALSE, stringsAsFactors = FALSE)
}))
print(tableS1, row.names = FALSE, right = FALSE)
cat("\nEvery variable above is skewed enough (or fails Shapiro-Wilk) that the\n",
    "median [IQR] is reported in Tables 1 and 2 unless 'Reported as' says otherwise.\n", sep = "")


# =============================================================================
# 10.  EXPORT
# =============================================================================

hdr("8. EXPORT")

tidy_out <- function(tb) {
  o <- data.frame(Characteristic = ifelse(tb$level == "", tb$variable, paste0("   ", tb$level)),
                  N = tb$n, Summary = tb$summary, Method = tb$method,
                  `Missing n` = tb$missing_n, `Missing %` = tb$missing_pct,
                  check.names = FALSE, stringsAsFactors = FALSE)
  o
}

sheets <- list(
  `Table 1 baseline`   = tidy_out(table1),
  `Table 1b supporting`= tidy_out(table1_supp),
  `Table 2 outcomes`   = tidy_out(table2),
  `VAS by PCIA`        = vas_pcia,
  `Table S1 numeric`   = tableS1,
  `Procedure audit`    = audit,
  `Data quality`       = QC
)
if (!is.null(table1_by)) sheets[[paste("Table 1 by", STRATIFY_BY)]] <- table1_by

for (nm in names(sheets)) {
  f <- file.path(OUT_DIR, paste0(gsub("[^A-Za-z0-9]+", "_", nm), ".csv"))
  write.csv(sheets[[nm]], f, row.names = FALSE, na = "")
  cat(" written: ", f, "\n", sep = "")
}

if (HAVE_WRITEXL) {
  xl <- file.path(OUT_DIR, "Cardiochirurgie_results.xlsx")
  writexl::write_xlsx(sheets, xl)
  cat(" written: ", xl, "\n", sep = "")
} else {
  cat(" (install.packages('writexl') to also get a single Excel workbook)\n")
}

saveRDS(d,     file.path(OUT_DIR, "analysis_dataset.rds"))
saveRDS(d_all, file.path(OUT_DIR, "analysis_dataset_all_records.rds"))
cat(" written: ", file.path(OUT_DIR, "analysis_dataset.rds"),
    " (derived variables, for further modelling)\n", sep = "")

hdr("DONE")
cat("Session: ", R.version.string, " | ", format(Sys.time(), "%Y-%m-%d %H:%M"), "\n", sep = "")
