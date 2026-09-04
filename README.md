# Cardiochirurgie AZ Delta — descriptive analysis

`Cardiochirurgie.R` produces the baseline table and the outcome tables for the
cardiac surgery ICU cohort (`iz_cardio_versie_04092026.xlsx`, 2066 records,
surgery 2020-06-09 → 2026-08-20). **31 TAVI procedures are excluded**, leaving
**2035 records in 2006 patients** in every figure reported.

Run it from the project root:

```r
# Session > Set Working Directory > To Source File Location
source("Cardiochirurgie.R")
```

Only `readxl` is required; `writexl` is optional (single Excel workbook).
Everything else is base R, so there is no tidyverse version to keep in step.
Results are printed to the console and written to `output/`.

---

## ⚠ Patient data must not leave the hospital drive

`Patiëntennummer` is birth-date derived — in 1215 of the 1674 ten-digit values
(72.6 %) the first six digits are exactly the patient's date of birth. Together
with `Geboortedatum` and the admission timestamps this is **directly identifying
personal data**, not a pseudonymised key.

`.gitignore` therefore excludes `data-raw/`, `output/` and all spreadsheet
formats. Only code is committed. Keep the export on the hospital drive, and
drop `Patiëntennummer` and `Geboortedatum` from any file you share (age is
already derived, so the birth date is not needed downstream).

---

## What the data actually is

| | |
|---|---|
| Unit of analysis | **one ICU admission after cardiac surgery**, not one patient |
| Records / patients | 2066 records in 2037 patients — 29 patients have a second record (re-operation: `Revisie`, `Thoraxdrain`, `Pericardfenestratie`) |
| Exclusion | `EXCLUDE_TAVI <- TRUE` drops the 31 TAVI records → **2035 analysed** |
| `Opname/Ontslag Datetime` | **ICU** admission/discharge. `Opname` falls a median of 5 min after the end of surgery (IQR −11 to +18), and `Ligduur` = `Ontslag − Opname` exactly |
| `Ligduur` | **ICU** length of stay in days |
| `Ontslag opname` | **hospital** discharge — never earlier than the ICU discharge |

Set `ANALYSIS_UNIT <- "first_per_patient"` to restrict to index procedures
(n = 2006) — worth doing for a baseline table, since the repeat rows are
re-interventions and land in the "Other" surgery group.

## Not available in this export

* **Total hospital length of stay** — `Ontslag opname` gives the hospital
  discharge but there is still no hospital *admission* date. What is reported is
  the **post-operative** hospital stay (surgery → discharge, median 7.9 d) and
  the ward stay after ICU discharge (median 4.9 d) — the post-operative stay is
  the usual ERAS endpoint anyway.
* **Urine output** — AKI is staged on creatinine and dialysis, so oliguric AKI
  without a creatinine rise is still missed and the rates remain a lower bound.
  Renal replacement therapy *is* now captured (`Dialyse_postop`).
* **Extubation timestamp** — reconstructed from the ventilation duration plus
  the ICU-admission offset (see below).

---

## Definitions and judgement calls

Everything below is a switch in the CONFIGURATION block at the top of the script.

### Elective vs non-elective — please confirm

`urgentie_graad` is **blank in 1255/2035 (61.7 %)**. The outcome gradient shows
what the codes mean:

| grade | n | ICU death | 180-day death | mech. support | median ICU LOS |
|---|---|---|---|---|---|
| *(blank)* | 1255 | 0.8 % | 2.5 % | 0.2 % | 2.98 d |
| E | 47 | 0.0 % | 6.4 % | 0.0 % | 3.13 d |
| S3 | 549 | 2.9 % | 6.4 % | 0.9 % | 3.07 d |
| S2 | 92 | 2.2 % | 4.3 % | 2.2 % | 3.90 d |
| S1 | 90 | **14.4 %** | **20.0 %** | 4.4 % | 3.90 d |

Blank and `E` behave like the lowest-risk cases and `S1` like the sickest, i.e.
urgency looks like it is **only coded when the case is not elective**. The
default `ELECTIVE_DEFINITION <- "blank_is_elective"` therefore reads blank + E
as elective and S1/S2/S3 as non-elective → **64.0 % elective**.

This is inference from outcomes, not documentation. **You know how the field is
filled in at AZ Delta — please confirm.** The alternative,
`"blank_is_missing"`, gives 47 elective / 731 non-elective / 1255 missing, which
would make the row unreportable.

### Type of surgery

Classified from the free-text `verrichting` field (212 distinct strings, several
procedures joined by `;`) by keyword:

* **CABG** — any string containing `CAB` (CABG, OPCAB, MIDCAB, Robot
  geassisteerde Midcab). No other procedure name in the file contains "cab", so
  the plain substring match is safe — `Carotis` does not match.
* **Valve** — surgical valve procedures only: `Aortaklep`, `Mitralisklep`,
  `Tricuspid/Tricuspied`, `AvR/MvR/TvR/AvP/MvP/TvP`, `Bentall` (composite valve
  conduit), `Ross`. `klepinterventie` is deliberately **not** a valve keyword,
  because in this file it only ever accompanies TAVI.

Two judgement calls, both switchable, both defaulting to *not* valve because
this is a cardiac **surgery** cohort:

| switch | default | effect |
|---|---|---|
| `EXCLUDE_TAVI` | `TRUE` | the 31 TAVI records are dropped from the analysis entirely |
| `DAVID_AS_VALVE` | `FALSE` | David / Tirone-David is valve-**sparing** root replacement → counted as aortic surgery |

With TAVI excluded the four groups are **1361 / 185 / 400 / 89** (CABG only /
CABG + valve / valve only / other). `TAVI_AS_VALVE` only matters if you set
`EXCLUDE_TAVI <- FALSE`.

**`output/verrichting_classification_AUDIT.csv` lists all 212 strings with the
group each was assigned to — please read it.** The classification is validated
by the fact that bypass and cross-clamp times track the groups exactly as
expected — median bypass / cross-clamp time is 90 / 66 min for CABG only,
107 / 80 for valve only and 124 / 90 for CABG + valve.

### Mortality

`mortaliteit` codes **one mutually exclusive** outcome per patient; blank =
alive. Four levels now: `overleden_IZ`, `overleden_hospitaal` (died in hospital
after ICU discharge), `30D mortaliteit` (died ≤ 30 d after hospital discharge)
and `180D mortaliteit`. They are reported both as coded and cumulatively:

| | n | % |
|---|---|---|
| ICU mortality | 41 | 2.0 % |
| **In-hospital mortality** | 63 | 3.1 % |
| 30-day mortality | 69 | 3.4 % |
| 180-day mortality | 91 | 4.5 % |

In-hospital mortality has an **independent cross-check**: `overleden_IZ` +
`overleden_hospitaal` = 63, and `Ontslag bestemming` = "8 Overleden" = 63. The
QC table verifies the two agree record by record; they do.

Caveat the script flags: 3 in-hospital deaths followed a stay longer than 30
days, so the cumulative 30-day figure is at most 3 deaths too high.

### AKI

The reference creatinine is **`laagste_creatinine_3m_voor_operatie`** — the
lowest value in the three months before surgery, which is the KDIGO baseline.
Creatinine is **mg/dL** here (reference median 0.93), auto-detected, so the
absolute-rise threshold is 0.3 mg/dL:

* **any AKI** — rise ≥ 0.3 mg/dL **or** ratio ≥ 1.5
* stage 1 — rise ≥ 0.3 mg/dL **or** ratio 1.5–1.9 (and ratio < 2)
* stage 2 — ratio 2.0–2.9
* stage 3 — ratio ≥ 3.0 **or post-operative dialysis**

`Dialyse_postop` is now recorded, so the KDIGO renal-replacement criterion is
applied: 56 patients had dialysis and **29 of them are stage 3 only because of
it** (their creatinine never reached 3× baseline). Set
`AKI_DIALYSIS_IS_3 <- FALSE` to score on creatinine alone.

Any AKI is **31.2 %** (610/1956 assessable). The reference is missing in 81
records (4.0 %) — the day-before value was missing in 196 (9.6 %), so the new
column is also the more complete one. Two sensitivity analyses are printed:

| definition | any AKI |
|---|---|
| lowest creatinine in 3 months (default) | 31.2 % |
| KDIGO time windows (rise ≤ 2 d, ratio ≤ 7 d) | 30.0 % |
| creatinine on the day before surgery | 26.9 % |

The day-before baseline misses roughly a fifth of the AKI the 3-month baseline
finds, which is the expected direction: a patient admitted with an acute rise
already under way has an elevated "day before" value, and measuring from it
hides the injury. `AKI_BASELINE <- "day_before"` switches back.

A sanity check in the QC table: the 3-month value is ≤ the day-before value in
1839 of 1839 records where both exist.

### Invasive ventilation

`duur_invbead_minuten` measures ventilation **from ICU admission** — patients
extubated in theatre have no value at all, which is how we know it excludes
theatre time. So:

* patients extubated on table contribute **0 minutes**, not a missing value;
* "time from end of surgery to extubation" adds the ICU-admission offset back
  on (median +5 min) and is clamped at 0;
* `> 6 h` is `VENT_THRESHOLD_MIN <- 360`, reported for the whole cohort, for
  the ICU-ventilated subgroup, and without the offset correction.

### Pain / PCIA

`PCIA` is a presence-only column (**312** `ja`, the rest blank), so "No PCIA"
means "not recorded as having had PCIA". 302 PCIA patients have a VAS score,
against 1639 without — a usable comparison, where the previous export had 14.

Mean VAS over 24 h is **1.45 [0.91–2.08] with PCIA vs 1.58 [1.00–2.20]
without**; Hodges–Lehmann shift −0.10 (95 % CI −0.21 to 0.00), p = 0.051. Note
this is the *opposite* direction to the previous export, where 14 patients gave
a large apparent increase — that was an artefact of the small, selectively
recorded group.

It still is not an effect of PCIA. PCIA is given to patients expected to have
more pain, so the comparison mixes any effect of the technique with confounding
by indication, in whichever direction it points. A difference of 0.1 VAS points
is also well below anything a patient would notice.

### Mean (SD) vs median [IQR]

Chosen per variable: mean (SD) if Shapiro–Wilk p > 0.05 **and** |skew| < 1,
otherwise median [IQR]. The chosen method is printed in every table, and
`output/Table_S1_numeric.csv` gives both summaries plus skewness and Shapiro p
for every numeric variable. In this cohort every variable is skewed enough that
median [IQR] is reported throughout.

Two places force one method across a row so that sub-groups stay comparable
(a group of 14 would otherwise pass Shapiro–Wilk and be reported as a mean):
the PCIA sub-groups, and every column of a stratified table.

Override with e.g. `FORCE_SUMMARY <- c(age_years = "mean_sd")`.

---

## Output

| file | contents |
|---|---|
| `Table_1_baseline.csv` | baseline characteristics |
| `Table_1b_supporting.csv` | raw urgency codes, composition of "Other", raw smoking categories, approach, CPB/cross-clamp |
| `Table_2_outcomes.csv` | mortality, AKI, ventilation, ICU readmission, pain, LOS, mobilisation |
| `VAS_by_PCIA.csv` | PCIA vs no PCIA for each VAS endpoint |
| `Table_S1_numeric.csv` | mean (SD) **and** median [IQR], skewness, Shapiro p for every numeric variable |
| `verrichting_classification_AUDIT.csv` | all 213 procedure strings → assigned group |
| `data_quality.txt` | consistency checks and missing-data counts |
| `Cardiochirurgie_results.xlsx` | all of the above as one workbook |
| `analysis_dataset.rds` | derived variables, for modelling — **contains identifiers** |

`STRATIFY_BY <- "surg_group"` adds a stratified Table 1 with p-values
(Kruskal–Wallis or ANOVA; χ², Fisher when expected counts < 5).

---

## `Slides_overview.R` — two-slide PowerPoint

```r
source("Cardiochirurgie.R")    # once, to build the derived dataset
source("Slides_overview.R")    # -> output/Cardiochirurgie_overview.pptx
```

Needs only `officer`. It re-uses `d` if it is already in the session, otherwise
reads `output/analysis_dataset.rds`, otherwise runs the analysis script itself.

* **Slide 1 — the cohort.** Hero band with the headline count, a five-figure KPI
  strip, type of surgery, age distribution, comorbidity and a baseline table.
* **Slide 2 — early recovery.** The six-step care pathway from theatre to
  hospital discharge, four headline indicators, the cumulative share extubated by
  time since end of surgery, mortality (ICU / in-hospital / 30-day / 180-day)
  and AKI stage with dialysis.

Every number is read from the data — nothing is typed in, so the deck is correct
by construction and re-running it after a data update refreshes both slides.

**Everything is a native PowerPoint shape**, not a picture: each rectangle, line
and word stays selectable and editable, the file is ~31 kB, and it stays crisp
at any zoom or print size. The trade-off is that the drawing is done by hand in
DrawingML rather than by a charting package, so a new chart type means a new
helper in section 2 (`el_hbars`, `el_vbars`, `el_table`, `el_step` are the ones
that exist).

Two implementation notes worth knowing if you edit it:

* officer 0.6.x cannot set the slide size and its template is 4:3, so the script
  rewrites `<p:sldSz>` inside the package to get 16:9 and then **reads the size
  back** and lays out from that — if the resize is ever refused the slides still
  fit the page instead of running off it. Repacking must use `zip::zipr()`;
  `zip::zip(mode = "cherry-pick")` flattens the directory structure and produces
  a file PowerPoint cannot open.
* The source is pure ASCII, with accented Dutch written as `\uXXXX` escapes, and
  the generated XML escapes every non-ASCII character as a numeric reference.
  That is deliberate: RStudio on Windows often runs in a non-UTF-8 locale, and
  without it "geëxtubeerd" and "patiënten" corrupt the file.

Change the look in the CONFIG block: `FONT` (default "Segoe UI") and the
palette constants. Slide text is Dutch; it is all in the `slide1` / `slide2`
blocks in sections 5 and 6.

The example report also carried a benchmark column (AZ Delta vs reference
ranges). That is deliberately **not** included — the reference ranges would have
to be quoted from a source rather than invented. Tell me which guideline values
you want to compare against and it is a short addition.

---

## Changes for the 04/09/2026 export

| | |
|---|---|
| Records | 1892 → 2066; period now starts 06/2020 |
| TAVI | 31 records **excluded** from every figure (`EXCLUDE_TAVI`) |
| AKI baseline | `laagste_creatinine_3m_voor_operatie` instead of the day-before value; any AKI 26.0 % → **31.2 %** |
| Dialysis | `Dialyse_postop` now recorded → KDIGO stage 3 by renal replacement (29 extra stage-3 patients) |
| Mortality | `overleden_hospitaal` added → **in-hospital mortality 3.1 %**, cross-checked against `Ontslag bestemming` |
| Hospital stay | `Ontslag opname` added → post-operative hospital stay, median **7.9 d** |
| PCIA | 16 → 312 recorded; the VAS difference reverses sign and is no longer significant |
| Column rename | `Patiëntennummer` → `Patientennummer`; the regex resolver matched both without a change |

To change the procedure classification, edit `RX_CABG` / `RX_VALVE` in
section 3.5. To add a variable to Table 1, add one `sp(...)` line to
`TABLE1_SPEC`.
