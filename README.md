# Cardiochirurgie AZ Delta — descriptive analysis

`Cardiochirurgie.R` produces the baseline table and the outcome tables for the
cardiac surgery ICU cohort (`iz_cardio_versie_28082026.xlsx`, 1892 records,
surgery 2020-12-29 → 2026-08-19).

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
| Records / patients | 1892 records in 1869 patients — 23 patients have a second record (re-operation: `Revisie`, `Thoraxdrain`, `Pericardfenestratie`) |
| `Opname/Ontslag Datetime` | **ICU** admission/discharge. `Opname` falls a median of 7 min after the end of surgery (IQR −6 to +18), and `Ligduur` = `Ontslag − Opname` exactly |
| `Ligduur` | **ICU** length of stay in days |

Set `ANALYSIS_UNIT <- "first_per_patient"` to restrict to index procedures
(n = 1869) — worth doing for a baseline table, since the 23 repeat rows are
re-interventions and land in the "Other" surgery group.

## Not available in this export

* **Hospital length of stay** — there is no hospital admission/discharge column.
  Only the ICU stay can be reported. The row exists in Table 2 marked
  "not available"; fill `d$hosp_los_days` if the columns are added.
* **Urine output and renal replacement therapy** — AKI is staged on creatinine
  alone, so the reported rates are a **lower bound** (oliguric AKI and stage 3
  by RRT are missed).
* **Extubation timestamp** — reconstructed from the ventilation duration plus
  the ICU-admission offset (see below).

---

## Definitions and judgement calls

Everything below is a switch in the CONFIGURATION block at the top of the script.

### Elective vs non-elective — please confirm

`urgentie_graad` is **blank in 1181/1892 (62.4 %)**. The outcome gradient shows
what the codes mean:

| grade | n | ICU death | any death | mech. support | median ICU LOS |
|---|---|---|---|---|---|
| *(blank)* | 1181 | 0.8 % | 2.3 % | 0.2 % | 2.97 d |
| E | 33 | 0.0 % | 6.1 % | 0.0 % | 3.28 d |
| S3 | 510 | 2.9 % | 6.1 % | 1.0 % | 3.07 d |
| S2 | 84 | 2.4 % | 4.8 % | 2.4 % | 3.97 d |
| S1 | 81 | **12.3 %** | **18.5 %** | 4.9 % | 4.02 d |

Blank and `E` behave like the lowest-risk cases and `S1` like the sickest, i.e.
urgency looks like it is **only coded when the case is not elective**. The
default `ELECTIVE_DEFINITION <- "blank_is_elective"` therefore reads blank + E
as elective and S1/S2/S3 as non-elective → **64.3 % elective**.

This is inference from outcomes, not documentation. **You know how the field is
filled in at AZ Delta — please confirm.** The alternative,
`"blank_is_missing"`, gives 33 elective / 675 non-elective / 1184 missing, which
would make the row unreportable.

### Type of surgery

Classified from the free-text `verrichting` field (213 distinct strings, several
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
| `TAVI_AS_VALVE` | `FALSE` | 31 TAVI records stay in "Other" rather than joining "Valve only" |
| `DAVID_AS_VALVE` | `FALSE` | David / Tirone-David is valve-**sparing** root replacement → counted as aortic surgery |

Setting both to `TRUE` gives 1242 / 174 / 423 / 53 instead of 1245 / 171 / 367 / 109.

**`output/verrichting_classification_AUDIT.csv` lists all 213 strings with the
group each was assigned to — please read it.** The classification is validated
by the fact that bypass and cross-clamp times track the groups exactly as
expected — median bypass / cross-clamp time is 90 / 66 min for CABG only,
107 / 81 for valve only and 120 / 88.5 for CABG + valve.

### Mortality

`mortaliteit` codes **one mutually exclusive** outcome per patient; blank =
alive. `30D mortaliteit` means died within 30 days **after ICU discharge**, so
the categories are reported both as coded and cumulatively.

Caveat the script flags: 2 ICU deaths occurred after an ICU stay longer than
30 days, so the cumulative 30-day figure is at most 2 deaths too high.

### AKI

Staged from `max_creatinine` vs `creatinine_dag_voor_operatie`, exactly as
requested. Creatinine is **mg/dL** here (median 0.97), auto-detected, so the
absolute-rise threshold is 0.3 mg/dL:

* stage 1 — rise ≥ 0.3 mg/dL **or** ratio 1.5–1.9 (and ratio < 2)
* stage 2 — ratio 2.0–2.9
* stage 3 — ratio ≥ 3.0

Baseline creatinine is missing in 188 records (9.9 %), which are not stageable.
Because `datum_max_creatinine` is a date without a time, a **day-resolution
KDIGO-windowed sensitivity analysis** is also reported (rise within 2 days,
ratio criteria within 7 days): it moves any-AKI from 26.0 % to 23.8 %.

### Invasive ventilation

`duur_invbead_minuten` measures ventilation **from ICU admission** — patients
extubated in theatre have no value at all, which is how we know it excludes
theatre time. So:

* patients extubated on table contribute **0 minutes**, not a missing value;
* "time from end of surgery to extubation" adds the ICU-admission offset back
  on (median +7 min) and is clamped at 0;
* `> 6 h` is `VENT_THRESHOLD_MIN <- 360`, reported for the whole cohort, for
  the ICU-ventilated subgroup, and without the offset correction.

### Pain / PCIA

`PCIA` is a presence-only column (16 `ja`, the rest blank), so "No PCIA" means
"not recorded as having had PCIA". Only **14** PCIA patients have a VAS score.
The comparison is reported with both Wilcoxon (Hodges–Lehmann shift + CI) and
Welch t, but PCIA is prescribed to patients expected to have more pain — a
higher VAS in the PCIA group is **confounding by indication**, not evidence
about the technique.

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

To change the procedure classification, edit `RX_CABG` / `RX_VALVE` in
section 3.5. To add a variable to Table 1, add one `sp(...)` line to
`TABLE1_SPEC`.
