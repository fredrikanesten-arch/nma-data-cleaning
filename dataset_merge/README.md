# LS + MS dataset merge

Harmonises the LS treatment/class codes with the MS codes and combines the
data from both sheets into one worksheet.

`input/` holds the original files from `fredrikanesten-arch/dataset-merge`.
`trt_to_class_ms.csv` is the base, and none of its codes are changed.

Run `python3 harmonise_codes.py` (needs `lxml`) to regenerate `output/`:

| File | Content |
|------|---------|
| `trt_to_class_merge.csv` | MS codes 1–99 / classes 1–50 unchanged. LS-only treatments are appended as 100–137 and LS-only classes as 51–59. LS treatments that also exist in MS use the MS codes. |
| `merge_dataset.xlsx` | One **Merged dataset** sheet with three outcome blocks (change from baseline, baseline/follow-up, dichotomous). Within each block, MS rows precede LS rows. LS `t[,1]`–`t[,5]` use merged codes; MS codes remain unchanged. The merged *TREATMENT AND CLASS CODES* table is alongside the data in columns AM–AP. |
| `ls_to_merge_code_map.csv` | Old LS code → merged code for every LS treatment and class (`matched` = existing MS code, `added` = new code). |

Treatments and classes are matched by name, ignoring case and extra whitespace.

Columns are aligned by header, not position. Missing LS columns such as
`CCFB[,4]` and `C[,4]`–`C[,5]` are filled with `NA`, keeping `#` and `studyid`
under their correct headers.
