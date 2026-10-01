# LS + MS dataset merge

Harmonises the LS treatment/class codes with the MS codes so the two sheets in
`merge_dataset.xlsx` can be merged without codes getting mixed up.

`input/` holds the original files from `fredrikanesten-arch/dataset-merge`.
`trt_to_class_ms.csv` is the base, and none of its codes are changed.

Run `python3 harmonise_codes.py` (needs `lxml`) to regenerate `output/`:

| File | Content |
|------|---------|
| `trt_to_class_merge.csv` | MS codes 1–99 / classes 1–50 unchanged. LS-only treatments are appended as 100–137 and LS-only classes as 51–59. LS treatments that also exist in MS use the MS codes. |
| `merge_dataset.xlsx` | `t[,1]`–`t[,5]` in all three data blocks of **LS SMD bias-adj** are recoded to the merged codes, and that sheet's *TREATMENT AND CLASS CODES* table is replaced with the merged table. **MS SMD bias-adj** is left unchanged. |
| `ls_to_merge_code_map.csv` | Old LS code → merged code for every LS treatment and class (`matched` = existing MS code, `added` = new code). |

Treatments and classes are matched by name, ignoring case and extra whitespace.

Note: some blocks have different columns in the two sheets. The LS CFB block has
no `CCFB[,4]`, and the LS dichotomous block has only `C[,2]`–`C[,3]`, while MS
has `C[,2]`–`C[,5]`. Line up these columns by header, not by position, when you
stack the two sheets.
