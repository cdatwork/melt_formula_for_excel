# melt_formula_for_excel
I created a MELT formula (similar to the melt function in pandas) for Excel to turn a wide table into a tall table.

To use in Excel:
* Press CTRL+F3 or go to 'Formulas' > 'Name Manager'
* Under 'Name', put "MELT"
* Under 'Comment', put "Turns a wide table into a tall table. USAGE: MELT(array, {headers}, 3)"
* Under 'Refers to', paste the formula below

```
=LAMBDA(tbl, idIdxs, ignore,
  LET(
    hdrs, TAKE(tbl,1),
    data, DROP(tbl,1),
    allIdx, SEQUENCE(, COLUMNS(tbl)),
    valIdxs, FILTER(allIdx, ISNA(MATCH(allIdx, idIdxs, 0))),
    ids,  CHOOSECOLS(data, idIdxs),
    vals, CHOOSECOLS(data, valIdxs),
    vars, CHOOSECOLS(hdrs, valIdxs),
    r, ROWS(vals), c, COLUMNS(vals),
    idIndex,  MOD(SEQUENCE(r*c)-1, r)+1,
    varIndex, CEILING(SEQUENCE(r*c)/r),
    VSTACK(
      HSTACK(CHOOSECOLS(hdrs, idIdxs), {"Variable","Value"}),
      HSTACK(
        INDEX(ids, idIndex, SEQUENCE(, COLUMNS(ids))),
        INDEX(vars, 1, varIndex),
        TOCOL(vals, ignore)
      )
    )
  )
)
```
