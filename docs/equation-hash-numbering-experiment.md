# Equation `#` numbering experiment

Branch: `experiment/equation-hash-numbering`

Goal: test whether Word's equation-internal `formula#number` mechanism can
replace the paragraph tab-stop layout, while keeping EqNB fields and cross
references.

## Tutorial Route

The pasted tutorial says the important part is that `#(...)` and the `SEQ` field
are both inside the Word equation object:

- Plain number: `formula#({ SEQ Equation })`
- Chapter number: `formula#({ SEQ Chapter \c }-{ SEQ Equation })`
- Chapter start marker: hidden `{ SEQ Chapter \h }` plus hidden
  `{ SEQ Equation \r \h }`
- Cross reference: bookmark the number text inside the parentheses, then insert
  a `REF` field to that bookmark.

## Findings

- In manual Word editing, `formula#(1)` can push `(1)` to the right side of an
  equation line.
- In VBA automation, inserting placeholder text plus `#(1)` into an equation
  object leaves `#` as literal text in the rendered PDF.
- Creating an `OMath` from a range containing placeholder text plus `#(SEQ
  field)` also leaves `#` literal in automation.
- Calling `OMath.BuildUp` from automation is not reliable in this environment.
- Driving the interactive equation editor with `EquationInsertNew`, `Alt+=`, or
  `SendKeys` is not stable enough for the default EqNB workflow.

## Branch Implementation

This branch keeps the stable paragraph-tab implementation and adds experimental
macros:

- `InsertHashEquationLinePlain`
- `InsertHashEquationLineChapterHyphen`
- `MarkEquationChapterStart`

The experimental macros insert `placeholder#(...)`, add `SEQ` fields inside the
parentheses, bookmark the visible number text, and convert the full range to a
Word equation object. If the user's interactive Word session honors the equation
editor's `#` parser after pressing Enter inside the equation, the result should
match the tutorial while keeping EqNB cross references.

## Current Risk

The `#` mechanism appears to depend on Word's interactive equation editor path,
not the regular VBA text insertion path. The experimental buttons are isolated
from the stable buttons until a reliable parser trigger is confirmed.
