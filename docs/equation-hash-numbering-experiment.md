# Equation `#` numbering experiment

Branch: `experiment/equation-hash-numbering`

Goal: test whether Word's equation-internal `formula#number` mechanism can replace
the current paragraph tab-stop layout, while keeping EqNB fields and cross references.

## Findings

- In manual Word editing, `formula#(1)` can push `(1)` to the right side of an
  equation line.
- In VBA automation, inserting `□#(1)` into an equation object leaves `#` as
  literal text.
- Creating an OMath from a range containing `□#(SEQ field)` also leaves `#` as
  literal text.
- Calling `OMath.BuildUp` on that range breaks the field/linear text structure
  into multiple paragraph-like runs and is not usable for EqNB.
- `Application.CommandBars.ExecuteMso "EquationProfessional"` was not callable
  in the tested automation context.

## Conclusion

The `#` mechanism appears to depend on Word's interactive equation editor path,
not the regular VBA text insertion path. For now, keep the stable paragraph
tab-stop implementation on `master`.

Do not merge a `#` implementation unless a reliable VBA trigger for Word's
interactive equation parser is found.
