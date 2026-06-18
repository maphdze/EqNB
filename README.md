# EqNB

EqNB is a Word `.dotm` add-in for inserting equation numbers and cross-references.

## Features

- Insert centered display equations with right-aligned sequential numbers.
- Insert inline equations without numbering or display style.
- Use Word fields for equation numbers and references: `SEQ Equation` and `REF bookmark \h`.
- Store a document-wide reference format such as `({n})`, `式({n})`, `Eq.({n})`, or `[{n}]`.
- Provide an `EqNB` ribbon tab whose labels and tips follow the Word UI language: English or Simplified Chinese.

## Build

Enable Word's VBA project object model access, then run:

```powershell
cd C:\Software\AI\Word公式编辑器
powershell -ExecutionPolicy Bypass -File tools\Build-Dotm.ps1
```

Output:

```text
release\EqNB.dotm
```

## Install

Copy `release\EqNB.dotm` to:

```text
%APPDATA%\Microsoft\Word\STARTUP
```

Restart Word. The `EqNB` tab should appear in the ribbon.

More details are in `docs\dotm-install.md`.
