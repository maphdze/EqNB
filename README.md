# EqNB

EqNB is a Microsoft Word add-in for inserting numbered display equations and
cross-references. It is implemented as a macro-enabled Word template (`.dotm`).

EqNB is currently intended for Microsoft Word on Windows. The macOS version of
Word has not been fully supported or tested, and some automation behavior is
known to be unreliable there.

The current implementation uses Word's equation-internal hash numbering syntax
(`equation#(number)`) so that the equation stays centered while the number is
aligned to the right.

## Features

- Insert display equations with right-aligned numbers.
- Insert inline equations without numbering.
- Insert chapter-based equation numbers, such as `(1.1)` or `(1-1)`.
- Add hidden chapter-start markers for chapter numbering.
- Insert refreshable cross-references to equation numbers.
- Configure document-level number formats and reference formats.
- Use Times New Roman upright text for equation numbers while leaving the
  equation font unchanged.
- Install as a Word Startup add-in with a double-click script.

## Quick Install

Download and unzip:

```text
release/EqNB-Installer.zip
```

Close all Word windows, then double-click:

```text
双击安装EqNB.cmd
```

Restart Word. The `EqNB` tab should appear in the ribbon.

To uninstall, close Word and double-click:

```text
双击卸载EqNB.cmd
```

## Manual Install

The installer copies:

```text
release/EqNB.dotm
```

to Word's Startup folder:

```text
%APPDATA%\Microsoft\Word\STARTUP
```

You can also run the PowerShell installer directly:

```powershell
powershell -ExecutionPolicy Bypass -File tools\Install-EqNB.ps1
```

Uninstall:

```powershell
powershell -ExecutionPolicy Bypass -File tools\Uninstall-EqNB.ps1
```

## Build From Source

Building the `.dotm` requires Microsoft Word and access to the VBA project
object model.

In Word, enable:

```text
File > Options > Trust Center > Trust Center Settings > Macro Settings >
Trust access to the VBA project object model
```

Then run from the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File tools\Build-Dotm.ps1
```

The output is:

```text
release/EqNB.dotm
```

## Usage

After installation, open the `EqNB` ribbon tab in Word.

- `#编号公式`: insert a numbered display equation.
- `行内公式`: insert an inline equation.
- `#章节编号`: insert a chapter-based numbered equation.
- `章节起点`: mark the start of a chapter for equation numbering.
- `编号格式`: choose common number formats such as `(1)`, `(1.1)`, or `(1-1)`.
- `引用格式`: set the reference text format, such as `({n})` or `Eq.({n})`.
- `插入引用`: insert a refreshable equation reference.
- `刷新`: refresh equation numbers and references.

## Notes

- Currently supported target: Microsoft Word for Windows.
- Microsoft Word for Mac is not recommended at this stage.
- Close Word before installing or uninstalling.
- If the add-in does not appear, check Word macro/security settings and confirm
  that `EqNB.dotm` is in Word's Startup folder.
- The add-in stores number and reference format settings in the current Word
  document.
