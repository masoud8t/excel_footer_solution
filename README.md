# KISS Print — Excel Bottom-of-Page Repeater

A deliberately simple Excel/VBA utility for forms and reports where a large block of cells must be repeated at the bottom of every printed page.

The project follows one rule:

> **Keep It Simple, Stupid.**

Instead of trying to replace Excel's pagination engine, the macro creates a working copy of the worksheet, inserts the selected footer block after a configurable number of data rows, adds manual page breaks, and keeps the generated print sheet available for manual adjustment.

## Why this exists

Excel can repeat rows at the **top** of printed pages through **Page Layout → Print Titles**.

It does not provide an equally simple built-in mechanism for repeating a large, formatted block of cells at the **bottom** of every page.

This is especially useful for controlled forms where the bottom section may contain:

- approval fields
- technical committee results
- explanations
- signatures
- revision information
- document/form codes
- large merged-cell layouts
- Persian/RTL text
- borders and other print formatting

For these forms, an Excel Header/Footer is often too restrictive because the repeated block is effectively a substantial part of the form rather than a small page footer.

## Design principles

The implementation intentionally avoids complex pagination logic.

### Original worksheet is not modified

The source worksheet is copied to a new worksheet before the print layout is generated.

The original form remains unchanged.

### The generated print sheet is kept

The temporary print sheet is **not deleted** after preview.

This is intentional.

The generated sheet can be manually adjusted in Excel before printing:

- row heights
- page breaks
- margins
- scaling
- spacing
- footer placement
- column widths
- print area

Running the macro again creates another uniquely named print sheet, such as:

`__KISS_PRINT`

`__KISS_PRINT_2`

`__KISS_PRINT_3`

Existing print sheets are preserved.

### Pagination is deliberately manual

The macro asks:

> **After how many data rows should the selected footer be pasted?**

For example:

`24` → footer after every 24 data rows.

This avoids trying to predict Excel's physical page layout from row heights, fonts, printer drivers, paper sizes, scaling, and other variables.

The user tunes the row count once for the specific form/printer combination.

## Workflow

1. Open the original Excel form.
2. Select the complete block that should be repeated at the bottom of each page.
3. Run `KISS_Print_Form`.
4. Enter the number of data rows that should appear before the repeated block.
5. Excel creates a new `__KISS_PRINT...` worksheet.
6. Review the generated print layout.
7. Make any final manual adjustments.
8. Print the generated worksheet.

The original worksheet is left intact.

## Configuration

The macro currently has one fixed configuration value:

`DATA_FIRST_ROW`

Example:

```vba
Const DATA_FIRST_ROW As Long = 8
```

This is the first row containing the actual repeating data.

The number of rows per page is entered interactively when the macro runs:

```
After how many data rows should the selected footer be pasted?

[ 24 ]
```

This makes testing different layouts very fast.

## Example

Assume the form contains:

- rows 1–7: form/header information
- rows 8–103: variable data
- rows 104–120: controlled bottom section

Select rows 104:120 and run the macro.

With `24` rows per page, the intended layout is:

```
Data rows 8–31
Footer block

Data rows 32–55
Footer block

Data rows 56–79
Footer block

Data rows 80–103
Footer block
```

The exact number should be tuned against the actual printed form.

## Why not use the Excel Page Footer?

Excel's native page footer is appropriate for small elements such as:

- page numbers
- dates
- short document identifiers
- simple approval text

It is not a good fit for a large multi-row form section containing merged cells, borders, labels, signatures, and structured fields.

For those layouts, keeping the repeated block as normal worksheet cells provides much better control.

## Why not calculate page height automatically?

Because that quickly becomes complicated and fragile.

Actual Excel pagination can depend on:

- paper size
- printer driver
- margins
- zoom
- Fit to Page settings
- row heights
- font rendering
- merged cells
- hidden rows/columns
- manual page breaks
- installed fonts
- Excel version

KISS Print intentionally leaves that final physical adjustment to Excel and the user.

## Scope

This utility is intended for **print-oriented Excel forms and reports**.

It is not intended to be a general replacement for Excel's pagination engine.

The current approach works best when:

- the source data is one continuous block
- the repeated footer is a contiguous range of rows
- a fixed number of data rows is acceptable per printed page
- manual final adjustment is acceptable

## File structure

```
excel/
└── kiss-print/
    ├── KISS_Print_Form.bas
    └── README.md
```

## Installation

### Option 1 — Import the VBA module

1. Open the target workbook in Excel.
2. Press **Alt + F11**.
3. Right-click the workbook in the VBA Project window.
4. Choose **Import File...**
5. Select `KISS_Print_Form.bas`.
6. Save the workbook as `.xlsm`.

### Option 2 — Copy the module

1. Open the `.bas` file.
2. Copy the VBA code.
3. In Excel press **Alt + F11**.
4. Choose **Insert → Module**.
5. Paste the code.
6. Save as `.xlsm`.

## Usage notes

The repeated footer must be selected before running the macro.

For example:

```
104:120
```

The selected block is copied from the original worksheet, so formatting from the original form remains the source of truth.

The print sheet is intentionally treated as a **working artifact** rather than an invisible implementation detail.

That means it is normal to create several versions while tuning a form:

```
__KISS_PRINT
__KISS_PRINT_2
__KISS_PRINT_3
```

Once the correct layout is found, the preferred print sheet can simply be used as the final print layout.

## Safety

The macro is designed to avoid modifying the original form data.

It operates by:

1. reading the active worksheet
2. creating a worksheet copy
3. modifying only that copy
4. leaving the generated copy in the workbook

Nevertheless, test the macro on a copy of an important workbook before deploying it in a controlled production document.

## Current limitations

This implementation intentionally does not attempt to automatically reproduce every possible Excel pagination scenario.

Known boundaries include:

- the data section is assumed to be continuous
- the footer is selected as one contiguous row block
- the number of data rows per page is user-configured
- the user may need to tune the value for the target printer
- the generated sheet may require manual adjustment
- complex Excel objects such as charts/shapes are outside the core design
- multiple independent print areas are outside the intended use case

## Philosophy

This project is intentionally small.

A print utility should not become a second spreadsheet application.

When Excel can handle something reliably, let Excel handle it.

When VBA can solve the remaining problem with a few predictable operations, use VBA.

When the final physical adjustment depends on the actual printer and form, expose one simple setting and let the user tune it.

**KISS.**
