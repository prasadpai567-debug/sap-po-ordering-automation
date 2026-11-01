# SAP Purchase Order Automation with Excel VBA

Excel VBA automation that uses SAP GUI Scripting to process material and MRP-area inputs, work through SAP MD04, convert eligible planned orders into purchase requisitions, create purchase orders through ME21N, and write the resulting purchase-order number or an error message back to Excel.

> **Important:** This project automates transactions in SAP. Use it only with authorized accounts, approved business processes, and a non-production test system until the result is fully validated.

This public repository intentionally excludes the original operational `.xlsm` workbook. The published VBA export uses placeholders for company-specific recipients and communication values; restore approved internal values only in an authorized private copy.

## Project overview

The application provides a spreadsheet-driven interface for a procurement process that would otherwise require repeated manual work in SAP:

1. Enter material and MRP-area information in Excel.
2. Start the macro from the worksheet button.
3. Let the automation navigate SAP MD04 for each row.
4. Locate a planned-order entry and convert it into a purchase requisition.
5. Continue into ME21N, populate configured purchasing information, and save the purchase order.
6. Capture the resulting document number in Excel and highlight rows that fail.

The main business problem addressed is repetitive, error-prone navigation between Excel and SAP during purchase-order processing. The workbook centralizes the input list and provides a simple row-level result area.

## Implemented features

- **Spreadsheet-driven processing:** Reads material and MRP-area values row by row from the active worksheet.
- **One-click launch:** A worksheet shape labeled `MD04` is assigned to the `SAP_Final_Solution` macro.
- **SAP GUI session attachment:** Connects to the first available SAP GUI connection and session through late-bound SAP GUI Scripting objects.
- **MD04 navigation:** Opens MD04, enters the material and MRP-area values, and submits the selection.
- **Dynamic MD04 handling:** Searches for the MD04 table container and detail icon instead of relying on only one fixed table position.
- **Planned-order processing:** Searches the MD04 list for the planned-order entry, opens its detail action, and converts it to a purchase requisition.
- **Warning handling:** Attempts to accept SAP warning or information dialogs during conversion, navigation, and saving.
- **ME21N automation:** Navigates to ME21N and enters the recorded purchasing-screen values across relevant header, item, delivery-schedule, confirmation, text, and communication areas.
- **Purchase-order result capture:** Reads the SAP status bar and extracts the first 10-digit number as the purchase-order result.
- **Row-level error marking:** Writes an error message in the output column and highlights the affected row in yellow when processing fails.
- **Completion notification:** Displays a completion message after all input rows have been attempted.

## Application workflow

### 1. Prepare the input sheet

The workbook contains one visible worksheet. The working columns are:

| Column | Purpose |
|---|---|
| A | Material identifier supplied to SAP |
| B | MRP area supplied to SAP |
| C | Purchase-order result or error message |

The current workbook contains sample operational data. Replace it with approved test data before running the macro.

### 2. Start the required systems

The user must already be logged in to SAP GUI for Windows. SAP GUI Scripting must be available, and the SAP user must have authorization for the relevant MD04, purchase-requisition, and ME21N activities.

### 3. Run the macro

Click the worksheet button labeled `MD04`, or run `SAP_Final_Solution` from Excel's Developer > Macros menu. The macro uses the active worksheet as its input sheet.

### 4. Process each row

For every non-empty material row, the macro:

1. Opens or resets MD04.
2. Enters the material and MRP-area values.
3. Applies the configured MD04 filter behavior.
4. Finds the planned-order row, including rows that require table scrolling.
5. Locates the detail icon across a range of possible columns.
6. Converts the planned order into a purchase requisition.
7. Refreshes MD04 and searches for the resulting purchase-requisition entry.
8. Navigates to ME21N.
9. Handles document-overview and warning dialogs.
10. Enters the configured ME21N values.
11. Saves the document and handles final save warnings.
12. Reads the status bar and writes the result to column C.

### 5. Handle failures

If a row cannot be completed, the macro writes an `ERROR:` message in column C and colors that row yellow. Processing then continues with the next row. The final completion message means that processing attempts have finished; it does not guarantee that every row succeeded.

## Technical details

### Technology stack

- Microsoft Excel desktop workbook in `.xlsm` format
- VBA, with `Option Explicit`
- SAP GUI for Windows Scripting API through late-bound `Object` variables
- SAP transactions used by the implementation: MD04 and ME21N
- Built-in Windows scripting objects:
  - `Scripting.Dictionary` for SAP GUI path caching
  - `VBScript.RegExp` for extracting a 10-digit document number
- Excel worksheet formatting, shapes, message boxes, and cell/row updates

### VBA project structure

The reviewed project contains:

- `Sheet1` — the input/output worksheet
- `ThisWorkbook` — workbook object
- `Module1` — the standard VBA module containing the main macro and helper routines
- No UserForms were present in the reviewed VBA project

The main routines are:

- `SAP_Final_Solution` — controls the end-to-end row loop and SAP workflow
- `GetMD04Sub` — discovers the active MD04 table container across likely SAP dynpro variants
- `ExecuteSAP` — performs common SAP GUI actions such as selecting, typing, focusing, pressing, and setting selection indexes
- `SafeGet` — resolves SAP GUI paths across common screen/container variations and caches successful paths
- `ExtractPONumber` — extracts the first 10-digit number from a SAP status message

### Workbook characteristics

- One visible worksheet
- Small row-oriented input/output layout
- Worksheet shape used as the macro launcher
- Conditional formatting for duplicate material values
- No Excel tables, charts, PivotTables, formulas, or external workbook links were identified in the reviewed workbook
- A hidden filter-database defined name and template-related custom XML are present but are not part of the business workflow

## Requirements

- Windows desktop environment
- Microsoft Excel desktop with permission to run macros
- SAP GUI for Windows
- SAP GUI Scripting enabled and permitted by the SAP environment
- An active SAP connection and logged-in session
- SAP authorization for the required transactions and document creation steps
- A stable SAP GUI layout that matches the automation's recorded controls

Excel Online, LibreOffice, and other spreadsheet viewers are not supported for running this automation.

## Installation and setup

### Recommended public-repository setup

Do not publish the attached workbook unchanged. First create a sanitized copy that contains only dummy data and approved automation code.

1. Create a new repository using the recommended structure below.
2. Add this `README.md`.
3. Add a sanitized `.xlsm` workbook only if publishing the compiled VBA project is approved.
4. Optionally export a reviewed, sanitized `Module1.bas` file for source visibility.
5. Replace all operational sample values with dummy values.
6. Remove workbook metadata, internal URLs, company identifiers, and screenshots containing live SAP information.
7. Confirm that the workbook opens in Excel and that the macro project is still intact.

### Local or internal setup

1. Open the `.xlsm` file in Microsoft Excel desktop.
2. Review Excel's macro-security prompt and enable content only when the file is trusted and approved.
3. Start SAP GUI for Windows and log in to the correct test system, client, and session.
4. Confirm that SAP GUI Scripting is enabled and that the user has the required authorizations.
5. Enter test material identifiers in column A and MRP areas in column B. Leave column C available for results.
6. Keep the input worksheet active.
7. Click the `MD04` worksheet button or run `SAP_Final_Solution`.
8. Review the result or error text in column C and inspect yellow rows.
9. Verify created documents directly in the approved SAP test environment.

Run one or two controlled test rows before processing a larger list. Avoid duplicate rows and do not run the macro concurrently in multiple Excel or SAP sessions unless the process owner has explicitly approved that usage.

## Example input

Use dummy data in public documentation and testing:

| Material | MRP Area | PO Number |
|---|---|---|
| `DEMO-MAT-001` | `TEST-MRP` | *(blank before run)* |
| `DEMO-MAT-002` | `TEST-MRP` | *(blank before run)* |

The actual application writes either a document number or an error/status message into the result column.

## Data protection and publication checklist

Before uploading the workbook or screenshots to GitHub, remove or replace:

- Material identifiers, MRP-area values, plant codes, purchase-order numbers, requisition numbers, and other operational identifiers
- Supplier, vendor, customer, employee, or user details
- SAP system names, client numbers, connection details, hostnames, and environment identifiers
- Company-specific process codes and hard-coded communication values
- SharePoint, OneDrive, or other internal absolute paths
- Personal names, email addresses, workbook author metadata, and revision metadata
- Screenshots showing live transactions, vendors, prices, quantities, account assignments, or user names
- Any credentials, tokens, passwords, or recorded SAP automation outside the reviewed module

The reviewed `Module1` code did not show a hard-coded SAP password or credential. This is not a substitute for a complete security review of the workbook, embedded VBA project, screenshots, and any future modules.

An `.xlsm` file contains compiled VBA and may expose source logic. Publish it only after confirming that disclosure is allowed by the employer or process owner. If the code is proprietary, publish a redacted demonstration workbook or screenshots instead.

## Limitations and operational risks

- The automation is attended, desktop-based SAP GUI automation; it is not a server-side integration.
- SAP GUI element paths and screen layouts can change after SAP upgrades, personalization, or transaction customization.
- The macro attaches to the first available SAP connection and first session, which may be unsafe when multiple sessions are open.
- The workflow includes extensive `On Error Resume Next` handling. This helps it pass transient warnings but can also hide the original cause of a failure.
- The result parser assumes that the first 10-digit number in the SAP status bar is the purchase-order number. If no such number is present, the full status text is written to the result cell.
- Several ME21N values and SAP GUI paths are embedded in the VBA implementation rather than stored in a configuration sheet.
- The macro does not provide a formal audit log, transaction-level rollback, duplicate prevention, production-environment guard, or structured retry history.
- The `Processing Complete` message means the row loop has ended, not that all purchase orders were successfully created.

## Implemented behavior vs. suggested improvements

### Implemented in the reviewed workbook

- Row-based Excel input
- One-click worksheet launch
- SAP GUI attachment
- MD04 lookup and dynamic table/icon discovery
- Planned-order conversion to purchase requisition
- ME21N navigation and save flow
- Warning-dialog handling and retries
- PO/status extraction
- Yellow row-level error marking

### Suggested future improvements

These are recommendations, not features confirmed in the current workbook:

- Add a configuration sheet for SAP environment, plant, purchasing values, communication values, and other process parameters.
- Add input validation for blank values, duplicate materials, invalid MRP areas, and already-processed rows.
- Let the user choose the SAP connection and session explicitly.
- Add a production-environment warning and a sandbox/test-system guard.
- Replace broad `On Error Resume Next` blocks with structured error handling and a detailed log sheet.
- Store timestamps, source row, SAP message, document type, and final document number in an audit log.
- Validate the extracted document number against the expected SAP message pattern and document type.
- Add idempotency checks so a rerun cannot unintentionally create duplicate documents.
- Separate SAP interaction, configuration, validation, and reporting into smaller testable modules.
- Add a controlled UserForm only if a richer user interface is genuinely needed.
- Create a mock or recorded test harness for development so code changes can be tested without creating live SAP documents.

## Recommended repository structure

```text
sap-po-ordering-automation/
├── README.md
├── LICENSE
├── .gitignore
├── docs/
│   ├── VBA_APPLICATION_ANALYSIS.md
│   └── screenshots/
│       ├── 01-input-sheet.png
│       ├── 02-md04-processing.png
│       ├── 03-me21n-entry.png
│       └── 04-completed-results.png
├── src/
│   └── Module1.bas
├── workbook/
│   └── SAP_PO_Order_Automation_Sanitized.xlsm
└── samples/
    └── sample_input.csv
```

### Files suitable for GitHub

- `README.md`
- A sanitized analysis document
- A reviewed, sanitized VBA export such as `src/Module1.bas`
- A sanitized demonstration workbook, if permitted
- Dummy input data and screenshots using test values
- A license selected by the owner

### Files not suitable for GitHub

- The original operational workbook
- Any workbook containing real material, MRP-area, PO, vendor, or SAP system data
- Unredacted screenshots
- Internal paths, credentials, exported SAP session recordings, or company-only documents
- Temporary analysis copies or extracted binary artifacts

## Suggested screenshots

Use screenshots from a test or sandbox environment with dummy values:

1. **Input sheet:** Material and MRP-area columns, blank output column, and the `MD04` button.
2. **MD04 processing:** The planned-order row and the detail action, with identifiers and system information redacted.
3. **ME21N processing:** The relevant header, delivery, confirmation, text, and communication areas using safe test values.
4. **Successful output:** The Excel result column populated with a dummy purchase-order number.
5. **Error handling:** A failed row highlighted in yellow with a non-sensitive error message.
6. **Technical view:** The VBA Project Explorer and sanitized `Module1` code, if source disclosure is approved.

Add a caption such as “Illustrative screenshot from a test environment — no production data.”

## Resume and portfolio description

### Resume version

Built an Excel VBA and SAP GUI Scripting automation tool that reads material and MRP-area inputs, navigates SAP MD04, converts planned orders into purchase requisitions, creates purchase orders in ME21N, captures document results back into Excel, and flags row-level errors. Demonstrated Excel automation, SAP GUI integration, dynamic UI handling, exception handling, and procurement-process improvement.

### GitHub or LinkedIn version

Excel VBA and SAP GUI Scripting automation for row-driven purchase-order processing, SAP transaction navigation, document-number capture, and exception reporting.

## Validation note

This documentation is based on a static review of the workbook structure and VBA source. The macro was not executed against SAP during the review. Validate every step in an authorized SAP test environment before operational use.

