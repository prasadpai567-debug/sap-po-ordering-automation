# VBA Application Analysis

## Scope and method

The attached `.xlsm` workbook was reviewed as an Excel/VBA application. The review covered:

- Workbook package structure and metadata
- Worksheet layout, values, formatting, defined names, and drawing assignments
- VBA Project Explorer structure
- Standard-module code and helper routines
- SAP GUI interaction logic and error handling

The original workbook was not modified. A separate analysis copy was used to inspect the workbook in Excel and the VBA editor. The macro was not enabled or executed, and no SAP session was connected.

## Executive summary

The application is a compact, spreadsheet-driven SAP procurement automation tool. It takes material and MRP-area values from Excel, uses SAP GUI Scripting to navigate MD04, identifies a planned-order entry, converts it into a purchase requisition, navigates to ME21N, fills configured purchasing fields, saves the resulting document, and writes a 10-digit document number or status text back to the worksheet.

The design is useful for reducing repetitive attended SAP work, but it is tightly coupled to SAP GUI screen paths and process-specific values. It should be treated as an internal operational tool or demonstration project unless it is first sanitized, parameterized, and tested against a controlled SAP environment.

## Application purpose and business problem

The workflow addresses a common procurement-operations problem: users must repeatedly transfer a list of materials and MRP areas into SAP, inspect MD04, convert eligible planning results, continue through purchase-document creation, and record the result.

The workbook centralizes the input list and automates repeated navigation. Its intended benefits are:

- Less repetitive typing and screen navigation
- More consistent execution of a documented SAP process
- Faster processing of multiple rows
- A visible result column for document numbers and errors
- Row-level identification of exceptions that require manual review

The application does not replace SAP authorization, business approval, or human review of the created documents.

## Workbook structure

The reviewed workbook has one visible worksheet, `Sheet1`.

| Area | Observed behavior |
|---|---|
| Input/output table | Columns A–C are used for material, MRP area, and result/document output |
| Launcher | A rounded worksheet shape labeled `MD04` is assigned to the main macro |
| Active input behavior | The macro uses the active worksheet rather than explicitly selecting a worksheet by name |
| Formatting | Duplicate-value conditional formatting exists for the material range; failed rows are colored yellow by VBA |
| Formula use | No meaningful business formulas were identified |
| Structured objects | No Excel Table, chart, PivotTable, or UserForm was identified |
| Workbook metadata | Author/revision metadata and an internal absolute path were present and should be removed before publication |

The workbook contains operational-looking sample rows. They are intentionally not reproduced in this report.

## VBA project inventory

The VBA Project Explorer showed:

- `Sheet1 (Sheet1)`
- `ThisWorkbook`
- `Module1`
- No UserForms
- No separate class modules

The business logic is concentrated in the standard module `Module1`.

### Main procedure: `SAP_Final_Solution`

The main procedure performs the complete workflow:

1. Attempts to attach to a running SAP GUI instance.
2. Obtains the first SAP GUI connection and first session.
3. Uses the active Excel sheet as the input source.
4. Finds the last used material row in column A.
5. Clears existing fill colors in the working range.
6. Loops through each input row.
7. Processes MD04, planned-order conversion, purchase-requisition discovery, and ME21N entry.
8. Extracts the SAP status-bar result.
9. Marks failures and continues to the next row.
10. Displays a completion message.

### Helper: `GetMD04Sub`

This routine searches likely MD04 subcontainer identifiers and returns the container that contains the MD04 table. It first checks a coarse range and then a more granular range before falling back to a default identifier.

This is a resilience technique for SAP dynpro variations, but it still depends on the underlying SAP screen structure remaining recognizable.

### Helper: `ExecuteSAP`

This wrapper centralizes common SAP GUI actions. The supported actions include selecting an item, setting text, focusing a control, setting caret position, selecting a checkbox, pressing a control, and setting selection indexes.

The wrapper uses `SafeGet` and raises a VBA error when the requested SAP element cannot be resolved.

### Helper: `SafeGet`

`SafeGet` attempts to resolve a recorded SAP GUI path across common SAPLMEGUI and SAPLMEVIEWS container variations. Successful paths are stored in a static `Scripting.Dictionary` cache so later calls can reuse the working path.

This approach reduces failures caused by predictable container-number changes, but it does not make the automation independent of SAP GUI layout changes.

### Helper: `ExtractPONumber`

This routine creates a late-bound `VBScript.RegExp` object and searches the SAP status message for the first standalone 10-digit number. If no match is found, it returns the full status-bar message.

This behavior should be improved before production use because a status message is not necessarily a purchase-order number. The result should be validated by message type, document type, and a stronger SAP-specific pattern.

## Detailed automation flow

### SAP session initialization

The macro calls `GetObject("SAPGUI")`, obtains the SAP GUI scripting engine, and selects the first connection and first child session. If SAP GUI is not running or no session is available, it displays an error message and exits.

The use of late-bound objects avoids a compile-time SAP reference, which simplifies deployment but moves many failures to runtime.

### Input preparation

The macro assigns the active worksheet to `wsInput`, finds the last used row in column A, and clears the fill color across the working input/output range. It then reads material and MRP-area values for each row.

Blank material rows are skipped. The macro does not provide a formal validation layer for duplicate values, invalid MRP areas, or already-processed rows.

### MD04 navigation and filtering

For each material, the macro maximizes the SAP window, enters MD04, fills the material and MRP-area controls, and submits the screen. It conditionally toggles a filter control and sets a filter-related value when the expected control is available.

The code intentionally avoids assuming that every SAP session starts in exactly the same filter state.

### Planned-order discovery

The macro locates the MD04 table through `GetMD04Sub`, loops through table rows, and adjusts the table scrollbar when the target row is outside the visible window. It reads the row text and searches for the planned-order marker.

After finding the row, it scans a range of likely columns for the detail icon. This dynamic icon search is intended to tolerate small column-position changes in the SAP grid.

If the planned-order row or its detail icon cannot be found, the macro raises a row-level error.

### Conversion to purchase requisition

The macro presses the SAP control that converts the planned order to a purchase requisition. It then switches the document into edit/change mode when needed, saves, and attempts to accept warning or information dialogs.

It refreshes MD04 and searches for the new purchase-requisition entry. Once that row is found, it opens the detail action and navigates toward ME21N.

### ME21N processing

The ME21N section includes recorded UI operations protected by `ExecuteSAP`. The visible logic includes:

- Handling the document overview area
- Copying or selecting a document-overview item
- Dismissing repeated SAP warnings using a bounded Enter loop
- Entering delivery-schedule information
- Selecting confirmation-related controls
- Opening or preserving header detail sections
- Entering text and communication values
- Saving the document and handling final save warnings

Some values are process-specific and are embedded in the module. They should be moved to a reviewed configuration area before the application is shared or reused in another organization.

### Result writing and row errors

On success, the macro reads the SAP status bar, passes it to `ExtractPONumber`, and writes the returned value into column C.

On failure, the row error handler:

- Colors the Excel row yellow
- Writes `ERROR: ` followed by the VBA error description into column C
- Attempts to close or dismiss a secondary SAP window
- Clears the error and continues with the next input row

The final message box is therefore a batch-completion notification, not a success report.

## Technologies and dependencies

### Languages and APIs

- VBA for application logic
- Excel object model for workbook, worksheet, cell, row, shape, and formatting operations
- SAP GUI Scripting API through late-bound COM objects
- Windows scripting objects for dictionary caching and regular expressions

### Required runtime conditions

- Microsoft Excel desktop
- Macros permitted by Excel security settings
- SAP GUI for Windows running and logged in
- SAP GUI Scripting enabled on the client/server where required
- Appropriate SAP authorization
- SAP screens and control paths compatible with the recorded automation

No external package manager, web service, database, or separate executable was observed.

## Hard-coded and environment-specific behavior

The following categories are embedded in the current implementation:

- SAP transaction names and GUI control paths
- MD04 table/container discovery assumptions
- Planned-order and purchase-requisition row markers
- ME21N screen paths and recorded actions
- At least one process-specific communication value
- Status-bar parsing assumptions
- Use of the first SAP connection and first SAP session

These are implementation details, not user-configurable features in the current workbook.

## Error-handling assessment

### Strengths

- Checks for missing SAP GUI/session conditions before starting
- Handles each input row independently
- Attempts to close warning dialogs
- Uses bounded retry loops for some ME21N warnings
- Highlights failed rows for follow-up
- Uses dynamic path and table discovery in several parts of the workflow

### Risks

- Broad `On Error Resume Next` blocks can suppress the original error
- There is no structured log with timestamp, row, SAP message, or screen state
- A completion message can be displayed even when one or more rows failed
- A status message without a 10-digit number is written verbatim to the output column
- The macro does not confirm that the document was created in the intended SAP system
- A rerun may create duplicate documents if the input rows are not checked first

## Security and privacy findings

The following items should be addressed before public release:

- The workbook contained live-looking material/MRP-area/document data in the input table.
- Workbook core properties contained a personal author name and revision information.
- Workbook package metadata contained an internal SharePoint-style absolute path.
- The compiled VBA project exposes the automation logic and should be treated as source code.
- SAP GUI paths and process values may reveal internal transaction design.
- Screenshots could expose vendors, quantities, prices, users, system names, or document numbers.

No hard-coded SAP password or credential was observed in the reviewed `Module1` code. A complete publication review should still inspect all future modules, workbook names/ranges, custom XML, connections, names, shapes, and screenshots.

## Recommended repository contents

```text
sap-po-ordering-automation/
├── README.md
├── LICENSE
├── .gitignore
├── docs/
│   ├── VBA_APPLICATION_ANALYSIS.md
│   └── screenshots/
├── src/
│   └── Module1.bas
├── workbook/
│   └── SAP_PO_Order_Automation_Sanitized.xlsm
└── samples/
    └── sample_input.csv
```

The original operational workbook should remain outside the public repository. The `.xlsm` and `.bas` files should be published only after code, metadata, data, and screenshots have been reviewed by the organization that owns the process.

## Suggested future improvements

1. Add a configuration sheet for environment-specific and process-specific values.
2. Validate inputs before any SAP transaction is opened.
3. Make SAP connection/session selection explicit.
4. Add a production-environment guard and a test-mode option.
5. Replace broad error suppression with structured error handling.
6. Create an audit log with row-level outcomes and SAP messages.
7. Verify the resulting document number with stronger SAP-specific checks.
8. Add duplicate prevention and safe rerun behavior.
9. Refactor the SAP workflow into smaller, testable procedures.
10. Create a non-SAP test harness for unit-level validation.

## Review conclusion

The workbook demonstrates practical Excel VBA and SAP GUI automation skills, especially in row-driven processing, dynamic SAP element discovery, warning handling, and result reporting. Its strongest portfolio value is as an example of solving a real procurement workflow with desktop automation.

For public GitHub use, present it as a sanitized demonstration and be explicit that the current implementation is attended, SAP-GUI-dependent, and process-specific. Avoid claiming unattended integration, generic SAP compatibility, or guaranteed transaction success unless those capabilities are added and tested.

