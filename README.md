# SAP Purchase-Order Automation with Excel VBA and SAP GUI Scripting

Excel VBA and SAP GUI Scripting automation for attended procurement purchase-order processing. The workflow reads material and MRP-area inputs from Excel, navigates SAP MD04, converts eligible planned orders into purchase requisitions, continues through ME21N, and writes a result or error message back to Excel.

> Important: This automation can create SAP business documents. Use it only with authorized accounts, approved processes, and a controlled non-production test system.

## Public repository scope

This public repository contains sanitized VBA source, documentation, and dummy CSV input. It intentionally excludes the operational XLSM workbook, live SAP data, credentials, screenshots, and a runnable public demo.

Workbook-specific details described in VBA_APPLICATION_ANALYSIS.md were obtained from a separate workbook review. They should be validated against an approved internal copy before use.

## Business problem

Procurement users may need to transfer a list of materials and MRP areas into SAP, inspect MD04, convert eligible planning results, create purchase orders, and record the outcomes. Repeating those steps manually can be time-consuming and inconsistent.

This project demonstrates how Excel can act as a controlled input and results sheet while SAP GUI Scripting performs the repetitive attended navigation.

No measured time-saving or throughput claim is published because this repository contains no benchmark data. Any business benefit should be measured in an approved test or operational environment.

## Application workflow

1. Enter material and MRP-area values in the approved Excel workbook.
2. Start the macro from the worksheet or Excel's macro list.
3. Open MD04 and submit the material and MRP-area values.
4. Search the MD04 table for an eligible planned-order entry.
5. Convert the planned order into a purchase requisition.
6. Refresh MD04 and locate the resulting purchase-requisition entry.
7. Navigate to ME21N and enter the configured purchasing values.
8. Save the purchase order and read the SAP status message.
9. Write the result to Excel or mark the failed row for review.

## Features verified in the published VBA source

The published Module1.bas source visibly contains:

- SAP GUI attachment through late-bound scripting objects.
- MD04 navigation using material and MRP-area inputs.
- Dynamic MD04 table-container discovery.
- Planned-order and purchase-requisition row searches.
- Detail-icon discovery across a range of possible columns.
- ME21N delivery, invoice, confirmation, text, and communication actions.
- SAP warning-dialog handling and bounded retry loops.
- Purchase-order result extraction from the SAP status bar.
- Row-level error messages and yellow highlighting.
- Helper routines for SAP path discovery, caching, and common actions.

These are source-level findings, not a claim that the macro has been successfully executed in every SAP environment.

## Excel input and output

The workbook structure described by the separate analysis uses:

| Column | Purpose |
|---|---|
| A | Material identifier supplied to SAP |
| B | MRP area supplied to SAP |
| C | Purchase-order result or error message |

The workbook itself is not included in this public repository. Use sample_input.csv or an approved internal test workbook with dummy values.

## Requirements

- Windows desktop environment
- Microsoft Excel desktop with macros permitted
- SAP GUI for Windows
- SAP GUI Scripting enabled and permitted by the SAP environment
- An authorized SAP connection and session
- Authorization for the relevant MD04, purchase-requisition, and ME21N activities
- SAP screen layouts compatible with the recorded control paths

No Python packages, web services, databases, or separate executables are required by the published VBA module.

## Internal or test-environment setup

1. Open the approved internal XLSM workbook in Excel desktop.
2. Import or update the reviewed Module1.bas module only in an approved copy.
3. Start SAP GUI and log into the intended test system, client, and session.
4. Confirm SAP GUI Scripting and required authorizations.
5. Enter one or two dummy material and MRP-area rows.
6. Confirm that the intended Excel worksheet is active.
7. Run SAP_Final_Solution.
8. Review column C and yellow rows.
9. Verify any created documents directly in the approved SAP test environment.

Do not run the public source against production data without process-owner approval. Do not run duplicate rows or multiple concurrent Excel/SAP sessions unless explicitly approved.

## Error handling and logging

The current module:

- Exits with a message if SAP GUI or an active session is unavailable.
- Processes rows independently.
- Writes an ERROR message into column C when a row fails.
- Highlights failed rows in yellow.
- Displays a completion message after the row loop ends.

The completion message means that processing attempts have finished; it does not prove that every purchase order was created. The current repository does not contain a structured audit log with timestamps, SAP messages, retry counts, or transaction outcomes.

## Security and publication guidance

Before publishing any workbook, source export, or screenshot, remove or replace:

- Material, MRP-area, plant, requisition, purchase-order, supplier, and price data.
- SAP system names, client numbers, hostnames, and internal connection details.
- Personal names, email addresses, workbook author metadata, and internal paths.
- Credentials, tokens, passwords, recordings, or company-only documents.

The reviewed Module1.bas source does not show a hard-coded SAP password or credential. This does not replace a complete review of the workbook package, embedded VBA project, custom XML, names, connections, and screenshots.

The compiled VBA project may expose proprietary process logic. Publish it only when disclosure is approved by the process or data owner.

## Limitations and operational risks

- This is attended desktop automation, not a server-side SAP integration.
- The macro attaches to the first available SAP connection and session.
- SAP GUI control paths can change after upgrades, personalization, or transaction customization.
- Several ME21N values and control paths are embedded in the VBA source.
- Broad On Error Resume Next handling can hide the original cause of a failure.
- Repeated warning acceptance can accept an unexpected SAP dialog.
- The result parser assumes that the first 10-digit number in the status bar is the purchase-order number.
- There is no production-environment guard, duplicate prevention, rollback, or structured retry history.
- The code uses the same input value for the MD04 MRP-area and plant fields; confirm that this is correct for the intended process.

## Testing status

This repository includes a non-SAP automated contract test suite and a GitHub Actions workflow that validate required public files, sanitized VBA source contracts, and sample CSV shape. The checks do not execute SAP GUI actions or create business documents. The macro was not executed against SAP during the review, and no claim is made that it has successfully created a purchase requisition or purchase order in every SAP environment.

## Suggested future improvements

- Allow the user to select the SAP connection and session explicitly.
- Add input validation for blank values, invalid MRP areas, duplicates, and already-processed rows.
- Add a test-system or production-environment guard.
- Move environment and purchasing values into a reviewed configuration area.
- Replace broad error suppression with structured handling and an audit log.
- Validate the SAP message type and document type before recording a purchase-order number.
- Add safe rerun and idempotency checks.
- Refactor SAP interaction, validation, configuration, and reporting into smaller modules.
- Add a mock or recorded harness that does not create live SAP documents.
- Add sanitized screenshots from an approved test environment.

## Sample input

The included sample_input.csv contains dummy values only:

| Material | MRP Area | PO Number |
|---|---|---|
| DEMO-MAT-001 | TEST-MRP | blank before run |
| DEMO-MAT-002 | TEST-MRP | blank before run |

## Suggested demonstration screenshots

Use only dummy test data and add the caption: “Illustrative screenshot from a test environment — no production data.”

1. Excel input sheet with material, MRP area, and blank result columns.
2. MD04 planned-order processing with identifiers redacted.
3. ME21N entry using safe test values.
4. Successful dummy result in Excel.
5. Failed row highlighted with a non-sensitive error message.

## Recruiter and portfolio summary

Built an Excel VBA and SAP GUI Scripting automation tool for attended procurement processing. The solution reads material and MRP-area inputs, navigates SAP MD04, converts eligible planned orders into purchase requisitions, creates purchase orders through ME21N, captures document results in Excel, and highlights row-level exceptions. The project demonstrates procurement process analysis, Excel automation, SAP GUI integration, dynamic UI handling, and operational error reporting.

## Related documentation

- VBA source: Module1.bas
- Workbook review: VBA_APPLICATION_ANALYSIS.md
- Dummy input: sample_input.csv
- Automated contract tests: tests/test_repository_contract.py
- Controlled validation matrix: TEST_PLAN.md
