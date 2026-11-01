# Test Plan

## Scope

This repository can validate public artifacts without connecting to SAP or creating business documents. SAP execution remains an attended, controlled test-environment activity.

## Automated checks

The GitHub Actions workflow runs on pushes to main, pull requests, and manual dispatch. It verifies:

- Required public files are present.
- Workbook binaries (.xlsm and .xlsb) are not committed.
- The VBA source contains the expected public contract and no unsupported MD04 fallback.
- The sample CSV has the expected three-column shape.
- The standard-library Python contract tests pass.

Run the same checks locally with:

~~~text
python -m unittest discover -s tests -v
~~~

## Controlled SAP test matrix

| Scenario | Expected result |
|---|---|
| SAP GUI is not running | Macro shows a connection error and exits. |
| No active SAP session | Macro shows a session error and exits. |
| Material is blank | Row is skipped without SAP navigation. |
| MRP area is blank | Row is marked yellow with an input-validation error. |
| MD04 table container cannot be identified | Row fails before using a guessed screen ID. |
| Planned order is not found | Row is marked yellow; later rows continue. |
| Purchase requisition is not found after conversion | Row is marked yellow; later rows continue. |
| ME21N warning persists after bounded retries | Row fails instead of silently skipping the action. |
| Successful controlled test row | Verify the created document directly in SAP and reconcile column C. |

## Safety prerequisites

- Use an approved non-production SAP system and authorized test account.
- Use dummy materials, MRP areas, suppliers, and purchasing values.
- Confirm the intended Excel worksheet and SAP session before running.
- Back up the approved test workbook before importing the VBA module.
- Do not run duplicate rows or concurrent sessions unless the process owner approves it.

## Limitations

There is no SAP mock layer, live-SAP test run, performance benchmark, rollback mechanism, or idempotency guarantee in this public repository. A successful GitHub Actions run proves only that the repository contract checks passed.
