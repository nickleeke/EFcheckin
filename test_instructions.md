# Running SPED Lead Tests

To run the tests in Google Apps Script:

1. Open the Apps Script editor
2. Select the `Tests.gs` file
3. From the function dropdown, select `runAllSpedLeadTests`
4. Click the Run button
5. Check the execution log for "All SPED Lead tests passed"

## Current Test Coverage

- `test_spedlead_getSpedLeadsReturnsArray` - Tests getSpedLeads_ returns array
- `test_spedlead_getSpedLeadsReturnsEmptyWhenNone` - Tests getSpedLeads_ returns empty array when no leads
- `test_spedlead_getUserStatusReturnsRole` - Tests SPED Lead role detection
- `test_spedlead_getEvalMetricsReturnsZeroForEmpty` - Tests eval metrics with empty sheet
- `test_spedlead_syncHandlesEmptyCaseloads` - Tests sync with no connected case managers

## Test: syncHandlesEmptyCaseloads

This test verifies that `syncSpedLeadDashboard()` correctly handles the case where:
- A SPED Lead exists
- The SPED Lead has a provisioned spreadsheet
- The SPED Lead has zero connected case managers

Expected result:
- `success: true`
- `syncedCount: 0`
- `failedCount: 0`
- Last sync timestamp is updated in ScriptProperties
