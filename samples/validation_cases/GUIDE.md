# Missing-group and invalid CSV tests

Use a separate, empty report folder. All files are UTF–16 and tab-separated; valid items are in HUF.

## Six choices
The missing_group_six_choices.csv file contains six invoice items on six separate documents. The pop-up windows appear in CSV order:

| Item code | Choice | Net value |
|---|---|---:|
| CHOOSE-NEW-PARTS | New parts | HUF 1000 |
| CHOOSE-USED-PARTS | Used parts | HUF 2000 |
| CHOOSE-LABOR | Labor | HUF 3000 |
| CHOOSE-SHIPPING | Shipping | HUF 4000 |
| CHOOSE-TRAILER | Trailer | HUF 5000 |
| CHOOSE-NONE | None of these | HUF 6000 |

For September 18, 2026, the preview should show increases of HUF 1000 for new parts, 2000 for used parts, 3000 for labor, 4000 for shipping and 5000 for Trailer. All other daily amounts and the monthly CONSUMABLES/DIAGNOSTICS/AIR_CONDITIONING totals should be zero. There is no workshop parts total because the labor is on a different document. Number of new documents: 6; the item assigned None of these is also recorded, but it has no aggregation category. After saving, reopen the report and check the daily values.

## Closing with X
When importing missing_group_cancel.csv, close the group-selection window using X. A cancellation message is expected, rather than an import error. There must be no preview or save; the Import button must become usable again. When importing again, the same item should prompt for a choice again.

## Non-invoice transactions
The missing_group_non_invoice.csv file must not prompt for a group: the ungrouped HUF 700 purchase does not increase the purchase total, while the inventory item is skipped. The purchase document is still recorded as processed. Number of new documents: 1.

## Invalid CSV
In invalid_net_value.csv, invalid_date.csv, invalid_currency.csv and missing_document_number.csv, the first data row is valid and the second is deliberately invalid. The error should point to line 3 of the file and identify the invalid value or the missing document number. There must be no preview or save, and even the first valid row must not be partially imported.

The Currency header is missing from missing_column.csv. A message about a missing CSV column is expected; no line number is required for this.

After the error window is closed, the program can start another CSV selection. Select valid_retry.csv: it should be possible to import it with a HUF 1000 increase in new parts and 1 new document. This valid file contains the same document as the first valid row in the invalid files, demonstrating that the failed read did not save anything partially. Perform each invalid-file test in a separate, empty folder, or import the valid test file only once, after all invalid-file tests.
