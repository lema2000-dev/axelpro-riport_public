# CSV–Excel Report Automation – User Guide

The application fills annual Excel sales reports based on movement CSV exports. The program writes daily and monthly net amounts in HUF and converts EUR items using the stored daily exchange rate.

The development work plan, the status of checks and the detailed processing rules can be found in [Statement_of_Work.md](Statement_of_Work.md).

## The version presented

This is the documentation for the Hungarian-language, anonymized portfolio version. The project demonstrates the automation of a real reporting workflow, without personal names or labels linked to the client. The demonstration data contains fictional transactions.

The Hungarian and English versions are on separate Git branches. You can switch between them using GitHub's branch selector.

In the Hungarian version, the CSV headers, categories and Hungarian interface names are retained. The executable filenames and settings paths below follow the current technical identifiers.

## Launch and initial setup

1. Launch the supplied `SalesReport.exe` file. The packaged application does not require Python to be installed, and no external template file is needed.
2. On first launch, browse to an existing folder that you can write to and where you want to save the annual reports. For example, create a **Sales Reports** folder under **Documents**. Do not select the root of `C:\` as the destination folder; the program does not need to be run as administrator for normal use.
3. Enter the names of additional new-part groups separated by semicolons, for example `Filters;Lubricants`, then click **Done**. If there are no additional groups, you can leave this field empty.
4. The main menu appears. The program retains the settings for the next launch.

The program already treats the `New parts` group as new parts, so you do not need to add it separately. Group names are recognized without regard to case. The names must match the group names in the export.

`SalesReport-dev.exe` launches the same graphical interface with a console window; `SalesReport.exe` is intended for normal use.

## Preparing and editing CSV files

You can import the original export directly. If you want to view or modify its contents beforehand, we recommend **Rons Data Edit**, which we also used during the Windows tests.

Rons Data Edit displays the file as a table, treats values as text and supports the export's UTF–16 encoding. This is important because Excel can automatically convert dates, numbers or identifiers when opening and resaving CSV files; such changes may cause the reporting program to reject the file.

- [Official product page](https://www.ronsplace.ca/products/ronsdataedit)
- [Official download](https://www.ronsplace.ca/products/download) – select **Download Install** in the **Rons Data Edit** section. The **Download Portable** version is intended for use without installation.
- The free **Lite** version can save up to **2500 rows**. Modifying larger files requires the unrestricted version. This is a limitation of the CSV editor, not an import limit of the reporting program.

When editing, preserve the **UTF–16 encoding**, **tab-separated format**, original column headers and the text of fields you have not changed. If the editor asks for settings, use **Tab** as the separator and **UTF–16** as the encoding. Do not run automatic date or number conversions. Do not resave CSV files from Excel for this import; you can, however, use the completed `.xlsx` reports in Excel.

Required columns:

| Column | Expected content |
| --- | --- |
| Movement timestamp | The export's original date format, for example `2026. 09. 18. 10:00:00`. |
| Transaction type | The transaction type used in the export. Inventory items are skipped during import. |
| Document number | The processed item's nonempty document number. |
| Partner | Required column; the program automatically skips items belonging to the two excluded partners. |
| Group | Group name; an empty invoice group prompts for a choice, while the amount of a purchase with an empty group is excluded. |
| Item code | The item's item code; separate monthly totals are maintained for the `SANY`, `DIAG`, `KJ` codes. |
| Net value | Net value, for example `1250,50`. |
| Currency | `HUF` or `EUR`. |

### Automatically skipped items

In addition to inventory items, the program **skips rows regardless of transaction type** if their Partner field contains one of the following names:

- Company 1
- Company 2

`Company 1` and `Company 2` are fictional names: the exclusion values in the public CSV reader and the partner names in the test CSVs must also match them. Renaming labels in the spreadsheet alone does not change the reader's filtering.

The comparison uses an exact full-name match, ignoring leading and trailing spaces and differences in case. These items are not included in totals or in the document register. The program does not request a classification even if an excluded partner's invoice has an empty group.

For **purchases, only the amounts of items with a group contribute** to the daily purchase total. Purchases with an empty group or a group containing only spaces are excluded from this total, but their documents are recorded as processed. There is no group-selection window for purchases. The rule applies to both HUF and EUR items.

## Importing

1. Close the affected annual reports in Excel so that the program can save changes to them.
2. Click **Import** and browse to the CSV file.
3. Wait for processing to finish. The Import button cannot be used during this time. If multiple years are involved, the program works with separate annual reports.
4. If the program requests a group selection or an exchange rate, complete the displayed window as described below.
5. In each yearly preview, check the destination file and the amounts to be added.
6. Choosing **Accept** saves that year. Choosing **Reject** leaves that year's changes unsaved and continues with the next year.
7. At the end of the process, the Import button becomes usable again. You can open the saved report in Excel, where the formulas will be recalculated.

The filename is `Sales_Report_YEAR.xlsx`, for example `Sales_Report_2026.xlsx`. A new file contains all twelve months, **Összesítő**, **Euro**, and the hidden exchange-rate and document registers.

### Using the monthly report

On monthly worksheets, the **first three rows and columns A–C are frozen**. They remain visible while scrolling; the first scrollable cell is **D4**.

Some fields in the demonstration template have been given neutral names or a `-` label. The hyphen only replaces the label: the underlying cells and remaining formulas continue to work.

### Missing group

For an invoice item with an empty group that is not excluded based on its partner, the window displays the document number and item code. Choose one of the six buttons:

- New parts
- Used parts
- Labor
- Shipping
- Trailer
- None of these

**None of these** leaves the group empty. As a result, the item is not included in group-based aggregation, but separate monthly aggregation by item code may still apply. The choice changes only the item in the current import; it does not rewrite the original CSV. The selection window's **X** button cancels the import.

### Entering an exchange rate

The program retrieves the current MNB EUR/HUF exchange rate and uses daily exchange rates already stored in the reports. For a missing day, it uses the rate with the latest known source date that is not later than the target day. If no such rate exists, or the current rate cannot be retrieved, manual input is requested.

Enter a positive exchange rate, for example `380,25`, and a valid source date in `YYYY-MM-DD` format, for example `2026-09-18`. With manual input, the source date may differ from the item's date and may even be later. If the number or date is invalid, the program requests new input.

The monthly exchange rates on the **Euro** worksheet are used by the formulas for manually completed EUR fields. In a new workbook, these start at zero. Import updates only the monthly rate for the current month of the current year; before using manual EUR amounts for an earlier month, check and, if necessary, fill in that month's exchange rate. Daily conversion of EUR items from the CSV works separately from this.

### Repeated imports and cancellation

The program records imported documents **by year and document number**. Importing the same complete file again does not increase the amounts again. From a file containing both known and new documents, only the items belonging to new documents are imported.

A CSV row added or modified after its document has already been imported does not update the report: all items belonging to that document are skipped. Such corrections must be agreed separately; reimporting the CSV is not a correction operation.

Closing the main window requests cancellation. The program waits for background processing to stop. Changes for years that have not yet been saved are lost, while years already accepted and successfully saved are retained.

## Main menu

| Button | Behavior |
| --- | --- |
| Import | CSV selection, processing and a preview for each year. |
| Add new-part groups | Enter additional new-part groups separated by semicolons. |
| Remove new-part groups | Remove previously added additional groups. |
| Change report folder | Select another existing destination folder. Previously created reports are not moved. |
| View current settings | Display the report folder and additional new-part groups. |
| Exit | Close the application. |
| Reset and close | Delete the settings after confirmation and close the application. |

Enter at least one group in the add or remove window. If the input is invalid, correct it or choose **Cancel**.

**Resetting does not delete the report folder or the Excel reports.** It deletes the saved destination-folder setting and group list. On the next launch, the program requests them again. If you select the same report folder, the exchange-rate and document registers in the annual files there are retained.

## Errors and warnings

If an error occurs, read the message and close it using **OK**. For invalid CSV files, the message identifies the problematic field and, for row errors, the line number as well. Correct the file in Rons Data Edit, or create a new export, then restart the import.

If the program cannot save, check that the report is not open in Excel and that the destination folder is available and writable. In a multi-year import, years successfully saved before a later save error are retained.

A workbook structure or formula error interrupts the import. Use the message to restore the indicated worksheet or cell. Of the monthly sheets, the program checks those affected by the import.

Import may continue if a formatting or dimension difference is found. The program displays only the first such warning; check the spreadsheet at the indicated location. It accepts formula and formatting variations that have the same meaning and result from simply resaving the file in Excel.

## Preserving the reports

Fields intended for manual data entry in the reports can still be filled in using Excel. Preserve worksheet names, dates, headers, merged ranges, formulas and hidden registers.

The program saves through a temporary file and **does not create automatic backups**. The user is responsible for preserving the reports and the document register stored in them. The program aggregates net amounts and rounds to whole HUF; importing in batches may produce a small rounding difference compared with a single complete import.

The settings file on Windows is: `%APPDATA%\axelpro_riport\settings.json`.
