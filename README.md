# Darwill AI Prospector 3.6

## Permanent Master Prospect Database CSV

Version 3.6 preserves all v3.5 Deep Contact Data Recovery features and adds a
permanent master CSV that can be shared across application versions.

### Select the file once

Open **Master Prospect Database** and use the **Permanent Master CSV** section.

Select:

`Darwill_Master_Prospect_Database.csv`

The application remembers the full file path in its settings.

### Startup behavior

On launch, the app:

1. Loads the selected permanent CSV
2. Imports or updates its records in the internal SQLite master database
3. Shows the exact loaded file path
4. Shows the number of master prospect records
5. Shows the file's last-updated time

If no path was previously saved, the app searches common Documents, OneDrive,
Desktop, and Darwill folders for the standard filename.

### Automatic synchronization

When enabled, every completed search run writes the complete internal master
database back to the same permanent CSV.

The CSV is also synchronized after important lifecycle changes.

### Automatic backups

Before replacing the permanent CSV, the application creates a timestamped
backup in:

`master_csv_backups`

The write uses a temporary file and an atomic replace to reduce corruption risk.

### Existing 3.4/3.5 records

The permanent CSV is imported without discarding later lifecycle statuses.
Approved, HubSpot-synced, in-sequence, meeting, customer, and lost statuses are
preserved.

### Internal database remains authoritative during a run

The app continues using SQLite for fast deduplication and lifecycle operations.
The permanent CSV is the portable, cross-version copy and is synchronized with
the internal database.

## Deep Contact Data Recovery retained

Version 3.6 includes all v3.5 features:

- person-specific searches
- company-page crawling
- public PDF extraction
- stronger email-pattern evidence
- MX/A-record domain checks
- Needs Verification queue
- default sequence block for predicted or unverified emails
- recovery evidence in Contact Intelligence and Excel

## First launch

Double-click `install_and_run.bat`.

Then open **Master Prospect Database**, confirm the permanent CSV path, and
click **Load / Import** once.

## Later launches

Double-click `Run Darwill AI Prospector.bat`.
