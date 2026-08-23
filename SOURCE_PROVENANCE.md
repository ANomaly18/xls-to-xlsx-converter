# Source provenance

The initial implementation was extracted and adapted from the standalone XLS conversion utility implemented for DCApps.

## Included

- Generic binary BIFF `.xls` conversion
- Generic HTML-disguised `.xls` conversion
- Desktop GUI and command-line interfaces
- Formatting and merge preservation helpers
- Runtime-generated synthetic tests

## Excluded

- DCApps application source, configuration, credentials, and database access
- Production spreadsheets and converted results
- Build artifacts, caches, logs, and temporary files
- Operational health records and private fixtures

The open-source extraction adds offline enforcement, resource bounds, atomic output, overwrite protection, safe hyperlink handling, formula-injection neutralization, package metadata, and continuous integration.
