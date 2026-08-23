# Security policy

## Supported versions

Security fixes apply to the latest revision on the default branch.

## Reporting

Use the repository's private GitHub security-advisory reporting flow. Do not attach confidential spreadsheets, credentials, personal records, or production exports. Provide a minimal synthetic reproduction instead.

## Threat model

Spreadsheet and HTML parsers handle attacker-controlled structured input. The converter therefore:

- limits input size and expanded HTML table dimensions;
- validates the `.xls` extension and inspects content signatures;
- ignores external CSS and performs no network fetch during conversion;
- neutralizes HTML text beginning with `=` before writing XLSX cells;
- retains only bounded HTTP(S) hyperlinks without embedded credentials;
- writes through a temporary file and atomically moves a successful result; and
- refuses to overwrite existing output unless the caller explicitly opts in.

These controls reduce risk but do not make untrusted workbooks safe. Scan unknown files and run high-risk conversions in a disposable, least-privilege environment.

## Dependencies

Binary `.xls` parsing uses `xlrd` 2.0.2 or newer. Version 2.x continues to support the historical `.xls` format while removing unrelated workbook formats. Keep parser dependencies current and run `python -m pip_audit` before releases.

The application does not preserve or execute macros, VBA, embedded programs, or spreadsheet formula definitions.
