# Contributing

Contributions are welcome through focused pull requests.

1. Install development dependencies with `python -m pip install -e ".[dev]"`.
2. Add or update tests for every behavior change.
3. Run `python -m pytest` and `python -m pip_audit`.
4. Describe security and compatibility implications in the pull request.

Never commit real exports, identifiers, health records, credentials, or customer workbooks. Tests and issue reproductions must generate obviously synthetic fixtures at runtime.
