# Contributing to pivoteer

Thanks for your interest in improving pivoteer. This guide explains how to
propose changes and how we review contributions.

## Getting Started

1. Fork the repository.
2. Create a feature branch from `main`.
3. Install dependencies and run tests.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .[dev]
pytest
```

## Development Guidelines

- Follow PEP 8 and keep changes Black-compatible.
- Use strict type hints for new code.
- Prefer small, focused commits with clear messages.
- Do not modify generated Excel binary fixtures unless requested.

## Tests

- Run `pytest` before opening a PR.
- If you add a feature, include tests for expected behavior.

## Pull Requests

- Use the PR template and describe the change clearly.
- Ensure CI is green.
- Link related issues.

## Release Process (Maintainers)

1. Update `CHANGELOG.md`.
2. Build artifacts: `python -m build`.
3. Verify metadata: `twine check dist/*`.
4. Create a GitHub release tagged `vX.Y.Z`.

## Security

Please report security issues via `SECURITY.md`.
