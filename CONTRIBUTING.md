# Contributing

Contributions are welcome: bug reports, fixes, tests, documentation,
and feature proposals.

## Before your first pull request

This project uses a Contributor License Agreement to keep its
licensing options intact (the project is AGPL-3.0 with the ability to
offer other license terms). Read [CLA.md](CLA.md) and include this
line in your first pull request description:

    I have read the CLA.md of this repository and agree to its terms.

Pull requests without CLA agreement cannot be merged, however good
the code is.

## Ground rules

- Bug reports: include the tool name, the exact call (parameters
  included), the observed result, and the expected result. A minimal
  reproducing document helps enormously.
- Fixes: every behavior change needs a test. The suite must stay
  green (`pytest tests -m "not live"`; live tests require Word
  installed and closed).
- Style: match the surrounding code. No em dashes in docstrings or
  documentation. Docstrings are token-budgeted and tested; run the
  docstring budget test after editing any tool description.
- Public copy (README, site, llms.txt) carries only script-generated
  counts; never hand-edit a number.

## What gets accepted

Small, well-tested fixes merge fastest. Larger features should start
as an issue describing the use case before any code is written; the
tool surface is deliberately consolidated and new tools need a strong
reason to exist.
