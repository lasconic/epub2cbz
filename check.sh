#!/bin/sh
#
# Small script to ensure quality checks pass before submitting a commit/PR.
#
python -m ruff format epub2cbz.py
python -m ruff check --fix epub2cbz.py
python -m mypy epub2cbz.py
