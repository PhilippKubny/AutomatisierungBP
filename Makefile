.PHONY: venv install lint type test smoke run-example

PYTHON ?= python3
VENVDIR ?= .venv
ifeq ($(OS),Windows_NT)
VENVBIN := $(VENVDIR)/Scripts
else
VENVBIN := $(VENVDIR)/bin
endif
PY := $(VENVBIN)/python
PIP := $(VENVBIN)/pip

venv:
	$(PYTHON) -m venv $(VENVDIR)

install: venv
	$(PIP) install -e .[dev]

lint:
	$(PY) -m ruff check .

type:
	$(PY) -m mypy src cli.py tests

test:
	$(PY) -m pytest -q

smoke:
	$(PY) -m tests.smoke_fetch

run-example:
	$(PY) cli.py --excel "Liste BP Cleaning Kreditoren.xlsx" --sheet "Tabelle1" --start 3 --name-col C
