.PHONY: venv install install-dev lint type test smoke run-example format

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
        $(PIP) install -e .

install-dev: venv
        $(PIP) install -e .[dev]

format:
        $(PY) -m ruff format src tests

lint:
        $(PY) -m ruff check .

type:
        $(PY) -m mypy src tests

test:
	$(PY) -m pytest -q

smoke:
        $(PY) -m tests.smoke_fetch

run-example:
        $(PY) -m bpauto.cli --excel "Liste BP Cleaning Kreditoren.xlsx" --sheet "Tabelle1" --start 3 --name-col C
