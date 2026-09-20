# Sunday Service slides app — local setup & run.
#
#   make install   create .venv and install dependencies
#   make run       install (if needed) then launch the Streamlit app
#   make stop      stop a running app
#   make verify    check the Google Sheets connection
#   make clean     remove .venv
#
# On this Mac, cryptography/google-auth must come from prebuilt wheels, hence
# --only-binary=:all: below.

VENV    ?= .venv
PYTHON  ?= python3
PIP     := $(VENV)/bin/pip
STREAMLIT := $(VENV)/bin/streamlit
PORT    ?= 8501

.PHONY: install run stop verify clean help

help:
	@grep -E '^[a-z]+:.*?##' $(MAKEFILE_LIST) | sed 's/:.*##/\t/' | sort

$(VENV)/bin/python:
	$(PYTHON) -m venv $(VENV)

$(VENV)/.installed: $(VENV)/bin/python requirements.txt
	$(PIP) install --upgrade pip
	$(PIP) install --only-binary=:all: -r requirements.txt
	@touch $@

install: $(VENV)/.installed  ## Create .venv and install dependencies

run: install  ## Launch the app locally
	$(STREAMLIT) run app.py --server.port $(PORT)

stop:  ## Stop a running app
	-pkill -f "streamlit run app.py"

verify: install  ## Check the Google Sheets connection
	$(VENV)/bin/python -c "import tomllib, store; \
	s = store.get_store(tomllib.load(open('.streamlit/secrets.toml','rb'))); \
	print('backend:', type(s).__name__); print('weeks:', list(s.all().keys()))"

clean:  ## Remove .venv
	rm -rf $(VENV)
