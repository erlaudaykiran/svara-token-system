#!/usr/bin/env bash
cd "$(dirname "$0")/backend" || exit 1
if [ ! -d ../venv ]; then
  echo "First run: installing requirements..."
  python3 -m venv ../venv && ../venv/bin/pip install -r requirements.txt || exit 1
fi
../venv/bin/python app.py
