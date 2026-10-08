#!/bin/bash
cd -- "$(dirname -- "$0")" || exit 1
if [ ! -f cohort_extract_gui.py ]; then
  echo "Extract the entire ZIP before running this launcher."
  read -r -p "Press Return to close. "
  exit 1
fi
python3 cohort_extract_gui.py
status=$?
if [ "$status" -ne 0 ]; then
  echo "Python 3 with Tkinter is required. See README.md."
  read -r -p "Press Return to close. "
fi
exit "$status"
