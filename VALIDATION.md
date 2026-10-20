# Validation

10 Python test cases passed on CPython3.10.21 and3.14.7 on Linux.

Test command: `uv run python -m unittest discover -s tests -v`

LibreOffice recalculated sample, changed quantities/payment and rounding fixture. All subtotal/tax/total/balance cells matched independent Decimal targets. Sample total INR2488.36 and balance INR1988.36. Formula cells remained formulas. Actual PDF rendering visually inspected. Microsoft Excel was not available.

Tests are a bounded check, not a production/security certification. Browser checks used Linux Chromium, not physical phones or Safari/Firefox. Local data is unencrypted. All examples are synthetic; screenshots show actual running output, not design mockups.
