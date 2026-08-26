HOLYOKE BANK TRANSFER CONVERTER - WINDOWS EXE
================================================

THIS VERSION HAS THE BANK-CODE MAP BUILT IN.

Embedded mapping:
- 44 account numbers currently mapped.
- Matching is based on the LAST 4 DIGITS of the bank account number.
- Staff do NOT need to upload a separate bank-code workbook.

NORMAL STAFF WORKFLOW
---------------------
1. Open HolyokeTransferConverter.exe.
2. Select the raw bank transfer CSV.
3. Choose where to save the output workbook.
4. Add any extra scrub pairs only if needed.
5. Click Convert.
6. The converter creates the CB-ready workbook.

BUILT-IN TRANSFER RULES
-----------------------
- TRANSFER FROM = negative.
- TRANSFER TO = positive.
- Each transfer pair nets to zero.
- Debit first, credit second.
- Dates use mm-dd-yyyy with no timestamp.
- Column D = TRANSFER.
- Columns G and K contain the signed transfer amount.
- Column L contains the bank code.
- Negative amounts are red; positive amounts are black.
- Output is split into monthly tabs.
- Descriptions use bank codes:
  TRANSFER FROM [FROM CODE] TO [TO CODE]

BUILT-IN SCRUBS
---------------
- 1171 <-> 1220
- 1171 <-> 1270
- Any transfer where BOTH sides are among:
  7189, 4237, 4245

UNRECOGNIZED ACCOUNTS
---------------------
If a new 4-digit account is encountered that is not in the embedded map,
the converter will finish the workbook and display the unmapped account(s).
Update the source mapping and rebuild the EXE before broad distribution.

BUILDING THE EXE WITH GITHUB ACTIONS
------------------------------------
Repository structure should be:

.github/
  workflows/
    build-windows.yml
BankTransferConverter_EXE/
  desktop_app.py
  requirements.txt
  build_windows.bat
  README.txt

Run the "Build Windows EXE" workflow under GitHub Actions.
Download the HolyokeTransferConverter-Windows artifact after a successful build.
