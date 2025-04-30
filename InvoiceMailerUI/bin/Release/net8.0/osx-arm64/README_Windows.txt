# InvoiceMailer User Guide - Windows Version

Welcome to InvoiceMailer!

This app lets you send invoice emails automatically using your Microsoft 365 tenant.

---

## How to Set Up:

1. Extract the InvoiceMailer-Windows.zip file to a folder on your computer.
2. Run the installer script first:
   - Double-click `install-windows.bat`
   - The script will verify your system is ready

3. After installation, run the application:
   - Double-click `run-invoices.bat` (recommended)
   - Or you can use `InvoiceMailerUI.exe` directly

4. On first launch, enter your:
   - TenantId
   - ClientId
   - ClientSecret
   (The app will save these automatically.)

5. Edit `recipients.csv` to map invoice keys to email addresses.
6. Place invoice files into the `invoices` folder.

---

## How to Run:

Every time you want to send invoices:
- Double-click `run-invoices.bat` (recommended)
OR
- Double-click `InvoiceMailerUI.exe`

---

## Troubleshooting:

If you get a security warning:
1. Right-click the file (`run-invoices.bat` or `InvoiceMailerUI.exe`)
2. Select "Properties" 
3. Check "Unblock" option
4. Click "Apply" and "OK"

If the console window closes immediately after an error:
- Use `run-invoices.bat` which keeps the console window open
- Check the log files in the `logs` folder 