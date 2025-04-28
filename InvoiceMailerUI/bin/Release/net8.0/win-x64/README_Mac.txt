# InvoiceMailer User Guide - macOS Version

Welcome to InvoiceMailer!

This app lets you send invoice emails automatically using your Microsoft 365 tenant.

---

## How to Set Up:

1. Extract the InvoiceMailer-macOS.zip file to a folder on your computer.
2. Run the installer script first:
   - Double-click `install-mac.sh`
   - If prompted, choose "Open" in the security dialog
   - Alternatively, open Terminal and run: `chmod +x install-mac.sh && ./install-mac.sh`

3. After installation, run the application:
   - Double-click `run-invoices.sh`
   - If prompted, choose "Open" in the security dialog

4. On first launch, enter your:
   - TenantId
   - ClientId
   - ClientSecret
   (The app will save these automatically.)

5. Edit `recipients.csv` to map invoice keys to email addresses.
6. Place invoice files into the `/invoices/` folder.

---

## How to Run:

Every time you want to send invoices:
- Double-click `run-invoices.sh`
OR
- Open Terminal and run:
   ./run-invoices.sh 

---

## What the Installer Does:

The `install-mac.sh` script:
- Sets executable permissions on required files
- Removes macOS quarantine restrictions
- Makes the application ready to run 