# InvoiceMailer User Guide - macOS Version

Welcome to InvoiceMailer!

This app lets you send invoice emails automatically using your Microsoft 365 tenant.

---

## How to Set Up:

1. Open the InvoiceMailer folder.
2. Open Terminal in this folder.
3. Run:
   ./run-invoices.sh
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