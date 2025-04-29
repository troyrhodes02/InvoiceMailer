#!/bin/bash
echo "🚀 Starting InvoiceMailer Mac Setup..."
cd "$(dirname "$0")"

echo "🔒 Setting executable permissions..."
chmod +x run-invoices.sh
chmod +x InvoiceMailerUI

echo "🛡️ Removing quarantine restrictions..."
xattr -d com.apple.quarantine * || true

echo "✅ Setup Complete! You can now double-click run-invoices.sh to start InvoiceMailer." 