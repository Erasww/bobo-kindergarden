const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs/promises');
const path = require('path');

const app = express();
const PORT = 3000;
const workbookPath = path.join(__dirname, 'enrollments.xlsx');
const headers = [
  'Submitted At',
  'First Name',
  'Last Name',
  'Phone',
  'Email',
  'Child Name',
  'Age',
  'Program',
  'Message'
];

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});
app.use(express.static(path.join(__dirname)));

app.get('/health', (req, res) => {
  res.json({ success: true, message: 'Server is running' });
});

async function getWorkbookAndSheet() {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(workbookPath);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  let sheet = workbook.getWorksheet(1);
  if (!sheet) sheet = workbook.addWorksheet('Enrollments');
  const firstRowValues = sheet.getRow(1).values.slice(1);
  if (sheet.rowCount === 0 || firstRowValues.length === 0) sheet.addRow(headers);
  return { workbook, sheet };
}

async function saveEnrollment(data) {
  const { parentFirstName, parentLastName, phone, email, childName, childAge, program, message } = data;
  if (!parentFirstName || !parentLastName || !phone || !email || !childName || !childAge || !program) {
    throw new Error('Please fill in all required fields.');
  }
  const { workbook, sheet } = await getWorkbookAndSheet();
  sheet.addRow([new Date().toISOString(), parentFirstName, parentLastName, phone, email, childName, childAge, program, message || '']);
  await workbook.xlsx.writeFile(workbookPath);
}

app.post('/enroll', async (req, res) => {
  try {
    await saveEnrollment(req.body);
    res.json({ success: true, message: 'Enrollment saved!' });
  } catch (error) {
    console.error('Excel write error:', error);
    const status = error.message === 'Please fill in all required fields.' ? 400 : 500;
    let message = error.message || 'Failed to save enrollment';
    if (error.code === 'EBUSY' || error.code === 'EPERM') {
      message = 'Cannot write to enrollments.xlsx. Close the Excel file if it is open, then try again.';
    }
    res.status(status).json({ success: false, message });
  }
});

// ─── Payments Excel ───────────────────────────────────────────────────────────
const paymentsPath = path.join(__dirname, 'payments.xlsx');
const pendingPaymentsPath = path.join(__dirname, 'pending-payments.json');

// Column layout (1-indexed):
//  1: Submitted At  2: Parent Name  3: Child Name  4: Phone
//  5: Plan          6: Amount (₸)   7: Receipt Number
//  8: Payment Method (Kaspi / Cash)   ← NEW column
//  9: Payment Month  10: Status
const paymentHeaders = [
  'Submitted At',
  'Parent Name',
  'Child Name',
  'Phone',
  'Plan',
  'Amount (₸)',
  'Receipt Number',
  'Payment Method', // col 8 — "Kaspi" or "Cash"
  'Payment Month',  // col 9
  'Status'          // col 10
];

function isPaymentHeaderRow(rowValues) {
  const n = rowValues.map((v) => String(v || '').trim());
  return n[0] === 'Submitted At' && n[1] === 'Parent Name' && n[2] === 'Child Name' && n[6] === 'Receipt Number';
}

async function getPaymentsSheet() {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(paymentsPath);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }

  const cleanedRows = [];
  const seenRows = new Set();

  workbook.worksheets.forEach((worksheet) => {
    for (let i = 1; i <= worksheet.rowCount; i++) {
      const rowValues = worksheet.getRow(i).values.slice(1);
      const hasValues = rowValues.some((v) => v !== null && v !== undefined && v !== '');
      if (!hasValues || isPaymentHeaderRow(rowValues)) continue;

      // Migrate old rows (9 cols, no Payment Method) → new format (10 cols)
      let row;
      const col8 = String(rowValues[7] || '').trim();
      const isOldFormat = rowValues.length <= 9 && col8 !== 'Kaspi' && col8 !== 'Cash';
      if (isOldFormat) {
        row = [
          rowValues[0] || '',
          rowValues[1] || '',
          rowValues[2] || '',
          rowValues[3] || '',
          rowValues[4] || '',
          rowValues[5] || '',
          rowValues[6] || '',
          'Kaspi',           // backfill Payment Method
          rowValues[7] || '', // Payment Month
          rowValues[8] || 'Pending'
        ];
      } else {
        row = [
          rowValues[0] || '',
          rowValues[1] || '',
          rowValues[2] || '',
          rowValues[3] || '',
          rowValues[4] || '',
          rowValues[5] || '',
          rowValues[6] || '',
          rowValues[7] || 'Kaspi',
          rowValues[8] || '',
          rowValues[9] || 'Pending'
        ];
      }

      const key = row.map((v) => String(v ?? '')).join('|');
      if (!seenRows.has(key)) {
        seenRows.add(key);
        cleanedRows.push(row);
      }
    }
  });

  while (workbook.worksheets.length > 0) {
    workbook.removeWorksheet(workbook.worksheets[0].id);
  }
  const sheet = workbook.addWorksheet('Payments');
  sheet.getRow(1).values = paymentHeaders;
  cleanedRows.forEach((r, i) => { sheet.getRow(i + 2).values = r; });

  return { workbook, sheet };
}

async function readPendingPayments() {
  try {
    const raw = await fs.readFile(pendingPaymentsPath, 'utf8');
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    if (error.code === 'ENOENT') {
      return [];
    }
    throw error;
  }
}

async function writePendingPayments(items) {
  await fs.writeFile(pendingPaymentsPath, JSON.stringify(items, null, 2), 'utf8');
}

function getReceiptFromProviderPayload(payload, fallbackId) {
  return String(
    payload?.reference ||
    payload?.receipt_number ||
    payload?.receiptNumber ||
    payload?.transaction_id ||
    payload?.transactionId ||
    payload?.id ||
    fallbackId ||
    ''
  ).trim();
}

async function upsertPendingPayment(entry) {
  const items = await readPendingPayments();
  const receiptNumber = String(entry.receiptNumber || '').trim();
  const invoiceId = String(entry.invoiceId || '').trim();

  const idx = items.findIndex((item) => {
    const sameReceipt = receiptNumber && String(item.receiptNumber || '').trim() === receiptNumber;
    const sameInvoice = invoiceId && String(item.invoiceId || '').trim() === invoiceId;
    return sameReceipt || sameInvoice;
  });

  const normalized = {
    invoiceId: invoiceId || '',
    receiptNumber,
    amount: parseInt(entry.amount || 0),
    phone: entry.phone || '',
    plan: entry.plan || '',
    paidAt: entry.paidAt || new Date().toISOString(),
    status: entry.status || 'pending'
  };

  if (idx >= 0) {
    items[idx] = { ...items[idx], ...normalized };
  } else {
    items.push(normalized);
  }

  await writePendingPayments(items);
  return normalized;
}

async function removePendingPayment(receiptNumber) {
  const target = String(receiptNumber || '').trim();
  if (!target) return;
  const items = await readPendingPayments();
  const filtered = items.filter((item) => String(item.receiptNumber || '').trim() !== target);
  if (filtered.length !== items.length) {
    await writePendingPayments(filtered);
  }
}

// ─── Kaspi Pay via ApiPay.kz ──────────────────────────────────────────────────
//
//  HOW TO GET YOUR API KEY:
//  1. Go to https://apipay.kz and register
//  2. Contact them on WhatsApp: +7 708 516 74 89
//  3. They connect your Kaspi Business account as a "Cashier"
//  4. You get an API key — paste it below OR set env var APIPAY_KEY=your_key
//
const APIPAY_KEY = process.env.APIPAY_KEY || 'YOUR_APIPAY_KEY_HERE';
const APIPAY_BASE = 'https://bpapi.bazarbay.site/api/v1';

// ─── POST /payment/verify-receipt ─────────────────────────────────────────────
// Verifies a Kaspi receipt number against ApiPay BEFORE saving to Excel.
// Frontend must call this first; only if { verified: true } should it proceed
// to POST /payment.
app.post('/payment/verify-receipt', async (req, res) => {
  const { receiptNumber, expectedAmount } = req.body;

  if (!receiptNumber) {
    return res.status(400).json({ success: false, message: 'Receipt number is required.' });
  }

  if (!APIPAY_KEY || APIPAY_KEY === 'YOUR_APIPAY_KEY_HERE') {
    // No API key configured — block the submission so random numbers can't get through
    return res.status(503).json({
      success: false,
      message: 'Kaspi verification is not configured yet. Set APIPAY_KEY in your environment. See server.js comments for setup instructions.'
    });
  }

  try {
    // Look up the transaction by its Kaspi receipt / reference number
    const response = await fetch(`${APIPAY_BASE}/transactions?reference=${encodeURIComponent(receiptNumber)}`, {
      headers: { 'X-API-Key': APIPAY_KEY }
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('ApiPay verify error:', data);
      return res.status(502).json({ success: false, message: data.message || 'Could not verify receipt with Kaspi.' });
    }

    const transactions = Array.isArray(data) ? data : (data.items || data.data || []);
    const match = transactions.find(
      (t) => String(t.reference || t.receipt_number || t.id) === String(receiptNumber)
    );

    if (!match) {
      return res.status(404).json({ success: false, message: 'Receipt not found. Make sure the Kaspi payment was completed.' });
    }

    const txStatus = String(match.status || '').toLowerCase();
    if (txStatus && txStatus !== 'paid' && txStatus !== 'success' && txStatus !== 'completed') {
      return res.status(400).json({ success: false, message: `Payment status is "${match.status}" — not paid yet.` });
    }

    // Verify the paid amount matches the selected plan
    if (expectedAmount) {
      const paidAmount = parseInt(match.amount || match.sum || 0);
      const expected = parseInt(expectedAmount);
      if (paidAmount !== expected) {
        return res.status(400).json({
          success: false,
          message: `Amount mismatch: receipt shows ₸${paidAmount.toLocaleString()} but the selected plan costs ₸${expected.toLocaleString()}.`
        });
      }
    }

    res.json({
      success: true,
      verified: true,
      amount: match.amount || match.sum,
      paidAt: match.created_at || match.paid_at || new Date().toISOString()
    });

  } catch (err) {
    console.error('Receipt verification error:', err);
    res.status(500).json({ success: false, message: 'Server error during receipt verification.' });
  }
});

app.get('/payment/pending', async (req, res) => {
  try {
    const items = await readPendingPayments();
    const pending = items
      .filter((item) => String(item.status || '').trim() === 'pending')
      .sort((a, b) => new Date(b.paidAt || 0).getTime() - new Date(a.paidAt || 0).getTime());

    res.json({ success: true, pending });
  } catch (error) {
    console.error('Pending payments read error:', error);
    res.status(500).json({ success: false, pending: [], message: 'Failed to load pending payments.' });
  }
});

// ─── POST /payment ─────────────────────────────────────────────────────────────
// Saves a payment to payments.xlsx.
//   Kaspi: receipt must be verified via /payment/verify-receipt first (frontend enforces this).
//   Cash:  no receipt needed — saved immediately with status "Cash - Confirmed".
async function savePayment(data) {
  const { parentName, childName, phone, plan, amount, receiptNumber, paymentMonth, paymentMethod } = data;

  if (!parentName || !childName || !phone || !plan || !amount || !paymentMonth || !paymentMethod) {
    throw new Error('Please fill in all required fields.');
  }

  if (paymentMethod !== 'Kaspi' && paymentMethod !== 'Cash') {
    throw new Error('Payment method must be "Kaspi" or "Cash".');
  }

  if (paymentMethod === 'Kaspi' && !receiptNumber) {
    throw new Error('Receipt number is required for Kaspi payments.');
  }

  const { workbook, sheet } = await getPaymentsSheet();

  // Prevent duplicate Kaspi receipt numbers
  if (paymentMethod === 'Kaspi') {
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
        const existing = String(row.getCell(7).value || '');
        if (existing && existing === String(receiptNumber)) {
          throw new Error(`Receipt number ${receiptNumber} has already been submitted.`);
        }
      }
    });
  }

  // Status:
  //   Cash  → "Cash - Confirmed" (staff physically accepted the money)
  //   Kaspi → "Confirmed" (receipt was verified against ApiPay before this call)
  const status = paymentMethod === 'Cash' ? 'Cash - Confirmed' : 'Confirmed';

  sheet.addRow([
    new Date().toISOString(), // col 1 — Submitted At
    parentName,               // col 2
    childName,                // col 3
    phone,                    // col 4
    plan,                     // col 5
    parseInt(amount),         // col 6 — Amount (₸)
    receiptNumber || 'N/A',  // col 7 — Receipt Number
    paymentMethod,            // col 8 — Payment Method
    paymentMonth,             // col 9 — Payment Month
    status                    // col 10 — Status
  ]);

  await workbook.xlsx.writeFile(paymentsPath);
}

app.post('/payment', async (req, res) => {
  try {
    await savePayment(req.body);
    res.json({ success: true, message: 'Payment saved!' });
  } catch (error) {
    console.error('Payment save error:', error);
    const status = error.message.includes('required') || error.message.includes('already') || error.message.includes('must be') ? 400 : 500;
    let message = error.message || 'Failed to save payment.';
    if (error.code === 'EBUSY' || error.code === 'EPERM') {
      message = 'Cannot write to payments.xlsx. Close the file if it is open in Excel, then try again.';
    }
    res.status(status).json({ success: false, message });
  }
});

app.post('/payment/confirm', async (req, res) => {
  try {
    await savePayment(req.body);
    if (req.body.paymentMethod === 'Kaspi') {
      await removePendingPayment(req.body.receiptNumber);
    }
    res.json({ success: true, message: 'Payment confirmed and saved!' });
  } catch (error) {
    console.error('Payment confirm error:', error);
    const status = error.message.includes('required') || error.message.includes('already') || error.message.includes('must be') ? 400 : 500;
    let message = error.message || 'Failed to confirm payment.';
    if (error.code === 'EBUSY' || error.code === 'EPERM') {
      message = 'Cannot write to payments.xlsx. Close the file if it is open in Excel, then try again.';
    }
    res.status(status).json({ success: false, message });
  }
});

// POST /kaspi/create-invoice
app.post('/kaspi/create-invoice', async (req, res) => {
  const { phone, amount, plan, description } = req.body;
  if (!phone || !amount || !plan) {
    return res.status(400).json({ success: false, message: 'Missing required fields.' });
  }
  if (!APIPAY_KEY || APIPAY_KEY === 'YOUR_APIPAY_KEY_HERE') {
    return res.status(503).json({
      success: false,
      message: 'Kaspi invoice creation is not configured yet. Set APIPAY_KEY in your environment.'
    });
  }
  try {
    const response = await fetch(`${APIPAY_BASE}/invoices`, {
      method: 'POST',
      headers: { 'X-API-Key': APIPAY_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        amount: parseInt(amount),
        phone_number: phone,
        description: description || `Bobo Kindergarten — ${plan}`,
        external_order_id: `bobo_${Date.now()}`
      })
    });
    const invoice = await response.json();
    if (!response.ok) {
      return res.status(502).json({ success: false, message: invoice.message || 'Failed to create Kaspi invoice.' });
    }

    await upsertPendingPayment({
      invoiceId: invoice.id,
      receiptNumber: '',
      amount,
      phone,
      plan,
      paidAt: new Date().toISOString(),
      status: 'awaiting_payment'
    });

    res.json({ success: true, invoiceId: invoice.id, qrUrl: invoice.qr_url || invoice.payment_url, status: invoice.status });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Server error. Try again.' });
  }
});

// GET /kaspi/status/:id
app.get('/kaspi/status/:id', async (req, res) => {
  if (!APIPAY_KEY || APIPAY_KEY === 'YOUR_APIPAY_KEY_HERE') {
    return res.status(503).json({
      status: 'error',
      message: 'Kaspi status check is not configured yet. Set APIPAY_KEY in your environment.'
    });
  }
  try {
    const response = await fetch(`${APIPAY_BASE}/invoices/${req.params.id}`, {
      headers: { 'X-API-Key': APIPAY_KEY }
    });
    const invoice = await response.json();
    if (!response.ok) return res.status(502).json({ status: 'error', message: invoice.message });
    if (invoice.status === 'paid') {
      const receiptNumber = getReceiptFromProviderPayload(invoice, req.params.id);
      const paidAt = invoice.created_at || invoice.paid_at || new Date().toISOString();
      const amount = invoice.amount || invoice.sum || 0;
      const phone = invoice.phone_number || invoice.phone || '';
      const plan = invoice.description || 'Kaspi Payment';
      await upsertPendingPayment({
        invoiceId: req.params.id,
        receiptNumber,
        amount,
        phone,
        plan,
        paidAt,
        status: 'pending'
      });

      try {
        const { workbook, sheet } = await getPaymentsSheet();
        let updated = false;
        sheet.eachRow((row, rowNum) => {
          if (rowNum > 1 && String(row.getCell(7).value) === receiptNumber) {
            row.getCell(10).value = 'Confirmed'; // col 10 = Status
            row.getCell(9).value = new Date().toLocaleDateString('en-KZ', { month: 'long', year: 'numeric' });
            updated = true;
          }
        });
        if (updated) await workbook.xlsx.writeFile(paymentsPath);
      } catch (excelErr) {
        console.warn('Excel update warning:', excelErr.message);
      }

      return res.json({
        status: invoice.status,
        receiptNumber,
        amount,
        phone,
        plan,
        paidAt
      });
    }
    res.json({ status: invoice.status });
  } catch (err) {
    res.status(500).json({ status: 'error' });
  }
});

// POST /kaspi/cancel/:id
app.post('/kaspi/cancel/:id', async (req, res) => {
  try {
    await fetch(`${APIPAY_BASE}/invoices/${req.params.id}/cancel`, {
      method: 'POST',
      headers: { 'X-API-Key': APIPAY_KEY }
    });
    res.json({ success: true });
  } catch (err) {
    res.json({ success: false });
  }
});

// ─── Server start ─────────────────────────────────────────────────────────────
function startServer() {
  return app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
  });
}

if (require.main === module) {
  startServer();
}
module.exports = { app, saveEnrollment, savePayment, startServer };
