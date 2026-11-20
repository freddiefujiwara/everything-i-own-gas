import { FLAG_VALUE, checkedRow_ } from './admin.js';
import { getSheet } from './common.js';

// --- Subscriptions ---
// Columns: Service, Purpose, Yearly Cost, Monthly Cost, Ratio, URL.
// The sheet's own convention is kept: exactly one of the cost cells
// holds the entered amount and the other derives it by formula
// (monthly = yearly / 12, or yearly = monthly * 12). The row's basis
// says which side is entered. The Ratio cell is a formula against the
// totals row's monthly sum; structural inserts/deletes shift it like
// the sheet does, so the admin only writes it for new rows. The totals
// row (the one with =SUM(...) in a cost column) is never editable,
// and new rows are inserted just above it so the SUM ranges and the
// ratio references keep covering every subscription.
const SUBSCRIPTION_SHEET = 'subscriptions';

const findTotalsRow_ = sheet => {
  const formulas = sheet.getDataRange().getFormulas();
  for (let i = 1; i < formulas.length; i++) {
    if (/^=SUM\(/i.test(formulas[i][2] || '') || /^=SUM\(/i.test(formulas[i][3] || '')) {
      return i + 1;
    }
  }
  return 0;
};

const subscriptionBasis_ = (cFormula, dFormula) => {
  if (/^=C\d+\/12/.test(dFormula || '')) return 'yearly';
  if (/^=D\d+\*12/.test(cFormula || '')) return 'monthly';
  return 'yearly';
};

const readSubscriptions_ = () => {
  const sheet = getSheet(SUBSCRIPTION_SHEET);
  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();
  const totalsRow = findTotalsRow_(sheet);
  return values
    .slice(1)
    .map((row, index) => ({
      row: index + 2,
      service: row[0] ?? '',
      purpose: row[1] ?? '',
      yearly: row[2] ?? '',
      monthly: row[3] ?? '',
      ratio: row[4] ?? '',
      url: row[5] ?? '',
      flag: row[6] ?? '',
      basis: subscriptionBasis_(formulas[index + 1][2], formulas[index + 1][3]),
    }))
    .filter(sub => sub.row !== totalsRow);
};

// Build the A:F cells for one subscription at a given row. amounts:
// sub.amount is the entered cost on the sub.basis side; the other
// side becomes the deriving formula, like the sheet's own rows.
const subscriptionCells_ = (sub, row, ratioFormula) => {
  const src = sub || {};
  // Text reads totally, like the item rows: a value whose String()
  // throws reads as blank instead of failing the whole mutation.
  const text = v => {
    if (v === null || v === undefined) return '';
    try {
      return String(v);
    } catch (e) {
      return '';
    }
  };
  // The amount is the entered cost: a finite number, or a string
  // that parses as one. Anything else (a boolean, an array, an
  // unstringifiable object) reads as no amount, like a blank cell —
  // a bare Number() would store a boolean as 1 and an array as its
  // single element.
  const amount =
    typeof src.amount === 'number'
      ? isFinite(src.amount)
        ? src.amount
        : ''
      : typeof src.amount === 'string' &&
          src.amount.trim() !== '' &&
          isFinite(Number(src.amount))
        ? Number(src.amount)
        : '';
  let yearly = '';
  let monthly = '';
  if (amount !== '') {
    if (src.basis === 'monthly') {
      monthly = amount;
      yearly = `=D${row}*12`;
    } else {
      yearly = amount;
      monthly = `=C${row}/12`;
    }
  }
  return [text(src.service), text(src.purpose), yearly, monthly, ratioFormula, text(src.url), text(src.flag)];
};

const createSubscription_ = sub => {
  const sheet = getSheet(SUBSCRIPTION_SHEET);
  const totalsRow = findTotalsRow_(sheet);
  let row;
  let newTotalsRow = totalsRow;
  if (totalsRow > 0) {
    sheet.insertRowsBefore(totalsRow);
    row = totalsRow;
    newTotalsRow = totalsRow + 1;
  } else {
    row = sheet.getLastRow() + 1;
  }
  const cells = subscriptionCells_(sub, row, '');
  if ((cells[2] !== '' || cells[3] !== '') && newTotalsRow > 0) {
    cells[4] = `=D${row}/$D$${newTotalsRow}`;
  }
  sheet.getRange(row, 1, 1, 7).setValues([cells]);
  return { row };
};

const checkedSubscriptionRow_ = (sheet, row) => {
  const n = checkedRow_(sheet, row);
  if (n === findTotalsRow_(sheet)) {
    throw new Error('the totals row is not editable');
  }
  return n;
};

const updateSubscription_ = (row, sub) => {
  const sheet = getSheet(SUBSCRIPTION_SHEET);
  const n = checkedSubscriptionRow_(sheet, row);
  const cells = subscriptionCells_(sub, n, '');
  // The Ratio cell (E) is left untouched: its formula keeps pointing
  // at the totals row, exactly as the sheet maintains it.
  sheet.getRange(n, 1, 1, 4).setValues([cells.slice(0, 4)]);
  sheet.getRange(n, 6).setValue(cells[5]);
  // The flag (G) follows the same rule as the item sheets: set_flag
  // owns it, and a plain edit that omits it keeps the stored value.
  const flag =
    sub && sub.flag !== undefined
      ? cells[6]
      : ((sheet.getRange(n, 7).getValues()[0] || [])[0] ?? '');
  sheet.getRange(n, 7).setValue(flag);
  return { row: n };
};

const setSubscriptionFlag_ = (row, flagged) => {
  const sheet = getSheet(SUBSCRIPTION_SHEET);
  const n = checkedSubscriptionRow_(sheet, row);
  const flag = flagged ? FLAG_VALUE : '';
  sheet.getRange(n, 7).setValue(flag);
  return { row: n, flag };
};

const deleteSubscription_ = row => {
  const sheet = getSheet(SUBSCRIPTION_SHEET);
  const n = checkedSubscriptionRow_(sheet, row);
  sheet.deleteRow(n);
  return { row: n };
};

export { SUBSCRIPTION_SHEET, createSubscription_, deleteSubscription_, readSubscriptions_, setSubscriptionFlag_, subscriptionCells_, updateSubscription_ };
