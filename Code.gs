/**
 * OHL Test Catalog — Google Apps Script Web App
 *
 * SETUP (one-time):
 *  1. Open the Google Sheet → Extensions → Apps Script
 *  2. Replace the default code with this entire file
 *  3. Click Deploy → New Deployment
 *     - Type: Web app
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  4. Click Deploy → copy the Web app URL
 *  5. Paste that URL into index.html as the value of SCRIPT_URL
 *  6. Commit & push index.html — done!
 *
 * To refresh after a sheet change, just reload the dashboard page.
 * To force re-deploy (e.g. after editing this script), click
 * Deploy → Manage Deployments → edit → new version.
 */

const SHEET_ID = '1FPI-QWoy3x6eATQDkDkoJ_EqW8E7MOTil4PTAT9ZQbM';

function doGet() {
  const data = aggregateSheetData();
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function aggregateSheetData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const rows  = sheet.getDataRange().getValues();

  const map   = {};   // key → item object
  const order = [];   // insertion order of keys

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const type    = String(row[0] || '').trim();
    const name    = String(row[1] || '').trim();
    const subtest = String(row[2] || '').trim();
    const param   = String(row[3] || '').trim();
    const city    = String(row[4] || '').trim();
    const price   = Number(row[5]) || 0;
    const tat     = String(row[6] || '').trim();
    const fasting = String(row[7] || 'No').trim();
    const gender  = String(row[8] || 'both').trim();

    if (!type || !name || !city) continue;

    const key = type + '||' + name;
    if (!map[key]) {
      map[key] = {
        type, name,
        fast: fasting,
        gender,
        cities: {},
        _params: [],
        _subsMap: {},
        _subsOrder: []
      };
      order.push(key);
    }

    const item = map[key];

    // First price seen for a city wins
    if (!item.cities[city] && price > 0) {
      item.cities[city] = { p: price, t: tat };
    }

    if (type === 'Test') {
      if (param && !item._params.includes(param)) item._params.push(param);
    } else {
      // Package: group parameters under their sub-test
      if (subtest && !item._subsMap[subtest]) {
        item._subsMap[subtest] = [];
        item._subsOrder.push(subtest);
      }
      if (subtest && param && !item._subsMap[subtest].includes(param)) {
        item._subsMap[subtest].push(param);
      }
    }
  }

  return order.map(key => {
    const item = map[key];
    const base = {
      type: item.type,
      name: item.name,
      fast: item.fast,
      gender: item.gender,
      cities: item.cities
    };
    if (item.type === 'Package') {
      return Object.assign({}, base, {
        subs: item._subsOrder.map(n => ({ name: n, params: item._subsMap[n] }))
      });
    }
    return Object.assign({}, base, { params: item._params });
  });
}
