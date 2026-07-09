/**
 * Stale offer controls for the Reflex-AI Training Scheduler.
 *
 * Purpose:
 * - Prevent old pending/offered rows from prior months from blocking new sims.
 * - Add a cutoff date setting: OFFER_IGNORE_BEFORE_DATE, for example 2026-07-01.
 * - Add a menu-callable cleanup function: expireOldPendingOffers.
 * - Maintain the original offer Status while marking old rows Inactive.
 *
 * IMPORTANT:
 * This file includes a replacement tsHasActivePendingOffer_ function. In Apps
 * Script, replace the existing scheduler version with this one, or ensure this
 * definition is the only function with that name in the project.
 */

function expireOldPendingOffers() {
  var cutoff = tsGetOfferIgnoreBeforeDate_();
  if (!cutoff) {
    SpreadsheetApp.getUi().alert(
      'Expire Old Pending Offers',
      'Set OFFER_IGNORE_BEFORE_DATE in TS Config first, for example 2026-07-01.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var ss = tsGetSpreadsheet_();
  var offerSheet = tsEnsureOffersSheet_(ss);
  var tokenSheet = ss.getSheetByName(TS.SHEETS.BOOKING_TOKENS);

  if (!offerSheet || offerSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert(
      'Expire Old Pending Offers',
      'No TS Offers rows found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var activeColumn = tsEnsureOfferActiveColumn_(offerSheet);
  var values = offerSheet.getRange(1, 1, offerSheet.getLastRow(), Math.max(offerSheet.getLastColumn(), activeColumn)).getValues();
  var scanned = 0;
  var beforeCutoff = 0;
  var activatedRows = 0;
  var inactivatedRows = 0;
  var pendingTokensStaled = 0;
  var beforeCutoff = 0;
  var unparsableCreatedAt = 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 1;
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    var createdAt = row[TS.OFFER_COLS.CREATED_AT - 1];
    scanned++;

    var parsedCreatedAt = tsParseOfferDateValue_(createdAt);
    if (!parsedCreatedAt) {
      unparsableCreatedAt++;
      continue;
    }

    if (!tsDateIsBeforeCutoff_(parsedCreatedAt, cutoff)) {
      if (String(row[activeColumn - 1] || '').trim() !== 'Active') {
        offerSheet.getRange(rowNum, activeColumn).setValue('Active');
        activatedRows++;
      }
      continue;
    }
    beforeCutoff++;

    if (String(row[activeColumn - 1] || '').trim() !== 'Inactive') {
      offerSheet.getRange(rowNum, activeColumn).setValue('Inactive');
      inactivatedRows++;
    }

    // Keep the original offer Status intact for audit/history, but disable old
    // pending booking tokens so stale links cannot still be booked.
    var tokenIds = String(row[TS.OFFER_COLS.TOKEN_IDS - 1] || '')
      .split(',')
      .map(function(value) { return value.trim(); })
      .filter(Boolean);

    tokenIds.forEach(function(tokenId) {
      var hit = tsLookupBookingTokenRow_(tokenId);
      if (!hit) return;

      var tokenStatus = String(hit.row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
      if (tokenStatus !== 'PENDING') return;

      tsUpdateBookingTokenStatus_(hit.rowNum, 'STALE');
      pendingTokensStaled++;
    });
  }

  SpreadsheetApp.flush();
  tsAudit_(
    'EXPIRE_OLD_OFFERS',
    '',
    'Cutoff=' + Utilities.formatDate(cutoff, TS.TZ, 'yyyy-MM-dd') +
      '; scanned=' + scanned +
      '; beforeCutoff=' + beforeCutoff +
      '; inactivatedRows=' + inactivatedRows +
      '; activatedRows=' + activatedRows +
      '; pendingTokensStaled=' + pendingTokensStaled +
      '; unparsableCreatedAt=' + unparsableCreatedAt,
    'OK'
  );

  SpreadsheetApp.getUi().alert(
    'Expire Old Pending Offers',
    'Old pending/offered offers expired.\n\n' +
      'Cutoff date: ' + Utilities.formatDate(cutoff, TS.TZ, 'yyyy-MM-dd') + '\n' +
      'Rows scanned: ' + scanned + '\n' +
      'Rows before cutoff marked Inactive: ' + inactivatedRows + '\n' +
      'Rows on/after cutoff marked Active: ' + activatedRows + '\n' +
      'Old pending tokens marked STALE: ' + pendingTokensStaled + '\n\n' +
      (unparsableCreatedAt ? 'Rows with unreadable Created At: ' + unparsableCreatedAt + '\n\n' : '') +
      'Run offers again to create new offers for current sims.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function expireoldpendingoffers() {
  expireOldPendingOffers();
}

/**
 * Replacement for the scheduler's existing tsHasActivePendingOffer_.
 *
 * Old behavior:
 * - Any active pending offer/token could block new offers.
 *
 * New behavior:
 * - Active BOOKED and ESCALATED rows still block.
 * - Inactive rows do not block, regardless of original Status.
 * - STALE/CANCELLED/EXPIRED/SUPERSEDED rows do not block.
 * - Rows before OFFER_IGNORE_BEFORE_DATE do not block.
 * - Still-valid pending tokens after the cutoff do block.
 */
function tsHasActivePendingOffer_(email, salesGroup) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var activeColumn = tsGetOfferActiveColumn_(sheet);
  var targetEmail = String(email || '').trim().toLowerCase();
  var targetGroup = String(salesGroup || '').trim().toLowerCase();
  var cutoff = tsGetOfferIgnoreBeforeDate_();
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowEmail = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var rowGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    var createdAt = row[TS.OFFER_COLS.CREATED_AT - 1];

    if (rowEmail !== targetEmail || rowGroup !== targetGroup) continue;
    if (activeColumn && tsOfferRowIsInactive_(row, activeColumn)) continue;
    if (cutoff && tsDateIsBeforeCutoff_(createdAt, cutoff)) continue;

    if (status === 'BOOKED' || status === 'ESCALATED') return true;
    if (status === 'STALE' || status === 'CANCELLED' || status === 'EXPIRED' || status === 'SUPERSEDED') continue;
    if (status !== 'OFFERED' && status !== 'PENDING') continue;

    var tokenIds = String(row[TS.OFFER_COLS.TOKEN_IDS - 1] || '')
      .split(',')
      .map(function(value) { return value.trim(); })
      .filter(Boolean);

    if (!tokenIds.length) continue;

    var anyPending = tokenIds.some(function(tokenId) {
      var hit = tsLookupBookingTokenRow_(tokenId);
      if (!hit) return false;

      var tokenStatus = String(hit.row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
      if (tokenStatus !== 'PENDING') return false;

      var tokenCreated = hit.row[TS.BOOK_COLS.CREATED_AT - 1];
      if (cutoff && tsDateIsBeforeCutoff_(tokenCreated, cutoff)) return false;

      return true;
    });

    if (anyPending) return true;
  }

  return false;
}

function tsGetOfferIgnoreBeforeDate_() {
  var raw = tsReadOfferIgnoreBeforeDateFromConfigSheet_();

  if (!raw) {
    try {
      raw = PropertiesService.getScriptProperties().getProperty('OFFER_IGNORE_BEFORE_DATE');
    } catch (ignore) {}
  }

  if (!raw) return null;

  var parsed = tsParseOfferDateValue_(raw);
  if (!parsed || isNaN(parsed.getTime())) {
    throw new Error('Invalid OFFER_IGNORE_BEFORE_DATE. Use yyyy-mm-dd, for example 2026-07-01.');
  }

  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

function tsReadOfferIgnoreBeforeDateFromConfigSheet_() {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return '';

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '').trim().toUpperCase();
    if (!key) continue;
    if (key === 'MANAGER NAME') break;

    if (
      key === 'OFFER_IGNORE_BEFORE_DATE' ||
      key === 'OFFER_IGNORE' ||
      key === 'OFFER_IGNOR' ||
      key.indexOf('OFFER_IGNORE_BEFORE') === 0
    ) {
      return values[i][1];
    }
  }

  return '';
}

function tsDateIsBeforeCutoff_(value, cutoff) {
  if (!cutoff) return false;
  var parsed = tsParseOfferDateValue_(value);
  if (!parsed) return false;
  return parsed.getTime() < cutoff.getTime();
}

function tsIsOfferStatusStaleEligible_(status) {
  return ['PENDING', 'OFFERED'].indexOf(String(status || '').trim().toUpperCase()) !== -1;
}

function tsEnsureOfferActiveColumn_(sheet) {
  var existingColumn = tsGetOfferActiveColumn_(sheet);
  if (existingColumn) return existingColumn;

  var newColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumn).setValue('Active or Inactive');
  sheet.getRange(1, newColumn).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#ffffff');
  return newColumn;
}

function tsGetOfferActiveColumn_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i] || '').trim().toLowerCase();
    if (header === 'active or inactive' || header === 'active/inactive' || header === 'active') {
      return i + 1;
    }
  }
  return 0;
}

function tsOfferRowIsInactive_(row, activeColumn) {
  return String(row[activeColumn - 1] || '').trim().toLowerCase() === 'inactive';
}

function tsParseOfferDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === 'number' && isFinite(value)) {
    // Google Sheets serial date number.
    return new Date(Math.round((value - 25569) * 86400 * 1000));
  }

  var raw = String(value || '').trim();
  if (!raw) return null;

  var direct = new Date(raw);
  if (!isNaN(direct.getTime())) {
    return direct;
  }

  var match = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    var month = Number(match[1]);
    var day = Number(match[2]);
    var year = Number(match[3]);
    var hour = Number(match[4] || 0);
    var minute = Number(match[5] || 0);
    var second = Number(match[6] || 0);
    var parsed = new Date(year, month - 1, day, hour, minute, second);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  return null;
}
