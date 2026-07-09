/**
 * Stale offer controls for the Reflex-AI Training Scheduler.
 *
 * Purpose:
 * - Prevent old pending/offered rows from prior months from blocking new sims.
 * - Add a cutoff date setting: OFFER_IGNORE_BEFORE_DATE, for example 2026-07-01.
 * - Add a menu-callable cleanup function: expireOldPendingOffers.
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

  var values = offerSheet.getDataRange().getValues();
  var staleOffers = 0;
  var staleTokens = 0;

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowNum = i + 1;
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    var createdAt = row[TS.OFFER_COLS.CREATED_AT - 1];

    if (!tsIsOfferStatusStaleEligible_(status)) {
      continue;
    }

    if (!tsDateIsBeforeCutoff_(createdAt, cutoff)) {
      continue;
    }

    offerSheet.getRange(rowNum, TS.OFFER_COLS.STATUS).setValue('STALE');
    staleOffers++;

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
      staleTokens++;
    });
  }

  SpreadsheetApp.flush();
  tsAudit_(
    'EXPIRE_OLD_OFFERS',
    '',
    'Cutoff=' + Utilities.formatDate(cutoff, TS.TZ, 'yyyy-MM-dd') +
      '; staleOffers=' + staleOffers +
      '; staleTokens=' + staleTokens,
    'OK'
  );

  SpreadsheetApp.getUi().alert(
    'Expire Old Pending Offers',
    'Old pending/offered offers expired.\n\n' +
      'Cutoff date: ' + Utilities.formatDate(cutoff, TS.TZ, 'yyyy-MM-dd') + '\n' +
      'Offers marked STALE: ' + staleOffers + '\n' +
      'Tokens marked STALE: ' + staleTokens + '\n\n' +
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
 * - BOOKED and ESCALATED still block.
 * - STALE/CANCELLED/EXPIRED/SUPERSEDED do not block.
 * - OFFERED/PENDING rows before OFFER_IGNORE_BEFORE_DATE do not block.
 * - Still-valid pending tokens after the cutoff do block.
 */
function tsHasActivePendingOffer_(email, salesGroup) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var targetEmail = String(email || '').trim().toLowerCase();
  var targetGroup = String(salesGroup || '').trim().toLowerCase();
  var cutoff = tsGetOfferIgnoreBeforeDate_();
  var now = new Date().getTime();
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowEmail = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var rowGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();
    var createdAt = row[TS.OFFER_COLS.CREATED_AT - 1];

    if (rowEmail !== targetEmail || rowGroup !== targetGroup) continue;

    if (status === 'BOOKED' || status === 'ESCALATED') return true;
    if (status === 'STALE' || status === 'CANCELLED' || status === 'EXPIRED' || status === 'SUPERSEDED') continue;
    if (cutoff && tsDateIsBeforeCutoff_(createdAt, cutoff)) continue;
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

      if (tokenCreated instanceof Date && !isNaN(tokenCreated.getTime())) {
        return now - tokenCreated.getTime() < TS.BOOKING_TOKEN_TTL_MS;
      }

      return true;
    });

    if (anyPending) return true;
  }

  return false;
}

function tsGetOfferIgnoreBeforeDate_() {
  var raw = '';

  try {
    raw = String(tsLoadConfig_().OFFER_IGNORE_BEFORE_DATE || '').trim();
  } catch (ignore) {}

  if (!raw) {
    raw = String(PropertiesService.getScriptProperties().getProperty('OFFER_IGNORE_BEFORE_DATE') || '').trim();
  }

  if (!raw) return null;

  var parsed = Utilities.parseDate(raw.substring(0, 10) + 'T00:00:00', TS.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  if (!parsed || isNaN(parsed.getTime())) {
    throw new Error('Invalid OFFER_IGNORE_BEFORE_DATE. Use yyyy-mm-dd, for example 2026-07-01.');
  }

  return parsed;
}

function tsDateIsBeforeCutoff_(value, cutoff) {
  if (!cutoff) return false;
  if (!(value instanceof Date) || isNaN(value.getTime())) return false;
  return value.getTime() < cutoff.getTime();
}

function tsIsOfferStatusStaleEligible_(status) {
  return ['PENDING', 'OFFERED'].indexOf(String(status || '').trim().toUpperCase()) !== -1;
}
