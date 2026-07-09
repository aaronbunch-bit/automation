/**
 * Reoffer support for Book It links that fail because staffing changed.
 *
 * Adds support for:
 * - A "Send me more options" web page button.
 * - A ?reoffer=<booking_token> web-app action.
 * - Reusing the original booking token row to generate a fresh 3-slot offer.
 *
 * REQUIRED MAIN SCHEDULER EDITS:
 *
 * 1) In doGet(e), add this after the existing escalate block and before token:
 *
 *    var reofferToken = String(params.reoffer || '').trim();
 *    if (reofferToken) {
 *      var reofferResult = tsExecuteReofferToken_(reofferToken);
 *      return HtmlService.createHtmlOutput(tsBookingResponsePage_(reofferResult.message, reofferResult.ok))
 *        .setTitle(TS.BRAND);
 *    }
 *
 * 2) In tsExecuteBookItToken_(token), replace the capacity-fail return:
 *
 *    return { ok: false, message: 'Staffing changed — this slot is no longer available. Watch Slack for new options.' };
 *
 *    with:
 *
 *    return tsCapacityChangedReofferResult_(token);
 */

function tsCapacityChangedReofferResult_(bookingToken) {
  return {
    ok: false,
    message: 'Staffing changed — this slot is no longer available.\n\n' +
      'Click this link to request three new training options:\n' +
      tsBuildReofferUrl_(bookingToken)
  };
}

function tsBuildReofferUrl_(bookingToken) {
  var base = tsGetWebAppUrl_();
  if (!base || !bookingToken) return '';
  return base + (base.indexOf('?') === -1 ? '?' : '&') +
    'reoffer=' + encodeURIComponent(bookingToken);
}

function tsExecuteReofferToken_(bookingToken) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: 'Could not acquire lock. Try again in a moment.' };
  }

  try {
    var hit = tsLookupBookingTokenRow_(bookingToken);
    if (!hit) {
      return { ok: false, message: 'Invalid or unknown booking link.' };
    }

    var tokenRow = hit.row;
    var consultantName = String(tokenRow[TS.BOOK_COLS.CONSULTANT_NAME - 1] || '').trim();
    var consultantEmail = String(tokenRow[TS.BOOK_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var managerName = String(tokenRow[TS.BOOK_COLS.MANAGER_NAME - 1] || '').trim();
    var salesGroup = String(tokenRow[TS.BOOK_COLS.SALES_GROUP - 1] || '').trim();
    var queue = String(tokenRow[TS.BOOK_COLS.QUEUE - 1] || '').trim();
    var simsCsv = String(tokenRow[TS.BOOK_COLS.SIMS_CSV - 1] || '').trim();
    var durationMin = Number(tokenRow[TS.BOOK_COLS.DURATION_MIN - 1] || 30);
    var oldOfferRow = Number(tokenRow[TS.BOOK_COLS.OFFER_ROW - 1] || 0) || 0;
    var sims = simsCsv.split(',').map(function(value) { return value.trim(); }).filter(Boolean);

    if (!consultantEmail || !queue || !salesGroup) {
      return { ok: false, message: 'This booking link is missing consultant or queue data. Contact WFM.' };
    }

    var config = tsLoadConfig_();
    var apiKey = tsGetApiKey_();
    var headers = tsAuthHeaders_(apiKey);
    var siteId = tsResolveSiteId_(headers, TS.ASSEMBLED.SITE_NAME);
    var queueId = tsResolveQueueId_(headers, queue);

    if (!queueId) {
      return { ok: false, message: 'Could not resolve Assembled queue. Contact WFM.' };
    }

    var durationMs = durationMin * 60 * 1000;
    var searchRange = tsComputeSearchRange_(config);
    var scheduleIdx = tsPullPhoneScheduleIndex_(headers, searchRange.start, searchRange.end);
    var forecastCache = tsBuildForecastCache_(headers, siteId, queueId, config);
    var result = tsFindTrainingWindows_(headers, siteId, queueId, consultantEmail, durationMs, config, {
      forecastCache: forecastCache,
      scheduleIdx: scheduleIdx
    });

    if (!result.windows.length) {
      tsAudit_('REOFFER', consultantEmail, 'No replacement windows found — ' + result.diag, 'WARN');
      tsNotifyManagerNoReofferCapacity_(managerName, consultantName, consultantEmail, durationMin, simsCsv);
      return {
        ok: false,
        message: 'No replacement training windows are currently available. Please contact your manager or WFM.'
      };
    }

    tsSupersedeOldOfferForReoffer_(oldOfferRow);

    var need = {
      name: consultantName,
      email: consultantEmail,
      salesGroup: salesGroup,
      queue: queue,
      managerName: managerName,
      incompleteCount: sims.length || 1,
      sims: sims,
      durationMin: durationMin
    };

    var offerRow = tsAppendOfferRow_(need);
    var tokens = [];
    var windows = result.windows.slice(0, 3);

    windows.forEach(function(window) {
      var newToken = tsCreateBookingToken_({
        consultantName: need.name,
        consultantEmail: need.email,
        managerName: need.managerName,
        salesGroup: need.salesGroup,
        queue: need.queue,
        simsCsv: need.sims.join(', '),
        durationMin: need.durationMin,
        dateStr: window.dateStr,
        startStr: window.startStr,
        endStr: window.endStr,
        offerRow: offerRow
      });
      tokens.push(newToken);
    });

    var offerSheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
    offerSheet.getRange(offerRow, TS.OFFER_COLS.TOKEN_IDS).setValue(tokens.join(','));
    offerSheet.getRange(offerRow, TS.OFFER_COLS.OFFER_SENT_AT).setValue(new Date());
    offerSheet.getRange(offerRow, TS.OFFER_COLS.STATUS).setValue('OFFERED');
    SpreadsheetApp.flush();

    var message = tsBuildOfferSlackMessage_(need, windows, tokens);
    tsSlackDmConsultant_(need.name, need.email, message);
    tsAudit_('REOFFER', consultantEmail, 'Sent replacement offer with ' + tokens.length + ' slot(s)', 'OK');

    return {
      ok: true,
      message: 'The option you chose is no longer available. New options were just sent to you in Slack. Please reassess the updated time blocks and select a new offer.'
    };
  } catch (err) {
    tsAudit_('REOFFER', String(bookingToken || ''), String(err), 'FAILED');
    return {
      ok: false,
      message: 'Something went wrong while requesting new options. Please contact your manager or WFM.'
    };
  } finally {
    lock.releaseLock();
  }
}

function tsSupersedeOldOfferForReoffer_(offerRow) {
  if (!offerRow || offerRow < 2) return;

  var offerSheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!offerSheet) return;

  var row = offerSheet.getRange(offerRow, 1, 1, offerSheet.getLastColumn()).getValues()[0];
  var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

  if (status !== 'BOOKED' && status !== 'ESCALATED') {
    offerSheet.getRange(offerRow, TS.OFFER_COLS.STATUS).setValue('SUPERSEDED');
  }

  var tokenIds = String(row[TS.OFFER_COLS.TOKEN_IDS - 1] || '')
    .split(',')
    .map(function(value) { return value.trim(); })
    .filter(Boolean);

  tokenIds.forEach(function(tokenId) {
    var hit = tsLookupBookingTokenRow_(tokenId);
    if (!hit) return;

    var tokenStatus = String(hit.row[TS.BOOK_COLS.STATUS - 1] || '').trim().toUpperCase();
    if (tokenStatus === 'PENDING') {
      tsUpdateBookingTokenStatus_(hit.rowNum, 'SUPERSEDED');
    }
  });
}

function tsNotifyManagerNoReofferCapacity_(managerName, consultantName, consultantEmail, durationMin, simsCsv) {
  if (!managerName) {
    tsAudit_('REOFFER_MANAGER_NOTIFY', consultantEmail, 'No manager available for no-capacity reoffer notice', 'WARN');
    return;
  }

  var message = [
    '*Reflex-AI training scheduling support needed*',
    '',
    'Capacity constraints do not allow *' + (consultantName || consultantEmail) + '* to be automatically scheduled.',
    'Please manually schedule a *' + durationMin + '-minute* time block for them to complete their simulations in Assembled.',
    simsCsv ? 'Sims: ' + simsCsv : ''
  ].filter(Boolean).join('\n');

  try {
    tsSlackDmManager_(managerName, message);
    tsAudit_('REOFFER_MANAGER_NOTIFY', consultantEmail, 'No-capacity manager DM sent to ' + managerName, 'OK');
  } catch (err) {
    tsAudit_('REOFFER_MANAGER_NOTIFY', consultantEmail, 'No-capacity manager DM failed: ' + err, 'WARN');
  }
}
