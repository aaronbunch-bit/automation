/**
 * Auto-escalation for reps who cannot be automatically scheduled because no
 * capacity-safe windows are available.
 *
 * REQUIRED MAIN SCHEDULER EDIT:
 *
 * In tsProcessOneConsultantOffer_, replace:
 *
 *   if (!result.windows.length) {
 *     tsAudit_('OFFERS', need.email,
 *       'No capacity-safe windows in next ' + (config.SEARCH_DAYS || TS.SEARCH_DAYS) + ' weekdays — ' + result.diag,
 *       'WARN');
 *     return;
 *   }
 *
 * with:
 *
 *   if (!result.windows.length) {
 *     tsAutoEscalateNoCapacity_(need, result.diag, config);
 *     return;
 *   }
 */

function tsAutoEscalateNoCapacity_(need, diag, config) {
  config = config || tsLoadConfig_();

  if (tsConsultantAlreadyBooked_(need.email, need.salesGroup)) {
    tsAudit_('AUTO_ESCALATE_NO_CAPACITY', need.email, 'Already booked — no escalation created', 'INFO');
    return;
  }

  if (tsHasEscalatedOffer_(need.email, need.salesGroup)) {
    tsAudit_('AUTO_ESCALATE_NO_CAPACITY', need.email, 'Already escalated — no duplicate escalation created', 'INFO');
    return;
  }

  var offerRow = tsAppendOfferRow_(need);
  var offerSheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  var now = new Date();
  var reason = 'No capacity-safe windows in next ' + (config.SEARCH_DAYS || TS.SEARCH_DAYS) + ' weekdays — ' + diag;

  offerSheet.getRange(offerRow, TS.OFFER_COLS.STATUS).setValue('ESCALATED');
  offerSheet.getRange(offerRow, TS.OFFER_COLS.OFFER_SENT_AT).setValue(now);
  offerSheet.getRange(offerRow, TS.OFFER_COLS.REMINDER_SENT_AT).setValue(now);
  SpreadsheetApp.flush();

  tsNotifyRepNoCapacity_(need);
  tsNotifyManagerNoCapacity_(need, reason);

  tsAudit_('AUTO_ESCALATE_NO_CAPACITY', need.email, reason, 'OK');
  tsRefreshAnalyticsSafe_();
}

function tsHasEscalatedOffer_(email, salesGroup) {
  var sheet = tsGetSpreadsheet_().getSheetByName(TS.SHEETS.OFFERS);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var targetEmail = String(email || '').trim().toLowerCase();
  var targetGroup = String(salesGroup || '').trim().toLowerCase();
  var activeColumn = typeof tsGetOfferActiveColumn_ === 'function'
    ? tsGetOfferActiveColumn_(sheet)
    : 0;
  var values = sheet.getDataRange().getValues();

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowEmail = String(row[TS.OFFER_COLS.CONSULTANT_EMAIL - 1] || '').trim().toLowerCase();
    var rowGroup = String(row[TS.OFFER_COLS.SALES_GROUP - 1] || '').trim().toLowerCase();
    var status = String(row[TS.OFFER_COLS.STATUS - 1] || '').trim().toUpperCase();

    if (rowEmail !== targetEmail || rowGroup !== targetGroup) continue;
    if (activeColumn && typeof tsOfferRowIsInactive_ === 'function' && tsOfferRowIsInactive_(row, activeColumn)) continue;
    if (status === 'ESCALATED') return true;
  }

  return false;
}

function tsNotifyRepNoCapacity_(need) {
  var message = [
    '*Reflex-AI training scheduling update*',
    '',
    'Your current schedule does not have enough open capacity-safe windows for automatic scheduling of your Reflex-AI simulations.',
    '',
    'Your manager has been notified and will help schedule time manually.',
    '',
    'No action is needed from you right now.'
  ].join('\n');

  try {
    tsSlackDmConsultant_(need.name, need.email, message);
    tsAudit_('AUTO_ESCALATE_REP_NOTIFY', need.email, 'Rep no-capacity DM sent', 'OK');
  } catch (err) {
    tsAudit_('AUTO_ESCALATE_REP_NOTIFY', need.email, 'Rep no-capacity DM failed: ' + err, 'WARN');
  }
}

function tsNotifyManagerNoCapacity_(need, reason) {
  if (!need.managerName) {
    tsAudit_('AUTO_ESCALATE_MANAGER_NOTIFY', need.email, 'No manager available for no-capacity escalation', 'WARN');
    return;
  }

  var message = [
    '*Reflex-AI training manual scheduling needed*',
    '',
    '*' + (need.name || need.email) + '* could not be automatically scheduled because there is not enough capacity within their current schedule.',
    '',
    'Please manually schedule a *' + need.durationMin + '-minute* Training block in Assembled for this rep to complete their simulations.',
    need.sims && need.sims.length ? 'Sims: ' + need.sims.join(', ') : '',
    '',
    'This has been auto-escalated because no capacity-safe automatic windows were available.'
  ].filter(Boolean).join('\n');

  try {
    tsSlackDmManager_(need.managerName, message);
    tsAudit_('AUTO_ESCALATE_MANAGER_NOTIFY', need.email, 'Manager no-capacity DM sent to ' + need.managerName, 'OK');
  } catch (err) {
    tsAudit_('AUTO_ESCALATE_MANAGER_NOTIFY', need.email, 'Manager no-capacity DM failed: ' + err, 'WARN');
  }
}
