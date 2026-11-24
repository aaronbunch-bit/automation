const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../TrainingScheduler_Weekly_Assigned_Mode.gs'), 'utf8');
function context() { const c = {Date}; vm.createContext(c); vm.runInContext(source,c); return c; }
const date = t => new Date('2026-11-24T'+t+':00Z');
const need = {salesGroup:'High School', sims:['Discovery'], weekStart:'2026-11-22'};
const window = (a,b) => ({start:date(a),end:date(b)});
function add(c,o,a,b,n=need) { c.tsIncrementWeeklySlotOccupancy_(o,n,window(a,b)); }
function count(c,o,a,b,n=need) { return c.tsWeeklySlotCount_(o,n,date(a),date(b)); }
test('mixed simulations and staggered starts share the same audience cap',()=>{
 const c=context(),o={};
 add(c,o,'10:00','10:20'); add(c,o,'10:05','10:25',{...need,sims:['Objections']}); add(c,o,'10:10','10:30');
 assert.equal(count(c,o,'10:15','10:35'),3);
 assert.equal(count(c,o,'09:55','10:15'),3);
 assert.equal(count(c,o,'10:20','10:40'),2);
 assert.equal(count(c,o,'10:30','10:50'),0);
 assert.equal(count(c,o,'10:15','10:35',{...need,salesGroup:'College'}),0);
});
test('count peak concurrency, not all touching bookings; adjacent blocks are allowed',()=>{
 const c=context(),o={};
 for(const [a,b] of [['10:00','10:10'],['10:10','10:20'],['10:20','10:30']]) add(c,o,a,b);
 assert.equal(count(c,o,'10:00','10:30'),1);
 assert.equal(count(c,o,'09:40','10:00'),0);
});
test('cap cannot exceed three and invalid audience or interval fails closed',()=>{
 const c=context();
 for(const value of [undefined,0,-1,'bad',4,99]) assert.equal(c.tsWeeklyConcurrencyCap_({WEEKLY_MAX_REPS_PER_SIM_SLOT:value}),3);
 assert.equal(c.tsWeeklyConcurrencyCap_({WEEKLY_MAX_REPS_PER_SIM_SLOT:2}),2);
 assert.throws(()=>count(c,{},'10:00','10:20',{salesGroup:''}),/audience/);
 assert.throws(()=>count(c,{},'10:20','10:00'),/window/);
});
function sheetContext(rows) {
 const c=context();
 c.TS={SHEETS:{OFFERS:'offers'},OFFER_COLS:{STATUS:1,BOOKED_WINDOW:2,SALES_GROUP:3,DURATION_MIN:4}};
 const sheet={getLastRow:()=>rows.length+1,getDataRange:()=>({getValues:()=>[[],...rows]})};
 c.tsGetSpreadsheet_=()=>({getSheetByName:()=>sheet});
 c.tsBuildDateTime_=(d,t)=>new Date(d+'T'+t.padStart(5,'0')+':00Z');
 c.tsGetOfferActiveColumn_=()=>5;
 c.tsOfferRowIsInactive_=r=>r[4]==='N';
 return c;
}
test('loads persisted bookings once, including legacy duration and midnight overlap',()=>{
 const c=sheetContext([
 ['BOOKED','2026-11-24 10:00-10:20','High School',20],
 ['BOOKED','2026-11-24 10:05','High School',20],
 ['BOOKED','2026-11-24 10:10-10:30','College',20],
 ['BOOKED','2026-11-24 10:10-10:30','High School',20,'N'],
 ['OFFERED','2026-11-24 10:10-10:30','High School',20],
 ['BOOKED','2026-11-21 23:50-00:10','High School',20]
 ]);
 const o=c.tsBuildWeeklySlotOccupancy_('2026-11-22');
 assert.equal(count(c,o,'10:10','10:20'),2);
 assert.equal(c.tsWeeklySlotCount_(o,need,new Date('2026-11-22T00:00Z'),new Date('2026-11-22T00:20Z')),1);
});
test('malformed persisted booking stops scheduling instead of undercounting',()=>{
 const c=sheetContext([['BOOKED','2026-11-24 10:00','High School','']]);
 assert.throws(()=>c.tsBuildWeeklySlotOccupancy_('2026-11-22'),/duration/);
});
test('booking rechecks fresh occupancy under lock and releases it on every path',()=>{
 const c=context(); let locked=false, commits=0; let o={};
 c.LockService={getScriptLock:()=>({waitLock:()=>{locked=true},releaseLock:()=>{locked=false}})};
 c.tsBuildWeeklySlotOccupancy_=()=>{assert.equal(locked,true);return o};
 c.tsLoadConfig_=()=>({}); c.tsAudit_=()=>{}; c.tsWeeklyAlreadyScheduled_=()=>false;
 c.tsCommitWeeklyAssignedWindow_=()=>{assert.equal(locked,true);commits++;return true};
 add(c,o,'10:00','10:20');add(c,o,'10:05','10:25');
 assert.equal(c.tsBookWeeklyAssignedWindow_({},need,window('10:10','10:30')),true);
 assert.equal(locked,false); assert.equal(commits,1);
 add(c,o,'10:10','10:30');
 assert.equal(c.tsBookWeeklyAssignedWindow_({},need,window('10:15','10:35')),false);
 assert.equal(locked,false); assert.equal(commits,1);
 c.tsBuildWeeklySlotOccupancy_=()=>{throw Error('read failed')};
 assert.throws(()=>c.tsBookWeeklyAssignedWindow_({},need,window('10:15','10:35')),/read failed/);
 assert.equal(locked,false);
});
test('window selection rejects the fourth overlap and accepts the next adjacent window',()=>{
 const c=context(),o={};
 const d=t=>new Date('2099-11-24T'+t+':00Z');
 c.TS={BUSINESS_HOUR_START:9,BUSINESS_HOUR_END:17,TZ:'UTC'};
 c.tsBuildDateTime_=(day,t)=>new Date(day+'T'+t+':00Z');
 c.tsPad2_=n=>String(n).padStart(2,'0'); c.tsConsultantOnShiftDuringWindow_=()=>true;
 c.Utilities={formatDate:dt=>dt.toISOString().slice(11,16)};
 c.tsNum_=Number; c.tsIsNum_=v=>Number.isFinite(v);
 for(const start of ['10:00','10:05','10:10']) c.tsIncrementWeeklySlotOccupancy_(o,need,{start:d(start),end:d('10:30')});
 const forecast={'2099-11-24':['10:15','10:30'].map(t=>({start_time:d(t).getTime()/1000,staffing_net:10}))};
 const result=c.tsFindBestWeeklyTrainingWindow_(forecast,{},need,'rep@example.test',1200000,{},o,3);
 assert.equal(result.startStr,'10:30');
});
