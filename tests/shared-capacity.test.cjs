const test=require('node:test'),assert=require('node:assert/strict'),vm=require('node:vm'),fs=require('node:fs');
const source=fs.readFileSync(require('node:path').join(__dirname, '../ReflexAI_Capacity_Guard.gs'),'utf8');
function harness(){
 const ledger={}, state={locked:false,posts:0,rows:[],fail:false};let seq=0;
 const c={Date,PropertiesService:{getScriptProperties:()=>({getProperties:()=>({...ledger}),setProperty:(k,v)=>ledger[k]=v,deleteProperty:k=>delete ledger[k]})},
 LockService:{getScriptLock:()=>({hasLock:()=>state.locked,waitLock:()=>state.locked=true,releaseLock:()=>state.locked=false})},
 Utilities:{getUuid:()=>String(++seq)},UrlFetchApp:{fetch:()=>{assert.equal(state.locked,true);state.posts++;if(state.fail)throw Error('timeout');return {getResponseCode:()=>200,getContentText:()=>JSON.stringify({id:'a'+seq})}}}};
 vm.createContext(c);vm.runInContext(source,c);
 c.rcContext_=()=>({trainingId:'training',people:{rep:{groups:['high school']},other:{groups:['college']}}});
 c.rcActivities_=()=>state.rows;
 const start=Math.floor(Date.now()/1000)+86400;
 const payload={start_time:start,end_time:start+1200,type_id:'training',agent_id:'rep',description:'ReflexAI test'};
 return {c,state,ledger,payload,start};
}
test('shared ledger blocks fourth write even before earlier writes appear in Assembled',()=>{
 const {c,state,payload}=harness();
 for(let i=0;i<3;i++)c.commitReflexActivity({}, {...payload,start_time:payload.start_time+i*60},'High School',99);
 assert.throws(()=>c.commitReflexActivity({},payload,'High School',3),/overlap/);
 assert.equal(state.posts,3);assert.equal(state.locked,false);
});
test('live training across simulations counts; separate audience does not',()=>{
 const {c,state,payload,start}=harness();
 state.rows=[0,60,120].map((offset,i)=>({id:String(i),type_id:'training',agent_id:'rep',start_time:start+offset,end_time:start+1800,description:'Different sim '+i}));
 assert.throws(()=>c.commitReflexActivity({},payload,'High School',3),/overlap/);assert.equal(state.posts,0);
 state.rows.forEach(a=>a.agent_id='other');
 c.commitReflexActivity({},payload,'High School',3);assert.equal(state.posts,1);
});
test('targeted two-person cap preserved and existing lock remains owned by caller',()=>{
 const {c,state,payload}=harness();state.locked=true;
 c.commitReflexActivity({},payload,'High School',2);c.commitReflexActivity({},payload,'High School',2);
 assert.throws(()=>c.commitReflexActivity({},payload,'High School',2),/overlap/);
 assert.equal(state.posts,2);assert.equal(state.locked,true);
});
test('uncertain write retains a capacity reservation',()=>{
 const {c,state,payload}=harness();state.fail=true;
 assert.throws(()=>c.commitReflexActivity({},payload,'High School',1),/timeout/);
 state.fail=false;
 assert.throws(()=>c.commitReflexActivity({},payload,'High School',1),/overlap/);
 assert.equal(state.posts,1);assert.equal(state.locked,false);
});
test('visible reservation is not double counted and adjacent bookings are allowed',()=>{
 const {c,state,payload,start}=harness();c.commitReflexActivity({},payload,'High School',2);
 state.rows=[{id:'a1',agent_id:'rep',type_id:'training',start_time:start,end_time:start+1200}];
 c.commitReflexActivity({},payload,'High School',2);
 c.commitReflexActivity({},{...payload,start_time:start+1200,end_time:start+2400},'High School',2);
 assert.equal(state.posts,3);
});
test('unknown audience, wrong type, and failed read prevent writes',()=>{
 const {c,state,payload}=harness();
 assert.throws(()=>c.commitReflexActivity({},payload,'College',3),/audience/);
 assert.throws(()=>c.commitReflexActivity({},{...payload,type_id:'phone'},'High School',3),/Training/);
 c.rcActivities_=()=>{throw Error('read failed')};
 assert.throws(()=>c.commitReflexActivity({},payload,'High School',3),/read failed/);
 assert.equal(state.posts,0);assert.equal(state.locked,false);
});
