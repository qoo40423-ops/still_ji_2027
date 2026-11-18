/** 紀家烘焙坊 116 中秋｜獨立管理 GAS（TEST）
 * 請建立另一個 Apps Script 專案，部署為「執行身分：我／存取權：只有自己」。
 * 不要貼到既有公開匯款回報的 Code.gs。
 */
const MANAGER_TEST_SHEET_ID='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4';
const MANAGER_TZ='Asia/Taipei';
function doGet(){return HtmlService.createHtmlOutputFromFile('Manager2027').setTitle('紀家烘焙坊｜中秋管理').addMetaTag('viewport','width=device-width, initial-scale=1');}
function managerSnapshot2027(){
 const ss=SpreadsheetApp.openById(MANAGER_TEST_SHEET_ID);
 const get=(name,n)=>{const sh=ss.getSheetByName(name);if(!sh)throw Error('缺少工作表：'+name);return sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,n).getValues().filter(r=>String(r[0]||'').trim());};
 const stamp=v=>v instanceof Date?Utilities.formatDate(v,MANAGER_TZ,'yyyy/MM/dd HH:mm'):String(v==null?'':v);
 const day=v=>v instanceof Date?Utilities.formatDate(v,MANAGER_TZ,'yyyy-MM-dd'):String(v||'').trim();
 const pickup=get('門市自取',18).map(v=>({type:'自取',id:String(v[0]),date:day(v[2]),slot:stamp(v[3]),boxes6:Number(v[6])||0,boxes15:Number(v[7])||0,name:String(v[4]),phone:String(v[5]),quantity:'6入×'+v[6]+'、15入×'+v[7],pieces:Number(v[8])||0,amount:Number(v[11])||0,deadline:stamp(v[12]),payment:String(v[13]),remittance:String(v[14]||''),verified:stamp(v[15]),status:String(v[16]),note:String(v[17]||'')}));
 const delivery=get('宅配到府',25).map(v=>({type:'宅配',id:String(v[0]),date:day(v[2]),slot:'出貨',boxes6:4*(Number(v[6])||0)+10*(Number(v[7])||0)+15*(Number(v[8])||0),boxes15:Number(v[9])||0,name:String(v[3]),phone:String(v[4]),address:String(v[5]),quantity:'4盒組×'+v[6]+'、10盒組×'+v[7]+'、15盒組×'+v[8]+'、15入×'+v[9],pieces:Number(v[11])||0,amount:Number(v[16])||0,deadline:stamp(v[17]),payment:String(v[18]),remittance:String(v[19]||''),verified:stamp(v[20]),slipConfirmation:String(v[21]||''),tracking:String(v[22]||''),status:String(v[23]),note:String(v[24]||'')}));
 const empty=()=>({boxes6:0,boxes15:0,pieces:0,pans:0});
 const add=(target,o)=>{target.boxes6+=o.boxes6;target.boxes15+=o.boxes15;target.pieces+=o.boxes6*6+o.boxes15*15;};
 const finish=x=>{x.pans=Math.ceil(x.pieces/60);return x;};
 const map={};
 const make=date=>map[date]||(map[date]={date,morning:empty(),afternoon:empty(),pickup:empty(),delivery:empty(),slots:{'09:00~12:00':empty(),'13:00~15:00':empty(),'17:00~18:00':empty()},unclassified:empty()});
 const active=o=>!['已取消','取消','作廢'].includes(o.status);
 const slotKey=slot=>{const t=String(slot||'').replace(/\\s/g,'').replace(/[～〜－–—]/g,'~');if(/09[:：]?00|09:00~12:00|上午/.test(t))return '09:00~12:00';if(/13[:：]?00|13:00~15:00/.test(t))return '13:00~15:00';if(/17[:：]?00|17:00~18:00/.test(t))return '17:00~18:00';return '';};
 pickup.filter(active).forEach(o=>{const d=make(o.date);add(d.pickup,o);const key=slotKey(o.slot);if(key){add(d.slots[key],o);add(key==='09:00~12:00'?d.morning:d.afternoon,o);}else add(d.unclassified,o);});
 delivery.filter(active).forEach(o=>{const d=make(o.date);add(d.delivery,o);add(d.morning,o);});
 const capacity=Object.values(map).sort((a,b)=>a.date.localeCompare(b.date)).map(d=>{Object.values(d.slots).forEach(finish);['morning','afternoon','pickup','delivery','unclassified'].forEach(k=>finish(d[k]));const total=d.pickup.pieces+d.delivery.pieces;return {...d,total,pans:Math.ceil(total/60),remaining:Math.max(0,4800-total),limit:4800};});
 return {generatedAt:stamp(new Date()),pickup,delivery,capacity,counts:{pickup:pickup.length,delivery:delivery.length,pendingReview:pickup.concat(delivery).filter(o=>o.payment==='待核帳').length,unpaid:pickup.concat(delivery).filter(o=>o.payment==='未付款'&&!['已取消','取消','作廢'].includes(o.status)).length}};
}


/** 僅允許已核帳的自取訂單標記已取貨；操作僅更新 Q 欄。 */
function markPickupCollected2027(orderId){
 const id=String(orderId||'').trim();
 if(!/^SJ27-[A-Z0-9-]{8,30}$/.test(id))throw Error('訂單編號格式不正確');
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(15000))throw Error('系統忙碌，請稍後重試');
 try{
  const sh=SpreadsheetApp.openById(MANAGER_TEST_SHEET_ID).getSheetByName('門市自取');
  if(!sh)throw Error('找不到門市自取工作表');
  const last=sh.getLastRow();
  if(last<2)throw Error('找不到訂單');
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues().map(r=>r[0].trim());
  const matches=ids.map((x,i)=>x===id?i+2:0).filter(Boolean);
  if(matches.length!==1)throw Error(matches.length?'訂單編號重複，請人工檢查':'找不到訂單');
  const row=matches[0];
  const pay=String(sh.getRange(row,14).getDisplayValue()).trim();
  const current=String(sh.getRange(row,17).getDisplayValue()).trim();
  if(current==='已取貨')return {ok:true,already:true,id,status:'已取貨'};
  if(pay!=='已付款')throw Error('此訂單尚未完成核帳，不能標記已取貨');
  if(!['已確認','製作中'].includes(current))throw Error('目前訂單狀態「'+current+'」不允許標記已取貨');
  sh.getRange(row,17).setValue('已取貨');
  SpreadsheetApp.flush();
  return {ok:true,already:false,id,status:'已取貨'};
 }finally{lock.releaseLock();}
}
