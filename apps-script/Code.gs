/**
 * 紀家烘焙坊｜116 年中秋訂單後端（目前仍連接 TEST 試算表）
 * 保留訂單建立、付款連動、逾期保護、產能統計及觸發器。
 * 已移除一次性測試／模擬函式；歷史完整版本見 archive。
 * S-11D 新增自取匯款回報 doPost；尚未部署或啟用前端送出。
 */

const SJ27 = Object.freeze({
  SPREADSHEET_ID: '1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4',
  PICKUP: '門市自取', DELIVERY: '宅配到府', SUMMARY: '每日產能統計',
  LIMIT: 4800, TZ: 'Asia/Taipei', PAY_HOURS: 72,
  PICKUP_HEADERS: ['訂單編號','成立時間','取貨日期','取貨時段','提貨人','聯絡電話','6入盒數','15入盒數','蛋黃酥顆數','商品金額','優惠類型','應付總額','付款截止時間','付款狀態','匯款後五碼','核帳日期','訂單狀態','備註'],
  DELIVERY_HEADERS: ['訂單編號','成立時間','出貨日期','收件人','聯絡電話','收件地址','6入4盒組','6入10盒組','6入15盒組','15入盒數','提袋數','蛋黃酥顆數','商品金額','提袋費','基本運費','偏遠加價','應付總額','付款截止時間','付款狀態','匯款後五碼','核帳日期','宅配單確認','黑貓單號','訂單狀態','備註'],
  SUMMARY_HEADERS: ['製作日期','自取顆數','宅配顆數','合計顆數','換算盤數','剩餘顆數','每日上限','狀態']
});
/** 五碼訂單號：跨自取／宅配表檢查，保留舊編號且不重複使用。 */
function sj27ShortOrderId_(ss){
 const used=new Set();
 for(const name of [SJ27.PICKUP,SJ27.DELIVERY]){
  const sh=ss.getSheetByName(name);
  if(sh&&sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().forEach(r=>used.add(String(r[0]).trim().toUpperCase()));
 }
 for(let n=0;n<200;n++){
  const id='SJ-'+String(Math.floor(Math.random()*90000)+10000);
  if(!used.has(id))return id;
 }
 throw Error('暫時無法產生唯一訂單編號，請稍後重試');
}
function sjSpreadsheet_(){return SpreadsheetApp.openById(SJ27.SPREADSHEET_ID);}
function sjSheet_(name){
  const sh=sjSpreadsheet_().getSheetByName(name);
  if(!sh)throw Error('找不到工作表：'+name);
  const expected=name===SJ27.PICKUP?SJ27.PICKUP_HEADERS:name===SJ27.DELIVERY?SJ27.DELIVERY_HEADERS:SJ27.SUMMARY_HEADERS;
  const actual=sh.getRange(1,1,1,expected.length).getDisplayValues()[0];
  if(expected.some((h,i)=>h!==actual[i]))throw Error(name+' 表頭不符，已停止寫入，請勿直接改欄位順序');
  return sh;
}
function sjText_(x,max,label){const v=String(x==null?'':x).trim();if(v.length>max)throw Error(label+'過長');return v;}
function sjNum_(x,label){const n=Number(x==null||x===''?0:x);if(!Number.isSafeInteger(n)||n<0)throw Error(label+'須為非負整數');return n;}
function sjDate_(v){
  const s=sjText_(v,10,'日期');
  if(!/^2027-(08|09)-\d{2}$/.test(s))throw Error('目前僅接受 2027 年 8～9 月測試日期');
  const d=new Date(s+'T00:00:00Z');
  if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==s||s>'2027-09-15')throw Error('日期無效或超出測試範圍');
  if(s<Utilities.formatDate(new Date(),SJ27.TZ,'yyyy-MM-dd'))throw Error('不可選擇過去日期');
  return s;
}
function sjPhone_(v){
  const s=sjText_(v,30,'電話').replace(/[\s()－-]/g,'');
  if(!/^(09\d{8}|0[2-8]\d{7,8}(?:#\d{1,6})?)$/.test(s))throw Error('請輸入有效的手機或市話');
  return s;
}
function sjRemote_(address){return /^(宜蘭縣|宜蘭市|花蓮縣|臺東縣|台東縣)/.test(address.replace(/\s/g,''));}
function sjAddress_(v){
  const a=sjText_(v,200,'地址');
  if(a.length<8)throw Error('請填寫完整宅配地址');
  if(/^(澎湖|金門|連江|馬祖|小琉球|綠島|蘭嶼)/.test(a))throw Error('目前僅提供臺灣本島宅配');
  return a;
}
function sjStatusActive_(v){return !['已取消','取消','作廢','退款完成','已退款'].includes(String(v||'').trim());}
function sjDataRows_(name){
  const sh=sjSheet_(name),width=name===SJ27.PICKUP?SJ27.PICKUP_HEADERS.length:SJ27.DELIVERY_HEADERS.length;
  if(sh.getLastRow()<2)return [];
  return sh.getRange(2,1,sh.getLastRow()-1,width).getValues().filter(r=>String(r[0]||'').trim());
}
function sjDateValue_(v){return v instanceof Date?Utilities.formatDate(v,SJ27.TZ,'yyyy-MM-dd'):String(v||'').trim();}
function sjUsage_(date){
  let pickup=0,delivery=0;
  sjDataRows_(SJ27.PICKUP).forEach(r=>{if(sjDateValue_(r[2])===date&&sjStatusActive_(r[16]))pickup+=Number(r[8])||0;});
  sjDataRows_(SJ27.DELIVERY).forEach(r=>{if(sjDateValue_(r[2])===date&&sjStatusActive_(r[23]))delivery+=Number(r[11])||0;});
  return {pickup,delivery,total:pickup+delivery};
}
function sjNextRow_(sh){
  const last=Math.max(2,sh.getLastRow());
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
  const index=ids.findIndex(r=>!String(r[0]).trim());
  const row=index<0?last+1:index+2;
  if(row>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),Math.max(100,row-sh.getMaxRows()));
  return row;
}
function sjDuplicate_(requestId){
  if(!requestId)return null;
  for(const name of [SJ27.PICKUP,SJ27.DELIVERY]){
    const sh=sjSheet_(name),last=sh.getLastRow();
    if(last<2)continue;
    const notesCol=name===SJ27.PICKUP?18:25;
    const rows=sh.getRange(2,1,last-1,notesCol).getDisplayValues();
    const hit=rows.find(r=>r[0]&&r[notesCol-1].includes('[RID:'+requestId+']'));
    if(hit)return {orderId:hit[0],duplicate:true};
  }
  return null;
}
/** 只做唯讀檢查；不建立工作表、不清除任何資料。 */
function inspect2027(){
  const result={spreadsheet:sjSpreadsheet_().getName(),timeZone:sjSpreadsheet_().getSpreadsheetTimeZone(),sheets:{}};
  [SJ27.PICKUP,SJ27.DELIVERY,SJ27.SUMMARY].forEach(name=>{
    const sh=sjSheet_(name);
    result.sheets[name]={headersOk:true,existingOrders:name===SJ27.SUMMARY?null:sjDataRows_(name).length};
  });
  Logger.log(JSON.stringify(result));return result;
}
/**
 * 內部測試建立訂單。需提供唯一 requestId (至少 12 字)。
 * pickup: {method:'自取',name,phone,date,pickupTime,sixBoxes,fifteenBoxes,notes,requestId}
 * delivery: {method:'宅配',name,phone,date,address,four,ten,fifteenGroup,fifteenBoxes,bags,notes,requestId}
 * 不信任前端傳入的價格、運費、優惠碼或付款狀態。
 */
function createOrder2027(data){
  if(!data||typeof data!=='object')throw Error('缺少訂單資料');
  const method=sjText_(data.method,10,'方式');
  if(method!=='自取'&&method!=='宅配')throw Error('訂購方式錯誤');
  const name=sjText_(data.name,60,'姓名'),phone=sjPhone_(data.phone),date=sjDate_(data.date||data.productionDate);
  if(!name)throw Error('姓名不可空白');
  const rid=sjText_(data.requestId,100,'請求識別碼');
  if(!/^[a-zA-Z0-9_-]{12,100}$/.test(rid))throw Error('請提供至少 12 字的唯一 requestId');
  const notes=sjText_(data.notes,300,'備註');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌，請稍後再試');
  try{
    const dup=sjDuplicate_(rid);if(dup)return dup;
    let pieces,values,sh;
    const created=new Date(),deadline=new Date(created.getTime()+SJ27.PAY_HOURS*3600000);
    const orderId=sj27ShortOrderId_(sjSpreadsheet_());
    if(method==='自取'){
      const six=sjNum_(data.sixBoxes,'6入盒數'),fifteen=sjNum_(data.fifteenBoxes,'15入盒數');
      pieces=six*6+fifteen*15;
      if(!pieces||(six&&!fifteen&&six<8)||(fifteen&&!six&&fifteen<3)||(six&&fifteen&&pieces<42))throw Error('未達訂單最小數量標準，僅能現場購買，敬請見諒');
      const time=sjText_(data.pickupTime,15,'取貨時段'),weekend=[0,6].includes(new Date(date+'T00:00:00Z').getUTCDay());
      const allowed=date==='2027-09-15'?['09:00–10:00']:weekend?['09:00–12:00','13:00–15:00','15:00–17:00','17:00–18:00']:['13:00–15:00','15:00–17:00','17:00–18:00'];
      if(!allowed.includes(time))throw Error('取貨時段不符合選定日期');
      // 員工價須經管理端授權；備註 8888 不會自行觸發折扣。
      const staff=data.staffApproved===true;
      if(staff)throw Error('員工優惠尚未啟用後端授權，請先以原價測試');
      const goods=six*300+fifteen*750;
      values=[orderId,created,date,time,name,phone,six,fifteen,pieces,goods,'一般售價',goods,deadline,'未付款','','','待確認',notes+' [RID:'+rid+']'];
      sh=sjSheet_(SJ27.PICKUP);
    }else{
      const address=sjAddress_(data.address),four=sjNum_(data.four,'4盒組'),ten=sjNum_(data.ten,'10盒組'),group=sjNum_(data.fifteenGroup,'15盒組'),fifteen=sjNum_(data.fifteenBoxes,'15入盒數'),bags=sjNum_(data.bags,'提袋');
      if([0,6].includes(new Date(date+'T00:00:00Z').getUTCDay()))throw Error('週末不提供出貨');
      if(fifteen>0&&fifteen<4)throw Error('15入宅配至少4盒');
      if(bags>0&&!group)throw Error('提袋僅限6入15盒組');
      const six=four*4+ten*10+group*15;pieces=six*6+fifteen*15;
      if(!pieces)throw Error('請選擇商品');
      const bagSplit=Array.from({length:group},(_,i)=>Math.floor(bags/group)+(i<bags%group?1:0));
      const goods=six*300+fifteen*750,bagFee=bags*20,cartons15=Math.ceil(fifteen/10);
      const shipping=four*150+ten*170+bagSplit.reduce((sum,n)=>sum+(n>=5?210:170),0)+cartons15*210;
      const remoteFee=sjRemote_(address)?(four+ten+group+cartons15)*10:0;
      values=[orderId,created,date,name,phone,address,four,ten,group,fifteen,bags,pieces,goods,bagFee,shipping,remoteFee,goods+bagFee+shipping+remoteFee,deadline,'未付款','','','待回傳','','待確認',notes+' [RID:'+rid+']'];
      sh=sjSheet_(SJ27.DELIVERY);
    }
    if(pieces>SJ27.LIMIT)throw Error('單筆訂單超過每日產能上限');
    const used=sjUsage_(date).total;
    if(used+pieces>SJ27.LIMIT)throw Error('當日產能不足，剩餘 '+(SJ27.LIMIT-used)+' 顆');
    const row=sjNextRow_(sh);
    // 電話必須以純文字保存，避免 Google Sheets 自動移除開頭的 0。
    sh.getRange(row,method==='自取'?6:5).setNumberFormat('@');
    sh.getRange(row,1,1,values.length).setValues([values]);
    sh.getRange(row,2).setNumberFormat('yyyy/mm/dd hh:mm');
    sh.getRange(row,method==='自取'?13:18).setNumberFormat('yyyy/mm/dd hh:mm');
    SpreadsheetApp.flush();
    refreshCapacitySummary2027_();
    return {ok:true,orderId,method,date,pieces,total:method==='自取'?values[11]:values[16],paymentDeadline:Utilities.formatDate(deadline,SJ27.TZ,'yyyy/MM/dd HH:mm'),testOnly:true};
  }finally{lock.releaseLock();}
}
function refreshCapacitySummary2027_(){
  const sh=sjSheet_(SJ27.SUMMARY),map={};
  sjDataRows_(SJ27.PICKUP).forEach(r=>{const date=sjDateValue_(r[2]);if(!/^2027-\d{2}-\d{2}$/.test(date)||!sjStatusActive_(r[16]))return;(map[date]||(map[date]={p:0,d:0})).p+=Number(r[8])||0;});
  sjDataRows_(SJ27.DELIVERY).forEach(r=>{const date=sjDateValue_(r[2]);if(!/^2027-\d{2}-\d{2}$/.test(date)||!sjStatusActive_(r[23]))return;(map[date]||(map[date]={p:0,d:0})).d+=Number(r[11])||0;});
  const rows=Object.keys(map).sort().map(date=>{const {p,d}=map[date],total=p+d;return [date,p,d,total,Math.ceil(total/60),Math.max(0,SJ27.LIMIT-total),SJ27.LIMIT,total>SJ27.LIMIT?'超量':total===SJ27.LIMIT?'額滿':'可接單'];});
  // 只更新統計表 A:H 的資料列，保留 I 欄以後原有說明。
  const last=Math.max(2,sh.getLastRow());
  sh.getRange(2,1,last-1,8).clearContent();
  if(rows.length){if(rows.length+1>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());sh.getRange(2,1,rows.length,8).setValues(rows);}
  return rows.length;
}
function refreshCapacitySummary2027(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(30000))throw Error('系統忙碌');
  try{return refreshCapacitySummary2027_();}finally{lock.releaseLock();}
}
function capacityForDate2027(date){
  const d=sjDate_(date),u=sjUsage_(d);
  return {date:d,pickup:u.pickup,delivery:u.delivery,used:u.total,remaining:Math.max(0,SJ27.LIMIT-u.total),limit:SJ27.LIMIT,pans:Math.ceil(u.total/60)};
}

function paymentAudit2027(){
  const now=new Date(),result={mode:'READ ONLY',checkedAt:Utilities.formatDate(now,SJ27.TZ,'yyyy/MM/dd HH:mm:ss'),expiredCandidates:[],awaitingReconciliation:[],warnings:[],counts:{pickup:0,delivery:0}};
  for(const [name,kind,deadlineIndex,payIndex,statusIndex,amountIndex] of [
    [SJ27.PICKUP,'自取',12,13,16,11],
    [SJ27.DELIVERY,'宅配',17,18,23,16]
  ]){
    const rows=sjDataRows_(name);result.counts[kind==='自取'?'pickup':'delivery']=rows.length;
    rows.forEach(r=>{
      const orderId=String(r[0]||''),pay=String(r[payIndex]||'').trim(),status=String(r[statusIndex]||'').trim();
      const deadline=r[deadlineIndex],deadlineTime=deadline instanceof Date?deadline.getTime():NaN;
      const info={orderId,method:kind,paymentStatus:pay,orderStatus:status,amount:Number(r[amountIndex])||0,deadline:deadline instanceof Date?Utilities.formatDate(deadline,SJ27.TZ,'yyyy/MM/dd HH:mm'):String(deadline||'')};
      if(pay==='待核帳')result.awaitingReconciliation.push(info);
      if(!sjStatusActive_(status)||pay==='已付款'||pay==='退款中'||pay==='已退款')return;
      if(!Number.isFinite(deadlineTime)){result.warnings.push({orderId,issue:'付款截止時間無法解析，未判定逾期'});return;}
      if(pay==='未付款'&&deadlineTime<now.getTime())result.expiredCandidates.push(info);
      else if(pay==='逾期未付')result.warnings.push({orderId,issue:'已標記逾期未付，但尚未取消；請人工核對'});
    });
  }
  Logger.log('===== S-02 付款稽核｜READ ONLY =====');
  Logger.log(JSON.stringify(result));
  return result;
}

function previewPaymentSync2027(){
  const result={mode:'READ ONLY',changes:[],warnings:[]};
  for(const spec of [
    {name:SJ27.PICKUP,pay:14,verified:16,status:17,target:'已確認'},
    {name:SJ27.DELIVERY,pay:19,verified:21,status:24,target:'待出貨'}
  ]){
    const sh=sjSheet_(spec.name),last=sh.getLastRow();
    if(last<2)continue;
    const rows=sh.getRange(2,1,last-1,spec.status).getValues();
    rows.forEach((r,i)=>{
      const orderId=String(r[0]||'').trim();
      if(!orderId||String(r[spec.pay-1]||'').trim()!=='已付款')return;
      const status=String(r[spec.status-1]||'').trim();
      if(!sjStatusActive_(status)){result.warnings.push({orderId,reason:'訂單已取消，不自動恢復'});return;}
      const updates=[];
      if(!r[spec.verified-1])updates.push({column:spec.verified,value:'NOW'});
      if(['待付款','待確認'].includes(status))updates.push({column:spec.status,value:spec.target});
      if(updates.length)result.changes.push({sheet:spec.name,row:i+2,orderId,updates});
    });
  }
  Logger.log('===== S-03 付款連動預覽｜READ ONLY =====');
  Logger.log(JSON.stringify(result));return result;
}

function applyPaymentSync2027(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌，請稍後重試');
  try{
    const plan=previewPaymentSync2027(),done=[],skipped=[];
    for(const item of plan.changes){
      const sh=sjSheet_(item.sheet),isPickup=item.sheet===SJ27.PICKUP;
      const payCol=isPickup?14:19,statusCol=isPickup?17:24,verifiedCol=isPickup?16:21;
      const id=String(sh.getRange(item.row,1).getValue()||'').trim();
      const pay=String(sh.getRange(item.row,payCol).getValue()||'').trim();
      const status=String(sh.getRange(item.row,statusCol).getValue()||'').trim();
      if(id!==item.orderId||pay!=='已付款'||!sjStatusActive_(status)){skipped.push(item.orderId);continue;}
      if(!sh.getRange(item.row,verifiedCol).getValue()){
        sh.getRange(item.row,verifiedCol).setValue(new Date()).setNumberFormat('yyyy/mm/dd hh:mm');
      }
      if(['待付款','待確認'].includes(status)){
        sh.getRange(item.row,statusCol).setValue(isPickup?'已確認':'待出貨');
      }
      done.push(item.orderId);
    }
    const result={mode:'WRITE',updated:done,skipped};
    Logger.log(JSON.stringify(result));return result;
  }finally{lock.releaseLock();}
}

/**
 * S-04｜安裝型編輯觸發器：只有人工把付款欄改為「已付款」才同步該列。
 * 安裝型觸發器須由具有試算表權限的管理者執行 setupPaymentEditTrigger2027 一次。
 * 不使用簡易 onEdit（跨檔 openById 權限可能不足）。
 */
function paymentEditHandler2027(e){
  if(!e||!e.range||!e.source)return;
  if(e.source.getId()!==SJ27.SPREADSHEET_ID)return;
  const sh=e.range.getSheet(),name=sh.getName();
  const isPickup=name===SJ27.PICKUP;
  if(!isPickup&&name!==SJ27.DELIVERY)return;
  const paymentCol=isPickup?14:19,statusCol=isPickup?17:24,verifiedCol=isPickup?16:21;
  const codeCol=isPickup?15:null;
  // 單格人工編輯；門市自取 O 欄輸入 88888，或原本的付款狀態欄改成已付款。
  if(e.range.getNumRows()!==1||e.range.getNumColumns()!==1||e.range.getRow()<2)return;
  const editedCol=e.range.getColumn(),editedValue=String(e.value||'').trim();
  const isCounterPayment=isPickup&&editedCol===codeCol&&editedValue==='88888';
  const isManualPaid=editedCol===paymentCol&&editedValue==='已付款';
  if(!isCounterPayment&&!isManualPaid)return;
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('付款連動鎖定逾時，請人工檢查該筆訂單');
  try{
    const row=e.range.getRow(),orderId=String(sh.getRange(row,1).getValue()||'').trim();
    const payCell=sh.getRange(row,paymentCol);
    const pay=String(payCell.getValue()||'').trim();
    const status=String(sh.getRange(row,statusCol).getValue()||'').trim();
    if(!orderId||!sjStatusActive_(status))return;
    if(isCounterPayment){
      // 防止退款中、待核帳、已付款等訂單因誤填 88888 而覆寫財務狀態。
      if(pay!=='未付款'){
        Logger.log(JSON.stringify({event:'COUNTER_PAYMENT_SKIPPED',sheet:name,row,orderId,payment:pay,reason:'只接受未付款訂單'}));
        return;
      }
      if(String(sh.getRange(row,codeCol).getDisplayValue()).trim()!=='88888')return;
      payCell.setValue('已付款');
    }else if(pay!=='已付款'){
      return;
    }
    const verified=sh.getRange(row,verifiedCol);
    if(!verified.getValue())verified.setValue(new Date()).setNumberFormat('yyyy/mm/dd hh:mm');
    if(['待付款','待確認'].includes(status))
      sh.getRange(row,statusCol).setValue(isPickup?'已確認':'待出貨');
    Logger.log(JSON.stringify({event:isCounterPayment?'COUNTER_PAYMENT_SYNC':'PAYMENT_SYNC',sheet:name,row,orderId,statusBefore:status}));
  }finally{lock.releaseLock();}
}

/**
 * S-04A｜門市付款識別碼 88888 僅限門市自取 O 欄。
 * 既有安裝型觸發器仍呼叫 paymentEditHandler2027，更新程式後無需重裝。
 * 只在原付款狀態為「未付款」時自動設「已付款」、補核帳日期及更新狀態。
 * 注意：88888 表示門市人員已實際收款，不是銀行匯款後五碼。
 */
/** 只需執行一次；重複執行會保留既有觸發器，不會重複安裝。 */
function setupPaymentEditTrigger2027(){
  const name='paymentEditHandler2027';
  const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===name);
  if(existing.length){
    const result={installed:false,existing:existing.length,message:'已存在付款編輯觸發器，未重複安裝'};
    Logger.log(JSON.stringify(result));return result;
  }
  ScriptApp.newTrigger(name).forSpreadsheet(SJ27.SPREADSHEET_ID).onEdit().create();
  const result={installed:true,message:'S-04 付款編輯觸發器已安裝（TEST）'};
  Logger.log(JSON.stringify(result));return result;
}

/** 若需暫停自動連動，執行此函式移除本功能的觸發器。 */
function removePaymentEditTrigger2027(){
  const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='paymentEditHandler2027');
  triggers.forEach(t=>ScriptApp.deleteTrigger(t));
  Logger.log(JSON.stringify({removed:triggers.length}));
}


function sjOverdueSafeDecision2027_(v,t,now){
  const id=String(v[0]||'').trim();
  const payment=String(v[t.pay-1]||'').trim();
  const status=String(v[t.status-1]||'').trim();
  if(!id)return {kind:'ignore'};
  if(!['未付款','逾期未付'].includes(payment)||status!=='待確認')
    return {kind:'protected',reason:'付款或訂單狀態不符'};
  const remittance=String(v[t.remit-1]||'').trim();
  if(remittance)return {kind:'protected',reason:'已有匯款資料，等待人工確認'};
  const created=v[1],deadline=v[t.deadline-1];
  const date=sjDateValue_(v[2]),pieces=Number(v[t.pieces-1]);
  if(!(created instanceof Date)||!Number.isFinite(created.getTime())||
     !(deadline instanceof Date)||!Number.isFinite(deadline.getTime())||
     !/^2027-(08|09)-\d{2}$/.test(date)||
     !Number.isSafeInteger(pieces)||pieces<=0||
     deadline.getTime()<created.getTime())
    return {kind:'warning',reason:'成立時間、截止時間、製作日期或顆數異常'};
  if(now<created.getTime()+SJ27.PAY_HOURS*3600000)
    return {kind:'protected',reason:'成立未滿72小時'};
  if(now<deadline.getTime())
    return {kind:'protected',reason:'付款期限尚未到'};
  return {kind:'candidate',date,pieces};
}
function applyOverdueAll2027(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('S-08 系統忙碌，未執行');
  try{
    const now=Date.now(),specs=[
      {name:SJ27.PICKUP,pay:14,status:17,deadline:13,remit:15,pieces:9,width:18},
      {name:SJ27.DELIVERY,pay:19,status:24,deadline:18,remit:20,pieces:12,width:25}
    ];
    const pending=[],warnings=[],protectedOrders=[],allIds=[];
    for(const t of specs){
      const sh=sjSheet_(t.name),last=sh.getLastRow();
      if(last<2)continue;
      const rows=sh.getRange(2,1,last-1,t.width).getValues();
      rows.forEach((v,i)=>{
        const id=String(v[0]||'').trim();
        if(!id)return;
        allIds.push(id);
        const d=sjOverdueSafeDecision2027_(v,t,now);
        if(d.kind==='warning')warnings.push({sheet:t.name,row:i+2,id,reason:d.reason});
        else if(d.kind==='protected')protectedOrders.push({sheet:t.name,id,reason:d.reason});
        else if(d.kind==='candidate')pending.push({sh,t,row:i+2,id,date:d.date,pieces:d.pieces});
      });
    }
    if(new Set(allIds).size!==allIds.length)
      warnings.push({reason:'跨表或同表訂單編號重複'});
    if(warnings.length){
      const result={mode:'ABORT NO WRITE',warnings,candidateCount:pending.length};
      Logger.log('===== S-08 安全中止 =====');
      Logger.log(JSON.stringify(result));return result;
    }
    const before={};
    pending.forEach(x=>{if(!(x.date in before))before[x.date]=sjUsage_(x.date).total;});
    for(const x of pending){
      x.sh.getRange(x.row,x.t.pay).setValue('逾期未付');
      x.sh.getRange(x.row,x.t.status).setValue('已取消');
    }
    if(pending.length){SpreadsheetApp.flush();refreshCapacitySummary2027_();}
    const releaseByDate={},capacityChecks={};
    for(const date of Object.keys(before)){
      const after=sjUsage_(date).total;
      const expected=pending.filter(x=>x.date===date).reduce((n,x)=>n+x.pieces,0);
      releaseByDate[date]=before[date]-after;
      capacityChecks[date]={before:before[date],after,released:releaseByDate[date],
        expected,pass:releaseByDate[date]===expected};
    }
    const result={mode:'WRITE OVERDUE S-08',cancelledCount:pending.length,
      cancelled:pending.map(x=>({sheet:x.t.name,row:x.row,id:x.id,pieces:x.pieces})),
      protectedCount:protectedOrders.length,releaseByDate,capacityChecks,
      passed:Object.values(capacityChecks).every(x=>x.pass),
      note:'逾期處理完成；請查看取消及產能驗證結果'};
    Logger.log('===== 逾期取消執行結果 =====');
    Logger.log(JSON.stringify(result));
    if(!result.passed)throw Error('S-08 產能驗證不符，請人工檢查');
    return result;
  }finally{lock.releaseLock();}
}

function previewOverdueSafety2027(){
  const specs=[
    {name:SJ27.PICKUP,pay:14,status:17,deadline:13,remit:15,pieces:9,width:18},
    {name:SJ27.DELIVERY,pay:19,status:24,deadline:18,remit:20,pieces:12,width:25}
  ];
  const now=Date.now(),candidates=[],warnings=[],protectedByReason={},scanned={};
  const allIds=[],releaseByDate={};
  for(const t of specs){
    const sh=sjSheet_(t.name),last=sh.getLastRow();
    const rows=last<2?[]:sh.getRange(2,1,last-1,t.width).getValues();
    scanned[t.name]=0;
    rows.forEach((v,i)=>{
      const id=String(v[0]||'').trim();
      if(!id)return;
      scanned[t.name]++;allIds.push(id);
      const d=sjOverdueSafeDecision2027_(v,t,now);
      if(d.kind==='warning')warnings.push({sheet:t.name,row:i+2,id,reason:d.reason});
      else if(d.kind==='protected')protectedByReason[d.reason]=(protectedByReason[d.reason]||0)+1;
      else if(d.kind==='candidate'){
        candidates.push({sheet:t.name,row:i+2,id,date:d.date,pieces:d.pieces});
        releaseByDate[d.date]=(releaseByDate[d.date]||0)+d.pieces;
      }
    });
  }
  if(new Set(allIds).size!==allIds.length)
    warnings.push({reason:'跨表或同表訂單編號重複'});
  const result={mode:'READ ONLY S-08B',checkedAt:Utilities.formatDate(new Date(),SJ27.TZ,'yyyy/MM/dd HH:mm:ss'),
    scanned,candidateCount:candidates.length,candidates,protectedByReason,
    warnings,releaseByDate,writePerformed:false,
    safeToApply:warnings.length===0,
    note:'這是預覽；safeToApply 不代表已完成所有正式上線測試'};
  Logger.log('===== S-08B 真實 TEST 訂單安全預覽｜READ ONLY =====');
  Logger.log(JSON.stringify(result));return result;
}


function overdueHourlyHandler2027(){
  if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4')
    throw Error('S-10B 非核准 TEST 試算表，拒絕執行');
  const result=applyOverdueAll2027();
  Logger.log('===== S-10B 自動逾期檢查結果 =====');
  Logger.log(JSON.stringify(result));
  if(result.mode==='ABORT NO WRITE'||!result.passed)
    throw Error('S-10B 安全檢查未通過，請查看執行紀錄');
  return result;
}
/** S-10B｜唯讀查看逾期排程，不新增或移除。 */
function inspectOverdueHourlyTrigger2027(){
  const ts=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='overdueHourlyHandler2027');
  const result={mode:'READ ONLY',handler:'overdueHourlyHandler2027',
    installedCount:ts.length,installed:ts.length===1,
    duplicate:ts.length>1,writePerformed:false};
  Logger.log('===== S-10B 逾期排程狀態 =====');
  Logger.log(JSON.stringify(result));return result;
}
/** S-10B｜需人工執行一次才會安裝，已存在則不重複安裝。 */
function setupOverdueHourlyTrigger2027(){
  if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4')
    throw Error('S-10B 非核准 TEST 試算表，拒絕安裝');
  const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='overdueHourlyHandler2027');
  if(existing.length){
    const result={installed:false,existing:existing.length,
      message:existing.length===1?'排程已存在，未重複安裝':'偵測到重複排程，請人工檢查'};
    Logger.log(JSON.stringify(result));return result;
  }
  const preview=previewOverdueSafety2027();
  if(preview.warnings.length||preview.candidateCount)
    throw Error('S-10B 安裝前預覽有警告或取消候選，請先人工檢查');
  ScriptApp.newTrigger('overdueHourlyHandler2027').timeBased().everyHours(1).create();
  const result={installed:true,frequency:'每小時',testOnly:true,
    note:'Apps Script 會在每小時區間內安排執行，並非整點精準執行'};
  Logger.log('===== S-10B 已安裝 TEST 自動逾期排程 =====');
  Logger.log(JSON.stringify(result));return result;
}
/** S-10B｜停止逾期自動檢查，不影響付款編輯觸發器。 */
function removeOverdueHourlyTrigger2027(){
  const ts=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='overdueHourlyHandler2027');
  ts.forEach(t=>ScriptApp.deleteTrigger(t));
  const result={removed:ts.length,handler:'overdueHourlyHandler2027',
    paymentEditTriggerUnaffected:true};
  Logger.log('===== S-10B 已移除 TEST 逾期排程 =====');
  Logger.log(JSON.stringify(result));return result;
}


/**
 * 自取匯款回報｜TEST 後端核心（尚未公開 HTTP 端點）。
 * 僅由受信任的伺服器端入口呼叫，勿直接開放匿名寫入。
 * 只更新門市自取 N/O 欄；P 核帳日期留給人工核帳。
 */
function submitPickupRemittance2027(data){
  if(!data||typeof data!=='object')throw Error('回報資料不完整');
  const orderId=String(data.orderId||'').trim().toUpperCase();
  const phone=sjPhone_(data.phone);
  const last5=String(data.last5||'').trim();
  if(!/^(?:SJ-\d{5}|SJ27-[0-9]{6}-[A-F0-9]{8})$/.test(orderId))throw Error('訂單編號格式不正確');
  if(!/^\d{5}$/.test(last5)||last5==='88888')throw Error('請填寫有效的匯款帳號後五碼');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('目前系統忙碌，請稍後再試');
  try{
    const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
    if(last<2)throw Error('無法核對訂單資料');
    const rows=sh.getRange(2,1,last-1,18).getValues();
    const matches=rows.map((v,i)=>({v,row:i+2})).filter(x=>String(x.v[0]||'').trim().toUpperCase()===orderId);
    if(matches.length!==1)throw Error('無法核對訂單資料');
    const hit=matches[0],v=hit.v;
    if(sjPhone_(v[5])!==phone)throw Error('無法核對訂單資料');
    const payment=String(v[13]||'').trim(),status=String(v[16]||'').trim();
    if(!sjStatusActive_(status))throw Error('訂單目前無法回報，請聯繫店家');
    if(payment==='待核帳'&&String(v[14]||'').trim()===last5)
      return {ok:true,alreadySubmitted:true,message:'這筆匯款資料已收到，請等待店家核帳'};
    if(payment!=='未付款')throw Error('訂單目前無法回報，請聯繫店家');
    if(String(v[14]||'').trim())throw Error('已有匯款資料，請聯繫店家確認');
    const now=Date.now(),created=v[1],deadline=v[12];
    if(!(created instanceof Date)||!(deadline instanceof Date)||
       !Number.isFinite(created.getTime())||!Number.isFinite(deadline.getTime())||
       now>=created.getTime()+SJ27.PAY_HOURS*3600000||now>=deadline.getTime())
      throw Error('付款回報期限已過，請聯繫店家');
    sh.getRange(hit.row,15).setNumberFormat('@').setValue(last5);
    sh.getRange(hit.row,14).setValue('待核帳');
    SpreadsheetApp.flush();
    Logger.log(JSON.stringify({event:'PICKUP_REMITTANCE_SUBMITTED',row:hit.row,orderId}));
    return {ok:true,alreadySubmitted:false,message:'已收到匯款回報，請等待店家核帳'};
  }finally{lock.releaseLock();}
}


/**
 * S-11A｜自取匯款回報安全檢查（唯讀）。
 * 不傳入客戶資料、不寫入任何欄位；確認欄位、測試環境與可回報候選數。
 */
function inspectPickupRemittance2027(){
  if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4')
    throw Error('非核准 TEST 試算表，拒絕執行 S-11A');
  const sh=sjSheet_(SJ27.PICKUP);
  const rows=sjDataRows_(SJ27.PICKUP);
  const now=Date.now();
  let eligible=0,notEligible=0,protectedRemittance=0;
  rows.forEach(v=>{
    const payment=String(v[13]||'').trim(),status=String(v[16]||'').trim();
    if(payment==='待核帳'&&String(v[14]||'').trim())protectedRemittance++;
    const validDates=v[1] instanceof Date&&v[12] instanceof Date;
    if(payment==='未付款'&&sjStatusActive_(status)&&!String(v[14]||'').trim()&&
       validDates&&now<v[1].getTime()+SJ27.PAY_HOURS*3600000&&now<v[12].getTime())eligible++;
    else notEligible++;
  });
  let reservedCodeRejected=false;
  try{submitPickupRemittance2027({orderId:'SJ27-261118-00000000',phone:'0912345678',last5:'88888'});}
  catch(e){reservedCodeRejected=String(e.message||e).includes('後五碼');}
  const result={mode:'READ ONLY S-11A',sheet:sh.getName(),pickupOrders:rows.length,
    eligible,notEligible,protectedRemittance,reservedCodeRejected,
    deliverySheetUntouched:true,writePerformed:false,
    publicEndpointEnabled:false};
  Logger.log('===== S-11A 自取匯款回報安全檢查 =====');
  Logger.log(JSON.stringify(result));return result;
}


/** S-11B｜只驗證匯款碼格式；不讀寫試算表、不提交訂單。 */
function testPickupRemittanceValidation2027(){
  const source=String(submitPickupRemittance2027);
  const checks={
    fiveDigitPatternFixed:source.includes('/^\\d{5}$/'),
    reserved88888Rejected:source.includes("last5==='88888'"),
    onlyPickupSheet:source.includes('sjSheet_(SJ27.PICKUP)'),
    setsPendingReview:source.includes("setValue('待核帳')"),
    postHandlerPresent:typeof doPost==='function'&&typeof doGet==='undefined'
  };
  const result={mode:'READ ONLY S-11B',checks,passed:Object.values(checks).every(Boolean),writePerformed:false};
  Logger.log('===== S-11B 自取回報程式安全檢查 =====');
  Logger.log(JSON.stringify(result));return result;
}


/**
 * S-11C｜建立專用自取匯款回報測試訂單並回報（僅 TEST）。
 * 使用正常建單流程；只操作本函式建立的訂單，不碰既有七筆。
 * 此函式會真的新增一筆 TEST 訂單，並將 N 設為待核帳、O 設為 12345。
 */
function testPickupRemittanceWrite2027(){
  if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4')
    throw Error('非核准 TEST 試算表，拒絕執行 S-11C');
  const date='2027-08-16',phone='0912345678',code='12345';
  const reqId='S11C'+Utilities.getUuid().replace(/-/g,'');
  const before=sjUsage_(date).total;
  if(before+48>SJ27.LIMIT)throw Error('測試日期產能不足，未建立訂單');
  const created=createOrder2027({method:'自取',name:'S11C測試訂單',phone,
    date,pickupTime:'13:00–15:00',sixBoxes:8,fifteenBoxes:0,
    notes:'S-11C 專用測試，請勿出貨',requestId:reqId});
  if(!created.ok||created.method!=='自取')throw Error('建立 TEST 訂單失敗');
  const first=submitPickupRemittance2027({orderId:created.orderId,phone,last5:code});
  const second=submitPickupRemittance2027({orderId:created.orderId,phone,last5:code});
  const sh=sjSheet_(SJ27.PICKUP),rows=sh.getRange(2,1,sh.getLastRow()-1,18).getValues();
  const matches=rows.filter(v=>String(v[0]||'')===created.orderId);
  if(matches.length!==1)throw Error('TEST 訂單查找異常');
  const v=matches[0];
  const checks={created:created.ok,pendingReview:String(v[13])==='待核帳',
    remittanceCode:String(v[14])===code,verifiedDateBlank:!v[15],
    orderStatusUnchanged:String(v[16])==='待確認',
    repeatIdempotent:second.ok===true&&second.alreadySubmitted===true,
    overdueProtected:sjOverdueSafeDecision2027_(v,
      {pay:14,status:17,deadline:13,remit:15,pieces:9},Date.now()+100*3600000).kind==='protected',
    capacityUnchanged:sjUsage_(date).total===before+48};
  const result={mode:'WRITE TEST S-11C',orderId:created.orderId,date,
    checks,passed:Object.values(checks).every(Boolean),
    addedTestOrder:true,writePerformed:true,
    note:'已新增一筆 TEST 自取訂單，保留供後續核帳測試；不會自動刪除'};
  Logger.log('===== S-11C 自取匯款回報實際寫入測試 =====');
  Logger.log(JSON.stringify(result));return result;
}


/**
 * S-11D｜僅限自取匯款回報的 Web App 入口（TEST）。
 * HTML form POST 跳轉至 Apps Script 回應頁，不使用跨網域 fetch。
 * 以 CacheService 對同一訂單、電話及全站請求限流；不回傳個資。
 * 注意：Apps Script 公開 Web App 無法可靠取得訪客 IP，這是基本限流，
 * 上線前仍須驗證防濫用方案與部署權限。
 */
function doPost(e){
  const p=e&&e.parameter||{};
  const isOrder=String(p.action||'')==='create-order';
  const respond=(message,success)=>HtmlService.createHtmlOutput(
    '<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8">'+
    '<meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<title>'+(isOrder?'TEST 訂單建立結果':'自取匯款回報結果')+'｜紀家烘焙坊</title></head>'+
    '<body style="font-family:serif;background:#faf6f0;color:#49352b;text-align:center;padding:50px 18px">'+
    '<main style="max-width:520px;margin:auto;background:#fffdfa;border:1px solid #e8dbcf;border-radius:16px;padding:28px">'+
    '<h2 style="color:#803d50">'+(isOrder?(success?'TEST 訂單建立成功':'TEST 訂單未成立'):(success?'匯款回報已收到':'匯款回報未完成'))+'</h2><p>'+
    message+'</p><p>您可以關閉此頁，返回紀家烘焙坊網站。</p></main></body></html>');
  if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4')
    return respond('目前無法受理，請聯繫店家。',false);
  if(String(p.action||'')==='create-order'&&String(p.website||'')==='stillji2027')return sj27PublicCreate_(p,respond);
  // 預設關閉；只有完成公開驗收後才可在 Script Properties 明確設為 ENABLED。
  if(PropertiesService.getScriptProperties().getProperty('SJ27_PUBLIC_REMITTANCE')!=='ENABLED')
    return respond('匯款回報服務尚未開放，請稍後再試。',false);
  if(String(p.action||'')!=='pickup-remittance'||String(p.website||'')!=='stillji2027')
    return respond('請由紀家烘焙坊自取匯款回報頁面送出。',false);
  if(String(p.company||'').trim())return respond('無法完成回報，請稍後再試。',false);
  const id=String(p.orderId||'').trim().toUpperCase(),phone=String(p.phone||'').trim();
  if(id.length>32||phone.length>30||String(p.last5||'').length>8)
    return respond('資料格式不正確，請返回檢查。',false);
  const cache=CacheService.getScriptCache();
  const key='sj27-remit-'+Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,id+'|'+phone)).slice(0,36);
  const attempts=Number(cache.get(key)||0);
  if(attempts>=5)return respond('嘗試次數過多，請稍後再試或聯繫店家。',false);
  cache.put(key,String(attempts+1),3600);
  try{
    const result=submitPickupRemittance2027({orderId:id,phone,last5:p.last5});
    return respond(result.message,true);
  }catch(err){
    Logger.log(JSON.stringify({event:'PICKUP_REMITTANCE_REJECTED',reason:String(err.message||err).slice(0,120)}));
    return respond('無法完成回報，請確認訂單資料及付款期限，或聯繫店家。',false);
  }
}
/** S-11D｜不發送請求的唯讀檢查。 */
function inspectPickupWebEndpoint2027(){
  const checks={testSpreadsheet:SJ27.SPREADSHEET_ID==='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4',
    postHandler:typeof doPost==='function',getHandlerNotEnabled:typeof doGet==='undefined',
    pickupOnly:doPost.toString().includes('pickup-remittance'),
    remittanceCore:typeof submitPickupRemittance2027==='function'};
  const result={mode:'READ ONLY S-11D',checks,passed:Object.values(checks).every(Boolean),
    writePerformed:false,note:'此檢查不代表 Web App 已部署或公開安全審查完成'};
  Logger.log(JSON.stringify(result));return result;
}

/** S-11F｜唯讀確認公開入口開關，絕不修改設定或訂單。 */
function inspectPublicRemittanceGate2027(){
  const value=PropertiesService.getScriptProperties().getProperty('SJ27_PUBLIC_REMITTANCE');
  const source=String(doPost);
  const checks={
    testSpreadsheet:SJ27.SPREADSHEET_ID==='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4',
    failClosed:source.includes("getProperty('SJ27_PUBLIC_REMITTANCE')!=='ENABLED'"),
    currentlyDisabled:value!=='ENABLED',
    noGetEndpoint:typeof doGet==='undefined'
  };
  const result={mode:'READ ONLY S-11F',checks,passed:Object.values(checks).every(Boolean),
    publicGate:value==='ENABLED'?'ENABLED':'DISABLED',
    writePerformed:false,
    reminder:'必須更新既有 Web App 部署版本，舊部署不會自動套用此開關'};
  Logger.log('===== S-11F 公開入口預設關閉檢查 =====');
  Logger.log(JSON.stringify(result));return result;
}


/**
 * S-12｜內部管理側欄（不新增 doGet，不向公開網站提供訂單資料）。
 * 由 Apps Script 編輯器手動執行 openOrderManager2027 開啟；
 * 若為獨立 Apps Script，請先確認 SpreadsheetApp.getUi() 是否可用。
 */
function managerSnapshot2027(){
  if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4')
    throw Error('管理端僅允許 TEST 試算表');
  const stamp=v=>v instanceof Date?Utilities.formatDate(v,SJ27.TZ,'yyyy/MM/dd HH:mm'):String(v==null?'':v);
  const pickup=sjDataRows_(SJ27.PICKUP).map(v=>({
    type:'自取',id:String(v[0]),date:sjDateValue_(v[2]),slot:stamp(v[3]),
    name:String(v[4]),phone:String(v[5]),quantity:'6入×'+v[6]+'、15入×'+v[7],
    pieces:Number(v[8])||0,amount:Number(v[11])||0,deadline:stamp(v[12]),
    payment:String(v[13]),remittance:String(v[14]||''),verified:stamp(v[15]),
    status:String(v[16]),note:String(v[17]||'')
  }));
  const delivery=sjDataRows_(SJ27.DELIVERY).map(v=>({
    type:'宅配',id:String(v[0]),date:sjDateValue_(v[2]),slot:'出貨',
    name:String(v[3]),phone:String(v[4]),address:String(v[5]),
    quantity:'4盒組×'+v[6]+'、10盒組×'+v[7]+'、15盒組×'+v[8]+'、15入×'+v[9],
    pieces:Number(v[11])||0,amount:Number(v[16])||0,deadline:stamp(v[17]),
    payment:String(v[18]),remittance:String(v[19]||''),verified:stamp(v[20]),
    slipConfirmation:String(v[21]||''),tracking:String(v[22]||''),
    status:String(v[23]),note:String(v[24]||'')
  }));
  const map={};
  pickup.concat(delivery).forEach(o=>{
    if(!sjStatusActive_(o.status))return;
    const d=map[o.date]||(map[o.date]={date:o.date,pickup:0,delivery:0});
    if(o.type==='自取')d.pickup+=o.pieces;else d.delivery+=o.pieces;
  });
  const capacity=Object.values(map).sort((a,b)=>a.date.localeCompare(b.date)).map(d=>({
    ...d,total:d.pickup+d.delivery,pans:Math.ceil((d.pickup+d.delivery)/60),
    remaining:Math.max(0,SJ27.LIMIT-d.pickup-d.delivery),limit:SJ27.LIMIT
  }));
  return {generatedAt:stamp(new Date()),pickup,delivery,capacity,
    counts:{pickup:pickup.length,delivery:delivery.length,
      pendingReview:pickup.concat(delivery).filter(o=>o.payment==='待核帳').length,
      unpaid:pickup.concat(delivery).filter(o=>o.payment==='未付款'&&sjStatusActive_(o.status)).length}};
}
function openOrderManager2027(){
  const html=HtmlService.createHtmlOutputFromFile('Manager2027')
    .setTitle('紀家烘焙坊｜2027 TEST 管理');
  SpreadsheetApp.getUi().showSidebar(html);
}
/** S-12 唯讀煙霧測試：無新增或修改資料。 */
function inspectOrderManager2027(){
  const data=managerSnapshot2027();
  const result={mode:'READ ONLY S-12',counts:data.counts,capacityDates:data.capacity.length,
    hasPickup:Array.isArray(data.pickup),hasDelivery:Array.isArray(data.delivery),
    writePerformed:false};
  Logger.log(JSON.stringify(result));return result;
}


/** 2027 TEST 網站訂單入口。預設關閉；獨立於匯款回報開關。 */
function sj27PublicCreate_(p,respond){
 if(SJ27.SPREADSHEET_ID!=='1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4'||PropertiesService.getScriptProperties().getProperty('SJ27_PUBLIC_ORDERS')!=='ENABLED')
   return respond('TEST 收單尚未啟用，請聯繫管理人員。',false);
 if(String(p.company||'').trim())return respond('資料未通過驗證。',false);
 const safe=(v,n)=>String(v==null?'':v).trim().slice(0,n);
 const count=(v,label)=>{const n=Number(v||0);if(!Number.isSafeInteger(n)||n<0||n>1000)throw Error(label+'數量不正確');return n;};
 const date=safe(p.date,10),kind=safe(p.kind,10),name=safe(p.name,60),phone=safe(p.phone,30),address=safe(p.address,200);
 const rid=safe(p.requestId,100);
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(20000))return respond('目前訂單較多，請稍後重試。',false);
 try{
  if(!/^[A-Za-z0-9_-]{16,100}$/.test(rid)||!/^2027-(08|09)-\d{2}$/.test(date)||date>'2027-09-15'||new Date(date+'T00:00:00Z').toISOString().slice(0,10)!==date||!name||!/^(09\d{8}|0[2-8]\d{7,8}(?:#\d{1,6})?)$/.test(phone.replace(/[\s()－-]/g,'')))throw Error('姓名、電話、日期或識別碼不正確');
  if(!['pickup','delivery'].includes(kind))throw Error('訂單類型不正確');
  const ss=SpreadsheetApp.openById(SJ27.SPREADSHEET_ID),sh=ss.getSheetByName(kind==='pickup'?'門市自取':'宅配到府'),width=kind==='pickup'?18:25;
  if(!sh)throw Error('訂單工作表不存在');
  const tag='[WEB:'+rid+']';
  if(sh.getLastRow()>1){const rows=sh.getRange(2,1,sh.getLastRow()-1,width).getDisplayValues(),found=rows.find(v=>String(v[width-1]).includes(tag));if(found)return respond('訂單已成立，編號：'+found[0]+'（重複送出未新增）',true);}
  const now=new Date(),deadline=new Date(now.getTime()+72*3600000),id=sj27ShortOrderId_(ss),weekday=new Date(date+'T00:00:00Z').getUTCDay();
  let pieces,values;
  if(kind==='pickup'){
   const a=count(p.six,'6入'),b=count(p.fifteen,'15入'),slot=safe(p.slot,20);
   pieces=a*6+b*15;
   if(!pieces||(a&&!b&&a<8)||(b&&!a&&b<3)||(a&&b&&pieces<42))throw Error('自取未達最低訂購量');
   const allowed=date==='2027-09-15'?['09:00–10:00']:[0,6].includes(weekday)?['09:00–12:00','13:00–15:00','15:00–17:00','17:00–18:00']:['13:00–15:00','15:00–17:00','17:00–18:00'];
   if(!allowed.includes(slot))throw Error('取貨時段不正確');
   const goods=a*300+b*750;
   values=[id,now,date,slot,name,phone,a,b,pieces,goods,'一般售價',goods,deadline,'未付款','','','待確認',safe(p.notes,250)+' '+tag];
  }else{
   if([0,6].includes(weekday))throw Error('週末無法出貨');
   if(address.length<8)throw Error('請填寫完整收件地址');
   const four=count(p.four,'4盒組'),ten=count(p.ten,'10盒組'),group=count(p.group,'15盒組'),fifteen=count(p.fifteen,'15入'),bags=count(p.bags,'提袋');
   if(fifteen>0&&fifteen<4||bags>0&&!group)throw Error('宅配商品或提袋數量不符規則');
   pieces=(four*4+ten*10+group*15)*6+fifteen*15;if(!pieces)throw Error('請選擇商品');
   const split=Array.from({length:group},(_,i)=>Math.floor(bags/group)+(i<bags%group?1:0));
   const goods=(four*4+ten*10+group*15)*300+fifteen*750,bagFee=bags*20,cartons15=Math.ceil(fifteen/10),shipping=four*150+ten*170+split.reduce((n,x)=>n+(x>=5?210:170),0)+cartons15*210,remote=/^(宜蘭縣|宜蘭市|花蓮縣|臺東縣|台東縣)/.test(address)?(four+ten+group+cartons15)*10:0;
   const invoice=safe(p.invoice,10),tax=safe(p.taxId,8),arrival=safe(p.arrival,10);
   if(!['normal','business'].includes(invoice)||invoice==='business'&&!/^\d{8}$/.test(tax)||!['any','before13','14to18'].includes(arrival))throw Error('發票或到貨時段資料不正確');
   values=[id,now,date,name,phone,address,four,ten,group,fifteen,bags,pieces,goods,bagFee,shipping,remote,goods+bagFee+shipping+remote,deadline,'未付款','','','待回傳','','待確認','發票:'+invoice+(invoice==='business'?'/統編:'+tax:'')+' 到貨:'+arrival+' '+tag];
  }
  if(pieces>4800)throw Error('超過每日 4800 顆產能上限');
  let used=0;
  [['門市自取',3,9,17],['宅配到府',3,12,24]].forEach(([n,dcol,pcol,scol])=>{const sheet=ss.getSheetByName(n);if(!sheet||sheet.getLastRow()<2)return;sheet.getRange(2,1,sheet.getLastRow()-1,scol).getValues().forEach(v=>{const d=v[dcol-1] instanceof Date?Utilities.formatDate(v[dcol-1],'Asia/Taipei','yyyy-MM-dd'):String(v[dcol-1]);if(d===date&&!['已取消','取消','作廢'].includes(String(v[scol-1])))used+=Number(v[pcol-1])||0;});});
  if(used+pieces>4800)throw Error('當日剩餘產能不足');
  const row=sjNextRow_(sh);sh.getRange(row,kind==='pickup'?6:5).setNumberFormat('@');sh.getRange(row,1,1,width).setValues([values]);
  return respond('TEST 訂單建立成功！訂單編號：'+id+'。請於 72 小時內完成匯款並通知店家。',true);
 }catch(err){Logger.log('PUBLIC_TEST_ORDER_REJECTED '+String(err));return respond('訂單未成立：'+String(err.message||err).replace(/[<>&"']/g,''),false);}
 finally{lock.releaseLock();}
}
