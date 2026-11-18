/**
 * 紀家烘焙坊｜116 年中秋 TEST 後端（僅限 Apps Script 編輯器內部執行）
 * 對應現有 Google 試算表：門市自取／宅配到府／每日產能統計。
 * 不含 doPost / doGet，不會公開接單。請勿將任何金鑰放進 GitHub。
 */
const SJ27 = Object.freeze({
  SPREADSHEET_ID: '1kL2ZO0O-rNoR6UsepheMV84_9YtY7fdk4UFKDhMt4W4',
  PICKUP: '門市自取', DELIVERY: '宅配到府', SUMMARY: '每日產能統計',
  LIMIT: 4800, TZ: 'Asia/Taipei', PAY_HOURS: 72,
  PICKUP_HEADERS: ['訂單編號','成立時間','取貨日期','取貨時段','提貨人','聯絡電話','6入盒數','15入盒數','蛋黃酥顆數','商品金額','優惠類型','應付總額','付款截止時間','付款狀態','匯款後五碼','核帳日期','訂單狀態','備註'],
  DELIVERY_HEADERS: ['訂單編號','成立時間','出貨日期','收件人','聯絡電話','收件地址','6入4盒組','6入10盒組','6入15盒組','15入盒數','提袋數','蛋黃酥顆數','商品金額','提袋費','基本運費','偏遠加價','應付總額','付款截止時間','付款狀態','匯款後五碼','核帳日期','宅配單確認','黑貓單號','訂單狀態','備註'],
  SUMMARY_HEADERS: ['製作日期','自取顆數','宅配顆數','合計顆數','換算盤數','剩餘顆數','每日上限','狀態']
});
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
    const orderId='SJ27-'+Utilities.formatDate(created,SJ27.TZ,'yyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
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
/** 首次驗證請先執行這個，只讀取資料，不新增訂單。 */
function testInspect2027(){return inspect2027();}
/** 確認唯讀檢查通過後，才手動執行這個測試新增一筆自取訂單。 */
function testCreatePickup2027(){
  return createOrder2027({method:'自取',name:'測試客人',phone:'0912345678',date:'2027-09-01',pickupTime:'13:00–15:00',sixBoxes:8,fifteenBoxes:0,notes:'TEST 測試資料',requestId:'TEST-PICKUP-2027-0001'});
}
/** 確認唯讀檢查通過後，才手動執行這個測試新增一筆宅配訂單。 */
function testCreateDelivery2027(){
  return createOrder2027({method:'宅配',name:'測試客人',phone:'0912345678',address:'彰化縣彰化市三民路402-2號',date:'2027-09-02',four:1,ten:0,fifteenGroup:0,fifteenBoxes:0,bags:0,notes:'TEST 測試資料',requestId:'TEST-DELIVERY-2027-0001'});
}

/**
 * 付款管理 S-02：唯讀逾期報告。
 * 不取消訂單、不改付款狀態、不寄 LINE、不釋放產能。
 * 待核帳、已付款、退款中等情況全部交由人工複核。
 */
/**
 * S-02｜付款管理安全稽核（READ ONLY）。
 * 只列出逾期候選與待核帳訂單；不更改任何欄位、不取消訂單。
 * 使用付款截止時間 Date 儲存值；只有「未付款」且訂單仍有效才是候選。
 */
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
/** S-02：第一次先執行此函式，完全不寫入試算表。 */
function testPaymentAudit2027(){return paymentAudit2027();}

/**
 * S-03｜同一張訂單表付款連動（預覽模式）。
 * 僅讀取：當付款狀態為「已付款」時，建議補上核帳日期；
 * 自取訂單狀態「待付款／待確認」改「已確認」；
 * 宅配訂單狀態「待付款／待確認」改「待出貨」。
 * 不改已取消或已進入其他流程的訂單，不操作黑貓單號。
 * 注意：核帳日期代表程式確認時間，不代表銀行實際入帳時間。
 */
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
function testPreviewPaymentSync2027(){return previewPaymentSync2027();}

/**
 * S-03｜手動確認後才執行寫入。
 * 再次核對付款狀態與訂單編號，避免預覽後狀態變動誤寫。
 * 不處理逾期取消，也不改動宅配單確認及黑貓單號。
 */
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

/**
 * S-05｜72 小時逾期取消安全預覽（READ ONLY）。
 * 僅「未付款」且已過截止時間、訂單尚有效者可列為候選。
 * 「待核帳／已付款／退款中／已退款」一律不取消。
 * 有效日期必須是 Sheets 日期值；文字或缺值會列警告，不會猜測。
 * 目前不取消、不修改付款狀態、不釋放產能、不建立定時觸發器。
 */
function previewOverdueCancellation2027(){
  const now=new Date();
  const result={
    mode:'READ ONLY',checkedAt:Utilities.formatDate(now,SJ27.TZ,'yyyy/MM/dd HH:mm:ss'),
    candidates:[],protected:[],warnings:[],counts:{pickup:0,delivery:0}
  };
  const specs=[
    {name:SJ27.PICKUP,key:'pickup',deadline:13,payment:14,status:17,pieces:9},
    {name:SJ27.DELIVERY,key:'delivery',deadline:18,payment:19,status:24,pieces:12}
  ];
  for(const spec of specs){
    const sh=sjSheet_(spec.name),last=sh.getLastRow();
    if(last<2)continue;
    const rows=sh.getRange(2,1,last-1,spec.status).getValues();
    for(let i=0;i<rows.length;i++){
      const row=rows[i],orderId=String(row[0]||'').trim();
      if(!orderId)continue;
      result.counts[spec.key]++;
      const payment=String(row[spec.payment-1]||'').trim();
      const status=String(row[spec.status-1]||'').trim();
      if(!sjStatusActive_(status))continue;
      if(['待核帳','已付款','退款中','已退款'].includes(payment)){
        result.protected.push({sheet:spec.name,row:i+2,orderId,payment});
        continue;
      }
      if(payment!=='未付款'){
        result.warnings.push({sheet:spec.name,row:i+2,orderId,reason:'非預期付款狀態，需人工確認',payment});
        continue;
      }
      const deadline=row[spec.deadline-1];
      if(!(deadline instanceof Date)||!Number.isFinite(deadline.getTime())){
        result.warnings.push({sheet:spec.name,row:i+2,orderId,reason:'付款截止時間不是有效日期，略過'});
        continue;
      }
      if(deadline.getTime()>=now.getTime())continue;
      const pieces=Number(row[spec.pieces-1]);
      if(!Number.isSafeInteger(pieces)||pieces<0){
        result.warnings.push({sheet:spec.name,row:i+2,orderId,reason:'蛋黃酥顆數異常，略過'});
        continue;
      }
      result.candidates.push({
        sheet:spec.name,row:i+2,orderId,payment,status,
        deadline:Utilities.formatDate(deadline,SJ27.TZ,'yyyy/MM/dd HH:mm'),
        pieces,proposedPayment:'逾期未付',proposedStatus:'已取消'
      });
    }
  }
  result.candidateCount=result.candidates.length;
  Logger.log('===== S-05 逾期取消預覽｜READ ONLY =====');
  Logger.log(JSON.stringify(result));
  return result;
}
function testPreviewOverdueCancellation2027(){
  return previewOverdueCancellation2027();
}

/**
 * S-05B｜純模擬測試，不接觸試算表、不建立或取消真實訂單。
 * 用固定案例確認 72 小時邊界、付款保護、取消冪等性及產能計算。
 * 正式寫入功能尚未啟用。
 */
function testOverdueRulesSimulation2027(){
  const now=new Date('2026-11-18T10:00:00+08:00').getTime();
  const cases=[
    {name:'逾期未付款',pay:'未付款',status:'待確認',hoursAgo:73,pieces:48,expected:true},
    {name:'剛好到期',pay:'未付款',status:'待確認',hoursAgo:0,pieces:24,expected:false},
    {name:'尚未到期',pay:'未付款',status:'待確認',hoursAgo:-1,pieces:24,expected:false},
    {name:'逾期待核帳',pay:'待核帳',status:'待確認',hoursAgo:100,pieces:48,expected:false},
    {name:'逾期已付款',pay:'已付款',status:'已確認',hoursAgo:100,pieces:48,expected:false},
    {name:'已取消不可重複',pay:'未付款',status:'已取消',hoursAgo:100,pieces:48,expected:false},
    {name:'退款中保護',pay:'退款中',status:'待確認',hoursAgo:100,pieces:48,expected:false}
  ];
  const results=cases.map(c=>{
    const deadline=now-c.hoursAgo*3600000;
    const eligible=c.pay==='未付款'&&sjStatusActive_(c.status)&&deadline<now;
    return {case:c.name,expected:c.expected,actual:eligible,pass:eligible===c.expected};
  });
  const usedBefore=72,cancelledPieces=48,usedAfter=usedBefore-cancelledPieces;
  const capacityPass=usedAfter===24&&SJ27.LIMIT-usedAfter===4776;
  const result={mode:'SIMULATION ONLY',passed:results.every(r=>r.pass)&&capacityPass,
    results,capacity:{usedBefore,cancelledPieces,usedAfter,remaining:SJ27.LIMIT-usedAfter,pass:capacityPass},
    note:'未讀寫試算表；正式取消與產能重算尚待另行驗證'};
  Logger.log('===== S-05B 逾期規則純模擬｜NO SHEET ACCESS =====');
  Logger.log(JSON.stringify(result));return result;
}

/**
 * S-05C｜建立專用逾期測試訂單（需人工執行；會寫入 TEST）。
 * 以唯一 RID 標記，只會建立一次；將付款截止時間設為昨天，模擬逾期。
 * 不影響原有兩筆已付款訂單；尚未執行取消。
 */
function setupOverdueTestOrder2027(){
  const rid='TEST-S05C-OVERDUE-2027-0001';
  const existing=sjDuplicate_(rid);
  if(existing){
    Logger.log(JSON.stringify({mode:'SETUP',existing:true,orderId:existing.orderId,message:'測試訂單已存在，不重複建立或修改'}));
    return existing;
  }
  const created=createOrder2027({
    method:'自取',name:'S05C逾期測試',phone:'0912345678',
    date:'2027-09-03',pickupTime:'13:00–15:00',
    sixBoxes:8,fifteenBoxes:0,notes:'S05C 專用逾期取消測試',
    requestId:rid
  });
  const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
  const idx=ids.findIndex(v=>v[0]===created.orderId);
  if(idx<0)throw Error('找不到剛建立的測試訂單，請勿繼續');
  const row=idx+2;
  // 測試專用：將截止時間調整到 24 小時前；正式訂單仍維持建立後 72 小時。
  sh.getRange(row,13).setValue(new Date(Date.now()-24*3600000)).setNumberFormat('yyyy/mm/dd hh:mm');
  const result={mode:'SETUP WRITE TEST ONLY',orderId:created.orderId,row,
    sheet:SJ27.PICKUP,pieces:48,deadline:'昨天（僅此測試單）'};
  Logger.log('===== S-05C 建立專用逾期測試單 =====');
  Logger.log(JSON.stringify(result));return result;
}

/**
 * S-05C｜只取消專用 RID 的測試訂單。
 * 重新確認 ID、RID、未付款、有效截止日期、已逾期及狀態；
 * 取消後重新彙整每日產能，不刪除訂單或更動其他訂單。
 * 不建立定時觸發器，絕不批次取消一般訂單。
 */
function applyOverdueTestCancellation2027(){
  const rid='[RID:TEST-S05C-OVERDUE-2027-0001]';
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌，未取消');
  try{
    const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
    if(last<2)throw Error('沒有測試訂單');
    const rows=sh.getRange(2,1,last-1,18).getValues();
    const matches=rows.map((v,i)=>({v,row:i+2})).filter(x=>String(x.v[17]||'').includes(rid));
    if(matches.length!==1)throw Error('S05C 專用測試單數量不為 1，已停止');
    const {v,row}=matches[0],orderId=String(v[0]||'').trim();
    const payment=String(v[13]||'').trim(),status=String(v[16]||'').trim();
    const deadline=v[12],pieces=Number(v[8]);
    if(!['未付款','逾期未付'].includes(payment)||!['待確認','待付款'].includes(status)||
       !(deadline instanceof Date)||!Number.isFinite(deadline.getTime())||
       deadline.getTime()>=Date.now()||pieces!==48)
      throw Error('付款、狀態、期限或數量不符合測試條件；未取消');
    const date=sjDateValue_(v[2]),before=sjUsage_(date);
    sh.getRange(row,14).setValue('逾期未付');
    sh.getRange(row,17).setValue('已取消');
    SpreadsheetApp.flush();
    const summaryRows=refreshCapacitySummary2027_();
    const after=sjUsage_(date);
    const result={mode:'WRITE TEST ONLY',sheet:SJ27.PICKUP,row,orderId,
      payment:'逾期未付',status:'已取消',date,
      capacityBefore:before.total,capacityAfter:after.total,
      released:before.total-after.total,expectedReleased:pieces,
      capacityMatches:before.total-after.total===pieces,summaryRows};
    Logger.log('===== S-05C 專用測試單取消與產能釋放 =====');
    Logger.log(JSON.stringify(result));
    if(!result.capacityMatches)throw Error('取消後產能差額不符，請檢查日產能統計');
    return result;
  }finally{lock.releaseLock();}
}

/**
 * S-06A｜批次取消規則預演（READ ONLY）。
 * 讀取兩張既有訂單表，核對已逾期的未付款／逾期未付訂單。
 * 只有「待確認／待付款」可進入候選；已付款、待核帳、退款中、已取消皆受保護。
 * 每筆候選計算可釋放顆數，依日期彙整，但完全不寫入、不裝定時觸發器。
 */
function previewBatchOverdue2027(){
  const now=new Date(),result={
    mode:'READ ONLY',checkedAt:Utilities.formatDate(now,SJ27.TZ,'yyyy/MM/dd HH:mm:ss'),
    candidates:[],protectedCount:0,warnings:[],releaseByDate:{},
    scanned:{pickup:0,delivery:0}
  };
  const specs=[
    {name:SJ27.PICKUP,key:'pickup',pay:14,status:17,deadline:13,pieces:9,date:3},
    {name:SJ27.DELIVERY,key:'delivery',pay:19,status:24,deadline:18,pieces:12,date:3}
  ];
  for(const spec of specs){
    const sh=sjSheet_(spec.name),last=sh.getLastRow();
    if(last<2)continue;
    const rows=sh.getRange(2,1,last-1,spec.status).getValues();
    for(let i=0;i<rows.length;i++){
      const v=rows[i],orderId=String(v[0]||'').trim();
      if(!orderId)continue;
      result.scanned[spec.key]++;
      const pay=String(v[spec.pay-1]||'').trim();
      const status=String(v[spec.status-1]||'').trim();
      if(!sjStatusActive_(status)){result.protectedCount++;continue;}
      if(['已付款','待核帳','退款中','已退款'].includes(pay)){
        result.protectedCount++;continue;
      }
      if(!['未付款','逾期未付'].includes(pay)){
        result.warnings.push({sheet:spec.name,row:i+2,orderId,reason:'不認識的付款狀態',pay});continue;
      }
      if(!['待確認','待付款'].includes(status)){
        result.protectedCount++;continue;
      }
      const deadline=v[spec.deadline-1],pieces=Number(v[spec.pieces-1]);
      const date=sjDateValue_(v[spec.date-1]);
      if(!(deadline instanceof Date)||!Number.isFinite(deadline.getTime())||
         !Number.isSafeInteger(pieces)||pieces<=0||!/^2027-(08|09)-\d{2}$/.test(date)){
        result.warnings.push({sheet:spec.name,row:i+2,orderId,reason:'期限、顆數或製作日期異常'});continue;
      }
      if(deadline.getTime()>=now.getTime())continue;
      const item={sheet:spec.name,row:i+2,orderId,payment:pay,status,
        deadline:Utilities.formatDate(deadline,SJ27.TZ,'yyyy/MM/dd HH:mm'),
        date,pieces,proposedPayment:'逾期未付',proposedStatus:'已取消'};
      result.candidates.push(item);
      result.releaseByDate[date]=(result.releaseByDate[date]||0)+pieces;
    }
  }
  result.candidateCount=result.candidates.length;
  Logger.log('===== S-06A 批次逾期取消預演｜READ ONLY =====');
  Logger.log(JSON.stringify(result));return result;
}
function testPreviewBatchOverdue2027(){return previewBatchOverdue2027();}

/**
 * S-06B｜批次取消純模擬：不讀寫試算表、不安裝觸發器。
 * 模擬兩張表、多筆逾期、待核帳保護及重複執行。
 */
function testBatchOverdueSimulation2027(){
  const now=Date.now();
  const cases=[
    {id:'P1',sheet:'門市自取',date:'2027-09-03',pieces:48,pay:'未付款',status:'待確認',hoursAgo:1},
    {id:'D1',sheet:'宅配到府',date:'2027-09-03',pieces:24,pay:'逾期未付',status:'待確認',hoursAgo:10},
    {id:'P2',sheet:'門市自取',date:'2027-09-04',pieces:60,pay:'未付款',status:'待確認',hoursAgo:2},
    {id:'P3',sheet:'門市自取',date:'2027-09-03',pieces:30,pay:'已付款',status:'已確認',hoursAgo:8},
    {id:'D2',sheet:'宅配到府',date:'2027-09-03',pieces:24,pay:'待核帳',status:'待確認',hoursAgo:8},
    {id:'D3',sheet:'宅配到府',date:'2027-09-03',pieces:24,pay:'未付款',status:'待確認',hoursAgo:-1},
    {id:'P4',sheet:'門市自取',date:'2027-09-04',pieces:48,pay:'逾期未付',status:'已取消',hoursAgo:8},
    {id:'D4',sheet:'宅配到府',date:'2027-09-04',pieces:24,pay:'退款中',status:'待確認',hoursAgo:8}
  ];
  const eligible=x=>['未付款','逾期未付'].includes(x.pay)&&
    ['待確認','待付款'].includes(x.status)&&
    now-x.hoursAgo*3600000<now;
  function pass(){
    const cancelled=[],released={};
    for(const x of cases){
      if(!eligible(x))continue;
      cancelled.push(x.id);
      released[x.date]=(released[x.date]||0)+x.pieces;
      x.pay='逾期未付';x.status='已取消';
    }
    return {cancelled,released};
  }
  const first=pass(),second=pass();
  const expected=['P1','D1','P2'];
  const passed=JSON.stringify(first.cancelled)===JSON.stringify(expected)&&
    first.released['2027-09-03']===72&&
    first.released['2027-09-04']===60&&
    second.cancelled.length===0&&Object.keys(second.released).length===0&&
    cases.find(x=>x.id==='P3').status==='已確認'&&
    cases.find(x=>x.id==='D2').pay==='待核帳';
  const result={mode:'SIMULATION ONLY',passed,first,second,
    protectedCases:cases.filter(x=>!expected.includes(x.id)).map(x=>({id:x.id,pay:x.pay,status:x.status})),
    note:'僅記憶體模擬；尚未驗證批次實際寫入與觸發器'};
  Logger.log('===== S-06B 批次取消及重複執行純模擬 =====');
  Logger.log(JSON.stringify(result));return result;
}

/**
 * S-06C 第一步：建立兩筆專用 TEST 逾期訂單，供後續實際批次取消驗證。
 * 僅手動執行；RID 去重；不取消任何訂單，不建立觸發器。
 * 一筆自取（48 顆）、一筆宅配（24 顆），同一出貨／取貨日。
 */
function setupBatchOverdueTestOrders2027(){
  const specs=[
    {method:'自取',rid:'TEST-S06C-PICKUP-2027-0001',name:'S06C自取測試',
      date:'2027-09-06',pickupTime:'13:00–15:00',sixBoxes:8,fifteenBoxes:0},
    {method:'宅配',rid:'TEST-S06C-DELIVERY-2027-0001',name:'S06C宅配測試',
      date:'2027-09-06',address:'彰化縣彰化市中山路一段1號',four:1,ten:0,fifteenGroup:0,fifteenBoxes:0,bags:0}
  ];
  const results=[];
  for(const spec of specs){
    const existing=sjDuplicate_(spec.rid);
    if(existing){
      results.push({method:spec.method,existing:true,orderId:existing.orderId,
        note:'既有測試單不重新修改期限'});continue;
    }
    const created=createOrder2027(Object.assign({},spec,{
      phone:'0912345678',requestId:spec.rid,notes:'S06C 專用批次取消測試'
    }));
    const name=spec.method==='自取'?SJ27.PICKUP:SJ27.DELIVERY;
    const sh=sjSheet_(name),last=sh.getLastRow();
    const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
    const idx=ids.findIndex(v=>v[0]===created.orderId);
    if(idx<0)throw Error('建立後找不到 '+created.orderId+'，請檢查表格');
    const row=idx+2,deadlineCol=spec.method==='自取'?13:18;
    sh.getRange(row,deadlineCol).setValue(new Date(Date.now()-24*3600000))
      .setNumberFormat('yyyy/mm/dd hh:mm');
    results.push({method:spec.method,orderId:created.orderId,row,
      pieces:created.pieces,date:spec.date,deadline:'昨天（僅測試單）'});
  }
  Logger.log('===== S-06C 建立批次專用測試訂單 =====');
  Logger.log(JSON.stringify({mode:'SETUP WRITE TEST ONLY',results,
    note:'尚未取消任何訂單；請先檢查兩筆資料及 READ ONLY 預演'}));
  return results;
}

/**
 * S-06C 第二階段｜僅取消兩筆指定 RID 的 TEST 訂單。
 * 不處理其他逾期訂單；每次重新讀取最新狀態，避免重複取消。
 * 若任一筆不符合條件則整批停止，不先寫入。
 */
function applyBatchOverdueTest2027(){
  const targets=[
    {sheet:SJ27.PICKUP,rid:'[RID:TEST-S06C-PICKUP-2027-0001]',id:'SJ27-261118-0021D4E8',
      paymentCol:14,statusCol:17,deadlineCol:13,piecesCol:9,expectedPieces:48},
    {sheet:SJ27.DELIVERY,rid:'[RID:TEST-S06C-DELIVERY-2027-0001]',id:'SJ27-261118-400CCA66',
      paymentCol:19,statusCol:24,deadlineCol:18,piecesCol:12,expectedPieces:24}
  ];
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('S-06C 系統忙碌，未執行');
  try{
    const now=Date.now(),prepared=[],alreadyCancelled=[];
    for(const t of targets){
      const sh=sjSheet_(t.sheet),last=sh.getLastRow();
      if(last<2)throw Error('找不到測試訂單：'+t.sheet);
      const width=t.sheet===SJ27.PICKUP?18:25;
      const rows=sh.getRange(2,1,last-1,width).getValues();
      const found=rows.map((v,i)=>({v,row:i+2}))
        .filter(x=>String(x.v[width-1]||'').includes(t.rid));
      if(found.length!==1)throw Error('測試單 RID 不唯一或不存在：'+t.sheet);
      const {v,row}=found[0];
      const id=String(v[0]||'').trim(),payment=String(v[t.paymentCol-1]||'').trim();
      const status=String(v[t.statusCol-1]||'').trim();
      const deadline=v[t.deadlineCol-1],pieces=Number(v[t.piecesCol-1]);
      const date=sjDateValue_(v[2]);
      if(id!==t.id||pieces!==t.expectedPieces||date!=='2027-09-06')
        throw Error('測試單識別、數量或日期不符，全部停止：'+t.sheet);
      if(payment==='逾期未付'&&status==='已取消'){
        alreadyCancelled.push({sheet:t.sheet,id,row});continue;
      }
      if(!['未付款','逾期未付'].includes(payment)||
         !['待確認','待付款'].includes(status)||
         !(deadline instanceof Date)||!Number.isFinite(deadline.getTime())||
         deadline.getTime()>=now)
        throw Error('測試單付款、狀態或截止時間不符，全部停止：'+t.sheet);
      prepared.push({sh,t,row,id,pieces,date});
    }
    if(prepared.length!==0&&prepared.length!==2)
      throw Error('偵測到部分測試單已取消、部分未取消，為避免部分寫入，請人工檢查');
    const before=sjUsage_('2027-09-06').total;
    if(prepared.length===2){
      for(const x of prepared){
        x.sh.getRange(x.row,x.t.paymentCol).setValue('逾期未付');
        x.sh.getRange(x.row,x.t.statusCol).setValue('已取消');
      }
      SpreadsheetApp.flush();
      refreshCapacitySummary2027_();
    }
    const after=sjUsage_('2027-09-06').total;
    const released=before-after;
    const result={mode:'WRITE TEST ONLY',cancelled:prepared.map(x=>({sheet:x.t.sheet,id:x.id,pieces:x.pieces})),
      alreadyCancelled,capacityBefore:before,capacityAfter:after,released,
      expectedReleased:prepared.reduce((sum,x)=>sum+x.pieces,0),
      capacityMatches:released===prepared.reduce((sum,x)=>sum+x.pieces,0)};
    Logger.log('===== S-06C 專用兩筆批次取消測試 =====');
    Logger.log(JSON.stringify(result));
    if(!result.capacityMatches)throw Error('產能差額不符，請檢查 TEST 統計表');
    return result;
  }finally{lock.releaseLock();}
}

/**
 * S-07｜通用逾期取消（手動執行；尚未安裝定時觸發器）。
 * 安全設計：
 * 1. ScriptLock 下重新讀取當前資料；不依賴先前預覽結果。
 * 2. 只處理付款狀態「未付款／逾期未付」且訂單「待確認」。
 * 3. 若任何訂單資料異常，整批停止、不寫入。
 * 4. 已付款、待核帳、退款、已取消、其他非待確認狀態不處理。
 * 5. 完成後從有效訂單重算產能；不刪除歷史資料。
 * 注意：試算表寫入不是交易；若中途失敗，重新執行前須人工檢查。
 */
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
    const result={mode:'MANUAL WRITE S-08',cancelledCount:pending.length,
      cancelled:pending.map(x=>({sheet:x.t.name,row:x.row,id:x.id,pieces:x.pieces})),
      protectedCount:protectedOrders.length,releaseByDate,capacityChecks,
      passed:Object.values(capacityChecks).every(x=>x.pass),
      note:'尚未安裝定時觸發器；須先完成 S-08 測試'};
    Logger.log('===== S-08 通用逾期取消（手動） =====');
    Logger.log(JSON.stringify(result));
    if(!result.passed)throw Error('S-08 產能驗證不符，請人工檢查');
    return result;
  }finally{lock.releaseLock();}
}
/** S-08A｜純記憶體安全規則測試：不讀寫試算表。 */
function testOverdueSafetyRules2027(){
  const now=new Date('2026-11-18T11:00:00+08:00').getTime();
  const t={pay:14,status:17,deadline:13,remit:15,pieces:9};
  function row(ageHours,deadlineHoursAgo,payment,status,remit){
    const v=Array(18).fill('');
    v[0]='TEST-S08';v[1]=new Date(now-ageHours*3600000);
    v[2]='2027-09-08';v[8]=48;
    v[12]=new Date(now-deadlineHoursAgo*3600000);
    v[13]=payment;v[14]=remit;v[16]=status;return v;
  }
  const cases=[
    {name:'未滿72小時且截止時間已過但早於成立（資料異常）',
      v:row(1,24,'未付款','待確認',''),want:'warning'},
    {name:'滿72小時且期限已過',v:row(96,24,'未付款','待確認',''),want:'candidate'},
    {name:'已填匯款後五碼',v:row(96,24,'未付款','待確認','12345'),want:'protected'},
    {name:'待核帳保護',v:row(96,24,'待核帳','待確認',''),want:'protected'},
    {name:'已付款保護',v:row(96,24,'已付款','已確認',''),want:'protected'},
    {name:'期限延長尚未到',v:row(96,-24,'未付款','待確認',''),want:'protected'},
    {name:'已取消不重複',v:row(96,24,'逾期未付','已取消',''),want:'protected'},
    {name:'成立時間缺失中止',v:row(96,24,'未付款','待確認',''),want:'warning'},
    {name:'未滿72小時且正常截止時間未到',v:row(1,-71,'未付款','待確認',''),want:'protected'},
    {name:'成立滿72小時但截止時間尚未到',v:row(80,-2,'未付款','待確認',''),want:'protected'}
  ];
  cases[7].v[1]='';
  const results=cases.map(x=>{
    const got=sjOverdueSafeDecision2027_(x.v,t,now).kind;
    return {name:x.name,want:x.want,got,pass:got===x.want};
  });
  const result={mode:'PURE SIMULATION NO SHEET ACCESS',
    passed:results.every(x=>x.pass),passedCount:results.filter(x=>x.pass).length,
    total:results.length,results};
  Logger.log('===== S-08A 修正版 72小時與匯款保護測試 =====');
  Logger.log(JSON.stringify(result));return result;
}

/**
 * S-07A｜建立唯一專用逾期測試單，供通用 applyOverdueAll2027 驗證。
 * 只建立一筆 48 顆門市自取測試單；不執行取消、不安裝排程。
 * RID 去重：若已建立，絕不再次修改其截止時間。
 */
function setupGenericOverdueTest2027(){
  const rid='TEST-S07A-GENERIC-2027-0001';
  const existing=sjDuplicate_(rid);
  if(existing){
    const result={mode:'SETUP NO WRITE',existing:true,orderId:existing.orderId};
    Logger.log('===== S-07A 測試單已存在 =====');
    Logger.log(JSON.stringify(result));return result;
  }
  const created=createOrder2027({
    method:'自取',name:'S07A通用取消測試',phone:'0912345678',
    date:'2027-09-07',pickupTime:'13:00–15:00',
    sixBoxes:8,fifteenBoxes:0,
    notes:'S07A 專用通用逾期取消測試',requestId:rid
  });
  const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
  const idx=ids.findIndex(x=>x[0]===created.orderId);
  if(idx<0)throw Error('建立後找不到測試單，請停止');
  const row=idx+2;
  sh.getRange(row,13).setValue(new Date(Date.now()-24*3600000))
    .setNumberFormat('yyyy/mm/dd hh:mm');
  const result={mode:'SETUP WRITE TEST ONLY',orderId:created.orderId,
    sheet:SJ27.PICKUP,row,date:'2027-09-07',pieces:48,
    deadline:'昨天（僅此測試單）',note:'尚未取消；下一步先 READ ONLY 預覽'};
  Logger.log('===== S-07A 建立通用逾期專用測試單 =====');
  Logger.log(JSON.stringify(result));return result;
}

/** S-08B｜真實 TEST 資料安全預覽，完全唯讀，使用與正式取消相同的判斷函式。 */
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

/** S-08C｜只建立兩筆專用 TEST 保護單，不取消、不安裝排程。 */
function setupOverdueProtectionFixtures2027(){
  const cases=[
    {rid:'TEST-S08C-YOUNG-2027-0001',name:'S08C未滿72小時',date:'2027-09-09',kind:'young'},
    {rid:'TEST-S08C-REMIT-2027-0001',name:'S08C已填匯款資料',date:'2027-09-10',kind:'remit'}
  ];
  const results=[];
  for(const x of cases){
    let existing=sjDuplicate_(x.rid);
    if(!existing){
      existing=createOrder2027({method:'自取',name:x.name,phone:'0912345678',
        date:x.date,pickupTime:'13:00–15:00',sixBoxes:8,fifteenBoxes:0,
        notes:'S08C 安全保護專用測試',requestId:x.rid});
    }
    const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
    const rows=sh.getRange(2,1,last-1,18).getValues();
    const hits=rows.map((v,i)=>({v,row:i+2})).filter(y=>
      String(y.v[0]||'')===existing.orderId&&
      String(y.v[17]||'').includes('[RID:'+x.rid+']'));
    if(hits.length!==1)throw Error('S08C 測試單識別不唯一，請停止：'+x.rid);
    const h=hits[0],v=h.v;
    if(String(v[13])!=='未付款'||String(v[16])!=='待確認')
      throw Error('S08C 測試單狀態已變動，請停止：'+x.rid);
    // 只在新建立的測試單上做一次設定；重跑不覆寫人工修改。
    if(!existing.duplicate){
      if(x.kind==='remit'){
        const created=new Date(Date.now()-96*3600000);
        sh.getRange(h.row,2).setValue(created).setNumberFormat('yyyy/mm/dd hh:mm');
        sh.getRange(h.row,13).setValue(new Date(created.getTime()+72*3600000))
          .setNumberFormat('yyyy/mm/dd hh:mm');
        sh.getRange(h.row,15).setNumberFormat('@').setValue('12345');
      }
    }
    results.push({orderId:existing.orderId,row:h.row,kind:x.kind,
      createdNew:!existing.duplicate,expectedProtection:x.kind==='young'?'成立未滿72小時':'已有匯款資料，等待人工確認'});
  }
  const result={mode:'SETUP TEST ONLY',sheet:SJ27.PICKUP,orders:results,
    note:'只建立或尋找兩筆專用測試單；請下一步執行 previewOverdueSafety2027 唯讀確認'};
  Logger.log('===== S-08C 建立安全保護測試單 =====');
  Logger.log(JSON.stringify(result));return result;
}

/** S-09A｜建立唯一的真實逾期候選 TEST 單；重跑只回報、不重設。 */
function setupOverdue72hCandidate2027(){
  const rid='TEST-S09A-72H-CANCEL-2027-0001';
  let found=sjDuplicate_(rid),isNew=false;
  if(!found){
    found=createOrder2027({method:'自取',name:'S09A逾期取消測試',
      phone:'0912345678',date:'2027-09-13',pickupTime:'13:00–15:00',
      sixBoxes:8,fifteenBoxes:0,notes:'S09A 專用取消驗收測試單',
      requestId:rid});
    isNew=true;
  }
  const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
  const rows=sh.getRange(2,1,last-1,18).getValues();
  const hits=rows.map((v,i)=>({v,row:i+2})).filter(x=>
    String(x.v[0]||'')===found.orderId&&
    String(x.v[17]||'').includes('[RID:'+rid+']'));
  if(hits.length!==1)throw Error('S09A 測試單識別不唯一，停止');
  const h=hits[0];
  if(isNew){
    if(String(h.v[13])!=='未付款'||String(h.v[16])!=='待確認'||String(h.v[14]||'').trim())
      throw Error('S09A 測試單初始狀態異常，停止');
    const created=new Date(Date.now()-96*3600000);
    const deadline=new Date(created.getTime()+72*3600000);
    sh.getRange(h.row,2).setValue(created).setNumberFormat('yyyy/mm/dd hh:mm');
    sh.getRange(h.row,13).setValue(deadline).setNumberFormat('yyyy/mm/dd hh:mm');
    SpreadsheetApp.flush();
  }
  const current=sh.getRange(h.row,1,1,18).getValues()[0];
  const decision=sjOverdueSafeDecision2027_(current,
    {pay:14,status:17,deadline:13,remit:15,pieces:9},Date.now());
  const result={mode:'S09A SETUP TEST ONLY',orderId:found.orderId,
    sheet:SJ27.PICKUP,row:h.row,createdNew:isNew,
    date:sjDateValue_(current[2]),pieces:current[8],
    payment:current[13],status:current[16],hasRemittance:!!String(current[14]||'').trim(),
    decision:decision.kind,reason:decision.reason||null,
    note:'尚未取消；下一步請執行 previewOverdueSafety2027（READ ONLY）'};
  Logger.log('===== S-09A 建立滿72小時逾期候選 TEST 單 =====');
  Logger.log(JSON.stringify(result));return result;
}

/** S-10A｜取消後重複執行安全性 + 排程狀態；完全唯讀。 */
function inspectOverdueSchedulerSafety2027(){
  const orderId='SJ27-261118-31901CB0';
  const sh=sjSheet_(SJ27.PICKUP),last=sh.getLastRow();
  const rows=last<2?[]:sh.getRange(2,1,last-1,18).getValues();
  const hits=rows.map((v,i)=>({v,row:i+2})).filter(x=>String(x.v[0]||'').trim()===orderId);
  const warnings=[];
  if(hits.length!==1)warnings.push('S-09 測試訂單不存在或重複');
  let testOrder=null;
  if(hits.length===1){
    const x=hits[0],v=x.v;
    const decision=sjOverdueSafeDecision2027_(v,
      {pay:14,status:17,deadline:13,remit:15,pieces:9},Date.now());
    testOrder={orderId,row:x.row,payment:v[13],status:v[16],pieces:v[8],
      date:sjDateValue_(v[2]),decision:decision.kind,
      excludedFromRepeatCancellation:decision.kind!=='candidate'};
    if(v[13]!=='逾期未付'||v[16]!=='已取消'||decision.kind!=='protected')
      warnings.push('S-09 測試單取消後狀態不符');
  }
  const preview=previewOverdueSafety2027();
  const triggers=ScriptApp.getProjectTriggers();
  const overdue=triggers.filter(t=>t.getHandlerFunction()==='applyOverdueAll2027');
  const result={mode:'READ ONLY S-10A',testOrder,
    candidateCount:preview.candidateCount,candidates:preview.candidates,
    warnings:warnings.concat(preview.warnings),
    overdueTriggerCount:overdue.length,
    triggerAlreadyInstalled:overdue.length>0,
    passed:warnings.length===0&&preview.warnings.length===0&&
      preview.candidateCount===0&&overdue.length===0,
    writePerformed:false,
    note:'本函式不執行取消，不安裝排程；passed 僅代表此時資料符合驗收預期'};
  Logger.log('===== S-10A 重複取消與排程狀態唯讀驗收 =====');
  Logger.log(JSON.stringify(result));return result;
}
