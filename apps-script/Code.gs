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
      values=[orderId,created,date,time,name,phone,six,fifteen,pieces,goods,'一般售價',goods,deadline,'未付款','','','待付款',notes+' [RID:'+rid+']'];
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
      values=[orderId,created,date,name,phone,address,four,ten,group,fifteen,bags,pieces,goods,bagFee,shipping,remoteFee,goods+bagFee+shipping+remoteFee,deadline,'未付款','','','待回傳宅配單','','待付款',notes+' [RID:'+rid+']'];
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
