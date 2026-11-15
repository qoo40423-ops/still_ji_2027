/**
 * STILL JI 2027 Mid-Autumn - Apps Script backend (TEST template)
 * Standalone Google Spreadsheet; 自取訂單 / 宅配訂單 / 每日產能統計.
 * Script Properties: SPREADSHEET_ID. Deploy as Web App only after testing.
 * IMPORTANT: Do not deploy as public write API until abuse controls are in place.
 */
const CONFIG = Object.freeze({
  PICKUP: '自取訂單',
  DELIVERY: '宅配訂單',
  SUMMARY: '每日產能統計',
  CAPACITY: 4800,
  TZ: 'Asia/Taipei',
  HEADERS: ['訂單編號','建立時間','訂購方式','姓名','電話','製作日期','出貨日期','取貨日期','商品明細JSON','6入盒數','15入盒數','蛋黃酥顆數','盤數(進位)','商品金額','提袋金額','運費','偏遠加價','應付總額','付款狀態','訂單狀態','備註','黑貓單號','客戶請求識別碼']
});
function setup2027(){
  const ss=SpreadsheetApp.openById(requiredProperty_('SPREADSHEET_ID'));
  for(const name of [CONFIG.PICKUP,CONFIG.DELIVERY]){
    let sh=ss.getSheetByName(name);
    if(!sh)sh=ss.insertSheet(name);
    if(sh.getLastRow()===0)sh.appendRow(CONFIG.HEADERS);
    const h=sh.getRange(1,1,1,CONFIG.HEADERS.length).getDisplayValues()[0];
    if(h.join('|')!==CONFIG.HEADERS.join('|'))throw Error(name+' 表頭不一致，為避免覆寫資料已停止');
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,CONFIG.HEADERS.length).setFontWeight('bold');
  }
  let summary=ss.getSheetByName(CONFIG.SUMMARY);
  if(!summary)summary=ss.insertSheet(CONFIG.SUMMARY);
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌');
  try{refreshCapacitySummary_();}finally{lock.releaseLock();}
  return '自取訂單、宅配訂單、每日產能統計已建立';
}
function requiredProperty_(key){
  const value=PropertiesService.getScriptProperties().getProperty(key);
  if(!value)throw Error('缺少 Script Property: '+key);
  return value;
}
function sheet_(method){
  const name=method==='自取'?CONFIG.PICKUP:method==='宅配'?CONFIG.DELIVERY:null;
  if(!name)throw Error('訂購方式錯誤');
  const sh=SpreadsheetApp.openById(requiredProperty_('SPREADSHEET_ID')).getSheetByName(name);
  if(!sh)throw Error('請先執行 setup2027：'+name);
  const h=sh.getRange(1,1,1,CONFIG.HEADERS.length).getDisplayValues()[0];
  if(h.join('|')!==CONFIG.HEADERS.join('|'))throw Error(name+' 表頭不符');
  return sh;
}
function dateCell_(value){
  if(value instanceof Date)return Utilities.formatDate(value,CONFIG.TZ,'yyyy-MM-dd');
  return String(value==null?'':value).trim();
}
function orderRows_(sh){
  return sh.getLastRow()<2?[]:sh.getRange(2,1,sh.getLastRow()-1,CONFIG.HEADERS.length).getValues();
}
function allRows_(){return orderRows_(sheet_('自取')).concat(orderRows_(sheet_('宅配')));}
function n_(x,name){const v=Number(x);if(!Number.isSafeInteger(v)||v<0)throw Error(name+'必須為非負整數');return v;}
function text_(x,max){const s=String(x==null?'':x).trim();if(s.length>max)throw Error('輸入字數超過限制');return s;}
function validDate_(s){
  if(!/^2027-\d{2}-\d{2}$/.test(s))throw Error('日期格式錯誤');
  const d=new Date(s+'T00:00:00Z');
  if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==s)throw Error('日期無效');
  return d;
}
function taiwanToday_(){return Utilities.formatDate(new Date(),CONFIG.TZ,'yyyy-MM-dd');}
function weekday_(date){return validDate_(date).getUTCDay();}
function normalizedOrder_(data){
  if(!data||typeof data!=='object')throw Error('缺少訂單內容');
  const method=text_(data.method,12);
  if(method!=='自取'&&method!=='宅配')throw Error('訂購方式錯誤');
  const name=text_(data.name,60),phone=text_(data.phone,30);
  if(!name||!phone)throw Error('姓名與電話不可空白');
  const date=text_(data.productionDate,10);
  validDate_(date);
  if(date<taiwanToday_())throw Error('不能選擇過去的日期');
  const six=n_(data.sixBoxes||0,'6入盒數'),fifteen=n_(data.fifteenBoxes||0,'15入盒數');
  const four=n_(data.four||0,'4盒組'),ten=n_(data.ten||0,'10盒組'),group=n_(data.fifteenGroup||0,'15盒組');
  const bags=n_(data.bags||0,'提袋數');
  if(method==='宅配'){
    if(weekday_(date)===0||weekday_(date)===6)throw Error('宅配週六、週日不可出貨');
    if(four>0&&date>'2027-08-31')throw Error('6入4盒組超過最晚出貨日');
    if(group>0&&date>'2027-09-10')throw Error('6入15盒組超過最晚出貨日');
    if(six!==four*4+ten*10+group*15)throw Error('6入宅配盒數與組合不一致');
    if(fifteen>0&&fifteen<4)throw Error('15入宅配至少4盒');
    if(bags>0&&group===0)throw Error('目前提袋僅支援6入15盒組');
  }else{
    if(four||ten||group)throw Error('自取不可填宅配組合');
    if(bags)throw Error('自取贈袋尚未定案，暫不接受提袋數');
    const pickupTime=text_(data.pickupTime,10);
    if(pickupTime!=='上午'&&pickupTime!=='下午')throw Error('請選擇提貨時段');
    if(date==='2027-09-15'&&pickupTime!=='上午')throw Error('中秋節當天僅開放上午10點前提貨');
    if((six>0&&fifteen===0&&six<8)||(fifteen>0&&six===0&&fifteen<3)||(six>0&&fifteen>0&&six*6+fifteen*15<42))throw Error('未達訂單最小數量標準，僅能現場購買，敬請見諒');
  }
  const pieces=six*6+fifteen*15;
  if(!pieces||pieces>CONFIG.CAPACITY)throw Error('訂單顆數必須介於1至4800顆');
  const bagSplit=group?Array.from({length:group},(_,i)=>Math.floor(bags/group)+(i<bags%group?1:0)):[];
  const shippingSix=method==='宅配'?four*150+ten*170+bagSplit.reduce((a,b)=>a+(b>=5?210:170),0):0;
  const cartons15=method==='宅配'?Math.ceil(fifteen/10):0;
  const shipping15=cartons15*210;
  const remote=method==='宅配'&&data.remote===true;
  const remoteFee=remote?(four+ten+group+cartons15)*10:0;
  const goods=six*300+fifteen*750;
  const bagAmount=bags*20;
  const detail=JSON.stringify({four,ten,fifteenGroup:group,bags,bagSplit,remote,pickupTime:method==='自取'?text_(data.pickupTime,10):''});
  return {method,name,phone,date,six,fifteen,pieces,goods,bagAmount,shipping:shippingSix+shipping15,remoteFee,total:goods+bagAmount+shippingSix+shipping15+remoteFee,detail,requestId:text_(data.requestId,100),notes:text_(data.notes,300)};
}
function activeStatus_(status){return !['已取消','取消','作廢','退款完成'].includes(String(status).trim());}
function usage_(date){
  return allRows_().reduce((sum,r)=>sum+(dateCell_(r[5])===date&&activeStatus_(r[19])?Number(r[11])||0:0),0);
}
function refreshCapacitySummary_(){
  const ss=SpreadsheetApp.openById(requiredProperty_('SPREADSHEET_ID'));
  const sh=ss.getSheetByName(CONFIG.SUMMARY);
  if(!sh)throw Error('請先執行 setup2027');
  const map={};
  for(const [method,rows] of [['自取',orderRows_(sheet_('自取'))],['宅配',orderRows_(sheet_('宅配'))]]){
    for(const r of rows){
      const date=dateCell_(r[5]);
      if(!/^2027-\d{2}-\d{2}$/.test(date)||!activeStatus_(r[19]))continue;
      const pieces=Number(r[11])||0;
      if(!map[date])map[date]={pickup:0,delivery:0};
      map[date][method==='自取'?'pickup':'delivery']+=pieces;
    }
  }
  const values=[['製作日期','自取顆數','宅配顆數','已使用顆數','已使用盤數(進位)','剩餘顆數','每日上限','狀態']];
  Object.keys(map).sort().forEach(date=>{
    const p=map[date].pickup,d=map[date].delivery,total=p+d;
    values.push([date,p,d,total,Math.ceil(total/60),Math.max(0,CONFIG.CAPACITY-total),CONFIG.CAPACITY,total>CONFIG.CAPACITY?'超量':total===CONFIG.CAPACITY?'額滿':'可接單']);
  });
  sh.clearContents();
  sh.getRange(1,1,values.length,values[0].length).setValues(values);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,values[0].length).setFontWeight('bold');
  return values.length-1;
}
/** Manual refresh after editing a status, date or quantity in the spreadsheet. */
function refreshCapacitySummary2027(){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌');
  try{return refreshCapacitySummary_();}finally{lock.releaseLock();}
}
/** Called only from trusted server-side order creation flow. */
function createOrder2027(data){
  const order=normalizedOrder_(data);
  if(!order.requestId||order.requestId.length<12)throw Error('缺少有效的請求識別碼');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌，請稍後重試');
  try{
    const sh=sheet_(order.method);
    // Idempotency: repeated submits return existing order without consuming capacity twice.
    const existing=allRows_().find(r=>String(r[22])===order.requestId);
    if(existing)return {ok:true,orderId:String(existing[0]),duplicate:true};
    const used=usage_(order.date);
    if(used+order.pieces>CONFIG.CAPACITY)throw Error('該日期剩餘產能不足：剩餘 '+Math.max(0,CONFIG.CAPACITY-used)+' 顆');
    const id='SJ27-'+Utilities.formatDate(new Date(),CONFIG.TZ,'yyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
    sh.appendRow([id,new Date(),order.method,order.name,order.phone,order.date,order.method==='宅配'?order.date:'',order.method==='自取'?order.date:'',order.detail,order.six,order.fifteen,order.pieces,Math.ceil(order.pieces/60),order.goods,order.bagAmount,order.shipping,order.remoteFee,order.total,'未付款','待確認',order.notes,'',order.requestId]);
    SpreadsheetApp.flush();
    refreshCapacitySummary_();
    return {ok:true,orderId:id,productionDate:order.date,pieces:order.pieces,total:order.total};
  }finally{lock.releaseLock();}
}
/** Read-only internal capacity view; do not expose publicly without controls. */
function capacityForDate2027(date){
  validDate_(date);
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌');
  try{const used=usage_(date);return {date,limit:CONFIG.CAPACITY,used,remaining:Math.max(0,CONFIG.CAPACITY-used),usedPans:Math.ceil(used/60)};}
  finally{lock.releaseLock();}
}
