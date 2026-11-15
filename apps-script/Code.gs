/**
 * STILL JI 2027 Mid-Autumn - Apps Script backend (TEST template)
 * Standalone Google Spreadsheet; one worksheet: 訂單總表; one order per row.
 * Script Properties: SPREADSHEET_ID. Deploy as Web App only after testing.
 * IMPORTANT: Do not deploy as public write API until abuse controls are in place.
 */
const CONFIG = Object.freeze({
  SHEET: '訂單總表',
  CAPACITY: 4800,
  TZ: 'Asia/Taipei',
  HEADERS: ['訂單編號','建立時間','訂購方式','姓名','電話','製作日期','出貨日期','取貨日期','商品明細JSON','6入盒數','15入盒數','蛋黃酥顆數','盤數(進位)','商品金額','提袋金額','運費','偏遠加價','應付總額','付款狀態','訂單狀態','備註','黑貓單號','客戶請求識別碼']
});
function setup2027(){
  const ss=SpreadsheetApp.openById(requiredProperty_('SPREADSHEET_ID'));
  let sh=ss.getSheetByName(CONFIG.SHEET);
  if(!sh)sh=ss.insertSheet(CONFIG.SHEET);
  if(sh.getLastRow()===0)sh.appendRow(CONFIG.HEADERS);
  else {
    const h=sh.getRange(1,1,1,CONFIG.HEADERS.length).getDisplayValues()[0];
    if(h.join('|')!==CONFIG.HEADERS.join('|'))throw Error('表頭與 2027 範本不一致，已停止以避免覆寫既有資料');
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,CONFIG.HEADERS.length).setFontWeight('bold');
  return '2027 訂單總表已準備完成';
}
function requiredProperty_(key){
  const value=PropertiesService.getScriptProperties().getProperty(key);
  if(!value)throw Error('缺少 Script Property: '+key);
  return value;
}
function sheet_(){
  const sh=SpreadsheetApp.openById(requiredProperty_('SPREADSHEET_ID')).getSheetByName(CONFIG.SHEET);
  if(!sh)throw Error('請先執行 setup2027');
  const h=sh.getRange(1,1,1,CONFIG.HEADERS.length).getDisplayValues()[0];
  if(h.join('|')!==CONFIG.HEADERS.join('|'))throw Error('訂單總表表頭不符');
  return sh;
}
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
  const detail=JSON.stringify({four,ten,fifteenGroup:group,bags,bagSplit,remote});
  return {method,name,phone,date,six,fifteen,pieces,goods,bagAmount,shipping:shippingSix+shipping15,remoteFee,total:goods+bagAmount+shippingSix+shipping15+remoteFee,detail,requestId:text_(data.requestId,100),notes:text_(data.notes,300)};
}
function activeStatus_(status){return !['已取消','取消','作廢','退款完成'].includes(String(status).trim());}
function usage_(sh,date){
  if(sh.getLastRow()<2)return 0;
  const rows=sh.getRange(2,1,sh.getLastRow()-1,CONFIG.HEADERS.length).getValues();
  return rows.reduce((sum,r)=>sum+(String(r[5]).trim()===date&&activeStatus_(r[19])?Number(r[11])||0:0),0);
}
/** Called only from trusted server-side order creation flow. */
function createOrder2027(data){
  const order=normalizedOrder_(data);
  if(!order.requestId||order.requestId.length<12)throw Error('缺少有效的請求識別碼');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌，請稍後重試');
  try{
    const sh=sheet_();
    // Idempotency: repeated submits return existing order without consuming capacity twice.
    if(sh.getLastRow()>1){
      const rows=sh.getRange(2,1,sh.getLastRow()-1,CONFIG.HEADERS.length).getValues();
      const existing=rows.find(r=>String(r[22])===order.requestId);
      if(existing)return {ok:true,orderId:String(existing[0]),duplicate:true};
    }
    const used=usage_(sh,order.date);
    if(used+order.pieces>CONFIG.CAPACITY)throw Error('該日期剩餘產能不足：剩餘 '+Math.max(0,CONFIG.CAPACITY-used)+' 顆');
    const id='SJ27-'+Utilities.formatDate(new Date(),CONFIG.TZ,'yyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
    sh.appendRow([id,new Date(),order.method,order.name,order.phone,order.date,order.method==='宅配'?order.date:'',order.method==='自取'?order.date:'',order.detail,order.six,order.fifteen,order.pieces,Math.ceil(order.pieces/60),order.goods,order.bagAmount,order.shipping,order.remoteFee,order.total,'未付款','待確認',order.notes,'',order.requestId]);
    SpreadsheetApp.flush();
    return {ok:true,orderId:id,productionDate:order.date,pieces:order.pieces,total:order.total};
  }finally{lock.releaseLock();}
}
/** Read-only internal capacity view; do not expose publicly without controls. */
function capacityForDate2027(date){
  validDate_(date);
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('系統忙碌');
  try{const used=usage_(sheet_(),date);return {date,limit:CONFIG.CAPACITY,used,remaining:Math.max(0,CONFIG.CAPACITY-used),usedPans:Math.ceil(used/60)};}
  finally{lock.releaseLock();}
}
