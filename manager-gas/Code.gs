/** 與公開 TEST 訂單共用編號空間，避免跨表重複。 */
function mgrShortOrderId_(ss){
 const used=new Set();
 for(const name of ['門市自取','宅配到府']){
  const sh=ss.getSheetByName(name);
  if(sh&&sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().forEach(r=>used.add(String(r[0]).trim().toUpperCase()));
 }
 for(let n=0;n<200;n++){
  const id='SJ-'+String(Math.floor(Math.random()*90000)+10000);
  if(!used.has(id))return id;
 }
 throw Error('暫時無法產生唯一訂單編號，請稍後重試');
}
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
 const make=date=>map[date]||(map[date]={date,morning:empty(),afternoon:empty(),pickup:empty(),delivery:empty(),slots:{'09:00~10:00':empty(),'09:00~12:00':empty(),'13:00~15:00':empty(),'15:00~17:00':empty(),'17:00~18:00':empty()},unclassified:empty()});
 const active=o=>!['已取消','取消','作廢'].includes(o.status);
 const slotKey=slot=>{
  const t=String(slot||'').replace(/\s/g,'').replace(/[～〜－–—]/g,'~').replace(/：/g,':');
  for(const k of ['09:00~10:00','09:00~12:00','13:00~15:00','15:00~17:00','17:00~18:00'])if(t.includes(k))return k;
  if(t==='上午')return '09:00~12:00';
  return '';
 };

 pickup.filter(active).forEach(o=>{const d=make(o.date);add(d.pickup,o);const key=slotKey(o.slot);if(key){add(d.slots[key],o);add(key.startsWith('09:')?d.morning:d.afternoon,o);}else add(d.unclassified,o);});
 delivery.filter(active).forEach(o=>{const d=make(o.date);add(d.delivery,o);add(d.morning,o);});
 const vipSh=ss.getSheetByName('老闆VIP公關單');
 const vip=vipSh&&vipSh.getLastRow()>1?vipSh.getRange(2,1,vipSh.getLastRow()-1,12).getValues().filter(v=>String(v[0]||'').trim()).map(v=>({id:String(v[0]),date:day(v[2]),shift:String(v[3]),method:String(v[4]),name:String(v[5]),boxes6:Number(v[6])||0,boxes15:Number(v[7])||0,pieces:Number(v[8])||0,status:String(v[10]),note:String(v[11]||'')})):[];
 vip.filter(o=>o.status!=='已取消').forEach(o=>{const d=make(o.date);if(!d.vip)d.vip=empty();add(d.vip,o);add(o.shift==='上午'?d.morning:d.afternoon,o);});
 const capacity=Object.values(map).sort((a,b)=>a.date.localeCompare(b.date)).map(d=>{Object.values(d.slots).forEach(finish);['morning','afternoon','pickup','delivery','unclassified'].forEach(k=>finish(d[k]));const total=d.pickup.pieces+d.delivery.pieces+(d.vip?d.vip.pieces:0);return {...d,total,pans:Math.ceil(total/60),remaining:Math.max(0,4800-total),limit:4800};});
 return {generatedAt:stamp(new Date()),pickup,delivery,vip,capacity,counts:{pickup:pickup.length,delivery:delivery.length,pendingReview:pickup.concat(delivery).filter(o=>o.payment==='待核帳').length,unpaid:pickup.concat(delivery).filter(o=>o.payment==='未付款'&&!['已取消','取消','作廢'].includes(o.status)).length}};
}


/** 僅允許已核帳的自取訂單標記已取貨；操作僅更新 Q 欄。 */
function markPickupCollected2027(orderId){
 const id=String(orderId||'').trim();
 if(!/^(?:SJ-\d{5}|SJ27-[A-Z0-9-]{8,30})$/.test(id))throw Error('訂單編號格式不正確');
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

/** 私人管理 App 寫入：只供獨立私人 GAS 專案使用。 */
const MGR_VIP_HEADERS=['VIP單號','建立時間','製作日期','製作時段','交付方式','對象／單位','6入盒數','15入盒數','顆數','聯絡／地址','狀態','備註'];
function mgrInt_(v,label){const n=Number(v);if(!Number.isSafeInteger(n)||n<0)throw Error(label+'必須是非負整數');return n;}
function mgrText_(v,max,label){const s=String(v==null?'':v).trim();if(s.length>max)throw Error(label+'字數過長');return s;}
function mgrDate_(v){const s=String(v||'');if(!/^2027-(08|09)-\d{2}$/.test(s)||s>'2027-09-15'||new Date(s+'T00:00:00Z').toISOString().slice(0,10)!==s)throw Error('請選擇有效的 2027/8/1～9/15 日期');return s;}
function mgrPhone_(v){const s=mgrText_(v,30,'電話').replace(/[\s()－-]/g,'');if(!/^(09\d{8}|0[2-8]\d{7,8}(?:#\d{1,6})?)$/.test(s))throw Error('請填寫有效電話');return s;}
function mgrVipSheet_(ss){let sh=ss.getSheetByName('老闆VIP公關單');if(!sh){sh=ss.insertSheet('老闆VIP公關單');sh.getRange(1,1,1,12).setValues([MGR_VIP_HEADERS]);sh.setFrozenRows(1);}else{const heads=sh.getRange(1,1,1,12).getDisplayValues()[0];if(heads.some((x,i)=>x!==MGR_VIP_HEADERS[i]))throw Error('VIP 工作表表頭不符，已停止寫入');}return sh;}
function mgrUsage_(ss,date){let ordinary=0;for(const [name,dcol,pcol,scol] of [['門市自取',3,9,17],['宅配到府',3,12,24]]){const sh=ss.getSheetByName(name);if(!sh||sh.getLastRow()<2)continue;sh.getRange(2,1,sh.getLastRow()-1,scol).getValues().forEach(v=>{const d=v[dcol-1] instanceof Date?Utilities.formatDate(v[dcol-1],MANAGER_TZ,'yyyy-MM-dd'):String(v[dcol-1]);if(d===date&&!['已取消','取消','作廢','已退款'].includes(String(v[scol-1])))ordinary+=Number(v[pcol-1])||0;});}return ordinary;}
function mgrRid_(){return Utilities.getUuid().replace(/-/g,'').slice(0,20);}
function mgrNewOrder2027(o){
 if(!o||!['自取','宅配'].includes(o.method))throw Error('請選擇自取或宅配');
 const ss=SpreadsheetApp.openById(MANAGER_TEST_SHEET_ID),date=mgrDate_(o.date),name=mgrText_(o.name,60,'姓名'),phone=mgrPhone_(o.phone),note=mgrText_(o.notes,250,'備註'),paid=o.payment==='已收款';
 if(!name||!['待付款','已收款'].includes(o.payment))throw Error('請填寫姓名及付款方式');
 if(paid&&!['現金','轉帳'].includes(o.payMethod))throw Error('請選擇現金或轉帳');
 const six=mgrInt_(o.six||0,'6入'),fifteen=mgrInt_(o.fifteen||0,'15入');
 let pieces,values,sh,phoneCol,deadlineCol;
 const lock=LockService.getScriptLock();if(!lock.tryLock(30000))throw Error('系統忙碌，請稍後重試');
 try{
 const created=new Date(),deadline=new Date(created.getTime()+72*3600000),id=mgrShortOrderId_(ss);
 const payment=paid?'已付款':'未付款',verified=paid?created:'',status=paid?(o.method==='自取'?'已確認':'待出貨'):'待確認';
 const tag='[APP:'+mgrText_(o.requestId,100,'請求識別碼')+']';
 if(!/^[a-zA-Z0-9_-]{12,100}$/.test(o.requestId||''))throw Error('缺少有效的送單識別碼');
 for(const [sheetName,n] of [['門市自取',18],['宅配到府',25]]){const x=ss.getSheetByName(sheetName);if(x&&x.getLastRow()>1){const v=x.getRange(2,1,x.getLastRow()-1,n).getDisplayValues();const hit=v.find(r=>String(r[n-1]).includes(tag));if(hit)return {ok:true,duplicate:true,orderId:hit[0]};}}
 if(o.method==='自取'){
 pieces=six*6+fifteen*15;if(!pieces||(six&&!fifteen&&six<8)||(fifteen&&!six&&fifteen<3)||(six&&fifteen&&pieces<42))throw Error('自取未達最低訂購量');
 const slot=mgrText_(o.slot,15,'取貨時段'),weekend=[0,6].includes(new Date(date+'T00:00:00Z').getUTCDay()),allowed=date==='2027-09-15'?['09:00–10:00']:weekend?['09:00–12:00','13:00–15:00','15:00–17:00','17:00–18:00']:['13:00–15:00','15:00–17:00','17:00–18:00'];
 if(!allowed.includes(slot))throw Error('取貨時段不符合日期');
 const goods=six*300+fifteen*750;sh=ss.getSheetByName('門市自取');phoneCol=6;deadlineCol=13;
 values=[id,created,date,slot,name,phone,six,fifteen,pieces,goods,'一般售價',goods,deadline,payment,paid?'88888':'',verified,status,note+' [收款:'+ (paid?o.payMethod:'待付款')+'] '+tag];
 }else{
 if([0,6].includes(new Date(date+'T00:00:00Z').getUTCDay()))throw Error('週末不提供宅配出貨');
 const invoiceType=String(o.invoiceType||'一般發票').trim(),taxId=String(o.taxId||'').trim();
 if(!['一般發票','公司發票'].includes(invoiceType))throw Error('發票類型不正確');
 if(invoiceType==='公司發票'&&!/^\d{8}$/.test(taxId))throw Error('公司發票請填寫8碼統一編號');
 const invoiceNote=invoiceType==='公司發票'?'[發票:公司／統編:'+taxId+']':'[發票:一般]';
 const address=mgrText_(o.address,200,'地址');if(address.length<8||/^(澎湖|金門|連江|馬祖|小琉球|綠島|蘭嶼)/.test(address))throw Error('請填寫台灣本島完整地址');
 const four=mgrInt_(o.four||0,'4盒組'),ten=mgrInt_(o.ten||0,'10盒組'),group=mgrInt_(o.group||0,'15盒組'),bags=mgrInt_(o.bags||0,'提袋');
 if(fifteen>0&&fifteen<4)throw Error('15入宅配至少4盒');if(bags>0&&!group)throw Error('提袋僅限6入15盒組');
 const sixBoxes=four*4+ten*10+group*15;pieces=sixBoxes*6+fifteen*15;if(!pieces)throw Error('請選擇商品');
 const split=Array.from({length:group},(_,i)=>Math.floor(bags/group)+(i<bags%group?1:0));
 const goods=sixBoxes*300+fifteen*750,bagFee=bags*20,cartons15=Math.ceil(fifteen/10),shipping=four*150+ten*170+split.reduce((a,n)=>a+(n>=5?210:170),0)+cartons15*210;
 const remote=/^(宜蘭縣|宜蘭市|花蓮縣|臺東縣|台東縣)/.test(address.replace(/\s/g,''))?(four+ten+group+cartons15)*10:0;
 sh=ss.getSheetByName('宅配到府');phoneCol=5;deadlineCol=18;
 values=[id,created,date,name,phone,address,four,ten,group,fifteen,bags,pieces,goods,bagFee,shipping,remote,goods+bagFee+shipping+remote,deadline,payment,'',verified,'待回傳','',status,note+' '+invoiceNote+' [收款:'+(paid?o.payMethod:'待付款')+'] '+tag];
 }
 if(!sh)throw Error('找不到訂單工作表');if(pieces>4800||mgrUsage_(ss,date)+pieces>4800)throw Error('一般訂單已超過每日 4800 顆上限');
 const row=sh.getLastRow()+1;sh.getRange(row,phoneCol).setNumberFormat('@');sh.getRange(row,1,1,values.length).setValues([values]);sh.getRange(row,2).setNumberFormat('yyyy/mm/dd hh:mm');sh.getRange(row,deadlineCol).setNumberFormat('yyyy/mm/dd hh:mm');
 return {ok:true,orderId:id,amount:o.method==='自取'?values[11]:values[16],pieces,paid};
 }finally{lock.releaseLock();}
}
function mgrNewVip2027(o){
 const date=mgrDate_(o.date),shift=mgrText_(o.shift,2,'時段'),method=mgrText_(o.method,2,'方式'),name=mgrText_(o.name,80,'對象'),six=mgrInt_(o.six||0,'6入'),fifteen=mgrInt_(o.fifteen||0,'15入'),contact=mgrText_(o.contact,200,'聯絡／地址'),notes=mgrText_(o.notes,300,'備註');
 if(!['上午','下午'].includes(shift)||!['自取','宅配'].includes(method)||!name||six*6+fifteen*15<=0)throw Error('請完整填寫 VIP 公關單');
 const lock=LockService.getScriptLock();if(!lock.tryLock(30000))throw Error('系統忙碌');
 try{const ss=SpreadsheetApp.openById(MANAGER_TEST_SHEET_ID),sh=mgrVipSheet_(ss),rid=mgrText_(o.requestId,100,'識別碼');if(!/^[a-zA-Z0-9_-]{12,100}$/.test(rid))throw Error('缺少送單識別碼');
 if(sh.getLastRow()>1){const rows=sh.getRange(2,1,sh.getLastRow()-1,12).getDisplayValues();const found=rows.find(r=>r[11].includes('[APP:'+rid+']'));if(found)return {ok:true,duplicate:true,orderId:found[0]};}
 const id='VIP27-'+Utilities.formatDate(new Date(),MANAGER_TZ,'yyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
 sh.appendRow([id,new Date(),date,shift,method,name,six,fifteen,six*6+fifteen*15,contact,'待製作',notes+' [APP:'+rid+']']);return {ok:true,orderId:id,pieces:six*6+fifteen*15};
 }finally{lock.releaseLock();}
}
function mgrCancelVip2027(id){
 const key=String(id||'').trim();if(!/^VIP27-[A-Z0-9-]+$/.test(key))throw Error('VIP 單號無效');
 const lock=LockService.getScriptLock();if(!lock.tryLock(30000))throw Error('系統忙碌');
 try{const sh=mgrVipSheet_(SpreadsheetApp.openById(MANAGER_TEST_SHEET_ID));if(sh.getLastRow()<2)throw Error('找不到 VIP 單');const rows=sh.getRange(2,1,sh.getLastRow()-1,11).getDisplayValues();const matches=rows.map((r,i)=>r[0]===key?i+2:0).filter(Boolean);if(matches.length!==1)throw Error('找不到唯一 VIP 單');const row=matches[0],status=sh.getRange(row,11).getDisplayValue();if(status==='已取消')return {ok:true,already:true};if(status==='已完成')throw Error('已完成公關單請人工處理');sh.getRange(row,11).setValue('已取消');return {ok:true};}finally{lock.releaseLock();}
}

/** 私人管理 APP：人工確認已實際收到款項，支援自取與宅配。 */
function mgrConfirmPayment2027(kind,orderId){
 if(!['pickup','delivery'].includes(kind))throw Error('訂單類型不正確');
 const id=String(orderId||'').trim();
 if(!/^(?:SJ-\d{5}|SJ27-[A-Z0-9-]{8,30})$/.test(id))throw Error('訂單編號格式不正確');
 const lock=LockService.getScriptLock();
 if(!lock.tryLock(15000))throw Error('系統忙碌，請稍後再試');
 try{
  const sh=SpreadsheetApp.openById(MANAGER_TEST_SHEET_ID).getSheetByName(kind==='pickup'?'門市自取':'宅配到府');
  if(!sh||sh.getLastRow()<2)throw Error('找不到訂單');
  const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().map(r=>String(r[0]).trim());
  const matches=ids.map((v,i)=>v===id?i+2:0).filter(Boolean);
  if(matches.length!==1)throw Error(matches.length?'訂單編號重複，請人工處理':'找不到訂單');
  const row=matches[0],payCol=kind==='pickup'?14:19,verifiedCol=kind==='pickup'?16:21,statusCol=kind==='pickup'?17:24;
  const status=String(sh.getRange(row,statusCol).getDisplayValue()).trim();
  if(['已取消','取消','作廢','已退款','已取貨','已出貨'].includes(status))throw Error('此訂單狀態不允許核帳：'+status);
  const payment=String(sh.getRange(row,payCol).getDisplayValue()).trim();
  if(payment==='已付款')return {ok:true,already:true,id};
  if(!['未付款','待核帳','逾期未付'].includes(payment))throw Error('目前付款狀態不允許核帳：'+payment);
  sh.getRange(row,payCol).setValue('已付款');
  sh.getRange(row,verifiedCol).setValue(new Date());
  if(status==='待確認')sh.getRange(row,statusCol).setValue('已確認');
  SpreadsheetApp.flush();
  return {ok:true,already:false,id};
 }finally{lock.releaseLock();}
}
