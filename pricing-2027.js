/* 2027 Mid-Autumn pricing engine. Pure functions; no production credentials. */
(function(root,factory){
  const api=factory();
  if(typeof module==="object"&&module.exports)module.exports=api;
  root.StillJi2027=api;
})(typeof globalThis!=="undefined"?globalThis:this,function(){
  "use strict";
  const PRICES=Object.freeze({six:300,fifteen:750,bag:20});
  function integer(value,label){const n=Number(value);if(!Number.isSafeInteger(n)||n<0)throw Error(label+"必須是非負整數");return n;}
  function allocateBags(groups,bags){
    groups=integer(groups,"15盒組數");bags=integer(bags,"提袋數");
    if(!groups&&bags)throw Error("沒有6入15盒組，不能配置提袋");
    if(!groups)return [];
    // Even allocation minimizes groups with >=5 bags when possible.
    const base=Math.floor(bags/groups),remainder=bags%groups;
    return Array.from({length:groups},(_,i)=>base+(i<remainder?1:0));
  }
  // These are last selectable SHIPPING dates, not order placement deadlines.
  const LAST_SHIPPING_DATES=Object.freeze({four:"2027-08-31",fifteenGroup:"2027-09-10"});
  const DAILY_CAPACITY_PIECES=4800;
  function shippingDateStatus(input,date){
    if(!/^2027-\d{2}-\d{2}$/.test(date||""))return {valid:false,message:"請選擇 2027 年的出貨日期"};
    const d=new Date(date+"T00:00:00Z");
    if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==date)return {valid:false,message:"出貨日期無效"};
    if(d.getUTCDay()===0||d.getUTCDay()===6)return {valid:false,message:"宅配週六、週日不出貨，請選擇平日"};
    if(Number(input.four||0)>0&&date>LAST_SHIPPING_DATES.four)return {valid:false,message:"6 入 4 盒組最晚出貨日為 2027/08/31"};
    if(Number(input.fifteenGroup||0)>0&&date>LAST_SHIPPING_DATES.fifteenGroup)return {valid:false,message:"6 入 15 盒組最晚出貨日為 2027/09/10"};
    return {valid:true,message:""};
  }
  function piecesForOrder(input){
    const four=integer(input.four||0,"6入4盒組"),ten=integer(input.ten||0,"6入10盒組"),group=integer(input.fifteenGroup||0,"6入15盒組"),fifteen=integer(input.fifteenBoxes||0,"15入盒數");
    return (four*4+ten*10+group*15)*6+fifteen*15;
  }
  function remoteFromAddress(address){return /(?:宜蘭縣|花蓮縣|臺東縣|台東縣)/.test(String(address||"").replace(/\s/g,""));}
  function calculate(input){
    const four=integer(input.four||0,"6入4盒組");
    const ten=integer(input.ten||0,"6入10盒組");
    const fifteenGroup=integer(input.fifteenGroup||0,"6入15盒組");
    const fifteenBoxes=integer(input.fifteenBoxes||0,"15入盒數");
    const bags=integer(input.bags||0,"不織布提袋");
    const remote=remoteFromAddress(input.address);
    if(bags>0&&!fifteenGroup)throw Error("提袋分配規則僅確認6入15盒組，其他組合需先確認");
    const bagAllocation=allocateBags(fifteenGroup,bags);
    const shippingSix=150*four+170*ten+bagAllocation.reduce((sum,n)=>sum+(n<=4?170:210),0);
    const sixBoxes=4*four+10*ten+15*fifteenGroup;
    if(fifteenBoxes>0&&fifteenBoxes<4)throw Error("15入宅配至少需訂購4盒");
    const fifteenCartons=Math.ceil(fifteenBoxes/10);
    const shippingFifteen=fifteenCartons*210;
    const shipping=shippingSix+shippingFifteen;
    const remoteBoxes=four+ten+fifteenGroup; // 6入固定組合一組一箱
    const remoteSurcharge=remote?(remoteBoxes+fifteenCartons)*10:0;
    const goods=sixBoxes*PRICES.six+fifteenBoxes*PRICES.fifteen;
    const bagAmount=bags*PRICES.bag;
    return {sixBoxes,fifteenBoxes,goods,bagAmount,bagAllocation,shippingSix,shippingFifteen,fifteenCartons,remoteSurcharge,shipping,total:goods+bagAmount+shipping+remoteSurcharge,requiresReview:false};
  }
  return Object.freeze({PRICES,LAST_SHIPPING_DATES,DAILY_CAPACITY_PIECES,shippingDateStatus,piecesForOrder,remoteFromAddress,allocateBags,calculate});
});
