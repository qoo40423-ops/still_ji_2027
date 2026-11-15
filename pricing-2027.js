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
  const DEADLINES=Object.freeze({four:"2027-08-31",fifteenGroup:"2027-09-10"});
  function taiwanDate(now){return new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Taipei",year:"numeric",month:"2-digit",day:"2-digit"}).format(now||new Date());}
  function checkDeadlines(input,now){const today=taiwanDate(now);const closed=[];for(const key of Object.keys(DEADLINES)){if(Number(input[key]||0)>0&&today>DEADLINES[key])closed.push(key);}return {today,closed};}
  function calculate(input){
    const four=integer(input.four||0,"6入4盒組");
    const ten=integer(input.ten||0,"6入10盒組");
    const fifteenGroup=integer(input.fifteenGroup||0,"6入15盒組");
    const fifteenBoxes=integer(input.fifteenBoxes||0,"15入盒數");
    const bags=integer(input.bags||0,"不織布提袋");
    const remote=Boolean(input.remote);
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
  return Object.freeze({PRICES,DEADLINES,taiwanDate,checkDeadlines,allocateBags,calculate});
});
