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
    const fifteenShipping=fifteenBoxes===0?0:fifteenBoxes<5?(fifteenBoxes<=2?170:210):0;
    const shippingFifteen=fifteenShipping;
    // 15入滿5盒免基本運費；非整箱餘數與多箱規則仍待確認。
    const shipping=shippingSix+shippingFifteen;
    const remoteBoxes=four+ten+fifteenGroup; // 6入固定組合一組一箱
    const remoteSurcharge=remote?remoteBoxes*10:0; // 15入偏遠費待確認，暫不計入
    const goods=sixBoxes*PRICES.six+fifteenBoxes*PRICES.fifteen;
    const bagAmount=bags*PRICES.bag;
    return {sixBoxes,fifteenBoxes,goods,bagAmount,bagAllocation,shippingSix,shippingFifteen,remoteSurcharge,shipping,total:goods+bagAmount+shipping+remoteSurcharge,requiresReview:(fifteenBoxes>=5&&fifteenBoxes%5!==0)||(fifteenBoxes>0&&remote)};
  }
  return Object.freeze({PRICES,allocateBags,calculate});
});
