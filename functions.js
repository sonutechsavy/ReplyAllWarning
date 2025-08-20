/* global Office */
const THRESHOLD = 3;

Office.onReady(() => {});

function getRecipientsCount(item) {
  return new Promise((resolve, reject) => {
    let toCount=0, ccCount=0;
    item.to.getAsync(toRes=>{
      if(toRes.status!==Office.AsyncResultStatus.Succeeded) return reject(toRes.error);
      toCount=(toRes.value||[]).length;
      item.cc.getAsync(ccRes=>{
        if(ccRes.status!==Office.AsyncResultStatus.Succeeded) return reject(ccRes.error);
        ccCount=(ccRes.value||[]).length;
        resolve(toCount+ccCount);
      });
    });
  });
}

function checkRecipientsOnSend(event){
  const item=Office.context.mailbox.item;
  getRecipientsCount(item).then(count=>{
    if(count>THRESHOLD){
      Office.context.ui.displayDialogAsync(
        "https://sonutechsavy.github.io/ReplyAllWarning/dialog.html?count="+count,
        {height:30,width:30,displayInIframe:true},
        result=>{
          if(result.status===Office.AsyncResultStatus.Succeeded){
            const dialog=result.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived,args=>{
              if(args.message==="send"){
                event.completed({allowEvent:true});
              }else{
                event.completed({allowEvent:false});
              }
              dialog.close();
            });
          }else{
            event.completed({allowEvent:true});
          }
        });
    }else{
      event.completed({allowEvent:true});
    }
  }).catch(err=>{
    console.error(err);
    event.completed({allowEvent:true});
  });
}

if(typeof window!=="undefined") window.checkRecipientsOnSend=checkRecipientsOnSend;
