Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("run").onclick = run;

      let params = new URLSearchParams(window.location.search).get("id");
  
      if(params){
          addTextToDialogHeader("Second Dialog")
      }
      else{
            addTextToDialogHeader("First Dialog")
      }
    }
  });
  
  export async function run() {
    /**
     * Insert your Outlook code here
     */
    let params = new URLSearchParams(window.location.search).get("id");

    if(params){
        Office.context.ui.messageParent("This is my message from second dialog");
    }
    else{
        Office.context.ui.messageParent("This is my message from first dialog");
    }
    
  }

  function addTextToDialogHeader(message){
    document.getElementById("dialog-header").textContent = message;
  }