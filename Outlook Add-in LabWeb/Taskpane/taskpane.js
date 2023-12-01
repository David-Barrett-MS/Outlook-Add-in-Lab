/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  //first dialog
  const urlString = await getAbsoluteURL(window.location.href, "./dialog.html");
  let url = new URL(urlString);
  url = url.toString();

  console.log("Opening first dialog")
  const dialogOptions = { width: 50, height: 30, displayInIframe: true };
  let firstArgs = await showDialog(url, dialogOptions,false);
  addMessageToTaskpane(firstArgs.message)


  //second dialog
  let url2 = new URL(urlString);
  const search_params = url2.searchParams;
  search_params.set("id", firstArgs.message);
  url2.search = search_params.toString();
  url2 = url2.toString();

  console.log("Opening second dialog")
  const dialogOptions2 = { width: 20, height: 50, displayInIframe: true };
  let secondArgs = await showDialog(url2, dialogOptions2,true);
  addMessageToTaskpane(secondArgs.message)
}

function addMessageToTaskpane(message){
  if(message)
    document.getElementById("message").textContent = message.toString();
}


export function showDialog(url, dialogOptions, secondDialog) {
  if (dialogOptions.callback) {
      dialogOptions.callback = undefined;
  }
  
  return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(url, dialogOptions, async(asyncResult) => {
          console.log(asyncResult)
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              console.log("Inside Failed Status of Dialog")

              //12007 : A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.
              if (secondDialog && asyncResult.error.code === 12007) {
                  try {
                      await sleep(1000);
                      const res = await showDialog(url, dialogOptions, secondDialog);
                      resolve(res);
                  } catch (e) {
                      reject(e);
                  }
                  // Recursive call
              } else {
                  asyncResult.value.close();
                  reject(asyncResult.error);
              }
          } else {
              asyncResult.value.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
                  // asyncResult.value.close();
                  console.log("Dialog event recieved . Is second dialog : " + secondDialog)
                  console.log(args)

                  //12006 : The dialog box was closed, usually because the user chooses the X button. Thrown within the dialog and triggers a DialogEventReceived event in the host page.
                  // if(args.error === 12006){
                  //   asyncResult.value.close();
                  // }
                  
                  resolve(args);
              });
              asyncResult.value.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {

                  console.log("Dialog message recieved . Is second dialog : " + secondDialog)
                  console.log(args)
                  asyncResult.value.close();
                  console.log("closed dialog")
                  resolve(args);
              });
          }
      });
  });
}


function getAbsoluteURL(base, relative) {
  const stack = base.split("/");
  const parts = relative.split("/");
  stack.pop();
  for (let i = 0; i < parts.length; i++) {
    if (parts[i] == ".") {
      continue;
    }
    if (parts[i] == "..") {
      stack.pop();
    } else {
      stack.push(parts[i]);
    }
  }
  return stack.join("/");
}

function sleep(time) {
  return new Promise((resolve) => setTimeout(resolve, time));
}