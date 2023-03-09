/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
/*
import $ from "jquery";
import jQuery from "jquery";

import tagit from "tag-it";*/
import axios from "axios";


var SERVICEURL = "https://whatsapp.scip.co/gmailparser/public/index.php/";
var ENCKEY = "#dk#";
var CONVERSATIONID;
var EMAILDATE;
var SUBJECT;
var FROMNAME;
var FROMEMAIL;
var SENDERNAME;
var SENDEREMAIL;
var TONAME;
var TOEMAIL;
var EMAILBODY;
var ATTACHMENTS_ARR = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;

    document.getElementById("saveButton").onclick = saveEmailData;

    setTimeout(function(){
      getEmailData();
      initTagit();
    }, 500);
    
  }
});

export async function run() {
  /** No need
   * Insert your Outlook code here
   */

  // Get a reference to the current message
 /*
  var item = Office.context.mailbox.item;
  console.log(item);
*/
}


export async function initTagit(){
  var sampleTags = [];

  $("#SCIP_industryTags").tagit({
    availableTags: sampleTags,
    allowSpaces: true,
    placeholderText: "Enter Industry tags"
  });

  $("#SCIP_technologyTags").tagit({
    availableTags: sampleTags,
    allowSpaces: true,
    placeholderText: "Enter Technology tags"
  });

  $("#SCIP_revenueModelTags").tagit({
    availableTags: sampleTags,
    allowSpaces: true,
    placeholderText: "Enter Revenue-Model tags."
  });

}


export async function getEmailData(){
  var item = Office.context.mailbox.item;
  // console.log("item:");
  // console.log(item);

  // console.log("item body:");
  // console.log(item.body);

  var coercionType = "text";
  item.body.getAsync(coercionType, function(response){
    //get body content
    // console.log("body response");
    // console.log(response);
    //status
    //value
    EMAILBODY = response.value;

  });


  
  CONVERSATIONID = item.conversationId;
  SUBJECT = item.subject;
  EMAILDATE = item.dateTimeCreated;
  FROMNAME = item.from.displayName;
  FROMEMAIL = item.from.emailAddress;
  SENDERNAME = item.sender.displayName;
  SENDEREMAIL = item.sender.emailAddress;
  TONAME = item.to[0].displayName;
  TOEMAIL = item.to[0].emailAddress;
  ATTACHMENTS_ARR = [];

  var attachments = item.attachments;
  // console.log("attachments");
  // console.log(attachments);

  if(attachments.length > 0){

    $.each(attachments, function(i,v){

      var tmpAttchType = v.attachmentType;
      var tmpAttchContType = v.contentType;
      var tmpAttchId = v.id;
      var tmpAttchIsInline = v.isInline;
      var tmpAttchName = v.name;
      var tmpAttchSize = v.size;

      //get attachment data
      // console.log("Office.AsyncContextOptions");
      // console.log(Office.AsyncContextOptions);
    
      var customValues = ['123dk'];  //user reference values
      item.getAttachmentContentAsync(tmpAttchId, customValues, function(attachResp){
          // console.log("attachResp");
          // console.log(attachResp);

          //status "succeeded"
          //value.content
          //value.format
          if(attachResp.status == "succeeded"){
            var tmpAttchObj = {
              "attachmentType":tmpAttchType,
              "contentType":tmpAttchContType,
              "id":tmpAttchId,
              "name":tmpAttchName,
              "size":tmpAttchSize,
              "content":attachResp.value.content,
              "format":attachResp.value.format
            };
          
            ATTACHMENTS_ARR.push(tmpAttchObj);

          }
          
      });
  
    });

    setTimeout(function(){
      // console.log("ATTACHMENTS_ARR");
      // console.log(ATTACHMENTS_ARR);
    },500);
    
  }


}


var TMP_elmId = "";
var TMP_bttnContent = "";

function saveEmailData(){

  /*
  CONVERSATIONID
  SUBJECT
  FROMNAME
  FROMEMAIL
  SENDERNAME
  SENDEREMAIL
  TONAME
  TOEMAIL
  //ATTACHMENTS_ARR
  */

  var SCIP_name = $("#SCIP_name").val();
  var SCIP_from = $("#SCIP_from").val();
  var SCIP_dealname = $("#SCIP_dealname").val();
  var SCIP_dealtype = $("#SCIP_dealtype").val();
  var SCIP_source = $("#SCIP_source").val();
  var SCIP_action = $("#SCIP_action").val();
  var SCIP_note = $("#SCIP_note").val();
  var SCIP_industryTags = $("#SCIP_industryTags").val();
  var SCIP_technologyTags = $("#SCIP_technologyTags").val();
  var SCIP_revenueModelTags = $("#SCIP_revenueModelTags").val();


  if(!isReal(CONVERSATIONID)){
    var msg = 'Please re-install the plugin';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_name)){
    var msg = 'Please enter the Name';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_from)){
    var msg = 'Please enter From Name';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_dealname)){
    var msg = 'Please enter the Deal Name';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_dealtype)){
    var msg = 'Please select the Deal Type';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_source)){
    var msg = 'Please select the Deal Source';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_action)){
    var msg = 'Please select the Action';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_note)){
    var msg = 'Please enter the Note';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_industryTags)){
    var msg = 'Please enter the Industry Tags';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_technologyTags)){
    var msg = 'Please enter the Technology Tags';
    var err = 1;
    showError(msg,err);
    return false;
  }else if(!isReal(SCIP_revenueModelTags)){
    var msg = 'Please enter the Revenue Model Tags';
    var err = 1;
    showError(msg,err);
    return false;
  }else{

    var SCIP_name = $("#SCIP_name").val();
    var SCIP_from = $("#SCIP_from").val();
    var SCIP_dealname = $("#SCIP_dealname").val();
    var SCIP_dealtype = $("#SCIP_dealtype").val();
    var SCIP_source = $("#SCIP_source").val();
    var SCIP_action = $("#SCIP_action").val();
    var SCIP_note = $("#SCIP_note").val();
    var SCIP_industryTags = $("#SCIP_industryTags").val();
    var SCIP_technologyTags = $("#SCIP_technologyTags").val();
    var SCIP_revenueModelTags = $("#SCIP_revenueModelTags").val();

      
    var dataJson = {
      "SCIP_name":SCIP_name,
      "SCIP_from":SCIP_from,
      "SCIP_dealname":SCIP_dealname,
      "SCIP_dealtype":SCIP_dealtype,
      "SCIP_source":SCIP_source,
      "SCIP_action":SCIP_action,
      "SCIP_note":SCIP_note,
      "SCIP_industryTags":SCIP_industryTags,
      "SCIP_technologyTags":SCIP_technologyTags,
      "SCIP_revenueModelTags":SCIP_revenueModelTags,
      "EMAIL_CONVERSATIONID":CONVERSATIONID,
      "EMAIL_SUBJECT":SUBJECT,
      "EMAIL_FROMNAME":FROMNAME,
      "EMAIL_FROMEMAIL":FROMEMAIL,
      "EMAIL_SENDERNAME":SENDERNAME,
      "EMAIL_SENDEREMAIL":SENDEREMAIL,
      "EMAIL_TONAME":TONAME,
      "EMAIL_TOEMAIL":TOEMAIL,
      "EMAIL_BODY":EMAILBODY,
      "EMAIL_DATE":EMAILDATE
    };
    
    //ATTACHMENTS_ARR


    //var path = "request.php";
    var path = "saveoutlookemail";
    var elmId = "saveButton";
    var bttnContent = $("#"+elmId).html();

    var TMP_elmId = elmId;
    var TMP_bttnContent = bttnContent;

    showLoader(elmId);

    callAjax(dataJson, path, function(ajaxResp){

      //console.log("ajaxResp:");
      //console.log(ajaxResp);
      //return false;

      if(ajaxResp.C == 100){
          // console.log("ATTACHMENTS_ARR:");
          // console.log(ATTACHMENTS_ARR);
          // debugger
          
          var messageId = ajaxResp.R.messageRecordId;
          var threadId = ajaxResp.R.threadId;
          var folderId = ajaxResp.R.folderId; 

          if(ATTACHMENTS_ARR.length > 0){
            
            /*
            var continueLoop = 0;
            for (let i = 0; i < ATTACHMENTS_ARR.length; i++) {
              var messageId = ajaxResp.R.messageRecordId;
              var threadId = ajaxResp.R.threadId;
              var folderId = ajaxResp.R.folderId; 
              var crrIdx = i;
              saveAttachments(threadId ,messageId,  folderId, crrIdx, function(attchRsp){
                // console.log("attchRsp");
                // console.log(attchRsp);
                if (attchRsp.C == 100) {
                  continueLoop = 1;
                  // console.log("continueLoop:");
                  // console.log(continueLoop);
                }else {
                  continueLoop = 0;
                }
              });
            }

            var msg = "Data saved successfully.";
            var err = 0;
            showError(msg, err);
            hideLoader(elmId, bttnContent);
            */

            //ATTACHMENTS_ARR
            var crrIdx = 0;
            saveAttachments(threadId ,messageId, folderId, crrIdx);


          }else{
            var msg = "Data saved successfully.";
            var err = 0;
            showError(msg, err);
            hideLoader(elmId, bttnContent);
          }

      }else{

        var msg = "Please try again";
        var err = 1;
        showError(msg, err);

        hideLoader(elmId, bttnContent);
      
      }

    });

  }

}



function saveAttachments(threadId ,messageId, folderId, crrIdx){

   //ATTACHMENTS_ARR
   //var crrIdx = 0;
  
  var totalAttchmnts = ATTACHMENTS_ARR.length;
  var lastAttchIdx = totalAttchmnts - 1;
  
  if(crrIdx <= lastAttchIdx){

      var path = "saveoutlookattachments";

      //var elmId = "saveButton";
      //var bttnContent = $("#"+elmId).html();
      //showLoader(elmId);
      var dataJson = {
        "threadId":threadId,
        "folder":folderId,
        "recordId":messageId,
        "attachmentType":ATTACHMENTS_ARR[crrIdx].attachmentType,
        "contentType":ATTACHMENTS_ARR[crrIdx].contentType,
        "id":ATTACHMENTS_ARR[crrIdx].id,
        "name":ATTACHMENTS_ARR[crrIdx].name,
        "size":ATTACHMENTS_ARR[crrIdx].size,
        "content":ATTACHMENTS_ARR[crrIdx].content,
        "format":ATTACHMENTS_ARR[crrIdx].format
      };
      
      var path = "saveoutlookattachments";
      callAjax(dataJson, path, function(ajaxResp){

        //console.log("ajaxResp");
        //console.log(ajaxResp);

        var newIdx = crrIdx + 1;
        
        //call itself for next attachment
        saveAttachments(threadId, messageId, folderId, newIdx);
      
      });

  }else{
    
    // hide loader
    var msg = "Data saved successfully.";
    var err = 0;
    showError(msg, err);
    hideLoader(TMP_elmId, TMP_bttnContent);
    
  }
}

/*
function saveAttachments(threadId ,messageId, folderId, crrIdx, callBack){

    var path = "saveoutlookattachments";

    //var elmId = "saveButton";
    //var bttnContent = $("#"+elmId).html();
    //showLoader(elmId);
    var dataJson = {
      "threadId":threadId,
      "folder":folderId,
      "recordId":messageId,
      "attachmentType":ATTACHMENTS_ARR[crrIdx].attachmentType,
      "contentType":ATTACHMENTS_ARR[crrIdx].contentType,
      "id":ATTACHMENTS_ARR[crrIdx].id,
      "name":ATTACHMENTS_ARR[crrIdx].name,
      "size":ATTACHMENTS_ARR[crrIdx].size,
      "content":ATTACHMENTS_ARR[crrIdx].content,
      "format":ATTACHMENTS_ARR[crrIdx].format
    };
    
    var path = "saveoutlookattachments";
    callAjax(dataJson, path, function(ajaxResp){

      //console.log("ajaxResp");
      //console.log(ajaxResp);
      return callBack(ajaxResp);
    
    });

}
*/


function isReal(arg){
  if(arg != "" && arg != null && arg != undefined){
    return true;
  }else{
    return false;
  }
}

function showError(msg, err){
  
  if(err == 1){
    //error  
    $("#alertMessage").addClass("errorClass");
    $("#alertMessage").removeClass("successClass");
  }else{
    //suucess
    $("#alertMessage").removeClass("errorClass");
    $("#alertMessage").addClass("successClass");
  }

  $("#alertMessage").html(msg);
  $("#alertMessage").fadeIn("slow");

  setTimeout(function(){
    $("#alertMessage").fadeOut("slow");
  },3000);
  
}

function showLoader(elmId){
  var loaderHtml = `<div class="spinner-border text-light" role="status"></div>`;
  $("#"+elmId).html(loaderHtml);
}

function hideLoader(elmId, content){
  $("#"+elmId).html(content);
}

function callAjax(dataJson, path, callback){

   $.ajax({
     //url:"https://cors-anywhere.herokuapp.com/"+SERVICEURL+path,
    //url:SERVICEURL+path,
     //url:"https://whatsapp.scip.co/gmailparser/request.php",
     url:"http://localhost:8080/"+SERVICEURL+path,
     data:dataJson,
     dataType:"json",
     type:"POST",
     method:"POST",
     crossDomain:true,
     cache:false,
     headers:{
       "accept":"application/json",
       "Access-Control-Allow-Origin":"*",
       "X-Requested-With":"XMLHttpRequest"
     },
    //headers:{"X-Requested-With":"XMLHttpRequest"},
     success:function(resp){
      //  console.log("resp");
      //  console.log(resp);
    
       return callback(resp); 
    
     },
     error:function(p1,p2,p3){
       console.log("p1");
       console.log(p1);
       console.log("p2");
       console.log(p2);
       console.log("p3");
       console.log(p3);

      return callback(p1+","+p2+","+p3); 
    }
  });

}