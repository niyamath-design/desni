function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || "";
  var callback = params.callback || "";
  try {
    if (action === "getOtpResult") return jsonResponse_(getOtpResult_(params.requestId), callback);
    if (action === "requestOtp" || action === "verifyOtp") {
      var otpResult = processOtpRequest_(action, params);
      if (params.requestId) CacheService.getScriptCache().put("otp-result:" + params.requestId, JSON.stringify({ready:true,result:otpResult}), 120);
      return jsonResponse_(otpResult, callback);
    }
    if (action === "createAuditPayment") return jsonResponse_(createAuditPayment_(params), callback);
    var sheet = getResponsesSheet_();
    if (action === "getResponses") {
      var responses = [], lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 14)).getValues();
        for (var i = data.length - 1; i >= 0; i--) {
          if (!(data[i][1] || data[i][2] || data[i][3] || data[i][4] || data[i][5])) continue;
          responses.push({
            id: "sheet-lead-" + (i + 1),
            timestamp: data[i][0] instanceof Date ? data[i][0].toISOString() : (data[i][0] || "").toString(),
            name: data[i][1] || "", email: data[i][2] || "", phone: data[i][3] || "",
            service: data[i][4] || "", message: data[i][5] || "", status: "new",
            category: data[i][6] || (data[i][7] === "Yes" ? "₹1 Paid Client" : "Website Inquiry"),
            paidStatus: data[i][7] || "No", paidAt: data[i][8] || "", emailStatus: data[i][9] || "",
            paymentLinkId: data[i][10] || "", paymentId: data[i][11] || "", paymentAmount: data[i][12] || 0,
            upiTransactionId: data[i][13] || ""
          });
        }
      }
      return jsonResponse_({success:true,data:responses}, callback);
    }
    var name=params.name||"", email=params.email||"", phone=params.phone||"", service=params.service||"", message=params.message||"", upiTransactionId=params.upiTransactionId||"";
    if (name || email || service) {
      var lock=LockService.getScriptLock(); lock.waitLock(10000);
      try { sheet.appendRow([new Date(),name,email,phone,service,message,"Website Inquiry","No","","","","",0,upiTransactionId]); }
      finally { lock.releaseLock(); }
    }
    return jsonResponse_({success:true}, callback);
  } catch (error) {
    return jsonResponse_({success:false,error:error.toString()}, callback);
  }
}

function getResponsesSheet_() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName("Responses");
  var headers=["Timestamp","Name","Email","Phone","Service","Message","Category","₹1 Paid","Paid At","Auto Email Status","Cashfree Payment Link ID","Cashfree Payment ID","Paid Amount (paise)","UPI Transaction ID"];
  if (!sheet) sheet=ss.insertSheet("Responses");
  if (sheet.getLastRow()===0) { sheet.appendRow(headers); sheet.getRange(1,1,1,headers.length).setFontWeight("bold"); }
  else {
    sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight("bold");
  }
  return sheet;
}

// Run installPaidClientTrigger once from the Apps Script editor to authorize email and install the edit trigger.
function installPaidClientTrigger() {
  ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === "onPaidStatusEdit"; }).forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === "syncCashfreePayments" || t.getHandlerFunction() === "syncRazorpayPayments"; }).forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("onPaidStatusEdit").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  ScriptApp.newTrigger("syncCashfreePayments").timeBased().everyMinutes(5).create();
  syncPaidClientsSheet_();
}

function onPaidStatusEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "Responses" || e.range.getColumn() !== 8 || e.range.getRow() < 2) return;
  if (String(e.value || "").trim().toLowerCase() !== "yes") return;
  var row = e.range.getRow(), values = sheet.getRange(row, 1, 1, 13).getValues()[0];
  if (!values[2] || String(values[9] || "").indexOf("Sent") === 0) return;
  markPaidClient_(sheet, row, values[11] || "", values[12] || 0, new Date());
}

function syncPaidClientsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), source = getResponsesSheet_();
  var target = ss.getSheetByName("Rs 1 Paid Clients") || ss.insertSheet("Rs 1 Paid Clients");
  var values = source.getLastRow() > 1 ? source.getRange(1, 1, source.getLastRow(), 13).getValues() : [];
  var out = [["Timestamp","Name","Email","Phone","Service","Category","₹1 Paid","Paid At","Auto Email Status","Cashfree Link ID"]];
  for (var i = 1; i < values.length; i++) if (String(values[i][7]).toLowerCase() === "yes") out.push([values[i][0],values[i][1],values[i][2],values[i][3],values[i][4],values[i][6],values[i][7],values[i][8],values[i][9],values[i][11]]);
  target.clearContents(); target.getRange(1, 1, out.length, out[0].length).setValues(out); target.getRange(1,1,1,out[0].length).setFontWeight("bold");
}

function createAuditPayment_(params) {
  var name=String(params.name||"").trim(), email=String(params.email||"").trim(), phone=String(params.phone||"").replace(/[\s()-]/g,"");
  var digits=phone.replace(/\D/g,""), localPhone=digits.length===10?digits:(digits.length===12&&digits.slice(0,2)==="91"?digits.slice(2):"");
  var service=String(params.service||"").trim(), message=String(params.message||"").trim();
  if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || !/^[6-9][0-9]{9}$/.test(localPhone) || !service) return {success:false,message:"Please enter a valid name, email, phone, and service."};
  var props=PropertiesService.getScriptProperties(), keyId=props.getProperty("CASHFREE_CLIENT_ID"), keySecret=props.getProperty("CASHFREE_CLIENT_SECRET");
  if (!keyId || !keySecret) return {success:false,message:"Cashfree checkout setup is incomplete. Add CASHFREE_CLIENT_ID and CASHFREE_CLIENT_SECRET in Apps Script Project Settings → Script Properties."};
  var apiVersion=props.getProperty("CASHFREE_API_VERSION")||"2026-01-01";
  var linkPayload={link_id:"desni_"+Utilities.getUuid().replace(/-/g,"").slice(0,24),link_amount:1,link_currency:"INR",link_purpose:"₹1 Brand Audit",link_partial_payments:false,link_auto_reminders:false,customer_details:{customer_name:name,customer_email:email,customer_phone:localPhone},link_notify:{send_email:false,send_sms:false,send_whatsapp:false},link_notes:{service:service,product:"₹1 Brand Audit"}};
  var response=UrlFetchApp.fetch("https://api.cashfree.com/pg/links",{method:"post",contentType:"application/json",payload:JSON.stringify(linkPayload),headers:{"x-api-version":apiVersion,"x-client-id":keyId,"x-client-secret":keySecret},muteHttpExceptions:true});
  var status=response.getResponseCode(), body={}; try { body=JSON.parse(response.getContentText()); } catch(ignored) {}
  if(status<200||status>=300||!body.link_id||!body.link_url) return {success:false,message:"Cashfree could not create the payment link. Check the API credentials and API version in Apps Script settings."};
  var sheet=getResponsesSheet_(), lock=LockService.getScriptLock(); lock.waitLock(10000);
  try { sheet.appendRow([new Date(),name,email,"+91"+localPhone,service,message,"₹1 Audit Lead","No","","",body.link_id,"",0,params.upiTransactionId||""]); }
  finally { lock.releaseLock(); }
  return {success:true,shortUrl:body.link_url,paymentLinkId:body.link_id,leadId:"sheet-lead-"+sheet.getLastRow()};
}

function syncCashfreePayments() {
  var props=PropertiesService.getScriptProperties(), keyId=props.getProperty("CASHFREE_CLIENT_ID"), keySecret=props.getProperty("CASHFREE_CLIENT_SECRET");
  if(!keyId||!keySecret) return;
  var sheet=getResponsesSheet_(), last=sheet.getLastRow(); if(last<2) return;
  var apiVersion=props.getProperty("CASHFREE_API_VERSION")||"2026-01-01", values=sheet.getRange(2,1,last-1,14).getValues(), changed=false;
  for(var i=0;i<values.length;i++) {
    var linkId=String(values[i][10]||""); if(!linkId||String(values[i][7]).toLowerCase()==="yes") continue;
    try {
      var response=UrlFetchApp.fetch("https://api.cashfree.com/pg/links/"+encodeURIComponent(linkId),{method:"get",headers:{"x-api-version":apiVersion,"x-client-id":keyId,"x-client-secret":keySecret},muteHttpExceptions:true});
      if(response.getResponseCode()!==200) continue;
      var link=JSON.parse(response.getContentText());
      if(link.link_status!=="PAID"||Number(link.link_amount_paid)<Number(link.link_amount)||link.link_currency!=="INR") continue;
      var row=i+2, paymentId=(link.payments&&link.payments.length&&link.payments[0].cf_payment_id)||link.cf_link_id||"";
      markPaidClient_(sheet,row,paymentId,Math.round(Number(link.link_amount_paid)*100),new Date()); changed=true;
    } catch(err) { console.error("Cashfree payment sync failed for row "+(i+2)+": "+err); }
  }
  if(changed) syncPaidClientsSheet_();
}

function markPaidClient_(sheet,row,paymentId,amountPaise,paidAt) {
  var values=sheet.getRange(row,1,1,13).getValues()[0];
  if(String(values[9]||"").indexOf("Sent")===0) return;
  sheet.getRange(row,7).setValue("₹1 Paid Client"); sheet.getRange(row,8).setValue("Yes"); sheet.getRange(row,9).setValue(paidAt||new Date());
  if(paymentId) sheet.getRange(row,12).setValue(paymentId);
  if(amountPaise) sheet.getRange(row,13).setValue(amountPaise);
  try {
    MailApp.sendEmail(String(values[2]),"Your ₹1 Brand Audit is confirmed",["Hi "+(values[1]||"there")+",","","Thanks for purchasing the ₹1 Brand Audit. Your payment has been confirmed. We’ll review your brand and email you with the next steps.","","— desni"].join(String.fromCharCode(10)));
    sheet.getRange(row,10).setValue("Sent");
  } catch(err) { sheet.getRange(row,10).setValue("Failed: "+err.toString().slice(0,150)); }
  syncPaidClientsSheet_();
}

function processOtpRequest_(action, params) {
  var requestId=params.requestId||"";
  if (!/^[a-zA-Z0-9_-]{12,80}$/.test(requestId)) return {success:false,message:"Invalid verification request. Please try again."};
  var phone=(params.phone||"").replace(/[\s()-]/g,"");
  if (!/^\+91[6-9][0-9]{9}$/.test(phone)) return {success:false,message:"Enter a valid Indian mobile number."};
  var props=PropertiesService.getScriptProperties();
  var sid=props.getProperty("TWILIO_ACCOUNT_SID"), token=props.getProperty("TWILIO_AUTH_TOKEN"), service=props.getProperty("TWILIO_VERIFY_SERVICE_SID");
  if (!sid || !token || !service) return {success:false,message:"SMS verification is not configured yet. Add the Twilio Verify settings in Apps Script Script Properties."};
  var cache=CacheService.getScriptCache(), phoneKey="otp-cooldown:"+phone.replace(/[^0-9]/g,"");
  var url="https://verify.twilio.com/v2/Services/"+encodeURIComponent(service), payload={To:phone};
  if (action==="requestOtp") {
    var rateLock=LockService.getScriptLock(); rateLock.waitLock(10000);
    try {
      if (cache.get(phoneKey)) return {success:false,message:"Please wait 60 seconds before requesting another code."};
      var phoneHash=Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,phone)).replace(/[^a-zA-Z0-9]/g,"");
      var dayKey="otp-sends:"+Utilities.formatDate(new Date(),"GMT","yyyyMMdd")+":"+phoneHash;
      var sentCount=parseInt(props.getProperty(dayKey)||"0",10);
      if (sentCount>=5) return {success:false,message:"Daily verification limit reached for this number. Try again tomorrow."};
      props.setProperty(dayKey,String(sentCount+1));
      cache.put(phoneKey,"1",60);
    } finally { rateLock.releaseLock(); }
    url+="/Verifications"; payload.Channel="sms";
  } else {
    var code=(params.code||"").trim();
    if (!/^[0-9]{6}$/.test(code)) return {success:false,message:"Enter the 6-digit code from the SMS."};
    url+="/VerificationCheck"; payload.Code=code;
  }
  var response=UrlFetchApp.fetch(url,{method:"post",payload:payload,headers:{Authorization:"Basic "+Utilities.base64Encode(sid+":"+token)},muteHttpExceptions:true});
  var status=response.getResponseCode(), body={};
  try { body=JSON.parse(response.getContentText()); } catch (ignored) {}
  if (status<200 || status>=300) return {success:false,message:action==="requestOtp"?"We could not send the SMS. Check the Twilio Verify settings and try again.":"The code could not be checked. Please request a new code and try again."};
  if (action==="requestOtp" && body.status==="pending") return {success:true,message:"Verification code sent by SMS."};
  if (action==="verifyOtp" && body.status==="approved") return {success:true,message:"Phone number verified."};
  return {success:false,message:action==="verifyOtp"?"That code is incorrect or expired.":"The SMS provider did not accept the request."};
}

function getOtpResult_(requestId) {
  if (!/^[a-zA-Z0-9_-]{12,80}$/.test(requestId||"")) return {ready:true,result:{success:false,message:"Invalid verification request."}};
  var value=CacheService.getScriptCache().get("otp-result:"+requestId);
  return value ? JSON.parse(value) : {ready:false};
}

function jsonResponse_(data, callback) {
  var json=JSON.stringify(data);
  if (callback && /^desniJsonp_[a-zA-Z0-9_]+$/.test(callback)) return ContentService.createTextOutput(callback+"("+json+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
