const APP={
  NAME:"Seeman Foods Business Management System",
  TZ:"Asia/Colombo",
  SHEETS:{
    Users:["id","username","passwordHash","salt","fullName","role","status","createdAt"],
    Customers:["id","customerName","phone","address","email","customerType","notes","createdAt"],
    Sales:["id","customerId","customerName","orderId","itemName","quantity","unitPrice","subtotal","serviceChargeRate","serviceCharge","commission","totalAmount","paymentStatus","saleDate","notes","createdAt","createdBy"],
    Costs:["id","description","category","quantity","unitPrice","totalPrice","costDate","notes","createdAt","createdBy"],
    Orders:["id","customerId","customerName","orderDetails","orderDate","deliveryDate","status","totalAmount","saleId","notes","createdAt","createdBy"],
    Payments:["id","customerId","customerName","saleId","orderId","referenceId","amount","paymentMethod","paymentDate","status","notes","createdAt","createdBy"],
    Receipts:["id","paymentId","customerId","customerName","amount","receiptDate","notes","createdAt","createdBy"],
    Income:["id","year","month","sales","costs","netIncome","createdAt"]
  }};

function doGet(){
  try{
    setupSystem();
    return json({ok:true,status:"online",application:APP.NAME});
  }catch(err){
    return json({ok:false,success:false,message:err.message});
  }
}

function doPost(e){
  try{
    if(!e||!e.postData||!e.postData.contents)throw new Error("POST data missing");
    const q=JSON.parse(e.postData.contents),fn=q.functionName,a=q.args||[];
    let r;
    switch(fn){
      case "login":r=login(a[0],a[1]);break;
      case "logout":r={ok:true,success:true};break;
      case "getDashboard":r=getDashboard();break;
      case "getRecords":r=getRecords(a[0]);break;
      case "saveRecord":r=saveRecord(a[0],a[1],a[2]);break;
      case "updateRecord":r=updateRecord(a[0],a[1],a[2],a[3]);break;
      case "deleteRecord":r=deleteRecord(a[0],a[1],a[2]);break;
      case "generateRecordPDF":r=generateRecordPDF(a[0],a[1],a[2]);break;
      case "generateIncome":r=generateIncome(a[0],a[1],a[2]);break;
      case "generateIncomePDF":r=generateIncomePDF(a[0],a[1],a[2]);break;
      default:throw new Error("Unknown function: "+fn);
    }
    return json(r);
  }catch(err){return json({ok:false,success:false,message:err.message})}
}

function json(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON)}

function setupSystem(){
  const ss=SpreadsheetApp.getActive();
  Object.keys(APP.SHEETS).forEach(n=>migrateSheet(ss,n));
  createDefaultUsers();
  return "Setup complete";
}

function migrateSheet(ss,name){
  const wanted=APP.SHEETS[name];
  let sh=ss.getSheetByName(name)||ss.insertSheet(name);
  const lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(lastRow===0 || lastCol===0){
    sh.getRange(1,1,1,wanted.length).setValues([wanted]);
  }else{
    const oldHeaders=sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
    const oldRows=lastRow>1?sh.getRange(2,1,lastRow-1,lastCol).getValues():[];
    const oldIndex={}; oldHeaders.forEach((h,i)=>oldIndex[h]=i);
    const data=oldRows.filter(r=>r.some(v=>v!=="")).map(r=>wanted.map(h=>oldIndex[h]!==undefined?r[oldIndex[h]]:""));
    // Add relationship IDs while preserving existing customer names/references.
    data.forEach(row=>{
      const obj={}; wanted.forEach((h,i)=>obj[h]=row[i]);
      if(name==="Sales" && !obj.customerId && obj.customerName) obj.customerId=findCustomerIdByName(obj.customerName);
      if(name==="Orders" && !obj.customerId && obj.customerName) obj.customerId=findCustomerIdByName(obj.customerName);
      if(name==="Payments" && !obj.customerId && obj.customerName) obj.customerId=findCustomerIdByName(obj.customerName);
      if(name==="Payments" && !obj.saleId && obj.referenceId) obj.saleId=findExistingId("Sales",obj.referenceId);
      if(name==="Payments" && !obj.orderId && obj.referenceId) obj.orderId=findExistingId("Orders",obj.referenceId);
      if(name==="Receipts" && obj.paymentId){ const pay=findById("Payments",obj.paymentId); if(pay){ obj.customerId=pay.customerId||obj.customerId; obj.customerName=pay.customerName||obj.customerName; if(!obj.amount) obj.amount=pay.amount; } }
      wanted.forEach((h,i)=>row[i]=obj[h]??"");
    });
    sh.clearContents();
    sh.getRange(1,1,1,wanted.length).setValues([wanted]);
    if(data.length) sh.getRange(2,1,data.length,wanted.length).setValues(data);
  }
  sh.setFrozenRows(1); sh.getRange(1,1,1,wanted.length).setFontWeight("bold");
}

function findById(n,id){ if(!id)return null; return all(n).find(x=>String(x.id)===String(id))||null; }
function findExistingId(n,id){ return findById(n,id)?String(id):""; }
function findCustomerIdByName(name){
  if(!name)return "";
  const x=all("Customers").find(c=>String(c.customerName).trim().toLowerCase()===String(name).trim().toLowerCase());
  return x?String(x.id):"";
}

function createDefaultUsers(){
  const users=all("Users");if(users.length)return;
  const a=hashCreate("admin123"),s=hashCreate("staff123"),now=new Date();
  append("Users",{id:"USR-"+year()+"-000001",username:"admin",passwordHash:a.hash,salt:a.salt,fullName:"System Administrator",role:"Admin",status:"Active",createdAt:now});
  append("Users",{id:"USR-"+year()+"-000002",username:"staff",passwordHash:s.hash,salt:s.salt,fullName:"System Staff",role:"Staff",status:"Active",createdAt:now});
}

function login(u,p){
  const user=all("Users").find(x=>String(x.username).toLowerCase()===String(u||"").toLowerCase());
  if(!user||String(user.status).toLowerCase()!=="active")return{ok:true,success:false,message:"Invalid username or password."};
  if(hashPassword(p,user.salt)!==user.passwordHash)return{ok:true,success:false,message:"Invalid username or password."};
  return{ok:true,success:true,user:{id:user.id,username:user.username,fullName:user.fullName,role:user.role,status:user.status}};
}

function getRecords(n){
  n=valid(n);let records=all(n);
  if(n==="Sales") records=records.map(x=>{const o=x.orderId?findById("Orders",x.orderId):null;return Object.assign({},x,{orderLabel:x.orderId?(x.orderId+(o?" - "+o.orderDetails:"")):"-"})});
  if(n==="Orders") records=records.map(x=>{const s=x.saleId?findById("Sales",x.saleId):null;return Object.assign({},x,{saleLabel:x.saleId?(x.saleId+(s?" - "+s.itemName:"")):"-"})});
  if(n==="Payments") records=records.map(x=>{const s=x.saleId?findById("Sales",x.saleId):null,o=x.orderId?findById("Orders",x.orderId):null;return Object.assign({},x,{saleLabel:x.saleId?(x.saleId+(s?" - "+s.itemName:"")):"-",orderLabel:x.orderId?(x.orderId+(o?" - "+o.orderDetails:"")):"-"})});
  if(n==="Receipts") records=records.map(x=>{const p=x.paymentId?findById("Payments",x.paymentId):null;return Object.assign({},x,{paymentLabel:x.paymentId?(x.paymentId+(p?" - Rs. "+p.amount:"")):"-"})});
  return{ok:true,success:true,records:records};
}

function saveRecord(n,d,u){
  n=valid(n); if(!d)throw new Error("No data received");
  const row=Object.assign({},d,{id:newId(n),createdAt:new Date(),createdBy:u&&u.username||""});
  applyRelations(n,row);
  calculateFields(n,row);
  if(n==="Users"){
    if(!d.password && !d.passwordHash) throw new Error("Password is required for a new user.");
    if(d.password){const hp=hashCreate(d.password);row.passwordHash=hp.hash;row.salt=hp.salt;}
    row.status=d.status||"Active"; delete row.password;
  }
  append(n,row); SpreadsheetApp.flush();
  syncRelatedRecords(n,row);
  return{ok:true,success:true,id:row.id,message:"Record saved successfully."};
}

function applyRelations(n,row){
  if(n==="Sales"){
    if(!row.customerId) row.customerId=findCustomerIdByName(row.customerName);
    if(!row.customerId) throw new Error("Please select a valid customer for this sale.");
    const c=findById("Customers",row.customerId); if(!c)throw new Error("Customer not found.");
    row.customerName=c.customerName;
    if(row.orderId){const o=findById("Orders",row.orderId);if(!o)throw new Error("Order not found.");if(String(o.customerId)!==String(row.customerId))throw new Error("Sale customer and order customer must match.");}
  }
  if(n==="Orders"){
    if(!row.customerId) row.customerId=findCustomerIdByName(row.customerName);
    if(!row.customerId) throw new Error("Please select a valid customer for this order.");
    const c=findById("Customers",row.customerId); if(!c)throw new Error("Customer not found.");
    row.customerName=c.customerName;
    if(row.saleId){const sale=findById("Sales",row.saleId);if(!sale)throw new Error("Sale not found.");if(String(sale.customerId)!==String(row.customerId))throw new Error("Order customer and sale customer must match.");}
  }
  if(n==="Payments"){
    if(row.saleId){const sale=findById("Sales",row.saleId);if(!sale)throw new Error("Sale not found.");row.customerId=row.customerId||sale.customerId;row.customerName=sale.customerName;row.orderId=row.orderId||sale.orderId||"";}
    if(row.orderId){const o=findById("Orders",row.orderId);if(!o)throw new Error("Order not found.");row.customerId=row.customerId||o.customerId;row.customerName=o.customerName;row.saleId=row.saleId||o.saleId||"";}
    if(!row.customerId) row.customerId=findCustomerIdByName(row.customerName);
    if(!row.customerId) throw new Error("Please select a valid customer.");
    const c=findById("Customers",row.customerId);if(!c)throw new Error("Customer not found.");row.customerName=c.customerName;
    if(row.saleId){const sale=findById("Sales",row.saleId);if(String(sale.customerId)!==String(row.customerId))throw new Error("Payment customer and sale customer must match.");}
    if(row.orderId){const o=findById("Orders",row.orderId);if(String(o.customerId)!==String(row.customerId))throw new Error("Payment customer and order customer must match.");}
    if(!row.saleId && !row.orderId && !row.referenceId) throw new Error("Link this payment to a Sale or Order.");
    if(!row.referenceId) row.referenceId=row.saleId||row.orderId;
  }
  if(n==="Receipts"){
    const p=findById("Payments",row.paymentId);if(!p)throw new Error("Please select a valid payment.");
    row.customerId=p.customerId;row.customerName=p.customerName;
    if(!row.amount)row.amount=p.amount;
  }
}

function calculateFields(n,row){
  if(n==="Sales"){
    const q=+row.quantity||0,p=+row.unitPrice||0,r=+row.serviceChargeRate||0,c=+row.commission||0,sub=q*p,sc=sub*r/100;
    row.subtotal=sub;row.serviceCharge=sc;row.totalAmount=Math.max(0,sub+sc-c);
  }
  if(n==="Costs")row.totalPrice=(+row.quantity||0)*(+row.unitPrice||0);
}

function syncRelatedRecords(n,row){
  if(n==="Customers") syncCustomerName(row.id,row.customerName);
  if(n==="Sales" && row.orderId){patchById("Orders",row.orderId,{saleId:row.id,totalAmount:row.totalAmount,customerId:row.customerId,customerName:row.customerName});}
  if(n==="Orders" && row.saleId){patchById("Sales",row.saleId,{orderId:row.id,customerId:row.customerId,customerName:row.customerName});}
  if(n==="Payments" && row.saleId) updateSalePaymentStatus(row.saleId);
  if(n==="Payments" && row.orderId) updateOrderPaymentStatus(row.orderId);
  if(n==="Receipts" && row.paymentId){const p=findById("Payments",row.paymentId);if(p&&p.saleId)updateSalePaymentStatus(p.saleId);}
}
function syncCustomerName(customerId,name){
  ["Sales","Orders","Payments","Receipts"].forEach(n=>{
    const rows=all(n);rows.filter(x=>String(x.customerId)===String(customerId)).forEach(x=>patchById(n,x.id,{customerName:name}));
  });
}

function updateSalePaymentStatus(saleId){
  const sale=findById("Sales",saleId);if(!sale)return;
  const paid=all("Payments").filter(p=>String(p.saleId)===String(saleId)&&(String(p.status).toLowerCase()==="paid"||String(p.status).toLowerCase()==="partial")).reduce((a,p)=>a+(+p.amount||0),0);
  const total=+sale.totalAmount||0; const status=paid<=0?"Pending":paid+0.01>=total?"Paid":"Partial";
  patchById("Sales",saleId,{paymentStatus:status});
}
function updateOrderPaymentStatus(orderId){
  // Order status remains operational (Pending/Processing/Completed/Cancelled); payment state is derived from linked payments.
  return true;
}
function patchById(n,id,changes){
  const s=sheet(n),v=s.getDataRange().getValues(),h=v[0],ix=h.indexOf("id");
  for(let i=1;i<v.length;i++)if(String(v[i][ix])===String(id)){const row={};h.forEach((k,j)=>row[k]=v[i][j]);Object.keys(changes).forEach(k=>{if(h.indexOf(k)>=0)row[k]=changes[k]});s.getRange(i+1,1,1,h.length).setValues([h.map(k=>row[k]??"")]);return;}
}

function updateRecord(n,id,d,u){
  n=valid(n);const s=sheet(n),v=s.getDataRange().getValues(),h=v[0],ix=h.indexOf("id");
  let rn=-1;for(let i=1;i<v.length;i++)if(String(v[i][ix])===String(id)){rn=i+1;break}
  if(rn<0)throw new Error("Record not found");
  const old=all(n).find(x=>String(x.id)===String(id));const row=Object.assign({},old,d,{id:id,createdAt:old.createdAt||new Date(),createdBy:old.createdBy||u&&u.username||""});
  applyRelations(n,row);calculateFields(n,row);
  if(n==="Users"){
    if(d.password){const hp=hashCreate(d.password);row.passwordHash=hp.hash;row.salt=hp.salt;}
    delete row.password;
  }
  s.getRange(rn,1,1,h.length).setValues([h.map(k=>row[k]??"")]);SpreadsheetApp.flush();syncRelatedRecords(n,row);
  return{ok:true,success:true,message:"Record updated successfully."};
}

function deleteRecord(n,id,u){
  if(!u||u.role!=="Admin")throw new Error("Only Admin can delete records");
  n=valid(n);
  const linked=linkedRecords(n,id);
  if(linked.length)throw new Error("Cannot delete this record because it is linked to: "+linked.join(", ")+". Remove/update those links first.");
  const s=sheet(n),v=s.getDataRange().getValues(),ix=v[0].indexOf("id");
  for(let i=1;i<v.length;i++)if(String(v[i][ix])===String(id)){s.deleteRow(i+1);return{ok:true,success:true,message:"Record deleted"}}
  throw new Error("Record not found");
}
function linkedRecords(n,id){
  const out=[];
  if(n==="Customers"){
    if(all("Sales").some(x=>String(x.customerId)===String(id)))out.push("Sales");
    if(all("Orders").some(x=>String(x.customerId)===String(id)))out.push("Orders");
    if(all("Payments").some(x=>String(x.customerId)===String(id)))out.push("Payments");
    if(all("Receipts").some(x=>String(x.customerId)===String(id)))out.push("Pay Receipts");
  }
  if(n==="Sales"){
    if(all("Orders").some(x=>String(x.saleId)===String(id)))out.push("Orders");
    if(all("Payments").some(x=>String(x.saleId)===String(id)))out.push("Payments");
  }
  if(n==="Orders"){
    if(all("Sales").some(x=>String(x.orderId)===String(id)))out.push("Sales");
    if(all("Payments").some(x=>String(x.orderId)===String(id)))out.push("Payments");
  }
  if(n==="Payments" && all("Receipts").some(x=>String(x.paymentId)===String(id)))out.push("Pay Receipts");
  return out;
}

function getDashboard(){
  const now=new Date(),y=now.getFullYear(),m=now.getMonth();
  const sales=all("Sales"),costs=all("Costs"),payments=all("Payments"),orders=all("Orders"),customers=all("Customers"),receipts=all("Receipts");
  const ts=sales.reduce((a,x)=>a+(sameMonth(x.saleDate,y,m)?+x.totalAmount||0:0),0);
  const tc=costs.reduce((a,x)=>a+(sameMonth(x.costDate,y,m)?+x.totalPrice||0:0),0);
  const pending=payments.filter(x=>String(x.status).toLowerCase()==="pending").length;
  const active=orders.filter(x=>["pending","processing"].includes(String(x.status).toLowerCase())).length;
  return{ok:true,success:true,sales:ts,costs:tc,profit:ts-tc,pendingPayments:pending,activeOrders:active,customerCount:customers.length,receiptCount:receipts.length,recentSales:sales.slice(-5).reverse(),recentCosts:costs.slice(-5).reverse()};
}

function generateIncome(y,m,u){
  if(!u||u.role!=="Admin")throw new Error("Only Admin can view Monthly Income");
  y=+y;m=+m;
  const sales=all("Sales"),costs=all("Costs");
  const s=sales.reduce((a,x)=>a+(sameYearMonth(x.saleDate,y,m)?+x.totalAmount||0:0),0);
  const c=costs.reduce((a,x)=>a+(sameYearMonth(x.costDate,y,m)?+x.totalPrice||0:0),0);
  return{ok:true,success:true,year:y,month:m,sales:s,costs:c,netIncome:s-c};
}

function generateRecordPDF(n,id,u){
  n=valid(n);
  const r=all(n).find(x=>String(x.id)===String(id));
  if(!r)throw new Error("Record not found");

  const doc=DocumentApp.create("Seeman Foods - "+n+" - "+id);
  const body=doc.getBody();
  body.clear();
  body.setMarginTop(36);
  body.setMarginBottom(42);
  body.setMarginLeft(42);
  body.setMarginRight(42);

  // Header - designed to match the supplied White Simple Invoice sample.
  const logo=body.appendParagraph("SF");
  logo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  logo.setFontFamily("Arial").setFontSize(22).setBold(true);

  const company=body.appendParagraph("Seeman Foods");
  company.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  company.setFontFamily("Arial").setFontSize(20).setBold(true);

  const contact=body.appendParagraph(
    "195/B Alubogahalanda Pinnawala South, Waga.\n"+
    "seemanfoodsofficeal@gmail.com\n"+
    "+94 77 754 9425"
  );
  contact.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contact.setFontFamily("Arial").setFontSize(9).setForegroundColor("#666666");

  body.appendParagraph("");

  const heading=body.appendParagraph(n==="Sales"?"SALES INVOICE":title(n)+" DETAILS");
  heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  heading.setFontFamily("Arial").setFontSize(14).setBold(true);
  heading.setForegroundColor("#222222");

  const ref=body.appendParagraph("Reference: "+String(r.id||id));
  ref.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  ref.setFontFamily("Arial").setFontSize(9).setForegroundColor("#777777");
  body.appendParagraph("");

  if(n==="Sales"){
    const rows=[
      ["ID", valueForPdf(r.id)],
      ["Sale Date", valueForPdf(r.saleDate)],
      ["Customer Name", valueForPdf(r.customerName)],
      ["Item Name", valueForPdf(r.itemName)],
      ["Quantity", valueForPdf(r.quantity)],
      ["Unit Price", moneyForPdf(r.unitPrice)],
      ["Subtotal", moneyForPdf(r.subtotal)],
      ["Service Charge Rate", rateForPdf(r.serviceChargeRate)],
      ["Service Charge", moneyForPdf(r.serviceCharge)],
      ["Total Amount", moneyForPdf(r.totalAmount)],
      ["Payment Status", valueForPdf(r.paymentStatus)],
      ["Created By", valueForPdf(r.createdBy)]
    ];
    appendInvoiceTable(body,rows);

    body.appendParagraph("");
    const note=body.appendParagraph("Thank you for your business.");
    note.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    note.setFontFamily("Arial").setFontSize(9).setItalic(true).setForegroundColor("#777777");
  }else{
    const rows=[];
    Object.keys(r).forEach(k=>rows.push([title(k),valueForPdf(r[k])]));
    appendInvoiceTable(body,rows);
  }

  const footer=doc.addFooter();
  const fp=footer.appendParagraph(
    "If you have any question please contact : +94 77 754 9425 | seemanfoodsofficeal@gmail.com"
  );
  fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fp.setFontFamily("Arial").setFontSize(8).setForegroundColor("#777777");

  doc.saveAndClose();
  const file=DriveApp.createFile(DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF));
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return{ok:true,success:true,url:file.getUrl(),fileId:file.getId()};
}

function appendInvoiceTable(body,rows){
  const table=body.appendTable(rows);
  table.setBorderWidth(0);
  for(let i=0;i<table.getNumRows();i++){
    const row=table.getRow(i);
    const label=row.getCell(0);
    const value=row.getCell(1);

    label.setWidth(155);
    value.setWidth(330);
    label.setBackgroundColor("#F7F7F7");

    const lp=label.getChild(0).asParagraph();
    lp.setFontFamily("Arial").setFontSize(10).setBold(true).setForegroundColor("#222222");
    const vp=value.getChild(0).asParagraph();
    vp.setFontFamily("Arial").setFontSize(10).setForegroundColor("#333333");

    if(i===rows.length-1 && String(rows[i][0]).toLowerCase().indexOf("total")>=0){
      label.setBackgroundColor("#EDEDED");
      value.setBackgroundColor("#EDEDED");
      lp.setBold(true);vp.setBold(true).setFontSize(11);
    }
  }
}

function valueForPdf(v){
  if(v===null||v===undefined||v==="")return "-";
  if(Object.prototype.toString.call(v)==="[object Date]"){
    return Utilities.formatDate(v,APP.TZ,"dd/MM/yyyy HH:mm");
  }
  return String(v);
}

function moneyForPdf(v){
  if(v===null||v===undefined||v===""||isNaN(Number(v)))return "-";
  return "Rs. "+Number(v).toLocaleString("en-LK",{minimumFractionDigits:2,maximumFractionDigits:2});
}

function rateForPdf(v){
  if(v===null||v===undefined||v==="")return "-";
  const x=Number(v);
  if(isNaN(x))return String(v);
  return x+"%";
}

function generateIncomePDF(y,m,u){
  const r=generateIncome(y,m,u);
  const doc=DocumentApp.create("Seeman Foods Monthly Income");
  const body=doc.getBody();
  body.clear();
  body.setMarginTop(36);
  body.setMarginBottom(42);
  body.setMarginLeft(42);
  body.setMarginRight(42);

  const logo=body.appendParagraph("SF");
  logo.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  logo.setFontFamily("Arial").setFontSize(22).setBold(true);

  const company=body.appendParagraph("Seeman Foods");
  company.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  company.setFontFamily("Arial").setFontSize(20).setBold(true);

  const contact=body.appendParagraph(
    "195/B Alubogahalanda Pinnawala South, Waga.\n"+
    "seemanfoodsofficeal@gmail.com\n"+
    "+94 77 754 9425"
  );
  contact.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  contact.setFontFamily("Arial").setFontSize(9).setForegroundColor("#666666");

  body.appendParagraph("");
  const heading=body.appendParagraph("MONTHLY INCOME REPORT");
  heading.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  heading.setFontFamily("Arial").setFontSize(14).setBold(true);

  const period=body.appendParagraph("Period: "+y+" / "+String(m).padStart(2,"0"));
  period.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  period.setFontFamily("Arial").setFontSize(10).setForegroundColor("#777777");
  body.appendParagraph("");

  appendInvoiceTable(body,[
    ["Sales",moneyForPdf(r.sales)],
    ["Costs",moneyForPdf(r.costs)],
    ["Net Income",moneyForPdf(r.netIncome)]
  ]);

  const footer=doc.addFooter();
  const fp=footer.appendParagraph(
    "If you have any question please contact : +94 77 754 9425 | seemanfoodsofficeal@gmail.com"
  );
  fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fp.setFontFamily("Arial").setFontSize(8).setForegroundColor("#777777");

  doc.saveAndClose();
  const file=DriveApp.createFile(DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF));
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return{ok:true,success:true,url:file.getUrl(),fileId:file.getId()};
}

function moduleName(n){
  const map={
    sales:"Sales",costs:"Costs",customers:"Customers",orders:"Orders",
    payments:"Payments",receipts:"Receipts",users:"Users",income:"Income"
  };
  const key=String(n||"").trim();
  return APP.SHEETS[key]?key:(map[key.toLowerCase()]||key);
}
function sheet(n){
  n=moduleName(n);
  const ss=SpreadsheetApp.getActive();
  let s=ss.getSheetByName(n);
  if(!s){
    if(!APP.SHEETS[n]) throw new Error("Invalid module: "+n);
    s=ss.insertSheet(n);
    s.getRange(1,1,1,APP.SHEETS[n].length).setValues([APP.SHEETS[n]]);
    s.setFrozenRows(1);
    s.getRange(1,1,1,APP.SHEETS[n].length).setFontWeight("bold");
  }else if(s.getLastRow()===0){
    s.getRange(1,1,1,APP.SHEETS[n].length).setValues([APP.SHEETS[n]]);
    s.setFrozenRows(1);
  }
  return s;
}
function valid(n){n=moduleName(n);if(!APP.SHEETS[n])throw new Error("Invalid module: "+n);return n}
function all(n){
  const s=sheet(n),v=s.getDataRange().getValues();if(v.length<2)return[];
  const h=v[0];return v.slice(1).filter(r=>r.some(x=>x!=="")).map(r=>{const o={};h.forEach((k,i)=>o[k]=norm(r[i]));return o});
}
function append(n,o){const h=APP.SHEETS[n];sheet(n).appendRow(h.map(k=>o[k]??""))}
function newId(n){
  const p={Sales:"SAL",Costs:"CST",Customers:"CUS",Orders:"ORD",Payments:"PAY",Receipts:"RCP",Users:"USR",Income:"INC"}[n];if(!p)throw new Error("ID prefix missing");
  let max=0;all(n).forEach(x=>{const m=String(x.id||"").match(/-(\d+)$/);if(m)max=Math.max(max,+m[1])});
  return p+"-"+year()+"-"+String(max+1).padStart(6,"0");
}
function norm(v){return v instanceof Date?Utilities.formatDate(v,APP.TZ,"yyyy-MM-dd HH:mm:ss"):v}
function year(){return Utilities.formatDate(new Date(),APP.TZ,"yyyy")}
function sameMonth(v,y,m){if(!v)return false;const d=new Date(v);return !isNaN(d)&&d.getFullYear()===y&&d.getMonth()===m}
function sameYearMonth(v,y,m){if(!v)return false;const d=new Date(v);return !isNaN(d)&&d.getFullYear()===y&&d.getMonth()+1===m}
function title(s){return String(s).replace(/([A-Z])/g," $1").replace(/^./,x=>x.toUpperCase())}
function hashCreate(p){const salt=Utilities.getUuid();return{salt:salt,hash:hashPassword(p,salt)}}
function hashPassword(p,s){
  let x=String(p)+String(s);
  for(let i=0;i<1000;i++){
    const b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,x);
    x=b.map(v=>("0"+(v<0?v+256:v).toString(16)).slice(-2)).join("");
  }
  return x;
}
