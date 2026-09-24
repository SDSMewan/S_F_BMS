let currentUser=null,currentModule=null,currentEditId=null,moduleRecords={};

async function api(functionName,...args){
  if(!CONFIG.WEB_APP_URL||CONFIG.WEB_APP_URL.includes("PASTE_YOUR")) throw new Error("index.html එකේ CONFIG.WEB_APP_URL එකට Apps Script /exec URL එක දාන්න.");
  const response=await fetch(CONFIG.WEB_APP_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify({functionName,args})});
  const text=await response.text();
  let result;
  try{result=JSON.parse(text)}catch(e){throw new Error("Apps Script response එක JSON නොවේ: "+text.slice(0,300))}
  if(result.ok===false||result.success===false) throw new Error(result.message||"Request failed");
  return result;
}

const CONFIG_MODULES={
 sales:{title:"Sales",fields:[
  ["saleDate","Sale Date","date",1],
  ["customerId","Customer","relation-customer",1],
  ["orderId","Related Order","relation-order",0],

  ["itemName","Item Name","select",1,["SB - 001","SB - 002","SB - 003"]],
  
  ["quantity","Quantity (kg)","number",1],
  ["unitPrice","Unit Price","number",1],
  ["serviceChargeRate","Service Charge Rate (%)","number",0],
  ["commission","Commission","number",0],
  ["paymentStatus","Payment Status","select",0,["Pending","Partial","Paid"]],
  ["notes","Notes","textarea",0,[],"full"]
 ]},
costs: {
  title: "Costs",
  fields: [
    ["description", "Description", "select", 1, [
      "Purchase of Raw Materials",
      "Transportation Costs",
      "Water Bill",
      "Electricity Bill",
      "Other"]],

    ["category", "Category", "select", 1, [
      "Fertilizer [Albert Solution]",
      "Fertilizer [YaraMila Complex]",
      "Fertilizer [Yaramila Target]",
      "Fertilizer [Yara Tera Calcinit Calcium Nitrate]",
      "Fungicides [Captan]",
      "Insecticide [Regent]",
      "Insecticide [Regent]",
      "Insecticide [Abamectin]",
      "Other"
    ]],

    ["quantity", "Quantity", "number", 1],

    ["unitPrice", "Unit Price", "number", 1],

    ["costDate", "Cost Date", "date", 1],

    ["notes", "Notes", "textarea", 0, [], "full"]
  ]
},
 customers:{title:"Customers",fields:[
  ["customerName","Customer Name","text",1], ["phone","Phone","text",1], ["address","Address","text",0],
  ["email","Email","email",0], ["customerType","Customer Type","select",1,["Regular","Wholesale","Retail"]],
  ["notes","Notes","textarea",0,[],"full"]
 ]},
 orders:{title:"Orders",fields:[
  ["customerId","Customer","relation-customer",1],
  ["orderDetails","Order Details","textarea",1,[],"full"],
  ["orderDate","Order Date","date",1], ["deliveryDate","Delivery Date","date",0],
  ["status","Status","select",1,["Pending","Processing","Completed","Cancelled"]],
  ["totalAmount","Total Amount","number",0], ["saleId","Linked Sale","relation-sale",0],
  ["notes","Notes","textarea",0,[],"full"]
 ]},
 payments:{title:"Payments",fields:[
  ["customerId","Customer","relation-customer",1],
  ["saleId","Linked Sale","relation-sale",0],
  ["orderId","Linked Order","relation-order",0],
  ["amount","Amount","number",1],
  ["paymentMethod","Payment Method","select",1,["Cash","Bank","Card","Online"]],
  ["paymentDate","Payment Date","date",1],
  ["status","Status","select",1,["Paid","Pending","Partial"]],
  ["notes","Notes","textarea",0,[],"full"]
 ]},
 receipts:{title:"Pay Receipts",fields:[
  ["paymentId","Payment","relation-payment",1],
  ["receiptDate","Receipt Date","date",1],
  ["notes","Notes","textarea",0,[],"full"]
 ]},
 users:{title:"User Management",fields:[
  ["username","Username","text",1], ["password","Password","password",1], ["fullName","Full Name","text",1],
  ["role","Role","select",1,["Admin","Staff"]], ["status","Status","select",1,["Active","Inactive"]]
 ]}
};

document.addEventListener("DOMContentLoaded",()=>{
  document.getElementById("loginForm").addEventListener("submit",loginSubmit);
  document.getElementById("logoutButton").addEventListener("click",logout);
  document.getElementById("mobileMenu").addEventListener("click",()=>toggleMenu());
  document.getElementById("backdrop").addEventListener("click",()=>toggleMenu(false));
  document.getElementById("pwToggle").addEventListener("click",e=>{const i=document.getElementById("loginPassword"),s=i.type==="password";i.type=s?"text":"password";e.target.textContent=s?"Hide":"Show"});
  document.getElementById("modal").addEventListener("mousedown",e=>{if(e.target.id==="modal")closeModal()});
  document.addEventListener("keydown",e=>{if(e.key==="Escape"){closeModal();toggleMenu(false)}});
  document.querySelectorAll(".nav").forEach(x=>x.addEventListener("click",()=>openPage(x.dataset.page)));
  document.getElementById("recordForm").addEventListener("submit",saveForm);
  // Always show the login page first when the website is opened/refreshed.
  // Do not automatically restore the previous user session from localStorage.
  currentUser=null;
  localStorage.removeItem("seemanUser");
  document.getElementById("loginPage").classList.remove("hidden");
  document.getElementById("appPage").classList.add("hidden");
  document.getElementById("loginUsername").focus();
});

function toggleMenu(force){
  const s=document.getElementById("sidebar"),open=force===undefined?!s.classList.contains("open"):force;
  s.classList.toggle("open",open);document.getElementById("backdrop").classList.toggle("show",open);
}
function setBusy(btn,busy){btn.disabled=busy;btn.classList.toggle("loading",busy)}
async function loginSubmit(e){
  e.preventDefault();
  const msg=document.getElementById("loginMessage"),lb=document.getElementById("loginButton"); msg.textContent="";setBusy(lb,true);
  try{
    const r=await api("login",document.getElementById("loginUsername").value.trim(),document.getElementById("loginPassword").value);
    if(!r.success) throw new Error(r.message||"Invalid login");
    currentUser=r.user;localStorage.setItem("seemanUser",JSON.stringify(currentUser));showApp();
  }catch(err){msg.textContent=err.message}
  finally{setBusy(lb,false)}
}
function showApp(){
  document.getElementById("loginPage").classList.add("hidden");document.getElementById("appPage").classList.remove("hidden");
  document.getElementById("currentUserName").textContent=currentUser.fullName||currentUser.username;
  document.getElementById("currentUserRole").textContent=currentUser.role;
  document.querySelectorAll(".admin-only").forEach(x=>x.style.display=currentUser.role==="Admin"?"":"none");
  openPage("dashboard");
}
async function logout(){
  try{await api("logout",currentUser)}catch(e){}
  currentUser=null;localStorage.removeItem("seemanUser");location.reload();
}
function allowed(page){return currentUser&&(currentUser.role==="Admin"||!["users","income"].includes(page))}
async function openPage(page){
  if(!allowed(page)){toast("Access denied","error");return}
  document.querySelectorAll(".page").forEach(x=>x.classList.remove("active"));
  document.getElementById("page-"+page).classList.add("active");
  document.querySelectorAll(".nav").forEach(x=>x.classList.toggle("active",x.dataset.page===page));
  document.getElementById("pageTitle").textContent={dashboard:"Dashboard",sales:"Sales",costs:"Costs",customers:"Customers",orders:"Orders",payments:"Payments",receipts:"Pay Receipts",users:"User Management",income:"Monthly Income"}[page];
  toggleMenu(false);window.scrollTo({top:0});
  if(page==="dashboard") await loadDashboard();
  else if(page==="income") await loadIncome();
  else await loadModule(page);
}

async function loadDashboard(){
  try{
    const r=await api("getDashboard");
    document.getElementById("dSales").textContent=money(r.sales);
    document.getElementById("dCosts").textContent=money(r.costs);
    const dp=document.getElementById("dProfit");dp.textContent=money(r.profit);dp.className=Number(r.profit)<0?"neg":"pos";
    document.getElementById("dPayments").textContent=r.pendingPayments||0;
    document.getElementById("dOrders").textContent=r.activeOrders||0;
    renderSimple("recentSales",r.recentSales||[]);
    renderSimple("recentCosts",r.recentCosts||[]);
  }catch(e){toast(e.message,"error")}
}
const MONEY_KEY=/(amount|price|total|commission|profit|cost)/i;
function cell(k,v){
  if(v===""||v==null)return {h:'<span style="color:#9aa8a4">-</span>',c:""};
  if(/status$/i.test(k)){const s=String(v).toLowerCase();const t=/paid|complete|active/.test(s)?"ok":/pending|partial|process/.test(s)?"warn":/cancel|inactive/.test(s)?"bad":"info";return {h:`<span class="pill ${t}">${safe(v)}</span>`,c:""}}
  if(MONEY_KEY.test(k)&&!/rate/i.test(k)&&v!==""&&!isNaN(v))return {h:money(v),c:"num"};
  if(/date$/i.test(k)&&/^\d{4}-\d{2}-\d{2}/.test(String(v)))return {h:safe(String(v).slice(0,10)),c:""};
  return {h:safe(v),c:""};
}
function renderSimple(id,records){
  if(!records.length){document.getElementById(id).innerHTML='<div class="empty">No records found.</div>';return}
  const keys=Object.keys(records[0]).slice(0,7);
  let h='<div class="table-wrap"><table class="table"><thead><tr>'+keys.map(k=>`<th>${header(k)}</th>`).join("")+'</tr></thead><tbody>';
  records.forEach(r=>h+="<tr>"+keys.map(k=>{const c=cell(k,r[k]);return `<td class="${c.c}" data-label="${header(k)}">${c.h}</td>`}).join("")+"</tr>");
  h+="</tbody></table></div>";document.getElementById(id).innerHTML=h;
}

async function loadModule(page){
  const c=CONFIG_MODULES[page]; if(!c)return;
  const box=document.getElementById(page+"Module");
  box.innerHTML=`<div class="module"><div class="module-head"><h2>${c.title}<span id="${page}Count" class="count"></span></h2><div class="tools"><input id="${page}Search" class="search" type="search" aria-label="Search ${c.title}" placeholder="Search ${c.title.toLowerCase()}..." oninput="filterModule('${page}')"><button class="btn primary" onclick="openAdd('${page}')">+ New</button></div></div><div id="${page}Table"><div class="skeleton"></div><div class="skeleton"></div><div class="skeleton"></div></div></div>`;
  try{const r=await api("getRecords",page);moduleRecords[page]=r.records||[];renderTable(page,moduleRecords[page])}catch(e){box.querySelector("#"+page+"Table").innerHTML=`<div class="empty">${safe(e.message)}</div>`}
}
function filterModule(page){
  const q=document.getElementById(page+"Search").value.toLowerCase();
  renderTable(page,(moduleRecords[page]||[]).filter(r=>Object.values(r).join(" ").toLowerCase().includes(q)));
}
function tableKeys(page,record){
  const hidden={users:["passwordHash","salt"],sales:["customerId"],orders:["customerId"],payments:["customerId"],receipts:["customerId","paymentId"]}[page]||[];
  const extraHidden={sales:["orderId"],orders:["saleId"],payments:["saleId","orderId"],receipts:[]}[page]||[];
  const hide=new Set([...hidden,...extraHidden]);
  return Object.keys(record).filter(k=>!hide.has(k));
}
function renderTable(page,records){
  const box=document.getElementById(page+"Table");if(!records.length){const cn=document.getElementById(page+"Count");if(cn)cn.textContent="0 records";const q=(document.getElementById(page+"Search")||{}).value;box.innerHTML=q?'<div class="empty">No results match "'+safe(q)+'".</div>':`<div class="empty">Nothing here yet.<br><button class="btn primary" onclick="openAdd('${page}')">+ Add the first record</button></div>`;return}
  const keys=tableKeys(page,records[0]);
  const cn=document.getElementById(page+"Count");if(cn)cn.textContent=records.length+(records.length===1?" record":" records");
  let h='<div class="table-wrap"><table class="table"><thead><tr>'+keys.map(k=>`<th class="${MONEY_KEY.test(k)&&!/rate/i.test(k)?"num":""}">${header(k)}</th>`).join("")+'<th>Actions</th></tr></thead><tbody>';
  records.forEach(r=>{
    const id=r.id;
    h+="<tr>"+keys.map(k=>{const c=cell(k,r[k]);return `<td class="${c.c}" data-label="${header(k)}">${c.h}</td>`}).join("")+`<td><div class="actions"><button class="action edit" onclick="editRecord('${page}','${escJs(id)}')">Edit</button><button class="action pdf" onclick="pdfRecord('${page}','${escJs(id)}')">PDF</button>${currentUser.role==="Admin"?`<button class="action del" onclick="deleteRecord('${page}','${escJs(id)}')">Delete</button>`:""}</div></td></tr>`;
  });
  h+="</tbody></table></div>";box.innerHTML=h;
}
async function openAdd(page){
  currentModule=page;currentEditId=null;
  document.getElementById("modalTitle").textContent="Add "+CONFIG_MODULES[page].title;
  try{await ensureRelationData(page);await buildForm(page,{});document.getElementById("modal").classList.remove("hidden");const f0=document.querySelector("#modalFields input,#modalFields select");if(f0&&window.innerWidth>650)f0.focus();}
  catch(e){toast(e.message,"error")}
}
async function editRecord(page,id){
  const r=(moduleRecords[page]||[]).find(x=>String(x.id)===String(id));if(!r)return;
  currentModule=page;currentEditId=id;
  document.getElementById("modalTitle").textContent="Edit "+CONFIG_MODULES[page].title;
  try{await ensureRelationData(page);await buildForm(page,r);document.getElementById("modal").classList.remove("hidden");const f0=document.querySelector("#modalFields input,#modalFields select");if(f0&&window.innerWidth>650)f0.focus();}
  catch(e){toast(e.message,"error")}
}
async function ensureRelationData(page){
  const needed={sales:["customers","orders"],orders:["customers","sales"],payments:["customers","sales","orders"],receipts:["payments"]}[page]||[];
  for(const n of needed){const r=await api("getRecords",n);moduleRecords[n]=r.records||[];}
}
function relationOptions(type,current){
  if(type==="relation-customer"){
    return (moduleRecords.customers||[]).map(x=>({value:x.id,label:`${x.customerName}${x.phone?" - "+x.phone:""}`,customerId:x.id}));
  }
  if(type==="relation-order"){
    return (moduleRecords.orders||[]).map(x=>({value:x.id,label:`${x.id} - ${x.customerName||""} - ${x.orderDetails||""}`,customerId:x.customerId}));
  }
  if(type==="relation-sale"){
    return (moduleRecords.sales||[]).map(x=>({value:x.id,label:`${x.id} - ${x.customerName||""} - ${x.itemName||""} - ${money(x.totalAmount)}`,customerId:x.customerId}));
  }
  if(type==="relation-payment"){
    return (moduleRecords.payments||[]).map(x=>({value:x.id,label:`${x.id} - ${x.customerName||""} - ${money(x.amount)} - ${x.status||""}`,customerId:x.customerId}));
  }
  return [];
}
async function buildForm(page,record){
  const fields=CONFIG_MODULES[page].fields; record={...record};
  ["saleDate","costDate","orderDate","paymentDate","receiptDate"].forEach(k=>{
    if(!record[k] && fields.some(f=>f[0]===k)) record[k]=new Date().toISOString().slice(0,10);
  });
  if(page==="payments" && !record.status)record.status="Paid";
  if(page==="sales" && !record.paymentStatus)record.paymentStatus="Pending";
  if(page==="orders" && !record.status)record.status="Pending";
  document.getElementById("modalFields").innerHTML=fields.map(f=>{
    const [name,label,type,required,opts=[],full=""]=f;
    let v=record[name]??"";
    if(type==="date"&&v)v=String(v).slice(0,10);
    if(type.startsWith("relation-")){
      let options=relationOptions(type,v);
      // Keep the selected value visible even if it is no longer in the current cache.
      if(v&&!options.some(o=>String(o.value)===String(v)))options.unshift({value:v,label:String(v)});
      return `<div class="field ${full}"><label class="${required?"req":""}">${label}</label><select name="${name}" ${required?"required":""} data-relation="${type}"><option value="">Select</option>${options.map(o=>`<option value="${attr(o.value)}" ${String(o.value)===String(v)?"selected":""}>${safe(o.label)}</option>`).join("")}</select></div>`;
    }
    if(type==="select")return `<div class="field ${full}"><label class="${required?"req":""}">${label}</label><select name="${name}" ${required?"required":""}><option value="">Select</option>${opts.map(o=>`<option value="${attr(o)}" ${String(o)===String(v)?"selected":""}>${safe(o)}</option>`).join("")}</select></div>`;
    if(type==="textarea")return `<div class="field ${full}"><label class="${required?"req":""}">${label}</label><textarea name="${name}" rows="4" ${required?"required":""}>${safe(v)}</textarea></div>`;
    const extra=type==="number"?' step="0.01" min="0"':"";
    const placeholder=(type==="password"&&v)?"Leave blank to keep current password":"";
    const req=required&&!v?"required":"";
    return `<div class="field ${full}"><label class="${required?"req":""}">${label}</label><input ${type==="number"?'inputmode="decimal"':""} name="${name}" type="${type}" value="${type==="password"?"":attr(v)}" placeholder="${placeholder}" ${req}${extra}></div>`;
  }).join("");
  attachRelationFiltering(page);
}
function attachRelationFiltering(page){
  const form=document.getElementById("recordForm");
  const customer=form.querySelector('[name="customerId"]');
  const sale=form.querySelector('[name="saleId"]');
  const order=form.querySelector('[name="orderId"]');
  if(customer){customer.addEventListener("change",()=>filterRelatedSelects(customer.value));}
  if(sale){sale.addEventListener("change",()=>{
    const x=(moduleRecords.sales||[]).find(r=>String(r.id)===String(sale.value));
    if(x&&customer&&!customer.value)customer.value=x.customerId||"";
    filterRelatedSelects(customer?customer.value:"");
  });}
  if(order){order.addEventListener("change",()=>{
    const x=(moduleRecords.orders||[]).find(r=>String(r.id)===String(order.value));
    if(x&&customer&&!customer.value)customer.value=x.customerId||"";
    filterRelatedSelects(customer?customer.value:"");
  });}
}
function filterRelatedSelects(customerId){
  const form=document.getElementById("recordForm");
  ["saleId","orderId"].forEach(name=>{
    const el=form.querySelector(`[name="${name}"]`);if(!el)return;
    const type=name==="saleId"?"relation-sale":"relation-order";
    const selected=el.value;let options=relationOptions(type,selected).filter(x=>!customerId||String(x.customerId)===String(customerId));
    el.innerHTML='<option value="">Select</option>'+options.map(o=>`<option value="${attr(o.value)}" ${String(o.value)===String(selected)?"selected":""}>${safe(o.label)}</option>`).join("");
    if(selected&&!options.some(o=>String(o.value)===String(selected)))el.value="";
  });
}

async function saveForm(e){
  e.preventDefault();const fd=new FormData(e.target),data={};fd.forEach((v,k)=>data[k]=v);
  const sb=document.getElementById("saveButton");setBusy(sb,true);
  try{
    const r=currentEditId?await api("updateRecord",currentModule,currentEditId,data,currentUser):await api("saveRecord",currentModule,data,currentUser);
    if(!r.success)throw new Error(r.message||"Save failed");
    closeModal();toast("Record saved","success");invalidateRelations(currentModule);await loadModule(currentModule);await loadDashboard();
  }catch(err){toast(err.message,"error")}
  finally{setBusy(sb,false)}
}
function invalidateRelations(page){
  const map={customers:["sales","orders","payments","receipts"],sales:["orders","payments"],orders:["sales","payments"],payments:["receipts"]};
  (map[page]||[]).forEach(x=>{delete moduleRecords[x]});
}
async function deleteRecord(page,id){
  if(currentUser.role!=="Admin")return toast("Only Admin can delete","error");
  if(!confirm("Delete this record? This cannot be undone."))return;
  try{await api("deleteRecord",page,id,currentUser);toast("Deleted","success");await loadModule(page);await loadDashboard()}catch(e){toast(e.message,"error")}
}
async function pdfRecord(page,id){
  try{const r=await api("generateRecordPDF",page,id,currentUser);if(r.url)window.open(r.url,"_blank")}catch(e){toast(e.message,"error")}
}
function closeModal(){document.getElementById("modal").classList.add("hidden");document.getElementById("recordForm").reset();currentEditId=null}
async function loadIncome(){
  const box=document.getElementById("incomeModule");
  box.innerHTML=`<div class="module"><div class="module-head"><h2>Monthly Income</h2><div class="tools"><input type="month" id="incomeMonth" class="search"><button class="btn primary" onclick="calculateIncome()">Generate</button><button class="btn success" onclick="incomePDF()">PDF</button></div></div><div id="incomeResult"><div class="empty">Pick a month and press Generate to see income.</div></div></div>`;
  document.getElementById("incomeMonth").value=new Date().toISOString().slice(0,7);
}
async function calculateIncome(){
  const v=document.getElementById("incomeMonth").value;if(!v)return toast("Select month","error");
  const [y,m]=v.split("-");try{const r=await api("generateIncome",+y,+m,currentUser);document.getElementById("incomeResult").innerHTML=`<div class="report-box"><div class="card"><small>Sales</small><strong>${money(r.sales)}</strong></div><div class="card"><small>Costs</small><strong>${money(r.costs)}</strong></div><div class="card"><small>Net Income</small><strong class="${Number(r.netIncome)<0?"neg":"pos"}">${money(r.netIncome)}</strong></div></div>`}catch(e){toast(e.message,"error")}
}
async function incomePDF(){
  const v=document.getElementById("incomeMonth").value;if(!v)return;
  const [y,m]=v.split("-");try{const r=await api("generateIncomePDF",+y,+m,currentUser);if(r.url)window.open(r.url,"_blank")}catch(e){toast(e.message,"error")}
}
function money(v){return "Rs. "+(Number(v)||0).toLocaleString("en-LK",{minimumFractionDigits:2,maximumFractionDigits:2})}
function header(s){return String(s).replace(/([A-Z])/g," $1").replace(/^./,x=>x.toUpperCase())}
function safe(v){return String(v??"").replace(/[&<>"']/g,x=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[x]))}
function attr(v){return safe(v)}
function escJs(v){return String(v??"").replace(/\\/g,"\\\\").replace(/'/g,"\\'")}
let toastTimer;function toast(msg,type=""){const t=document.getElementById("toast");t.textContent=msg;t.className="toast show "+type;clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.className="toast",3200)}