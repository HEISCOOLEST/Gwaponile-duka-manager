/* ============================================================
   FIREBASE SETUP
============================================================ */
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, signOut,
  createUserWithEmailAndPassword, sendPasswordResetEmail,
  EmailAuthProvider, reauthenticateWithCredential, updatePassword
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore, collection, doc, setDoc, addDoc, updateDoc, deleteDoc, runTransaction, writeBatch,
  onSnapshot, getDoc, serverTimestamp, query, where, orderBy, enableIndexedDbPersistence
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyD-DXZbSx_byeXUe8zPqidM0isZud0Sjww",
  authDomain: "duka-management-c2567.firebaseapp.com",
  projectId: "duka-management-c2567",
  storageBucket: "duka-management-c2567.firebasestorage.app",
  messagingSenderId: "941323216259",
  appId: "1:941323216259:web:ca3dca74a77209694e82b6"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
let offlineEnabled = false;
enableIndexedDbPersistence(db).then(()=>{ offlineEnabled=true; updateConnectionBadge(); }).catch(()=>{ updateConnectionBadge(); });

// Secondary app instance — used ONLY to create new employee accounts
// without signing the currently logged-in Admin out of their session.
const secondaryApp = initializeApp(firebaseConfig, "Secondary");
const secondaryAuth = getAuth(secondaryApp);

/* ============================================================
   CONSTANTS / HELPERS
============================================================ */
const SUPPORT = {
  phone: "+255671416562",
  whatsapp: "255671416562",
  email: "nehemiagwaponile7@gmail.com"
};

const ROLES = ["Administrator","Manager","Cashier","Storekeeper"];
const CAN_EXPENSES = r => r==="Administrator" || r==="Manager";
const CAN_PURCHASE = r => r==="Administrator" || r==="Manager" || r==="Storekeeper";
const CAN_BACKUP = r => r==="Administrator" || r==="Manager";
const EXPENSE_CATEGORIES = ["Rent","Electricity","Transport","Salaries","Internet","Supplies","Other"];
const uidSafe = () => crypto.randomUUID ? crypto.randomUUID() : uid();
const saleNumber = () => `SALE-${todayISO().replaceAll("-","")}-${String(Date.now()).slice(-6)}`;
function csvCell(v){ const x=String(v??""); return `"${x.replaceAll('"','""')}"`; }
function downloadText(filename, text, type="text/plain"){
  const blob=new Blob([text],{type});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=filename; a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}
function exportRows(filename, rows){
  if(!rows.length){ showToast("Hakuna data ya ku-export"); return; }
  const keys=[...new Set(rows.flatMap(r=>Object.keys(r)))];
  const csv=[keys.map(csvCell).join(","),...rows.map(r=>keys.map(k=>csvCell(typeof r[k]==="object"?JSON.stringify(r[k]):r[k])).join(","))].join("\n");
  downloadText(filename,csv,"text/csv;charset=utf-8");
}

/* ============================================================
   PROFESSIONAL EXCEL EXPORT (.xlsx)
   Badala ya CSV ghafi (maneno/JSON yanayojipanga hovyo),
   hii inatengeneza faili la Excel lenye column zilizopangwa,
   header ya rangi, tarehe zinazosomeka na currency format.
============================================================ */
const XLSX_HEADER_LABELS = {
  id:"ID", number:"Number", date:"Date", supplierId:"Supplier ID", supplierName:"Supplier",
  customerId:"Customer ID", customerName:"Customer", productId:"Product ID", name:"Name",
  purchaseId:"Purchase ID", createdBy:"Recorded By", createdAt:"Recorded On", updatedAt:"Updated On",
  paymentMethod:"Payment Method", paymentStatus:"Payment Status", dueDate:"Due Date",
  paidAmount:"Amount Paid", balance:"Balance", total:"Total", amount:"Amount",
  lineTotal:"Line Total", unitCost:"Unit Cost", costPrice:"Cost Price", sellingPrice:"Selling Price",
  totalSpent:"Total Spent", phone:"Phone Number", category:"Category", stock:"Stock",
  items:"Items", actorName:"Actor", actorUid:"Actor ID", referenceId:"Reference ID",
  status:"Status", qty:"Quantity", orders:"Orders", notes:"Notes"
};
function xlsxHeaderLabel(k){
  if(XLSX_HEADER_LABELS[k]) return XLSX_HEADER_LABELS[k];
  const spaced=String(k).replace(/([a-z0-9])([A-Z])/g,"$1 $2").replace(/_/g," ");
  return spaced.charAt(0).toUpperCase()+spaced.slice(1);
}
function xlsxIsMoneyKey(k){ return /amount|total|balance|price|cost|revenue|expense|spent|paid/i.test(k); }
function xlsxIsIdKey(k){ return /(^id$|Id$|uid$|Uid$)/.test(k); }
function xlsxReadableValue(k,v){
  if(v===null||v===undefined) return "";
  if(typeof v==="object"){
    if(typeof v.toDate==="function"){ try{ return v.toDate(); }catch(e){} }
    if(typeof v.seconds==="number"){ return new Date(v.seconds*1000); }
    if(Array.isArray(v)){
      return v.map(it=>(it&&typeof it==="object")?`${it.name||it.productName||"Item"} x${it.qty??""}`.trim():String(it)).join(", ");
    }
    return JSON.stringify(v);
  }
  if(typeof v==="string" && /^\d{4}-\d{2}-\d{2}$/.test(v)) return new Date(v+"T00:00:00");
  return v;
}
async function exportRowsXlsx(filename, rows, sheetTitle){
  if(!rows || !rows.length){ showToast("Hakuna data ya ku-export"); return; }
  if(typeof ExcelJS==="undefined"){ showToast("Excel engine haijapakia — angalia mtandao wako kisha jaribu tena."); return; }
  const keys=[...new Set(rows.flatMap(r=>Object.keys(r)))];
  const wb=new ExcelJS.Workbook();
  wb.creator="SALES MANAGEMENT SYSTEM"; wb.created=new Date();
  const ws=wb.addWorksheet((sheetTitle||"Export").slice(0,31),{views:[{state:"frozen",ySplit:1}]});
  ws.columns=keys.map(k=>({header:xlsxHeaderLabel(k),key:k,width:16}));
  rows.forEach(r=>{
    const rowObj={}; keys.forEach(k=>{ rowObj[k]=xlsxReadableValue(k,r[k]); });
    const row=ws.addRow(rowObj);
    keys.forEach((k,i)=>{
      const cell=row.getCell(i+1), val=rowObj[k];
      if(val instanceof Date){ cell.numberFmt = /Date$|^date$/.test(xlsxHeaderLabel(k))&&!(String(r[k]||"").length>10) ? "dd/mm/yyyy" : "dd/mm/yyyy hh:mm"; cell.alignment={horizontal:"left"}; }
      else if(typeof val==="number" && xlsxIsMoneyKey(k)){ cell.numberFmt='#,##0 "TZS";[Red]-#,##0 "TZS"'; cell.alignment={horizontal:"right"}; }
      else if(typeof val==="number"){ cell.alignment={horizontal:xlsxIsIdKey(k)?"left":"right"}; }
      else { cell.alignment={horizontal:"left"}; }
    });
  });
  const header=ws.getRow(1);
  header.eachCell(cell=>{
    cell.font={bold:true,color:{argb:"FFFFFFFF"},size:11};
    cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF0F172A"}};
    cell.alignment={vertical:"middle",horizontal:"center",wrapText:true};
    cell.border={bottom:{style:"medium",color:{argb:"FFFBBF24"}}};
  });
  header.height=24;
  ws.eachRow((row,rowNumber)=>{
    if(rowNumber===1) return;
    row.eachCell({includeEmpty:true},cell=>{
      cell.border={bottom:{style:"thin",color:{argb:"FFE7E5E4"}}};
      if(rowNumber%2===0) cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFAFAF9"}};
    });
    row.height=19;
  });
  if(keys.length>0) ws.autoFilter={from:{row:1,column:1},to:{row:1,column:keys.length}};
  ws.columns.forEach((col,idx)=>{
    const k=keys[idx]; let maxLen=xlsxHeaderLabel(k).length;
    rows.forEach(r=>{
      const v=xlsxReadableValue(k,r[k]);
      const len = v instanceof Date ? 16 : String(v??"").length;
      if(len>maxLen) maxLen=len;
    });
    col.width=Math.min(Math.max(maxLen+3,12),42);
  });
  try{
    const buf=await wb.xlsx.writeBuffer();
    const blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=filename; a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),1000);
  }catch(err){ showToast("Excel export imeshindikana: "+err.message); }
}

const PAYMENT_METHODS = ["Cash","M-Pesa","Tigo Pesa","Airtel Money","Bank"];
const CAN_SEE_FINANCE = r => r==="Administrator" || r==="Manager";
const CAN_DISCOUNT = r => r==="Administrator" || r==="Manager";
const CAN_MANAGE_USERS = r => r==="Administrator" || r==="Manager";
const CAN_EDIT_PRODUCTS = r => r==="Administrator" || r==="Manager" || r==="Storekeeper";
const CAN_SELL = r => r==="Administrator" || r==="Cashier";
const CAN_VIEW_SALES = r => r==="Administrator" || r==="Manager" || r==="Cashier";
const CAN_REFUND_SALE = r => r==="Administrator" || r==="Manager";
const CAN_MANAGE_SUPPLIERS = r => r==="Administrator" || r==="Manager" || r==="Storekeeper";
// Shift access: Admin/Manager/Cashier may READ shifts.
// Only Manager/Cashier use the My Shift workflow.
const CAN_READ_SHIFTS = r => r==="Administrator" || r==="Manager" || r==="Cashier";
const CAN_TRACK_SHIFT = r => r==="Manager" || r==="Cashier";
const CAN_SEE_SHIFTS_REPORT = r => r==="Administrator" || r==="Manager";
// Matches rules: customers/sales/shifts read = Admin/Manager/Cashier only (Inventory excluded)
const CAN_SEE_CUSTOMERS = r => r!=="Storekeeper";
// Matches rules: customers/suppliers delete = Admin/Manager only
const CAN_DELETE_CUSTOMER = r => r==="Administrator" || r==="Manager";
const CAN_DELETE_SUPPLIER = r => r==="Administrator" || r==="Manager";
// Matches rules: expenses delete = Administrator only
const CAN_DELETE_EXPENSE = r => r==="Administrator";
// Matches rules: Users/auditLog read = Admin OR Manager (management actions stay Admin-only)
const CAN_VIEW_USERS = r => r==="Administrator" || r==="Manager";

const fmt = n => "TZS " + Math.round(Number(n)||0).toLocaleString("en-US");
const uid = () => Math.random().toString(36).slice(2,10)+Date.now().toString(36);
const todayISO = () => new Date().toISOString().slice(0,10);
const monthKey = d => d.slice(0,7); // "YYYY-MM"
const escapeHtml = s => String(s??"").replace(/[&<>"']/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const initials = name => (name||"?").trim().split(/\s+/).slice(0,2).map(w=>w[0]?.toUpperCase()).join("");

/* ============================================================
   STATE
============================================================ */
let state = {
  authReady:false, user:null, profile:null, profileLoadError:"", // {name,email,role,status}
  tab:"dashboard",
  products:[], customers:[], customerPayments:[], sales:[], users:[],
  suppliers:[], shifts:[], expenses:[], purchases:[], stockMoves:[],
  cart:[], saleCustomerId:null, walkInName:"", walkInPhone:"", saveWalkInCustomer:false, productSearch:"", invSearch:"", supplierSearch:"",
  discountType:"none", discountValue:"", paymentMethod:"Cash", amountPaid:"", receiptSearch:"",
  saleOnCredit:false, saleDueDate:"", printerName: (localStorage.getItem("duka_printer_name")||""),
  auditLog:[],
  reportMonth: todayISO().slice(0,7),
  loginError:"",
  shiftStartCash:"", shiftEndCash:"",
  chatMessages:[], chatText:"", chatReplyTo:null,
  // Theme preferences: localTheme belongs to each employee; globalThemeOverride
  // is controlled by Administrator and temporarily overrides everyone.
  localTheme:"ocean",
  globalThemeOverride:null,
  themeControlLoaded:false,
};

let unsubs = [];
function clearListeners(){ unsubs.forEach(u=>u()); unsubs=[]; }
let notifiedLowStockIds = new Set(); // avoid repeat browser notifications in same session

/* ============================================================
   ICONS
============================================================ */
const ICONS = {
  chat:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5a8.4 8.4 0 0 1-9 8.5 9.7 9.7 0 0 1-4.2-1L3 20l1.5-4.2A8.3 8.3 0 0 1 3 11.5 8.4 8.4 0 0 1 12 3a8.4 8.4 0 0 1 9 8.5Z"/><path d="M8 11h.01M12 11h.01M16 11h.01"/></svg>',
  spark:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8L12 3z"/><path d="M19 16l.7 2.3L22 19l-2.3.7L19 22l-.7-2.3L16 19l2.3-.7L19 16z"/></svg>',

  dashboard:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9"/><rect x="14" y="3" width="7" height="5"/><rect x="14" y="12" width="7" height="9"/><rect x="3" y="16" width="7" height="5"/></svg>',
  sale:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>',
  inventory:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="M3.3 7 12 12l8.7-5"/><path d="M12 22V12"/></svg>',
  customers:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
  history:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v5h5"/><path d="M3.05 13A9 9 0 1 0 6 5.3L3 8"/><path d="M12 7v5l4 2"/></svg>',
  report:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg>',
  users:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
  support:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>',
  plus:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M5 12h14"/></svg>',
  minus:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/></svg>',
  trash:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>',
  edit:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/></svg>',
  close:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18M6 6l12 12"/></svg>',
  search:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></svg>',
  alert:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0Z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg>',
  trend:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>',
  profit:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>',
  receipt:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>',
  chevron:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round"><path d="m9 18 6-6-6-6"/></svg>',
  cart:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 2 1.58h9.78a2 2 0 0 0 2-1.58l1.65-7.42H5.12"/></svg>',
  phone:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>',
  mail:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16v16H4z"/><path d="m22 6-10 7L2 6"/></svg>',
  whatsapp:'<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2.05 22l5.25-1.38a9.9 9.9 0 0 0 4.74 1.21h.01c5.46 0 9.91-4.45 9.91-9.91 0-2.65-1.03-5.14-2.9-7.01A9.82 9.82 0 0 0 12.04 2zm0 18.13h-.01a8.2 8.2 0 0 1-4.18-1.14l-.3-.18-3.12.82.83-3.04-.2-.31a8.2 8.2 0 0 1-1.26-4.37c0-4.54 3.7-8.24 8.25-8.24 2.2 0 4.27.86 5.83 2.42a8.18 8.18 0 0 1 2.41 5.83c0 4.55-3.7 8.21-8.25 8.21zm4.52-6.16c-.25-.12-1.47-.72-1.7-.81-.23-.08-.39-.12-.56.13-.17.24-.64.81-.78.97-.14.17-.29.19-.53.06-.25-.12-1.05-.39-2-1.23-.74-.66-1.24-1.47-1.39-1.72-.14-.25-.02-.38.11-.51.11-.11.25-.29.37-.43.12-.14.16-.25.25-.41.08-.17.04-.31-.02-.43-.06-.12-.56-1.34-.76-1.84-.2-.48-.4-.42-.56-.42-.14-.01-.31-.01-.47-.01a.9.9 0 0 0-.65.31c-.23.24-.86.85-.86 2.06 0 1.22.89 2.4 1.01 2.56.12.17 1.75 2.67 4.23 3.74.59.26 1.05.41 1.41.52.59.19 1.13.16 1.55.1.47-.07 1.47-.6 1.68-1.18.21-.58.21-1.08.14-1.18-.06-.11-.23-.17-.48-.29z"/></svg>',
  logout:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>',
  sun:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.41-1.41M17.66 6.34l1.41-1.41"/></svg>',
  moon:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M21 12.8A8.5 8.5 0 1 1 11.2 3 6.7 6.7 0 0 0 21 12.8z"/></svg>',
  lock:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>',
  camera:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.5 4h-5L7 7H4a2 2 0 0 0-2 2v9a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-3l-2.5-3z"/><circle cx="12" cy="13" r="3.5"/></svg>',
  truck:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 3h15v13H1z"/><path d="M16 8h4l3 3v5h-7V8z"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>',
  clock:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
  print:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>',
};

/* ============================================================
   AUTH FLOW
============================================================ */
async function fetchUserProfileWithRetry(uid, attempts=3){
  let lastErr = null;
  for(let i=0;i<attempts;i++){
    try{
      return await getDoc(doc(db,"Users",uid));
    }catch(e){
      lastErr = e;
      // Errors that are typically transient (blocked/interrupted request, offline,
      // ad-blocker/extension interference, slow network) — worth retrying instead
      // of immediately kicking the user back to the login screen.
      const code = (e && e.code) || "";
      const transient = code.includes("unavailable") || code.includes("deadline-exceeded")
        || code.includes("network-request-failed") || code.includes("internal")
        || code.includes("cancelled") || !navigator.onLine;
      if(!transient || i === attempts-1) throw e;
      await new Promise(r=>setTimeout(r, 1200*(i+1)));
    }
  }
  throw lastErr;
}

onAuthStateChanged(auth, async (fbUser) => {
  clearListeners();
  if(!fbUser){
    state.user = null; state.profile = null;
    state.globalThemeOverride = null;
    state.themeControlLoaded = false;
    state.localTheme = getSavedTheme();
    refreshEffectiveTheme();
    state.authReady = true;
    render();
    return;
  }
  state.user = fbUser;
  try{
    const snap = await fetchUserProfileWithRetry(fbUser.uid);
    if(!snap.exists()){
      state.loginError = "Akaunti hii haijasajiliwa kwenye mfumo. Wasiliana na Admin.";
      await signOut(auth);
      state.authReady = true; render();
      return;
    }
    const profile = snap.data();
    // Normalize to tolerate stray whitespace / autocapitalized letters from mobile keyboards
    profile.status = (profile.status || "").toString().trim().toLowerCase();
    const roleRaw = (profile.role || "").toString().trim().toLowerCase();
    const matchedRole = ROLES.find(r => r.toLowerCase() === roleRaw);
    if(matchedRole) profile.role = matchedRole;

    if(profile.status !== "active"){
      state.loginError = "Akaunti yako imezimwa. Wasiliana na Admin.";
      await signOut(auth);
      state.authReady = true; render();
      return;
    }
    if(!matchedRole){
      state.loginError = "Role ya akaunti yako si sahihi. Wasiliana na Admin.";
      await signOut(auth);
      state.authReady = true; render();
      return;
    }
    state.profile = profile;
    state.profileLoadError = "";
    state.tab = "dashboard";
    state.localTheme = getSavedTheme();
    state.globalThemeOverride = null;
    state.themeControlLoaded = false;
    refreshEffectiveTheme();
    state.authReady = true;
    startListeners();
    requestNotifyPermission();
    render();
  }catch(e){
    console.error(e);
    // Do NOT force-logout on network/connection errors — that was kicking users
    // back to the login screen after brief network hiccups. Keep the Firebase
    // session alive and let them retry loading their profile instead.
    state.profileLoadError = "Imeshindwa kupakia akaunti yako: " + (e.code || e.message || e) + ". Angalia mtandao wako kisha jaribu tena.";
    state.authReady = true; render();
  }
});
async function retryProfileLoad(){
  if(!state.user) return;
  state.profileLoadError = "";
  render();
  try{
    const snap = await fetchUserProfileWithRetry(state.user.uid, 1);
    if(snap.exists()){
      const profile = snap.data();
      profile.status = (profile.status || "").toString().trim().toLowerCase();
      const roleRaw = (profile.role || "").toString().trim().toLowerCase();
      const matchedRole = ROLES.find(r => r.toLowerCase() === roleRaw);
      if(matchedRole) profile.role = matchedRole;
      if(profile.status === "active" && matchedRole){
        state.profile = profile;
        state.tab = "dashboard";
        state.localTheme = getSavedTheme();
        state.globalThemeOverride = null;
        state.themeControlLoaded = false;
        refreshEffectiveTheme();
        startListeners();
        requestNotifyPermission();
      }
    }
  }catch(e){
    state.profileLoadError = "Imeshindwa tena: " + (e.code || e.message || e) + ". Angalia mtandao wako.";
  }
  render();
}

function startListeners(){
  unsubs.push(onSnapshot(collection(db,"products"), snap=>{
    state.products = snap.docs.map(d=>({id:d.id, ...d.data()}));
    checkLowStockNotifications(state.products);
    render();
  }));
  // Rules: suppliers read = Admin/Manager/Inventory only. Cashier must not subscribe.
  if(CAN_MANAGE_SUPPLIERS(state.profile.role)){
    unsubs.push(onSnapshot(collection(db,"suppliers"), snap=>{
      state.suppliers = snap.docs.map(d=>({id:d.id, ...d.data()}));
      render();
    }));
  }
  // Rules: shifts read = Admin/Manager/Cashier only. Inventory must not subscribe.
  // Firestore rules only allow a Cashier to read shifts where cashierUid == their own uid.
  // Because that condition depends on resource.data, an unfiltered collection(db,"shifts")
  // query is rejected outright for Cashiers (Firestore denies queries that COULD return
  // documents the caller isn't allowed to see) — so a Cashier's shifts list never populates
  // and endShift() looks successful (the direct doc update is allowed) while the "My Shift"
  // screen keeps showing the shift as open, because state.shifts was never refreshed at all.
  // Fix: Cashiers subscribe to a query scoped with where("cashierUid","==", their uid),
  // which matches the rule and is allowed; Admin/Manager keep the full collection.
  if(CAN_READ_SHIFTS(state.profile.role)){
    const shiftsQuery = state.profile.role === "Cashier"
      ? query(collection(db,"shifts"), where("cashierUid","==", state.user.uid))
      : collection(db,"shifts");
    unsubs.push(onSnapshot(shiftsQuery, snap=>{
      state.shifts = snap.docs.map(d=>({id:d.id, ...d.data()}));
      render();
    }, err=>{
      console.warn("Shifts listener error:", err);
      showToast("Imeshindwa kupakia shift: "+(err.message||err));
    }));
  }
  if(CAN_EXPENSES(state.profile.role)){
    unsubs.push(onSnapshot(query(collection(db,"expenses"), orderBy("createdAt","desc")), snap=>{
      state.expenses = snap.docs.map(d=>({id:d.id, ...d.data()})); render();
    }));
  }
  if(CAN_PURCHASE(state.profile.role)){
    unsubs.push(onSnapshot(query(collection(db,"purchases"), orderBy("createdAt","desc")), snap=>{
      state.purchases = snap.docs.map(d=>({id:d.id, ...d.data()})); render();
    }));
    unsubs.push(onSnapshot(query(collection(db,"stockMoves"), orderBy("createdAt","desc")), snap=>{
      state.stockMoves = snap.docs.map(d=>({id:d.id, ...d.data()})); render();
    }));
  }
  // Rules: customers read = Admin/Manager/Cashier only. Inventory must not subscribe.
  if(CAN_SEE_CUSTOMERS(state.profile.role)){
    unsubs.push(onSnapshot(collection(db,"customers"), snap=>{
      state.customers = snap.docs.map(d=>({id:d.id, ...d.data()}));
      if(!state.saleCustomerId && state.customers[0]) state.saleCustomerId = state.customers[0].id;
      render();
    }));
    // Customer credit (Deni) payment history — same visibility as customers themselves.
    unsubs.push(onSnapshot(query(collection(db,"customerPayments"), orderBy("createdAt","desc")), snap=>{
      state.customerPayments = snap.docs.map(d=>({id:d.id, ...d.data()}));
      render();
    }, err=>console.warn("Customer payments listener:", err)));
  }
  // Sales visibility is separate from sale creation:
  // Admin + Manager can monitor all sales; Cashier sees only their own sales.
  // IMPORTANT: the Cashier query must include cashierUid == current UID so Firestore
  // can prove every returned document is allowed by the security rules.
  if(CAN_VIEW_SALES(state.profile.role)){
    const salesQuery = state.profile.role === "Cashier"
      ? query(collection(db,"sales"), where("cashierUid","==",state.user.uid))
      : query(collection(db,"sales"), orderBy("createdAt","desc"));
    unsubs.push(onSnapshot(salesQuery, snap=>{
      state.sales = snap.docs.map(d=>({id:d.id, ...d.data()}));
      state.sales.sort((a,b)=>{
        const ta = a.createdAt?.seconds ? a.createdAt.seconds : 0;
        const tb = b.createdAt?.seconds ? b.createdAt.seconds : 0;
        return tb-ta || String(b.date||"").localeCompare(String(a.date||""));
      });
      render();
    }, err=>{
      console.error("Sales listener:",err);
      showToast("Imeshindikana kupakia mauzo yako: "+err.message);
    }));
  }
  // Rules: auditLog read = Admin OR Manager. Users list = Admin only (Manage Users).
  if(CAN_MANAGE_USERS(state.profile.role)){
    unsubs.push(onSnapshot(collection(db,"Users"), snap=>{
      state.users = snap.docs.map(d=>({id:d.id, ...d.data()}));
      render();
    }));
  }
  if(CAN_VIEW_USERS(state.profile.role)){
    unsubs.push(onSnapshot(query(collection(db,"auditLog"), orderBy("createdAt","desc")), snap=>{
      state.auditLog = snap.docs.map(d=>({id:d.id, ...d.data()}));
      render();
    }));
  }
}

/* ============================================================
   LOW STOCK BROWSER NOTIFICATIONS
============================================================ */
function requestNotifyPermission(){
  if(typeof Notification === "undefined") return;
  if(Notification.permission === "default") Notification.requestPermission();
}
function checkLowStockNotifications(products){
  if(typeof Notification === "undefined" || Notification.permission !== "granted") return;
  if(!state.profile || !CAN_EDIT_PRODUCTS(state.profile.role)) return;
  products.forEach(p=>{
    const isLow = (p.stock||0) <= (p.reorder||0);
    if(isLow && !notifiedLowStockIds.has(p.id)){
      notifiedLowStockIds.add(p.id);
      try{
        new Notification("Stock ndogo — "+p.name, {
          body: `Kimebaki ${p.stock} tu (reorder ${p.reorder}). Agiza tena.`,
          tag: "lowstock-"+p.id
        });
      }catch(e){ /* ignore notification errors */ }
    } else if(!isLow && notifiedLowStockIds.has(p.id)){
      notifiedLowStockIds.delete(p.id); // stock replenished — allow re-alert if it drops again
    }
  });
}

async function logAudit(action, details){
  try{
    await addDoc(collection(db,"auditLog"), {
      action, details,
      actorUid: state.user?.uid || "",
      actorName: state.profile?.name || "System",
      date: todayISO(),
      time: new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}),
      createdAt: serverTimestamp()
    });
  }catch(e){ console.error("Audit log failed", e); }
}

async function doLogin(email, password){
  state.loginError = "";
  try{
    await signInWithEmailAndPassword(auth, email, password);
  }catch(e){
    state.loginError = "Email au password si sahihi.";
    render();
  }
}
async function doLogout(){
  clearListeners();
  await signOut(auth);
  state = {...state, tab:"dashboard", products:[], customers:[], customerPayments:[], sales:[], users:[], expenses:[], purchases:[], stockMoves:[], cart:[], globalThemeOverride:null, themeControlLoaded:false};
}

function updateConnectionBadge(){
  const el=document.getElementById("connectionBadge");
  if(!el) return;
  const online=navigator.onLine;
  el.textContent=online ? (offlineEnabled ? "ONLINE • SYNC READY" : "ONLINE") : "OFFLINE • LOCAL QUEUE";
  el.className="connection-badge "+(online?"online":"offline");
}
window.addEventListener("online", updateConnectionBadge);
window.addEventListener("offline", updateConnectionBadge);

/* ============================================================
   RENDER ROOT
============================================================ */
function initRevenueChartAnimation(){
  const stage=document.getElementById("revenueChartStage");
  const btn=document.getElementById("replayRevenueChart");
  if(!stage || !btn) return;
  btn.onclick=()=>{
    stage.classList.remove("is-replaying");
    void stage.offsetWidth;
    stage.classList.add("is-replaying");
    requestAnimationFrame(()=>{
      stage.classList.remove("is-replaying");
    });
  };
}

function render(){
  const root = document.getElementById("root");
  if(state.profile){
    const allowedTabs=new Set(navItemsFor(state.profile.role).map(x=>x.key));
    if(!allowedTabs.has(state.tab)) state.tab="dashboard";
  }
  if(!state.authReady){ root.innerHTML = `<div class="login-loading">Loading...</div>`; return; }
  if(state.user && !state.profile && state.profileLoadError){
    root.innerHTML = `
    <div class="login-wrap">
      <div class="login-card">
        <div class="login-logo">${ICONS.lock}</div>
        <h1 class="login-title">SALES MANAGEMENT SYSTEM</h1>
        <p class="login-sub">Imeshindwa kupakia akaunti yako</p>
        <div class="login-error" style="display:block;">${escapeHtml(state.profileLoadError)}</div>
        <button type="button" class="btn btn-amber btn-block" id="retryProfileBtn">Jaribu Tena</button>
        <p style="text-align:center;font-size:11.5px;color:var(--muted);margin-top:16px;cursor:pointer;text-decoration:underline;" id="retryLogoutLink">
          Toka na uingie tena
        </p>
      </div>
    </div>`;
    document.getElementById("retryProfileBtn")?.addEventListener("click", retryProfileLoad);
    document.getElementById("retryLogoutLink")?.addEventListener("click", async ()=>{ await signOut(auth); });
    return;
  }
  if(!state.user || !state.profile){ root.innerHTML = renderLogin(); bindLoginEvents(); return; }
  root.innerHTML = renderApp();
  bindAppEvents();
  requestAnimationFrame(initRevenueChartAnimation);
}

/* ============================================================
   LOGIN VIEW
============================================================ */
function renderLogin(){
  return `
  <div class="login-wrap">
    <div class="login-card">
      <div class="login-logo">${ICONS.lock}</div>
      <h1 class="login-title">SALES MANAGEMENT SYSTEM</h1>
      <p class="login-sub">Ingia kwa akaunti yako ya kazi</p>
      <div class="login-error" id="loginErrorBox">${escapeHtml(state.loginError)}</div>
      <form id="loginForm">
        <div class="field"><label>Email</label><input required type="email" id="loginEmail" placeholder="jina@mfano.com"/></div>
        <div class="field"><label>Password</label><input required type="password" id="loginPassword" placeholder="••••••••"/></div>
        <button type="submit" class="btn btn-amber btn-block" id="loginBtn">Ingia</button>
      </form>
      <p style="text-align:center;font-size:11.5px;color:var(--muted);margin-top:16px;">
        Umesahau password? Wasiliana na Admin wako.
      </p>
      <p style="text-align:center;font-size:9.5px;color:#d6d3d1;margin-top:10px;">SALES MANAGEMENT SYSTEM V1.4 • Professional Edition</p>
    </div>
  </div>`;
}
function bindLoginEvents(){
  const errBox = document.getElementById("loginErrorBox");
  if(state.loginError) errBox.style.display = "block"; else errBox.style.display="none";
  document.getElementById("loginForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const email = document.getElementById("loginEmail").value.trim();
    const password = document.getElementById("loginPassword").value;
    const btn = document.getElementById("loginBtn");
    btn.disabled = true; btn.innerHTML = `<span class="spinner"></span> Inaingia...`;
    await doLogin(email, password);
    if(document.getElementById("loginBtn")){ btn.disabled=false; btn.textContent="Ingia"; }
  });
}

/* ============================================================
   APP SHELL
============================================================ */
/* ============================================================
   ROLE-BASED UI
   Each role gets only the navigation relevant to its job.
   Firestore Rules remain the final security boundary.
============================================================ */
const ROLE_NAV = {
  Administrator: [
    ["dashboard","Dashboard","dashboard"],["sale","New Sale","sale"],
    ["inventory","Inventory","inventory"],["customers","Customers","customers"],
    ["suppliers","Suppliers","truck"],["history","Sales History","history"],
    ["receipts","Receipt Center","receipt"],["report","Monthly Report","report"],
    ["shiftsreport","Shifts Report","clock"],["expenses","Expenses","receipt"],
    ["purchases","Purchases","truck"],["backup","Backup & Export","history"],
    ["users","Manage Users","users"],["audit","Audit Log","history"],
    ["chat","Staff Chat","chat"],["support","Support","support"]
  ],
  Manager: [
    ["dashboard","Dashboard","dashboard"],
    ["inventory","Inventory","inventory"],["customers","Customers","customers"],
    ["suppliers","Suppliers","truck"],["history","Sales History","history"],
    ["receipts","Receipt Center","receipt"],["myshift","My Shift","clock"],
    ["report","Monthly Report","report"],["shiftsreport","Shifts Report","clock"],
    ["expenses","Expenses","receipt"],["purchases","Purchases","truck"],
    ["backup","Backup & Export","history"],["users","Manage Users","users"],["audit","Audit Log","history"],
    ["chat","Staff Chat","chat"],["support","Support","support"]
  ],
  Cashier: [
    ["dashboard","Dashboard","dashboard"],["sale","New Sale","sale"],
    ["inventory","Inventory","inventory"],["customers","Customers","customers"],
    ["history","Sales History","history"],["receipts","Receipt Center","receipt"],
    ["myshift","My Shift","clock"],["chat","Staff Chat","chat"],["support","Support","support"]
  ],
  Storekeeper: [
    ["dashboard","Dashboard","dashboard"],["inventory","Inventory","inventory"],
    ["suppliers","Suppliers","truck"],["purchases","Purchases","truck"],
    ["chat","Staff Chat","chat"],["support","Support","support"]
  ]
};

function navItemsFor(role){
  return (ROLE_NAV[role] || ROLE_NAV.Cashier).map(([key,label,icon])=>({key,label,icon}));
}


/* ============================================================
   THEME SYSTEM — ORIGINAL + ADMIN GLOBAL HOLIDAY THEMES
============================================================ */

const ORIGINAL_THEMES = [
  {id:"ocean", name:"Ocean", icon:"🌊", desc:"Coastal glass • calm professional"},
  {id:"emerald", name:"Emerald", icon:"💚", desc:"Botanical luxury • fresh business"},
  {id:"royal", name:"Royal", icon:"👑", desc:"Velvet purple • elegant premium"},
  {id:"sunset", name:"Sunset", icon:"🌅", desc:"Warm horizon • amber glow"},
  {id:"mint", name:"Mint", icon:"🍃", desc:"Fresh glass • modern wellness"}
];

const HOLIDAY_THEMES = [
  {id:"christmas", name:"Christmas", icon:"🎄", desc:"Red • green • gold festive"},
  {id:"newyear", name:"New Year", icon:"🎆", desc:"Midnight • violet • champagne"},
  {id:"eidfitr", name:"Eid al-Fitr", icon:"🌙", desc:"Emerald • teal • gold"},
  {id:"eidadha", name:"Eid al-Adha", icon:"🕌", desc:"Teal • emerald • gold"},
  {id:"ramadan", name:"Ramadan", icon:"🌙", desc:"Indigo • emerald • gold"},
  {id:"easter", name:"Easter", icon:"🌸", desc:"Lavender • rose • teal"},
  {id:"independence", name:"Independence Day", icon:"🇹🇿", desc:"Tanzania • green • yellow • black • blue"},
  {id:"union", name:"Union Day", icon:"🇹🇿", desc:"Tanzania • green • yellow • black • blue"},
  {id:"sabasaba", name:"Saba Saba", icon:"🏭", desc:"Tanzania Trade Fair • green • gold • blue"},
  {id:"naneNane", name:"Nane Nane", icon:"🌾", desc:"Agriculture • earth • gold"},
  {id:"labour", name:"Labour Day", icon:"🛠️", desc:"Workers • red • gold • Tanzania"},
  {id:"revolution", name:"Zanzibar Revolution", icon:"🇹🇿", desc:"Zanzibar • green • blue • gold"},
  {id:"karume", name:"Karume Day", icon:"🇹🇿", desc:"Zanzibar • green • blue • gold"},
  {id:"nyerere", name:"Nyerere Day", icon:"🇹🇿", desc:"Tanzania • green • gold • blue"},
  {id:"boxing", name:"Boxing Day", icon:"🎁", desc:"Festive • emerald • gold"},
  {id:"maulid", name:"Maulid", icon:"🌙", desc:"Emerald • gold • midnight"},
  {id:"valentine", name:"Valentine", icon:"❤️", desc:"Burgundy • rose • champagne"}
];

function themeCard(theme, active, disabled=false, actionAttr="") {
  const a = theme.previewA || theme.id;
  const b = theme.previewB || theme.id;
  return `<button type="button" class="theme-choice ${active ? 'active' : ''}" ${disabled ? 'disabled' : ''} ${actionAttr}
    style="--theme-preview-a:${escapeHtml(theme.previewA || 'var(--accent)')};--theme-preview-b:${escapeHtml(theme.previewB || 'var(--accent-strong)')}">
    <span class="theme-choice-icon">${theme.icon}</span>
    <span class="theme-choice-copy"><b>${escapeHtml(theme.name)}</b><small>${escapeHtml(theme.desc)}</small></span>
    ${active ? '<span class="theme-choice-check">✓</span>' : ''}
  </button>`;
}

const ALL_THEMES = [...ORIGINAL_THEMES, ...HOLIDAY_THEMES];
const THEME_IDS = new Set(ALL_THEMES.map(x=>x.id));
const HOLIDAY_IDS = new Set(HOLIDAY_THEMES.map(x=>x.id));

function themeInfo(id){
  return ALL_THEMES.find(x=>x.id===id) || ORIGINAL_THEMES[0];
}
function isValidTheme(id){ return THEME_IDS.has(id); }

function getThemeStorageKey(){
  return state.user?.uid ? `duka_theme_${state.user.uid}` : "duka_theme";
}
function getSavedTheme(){
  const key=getThemeStorageKey();
  const saved=localStorage.getItem(key) || (key!=="duka_theme" ? localStorage.getItem("duka_theme") : null);
  // Light/Dark were removed in V17.3. Migrate old saved choices to Ocean.
  if(saved === "light" || saved === "dark") return "ocean";
  if(isValidTheme(saved) && !HOLIDAY_IDS.has(saved)) return saved;
  return "ocean";
}

/* ============================================================
   AUTOMATIC HOLIDAY THEME ENGINE
   Active window: 1 day before the holiday through 2 days after.
   Date is evaluated in Tanzania time (Africa/Dar_es_Salaam).
============================================================ */
const AUTO_HOLIDAY_FIXED = [
  ["newyear","01-01"],
  ["revolution","01-12"],
  ["union","04-26"],
  ["labour","05-01"],
  ["sabasaba","07-07"],
  ["naneNane","08-08"],
  ["nyerere","10-14"],
  ["independence","12-09"],
  ["christmas","12-25"],
  ["boxing","12-26"]
];

/* Lunar holidays can move with moon sighting, so yearly dates are kept
   in a small calendar table and can be updated when official dates change. */
const AUTO_HOLIDAY_MOVABLE = {
  2026: [
    ["eidfitr","2026-03-20"],
    ["eidadha","2026-05-27"],
    ["maulid","2026-08-25"]
  ],
  2027: [
    ["eidfitr","2027-03-09"],
    ["eidadha","2027-05-17"],
    ["maulid","2027-08-14"]
  ]
};

function tzTodayISO(){
  try{
    return new Intl.DateTimeFormat("en-CA",{
      timeZone:"Africa/Dar_es_Salaam",
      year:"numeric",month:"2-digit",day:"2-digit"
    }).format(new Date());
  }catch(e){
    return new Date().toISOString().slice(0,10);
  }
}
function isoAddDays(iso,days){
  const d=new Date(iso+"T00:00:00Z");
  d.setUTCDate(d.getUTCDate()+days);
  return d.toISOString().slice(0,10);
}
function easterSundayISO(year){
  const a=year%19,b=Math.floor(year/100),c=year%100,d=Math.floor(b/4),e=b%4,
    f=Math.floor((b+8)/25),g=Math.floor((b-f+1)/3),
    h=(19*a+b-d-g+15)%30,i=Math.floor(c/4),k=c%4,
    l=(32+2*e+2*i-h-k)%7,m=Math.floor((a+11*h+22*l)/451),
    month=Math.floor((h+l-7*m+114)/31),day=((h+l-7*m+114)%31)+1;
  return `${year}-${String(month).padStart(2,"0")}-${String(day).padStart(2,"0")}`;
}
function autoHolidayThemeForToday(){
  const today=tzTodayISO();
  const year=Number(today.slice(0,4));
  const candidates=[];

  AUTO_HOLIDAY_FIXED.forEach(([theme,md])=>{
    candidates.push({theme,date:`${year}-${md}`,priority:10});
  });
  (AUTO_HOLIDAY_MOVABLE[year]||[]).forEach(([theme,date])=>{
    candidates.push({theme,date,priority:20});
  });

  // Good Friday and Easter Monday use the Easter theme.
  const easter=easterSundayISO(year);
  candidates.push({theme:"easter",date:isoAddDays(easter,-2),priority:30});
  candidates.push({theme:"easter",date:isoAddDays(easter,1),priority:30});

  // Valentine is included as a festive theme using the same requested window.
  candidates.push({theme:"valentine",date:`${year}-02-14`,priority:5});

  const matches=candidates.filter(x=>{
    const start=isoAddDays(x.date,-1);
    const end=isoAddDays(x.date,2);
    return today>=start && today<=end;
  }).sort((a,b)=>b.priority-a.priority || a.date.localeCompare(b.date));

  return matches.length && HOLIDAY_IDS.has(matches[0].theme) ? matches[0].theme : null;
}

function effectiveTheme(){
  // Automatic holiday mode has priority during its configured window.
  const autoHoliday=autoHolidayThemeForToday();
  if(autoHoliday) return autoHoliday;

  // Administrator-selected holiday themes must actually become the active
  // visual theme for everyone, not only change the header/icon.
  if(state.globalThemeOverride && HOLIDAY_IDS.has(state.globalThemeOverride)){
    return state.globalThemeOverride;
  }

  // When no holiday is active, everyone returns to their saved normal theme.
  return state.localTheme || getSavedTheme();
}

function applyTheme(theme){
  const t=(theme === "light" || theme === "dark") ? "ocean" : (isValidTheme(theme) ? theme : "ocean");
  document.documentElement.setAttribute("data-theme",t);
  if(!HOLIDAY_IDS.has(t)){
    state.localTheme=t;
    localStorage.setItem(getThemeStorageKey(),t);
  }
  document.documentElement.setAttribute(
    "data-holiday-override",
    (autoHolidayThemeForToday() || (state.globalThemeOverride && HOLIDAY_IDS.has(state.globalThemeOverride))) ? "on" : "off"
  );
}

function refreshEffectiveTheme(){ applyTheme(effectiveTheme()); }

let __holidayThemeLastDate = tzTodayISO();
setInterval(()=>{
  const nowDate=tzTodayISO();
  if(nowDate!==__holidayThemeLastDate){
    __holidayThemeLastDate=nowDate;
    refreshEffectiveTheme();
    if(typeof render==="function") render();
  }
},60000);

function toggleTheme(){ openThemeChooser(); }

function closeThemeChooser(){
  const modal=document.getElementById("themeChooserOverlay");
  if(modal) modal.remove();
}


function openThemeChooser(){
  if(!state.profile) return;
  const isAdmin=state.profile.role==="Administrator";
  const current=effectiveTheme();
  const global=state.globalThemeOverride && HOLIDAY_IDS.has(state.globalThemeOverride)
    ? themeInfo(state.globalThemeOverride) : null;

  const originalHtml=ORIGINAL_THEMES.map(t=>
    themeCard(t,current===t.id && !global,false,`data-local-theme="${t.id}"`)
  ).join("");

  const holidayHtml=isAdmin ? HOLIDAY_THEMES.map(t=>
    themeCard(t,current===t.id && !!global,false,`data-global-holiday="${t.id}"`)
  ).join("") : "";

  document.getElementById("modalRoot").innerHTML=`
    <div class="modal-overlay theme-modal-overlay" id="themeChooserOverlay">
      <div class="modal theme-modal">
        <div class="modal-head">
          <div>
            <h3>🎨 SALES MANAGEMENT SYSTEM Themes</h3>
            <div class="muted theme-modal-sub">
              ${global
                ? `${global.icon} <b>${escapeHtml(global.name)}</b> imewashwa na Administrator kwa wafanyakazi wote.`
                : "Kila mfanyakazi anaweza kuchagua theme yake mwenyewe."}
            </div>
          </div>
          <button class="icon-btn" id="closeThemeChooserBtn">${ICONS.close}</button>
        </div>
        <div class="modal-body">
          <div class="theme-section">
            <div class="theme-section-title">✨ Original Themes <span>Personal • Kila Mfanyakazi</span></div>
            <div class="theme-grid">${originalHtml}</div>
          </div>
          ${isAdmin ? `
          <div class="theme-section holiday-section">
            <div class="theme-section-title">🎉 Holiday Themes <span>Administrator • Global kwa Wote</span></div>
            <div class="theme-grid">${holidayHtml}</div>
            <div class="theme-admin-note">
              <b>Automatic Holiday Mode:</b> Sikukuu hujiwasha yenyewe kuanzia siku 1 kabla
              hadi siku 2 baada ya sikukuu. Baada ya hapo kila mfanyakazi anarudi kwenye Normal Theme yake.
              Administrator bado anaweza kuchagua holiday theme manually.
            </div>
            <button class="btn btn-outline btn-block" id="returnNormalThemesBtn">
              ↩️ Rudisha Normal Themes kwa Wafanyakazi Wote
            </button>
          </div>` : `
          <div class="theme-admin-note staff-theme-note">
            🎉 Holiday themes zinawezeshwa na Administrator pekee. Theme yako binafsi
            itarudi moja kwa moja pale Administrator atakaporudisha Normal Themes.
          </div>`}
        </div>
      </div>
    </div>`;

  const overlay=document.getElementById("themeChooserOverlay");
  document.getElementById("closeThemeChooserBtn").onclick=closeThemeChooser;
  overlay.onclick=e=>{if(e.target===overlay) closeThemeChooser();};

  document.querySelectorAll("[data-local-theme]").forEach(btn=>{
    btn.addEventListener("click",()=>{
      const id=btn.dataset.localTheme;
      if(!isValidTheme(id) || HOLIDAY_IDS.has(id)) return;
      state.localTheme=id;
      localStorage.setItem(getThemeStorageKey(),id);
      refreshEffectiveTheme();
      closeThemeChooser();
      render();
    });
  });

  if(isAdmin){
    document.querySelectorAll("[data-global-holiday]").forEach(btn=>{
      btn.addEventListener("click",()=>setGlobalHolidayTheme(btn.dataset.globalHoliday));
    });
    const normalBtn=document.getElementById("returnNormalThemesBtn");
    if(normalBtn) normalBtn.onclick=clearGlobalHolidayTheme;
  }
}

async function setGlobalHolidayTheme(theme){
  if(state.profile?.role!=="Administrator" || !HOLIDAY_IDS.has(theme)) return;
  try{
    await setDoc(doc(db,"settings","themeControl"),{
      active:true, theme,
      updatedBy:state.user.uid,
      updatedByName:state.profile.name,
      updatedAt:serverTimestamp()
    },{merge:true});
    state.globalThemeOverride=theme;
    state.themeControlLoaded=true;
    refreshEffectiveTheme();
    closeThemeChooser();
    showToast(`${themeInfo(theme).icon} ${themeInfo(theme).name} imewashwa kwa wafanyakazi wote.`);
    render();
  }catch(e){
    console.error("Global theme save failed:",e);
    showToast("Theme haijahifadhiwa: "+(e.message||e));
  }
}

async function clearGlobalHolidayTheme(){
  if(state.profile?.role!=="Administrator") return;
  try{
    await setDoc(doc(db,"settings","themeControl"),{
      active:false, theme:"",
      updatedBy:state.user.uid,
      updatedByName:state.profile.name,
      updatedAt:serverTimestamp()
    },{merge:true});
    state.globalThemeOverride=null;
    state.themeControlLoaded=true;
    refreshEffectiveTheme();
    closeThemeChooser();
    showToast("Normal Themes zimerudishwa. Kila mfanyakazi atatumia theme yake.");
    render();
  }catch(e){
    console.error("Global theme reset failed:",e);
    showToast("Theme haikurudishwa: "+(e.message||e));
  }
}

state.localTheme=getSavedTheme();
applyTheme(effectiveTheme());
setTimeout(refreshEffectiveTheme,1000);


function renderApp(){
  const role = state.profile.role;
  const lowStock = state.products.filter(p=>p.stock<=p.reorder).length;
  const items = navItemsFor(role);

  const navHtml = items.map(item=>{
    const active = state.tab===item.key;
    const badge = item.key==="inventory" && lowStock>0 ? `<span class="nav-badge">${lowStock}</span>` : "";
    return `<button class="nav-item ${active?'active':''}" data-tab="${item.key}">${ICONS[item.icon]}<span>${item.label}</span>${badge}</button>`;
  }).join("");

  let viewHtml = "";
  if(state.tab==="dashboard") viewHtml = renderDashboard();
  else if(state.tab==="sale" && CAN_SELL(role)) viewHtml = renderSale();
  else if(state.tab==="inventory") viewHtml = renderInventory();
  else if(state.tab==="customers" && CAN_SEE_CUSTOMERS(role)) viewHtml = renderCustomers();
  else if(state.tab==="suppliers" && CAN_MANAGE_SUPPLIERS(role)) viewHtml = renderSuppliers();
  else if(state.tab==="history" && CAN_VIEW_SALES(role)) viewHtml = renderHistory();
  else if(state.tab==="receipts" && CAN_VIEW_SALES(role)) viewHtml = renderReceiptCenter();
  else if(state.tab==="myshift" && CAN_TRACK_SHIFT(role)) viewHtml = renderMyShift();
  else if(state.tab==="report" && CAN_SEE_FINANCE(role)) viewHtml = renderReport();
  else if(state.tab==="shiftsreport" && CAN_SEE_SHIFTS_REPORT(role)) viewHtml = renderShiftsReport();
  else if(state.tab==="expenses" && CAN_EXPENSES(role)) viewHtml = renderExpenses();
  else if(state.tab==="purchases" && CAN_PURCHASE(role)) viewHtml = renderPurchases();
  else if(state.tab==="backup" && CAN_BACKUP(role)) viewHtml = renderBackup();
  else if(state.tab==="users" && CAN_MANAGE_USERS(role)) viewHtml = renderUsers();
  else if(state.tab==="audit" && CAN_VIEW_USERS(role)) viewHtml = renderAuditLog();
  else if(state.tab==="chat") viewHtml = renderChat();
  else if(state.tab==="support") viewHtml = renderSupport();
  else viewHtml = `<div class="view"><p class="empty-note">Huna ruhusa ya sehemu hii.</p></div>`;

  return `
  <div class="app" data-tab="${escapeHtml(state.tab)}">
    <aside class="sidebar">
      <div class="sidebar-head">
        <div class="logo-badge">${ICONS.sale}</div>
        <div><div class="brand-name">SALES MANAGEMENT SYSTEM</div><div class="brand-sub">sales &amp; inventory</div></div>
      </div>
      <div class="sidebar-id-card">
        <div class="id-name">${escapeHtml(state.profile.name)}</div>
        <div class="id-role">${escapeHtml(role)}</div>
      </div>
      <nav class="sidebar-nav">${navHtml}</nav>
      <div class="sidebar-foot">
        <div class="user-chip">
          <div class="user-avatar">${initials(state.profile.name)}</div>
          <div>
            <div class="user-name">${escapeHtml(state.profile.name)}</div>
            <div class="user-role">${escapeHtml(role)}</div>
          </div>
        </div>
        <button class="logout-btn" id="changePasswordBtn" style="margin-bottom:8px;">${ICONS.lock||ICONS.edit} Badilisha Password</button>
        <button class="logout-btn" id="logoutBtn">${ICONS.logout} Toka</button>
      </div>
    </aside>
    <main>
      <div class="app-mainbar">
        <div class="theme-status">
          <span class="theme-status-icon">${state.globalThemeOverride ? themeInfo(state.globalThemeOverride).icon : "🎨"}</span>
          <span class="theme-status-copy">
            <b>${state.globalThemeOverride ? escapeHtml(themeInfo(state.globalThemeOverride).name) : "Personal Theme"}</b>
            <small>${state.globalThemeOverride ? "Global • Administrator" : "Your workspace theme"}</small>
          </span>
        </div>
        <button class="theme-toggle theme-control-btn" id="themeToggle" type="button"
                title="${state.profile.role==="Administrator" ? "Badilisha Original au Holiday Theme" : "Badilisha Personal Theme"}"
                aria-label="Badilisha theme">
          🎨 <span class="theme-control-text">Themes</span>
        </button>
        ${state.globalThemeOverride ? `<span class="global-theme-badge">${themeInfo(state.globalThemeOverride).icon} ${escapeHtml(themeInfo(state.globalThemeOverride).name)} • Global</span>` : ""}

        ${(() => {
          const notices = v15Notifications();
          const count = notices.length;
          const rows = notices.length
            ? notices.map(n => `<button class="v15-notify-row ${n.level||"amber"}" data-v15-notify-tab="${escapeHtml(n.tab||"dashboard")}">
                <b>${escapeHtml(n.title||"Warning")}</b>
                <small>${escapeHtml(n.text||"")}</small>
              </button>`).join("")
            : `<div class="v15-empty">Hakuna warning kwa sasa 🎉</div>`;
          return `<div class="v15-notify-wrap" id="v15NotifyWrap">
            <button class="v15-notify-btn" id="v15NotifyBtn" type="button" aria-label="Warnings na notifications" title="Warnings & Notifications">
              ${ICONS.alert}
              ${count ? `<span>${count > 99 ? "99+" : count}</span>` : ""}
            </button>
            <div class="v15-notify-panel" id="v15NotifyPanel" role="dialog" aria-label="Warnings na notifications">
              <div class="v15-notify-head"><b>Warnings & Notifications</b><span>${count} ${count===1?"alert":"alerts"}</span></div>
              ${rows}
            </div>
          </div>`;
        })()}

        <div id="connectionBadge" class="connection-badge online">ONLINE</div>
      </div>
      <div class="view">${viewHtml}</div>
      ${renderMobileBottomNav(role)}
    </main>
  </div>
  <div id="modalRoot"></div>`;
}

function renderMobileBottomNav(role){
  const btn=(key,label,icon,active)=>`<button class="mobile-bottom-nav-item ${active?'active':''}" data-bottom-tab="${key}">${ICONS[icon]}<span>${label}</span></button>`;
  let items=[btn("dashboard","Home","dashboard",state.tab==="dashboard")];
  if(role==="Cashier"){
    items.push(btn("sale","New sale","sale",state.tab==="sale"));
    items.push(btn("history","Sales","history",state.tab==="history"));
  }else if(role==="Storekeeper"){
    items.push(btn("inventory","Stock","inventory",state.tab==="inventory"));
    items.push(btn("purchases","Purchases","truck",state.tab==="purchases"));
  }else if(role==="Manager"){
    items.push(btn("history","Sales","history",state.tab==="history"));
    items.push(btn("inventory","Stock","inventory",state.tab==="inventory"));
  }else{
    items.push(btn("sale","New sale","sale",state.tab==="sale"));
    items.push(btn("history","Sales","history",state.tab==="history"));
  }
  items.push(`<button class="mobile-bottom-nav-item" id="bottomNavMore">${ICONS.menu || `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 6h16M4 12h16M4 18h16"/></svg>`}<span>More</span></button>`);
  return `<nav class="mobile-bottom-nav">${items.join("")}</nav>`;
}

function bindAppEvents(){
  document.querySelectorAll(".nav-item").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      state.tab = btn.dataset.tab;
      state.cart=[];
      state.discountType="none";
      state.discountValue="";
      render();
    });
  });

  const logoutBtn = document.getElementById("logoutBtn");
  if(logoutBtn) logoutBtn.addEventListener("click", async ()=>{ await doLogout(); render(); });
  const changePwBtn = document.getElementById("changePasswordBtn");
  if(changePwBtn) changePwBtn.addEventListener("click", openChangePasswordModal);

  const dashNewSaleBtn = document.getElementById("dashNewSaleBtn");
  if(dashNewSaleBtn) dashNewSaleBtn.addEventListener("click", ()=>{ state.tab="sale"; render(); });

  const tickerViewAllBtn = document.getElementById("tickerViewAllBtn");
  if(tickerViewAllBtn) tickerViewAllBtn.addEventListener("click", ()=>{ state.tab="history"; render(); });

  const themeToggle = document.getElementById("themeToggle");
  if(themeToggle) themeToggle.addEventListener("click", toggleTheme);

  // Mobile hamburger toggle imeondolewa — "More" kwenye bottom nav
  // ndiyo pekee inayofungua/kufunga sidebar sasa (tazama chini).
  const sidebar = document.querySelector(".sidebar");

  // Mobile bottom nav bar
  document.querySelectorAll(".mobile-bottom-nav-item[data-bottom-tab]").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      state.tab = btn.dataset.bottomTab;
      state.cart=[];
      state.discountType="none";
      state.discountValue="";
      render();
    });
  });
  const bottomNavMore = document.getElementById("bottomNavMore");
  if(bottomNavMore && sidebar){
    bottomNavMore.addEventListener("click", ()=>{
      sidebar.classList.toggle("expanded");
    });
  }

  bindViewEvents();
}

function showToast(msg){
  const root = document.getElementById("toastRoot");
  if(!root) return;
  root.innerHTML = `<div class="toast">${escapeHtml(msg)}</div>`;
  setTimeout(()=>{ if(root) root.innerHTML=""; }, 2400);
}

/* ============================================================
   DASHBOARD
============================================================ */
function creditReminderBanner(){
  if(!CAN_SEE_CUSTOMERS(state.profile.role)) return "";
  const today = todayISO();
  const due = state.customers
    .filter(c=>(Number(c.balance)||0)>0 && c.nextDueDate)
    .map(c=>({...c, diffDays: Math.round((new Date(c.nextDueDate)-new Date(today))/86400000)}))
    .filter(c=>c.diffDays<=3)
    .sort((a,b)=>a.diffDays-b.diffDays);
  if(!due.length) return "";
  return `<div class="v15-info" style="background:#fff7ed;border:1px solid #fed7aa;margin-bottom:16px;display:block;">
    <div style="font-weight:800;margin-bottom:8px;">🔔 Ukumbusho wa Malipo ya Wateja (Deni)</div>
    ${due.map(c=>{
      const label = c.diffDays<0 ? `Imechelewa siku ${Math.abs(c.diffDays)}` : c.diffDays===0 ? "Inatakiwa Leo" : c.diffDays===1 ? "Kesho" : `Siku ${c.diffDays} zijazo`;
      const color = c.diffDays<0 ? "var(--red)" : "#b45309";
      return `<div class="reminder-row" data-remind-customer="${c.id}" style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-top:1px dashed #fed7aa;cursor:pointer;">
        <span>${escapeHtml(c.name)} — <span style="color:${color};font-weight:700;">${label}</span> <span class="muted">(${escapeHtml(c.nextDueDate)})</span></span>
        <span class="mono" style="font-weight:700;">${fmt(c.balance)}</span>
      </div>`;
    }).join("")}
  </div>`;
}

function renderDashboard(){
  const role = state.profile.role;
  const totalRevenue = state.sales.reduce((sum,x)=>sum+(Number(x.total)||0),0);
  const totalOrders = state.sales.length;
  const avgOrder = totalOrders ? totalRevenue/totalOrders : 0;
  const lowStock = state.products.filter(p=>Number(p.stock||0)<=Number(p.reorder||0)).length;
  const stockUnits = state.products.reduce((sum,p)=>sum+(Number(p.stock)||0),0);
  const mySales = role==="Cashier" ? state.sales.filter(s=>s.cashierUid===state.user.uid) : state.sales;
  const myTotal = mySales.reduce((sum,x)=>sum+(Number(x.total)||0),0);

  if(role==="Administrator" || role==="Manager"){
    const days=[...Array(14)].map((_,i)=>{
      const d=new Date(); d.setDate(d.getDate()-(13-i));
      return d.toISOString().slice(0,10);
    });
    const chartData=days.map(d=>({
      date:d.slice(5),
      revenue:state.sales.filter(s=>s.date===d).reduce((sum,s)=>sum+(Number(s.total)||0),0)
    }));
    return `
      <div class="view-header"><div>
        <p class="view-eyebrow">Management workspace</p>
        <h1>Dashboard</h1>
        <p>Karibu, ${escapeHtml(state.profile.name.split(" ")[0])}. Taarifa za biashara na uendeshaji kwa nafasi yako.</p>
      </div></div>
      ${CAN_SELL(role) ? `<button class="hero-cta" id="dashNewSaleBtn">${ICONS.sale} New sale</button>` : ""}
      ${creditReminderBanner()}
      <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Total Revenue</span><div class="kpi-icon green">${ICONS.trend}</div></div><div class="kpi-value">${fmt(totalRevenue)}</div><div class="kpi-sub">${totalOrders} sales recorded</div></div>
        <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Average Revenue</span><div class="kpi-icon slate">${ICONS.receipt}</div></div><div class="kpi-value">${fmt(avgOrder)}</div></div>
        <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Customers</span><div class="kpi-icon amber">${ICONS.customers}</div></div><div class="kpi-value">${state.customers.length}</div></div>
        <button class="kpi-card" id="lowStockCard" style="cursor:pointer;"><div class="kpi-top"><span class="kpi-label">Low Stock</span><div class="kpi-icon ${lowStock?'red':'slate'}">${ICONS.alert}</div></div><div class="kpi-value">${lowStock}</div><div class="kpi-sub">tap to view inventory</div></button>
      </div>
      <div class="v17-target-card ${state.businessTarget>0 && v15TodayRevenue()>=state.businessTarget ? "done" : ""}" id="v17SalesTargetCard">
        <div class="v17-target-top">
          <div>
            <div class="v17-target-title"><span>🎯</span> Target ya Mauzo ya Leo</div>
            <div class="v17-target-sub">Lengo la mauzo la siku hii — progress inasasishwa moja kwa moja.</div>
          </div>
          <div class="v17-target-actions">
            <div class="v17-target-amount"><strong>${state.businessTarget>0 ? fmt(v15TodayRevenue()) : fmt(v15TodayRevenue())}</strong><small>mauzo leo</small></div>
            <div class="v17-target-badge ${state.businessTarget>0 && v15TodayRevenue()>=state.businessTarget ? "done" : ""}">${state.businessTarget>0 ? (v15TodayRevenue()>=state.businessTarget ? "🎉 TARGET IMEFIKIWA" : Math.round(v15TargetPercent())+"%") : "Haijawekwa"}</div>
            <button class="btn btn-outline" id="v15TargetBtn" type="button">${state.businessTarget>0 ? "Badilisha Target" : "Weka Target"}</button>
          </div>
        </div>
        <div class="v17-target-progress"><span style="width:${v15TargetPercent()}%"></span></div>
        <div class="v17-target-meta"><span>Target: <b>${state.businessTarget>0 ? fmt(state.businessTarget) : "Haijawekwa"}</b></span><span>${state.businessTarget>0 ? (v15TodayRevenue()>=state.businessTarget ? "Umevuka lengo kwa "+fmt(v15TodayRevenue()-state.businessTarget) : "Bado "+fmt(Math.max(0,state.businessTarget-v15TodayRevenue()))+" kufikia target") : "Administrator/Manager anaweza kuweka target."}</span></div>
      </div>
      <div class="grid-2">
        <div class="card"><h2>Revenue — last 14 days</h2><div class="chart-wrap">${renderLineChart(chartData)}</div></div>
        ${renderTicker()}
      </div>`;
  }

  if(role==="Cashier"){
    return `
      <div class="view-header"><div>
        <p class="view-eyebrow">Cashier workspace</p>
        <h1>Dashboard</h1>
        <p>Karibu, ${escapeHtml(state.profile.name.split(" ")[0])}. Hapa unaona mauzo yako na taarifa za kazi yako.</p>
      </div></div>
      <button class="hero-cta" id="dashNewSaleBtn">${ICONS.sale} New sale</button>
      ${creditReminderBanner()}
      <div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Mauzo Yangu</span><div class="kpi-icon green">${ICONS.trend}</div></div><div class="kpi-value">${fmt(myTotal)}</div><div class="kpi-sub">${mySales.length} sales</div></div>
        <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Products</span><div class="kpi-icon slate">${ICONS.inventory}</div></div><div class="kpi-value">${state.products.length}</div></div>
        <button class="kpi-card" id="lowStockCard" style="cursor:pointer;"><div class="kpi-top"><span class="kpi-label">Low Stock</span><div class="kpi-icon ${lowStock?'red':'slate'}">${ICONS.alert}</div></div><div class="kpi-value">${lowStock}</div><div class="kpi-sub">tap to view inventory</div></button>
      </div>
      <div class="grid-2" style="grid-template-columns:1fr;">${renderTicker()}</div>`;
  }

  return `
    <div class="view-header"><div>
      <p class="view-eyebrow">Inventory workspace</p>
      <h1>Dashboard</h1>
      <p>Karibu, ${escapeHtml(state.profile.name.split(" ")[0])}. Taarifa za stock, suppliers na purchases zinazokuhusu.</p>
    </div></div>
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Products</span><div class="kpi-icon slate">${ICONS.inventory}</div></div><div class="kpi-value">${state.products.length}</div><div class="kpi-sub">${stockUnits} units in stock</div></div>
      <button class="kpi-card" id="lowStockCard" style="cursor:pointer;"><div class="kpi-top"><span class="kpi-label">Low Stock</span><div class="kpi-icon ${lowStock?'red':'slate'}">${ICONS.alert}</div></div><div class="kpi-value">${lowStock}</div><div class="kpi-sub">tap to view inventory</div></button>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Suppliers</span><div class="kpi-icon amber">${ICONS.truck}</div></div><div class="kpi-value">${state.suppliers.length}</div></div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Purchases</span><div class="kpi-icon green">${ICONS.receipt}</div></div><div class="kpi-value">${state.purchases.length}</div></div>
    </div>
    <div class="card"><h2>Stock workspace</h2><p class="empty-note">Tumia Inventory, Suppliers na Purchases kusimamia stock na bidhaa.</p></div>`;
}

function renderTicker(){
  const role = state.profile.role;
  const list = role==="Cashier" ? state.sales.filter(s=>s.cashierUid===state.user.uid) : state.sales;
  const recent = list.slice(0,8);
  const canViewAll = CAN_SELL(role);
  return `
    <div class="ticker">
      <h2>${ICONS.receipt} Recent sales${canViewAll ? `<button class="ticker-viewall" id="tickerViewAllBtn" type="button">View all</button>` : ""}</h2>
      <p class="ticker-desc">Latest completed transactions.</p>
      <div class="ticker-list">
        ${recent.length===0 ? `<p class="empty-note">Hakuna mauzo bado.</p>` :
          recent.map(s=>`<div class="ticker-row"><div class="ticker-row-top"><span>#${s.number}</span><span>${s.date}</span></div><div class="ticker-row-bottom"><span class="ticker-name">${escapeHtml(s.customerName)}</span><span class="ticker-amt">${fmt(s.total)}</span></div></div>`).join("")}
      </div>
    </div>`;
}
function renderLineChart(data){
  const w=560,h=210,padL=42,padB=22,padT=10,padR=8;
  const target = Number(state.businessTarget)||0;
  // Include today's target in the scale so the target line is always visible.
  const max = Math.max(1, ...data.map(d=>Number(d.revenue)||0), target);
  const stepX = (w-padL-padR)/(data.length-1||1);
  const plotH = h-padT-padB;
  const points = data.map((d,i)=>({x:padL+i*stepX, y:padT+plotH*(1-(Number(d.revenue)||0)/max), ...d}));
  const path = points.map((p,i)=>(i===0?"M":"L")+p.x.toFixed(1)+" "+p.y.toFixed(1)).join(" ");
  const gridY=[0,0.5,1].map(f=>padT+plotH*f);
  const targetY = target>0 ? padT+plotH*(1-target/max) : null;
  const targetMarkup = targetY===null ? "" : `
        <line class="chart-target-line" x1="${padL}" y1="${targetY}" x2="${w-padR}" y2="${targetY}" stroke="var(--green,#16a34a)" stroke-width="1.8" stroke-opacity=".78"/>
        <rect x="${w-padR-66}" y="${Math.max(2,targetY-11)}" width="62" height="17" rx="8.5" fill="var(--surface)" stroke="var(--green,#16a34a)" stroke-opacity=".35"/>
        <text class="chart-target-label" x="${w-padR-35}" y="${Math.max(13,targetY+1)}" font-size="8.5" fill="var(--green,#16a34a)" text-anchor="middle">TARGET ${fmt(target).replace('TZS ','')}</text>`;
  return `<div class="chart-stage" id="revenueChartStage">
    <button type="button" class="chart-replay-btn" id="replayRevenueChart">▶ Play animation</button>
    <svg class="chart" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" aria-label="Revenue chart ya siku 14">
      ${gridY.map(y=>`<line class="chart-grid" x1="${padL}" y1="${y}" x2="${w-padR}" y2="${y}" stroke="currentColor" stroke-opacity=".12" stroke-width="1"/>`).join("")}
      ${targetMarkup}
      <path class="chart-line-glow" d="${path}" fill="none" stroke="var(--accent)" stroke-opacity=".16" stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/>
      <path class="chart-line" d="${path}" fill="none" stroke="var(--accent)" stroke-width="3.4" stroke-linecap="round" stroke-linejoin="round"/>
      ${points.map((p,i)=>`<circle class="chart-point-ring" cx="${p.x}" cy="${p.y}" r="8" fill="none" stroke="var(--accent)" stroke-width="1.6" style="animation-delay:${350+i*145}ms"></circle><circle class="chart-point" cx="${p.x}" cy="${p.y}" r="3.4" fill="var(--accent)" style="animation-delay:${350+i*145}ms"><title>${p.date}: ${fmt(p.revenue)}${target>0 ? ` • Target ${fmt(target)}` : ""}</title></circle>`).join("")}
      ${points.map((p,i)=> i%2===0 ? `<text x="${p.x}" y="${h-4}" font-size="8.5" fill="currentColor" opacity=".62" text-anchor="middle">${p.date}</text>` : "").join("")}
      <text x="${padL-6}" y="${padT+6}" font-size="8.5" fill="currentColor" opacity=".62" text-anchor="end">${fmt(max).replace('TZS ','')}</text>
      <text x="${padL-6}" y="${h-padB}" font-size="8.5" fill="currentColor" opacity=".62" text-anchor="end">0</text>
    </svg>
  </div>`;
}

/* ============================================================
   INVENTORY
============================================================ */
function renderInventory(){
  const role = state.profile.role;
  const canEdit = CAN_EDIT_PRODUCTS(role);
  const canSeeFin = CAN_SEE_FINANCE(role);
  const q = state.invSearch.toLowerCase();
  const filtered = state.products.filter(p=>p.name.toLowerCase().includes(q)||p.sku.toLowerCase().includes(q)||(p.barcode||"").toLowerCase().includes(q));
  return `
    <div class="view-header">
      <div><h1>Inventory</h1><p>${state.products.length} products tracked</p></div>
      ${canEdit ? `<button class="btn btn-dark" id="addProductBtn">${ICONS.plus} Add product</button>` : ""}
    </div>
    <div class="search-wrap">${ICONS.search}<input id="invSearchInput" placeholder="Search products or SKU" value="${escapeHtml(state.invSearch)}"/></div>
    <div class="table-card">
      <table>
        <thead><tr><th>Product</th><th>Category</th>${canSeeFin?'<th class="right">Cost</th>':''}<th class="right">Selling Price</th><th class="right">Stock</th>${canEdit?'<th></th>':''}</tr></thead>
        <tbody>
          ${filtered.length===0 ? `<tr class="empty-row"><td colspan="6">No products found.</td></tr>` :
          filtered.map(p=>{
            const low = p.stock<=p.reorder;
            return `<tr>
              <td><div class="cell-title">${escapeHtml(p.name)}</div><div class="cell-sub">${escapeHtml(p.sku)}</div></td>
              <td>${escapeHtml(p.category)}</td>
              ${canSeeFin?`<td class="right mono">${fmt(p.costPrice||0)}</td>`:''}
              <td class="right mono">${fmt(p.price)}</td>
              <td class="right mono">${low?`<span class="pill-low">${p.stock}</span>`:p.stock}</td>
              ${canEdit?`<td><div class="row-actions">
                <button class="icon-btn" data-edit-product="${p.id}">${ICONS.edit}</button>
                <button class="icon-btn danger" data-del-product="${p.id}">${ICONS.trash}</button>
              </div></td>`:''}
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>
  `;
}

/* ============================================================
   CUSTOMERS
============================================================ */
function renderCustomers(){
  const canDelete = CAN_DELETE_CUSTOMER(state.profile.role); // Admin/Manager only, per rules
  return `
    <div class="view-header">
      <div><h1>Customers</h1><p>${state.customers.length} customers</p></div>
      <button class="btn btn-dark" id="addCustomerBtn">${ICONS.plus} Add customer</button>
    </div>
    <div class="table-card">
      <table>
        <thead><tr><th>Name</th><th>Phone</th><th class="right">Orders</th><th class="right">Total spent</th><th class="right">Deni</th><th class="right">Credit Limit</th><th>Tarehe ya Kulipa</th><th></th></tr></thead>
        <tbody>
          ${state.customers.map(c=>{
            const bal = Number(c.balance)||0;
            const limit = Number(c.creditLimit)||0;
            const overLimit = limit>0 && bal>limit;
            let dueLabel = "—";
            if(bal>0 && c.nextDueDate){
              const diffDays = Math.round((new Date(c.nextDueDate)-new Date(todayISO()))/86400000);
              const txt = diffDays<0 ? `Imechelewa siku ${Math.abs(diffDays)}` : diffDays===0 ? "Leo" : diffDays===1 ? "Kesho" : `Siku ${diffDays}`;
              const color = diffDays<=0 ? "var(--red)" : (diffDays<=3 ? "#b45309" : "inherit");
              dueLabel = `<span style="color:${color};font-weight:${diffDays<=3?700:400};">${escapeHtml(c.nextDueDate)}</span><br/><span class="muted" style="font-size:11px;">${txt}</span>`;
            }
            return `
            <tr>
              <td class="cell-title">${escapeHtml(c.name)}</td>
              <td class="mono" style="font-size:12px;">${escapeHtml(c.phone||"—")}</td>
              <td class="right mono">${c.orders||0}</td>
              <td class="right mono">${fmt(c.totalSpent||0)}</td>
              <td class="right mono" style="${bal>0?`color:${overLimit?'var(--red)':'#b45309'};font-weight:700;`:''}">${fmt(bal)}</td>
              <td class="right mono">${fmt(limit)}</td>
              <td style="font-size:12px;">${dueLabel}</td>
              <td><div class="row-actions">
                ${bal>0 ? `<button class="icon-btn" data-pay-customer="${c.id}" title="Lipa Deni">${ICONS.receipt}</button>` : ""}
                ${bal>0 ? `<button class="icon-btn" data-set-duedate="${c.id}" title="Weka/Badilisha Tarehe ya Kulipa">${ICONS.clock}</button>` : ""}
                <button class="icon-btn" data-edit-customer="${c.id}">${ICONS.edit}</button>
                ${canDelete ? `<button class="icon-btn danger" data-del-customer="${c.id}">${ICONS.trash}</button>` : ""}
              </div></td>
            </tr>`;}).join("")}
        </tbody>
      </table>
    </div>
  `;
}

/* ============================================================
   SUPPLIERS
============================================================ */
function renderSuppliers(){
  const q = state.supplierSearch.toLowerCase();
  const filtered = state.suppliers.filter(s=> s.name.toLowerCase().includes(q) || (s.phone||"").includes(q));
  const canDelete = CAN_DELETE_SUPPLIER(state.profile.role); // Admin/Manager only, per rules
  return `
    <div class="view-header">
      <div><h1>Suppliers</h1><p>${state.suppliers.length} wasambazaji waliosajiliwa</p></div>
      <button class="btn btn-dark" id="addSupplierBtn">${ICONS.plus} Ongeza msambazaji</button>
    </div>
    <div class="search-wrap">${ICONS.search}<input id="supplierSearchInput" placeholder="Tafuta jina au simu" value="${escapeHtml(state.supplierSearch)}"/></div>
    <div class="table-card">
      <table>
        <thead><tr><th>Jina</th><th>Simu</th><th>Bidhaa anazoleta</th><th></th></tr></thead>
        <tbody>
          ${filtered.length===0 ? `<tr class="empty-row"><td colspan="4">Hakuna wasambazaji bado.</td></tr>` :
          filtered.map(s=>`
            <tr>
              <td class="cell-title">${escapeHtml(s.name)}</td>
              <td class="mono" style="font-size:12px;">${escapeHtml(s.phone||"—")}</td>
              <td style="font-size:12.5px;color:var(--muted);">${escapeHtml(s.suppliesText||"—")}</td>
              <td><div class="row-actions">
                <button class="icon-btn" data-edit-supplier="${s.id}">${ICONS.edit}</button>
                ${canDelete ? `<button class="icon-btn danger" data-del-supplier="${s.id}">${ICONS.trash}</button>` : ""}
              </div></td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

/* ============================================================
   MY SHIFT (Shift Reconciliation)
============================================================ */
function currentOpenShift(){
  return state.shifts.find(s=> s.cashierUid===state.user.uid && s.status==="open");
}
function salesDuringShift(shift){
  if(!shift) return [];
  const start = new Date(shift.startTime).getTime();
  const end = shift.endTime ? new Date(shift.endTime).getTime() : Date.now();
  return state.sales.filter(s=>{
    if(s.cashierUid!==shift.cashierUid || s.status==="refunded") return false;
    const t = s.createdAt && s.createdAt.toDate ? s.createdAt.toDate().getTime() : new Date(s.date+"T"+s.time).getTime();
    return t>=start && t<=end;
  });
}
function renderMyShift(){
  const shift = currentOpenShift();
  if(!shift){
    const lastClosed = state.shifts.filter(s=>s.cashierUid===state.user.uid && s.status==="closed").sort((a,b)=> new Date(b.endTime)-new Date(a.endTime))[0];
    return `
      <div class="view-header"><div><h1>My Shift</h1><p>Anza shift yako kabla ya kuanza kuuza.</p></div></div>
      <div class="card" style="max-width:420px;">
        <h2>Anza Shift Mpya</h2>
        <div class="field"><label>Fedha taslimu mfukoni (Starting Cash)</label>
          <input type="number" min="0" id="startCashInput" placeholder="0" value="${escapeHtml(state.shiftStartCash)}"/></div>
        <button class="btn btn-amber btn-block" id="startShiftBtn">${ICONS.clock} Anza Shift</button>
      </div>
      ${lastClosed ? `
      <div class="card" style="max-width:420px;margin-top:16px;">
        <h2>Shift ya Mwisho</h2>
        <div class="subtotal-row"><span>Ilifungwa</span><span>${new Date(lastClosed.endTime).toLocaleString("en-GB")}</span></div>
        <div class="subtotal-row"><span>Tofauti (Difference)</span><span class="mono" style="color:${Math.abs(lastClosed.difference)<1?'var(--green)':'var(--red)'};">${fmt(lastClosed.difference)}</span></div>
      </div>` : ""}
    `;
  }
  const mySales = salesDuringShift(shift);
  const cashSales = mySales.filter(s=>s.paymentMethod==="Cash").reduce((sum,s)=>sum+s.total,0);
  const otherSales = mySales.reduce((sum,s)=>sum+s.total,0) - cashSales;
  const expectedCash = (shift.startCash||0) + cashSales;
  return `
    <div class="view-header"><div><h1>My Shift</h1><p>Shift ilianza ${new Date(shift.startTime).toLocaleString("en-GB")}</p></div></div>
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Mauzo (jumla)</span><div class="kpi-icon green">${ICONS.trend}</div></div><div class="kpi-value">${fmt(cashSales+otherSales)}</div><div class="kpi-sub">${mySales.length} sales</div></div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Mauzo ya Cash</span><div class="kpi-icon amber">${ICONS.receipt}</div></div><div class="kpi-value">${fmt(cashSales)}</div></div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Njia Nyingine</span><div class="kpi-icon slate">${ICONS.report}</div></div><div class="kpi-value">${fmt(otherSales)}</div></div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Cash Inayotarajiwa</span><div class="kpi-icon slate">${ICONS.clock}</div></div><div class="kpi-value">${fmt(expectedCash)}</div><div class="kpi-sub">Starting: ${fmt(shift.startCash||0)}</div></div>
    </div>
    <div class="card" style="max-width:420px;">
      <h2>Maliza Shift (Reconciliation)</h2>
      <div class="field"><label>Fedha taslimu ulizohesabu sasa (Counted Cash)</label>
        <input type="number" min="0" id="endCashInput" placeholder="0" value="${escapeHtml(state.shiftEndCash)}"/></div>
      <button class="btn btn-amber btn-block" id="endShiftBtn">${ICONS.clock} Funga Shift</button>
    </div>
  `;
}

/* ============================================================
   SHIFTS REPORT (Admin / Manager)
============================================================ */
function renderShiftsReport(){
  const closed = state.shifts.filter(s=>s.status==="closed").sort((a,b)=> new Date(b.endTime)-new Date(a.endTime));
  const open = state.shifts.filter(s=>s.status==="open");
  return `
    <div class="view-header"><div><h1>Shifts Report</h1><p>Historia ya shift za wafanyakazi na reconciliation.</p>
      <p style="font-size:11px;color:var(--muted);margin-top:4px;">${state.shifts.length} shift record${state.shifts.length===1?'':'s'} zimepakiwa.</p></div></div>
    ${open.length>0 ? `
    <div class="card" style="margin-bottom:16px;">
      <h2>Shift Zilizo Wazi Sasa</h2>
      ${open.length>1 ? `<p style="font-size:12px;color:var(--red);margin:-6px 0 10px;">Onyo: baadhi ya cashier wana shift zaidi ya moja iliyo wazi (mara nyingi husababishwa na kubonyeza "Anza Shift" mara kadhaa). Funga zile za ziada hapa chini.</p>` : ""}
      <table><tbody>
        ${open.map(s=>`<tr>
          <td class="cell-title">${escapeHtml(s.cashierName)}</td>
          <td style="font-size:12px;color:var(--muted);">Ilianza ${new Date(s.startTime).toLocaleString("en-GB")}</td>
          <td class="right mono">${fmt(s.startCash||0)} start</td>
          <td class="right"><button class="btn btn-outline" style="font-size:11px;padding:6px 10px;" data-close-stray-shift="${s.id}">Funga Shift Hii</button></td>
        </tr>`).join("")}
      </tbody></table>
    </div>` : ""}
    <div class="table-card">
      <table>
        <thead><tr><th>Cashier</th><th>Ilianza</th><th>Ilifungwa</th><th class="right">Starting Cash</th><th class="right">Expected</th><th class="right">Counted</th><th class="right">Difference</th></tr></thead>
        <tbody>
          ${closed.length===0 ? `<tr class="empty-row"><td colspan="7">Hakuna shift zilizofungwa bado.</td></tr>` :
          closed.map(s=>`
            <tr>
              <td class="cell-title">${escapeHtml(s.cashierName)}</td>
              <td style="font-size:12px;color:var(--muted);">${new Date(s.startTime).toLocaleString("en-GB")}</td>
              <td style="font-size:12px;color:var(--muted);">${new Date(s.endTime).toLocaleString("en-GB")}</td>
              <td class="right mono">${fmt(s.startCash||0)}</td>
              <td class="right mono">${fmt(s.expectedCash||0)}</td>
              <td class="right mono">${fmt(s.endCash||0)}</td>
              <td class="right mono" style="font-weight:700;color:${Math.abs(s.difference)<1?'var(--green)':'var(--red)'};">${fmt(s.difference)}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

/* ============================================================
   NEW SALE
============================================================ */
function computeSaleLines(){
  const lines = state.cart.map(c=>{
    const p = state.products.find(x=>x.id===c.productId);
    return {...c, product:p, lineTotal: p.price*c.qty};
  });
  const subtotal = lines.reduce((s,l)=>s+l.lineTotal,0);
  let discountAmount = 0;
  if(state.discountType==="percent"){
    const pct = Math.min(100, Math.max(0, Number(state.discountValue)||0));
    discountAmount = subtotal * pct/100;
  } else if(state.discountType==="amount"){
    discountAmount = Math.min(subtotal, Math.max(0, Number(state.discountValue)||0));
  }
  const total = Math.max(0, subtotal - discountAmount);
  return {lines, subtotal, discountAmount, total};
}

function updateDiscountPreview(){
  const {subtotal, discountAmount, total} = computeSaleLines();
  const totalEl = document.querySelector(".cart-total-val");
  if(totalEl) totalEl.textContent = fmt(total);
  const discountRows = document.querySelectorAll(".subtotal-row");
  if(discountRows.length >= 2){
    discountRows[0].querySelector(".mono").textContent = fmt(subtotal);
    discountRows[1].querySelector(".mono").textContent = "−" + fmt(discountAmount);
  }
  const paid = Number(state.amountPaid)||0;
  const changeEl = document.getElementById("changePreview");
  if(changeEl) changeEl.textContent = fmt(Math.max(0, paid-total));
}

function renderSale(){
  const role = state.profile.role;
  const activeShift = role === "Cashier" ? currentOpenShift() : null;
  const shiftRequired = role === "Cashier" && !activeShift;
  const q = state.productSearch.toLowerCase();
  const inStock = state.products.filter(p=>p.stock>0);
  const filtered = inStock.filter(p=>p.name.toLowerCase().includes(q));
  const {lines, subtotal, discountAmount, total} = computeSaleLines();
  const canDiscount = CAN_DISCOUNT(role);

  return `
    <div class="view-header"><div><h1>New Sale</h1><p>Add products to the cart, then complete the sale.</p></div></div>
    ${shiftRequired ? `<div class="shift-required"><strong>${ICONS.lock} Shift haijaanza</strong>Cashier lazima aanze shift yake kwenye <b>My Shift</b> kabla ya kukamilisha sale.</div>` : ""}
    ${activeShift ? `<div class="locked-note">${ICONS.clock} Shift iko wazi tangu ${escapeHtml(new Date(activeShift.startTime).toLocaleString("en-GB"))}. Mauzo mapya yatahusishwa na shift hii.</div>` : ""}
    <div class="sale-grid">
      <div>
        <div style="display:flex;gap:8px;align-items:flex-start;margin-bottom:14px;">
          <div class="search-wrap" style="margin-bottom:0;flex:1;">${ICONS.search}<input id="saleSearchInput" placeholder="Search product to add..." value="${escapeHtml(state.productSearch)}"/></div>
          <button class="btn btn-outline" id="scanSaleBarcodeBtn" style="flex-shrink:0;">${ICONS.camera} Scan</button>
        </div>
        <div class="product-grid">
          ${filtered.length===0 ? `<p class="empty-note" style="grid-column:1/-1;text-align:center;padding:24px;">No products found.</p>` :
          filtered.map(p=>`
            <button class="product-tile" data-add-product="${p.id}">
              <div class="pt-name">${escapeHtml(p.name)}</div>
              <div class="pt-stock">${p.stock} in stock</div>
              <div class="pt-price">${fmt(p.price)}</div>
            </button>`).join("")}
        </div>
      </div>
      <div class="cart-panel">
        <div class="field">
          <label>Customer</label>
          <select id="saleCustomerSelect">
            <option value="__WALK_IN__" ${state.saleCustomerId==="__WALK_IN__"?'selected':''}>🚶 Walk-in Customer — Mteja wa dukani</option>
            ${state.customers.map(c=>`<option value="${c.id}" ${c.id===state.saleCustomerId?'selected':''}>${escapeHtml(c.name)}</option>`).join("")}
          </select>
          ${state.saleCustomerId==="__WALK_IN__" ? `
          <div class="v15-walkin" style="margin-top:10px;">
            <div class="v15-section-title">Walk-in Customer <span>Si lazima awe kwenye mfumo</span></div>
            <div class="field" style="margin-bottom:8px;">
              <label>Jina la mteja (hiari)</label>
              <input id="walkInName" type="text" placeholder="Walk-in Customer" value="${escapeHtml(state.walkInName||'')}" maxlength="80"/>
            </div>
            <div class="field" style="margin-bottom:8px;">
              <label>Simu (hiari)</label>
              <input id="walkInPhone" type="tel" placeholder="07XXXXXXXX" value="${escapeHtml(state.walkInPhone||'')}" maxlength="20"/>
            </div>
            <label class="v15-check"><input id="saveWalkInCustomer" type="checkbox" ${state.saveWalkInCustomer?'checked':''}/> Hifadhi huyu kwenye Customers baada ya sale</label>
            <div class="v15-info">Receipt itatengenezwa bila kuhitaji customer wa kwenye mfumo. Ukiweka jina tupu, receipt itaonyesha <b>Walk-in Customer</b>.</div>
          </div>` : ""}
        </div>
        <div class="cart-title">${ICONS.cart} Cart</div>
        ${lines.length===0 ? `<p class="empty-note" style="text-align:center;padding:20px 0;">Cart is empty — tap a product to add it.</p>` : ""}
        <div class="cart-list">
          ${lines.map(l=>`
            <div class="cart-row">
              <div>
                <div class="cart-row-name">${escapeHtml(l.product.name)}</div>
                <div class="qty-ctrl">
                  <button class="qty-btn" data-qty-minus="${l.productId}">${ICONS.minus}</button>
                  <span class="qty-val">${l.qty}</span>
                  <button class="qty-btn" data-qty-plus="${l.productId}">${ICONS.plus}</button>
                </div>
              </div>
              <div style="text-align:right;">
                <div class="mono">${fmt(l.lineTotal)}</div>
                <button class="icon-btn danger" data-remove-cart="${l.productId}" style="margin-top:6px;">${ICONS.trash}</button>
              </div>
            </div>`).join("")}
        </div>
        ${canDiscount ? `
        <div class="discount-box">
          <label style="display:block;font-size:11.5px;font-weight:700;color:#57534e;margin-bottom:5px;">Discount (hiari)</label>
          <div class="discount-row">
            <select id="discountType">
              <option value="none" ${state.discountType==="none"?"selected":""}>Hakuna</option>
              <option value="percent" ${state.discountType==="percent"?"selected":""}>%</option>
              <option value="amount" ${state.discountType==="amount"?"selected":""}>TZS</option>
            </select>
            <input id="discountValue" type="number" min="0" placeholder="0" value="${escapeHtml(state.discountValue)}" ${state.discountType==="none"?"disabled":""}/>
          </div>
        </div>` : ""}
        ${discountAmount>0 ? `
        <div class="subtotal-row"><span>Subtotal</span><span class="mono">${fmt(subtotal)}</span></div>
        <div class="subtotal-row"><span>Discount</span><span class="mono">−${fmt(discountAmount)}</span></div>` : ""}
        <div class="cart-total-row"><span class="cart-total-label">Total</span><span class="cart-total-val mono">${fmt(total)}</span></div>
        ${(()=>{
          const creditCustomer = state.saleCustomerId!=="__WALK_IN__" ? state.customers.find(c=>c.id===state.saleCustomerId) : null;
          if(!creditCustomer) return "";
          const bal = Number(creditCustomer.balance)||0;
          const limit = Number(creditCustomer.creditLimit)||0;
          const available = Math.max(0, limit-bal);
          return `<label class="v15-check" style="margin-bottom:10px;"><input id="saleOnCreditCheck" type="checkbox" ${state.saleOnCredit?'checked':''}/> Muuzie kwa Mkopo (Deni)</label>
          ${state.saleOnCredit ? `<div class="v15-info" style="margin-bottom:10px;">Deni la sasa: <b>${fmt(bal)}</b> • Limit: <b>${fmt(limit)}</b> • Mkopo unaobaki: <b>${fmt(available)}</b></div>
          <div class="field" style="margin-bottom:10px;"><label>Tarehe ya Kulipa (hiari)</label><input id="saleDueDate" type="date" min="${todayISO()}" value="${escapeHtml(state.saleDueDate||'')}"/></div>` : ""}`;
        })()}
        <div class="field"><label>Njia ya Malipo</label>
          <select id="paymentMethodSelect">
            ${PAYMENT_METHODS.map(m=>`<option value="${m}" ${state.paymentMethod===m?"selected":""}>${m}</option>`).join("")}
          </select>
        </div>
        <div class="field"><label>${state.saleOnCredit ? "Malipo ya Awali (Deposit) — hiari" : "Kiasi Kilicholipwa"}</label>
          <input id="amountPaidInput" type="number" min="0" step="1" placeholder="${state.saleOnCredit ? '0' : fmt(total).replace('TZS ','')}" value="${escapeHtml(state.amountPaid)}"/>
          ${state.saleOnCredit
            ? `<div class="muted" style="margin-top:5px;">Deni jipya la sale hii: <strong id="creditRemainPreview">${fmt(Math.max(0,total-(Number(state.amountPaid)||0)))}</strong></div>`
            : `<div class="muted" style="margin-top:5px;">Change: <strong id="changePreview">${fmt(Math.max(0,(Number(state.amountPaid)||0)-total))}</strong></div>`}
        </div>
        <button class="btn btn-amber btn-block" id="completeSaleBtn" ${lines.length===0 || shiftRequired?'disabled':''}>${shiftRequired ? "Anza Shift Kwanza" : "Complete sale"} ${ICONS.chevron}</button>
      </div>
    </div>
  `;
}

/* ============================================================
   SALES HISTORY
============================================================ */
function renderHistory(){
  const role = state.profile.role;
  const isCashier = role === "Cashier";
  // Firestore already restricts Cashier reads to cashierUid == current UID.
  // Keep this client-side filter too as a second UI safeguard.
  const list = isCashier
    ? state.sales.filter(s=>s.cashierUid===state.user.uid)
    : state.sales;
  return `
    <div class="view-header">
      <div>
        <p class="view-eyebrow">${isCashier ? "Mauzo yako" : "Mauzo yote"}</p>
        <h1>${isCashier ? "Mauzo Yangu" : "Sales History"}</h1>
        <p>${list.length} ${isCashier ? "mauzo uliyofanya" : "sales recorded"}</p>
      </div>
    </div>
    <div class="table-card">
      <table>
        <thead><tr>
          <th>Sale #</th><th>Date</th><th>Customer</th>
          ${isCashier ? "" : "<th>Cashier</th>"}
          <th>Payment</th><th class="right">Items</th><th class="right">Total</th><th>Receipt</th>
        </tr></thead>
        <tbody>
          ${list.length===0 ? `<tr class="empty-row"><td colspan="${isCashier?7:8}">${isCashier ? "Bado hujafanya sale yoyote." : "No sales yet."}</td></tr>` :
          list.map(s=>`
            <tr class="clickable" data-view-receipt="${s.id}" title="Fungua receipt">
              <td class="mono" style="font-size:12px;color:#78716c;">#${s.number}</td>
              <td>${s.date} <span style="color:#a8a29e;font-size:11px;">${s.time||""}</span></td>
              <td class="cell-title">${escapeHtml(s.customerName||"Walk-in Customer")}</td>
              ${isCashier ? "" : `<td style="font-size:12px;color:var(--muted);">${escapeHtml(s.cashierName||"—")}</td>`}
              <td style="font-size:12px;">${escapeHtml(s.paymentMethod||"—")}</td>
              <td class="right">${s.items?.length||0}</td>
              <td class="right mono" style="font-weight:700;${s.status==='refunded'?'text-decoration:line-through;color:#a8a29e;':''}">${fmt(s.total)}${s.status==='refunded'?' <span class="pill-low" style="text-decoration:none;">Refunded</span>':''}</td>
              <td class="right">${ICONS.chevron}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}

/* ============================================================
   MONTHLY REPORT
============================================================ */
function renderReport(){
  const [y,m] = state.reportMonth.split("-").map(Number);
  const prevDate = new Date(y, m-2, 1);
  const prevKey = `${prevDate.getFullYear()}-${String(prevDate.getMonth()+1).padStart(2,"0")}`;

  const salesThis = state.sales.filter(s=>monthKey(s.date)===state.reportMonth);
  const salesPrev = state.sales.filter(s=>monthKey(s.date)===prevKey);

  const calc = (list)=>{
    let revenue=0, cogs=0;
    const productTotals = {};
    list.forEach(s=>{
      revenue += s.total;
      s.items.forEach(it=>{
        const cost = (it.cost||0)*it.qty;
        cogs += cost;
        productTotals[it.name] = (productTotals[it.name]||0) + it.lineTotal;
      });
    });
    const profit = revenue - cogs;
    return {revenue, cogs, profit, productTotals};
  };
  const cur = calc(salesThis);
  const prev = calc(salesPrev);

  const pctChange = (a,b)=> b===0 ? (a===0?0:100) : ((a-b)/b*100);
  const revChange = pctChange(cur.revenue, prev.revenue);
  const profitChange = pctChange(cur.profit, prev.profit);

  const topProducts = Object.entries(cur.productTotals).sort((a,b)=>b[1]-a[1]).slice(0,5);

  const monthOptions = [...Array(12)].map((_,i)=>{
    const d = new Date(); d.setDate(1); d.setMonth(d.getMonth()-i);
    const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
    const label = d.toLocaleDateString("en-GB",{month:"long", year:"numeric"});
    return `<option value="${key}" ${key===state.reportMonth?"selected":""}>${label}</option>`;
  }).join("");

  const trendBadge = (val)=> val>=0
    ? `<span class="kpi-sub up">▲ ${val.toFixed(1)}% vs last month</span>`
    : `<span class="kpi-sub down">▼ ${Math.abs(val).toFixed(1)}% vs last month</span>`;

  return `
    <div class="view-header">
      <div><h1>Monthly Report</h1><p>Faida na mauzo kwa mwezi uliochaguliwa.</p></div>
      <select id="reportMonthSelect" style="max-width:200px;">${monthOptions}</select>
    </div>
    <div class="kpi-grid">
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Revenue</span><div class="kpi-icon green">${ICONS.trend}</div></div><div class="kpi-value">${fmt(cur.revenue)}</div>${trendBadge(revChange)}</div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Cost of Goods</span><div class="kpi-icon slate">${ICONS.inventory}</div></div><div class="kpi-value">${fmt(cur.cogs)}</div></div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Net Profit</span><div class="kpi-icon amber">${ICONS.profit}</div></div><div class="kpi-value">${fmt(cur.profit)}</div>${trendBadge(profitChange)}</div>
      <div class="kpi-card"><div class="kpi-top"><span class="kpi-label">Sales Count</span><div class="kpi-icon slate">${ICONS.receipt}</div></div><div class="kpi-value">${salesThis.length}</div></div>
    </div>
    <div class="card">
      <h2>Bidhaa zinazouza zaidi (mwezi huu)</h2>
      ${topProducts.length===0 ? `<p class="empty-note">Hakuna mauzo mwezi huu.</p>` :
        `<table><tbody>${topProducts.map(([name,total])=>`<tr><td>${escapeHtml(name)}</td><td class="right mono">${fmt(total)}</td></tr>`).join("")}</tbody></table>`}
    </div>
  `;
}

/* ============================================================
   MANAGE USERS (Admin only)
============================================================ */
function renderUsers(){
  const role = state.profile.role;
  const canCreate = role === "Administrator";
  const canManageExisting = role === "Administrator" || role === "Manager";
  return `
    <div class="view-header">
      <div><h1>Manage Users</h1><p>${state.users.length} wafanyakazi</p></div>
      ${canCreate ? `<button class="btn btn-dark" id="addUserBtn">${ICONS.plus} Ongeza mfanyakazi</button>` : ""}
    </div>
    ${!canCreate ? `<div class="card" style="margin-bottom:14px;padding:12px 14px;"><div class="muted">Manager anaweza kusimamia wafanyakazi waliopo lakini hawezi kuongeza mfanyakazi mpya.</div></div>` : ""}
    <div class="table-card">
      <table>
        <thead><tr><th>Jina</th><th>Email</th><th>Role</th><th>Status</th><th></th></tr></thead>
        <tbody>
          ${state.users.length===0 ? `<tr class="empty-row"><td colspan="5">Hakuna wafanyakazi bado.</td></tr>` :
          state.users.map(u=>`
            <tr>
              <td class="cell-title">${escapeHtml(u.name)}${u.id===state.user.uid?' <span style="color:var(--muted);font-weight:400;">(wewe)</span>':''}</td>
              <td class="mono" style="font-size:12px;">${escapeHtml(u.email)}</td>
              <td><span class="pill-role pill-${u.role}">${escapeHtml(u.role)}</span></td>
              <td><span class="pill-role ${u.status==='active'?'pill-Cashier':'pill-disabled'}">${u.status==='active'?'Active':'Disabled'}</span></td>
              <td>${canManageExisting ? `<div class="row-actions">
                <button class="icon-btn" data-reset-pw="${u.email}" title="Tuma password reset">${ICONS.mail}</button>
                ${u.id!==state.user.uid && (role === "Administrator" || u.role !== "Administrator") ? `<button class="icon-btn" data-toggle-user="${u.id}" title="Washa/Zima">${u.status==='active'?ICONS.lock:ICONS.plus}</button>` : ""}
              </div>` : ""}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}



/* ============================================================
   STAFF CHAT — TEXT ONLY + REPLY
============================================================ */
function chatTime(ts){
  try{
    const d=ts?.toDate ? ts.toDate() : new Date(ts||Date.now());
    return d.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"});
  }catch(e){ return ""; }
}

function chatMessageTimestamp(m){
  return m.createdAt || m.clientCreatedAt || Date.now();
}


function cancelChatReply(){
  state.chatReplyTo=null;
  render();
  requestAnimationFrame(()=>document.getElementById("staffChatInput")?.focus());
}

function chooseChatReply(messageId){
  const m=state.chatMessages.find(x=>x.id===messageId);
  if(!m) return;
  state.chatReplyTo={
    messageId:m.id,
    senderUid:m.senderUid||"",
    senderName:m.senderName||"Staff",
    text:String(m.text||"").slice(0,2000)
  };
  render();
  requestAnimationFrame(()=>{
    const input=document.getElementById("staffChatInput");
    input?.focus();
  });
}


function renderChat(){
  const messages=[...state.chatMessages].sort((a,b)=>{
    const ta=a.createdAt?.toMillis ? a.createdAt.toMillis() : (a.clientCreatedAt||0);
    const tb=b.createdAt?.toMillis ? b.createdAt.toMillis() : (b.clientCreatedAt||0);
    return ta-tb;
  });
  return `
    <div class="view-header"><div><h1>Staff Chat</h1><p>Wasiliana na Administrator, Manager, Cashier na Storekeeper ndani ya mfumo.</p></div></div>
    <div class="staff-chat">
      <div class="staff-chat-head">
        <div><div class="staff-chat-title">SALES MANAGEMENT SYSTEM — Staff Chat</div><div class="staff-chat-sub">Team room • Text only</div></div>
        <div class="staff-chat-online">● TEAM</div>
      </div>
      <div class="staff-chat-messages" id="staffChatMessages">
        ${messages.length ? messages.map(m=>{
          const mine=m.senderUid===state.user?.uid;
          const name=m.senderName || "Staff";
          const role=m.senderRole || "";
          const replyPreview=m.replyToMessageId ? `
            <div class="chat-quoted">
              <div class="chat-quoted-name">${escapeHtml(m.replyToSenderName||"Staff")}</div>
              <div class="chat-quoted-text">${escapeHtml(m.replyToText||"Ujumbe")}</div>
            </div>` : "";
          return `<div class="chat-row ${mine?'mine':''}" data-chat-message="${escapeHtml(m.id)}">
            <div class="chat-avatar">${initials(name)}</div>
            <div class="chat-bubble">
              ${mine?'':`<div class="chat-sender">${escapeHtml(name)}${role?` • ${escapeHtml(role)}`:''}</div>`}
              ${replyPreview}
              <div class="chat-text">${escapeHtml(m.text||"")}</div>
              <div class="chat-meta">${chatTime(chatMessageTimestamp(m))}</div>
              <div class="chat-tools">
                <button class="chat-reply-btn" data-chat-reply="${escapeHtml(m.id)}">↩ Jibu</button>
              </div>
            </div>
          </div>`;
        }).join("") : `<div class="chat-empty">Hakuna ujumbe bado.<br/>Anza mawasiliano na timu yako.</div>`}
      </div>
      ${state.chatReplyTo ? `<div class="chat-reply-compose">
        <div><strong>Unamjibu ${escapeHtml(state.chatReplyTo.senderName)}</strong><div>${escapeHtml(state.chatReplyTo.text||"Ujumbe")}</div></div>
        <button id="cancelChatReply" title="Futa reply">×</button>
      </div>` : ""}
      <div class="chat-compose">
        <textarea id="staffChatInput" maxlength="2000" placeholder="Andika ujumbe...">${escapeHtml(state.chatText||"")}</textarea>
        <button class="chat-action send" id="staffSendBtn" title="Tuma">➤</button>
      </div>
    </div>`;
}

async function sendStaffChatText(){
  const input=document.getElementById("staffChatInput");
  const text=String(input?.value||state.chatText||"").trim();
  if(!text) return;
  if(text.length>2000){showToast("Ujumbe ni mrefu sana. Maximum ni characters 2,000.");return;}
  try{
    const now=Date.now();
    const payload={
      type:"text", text,
      senderUid:state.user.uid,
      senderName:state.profile.name||"Staff",
      senderRole:state.profile.role||"",
      clientCreatedAt:now,
      createdAt:serverTimestamp()
    };
    if(state.chatReplyTo){
      payload.replyToMessageId=state.chatReplyTo.messageId;
      payload.replyToSenderUid=state.chatReplyTo.senderUid;
      payload.replyToSenderName=state.chatReplyTo.senderName;
      payload.replyToText=state.chatReplyTo.text;
    }
    const sentRef = await addDoc(collection(db,"staffChat"),payload);

    // The Firestore listener will normally render this again from the server.
    // Keeping a temporary local copy prevents the message from visually
    // disappearing during the short serverTimestamp/snapshot round-trip.
    state.chatMessages = [
      ...state.chatMessages,
      {id:sentRef.id, ...payload}
    ].filter(m=>m.type==="text").slice(-200);

    state.chatText="";
    state.chatReplyTo=null;
    render();
    scrollStaffChatToBottom();
  }catch(e){
    console.error("sendStaffChatText:", e);
    showToast("Ujumbe haukutumwa: " + (e.message||e));
  }
}

function scrollStaffChatToBottom(){
  requestAnimationFrame(()=>{
    const el=document.getElementById("staffChatMessages");
    if(el) el.scrollTop=el.scrollHeight;
  });
}

/* ============================================================
   SUPPORT
============================================================ */
function renderSupport(){
  const waLink = `https://wa.me/${SUPPORT.whatsapp}`;
  return `
    <div class="view-header"><div><h1>Support</h1><p>Una tatizo au swali? Wasiliana nasi.</p></div></div>
    <a class="support-card" href="tel:${SUPPORT.phone}">
      <div class="support-icon" style="background:#eff6ff;color:#1d4ed8;">${ICONS.phone}</div>
      <div><div class="support-title">Piga Simu</div><div class="support-sub">${SUPPORT.phone}</div></div>
    </a>
    <a class="support-card" href="${waLink}" target="_blank" rel="noopener">
      <div class="support-icon" style="background:#f0fdf4;color:#15803d;">${ICONS.whatsapp}</div>
      <div><div class="support-title">WhatsApp</div><div class="support-sub">${SUPPORT.phone}</div></div>
    </a>
    <a class="support-card" href="mailto:${SUPPORT.email}">
      <div class="support-icon" style="background:#fffbeb;color:#b45309;">${ICONS.mail}</div>
      <div><div class="support-title">Email</div><div class="support-sub">${SUPPORT.email}</div></div>
    </a>
  `;
}

/* ============================================================
   EVENT BINDING (per view)
============================================================ */
function bindViewEvents(){
  const role = state.profile.role;

  if(state.tab==="dashboard"){
    const lowCard = document.getElementById("lowStockCard");
    if(lowCard) lowCard.addEventListener("click", ()=>{ state.tab="inventory"; render(); });
    document.querySelectorAll("[data-remind-customer]").forEach(row=> row.addEventListener("click", ()=>{
      const c = state.customers.find(x=>x.id===row.dataset.remindCustomer);
      if(c) openCustomerPaymentModal(c);
    }));
  }

  if(state.tab==="inventory"){
    const addBtn = document.getElementById("addProductBtn");
    if(addBtn) addBtn.addEventListener("click", ()=>openProductModal(null));
    document.getElementById("invSearchInput").addEventListener("input", e=>{
      state.invSearch = e.target.value; render();
      const el = document.getElementById("invSearchInput"); el.focus(); el.selectionStart = el.selectionEnd = el.value.length;
    });
    document.querySelectorAll("[data-edit-product]").forEach(b=> b.addEventListener("click", ()=> openProductModal(state.products.find(p=>p.id===b.dataset.editProduct))));
    document.querySelectorAll("[data-del-product]").forEach(b=> b.addEventListener("click", async ()=>{
      if(confirm("Remove this product?")){
        const p = state.products.find(x=>x.id===b.dataset.delProduct);
        await deleteDoc(doc(db,"products", b.dataset.delProduct));
        await logAudit("Product deleted", p ? p.name : b.dataset.delProduct);
        showToast("Product removed");
      }
    }));
  }

  if(state.tab==="customers"){
    document.getElementById("addCustomerBtn").addEventListener("click", ()=>openCustomerModal(null));
    document.querySelectorAll("[data-edit-customer]").forEach(b=> b.addEventListener("click", ()=> openCustomerModal(state.customers.find(c=>c.id===b.dataset.editCustomer))));
    document.querySelectorAll("[data-pay-customer]").forEach(b=> b.addEventListener("click", ()=> openCustomerPaymentModal(state.customers.find(c=>c.id===b.dataset.payCustomer))));
    document.querySelectorAll("[data-set-duedate]").forEach(b=> b.addEventListener("click", ()=> openSetDueDateModal(state.customers.find(c=>c.id===b.dataset.setDuedate))));
    document.querySelectorAll("[data-del-customer]").forEach(b=> b.addEventListener("click", async ()=>{
      if(confirm("Remove this customer?")){
        await deleteDoc(doc(db,"customers", b.dataset.delCustomer));
        showToast("Customer removed");
      }
    }));
  }

  if(state.tab==="suppliers"){
    document.getElementById("addSupplierBtn").addEventListener("click", ()=>openSupplierModal(null));
    document.getElementById("supplierSearchInput").addEventListener("input", e=>{
      state.supplierSearch = e.target.value; render();
      const el = document.getElementById("supplierSearchInput"); el.focus(); el.selectionStart=el.selectionEnd=el.value.length;
    });
    document.querySelectorAll("[data-edit-supplier]").forEach(b=> b.addEventListener("click", ()=> openSupplierModal(state.suppliers.find(s=>s.id===b.dataset.editSupplier))));
    document.querySelectorAll("[data-del-supplier]").forEach(b=> b.addEventListener("click", async ()=>{
      if(confirm("Ondoa msambazaji huyu?")){
        await deleteDoc(doc(db,"suppliers", b.dataset.delSupplier));
        showToast("Msambazaji ameondolewa");
      }
    }));
  }

  if(state.tab==="myshift"){
    const startBtn = document.getElementById("startShiftBtn");
    if(startBtn){
      const cashInput = document.getElementById("startCashInput");
      cashInput.addEventListener("input", e=>{ state.shiftStartCash = e.target.value; });
      startBtn.addEventListener("click", async ()=>{
        if(startBtn.disabled) return;
        startBtn.disabled = true;
        const originalLabel = startBtn.innerHTML;
        startBtn.innerHTML = `<span class="spinner"></span> Inaanza...`;
        try{
          await startShift();
        } finally {
          if(startBtn){ startBtn.disabled = false; startBtn.innerHTML = originalLabel; }
        }
      });
    }
    const endBtn = document.getElementById("endShiftBtn");
    if(endBtn){
      const cashInput = document.getElementById("endCashInput");
      cashInput.addEventListener("input", e=>{ state.shiftEndCash = e.target.value; });
      endBtn.addEventListener("click", async ()=>{
        if(endBtn.disabled) return;
        if(!confirm("Una uhakika unataka kufunga shift yako?")) return;
        endBtn.disabled = true;
        const originalLabel = endBtn.innerHTML;
        endBtn.innerHTML = `<span class="spinner"></span> Inafunga...`;
        try{
          await endShift();
        } finally {
          if(endBtn){ endBtn.disabled = false; endBtn.innerHTML = originalLabel; }
        }
      });
    }
  }

  if(state.tab==="shiftsreport"){
    document.querySelectorAll("[data-close-stray-shift]").forEach(b=>{
      b.addEventListener("click", async ()=>{
        if(!confirm("Funga shift hii ya ziada? (counted cash itawekwa sawa na expected cash)")) return;
        b.disabled = true;
        await closeStrayShift(b.dataset.closeStrayShift);
      });
    });
  }

  if(state.tab==="sale"){
    const searchInput = document.getElementById("saleSearchInput");
    searchInput.addEventListener("input", e=>{
      state.productSearch = e.target.value; render();
      const el = document.getElementById("saleSearchInput"); el.focus(); el.selectionStart=el.selectionEnd=el.value.length;
    });
    document.getElementById("saleCustomerSelect").addEventListener("change", e=>{ state.saleCustomerId = e.target.value; if(e.target.value==="__WALK_IN__") state.saleOnCredit=false; render(); });
    const creditCheck=document.getElementById("saleOnCreditCheck");
    if(creditCheck) creditCheck.addEventListener("change", e=>{
      state.saleOnCredit = e.target.checked;
      state.amountPaid = "";
      state.saleDueDate = "";
      render();
    });
    const dueDateInput=document.getElementById("saleDueDate");
    if(dueDateInput) dueDateInput.addEventListener("change", e=>{ state.saleDueDate = e.target.value; });
    const walkName=document.getElementById("walkInName"); if(walkName) walkName.addEventListener("input",e=>{state.walkInName=e.target.value;});
    const walkPhone=document.getElementById("walkInPhone"); if(walkPhone) walkPhone.addEventListener("input",e=>{state.walkInPhone=e.target.value;});
    const saveWalk=document.getElementById("saveWalkInCustomer"); if(saveWalk) saveWalk.addEventListener("change",e=>{state.saveWalkInCustomer=e.target.checked;});
    document.getElementById("scanSaleBarcodeBtn").addEventListener("click", ()=>{
      openScannerModal((value)=>{
        const p = state.products.find(x=>x.barcode && x.barcode===value);
        if(!p){ showToast("Bidhaa yenye barcode hiyo haikupatikana"); return; }
        if(p.stock<=0){ showToast(`${p.name} — stock imekwisha`); return; }
        const existing = state.cart.find(c=>c.productId===p.id);
        if(existing){ if(existing.qty<p.stock) existing.qty++; }
        else state.cart.push({productId:p.id, qty:1});
        showToast(`${p.name} imeongezwa kwenye cart`);
        render();
      });
    });
    document.querySelectorAll("[data-add-product]").forEach(b=> b.addEventListener("click", ()=>{
      const pid = b.dataset.addProduct;
      const p = state.products.find(x=>x.id===pid);
      const existing = state.cart.find(c=>c.productId===pid);
      if(existing){ if(existing.qty<p.stock) existing.qty++; }
      else state.cart.push({productId:pid, qty:1});
      render();
    }));
    document.querySelectorAll("[data-qty-plus]").forEach(b=> b.addEventListener("click", ()=>changeQty(b.dataset.qtyPlus,1)));
    document.querySelectorAll("[data-qty-minus]").forEach(b=> b.addEventListener("click", ()=>changeQty(b.dataset.qtyMinus,-1)));
    document.querySelectorAll("[data-remove-cart]").forEach(b=> b.addEventListener("click", ()=>{
      state.cart = state.cart.filter(c=>c.productId!==b.dataset.removeCart); render();
    }));
    const dType = document.getElementById("discountType");
    if(dType) dType.addEventListener("change", e=>{ state.discountType = e.target.value; if(e.target.value==="none") state.discountValue=""; render(); });
    const dVal = document.getElementById("discountValue");
    if(dVal){
      // Do NOT re-render on every keystroke. Re-rendering replaces the
      // number input and causes mobile Chrome to throw InvalidStateError
      // when selectionStart/selectionEnd are used on type="number".
      dVal.addEventListener("input", e=>{
        state.discountValue = e.target.value;
        updateDiscountPreview();
      });
      dVal.addEventListener("change", ()=>render());
      dVal.addEventListener("blur", ()=>render());
    }
    const completeBtn = document.getElementById("completeSaleBtn");
    if(completeBtn) completeBtn.addEventListener("click", completeSale);
    const pmSel = document.getElementById("paymentMethodSelect");
    if(pmSel) pmSel.addEventListener("change", e=>{
      state.paymentMethod = e.target.value;
      if(!state.saleOnCredit){
        if(e.target.value !== "Cash") state.amountPaid = String(computeSaleLines().total);
        else if(Number(state.amountPaid) <= 0) state.amountPaid = "";
      }
      render();
    });
    const paidInput = document.getElementById("amountPaidInput");
    if(paidInput) paidInput.addEventListener("input", e=>{
      state.amountPaid = e.target.value;
      const t=computeSaleLines().total;
      const ch=document.getElementById("changePreview");
      if(ch) ch.textContent=fmt(Math.max(0,(Number(state.amountPaid)||0)-t));
      const cr=document.getElementById("creditRemainPreview");
      if(cr) cr.textContent=fmt(Math.max(0,t-(Number(state.amountPaid)||0)));
    });
  }

  if(state.tab==="history"){
    document.querySelectorAll("[data-view-receipt]").forEach(row=> row.addEventListener("click", ()=>{
      const sale = state.sales.find(s=>s.id===row.dataset.viewReceipt);
      openReceiptModal(sale);
    }));
    document.querySelectorAll("[data-admin-reverse-sale]").forEach(btn=> btn.addEventListener("click", e=>{
      e.stopPropagation();
      if(!CAN_REFUND_SALE(state.profile.role)) return showToast("Administrator au Manager pekee ndiye anaweza kurefund sale.");
      const sale=state.sales.find(s=>s.id===btn.dataset.adminReverseSale);
      if(sale) refundSale(sale);
    }));
  }

  if(state.tab==="receipts"){
    const rs=document.getElementById("receiptSearchInput");
    if(rs) rs.addEventListener("input", e=>{ state.receiptSearch=e.target.value; render(); const el=document.getElementById("receiptSearchInput"); if(el){el.focus();el.selectionStart=el.selectionEnd=el.value.length;} });
    const cpb=document.getElementById("connectPrinterBtn"); if(cpb) cpb.addEventListener("click", connectPrinter);
    const dpb=document.getElementById("disconnectPrinterBtn"); if(dpb) dpb.addEventListener("click", disconnectPrinter);
    document.querySelectorAll("[data-receipt-open]").forEach(b=>b.addEventListener("click",()=>{ const sale=state.sales.find(x=>x.id===b.dataset.receiptOpen); if(sale) openReceiptModal(sale); }));
    document.querySelectorAll("[data-receipt-print]").forEach(b=>b.addEventListener("click",()=>{ const sale=state.sales.find(x=>x.id===b.dataset.receiptPrint); if(sale) openReceiptModal(sale); setTimeout(()=>document.getElementById("printReceiptBtn")?.click(),100); }));
    document.querySelectorAll("[data-receipt-wa]").forEach(b=>b.addEventListener("click",()=>{ const sale=state.sales.find(x=>x.id===b.dataset.receiptWa); if(sale) sendReceiptWhatsApp(sale); }));
    document.querySelectorAll("[data-receipt-sms]").forEach(b=>b.addEventListener("click",()=>{ const sale=state.sales.find(x=>x.id===b.dataset.receiptSms); if(sale) sendReceiptSMS(sale); }));
  }

  if(state.tab==="report"){
    document.getElementById("reportMonthSelect").addEventListener("change", e=>{ state.reportMonth = e.target.value; render(); });
  }

  if(state.tab==="expenses"){
    const b=document.getElementById("addExpenseBtn"); if(b) b.addEventListener("click", openExpenseModal);
    document.querySelectorAll("[data-del-expense]").forEach(b=>b.addEventListener("click",()=>deleteExpense(b.dataset.delExpense)));
  }
  if(state.tab==="purchases"){
    const b=document.getElementById("receiveStockBtn"); if(b) b.addEventListener("click", openPurchaseModal);
  }
  if(state.tab==="audit"){
    const cleanupBtn=document.getElementById("deleteOldAuditBtn");
    if(cleanupBtn) cleanupBtn.addEventListener("click", deleteAuditLogsOlderThan30Days);
  }

  if(state.tab==="backup"){
    const b=document.getElementById("exportSalesBtn"); if(b) b.addEventListener("click",()=>exportRowsXlsx(`sales-${todayISO()}.xlsx`, state.sales, "Sales"));
    const c=document.getElementById("exportProductsBtn"); if(c) c.addEventListener("click",()=>exportRowsXlsx(`inventory-${todayISO()}.xlsx`,state.products,"Inventory"));
    const d=document.getElementById("exportCustomersBtn"); if(d) d.addEventListener("click",()=>exportRowsXlsx(`customers-${todayISO()}.xlsx`,state.customers,"Customers"));
    const e=document.getElementById("exportExpensesBtn"); if(e) e.addEventListener("click",()=>exportRowsXlsx(`expenses-${todayISO()}.xlsx`,state.expenses,"Expenses"));
    const f=document.getElementById("exportPurchasesBtn"); if(f) f.addEventListener("click",()=>exportRowsXlsx(`purchases-${todayISO()}.xlsx`,state.purchases,"Purchases"));
    const j=document.getElementById("exportJsonBtn"); if(j) j.addEventListener("click",exportFullBackup);
    const sw=document.getElementById("shareBackupWaBtn"); if(sw) sw.addEventListener("click",downloadAndShareBackupWhatsApp);
    const ow=document.getElementById("saveOwnerWaBtn"); if(ow) ow.addEventListener("click",saveOwnerWhatsApp);
  }

  if(state.tab==="users"){
    document.getElementById("addUserBtn")?.addEventListener("click", openUserModal);
    document.querySelectorAll("[data-reset-pw]").forEach(b=> b.addEventListener("click", async ()=>{
      if(confirm(`Tuma password reset email kwa ${b.dataset.resetPw}?`)){
        try{
          await sendPasswordResetEmail(auth, b.dataset.resetPw);
          showToast("Password reset email imetumwa");
        }catch(e){ showToast("Imeshindikana: "+e.message); }
      }
    }));
    document.querySelectorAll("[data-toggle-user]").forEach(b=> b.addEventListener("click", async ()=>{
      const u = state.users.find(x=>x.id===b.dataset.toggleUser);
      const newStatus = u.status==="active" ? "disabled" : "active";
      await updateDoc(doc(db,"Users",u.id), {status:newStatus});
      await logAudit(newStatus==="active"?"User enabled":"User disabled", u.name+" ("+u.email+")");
      showToast(newStatus==="active" ? "Mtumiaji amewashwa" : "Mtumiaji amezimwa");
    }));
  }
}

async function startShift(){
  if(state.profile.role === "Cashier" && currentOpenShift()){
    showToast("Una shift ambayo bado iko wazi.");
    return;
  }
  const startCash = Number(state.shiftStartCash)||0;
  const startTime = new Date().toISOString();
  try{
    const newShift = {
      cashierUid: state.user.uid,
      cashierName: state.profile.name,
      startCash,
      startTime,
      status: "open",
      endTime: null, endCash: null, expectedCash: null, difference: null,
      createdAt: serverTimestamp()
    };
    const ref = await addDoc(collection(db,"shifts"), newShift);
    // Reflect the new shift locally right away (same reasoning as endShift):
    // don't rely solely on the onSnapshot round-trip before the next render().
    if(!state.shifts.some(s=>s.id===ref.id)) state.shifts.push({id:ref.id, ...newShift});
    await logAudit("Shift started", `${state.profile.name} — starting cash ${fmt(startCash)}`);
    state.shiftStartCash = "";
    showToast("Shift imeanza");
    render();
  }catch(e){ showToast("Hitilafu: "+e.message); }
}

/*
 * Administrator/Manager utility: force-close a stray/duplicate open
 * shift (e.g. created by a cashier double-tapping "Anza Shift" before
 * the UI confirmed the first one). Counted cash is set equal to the
 * expected cash so the difference is recorded as zero — Admin/Manager
 * can edit it later from the report if the real counted amount is known.
 */
async function closeStrayShift(shiftId){
  const shift = state.shifts.find(s=>s.id===shiftId);
  if(!shift) return;
  const mySales = salesDuringShift(shift);
  const cashSales = mySales.filter(s=>s.paymentMethod==="Cash").reduce((sum,s)=>sum+s.total,0);
  const expectedCash = (shift.startCash||0) + cashSales;
  const endTime = new Date().toISOString();
  try{
    await updateDoc(doc(db,"shifts", shiftId), {
      endTime,
      endCash: expectedCash, expectedCash, difference: 0,
      status: "closed"
    });
    shift.endTime = endTime; shift.endCash = expectedCash;
    shift.expectedCash = expectedCash; shift.difference = 0; shift.status = "closed";
    await logAudit("Shift force-closed (duplicate)", `${shift.cashierName} — shift #${shiftId}`);
    showToast("Shift ya ziada imefungwa.");
    render();
  }catch(e){ showToast("Hitilafu: "+e.message); }
}

async function endShift(){
  const shift = currentOpenShift();
  if(!shift) return;
  const endCash = Number(state.shiftEndCash)||0;
  const mySales = salesDuringShift(shift);
  const cashSales = mySales.filter(s=>s.paymentMethod==="Cash").reduce((sum,s)=>sum+s.total,0);
  const expectedCash = (shift.startCash||0) + cashSales;
  const difference = endCash - expectedCash;
  const endTime = new Date().toISOString();
  try{
    await updateDoc(doc(db,"shifts", shift.id), {
      endTime, endCash, expectedCash, difference,
      status: "closed"
    });
    // Update local state immediately instead of waiting for the onSnapshot
    // listener to round-trip — otherwise the very next render() below can
    // still see the shift as "open" and re-show the close-shift screen,
    // even though the write already succeeded (this is what produced the
    // "Shift imefungwa" toast while the Funga Shift screen stayed visible).
    const localShift = state.shifts.find(s=>s.id===shift.id);
    if(localShift){
      localShift.endTime = endTime; localShift.endCash = endCash;
      localShift.expectedCash = expectedCash; localShift.difference = difference;
      localShift.status = "closed";
    }
    await logAudit("Shift ended", `${state.profile.name} — difference ${fmt(difference)}`);
    state.shiftEndCash = "";
    showToast(`Shift imefungwa. Tofauti: ${fmt(difference)}`);
    render();
  }catch(e){ showToast("Hitilafu: "+e.message); }
}

function changeQty(productId, delta){
  const p = state.products.find(x=>x.id===productId);
  const c = state.cart.find(c=>c.productId===productId);
  if(!c) return;
  c.qty = Math.min(p.stock, Math.max(1, c.qty+delta));
  render();
}

async function completeSale(){
  // Sales may only be created by Administrator or Cashier.
  if(!CAN_SELL(state.profile.role)){
    showToast("Huna ruhusa ya kufanya mauzo. Mauzo ni ya Administrator/Cashier pekee.");
    state.tab = "dashboard";
    render();
    return;
  }
  const isCashier = state.profile.role === "Cashier";
  const activeShift = isCashier ? currentOpenShift() : null;

  // Cashier MUST have an open shift before completing a sale.
  // This is also enforced by Firestore Rules as the final security layer.
  if(isCashier && !activeShift){
    showToast("Anza shift yako kwanza kabla ya kufanya sale.");
    state.tab = "myshift";
    render();
    return;
  }

  const {lines, subtotal, discountAmount, total} = computeSaleLines();
  if(lines.length===0) return;
  const isWalkIn = state.saleCustomerId === "__WALK_IN__";
  const customer = isWalkIn ? null : (state.customers.find(c=>c.id===state.saleCustomerId) || state.customers[0]);
  if(!customer && !isWalkIn){ showToast("Chagua customer au Walk-in Customer"); return; }
  if(lines.some(l=>l.qty > l.product.stock)){ showToast("Stock haitoshi kwa bidhaa moja au zaidi."); return; }

  const onCredit = !!state.saleOnCredit;
  if(onCredit && !customer){ showToast("Mauzo ya mkopo (Deni) yanahitaji customer aliyesajiliwa, si Walk-in."); return; }

  let amountPaid, creditAmount = 0;
  if(onCredit){
    amountPaid = Math.max(0, Number(state.amountPaid)||0);
    creditAmount = Math.max(0, total-amountPaid);
    const limit = Number(customer.creditLimit)||0;
    const currentBalance = Number(customer.balance)||0;
    if(creditAmount>0){
      if(limit<=0){ showToast("Mteja huyu hana Credit Limit iliyowekwa. Weka limit kwenye Customers (Admin/Manager)."); return; }
      if(currentBalance+creditAmount > limit){ showToast(`Mkopo unaozidi limit. Mkopo unaobaki: ${fmt(Math.max(0,limit-currentBalance))}.`); return; }
    }
  } else {
    amountPaid = state.paymentMethod === "Cash" ? Number(state.amountPaid) : total;
    if(!Number.isFinite(amountPaid) || amountPaid < total){ showToast(`Kiasi kilicholipwa lazima kiwe angalau ${fmt(total)}.`); return; }
  }
  const change = Math.max(0, amountPaid-total);

  const btn = document.getElementById("completeSaleBtn");
  if(btn){ btn.disabled=true; btn.innerHTML=`<span class="spinner"></span> Inakamilisha...`; }

  try{
    const sale = {
      number:saleNumber(), date:todayISO(),
      time:new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}),
      customerId:customer ? customer.id : null,
      customerName:customer ? customer.name : (state.walkInName.trim() || "Walk-in Customer"),
      customerPhone:customer ? (customer.phone || "") : (state.walkInPhone.trim() || ""),
      cashierUid:state.user.uid, cashierName:state.profile.name,
      // Cashier sales MUST belong to the currently open shift.
      shiftId: activeShift ? activeShift.id : null,
      items:lines.map(l=>({productId:l.product.id,name:l.product.name,qty:l.qty,cost:l.product.costPrice||0,price:l.product.price,lineTotal:l.lineTotal})),
      subtotal, discountType:state.discountType, discountValue:Number(state.discountValue)||0,
      discountAmount,total,paymentMethod:onCredit?`${state.paymentMethod} + Deni`:state.paymentMethod,amountPaid,change,status:"completed",
      isCredit:onCredit, creditAmount, paymentStatus: creditAmount>0 ? (amountPaid>0?"partial":"unpaid") : "paid",
      dueDate: (onCredit && creditAmount>0 && state.saleDueDate) ? state.saleDueDate : null,
      createdAt:serverTimestamp()
    };
    const saleRef=doc(collection(db,"sales"));
    const customerRef = customer ? doc(db,"customers",customer.id) : (state.saveWalkInCustomer ? doc(collection(db,"customers")) : null);

    await runTransaction(db, async tx=>{
      let customerSnap = null;
      if(customer){
        customerSnap=await tx.get(customerRef);
        if(!customerSnap.exists()) throw new Error("Mteja hakupatikana.");
      }

      // Re-read the shift inside the transaction so the client cannot rely
      // only on an old cached state.
      let shiftSnap = null;
      if(isCashier){
        const shiftRef = doc(db,"shifts",activeShift.id);
        shiftSnap = await tx.get(shiftRef);
        if(!shiftSnap.exists()) throw new Error("Shift haikupatikana. Anza shift tena.");
        const liveShift = shiftSnap.data();
        if(liveShift.cashierUid !== state.user.uid || liveShift.status !== "open"){
          throw new Error("Shift yako haiko wazi. Anza shift tena kabla ya kuuza.");
        }
      }

      const productReads=[];
      for(const l of lines){
        const pRef=doc(db,"products",l.product.id);
        const pSnap=await tx.get(pRef);
        productReads.push({line:l,ref:pRef,snap:pSnap});
      }
      for(const x of productReads){
        if(!x.snap.exists()) throw new Error(`Bidhaa ${x.line.product.name} haipo.`);
        const stock=Number(x.snap.data().stock)||0;
        if(stock<x.line.qty) throw new Error(`Stock haitoshi: ${x.line.product.name} (imebaki ${stock}).`);
      }
      for(const x of productReads){
        const stock=Number(x.snap.data().stock)||0;
        tx.update(x.ref,{stock:stock-x.line.qty});
        tx.set(doc(collection(db,"stockMoves")),{
          productId:x.line.product.id,
          productName:x.line.product.name,
          type:"sale",
          qty:-x.line.qty,
          referenceId:saleRef.id,
          shiftId: activeShift ? activeShift.id : null,
          actorUid:state.user.uid,
          actorName:state.profile.name,
          createdAt:serverTimestamp()
        });
      }
      if(customer){
        const c=customerSnap.data();
        const newBalance = (c.balance||0)+creditAmount;
        let nextDueDate = c.nextDueDate || null;
        if(sale.dueDate && (!nextDueDate || sale.dueDate < nextDueDate)) nextDueDate = sale.dueDate;
        if(newBalance<=0) nextDueDate = null;
        tx.update(customerRef,{totalSpent:(c.totalSpent||0)+total,orders:(c.orders||0)+1,balance:newBalance,nextDueDate});
      } else if(state.saveWalkInCustomer && customerRef){
        const walkName=state.walkInName.trim() || "Walk-in Customer";
        tx.set(customerRef,{name:walkName,phone:state.walkInPhone.trim()||"",email:"",totalSpent:total,orders:1,createdAt:serverTimestamp(),source:"walk-in"});
        sale.customerId=customerRef.id;
      }
      tx.set(saleRef,sale);
    });

    await logAudit("Sale created",`#${sale.number} — ${fmt(sale.total)} (${sale.paymentMethod})`);
    state.cart=[]; state.discountType="none"; state.discountValue=""; state.amountPaid=""; state.saleOnCredit=false; state.saleDueDate="";
    state.saleCustomerId=state.customers[0]?.id || "__WALK_IN__"; state.walkInName=""; state.walkInPhone=""; state.saveWalkInCustomer=false;
    showToast(onCredit ? `Sale #${sale.number} imekamilika kwa Deni — Deni: ${fmt(creditAmount)}` : `Sale #${sale.number} completed — ${fmt(sale.total)}`);
    state.tab="history"; render(); openReceiptModal({...sale,id:saleRef.id});
  }catch(e){
    console.error(e);
    showToast("Sale haikukamilika: "+e.message);
    render();
  }
}

/* ============================================================
   MODALS
============================================================ */
function closeModal(){ const r=document.getElementById("modalRoot"); if(r) r.innerHTML=""; }

/* ============================================================
   BARCODE SCANNER
   Uses the native BarcodeDetector API where available (Chrome/Android).
   Falls back to a message telling the user to type the code manually
   on devices/browsers that don't support it (e.g. iOS Safari).
============================================================ */
let scannerStream = null;
function stopScanner(){
  if(scannerStream){ scannerStream.getTracks().forEach(t=>t.stop()); scannerStream=null; }
}
async function openScannerModal(onResult){
  const supported = "BarcodeDetector" in window;
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>Scan Barcode</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          ${supported ? `
            <div class="scanner-wrap"><video id="scannerVideo" playsinline autoplay muted></video><div class="scanner-frame"></div></div>
            <p class="scanner-hint">Elekeza kamera kwenye barcode ya bidhaa.</p>
          ` : `
            <p class="empty-note">Kifaa/kivinjari chako hakina msaada wa kusoma barcode moja kwa moja. Andika namba ya barcode chini.</p>
            <div class="field"><label>Namba ya Barcode</label><input id="manualBarcodeInput" placeholder="Andika barcode"/></div>
            <button class="btn btn-amber btn-block" id="manualBarcodeBtn">Tumia namba hii</button>
          `}
        </div>
      </div>
    </div>`;
  const close = ()=>{ stopScanner(); closeModal(); };
  document.getElementById("closeModalBtn").addEventListener("click", close);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") close(); });

  if(!supported){
    document.getElementById("manualBarcodeBtn").addEventListener("click", ()=>{
      const val = document.getElementById("manualBarcodeInput").value.trim();
      if(val){ close(); onResult(val); }
    });
    return;
  }

  try{
    scannerStream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } });
    const video = document.getElementById("scannerVideo");
    if(!video) { stopScanner(); return; } // modal was closed before camera opened
    video.srcObject = scannerStream;
    const detector = new BarcodeDetector();
    let stopped = false;
    const tick = async ()=>{
      if(stopped) return;
      try{
        const codes = await detector.detect(video);
        if(codes.length>0){
          stopped = true;
          const value = codes[0].rawValue;
          close();
          onResult(value);
          return;
        }
      }catch(e){ /* keep trying */ }
      requestAnimationFrame(tick);
    };
    requestAnimationFrame(tick);
  }catch(e){
    document.getElementById("modalRoot").querySelector(".modal-body").innerHTML =
      `<p class="empty-note">Imeshindikana kufungua kamera: ${escapeHtml(e.message)}. Hakikisha umeruhusu ufikiaji wa kamera.</p>`;
  }
}

function nextProductSku(category){
  const prefix = ((category||"").trim().slice(0,3).toUpperCase().replace(/[^A-Z0-9]/g,"")) || "PRD";
  let maxNum = 0;
  (state.products||[]).forEach(p=>{
    const m = String(p.sku||"").toUpperCase().match(new RegExp("^"+prefix+"-(\\d+)$"));
    if(m){ const n=parseInt(m[1],10); if(n>maxNum) maxNum=n; }
  });
  return prefix + "-" + String(maxNum+1).padStart(3,"0");
}

function openProductModal(product){
  const isEdit = !!product;
  const p = product || {name:"",sku:"",category:"",costPrice:"",price:"",stock:"",reorder:"5",barcode:""};
  if(!isEdit) p.sku = nextProductSku(p.category);
  const canSeeFin = CAN_SEE_FINANCE(state.profile.role);
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>${isEdit?"Edit product":"Add product"}</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          <form id="productForm">
            <div class="field"><label>Product name</label><input required id="f_name" value="${escapeHtml(p.name)}" placeholder="e.g. Cooking Oil 5L"/></div>
            <div class="field-row">
              <div class="field"><label>SKU / code ${isEdit?"":`<span style="font-weight:400;color:var(--muted);">(auto)</span>`}</label><input id="f_sku" value="${escapeHtml(p.sku)}" placeholder="GRC-002"/></div>
              <div class="field"><label>Category</label><input id="f_category" value="${escapeHtml(p.category)}" placeholder="Groceries"/></div>
            </div>
            <div class="field"><label>Barcode</label>
              <div style="display:flex;gap:8px;">
                <input id="f_barcode" value="${escapeHtml(p.barcode||"")}" placeholder="Scan au andika barcode" style="flex:1;"/>
                <button type="button" class="btn btn-outline" id="scanBarcodeBtn" style="flex-shrink:0;">${ICONS.camera}</button>
              </div>
            </div>
            <div class="field-row">
              ${canSeeFin ? `<div class="field"><label>Cost Price (TZS)</label><input type="number" min="0" id="f_cost" value="${p.costPrice}"/></div>` : `<input type="hidden" id="f_cost" value="${p.costPrice||0}"/>`}
              <div class="field"><label>Selling Price (TZS)</label><input required type="number" min="0" id="f_price" value="${p.price}"/></div>
            </div>
            <div class="field-row">
              <div class="field"><label>Stock qty</label><input required type="number" min="0" id="f_stock" value="${p.stock}"/></div>
              <div class="field"><label>Reorder at</label><input type="number" min="0" id="f_reorder" value="${p.reorder}"/></div>
            </div>
            <button type="submit" class="btn btn-amber btn-block">${isEdit?"Save changes":"Add product"}</button>
          </form>
        </div>
      </div>
    </div>`;
  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeModal(); });
  document.getElementById("scanBarcodeBtn").addEventListener("click", ()=>{
    openScannerModal((value)=>{
      openProductModal(product); // re-open modal (scanner closed it) with same product context
      document.getElementById("f_barcode").value = value;
    });
  });
  // Auto-regenerate the SKU as the category changes, unless the user has
  // manually typed into the SKU field themselves (new products only).
  if(!isEdit){
    let skuTouchedByUser = false;
    const skuInput = document.getElementById("f_sku");
    const categoryInput = document.getElementById("f_category");
    skuInput.addEventListener("input", ()=>{ skuTouchedByUser = true; });
    categoryInput.addEventListener("input", ()=>{
      if(!skuTouchedByUser) skuInput.value = nextProductSku(categoryInput.value);
    });
  }
  document.getElementById("productForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const name = document.getElementById("f_name").value.trim();
    if(!name) return;
    let sku = document.getElementById("f_sku").value.trim();
    if(!sku) sku = nextProductSku(document.getElementById("f_category").value);
    const data = {
      name, sku, category: document.getElementById("f_category").value,
      barcode: document.getElementById("f_barcode").value.trim(),
      costPrice: Number(document.getElementById("f_cost").value)||0,
      price: Number(document.getElementById("f_price").value)||0,
      stock: Number(document.getElementById("f_stock").value)||0,
      reorder: Number(document.getElementById("f_reorder").value)||0,
    };
    try{
      if(isEdit) await updateDoc(doc(db,"products",product.id), data);
      else await addDoc(collection(db,"products"), data);
      showToast(isEdit?"Product updated":"Product added");
      closeModal();
    }catch(err){ showToast("Hitilafu: "+err.message); }
  });
}

/* ============================================================
   CHANGE MY PASSWORD (self-service, any logged-in employee)
   Requires the current password (re-authenticates first) so that
   if an old password has leaked to someone else, only the real
   account owner — who still knows the current password — can
   change it out from under them.
============================================================ */
function openChangePasswordModal(){
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>Badilisha Password</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          <form id="changePwForm">
            <div class="field"><label>Password ya Zamani</label><input required id="cpwOld" type="password" autocomplete="current-password"/></div>
            <div class="field"><label>Password Mpya</label><input required id="cpwNew" type="password" minlength="6" autocomplete="new-password"/></div>
            <div class="field"><label>Rudia Password Mpya</label><input required id="cpwConfirm" type="password" minlength="6" autocomplete="new-password"/></div>
            <div class="muted" style="margin-bottom:14px;">Password mpya iwe angalau herufi/namba 6.</div>
            <button type="submit" class="btn btn-amber btn-block" id="cpwSubmitBtn">Hifadhi Password Mpya</button>
          </form>
        </div>
      </div>
    </div>`;
  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeModal(); });
  document.getElementById("changePwForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const oldPw = document.getElementById("cpwOld").value;
    const newPw = document.getElementById("cpwNew").value;
    const confirmPw = document.getElementById("cpwConfirm").value;
    if(newPw.length<6){ showToast("Password mpya lazima iwe angalau herufi/namba 6."); return; }
    if(newPw !== confirmPw){ showToast("Password mpya na urudiaji hazifanani."); return; }
    if(newPw === oldPw){ showToast("Chagua password mpya tofauti na ya zamani."); return; }
    const btn = document.getElementById("cpwSubmitBtn");
    btn.disabled = true; const originalLabel = btn.textContent; btn.innerHTML = `<span class="spinner"></span> Inahifadhi...`;
    try{
      const user = auth.currentUser;
      if(!user || !user.email) throw new Error("Session imeisha. Ingia tena.");
      const cred = EmailAuthProvider.credential(user.email, oldPw);
      await reauthenticateWithCredential(user, cred);
      await updatePassword(user, newPw);
      await logAudit("Password changed", `${state.profile.name} alibadilisha password yake.`);
      showToast("Password imebadilishwa. Tumia mpya kuingia safari ijayo.");
      closeModal();
    }catch(err){
      let msg = "Imeshindikana kubadilisha password.";
      if(err.code==="auth/wrong-password" || err.code==="auth/invalid-credential") msg = "Password ya zamani si sahihi.";
      else if(err.code==="auth/weak-password") msg = "Password mpya ni dhaifu mno.";
      else if(err.code==="auth/too-many-requests") msg = "Majaribio mengi. Subiri kidogo kisha jaribu tena.";
      else if(err.code==="auth/requires-recent-login") msg = "Ingia upya kisha jaribu tena.";
      else if(err.message) msg = err.message;
      showToast(msg);
      btn.disabled = false; btn.textContent = originalLabel;
    }
  });
}

function openSetDueDateModal(customer){
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal"><div class="modal-head"><h3>Tarehe ya Kulipa — ${escapeHtml(customer.name)}</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><form id="dueDateForm"><div class="field"><label>Tarehe</label><input id="ddDate" type="date" value="${escapeHtml(customer.nextDueDate||'')}"/></div><div class="stack"><button type="submit" class="btn btn-amber btn-block">Hifadhi</button>${customer.nextDueDate?`<button type="button" id="ddClearBtn" class="btn btn-outline btn-block">Ondoa Tarehe</button>`:""}</div></form></div></div></div>`;
  document.getElementById("closeModalBtn").onclick=closeModal;
  document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal();};
  const clearBtn=document.getElementById("ddClearBtn");
  if(clearBtn) clearBtn.onclick=async()=>{
    try{ await updateDoc(doc(db,"customers",customer.id),{nextDueDate:null}); showToast("Tarehe imeondolewa"); closeModal(); }
    catch(err){ showToast("Hitilafu: "+err.message); }
  };
  document.getElementById("dueDateForm").onsubmit=async e=>{
    e.preventDefault();
    const val=document.getElementById("ddDate").value;
    try{ await updateDoc(doc(db,"customers",customer.id),{nextDueDate: val||null}); showToast("Tarehe imehifadhiwa"); closeModal(); }
    catch(err){ showToast("Hitilafu: "+err.message); }
  };
}

function openCustomerModal(customer){
  const isEdit = !!customer;
  const c = customer || {name:"",phone:"",email:"",creditLimit:0,balance:0};
  const canSetLimit = CAN_DISCOUNT(state.profile.role); // Admin/Manager only set credit policy
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>${isEdit?"Edit customer":"Add customer"}</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          <form id="customerForm">
            <div class="field"><label>Full name</label><input required id="f_cname" value="${escapeHtml(c.name)}"/></div>
            <div class="field"><label>Phone</label><input id="f_cphone" value="${escapeHtml(c.phone)}" placeholder="07XX XXX XXX"/></div>
            <div class="field"><label>Email (optional)</label><input type="email" id="f_cemail" value="${escapeHtml(c.email)}"/></div>
            <div class="field">
              <label>Kiwango cha Mkopo / Credit Limit (TZS)</label>
              <input id="f_climit" type="number" min="0" step="1" value="${Number(c.creditLimit)||0}" ${canSetLimit?"":"disabled"}/>
              ${!canSetLimit ? `<div class="muted" style="margin-top:5px;">Administrator/Manager pekee ndio wanaweza kubadilisha limit.</div>` : ""}
            </div>
            ${isEdit ? `<div class="kpi-card" style="margin-bottom:14px;"><div class="kpi-label">Deni la sasa (Balance)</div><div class="kpi-value">${fmt(c.balance||0)}</div></div>` : ""}
            <button type="submit" class="btn btn-amber btn-block">${isEdit?"Save changes":"Add customer"}</button>
          </form>
        </div>
      </div>
    </div>`;
  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeModal(); });
  document.getElementById("customerForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const name = document.getElementById("f_cname").value.trim();
    if(!name) return;
    const data = {name, phone:document.getElementById("f_cphone").value, email:document.getElementById("f_cemail").value};
    if(canSetLimit) data.creditLimit = Math.max(0, Number(document.getElementById("f_climit").value)||0);
    try{
      if(isEdit) await updateDoc(doc(db,"customers",customer.id), data);
      else await addDoc(collection(db,"customers"), {...data, creditLimit:data.creditLimit||0, balance:0, totalSpent:0, orders:0});
      showToast(isEdit?"Customer updated":"Customer added");
      closeModal();
    }catch(err){ showToast("Hitilafu: "+err.message); }
  });
}

/* ============================================================
   CUSTOMER CREDIT (Deni) — record a payment against a customer's
   outstanding balance. Mirrors the supplier-payment pattern above,
   but the balance itself lives directly on the customer doc since
   debt can span many sales rather than a single purchase.
============================================================ */
function openCustomerPaymentModal(customer){
  const balance = Number(customer.balance)||0;
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal"><div class="modal-head"><h3>Lipa Deni — ${escapeHtml(customer.name)}</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><div class="kpi-card" style="margin-bottom:14px;"><div class="kpi-label">Deni la sasa</div><div class="kpi-value">${fmt(balance)}</div></div><form id="custPayForm"><div class="field"><label>Kiasi (TZS)</label><input required id="cpAmount" type="number" min="1" max="${balance}"/></div><div class="field"><label>Njia ya Malipo</label><select id="cpMethod">${PAYMENT_METHODS.map(m=>`<option>${m}</option>`).join("")}</select></div><button class="btn btn-amber btn-block">Save Payment</button></form></div></div></div>`;
  document.getElementById("closeModalBtn").onclick=closeModal;
  document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal();};
  document.getElementById("custPayForm").onsubmit=async e=>{
    e.preventDefault();
    const amount=Number(document.getElementById("cpAmount").value);
    if(!(amount>0) || amount>balance){ showToast("Kiasi si sahihi."); return; }
    try{
      const cRef=doc(db,"customers",customer.id);
      await runTransaction(db, async tx=>{
        const snap=await tx.get(cRef);
        if(!snap.exists()) throw new Error("Mteja hakupatikana.");
        const cur=Number(snap.data().balance)||0;
        const newBal=Math.max(0,cur-amount);
        tx.update(cRef,{balance:newBal, nextDueDate: newBal<=0 ? null : (snap.data().nextDueDate||null)});
      });
      await addDoc(collection(db,"customerPayments"),{customerId:customer.id,customerName:customer.name,amount,paymentMethod:document.getElementById("cpMethod").value,date:todayISO(),createdBy:state.user.uid,createdAt:serverTimestamp()});
      await logAudit("Customer payment", `${customer.name} — ${fmt(amount)}`);
      showToast("Malipo yamehifadhiwa");
      closeModal();
    }catch(err){ showToast("Hitilafu: "+err.message); }
  };
}

function openSupplierModal(supplier){
  const isEdit = !!supplier;
  const s = supplier || {name:"",phone:"",email:"",suppliesText:""};
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>${isEdit?"Hariri msambazaji":"Ongeza msambazaji"}</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          <form id="supplierForm">
            <div class="field"><label>Jina la kampuni / mtu</label><input required id="f_sname" value="${escapeHtml(s.name)}"/></div>
            <div class="field"><label>Simu</label><input id="f_sphone" value="${escapeHtml(s.phone)}" placeholder="07XX XXX XXX"/></div>
            <div class="field"><label>Email (hiari)</label><input type="email" id="f_semail" value="${escapeHtml(s.email)}"/></div>
            <div class="field"><label>Bidhaa anazoleta</label><textarea id="f_ssupplies" rows="2" placeholder="mf. Sukari, Mafuta, Unga">${escapeHtml(s.suppliesText)}</textarea></div>
            <button type="submit" class="btn btn-amber btn-block">${isEdit?"Hifadhi mabadiliko":"Ongeza msambazaji"}</button>
          </form>
        </div>
      </div>
    </div>`;
  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeModal(); });
  document.getElementById("supplierForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const name = document.getElementById("f_sname").value.trim();
    if(!name) return;
    const data = {
      name,
      phone: document.getElementById("f_sphone").value,
      email: document.getElementById("f_semail").value,
      suppliesText: document.getElementById("f_ssupplies").value,
    };
    try{
      if(isEdit) await updateDoc(doc(db,"suppliers",supplier.id), data);
      else await addDoc(collection(db,"suppliers"), data);
      showToast(isEdit?"Msambazaji amesasishwa":"Msambazaji ameongezwa");
      closeModal();
    }catch(err){ showToast("Hitilafu: "+err.message); }
  });
}

function openUserModal(){
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>Ongeza mfanyakazi</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          <form id="userForm">
            <div class="field"><label>Jina kamili</label><input required id="f_uname"/></div>
            <div class="field"><label>Email</label><input required type="email" id="f_uemail"/></div>
            <div class="field"><label>Password ya awali</label><input required type="text" id="f_upass" placeholder="Angalau herufi 6" minlength="6"/></div>
            <div class="field"><label>Role</label>
              <select id="f_urole">${ROLES.map(r=>`<option value="${r}">${r}</option>`).join("")}</select>
            </div>
            <button type="submit" class="btn btn-amber btn-block" id="userSubmitBtn">Ongeza mfanyakazi</button>
          </form>
        </div>
      </div>
    </div>`;
  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeModal(); });
  document.getElementById("userForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const name = document.getElementById("f_uname").value.trim();
    const email = document.getElementById("f_uemail").value.trim();
    const password = document.getElementById("f_upass").value;
    const role = document.getElementById("f_urole").value;
    const btn = document.getElementById("userSubmitBtn");
    btn.disabled = true; btn.innerHTML = `<span class="spinner" style="border-top-color:#0f172a;border-color:rgba(15,23,42,0.3);"></span> Inaongeza...`;
    try{
      const cred = await createUserWithEmailAndPassword(secondaryAuth, email, password);
      await setDoc(doc(db,"Users",cred.user.uid), {name, email, emailLower:email.toLowerCase(), role, status:"active", createdByUid:state.user.uid, createdAt: serverTimestamp()});
      await signOut(secondaryAuth);
      await logAudit("User created", `${name} (${email}) — ${role}`);
      showToast("Mfanyakazi ameongezwa");
      closeModal();
    }catch(err){
      showToast("Hitilafu: "+err.message);
      btn.disabled=false; btn.textContent="Ongeza mfanyakazi";
    }
  });
}

function normalizePhone(phone){
  let p=String(phone||"").replace(/[^0-9+]/g,"");
  if(p.startsWith("+") ) p=p.slice(1);
  if(p.startsWith("0")) p="255"+p.slice(1);
  return p;
}
function receiptInnerHTML(sale){
  const customer=state.customers.find(c=>c.id===sale.customerId);
  const customerName=sale.customerName || customer?.name || "Walk-in Customer";
  const customerPhone=sale.customerPhone || customer?.phone || "";
  return `
    <div class="receipt-head">
      <div class="rname">ST JOHN'S SHOP</div>
      <div class="rmeta">Dodoma, Tanzania</div>
      <div class="rmeta">Tel: 0755816562</div>
      <div class="rmeta">RECEIPT: ${escapeHtml(sale.number)}</div>
      <div class="rmeta">${escapeHtml(sale.date)} ${escapeHtml(sale.time)}</div>
      <div class="rmeta">Cashier: ${escapeHtml(sale.cashierName||"—")}</div>
      <div class="rmeta">Customer: ${escapeHtml(customerName)}</div>
      ${customerPhone?`<div class="rmeta">Tel: ${escapeHtml(customerPhone)}</div>`:""}
    </div>
    <div style="font-size:11px;font-weight:800;margin-bottom:5px;">ITEM             QTY     AMOUNT</div>
    <div class="receipt-items">
      ${(sale.items||[]).map(it=>`<div class="receipt-item"><span>${escapeHtml(it.name)} x${it.qty}</span><span>${fmt(it.lineTotal)}</span></div>`).join("")}
    </div>
    <div class="receipt-sub"><span>SUBTOTAL</span><span>${fmt(sale.subtotal)}</span></div>
    <div class="receipt-sub"><span>DISCOUNT</span><span>${fmt(sale.discountAmount||0)}</span></div>
    <div class="receipt-total"><span>TOTAL</span><span>${fmt(sale.total)}</span></div>
    <div class="receipt-sub"><span>Payment</span><span>${escapeHtml(sale.paymentMethod||"—")}</span></div>
    <div class="receipt-sub"><span>Paid</span><span>${fmt(sale.amountPaid ?? sale.total)}</span></div>
    <div class="receipt-sub"><span>Change</span><span>${fmt(sale.change ?? 0)}</span></div>
    ${sale.status==='refunded' ? `<div class="rmeta" style="color:var(--red);font-weight:700;text-align:center;margin-top:8px;">REFUNDED</div>` : ""}
    <div class="receipt-thanks">THANK YOU!<br/>Karibu tena.</div>
  `;
}

// Lazily loads html2canvas (only needed the first time someone sends a receipt
// as an image over WhatsApp) so it never slows down normal page load.
let _html2canvasPromise = null;
function ensureHtml2Canvas(){
  if(window.html2canvas) return Promise.resolve();
  if(_html2canvasPromise) return _html2canvasPromise;
  _html2canvasPromise = new Promise((resolve, reject)=>{
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
    s.onload = ()=>resolve();
    s.onerror = ()=>{ _html2canvasPromise = null; reject(new Error("load-failed")); };
    document.head.appendChild(s);
  });
  return _html2canvasPromise;
}

// Renders the receipt off-screen (reusing the exact same .receipt styling
// shown in the modal) and rasterizes it to a PNG blob for sharing as an image.
async function receiptImageBlob(sale){
  await ensureHtml2Canvas();
  const wrapper = document.createElement("div");
  wrapper.className = "receipt";
  wrapper.style.cssText = "position:fixed;left:-9999px;top:0;width:320px;background:#ffffff;padding:16px;";
  wrapper.innerHTML = receiptInnerHTML(sale);
  document.body.appendChild(wrapper);
  try{
    const canvas = await window.html2canvas(wrapper, { backgroundColor:"#ffffff", scale:2 });
    return await new Promise(resolve=> canvas.toBlob(resolve, "image/png"));
  }finally{
    wrapper.remove();
  }
}

function receiptText(sale, whatsapp=false){
  const customer=state.customers.find(c=>c.id===sale.customerId);
  const lines=(sale.items||[]).map(it=>`${it.name} x${it.qty} = ${fmt(it.lineTotal)}`).join("\n");
  const sep=whatsapp?"━━━━━━━━━━━━━━━━": "--------------------------------";
  const bold=whatsapp?"*":"";
  return `${bold}ST JOHN'S SHOP${bold}\nDodoma, Tanzania\nTel: 0755816562\n\n${sep}\nRECEIPT: ${sale.number}\nDate: ${sale.date} ${sale.time}\nCashier: ${sale.cashierName||"—"}${customer?.phone?`\nCustomer: ${customer.name} (${customer.phone})`:""}\n${sep}\nITEMS\n${lines}\n${sep}\nSUBTOTAL: ${fmt(sale.subtotal)}\nDISCOUNT: ${fmt(sale.discountAmount||0)}\n${bold}TOTAL: ${fmt(sale.total)}${bold}\n\nPayment: ${sale.paymentMethod||"—"}\nPaid: ${fmt(sale.amountPaid ?? sale.total)}\nChange: ${fmt(sale.change ?? 0)}\n${sep}\n\n${bold}THANK YOU!${bold}\n${bold}Karibu tena.${bold}`;
}
// Sends the receipt as an IMAGE (like a photo) via the device's native share
// sheet — so on phones the person just taps the WhatsApp icon and it goes as
// a picture, exactly like a screenshot. Falls back to a downloaded PNG (with
// WhatsApp opened for manual attach) on desktop browsers that can't share
// files, and to the old text-only link only if image generation itself fails.
// Sends the receipt as an IMAGE (like a photo). WhatsApp gives no way for a
// web app to open a specific chat AND pre-attach a file at the same time
// (that needs the paid WhatsApp Business API) — but we get very close:
// 1) copy the receipt image to the clipboard, 2) jump straight into that
// customer's chat via wa.me (works even for numbers not saved on the phone,
// no searching needed), 3) the person just long-presses and taps Paste.
async function sendReceiptWhatsApp(sale){
  const customer=state.customers.find(c=>c.id===sale.customerId);
  const phone=normalizePhone(sale.customerPhone || customer?.phone);
  if(!phone){ showToast("Mteja hana namba ya simu."); return; }
  showToast("Inatengeneza picha ya receipt...");
  let blob = null;
  try{
    blob = await receiptImageBlob(sale);
  }catch(e){ /* fall through to text-only link below */ }

  if(blob){
    let copied = false;
    try{
      if(navigator.clipboard && window.ClipboardItem){
        await navigator.clipboard.write([ new ClipboardItem({ "image/png": blob }) ]);
        copied = true;
      }
    }catch(e){ /* clipboard image copy not supported on this browser — fall back to download */ }

    if(copied){
      showToast("Picha imewekwa kwenye clipboard — fungua chat, bonyeza kwa muda kisanduku cha kuandika kisha 'Paste'.");
    }else{
      // Clipboard image copy isn't supported here — download it instead so it's
      // one tap away in the gallery/downloads while the correct chat opens.
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `Receipt-${sale.number}.png`;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(()=> URL.revokeObjectURL(url), 15000);
      showToast("Picha ya receipt imepakuliwa — iambatanishe (📎) ndani ya chat itakayofunguka.");
    }
    // Always jump straight into that customer's own chat — no searching needed,
    // works even for numbers that aren't saved as contacts.
    window.open(`https://wa.me/${phone}`,"_blank");
    return;
  }

  showToast("Imeshindikana kutengeneza picha ya receipt. Jaribu tena.");
}
function sendReceiptSMS(sale){
  const customer=state.customers.find(c=>c.id===sale.customerId);
  const phone=normalizePhone(sale.customerPhone || customer?.phone);
  if(!phone){ showToast("Mteja hana namba ya simu."); return; }
  window.location.href=`sms:${phone}?body=${encodeURIComponent(receiptText(sale,false))}`;
}
function copyReceiptText(sale){
  navigator.clipboard?.writeText(receiptText(sale,false)).then(()=>showToast("Receipt ime-copywa"),()=>showToast("Imeshindikana ku-copy"));
}

/* ============================================================
   BLUETOOTH RECEIPT PRINTER (ESC/POS thermal printer)
   Uses the Web Bluetooth API to pair with a thermal receipt
   printer and send raw ESC/POS bytes to it. The connection is
   per-browser/per-device (each cashier's phone pairs with its
   own printer) — there is nothing to store in Firestore, only
   the printer's name locally so the UI can show it's registered.
============================================================ */
let btPrinterDevice = null, btPrinterServer = null, btPrinterChar = null;
// Known "transparent UART" service/characteristic UUIDs used by most
// cheap ESC/POS thermal printers (58mm/80mm Bluetooth receipt printers).
const BT_PRINTER_SERVICES = [
  "000018f0-0000-1000-8000-00805f9b34fb", // common printer service
  "0000ff00-0000-1000-8000-00805f9b34fb",
  "e7810a71-73ae-499d-8c15-faa9aef0c3f2"  // some generic serial-over-BLE modules
];
function isPrinterConnected(){ return !!(btPrinterDevice && btPrinterDevice.gatt && btPrinterDevice.gatt.connected && btPrinterChar); }
async function findWritableCharacteristic(server){
  const services = await server.getPrimaryServices();
  for(const svc of services){
    const chars = await svc.getCharacteristics();
    const writable = chars.find(c=> c.properties.write || c.properties.writeWithoutResponse);
    if(writable) return writable;
  }
  return null;
}
async function connectPrinter(){
  if(!navigator.bluetooth){ showToast("Bluetooth haitumiki kwenye browser hii. Tumia Chrome kwenye Android."); return; }
  try{
    btPrinterDevice = await navigator.bluetooth.requestDevice({
      acceptAllDevices:true,
      optionalServices: BT_PRINTER_SERVICES
    });
    btPrinterServer = await btPrinterDevice.gatt.connect();
    btPrinterChar = await findWritableCharacteristic(btPrinterServer);
    if(!btPrinterChar){ showToast("Printa imeunganishwa lakini haikuonyesha channel ya kuchapisha."); return; }
    state.printerName = btPrinterDevice.name || "Printer";
    localStorage.setItem("duka_printer_name", state.printerName);
    btPrinterDevice.addEventListener("gattserverdisconnected", ()=>{ btPrinterChar=null; render(); });
    showToast(`Printa "${state.printerName}" imesajiliwa`);
    render();
  }catch(err){
    if(err.name!=="NotFoundError") showToast("Imeshindikana kuunganisha printa: "+err.message);
  }
}
function disconnectPrinter(){
  try{ btPrinterDevice?.gatt?.disconnect(); }catch(e){}
  btPrinterDevice=null; btPrinterServer=null; btPrinterChar=null;
  state.printerName=""; localStorage.removeItem("duka_printer_name");
  showToast("Printa imeondolewa");
  render();
}
// Encodes plain receipt text into ESC/POS bytes: init printer, plain text
// (thermal printers are single-byte/ASCII-friendly), then feed + partial cut.
function escposEncodeReceipt(text){
  const ESC=0x1B, GS=0x1D;
  const head=[ESC,0x40]; // ESC @ = initialize
  const bodyBytes=[...text].map(ch=>{
    const code=ch.codePointAt(0);
    return code<256 ? code : 0x3F; // fallback '?' for unsupported glyphs (emoji etc.)
  });
  const tail=[0x0A,0x0A,0x0A,GS,0x56,0x42,0x00]; // feed 3 lines, GS V B 0 = partial cut
  return new Uint8Array([...head, ...bodyBytes, ...tail]);
}
async function writeToPrinter(bytes){
  const CHUNK=180; // stay under typical BLE MTU
  for(let i=0;i<bytes.length;i+=CHUNK){
    const chunk=bytes.slice(i,i+CHUNK);
    if(btPrinterChar.properties.writeWithoutResponse) await btPrinterChar.writeValueWithoutResponse(chunk);
    else await btPrinterChar.writeValue(chunk);
  }
}
async function printSaleToBluetoothPrinter(sale){
  if(!isPrinterConnected()){
    showToast("Hakuna printa iliyosajiliwa. Bofya 'Sajili Printa' kwanza.");
    return;
  }
  try{
    const bytes=escposEncodeReceipt(receiptText(sale,false));
    await writeToPrinter(bytes);
    showToast("Receipt imetumwa kwa printa");
  }catch(err){
    showToast("Imeshindikana kuchapisha: "+err.message);
  }
}
function printerStatusHtml(){
  const connected=isPrinterConnected();
  return `<div class="v15-info" style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:14px;">
    <span>${ICONS.print} Printa: <b>${connected ? escapeHtml(state.printerName||"Imeunganishwa") : (state.printerName ? "Imekatika — unganisha tena" : "Hakuna iliyosajiliwa")}</b></span>
    <span class="stack" style="display:flex;gap:6px;">
      <button class="btn btn-outline" id="connectPrinterBtn">${connected?"Badilisha":"Sajili Printa"}</button>
      ${connected ? `<button class="btn btn-outline" id="disconnectPrinterBtn">Ondoa</button>` : ""}
    </span>
  </div>`;
}
function receiptActionButtons(sale){
  return `<div class="stack" style="margin-top:14px;">
    <button class="btn btn-outline" id="printReceiptBtn">${ICONS.print} Print / PDF</button>
    <button class="btn btn-outline" id="printBtReceiptBtn">${ICONS.print} Printa ya Risiti (Bluetooth)</button>
    <button class="btn btn-outline" id="whatsappReceiptBtn">${ICONS.whatsapp} WhatsApp</button>
    <button class="btn btn-outline" id="smsReceiptBtn">${ICONS.phone} SMS</button>
    <button class="btn btn-outline" id="copyReceiptBtn">${ICONS.receipt} Copy Text</button>
  </div>`;
}
function openReceiptModal(sale){
  if(!sale) return;
  const customer=state.customers.find(c=>c.id===sale.customerId);
  document.getElementById("modalRoot").innerHTML = `
    <div class="modal-overlay" id="overlay">
      <div class="modal">
        <div class="modal-head"><h3>Receipt #${escapeHtml(sale.number)}</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div>
        <div class="modal-body">
          <div class="receipt" id="printArea">
            ${receiptInnerHTML(sale)}
          </div>
          ${receiptActionButtons(sale)}
          ${CAN_REFUND_SALE(state.profile.role) && sale.status!=='refunded' ? `
          <div class="stack" style="margin-top:10px;">
            <button class="btn btn-outline btn-block" id="refundSaleBtn" style="color:#b45309;border-color:#fde68a;">${ICONS.trash} Refund Sale</button>
          </div>` : ""}
        </div>
      </div>
    </div>`;
  document.getElementById("closeModalBtn").addEventListener("click", closeModal);
  document.getElementById("overlay").addEventListener("click", e=>{ if(e.target.id==="overlay") closeModal(); });
  document.getElementById("printReceiptBtn").addEventListener("click", ()=> window.print());
  document.getElementById("printBtReceiptBtn").addEventListener("click", ()=> printSaleToBluetoothPrinter(sale));
  document.getElementById("whatsappReceiptBtn").addEventListener("click", ()=>sendReceiptWhatsApp(sale));
  document.getElementById("smsReceiptBtn").addEventListener("click", ()=>sendReceiptSMS(sale));
  document.getElementById("copyReceiptBtn").addEventListener("click", ()=>copyReceiptText(sale));
  const refundBtn = document.getElementById("refundSaleBtn");
  if(refundBtn) refundBtn.addEventListener("click", ()=> refundSale(sale));
}

function renderReceiptCenter(){
  const role=state.profile.role;
  let list=role==="Cashier" ? state.sales.filter(s=>s.cashierUid===state.user.uid) : state.sales;
  const q=(state.receiptSearch||"").trim().toLowerCase();
  if(q) list=list.filter(s=>String(s.number||"").toLowerCase().includes(q)||String(s.customerName||"").toLowerCase().includes(q)||String(s.cashierName||"").toLowerCase().includes(q));
  list=[...list].sort((a,b)=>`${b.date} ${b.time}`.localeCompare(`${a.date} ${a.time}`));
  return `<div class="view-header"><div><h1>Receipt Center</h1><p>Tengeneza, print, SMS au WhatsApp receipt za mauzo.</p></div></div>
    ${printerStatusHtml()}
    <div class="table-card" style="padding:16px;">
      <div class="search-wrap" style="margin-bottom:14px;">${ICONS.search}<input id="receiptSearchInput" placeholder="Search receipt number, customer or cashier..." value="${escapeHtml(state.receiptSearch||"")}"/></div>
      <div class="muted" style="margin-bottom:10px;">${list.length} receipt(s)</div>
      <div style="overflow:auto;"><table><thead><tr><th>Receipt</th><th>Date</th><th>Customer</th><th>Payment</th><th class="right">Total</th><th>Actions</th></tr></thead><tbody>
      ${list.length===0?`<tr class="empty-row"><td colspan="6">Hakuna receipt zilizopatikana.</td></tr>`:list.map(s=>`<tr>
        <td class="mono" style="font-weight:700;">${escapeHtml(s.number)}</td><td>${escapeHtml(s.date)} ${escapeHtml(s.time)}</td><td>${escapeHtml(s.customerName||"—")}</td><td>${escapeHtml(s.paymentMethod||"—")}</td><td class="right mono">${fmt(s.total)}</td>
        <td><div class="stack"><button class="btn btn-outline" data-receipt-open="${s.id}">View</button><button class="btn btn-outline" data-receipt-print="${s.id}">${ICONS.print}</button><button class="btn btn-outline" data-receipt-wa="${s.id}">${ICONS.whatsapp}</button><button class="btn btn-outline" data-receipt-sms="${s.id}">${ICONS.phone}</button></div></td>
      </tr>`).join("")}
      </tbody></table></div>
    </div>`;
}

function auditLogTime(a){
  if(a.createdAt?.toDate) return a.createdAt.toDate().getTime();
  if(a.date){
    const t=String(a.time||"00:00");
    const parsed=new Date(`${a.date}T${t}`);
    if(!Number.isNaN(parsed.getTime())) return parsed.getTime();
  }
  return NaN;
}

async function deleteAuditLogsOlderThan30Days(){
  if(!CAN_MANAGE_USERS(state.profile.role)){ showToast("Administrator pekee ndiye anaweza kufanya audit cleanup."); return; }
  const cutoff=Date.now()-(30*24*60*60*1000);
  const oldLogs=state.auditLog.filter(a=>{ const t=auditLogTime(a); return Number.isFinite(t) && t<cutoff; });
  if(oldLogs.length===0){ showToast("Hakuna audit logs za zaidi ya siku 30."); return; }
  const ok=confirm(`Futa audit logs ${oldLogs.length} za zaidi ya siku 30?\n\nLogs hizi zitaondolewa kabisa na hatua haiwezi kutenduliwa.`);
  if(!ok) return;
  try{
    // Firestore batches support up to 500 writes; use 450 for a safe margin.
    for(let i=0;i<oldLogs.length;i+=450){
      const batch=writeBatch(db);
      oldLogs.slice(i,i+450).forEach(a=>batch.delete(doc(db,"auditLog",a.id)));
      await batch.commit();
    }
    showToast(`${oldLogs.length} audit logs zimefutwa.`);
    render();
  }catch(e){ showToast("Audit cleanup haikukamilika: "+e.message); }
}

async function refundSale(sale){
  if(!CAN_REFUND_SALE(state.profile.role)){
    showToast("Administrator au Manager pekee ndiye anaweza kurefund sale.");
    return;
  }
  if(!confirm(`Una uhakika unataka kurefund Sale #${sale.number} — ${fmt(sale.total)}?`)) return;
  try{
    const saleRef=doc(db,"sales",sale.id);
    const customerRef=doc(db,"customers",sale.customerId);
    await runTransaction(db,async tx=>{
      const saleSnap=await tx.get(saleRef);
      if(!saleSnap.exists()) throw new Error("Sale haipo.");
      const liveSale=saleSnap.data();
      if(liveSale.status==="refunded") throw new Error("Sale tayari ime-refund.");
      const cSnap=await tx.get(customerRef);
      const productReads=[];
      for(const it of (liveSale.items||[])){
        if(!it.productId) throw new Error(`Sale ya zamani haina productId: ${it.name}`);
        const pRef=doc(db,"products",it.productId);
        const pSnap=await tx.get(pRef);
        productReads.push({it,pRef,pSnap});
      }
      if(cSnap.exists()){
        const c=cSnap.data();
        const creditToReverse = liveSale.isCredit ? Math.max(0,Number(liveSale.creditAmount)||0) : 0;
        tx.update(customerRef,{totalSpent:Math.max(0,(c.totalSpent||0)-liveSale.total),orders:Math.max(0,(c.orders||0)-1),balance:Math.max(0,(c.balance||0)-creditToReverse)});
      }
      for(const x of productReads){
        if(x.pSnap.exists()){
          const p=x.pSnap.data();
          tx.update(x.pRef,{stock:(p.stock||0)+x.it.qty});
          tx.set(doc(collection(db,"stockMoves")),{productId:x.it.productId,productName:x.it.name,type:"refund",qty:x.it.qty,referenceId:sale.id,shiftId:liveSale.shiftId || null,actorUid:state.user.uid,actorName:state.profile.name,createdAt:serverTimestamp()});
        }
      }
      tx.update(saleRef,{status:"refunded",refundedAt:serverTimestamp(),refundedBy:state.user.uid});
    });
    await logAudit("Sale refunded",`#${sale.number} — ${fmt(sale.total)}`);
    showToast(`Sale #${sale.number} ime-refund`);
    closeModal();
  }catch(e){ showToast("Refund haikukamilika: "+e.message); }
}

function renderExpenses(){
  const total=state.expenses.reduce((a,e)=>a+(Number(e.amount)||0),0);
  const canDelete = CAN_DELETE_EXPENSE(state.profile.role); // Administrator only, per rules
  return `<div class="view-header"><div><h1>Expenses</h1><p>Gharama za biashara na net profit.</p></div><button class="btn btn-dark" id="addExpenseBtn">${ICONS.plus} Ongeza expense</button></div>
    <div class="kpi-grid"><div class="kpi-card"><div class="kpi-label">Total Expenses</div><div class="kpi-value">${fmt(total)}</div></div></div>
    <div class="table-card"><table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th class="right">Amount</th><th></th></tr></thead><tbody>
    ${state.expenses.length?state.expenses.map(e=>`<tr><td>${escapeHtml(e.date||"")}</td><td>${escapeHtml(e.category||"")}</td><td>${escapeHtml(e.description||"")}</td><td class="right mono">${fmt(e.amount)}</td><td>${canDelete?`<button class="icon-btn danger" data-del-expense="${e.id}">${ICONS.trash}</button>`:""}</td></tr>`).join(""):`<tr class="empty-row"><td colspan="5">Hakuna expenses.</td></tr>`}
    </tbody></table></div>`;
}
function openExpenseModal(){
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal"><div class="modal-head"><h3>Ongeza Expense</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><form id="expenseForm">
  <div class="field"><label>Category</label><select id="exCat">${EXPENSE_CATEGORIES.map(x=>`<option>${x}</option>`).join("")}</select></div>
  <div class="field"><label>Description</label><input required id="exDesc" placeholder="Mf. Umeme wa mwezi"/></div>
  <div class="field"><label>Amount (TZS)</label><input required min="1" type="number" id="exAmount"/></div>
  <button class="btn btn-amber btn-block">Hifadhi Expense</button></form></div></div></div>`;
  document.getElementById("closeModalBtn").onclick=closeModal; document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal()};
  document.getElementById("expenseForm").onsubmit=async e=>{e.preventDefault();
    const amount=Number(document.getElementById("exAmount").value); if(amount<=0){showToast("Weka amount sahihi");return;}
    try{await addDoc(collection(db,"expenses"),{category:document.getElementById("exCat").value,description:document.getElementById("exDesc").value.trim(),amount,date:todayISO(),createdBy:state.user.uid,createdAt:serverTimestamp()});await logAudit("Expense created",`${fmt(amount)} — ${document.getElementById("exCat").value}`);closeModal();showToast("Expense imehifadhiwa");}catch(err){showToast("Hitilafu: "+err.message)}
  };
}
async function deleteExpense(id){
  if(!confirm("Futa expense hii?"))return;
  try{await deleteDoc(doc(db,"expenses",id));await logAudit("Expense deleted",id);showToast("Expense imefutwa");}catch(e){showToast("Hitilafu: "+e.message)}
}
function renderPurchases(){
  const total=state.purchases.reduce((a,p)=>a+(Number(p.total)||0),0);
  return `<div class="view-header"><div><h1>Purchases</h1><p>Receive stock kutoka kwa suppliers.</p></div><button class="btn btn-dark" id="receiveStockBtn">${ICONS.plus} Receive Stock</button></div>
  <div class="kpi-grid"><div class="kpi-card"><div class="kpi-label">Purchase Value</div><div class="kpi-value">${fmt(total)}</div></div><div class="kpi-card"><div class="kpi-label">Purchases</div><div class="kpi-value">${state.purchases.length}</div></div></div>
  <div class="table-card"><table><thead><tr><th>Date</th><th>Supplier</th><th>Items</th><th class="right">Total</th></tr></thead><tbody>
  ${state.purchases.length?state.purchases.map(p=>`<tr><td>${escapeHtml(p.date||"")}</td><td>${escapeHtml(p.supplierName||"—")}</td><td>${(p.items||[]).map(i=>escapeHtml(i.name)+" × "+i.qty).join(", ")}</td><td class="right mono">${fmt(p.total)}</td></tr>`).join(""):`<tr class="empty-row"><td colspan="4">Hakuna purchases.</td></tr>`}
  </tbody></table></div>`;
}
function openPurchaseModal(){
  const options=state.products.map(p=>`<option value="${p.id}">${escapeHtml(p.name)} — ${p.stock} stock</option>`).join("");
  const suppliers=state.suppliers.map(s=>`<option value="${s.id}">${escapeHtml(s.name)}</option>`).join("");
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal"><div class="modal-head"><h3>Receive Stock</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><form id="purchaseForm">
  <div class="field"><label>Supplier</label><select id="puSupplier">${suppliers||"<option value=''>Hakuna supplier</option>"}</select></div>
  <div class="field"><label>Product</label><select id="puProduct">${options}</select></div>
  <div class="field-row"><div class="field"><label>Quantity</label><input type="number" min="1" required id="puQty"/></div><div class="field"><label>Unit Cost (TZS)</label><input type="number" min="0" required id="puCost"/></div></div>
  <button class="btn btn-amber btn-block">Receive & Increase Stock</button></form></div></div></div>`;
  document.getElementById("closeModalBtn").onclick=closeModal; document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal()};
  document.getElementById("purchaseForm").onsubmit=async e=>{e.preventDefault();
    const productId=document.getElementById("puProduct").value, qty=Number(document.getElementById("puQty").value), unitCost=Number(document.getElementById("puCost").value);
    const supplierId=document.getElementById("puSupplier").value, supplier=state.suppliers.find(x=>x.id===supplierId);
    if(!productId||qty<=0||unitCost<0){showToast("Weka taarifa sahihi");return}
    try{
      const pRef=doc(db,"products",productId), purchaseRef=doc(collection(db,"purchases")), moveRef=doc(collection(db,"stockMoves"));
      await runTransaction(db,async tx=>{
        const pSnap=await tx.get(pRef); if(!pSnap.exists())throw new Error("Product haipo");
        const p=pSnap.data(); const item={productId,name:p.name,qty,unitCost,lineTotal:qty*unitCost};
        tx.update(pRef,{stock:(p.stock||0)+qty,costPrice:unitCost});
        tx.set(purchaseRef,{number:`PUR-${todayISO().replaceAll("-","")}-${String(Date.now()).slice(-6)}`,date:todayISO(),supplierId,supplierName:supplier?.name||"—",items:[item],total:qty*unitCost,createdBy:state.user.uid,createdAt:serverTimestamp()});
        tx.set(moveRef,{productId,productName:p.name,type:"purchase",qty,referenceId:purchaseRef.id,actorUid:state.user.uid,actorName:state.profile.name,createdAt:serverTimestamp()});
      });
      await logAudit("Stock received",`${qty} × ${state.products.find(x=>x.id===productId)?.name||productId}`);
      closeModal();showToast("Stock imeongezwa"); 
    }catch(err){showToast("Hitilafu: "+err.message)}
  };
}
function renderBackup(){
  return `<div class="view-header"><div><h1>Backup & Export</h1><p>Pakua data ya biashara yako kabla ya maintenance au migration.</p></div></div>
  <div class="card-grid">
    <div class="card"><h2>CSV Exports</h2><p class="muted">Export data moja moja kwa Excel.</p><div class="stack">
    <button class="btn btn-outline" id="exportSalesBtn">Sales CSV</button><button class="btn btn-outline" id="exportProductsBtn">Inventory CSV</button><button class="btn btn-outline" id="exportCustomersBtn">Customers CSV</button><button class="btn btn-outline" id="exportExpensesBtn">Expenses CSV</button><button class="btn btn-outline" id="exportPurchasesBtn">Purchases CSV</button></div></div>
    <div class="card"><h2>Full JSON Backup</h2><p class="muted">Backup ya data inayopatikana kwenye account hii.</p>
      <div class="stack"><button class="btn btn-dark" id="exportJsonBtn">Download Full Backup</button>
      <button class="btn btn-outline" id="shareBackupWaBtn">${ICONS.whatsapp} Download &amp; Share WhatsApp</button></div>
    </div>
    <div class="card"><h2>WhatsApp Namba ya Backup</h2><p class="muted">Namba itakayotumika ukibofya "Download &amp; Share WhatsApp".</p>
      <div class="field"><input id="ownerWaInput" type="tel" placeholder="07XX XXX XXX" value="${escapeHtml(state.ownerWhatsApp||"")}"/></div>
      <button class="btn btn-outline" id="saveOwnerWaBtn">Hifadhi Namba</button>
    </div>
  </div>`;
}
function exportFullBackup(){
  const backup={version:"V1.4",exportedAt:new Date().toISOString(),products:state.products,customers:state.customers,sales:state.sales,expenses:state.expenses,purchases:state.purchases,stockMoves:state.stockMoves,suppliers:state.suppliers,shifts:state.shifts};
  downloadText(`duka-manager-backup-${todayISO()}.json`,JSON.stringify(backup,null,2),"application/json");
  logAudit("Backup exported","Full JSON backup");
}
async function saveOwnerWhatsApp(){
  const raw=document.getElementById("ownerWaInput").value.trim();
  const phone=normalizePhone(raw);
  if(!phone){ showToast("Weka namba sahihi ya simu."); return; }
  try{
    await setDoc(doc(db,"settings","business"),{ownerWhatsApp:phone,updatedBy:state.user.uid,updatedAt:serverTimestamp()},{merge:true});
    state.ownerWhatsApp=phone; localStorage.setItem("duka_owner_wa",phone);
    showToast("Namba ya WhatsApp imehifadhiwa");
  }catch(e){ showToast("Hitilafu: "+e.message); }
}
// Backup faili la JSON linapakuliwa KISHA "share sheet" ya simu inafunguliwa
// (kama ilivyo kwa receipt) ili mtumiaji abofye WhatsApp na afungue chat
// ya namba aliyohifadhi bila kutafuta. Kwa simu zinazounga mkono
// Web Share API na files (Android Chrome), faili halisi linaambatanishwa
// moja kwa moja kwenye share sheet — hatua moja tu badala ya kubandika.
async function downloadAndShareBackupWhatsApp(){
  if(!state.ownerWhatsApp){ showToast("Weka na uhifadhi namba ya WhatsApp kwanza."); return; }
  const backup={version:"V1.4",exportedAt:new Date().toISOString(),products:state.products,customers:state.customers,sales:state.sales,expenses:state.expenses,purchases:state.purchases,stockMoves:state.stockMoves,suppliers:state.suppliers,shifts:state.shifts};
  const jsonStr=JSON.stringify(backup,null,2);
  const filename=`duka-manager-backup-${todayISO()}.json`;
  const blob=new Blob([jsonStr],{type:"application/json"});
  const file=new File([blob],filename,{type:"application/json"});

  if(navigator.canShare && navigator.canShare({files:[file]})){
    try{
      await navigator.share({files:[file], title:"SALES MANAGEMENT SYSTEM Backup", text:`Backup ya ${todayISO()}`});
      logAudit("Backup shared","Full JSON backup via share sheet");
      return;
    }catch(e){ /* mtumiaji ame-cancel share, au haifanyi kazi — endelea chini */ }
  }
  // Fallback: pakua faili, kisha fungua WhatsApp chat ili ushike faili mwenyewe (📎).
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a"); a.href=url; a.download=filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 15000);
  showToast("Backup imepakuliwa — iambatanishe (📎) kwenye chat itakayofunguka.");
  window.open(`https://wa.me/${state.ownerWhatsApp}`,"_blank");
  logAudit("Backup exported","Full JSON backup (download + manual WhatsApp attach)");
}

function renderAuditLog(){
  return `
    <div class="view-header">
      <div><h1>Audit Log</h1><p>Historia ya matendo muhimu kwenye mfumo</p></div>
      ${CAN_MANAGE_USERS(state.profile.role) ? `<button class="btn btn-outline" id="deleteOldAuditBtn" style="color:var(--red);border-color:#fecaca;">${ICONS.trash} Futa Logs za Siku 30+</button>` : ""}
    </div>
    <div class="card" style="margin-bottom:14px;padding:12px 14px;">
      <strong>Maintenance ya Audit Log</strong>
      <div class="muted" style="margin-top:4px;">Kitufe hiki kinafuta audit logs zenye umri wa zaidi ya siku 30. Administrator pekee ndiye anaweza kufanya cleanup.</div>
    </div>
    <div class="table-card">
      <table>
        <thead><tr><th>Tarehe</th><th>Aliyefanya</th><th>Kitendo</th><th>Maelezo</th></tr></thead>
        <tbody>
          ${state.auditLog.length===0 ? `<tr class="empty-row"><td colspan="4">Hakuna rekodi bado.</td></tr>` :
          state.auditLog.map(a=>`
            <tr>
              <td style="font-size:12px;color:var(--muted);">${a.date} <span style="color:#a8a29e;">${a.time}</span></td>
              <td class="cell-title">${escapeHtml(a.actorName)}</td>
              <td>${escapeHtml(a.action)}</td>
              <td style="font-size:12.5px;color:var(--muted);">${escapeHtml(a.details||"")}</td>
            </tr>`).join("")}
        </tbody>
      </table>
    </div>
  `;
}


/* ============================================================
   V1.5 BUSINESS PROFESSIONAL ENHANCEMENTS
   17 requested improvements: supplier credit, temporary customers,
   role dashboards, notifications, targets, valuation, dead stock,
   supplier statements, profit analysis, stronger audit/export.
============================================================ */

// Admin has no shift UI. Cashier owns a shift; Manager may review shifts.
const V15_CAN_TRACK_SHIFT = r => r==="Manager" || r==="Cashier";
const V15_CAN_SEE_SHIFTS_REPORT = r => r==="Administrator" || r==="Manager";
const V15_CAN_MANAGE_SUPPLIER_PAYMENTS = r => r==="Administrator" || r==="Manager";

state.tempCustomerName = state.tempCustomerName || "";
state.tempCustomerPhone = state.tempCustomerPhone || "";
state.saveWalkInCustomer = !!state.saveWalkInCustomer;
state.supplierPayments = state.supplierPayments || [];
state.businessTarget = Number(localStorage.getItem("duka_sales_target") || 0);
state.ownerWhatsApp = localStorage.getItem("duka_owner_wa") || "";

function v15Role(){ return state.profile?.role || ""; }
function v15DateValue(v){
  if(!v) return null;
  if(v.toDate) return v.toDate();
  const d=new Date(v);
  return Number.isNaN(d.getTime()) ? null : d;
}
function v15DaysUntil(dateStr){
  const d=new Date(dateStr+"T23:59:59");
  return Math.ceil((d.getTime()-Date.now())/86400000);
}
function v15PurchaseBalance(p){
  const total=Number(p.total)||0;
  const paid=Number(p.paidAmount)||0;
  const payments=state.supplierPayments.filter(x=>x.purchaseId===p.id).reduce((a,x)=>a+(Number(x.amount)||0),0);
  return Math.max(0,total-paid-payments);
}
function v15PurchasePaid(p){
  const paid=Number(p.paidAmount)||0;
  const payments=state.supplierPayments.filter(x=>x.purchaseId===p.id).reduce((a,x)=>a+(Number(x.amount)||0),0);
  return Math.min(Number(p.total)||0,paid+payments);
}
function v15SupplierDebt(supplierId){
  return state.purchases.filter(p=>p.supplierId===supplierId).reduce((a,p)=>a+v15PurchaseBalance(p),0);
}
function v15OverduePurchases(){
  return state.purchases.filter(p=>v15PurchaseBalance(p)>0 && p.dueDate && v15DaysUntil(p.dueDate)<0);
}
function v15DueSoonPurchases(){
  return state.purchases.filter(p=>v15PurchaseBalance(p)>0 && p.dueDate && v15DaysUntil(p.dueDate)>=0 && v15DaysUntil(p.dueDate)<=3);
}
function v15StockValue(){
  return state.products.reduce((a,p)=>a+(Number(p.stock)||0)*(Number(p.costPrice)||0),0);
}
function v15PotentialSalesValue(){
  return state.products.reduce((a,p)=>a+(Number(p.stock)||0)*(Number(p.price)||0),0);
}
function v15ProductSalesMap(){
  const map={};
  state.sales.filter(s=>s.status!=="refunded").forEach(s=>(s.items||[]).forEach(i=>{
    map[i.productId] = (map[i.productId]||0) + (Number(i.qty)||0);
  }));
  return map;
}
function v15LastSaleMap(){
  const map={};
  state.sales.filter(s=>s.status!=="refunded").forEach(s=>(s.items||[]).forEach(i=>{
    const old=map[i.productId];
    const d=new Date(`${s.date}T${s.time||"00:00"}`);
    if(!old || d>old) map[i.productId]=d;
  }));
  return map;
}
function v15DeadStock(days=60){
  const last=v15LastSaleMap(); const cutoff=Date.now()-days*86400000;
  return state.products.filter(p=>Number(p.stock)>0 && (!last[p.id] || last[p.id].getTime()<cutoff));
}
function v15SalesForDate(date, role=v15Role()){
  return state.sales.filter(s=>s.status!=="refunded" && s.date===date && (role!=="Cashier" || s.cashierUid===state.user.uid));
}
function v15TodayRevenue(){ return v15SalesForDate(todayISO()).reduce((a,s)=>a+(Number(s.total)||0),0); }
function v15TodayCOGS(){ return v15SalesForDate(todayISO()).reduce((a,s)=>(s.items||[]).reduce((x,i)=>x+(Number(i.cost)||0)*(Number(i.qty)||0),a),0); }
function v15TodayExpenses(){ return state.expenses.filter(e=>e.date===todayISO()).reduce((a,e)=>a+(Number(e.amount)||0),0); }
function v15TodayNet(){ return v15TodayRevenue()-v15TodayCOGS()-v15TodayExpenses(); }
function v15TargetPercent(){ return state.businessTarget>0 ? Math.min(100,(v15TodayRevenue()/state.businessTarget)*100) : 0; }
function v17CelebrateTarget(){
  const key=`${todayISO()}_${state.businessTarget}`;
  if(!state.businessTarget || v15TodayRevenue()<state.businessTarget) return;
  if(localStorage.getItem("gwaponile_target_celebrated")===key) return;
  localStorage.setItem("gwaponile_target_celebrated",key);
  const old=document.getElementById("v17Celebration"); if(old) old.remove();
  const layer=document.createElement("div"); layer.id="v17Celebration"; layer.className="v17-celebration";
  layer.innerHTML=`<div class="v17-celebration-box"><div class="v17-celebration-icon">🏆</div><h2>Target Imefikiwa! 🎉</h2><p>${fmt(v15TodayRevenue())} ya mauzo leo — umefikia lengo la ${fmt(state.businessTarget)}.</p></div>`;
  document.body.appendChild(layer);
  for(let i=0;i<42;i++){
    const c=document.createElement("i"); c.className="v17-confetti";
    c.style.left=(5+Math.random()*90)+"%";
    c.style.setProperty("--dx",((Math.random()-.5)*260)+"px");
    c.style.setProperty("--rot",((Math.random()*900)-450)+"deg");
    c.style.animationDelay=(Math.random()*.45)+"s";
    c.style.background=`hsl(${Math.floor(Math.random()*360)} 78% 55%)`;
    c.style.transform=`rotate(${Math.random()*180}deg)`;
    layer.appendChild(c);
  }
  setTimeout(()=>layer.remove(),5000);
}

async function v15SetTarget(){
  if(!["Administrator","Manager"].includes(v15Role())) return;
  const value=Number(prompt("Weka sales target ya leo (TZS):",state.businessTarget||0));
  if(!Number.isFinite(value)||value<0) return;
  try{
    await setDoc(doc(db,"settings","business"),{salesTarget:value,updatedBy:state.user.uid,updatedAt:serverTimestamp()},{merge:true});
    state.businessTarget=value;localStorage.setItem("duka_sales_target",String(value));showToast("Sales target imehifadhiwa");render();
  }catch(e){showToast("Target haijahifadhiwa: "+e.message);}
}
function v15Notifications(){
  const out=[];
  if(v15MonthlyReportDue()) out.push({level:"amber",title:"Monthly Report tayari",text:`Angalia ripoti ya ${v15PreviousMonthLabel()}`,tab:"report"});
  v15OverduePurchases().forEach(p=>out.push({level:"red",title:"Supplier payment overdue",text:`${p.supplierName||"Supplier"} — ${fmt(v15PurchaseBalance(p))}`,tab:"purchases"}));
  v15DueSoonPurchases().forEach(p=>out.push({level:"amber",title:p.dueDate===todayISO()?"Supplier payment due today":"Supplier payment due soon",text:`${p.supplierName||"Supplier"} — ${fmt(v15PurchaseBalance(p))} • ${p.dueDate}`,tab:"purchases"}));
  state.products.filter(p=>Number(p.stock)<=Number(p.reorder||0)).slice(0,20).forEach(p=>out.push({level:"amber",title:"Low stock",text:`${p.name} — ${p.stock} remaining`,tab:"inventory"}));
  state.products.filter(p=>Number(p.stock)<=0).slice(0,20).forEach(p=>out.push({level:"red",title:"Out of stock",text:p.name,tab:"inventory"}));
  return out.slice(0,30);
}
function v15NotifyDuePayments(){
  if(typeof Notification==="undefined" || Notification.permission!=="granted") return;
  const key=todayISO();
  const already=localStorage.getItem("duka_due_notified")===key;
  if(already) return;
  const urgent=[...v15OverduePurchases(),...v15DueSoonPurchases()];
  if(!urgent.length) return;
  try{ new Notification("SALES MANAGEMENT SYSTEM — Payment Alert",{body:`${urgent.length} supplier payment alert(s) zinahitaji attention.`}); localStorage.setItem("duka_due_notified",key); }catch(e){}
}
// Siku 1-3 za mwezi mpya: kumbusha Admin/Manager kuangalia Monthly Report
// ya mwezi uliopita. Inaonekana mara moja tu kwa mwezi (localStorage flag),
// na inahitaji mtu awe amefungua mfumo (siyo push ya nje kabisa).
function v15PreviousMonthLabel(){
  const d=new Date(); d.setDate(1); d.setMonth(d.getMonth()-1);
  return d.toLocaleDateString("en-GB",{month:"long",year:"numeric"});
}
function v15MonthlyReportDue(){
  if(!["Administrator","Manager"].includes(v15Role())) return false;
  const dayOfMonth=new Date().getDate();
  if(dayOfMonth>3) return false;
  const monthKey=todayISO().slice(0,7);
  return localStorage.getItem("duka_monthly_notified")!==monthKey;
}
function v15NotifyMonthlyReport(){
  if(!v15MonthlyReportDue()) return;
  const monthKey=todayISO().slice(0,7);
  const label=v15PreviousMonthLabel();
  if(typeof Notification!=="undefined" && Notification.permission==="granted"){
    try{ new Notification("SALES MANAGEMENT SYSTEM — Monthly Report",{body:`Monthly Report ya ${label} tayari — bofya kuangalia.`}); }catch(e){}
  }
  localStorage.setItem("duka_monthly_notified",monthKey);
}

// Extend listeners without breaking the original V1.4 listeners.
const v15OriginalStartListeners = startListeners;
function v15StartListeners(){
  v15OriginalStartListeners();
  // Sync the daily sales target from Firestore for Administrator/Manager.
  // This keeps the target visible and consistent across PC and phone.
  if(["Administrator","Manager"].includes(v15Role())){
    unsubs.push(onSnapshot(doc(db,"settings","business"),snap=>{
      const data=snap.exists() ? snap.data() : {};
      const next=Number(data.salesTarget)||0;
      state.businessTarget=next;
      localStorage.setItem("duka_sales_target",String(next));
      if(data.ownerWhatsApp) localStorage.setItem("duka_owner_wa",String(data.ownerWhatsApp));
      render();
    },err=>console.warn("Business settings listener:",err)));
  }
  if(["Administrator","Manager","Storekeeper"].includes(v15Role())){
    unsubs.push(onSnapshot(query(collection(db,"supplierPayments"),orderBy("createdAt","desc")),snap=>{
      state.supplierPayments=snap.docs.map(d=>({id:d.id,...d.data()}));
      render();
    }));
  }
  // Staff Chat messages: all active roles can read the common team room.
  // IMPORTANT: this listener must be attached to the actual startListeners()
  // function used after authentication; otherwise a sent message is written
  // to Firestore but never loaded back into the UI.
  unsubs.push(onSnapshot(collection(db,"staffChat"),snap=>{
    state.chatMessages=snap.docs.map(d=>({id:d.id,...d.data()}))
      .filter(m=>m.type==="text")
      .slice(-200);
    if(state.tab==="chat"){
      render();
      scrollStaffChatToBottom();
    }
  }, err=>{
    console.error("Staff Chat listener:", err);
    if(state.tab==="chat"){
      showToast("Chat haijasomeka kutoka Firebase: " + (err.message || err));
    }
  }));

  // Global holiday theme: all active employees can read this document.
  // Only Administrator can write it (enforced by Firestore Rules).
  unsubs.push(onSnapshot(doc(db,"settings","themeControl"),snap=>{
    const data=snap.exists() ? snap.data() : {};
    const next=(data.active && HOLIDAY_IDS.has(data.theme)) ? data.theme : null;
    const changed=state.globalThemeOverride!==next || !state.themeControlLoaded;
    state.globalThemeOverride=next;
    state.themeControlLoaded=true;
    if(changed){
      refreshEffectiveTheme();
      render();
    }
  }, err=>{
    console.error("Global theme listener:",err);
    state.themeControlLoaded=true;
  }));
}

// Use the extended listener set everywhere the app calls startListeners().
startListeners = v15StartListeners;

receiptInnerHTML = function(sale){
  const name=sale.customerName||"Walk-in Customer"; const phone=sale.customerPhone||"";
  const creditLine = sale.isCredit && (sale.creditAmount||0)>0 ? `<div class="receipt-sub" style="color:#b45309;font-weight:700;"><span>DENI (Outstanding)</span><span>${fmt(sale.creditAmount||0)}</span></div>` : "";
  return `<div class="receipt-head"><div class="rname">ST JOHN'S SHOP</div><div class="rmeta">Dodoma, Tanzania</div><div class="rmeta">Tel: 0755816562</div><div class="rmeta">RECEIPT: ${escapeHtml(sale.number)}</div><div class="rmeta">${escapeHtml(sale.date)} ${escapeHtml(sale.time)}</div><div class="rmeta">Cashier: ${escapeHtml(sale.cashierName||"—")}</div><div class="rmeta">Customer: ${escapeHtml(name)}${phone?` • ${escapeHtml(phone)}`:""}</div></div><div style="font-size:11px;font-weight:800;margin-bottom:5px;">ITEM             QTY     AMOUNT</div><div class="receipt-items">${(sale.items||[]).map(it=>`<div class="receipt-item"><span>${escapeHtml(it.name)} x${it.qty}</span><span>${fmt(it.lineTotal)}</span></div>`).join("")}</div><div class="receipt-sub"><span>SUBTOTAL</span><span>${fmt(sale.subtotal)}</span></div><div class="receipt-sub"><span>DISCOUNT</span><span>${fmt(sale.discountAmount||0)}</span></div><div class="receipt-total"><span>TOTAL</span><span>${fmt(sale.total)}</span></div><div class="receipt-sub"><span>Payment</span><span>${escapeHtml(sale.paymentMethod||"—")}</span></div><div class="receipt-sub"><span>Paid</span><span>${fmt(sale.amountPaid??sale.total)}</span></div><div class="receipt-sub"><span>Change</span><span>${fmt(sale.change??0)}</span></div>${creditLine}${sale.status==='refunded'?`<div class="rmeta" style="color:var(--red);font-weight:700;text-align:center;margin-top:8px;">REFUNDED</div>`:""}<div class="receipt-thanks">THANK YOU!<br/>Karibu tena.</div>`;
}
receiptText = function(sale,whatsapp=false){
  const lines=(sale.items||[]).map(it=>`${it.name} x${it.qty} = ${fmt(it.lineTotal)}`).join("\n"); const sep=whatsapp?"━━━━━━━━━━━━━━━━":"--------------------------------";const bold=whatsapp?"*":""; const customer=sale.customerName||"Walk-in Customer"; const phone=sale.customerPhone?` (${sale.customerPhone})`:"";
  const creditLine = sale.isCredit && (sale.creditAmount||0)>0 ? `\nDENI (Outstanding): ${fmt(sale.creditAmount||0)}` : "";
  return `${bold}ST JOHN'S SHOP${bold}\nDodoma, Tanzania\nTel: 0755816562\n\n${sep}\nRECEIPT: ${sale.number}\nDate: ${sale.date} ${sale.time}\nCashier: ${sale.cashierName||"—"}\nCustomer: ${customer}${phone}\n${sep}\nITEMS\n${lines}\n${sep}\nSUBTOTAL: ${fmt(sale.subtotal)}\nDISCOUNT: ${fmt(sale.discountAmount||0)}\n${bold}TOTAL: ${fmt(sale.total)}${bold}\n\nPayment: ${sale.paymentMethod||"—"}\nPaid: ${fmt(sale.amountPaid??sale.total)}\nChange: ${fmt(sale.change??0)}${creditLine}\n${sep}\n\n${bold}THANK YOU!${bold}\n${bold}Karibu tena.${bold}`;
}
// Sends the receipt as an IMAGE (like a real photo receipt) — always both
// downloads the PNG AND copies it to the clipboard, then opens the correct
// customer WhatsApp chat so the person just long-presses the message box and
// taps Paste. Falls back to a text-only message only if image generation
// itself fails (e.g. html2canvas can't load).
sendReceiptWhatsApp = async function(sale){
  const phone=normalizePhone(sale.customerPhone);
  if(!phone){showToast("Receipt haina namba ya simu ya customer.");return;}
  showToast("Inatengeneza picha ya receipt...");
  let blob=null;
  try{ blob=await receiptImageBlob(sale); }catch(e){ /* fall through to text-only link below */ }

  if(!blob){
    // Image generation failed entirely — never fall back to prefilled text,
    // just tell the person and stop (they can retry).
    showToast("Imeshindikana kutengeneza picha ya receipt. Jaribu tena.");
    return;
  }

  // 1) Always download the image so it's saved locally too.
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");
  a.href=url; a.download=`Receipt-${sale.number}.png`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),15000);

  // 2) Also copy it to the clipboard as an image, ready to paste.
  let copied=false;
  try{
    if(navigator.clipboard && window.ClipboardItem){
      await navigator.clipboard.write([ new ClipboardItem({ "image/png": blob }) ]);
      copied=true;
    }
  }catch(e){ /* clipboard image copy not supported on this browser */ }

  showToast(copied
    ? "Picha imepakuliwa na ime-wekwa kwenye clipboard — fungua chat, bonyeza kwa muda kisanduku cha kuandika kisha 'Paste'."
    : "Picha ya receipt imepakuliwa — iambatanishe (📎) ndani ya chat itakayofunguka.");

  // 3) Jump straight into that customer's WhatsApp chat — no text prefilled,
  // no searching needed — just paste (or attach) the image.
  window.open(`https://wa.me/${phone}`,"_blank");
}
sendReceiptSMS = function(sale){const phone=normalizePhone(sale.customerPhone);if(!phone){showToast("Receipt haina namba ya simu ya customer.");return;}window.location.href=`sms:${phone}?body=${encodeURIComponent(receiptText(sale,false))}`;}

renderSuppliers = function(){
  const q=(state.supplierSearch||"").toLowerCase(); const filtered=state.suppliers.filter(s=>String(s.name||"").toLowerCase().includes(q)||String(s.phone||"").includes(q));
  return `<div class="view-header"><div><h1>Supplier Ledger</h1><p>${state.suppliers.length} suppliers • credit, payments na balances.</p></div><button class="btn btn-dark" id="addSupplierBtn">${ICONS.plus} Ongeza supplier</button></div>
  <div class="search-wrap">${ICONS.search}<input id="supplierSearchInput" placeholder="Tafuta supplier..." value="${escapeHtml(state.supplierSearch)}"/></div>
  <div class="table-card"><table><thead><tr><th>Supplier</th><th>Phone</th><th class="right">Outstanding</th><th>Next Due</th><th>Status</th><th></th></tr></thead><tbody>${filtered.length?filtered.map(s=>{const debt=v15SupplierDebt(s.id);const ps=state.purchases.filter(p=>p.supplierId===s.id&&v15PurchaseBalance(p)>0).sort((a,b)=>String(a.dueDate||"9999").localeCompare(String(b.dueDate||"9999")));const due=ps[0]?.dueDate||"—";const overdue=ps.some(p=>p.dueDate&&v15DaysUntil(p.dueDate)<0);return `<tr><td class="cell-title">${escapeHtml(s.name)}</td><td>${escapeHtml(s.phone||"—")}</td><td class="right mono" style="font-weight:800;">${fmt(debt)}</td><td>${escapeHtml(due)}</td><td>${debt<=0?'<span class="pill-good">PAID</span>':overdue?'<span class="pill-low">OVERDUE</span>':'<span class="pill-mid">CREDIT</span>'}</td><td><div class="row-actions"><button class="btn btn-outline" data-v15-supplier-statement="${s.id}">Statement</button><button class="icon-btn" data-edit-supplier="${s.id}">${ICONS.edit}</button>${CAN_DELETE_SUPPLIER(v15Role())?`<button class="icon-btn danger" data-del-supplier="${s.id}">${ICONS.trash}</button>`:""}</div></td></tr>`;}).join(""):`<tr class="empty-row"><td colspan="6">Hakuna suppliers.</td></tr>`}</tbody></table></div>`;
}

function openSupplierStatement(supplierId){
  const s=state.suppliers.find(x=>x.id===supplierId); if(!s)return;
  const purchases=state.purchases.filter(p=>p.supplierId===supplierId); const balance=purchases.reduce((a,p)=>a+v15PurchaseBalance(p),0);
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal wide"><div class="modal-head"><h3>${escapeHtml(s.name)} — Supplier Statement</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><div class="kpi-grid"><div class="kpi-card"><div class="kpi-label">Purchases</div><div class="kpi-value">${fmt(purchases.reduce((a,p)=>a+p.total,0))}</div></div><div class="kpi-card"><div class="kpi-label">Outstanding</div><div class="kpi-value">${fmt(balance)}</div></div></div><div class="table-card"><table><thead><tr><th>Date</th><th>Purchase</th><th>Status</th><th>Due</th><th class="right">Balance</th><th></th></tr></thead><tbody>${purchases.map(p=>`<tr><td>${p.date||"—"}</td><td class="mono">${escapeHtml(p.number||p.id)}</td><td>${escapeHtml(p.paymentStatus||"PAID")}</td><td>${escapeHtml(p.dueDate||"—")}</td><td class="right mono">${fmt(v15PurchaseBalance(p))}</td><td>${V15_CAN_MANAGE_SUPPLIER_PAYMENTS(v15Role())&&v15PurchaseBalance(p)>0?`<button class="btn btn-outline" data-v15-pay-purchase="${p.id}">Record Payment</button>`:""}</td></tr>`).join("")}</tbody></table></div></div></div></div>`;
  document.getElementById("closeModalBtn").onclick=closeModal;document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal()};
  document.querySelectorAll("[data-v15-pay-purchase]").forEach(b=>b.onclick=()=>openSupplierPaymentModal(state.purchases.find(p=>p.id===b.dataset.v15PayPurchase)));
}
function openSupplierPaymentModal(purchase){
  if(!purchase||!V15_CAN_MANAGE_SUPPLIER_PAYMENTS(v15Role()))return;
  const balance=v15PurchaseBalance(purchase);
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal"><div class="modal-head"><h3>Record Supplier Payment</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><p class="muted">${escapeHtml(purchase.supplierName||"Supplier")} • ${escapeHtml(purchase.number||purchase.id)}</p><div class="kpi-card" style="margin-bottom:14px;"><div class="kpi-label">Outstanding</div><div class="kpi-value">${fmt(balance)}</div></div><form id="v15PaymentForm"><div class="field"><label>Amount (TZS)</label><input required id="v15PayAmount" type="number" min="1" max="${balance}"/></div><div class="field"><label>Payment Method</label><select id="v15PayMethod">${PAYMENT_METHODS.map(m=>`<option>${m}</option>`).join("")}</select></div><div class="field"><label>Payment Date</label><input id="v15PayDate" type="date" value="${todayISO()}"/></div><button class="btn btn-amber btn-block">Save Payment</button></form></div></div></div>`;
  document.getElementById("closeModalBtn").onclick=closeModal;document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal()};
  document.getElementById("v15PaymentForm").onsubmit=async e=>{e.preventDefault();const amount=Number(document.getElementById("v15PayAmount").value);if(amount<=0||amount>balance){showToast("Amount si sahihi.");return;}try{await addDoc(collection(db,"supplierPayments"),{purchaseId:purchase.id,supplierId:purchase.supplierId,supplierName:purchase.supplierName,amount,paymentMethod:document.getElementById("v15PayMethod").value,date:document.getElementById("v15PayDate").value,createdBy:state.user.uid,createdAt:serverTimestamp()});await logAudit("Supplier payment",`${purchase.supplierName} — ${fmt(amount)}`);closeModal();showToast("Supplier payment imehifadhiwa");}catch(err){showToast("Payment haikuhifadhiwa: "+err.message)}};
}

renderPurchases = function(){
  const total=state.purchases.reduce((a,p)=>a+(Number(p.total)||0),0); const outstanding=state.purchases.reduce((a,p)=>a+v15PurchaseBalance(p),0); const overdue=v15OverduePurchases().reduce((a,p)=>a+v15PurchaseBalance(p),0);
  return `<div class="view-header"><div><h1>Purchases & Supplier Credit</h1><p>Receive stock, track paid/credit purchases na supplier debt.</p></div><button class="btn btn-dark" id="receiveStockBtn">${ICONS.plus} Receive Stock</button></div>
  <div class="kpi-grid"><div class="kpi-card"><div class="kpi-label">Purchase Value</div><div class="kpi-value">${fmt(total)}</div></div><div class="kpi-card"><div class="kpi-label">Outstanding Debt</div><div class="kpi-value" style="color:${outstanding?'var(--red)':'inherit'}">${fmt(outstanding)}</div></div><div class="kpi-card"><div class="kpi-label">Overdue</div><div class="kpi-value" style="color:${overdue?'var(--red)':'inherit'}">${fmt(overdue)}</div></div></div>
  <div class="table-card"><table><thead><tr><th>Date</th><th>Supplier</th><th>Purchase</th><th>Payment</th><th>Due Date</th><th class="right">Total</th><th class="right">Balance</th><th></th></tr></thead><tbody>${state.purchases.length?state.purchases.map(p=>{const bal=v15PurchaseBalance(p);const status=bal<=0?'PAID':(p.dueDate&&v15DaysUntil(p.dueDate)<0?'OVERDUE':(v15PurchasePaid(p)>0?'PARTIAL':'CREDIT'));return `<tr><td>${escapeHtml(p.date||"")}</td><td>${escapeHtml(p.supplierName||"—")}</td><td class="mono">${escapeHtml(p.number||p.id)}</td><td>${status==='PAID'?'<span class="pill-good">PAID</span>':status==='OVERDUE'?'<span class="pill-low">OVERDUE</span>':'<span class="pill-mid">'+status+'</span>'}</td><td>${escapeHtml(p.dueDate||"—")}</td><td class="right mono">${fmt(p.total)}</td><td class="right mono" style="font-weight:800;">${fmt(bal)}</td><td>${V15_CAN_MANAGE_SUPPLIER_PAYMENTS(v15Role())&&bal>0?`<button class="btn btn-outline" data-v15-pay-purchase="${p.id}">Pay</button>`:""}</td></tr>`;}).join(""):`<tr class="empty-row"><td colspan="8">Hakuna purchases.</td></tr>`}</tbody></table></div>`;
}

openPurchaseModal = function(){
  const options=state.products.map(p=>`<option value="${p.id}">${escapeHtml(p.name)} — ${p.stock} stock</option>`).join(""); const suppliers=state.suppliers.map(s=>`<option value="${s.id}">${escapeHtml(s.name)}</option>`).join("");
  document.getElementById("modalRoot").innerHTML=`<div class="modal-overlay" id="overlay"><div class="modal"><div class="modal-head"><h3>Receive Stock & Record Payment</h3><button class="icon-btn" id="closeModalBtn">${ICONS.close}</button></div><div class="modal-body"><form id="purchaseForm"><div class="field"><label>Supplier</label><select id="puSupplier" required>${suppliers||"<option value=''>Hakuna supplier</option>"}</select></div><div class="field"><label>Product</label><select id="puProduct" required>${options}</select></div><div class="field-row"><div class="field"><label>Quantity</label><input type="number" min="1" required id="puQty"/></div><div class="field"><label>Unit Cost (TZS)</label><input type="number" min="0" required id="puCost"/></div></div><div class="field" style="background:var(--paper);border-radius:8px;padding:10px 12px;"><label style="margin-bottom:2px;">Total Cost (TZS)</label><div id="puTotalCost" style="font-size:19px;font-weight:800;">TZS 0</div></div><div class="field"><label>Payment Status</label><select id="puPaymentStatus"><option value="paid">Paid in Full</option><option value="credit">Credit / Kopa</option><option value="partial">Partial Payment</option></select></div><div id="puPaymentBox"><div class="field"><label>Amount Paid Now (TZS)</label><input type="number" min="0" id="puPaid" value="0"/></div><div class="field"><label>Due Date</label><input type="date" id="puDueDate"/></div></div><div class="v15-info">Credit purchases will appear in Supplier Ledger. Due/overdue alerts zitaonekana kwenye notification center.</div><button class="btn btn-amber btn-block">Receive Stock & Save Purchase</button></form></div></div></div>`;
  const statusEl=document.getElementById("puPaymentStatus"), box=document.getElementById("puPaymentBox"); const sync=()=>{box.style.display=statusEl.value==="paid"?"none":"block";if(statusEl.value==="paid")document.getElementById("puPaid").value="0";}; statusEl.onchange=sync;sync();
  const qtyEl=document.getElementById("puQty"), costEl=document.getElementById("puCost"), totalEl=document.getElementById("puTotalCost");
  const syncTotal=()=>{ const t=(Number(qtyEl.value)||0)*(Number(costEl.value)||0); totalEl.textContent=fmt(t); };
  qtyEl.addEventListener("input", syncTotal); costEl.addEventListener("input", syncTotal); syncTotal();
  document.getElementById("closeModalBtn").onclick=closeModal;document.getElementById("overlay").onclick=e=>{if(e.target.id==="overlay")closeModal()};
  document.getElementById("purchaseForm").onsubmit=async e=>{e.preventDefault();const productId=document.getElementById("puProduct").value,qty=Number(document.getElementById("puQty").value),unitCost=Number(document.getElementById("puCost").value),supplierId=document.getElementById("puSupplier").value,supplier=state.suppliers.find(x=>x.id===supplierId),status=statusEl.value;const total=qty*unitCost;let paid=status==="paid"?total:Number(document.getElementById("puPaid").value)||0;if(status==="partial"&&(paid<=0||paid>=total)){showToast("Partial payment lazima iwe chini ya total na zaidi ya zero.");return;}if(status==="credit")paid=0;if(!supplierId||!productId||qty<=0||unitCost<0){showToast("Weka taarifa sahihi.");return;}const due=status==="paid"?null:document.getElementById("puDueDate").value;if(status!=="paid"&&!due){showToast("Weka Due Date kwa credit/partial purchase.");return;}try{const pRef=doc(db,"products",productId),purchaseRef=doc(collection(db,"purchases")),moveRef=doc(collection(db,"stockMoves"));await runTransaction(db,async tx=>{const ps=await tx.get(pRef);if(!ps.exists())throw new Error("Product haipo");const p=ps.data();const item={productId,name:p.name,qty,unitCost,lineTotal:total};tx.update(pRef,{stock:(p.stock||0)+qty,costPrice:unitCost});tx.set(purchaseRef,{number:`PUR-${todayISO().replaceAll("-","")}-${String(Date.now()).slice(-6)}`,date:todayISO(),supplierId,supplierName:supplier?.name||"—",items:[item],total,paidAmount:paid,balance:Math.max(0,total-paid),paymentStatus:status,dueDate:due,createdBy:state.user.uid,createdAt:serverTimestamp()});tx.set(moveRef,{productId,productName:p.name,type:"purchase",qty,referenceId:purchaseRef.id,actorUid:state.user.uid,actorName:state.profile.name,createdAt:serverTimestamp()});});await logAudit("Purchase received",`${supplier?.name||"Supplier"} — ${fmt(total)} — ${status}`);closeModal();showToast(status==="paid"?"Purchase paid na stock imeongezwa":"Credit purchase imehifadhiwa");}catch(err){showToast("Hitilafu: "+err.message)}};
}

renderReport = function(){
  const key=state.reportMonth; const sales=state.sales.filter(s=>monthKey(s.date)===key&&s.status!=="refunded");
  let revenue=0,cogs=0;const totals={};sales.forEach(s=>{revenue+=Number(s.total)||0;(s.items||[]).forEach(i=>{cogs+=(Number(i.cost)||0)*(Number(i.qty)||0);totals[i.name]=(totals[i.name]||0)+(Number(i.lineTotal)||0);});});
  const expenses=state.expenses.filter(e=>monthKey(e.date||"")===key).reduce((a,e)=>a+(Number(e.amount)||0),0);const gross=revenue-cogs;const net=gross-expenses;const top=Object.entries(totals).sort((a,b)=>b[1]-a[1]).slice(0,8);
  return `<div class="view-header"><div><h1>Monthly Business Report</h1><p>Revenue, COGS, expenses na true net profit.</p></div><input type="month" id="reportMonthSelect" value="${escapeHtml(key)}" style="max-width:190px;"/></div><div class="kpi-grid"><div class="kpi-card"><div class="kpi-label">Revenue</div><div class="kpi-value">${fmt(revenue)}</div></div><div class="kpi-card"><div class="kpi-label">COGS</div><div class="kpi-value">${fmt(cogs)}</div></div><div class="kpi-card"><div class="kpi-label">Gross Profit</div><div class="kpi-value">${fmt(gross)}</div></div><div class="kpi-card"><div class="kpi-label">Net Profit</div><div class="kpi-value" style="color:${net>=0?'var(--green)':'var(--red)'}">${fmt(net)}</div></div></div><div class="grid-2"><div class="card"><h2>Top Products</h2>${top.length?`<table><tbody>${top.map(([n,v])=>`<tr><td>${escapeHtml(n)}</td><td class="right mono">${fmt(v)}</td></tr>`).join("")}</tbody></table>`:`<p class="empty-note">Hakuna data.</p>`}</div><div class="card"><h2>Inventory Position</h2><div class="v15-snapshot-grid"><div><span>Stock Cost Value</span><b>${fmt(v15StockValue())}</b></div><div><span>Potential Sales</span><b>${fmt(v15PotentialSalesValue())}</b></div><div><span>Potential Gross Profit</span><b>${fmt(v15PotentialSalesValue()-v15StockValue())}</b></div><div><span>Dead Stock Value</span><b>${fmt(v15DeadStock(60).reduce((a,p)=>a+(Number(p.stock)||0)*(Number(p.costPrice)||0),0))}</b></div></div></div></div>`;
}

renderBackup = function(){
  return `<div class="view-header"><div><h1>Backup & Export Center</h1><p>Export sales, stock, customers, supplier credit na finance — professional Excel files.</p></div></div><div class="card-grid"><div class="card"><h2>Excel Exports</h2><div class="stack"><button class="btn btn-outline" id="exportSalesBtn">${ICONS.file||""} Sales (Excel)</button><button class="btn btn-outline" id="exportProductsBtn">Inventory (Excel)</button><button class="btn btn-outline" id="exportCustomersBtn">Customers (Excel)</button><button class="btn btn-outline" id="exportExpensesBtn">Expenses (Excel)</button><button class="btn btn-outline" id="exportPurchasesBtn">Purchases (Excel)</button><button class="btn btn-outline" id="exportSupplierPaymentsBtn">Supplier Payments (Excel)</button></div></div><div class="card"><h2>Full Business Backup</h2><p class="muted">Includes supplier payments and all operational collections visible to your role.</p><button class="btn btn-dark" id="exportJsonBtn">Download Full Backup</button></div></div>`;
}
exportFullBackup = function(){const backup={version:"V1.5",exportedAt:new Date().toISOString(),products:state.products,customers:state.customers,sales:state.sales,expenses:state.expenses,purchases:state.purchases,stockMoves:state.stockMoves,suppliers:state.suppliers,supplierPayments:state.supplierPayments,shifts:state.shifts};downloadText(`duka-manager-v1.5-backup-${todayISO()}.json`,JSON.stringify(backup,null,2),"application/json");logAudit("Backup exported","Full V1.5 JSON backup");}

// Keep the original view/event system and add V1.5 handlers on top.
const v15OriginalBindViewEvents=bindViewEvents;
bindViewEvents = function(){
  v15OriginalBindViewEvents();
  if(state.tab==="dashboard"){
    document.querySelectorAll("[data-v15-tab]").forEach(b=>b.onclick=()=>{state.tab=b.dataset.v15Tab;render();});
    const t=document.getElementById("v15TargetBtn");if(t)t.onclick=v15SetTarget;
    requestAnimationFrame(()=>{
      const bar=document.querySelector(".v17-target-progress span");
      if(bar){const targetWidth=bar.style.width;bar.style.width="0%";requestAnimationFrame(()=>bar.style.width=targetWidth);}
      v17CelebrateTarget();
    });
  }
  if(state.tab==="sale"){
    const sel=document.getElementById("saleCustomerSelect");if(sel)sel.onchange=e=>{state.saleCustomerId=e.target.value||null;state.tempCustomerName="";state.tempCustomerPhone="";render();};
    const n=document.getElementById("tempCustomerName");if(n)n.oninput=e=>state.tempCustomerName=e.target.value;
    const ph=document.getElementById("tempCustomerPhone");if(ph)ph.oninput=e=>state.tempCustomerPhone=e.target.value;
    const save=document.getElementById("saveWalkInCustomer");if(save)save.onchange=e=>state.saveWalkInCustomer=e.target.checked;
  }
  if(state.tab==="suppliers"){
    document.querySelectorAll("[data-v15-supplier-statement]").forEach(b=>b.onclick=e=>openSupplierStatement(b.dataset.v15SupplierStatement));
    document.querySelectorAll("[data-v15-pay-purchase]").forEach(b=>b.onclick=e=>openSupplierPaymentModal(state.purchases.find(p=>p.id===b.dataset.v15PayPurchase)));
  }
  if(state.tab==="purchases") document.querySelectorAll("[data-v15-pay-purchase]").forEach(b=>b.onclick=e=>openSupplierPaymentModal(state.purchases.find(p=>p.id===b.dataset.v15PayPurchase)));
  if(state.tab==="audit"){const a=document.getElementById("auditSearchInput");if(a)a.oninput=e=>{state.auditSearch=e.target.value;render();const el=document.getElementById("auditSearchInput");if(el){el.focus();el.selectionStart=el.selectionEnd=el.value.length;}};const c=document.getElementById("deleteOldAuditBtn");if(c)c.onclick=deleteAuditLogsOlderThan30Days;}
  if(state.tab==="backup"){const sp=document.getElementById("exportSupplierPaymentsBtn");if(sp)sp.onclick=()=>exportRowsXlsx(`supplier-payments-${todayISO()}.xlsx`,state.supplierPayments,"Supplier Payments");}
  if(state.tab==="chat"){
    const input=document.getElementById("staffChatInput");
    if(input){
      input.addEventListener("input",e=>{state.chatText=e.target.value;});
      input.addEventListener("keydown",e=>{if(e.key==="Enter"&&!e.shiftKey){e.preventDefault();sendStaffChatText();}});
    }
    const send=document.getElementById("staffSendBtn");
    if(send) send.onclick=sendStaffChatText;
    const cancel=document.getElementById("cancelChatReply");
    if(cancel) cancel.onclick=cancelChatReply;
    document.querySelectorAll("[data-chat-reply]").forEach(btn=>{
      btn.onclick=()=>chooseChatReply(btn.dataset.chatReply);
    });
    scrollStaffChatToBottom();
  }
  const notify=document.getElementById("v15NotifyBtn");
  if(notify){
    notify.onclick=e=>{
      e.stopPropagation();
      document.getElementById("v15NotifyPanel")?.classList.toggle("show");
    };
    document.querySelectorAll("[data-v15-notify-tab]").forEach(b=>b.onclick=()=>{
      state.tab=b.dataset.v15NotifyTab;
      render();
    });
  }
  const notifyWrap=document.getElementById("v15NotifyWrap");
  if(notifyWrap){
    document.addEventListener("click", function closeNotifyOutside(e){
      if(!notifyWrap.contains(e.target)) document.getElementById("v15NotifyPanel")?.classList.remove("show");
    }, {once:true});
  }
}

// Fix receipt refund for temporary Walk-in customers: customerId may be null.
refundSale = async function(sale){
  if(!CAN_REFUND_SALE(v15Role())){showToast("Administrator au Manager pekee.");return;}
  if(!confirm(`Una uhakika unataka kurefund Sale #${sale.number} — ${fmt(sale.total)}?`))return;
  try{const saleRef=doc(db,"sales",sale.id);await runTransaction(db,async tx=>{const ss=await tx.get(saleRef);if(!ss.exists())throw new Error("Sale haipo.");const live=ss.data();if(live.status==="refunded")throw new Error("Sale tayari ime-refund.");let cSnap=null;if(live.customerId)cSnap=await tx.get(doc(db,"customers",live.customerId));const reads=[];for(const it of live.items||[]){const ref=doc(db,"products",it.productId);reads.push({it,ref,snap:await tx.get(ref)});}if(cSnap?.exists()){const c=cSnap.data();tx.update(doc(db,"customers",live.customerId),{totalSpent:Math.max(0,(c.totalSpent||0)-live.total),orders:Math.max(0,(c.orders||0)-1)});}for(const x of reads)if(x.snap.exists()){const p=x.snap.data();tx.update(x.ref,{stock:(p.stock||0)+x.it.qty});tx.set(doc(collection(db,"stockMoves")),{productId:x.it.productId,productName:x.it.name,type:"refund",qty:x.it.qty,referenceId:sale.id,shiftId:live.shiftId||null,actorUid:state.user.uid,actorName:state.profile.name,createdAt:serverTimestamp()});}tx.update(saleRef,{status:"refunded",refundedAt:serverTimestamp(),refundedBy:state.user.uid});});await logAudit("Sale refunded",`#${sale.number} — ${fmt(sale.total)}`);closeModal();showToast(`Sale #${sale.number} ime-refund`);}catch(e){showToast("Refund haikukamilika: "+e.message);}}

renderAuditLog = function(){
  const q=(state.auditSearch||"").toLowerCase();
  const logs=state.auditLog.filter(a=>!q||String(a.actorName||"").toLowerCase().includes(q)||String(a.action||"").toLowerCase().includes(q)||String(a.details||"").toLowerCase().includes(q));
  return `<div class="view-header"><div><h1>Audit Log</h1><p>History ya actions muhimu — ${logs.length} records shown.</p></div>${CAN_MANAGE_USERS(v15Role())?`<button class="btn btn-outline" id="deleteOldAuditBtn">${ICONS.trash} Cleanup 30+ days</button>`:""}</div><div class="search-wrap" style="margin-bottom:14px;">${ICONS.search}<input id="auditSearchInput" placeholder="Search actor, action or details..." value="${escapeHtml(state.auditSearch||"")}"/></div><div class="table-card"><table><thead><tr><th>Date</th><th>Actor</th><th>Action</th><th>Details</th></tr></thead><tbody>${logs.length?logs.map(a=>`<tr><td>${escapeHtml(a.date||"")} ${escapeHtml(a.time||"")}</td><td class="cell-title">${escapeHtml(a.actorName||"—")}</td><td>${escapeHtml(a.action||"")}</td><td style="font-size:12px;color:var(--muted);">${escapeHtml(a.details||"")}</td></tr>`).join(""):`<tr class="empty-row"><td colspan="4">Hakuna audit records.</td></tr>`}</tbody></table></div>`;
}

/* ============================================================
   INIT
============================================================ */
render();
