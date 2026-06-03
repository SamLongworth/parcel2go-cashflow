import { useState, useEffect, useRef } from "react";
import { AreaChart, Area, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine, Legend, LineChart, Line } from "recharts";
import * as XLSX from "xlsx";

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const TODAY = new Date();
const fmt = (d) => new Date(d).toLocaleDateString("en-GB", { day: "2-digit", month: "short" });
const fmtGBP = (n) => `£${Math.abs(n || 0).toLocaleString("en-GB", { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
const monthEnd = new Date(TODAY.getFullYear(), TODAY.getMonth() + 1, 0);

// ─── SAMPLE DATA ──────────────────────────────────────────────────────────────
const SAMPLE_RECEIPTS = [
  { customer_name: "Amazon Logistics UK",    invoice_number: "INV-2025-0441", amount: 48500, expected_payment_date: addDays(TODAY,  2), status: "Confirmed", source: "sample" },
  { customer_name: "DHL Express",            invoice_number: "INV-2025-0442", amount: 32000, expected_payment_date: addDays(TODAY,  4), status: "Confirmed", source: "sample" },
  { customer_name: "Next Retail Ltd",        invoice_number: "INV-2025-0443", amount: 19750, expected_payment_date: addDays(TODAY,  5), status: "Expected",  source: "sample" },
  { customer_name: "ASOS plc",               invoice_number: "INV-2025-0444", amount: 27300, expected_payment_date: addDays(TODAY,  7), status: "Expected",  source: "sample" },
  { customer_name: "Marks & Spencer",        invoice_number: "INV-2025-0446", amount: 38600, expected_payment_date: addDays(TODAY, 12), status: "Expected",  source: "sample" },
  { customer_name: "Boohoo Group",           invoice_number: "INV-2025-0447", amount: 11400, expected_payment_date: addDays(TODAY, 14), status: "Uncertain", source: "sample" },
  { customer_name: "John Lewis Partnership", invoice_number: "INV-2025-0449", amount: 54000, expected_payment_date: addDays(TODAY, 19), status: "Confirmed", source: "sample" },
  { customer_name: "Tesco Direct",           invoice_number: "INV-2025-0451", amount: 29500, expected_payment_date: addDays(TODAY, 24), status: "Expected",  source: "sample" },
  { customer_name: "Argos Ltd",              invoice_number: "INV-2025-0452", amount: 18200, expected_payment_date: addDays(TODAY, 26), status: "Uncertain", source: "sample" },
];
const SAMPLE_PAYMENTS = [
  { supplier_name: "Royal Mail Group",      invoice_number: "SUP-2025-1101", amount: 28500, expected_payment_date: addDays(TODAY,  1), status: "Due",   source: "sample" },
  { supplier_name: "Fuel & Fleet Services", invoice_number: "SUP-2025-1102", amount: 15600, expected_payment_date: addDays(TODAY,  3), status: "Due",   source: "sample" },
  { supplier_name: "Warehouse Ops Ltd",     invoice_number: "SUP-2025-1103", amount: 22000, expected_payment_date: addDays(TODAY,  6), status: "Due",   source: "sample" },
  { supplier_name: "HMRC VAT Payment",      invoice_number: "SUP-2025-1105", amount: 41200, expected_payment_date: addDays(TODAY, 10), status: "Fixed", source: "sample" },
  { supplier_name: "Staff Payroll",         invoice_number: "SUP-2025-1106", amount: 67500, expected_payment_date: addDays(TODAY, 13), status: "Fixed", source: "sample" },
  { supplier_name: "Vehicle Leasing Co",    invoice_number: "SUP-2025-1107", amount: 12300, expected_payment_date: addDays(TODAY, 15), status: "Due",   source: "sample" },
  { supplier_name: "Property Rent",         invoice_number: "SUP-2025-1109", amount: 18500, expected_payment_date: addDays(TODAY, 20), status: "Fixed", source: "sample" },
  { supplier_name: "Marketing Agency",      invoice_number: "SUP-2025-1110", amount:  6400, expected_payment_date: addDays(TODAY, 22), status: "Due",   source: "sample" },
  { supplier_name: "Fuel & Fleet Services", invoice_number: "SUP-2025-1112", amount: 13800, expected_payment_date: addDays(TODAY, 27), status: "Due",   source: "sample" },
];
const SAMPLE_BALANCE = 185000;

// ─── CONNECTORS ───────────────────────────────────────────────────────────────
async function fetchStripeData(secretKey) {
  const h = { Authorization: `Bearer ${secretKey}` };
  const base = "https://api.stripe.com/v1";
  const piRes = await fetch(`${base}/payment_intents?limit=50&created[gte]=${Math.floor(addDays(TODAY,-30).getTime()/1000)}`, { headers: h });
  if (!piRes.ok) throw new Error(`Stripe ${piRes.status}`);
  const piData = await piRes.json();
  const poRes  = await fetch(`${base}/payouts?limit=20`, { headers: h });
  const poData = poRes.ok ? await poRes.json() : { data: [] };
  const balRes = await fetch(`${base}/balance`, { headers: h });
  const balData = balRes.ok ? await balRes.json() : null;
  return {
    receipts: piData.data.filter(p=>["requires_capture","processing","succeeded"].includes(p.status)).map(pi=>({ customer_name:pi.description||pi.metadata?.customer_name||"Stripe Customer", invoice_number:pi.id, amount:pi.amount/100, expected_payment_date:new Date(pi.created*1000), status:pi.status==="succeeded"?"Confirmed":"Expected", source:"stripe" })),
    payments: poData.data.filter(p=>p.status!=="failed").map(po=>({ supplier_name:`Stripe Payout → ${po.destination||"Bank"}`, invoice_number:po.id, amount:po.amount/100, expected_payment_date:new Date(po.arrival_date*1000), status:po.status==="paid"?"Fixed":"Due", source:"stripe" })),
    stripeBalance: balData ? (balData.available?.reduce((s,b)=>s+b.amount,0)||0)/100 : 0,
  };
}
async function fetchXeroData(accessToken, tenantId) {
  const h = { Authorization:`Bearer ${accessToken}`, "Xero-tenant-id":tenantId, Accept:"application/json" };
  const base = "https://api.xero.com/api.xro/2.0";
  const sRes = await fetch(`${base}/Invoices?where=Type%3D%3D%22ACCREC%22%26%26Status%3D%3D%22AUTHORISED%22&order=DueDate`, { headers:h });
  if (!sRes.ok) throw new Error(`Xero ${sRes.status}`);
  const sData = await sRes.json();
  const bRes  = await fetch(`${base}/Invoices?where=Type%3D%3D%22ACCPAY%22%26%26Status%3D%3D%22AUTHORISED%22&order=DueDate`, { headers:h });
  const bData = bRes.ok ? await bRes.json() : { Invoices:[] };
  const parseXeroDate = s => new Date(parseInt((s||"").replace(/\/Date\((\d+).*\)\//,"$1"))||Date.now());
  return {
    receipts: (sData.Invoices||[]).map(inv=>({ customer_name:inv.Contact?.Name||"Unknown", invoice_number:inv.InvoiceNumber||inv.InvoiceID, amount:inv.AmountDue, expected_payment_date:parseXeroDate(inv.DueDate), status:"Expected", source:"xero" })),
    payments: (bData.Invoices||[]).map(inv=>({ supplier_name:inv.Contact?.Name||"Unknown", invoice_number:inv.InvoiceNumber||inv.InvoiceID, amount:inv.AmountDue, expected_payment_date:parseXeroDate(inv.DueDate), status:"Due", source:"xero" })),
  };
}
async function fetchBankingData(token) {
  const h = { Authorization:`Bearer ${token}` };
  const base = "https://api.truelayer.com/data/v1";
  const aRes = await fetch(`${base}/accounts`, { headers:h });
  if (!aRes.ok) throw new Error(`TrueLayer ${aRes.status}`);
  const aData = await aRes.json();
  let totalBalance=0; const allTx=[];
  for (const acct of (aData.results||[]).slice(0,3)) {
    const bRes = await fetch(`${base}/accounts/${acct.account_id}/balance`,{headers:h});
    if (bRes.ok) { const bd=await bRes.json(); totalBalance+=bd.results?.[0]?.available||0; }
    const from=addDays(TODAY,-30).toISOString().split("T")[0];
    const tRes = await fetch(`${base}/accounts/${acct.account_id}/transactions?from=${from}`,{headers:h});
    if (tRes.ok) { const td=await tRes.json(); allTx.push(...(td.results||[]).map(t=>({...t,accountName:acct.display_name}))); }
  }
  return {
    receipts: allTx.filter(t=>t.amount>0&&new Date(t.timestamp)>=TODAY).map(t=>({ customer_name:t.merchant_name||t.description||"Bank Credit", invoice_number:t.transaction_id, amount:t.amount, expected_payment_date:new Date(t.timestamp), status:"Confirmed", source:"bank" })),
    payments: allTx.filter(t=>t.amount<0&&new Date(t.timestamp)>=TODAY).map(t=>({ supplier_name:t.merchant_name||t.description||"Bank Debit", invoice_number:t.transaction_id, amount:Math.abs(t.amount), expected_payment_date:new Date(t.timestamp), status:"Fixed", source:"bank" })),
    bankBalance: totalBalance,
  };
}
async function runLiveSync(keys) {
  const result = { receipts:[], payments:[], balance:SAMPLE_BALANCE, log:[], errors:[] };
  if (keys.stripe)  { try { const d=await fetchStripeData(keys.stripe);  result.receipts.push(...d.receipts); result.payments.push(...d.payments); result.balance+=d.stripeBalance; result.log.push({source:"stripe",status:"ok",count:d.receipts.length+d.payments.length}); } catch(e){ result.errors.push({source:"stripe",message:e.message}); result.log.push({source:"stripe",status:"error",message:e.message}); } }
  if (keys.xeroToken&&keys.xeroTenant) { try { const d=await fetchXeroData(keys.xeroToken,keys.xeroTenant); result.receipts.push(...d.receipts); result.payments.push(...d.payments); result.log.push({source:"xero",status:"ok",count:d.receipts.length+d.payments.length}); } catch(e){ result.errors.push({source:"xero",message:e.message}); result.log.push({source:"xero",status:"error",message:e.message}); } }
  if (keys.truelayer) { try { const d=await fetchBankingData(keys.truelayer); result.receipts.push(...d.receipts); result.payments.push(...d.payments); result.balance=d.bankBalance; result.log.push({source:"bank",status:"ok",count:d.receipts.length+d.payments.length}); } catch(e){ result.errors.push({source:"bank",message:e.message}); result.log.push({source:"bank",status:"error",message:e.message}); } }
  if (!result.receipts.length) result.receipts=SAMPLE_RECEIPTS;
  if (!result.payments.length) result.payments=SAMPLE_PAYMENTS;
  return result;
}

// ─── FORECAST ENGINE ──────────────────────────────────────────────────────────
// paymentDelays: { [invoice_number]: daysToDelay }
// globalPaymentDelay: applied to ALL payments on top of individual delays
function buildForecast(receipts, payments, opening, receiptDelayDays=0, receiptHaircut=0, paymentDelays={}, globalPaymentDelay=0) {
  const adjReceipts = receipts.map(r => ({
    ...r,
    amount: r.amount * (1 - receiptHaircut / 100),
    expected_payment_date: receiptDelayDays ? addDays(r.expected_payment_date, receiptDelayDays) : r.expected_payment_date,
  }));
  const adjPayments = payments.map(p => {
    const indivDelay = paymentDelays[p.invoice_number] || 0;
    const totalDelay = indivDelay + globalPaymentDelay;
    return {
      ...p,
      expected_payment_date: totalDelay ? addDays(p.expected_payment_date, totalDelay) : p.expected_payment_date,
      _delayApplied: totalDelay,
    };
  });
  let balance = opening;
  const days = [];
  for (let d = new Date(TODAY); d <= monthEnd; d = addDays(d, 1)) {
    const dr = adjReceipts.filter(r=>fmt(r.expected_payment_date)===fmt(d)).reduce((s,r)=>s+r.amount,0);
    const dp = adjPayments.filter(p=>fmt(p.expected_payment_date)===fmt(d)).reduce((s,p)=>s+p.amount,0);
    balance += dr - dp;
    days.push({ date:fmt(d), balance:Math.round(balance), receipts:Math.round(dr), payments:Math.round(dp), net:Math.round(dr-dp) });
  }
  return days;
}

// ─── AI COMMENTARY ────────────────────────────────────────────────────────────
function generateCommentary(opening, receipts, payments, forecast, receiptDelayDays, receiptHaircut, globalPaymentDelay, paymentDelays) {
  const totalR = receipts.reduce((s,r)=>s+r.amount,0);
  const totalP = payments.reduce((s,p)=>s+p.amount,0);
  const me = forecast[forecast.length-1]?.balance ?? opening;
  const topR = [...receipts].sort((a,b)=>b.amount-a.amount)[0];
  const topP = [...payments].sort((a,b)=>b.amount-a.amount)[0];
  const uncertain = receipts.filter(r=>r.status==="Uncertain");
  const low = forecast.reduce((m,d)=>d.balance<m.balance?d:m, forecast[0]||{balance:opening,date:"—"});
  const liveCount = receipts.filter(r=>r.source!=="sample").length;
  const dataNote = liveCount>0 ? " Data sourced live." : " Using sample data.";
  const activePaymentDelays = Object.entries(paymentDelays).filter(([,v])=>v>0);
  const scenParts = [];
  if (receiptDelayDays>0) scenParts.push(`receipts delayed ${receiptDelayDays}d`);
  if (receiptHaircut>0)   scenParts.push(`${receiptHaircut}% receipt haircut`);
  if (globalPaymentDelay>0) scenParts.push(`all payments deferred ${globalPaymentDelay}d`);
  if (activePaymentDelays.length>0) scenParts.push(`${activePaymentDelays.length} individual payment${activePaymentDelays.length>1?"s":""}  deferred`);
  const scenNote = scenParts.length>0 ? ` ⚠ Scenario: ${scenParts.join(", ")}.` : "";
  return {
    insights: `Current cash stands at ${fmtGBP(opening)}.${dataNote}${scenNote} ${receipts.length} receipts (${fmtGBP(totalR)}) vs ${payments.length} payments (${fmtGBP(totalP)}) gives a forecast month-end of ${fmtGBP(me)} — a net ${me>=opening?"improvement":"reduction"} of ${fmtGBP(Math.abs(me-opening))}. Largest inflow: ${topR?.customer_name} ${fmtGBP(topR?.amount)}.`,
    risks: uncertain.length>0
      ? `${uncertain.length} receipt${uncertain.length>1?"s":""}  flagged Uncertain (${uncertain.map(r=>r.customer_name).join(", ")}, ${fmtGBP(uncertain.reduce((s,r)=>s+r.amount,0))} combined). Liquidity trough: ${low?.date} at ${fmtGBP(low?.balance)}. Largest payment: ${topP?.supplier_name} ${fmtGBP(topP?.amount)}.`
      : `No uncertain receipts. Trough: ${low?.date} at ${fmtGBP(low?.balance)}. Largest payment: ${topP?.supplier_name} ${fmtGBP(topP?.amount)}.`,
    opportunities: me>opening
      ? `Net inflow of ${fmtGBP(me-opening)} — consider sweeping surplus to a higher-yield account or reducing revolving credit facility mid-month.`
      : `Accelerate uncertain collections to restore headroom. Review deferrable discretionary supplier payments beyond month-end.`,
  };
}

async function callClaude(system, user) {
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method:"POST", headers:{"Content-Type":"application/json"},
    body: JSON.stringify({ model:"claude-sonnet-4-20250514", max_tokens:1000, system, messages:[{role:"user",content:user}] }),
  });
  const d = await r.json();
  return d.content?.map(b=>b.text||"").join("")||"";
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function exportToExcel(receipts, payments, forecast, balance, commentary, paymentDelays, globalPaymentDelay) {
  const wb = XLSX.utils.book_new();
  const summaryData = [
    ["Parcel2Go Cash Flow Model","",`Generated: ${new Date().toLocaleString("en-GB")}`],[],
    ["KEY METRICS","",""],
    ["Current Cash", balance, ""],
    ["Total Outstanding Receipts", receipts.reduce((s,r)=>s+r.amount,0), ""],
    ["Total Outstanding Payments", payments.reduce((s,p)=>s+p.amount,0), ""],
    ["Forecast Month-End Cash", forecast[forecast.length-1]?.balance||0, ""],
    [],["AI COMMENTARY","",""],
    ["Position & Forecast", commentary?.insights||"",""],
    ["Key Risks", commentary?.risks||"",""],
    ["Opportunities", commentary?.opportunities||"",""],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(summaryData); ws1["!cols"]=[{wch:28},{wch:18},{wch:60}];
  XLSX.utils.book_append_sheet(wb, ws1, "Summary");
  const receiptRows = [["Customer","Invoice","Amount (£)","Expected Date","Status","Source"],...receipts.map(r=>[r.customer_name,r.invoice_number,r.amount,fmt(r.expected_payment_date),r.status,r.source]),[],["TOTAL","",receipts.reduce((s,r)=>s+r.amount,0),"","",""]];
  const ws2 = XLSX.utils.aoa_to_sheet(receiptRows); ws2["!cols"]=[{wch:28},{wch:18},{wch:14},{wch:14},{wch:12},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws2, "Customer Receipts");
  const paymentRows = [["Supplier","Invoice","Amount (£)","Original Due Date","Delay Applied","Adjusted Date","Status","Source"],...payments.map(p=>{const d=paymentDelays[p.invoice_number]||0; const tot=d+globalPaymentDelay; return [p.supplier_name,p.invoice_number,p.amount,fmt(p.expected_payment_date),tot>0?`+${tot} days`:"None",tot>0?fmt(addDays(p.expected_payment_date,tot)):fmt(p.expected_payment_date),p.status,p.source];}),[], ["TOTAL","",payments.reduce((s,p)=>s+p.amount,0),"","","","",""]];
  const ws3 = XLSX.utils.aoa_to_sheet(paymentRows); ws3["!cols"]=[{wch:26},{wch:18},{wch:13},{wch:16},{wch:13},{wch:13},{wch:10},{wch:10}];
  XLSX.utils.book_append_sheet(wb, ws3, "Supplier Payments");
  const forecastRows = [["Date","Balance (£)","Receipts (£)","Payments (£)","Net (£)"],...forecast.map(d=>[d.date,d.balance,d.receipts,d.payments,d.net])];
  const ws4 = XLSX.utils.aoa_to_sheet(forecastRows); ws4["!cols"]=[{wch:12},{wch:16},{wch:14},{wch:14},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws4, "Daily Forecast");
  XLSX.writeFile(wb, `Parcel2Go_CashFlow_${new Date().toISOString().slice(0,10)}.xlsx`);
}
function exportToCSV(receipts, payments, forecast) {
  const lines = ["CUSTOMER RECEIPTS","Customer,Invoice,Amount,Expected Date,Status,Source",...receipts.map(r=>`${r.customer_name},${r.invoice_number},${r.amount},${fmt(r.expected_payment_date)},${r.status},${r.source}`),"","SUPPLIER PAYMENTS","Supplier,Invoice,Amount,Due Date,Status,Source",...payments.map(p=>`${p.supplier_name},${p.invoice_number},${p.amount},${fmt(p.expected_payment_date)},${p.status},${p.source}`),"","DAILY FORECAST","Date,Balance,Receipts,Payments,Net",...forecast.map(d=>`${d.date},${d.balance},${d.receipts},${d.payments},${d.net}`)];
  const a = document.createElement("a"); a.href=URL.createObjectURL(new Blob([lines.join("\n")],{type:"text/csv"}));
  a.download=`Parcel2Go_CashFlow_${new Date().toISOString().slice(0,10)}.csv`; a.click();
}

// ─── SMALL COMPONENTS ─────────────────────────────────────────────────────────
const SBADGE = { Confirmed:"#10b981", Expected:"#3b82f6", Uncertain:"#f59e0b", Due:"#ef4444", Fixed:"#8b5cf6", Partial:"#f97316" };
const SRC_META = { stripe:{color:"#818cf8",label:"Stripe"}, xero:{color:"#06b6d4",label:"Xero"}, bank:{color:"#10b981",label:"Bank"}, sample:{color:"#334155",label:"Sample"} };
function Badge({status}){ const c=SBADGE[status]||"#6b7280"; return <span style={{background:c+"22",color:c,border:`1px solid ${c}44`,borderRadius:4,padding:"2px 8px",fontSize:11,fontWeight:700,fontFamily:"monospace"}}>{status}</span>; }
function SrcBadge({source}){ const s=SRC_META[source]||SRC_META.sample; return <span style={{background:s.color+"18",color:s.color,border:`1px solid ${s.color}33`,borderRadius:3,padding:"1px 6px",fontSize:10,fontWeight:700,fontFamily:"monospace"}}>{s.label}</span>; }
function CashTip({active,payload,label}){ if(!active||!payload?.length) return null; return <div style={{background:"#0a0f1e",border:"1px solid #1e293b",borderRadius:8,padding:"10px 14px",fontSize:12}}><div style={{color:"#64748b",marginBottom:6}}>{label}</div>{payload.map(p=><div key={p.name} style={{color:p.color,fontWeight:700}}>{p.name}: {fmtGBP(p.value)}</div>)}</div>; }

function Slider({ label, value, min, max, step=1, unit="", color="#3b82f6", compact=false, onChange }) {
  return (
    <div style={{marginBottom: compact ? 0 : 18}}>
      {!compact && <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
        <span style={{fontSize:12,color:"#94a3b8",fontWeight:600}}>{label}</span>
        <span style={{fontSize:16,fontWeight:800,color:value===0?"#334155":color,fontVariantNumeric:"tabular-nums",minWidth:60,textAlign:"right"}}>{value===0?"No change":`+${value}${unit}`}</span>
      </div>}
      <div style={{position:"relative",height:compact?4:6,background:"#1e293b",borderRadius:3}}>
        <div style={{position:"absolute",left:0,top:0,height:"100%",width:`${((value-min)/(max-min))*100}%`,background:value===0?"#1e293b":`linear-gradient(90deg,${color}88,${color})`,borderRadius:3,transition:"width 0.12s"}}/>
        <input type="range" min={min} max={max} step={step} value={value} onChange={e=>onChange(Number(e.target.value))}
          style={{position:"absolute",inset:0,width:"100%",opacity:0,cursor:"pointer",height:"100%",margin:0}}/>
      </div>
      {!compact && <div style={{display:"flex",justifyContent:"space-between",marginTop:4,fontSize:10,color:"#1e3a5f"}}><span>{min}{unit}</span><span>{max}{unit}</span></div>}
    </div>
  );
}

// ─── EXPANDABLE PAYMENT ROW ───────────────────────────────────────────────────
function PaymentRow({ payment, index, individualDelay, globalPaymentDelay, baseMonthEnd, onDelayChange, payments, balance, receipts, receiptDelayDays, receiptHaircut, paymentDelays }) {
  const [expanded, setExpanded] = useState(false);

  const indivDelay = individualDelay || 0;
  const totalDelay = indivDelay + globalPaymentDelay;
  const hasDelay = totalDelay > 0;
  const hasIndivDelay = indivDelay > 0;

  // Calculate impact of THIS payment's individual delay on month-end (holding everything else constant)
  const scenarioDelays = { ...paymentDelays, [payment.invoice_number]: indivDelay };
  const scenForecast = buildForecast(receipts, payments, balance, receiptDelayDays, receiptHaircut, scenarioDelays, globalPaymentDelay);
  const scenME = scenForecast[scenForecast.length-1]?.balance ?? balance;
  const impact = scenME - baseMonthEnd;

  // Adjusted due date
  const adjDate = totalDelay > 0 ? addDays(payment.expected_payment_date, totalDelay) : null;

  return (
    <>
      {/* MAIN ROW */}
      <tr
        onClick={() => setExpanded(e => !e)}
        style={{
          background: expanded ? "#0d1f38" : index % 2 ? "transparent" : "#050b18",
          cursor: "pointer",
          borderLeft: expanded ? "3px solid #3b82f6" : "3px solid transparent",
          transition: "background 0.15s",
        }}
      >
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020",fontSize:13}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:10,color: expanded ? "#3b82f6" : "#334155",transition:"transform 0.15s",display:"inline-block",transform: expanded?"rotate(90deg)":"rotate(0deg)"}}>▶</span>
            <span style={{fontWeight:600,color:"#e2e8f0"}}>{payment.supplier_name}</span>
          </div>
        </td>
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020",fontFamily:"monospace",fontSize:11,color:"#334155"}}>{payment.invoice_number}</td>
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020",color:"#ef4444",fontWeight:700,fontVariantNumeric:"tabular-nums",fontSize:13}}>{fmtGBP(payment.amount)}</td>
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020",fontSize:13}}>
          <div style={{display:"flex",flexDirection:"column",gap:2}}>
            <span style={{color: hasDelay ? "#94a3b8" : "#94a3b8", textDecoration: hasDelay ? "line-through" : "none", fontSize: hasDelay ? 11 : 13}}>
              {fmt(payment.expected_payment_date)}
            </span>
            {hasDelay && <span style={{color:"#f59e0b",fontSize:12,fontWeight:600}}>{fmt(adjDate)} <span style={{fontSize:10,opacity:0.7}}>(+{totalDelay}d)</span></span>}
          </div>
        </td>
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020"}}><Badge status={payment.status}/></td>
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020"}}><SrcBadge source={payment.source}/></td>
        <td style={{padding:"10px 12px",borderBottom: expanded ? "none" : "1px solid #0a1020",textAlign:"right"}}>
          {hasIndivDelay
            ? <span style={{fontSize:11,color:"#10b981",fontWeight:700,background:"#10b98115",border:"1px solid #10b98133",borderRadius:4,padding:"2px 7px"}}>+{fmtGBP(Math.abs(impact))}</span>
            : <span style={{fontSize:11,color:"#334155"}}>—</span>
          }
        </td>
      </tr>

      {/* EXPANDED PANEL */}
      {expanded && (
        <tr style={{background:"#0a1828"}}>
          <td colSpan={7} style={{padding:"0 0 2px 0",borderBottom:"1px solid #1a2740"}}>
            <div style={{
              padding:"16px 20px 18px 36px",
              borderTop:"1px solid #1e3a5f33",
              animation:"expandRow 0.18s ease",
            }}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:20,alignItems:"start"}}>

                {/* Slider */}
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:10}}>
                    Defer this payment
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
                    <div style={{flex:1}}>
                      <Slider value={indivDelay} min={0} max={30} unit=" days" color="#f59e0b" compact onChange={v => onDelayChange(payment.invoice_number, v)}/>
                    </div>
                    <span style={{
                      fontSize:18,fontWeight:800,
                      color: indivDelay===0 ? "#334155" : "#f59e0b",
                      fontVariantNumeric:"tabular-nums",
                      minWidth:70,textAlign:"right",
                      transition:"color 0.15s"
                    }}>
                      {indivDelay===0 ? "No delay" : `+${indivDelay} days`}
                    </span>
                  </div>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    {[0,7,14,30].map(d=>(
                      <button key={d} onClick={e=>{e.stopPropagation();onDelayChange(payment.invoice_number,d);}}
                        style={{background:indivDelay===d?"#f59e0b22":"#1e293b",color:indivDelay===d?"#f59e0b":"#475569",border:`1px solid ${indivDelay===d?"#f59e0b44":"#334155"}`,borderRadius:5,padding:"4px 10px",fontSize:11,fontWeight:600,cursor:"pointer"}}>
                        {d===0?"Reset":`+${d}d`}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Date impact */}
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:10}}>Date Impact</div>
                  <div style={{display:"flex",flexDirection:"column",gap:6}}>
                    <div style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",background:"#080f1e",borderRadius:6,border:"1px solid #1a2740"}}>
                      <span style={{fontSize:12,color:"#475569"}}>Original due</span>
                      <span style={{fontSize:12,color:"#94a3b8",fontWeight:600}}>{fmt(payment.expected_payment_date)}</span>
                    </div>
                    {globalPaymentDelay>0 && (
                      <div style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",background:"#080f1e",borderRadius:6,border:"1px solid #1a2740"}}>
                        <span style={{fontSize:12,color:"#475569"}}>Global defer (+{globalPaymentDelay}d)</span>
                        <span style={{fontSize:12,color:"#f59e0b",fontWeight:600}}>{fmt(addDays(payment.expected_payment_date, globalPaymentDelay))}</span>
                      </div>
                    )}
                    <div style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",background: indivDelay>0||globalPaymentDelay>0 ? "#0d1f38" : "#080f1e",borderRadius:6,border:`1px solid ${indivDelay>0||globalPaymentDelay>0?"#f59e0b44":"#1a2740"}`}}>
                      <span style={{fontSize:12,color: indivDelay>0||globalPaymentDelay>0 ? "#f59e0b" : "#475569",fontWeight:600}}>Effective date</span>
                      <span style={{fontSize:12,color: indivDelay>0||globalPaymentDelay>0 ? "#f59e0b" : "#94a3b8",fontWeight:700}}>{fmt(addDays(payment.expected_payment_date, totalDelay))}</span>
                    </div>
                  </div>
                </div>

                {/* Cash impact */}
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:10}}>Cash Impact</div>
                  <div style={{display:"flex",flexDirection:"column",gap:6}}>
                    <div style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",background:"#080f1e",borderRadius:6,border:"1px solid #1a2740"}}>
                      <span style={{fontSize:12,color:"#475569"}}>Payment amount</span>
                      <span style={{fontSize:12,color:"#ef4444",fontWeight:700}}>{fmtGBP(payment.amount)}</span>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",background:"#080f1e",borderRadius:6,border:"1px solid #1a2740"}}>
                      <span style={{fontSize:12,color:"#475569"}}>Base month-end</span>
                      <span style={{fontSize:12,color:"#3b82f6",fontWeight:700}}>{fmtGBP(baseMonthEnd)}</span>
                    </div>
                    <div style={{display:"flex",justifyContent:"space-between",padding:"6px 10px",background: hasIndivDelay ? "#0d1f38" : "#080f1e",borderRadius:6,border:`1px solid ${hasIndivDelay?"#10b98144":"#1a2740"}`}}>
                      <span style={{fontSize:12,color: hasIndivDelay?"#10b981":"#475569",fontWeight: hasIndivDelay?600:400}}>Month-end impact</span>
                      <span style={{fontSize:13,fontWeight:800,color: hasIndivDelay ? "#10b981" : "#334155"}}>
                        {hasIndivDelay ? `+${fmtGBP(impact)}` : "No delay set"}
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </td>
        </tr>
      )}
    </>
  );
}

// ═════════════════════════════════════════════════════════════════════════════
// MAIN APP
// ═════════════════════════════════════════════════════════════════════════════
export default function App() {
  const [keys, setKeys]           = useState({ stripe:"", xeroToken:"", xeroTenant:"", truelayer:"" });
  const [savedKeys, setSavedKeys] = useState({ stripe:"", xeroToken:"", xeroTenant:"", truelayer:"" });
  const [showKeys, setShowKeys]   = useState(false);

  const [receipts, setReceipts]   = useState(SAMPLE_RECEIPTS);
  const [payments, setPayments]   = useState(SAMPLE_PAYMENTS);
  const [balance, setBalance]     = useState(SAMPLE_BALANCE);
  const [dataMode, setDataMode]   = useState("sample");
  const [syncLog, setSyncLog]     = useState([]);

  // Scenario: receipts
  const [receiptDelayDays, setReceiptDelayDays] = useState(0);
  const [receiptHaircut, setReceiptHaircut]     = useState(0);
  // Scenario: payments — global + per-invoice
  const [globalPaymentDelay, setGlobalPaymentDelay]   = useState(0);
  const [paymentDelays, setPaymentDelays]               = useState({}); // { invoice_number: days }

  const [forecast, setForecast]         = useState([]);
  const [scenForecast, setScenForecast] = useState([]);
  const [commentary, setCommentary]     = useState(null);

  const [tab, setTab]             = useState("overview");
  const [syncing, setSyncing]     = useState(false);
  const [syncMsg, setSyncMsg]     = useState("");
  const [lastSync, setLastSync]   = useState(null);
  const [cfoPanel, setCfoPanel]   = useState(null);
  const [cfoLoading, setCfoLoading] = useState(false);
  const [chat, setChat]           = useState([{ role:"assistant", text:"Hello. I'm your Cash Flow Assistant. Try: 'What if we delay the HMRC payment by 10 days?' or 'What are our largest upcoming payments?'" }]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const chatRef = useRef(null);

  const anyScenarioActive = receiptDelayDays>0 || receiptHaircut>0 || globalPaymentDelay>0 || Object.values(paymentDelays).some(v=>v>0);

  // Recalculate forecasts on any data/scenario change
  useEffect(()=>{
    const base = buildForecast(receipts, payments, balance, 0, 0, {}, 0);
    const scen = anyScenarioActive ? buildForecast(receipts, payments, balance, receiptDelayDays, receiptHaircut, paymentDelays, globalPaymentDelay) : null;
    setForecast(base);
    setScenForecast(scen || []);
    setCommentary(generateCommentary(balance, receipts, payments, scen||base, receiptDelayDays, receiptHaircut, globalPaymentDelay, paymentDelays));
  }, [receipts, payments, balance, receiptDelayDays, receiptHaircut, globalPaymentDelay, paymentDelays]);

  useEffect(()=>{ chatRef.current?.scrollIntoView({behavior:"smooth"}); }, [chat]);

  const handleIndividualPaymentDelay = (invoiceNumber, days) => {
    setPaymentDelays(prev => {
      const next = {...prev};
      if (days===0) delete next[invoiceNumber];
      else next[invoiceNumber] = days;
      return next;
    });
  };

  const resetAllPaymentDelays = () => { setPaymentDelays({}); setGlobalPaymentDelay(0); };

  const handleSync = async () => {
    setSyncing(true); setSyncLog([]);
    const hasLive = savedKeys.stripe||savedKeys.xeroToken||savedKeys.truelayer;
    if (!hasLive) {
      setSyncMsg("Refreshing sample data…");
      await new Promise(r=>setTimeout(r,900));
      setReceipts(SAMPLE_RECEIPTS.map(r=>({...r,amount:Math.round(r.amount*(1+(Math.random()-0.5)*0.03))})));
      setPayments(SAMPLE_PAYMENTS.map(p=>({...p,amount:Math.round(p.amount*(1+(Math.random()-0.5)*0.01))})));
      setBalance(Math.round(SAMPLE_BALANCE*(1+(Math.random()-0.5)*0.02)));
      setPaymentDelays({}); // reset per-payment delays on fresh sync
      setDataMode("sample"); setSyncMsg("✓ Sample data refreshed");
    } else {
      const steps=[savedKeys.stripe&&"Fetching Stripe…",savedKeys.xeroToken&&"Fetching Xero…",savedKeys.truelayer&&"Fetching bank feed…"].filter(Boolean);
      for(const s of steps){setSyncMsg(s);await new Promise(r=>setTimeout(r,400));}
      try {
        const res=await runLiveSync(savedKeys);
        setReceipts(res.receipts); setPayments(res.payments); setBalance(res.balance); setSyncLog(res.log);
        setPaymentDelays({});
        const live=res.log.filter(l=>l.status==="ok").map(l=>l.source);
        setDataMode(live.length?"live":"sample");
        setSyncMsg(res.errors.length?`⚠ ${res.errors.map(e=>`${e.source}: ${e.message}`).join(" | ")}`:`✓ Synced — ${live.join(", ")}`);
      } catch(e){setSyncMsg(`Error: ${e.message}`);}
    }
    setLastSync(new Date()); setSyncing(false);
  };

  const generateCFO = async () => {
    setCfoLoading(true);
    const me=forecast[forecast.length-1]?.balance??balance;
    const low=forecast.reduce((m,d)=>d.balance<m.balance?d:m,forecast[0]||{balance,date:"—"});
    const activePayDeferrals=Object.entries(paymentDelays).filter(([,v])=>v>0).map(([inv,d])=>{const p=payments.find(x=>x.invoice_number===inv);return p?`${p.supplier_name} deferred ${d}d`:inv;});
    const ctx=`Parcel2Go ${new Date().toLocaleDateString("en-GB")}. Cash: ${fmtGBP(balance)}. Receipts: ${fmtGBP(receipts.reduce((s,r)=>s+r.amount,0))} (${receipts.length}). Payments: ${fmtGBP(payments.reduce((s,p)=>s+p.amount,0))} (${payments.length}). Month-end: ${fmtGBP(me)}. Trough: ${fmtGBP(low?.balance)} on ${low?.date}. Uncertain receipts: ${receipts.filter(r=>r.status==="Uncertain").map(r=>`${r.customer_name} ${fmtGBP(r.amount)}`).join(", ")||"none"}. Active scenarios: receipt delay ${receiptDelayDays}d, haircut ${receiptHaircut}%, global payment deferral ${globalPaymentDelay}d${activePayDeferrals.length?`, individual deferrals: ${activePayDeferrals.join(", ")}`:""}.`;
    try {
      const raw=await callClaude("You are a commercial finance manager writing a daily CFO briefing for Parcel2Go UK. Return ONLY valid JSON: {keyInsights:string, risks:string, recommendedActions:string[3]}. No markdown.",ctx);
      setCfoPanel(JSON.parse(raw.replace(/```json|```/g,"").trim()));
    } catch {
      const me2=forecast[forecast.length-1]?.balance??balance;
      setCfoPanel({keyInsights:`Opening cash of ${fmtGBP(balance)} is forecast to ${me2>balance?"improve to":"fall to"} ${fmtGBP(me2)} by month-end.${anyScenarioActive?" Active payment deferral scenarios are improving the intra-month liquidity profile.":""}`,risks:`Intra-month trough of ${fmtGBP(low?.balance)} on ${low?.date}. ${receipts.filter(r=>r.status==="Uncertain").length>0?`Uncertain receipts from ${receipts.filter(r=>r.status==="Uncertain").map(r=>r.customer_name).join(", ")} represent downside risk.`:"No uncertain receipts currently."}`,recommendedActions:["Chase Uncertain-status customers for payment confirmation.","Confirm banking facility headroom ahead of the intra-month liquidity trough.","Review deferred payment scenarios against supplier terms to avoid penalty interest."]});
    }
    setCfoLoading(false);
  };

  const handleChat = async () => {
    if (!chatInput.trim()) return;
    const q=chatInput.trim(); setChat(p=>[...p,{role:"user",text:q}]); setChatInput(""); setChatLoading(true);
    const me=forecast[forecast.length-1]?.balance??balance;
    const scenMe=scenForecast[scenForecast.length-1]?.balance??me;
    const ctx=`Data (${dataMode}): cash ${fmtGBP(balance)}, receipts ${fmtGBP(receipts.reduce((s,r)=>s+r.amount,0))}, payments ${fmtGBP(payments.reduce((s,p)=>s+p.amount,0))}, base month-end ${fmtGBP(me)}, scenario month-end ${fmtGBP(scenMe)}. Scenarios active: receipt delay ${receiptDelayDays}d, haircut ${receiptHaircut}%, global payment deferral ${globalPaymentDelay}d, ${Object.keys(paymentDelays).length} individual payment deferrals. Top payments: ${[...payments].sort((a,b)=>b.amount-a.amount).slice(0,5).map(p=>`${p.supplier_name} ${fmtGBP(p.amount)} due ${fmt(p.expected_payment_date)}`).join("; ")}.`;
    try {
      const ans=await callClaude("You are a cash flow analyst for Parcel2Go UK. Answer concisely in 2–3 sentences with specific £ figures.",`${ctx}\n\nQ: ${q}`);
      setChat(p=>[...p,{role:"assistant",text:ans}]);
    } catch {
      const ql=q.toLowerCase();
      let ans=`Base month-end: ${fmtGBP(me)}. ${anyScenarioActive?`With active scenarios: ${fmtGBP(scenMe)}.`:""}`;
      if(ql.includes("largest")||ql.includes("biggest")) ans=`Top 3 payments: ${[...payments].sort((a,b)=>b.amount-a.amount).slice(0,3).map(p=>`${p.supplier_name} ${fmtGBP(p.amount)} (${fmt(p.expected_payment_date)})`).join("; ")}.`;
      if(ql.includes("delay")&&(ql.includes("hmrc")||ql.includes("payroll")||ql.includes("royal"))) {
        const match=payments.find(p=>ql.includes(p.supplier_name.toLowerCase().split(" ")[0].toLowerCase()));
        if(match){const sc=buildForecast(receipts,payments,balance,0,0,{[match.invoice_number]:10},0);const sm=sc[sc.length-1]?.balance??me;ans=`Deferring ${match.supplier_name} by 10 days moves month-end to ${fmtGBP(sm)} — an improvement of ${fmtGBP(sm-me)}. Use the Payments tab to set this interactively.`;}
      }
      setChat(p=>[...p,{role:"assistant",text:ans}]);
    }
    setChatLoading(false);
  };

  // Derived
  const totalR=receipts.reduce((s,r)=>s+r.amount,0);
  const totalP=payments.reduce((s,p)=>s+p.amount,0);
  const meBase=forecast[forecast.length-1]?.balance??balance;
  const meScen=scenForecast[scenForecast.length-1]?.balance??meBase;
  const scenImpact=meScen-meBase;
  const activeIndivDelays=Object.values(paymentDelays).filter(v=>v>0).length;
  const dualForecast=forecast.map((d,i)=>({...d,scenarioBalance:scenForecast[i]?.balance??null}));
  const weeklyData=[];
  for(let i=0;i<forecast.length;i+=7){const sl=forecast.slice(i,i+7);weeklyData.push({week:`Wk${Math.floor(i/7)+1}`,receipts:sl.reduce((s,d)=>s+d.receipts,0),payments:sl.reduce((s,d)=>s+d.payments,0)});}
  const hasLiveKeys=savedKeys.stripe||savedKeys.xeroToken||savedKeys.truelayer;

  const connStatus=[
    {id:"stripe", label:"Stripe",    icon:"⚡", color:"#818cf8", connected:!!savedKeys.stripe,     desc:"Payment intents + payouts",  keyFields:["stripe"],               keyLabels:["Restricted API Key"],      keyPlaceholders:["sk_live_..."]},
    {id:"xero",   label:"Xero",      icon:"☁",  color:"#06b6d4", connected:!!(savedKeys.xeroToken&&savedKeys.xeroTenant), desc:"Sales invoices + bills", keyFields:["xeroToken","xeroTenant"],keyLabels:["Access Token","Tenant ID"],keyPlaceholders:["eyJ...","uuid"]},
    {id:"bank",   label:"Bank Feed", icon:"🏦", color:"#10b981", connected:!!savedKeys.truelayer,  desc:"Live balance via TrueLayer", keyFields:["truelayer"],             keyLabels:["TrueLayer Access Token"],  keyPlaceholders:["eyJ..."]},
  ];

  const C = {
    app:   {background:"#03070f",minHeight:"100vh",color:"#e2e8f0",fontFamily:"'DM Sans',sans-serif"},
    hdr:   {background:"linear-gradient(90deg,#070e1c,#0d1f3c,#070e1c)",borderBottom:"1px solid #1e293b",padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",height:64},
    main:  {padding:"22px 28px"},
    kpiRow:{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:13,marginBottom:22},
    kpi:   (c)=>({background:"#080f1e",border:`1px solid ${c}2a`,borderRadius:12,padding:"15px 18px",position:"relative",overflow:"hidden"}),
    glow:  (c)=>({position:"absolute",top:-20,right:-20,width:80,height:80,background:`radial-gradient(circle,${c}28,transparent 70%)`,pointerEvents:"none"}),
    card:  {background:"#080f1e",border:"1px solid #1a2740",borderRadius:12,padding:"18px 20px"},
    cardT: {fontSize:11,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:13},
    g2:    {display:"grid",gridTemplateColumns:"1fr 1fr",gap:15,marginBottom:15},
    tabs:  {display:"flex",gap:3,background:"#080f1e",border:"1px solid #1a2740",borderRadius:9,padding:3,width:"fit-content",marginBottom:18},
    tab:   (a)=>({padding:"6px 14px",borderRadius:7,fontSize:12,fontWeight:600,cursor:"pointer",background:a?"#1e3a5f":"transparent",color:a?"#60a5fa":"#475569",border:"none",transition:"all 0.15s"}),
    th:    {padding:"8px 12px",fontSize:10,color:"#334155",textTransform:"uppercase",letterSpacing:"0.1em",borderBottom:"1px solid #1a2740",textAlign:"left",fontWeight:700},
    pill:  (c,a)=>({background:a?c+"22":"transparent",color:a?c:"#475569",border:`1px solid ${a?c+"55":"transparent"}`,borderRadius:6,padding:"3px 10px",fontSize:10,fontWeight:700,letterSpacing:"0.08em"}),
    cbox:  (c)=>({background:c+"0d",border:`1px solid ${c}30`,borderRadius:9,padding:"11px 15px",marginBottom:10}),
    clabel:(c)=>({fontSize:10,fontWeight:800,color:c,textTransform:"uppercase",letterSpacing:"0.14em",marginBottom:6}),
    ctext: {fontSize:13,color:"#cbd5e1",lineHeight:1.75},
    inp:   {width:"100%",background:"#0d1a2e",border:"1px solid #1e3a5f",borderRadius:7,padding:"9px 12px",color:"#e2e8f0",fontSize:13,outline:"none",fontFamily:"monospace",boxSizing:"border-box"},
    modal: {position:"fixed",inset:0,background:"rgba(0,0,0,0.78)",backdropFilter:"blur(4px)",zIndex:100,display:"flex",alignItems:"center",justifyContent:"center"},
    mbox:  {background:"#080f1e",border:"1px solid #1e3a5f",borderRadius:16,padding:28,width:540,maxHeight:"88vh",overflowY:"auto"},
  };

  const TABS = ["overview","scenario","payments","receipts","ai cfo","integrations"];

  return (
    <div style={C.app}>
      <style>{`*{box-sizing:border-box}::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:#030710}::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:3px}@keyframes spin{to{transform:rotate(360deg)}}@keyframes fade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}@keyframes expandRow{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:translateY(0)}}input[type=range]{-webkit-appearance:none;appearance:none}input[type=range]::-webkit-slider-thumb{-webkit-appearance:none;width:16px;height:16px;border-radius:50%;background:#f59e0b;cursor:pointer;border:2px solid #1e3a5f}`}</style>

      {/* KEY MODAL */}
      {showKeys&&(
        <div style={C.modal} onClick={e=>e.target===e.currentTarget&&setShowKeys(false)}>
          <div style={C.mbox}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
              <div><div style={{fontSize:16,fontWeight:800,color:"#f1f5f9",marginBottom:3}}>Connect Live Data</div><div style={{fontSize:12,color:"#475569"}}>Keys stored in browser memory only</div></div>
              <button onClick={()=>setShowKeys(false)} style={{background:"none",border:"none",color:"#475569",fontSize:20,cursor:"pointer"}}>✕</button>
            </div>
            {connStatus.map(conn=>(
              <div key={conn.id} style={{marginBottom:24}}>
                <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:11}}>
                  <span style={{fontSize:20}}>{conn.icon}</span>
                  <div style={{flex:1}}><div style={{fontWeight:700,color:"#f1f5f9",fontSize:13}}>{conn.label}</div><div style={{fontSize:11,color:"#475569"}}>{conn.desc}</div></div>
                  <span style={C.pill(conn.color,conn.connected)}>{conn.connected?"● CONNECTED":"○ NOT SET"}</span>
                </div>
                {conn.keyFields.map((field,fi)=>(
                  <div key={field} style={{marginBottom:6}}>
                    <div style={{fontSize:11,color:"#475569",marginBottom:3,fontWeight:600}}>{conn.keyLabels[fi]}</div>
                    <input style={C.inp} type="password" value={keys[field]} onChange={e=>setKeys(p=>({...p,[field]:e.target.value}))} placeholder={conn.keyPlaceholders[fi]}/>
                  </div>
                ))}
                <div style={{background:"#050b18",borderRadius:6,padding:"8px 12px",fontSize:10,color:"#475569",marginTop:6,fontFamily:"monospace",lineHeight:1.7}}>
                  {conn.id==="stripe"&&<><div style={{color:"#818cf8",fontWeight:700,marginBottom:2}}>GET YOUR KEY → dashboard.stripe.com → Developers → API keys → Restricted Key (read: PaymentIntents, Payouts, Balance)</div></>}
                  {conn.id==="xero"&&<><div style={{color:"#06b6d4",fontWeight:700,marginBottom:2}}>GET YOUR TOKEN → developer.xero.com → New App → OAuth 2.0 flow → Access Token + Tenant ID</div><div style={{color:"#f59e0b"}}>⚠ Tokens expire 30 mins — use backend proxy in production</div></>}
                  {conn.id==="bank"&&<><div style={{color:"#10b981",fontWeight:700,marginBottom:2}}>GET TOKEN → console.truelayer.com → connect your UK bank → Access Token (Barclays, HSBC, Lloyds, Monzo, Starling…)</div></>}
                </div>
              </div>
            ))}
            <button onClick={()=>{setSavedKeys({...keys});setShowKeys(false);}} style={{background:"linear-gradient(135deg,#2563eb,#1d4ed8)",color:"#fff",border:"none",borderRadius:8,padding:"10px 22px",fontSize:13,fontWeight:700,cursor:"pointer",width:"100%",marginTop:4}}>Save & Activate Connectors</button>
          </div>
        </div>
      )}

      {/* HEADER */}
      <header style={C.hdr}>
        <div style={{display:"flex",alignItems:"center",gap:12}}>
          <div style={{width:34,height:34,background:"linear-gradient(135deg,#3b82f6,#1d4ed8)",borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:12,color:"#fff"}}>P2G</div>
          <div>
            <div style={{fontSize:15,fontWeight:800,color:"#f1f5f9",letterSpacing:"-0.03em"}}>Parcel2Go</div>
            <div style={{fontSize:10,color:"#334155",textTransform:"uppercase",letterSpacing:"0.1em"}}>Cash Flow Intelligence</div>
          </div>
          <span style={{...C.pill(dataMode==="live"?"#10b981":"#475569",true),marginLeft:10}}>{dataMode==="live"?"● LIVE DATA":"○ SAMPLE DATA"}</span>
          {anyScenarioActive&&<span style={{...C.pill("#f59e0b",true)}}>⚠ SCENARIO{activeIndivDelays>0?` (+${activeIndivDelays} individual)`:""}</span>}
        </div>
        <div style={{display:"flex",alignItems:"center",gap:9}}>
          {lastSync&&<div style={{fontSize:10,color:"#334155",textAlign:"right"}}><div>Last sync</div><div style={{color:"#475569"}}>{lastSync.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit",second:"2-digit"})}</div></div>}
          {syncMsg&&!syncing&&<div style={{fontSize:10,color:syncMsg.startsWith("✓")?"#10b981":"#f59e0b",maxWidth:160,textAlign:"right"}}>{syncMsg}</div>}
          <button onClick={()=>exportToCSV(receipts,payments,forecast)} style={{background:"transparent",color:"#475569",border:"1px solid #1e293b",borderRadius:6,padding:"6px 11px",fontSize:11,fontWeight:600,cursor:"pointer"}}>↓ CSV</button>
          <button onClick={()=>exportToExcel(receipts,payments,forecast,balance,commentary,paymentDelays,globalPaymentDelay)} style={{background:"#10b98118",color:"#10b981",border:"1px solid #10b98144",borderRadius:6,padding:"6px 11px",fontSize:11,fontWeight:600,cursor:"pointer"}}>↓ Excel</button>
          <button onClick={()=>setShowKeys(true)} style={{background:"transparent",color:"#3b82f6",border:"1px solid #1e3a5f",borderRadius:6,padding:"6px 11px",fontSize:11,fontWeight:600,cursor:"pointer"}}>{hasLiveKeys?"⚙ Connections":"🔗 Connect Live"}</button>
          <button onClick={handleSync} disabled={syncing} style={{background:syncing?"#1e293b":"linear-gradient(135deg,#2563eb,#1d4ed8)",color:syncing?"#475569":"#fff",border:"none",borderRadius:7,padding:"7px 14px",fontSize:12,fontWeight:600,cursor:syncing?"default":"pointer",display:"flex",alignItems:"center",gap:6}}>
            <span style={{display:"inline-block",animation:syncing?"spin 0.8s linear infinite":"none",fontSize:14}}>⟳</span>
            {syncing?syncMsg||"Syncing…":"Run Daily Update"}
          </button>
        </div>
      </header>

      <main style={C.main}>
        {/* KPI CARDS */}
        <div style={C.kpiRow}>
          {[
            {label:"Current Cash",         val:fmtGBP(balance),  sub:`As at ${new Date().toLocaleDateString("en-GB")}`,                                               c:"#10b981"},
            {label:"Forecast Month-End",   val:fmtGBP(meBase),   sub:meBase>=balance?`▲ ${fmtGBP(meBase-balance)} vs today`:`▼ ${fmtGBP(balance-meBase)} vs today`,   c:meBase>=balance?"#3b82f6":"#f59e0b"},
            {label:"Outstanding Receipts", val:fmtGBP(totalR),   sub:`${receipts.length} invoices · ${receipts.filter(r=>r.status==="Uncertain").length} uncertain`,  c:"#8b5cf6"},
            {label:"Outstanding Payments", val:fmtGBP(totalP),   sub:`${payments.length} invoices · ${payments.filter(p=>p.status==="Fixed").length} fixed`,          c:"#ef4444"},
          ].map(k=>(
            <div key={k.label} style={C.kpi(k.c)}>
              <div style={C.glow(k.c)}/>
              <div style={{fontSize:10,color:"#334155",textTransform:"uppercase",letterSpacing:"0.12em",marginBottom:6,fontWeight:700}}>{k.label}</div>
              <div style={{fontSize:24,fontWeight:800,color:k.c,letterSpacing:"-0.03em",fontVariantNumeric:"tabular-nums"}}>{k.val}</div>
              <div style={{fontSize:11,color:"#334155",marginTop:3}}>{k.sub}</div>
            </div>
          ))}
        </div>

        {/* TABS */}
        <div style={C.tabs}>
          {TABS.map(t=><button key={t} style={C.tab(tab===t)} onClick={()=>setTab(t)}>{t.toUpperCase()}</button>)}
        </div>

        {/* ══ OVERVIEW ══════════════════════════════════════════════════════ */}
        {tab==="overview"&&(
          <div style={{animation:"fade 0.25s ease"}}>
            <div style={{...C.card,marginBottom:14}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
                <div style={C.cardT}>Daily Cash Balance Forecast</div>
                {anyScenarioActive&&<span style={{fontSize:11,color:"#f59e0b",background:"#f59e0b0d",border:"1px solid #f59e0b33",borderRadius:5,padding:"3px 10px"}}>Scenario overlay active</span>}
              </div>
              <ResponsiveContainer width="100%" height={210}>
                <AreaChart data={dualForecast} margin={{top:4,right:18,bottom:0,left:6}}>
                  <defs>
                    <linearGradient id="cg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#3b82f6" stopOpacity={0.28}/><stop offset="95%" stopColor="#3b82f6" stopOpacity={0.01}/></linearGradient>
                    <linearGradient id="sg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor="#f59e0b" stopOpacity={0.18}/><stop offset="95%" stopColor="#f59e0b" stopOpacity={0.01}/></linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#111827"/>
                  <XAxis dataKey="date" tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false}/>
                  <YAxis tickFormatter={v=>`£${(v/1000).toFixed(0)}k`} tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false}/>
                  <Tooltip content={<CashTip/>}/>
                  <ReferenceLine y={balance} stroke="#10b98144" strokeDasharray="4 4"/>
                  <Area type="monotone" dataKey="balance" stroke="#3b82f6" strokeWidth={2.5} fill="url(#cg)" name="Base Forecast"/>
                  {anyScenarioActive&&<Area type="monotone" dataKey="scenarioBalance" stroke="#f59e0b" strokeWidth={2} fill="url(#sg)" strokeDasharray="5 3" name="Scenario"/>}
                </AreaChart>
              </ResponsiveContainer>
            </div>
            <div style={C.g2}>
              <div style={C.card}>
                <div style={C.cardT}>Weekly Receipts vs Payments</div>
                <ResponsiveContainer width="100%" height={185}>
                  <BarChart data={weeklyData} barGap={4}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#111827"/>
                    <XAxis dataKey="week" tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false}/>
                    <YAxis tickFormatter={v=>`£${(v/1000).toFixed(0)}k`} tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false}/>
                    <Tooltip content={<CashTip/>}/><Legend wrapperStyle={{fontSize:11,color:"#475569"}}/>
                    <Bar dataKey="receipts" fill="#3b82f6" radius={[4,4,0,0]} name="Receipts"/>
                    <Bar dataKey="payments" fill="#ef444455" radius={[4,4,0,0]} name="Payments"/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
              <div style={C.card}>
                <div style={C.cardT}>AI Management Summary</div>
                {commentary&&<>
                  <div style={C.cbox("#3b82f6")}><div style={C.clabel("#3b82f6")}>Position & Forecast</div><div style={C.ctext}>{commentary.insights}</div></div>
                  <div style={C.cbox("#f59e0b")}><div style={C.clabel("#f59e0b")}>Key Risks</div><div style={C.ctext}>{commentary.risks}</div></div>
                  <div style={C.cbox("#10b981")}><div style={C.clabel("#10b981")}>Opportunities</div><div style={C.ctext}>{commentary.opportunities}</div></div>
                </>}
              </div>
            </div>
            <div style={C.card}>
              <div style={C.cardT}>💬 Cash Flow Assistant{dataMode==="live"&&<span style={{color:"#10b981",fontSize:10,marginLeft:6}}>· Live data active</span>}</div>
              <div style={{display:"flex",flexDirection:"column",height:340}}>
                <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:8,paddingRight:4}}>
                  {chat.map((m,i)=><div key={i} style={{alignSelf:m.role==="user"?"flex-end":"flex-start",maxWidth:"84%",background:m.role==="user"?"#1e3a5f":"#111827",border:`1px solid ${m.role==="user"?"#2563eb44":"#1f2937"}`,borderRadius:m.role==="user"?"12px 12px 3px 12px":"12px 12px 12px 3px",padding:"9px 13px",fontSize:13,color:"#e2e8f0",lineHeight:1.65}}>{m.text}</div>)}
                  {chatLoading&&<div style={{alignSelf:"flex-start",background:"#111827",border:"1px solid #1f2937",borderRadius:"12px 12px 12px 3px",padding:"9px 13px",fontSize:13,color:"#475569"}}>Analysing…</div>}
                  <div ref={chatRef}/>
                </div>
                <div style={{display:"flex",gap:8,marginTop:10}}>
                  <input style={{flex:1,background:"#111827",border:"1px solid #1f2937",borderRadius:8,padding:"9px 13px",color:"#e2e8f0",fontSize:13,outline:"none"}} value={chatInput} onChange={e=>setChatInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleChat()} placeholder="Ask about cash, payments, or delay scenarios…"/>
                  <button onClick={handleChat} style={{background:"#2563eb",color:"#fff",border:"none",borderRadius:8,padding:"9px 15px",fontSize:13,fontWeight:700,cursor:"pointer"}}>Send</button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ══ SCENARIO ══════════════════════════════════════════════════════ */}
        {tab==="scenario"&&(
          <div style={{animation:"fade 0.25s ease"}}>
            <div style={C.g2}>
              <div style={C.card}>
                <div style={C.cardT}>Receipt Scenarios</div>
                <Slider label="Receipt Delay" value={receiptDelayDays} min={0} max={30} unit=" days" color="#f59e0b" onChange={setReceiptDelayDays}/>
                <Slider label="Receipt Haircut (bad debt)" value={receiptHaircut} min={0} max={50} unit="%" color="#ef4444" onChange={setReceiptHaircut}/>
                <div style={{marginTop:4,display:"flex",gap:7}}>
                  <button onClick={()=>{setReceiptDelayDays(0);setReceiptHaircut(0);}} style={{background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:6,padding:"7px 12px",fontSize:11,fontWeight:600,cursor:"pointer",flex:1}}>Reset</button>
                </div>
              </div>
              <div style={C.card}>
                <div style={C.cardT}>Payment Scenarios — Global</div>
                <div style={{fontSize:12,color:"#334155",marginBottom:14}}>Defer <strong style={{color:"#94a3b8"}}>all</strong> supplier payments by this many days. Individual per-payment delays can be set in the Payments tab.</div>
                <Slider label="Defer All Payments" value={globalPaymentDelay} min={0} max={30} unit=" days" color="#10b981" onChange={setGlobalPaymentDelay}/>
                <div style={{marginTop:4,display:"flex",gap:7}}>
                  <button onClick={resetAllPaymentDelays} style={{background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:6,padding:"7px 12px",fontSize:11,fontWeight:600,cursor:"pointer",flex:1}}>Reset All Payment Delays</button>
                  {activeIndivDelays>0&&<div style={{background:"#f59e0b18",color:"#f59e0b",border:"1px solid #f59e0b44",borderRadius:6,padding:"7px 12px",fontSize:11,fontWeight:600,textAlign:"center",flex:1}}>{activeIndivDelays} individual deferral{activeIndivDelays>1?"s":""} active</div>}
                </div>
              </div>
            </div>

            {/* Impact summary */}
            <div style={{...C.card,marginBottom:14}}>
              <div style={C.cardT}>Scenario Impact Summary</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
                {[
                  {label:"Base Month-End Cash",    val:fmtGBP(meBase), c:"#3b82f6"},
                  {label:"Scenario Month-End Cash", val:fmtGBP(meScen), c:meScen<meBase?"#ef4444":"#10b981"},
                  {label:"Impact vs Base",           val:(scenImpact>=0?"+":"")+fmtGBP(scenImpact), c:scenImpact>=0?"#10b981":"#ef4444"},
                ].map(r=>(
                  <div key={r.label} style={{background:"#050b18",border:`1px solid ${r.c}22`,borderRadius:9,padding:"14px 16px",textAlign:"center"}}>
                    <div style={{fontSize:10,color:"#334155",textTransform:"uppercase",letterSpacing:"0.1em",marginBottom:8,fontWeight:700}}>{r.label}</div>
                    <div style={{fontSize:22,fontWeight:800,color:r.c,fontVariantNumeric:"tabular-nums"}}>{r.val}</div>
                  </div>
                ))}
              </div>
              <div style={{marginTop:12,background: scenImpact<-50000?"#ef444411":"#f59e0b0a",border:`1px solid ${scenImpact<-50000?"#ef444433":"#f59e0b22"}`,borderRadius:7,padding:"9px 14px",fontSize:12,color:"#cbd5e1",lineHeight:1.7}}>
                {!anyScenarioActive&&"No scenario active. Use the sliders above or click individual payments in the Payments tab to model deferrals."}
                {anyScenarioActive&&scenImpact>=-20000&&`Moderate scenario. Month-end cash remains healthy at ${fmtGBP(meScen)}.`}
                {anyScenarioActive&&scenImpact<-20000&&scenImpact>=-60000&&`Meaningful deterioration of ${fmtGBP(Math.abs(scenImpact))}. Review facility headroom.`}
                {anyScenarioActive&&scenImpact<-60000&&`⚠ Severe stress. Month-end cash of ${fmtGBP(meScen)} may breach minimum operating balance.`}
              </div>
            </div>

            {/* Dual chart */}
            <div style={C.card}>
              <div style={C.cardT}>Base vs Scenario — Daily Cash Balance</div>
              <ResponsiveContainer width="100%" height={230}>
                <LineChart data={dualForecast} margin={{top:4,right:18,bottom:0,left:6}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#111827"/>
                  <XAxis dataKey="date" tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false}/>
                  <YAxis tickFormatter={v=>`£${(v/1000).toFixed(0)}k`} tick={{fill:"#334155",fontSize:10}} axisLine={false} tickLine={false}/>
                  <Tooltip content={<CashTip/>}/><Legend wrapperStyle={{fontSize:11,color:"#64748b"}}/>
                  <ReferenceLine y={balance} stroke="#10b98133" strokeDasharray="4 4" label={{value:"Opening",fill:"#334155",fontSize:9}}/>
                  <Line type="monotone" dataKey="balance" stroke="#3b82f6" strokeWidth={2.5} dot={false} name="Base Forecast"/>
                  {anyScenarioActive&&<Line type="monotone" dataKey="scenarioBalance" stroke="#f59e0b" strokeWidth={2} dot={false} strokeDasharray="5 3" name="Scenario"/>}
                </LineChart>
              </ResponsiveContainer>
            </div>
          </div>
        )}

        {/* ══ PAYMENTS ══════════════════════════════════════════════════════ */}
        {tab==="payments"&&(
          <div style={{animation:"fade 0.25s ease"}}>
            <div style={C.card}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:10}}>
                <div>
                  <div style={C.cardT}>Supplier Payments</div>
                  <div style={{fontSize:12,color:"#475569",marginTop:-6}}>Click any row to set an individual deferral scenario</div>
                </div>
                <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
                  {/* Global delay mini-control */}
                  <div style={{display:"flex",alignItems:"center",gap:10,background:"#050b18",border:"1px solid #1a2740",borderRadius:8,padding:"8px 14px"}}>
                    <span style={{fontSize:11,color:"#475569",fontWeight:600,whiteSpace:"nowrap"}}>Defer all payments:</span>
                    <div style={{width:120}}>
                      <Slider value={globalPaymentDelay} min={0} max={30} unit=" days" color="#10b981" compact onChange={setGlobalPaymentDelay}/>
                    </div>
                    <span style={{fontSize:13,fontWeight:800,color:globalPaymentDelay>0?"#10b981":"#334155",minWidth:56,fontVariantNumeric:"tabular-nums"}}>
                      {globalPaymentDelay>0?`+${globalPaymentDelay}d`:"None"}
                    </span>
                  </div>
                  {(activeIndivDelays>0||globalPaymentDelay>0)&&(
                    <button onClick={resetAllPaymentDelays} style={{background:"#1e293b",color:"#94a3b8",border:"1px solid #334155",borderRadius:6,padding:"7px 12px",fontSize:11,fontWeight:600,cursor:"pointer"}}>Reset all</button>
                  )}
                  <div style={{fontSize:11,color:"#334155",display:"flex",gap:5}}>Sources: {[...new Set(payments.map(p=>p.source))].map(s=><SrcBadge key={s} source={s}/>)}</div>
                </div>
              </div>

              {/* Scenario summary bar — only shown when delays are active */}
              {anyScenarioActive&&(
                <div style={{display:"flex",gap:10,marginBottom:14,padding:"10px 14px",background:"#0d1f38",borderRadius:8,border:"1px solid #1e3a5f44",alignItems:"center",flexWrap:"wrap"}}>
                  <span style={{fontSize:11,color:"#60a5fa",fontWeight:700}}>ACTIVE DEFERRALS</span>
                  {globalPaymentDelay>0&&<span style={{fontSize:11,color:"#10b981",background:"#10b98115",border:"1px solid #10b98133",borderRadius:4,padding:"2px 8px"}}>All payments: +{globalPaymentDelay}d</span>}
                  {Object.entries(paymentDelays).filter(([,v])=>v>0).map(([inv,d])=>{
                    const p=payments.find(x=>x.invoice_number===inv);
                    return <span key={inv} style={{fontSize:11,color:"#f59e0b",background:"#f59e0b15",border:"1px solid #f59e0b33",borderRadius:4,padding:"2px 8px"}}>{p?.supplier_name||inv}: +{d}d</span>;
                  })}
                  <span style={{marginLeft:"auto",fontSize:12,fontWeight:700,color:scenImpact>=0?"#10b981":"#ef4444"}}>{scenImpact>=0?"+":""}{fmtGBP(scenImpact)} vs base</span>
                </div>
              )}

              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead>
                  <tr>
                    {["Supplier","Invoice","Amount","Due Date","Status","Source","Cash Impact"].map(h=><th key={h} style={C.th}>{h}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {[...payments].sort((a,b)=>new Date(a.expected_payment_date)-new Date(b.expected_payment_date)).map((p,i)=>(
                    <PaymentRow
                      key={p.invoice_number}
                      payment={p}
                      index={i}
                      individualDelay={paymentDelays[p.invoice_number]||0}
                      globalPaymentDelay={globalPaymentDelay}
                      baseMonthEnd={meBase}
                      onDelayChange={handleIndividualPaymentDelay}
                      payments={payments}
                      balance={balance}
                      receipts={receipts}
                      receiptDelayDays={receiptDelayDays}
                      receiptHaircut={receiptHaircut}
                      paymentDelays={paymentDelays}
                    />
                  ))}
                  <tr style={{background:"#0d1a2e"}}>
                    <td style={{padding:"9px 12px",fontWeight:800,color:"#f1f5f9",fontSize:13}} colSpan={2}>TOTAL</td>
                    <td style={{padding:"9px 12px",color:"#ef4444",fontWeight:800,fontSize:14,fontVariantNumeric:"tabular-nums"}}>{fmtGBP(totalP)}</td>
                    <td colSpan={4}/>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ══ RECEIPTS ══════════════════════════════════════════════════════ */}
        {tab==="receipts"&&(
          <div style={{animation:"fade 0.25s ease"}}>
            <div style={C.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div style={C.cardT}>Customer Receipts</div>
                <div style={{fontSize:11,color:"#334155",display:"flex",gap:5}}>Sources: {[...new Set(receipts.map(r=>r.source))].map(s=><SrcBadge key={s} source={s}/>)}</div>
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr>{["Customer","Invoice","Amount","Expected Date","Status","Source"].map(h=><th key={h} style={C.th}>{h}</th>)}</tr></thead>
                <tbody>
                  {[...receipts].sort((a,b)=>new Date(a.expected_payment_date)-new Date(b.expected_payment_date)).map((r,i)=>(
                    <tr key={i} style={{background:i%2?"transparent":"#050b18"}}>
                      <td style={{padding:"9px 12px",borderBottom:"1px solid #0a1020",fontWeight:600,color:"#e2e8f0"}}>{r.customer_name}</td>
                      <td style={{padding:"9px 12px",borderBottom:"1px solid #0a1020",fontFamily:"monospace",fontSize:11,color:"#334155"}}>{r.invoice_number}</td>
                      <td style={{padding:"9px 12px",borderBottom:"1px solid #0a1020",color:"#10b981",fontWeight:700,fontVariantNumeric:"tabular-nums"}}>{fmtGBP(r.amount)}</td>
                      <td style={{padding:"9px 12px",borderBottom:"1px solid #0a1020",color:"#94a3b8"}}>{fmt(r.expected_payment_date)}</td>
                      <td style={{padding:"9px 12px",borderBottom:"1px solid #0a1020"}}><Badge status={r.status}/></td>
                      <td style={{padding:"9px 12px",borderBottom:"1px solid #0a1020"}}><SrcBadge source={r.source}/></td>
                    </tr>
                  ))}
                  <tr style={{background:"#0d1a2e"}}>
                    <td style={{padding:"9px 12px",fontWeight:800,color:"#f1f5f9"}} colSpan={2}>TOTAL</td>
                    <td style={{padding:"9px 12px",color:"#10b981",fontWeight:800,fontSize:14}}>{fmtGBP(totalR)}</td>
                    <td colSpan={3}/>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ══ AI CFO ════════════════════════════════════════════════════════ */}
        {tab==="ai cfo"&&(
          <div style={{animation:"fade 0.25s ease"}}>
            <div style={C.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:18}}>
                <div><div style={C.cardT}>AI CFO Daily Briefing</div><div style={{fontSize:12,color:"#334155"}}>Written in the style of a commercial finance manager reporting to the CFO</div></div>
                <button onClick={generateCFO} disabled={cfoLoading} style={{background:"linear-gradient(135deg,#7c3aed,#6d28d9)",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontSize:12,fontWeight:700,cursor:cfoLoading?"default":"pointer",display:"flex",gap:6,alignItems:"center"}}>
                  <span style={{animation:cfoLoading?"spin 0.8s linear infinite":"none",display:"inline-block"}}>✦</span>
                  {cfoLoading?"Generating…":"Generate AI Briefing"}
                </button>
              </div>
              {!cfoPanel&&!cfoLoading&&<div style={{textAlign:"center",padding:"44px 0",color:"#1e3a5f"}}><div style={{fontSize:36,marginBottom:10}}>✦</div><div style={{fontSize:14,fontWeight:600,color:"#334155"}}>Click "Generate AI Briefing" for today's CFO summary</div></div>}
              {cfoLoading&&<div style={{textAlign:"center",padding:"44px 0",color:"#7c3aed",fontSize:14}}>Generating briefing…</div>}
              {cfoPanel&&!cfoLoading&&<>
                <div style={C.cbox("#3b82f6")}><div style={C.clabel("#3b82f6")}>Key Insights</div><div style={C.ctext}>{cfoPanel.keyInsights}</div></div>
                <div style={C.cbox("#ef4444")}><div style={C.clabel("#ef4444")}>Risks</div><div style={C.ctext}>{cfoPanel.risks}</div></div>
                <div style={C.cbox("#10b981")}>
                  <div style={C.clabel("#10b981")}>Recommended Actions</div>
                  <div style={C.ctext}>{(Array.isArray(cfoPanel.recommendedActions)?cfoPanel.recommendedActions:cfoPanel.recommendedActions.split("\n")).map((a,i)=>(
                    <div key={i} style={{display:"flex",gap:8,marginBottom:7}}><span style={{color:"#10b981",fontWeight:800,minWidth:16}}>{i+1}.</span><span>{a.replace(/^[-•\d.]\s*/,"")}</span></div>
                  ))}</div>
                </div>
                <div style={{fontSize:10,color:"#1e3a5f",textAlign:"right",marginTop:10}}>Generated {new Date().toLocaleString("en-GB")} · {dataMode.toUpperCase()}{anyScenarioActive?" · Scenario active":""}</div>
              </>}
            </div>
          </div>
        )}

        {/* ══ INTEGRATIONS ══════════════════════════════════════════════════ */}
        {tab==="integrations"&&(
          <div style={{animation:"fade 0.25s ease"}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:13,marginBottom:14}}>
              {connStatus.map(conn=>(
                <div key={conn.id} style={{...C.card,borderColor:conn.connected?conn.color+"55":"#1a2740"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:11}}>
                    <span style={{fontSize:22}}>{conn.icon}</span>
                    <div><div style={{fontWeight:700,fontSize:13,color:"#f1f5f9"}}>{conn.label}</div><div style={{fontSize:11,color:"#334155"}}>{conn.desc}</div></div>
                  </div>
                  <div style={{marginBottom:11}}><span style={C.pill(conn.color,conn.connected)}>{conn.connected?"● CONNECTED":"○ NOT CONFIGURED"}</span></div>
                  {syncLog.find(l=>l.source===conn.id)&&<div style={{fontSize:11,color:"#10b981",marginBottom:9}}>✓ {syncLog.find(l=>l.source===conn.id)?.count} records synced</div>}
                  <div style={{fontSize:11,color:"#334155",lineHeight:1.6,marginBottom:11}}>
                    {conn.id==="stripe"&&"Reads PaymentIntents + Payouts. Amounts converted from pence. Read-only restricted key required."}
                    {conn.id==="xero"&&"Reads ACCREC invoices (receipts) + ACCPAY bills (payments). OAuth 2.0 Access Token + Tenant ID."}
                    {conn.id==="bank"&&"40+ UK banks via TrueLayer Open Banking. Authoritative balance + transaction feed."}
                  </div>
                  <button onClick={()=>setShowKeys(true)} style={{background:conn.color+"18",color:conn.color,border:`1px solid ${conn.color}44`,borderRadius:6,padding:"6px 12px",fontSize:11,fontWeight:700,cursor:"pointer",width:"100%"}}>{conn.connected?"Update →":"Configure →"}</button>
                </div>
              ))}
            </div>
            <div style={C.card}>
              <div style={C.cardT}>Data Flow Architecture</div>
              <div style={{display:"flex",alignItems:"center",overflowX:"auto",padding:"8px 0",gap:0}}>
                {[{label:"Stripe",sub:"Intents\nPayouts",c:"#818cf8"},null,{label:"Xero",sub:"Invoices\nBills",c:"#06b6d4"},null,{label:"TrueLayer",sub:"Balance\nTx feed",c:"#10b981"},null,{label:"Connectors",sub:"Normalise\nMerge",c:"#f59e0b"},null,{label:"Forecast",sub:"Day-by-day\nWeekly",c:"#3b82f6"},null,{label:"Scenarios",sub:"Delay sliders\nHaircut",c:"#f97316"},null,{label:"AI Layer",sub:"Commentary\nCFO brief",c:"#8b5cf6"},null,{label:"Dashboard",sub:"KPIs · Export\nChat",c:"#ef4444"}].map((n,i)=>
                  n===null?<div key={i} style={{fontSize:15,color:"#1e3a5f",padding:"0 5px",flexShrink:0}}>→</div>
                  :<div key={i} style={{background:n.c+"15",border:`1px solid ${n.c}40`,borderRadius:7,padding:"8px 10px",textAlign:"center",flexShrink:0,minWidth:76}}>
                    <div style={{fontSize:10,fontWeight:800,color:n.c,whiteSpace:"pre-line",lineHeight:1.4}}>{n.label}</div>
                    <div style={{fontSize:9,color:"#334155",marginTop:3,whiteSpace:"pre-line",lineHeight:1.4}}>{n.sub}</div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
