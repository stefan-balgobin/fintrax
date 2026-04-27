import React, { useState, useContext, createContext } from "react";
import * as XLSX from "xlsx";

// ── THEME ─────────────────────────────────────────────────────────────────────
const DARK = {
  bg:"#0a0e1a", surface:"#111827", card:"#1a2235", border:"#1e2d45",
  accent:"#00d4ff", accentGlow:"rgba(0,212,255,0.12)", green:"#00e5a0",
  red:"#ff4d6d", yellow:"#ffd166", purple:"#a78bfa", text:"#e2e8f0", muted:"#64748b",
};
const LIGHT = {
  bg:"#f1f5f9", surface:"#ffffff", card:"#ffffff", border:"#e2e8f0",
  accent:"#0369a1", accentGlow:"rgba(3,105,161,0.08)", green:"#047857",
  red:"#dc2626", yellow:"#b45309", purple:"#6d28d9", text:"#0f172a", muted:"#64748b",
};

// Theme context — all components read T from here
const ThemeCtx = createContext(DARK);
const useT = () => useContext(ThemeCtx);

// T is only used as a fallback for the AuthScreen which renders before ThemeCtx
// All other components get T from useT()
const T = DARK;

// ── HELPERS ───────────────────────────────────────────────────────────────────
const fmt = n => n == null || n === "" ? "—" : `$${Number(n).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
const nextId = arr => arr.length ? Math.max(...arr.map(x => x.id)) + 1 : 1;
const CY = new Date().getFullYear();
const fmtDate = str => { if(!str||str==="—") return "—"; const d=new Date(str); return isNaN(d)?str:d.toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}); };
const timeLeft = exp => { if(!exp) return "—"; const diff=new Date(exp)-new Date(); if(diff<0) return "Expired"; const days=Math.floor(diff/86400000),y=Math.floor(days/365),m=Math.floor((days%365)/30),d=days%30; if(y>0) return `${y}y ${m}m ${d}d`; if(m>0) return `${m}m ${d}d`; return `${d}d`; };
const urgencyColor = (tl, t=T) => (!tl||tl==="Expired") ? t.red : !tl.includes("y") ? t.yellow : t.green;
const expiryYear = s => s ? new Date(s).getFullYear() : null;
const logYear = s => s && s!=="-" ? new Date(s).getFullYear() : null;
const effectiveRenews = sub => {
  if(!sub.autoRenew||!sub.renews||sub.cancelled) return sub.renews||"";
  let d=new Date(sub.renews);
  const today=new Date();
  if(d>today) return sub.renews;
  while(d<=today){ if(sub.type==="Annual") d.setFullYear(d.getFullYear()+1); else d.setMonth(d.getMonth()+1); }
  return d.toISOString().slice(0,10);
};

// ── STYLES — generated fresh from current theme T ─────────────────────────────
const mkS = (T) => ({
  card: { background:T.card, border:`1px solid ${T.border}`, borderRadius:12, padding:"14px 16px", marginBottom:12, width:"100%", boxSizing:"border-box", overflow:"hidden", minWidth:0 },
  label: { fontSize:10, letterSpacing:2, color:T.muted, textTransform:"uppercase", marginBottom:6, display:"block" },
  input: { width:"100%", background:T.surface, border:`1px solid ${T.border}`, borderRadius:8, color:T.text, padding:"10px 12px", fontSize:14, fontFamily:"inherit", boxSizing:"border-box", minWidth:0 },
  btn: (c=T.accent) => ({ padding:"10px 16px", borderRadius:8, border:`1px solid ${c}`, background:c+"20", color:c, cursor:"pointer", fontSize:13, fontFamily:"inherit", fontWeight:700, whiteSpace:"nowrap" }),
  badge: c => ({ display:"inline-block", padding:"3px 10px", borderRadius:20, fontSize:11, fontWeight:700, background:c+"20", color:c, border:`1px solid ${c}40`, whiteSpace:"nowrap" }),
  outlineBtn: (active, c=T.accent) => ({ background:"none", border:`1px solid ${active?c:T.border}`, borderRadius:6, color:active?c:T.muted, fontSize:12, fontFamily:"inherit", padding:"6px 14px", cursor:"pointer", fontWeight:600 }),
});
// Static fallback S for AuthScreen (renders before ThemeCtx)
const S = mkS(DARK);
// Hook for all other components to get live theme-aware styles
const useS = () => mkS(useContext(ThemeCtx));

// ── INITIAL DATA ──────────────────────────────────────────────────────────────
const INIT = {
  expenses:[],
  certs:[],
  personalDocs:[],
  subscriptions:[],
  vehicles:[],
  carLog:[],
  leisure:[],
  investments:[],
};

// ── PERSIST UI TOGGLES TO LOCALSTORAGE ───────────────────────────────────────
function useLocalState(key,init){
  const [val,setVal] = useState(()=>{
    try{const v=localStorage.getItem("fintrax_ui_"+key);return v!==null?JSON.parse(v):init;}catch{return init;}
  });
  const set = React.useCallback(v=>{
    setVal(prev=>{
      const next=typeof v==="function"?v(prev):v;
      try{localStorage.setItem("fintrax_ui_"+key,JSON.stringify(next));}catch{}
      return next;
    });
  },[key]);
  return [val,set];
}

// ── SHARED COMPONENTS ─────────────────────────────────────────────────────────
function StatCard({label,value,sub,color,onClick,active}){
  const T = useT(); const S = useS();
  return(
    <div onClick={onClick} style={{...S.card,marginBottom:0,borderColor:color+(active?"99":"40"),boxShadow:`0 0 14px ${color}${active?"30":"10"}`,cursor:onClick?"pointer":"default",overflow:"hidden",minWidth:0}}>
      <div style={{fontSize:9,letterSpacing:1.5,color:T.muted,textTransform:"uppercase",marginBottom:4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{label}</div>
      <div style={{fontSize:18,fontWeight:700,color,lineHeight:1.2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:T.muted,marginTop:2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{sub}</div>}
      {active&&<div style={{fontSize:9,color,letterSpacing:2,marginTop:4}}>● ACTIVE</div>}
    </div>
  );
}

function Modal({title,onClose,children}){
  const T = useT(); const S = useS();
  React.useEffect(()=>{
    const prev=document.body.style.overflow;
    document.body.style.overflow="hidden";
    return()=>{ document.body.style.overflow=prev; };
  },[]);
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.8)",zIndex:1000,display:"flex",alignItems:"flex-end",justifyContent:"center",padding:0}}>
      <div style={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:"16px 16px 0 0",padding:"20px 16px 32px",width:"100%",maxWidth:600,maxHeight:"92vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
          <div style={{fontWeight:700,fontSize:15,color:T.accent,letterSpacing:1}}>{title}</div>
          <button onClick={onClose} style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:22,lineHeight:1,padding:"0 4px"}}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function ConfirmDelete({label,onCancel,onConfirm}){
  const T = useT(); const S = useS();
  return(
    <Modal title="Confirm Delete" onClose={onCancel}>
      <div style={{fontSize:14,color:T.text,marginBottom:20,lineHeight:1.6}}>Delete <span style={{color:T.red,fontWeight:700}}>{label}</span>?<br/><span style={{color:T.muted,fontSize:12}}>This cannot be undone.</span></div>
      <div style={{display:"flex",gap:10}}>
        <button style={{...S.btn(T.muted),flex:1}} onClick={onCancel}>Cancel</button>
        <button style={{...S.btn(T.red),flex:1}} onClick={onConfirm}>Delete</button>
      </div>
    </Modal>
  );
}

function Field({label,children}){
  const T = useT(); const S = useS();return <div style={{marginBottom:14}}><label style={S.label}>{label}</label>{children}</div>;}
function G2({children}){return <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit, minmax(140px, 1fr))",gap:10}}>{children}</div>;}
function SaveCancel({onCancel,onSave}){
  const T = useT(); const S = useS();
  return(
    <div style={{display:"flex",gap:10,marginTop:16}}>
      <button style={{...S.btn(T.muted),flex:1}} onClick={onCancel}>Cancel</button>
      <button style={{...S.btn(T.green),flex:1}} onClick={onSave}>Save</button>
    </div>
  );
}

function ItemActions({label,onEdit,onDelete}){
  const T = useT(); const S = useS();
  const [open,setOpen] = useState(false);
  const [conf,setConf] = useState(false);
  return(
    <div>
      <button style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:14,padding:"4px 8px"}} onClick={()=>setOpen(o=>!o)}>{open?"▲":"▼"}</button>
      {open&&(
        <div style={{display:"flex",gap:8,marginTop:8,paddingTop:8,borderTop:`1px solid ${T.border}30`,justifyContent:"flex-end"}}>
          <button style={{...S.btn(T.accent),padding:"6px 16px",fontSize:12}} onClick={()=>{setOpen(false);onEdit();}}>✏ Edit</button>
          <button style={{...S.btn(T.red),padding:"6px 16px",fontSize:12}} onClick={()=>setConf(true)}>🗑 Delete</button>
        </div>
      )}
      {conf&&<ConfirmDelete label={label} onCancel={()=>setConf(false)} onConfirm={()=>{setConf(false);setOpen(false);onDelete();}}/>}
    </div>
  );
}

function SearchBar({value,onChange,placeholder}){
  const T = useT(); const S = useS();
  return <input style={{...S.input,marginBottom:14}} value={value} onChange={e=>onChange(e.target.value)} placeholder={placeholder||"Search..."}/>;
}

function ListToggle({collapsed,onToggle,count,label}){
  const T = useT(); const S = useS();
  return(
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
      <div style={{fontSize:11,letterSpacing:2,color:T.accent,textTransform:"uppercase",fontWeight:700}}>{label}</div>
      <button onClick={onToggle} style={{...S.outlineBtn(false),padding:"5px 12px",fontSize:11}}>{collapsed?`▼ Show (${count})`:"▲ Hide"}</button>
    </div>
  );
}

const VIEWABLE = /\.(pdf|png|jpg|jpeg|gif|webp|svg|bmp)$/i;
function openFile(f){
  if(VIEWABLE.test(f.name)){window.open(f.dataUrl,"_blank","noopener");}
  else{const a=document.createElement("a");a.href=f.dataUrl;a.download=f.name;a.click();}
}

function FilesSection({files,onUpload,onDownload,onRemove}){
  const T = useT(); const S = useS();
  const [open,setOpen] = useState(false);
  const [confirmFile,setConfirmFile] = useState(null);
  const list = files||[];
  return(
    <div style={{borderTop:`1px solid ${T.border}30`,paddingTop:10,marginTop:4}}>
      <button style={S.outlineBtn(open)} onClick={()=>setOpen(o=>!o)}>Attached Files ({list.length})</button>
      {open&&(
        <div style={{marginTop:10}}>
          <label style={{...S.btn(T.purple),padding:"6px 14px",fontSize:12,cursor:"pointer",display:"inline-block",marginBottom:10}}>
            + Upload<input type="file" multiple style={{display:"none"}} onChange={onUpload}/>
          </label>
          {list.length===0&&<div style={{fontSize:12,color:T.muted}}>No files yet.</div>}
          {list.map((f,i)=>(
            <div key={`${f.name}-${i}`} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:T.bg,borderRadius:8,padding:"8px 10px",marginBottom:6}}>
              <div style={{minWidth:0,flex:1}}>
                <div onClick={()=>openFile(f)} style={{color:T.accent,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",cursor:"pointer",textDecoration:"underline"}}>{f.name}</div>
                <div style={{color:T.muted,fontSize:10}}>{f.uploadedAt} · {(f.size/1024).toFixed(1)} KB</div>
              </div>
              <div style={{display:"flex",gap:6,marginLeft:8,flexShrink:0}}>
                <button style={{...S.btn(T.green),padding:"4px 10px",fontSize:11}} onClick={()=>onDownload(f)}>⬇</button>
                <button style={{...S.btn(T.red),padding:"4px 8px",fontSize:11}} onClick={()=>setConfirmFile(f.name)}>✕</button>
              </div>
            </div>
          ))}
        </div>
      )}
      {confirmFile&&(
        <ConfirmDelete label={confirmFile} onCancel={()=>setConfirmFile(null)} onConfirm={()=>{onRemove(confirmFile);setConfirmFile(null);}}/>
      )}
    </div>
  );
}

function useFiles(key,setData){
  const upload = (itemId,e) => {
    Array.from(e.target.files).forEach(file=>{
      const r=new FileReader();
      r.onload=ev=>setData(d=>({...d,[key]:d[key].map(x=>x.id===itemId?{...x,files:[...(x.files||[]),{name:file.name,type:file.type,size:file.size,dataUrl:ev.target.result,uploadedAt:new Date().toLocaleDateString()}]}:x)}));
      r.readAsDataURL(file);
    });
    e.target.value="";
  };
  const download = f=>{const a=document.createElement("a");a.href=f.dataUrl;a.download=f.name;a.click();};
  const remove = (itemId,name)=>setData(d=>({...d,[key]:d[key].map(x=>x.id===itemId?{...x,files:(x.files||[]).filter(f=>f.name!==name)}:x)}));
  return {upload,download,remove};
}

// ── OVERVIEW ──────────────────────────────────────────────────────────────────
function Overview({data, enabledPages}){
  const T = useT(); const S = useS();
  const ep = enabledPages||{};

  // ── calculations (mirror Expenses page logic, gated by enabledPages) ──
  const parseM = str=>{if(!str)return[];return str.split(",").map(m=>{const p=m.trim().split(":");return{name:p[0]?.trim()||"",paid:(p[1]?.trim()||"").toLowerCase()==="paid",date:p[2]?.trim()||""};}).filter(m=>m.name);};
  const subUserCost = sub=>{ const members=parseM(sub.members); const amt=Number(sub.amount||0); return members.length>0?amt/members.length:amt; };
  const subTotal     = ep.subscriptions!==false ? data.subscriptions.reduce((s,sub)=>s+(sub.type==="Annual"?subUserCost(sub):subUserCost(sub)*12),0) : 0;
  const certTotal    = ep.certs!==false ? data.certs.filter(c=>expiryYear(c.expiry)===CY).reduce((s,c)=>s+Number(c.costTTD||0),0) : 0;
  const docsTotal    = ep.personal!==false ? data.personalDocs.filter(d=>expiryYear(d.expiry)===CY).reduce((s,d)=>s+Number(d.cost||0),0) : 0;
  const carTotal     = ep.car!==false ? data.carLog.filter(r=>logYear(r.date)===CY).reduce((s,r)=>s+Number(r.cost||0),0) : 0;
  const leisureTotal = ep.leisure!==false ? data.leisure.reduce((s,r)=>{
    const legs=(r.legs||[]).slice().sort((a,b)=>((a.flightType==="Multi-City"?a.segments?.[0]?.departDate:a.departDate)||"").localeCompare((b.flightType==="Multi-City"?b.segments?.[0]?.departDate:b.departDate)||""));
    const fd=legs.length>0?(legs[0].flightType==="Multi-City"?legs[0].segments?.[0]?.departDate:legs[0].departDate):null;
    if(fd&&logYear(fd)!==CY) return s;
    return s+(r.legs||[]).reduce((sf,l)=>sf+Number(l.flightCost||0),0)+(r.accommodations||[]).reduce((sa,a)=>sa+Number(a.cost||0),0);
  },0) : 0;
  const manualTotal  = data.expenses.reduce((s,e)=>s+Number(e.amount||0),0);
  const totalCY      = subTotal+certTotal+docsTotal+carTotal+leisureTotal+manualTotal;
  const totalInv     = data.investments.reduce((s,r)=>s+Number(r.value||0),0);
  const monthsElapsed= new Date().getMonth()+1;
  const monthlyAvg   = totalCY/monthsElapsed;

  // ── alerts & upcoming (only from enabled pages) ──
  const alerts = [
    ...(ep.certs!==false ? data.certs.map(c=>({name:c.name,expiry:c.expiry,type:"Cert"})) : []),
    ...(ep.personal!==false ? data.personalDocs.map(d=>({name:d.type,expiry:d.expiry,type:"Doc"})) : []),
  ].filter(d=>d.expiry&&urgencyColor(timeLeft(d.expiry),T)!==T.green).sort((a,b)=>new Date(a.expiry)-new Date(b.expiry));
  const soon30 = ep.subscriptions!==false ? data.subscriptions.filter(s=>{if(s.cancelled)return false;const er=effectiveRenews(s);if(!er)return false;const diff=new Date(er)-new Date();return diff>0&&diff<=30*86400000;}).sort((a,b)=>new Date(effectiveRenews(a))-new Date(effectiveRenews(b))) : [];

  // ── spending breakdown rows (only enabled pages) ──
  const breakdown = [
    ep.subscriptions!==false && {label:"Subscriptions",   amount:subTotal,     color:T.accent},
    ep.certs!==false          && {label:"Certifications",  amount:certTotal,    color:T.yellow},
    ep.personal!==false       && {label:"Personal Docs",   amount:docsTotal,    color:T.purple},
    ep.car!==false            && {label:"Car Maintenance", amount:carTotal,     color:T.green},
    ep.leisure!==false        && {label:"Leisure & Travel",amount:leisureTotal, color:T.red},
    ...data.expenses.map(e=>({label:e.name,amount:Number(e.amount||0),color:e.category==="Fixed"?T.purple:T.muted})),
  ].filter(Boolean).filter(r=>r.amount>0);

  const showTotalSpend    = ep.expenses!==false;
  const showSubs          = ep.subscriptions!==false;
  const showAlerts        = ep.certs!==false||ep.personal!==false;
  const showInvCard       = ep.investments!==false;
  const [breakdownOpen,   setBreakdownOpen]   = useLocalState("overview_breakdown_open",true);
  const [invSnapshotOpen, setInvSnapshotOpen] = useLocalState("overview_inv_snapshot_open",true);

  const Section = ({title,color,onToggle,collapsed,children})=>(
    <div style={{...S.card,padding:"14px 16px",marginBottom:10}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:collapsed?0:12}}>
        <div style={{fontSize:10,letterSpacing:2,color:color||T.accent,textTransform:"uppercase",fontWeight:700}}>{title}</div>
        {onToggle&&<button onClick={onToggle} style={{...S.outlineBtn(false),padding:"3px 10px",fontSize:11}}>{collapsed?"▼ Show":"▲ Hide"}</button>}
      </div>
      {!collapsed&&children}
    </div>
  );

  return(
    <div>
      {/* ── header ── */}
      <div style={{marginBottom:16}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>OVERVIEW</div>
        <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginTop:2}}>// FINANCE DASHBOARD · {CY} · {new Date().toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"})}</div>
      </div>

      {/* ── key metrics ── */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:10,marginBottom:14}}>
        {showTotalSpend&&<StatCard label={`${CY} Total Spend`}  value={fmt(totalCY)}    sub={`avg ${fmt(monthlyAvg)}/mo`} color={T.red}/>}
        {showInvCard&&<StatCard label="Investments Value"        value={fmt(totalInv)}   sub={`${data.investments.length} holding${data.investments.length!==1?"s":""}`} color={T.green}/>}
        {showSubs&&<StatCard label="Subs Cost/yr"                value={fmt(subTotal)}   sub={`${data.subscriptions.filter(s=>!s.cancelled).length} active`} color={T.accent}/>}
        {showAlerts&&<StatCard label="Expiry Alerts"             value={alerts.length}   sub={alerts.length>0?`${alerts.filter(a=>timeLeft(a.expiry)==="Expired").length} expired`:"all clear"} color={alerts.length>0?T.yellow:T.green}/>}
      </div>

      {/* ── alerts + upcoming renewals ── */}
      {(alerts.length>0||soon30.length>0)&&(
        <div style={{display:"grid",gridTemplateColumns:alerts.length>0&&soon30.length>0?"1fr 1fr":"1fr",gap:10,marginBottom:10}}>
          {alerts.length>0&&(
            <Section title="⚠ Expiry Alerts" color={T.yellow}>
              {alerts.map((d,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${T.border}20`}}>
                  <div style={{minWidth:0,flex:1,paddingRight:6}}>
                    <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.name}</div>
                    <div style={{fontSize:10,color:T.muted}}>{d.type}</div>
                  </div>
                  <span style={{...S.badge(urgencyColor(timeLeft(d.expiry),T)),fontSize:10,flexShrink:0}}>{timeLeft(d.expiry)}</span>
                </div>
              ))}
            </Section>
          )}
          {soon30.length>0&&(
            <Section title="↻ Renewing Soon" color={T.accent}>
              {soon30.map((s,i)=>(
                <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${T.border}20`}}>
                  <div style={{minWidth:0,flex:1,paddingRight:6}}>
                    <div style={{fontSize:12,fontWeight:600,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.service}</div>
                    <div style={{fontSize:10,color:T.muted}}>{s.type}</div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0}}>
                    <div style={{fontSize:12,fontWeight:700,color:T.accent}}>{fmt(subUserCost(s))}</div>
                    <div style={{fontSize:10,color:T.muted}}>{fmtDate(effectiveRenews(s))}</div>
                  </div>
                </div>
              ))}
            </Section>
          )}
        </div>
      )}

      {/* ── spending breakdown ── */}
      {showTotalSpend&&breakdown.length>0&&(
        <Section title={`${CY} Spending Breakdown`} onToggle={()=>setBreakdownOpen(o=>!o)} collapsed={!breakdownOpen}>
          {breakdown.map((r,i)=>{
            const pct=totalCY>0?(r.amount/totalCY)*100:0;
            return(
              <div key={i} style={{marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:4}}>
                  <span style={{fontSize:12,color:T.text}}>{r.label}</span>
                  <div style={{display:"flex",gap:8,alignItems:"baseline"}}>
                    <span style={{fontSize:11,color:T.muted}}>{pct.toFixed(1)}%</span>
                    <span style={{fontSize:13,fontWeight:700,color:r.color}}>{fmt(r.amount)}</span>
                  </div>
                </div>
                <div style={{height:6,background:T.border+"60",borderRadius:4,overflow:"hidden"}}>
                  <div style={{height:"100%",width:`${pct}%`,background:r.color,borderRadius:4,transition:"width 0.3s ease"}}/>
                </div>
              </div>
            );
          })}
          <div style={{borderTop:`1px solid ${T.border}40`,marginTop:12,paddingTop:10,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:11,color:T.muted,letterSpacing:1}}>TOTAL</span>
            <span style={{fontSize:16,fontWeight:700,color:T.red}}>{fmt(totalCY)}</span>
          </div>
        </Section>
      )}

      {/* ── investments snapshot ── */}
      {ep.investments!==false&&data.investments.length>0&&(
        <Section title="Investments Snapshot" color={T.green} onToggle={()=>setInvSnapshotOpen(o=>!o)} collapsed={!invSnapshotOpen}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <span style={{fontSize:22,fontWeight:700,color:T.green}}>{fmt(totalInv)}</span>
            <span style={{fontSize:11,color:T.muted}}>{data.investments.length} holding{data.investments.length!==1?"s":""}</span>
          </div>
          {data.investments.slice(0,5).map((inv,i)=>{
            const pct=totalInv>0?(Number(inv.value||0)/totalInv)*100:0;
            return(
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:`1px solid ${T.border}20`}}>
                <span style={{fontSize:12,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",paddingRight:8}}>{inv.name||inv.ticker||"—"}</span>
                <div style={{display:"flex",gap:10,alignItems:"center",flexShrink:0}}>
                  <span style={{fontSize:10,color:T.muted}}>{pct.toFixed(1)}%</span>
                  <span style={{fontSize:12,fontWeight:700,color:T.green}}>{fmt(inv.value)}</span>
                </div>
              </div>
            );
          })}
          {data.investments.length>5&&<div style={{fontSize:11,color:T.muted,marginTop:8,textAlign:"center"}}>+{data.investments.length-5} more holdings</div>}
        </Section>
      )}
    </div>
  );
}

// ── EXPENSES ──────────────────────────────────────────────────────────────────
function Expenses({data,setData,enabledPages}){
  const T = useT(); const S = useS();
  const ep = enabledPages||{};
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [search,setSearch] = useState("");
  const [collapsed,setCollapsed] = useLocalState("expenses_collapsed",false);
  const [cardFilter,setCardFilter] = useLocalState("expenses_card_filter","All");
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const parseM = str=>{if(!str)return[];return str.split(",").map(m=>{const p=m.trim().split(":");return{name:p[0]?.trim()||"",paid:(p[1]?.trim()||"").toLowerCase()==="paid",date:p[2]?.trim()||""};}).filter(m=>m.name);};
  const subUserCost = sub=>{ const members=parseM(sub.members); const amt=Number(sub.amount||0); return members.length>0?amt/members.length:amt; };
  const subTotal = ep.subscriptions!==false ? data.subscriptions.reduce((s,sub)=>s+(sub.type==="Annual"?subUserCost(sub):subUserCost(sub)*12),0) : 0;
  const certTotal = ep.certs!==false ? data.certs.filter(c=>expiryYear(c.expiry)===CY).reduce((s,c)=>s+Number(c.costTTD||0),0) : 0;
  const personalDocsTotal = ep.personal!==false ? data.personalDocs.filter(d=>expiryYear(d.expiry)===CY).reduce((s,d)=>s+Number(d.cost||0),0) : 0;
  const carTotal = ep.car!==false ? data.carLog.filter(r=>logYear(r.date)===CY).reduce((s,r)=>s+Number(r.cost||0),0) : 0;
  const leisureTotal = ep.leisure!==false ? data.leisure.reduce((s,r)=>{
    const legs=(r.legs||[]).slice().sort((a,b)=>((a.flightType==="Multi-City"?a.segments?.[0]?.departDate:a.departDate)||"").localeCompare((b.flightType==="Multi-City"?b.segments?.[0]?.departDate:b.departDate)||""));
    const fd=legs.length>0?(legs[0].flightType==="Multi-City"?legs[0].segments?.[0]?.departDate:legs[0].departDate):null;
    if(fd&&logYear(fd)!==CY) return s;
    return s+(r.legs||[]).reduce((sf,l)=>sf+Number(l.flightCost||0),0)+(r.accommodations||[]).reduce((sa,a)=>sa+Number(a.cost||0),0);
  },0) : 0;
  const AUTO_EXPENSES = [
    ep.subscriptions!==false && {id:"auto-subs",    name:"Subscriptions",     category:"Variable", amount:subTotal,         live:true},
    ep.certs!==false          && {id:"auto-certs",   name:"Certifications",    category:"Variable", amount:certTotal,        live:true},
    ep.personal!==false       && {id:"auto-docs",    name:"Personal Documents",category:"Variable", amount:personalDocsTotal,live:true},
    ep.car!==false            && {id:"auto-car",     name:"Car Maintenance",   category:"Variable", amount:carTotal,         live:true},
    ep.leisure!==false        && {id:"auto-leisure", name:"Leisure & Travel",  category:"Variable", amount:leisureTotal,     live:true},
  ].filter(Boolean);
  const allExpenses = [...AUTO_EXPENSES, ...data.expenses];
  const total = allExpenses.reduce((s,e)=>s+Number(e.amount||0),0);
  const fixedTotal = data.expenses.filter(e=>e.category==="Fixed").reduce((s,e)=>s+Number(e.amount||0),0);
  const varTotal = AUTO_EXPENSES.reduce((s,e)=>s+Number(e.amount||0),0) + data.expenses.filter(e=>e.category!=="Fixed").reduce((s,e)=>s+Number(e.amount||0),0);
  const baseShown = allExpenses.filter(e=>cardFilter==="Fixed"?e.category==="Fixed":cardFilter==="Variable"?e.category!=="Fixed":true);
  const shown = baseShown.filter(e=>!search||e.name.toLowerCase().includes(search.toLowerCase()));
  function save(){const entry={...form,amount:Number(form.amount||0)};setData(d=>({...d,expenses:modal==="add"?[...d.expenses,{...entry,id:nextId(d.expenses)}]:d.expenses.map(e=>e.id===form.id?entry:e)}));setModal(null);}
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>EXPENSES</div>
        <button style={{...S.btn(T.green),padding:"8px 14px"}} onClick={()=>{setForm({category:"Fixed",name:"",amount:""});setModal("add");}}>+ Add</button>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// {CY} ANNUAL EXPENSE TRACKER</div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search expenses..."/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
        <StatCard label={`${CY} Total`} value={fmt(total)} sub={`${allExpenses.length} entries`} color={T.red} onClick={()=>setCardFilter("All")} active={cardFilter==="All"}/>
        <StatCard label="Fixed" value={fmt(fixedTotal)} sub={`${data.expenses.filter(e=>e.category==="Fixed").length} entries`} color={T.purple} onClick={()=>setCardFilter(f=>f==="Fixed"?"All":"Fixed")} active={cardFilter==="Fixed"}/>
        <StatCard label="Variable" value={fmt(varTotal)} sub={`${allExpenses.filter(e=>e.category!=="Fixed").length} entries`} color={T.yellow} onClick={()=>setCardFilter(f=>f==="Variable"?"All":"Variable")} active={cardFilter==="Variable"}/>
      </div>
      <ListToggle collapsed={collapsed} onToggle={()=>setCollapsed(c=>!c)} count={shown.length} label="Expenses"/>
      {!collapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {shown.map(e=>(
            <div key={e.id} style={S.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:600,fontSize:14,marginBottom:4}}>{e.name}</div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                    <span style={S.badge(e.category==="Fixed"?T.purple:T.yellow)}>{e.category}</span>
                    {e.live&&<span style={{fontSize:10,color:T.accent}}>⟳ auto</span>}
                  </div>
                </div>
                <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4,marginLeft:8}}>
                  <div style={{fontSize:17,fontWeight:700,color:T.accent}}>{fmt(e.amount)}</div>
                  <div style={{fontSize:10,color:T.muted}}>{total>0?(Number(e.amount)/total*100).toFixed(1)+"%":""}</div>
                  {!e.live&&<ItemActions label={e.name} onEdit={()=>{setForm({...e});setModal("edit");}} onDelete={()=>setData(d=>({...d,expenses:d.expenses.filter(x=>x.id!==e.id)}))}/>}
                </div>
              </div>
            </div>
          ))}
        </div>
      )}
      {modal&&(
        <Modal title={modal==="add"?"Add Expense":"Edit Expense"} onClose={()=>setModal(null)}>
          <Field label="Category"><select style={S.input} value={form.category} onChange={e=>upd("category",e.target.value)}><option>Fixed</option><option>Variable</option></select></Field>
          <Field label="Name"><input style={S.input} value={form.name||""} onChange={e=>upd("name",e.target.value)} placeholder="e.g. Rent"/></Field>
          <Field label="Amount (TTD)"><input style={S.input} type="number" value={form.amount||""} onChange={e=>upd("amount",e.target.value)} placeholder="0.00"/></Field>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
    </div>
  );
}

// ── CERTIFICATIONS ────────────────────────────────────────────────────────────
function Certifications({data,setData}){
  const T = useT(); const S = useS();
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [yearFilter,setYearFilter] = useLocalState("certs_year_filter","All");
  const [search,setSearch] = useState("");
  const [collapsed,setCollapsed] = useLocalState("certs_collapsed",false);
  const [cardFilter,setCardFilter] = useLocalState("certs_card_filter","All");
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const {upload,download,remove} = useFiles("certs",setData);
  const allYears = [...new Set(data.certs.map(c=>expiryYear(c.expiry)).filter(Boolean))].sort();
  const byYear = yearFilter==="All"?data.certs:data.certs.filter(c=>expiryYear(c.expiry)===Number(yearFilter));
  const expiringNow = data.certs.filter(c=>expiryYear(c.expiry)===CY).length;
  let shown = yearFilter==="All"&&cardFilter==="ExpiringThisYear"?byYear.filter(c=>expiryYear(c.expiry)===CY):byYear;
  shown = shown.filter(c=>!search||c.name.toLowerCase().includes(search.toLowerCase()));
  const totalCost = shown.reduce((s,c)=>s+Number(c.costTTD||0),0);
  function save(){const entry={...form,files:form.files||[]};setData(d=>({...d,certs:modal==="add"?[...d.certs,{...entry,id:nextId(d.certs)}]:d.certs.map(c=>c.id===form.id?entry:c)}));setModal(null);}
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>CERTIFICATIONS</div>
        <button style={{...S.btn(T.green),padding:"8px 14px"}} onClick={()=>{setForm({name:"",issued:"",expiry:"",costTTD:"",costUSD:"",files:[]});setModal("add");}}>+ Add</button>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// PROFESSIONAL CERTIFICATIONS</div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search certifications..."/>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
        <span style={{fontSize:11,color:T.muted}}>Year:</span>
        <select style={{...S.input,width:"auto",minWidth:100}} value={yearFilter} onChange={e=>{setYearFilter(e.target.value);setCardFilter("All");}}>
          {["All",...allYears.map(String)].map(y=><option key={y}>{y}</option>)}
        </select>
        <span style={{fontSize:11,color:T.muted}}>{yearFilter==="All"?`${data.certs.length} total`:`${shown.length} expiring ${yearFilter}`}</span>
      </div>
      {yearFilter==="All"?(
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
          <StatCard label="Total" value={byYear.length} color={T.accent} onClick={()=>setCardFilter(f=>f==="TotalCerts"?"All":"TotalCerts")} active={cardFilter==="TotalCerts"}/>
          <StatCard label={`Exp. ${CY}`} value={expiringNow} color={T.yellow} onClick={()=>setCardFilter(f=>f==="ExpiringThisYear"?"All":"ExpiringThisYear")} active={cardFilter==="ExpiringThisYear"}/>
          <StatCard label="Cost TTD" value={fmt(totalCost)} color={T.red}/>
        </div>
      ):(
        <div style={{marginBottom:14}}><StatCard label={`Cost (TTD) · ${yearFilter}`} value={fmt(totalCost)} color={T.red}/></div>
      )}
      <ListToggle collapsed={collapsed} onToggle={()=>setCollapsed(c=>!c)} count={shown.length} label="Certifications"/>
      {!collapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {shown.length===0&&<div style={{...S.card,textAlign:"center",color:T.muted}}>No certifications match.</div>}
          {shown.map(c=>{const tl=timeLeft(c.expiry),uc=urgencyColor(tl,T);return(
            <div key={c.id} style={{...S.card,borderColor:uc+"40"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
                <div style={{fontWeight:700,fontSize:14,color:T.text,flex:1}}>{c.name}</div>
                <span style={{...S.badge(uc),flexShrink:0}}>{tl}</span>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:T.muted,marginBottom:8}}>
                <span>Issued: <span style={{color:T.text}}>{fmtDate(c.issued)}</span></span>
                <span>Expires: <span style={{color:T.text}}>{fmtDate(c.expiry)}</span></span>
              </div>
              {(c.costTTD||c.costUSD)&&(
                <div style={{display:"flex",gap:14,marginBottom:10,flexWrap:"wrap"}}>
                  {c.costTTD&&<span style={{fontSize:12,color:T.accent}}>TTD: {fmt(c.costTTD)}</span>}
                  {c.costUSD&&<span style={{fontSize:12,color:T.purple}}>USD: ${c.costUSD}</span>}
                </div>
              )}
              <FilesSection files={c.files} onUpload={e=>upload(c.id,e)} onDownload={download} onRemove={name=>remove(c.id,name)}/>
              <div style={{display:"flex",justifyContent:"flex-end",marginTop:4}}>
                <ItemActions label={c.name} onEdit={()=>{setForm({...c,files:c.files||[]});setModal("edit");}} onDelete={()=>setData(d=>({...d,certs:d.certs.filter(x=>x.id!==c.id)}))}/>
              </div>
            </div>
          );})}
        </div>
      )}
      {modal&&(
        <Modal title={modal==="add"?"Add Cert":"Edit Cert"} onClose={()=>setModal(null)}>
          <Field label="Name"><input style={S.input} value={form.name||""} onChange={e=>upd("name",e.target.value)}/></Field>
          <G2>
            <Field label="Issued"><input style={S.input} type="date" value={form.issued||""} onChange={e=>upd("issued",e.target.value)}/></Field>
            <Field label="Expiry"><input style={S.input} type="date" value={form.expiry||""} onChange={e=>upd("expiry",e.target.value)}/></Field>
          </G2>
          <G2>
            <Field label="Cost TTD"><input style={S.input} type="number" value={form.costTTD||""} onChange={e=>upd("costTTD",e.target.value)}/></Field>
            <Field label="Cost USD"><input style={S.input} type="number" value={form.costUSD||""} onChange={e=>upd("costUSD",e.target.value)}/></Field>
          </G2>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
    </div>
  );
}

// ── PERSONAL DOCS ─────────────────────────────────────────────────────────────
function PersonalDocs({data,setData}){
  const T = useT(); const S = useS();
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [search,setSearch] = useState("");
  const [collapsed,setCollapsed] = useLocalState("docs_collapsed",false);
  const [qFilter,setQFilter] = useLocalState("docs_q_filter","All");
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const {upload,download,remove} = useFiles("personalDocs",setData);
  const expiredCount = data.personalDocs.filter(d=>timeLeft(d.expiry)==="Expired").length;
  const expiringCount = data.personalDocs.filter(d=>expiryYear(d.expiry)===CY&&timeLeft(d.expiry)!=="Expired").length;
  const shown = data.personalDocs.filter(d=>{
    if(search&&!d.type.toLowerCase().includes(search.toLowerCase())) return false;
    if(qFilter==="Expired") return timeLeft(d.expiry)==="Expired";
    if(qFilter==="ExpiringThisYear") return expiryYear(d.expiry)===CY&&timeLeft(d.expiry)!=="Expired";
    return true;
  });
  function save(){const entry={...form,files:form.files||[]};setData(d=>({...d,personalDocs:modal==="add"?[...d.personalDocs,{...entry,id:nextId(d.personalDocs)}]:d.personalDocs.map(x=>x.id===form.id?entry:x)}));setModal(null);}
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>PERSONAL DOCS</div>
        <button style={{...S.btn(T.green),padding:"8px 14px"}} onClick={()=>{setForm({type:"",issued:"",expiry:"",cost:"",files:[]});setModal("add");}}>+ Add</button>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// CARDS, IDs & DOCUMENT TRACKER</div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search documents..."/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
        <StatCard label="Total" value={data.personalDocs.length} color={T.accent} onClick={()=>setQFilter("All")} active={qFilter==="All"}/>
        <StatCard label="Expired" value={expiredCount} color={T.red} onClick={()=>setQFilter(q=>q==="Expired"?"All":"Expired")} active={qFilter==="Expired"}/>
        <StatCard label={`Exp. ${CY}`} value={expiringCount} color={T.yellow} onClick={()=>setQFilter(q=>q==="ExpiringThisYear"?"All":"ExpiringThisYear")} active={qFilter==="ExpiringThisYear"}/>
      </div>
      <ListToggle collapsed={collapsed} onToggle={()=>setCollapsed(c=>!c)} count={shown.length} label="Documents"/>
      {!collapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {shown.length===0&&<div style={{...S.card,textAlign:"center",color:T.muted}}>No documents match.</div>}
          {shown.map(d=>{const tl=timeLeft(d.expiry),uc=urgencyColor(tl,T);return(
            <div key={d.id} style={{...S.card,borderColor:uc+"40"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
                <div style={{fontWeight:700,fontSize:14,flex:1}}>{d.type}</div>
                <span style={{...S.badge(uc),flexShrink:0}}>{tl}</span>
              </div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:T.muted,marginBottom:8}}>
                <span>Issued: <span style={{color:T.text}}>{fmtDate(d.issued)}</span></span>
                <span>Expires: <span style={{color:T.text}}>{fmtDate(d.expiry)}</span></span>
              </div>
              {d.cost&&<div style={{fontSize:12,color:T.accent,marginBottom:10}}>Cost: {fmt(d.cost)}</div>}
              <FilesSection files={d.files} onUpload={e=>upload(d.id,e)} onDownload={download} onRemove={name=>remove(d.id,name)}/>
              <div style={{display:"flex",justifyContent:"flex-end",marginTop:4}}>
                <ItemActions label={d.type} onEdit={()=>{setForm({...d,files:d.files||[]});setModal("edit");}} onDelete={()=>setData(d2=>({...d2,personalDocs:d2.personalDocs.filter(x=>x.id!==d.id)}))}/>
              </div>
            </div>
          );})}
        </div>
      )}
      {modal&&(
        <Modal title={modal==="add"?"Add Document":"Edit Document"} onClose={()=>setModal(null)}>
          <Field label="Document Type"><input style={S.input} value={form.type||""} onChange={e=>upd("type",e.target.value)} placeholder="e.g. Passport"/></Field>
          <G2>
            <Field label="Issue Date"><input style={S.input} type="date" value={form.issued||""} onChange={e=>upd("issued",e.target.value)}/></Field>
            <Field label="Expiry Date"><input style={S.input} type="date" value={form.expiry||""} onChange={e=>upd("expiry",e.target.value)}/></Field>
          </G2>
          <Field label="Cost"><input style={S.input} type="number" value={form.cost||""} onChange={e=>upd("cost",e.target.value)} placeholder="0.00"/></Field>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
    </div>
  );
}

// ── SUBSCRIPTIONS ─────────────────────────────────────────────────────────────
function Subscriptions({data,setData}){
  const T = useT(); const S = useS();
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [formMembers,setFormMembers] = useState([]);
  const [search,setSearch] = useState("");
  const [collapsed,setCollapsed] = useLocalState("subs_collapsed",false);
  const [mOpen,setMOpen] = useState({});
  const [editM,setEditM] = useState({});
  const [cardFilter,setCardFilter] = useLocalState("subs_card_filter","All");
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const parseM = str=>{if(!str)return[];return str.split(",").map(m=>{const p=m.trim().split(":");return{name:p[0]?.trim()||"",paid:(p[1]?.trim()||"").toLowerCase()==="paid",date:p[2]?.trim()||""};}).filter(m=>m.name);};
  const serM = arr=>arr.filter(m=>m.name.trim()).map(m=>`${m.name}:${m.paid?"Paid":"Unpaid"}:${m.date}`).join(", ");
  const userCost = sub=>{ const members=parseM(sub.members); const amt=Number(sub.amount||0); return members.length>0?amt/members.length:amt; };
  const annualSubs = data.subscriptions.filter(s=>s.type==="Annual");
  const monthlySubs = data.subscriptions.filter(s=>s.type==="Monthly");
  const annualTotal = annualSubs.reduce((s,sub)=>s+userCost(sub),0);
  const monthlyTotal = monthlySubs.reduce((s,sub)=>s+userCost(sub),0);
  const annual = annualTotal + monthlyTotal*12;
  const shown = data.subscriptions.filter(s=>{
    if(search&&!s.service.toLowerCase().includes(search.toLowerCase())) return false;
    if(cardFilter==="Annual") return s.type==="Annual";
    if(cardFilter==="Monthly") return s.type==="Monthly";
    return true;
  });
  const togglePaid = (sid,i)=>setData(d=>({...d,subscriptions:d.subscriptions.map(sub=>{if(sub.id!==sid)return sub;const ms=parseM(sub.members);ms[i]={...ms[i],paid:!ms[i].paid,date:!ms[i].paid?new Date().toISOString().slice(0,10):ms[i].date};return{...sub,members:serM(ms)};})}));
  const updMF = (sid,i,field,val)=>setData(d=>({...d,subscriptions:d.subscriptions.map(sub=>{if(sub.id!==sid)return sub;const ms=parseM(sub.members);ms[i]={...ms[i],[field]:val};return{...sub,members:serM(ms)};})}));
  function openAdd(){setForm({service:"",type:"Monthly",amount:"",renews:"",autoRenew:true,cancelled:false});setFormMembers([{name:"Me",paid:false,date:""}]);setModal("add");}
  function openEdit(sub){setForm({...sub});const ms=parseM(sub.members);if(ms.length===0||ms[0].name!=="Me")ms.unshift({name:"Me",paid:false,date:""});setFormMembers(ms);setModal("edit");}
  function save(){
    const membersStr=serM(formMembers);
    setData(d=>({...d,subscriptions:modal==="add"?[...d.subscriptions,{...form,members:membersStr,id:nextId(d.subscriptions)}]:d.subscriptions.map(s=>s.id===form.id?{...form,members:membersStr}:s)}));
    setModal(null);
  }
  function addMember(){setFormMembers(m=>[...m,{name:"",paid:false,date:""}]);}
  function removeMember(i){setFormMembers(m=>m.filter((_,idx)=>idx!==i));}
  function updMember(i,field,val){setFormMembers(m=>m.map((x,idx)=>idx===i?{...x,[field]:val}:x));}
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>SUBSCRIPTIONS</div>
        <button style={{...S.btn(T.green),padding:"8px 14px"}} onClick={openAdd}>+ Add</button>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// RECURRING SUBSCRIPTION MANAGER</div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search subscriptions..."/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
        <StatCard label="CY Subs Total" value={fmt(annual)} sub={`${data.subscriptions.length} subscription${data.subscriptions.length!==1?"s":""}`} color={T.accent} onClick={()=>setCardFilter("All")} active={cardFilter==="All"}/>
        <StatCard label="Annual Subs" value={fmt(annualTotal)} sub={`${annualSubs.length} sub${annualSubs.length!==1?"s":""}`} color={T.red} onClick={()=>setCardFilter(f=>f==="Annual"?"All":"Annual")} active={cardFilter==="Annual"}/>
        <StatCard label="Monthly Subs" value={fmt(monthlyTotal)} sub={`${monthlySubs.length} sub${monthlySubs.length!==1?"s":""}`} color={T.yellow} onClick={()=>setCardFilter(f=>f==="Monthly"?"All":"Monthly")} active={cardFilter==="Monthly"}/>
      </div>
      <ListToggle collapsed={collapsed} onToggle={()=>setCollapsed(c=>!c)} count={shown.length} label="Subscriptions"/>
      {!collapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {shown.map(sub=>{
            const er=sub.cancelled?"":effectiveRenews(sub);
            const tl=sub.cancelled?"Cancelled":timeLeft(er);
            const uc=sub.cancelled?T.muted:urgencyColor(tl,T);
            const members=parseM(sub.members);
            const myCost=userCost(sub);
            const split=members.length>0?Number(sub.amount||0)/members.length:Number(sub.amount||0);
            const isOpen=!!mOpen[sub.id];
            return(
              <div key={sub.id} style={{...S.card,borderColor:(sub.cancelled?T.muted:T.accent)+"30"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:6}}>
                  <div style={{fontWeight:700,fontSize:14,flex:1}}>{sub.service}</div>
                  <span style={{...S.badge(uc),flexShrink:0}}>{tl}</span>
                </div>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:6}}>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    <span style={S.badge(T.purple)}>{sub.type}</span>
                    <span style={S.badge(sub.autoRenew!==false?T.green:T.yellow)}>{sub.autoRenew!==false?"Auto":"Manual"}</span>
                  </div>
                  {er&&<span style={{fontSize:12}}><span style={{color:T.muted}}>Renews </span><span style={{color:T.text,fontWeight:600}}>{fmtDate(er)}</span></span>}
                </div>
                <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:2}}>
                  <div style={{fontSize:20,fontWeight:700,color:T.accent}}>{fmt(myCost)}</div>
                  {members.length>1&&<div style={{fontSize:11,color:T.muted}}>my share · {fmt(sub.amount)} total</div>}
                </div>
                <div style={{fontSize:11,color:T.muted,marginBottom:members.length>0?10:4}}>{sub.type==="Annual"?"per year":"per month"}</div>
                {members.length>0&&(
                  <div style={{borderTop:`1px solid ${T.border}30`,paddingTop:10}}>
                    <button style={S.outlineBtn(isOpen)} onClick={()=>setMOpen(p=>({...p,[sub.id]:!p[sub.id]}))}>Members ({members.length})</button>
                    {isOpen&&(
                      <div style={{marginTop:12,display:"flex",flexDirection:"column",gap:8}}>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 80px 64px 28px",gap:8,fontSize:10,color:T.muted,letterSpacing:1,padding:"0 4px",marginBottom:2}}>
                          <span>MEMBER</span><span style={{textAlign:"center"}}>STATUS</span><span style={{textAlign:"center"}}>SPLIT</span><span/>
                        </div>
                        {members.map((m,i)=>{const ek=`${sub.id}_${i}`;const isEditing=!!editM[ek];return(
                          <div key={i} style={{background:T.bg,borderRadius:8,padding:"10px 12px",border:`1px solid ${i===0?T.accent+"50":T.border}30`}}>
                            {isEditing?(
                              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                                  <input style={{...S.input,flex:1}} value={m.name} onChange={e=>updMF(sub.id,i,"name",e.target.value)} placeholder="Name"/>
                                  <button style={{...S.btn(m.paid?T.green:T.red),padding:"8px 12px",fontSize:12,flexShrink:0}} onClick={()=>togglePaid(sub.id,i)}>{m.paid?"Paid":"Unpaid"}</button>
                                </div>
                                <div style={{display:"flex",gap:8,alignItems:"center"}}>
                                  <span style={{fontSize:11,color:T.muted,whiteSpace:"nowrap"}}>Paid on:</span>
                                  <input style={{...S.input,flex:1}} type="date" value={m.date||""} onChange={e=>updMF(sub.id,i,"date",e.target.value)}/>
                                  <button style={{...S.btn(T.accent),padding:"8px 12px",fontSize:12,flexShrink:0}} onClick={()=>setEditM(p=>({...p,[ek]:false}))}>Done</button>
                                </div>
                              </div>
                            ):(
                              <div style={{display:"grid",gridTemplateColumns:"1fr 80px 64px 28px",gap:8,alignItems:"center"}}>
                                <div>
                                  <div style={{fontWeight:600,fontSize:13,color:i===0?T.accent:T.text}}>{m.name}{i===0&&" (You)"}</div>
                                  {m.date&&<div style={{fontSize:10,color:T.muted}}>Paid: {fmtDate(m.date)}</div>}
                                </div>
                                <div style={{display:"flex",justifyContent:"center"}}>
                                  <button style={{...S.badge(m.paid?T.green:T.red),cursor:"pointer",border:`1px solid ${m.paid?T.green:T.red}60`,background:"none"}} onClick={()=>togglePaid(sub.id,i)}>{m.paid?"Paid":"Unpaid"}</button>
                                </div>
                                <span style={{color:T.accent,fontSize:12,textAlign:"center"}}>{fmt(split)}</span>
                                <button style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:14,padding:"0 4px"}} onClick={()=>setEditM(p=>({...p,[ek]:true}))}>✏</button>
                              </div>
                            )}
                          </div>
                        );})}
                      </div>
                    )}
                  </div>
                )}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8}}>
                  <button onClick={()=>setData(d=>({...d,subscriptions:d.subscriptions.map(s=>s.id===sub.id?{...s,cancelled:!s.cancelled}:s)}))} style={{...S.btn(sub.cancelled?T.green:T.muted),padding:"6px 14px",fontSize:12}}>
                    {sub.cancelled?"Reactivate":"Cancel"}
                  </button>
                  <ItemActions label={sub.service} onEdit={()=>openEdit(sub)} onDelete={()=>setData(d=>({...d,subscriptions:d.subscriptions.filter(x=>x.id!==sub.id)}))}/>
                </div>
              </div>
            );
          })}
        </div>
      )}
      {modal&&(
        <Modal title={modal==="add"?"Add Subscription":"Edit Subscription"} onClose={()=>setModal(null)}>
          <Field label="Service Name"><input style={S.input} value={form.service||""} onChange={e=>upd("service",e.target.value)} placeholder="e.g. Netflix"/></Field>
          <G2>
            <Field label="Type"><select style={S.input} value={form.type||"Monthly"} onChange={e=>upd("type",e.target.value)}><option>Monthly</option><option>Annual</option></select></Field>
            <Field label="Amount"><input style={S.input} type="number" value={form.amount||""} onChange={e=>upd("amount",e.target.value)} placeholder="0.00"/></Field>
          </G2>
          <Field label="Renewal Date"><input style={S.input} type="date" value={form.renews||""} onChange={e=>upd("renews",e.target.value)}/></Field>
          <Field label="Renewal Plan">
            <div style={{display:"flex",gap:8}}>
              <button type="button" style={{...S.outlineBtn(form.autoRenew!==false),flex:1}} onClick={()=>upd("autoRenew",true)}>Auto-Renew</button>
              <button type="button" style={{...S.outlineBtn(form.autoRenew===false),flex:1}} onClick={()=>upd("autoRenew",false)}>Manual Renew</button>
            </div>
          </Field>
          <div style={{marginBottom:14}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
              <label style={S.label}>Members</label>
              <button type="button" style={{...S.btn(T.accent),padding:"4px 12px",fontSize:12}} onClick={addMember}>+ Add Member</button>
            </div>
            {formMembers.length===0&&<div style={{fontSize:12,color:T.muted,padding:"8px 0"}}>No members added.</div>}
            {formMembers.map((m,i)=>(
              <div key={i} style={{background:T.bg,borderRadius:8,padding:"10px 12px",border:`1px solid ${i===0?T.accent+"50":T.border+"50"}`,marginBottom:8}}>
                <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:m.paid?8:0}}>
                  {i===0
                    ? <span style={{flex:1,fontSize:13,fontWeight:700,color:T.accent}}>Me (You)</span>
                    : <input style={{...S.input,flex:1}} value={m.name} onChange={e=>updMember(i,"name",e.target.value)} placeholder="Member name"/>
                  }
                  <button type="button" style={{...S.btn(m.paid?T.green:T.red),padding:"8px 12px",fontSize:12,flexShrink:0}} onClick={()=>updMember(i,"paid",!m.paid)}>{m.paid?"Paid":"Unpaid"}</button>
                  {i>0&&<button type="button" style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:16,padding:"0 4px",flexShrink:0}} onClick={()=>removeMember(i)}>✕</button>}
                </div>
                {m.paid&&(
                  <div style={{display:"flex",gap:8,alignItems:"center",marginTop:8}}>
                    <span style={{fontSize:11,color:T.muted,whiteSpace:"nowrap"}}>Paid on:</span>
                    <input style={{...S.input,flex:1}} type="date" value={m.date||""} onChange={e=>updMember(i,"date",e.target.value)}/>
                  </div>
                )}
              </div>
            ))}
          </div>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
    </div>
  );
}

// ── CAR MAINTENANCE ───────────────────────────────────────────────────────────
function CarMaintenance({data,setData}){
  const T = useT(); const S = useS();
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [search,setSearch] = useState("");
  const [fAction,setFAction] = useLocalState("car_filter_action","All");
  const [fVehicle,setFVehicle] = useLocalState("car_filter_vehicle","All");
  const [fYear,setFYear] = useLocalState("car_filter_year","All");
  const [logCollapsed,setLogCollapsed] = useLocalState("car_log_collapsed",false);
  const [vehiclesOpen,setVehiclesOpen] = useLocalState("car_vehicles_open",true);
  const [expV,setExpV] = useState({});
  const [confDelV,setConfDelV] = useState(null);
  const [vModal,setVModal] = useState(false);
  const [vForm,setVForm] = useState({});
  const [activeVE,setActiveVE] = useState(null);
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const vehicles = data.vehicles||[];
  const actions = ["All","Replaced","Serviced","Inspected","Pending"];
  const logYears = ["All",...[...new Set(data.carLog.map(r=>logYear(r.date)).filter(Boolean))].sort((a,b)=>b-a).map(String)];
  const aColor = a=>({Replaced:T.green,Serviced:T.accent,Inspected:T.purple,Pending:T.yellow}[a]||T.muted);
  const shown = [...data.carLog].reverse().filter(r=>{
    if(search&&!r.item.toLowerCase().includes(search.toLowerCase())&&!(r.supplier||"").toLowerCase().includes(search.toLowerCase())) return false;
    if(fAction!=="All"&&r.action!==fAction) return false;
    if(fVehicle!=="All"&&r.vehicleId!==fVehicle) return false;
    if(fYear!=="All"&&String(logYear(r.date))!==fYear) return false;
    return true;
  });
  const shownTotal = shown.reduce((s,r)=>s+Number(r.cost||0),0);
  const anyFilter = search||fAction!=="All"||fVehicle!=="All"||fYear!=="All";
  function save(){const entry={...form,vehicleId:form.vehicleId||(vehicles[0]?.id||"")};setData(d=>({...d,carLog:modal==="add"?[...d.carLog,{...entry,id:nextId(d.carLog)}]:d.carLog.map(r=>r.id===form.id?entry:r)}));setModal(null);}
  function saveV(){let vf={...vForm};if(!vf.id)vf.id="CAR-"+Date.now();if(activeVE){const o=activeVE,n=vf.id;setData(d=>({...d,vehicles:d.vehicles.map(v=>v.id===o?vf:v),carLog:n!==o?d.carLog.map(r=>r.vehicleId===o?{...r,vehicleId:n}:r):d.carLog}));if(fVehicle===o)setFVehicle(n);}else{setData(d=>({...d,vehicles:[...(d.vehicles||[]),vf]}));}setVModal(false);setActiveVE(null);}
  function delV(id){setData(d=>({...d,vehicles:d.vehicles.filter(v=>v.id!==id)}));if(fVehicle===id)setFVehicle("All");setConfDelV(null);}
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>CAR MAINTENANCE</div>
        <div style={{display:"flex",gap:8}}>
          <button style={{...S.btn(T.purple),padding:"8px 12px"}} onClick={()=>{setVForm({id:"",label:"",lastServicedDate:"",lastServicedMileage:""});setActiveVE(null);setVModal(true);}}>+ Vehicle</button>
          <button style={{...S.btn(T.green),padding:"8px 12px"}} onClick={()=>{setForm({item:"",action:"Replaced",date:"",mileage:"",cost:"",supplier:"",vehicleId:vehicles[0]?.id||""});setModal("add");}}>+ Entry</button>
        </div>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// VEHICLE MAINTENANCE TRACKER</div>
      {vehicles.length>0&&(
        <div style={{...S.card,padding:0,overflow:"hidden",marginBottom:14}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px",cursor:"pointer",borderBottom:vehiclesOpen?`1px solid ${T.border}30`:"none"}} onClick={()=>setVehiclesOpen(o=>!o)}>
            <div style={{fontSize:12,letterSpacing:2,color:T.purple,textTransform:"uppercase",fontWeight:700}}>Vehicles ({vehicles.length})</div>
            <span style={{color:T.muted,fontSize:13}}>{vehiclesOpen?"▲":"▼"}</span>
          </div>
          {vehiclesOpen&&(
            <div>
              {vehicles.map((v,vi)=>{
                const ie=!!expV[v.id];
                return(
                  <div key={v.id} style={{padding:"12px 16px",borderBottom:vi<vehicles.length-1?`1px solid ${T.border}20`:"none",background:fVehicle===v.id?T.purple+"10":"transparent",cursor:"pointer"}} onClick={()=>setFVehicle(fv=>fv===v.id?"All":v.id)}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                      <div style={{flex:1}}>
                        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,flexWrap:"wrap"}}>
                          <span style={{fontWeight:700,color:T.purple,fontSize:14}}>🚗 {v.id}</span>
                          {v.label&&v.label!==v.id&&<span style={{fontSize:12,color:T.muted}}>— {v.label}</span>}
                          {fVehicle===v.id&&<span style={{...S.badge(T.purple),fontSize:9}}>ACTIVE</span>}
                        </div>
                        {v.lastServicedDate&&(
                          <div style={{fontSize:12,color:T.muted,display:"flex",gap:12,flexWrap:"wrap"}}>
                            <span>Last: <span style={{color:T.text}}>{fmtDate(v.lastServicedDate)}</span></span>
                            {v.lastServicedMileage&&<span>Mileage: <span style={{color:T.accent}}>{Number(v.lastServicedMileage).toLocaleString()} km</span></span>}
                          </div>
                        )}
                        {ie&&(
                          <div style={{display:"flex",gap:8,marginTop:10}} onClick={e=>e.stopPropagation()}>
                            <button style={{...S.btn(T.accent),padding:"6px 14px",fontSize:12}} onClick={e=>{e.stopPropagation();setVForm({...v});setActiveVE(v.id);setVModal(true);setExpV(p=>({...p,[v.id]:false}));}}>✏ Edit</button>
                            <button style={{...S.btn(T.red),padding:"6px 14px",fontSize:12}} onClick={e=>{e.stopPropagation();setConfDelV(v.id);}}>🗑 Delete</button>
                          </div>
                        )}
                      </div>
                      <button style={{background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:14,padding:"4px 8px",flexShrink:0}} onClick={e=>{e.stopPropagation();setExpV(p=>({...p,[v.id]:!p[v.id]}));}}>{ie?"▲":"▼"}</button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
        <StatCard label="Total Spent" value={fmt(shownTotal)} color={T.red}/>
        <StatCard label="Log Entries" value={shown.length} color={T.accent}/>
      </div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search items or suppliers..."/>
      <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:10}}>
        <select style={{...S.input,flex:"1 1 120px",minWidth:100}} value={fAction} onChange={e=>setFAction(e.target.value)}>{actions.map(a=><option key={a}>{a}</option>)}</select>
        <select style={{...S.input,flex:"1 1 120px",minWidth:100}} value={fVehicle} onChange={e=>setFVehicle(e.target.value)}>{["All",...vehicles.map(v=>v.id)].map(v=><option key={v}>{v}</option>)}</select>
        <select style={{...S.input,flex:"1 1 100px",minWidth:80}} value={fYear} onChange={e=>setFYear(e.target.value)}>{logYears.map(y=><option key={y}>{y}</option>)}</select>
        {anyFilter&&<button style={{...S.btn(T.muted),padding:"8px 14px"}} onClick={()=>{setSearch("");setFAction("All");setFVehicle("All");setFYear("All");}}>✕ Clear</button>}
      </div>
      {anyFilter&&(
        <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>
          {search&&<span style={S.badge(T.accent)}>"{search}"</span>}
          {fAction!=="All"&&<span style={S.badge(aColor(fAction))}>{fAction}</span>}
          {fVehicle!=="All"&&<span style={S.badge(T.purple)}>{fVehicle}</span>}
          {fYear!=="All"&&<span style={S.badge(T.yellow)}>{fYear}</span>}
          <span style={{fontSize:11,color:T.muted,alignSelf:"center"}}>{shown.length} result{shown.length!==1?"s":""}</span>
        </div>
      )}
      <ListToggle collapsed={logCollapsed} onToggle={()=>setLogCollapsed(c=>!c)} count={shown.length} label="Maintenance Log"/>
      {!logCollapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {shown.length===0&&<div style={{...S.card,textAlign:"center",color:T.muted}}>No entries match.</div>}
          {shown.map(r=>{
            const vInfo=vehicles.find(v=>v.id===r.vehicleId);
            const vLabel=vInfo?.label&&vInfo.label!==vInfo.id?vInfo.label:null;
            return(
              <div key={r.id} style={S.card}>
                <div style={{display:"flex",justifyContent:"space-between",gap:12}}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:700,fontSize:14,marginBottom:6}}>{r.item}</div>
                    <div style={{fontSize:12,display:"flex",flexDirection:"column",gap:3}}>
                      <span style={{color:r.supplier?T.muted:T.border}}>📦 {r.supplier||"No supplier"}</span>
                      <span style={{color:r.cost?T.accent:T.border}}>💰 {r.cost?fmt(r.cost):"No cost logged"}</span>
                      {vLabel&&<span style={{color:T.purple,fontSize:11}}>🚗 {vLabel}</span>}
                    </div>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4,flexShrink:0}}>
                    <span style={S.badge(aColor(r.action))}>{r.action}</span>
                    <div style={{fontSize:12,color:T.muted,textAlign:"right"}}>
                      {r.date&&<div>{fmtDate(r.date)}</div>}
                      {r.mileage&&<div style={{color:T.accent}}>{Number(r.mileage).toLocaleString()} km</div>}
                    </div>
                    <ItemActions label={r.item} onEdit={()=>{setForm({...r});setModal("edit");}} onDelete={()=>setData(d=>({...d,carLog:d.carLog.filter(x=>x.id!==r.id)}))}/>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
      {confDelV&&<ConfirmDelete label={confDelV} onCancel={()=>setConfDelV(null)} onConfirm={()=>delV(confDelV)}/>}
      {modal&&(
        <Modal title={modal==="add"?"Add Log Entry":"Edit Log Entry"} onClose={()=>setModal(null)}>
          {vehicles.length>0&&<Field label="Vehicle"><select style={S.input} value={form.vehicleId||""} onChange={e=>upd("vehicleId",e.target.value)}>{vehicles.map(v=><option key={v.id} value={v.id}>{v.id}{v.label&&v.label!==v.id?` — ${v.label}`:""}</option>)}</select></Field>}
          <Field label="Item / Service"><input style={S.input} value={form.item||""} onChange={e=>upd("item",e.target.value)} placeholder="e.g. Engine Oil"/></Field>
          <G2>
            <Field label="Action"><select style={S.input} value={form.action||"Replaced"} onChange={e=>upd("action",e.target.value)}><option>Replaced</option><option>Serviced</option><option>Inspected</option><option>Pending</option></select></Field>
            <Field label="Date"><input style={S.input} type="date" value={form.date||""} onChange={e=>upd("date",e.target.value)}/></Field>
          </G2>
          <G2>
            <Field label="Mileage (km)"><input style={S.input} type="number" value={form.mileage||""} onChange={e=>upd("mileage",e.target.value)} placeholder="141767"/></Field>
            <Field label="Cost"><input style={S.input} type="number" value={form.cost||""} onChange={e=>upd("cost",e.target.value)} placeholder="0.00"/></Field>
          </G2>
          <Field label="Supplier"><input style={S.input} value={form.supplier||""} onChange={e=>upd("supplier",e.target.value)} placeholder="e.g. Massy"/></Field>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
      {vModal&&(
        <Modal title={activeVE?"Edit Vehicle":"Add Vehicle"} onClose={()=>{setVModal(false);setActiveVE(null);}}>
          <Field label="Vehicle ID"><input style={S.input} value={vForm.id||""} onChange={e=>setVForm(f=>({...f,id:e.target.value}))} placeholder="CAR-002"/></Field>
          <Field label="Label / Nickname"><input style={S.input} value={vForm.label||""} onChange={e=>setVForm(f=>({...f,label:e.target.value}))} placeholder="e.g. Toyota Hilux"/></Field>
          <G2>
            <Field label="Last Serviced"><input style={S.input} type="date" value={vForm.lastServicedDate||""} onChange={e=>setVForm(f=>({...f,lastServicedDate:e.target.value}))}/></Field>
            <Field label="Mileage (km)"><input style={S.input} type="number" value={vForm.lastServicedMileage||""} onChange={e=>setVForm(f=>({...f,lastServicedMileage:e.target.value}))}/></Field>
          </G2>
          <SaveCancel onCancel={()=>{setVModal(false);setActiveVE(null);}} onSave={saveV}/>
        </Modal>
      )}
    </div>
  );
}

// ── LEISURE — LegCard (top-level) ─────────────────────────────────────────────
function LegCard({leg,onEdit,onDelete}){
  const T = useT(); const S = useS();
  const typeColors = {"One Way":T.yellow,"Round Trip":T.accent,"Multi-City":T.purple};
  const tc = typeColors[leg.flightType]||T.muted;
  const segs = leg.segments||[];
  const routeText = leg.flightType==="Multi-City"
    ? segs.map(s=>s.from).concat(segs.length>0?[segs[segs.length-1].to]:[]).join(" / ")
    : leg.flightType==="Round Trip"
      ? `${leg.from||"?"} / ${leg.to||"?"} / ${leg.from||"?"}`
      : `${leg.from||"?"} / ${leg.to||"?"}`;
  return(
    <div style={{background:T.bg,borderRadius:10,padding:"12px",border:`1px solid ${T.border}30`,marginBottom:8}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
        <div style={{fontWeight:700,fontSize:13,color:T.text,flex:1,lineHeight:1.3}}>{routeText}</div>
        <span style={{...S.badge(tc),fontSize:10,flexShrink:0}}>{leg.flightType}</span>
      </div>
      {(leg.airline||leg.flightNo)&&(
        <div style={{fontSize:12,color:T.muted,marginBottom:8}}>
          {leg.airline&&<span>✈ {leg.airline}</span>}
          {leg.airline&&leg.flightNo&&<span style={{margin:"0 6px"}}>·</span>}
          {leg.flightNo&&<span>#{leg.flightNo}</span>}
        </div>
      )}
      <div style={{fontSize:12,color:T.muted,display:"flex",flexDirection:"column",gap:5,marginBottom:10}}>
        {leg.flightType==="One Way"&&(
          <div>
            <div><span style={{color:T.accent,fontWeight:600}}>Depart: </span><span style={{color:T.text}}>{leg.from||"—"}</span>{leg.departDate&&<span> · {fmtDate(leg.departDate)}{leg.departTime&&` ${leg.departTime}`}</span>}</div>
            <div><span style={{color:T.purple,fontWeight:600}}>Arrive: </span><span style={{color:T.text}}>{leg.to||"—"}</span>{leg.arriveDate&&<span> · {fmtDate(leg.arriveDate)}{leg.arriveTime&&` ${leg.arriveTime}`}</span>}</div>
          </div>
        )}
        {leg.flightType==="Round Trip"&&(
          <div>
            <div style={{fontSize:10,color:T.accent,letterSpacing:1,marginBottom:4,fontWeight:700}}>OUTBOUND</div>
            <div><span style={{color:T.accent,fontWeight:600}}>Depart: </span><span style={{color:T.text}}>{leg.from||"—"}</span>{leg.departDate&&<span> · {fmtDate(leg.departDate)}{leg.departTime&&` ${leg.departTime}`}</span>}</div>
            <div><span style={{color:T.purple,fontWeight:600}}>Arrive: </span><span style={{color:T.text}}>{leg.to||"—"}</span>{leg.arriveDate&&<span> · {fmtDate(leg.arriveDate)}{leg.arriveTime&&` ${leg.arriveTime}`}</span>}</div>
            {leg.returnDate&&(
              <div style={{marginTop:6}}>
                <div style={{fontSize:10,color:T.purple,letterSpacing:1,marginBottom:4,fontWeight:700}}>RETURN</div>
                <div><span style={{color:T.accent,fontWeight:600}}>Depart: </span><span style={{color:T.text}}>{leg.to||"—"}</span><span> · {fmtDate(leg.returnDate)}{leg.returnTime&&` ${leg.returnTime}`}</span></div>
                <div><span style={{color:T.purple,fontWeight:600}}>Arrive: </span><span style={{color:T.text}}>{leg.from||"—"}</span>{leg.returnArriveDate&&<span> · {fmtDate(leg.returnArriveDate)}{leg.returnArriveTime&&` ${leg.returnArriveTime}`}</span>}</div>
              </div>
            )}
          </div>
        )}
        {leg.flightType==="Multi-City"&&segs.map((seg,i)=>(
          <div key={seg.id} style={{paddingLeft:10,borderLeft:`2px solid ${T.purple}40`,marginBottom:4}}>
            <div style={{fontSize:10,color:T.purple,fontWeight:700,marginBottom:2}}>SEG {i+1}</div>
            <div><span style={{color:T.accent,fontWeight:600}}>Depart: </span><span style={{color:T.text}}>{seg.from||"—"}</span>{seg.departDate&&<span> · {fmtDate(seg.departDate)}{seg.departTime&&` ${seg.departTime}`}</span>}</div>
            <div><span style={{color:T.purple,fontWeight:600}}>Arrive: </span><span style={{color:T.text}}>{seg.to||"—"}</span>{seg.arriveDate&&<span> · {fmtDate(seg.arriveDate)}{seg.arriveTime&&` ${seg.arriveTime}`}</span>}</div>
          </div>
        ))}
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",borderTop:`1px solid ${T.border}20`,paddingTop:8}}>
        <span style={{fontSize:14,fontWeight:700,color:leg.flightCost?T.green:T.muted}}>{leg.flightCost?fmt(leg.flightCost):"No cost"}</span>
        <ItemActions label={leg.flightType+" flight"} onEdit={onEdit} onDelete={onDelete}/>
      </div>
    </div>
  );
}

function FlightModal({initial,onSave,onClose}){
  const T = useT(); const S = useS();
  const blank = {flightType:"One Way",from:"",to:"",departDate:"",departTime:"",arriveDate:"",arriveTime:"",returnDate:"",returnTime:"",returnArriveDate:"",returnArriveTime:"",flightCost:"",airline:"",flightNo:"",segments:[]};
  const [f,setF] = useState(initial?{...blank,...initial}:blank);
  const upd = (k,v)=>setF(p=>({...p,[k]:v}));
  const typeColors = {"One Way":T.yellow,"Round Trip":T.accent,"Multi-City":T.purple};
  const addSeg = ()=>setF(p=>({...p,segments:[...(p.segments||[]),{id:Date.now(),from:"",to:"",departDate:"",departTime:"",arriveDate:"",arriveTime:""}]}));
  const updSeg = (id,k,v)=>setF(p=>({...p,segments:p.segments.map(s=>s.id===id?{...s,[k]:v}:s)}));
  const delSeg = id=>setF(p=>({...p,segments:p.segments.filter(s=>s.id!==id)}));
  return(
    <Modal title={initial?"Edit Flight":"Add Flight"} onClose={onClose}>
      <div style={{marginBottom:16}}>
        <div style={S.label}>Flight Type</div>
        <div style={{display:"flex",gap:8}}>
          {["One Way","Round Trip","Multi-City"].map(type=>(
            <button key={type} onClick={()=>upd("flightType",type)} style={{flex:1,padding:"10px 4px",borderRadius:8,fontSize:12,fontFamily:"inherit",cursor:"pointer",fontWeight:700,border:`1px solid ${f.flightType===type?typeColors[type]:T.border}`,background:f.flightType===type?typeColors[type]+"20":"none",color:f.flightType===type?typeColors[type]:T.muted}}>
              {type}
            </button>
          ))}
        </div>
      </div>
      <G2>
        <Field label="Airline"><input style={S.input} value={f.airline||""} onChange={e=>upd("airline",e.target.value)} placeholder="Caribbean Airlines"/></Field>
        <Field label="Flight No."><input style={S.input} value={f.flightNo||""} onChange={e=>upd("flightNo",e.target.value)} placeholder="BW400"/></Field>
      </G2>
      <Field label="Total Flight Cost (TTD)"><input style={S.input} type="number" value={f.flightCost||""} onChange={e=>upd("flightCost",e.target.value)} placeholder="0.00"/></Field>
      {f.flightType==="One Way"&&(
        <div>
          <G2><Field label="From"><input style={S.input} value={f.from||""} onChange={e=>upd("from",e.target.value)} placeholder="Port of Spain"/></Field><Field label="To"><input style={S.input} value={f.to||""} onChange={e=>upd("to",e.target.value)} placeholder="New York"/></Field></G2>
          <G2><Field label="Depart Date"><input style={S.input} type="date" value={f.departDate||""} onChange={e=>upd("departDate",e.target.value)}/></Field><Field label="Depart Time"><input style={S.input} type="time" value={f.departTime||""} onChange={e=>upd("departTime",e.target.value)}/></Field></G2>
          <G2><Field label="Arrive Date"><input style={S.input} type="date" value={f.arriveDate||""} onChange={e=>upd("arriveDate",e.target.value)}/></Field><Field label="Arrive Time"><input style={S.input} type="time" value={f.arriveTime||""} onChange={e=>upd("arriveTime",e.target.value)}/></Field></G2>
        </div>
      )}
      {f.flightType==="Round Trip"&&(
        <div>
          <G2><Field label="Origin"><input style={S.input} value={f.from||""} onChange={e=>upd("from",e.target.value)} placeholder="Port of Spain"/></Field><Field label="Destination"><input style={S.input} value={f.to||""} onChange={e=>upd("to",e.target.value)} placeholder="New York"/></Field></G2>
          <div style={{fontSize:11,color:T.accent,letterSpacing:1,marginBottom:10,fontWeight:700}}>OUTBOUND</div>
          <G2><Field label="Depart Date"><input style={S.input} type="date" value={f.departDate||""} onChange={e=>upd("departDate",e.target.value)}/></Field><Field label="Depart Time"><input style={S.input} type="time" value={f.departTime||""} onChange={e=>upd("departTime",e.target.value)}/></Field></G2>
          <G2><Field label="Arrive Date"><input style={S.input} type="date" value={f.arriveDate||""} onChange={e=>upd("arriveDate",e.target.value)}/></Field><Field label="Arrive Time"><input style={S.input} type="time" value={f.arriveTime||""} onChange={e=>upd("arriveTime",e.target.value)}/></Field></G2>
          <div style={{fontSize:11,color:T.purple,letterSpacing:1,marginBottom:10,fontWeight:700}}>RETURN ({f.to||"Dest"} to {f.from||"Origin"})</div>
          <G2><Field label="Return Depart Date"><input style={S.input} type="date" value={f.returnDate||""} onChange={e=>upd("returnDate",e.target.value)}/></Field><Field label="Return Depart Time"><input style={S.input} type="time" value={f.returnTime||""} onChange={e=>upd("returnTime",e.target.value)}/></Field></G2>
          <G2><Field label="Return Arrive Date"><input style={S.input} type="date" value={f.returnArriveDate||""} onChange={e=>upd("returnArriveDate",e.target.value)}/></Field><Field label="Return Arrive Time"><input style={S.input} type="time" value={f.returnArriveTime||""} onChange={e=>upd("returnArriveTime",e.target.value)}/></Field></G2>
        </div>
      )}
      {f.flightType==="Multi-City"&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:11,color:T.purple,fontWeight:700}}>SEGMENTS</div>
            <button style={{...S.btn(T.purple),padding:"6px 12px",fontSize:12}} onClick={addSeg}>+ Add Segment</button>
          </div>
          {(f.segments||[]).length===0&&<div style={{fontSize:12,color:T.muted,marginBottom:10}}>No segments yet.</div>}
          {(f.segments||[]).map((seg,i)=>(
            <div key={seg.id} style={{background:T.surface,borderRadius:8,padding:"12px",marginBottom:10,border:`1px solid ${T.border}30`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{fontSize:11,color:T.purple,fontWeight:700}}>SEG {i+1}</div>
                <button style={{background:"none",border:"none",color:T.red,cursor:"pointer",fontSize:16}} onClick={()=>delSeg(seg.id)}>✕</button>
              </div>
              <G2><Field label="From"><input style={S.input} value={seg.from||""} onChange={e=>updSeg(seg.id,"from",e.target.value)} placeholder="City"/></Field><Field label="To"><input style={S.input} value={seg.to||""} onChange={e=>updSeg(seg.id,"to",e.target.value)} placeholder="City"/></Field></G2>
              <G2><Field label="Depart Date"><input style={S.input} type="date" value={seg.departDate||""} onChange={e=>updSeg(seg.id,"departDate",e.target.value)}/></Field><Field label="Depart Time"><input style={S.input} type="time" value={seg.departTime||""} onChange={e=>updSeg(seg.id,"departTime",e.target.value)}/></Field></G2>
              <G2><Field label="Arrive Date"><input style={S.input} type="date" value={seg.arriveDate||""} onChange={e=>updSeg(seg.id,"arriveDate",e.target.value)}/></Field><Field label="Arrive Time"><input style={S.input} type="time" value={seg.arriveTime||""} onChange={e=>updSeg(seg.id,"arriveTime",e.target.value)}/></Field></G2>
            </div>
          ))}
        </div>
      )}
      <SaveCancel onCancel={onClose} onSave={()=>onSave({...f,id:initial?.id||Date.now()})}/>
    </Modal>
  );
}

// ── TRIP DETAIL ───────────────────────────────────────────────────────────────
function TripDetail({trip,onBack,onSave}){
  const T = useT(); const S = useS();
  const [form,setForm] = useState({...trip,legs:trip.legs||[],accommodations:trip.accommodations||[],files:trip.files||[]});
  const [fModal,setFModal] = useState(null);
  const [aModal,setAModal] = useState(null);
  const [aForm,setAForm] = useState({});
  const [filesOpen,setFilesOpen] = useState(false);
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const updA = (k,v)=>setAForm(f=>({...f,[k]:v}));
  const statColors = {Booked:T.accent,Planning:T.muted};
  const getFirst = leg=>{if(!leg)return null;if(leg.flightType==="Multi-City")return leg.segments?.[0]?.departDate||null;return leg.departDate||null;};
  const getLast = leg=>{if(!leg)return null;if(leg.flightType==="Multi-City"){const s=leg.segments||[];return s[s.length-1]?.arriveDate||s[s.length-1]?.departDate||null;}if(leg.flightType==="Round Trip")return leg.returnArriveDate||leg.returnDate||null;return leg.arriveDate||leg.departDate||null;};
  const sorted = [...form.legs].sort((a,b)=>(getFirst(a)||"").localeCompare(getFirst(b)||""));
  const autoStart = sorted.length>0?getFirst(sorted[0]):null;
  const autoEnd = sorted.length>0?getLast(sorted[sorted.length-1]):null;
  const totalF = form.legs.reduce((s,l)=>s+Number(l.flightCost||0),0);
  const totalA = form.accommodations.reduce((s,a)=>s+Number(a.cost||0),0);
  function saveLeg(leg){setForm(f=>({...f,legs:fModal==="add"?[...f.legs,leg]:f.legs.map(l=>l.id===leg.id?leg:l)}));setFModal(null);}
  function saveAcc(){const entry={...aForm,id:aModal==="add"?nextId(form.accommodations):aForm.id};setForm(f=>({...f,accommodations:aModal==="add"?[...f.accommodations,entry]:f.accommodations.map(a=>a.id===entry.id?entry:a)}));setAModal(null);}
  function uploadFile(e){Array.from(e.target.files).forEach(file=>{const r=new FileReader();r.onload=ev=>setForm(f=>({...f,files:[...f.files,{name:file.name,type:file.type,size:file.size,dataUrl:ev.target.result,uploadedAt:new Date().toLocaleDateString()}]}));r.readAsDataURL(file);});e.target.value="";}
  function dlFile(f){const a=document.createElement("a");a.href=f.dataUrl;a.download=f.name;a.click();}
  return(
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:4}}>
        <button style={{background:"none",border:"none",color:T.accent,cursor:"pointer",fontSize:22,padding:"0 4px",lineHeight:1}} onClick={()=>{onSave(form);onBack();}}>←</button>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:1}}>{form.trip||"TRIP DETAILS"}</div>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:16}}>// ← tap arrow to save & return</div>
      <div style={S.card}>
        <G2>
          <Field label="Trip Name"><input style={S.input} value={form.trip||""} onChange={e=>upd("trip",e.target.value)} placeholder="e.g. World Cup 2026"/></Field>
          <Field label="Status"><select style={S.input} value={form.status||"Planning"} onChange={e=>upd("status",e.target.value)}><option>Planning</option><option>Booked</option></select></Field>
        </G2>
        <G2>
          <div>
            <div style={S.label}>Start <span style={{color:T.accent,fontSize:9}}>(auto)</span></div>
            <div style={{...S.input,color:autoStart?T.text:T.muted,background:T.bg,cursor:"default",fontSize:13}}>{autoStart?fmtDate(autoStart):"From first flight"}</div>
          </div>
          <div>
            <div style={S.label}>End <span style={{color:T.accent,fontSize:9}}>(auto)</span></div>
            <div style={{...S.input,color:autoEnd?T.text:T.muted,background:T.bg,cursor:"default",fontSize:13}}>{autoEnd?fmtDate(autoEnd):"From last flight"}</div>
          </div>
        </G2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:4}}>
          <span style={S.badge(statColors[form.status]||T.muted)}>{form.status}</span>
          {(autoStart||autoEnd)&&<span style={{fontSize:12,color:T.muted}}>{fmtDate(autoStart)||"?"} → {fmtDate(autoEnd)||"?"}</span>}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:12}}>
        <StatCard label="Flights" value={fmt(totalF)} color={T.accent}/>
        <StatCard label="Stays" value={fmt(totalA)} color={T.purple}/>
        <StatCard label="Total" value={fmt(totalF+totalA)} color={T.green}/>
      </div>
      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:12,letterSpacing:2,color:T.accent,textTransform:"uppercase",fontWeight:700}}>✈ Flights</div>
          <button style={{...S.btn(T.green),padding:"6px 14px",fontSize:12}} onClick={()=>setFModal("add")}>+ Add Flight</button>
        </div>
        {form.legs.length===0&&<div style={{fontSize:12,color:T.muted,textAlign:"center",padding:"8px 0"}}>No flights added yet.</div>}
        {form.legs.map(leg=>(
          <LegCard key={leg.id} leg={leg} onEdit={()=>setFModal(leg)} onDelete={()=>setForm(f=>({...f,legs:f.legs.filter(l=>l.id!==leg.id)}))}/>
        ))}
      </div>
      <div style={S.card}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:12,letterSpacing:2,color:T.accent,textTransform:"uppercase",fontWeight:700}}>🏨 Accommodation</div>
          <button style={{...S.btn(T.green),padding:"6px 14px",fontSize:12}} onClick={()=>{setAForm({city:"",hotel:"",checkIn:"",checkOut:"",checkInTime:"",cost:""});setAModal("add");}}>+ Add Stay</button>
        </div>
        {form.accommodations.length===0&&<div style={{fontSize:12,color:T.muted,textAlign:"center",padding:"8px 0"}}>No accommodation added yet.</div>}
        {form.accommodations.map(acc=>(
          <div key={acc.id} style={{background:T.bg,borderRadius:10,padding:"12px",border:`1px solid ${T.border}30`,marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:13,marginBottom:6}}>{acc.hotel||"Accommodation"} <span style={{color:T.muted,fontWeight:400,fontSize:12}}>· {acc.city}</span></div>
                <div style={{fontSize:12,color:T.muted,display:"flex",flexDirection:"column",gap:3}}>
                  {acc.checkIn&&<span>In: {fmtDate(acc.checkIn)}{acc.checkInTime?` ${acc.checkInTime}`:""}</span>}
                  {acc.checkOut&&<span>Out: {fmtDate(acc.checkOut)}</span>}
                  {acc.cost&&<span style={{color:T.purple}}>💰 {fmt(acc.cost)}</span>}
                </div>
              </div>
              <ItemActions label={acc.hotel||acc.city} onEdit={()=>{setAForm({...acc});setAModal("edit");}} onDelete={()=>setForm(f=>({...f,accommodations:f.accommodations.filter(a=>a.id!==acc.id)}))}/>
            </div>
          </div>
        ))}
      </div>
      <div style={S.card}>
        <button style={S.outlineBtn(filesOpen)} onClick={()=>setFilesOpen(o=>!o)}>Attached Files ({form.files.length})</button>
        {filesOpen&&(
          <div style={{marginTop:10}}>
            <label style={{...S.btn(T.purple),padding:"6px 14px",fontSize:12,cursor:"pointer",display:"inline-block",marginBottom:10}}>+ Upload<input type="file" multiple style={{display:"none"}} onChange={uploadFile}/></label>
            {form.files.length===0&&<div style={{fontSize:12,color:T.muted}}>No files yet.</div>}
            {form.files.map((f,fi)=>(
              <div key={fi} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:T.bg,borderRadius:8,padding:"8px 10px",marginBottom:6}}>
                <div style={{minWidth:0,flex:1}}><div onClick={()=>openFile(f)} style={{color:T.accent,fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",cursor:"pointer",textDecoration:"underline"}}>{f.name}</div><div style={{color:T.muted,fontSize:10}}>{f.uploadedAt} · {(f.size/1024).toFixed(1)} KB</div></div>
                <div style={{display:"flex",gap:6,marginLeft:8}}>
                  <button style={{...S.btn(T.green),padding:"4px 10px",fontSize:11}} onClick={()=>dlFile(f)}>⬇</button>
                  <button style={{...S.btn(T.red),padding:"4px 8px",fontSize:11}} onClick={()=>setForm(f2=>({...f2,files:f2.files.filter((_,j)=>j!==fi)}))}>✕</button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      {fModal&&<FlightModal initial={fModal==="add"?null:fModal} onSave={saveLeg} onClose={()=>setFModal(null)}/>}
      {aModal&&(
        <Modal title={aModal==="add"?"Add Accommodation":"Edit Accommodation"} onClose={()=>setAModal(null)}>
          <Field label="City"><input style={S.input} value={aForm.city||""} onChange={e=>updA("city",e.target.value)} placeholder="e.g. New York"/></Field>
          <Field label="Hotel / Property"><input style={S.input} value={aForm.hotel||""} onChange={e=>updA("hotel",e.target.value)} placeholder="e.g. Hilton Midtown"/></Field>
          <G2>
            <Field label="Check-in Date"><input style={S.input} type="date" value={aForm.checkIn||""} onChange={e=>updA("checkIn",e.target.value)}/></Field>
            <Field label="Check-in Time"><input style={S.input} type="time" value={aForm.checkInTime||""} onChange={e=>updA("checkInTime",e.target.value)}/></Field>
          </G2>
          <Field label="Check-out Date"><input style={S.input} type="date" value={aForm.checkOut||""} onChange={e=>updA("checkOut",e.target.value)}/></Field>
          <Field label="Total Cost (TTD)"><input style={S.input} type="number" value={aForm.cost||""} onChange={e=>updA("cost",e.target.value)} placeholder="0.00"/></Field>
          <SaveCancel onCancel={()=>setAModal(null)} onSave={saveAcc}/>
        </Modal>
      )}
    </div>
  );
}

// ── LEISURE MAIN ──────────────────────────────────────────────────────────────
function Leisure({data,setData}){
  const T = useT(); const S = useS();
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [search,setSearch] = useState("");
  const [collapsed,setCollapsed] = useLocalState("leisure_collapsed",false);
  const [activeTrip,setActiveTrip] = useState(null);
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const statColors = {Booked:T.accent,Planning:T.muted};
  const getFirst = leg=>{if(!leg)return null;if(leg.flightType==="Multi-City")return leg.segments?.[0]?.departDate||null;return leg.departDate||null;};
  const getLast = leg=>{if(!leg)return null;if(leg.flightType==="Multi-City"){const s=leg.segments||[];return s[s.length-1]?.arriveDate||s[s.length-1]?.departDate||null;}if(leg.flightType==="Round Trip")return leg.returnArriveDate||leg.returnDate||null;return leg.arriveDate||leg.departDate||null;};
  const tripDates = r=>{const sl=[...(r.legs||[])].sort((a,b)=>(getFirst(a)||"").localeCompare(getFirst(b)||""));return{start:sl.length>0?getFirst(sl[0]):null,end:sl.length>0?getLast(sl[sl.length-1]):null};};
  const tripCost = r=>(r.legs||[]).reduce((s,l)=>s+Number(l.flightCost||0),0)+(r.accommodations||[]).reduce((s,a)=>s+Number(a.cost||0),0);
  const cyTrips = data.leisure.filter(r=>{const fd=getFirst([...(r.legs||[])].sort((a,b)=>(getFirst(a)||"").localeCompare(getFirst(b)||""))[0]);return !fd||logYear(fd)===CY;});
  const total = cyTrips.reduce((s,r)=>s+tripCost(r),0);
  const shown = data.leisure.filter(r=>!search||r.trip.toLowerCase().includes(search.toLowerCase())||(r.status||"").toLowerCase().includes(search.toLowerCase()));
  function save(){const entry={...form,legs:form.legs||[],accommodations:form.accommodations||[],files:form.files||[]};setData(d=>({...d,leisure:modal==="add"?[...d.leisure,{...entry,id:nextId(d.leisure)}]:d.leisure.map(r=>r.id===form.id?entry:r)}));setModal(null);}
  function saveTrip(u){setData(d=>({...d,leisure:d.leisure.map(r=>r.id===u.id?u:r)}));}
  if(activeTrip!==null){const trip=data.leisure.find(r=>r.id===activeTrip);if(trip)return <TripDetail trip={trip} onBack={()=>setActiveTrip(null)} onSave={saveTrip}/>;}
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>LEISURE & TRAVEL</div>
        <button style={{...S.btn(T.green),padding:"8px 14px"}} onClick={()=>{setForm({trip:"",status:"Planning",legs:[],accommodations:[],files:[]});setModal("add");}}>+ Add</button>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// TRIPS & ACTIVITIES · Tap card for details</div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search trips..."/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
        <StatCard label={`${CY} Trips`} value={cyTrips.length} color={T.accent}/>
        <StatCard label={`${CY} Trip Cost`} value={fmt(total)} color={T.green}/>
      </div>
      <ListToggle collapsed={collapsed} onToggle={()=>setCollapsed(c=>!c)} count={shown.length} label="Trips"/>
      {!collapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          {shown.map(r=>{
            const {start,end}=tripDates(r);
            const tc=tripCost(r);
            return(
              <div key={r.id} onClick={()=>setActiveTrip(r.id)} style={{...S.card,cursor:"pointer",position:"relative",overflow:"hidden",borderColor:T.green+"40"}} onMouseEnter={e=>e.currentTarget.style.boxShadow=`0 0 16px ${T.green}20`} onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
                <div style={{position:"absolute",top:-8,right:-8,fontSize:60,opacity:0.05}}>✈</div>
                <div style={{fontWeight:700,fontSize:16,marginBottom:6,color:T.text}}>{r.trip}</div>
                {(start||end)&&<div style={{fontSize:12,color:T.muted,marginBottom:8}}>{fmtDate(start)}{end?` → ${fmtDate(end)}`:""}</div>}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",marginBottom:8}}>
                  <div>
                    <div style={{fontSize:10,color:T.muted,letterSpacing:1,marginBottom:2}}>TOTAL TRIP COST</div>
                    <div style={{fontSize:18,fontWeight:700,color:T.green}}>{tc>0?fmt(tc):"TBD"}</div>
                  </div>
                  <span style={S.badge(statColors[r.status]||T.muted)}>{r.status}</span>
                </div>
                <div style={{display:"flex",gap:10,fontSize:11,color:T.muted,marginBottom:4}}>
                  {r.legs?.length>0&&<span>✈ {r.legs.length} flight{r.legs.length!==1?"s":""}</span>}
                  {r.accommodations?.length>0&&<span>🏨 {r.accommodations.length} stay{r.accommodations.length!==1?"s":""}</span>}
                </div>
                <div style={{fontSize:9,color:T.border,letterSpacing:1}}>TAP FOR DETAILS →</div>
                <div style={{display:"flex",justifyContent:"flex-end",marginTop:8}} onClick={e=>e.stopPropagation()}>
                  <ItemActions label={r.trip} onEdit={()=>{setForm({...r});setModal("edit");}} onDelete={()=>setData(d=>({...d,leisure:d.leisure.filter(x=>x.id!==r.id)}))}/>
                </div>
              </div>
            );
          })}
        </div>
      )}
      {modal&&(
        <Modal title={modal==="add"?"Add Trip":"Edit Trip"} onClose={()=>setModal(null)}>
          <Field label="Trip Name"><input style={S.input} value={form.trip||""} onChange={e=>upd("trip",e.target.value)} placeholder="e.g. World Cup 2026"/></Field>
          <Field label="Status"><select style={S.input} value={form.status||"Planning"} onChange={e=>upd("status",e.target.value)}><option>Planning</option><option>Booked</option></select></Field>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
    </div>
  );
}

// ── INVESTMENTS ───────────────────────────────────────────────────────────────
function Investments({data,setData}){
  const T = useT(); const S = useS();
  const [modal,setModal] = useState(null);
  const [form,setForm] = useState({});
  const [search,setSearch] = useState("");
  const [collapsed,setCollapsed] = useLocalState("investments_collapsed",false);
  const [cardFilter,setCardFilter] = useLocalState("investments_card_filter","All");
  const upd = (k,v)=>setForm(f=>({...f,[k]:v}));
  const total = data.investments.reduce((s,r)=>s+Number(r.value||0),0);
  const etfTotal = data.investments.filter(r=>r.type==="ETF").reduce((s,r)=>s+Number(r.value||0),0);
  const stockTotal = data.investments.filter(r=>r.type==="Stock").reduce((s,r)=>s+Number(r.value||0),0);
  const etfPct = total>0?((etfTotal/total)*100).toFixed(1):"0";
  const stockPct = total>0?((stockTotal/total)*100).toFixed(1):"0";
  const baseShown = [...data.investments].sort((a,b)=>b.value-a.value).filter(r=>cardFilter==="ETF"?r.type==="ETF":cardFilter==="Stock"?r.type==="Stock":true);
  const shown = baseShown.filter(r=>!search||(r.ticker||"").toLowerCase().includes(search.toLowerCase())||(r.type||"").toLowerCase().includes(search.toLowerCase()));
  function save(){const entry={...form,value:Number(form.value||0)};setData(d=>({...d,investments:modal==="add"?[...d.investments,{...entry,id:nextId(d.investments)}]:d.investments.map(r=>r.id===form.id?entry:r)}));setModal(null);}
  const typeColor = t=>t==="ETF"?T.accent:t==="Stock"?T.purple:T.yellow;
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
        <div style={{fontSize:18,fontWeight:700,letterSpacing:2}}>INVESTMENTS</div>
        <button style={{...S.btn(T.green),padding:"8px 14px"}} onClick={()=>{setForm({ticker:"",type:"ETF",value:""});setModal("add");}}>+ Add</button>
      </div>
      <div style={{color:T.muted,fontSize:11,letterSpacing:1,marginBottom:14}}>// PORTFOLIO TRACKER</div>
      <SearchBar value={search} onChange={setSearch} placeholder="Search ticker or type..."/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(155px, 1fr))",gap:10,marginBottom:14}}>
        <StatCard label="Total" value={fmt(total)} color={T.green} onClick={()=>setCardFilter("All")} active={cardFilter==="All"}/>
        <StatCard label={`ETF ${etfPct}%`} value={fmt(etfTotal)} color={T.accent} onClick={()=>setCardFilter(f=>f==="ETF"?"All":"ETF")} active={cardFilter==="ETF"}/>
        <StatCard label={`Stock ${stockPct}%`} value={fmt(stockTotal)} color={T.purple} onClick={()=>setCardFilter(f=>f==="Stock"?"All":"Stock")} active={cardFilter==="Stock"}/>
      </div>
      <ListToggle collapsed={collapsed} onToggle={()=>setCollapsed(c=>!c)} count={shown.length} label="Holdings"/>
      {!collapsed&&(
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {shown.length===0&&<div style={{...S.card,textAlign:"center",color:T.muted}}>No holdings match.</div>}
          {shown.map(r=>{
            const pct = total>0?((r.value/total)*100).toFixed(1):"0";
            const tc = typeColor(r.type);
            const typeTotal = r.type==="ETF"?etfTotal:r.type==="Stock"?stockTotal:0;
            const typePct = typeTotal>0?((r.value/typeTotal)*100).toFixed(1):"—";
            return(
              <div key={r.id} style={S.card}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div>
                    <div style={{fontWeight:700,fontSize:16,color:tc,marginBottom:4}}>{r.ticker}</div>
                    <span style={S.badge(tc)}>{r.type}</span>
                  </div>
                  <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:4}}>
                    <div style={{fontSize:18,fontWeight:700,color:T.green}}>{fmt(r.value)}</div>
                    <div style={{fontSize:11,color:T.muted,textAlign:"right"}}>
                      <span style={{color:T.text}}>{pct}%</span> of total
                    </div>
                    <div style={{fontSize:11,color:T.muted,textAlign:"right"}}>
                      <span style={{color:tc}}>{typePct}%</span> of {r.type}s
                    </div>
                    <ItemActions label={r.ticker} onEdit={()=>{setForm({...r});setModal("edit");}} onDelete={()=>setData(d=>({...d,investments:d.investments.filter(x=>x.id!==r.id)}))}/>
                  </div>
                </div>
                <div style={{marginTop:10,height:4,background:T.border,borderRadius:4}}>
                  <div style={{height:"100%",width:`${pct}%`,background:tc,borderRadius:4}}/>
                </div>
              </div>
            );
          })}
        </div>
      )}
      {modal&&(
        <Modal title={modal==="add"?"Add Holding":"Edit Holding"} onClose={()=>setModal(null)}>
          <Field label="Ticker"><input style={S.input} value={form.ticker||""} onChange={e=>upd("ticker",e.target.value.toUpperCase())} placeholder="e.g. VOO"/></Field>
          <Field label="Type"><select style={S.input} value={form.type||"ETF"} onChange={e=>upd("type",e.target.value)}><option>ETF</option><option>Stock</option><option>Other</option></select></Field>
          <Field label="Current Value (USD)"><input style={S.input} type="number" value={form.value||""} onChange={e=>upd("value",e.target.value)} placeholder="0.00"/></Field>
          <SaveCancel onCancel={()=>setModal(null)} onSave={save}/>
        </Modal>
      )}
    </div>
  );
}

// ── NAV + APP SHELL ───────────────────────────────────────────────────────────

import { createClient } from "@supabase/supabase-js";

const TEST_MODE = false; // ← set to true to test locally with empty data, false for real app

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: { persistSession: true, storageKey: "fintrax_auth" },
  global: { headers: { apikey: SUPABASE_ANON_KEY } }
});

// ── AUTH SCREEN ───────────────────────────────────────────────────────────────
function AuthScreen({onAuth}){
  const [screen,setScreen] = useState("login");
  const [username,setUsername] = useState("");
  const [email,setEmail] = useState("");
  const [password,setPassword] = useState("");
  const [error,setError] = useState("");
  const [loading,setLoading] = useState(false);
  const [showPass,setShowPass] = useState(false);

  function switchScreen(s){ setScreen(s); setError(""); setUsername(""); setEmail(""); setPassword(""); }

  async function handleLogin(e){
    e.preventDefault();
    const u = username.trim();
    const p = password;
    if(!u||!p){ setError("Please enter your username and password."); return; }
    setLoading(true); setError("");
    try {
      // Look up email from username
      const {data:profile,error:lookupErr} = await supabase
        .from("user_profiles")
        .select("email")
        .eq("username", u)
        .maybeSingle();
      if(lookupErr||!profile){
        setLoading(false);
        setError("Username not found. Please check your username or sign up.");
        return;
      }
      // Sign in with the email we found
      const {data,error:signInErr} = await supabase.auth.signInWithPassword({
        email: profile.email,
        password: p,
      });
      setLoading(false);
      if(signInErr){ setError("Incorrect password. Please try again."); return; }
      onAuth(data.user);
    } catch(e){
      setLoading(false);
      setError("Something went wrong. Please try again.");
    }
  }

  async function handleSignup(e){
    e.preventDefault();
    const u = username.trim();
    const em = email.trim().toLowerCase();
    const p = password;
    if(!u||!em||!p){ setError("Please fill in all fields."); return; }
    if(u.length<3){ setError("Username must be at least 3 characters."); return; }
    if(p.length<6){ setError("Password must be at least 6 characters."); return; }
    if(!em.includes("@")){ setError("Please enter a valid email address."); return; }
    setLoading(true); setError("");
    try {
      // Check if username already taken
      const {data:existing} = await supabase
        .from("user_profiles")
        .select("id")
        .eq("username", u)
        .maybeSingle();
      if(existing){ setLoading(false); setError("That username is already taken. Please choose another."); return; }
      // Create auth account
      const {data,error:signUpErr} = await supabase.auth.signUp({
        email: em,
        password: p,
        options: { data: { username: u } }
      });
      if(signUpErr){
        setLoading(false);
        if(signUpErr.message.toLowerCase().includes("already registered")){
          setError("An account with that email already exists.");
        } else {
          setError(`Sign up failed: ${signUpErr.message}`);
        }
        return;
      }
      if(!data?.user){ setLoading(false); setError("Sign up failed. Please try again."); return; }
      // Save username + email to profiles table
      const {error:profileErr} = await supabase
        .from("user_profiles")
        .insert({ id: data.user.id, username: u, email: em });
      setLoading(false);
      if(profileErr){ setError(`Profile save failed: ${profileErr.message}`); return; }
      onAuth(data.user);
    } catch(e){
      setLoading(false);
      setError("Something went wrong. Please try again.");
    }
  }

  const isLogin = screen==="login";
  const inp = {width:"100%",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.text,padding:"12px 14px",fontSize:14,fontFamily:"inherit",boxSizing:"border-box"};
  return(
    <div style={{minHeight:"100vh",background:T.bg,color:T.text,fontFamily:"'IBM Plex Mono','Courier New',monospace",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"24px 20px",boxSizing:"border-box",overflowX:"hidden"}}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;700&display=swap" rel="stylesheet"/>
      <style>{`*{box-sizing:border-box;}input:focus{border-color:${T.accent}!important;outline:none;}button:active{opacity:0.7;}`}</style>
      <div style={{width:"100%",maxWidth:380}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontSize:40,color:T.accent,marginBottom:8}}>⬡</div>
          <div style={{fontSize:22,fontWeight:700,letterSpacing:4,color:T.accent}}>FINTRAX</div>
          <div style={{fontSize:11,color:T.muted,letterSpacing:2,marginTop:4}}>// PERSONAL FINANCE DASHBOARD</div>
        </div>
        <div style={{background:T.surface,borderRadius:16,padding:"24px 20px",border:`1px solid ${T.border}`}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:24}}>
            <button type="button" onClick={()=>switchScreen("login")} style={{padding:"10px",borderRadius:8,border:`1px solid ${isLogin?T.accent:T.border}`,background:isLogin?T.accentGlow:"none",color:isLogin?T.accent:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:700,letterSpacing:1}}>LOG IN</button>
            <button type="button" onClick={()=>switchScreen("signup")} style={{padding:"10px",borderRadius:8,border:`1px solid ${!isLogin?T.green:T.border}`,background:!isLogin?T.green+"20":"none",color:!isLogin?T.green:T.muted,cursor:"pointer",fontFamily:"inherit",fontSize:12,fontWeight:700,letterSpacing:1}}>SIGN UP</button>
          </div>
          <form onSubmit={isLogin?handleLogin:handleSignup}>
            <div style={{marginBottom:14}}>
              <label style={{fontSize:10,letterSpacing:2,color:T.muted,textTransform:"uppercase",display:"block",marginBottom:6}}>Username</label>
              <input style={inp} value={username} onChange={e=>setUsername(e.target.value)} placeholder="Enter your username" autoCapitalize="none" autoCorrect="off" autoComplete="username" spellCheck="false"/>
            </div>
            {!isLogin&&(
              <div style={{marginBottom:14}}>
                <label style={{fontSize:10,letterSpacing:2,color:T.muted,textTransform:"uppercase",display:"block",marginBottom:6}}>Email</label>
                <input style={inp} type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="Enter your email address" autoCapitalize="none" autoCorrect="off" autoComplete="email"/>
              </div>
            )}
            <div style={{marginBottom:20}}>
              <label style={{fontSize:10,letterSpacing:2,color:T.muted,textTransform:"uppercase",display:"block",marginBottom:6}}>Password</label>
              <div style={{position:"relative"}}>
                <input style={{...inp,paddingRight:44}} type={showPass?"text":"password"} value={password} onChange={e=>setPassword(e.target.value)} placeholder={isLogin?"Enter your password":"Min. 6 characters"} autoComplete={isLogin?"current-password":"new-password"}/>
                <button type="button" onClick={()=>setShowPass(s=>!s)} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:T.muted,cursor:"pointer",fontSize:16,padding:4,lineHeight:1}}>
                  {showPass?"🙈":"👁"}
                </button>
              </div>
            </div>
            {error&&<div style={{background:T.red+"15",border:`1px solid ${T.red}40`,borderRadius:8,padding:"10px 14px",fontSize:12,color:T.red,marginBottom:16,lineHeight:1.6}}>{error}</div>}
            <button type="submit" disabled={loading} style={{width:"100%",padding:"13px",borderRadius:8,border:"none",background:isLogin?T.accent:T.green,color:"#ffffff",cursor:loading?"not-allowed":"pointer",fontFamily:"inherit",fontSize:14,fontWeight:700,letterSpacing:1,opacity:loading?0.7:1}}>
              {loading?(isLogin?"Logging in...":"Creating account..."):(isLogin?"LOG IN":"CREATE ACCOUNT")}
            </button>
          </form>
          {!isLogin&&<div style={{fontSize:11,color:T.muted,textAlign:"center",marginTop:14,lineHeight:1.6}}>Your email is only used for account recovery. Your data is private and only accessible by you.</div>}
        </div>
      </div>
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────────
// ── BUBBLE TOGGLE ─────────────────────────────────────────────────────────────
function BubbleToggle({on, onToggle, size="md"}){
  const T = useT();
  const sz = size==="sm"?18:22;
  return(
    <div onClick={onToggle} style={{display:"inline-flex",alignItems:"center",gap:8,cursor:"pointer"}}>
      <div style={{width:sz*2.4,height:sz,borderRadius:sz,background:on?T.accent:T.border,position:"relative",transition:"background 0.2s",flexShrink:0}}>
        <div style={{position:"absolute",top:2,left:on?`calc(100% - ${sz-4}px - 2px)`:2,width:sz-4,height:sz-4,borderRadius:"50%",background:"white",transition:"left 0.2s"}}/>
      </div>
      <span style={{fontSize:11,color:on?T.accent:T.muted,fontWeight:700,letterSpacing:1}}>{on?"ENABLED":"DISABLED"}</span>
    </div>
  );
}

const ALL_PAGES = [
  {id:"overview",  label:"Overview",  icon:"◈"},
  {id:"expenses",  label:"Expenses",  icon:"◉"},
  {id:"certs",     label:"Certs",     icon:"◎"},
  {id:"personal",  label:"Docs",      icon:"◫"},
  {id:"subscriptions",label:"Subs",   icon:"◌"},
  {id:"car",       label:"Car",       icon:"◧"},
  {id:"leisure",   label:"Leisure",   icon:"◑"},
  {id:"investments",label:"Invest",   icon:"◐"},
];
const PAGE_COMPONENTS = {overview:Overview,expenses:Expenses,certs:Certifications,personal:PersonalDocs,subscriptions:Subscriptions,car:CarMaintenance,leisure:Leisure,investments:Investments};

const DEFAULT_ENABLED = {overview:true,expenses:true,certs:true,personal:true,subscriptions:true,car:true,leisure:true,investments:true};

export default function App(){
  const [active,setActive]           = useState("overview");
  const [data,setData]               = useState(INIT);
  const [menuOpen,setMenuOpen]       = useState(false);
  const [menuSection,setMenuSection] = useState("main");
  const [syncing,setSyncing]         = useState(false);
  const [loaded,setLoaded]           = useState(false);
  const [user,setUser]               = useState(null);
  const [authChecked,setAuthChecked] = useState(false);
  const [darkMode,setDarkMode]       = useState(true);
  const [enabledPages,setEnabledPages] = useState(DEFAULT_ENABLED);
  const [bottomExpanded,setBottomExpanded] = useState(false);
  const [exportSel,setExportSel] = useState({expenses:true,certs:true,personal:true,subscriptions:true,car:true,leisure:true,investments:true});
  const dataRef     = React.useRef(data);
  const pendingSave = React.useRef(null);
  React.useEffect(()=>{ dataRef.current = data; },[data]);

  // ── INACTIVITY AUTO-LOGOUT (15 min) ──────────────────────────────────────────
  React.useEffect(()=>{
    if(!user||TEST_MODE) return;
    const TIMEOUT = 15*60*1000;
    let timer = setTimeout(handleLogout, TIMEOUT);
    const reset = ()=>{ clearTimeout(timer); timer=setTimeout(handleLogout, TIMEOUT); };
    const events = ["mousemove","mousedown","keydown","touchstart","scroll","click"];
    events.forEach(e=>window.addEventListener(e, reset, {passive:true}));
    return()=>{ clearTimeout(timer); events.forEach(e=>window.removeEventListener(e, reset)); };
  },[user]);

  // Compute current theme
  const currentT = darkMode ? DARK : LIGHT;

  // Update document background when theme changes — runs on mount and every toggle
  React.useEffect(()=>{
    const bg = darkMode ? DARK.bg : LIGHT.bg;
    const text = darkMode ? DARK.text : LIGHT.text;
    document.body.style.cssText = `background:${bg} !important;color:${text};margin:0;padding:0;overflow-x:hidden;`;
    document.documentElement.style.cssText = `background:${bg} !important;`;
  },[darkMode]);

  // Check if user is already logged in on app open
  React.useEffect(()=>{
    if(TEST_MODE){setAuthChecked(true);setLoaded(true);return;}
    supabase.auth.getSession().then(({data:{session}})=>{
      if(session?.user){setUser(session.user);loadData(session.user.id);}
      else{setAuthChecked(true);}
    });
    const {data:{subscription}} = supabase.auth.onAuthStateChange((_event,session)=>{
      if(session?.user){ setUser(session.user); }
      else { setUser(null);setData(INIT);setLoaded(false);setAuthChecked(true); }
    });
    return()=>subscription.unsubscribe();
  },[]);

  async function loadData(userId){
    try{
      const {data:rows,error} = await supabase
        .from("fintrax_data").select("data").eq("user_id",userId)
        .order("updated_at",{ascending:false}).limit(1).maybeSingle();
      if(error){ console.error("Load error:",error); }
      else if(rows?.data&&Object.keys(rows.data).length>0){
        const saved = rows.data;
        setData({...INIT,...saved});
        if(saved.__darkMode !== undefined) setDarkMode(saved.__darkMode);
        if(saved.__enabledPages) setEnabledPages({...DEFAULT_ENABLED,...saved.__enabledPages});
      }
    }catch(e){console.error("Load error:",e);}
    setLoaded(true);setAuthChecked(true);
  }

  function handleAuth(u){ setUser(u); if(u) loadData(u.id); }

  async function handleLogout(){
    if(pendingSave.current){ await pendingSave.current; pendingSave.current=null; }
    await supabase.auth.signOut();
    setUser(null);setData(INIT);setLoaded(false);setAuthChecked(true);
    setMenuOpen(false);setMenuSection("main");
  }

  const setDataAndSave = (updater)=>{
    const next = typeof updater==="function"?updater(dataRef.current):updater;
    dataRef.current = next;
    setData(next);
    if(!TEST_MODE&&user){
      setSyncing(true);
      pendingSave.current = supabase.from("fintrax_data")
        .upsert({user_id:user.id,data:next,updated_at:new Date().toISOString()},{onConflict:"user_id"})
        .then(({error})=>{ if(error) console.error("Save error:",error); setSyncing(false); })
        .catch((e)=>{ console.error("Save error:",e); setSyncing(false); });
    }
  };

  function exportToExcel(){
    const wb = XLSX.utils.book_new();
    const addSheet = (rows,label)=>{ const ws=XLSX.utils.json_to_sheet(rows.length>0?rows:[{"No data":"—"}]); XLSX.utils.book_append_sheet(wb,ws,label); };
    const parseM = str=>{if(!str)return[];return str.split(",").map(m=>{const p=m.trim().split(":");return{name:p[0]?.trim()||"",paid:(p[1]?.trim()||"").toLowerCase()==="paid",date:p[2]?.trim()||""};}).filter(m=>m.name);};

    const sections = [
      {key:"expenses",      label:"Expenses",        rows: data.expenses.map(e=>       ({"Name":e.name,"Category":e.category,"Amount (TTD)":Number(e.amount||0)}))},
      {key:"certs",         label:"Certifications",  rows: data.certs.map(c=>          ({"Name":c.name,"Issued":c.issued||"","Expiry":c.expiry||"","Cost TTD":Number(c.costTTD||0),"Cost USD":Number(c.costUSD||0)}))},
      {key:"personal",      label:"Personal Docs",   rows: data.personalDocs.map(d=>   ({"Type":d.type,"Issued":d.issued||"","Expiry":d.expiry||"","Cost (TTD)":Number(d.cost||0)}))},
      {key:"subscriptions", label:"Subscriptions",   rows: data.subscriptions.flatMap(s=>{
        const members=parseM(s.members);
        const base={"Service":s.service,"Type":s.type,"Amount (TTD)":Number(s.amount||0),"Renewal Date":s.renews||""};
        if(members.length===0) return [{...base,"Member":"","Status":"","Paid Date":"","My Share (TTD)":""}];
        const share=Number(s.amount||0)/members.length;
        return members.map(m=>({...base,"Member":m.name,"Status":m.paid?"Paid":"Unpaid","Paid Date":m.date||"","My Share (TTD)":Number(share.toFixed(2))}));
      })},
      {key:"car",           label:"Car Maintenance", rows: data.carLog.map(r=>         ({"Vehicle":r.vehicleId||"","Item":r.item,"Action":r.action,"Date":r.date||"","Mileage (km)":Number(r.mileage||0),"Cost (TTD)":Number(r.cost||0),"Supplier":r.supplier||""}))},
      {key:"investments",   label:"Investments",     rows: data.investments.map(r=>    ({"Ticker":r.ticker,"Type":r.type,"Value (USD)":Number(r.value||0)}))},
    ];

    let added = 0;
    sections.forEach(s=>{
      if(!exportSel[s.key]) return;
      addSheet(s.rows, s.label);
      added++;
    });

    // ── Leisure: structured layout using array-of-arrays ─────────────────────
    if(exportSel.leisure){
      const getFirst = leg=>{ if(!leg)return""; if(leg.flightType==="Multi-City")return leg.segments?.[0]?.departDate||""; return leg.departDate||""; };
      const getLast  = leg=>{ if(!leg)return""; if(leg.flightType==="Multi-City"){const s=leg.segments||[];return s[s.length-1]?.arriveDate||s[s.length-1]?.departDate||"";} if(leg.flightType==="Round Trip")return leg.returnArriveDate||leg.returnDate||""; return leg.arriveDate||leg.departDate||""; };
      // Columns: [#/Label] [Detail A] [Detail B] [Detail C] [Detail D] [Depart/From Date] [Depart/From Time] [Arrive/To Date] [Arrive/To Time] [Cost (TTD)]
      const aoa = []; // array of arrays

      data.leisure.forEach((r, tripIdx)=>{
        const sl=[...(r.legs||[])].sort((a,b)=>(getFirst(a)||"").localeCompare(getFirst(b)||""));
        const flightCost=(r.legs||[]).reduce((s,l)=>s+Number(l.flightCost||0),0);
        const accCost=(r.accommodations||[]).reduce((s,a)=>s+Number(a.cost||0),0);
        const startDate=sl.length>0?getFirst(sl[0]):"", endDate=sl.length>0?getLast(sl[sl.length-1]):"";

        // ── Trip title ──
        aoa.push([`TRIP: ${r.trip}`,"","","","","","","","",""]);
        // ── Trip summary ──
        aoa.push(["Status:", r.status||"Planning", "", "Period:", startDate, "→", endDate, "", "Total Cost (TTD):", flightCost+accCost]);
        aoa.push(["Flights:", (r.legs||[]).length, "", "Stays:", (r.accommodations||[]).length, "", "Flight Cost (TTD):", flightCost, "Stay Cost (TTD):", accCost]);
        aoa.push([""]);

        // ── Flights ──
        if((r.legs||[]).length>0){
          aoa.push(["FLIGHTS","","","","","","","","",""]);
          aoa.push(["#","Type","From","To","Airline","Flight No.","Depart Date","Depart Time","Arrive Date","Arrive Time","Return Date","Return Time","Cost (TTD)"]);
          r.legs.forEach((leg,li)=>{
            if(leg.flightType==="Multi-City"){
              const segs=leg.segments||[];
              const fullRoute=segs.map(s=>s.from).concat(segs.length>0?[segs[segs.length-1].to]:[]).join(" → ");
              aoa.push([li+1,"Multi-City",fullRoute,"",leg.airline||"",leg.flightNo||"","","","","","","",Number(leg.flightCost||0)]);
              segs.forEach((seg,si)=>{
                aoa.push(["",`  Seg ${si+1}`,seg.from||"",seg.to||"","","",seg.departDate||"",seg.departTime||"",seg.arriveDate||"",seg.arriveTime||"","","",""]);
              });
            } else if(leg.flightType==="Round Trip"){
              aoa.push([li+1,"Round Trip",leg.from||"",leg.to||"",leg.airline||"",leg.flightNo||"","","","","","","",Number(leg.flightCost||0)]);
              aoa.push(["","  Outbound",leg.from||"",leg.to||"","","",leg.departDate||"",leg.departTime||"",leg.arriveDate||"",leg.arriveTime||"","","",""]);
              aoa.push(["","  Return",leg.to||"",leg.from||"","","",leg.returnDate||"",leg.returnTime||"",leg.returnArriveDate||"",leg.returnArriveTime||"","","",""]);
            } else {
              aoa.push([li+1,"One Way",leg.from||"",leg.to||"",leg.airline||"",leg.flightNo||"",leg.departDate||"",leg.departTime||"",leg.arriveDate||"",leg.arriveTime||"","","",Number(leg.flightCost||0)]);
            }
          });
          aoa.push([""]);
        }

        // ── Accommodation ──
        if((r.accommodations||[]).length>0){
          aoa.push(["ACCOMMODATION","","","","","","","","",""]);
          aoa.push(["#","Hotel / Property","City","","","","Check-in Date","Check-in Time","Check-out Date","","Cost (TTD)"]);
          r.accommodations.forEach((a,ai)=>{
            aoa.push([ai+1,a.hotel||"",a.city||"","","","",a.checkIn||"",a.checkInTime||"",a.checkOut||"","",Number(a.cost||0)]);
          });
          aoa.push([""]);
        }

        // ── Trip total ──
        aoa.push(["","","","","","","","","TRIP TOTAL (TTD):", flightCost+accCost]);

        // ── Separator between trips ──
        if(tripIdx<data.leisure.length-1){
          aoa.push([""]);
          aoa.push(["────────────────────────────────────────────────────────────────────────────────"]);
          aoa.push([""]);
        }
      });

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      // Set column widths
      ws["!cols"] = [{wch:4},{wch:14},{wch:20},{wch:20},{wch:20},{wch:14},{wch:14},{wch:12},{wch:14},{wch:12},{wch:14},{wch:12},{wch:14}];
      XLSX.utils.book_append_sheet(wb, ws, "Leisure & Travel");
      added++;
    }

    if(added===0) return;
    XLSX.writeFile(wb,`fintrax_export_${CY}.xlsx`);
  }

  // Save settings to Supabase when they change
  function saveSettings(dm, ep){
    if(!TEST_MODE&&user){
      const settingsData = {...dataRef.current,__darkMode:dm,__enabledPages:ep};
      supabase.from("fintrax_data")
        .upsert({user_id:user.id,data:settingsData,updated_at:new Date().toISOString()},{onConflict:"user_id"})
        .catch((e)=>console.error("Settings save error:",e));
    }
  }

  function toggleDark(){
    const nd = !darkMode;
    setDarkMode(nd);
    saveSettings(nd, enabledPages);
  }

  function togglePage(id){
    if(id==="overview") return;
    const ne = {...enabledPages,[id]:!enabledPages[id]};
    setEnabledPages(ne);
    if(!ne[active]) setActive("overview");
    saveSettings(darkMode, ne);
  }

  const username = user?.user_metadata?.username||user?.email?.split("@")?.[0]||"User";
  const visiblePages = ALL_PAGES.filter(p => enabledPages[p.id]);
  const Page = PAGE_COMPONENTS[active] || Overview;

  // Loading / auth states
  if(!authChecked) return(
    <div style={{minHeight:"100vh",background:DARK.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:12,fontFamily:"'IBM Plex Mono',monospace"}}>
      <div style={{fontSize:32,color:DARK.accent}}>⬡</div>
      <div style={{fontSize:11,color:DARK.muted,letterSpacing:2}}>FINTRAX</div>
    </div>
  );
  if(!TEST_MODE&&!user) return <AuthScreen onAuth={handleAuth}/>;
  if(!loaded) return(
    <div style={{minHeight:"100vh",background:currentT.bg,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:12,fontFamily:"'IBM Plex Mono',monospace"}}>
      <div style={{fontSize:32,color:currentT.accent}}>⬡</div>
      <div style={{fontSize:11,color:currentT.muted,letterSpacing:2}}>Loading your data...</div>
    </div>
  );

  // Bottom nav — first 5 visible pages + More button
  const bottomPages = visiblePages.slice(0,5);
  const morePagesExist = visiblePages.length > 5;
  const extraPages = visiblePages.slice(5);

  return(
    <ThemeCtx.Provider value={currentT}>
    <div style={{minHeight:"100vh",background:currentT.bg,color:currentT.text,fontFamily:"'IBM Plex Mono','Courier New',monospace",fontSize:14,overflowX:"hidden",width:"100%"}}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;700&display=swap" rel="stylesheet"/>
      <style>{`
        *{box-sizing:border-box;-webkit-tap-highlight-color:transparent;max-width:100%;}
        html,body{overflow-x:hidden;width:100%;margin:0;padding:0;}
        input,select,button{font-family:'IBM Plex Mono','Courier New',monospace;max-width:100%;}
        input:focus,select:focus{border-color:${currentT.accent}!important;outline:none;box-shadow:0 0 0 2px ${currentT.accent}20;}
        button:active{opacity:0.7;}
        ::-webkit-scrollbar{width:4px;height:4px;}
        ::-webkit-scrollbar-thumb{background:${currentT.border};border-radius:4px;}
        input[type="date"],input[type="time"]{color-scheme:${darkMode?"dark":"light"};}
        img{max-width:100%;height:auto;}
      `}</style>

      {/* ── TOP BAR ── */}
      <div style={{position:"sticky",top:0,zIndex:200,background:currentT.surface,borderBottom:`1px solid ${currentT.border}`,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 16px",height:52}}>
        {/* LEFT — hamburger */}
        <button onClick={()=>{setMenuOpen(m=>!m);setMenuSection("main");}} style={{background:"none",border:`1px solid ${currentT.border}`,color:currentT.accent,fontSize:20,cursor:"pointer",padding:"4px 10px",borderRadius:8,fontFamily:"inherit",lineHeight:1}}>
          {menuOpen?"✕":"☰"}
        </button>
        {/* RIGHT — app name */}
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {syncing&&<span style={{fontSize:10,color:currentT.muted,letterSpacing:1}}>saving...</span>}
          {TEST_MODE&&<span style={{fontSize:9,background:currentT.yellow+"20",color:currentT.yellow,border:`1px solid ${currentT.yellow}40`,borderRadius:4,padding:"2px 6px",letterSpacing:1,fontWeight:700}}>TEST</span>}
          <div style={{fontWeight:700,letterSpacing:4,color:currentT.accent,fontSize:15}}>FINTRAX ⬡</div>
        </div>
      </div>

      {/* ── CASCADE MENU ── */}
      {menuOpen&&(
        <div style={{position:"fixed",top:52,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.5)",zIndex:199}} onClick={()=>{setMenuOpen(false);setMenuSection("main");}}>
          <div style={{background:currentT.surface,borderBottom:`1px solid ${currentT.border}`,maxHeight:"80vh",overflowY:"auto"}} onClick={e=>e.stopPropagation()}>

            {/* MAIN MENU */}
            {menuSection==="main"&&(
              <div style={{padding:16}}>
                {/* Profile row */}
                <div onClick={()=>setMenuSection("profile")} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 16px",background:currentT.card,borderRadius:12,marginBottom:10,cursor:"pointer",border:`1px solid ${currentT.border}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <div style={{width:38,height:38,borderRadius:"50%",background:currentT.accent+"20",border:`1px solid ${currentT.accent}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,color:currentT.accent,flexShrink:0}}>👤</div>
                    <div>
                      <div style={{fontSize:10,color:currentT.muted,letterSpacing:2,marginBottom:2}}>PROFILE</div>
                      <div style={{fontSize:14,fontWeight:700,color:currentT.text}}>{username}</div>
                    </div>
                  </div>
                  <span style={{color:currentT.muted,fontSize:16}}>›</span>
                </div>
                {/* Settings row */}
                <div onClick={()=>setMenuSection("settings")} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"14px 16px",background:currentT.card,borderRadius:12,cursor:"pointer",border:`1px solid ${currentT.border}`}}>
                  <div style={{display:"flex",alignItems:"center",gap:12}}>
                    <div style={{width:38,height:38,borderRadius:"50%",background:currentT.purple+"20",border:`1px solid ${currentT.purple}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,flexShrink:0}}>⚙️</div>
                    <div>
                      <div style={{fontSize:10,color:currentT.muted,letterSpacing:2,marginBottom:2}}>SETTINGS</div>
                      <div style={{fontSize:14,fontWeight:700,color:currentT.text}}>App Preferences</div>
                    </div>
                  </div>
                  <span style={{color:currentT.muted,fontSize:16}}>›</span>
                </div>
              </div>
            )}

            {/* PROFILE SECTION */}
            {menuSection==="profile"&&(
              <div style={{padding:16}}>
                <button onClick={()=>setMenuSection("main")} style={{background:"none",border:"none",color:currentT.accent,cursor:"pointer",fontSize:13,padding:"0 0 16px",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
                  ‹ Back
                </button>
                <div style={{background:currentT.card,borderRadius:12,padding:20,border:`1px solid ${currentT.border}`,marginBottom:14,textAlign:"center"}}>
                  <div style={{width:60,height:60,borderRadius:"50%",background:currentT.accent+"20",border:`2px solid ${currentT.accent}40`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:26,margin:"0 auto 12px"}}>👤</div>
                  <div style={{fontSize:18,fontWeight:700,color:currentT.text,marginBottom:4}}>{username}</div>
                  <div style={{fontSize:11,color:currentT.muted,letterSpacing:1}}>FINTRAX ACCOUNT</div>
                </div>
                <button onClick={handleLogout} style={{...S.btn(currentT.red),width:"100%",textAlign:"center",padding:"13px"}}>
                  Log Out
                </button>
              </div>
            )}

            {/* SETTINGS SECTION */}
            {menuSection==="settings"&&(
              <div style={{padding:16}}>
                <button onClick={()=>setMenuSection("main")} style={{background:"none",border:"none",color:currentT.accent,cursor:"pointer",fontSize:13,padding:"0 0 16px",fontFamily:"inherit",display:"flex",alignItems:"center",gap:6}}>
                  ‹ Back
                </button>
                {/* Dark / Light mode */}
                <div style={{background:currentT.card,borderRadius:12,padding:16,border:`1px solid ${currentT.border}`,marginBottom:12}}>
                  <div style={{fontSize:10,letterSpacing:2,color:currentT.muted,textTransform:"uppercase",marginBottom:12}}>Appearance</div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{fontSize:14,fontWeight:600,color:currentT.text}}>{darkMode?"Dark Mode":"Light Mode"}</div>
                      <div style={{fontSize:11,color:currentT.muted,marginTop:2}}>{darkMode?"Switch to light":"Switch to dark"}</div>
                    </div>
                    <BubbleToggle on={darkMode} onToggle={toggleDark}/>
                  </div>
                </div>
                {/* Page toggles */}
                <div style={{background:currentT.card,borderRadius:12,padding:16,border:`1px solid ${currentT.border}`}}>
                  <div style={{fontSize:10,letterSpacing:2,color:currentT.muted,textTransform:"uppercase",marginBottom:12}}>Enabled Pages</div>
                  {ALL_PAGES.filter(p=>p.id!=="overview").map(p=>(
                    <div key={p.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",paddingBottom:12,marginBottom:12,borderBottom:`1px solid ${currentT.border}30`}}>
                      <div style={{display:"flex",alignItems:"center",gap:10}}>
                        <span style={{fontSize:16}}>{p.icon}</span>
                        <span style={{fontSize:13,fontWeight:600,color:currentT.text}}>{p.label}</span>
                      </div>
                      <BubbleToggle on={!!enabledPages[p.id]} onToggle={()=>togglePage(p.id)} size="sm"/>
                    </div>
                  ))}
                </div>

                {/* ── Export Data ── */}
                <div style={{background:currentT.card,borderRadius:12,padding:16,border:`1px solid ${currentT.border}`,marginTop:12}}>
                  <div style={{fontSize:10,letterSpacing:2,color:currentT.muted,textTransform:"uppercase",marginBottom:4}}>Export Data</div>
                  <div style={{fontSize:11,color:currentT.muted,marginBottom:14}}>Select sections to export as a single Excel file (.xlsx), one sheet per section.</div>
                  {[
                    {key:"expenses",      label:"Expenses",        icon:"◉"},
                    {key:"certs",         label:"Certifications",  icon:"◎"},
                    {key:"personal",      label:"Personal Docs",   icon:"◫"},
                    {key:"subscriptions", label:"Subscriptions",   icon:"◌"},
                    {key:"car",           label:"Car Maintenance", icon:"◧"},
                    {key:"leisure",       label:"Leisure & Travel",icon:"◑"},
                    {key:"investments",   label:"Investments",     icon:"◐"},
                  ].map(s=>(
                    <div key={s.key} onClick={()=>setExportSel(p=>({...p,[s.key]:!p[s.key]}))}
                      style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${currentT.border}20`,cursor:"pointer"}}>
                      <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${exportSel[s.key]?currentT.accent:currentT.border}`,background:exportSel[s.key]?currentT.accent:"none",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                        {exportSel[s.key]&&<span style={{color:"#fff",fontSize:11,fontWeight:700,lineHeight:1}}>✓</span>}
                      </div>
                      <span style={{fontSize:13,color:currentT.text}}>{s.icon} {s.label}</span>
                    </div>
                  ))}
                  <div style={{display:"flex",gap:8,marginTop:14}}>
                    <button onClick={()=>setExportSel({expenses:true,certs:true,personal:true,subscriptions:true,car:true,leisure:true,investments:true})}
                      style={{...S.outlineBtn(false,currentT.muted),padding:"6px 12px",fontSize:11}}>All</button>
                    <button onClick={()=>setExportSel({expenses:false,certs:false,personal:false,subscriptions:false,car:false,leisure:false,investments:false})}
                      style={{...S.outlineBtn(false,currentT.muted),padding:"6px 12px",fontSize:11}}>None</button>
                    <button onClick={exportToExcel}
                      style={{...S.btn(currentT.green),flex:1,padding:"8px 14px",fontSize:12,textAlign:"center"}}>
                      ⬇ Export to Excel
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── PAGE CONTENT ── */}
      <div style={{maxWidth:700,margin:"0 auto",padding:"20px 14px 90px"}}>
        <Page data={data} setData={setDataAndSave} enabledPages={enabledPages}/>
      </div>

      {/* ── EXPANDED BOTTOM NAV (More tray) ── */}
      {bottomExpanded&&(
        <div style={{position:"fixed",bottom:58,left:0,right:0,background:currentT.surface,borderTop:`1px solid ${currentT.border}`,zIndex:188,padding:12,display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8}} onClick={()=>setBottomExpanded(false)}>
          {extraPages.map(n=>(
            <button key={n.id} onClick={()=>{setActive(n.id);setBottomExpanded(false);}} style={{padding:"10px 6px",borderRadius:10,border:`1px solid ${active===n.id?currentT.accent:currentT.border}`,background:active===n.id?currentT.accentGlow:currentT.card,color:active===n.id?currentT.accent:currentT.muted,cursor:"pointer",fontFamily:"inherit",display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
              <span style={{fontSize:18}}>{n.icon}</span>
              <span style={{fontSize:10,letterSpacing:1}}>{n.label}</span>
            </button>
          ))}
        </div>
      )}

      {/* ── BOTTOM NAV BAR ── */}
      <div style={{position:"fixed",bottom:0,left:0,right:0,background:currentT.surface,borderTop:`1px solid ${currentT.border}`,display:"flex",zIndex:190,paddingBottom:"env(safe-area-inset-bottom,0px)",height:58}}>
        {bottomPages.map(n=>(
          <button key={n.id} onClick={()=>{setActive(n.id);setBottomExpanded(false);}} style={{flex:1,padding:"8px 2px 6px",background:"none",border:"none",color:active===n.id?currentT.accent:currentT.muted,cursor:"pointer",fontFamily:"inherit",display:"flex",flexDirection:"column",alignItems:"center",gap:2,borderTop:`2px solid ${active===n.id?currentT.accent:"transparent"}`,transition:"color 0.15s"}}>
            <span style={{fontSize:16}}>{n.icon}</span>
            <span style={{fontSize:9,letterSpacing:1}}>{n.label}</span>
          </button>
        ))}
        {morePagesExist&&(
          <button onClick={()=>setBottomExpanded(e=>!e)} style={{flex:1,padding:"8px 2px 6px",background:"none",border:"none",color:bottomExpanded?currentT.accent:currentT.muted,cursor:"pointer",fontFamily:"inherit",display:"flex",flexDirection:"column",alignItems:"center",gap:2,borderTop:`2px solid ${bottomExpanded?currentT.accent:"transparent"}`}}>
            <span style={{fontSize:16}}>{bottomExpanded?"▼":"⋯"}</span>
            <span style={{fontSize:9,letterSpacing:1}}>More</span>
          </button>
        )}
      </div>
    </div>
    </ThemeCtx.Provider>
  );
}