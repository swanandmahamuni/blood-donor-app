import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Saffron + Maroon Design Tokens ──────────────────────────────────────────
const SAF = "#E8641A";
const SAF_D = "#C2410C";
const SAF_L = "#FB923C";
const SAF_XL = "#FED7AA";
const SAF_G = "rgba(232,100,26,0.3)";
const MAR = "#7F1D1D";
const MAR_M = "#991B1B";
const MAR_L = "#B91C1C";
const APP_BG = "#F7EFE3";
const APP_BG2 = "#F2E6D4";
const CARD_BG = "#FFFFFF";
const SB_BG = "#3B0C0C";
const SB_BG2 = "#4D1010";
const TX_P = "#2C0A04";
const TX_S = "#6B2A10";
const TX_M = "#A0522D";
const TX_W = "rgba(245,203,167,0.9)";
const TX_WM = "rgba(245,203,167,0.5)";
const BDR = "rgba(139,60,20,0.12)";
const BDR_H = "rgba(232,100,26,0.35)";
const GRN = "#059669";
const GRN_L = "#D1FAE5";
const RED = "#DC2626";
const RED_L = "#FEE2E2";
const BLUE = "#2563EB";
const PURPLE = "#7C3AED";
const GOLD = "#D97706";

const BLOOD_GROUPS = ["A+","A-","B+","B-","AB+","AB-","O+","O-"];
const AF = "'Palatino Linotype','Book Antiqua',Palatino,Georgia,serif";
const SF = "'Inter','Segoe UI',system-ui,sans-serif";
const MF = "'Courier New',monospace";
const MONTHS_F = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const WDAYS = ["Su","Mo","Tu","We","Th","Fr","Sa"];

const BG_C = {"A+":"#FECDD3","A-":"#FCA5A5","B+":"#BFDBFE","B-":"#93C5FD","AB+":"#DDD6FE","AB-":"#C4B5FD","O+":"#BBF7D0","O-":"#86EFAC"};
const BG_TX = {"A+":"#9F1239","A-":"#991B1B","B+":"#1E40AF","B-":"#1D4ED8","AB+":"#4C1D95","AB-":"#5B21B6","O+":"#065F46","O-":"#047857"};
const GEN_C = {Male:BLUE, Female:"#DB2777", Others:PURPLE};

const SK = {
  DONORS:"bdms-donors", BG:"bdms-bg", OPACITY:"bdms-opacity",
  GAL_IDX:"bdms-gal-index", BANNER_ICON:"bdms-banner-icon",
  SIDEBAR_ICON:"bdms-sidebar-icon", BANNER_IMG:"bdms-banner-img",
  BANNER_OPACITY:"bdms-banner-opacity", BACKUP_INTERVAL:"bdms-backup-interval",
  SIDEBAR_OPEN:"bdms-sidebar-open", SIDEBAR_DECOR:"bdms-sidebar-decor",
  SIDEBAR_DECOR_OP:"bdms-sidebar-decor-op", SIDEBAR_DECOR_X:"bdms-sidebar-decor-x",
  SIDEBAR_DECOR_Y:"bdms-sidebar-decor-y",   SIDEBAR_DECOR_SC:"bdms-sidebar-decor-sc",
};
const galKey = id => `bdms-gal-img-${id}`;
const INIT_FORM = {
  id:null, srNo:null, shraddhavanaType:"", upasanaKendra:"",
  name:"", mobile:"", dob:"", age:"", gender:"",
  bloodGroup:"", status:"", rejectionReason:"", registeredOn:"", registeredAt:"",
};

// ─── localStorage helpers ──────────────────────────────────────────────────
const sSet = (k,v) => { try { localStorage.setItem(k,v); return true; } catch { return false; } };
const sGet = k => { try { return localStorage.getItem(k); } catch { return null; } };
const sDel = k => { try { localStorage.removeItem(k); } catch {} };

// ─── Utils ────────────────────────────────────────────────────────────────
const calcAge = dob => {
  if (!dob || dob.length !== 10) return "";
  const [d,m,y] = dob.split('/').map(Number);
  if (!d||!m||!y||y<1900) return "";
  const t = new Date(), b = new Date(y,m-1,d);
  let a = t.getFullYear() - b.getFullYear();
  if (t.getMonth()<b.getMonth() || (t.getMonth()===b.getMonth() && t.getDate()<b.getDate())) a--;
  return a>0 && a<120 ? String(a) : "";
};
const todayStr = () => {
  const t = new Date();
  return `${String(t.getDate()).padStart(2,'0')}/${String(t.getMonth()+1).padStart(2,'0')}/${t.getFullYear()}`;
};
const nowDT = () => {
  const t = new Date();
  return `${String(t.getDate()).padStart(2,'0')}/${String(t.getMonth()+1).padStart(2,'0')}/${t.getFullYear()} ${String(t.getHours()).padStart(2,'0')}:${String(t.getMinutes()).padStart(2,'0')}:${String(t.getSeconds()).padStart(2,'0')}`;
};
const nowTs = () => new Date().toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit",second:"2-digit"});
const bkFname = () => {
  const t = new Date();
  return `MBDC_Backup_${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,'0')}-${String(t.getDate()).padStart(2,'0')}_${String(t.getHours()).padStart(2,'0')}-${String(t.getMinutes()).padStart(2,'0')}.json`;
};

// ─── Donut slice helper (no JSX, returns path data only) ─────────────────
const donutSlices = (data, r, cx, cy) => {
  const total = data.reduce((s,d) => s+d.val, 0) || 1;
  let offset = -0.25;
  return data.map(d => {
    const pct = d.val / total;
    const startA = offset * Math.PI * 2;
    offset += pct;
    const endA = offset * Math.PI * 2;
    const largeArc = pct > 0.5 ? 1 : 0;
    const x1 = cx + r * Math.cos(startA), y1 = cy + r * Math.sin(startA);
    const x2 = cx + r * Math.cos(endA),   y2 = cy + r * Math.sin(endA);
    return { ...d, pathD: `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z` };
  });
};

// ═══ DatePicker ───────────────────────────────────────────────────────────
function DatePicker({ value, onChange, inputStyle }) {
  const [open, setOpen] = useState(false);
  const [txt, setTxt] = useState(value || "");
  const [vY, setVY] = useState(() => { if (value) { const p = value.split('/'); return parseInt(p[2])||1990; } return 1990; });
  const [vM, setVM] = useState(() => { if (value) { const p = value.split('/'); return (parseInt(p[1])-1)||0; } return 0; });
  const wRef = useRef();

  useEffect(() => { setTxt(value || ""); }, [value]);
  useEffect(() => {
    const fn = e => { if (wRef.current && !wRef.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", fn);
    return () => document.removeEventListener("mousedown", fn);
  }, []);

  const parse = s => {
    if (!s || s.length !== 10) return null;
    const [d,m,y] = s.split('/').map(Number);
    if (!d||!m||!y||d<1||d>31||m<1||m>12||y<1900||y>new Date().getFullYear()+1) return null;
    return {d,m,y};
  };
  const handleTyping = e => {
    const raw = e.target.value.replace(/\D/g,'').slice(0,8);
    const v = raw.length>4 ? `${raw.slice(0,2)}/${raw.slice(2,4)}/${raw.slice(4)}` : raw.length>2 ? `${raw.slice(0,2)}/${raw.slice(2)}` : raw;
    setTxt(v);
    if (v.length===10 && parse(v)) { onChange(v); const p=parse(v); setVY(p.y); setVM(p.m-1); }
  };
  const pick = day => { const s=`${String(day).padStart(2,'0')}/${String(vM+1).padStart(2,'0')}/${vY}`; setTxt(s); onChange(s); setOpen(false); };
  const navM = dir => {
    if (dir<0) { if (vM===0) { setVM(11); setVY(y=>y-1); } else setVM(m=>m-1); }
    else { if (vM===11) { setVM(0); setVY(y=>y+1); } else setVM(m=>m+1); }
  };

  const dim = new Date(vY,vM+1,0).getDate(), fd = new Date(vY,vM,1).getDay();
  const parsed = parse(txt), today = new Date(), curY = today.getFullYear();
  const years = Array.from({length:curY-1924+1}, (_,i) => curY-i);

  const empties = Array.from({length:fd}, (_,i) => i);
  const days = Array.from({length:dim}, (_,i) => i+1);

  const nBtn = { background:SAF_XL, border:"none", borderRadius:7, color:SAF_D, cursor:"pointer", padding:"5px 11px", fontSize:15, lineHeight:1, fontWeight:"bold" };
  const sel = { background:APP_BG, border:`1px solid ${BDR}`, borderRadius:7, color:TX_P, padding:"6px", fontSize:12, fontFamily:SF };

  return (
    <div ref={wRef} style={{position:"relative"}}>
      <div style={{position:"relative"}}>
        <input value={txt} onChange={handleTyping} placeholder="DD/MM/YYYY" maxLength={10} autoComplete="off"
          onFocus={() => { setOpen(true); if (txt.length===10 && parse(txt)) { const p=parse(txt); setVY(p.y); setVM(p.m-1); } }}
          style={{...inputStyle, paddingRight:44}} />
        <span onClick={() => setOpen(o=>!o)} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",cursor:"pointer",fontSize:16,color:SAF,lineHeight:1,userSelect:"none"}}>📅</span>
      </div>
      {open && (
        <div style={{position:"absolute",top:"calc(100% + 8px)",left:0,zIndex:9999,background:CARD_BG,border:`1px solid ${BDR_H}`,borderRadius:14,padding:18,width:300,boxShadow:`0 20px 60px rgba(139,60,20,0.2)`}}>
          <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:12}}>
            <button onClick={() => navM(-1)} style={nBtn}>‹</button>
            <select value={vM} onChange={e => setVM(Number(e.target.value))} style={{...sel,flex:1}}>
              {MONTHS_F.map((mo,i) => <option key={mo} value={i}>{mo}</option>)}
            </select>
            <select value={vY} onChange={e => setVY(Number(e.target.value))} style={{...sel,width:74}}>
              {years.map(y => <option key={y} value={y}>{y}</option>)}
            </select>
            <button onClick={() => navM(1)} style={nBtn}>›</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2,marginBottom:5}}>
            {WDAYS.map(d => <div key={d} style={{textAlign:"center",fontSize:10,color:TX_M,padding:"4px 0",fontWeight:"600",fontFamily:SF}}>{d}</div>)}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:2}}>
            {empties.map(i => <div key={`e${i}`} />)}
            {days.map(day => {
              const isSel = parsed && parsed.d===day && parsed.m===vM+1 && parsed.y===vY;
              const isT = today.getDate()===day && today.getMonth()===vM && today.getFullYear()===vY;
              return (
                <button key={day} onClick={() => pick(day)} style={{padding:"7px 0",borderRadius:7,border:"none",cursor:"pointer",fontSize:12,fontFamily:SF,
                  background: isSel ? `linear-gradient(135deg,${SAF},${SAF_D})` : isT ? SAF_XL : "transparent",
                  color: isSel ? "#fff" : isT ? SAF_D : TX_P,
                  fontWeight: (isSel||isT) ? "700" : "400"}}>
                  {day}
                </button>
              );
            })}
          </div>
          <div style={{marginTop:12,paddingTop:10,borderTop:`1px solid ${BDR}`,display:"flex",gap:8}}>
            <button onClick={() => { const s=todayStr(); setTxt(s); onChange(s); setOpen(false); }} style={{flex:1,background:SAF_XL,border:"none",borderRadius:7,color:SAF_D,cursor:"pointer",padding:"7px",fontSize:11,fontFamily:SF,fontWeight:"600"}}>Today</button>
            <button onClick={() => { setTxt(""); onChange(""); }} style={{flex:1,background:APP_BG,border:`1px solid ${BDR}`,borderRadius:7,color:TX_M,cursor:"pointer",padding:"7px",fontSize:11,fontFamily:SF}}>Clear</button>
            <button onClick={() => setOpen(false)} style={{flex:1,background:APP_BG,border:`1px solid ${BDR}`,borderRadius:7,color:TX_M,cursor:"pointer",padding:"7px",fontSize:11,fontFamily:SF}}>Close</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ═══ BgEditorModal ────────────────────────────────────────────────────────
function BgEditorModal({ mode, currentSrc, currentOpacity, onClose, onApply }) {
  const [img, setImg] = useState(currentSrc || null);
  const [opacity, setOpacity] = useState(currentOpacity ?? 0.12);
  const isBanner = mode === "banner";

  const load = f => {
    if (!f || !f.type.startsWith('image/')) return;
    const r = new FileReader();
    r.onload = ev => setImg(ev.target.result);
    r.readAsDataURL(f);
  };

  const prevBg = img
    ? `linear-gradient(rgba(${isBanner?"59,12,12":"247,239,227"},${1-opacity}),rgba(${isBanner?"77,16,16":"247,239,227"},${1-opacity})),url(${img}) center/cover`
    : isBanner ? `linear-gradient(135deg,${SB_BG},${MAR})` : `linear-gradient(135deg,${APP_BG},${APP_BG2})`;

  const lS = { fontSize:11, color:TX_M, textTransform:"uppercase", letterSpacing:1, display:"block", marginBottom:7, fontFamily:SF, fontWeight:"600" };

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(44,10,4,0.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9997,backdropFilter:"blur(8px)"}}>
      <div style={{background:CARD_BG,border:`1px solid ${BDR_H}`,borderRadius:20,padding:32,width:520,maxWidth:"92vw",boxShadow:`0 32px 80px rgba(139,60,20,0.25)`}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:22}}>
          <div>
            <h2 style={{fontSize:18,fontWeight:"700",color:MAR,fontFamily:AF,margin:0}}>{isBanner ? "🎨 Banner Image" : "🖼 Wallpaper"} Editor</h2>
            <p style={{fontSize:12,color:TX_M,margin:"4px 0 0",fontFamily:SF}}>Upload & adjust with live preview</p>
          </div>
          <button onClick={onClose} style={{background:APP_BG,border:`1px solid ${BDR}`,borderRadius:9,color:TX_M,cursor:"pointer",padding:"7px 13px",fontSize:13,fontFamily:SF}}>✕</button>
        </div>
        <div style={{marginBottom:16}} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();load(e.dataTransfer.files[0]);}}>
          <label style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,padding:"16px",background:APP_BG,border:`2px dashed ${BDR_H}`,borderRadius:12,cursor:"pointer",color:SAF,fontSize:13,fontFamily:SF,fontWeight:"600"}}>
            📁 {img ? "Change" : "Upload"} Image (or drag & drop)
            <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>load(e.target.files[0])} />
          </label>
        </div>
        <div style={{marginBottom:18}}>
          <label style={lS}>Overlay Opacity: <span style={{color:SAF,fontWeight:"700"}}>{Math.round(opacity*100)}%</span></label>
          <input type="range" min="0" max="0.95" step="0.05" value={opacity} onChange={e=>setOpacity(parseFloat(e.target.value))} style={{width:"100%",accentColor:SAF}} />
          <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:TX_M,marginTop:3,fontFamily:SF}}><span>Image visible</span><span>Image hidden</span></div>
        </div>
        <div style={{marginBottom:22}}>
          <label style={lS}>Live Preview</label>
          <div style={{height:isBanner?68:100,borderRadius:12,background:prevBg,display:"flex",alignItems:"center",justifyContent:"center",border:`1px solid ${BDR}`,overflow:"hidden"}}>
            <span style={{fontSize:12,fontWeight:"700",color:isBanner?"rgba(255,255,255,0.8)":TX_M,letterSpacing:2,textTransform:"uppercase",fontFamily:SF}}>{isBanner?"BANNER PREVIEW":"Wallpaper Preview"}</span>
          </div>
        </div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={() => onApply(img,opacity)} style={{flex:1,padding:"13px",background:`linear-gradient(135deg,${SAF},${SAF_D})`,border:"none",borderRadius:10,color:"#fff",cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",boxShadow:`0 6px 20px ${SAF_G}`}}>✅ Apply</button>
          {(img||currentSrc) && <button onClick={() => onApply(null,opacity)} style={{padding:"13px 18px",background:RED_L,border:`1px solid rgba(220,38,38,0.2)`,borderRadius:10,color:MAR,cursor:"pointer",fontSize:12,fontFamily:SF,fontWeight:"600"}}>Remove</button>}
          <button onClick={onClose} style={{padding:"13px 18px",background:APP_BG,border:`1px solid ${BDR}`,borderRadius:10,color:TX_M,cursor:"pointer",fontSize:12,fontFamily:SF}}>Cancel</button>
        </div>
      </div>
    </div>
  );
}

// ═══ ImageEditorModal ─────────────────────────────────────────────────────
function ImageEditorModal({ onClose, onSave }) {
  const [img, setImg] = useState(null);
  const [title, setTitle] = useState("");
  const [scale, setScale] = useState(100);
  const [bright, setBright] = useState(100);
  const [contrast, setContrast] = useState(100);
  const [saturate, setSaturate] = useState(100);
  const [origW, setOrigW] = useState(0);
  const [origH, setOrigH] = useState(0);

  const load = f => {
    if (!f || !f.type.startsWith('image/')) return;
    const r = new FileReader();
    r.onload = ev => { const src=ev.target.result, i=new Image(); i.onload=()=>{setOrigW(i.width);setOrigH(i.height);setImg(src);}; i.src=src; };
    r.readAsDataURL(f);
  };
  const dispW = Math.round(origW*scale/100), dispH = Math.round(origH*scale/100);
  const lS = { fontSize:11,color:TX_M,textTransform:"uppercase",letterSpacing:1,display:"block",marginBottom:5,fontFamily:SF,fontWeight:"600" };
  const iS2 = { width:"100%",padding:"10px 14px",background:APP_BG,border:`1.5px solid ${BDR}`,borderRadius:9,color:TX_P,fontSize:13,outline:"none",fontFamily:SF,boxSizing:"border-box" };
  const sliders = [
    {l:"Scale",v:scale,s:setScale,mn:10,mx:200,u:"%"},
    {l:"Brightness",v:bright,s:setBright,mn:20,mx:200,u:"%"},
    {l:"Contrast",v:contrast,s:setContrast,mn:20,mx:200,u:"%"},
    {l:"Saturation",v:saturate,s:setSaturate,mn:0,mx:200,u:"%"},
  ];

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(44,10,4,0.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9998,backdropFilter:"blur(8px)"}}>
      <div style={{background:CARD_BG,border:`1px solid ${BDR_H}`,borderRadius:20,padding:28,width:680,maxWidth:"92vw",maxHeight:"90vh",overflowY:"auto",boxShadow:`0 32px 80px rgba(139,60,20,0.25)`}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:22}}>
          <div>
            <h2 style={{fontSize:18,fontWeight:"700",color:MAR,fontFamily:AF,margin:0}}>🖼 Gallery Image Editor</h2>
            <p style={{fontSize:12,color:TX_M,margin:"4px 0 0",fontFamily:SF}}>Upload, adjust and place on Gallery Board</p>
          </div>
          <button onClick={onClose} style={{background:APP_BG,border:`1px solid ${BDR}`,borderRadius:9,color:TX_M,cursor:"pointer",padding:"7px 14px",fontSize:13,fontFamily:SF}}>✕</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:22}}>
          <div>
            <div style={{marginBottom:14}}><label style={lS}>Title</label><input value={title} onChange={e=>setTitle(e.target.value)} placeholder="e.g. Camp Photo…" style={iS2} /></div>
            <div style={{marginBottom:14}} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();load(e.dataTransfer.files[0]);}}>
              <label style={lS}>Upload Image</label>
              <label style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,padding:"16px",background:APP_BG,border:`2px dashed ${BDR_H}`,borderRadius:10,cursor:"pointer",color:SAF,fontSize:13,fontFamily:SF,fontWeight:"600"}}>
                📁 Choose File <input type="file" accept="image/*" style={{display:"none"}} onChange={e=>load(e.target.files[0])} />
              </label>
            </div>
            {img && (
              <div>
                {sliders.map(({l,v,s,mn,mx,u}) => (
                  <div key={l} style={{marginBottom:11}}>
                    <label style={lS}>{l}: <span style={{color:SAF,fontWeight:"700"}}>{v}{u}</span></label>
                    <input type="range" min={mn} max={mx} value={v} onChange={e=>s(Number(e.target.value))} style={{width:"100%",accentColor:SAF}} />
                  </div>
                ))}
                <div style={{display:"flex",gap:8,marginTop:8}}>
                  <button onClick={() => {setScale(100);setBright(100);setContrast(100);setSaturate(100);}} style={{padding:"8px 14px",background:APP_BG,border:`1px solid ${BDR}`,borderRadius:8,color:TX_M,cursor:"pointer",fontSize:11,fontFamily:SF}}>↺ Reset</button>
                  <button onClick={() => onSave({src:img,title:title||"Image",displayW:Math.min(dispW,360),displayH:Math.min(dispH,280),filter:`brightness(${bright}%) contrast(${contrast}%) saturate(${saturate}%)`})}
                    style={{flex:1,padding:"11px",background:`linear-gradient(135deg,${SAF},${SAF_D})`,border:"none",borderRadius:9,color:"#fff",cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",boxShadow:`0 4px 16px ${SAF_G}`}}>
                    ✅ Save to Gallery
                  </button>
                </div>
              </div>
            )}
          </div>
          <div>
            <label style={lS}>Preview</label>
            <div onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();load(e.dataTransfer.files[0]);}}
              style={{border:`2px dashed ${BDR_H}`,borderRadius:12,minHeight:260,display:"flex",alignItems:"center",justifyContent:"center",background:APP_BG,overflow:"hidden"}}>
              {img
                ? <img src={img} alt="preview" style={{maxWidth:"100%",maxHeight:280,filter:`brightness(${bright}%) contrast(${contrast}%) saturate(${saturate}%)`,transform:`scale(${scale/100})`,transformOrigin:"center",objectFit:"contain"}} />
                : <div style={{textAlign:"center",color:TX_M,padding:24,fontFamily:SF}}><div style={{fontSize:48,marginBottom:10}}>🖼</div><div style={{fontSize:12}}>Upload or drag image</div></div>
              }
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ═══ Loading Splash ────────────────────────────────────────────────────────
function LoadingSplash({ stage }) {
  const stages = ["Initialising…","Loading records…","Restoring settings…","Almost ready…"];
  return (
    <div style={{position:"fixed",inset:0,background:`linear-gradient(135deg,${SB_BG},${SB_BG2},${MAR})`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",zIndex:99999,fontFamily:AF}}>
      <div style={{fontSize:64,marginBottom:20}}>🩸</div>
      <div style={{fontSize:20,fontWeight:"700",color:"#fff",letterSpacing:3,textTransform:"uppercase",marginBottom:4,textAlign:"center"}}>Aniruddha's Academy of Disaster Management</div>
      <div style={{fontSize:11,color:SAF_L,letterSpacing:2,textTransform:"uppercase",marginBottom:36,fontFamily:SF,fontWeight:"600"}}>Mega Blood Donation Camp</div>
      <div style={{width:280,height:5,background:"rgba(255,255,255,0.12)",borderRadius:3,overflow:"hidden",marginBottom:16}}>
        <div style={{height:"100%",borderRadius:3,background:`linear-gradient(90deg,${SAF_D},${SAF},${SAF_L})`,width:`${(stage+1)*25}%`,transition:"width 0.5s ease"}} />
      </div>
      <div style={{fontSize:12,color:TX_WM,fontFamily:SF}}>{stages[Math.min(stage,3)]}</div>
    </div>
  );
}

// ═══ Main App ══════════════════════════════════════════════════════════════
export default function BloodDonorApp() {
  const [appReady,    setAppReady]    = useState(false);
  const [loadStage,   setLoadStage]   = useState(0);
  const [donors,      setDonors]      = useState([]);
  const [form,        setForm]        = useState(INIT_FORM);
  const [view,        setView]        = useState("dashboard");
  const [bg,          setBg]          = useState(null);
  const [bgOpacity,   setBgOpacity]   = useState(0.12);
  const [bannerImg,   setBannerImg]   = useState(null);
  const [bannerOp,    setBannerOp]    = useState(0.3);
  const [bannerIcon,  setBannerIcon]  = useState(null);
  const [sidebarIcon, setSidebarIcon] = useState(null);
  const [sidebarDecorImg,  setSidebarDecorImg]  = useState(null);
  const [sidebarDecorOp,   setSidebarDecorOp]   = useState(0.5);   // blend opacity
  const [sidebarDecorX,    setSidebarDecorX]    = useState(50);    // objectPosition X %
  const [sidebarDecorY,    setSidebarDecorY]    = useState(50);    // objectPosition Y %
  const [sidebarDecorSc,   setSidebarDecorSc]   = useState(100);   // scale / zoom %
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [search,      setSearch]      = useState("");
  const [filterBG,    setFilterBG]    = useState("All");
  const [filterSt,    setFilterSt]    = useState("All");
  const [filterTy,    setFilterTy]    = useState("All");
  const [editId,      setEditId]      = useState(null);
  const [toast,       setToast]       = useState(null);
  const [bkNotif,     setBkNotif]     = useState(null);
  const [sortField,   setSortField]   = useState("srNo");
  const [sortDir,     setSortDir]     = useState("asc");
  const [delConfirm,  setDelConfirm]  = useState(null);
  const [clearStep,   setClearStep]   = useState(0);
  const [activeTab,   setActiveTab]   = useState("form");
  const [lastSaved,   setLastSaved]   = useState(null);
  const [saveErr,     setSaveErr]     = useState(false);
  const [showImgEd,   setShowImgEd]   = useState(false);
  const [bgEdMode,    setBgEdMode]    = useState(null);
  const [gallery,     setGallery]     = useState([]);
  const [interact,    setInteract]    = useState(null);
  const [nextSrNo,    setNextSrNo]    = useState(1);
  const [bkInterval,  setBkInterval]  = useState(15);
  const [lastBkTime,  setLastBkTime]  = useState(null);
  const [restLog,     setRestLog]     = useState([]);
  const [storedBks,   setStoredBks]   = useState([]);
  const boardRef    = useRef();
  const fileRef     = useRef();
  const autoBkRef   = useRef(null);

  // ── Boot ─────────────────────────────────────────────────────────────
  useEffect(() => {
    const log = [];
    setLoadStage(0);
    const dr = sGet(SK.DONORS);
    if (dr) {
      try {
        const d = JSON.parse(dr);
        setDonors(d);
        if (d.length > 0) setNextSrNo(Math.max(...d.map(x=>x.srNo||0))+1);
        log.push(`✅ Donors: ${d.length} record(s) restored`);
      } catch { log.push("⚠️ Donor data corrupted"); }
    } else { log.push("ℹ️ Fresh start — no records yet"); }

    setLoadStage(1);
    const gi = sGet(SK.GAL_IDX);
    if (gi) {
      try {
        const idx = JSON.parse(gi);
        const full = idx.map(m => { const src=sGet(galKey(m.id)); return {...m,src:src||null}; });
        setGallery(full);
        log.push(`✅ Gallery: ${full.length} image(s) restored`);
      } catch { log.push("⚠️ Gallery data corrupted"); }
    }

    setLoadStage(2);
    const bgR = sGet(SK.BG); if (bgR) setBg(bgR);
    const opR = sGet(SK.OPACITY); if (opR) setBgOpacity(parseFloat(opR));
    const biR = sGet(SK.BANNER_IMG); if (biR) setBannerImg(biR);
    const boR = sGet(SK.BANNER_OPACITY); if (boR) setBannerOp(parseFloat(boR));
    const bnR = sGet(SK.BANNER_ICON); if (bnR) setBannerIcon(bnR);
    const siR = sGet(SK.SIDEBAR_ICON); if (siR) setSidebarIcon(siR);
    const sdR  = sGet(SK.SIDEBAR_DECOR);    if (sdR)  setSidebarDecorImg(sdR);
    const sdOp = sGet(SK.SIDEBAR_DECOR_OP); if (sdOp) setSidebarDecorOp(parseFloat(sdOp));
    const sdX  = sGet(SK.SIDEBAR_DECOR_X);  if (sdX)  setSidebarDecorX(parseFloat(sdX));
    const sdY  = sGet(SK.SIDEBAR_DECOR_Y);  if (sdY)  setSidebarDecorY(parseFloat(sdY));
    const sdSc = sGet(SK.SIDEBAR_DECOR_SC); if (sdSc) setSidebarDecorSc(parseFloat(sdSc));
    const bkR = sGet(SK.BACKUP_INTERVAL); if (bkR) setBkInterval(parseInt(bkR)||15);
    const soR = sGet(SK.SIDEBAR_OPEN); if (soR !== null) setSidebarOpen(soR === "true");
    const bks = Object.keys(localStorage).filter(k=>k.startsWith("bdms-backup-")).sort().reverse().slice(0,5);
    setStoredBks(bks);
    log.push("✅ Settings restored");

    setLoadStage(3);
    setTimeout(() => {
      setRestLog(log);
      setAppReady(true);
      const dl = log.find(l=>l.includes("record"));
      if (dl) setTimeout(() => showToast(dl.replace(/[✅ℹ️⚠️] /g,""),"info"), 400);
    }, 400);
  }, []);

  // ── Auto-backup timer ────────────────────────────────────────────────
  useEffect(() => {
    if (autoBkRef.current) clearInterval(autoBkRef.current);
    autoBkRef.current = setInterval(() => { if (donors.length > 0) doBackup(); }, bkInterval*60*1000);
    return () => { if (autoBkRef.current) clearInterval(autoBkRef.current); };
  }, [bkInterval, donors]);

  // ── Auto-save every 30s ──────────────────────────────────────────────
  useEffect(() => {
    const t = setInterval(() => {
      if (donors.length > 0) { const ok=sSet(SK.DONORS,JSON.stringify(donors)); if(ok) setLastSaved(nowTs()); }
    }, 30000);
    return () => clearInterval(t);
  }, [donors]);

  const doBackup = () => {
    const ts = Date.now(), fname = bkFname();
    const data = JSON.stringify({donors,exportedAt:nowDT(),backupFile:fname,version:"1.0"},null,2);
    sSet(`bdms-backup-${ts}`, data);
    const allBk = Object.keys(localStorage).filter(k=>k.startsWith("bdms-backup-")).sort();
    while (allBk.length > 5) { sDel(allBk.shift()); }
    setStoredBks(Object.keys(localStorage).filter(k=>k.startsWith("bdms-backup-")).sort().reverse().slice(0,5));
    try { const blob=new Blob([data],{type:"application/json"}); const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=fname; a.click(); } catch {}
    setLastBkTime(nowTs());
    setBkNotif({fname, time:nowTs()});
    setTimeout(() => setBkNotif(null), 8000);
  };

  const showToast = (msg, type="success") => { setToast({msg,type}); setTimeout(()=>setToast(null),3800); };

  const saveDonors = list => {
    setDonors(list);
    if (list.length > 0) setNextSrNo(Math.max(...list.map(x=>x.srNo||0))+1); else setNextSrNo(1);
    const ok = sSet(SK.DONORS, JSON.stringify(list));
    if (ok) { setSaveErr(false); setLastSaved(nowTs()); }
    else { setSaveErr(true); showToast("⚠️ Save failed! Export a backup.","error"); }
  };

  const saveGallery = list => {
    setGallery(list);
    const idx = list.map(({src:_s,...m}) => m);
    sSet(SK.GAL_IDX, JSON.stringify(idx));
    list.forEach(img => { if (img.src) sSet(galKey(img.id), img.src); });
  };

  const handleFormChange = (k, v) => setForm(f => { const u={...f,[k]:v}; if(k==="dob") u.age=calcAge(v); return u; });

  const validate = () => {
    if (!form.name.trim()) return "Name of Shraddhavan is required";
    if (!form.shraddhavanaType) return "Please select New or Old";
    if (!form.mobile.trim() || !/^\d{10}$/.test(form.mobile)) return "Mobile must be 10 digits";
    if (!form.gender) return "Please select Gender";
    if (!form.bloodGroup) return "Please select Blood Group";
    if (!form.status) return "Please select Accepted or Rejected";
    if (form.status==="Rejected" && !form.rejectionReason.trim()) return "Reason for Rejection is required";
    return null;
  };

  const handleSubmit = () => {
    const err = validate(); if (err) { showToast(err,"error"); return; }
    if (editId) {
      saveDonors(donors.map(d => d.id===editId ? {...form,id:editId,registeredAt:d.registeredAt,registeredOn:d.registeredOn} : d));
      showToast("Record updated! ✅"); setEditId(null);
    } else {
      saveDonors([...donors, {...form, id:Date.now(), srNo:nextSrNo, registeredOn:todayStr(), registeredAt:nowDT()}]);
      showToast("Shraddhavan registered! ✅");
    }
    setForm(INIT_FORM); setView("database");
  };

  const handleEdit   = d => { setForm({...d}); setEditId(d.id); setView("add"); setActiveTab("form"); };
  const handleDelete = id => { saveDonors(donors.filter(d=>d.id!==id)); setDelConfirm(null); showToast("Record removed.","info"); };
  const handleSort   = f  => { if(sortField===f) setSortDir(d=>d==="asc"?"desc":"asc"); else { setSortField(f); setSortDir("asc"); } };

  const exportExcel = () => {
    if (!donors.length) { showToast("No data to export!","error"); return; }
    const wb = XLSX.utils.book_new();
    const data = donors.map(d => ({"Sr.No":d.srNo,"Type":d.shraddhavanaType,"Upasana Kendra":d.upasanaKendra,"Name of Shraddhavan":d.name,"Mobile":d.mobile,"Date of Birth":d.dob,"Age":d.age,"Gender":d.gender,"Blood Group":d.bloodGroup,"Status":d.status,"Reason for Rejection":d.rejectionReason||"","Registered On":d.registeredOn,"Registered At":d.registeredAt||d.registeredOn||""}));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{wch:7},{wch:10},{wch:20},{wch:24},{wch:14},{wch:14},{wch:6},{wch:10},{wch:12},{wch:10},{wch:24},{wch:14},{wch:22}];
    XLSX.utils.book_append_sheet(wb, ws, "Shraddhavans");
    const S2 = [];
    const kv = (c,v) => ({"Category":c,"Details":"","Count / Value":v});
    const hd = t  => ({"Category":`── ${t} ──`,"Details":"","Count / Value":""});
    const bl = () => ({"Category":"","Details":"","Count / Value":""});
    S2.push(kv("Report Generated On",nowDT()));
    S2.push(kv("Organisation","Sadguru Shree Aniruddha Upasana Kendra, Wai"));
    S2.push(kv("Camp","Mega Blood Donation Camp"));
    S2.push(bl());
    S2.push(hd("OVERALL STATISTICS"));
    S2.push(kv("Total",stats.total)); S2.push(kv("New",stats.newS)); S2.push(kv("Old",stats.oldS));
    S2.push(kv("Accepted",stats.accepted)); S2.push(kv("Rejected",stats.rejected)); S2.push(bl());
    S2.push(hd("BLOOD GROUPS"));
    BLOOD_GROUPS.forEach(g => { const c=stats.byGroup[g]||0, p=stats.total?Math.round(c/stats.total*100):0; S2.push({"Category":g,"Details":`${p}%`,"Count / Value":c}); });
    S2.push(bl()); S2.push(hd("GENDER"));
    Object.entries(stats.byGender).forEach(([g,c]) => S2.push({"Category":g,"Details":`${stats.total?Math.round(c/stats.total*100):0}%`,"Count / Value":c}));
    S2.push(bl()); S2.push(hd("TOP UPASANA KENDRAS"));
    stats.byKendra.forEach((k,i) => S2.push({"Category":`${i+1}. ${k.name}`,"Details":"","Count / Value":k.count}));
    const ws2 = XLSX.utils.json_to_sheet(S2); ws2["!cols"] = [{wch:36},{wch:16},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws2, "Summary");
    XLSX.writeFile(wb, "AADM_Shraddhavan_Donors_Data.xlsx");
    showToast("Excel downloaded — includes Summary sheet!");
  };

  const importExcel = e => {
    const f = e.target.files[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => {
      const wb = XLSX.read(ev.target.result,{type:"binary"});
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const maxSr = donors.length ? Math.max(...donors.map(x=>x.srNo||0)) : 0;
      const imp = rows.map((r,i) => ({id:Date.now()+i,srNo:maxSr+i+1,shraddhavanaType:r["Type"]||"",upasanaKendra:r["Upasana Kendra"]||"",name:r["Name of Shraddhavan"]||"",mobile:r["Mobile"]||"",dob:r["Date of Birth"]||"",age:r["Age"]||"",gender:r["Gender"]||"",bloodGroup:r["Blood Group"]||"",status:r["Status"]||"",rejectionReason:r["Reason for Rejection"]||"",registeredOn:r["Registered On"]||todayStr(),registeredAt:r["Registered At"]||r["Registered On"]||nowDT()}));
      saveDonors([...donors,...imp]); showToast(`${imp.length} records imported!`);
    };
    reader.readAsBinaryString(f); e.target.value = "";
  };

  const exportJSON = () => {
    if (!donors.length) { showToast("No data!","error"); return; }
    const blob = new Blob([JSON.stringify({donors,exportedAt:nowDT(),version:"1.0"},null,2)],{type:"application/json"});
    const a = document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=`MBDC_Backup_${new Date().toISOString().slice(0,10)}.json`; a.click();
    showToast("JSON backup downloaded!");
  };

  const importJSON = e => {
    const f = e.target.files[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = ev => {
      try { const d=JSON.parse(ev.target.result); const list=Array.isArray(d)?d:(d.donors||[]); if(!list.length){showToast("No records in file","error");return;} saveDonors(list); showToast(`${list.length} records restored!`); }
      catch { showToast("Invalid backup file","error"); }
    };
    reader.readAsText(f); e.target.value = "";
  };

  const dlStoredBk = k => {
    const data = sGet(k); if (!data) return;
    const blob = new Blob([data],{type:"application/json"});
    const a = document.createElement("a"); a.href=URL.createObjectURL(blob);
    const ts = k.replace("bdms-backup-",""); a.download=`MBDC_Backup_${new Date(parseInt(ts)).toISOString().slice(0,16).replace("T","_").replace(":","-")}.json`; a.click();
  };

  const removeGalleryImg = id => { const u=gallery.filter(i=>i.id!==id); saveGallery(u); sDel(galKey(id)); showToast("Image removed","info"); };

  const startInteract = (e, id, type) => {
    e.preventDefault(); e.stopPropagation();
    const img = gallery.find(i=>i.id===id); if (!img) return;
    if (type==="drag") setInteract({type:"drag",id,offsetX:e.clientX-img.x,offsetY:e.clientY-img.y});
    else setInteract({type:"resize",id,startX:e.clientX,startY:e.clientY,origW:img.displayW,origH:img.displayH});
  };
  const handleBoardMove = e => {
    if (!interact || !boardRef.current) return;
    const rect = boardRef.current.getBoundingClientRect();
    if (interact.type==="drag") {
      const x=Math.max(0,Math.min(e.clientX-interact.offsetX,rect.width-60)), y=Math.max(0,Math.min(e.clientY-interact.offsetY,rect.height-60));
      setGallery(gs => gs.map(g => g.id===interact.id ? {...g,x,y} : g));
    } else {
      const dx=e.clientX-interact.startX, newW=Math.max(60,interact.origW+dx), ratio=interact.origH/interact.origW;
      setGallery(gs => gs.map(g => g.id===interact.id ? {...g,displayW:newW,displayH:Math.round(newW*ratio)} : g));
    }
  };
  const stopInteract = () => {
    if (interact) { const idx=gallery.map(({src:_s,...m})=>m); sSet(SK.GAL_IDX,JSON.stringify(idx)); setInteract(null); }
  };

  const kens = [...new Set(donors.map(d=>d.upasanaKendra).filter(Boolean))];
  const stats = {
    total: donors.length,
    newS:  donors.filter(d=>d.shraddhavanaType==="New").length,
    oldS:  donors.filter(d=>d.shraddhavanaType==="Old").length,
    accepted: donors.filter(d=>d.status==="Accepted").length,
    rejected: donors.filter(d=>d.status==="Rejected").length,
    byGroup: BLOOD_GROUPS.reduce((a,g) => { a[g]=donors.filter(d=>d.bloodGroup===g).length; return a; }, {}),
    byGender: {Male:donors.filter(d=>d.gender==="Male").length, Female:donors.filter(d=>d.gender==="Female").length, Others:donors.filter(d=>d.gender==="Others").length},
    byKendra: kens.map(k=>({name:k,count:donors.filter(d=>d.upasanaKendra===k).length})).sort((a,b)=>b.count-a.count),
  };

  const filtered = donors
    .filter(d => {
      const q = search.toLowerCase();
      return (!q || d.name?.toLowerCase().includes(q) || d.mobile?.includes(q) || d.upasanaKendra?.toLowerCase().includes(q))
        && (filterBG==="All" || d.bloodGroup===filterBG)
        && (filterSt==="All" || d.status===filterSt)
        && (filterTy==="All" || d.shraddhavanaType===filterTy);
    })
    .sort((a,b) => {
      let av=a[sortField]||"", bv=b[sortField]||"";
      if (["srNo","age"].includes(sortField)) { av=+av; bv=+bv; }
      return sortDir==="asc" ? (av>bv?1:-1) : (av<bv?1:-1);
    });

  const SW = sidebarOpen ? 256 : 68;
  const bannerBg = bannerImg
    ? `linear-gradient(rgba(59,12,12,${1-bannerOp}),rgba(77,16,16,${1-bannerOp})),url(${bannerImg}) center/cover`
    : `linear-gradient(135deg,${SB_BG} 0%,${SB_BG2} 40%,${MAR} 100%)`;
  const appBgStyle = bg
    ? {background:`linear-gradient(rgba(247,239,227,${1-bgOpacity}),rgba(247,239,227,${1-bgOpacity})),url(${bg}) center/cover fixed`}
    : {background:APP_BG};

  const iS = {width:"100%",padding:"11px 14px",background:CARD_BG,border:`1.5px solid ${BDR}`,borderRadius:9,color:TX_P,fontSize:14,outline:"none",fontFamily:SF,transition:"all 0.2s",boxSizing:"border-box",minHeight:44};
  const lS = {fontSize:11,color:TX_M,textTransform:"uppercase",letterSpacing:1,display:"block",marginBottom:5,fontFamily:SF,fontWeight:"600"};

  const card = {background:CARD_BG,border:`1px solid ${BDR}`,borderRadius:16,overflow:"hidden",boxShadow:`0 2px 12px rgba(139,60,20,0.08)`};

  const pBtn = {background:`linear-gradient(135deg,${SAF},${SAF_D})`,color:"#fff",border:"none",padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",boxShadow:`0 4px 16px ${SAF_G}`,transition:"all 0.18s"};
  const gBtn = {background:"transparent",color:SAF,border:`1.5px solid ${BDR_H}`,padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",transition:"all 0.18s"};
  const g2Btn = {background:APP_BG,color:TX_S,border:`1.5px solid ${BDR}`,padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",transition:"all 0.18s"};
  const dBtn = {background:RED_L,color:MAR,border:`1px solid rgba(220,38,38,0.2)`,padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",transition:"all 0.18s"};
  const sBtn = {background:`linear-gradient(135deg,${GRN},#047857)`,color:"#fff",border:"none",padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",transition:"all 0.18s"};
  const gdBtn = {background:`linear-gradient(135deg,${GOLD},#B45309)`,color:"#fff",border:"none",padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",transition:"all 0.18s"};
  const iBtn = {background:"rgba(37,99,235,0.1)",color:BLUE,border:`1px solid rgba(37,99,235,0.25)`,padding:"11px 20px",borderRadius:9,cursor:"pointer",fontSize:13,fontFamily:SF,fontWeight:"700",transition:"all 0.18s"};

  const navItemStyle = active => ({
    display:"flex", alignItems:"center", gap:sidebarOpen?12:0, padding:sidebarOpen?"10px 16px":"11px",
    borderRadius:10, cursor:"pointer", transition:"all 0.22s",
    justifyContent:sidebarOpen?"flex-start":"center",
    background: active ? `linear-gradient(135deg,${SAF},${SAF_D})` : "transparent",
    color: active ? "#fff" : TX_WM,
    fontSize:13, fontFamily:SF, fontWeight:active?"700":"500",
    whiteSpace:"nowrap", overflow:"hidden",
    boxShadow: active ? `0 4px 14px ${SAF_G}, inset 0 1px 0 rgba(255,255,255,0.15)` : "none",
  });

  const NAVS = [
    {key:"dashboard", icon:"▦", label:"Dashboard"},
    {key:"add",       icon:"✦", label:editId?"Edit Record":"Add Shraddhavan"},
    {key:"database",  icon:"⊟", label:"Database"},
    {key:"gallery",   icon:"◫", label:"Gallery Board"},
    {key:"settings",  icon:"✧", label:"Settings"},
  ];

  if (!appReady) return <LoadingSplash stage={loadStage} />;

  const donutData = [
    {val:stats.byGender.Male,   color:BLUE,   label:"Male"},
    {val:stats.byGender.Female, color:"#DB2777", label:"Female"},
    {val:stats.byGender.Others, color:PURPLE,  label:"Others"},
  ];
  const slices = donutSlices(donutData, 44, 60, 60);

  const thStyle = {padding:"11px 13px",textAlign:"left",fontSize:10,textTransform:"uppercase",letterSpacing:1.5,color:TX_M,borderBottom:`1px solid ${BDR}`,cursor:"pointer",userSelect:"none",whiteSpace:"nowrap",fontFamily:SF,fontWeight:"700"};
  const tdStyle = {padding:"11px 13px",fontSize:12,color:TX_P,background:CARD_BG,fontFamily:SF,transition:"background 0.15s"};

  return (
    <div style={{minHeight:"100vh",fontFamily:SF,color:TX_P,...appBgStyle}} onMouseMove={handleBoardMove} onMouseUp={stopInteract}>

      {/* TOP BANNER */}
      <div style={{position:"fixed",top:0,left:0,right:0,height:64,background:bannerBg,display:"flex",alignItems:"center",zIndex:200,boxShadow:`0 4px 20px rgba(44,10,4,0.35)`}}>
        <div style={{width:SW,display:"flex",alignItems:"center",justifyContent:"center",transition:"width 0.32s",flexShrink:0}}>
          {bannerIcon
            ? <img src={bannerIcon} alt="icon" style={{width:40,height:40,objectFit:"contain",borderRadius:"50%",border:"2px solid rgba(255,255,255,0.25)"}} />
            : <div style={{width:40,height:40,borderRadius:"50%",background:`linear-gradient(135deg,${SAF},${SAF_D})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20}}>🩸</div>
          }
        </div>
        <div style={{flex:1}}>
          <div style={{fontSize:15,fontWeight:"700",color:"#fff",letterSpacing:2,textTransform:"uppercase",fontFamily:AF,textShadow:"0 1px 12px rgba(0,0,0,0.5)",lineHeight:1.2}}>Aniruddha's Academy of Disaster Management</div>
          <div style={{fontSize:11,color:SAF_L,letterSpacing:3,textTransform:"uppercase",marginTop:3,fontFamily:SF,fontWeight:"600"}}>Mega Blood Donation Camp</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,marginRight:20,flexShrink:0}}>
			{/*<button onClick={() => setBgEdMode("banner")} style={{background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:9,color:"rgba(255,255,255,0.85)",cursor:"pointer",padding:"8px 12px",fontSize:12,fontFamily:SF,fontWeight:"600"}}>🎨 Banner</button>*/}
          <button onClick={() => { setForm(INIT_FORM); setEditId(null); setView("add"); setActiveTab("form"); }} style={{...pBtn, padding:"10px 18px"}}>+ New Shraddhavan Donor</button>
			  {/*<label style={{background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:9,color:"rgba(255,255,255,0.85)",cursor:"pointer",padding:"8px 12px",fontSize:12,fontFamily:SF,fontWeight:"600"}}>
            🏷 Icon
            <input type="file" accept="image/*" style={{display:"none"}} onChange={e => {
              const f=e.target.files[0]; if(!f) return;
              const r=new FileReader(); r.onload=ev=>{ const u=ev.target.result; setBannerIcon(u); setSidebarIcon(u); sSet(SK.BANNER_ICON,u); sSet(SK.SIDEBAR_ICON,u); showToast("Icon updated!"); }; r.readAsDataURL(f);
            }} />
			  </label>*/}
        </div>
      </div>

      {/* SIDEBAR */}
      <div style={{position:"fixed",left:0,top:64,bottom:0,width:SW,background:`linear-gradient(175deg,${SB_BG} 0%,${SB_BG2} 35%,#3F0D0D 65%,${SB_BG} 100%)`,borderRight:"1px solid rgba(255,255,255,0.05)",display:"flex",flexDirection:"column",zIndex:100,transition:"width 0.32s cubic-bezier(0.4,0,0.2,1)",overflow:"hidden",boxShadow:`4px 0 32px rgba(44,10,4,0.5)`}}>
        {/* Glowing top accent */}
        <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:`linear-gradient(90deg,transparent,${SAF},transparent)`,zIndex:1}} />
        {/* Subtle gradient overlay */}
        <div style={{position:"absolute",inset:0,background:`radial-gradient(ellipse at 30% 20%, rgba(232,100,26,0.06) 0%,transparent 60%)`,pointerEvents:"none",zIndex:0}} />

        {/* Logo */}
        <div style={{padding:sidebarOpen?"20px 18px 14px":"12px 8px",borderBottom:"1px solid rgba(255,255,255,0.06)",position:"relative",zIndex:1}}>
          {sidebarOpen ? (
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              {sidebarIcon
                ? <img src={sidebarIcon} alt="logo" style={{width:44,height:44,objectFit:"contain",borderRadius:10,border:`2px solid ${SAF}40`,flexShrink:0}} />
                : <div style={{width:44,height:44,borderRadius:10,background:`linear-gradient(135deg,${SAF},${SAF_D})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0,boxShadow:`0 4px 14px ${SAF_G}`}}>🩸</div>
              }
              <div style={{overflow:"hidden"}}>
                <div style={{fontSize:11.5,fontWeight:"700",color:"#fff",lineHeight:1.4,fontFamily:AF}}>Sadguru Shree Aniruddha<br/>Upasana Kendra, Wai</div>
                <div style={{fontSize:9,color:SAF,letterSpacing:2,textTransform:"uppercase",marginTop:5,fontFamily:SF,fontWeight:"700"}}>Donor Registry</div>
              </div>
            </div>
          ) : (
            <div style={{display:"flex",justifyContent:"center"}}>
              {sidebarIcon
                ? <img src={sidebarIcon} alt="logo" style={{width:36,height:36,objectFit:"contain",borderRadius:8}} />
                : <div style={{width:36,height:36,borderRadius:8,background:`linear-gradient(135deg,${SAF},${SAF_D})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,boxShadow:`0 4px 12px ${SAF_G}`}}>🩸</div>
              }
            </div>
          )}
          {sidebarOpen && (
            <label style={{display:"block",marginTop:8,cursor:"pointer",fontSize:9.5,color:"rgba(255,255,255,0.25)",textAlign:"right",fontFamily:SF}}>
              Set logo ✎
              <input type="file" accept="image/*" style={{display:"none"}} onChange={e => { const f=e.target.files[0]; if(!f) return; const r=new FileReader(); r.onload=ev=>{ setSidebarIcon(ev.target.result); sSet(SK.SIDEBAR_ICON,ev.target.result); showToast("Sidebar logo updated!"); }; r.readAsDataURL(f); }} />
            </label>
          )}
        </div>

        {sidebarOpen && <div style={{padding:"16px 18px 6px",position:"relative",zIndex:1}}><div style={{fontSize:9,color:"rgba(255,255,255,0.25)",textTransform:"uppercase",letterSpacing:2,fontFamily:SF,fontWeight:"700"}}>Menu</div></div>}

        <nav style={{flex:1,padding:"6px 8px",display:"flex",flexDirection:"column",gap:3,position:"relative",zIndex:1}}>
          {NAVS.map(item => (
            <div key={item.key} style={navItemStyle(view===item.key)} onClick={() => setView(item.key)} title={!sidebarOpen?item.label:""}>
              <span style={{fontSize:15,flexShrink:0,fontFamily:MF}}>{item.icon}</span>
              {sidebarOpen && <span style={{fontSize:13}}>{item.label}</span>}
              {sidebarOpen && item.key==="database" && donors.length>0 && (
                <span style={{marginLeft:"auto",background:view==="database"?"rgba(255,255,255,0.25)":`${SAF}22`,color:view==="database"?"#fff":SAF,fontSize:10,fontFamily:MF,padding:"2px 8px",borderRadius:8,fontWeight:"700"}}>{donors.length}</span>
              )}
            </div>
          ))}
        </nav>

        <div style={{padding:"8px",borderTop:"1px solid rgba(255,255,255,0.06)",position:"relative",zIndex:1}}>

          {/* ── Sidebar decorative image (controls are in Settings) ── */}
          {sidebarOpen && (
            sidebarDecorImg ? (
              <div style={{position:"relative",marginBottom:10,borderRadius:12,overflow:"hidden",height:110}}>
                <img
                  src={sidebarDecorImg}
                  alt="sidebar decor"
                  style={{
                    position:"absolute",
                    width:`${sidebarDecorSc}%`, height:`${sidebarDecorSc}%`,
                    minWidth:"100%", minHeight:"100%",
                    objectFit:"cover",
                    objectPosition:`${sidebarDecorX}% ${sidebarDecorY}%`,
                    left:"50%", top:"50%",
                    transform:"translate(-50%,-50%)",
                    mixBlendMode:"luminosity",
                    opacity:sidebarDecorOp,
                    filter:"sepia(50%) saturate(70%) brightness(0.8) hue-rotate(340deg)",
                    display:"block",
                  }}
                />
                <div style={{position:"absolute",inset:0,pointerEvents:"none",background:`linear-gradient(180deg,${SB_BG} 0%,transparent 30%,transparent 70%,${SB_BG} 100%)`}}/>
                <div style={{position:"absolute",inset:0,pointerEvents:"none",background:`linear-gradient(90deg,${SB_BG} 0%,transparent 18%,transparent 82%,${SB_BG} 100%)`}}/>
                <div style={{position:"absolute",inset:0,pointerEvents:"none",background:`radial-gradient(ellipse at 50% 60%,rgba(232,100,26,0.07) 0%,transparent 70%)`}}/>
              </div>
            ) : (
              <label style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,marginBottom:10,padding:"9px",background:"rgba(255,255,255,0.03)",border:"1px dashed rgba(255,255,255,0.1)",borderRadius:10,cursor:"pointer",color:"rgba(255,255,255,0.22)",fontSize:10,fontFamily:SF,transition:"all 0.2s"}}>
                🖼 Add image to sidebar
                <input type="file" accept="image/*" style={{display:"none"}} onChange={e => {
                  const f=e.target.files[0]; if(!f) return;
                  const r=new FileReader();
                  r.onload=ev=>{ setSidebarDecorImg(ev.target.result); sSet(SK.SIDEBAR_DECOR,ev.target.result); };
                  r.readAsDataURL(f);
                }} />
              </label>
            )
          )}

          {/* ── Autosaved indicator + author ──────────────────────── */}
          {sidebarOpen && (
            <div style={{paddingLeft:8,marginBottom:10}}>
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:2}}>
                <div style={{width:7,height:7,borderRadius:"50%",background:saveErr?"#EF4444":GRN,flexShrink:0}} />
                <span style={{fontSize:10,color:"rgba(255,255,255,0.6)",fontFamily:SF,fontWeight:"600"}}>{saveErr?"Save Error":"Autosaved"}</span>
              </div>
              {lastSaved && <div style={{fontSize:9.5,color:"rgba(255,255,255,0.3)",fontFamily:MF,paddingLeft:13}}>{todayStr()} {lastSaved}</div>}
              <div style={{fontSize:9,color:"rgba(255,255,255,0.2)",fontFamily:SF,paddingLeft:13,marginTop:2}}>Crafted by Swanandsinh Ashok Mahamuni</div>
            </div>
          )}

          {/* ── Collapse / Expand button ──────────────────────────── */}
          <button
            onClick={() => { const n=!sidebarOpen; setSidebarOpen(n); sSet(SK.SIDEBAR_OPEN,String(n)); }}
            style={{width:"100%",background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:9,color:"rgba(255,255,255,0.5)",cursor:"pointer",padding:"9px",display:"flex",alignItems:"center",justifyContent:"center",gap:sidebarOpen?8:0,fontSize:13,fontFamily:SF,transition:"all 0.2s"}}>
            <span style={{transition:"transform 0.3s",transform:sidebarOpen?"none":"rotate(180deg)",lineHeight:1}}>◀</span>
				{/*{sidebarOpen && <span style={{fontSize:11}}>Collapse</span>}*/}
          </button>
        </div>
      </div>

      {/* MAIN */}
      <div style={{marginLeft:SW,paddingTop:64,padding:`${64+24}px 28px 28px 28px`,minHeight:"100vh",transition:"margin-left 0.32s cubic-bezier(0.4,0,0.2,1)"}}>

        <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:26}}>
          <div>
            {view==="dashboard" && <><div style={{fontSize:26,fontWeight:"300",color:TX_P,fontFamily:AF}}>Welcome 🌸</div><div style={{fontSize:14,color:TX_M,marginTop:4}}>Overview of registered blood donors</div></>}
            {view!=="dashboard" && <h1 style={{fontSize:22,fontWeight:"700",color:MAR,margin:0,fontFamily:AF}}>{view==="add"&&(editId?"Edit Shraddhavan":"Add New Shraddhavan")}{view==="database"&&"Shraddhavan Database"}{view==="gallery"&&"Gallery Image Board"}{view==="settings"&&"Settings & Configuration"}</h1>}
            <div style={{fontSize:11,color:TX_M,marginTop:5}}>{new Date().toLocaleDateString("en-IN",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</div>
          </div>
          <button onClick={() => setBgEdMode("wallpaper")} style={{...g2Btn,fontSize:12,padding:"9px 14px"}}>🖼 Wallpaper</button>
        </div>

        {/* ═══ DASHBOARD ═══════════════════════════════════════════════ */}
        {view==="dashboard" && (
          <div>
            {/* KPI row 1 */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:16}}>
              {[
                {label:"TOTAL DONORS",  value:stats.total,    color:MAR_M, iconBg:"rgba(153,27,27,0.1)", icon:"👥"},
                {label:"ACCEPTED",      value:stats.accepted, color:GRN,   iconBg:"rgba(5,150,105,0.1)", icon:"✅"},
                {label:"REJECTED",      value:stats.rejected, color:RED,   iconBg:"rgba(220,38,38,0.1)", icon:"❌"},
              ].map(s => (
                <div key={s.label} style={{...card,padding:"24px 26px",position:"relative",overflow:"hidden"}}>
                  <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:`linear-gradient(90deg,transparent,${s.color}60,transparent)`}} />
                  <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
                    <div>
                      <div style={{fontSize:11,fontWeight:"700",color:TX_M,letterSpacing:1.5,textTransform:"uppercase",marginBottom:10,fontFamily:SF}}>{s.label}</div>
                      <div style={{fontSize:38,fontWeight:"300",color:TX_P,fontFamily:AF,lineHeight:1}}>{s.value}</div>
                    </div>
                    <div style={{width:52,height:52,borderRadius:"50%",background:s.iconBg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>{s.icon}</div>
                  </div>
                </div>
              ))}
            </div>
            {/* KPI row 2 */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:16,marginBottom:20}}>
              {[
                {label:"NEW SHRADDHAVAN", value:stats.newS, color:SAF,    iconBg:SAF_XL,                   icon:"🆕"},
                {label:"OLD SHRADDHAVAN", value:stats.oldS, color:PURPLE, iconBg:"rgba(124,58,237,0.1)",   icon:"🔄"},
              ].map(s => (
                <div key={s.label} style={{...card,padding:"24px 26px",position:"relative",overflow:"hidden"}}>
                  <div style={{position:"absolute",top:0,left:0,right:0,height:3,background:`linear-gradient(90deg,transparent,${s.color}60,transparent)`}} />
                  <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
                    <div>
                      <div style={{fontSize:11,fontWeight:"700",color:TX_M,letterSpacing:1.5,textTransform:"uppercase",marginBottom:10,fontFamily:SF}}>{s.label}</div>
                      <div style={{fontSize:38,fontWeight:"300",color:TX_P,fontFamily:AF,lineHeight:1}}>{s.value}</div>
                    </div>
                    <div style={{width:52,height:52,borderRadius:"50%",background:s.iconBg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>{s.icon}</div>
                  </div>
                </div>
              ))}
            </div>

            {/* Charts */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:18,marginBottom:20}}>
              {/* Blood groups */}
              <div style={{...card,padding:22}}>
                <h3 style={{fontSize:13,fontWeight:"700",color:MAR,marginBottom:14,fontFamily:AF,display:"flex",alignItems:"center",gap:6}}><span style={{color:SAF}}>⬡</span> Blood Groups</h3>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  {BLOOD_GROUPS.map(g => {
                    const c=stats.byGroup[g]||0, pct=stats.total?Math.round(c/stats.total*100):0;
                    return (
                      <div key={g} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 10px",borderRadius:9,background:c>0?`${BG_C[g]}40`:APP_BG,border:`1px solid ${c>0?BG_C[g]+"60":BDR}`}}>
                        <span style={{fontWeight:"700",fontSize:13,color:c>0?BG_TX[g]:TX_M,fontFamily:SF}}>{g}</span>
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:16,fontWeight:"700",color:c>0?BG_TX[g]:TX_M,fontFamily:MF}}>{c}</div>
                          <div style={{fontSize:9,color:TX_M}}>{pct}%</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
              {/* Gender donut */}
              <div style={{...card,padding:22}}>
                <h3 style={{fontSize:13,fontWeight:"700",color:MAR,marginBottom:12,fontFamily:AF}}>Gender Split</h3>
                <div style={{display:"flex",justifyContent:"center",marginBottom:14}}>
                  <svg width="120" height="120" viewBox="0 0 120 120">
                    {slices.map((s,i) => s.val > 0 && (
                      <path key={i} d={s.pathD} fill={s.color} opacity="0.9" />
                    ))}
                    <circle cx="60" cy="60" r="30" fill={CARD_BG} />
                    <text x="60" y="64" textAnchor="middle" fontSize="13" fontWeight="700" fill={TX_P} fontFamily={SF}>{stats.total}</text>
                  </svg>
                </div>
                {Object.entries(stats.byGender).map(([g,c]) => (
                  <div key={g} style={{display:"flex",alignItems:"center",gap:8,marginBottom:7}}>
                    <div style={{width:10,height:10,borderRadius:3,background:GEN_C[g],flexShrink:0}} />
                    <span style={{flex:1,fontSize:12,color:TX_S,fontFamily:SF}}>{g}</span>
                    <span style={{fontSize:12,fontWeight:"700",color:GEN_C[g],fontFamily:MF}}>{c}</span>
                    <span style={{fontSize:10,color:TX_M,minWidth:30,textAlign:"right"}}>{stats.total?Math.round(c/stats.total*100):0}%</span>
                  </div>
                ))}
              </div>
              {/* Top kendras */}
              <div style={{...card,padding:22}}>
                <h3 style={{fontSize:13,fontWeight:"700",color:MAR,marginBottom:14,fontFamily:AF}}>🏛 Top Upasana Kendras</h3>
                {stats.byKendra.length===0
                  ? <div style={{textAlign:"center",color:TX_M,paddingTop:30,fontSize:13}}>No data yet</div>
                  : stats.byKendra.slice(0,6).map((k,i) => (
                    <div key={k.name} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10,padding:"7px 10px",borderRadius:8,background:i===0?`${SAF_XL}80`:APP_BG,border:i===0?`1px solid ${SAF}30`:"none"}}>
                      <span style={{fontSize:11,color:i===0?SAF_D:TX_M,fontFamily:MF,width:16,fontWeight:"700"}}>{i+1}</span>
                      <span style={{flex:1,fontSize:12,color:TX_P,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={k.name}>{k.name}</span>
                      <span style={{fontSize:13,fontWeight:"700",color:i===0?SAF_D:MAR,fontFamily:MF}}>{k.count}</span>
                    </div>
                  ))
                }
              </div>
            </div>

            {/* Recent cards */}
            <div style={card}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <h3 style={{fontSize:14,fontWeight:"700",color:MAR,fontFamily:AF}}>Recent Shraddhavans</h3>
                <button style={gBtn} onClick={() => setView("database")}>View All →</button>
              </div>
              {donors.length===0 ? (
                <div style={{padding:48,textAlign:"center",color:TX_M}}>
                  <div style={{fontSize:44,marginBottom:12}}>🩸</div>
                  <div style={{fontSize:14}}>No records yet.</div>
                  <button style={{...pBtn,marginTop:16}} onClick={() => { setForm(INIT_FORM); setEditId(null); setView("add"); }}>+ Register First Shraddhavan</button>
                </div>
              ) : (
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14,padding:18}}>
                  {donors.slice(-6).reverse().map(d => (
                    <div key={d.id} style={{background:APP_BG,border:`1px solid ${d.status==="Accepted"?"rgba(5,150,105,0.2)":"rgba(220,38,38,0.15)"}`,borderRadius:12,padding:16,position:"relative",overflow:"hidden",boxShadow:"0 2px 8px rgba(139,60,20,0.07)"}}>
                      <div style={{position:"absolute",top:0,left:0,right:0,height:2,background:`linear-gradient(90deg,transparent,${d.status==="Accepted"?GRN:RED}50,transparent)`}} />
                      <div style={{position:"absolute",top:12,right:12}}>
                        <span style={{display:"inline-block",padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:"700",background:d.status==="Accepted"?GRN_L:RED_L,color:d.status==="Accepted"?GRN:RED}}>{d.status}</span>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
                        <div style={{width:40,height:40,borderRadius:"50%",background:`${BG_C[d.bloodGroup]||SAF_XL}60`,border:`2px solid ${BG_C[d.bloodGroup]||BDR}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:"700",color:BG_TX[d.bloodGroup]||MAR,flexShrink:0}}>{d.bloodGroup||"?"}</div>
                        <div>
                          <div style={{fontSize:13,fontWeight:"700",color:TX_P,lineHeight:1.2}}>{d.name}</div>
                          <div style={{fontSize:10,color:TX_M,marginTop:2}}>#{d.srNo} · {d.shraddhavanaType}</div>
                        </div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"4px 10px",fontSize:11,color:TX_S}}>
                        <span>📞 {d.mobile||"—"}</span>
                        <span>🎂 {d.age ? `${d.age} yrs` : "—"}</span>
                        <span>{d.gender==="Male"?"👨":"👩"} {d.gender||"—"}</span>
                        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>🏛 {d.upasanaKendra||"—"}</span>
                      </div>
                      {d.rejectionReason && <div style={{marginTop:8,fontSize:10,color:RED,background:RED_L,borderRadius:6,padding:"4px 8px"}}>⚠ {d.rejectionReason}</div>}
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ═══ ADD / EDIT ══════════════════════════════════════════════ */}
        {view==="add" && (
          <div style={{maxWidth:820}}>
            <div style={{display:"flex",gap:8,marginBottom:20}}>
              <button onClick={() => setActiveTab("form")} style={activeTab==="form"?pBtn:g2Btn}>📝 Registration Form</button>
              <button onClick={() => setActiveTab("preview")} style={activeTab==="preview"?pBtn:g2Btn}>👁 Card Preview</button>
            </div>
            {activeTab==="form" ? (
              <div style={card}>
                <div style={{padding:"18px 24px",borderBottom:`1px solid ${BDR}`,display:"flex",alignItems:"center",justifyContent:"space-between",background:`linear-gradient(90deg,${SAF_XL}40,transparent)`}}>
                  <div>
                    <div style={{fontSize:15,fontWeight:"700",color:MAR,fontFamily:AF}}>{editId?"✏️ Edit Shraddhavan Record":"🩸 New Shraddhavan Registration"}</div>
                    <div style={{fontSize:11,color:TX_M,marginTop:3}}>Fields * required · Sr.No & timestamp auto-assigned</div>
                  </div>
                  {!editId && <div style={{background:SAF_XL,border:`1px solid ${SAF}40`,borderRadius:9,padding:"8px 16px",fontSize:12,color:SAF_D,fontFamily:SF,fontWeight:"700"}}>Sr.No: #{nextSrNo}</div>}
                </div>
                <div style={{padding:26}}>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18,marginBottom:18}}>
                    <div><label style={lS}>New / Old Shraddhavan *</label>
                      <select style={iS} value={form.shraddhavanaType} onChange={e=>handleFormChange("shraddhavanaType",e.target.value)}>
                        <option value="">Select Type</option><option value="New">New</option><option value="Old">Old</option>
                      </select></div>
                    <div><label style={lS}>Upasana Kendra</label>
                      <input style={iS} type="text" placeholder="Enter Upasana Kendra" value={form.upasanaKendra} autoComplete="off" onChange={e=>handleFormChange("upasanaKendra",e.target.value)} /></div>
                  </div>
                  <div style={{marginBottom:18}}><label style={lS}>Name of Shraddhavan *</label>
                    <input style={iS} type="text" placeholder="Full name" value={form.name} autoComplete="off" onChange={e=>handleFormChange("name",e.target.value)} /></div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:18,marginBottom:18}}>
                    <div><label style={lS}>Mobile Number * (10 digits)</label>
                      <input style={iS} type="tel" placeholder="e.g. 9876543210" value={form.mobile} maxLength={10} autoComplete="off" onChange={e=>handleFormChange("mobile",e.target.value.replace(/\D/g,''))} /></div>
                    <div><label style={lS}>Date of Birth (DD/MM/YYYY)</label>
                      <DatePicker value={form.dob} onChange={v=>handleFormChange("dob",v)} inputStyle={iS} /></div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:18,marginBottom:18}}>
                    <div><label style={lS}>Age (auto from DOB)</label>
                      <input style={{...iS,background:APP_BG,color:TX_M}} type="number" placeholder="Auto" value={form.age} onChange={e=>handleFormChange("age",e.target.value)} /></div>
                    <div><label style={lS}>Gender *</label>
                      <select style={iS} value={form.gender} onChange={e=>handleFormChange("gender",e.target.value)}>
                        <option value="">Select Gender</option><option value="Male">Male</option><option value="Female">Female</option><option value="Others">Others</option>
                      </select></div>
                    <div><label style={lS}>Blood Group *</label>
                      <select style={iS} value={form.bloodGroup} onChange={e=>handleFormChange("bloodGroup",e.target.value)}>
                        <option value="">Select</option>
                        {BLOOD_GROUPS.map(g => <option key={g} value={g}>{g}</option>)}
                      </select></div>
                  </div>
                  <div style={{marginBottom:18}}><label style={lS}>Status *</label>
                    <div style={{display:"flex",gap:10}}>
                      {["Accepted","Rejected"].map(s => (
                        <button key={s} onClick={() => handleFormChange("status",s)} style={{flex:1,padding:14,borderRadius:10,cursor:"pointer",fontSize:14,fontFamily:SF,fontWeight:"700",transition:"all 0.2s",
                          border:`2px solid ${form.status===s?(s==="Accepted"?GRN:RED):BDR}`,
                          background:form.status===s?(s==="Accepted"?GRN_L:RED_L):APP_BG,
                          color:form.status===s?(s==="Accepted"?GRN:RED):TX_M}}>
                          {s==="Accepted"?"✅":"❌"} {s}
                        </button>
                      ))}
                    </div>
                  </div>
                  {form.status==="Rejected" && (
                    <div style={{marginBottom:18}}><label style={{...lS,color:RED}}>Reason for Rejection *</label>
                      <input style={{...iS,borderColor:`${RED}40`}} type="text" placeholder="Specify reason…" value={form.rejectionReason} autoComplete="off" onChange={e=>handleFormChange("rejectionReason",e.target.value)} /></div>
                  )}
                  <div style={{display:"flex",gap:12,flexWrap:"wrap",marginTop:8}}>
                    <button style={pBtn} onClick={handleSubmit}>{editId?"💾 Update Record":"✅ Register Shraddhavan"}</button>
                    <button style={g2Btn} onClick={() => { setForm(INIT_FORM); showToast("Form cleared","info"); }}>🔄 Clear</button>
                    {editId && <button style={g2Btn} onClick={() => { setEditId(null); setForm(INIT_FORM); setView("database"); }}>← Cancel</button>}
                  </div>
                </div>
              </div>
            ) : (
              <div style={{...card,padding:28}}>
                <h3 style={{fontSize:14,fontWeight:"700",color:MAR,marginBottom:18,fontFamily:AF}}>Shraddhavan Card Preview</h3>
                <div style={{background:APP_BG,border:`1.5px solid ${form.status==="Accepted"?"rgba(5,150,105,0.3)":"rgba(220,38,38,0.2)"}`,borderRadius:14,padding:22,maxWidth:460}}>
                  <div style={{display:"flex",alignItems:"flex-start",gap:14,marginBottom:14}}>
                    <div style={{width:52,height:52,borderRadius:"50%",background:`${BG_C[form.bloodGroup]||SAF_XL}50`,border:`2px solid ${BG_C[form.bloodGroup]||BDR}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:"700",color:BG_TX[form.bloodGroup]||MAR,flexShrink:0}}>{form.bloodGroup||"?"}</div>
                    <div style={{flex:1}}>
                      <div style={{fontSize:17,fontWeight:"700",color:TX_P,marginBottom:3}}>{form.name||"—"}</div>
                      <div style={{fontSize:11,color:TX_M}}>Sr.No: #{editId?form.srNo:nextSrNo}{form.shraddhavanaType&&` · ${form.shraddhavanaType}`}</div>
                    </div>
                    {form.status && <span style={{display:"inline-block",padding:"4px 10px",borderRadius:20,fontSize:11,fontWeight:"700",background:form.status==="Accepted"?GRN_L:RED_L,color:form.status==="Accepted"?GRN:RED}}>{form.status}</span>}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px 16px",fontSize:12,color:TX_S}}>
                    <span>📞 {form.mobile||"—"}</span><span>🎂 {form.dob||"—"}</span>
                    <span>Age: {form.age?`${form.age} yrs`:"—"}</span>
                    <span>{form.gender==="Male"?"👨":form.gender==="Female"?"👩":"🧑"} {form.gender||"—"}</span>
                    <span style={{gridColumn:"1/-1"}}>🏛 {form.upasanaKendra||"—"}</span>
                  </div>
                  {form.rejectionReason && <div style={{marginTop:12,fontSize:11,color:RED,background:RED_L,borderRadius:7,padding:"6px 10px"}}>⚠ {form.rejectionReason}</div>}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ═══ DATABASE ════════════════════════════════════════════════ */}
        {view==="database" && (
          <div>
            <div style={{...card,padding:"14px 18px",marginBottom:14}}>
              <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                <input style={{...iS,flex:1,minWidth:160}} placeholder="🔍 Search name, mobile, kendra…" value={search} onChange={e=>setSearch(e.target.value)} />
                <select style={{...iS,width:120}} value={filterBG} onChange={e=>setFilterBG(e.target.value)}><option value="All">All Groups</option>{BLOOD_GROUPS.map(g=><option key={g} value={g}>{g}</option>)}</select>
                <select style={{...iS,width:120}} value={filterSt} onChange={e=>setFilterSt(e.target.value)}><option value="All">All Status</option><option value="Accepted">Accepted</option><option value="Rejected">Rejected</option></select>
                <select style={{...iS,width:110}} value={filterTy} onChange={e=>setFilterTy(e.target.value)}><option value="All">All Types</option><option value="New">New</option><option value="Old">Old</option></select>
                <button style={sBtn} onClick={exportExcel}>📥 Export Excel</button>
                <label style={{...g2Btn,display:"flex",alignItems:"center",gap:5,cursor:"pointer"}}>📤 Import<input ref={fileRef} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={importExcel} /></label>
              </div>
              <div style={{marginTop:8,fontSize:11,color:TX_M}}>Showing {filtered.length} of {donors.length} records</div>
            </div>
            <div style={{...card,overflowX:"auto"}}>
              {filtered.length===0 ? (
                <div style={{padding:52,textAlign:"center",color:TX_M}}><div style={{fontSize:44,marginBottom:10}}>🔍</div><div>No records found.</div></div>
              ) : (
                <table style={{width:"100%",borderCollapse:"separate",borderSpacing:"0 2px"}}>
                  <thead>
                    <tr style={{background:APP_BG}}>
                      {[["Sr#","srNo"],["Type","shraddhavanaType"],["Name","name"],["Mobile","mobile"],["DOB","dob"],["Age","age"],["Gender","gender"],["Blood","bloodGroup"],["Kendra","upasanaKendra"],["Status","status"],["Registered At","registeredAt"],["Actions",null]].map(([h,f]) => (
                        <th key={h} style={thStyle} onClick={() => f && handleSort(f)}>
                          {h}{f && sortField===f ? (sortDir==="asc"?" ↑":" ↓") : ""}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map(d => (
                      <tr key={d.id}
                        onMouseEnter={e => e.currentTarget.querySelectorAll("td").forEach(td => td.style.background=`${SAF_XL}40`)}
                        onMouseLeave={e => e.currentTarget.querySelectorAll("td").forEach(td => td.style.background=CARD_BG)}>
                        <td style={{...tdStyle,fontFamily:MF,color:TX_M,fontSize:11}}>#{d.srNo}</td>
                        <td style={tdStyle}><span style={{padding:"2px 9px",borderRadius:5,fontSize:10,fontWeight:"700",background:d.shraddhavanaType==="New"?SAF_XL:"rgba(124,58,237,0.1)",color:d.shraddhavanaType==="New"?SAF_D:PURPLE,border:`1px solid ${d.shraddhavanaType==="New"?`${SAF}40`:"rgba(124,58,237,0.3)"}`}}>{d.shraddhavanaType||"—"}</span></td>
                        <td style={tdStyle}><strong style={{color:TX_P}}>{d.name}</strong></td>
                        <td style={{...tdStyle,fontFamily:MF,fontSize:11}}>{d.mobile}</td>
                        <td style={{...tdStyle,fontSize:11}}>{d.dob||"—"}</td>
                        <td style={{...tdStyle,fontFamily:MF}}>{d.age||"—"}</td>
                        <td style={tdStyle}><span style={{color:GEN_C[d.gender]||TX_S,fontWeight:"600"}}>{d.gender||"—"}</span></td>
                        <td style={tdStyle}>{d.bloodGroup ? <span style={{padding:"3px 9px",borderRadius:6,fontSize:11,fontWeight:"700",background:`${BG_C[d.bloodGroup]}50`,color:BG_TX[d.bloodGroup],border:`1px solid ${BG_C[d.bloodGroup]}`}}>{d.bloodGroup}</span> : "—"}</td>
                        <td style={{...tdStyle,fontSize:11,maxWidth:110,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={d.upasanaKendra}>{d.upasanaKendra||"—"}</td>
                        <td style={tdStyle}><span style={{display:"inline-block",padding:"3px 9px",borderRadius:20,fontSize:10,fontWeight:"700",background:d.status==="Accepted"?GRN_L:RED_L,color:d.status==="Accepted"?GRN:RED}}>{d.status}</span></td>
                        <td style={{...tdStyle,fontSize:10,color:TX_M,whiteSpace:"nowrap",fontFamily:MF}}>{d.registeredAt||d.registeredOn||"—"}</td>
                        <td style={tdStyle}>
                          <div style={{display:"flex",gap:5}}>
                            <button style={{...gBtn,padding:"6px 12px",fontSize:11}} onClick={() => handleEdit(d)}>✏️</button>
                            <button style={{...dBtn,padding:"6px 12px",fontSize:11}} onClick={() => setDelConfirm(d.id)}>🗑</button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        )}

        {/* ═══ GALLERY ═════════════════════════════════════════════════ */}
        {view==="gallery" && (
          <div>
            <div style={{display:"flex",gap:12,marginBottom:16,alignItems:"center"}}>
              <button style={pBtn} onClick={() => setShowImgEd(true)}>+ Upload & Edit Image</button>
              <div style={{fontSize:12,color:TX_M}}>{gallery.length} images · Drag to move · ⤡ to resize</div>
            </div>
            <div ref={boardRef} style={{position:"relative",minHeight:520,background:APP_BG2,border:`1px solid ${BDR}`,borderRadius:16,overflow:"hidden",userSelect:"none",backgroundImage:`radial-gradient(circle,${BDR} 1px,transparent 1px)`,backgroundSize:"28px 28px"}}>
              {gallery.length===0 ? (
                <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",color:TX_M}}>
                  <div style={{fontSize:56,marginBottom:14}}>🖼</div>
                  <div style={{fontSize:14}}>Upload images to place them here</div>
                </div>
              ) : gallery.map(img => (
                <div key={img.id} style={{position:"absolute",left:img.x,top:img.y,cursor:interact?.id===img.id&&interact?.type==="drag"?"grabbing":"grab"}} onMouseDown={e=>startInteract(e,img.id,"drag")}>
                  <div style={{position:"relative",background:CARD_BG,border:`1px solid ${BDR_H}`,borderRadius:10,overflow:"hidden",boxShadow:`0 8px 28px rgba(139,60,20,0.15)`}}>
                    <div style={{padding:"5px 10px",display:"flex",alignItems:"center",justifyContent:"space-between",background:SAF_XL,borderBottom:`1px solid ${BDR}`}}>
                      <span style={{fontSize:10,color:SAF_D,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:140,fontFamily:SF,fontWeight:"600"}}>{img.title}</span>
                      <button onMouseDown={e=>e.stopPropagation()} onClick={() => removeGalleryImg(img.id)} style={{background:"none",border:"none",color:RED,cursor:"pointer",fontSize:12,padding:"0 2px"}}>✕</button>
                    </div>
                    {img.src
                      ? <img src={img.src} alt={img.title} draggable={false} style={{display:"block",width:img.displayW,height:img.displayH,filter:img.filter||"none",objectFit:"cover"}} />
                      : <div style={{width:img.displayW||120,height:img.displayH||80,display:"flex",alignItems:"center",justifyContent:"center",color:TX_M,fontSize:11}}>Missing</div>
                    }
                    <div onMouseDown={e=>{e.stopPropagation();startInteract(e,img.id,"resize");}} style={{position:"absolute",bottom:0,right:0,width:18,height:18,cursor:"se-resize",background:SAF,borderTopLeftRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:9,color:"#fff"}}>⤡</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ═══ SETTINGS ════════════════════════════════════════════════ */}
        {view==="settings" && (
          <div style={{maxWidth:660}}>
            {/* Appearance */}
            <div style={{...card,marginBottom:18}}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`,background:`linear-gradient(90deg,${SAF_XL}40,transparent)`}}><div style={{fontSize:14,fontWeight:"700",color:MAR,fontFamily:AF}}>🎨 Appearance & Branding</div></div>
              <div style={{padding:22,display:"flex",flexDirection:"column",gap:12}}>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <button style={g2Btn} onClick={() => setBgEdMode("wallpaper")}>🖼 Edit Wallpaper</button>
                  <button style={g2Btn} onClick={() => setBgEdMode("banner")}>🎨 Edit Banner Image</button>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                  <label style={{...g2Btn,display:"flex",alignItems:"center",gap:7,cursor:"pointer"}}>📌 Set Banner/Sidebar Icon<input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{const u=ev.target.result;setBannerIcon(u);setSidebarIcon(u);sSet(SK.BANNER_ICON,u);sSet(SK.SIDEBAR_ICON,u);showToast("Icon updated!");};r.readAsDataURL(f);}}/></label>
                  <label style={{...g2Btn,display:"flex",alignItems:"center",gap:7,cursor:"pointer"}}>🏷 Sidebar Logo Only<input type="file" accept="image/*" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{setSidebarIcon(ev.target.result);sSet(SK.SIDEBAR_ICON,ev.target.result);showToast("Sidebar logo updated!");};r.readAsDataURL(f);}}/></label>
                </div>
              </div>
            </div>

            {/* ── Sidebar Image ─────────────────────────────────────── */}
            <div style={{...card,marginBottom:18}}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`,background:`linear-gradient(90deg,${SAF_XL}40,transparent)`}}>
                <div style={{fontSize:14,fontWeight:"700",color:MAR,fontFamily:AF}}>🖼 Sidebar Decorative Image</div>
                <div style={{fontSize:11,color:TX_M,marginTop:3,fontFamily:SF}}>Upload, adjust opacity, zoom and crop — changes are saved automatically</div>
              </div>
              <div style={{padding:22}}>
                {/* Live preview */}
                <div style={{marginBottom:18}}>
                  <div style={{fontSize:11,color:TX_M,textTransform:"uppercase",letterSpacing:1,fontWeight:"600",marginBottom:8,fontFamily:SF}}>Live Preview</div>
                  <div style={{position:"relative",height:90,borderRadius:10,overflow:"hidden",background:`linear-gradient(175deg,${SB_BG},${SB_BG2})`,border:`1px solid rgba(255,255,255,0.06)`}}>
                    {sidebarDecorImg ? (
                      <>
                        <img src={sidebarDecorImg} alt="preview" style={{position:"absolute",width:`${sidebarDecorSc}%`,height:`${sidebarDecorSc}%`,minWidth:"100%",minHeight:"100%",objectFit:"cover",objectPosition:`${sidebarDecorX}% ${sidebarDecorY}%`,left:"50%",top:"50%",transform:"translate(-50%,-50%)",mixBlendMode:"luminosity",opacity:sidebarDecorOp,filter:"sepia(50%) saturate(70%) brightness(0.8) hue-rotate(340deg)"}}/>
                        <div style={{position:"absolute",inset:0,pointerEvents:"none",background:`linear-gradient(180deg,${SB_BG} 0%,transparent 30%,transparent 70%,${SB_BG} 100%)`}}/>
                        <div style={{position:"absolute",inset:0,pointerEvents:"none",background:`linear-gradient(90deg,${SB_BG} 0%,transparent 18%,transparent 82%,${SB_BG} 100%)`}}/>
                        <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center"}}><span style={{fontSize:9,color:"rgba(255,255,255,0.3)",fontFamily:SF,letterSpacing:1,textTransform:"uppercase"}}>Sidebar Preview</span></div>
                      </>
                    ) : (
                      <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",color:"rgba(255,255,255,0.2)",fontSize:12,fontFamily:SF}}>No image set — upload one below</div>
                    )}
                  </div>
                </div>

                {/* Upload / Change / Remove */}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:18}}>
                  <label style={{...g2Btn,display:"flex",alignItems:"center",justifyContent:"center",gap:7,cursor:"pointer"}}>
                    {sidebarDecorImg ? "✎ Change Image" : "📁 Upload Image"}
                    <input type="file" accept="image/*" style={{display:"none"}} onChange={e => {
                      const f=e.target.files[0]; if(!f) return;
                      const r=new FileReader();
                      r.onload=ev=>{ setSidebarDecorImg(ev.target.result); sSet(SK.SIDEBAR_DECOR,ev.target.result); showToast("Sidebar image updated!"); };
                      r.readAsDataURL(f);
                    }}/>
                  </label>
                  {sidebarDecorImg && (
                    <button style={dBtn} onClick={() => {
                      setSidebarDecorImg(null); setSidebarDecorOp(0.5); setSidebarDecorX(50); setSidebarDecorY(50); setSidebarDecorSc(100);
                      sDel(SK.SIDEBAR_DECOR); sDel(SK.SIDEBAR_DECOR_OP); sDel(SK.SIDEBAR_DECOR_X); sDel(SK.SIDEBAR_DECOR_Y); sDel(SK.SIDEBAR_DECOR_SC);
                      showToast("Sidebar image removed","info");
                    }}>🗑 Remove Image</button>
                  )}
                </div>

                {/* Adjustment sliders — only when image is loaded */}
                {sidebarDecorImg && (
                  <div>
                    {[
                      { label:"Blend Opacity", val:sidebarDecorOp,  min:0.05, max:0.95, step:0.05, unit:"%", display:v=>Math.round(v*100)+"%", set:v=>{ setSidebarDecorOp(v);  sSet(SK.SIDEBAR_DECOR_OP,String(v));  }, leftTip:"Very subtle",    rightTip:"More visible"   },
                      { label:"Zoom / Scale",  val:sidebarDecorSc,  min:100,  max:300,  step:5,    unit:"%", display:v=>v+"%",                 set:v=>{ setSidebarDecorSc(v);  sSet(SK.SIDEBAR_DECOR_SC,String(v));  }, leftTip:"Fit (100%)",    rightTip:"Zoomed (300%)"  },
                      { label:"Horizontal Position (pan / crop)", val:sidebarDecorX, min:0, max:100, step:1, unit:"%", display:v=>v+"%", set:v=>{ setSidebarDecorX(v); sSet(SK.SIDEBAR_DECOR_X,String(v)); }, leftTip:"← Left", rightTip:"Right →" },
                      { label:"Vertical Position (pan / crop)",   val:sidebarDecorY, min:0, max:100, step:1, unit:"%", display:v=>v+"%", set:v=>{ setSidebarDecorY(v); sSet(SK.SIDEBAR_DECOR_Y,String(v)); }, leftTip:"↑ Top",  rightTip:"Bottom ↓" },
                    ].map(ctrl => (
                      <div key={ctrl.label} style={{marginBottom:16}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
                          <label style={lS}>{ctrl.label}</label>
                          <span style={{fontSize:11,color:SAF,fontWeight:"700",fontFamily:MF}}>{ctrl.display(ctrl.val)}</span>
                        </div>
                        <input type="range" min={ctrl.min} max={ctrl.max} step={ctrl.step} value={ctrl.val}
                          onChange={e => ctrl.set(parseFloat(e.target.value))}
                          style={{width:"100%",accentColor:SAF}}/>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:TX_M,marginTop:2,fontFamily:SF}}>
                          <span>{ctrl.leftTip}</span><span>{ctrl.rightTip}</span>
                        </div>
                      </div>
                    ))}
                    <button style={{...g2Btn,padding:"7px 14px",fontSize:11,marginTop:6}} onClick={() => {
                      setSidebarDecorOp(0.5); setSidebarDecorX(50); setSidebarDecorY(50); setSidebarDecorSc(100);
                      sSet(SK.SIDEBAR_DECOR_OP,"0.5"); sSet(SK.SIDEBAR_DECOR_X,"50"); sSet(SK.SIDEBAR_DECOR_Y,"50"); sSet(SK.SIDEBAR_DECOR_SC,"100");
                      showToast("Image controls reset to defaults","info");
                    }}>↺ Reset Adjustments to Default</button>
                  </div>
                )}
              </div>
            </div>

            {/* Auto-backup */}
            <div style={{...card,marginBottom:18}}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`,background:`linear-gradient(90deg,${GRN_L}60,transparent)`}}><div style={{fontSize:14,fontWeight:"700",color:GRN,fontFamily:AF}}>💾 Auto-Backup & Restore</div></div>
              <div style={{padding:22}}>
                <div style={{background:GRN_L,border:"1px solid rgba(5,150,105,0.2)",borderRadius:10,padding:14,marginBottom:18,fontSize:12,color:"rgb(6 78 59)",lineHeight:1.85}}>
                  🔄 Auto-backup every <strong>{bkInterval} minutes</strong> · 📁 <code style={{fontFamily:MF,fontSize:10}}>MBDC_Backup_YYYY-MM-DD_HH-MM.json</code><br/>
                  {lastBkTime && <>⏱ Last backup: <strong>{lastBkTime}</strong></>}
                </div>
                <div style={{marginBottom:18}}>
                  <label style={{...lS,marginBottom:8}}>Interval: <span style={{color:GRN,fontWeight:"700"}}>{bkInterval} minutes</span></label>
                  <input type="range" min="5" max="60" step="5" value={bkInterval} onChange={e=>{const v=parseInt(e.target.value);setBkInterval(v);sSet(SK.BACKUP_INTERVAL,String(v));}} style={{width:"100%",accentColor:GRN}} />
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:TX_M,marginTop:2}}><span>5 min</span><span>30 min</span><span>60 min</span></div>
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                  <button style={sBtn} onClick={doBackup}>🔄 Backup Now</button>
                  <button style={gdBtn} onClick={exportExcel}>📥 Export Excel</button>
                  <button style={iBtn} onClick={exportJSON}>📦 Export JSON</button>
                  <label style={{...g2Btn,display:"flex",alignItems:"center",gap:6,cursor:"pointer"}}>🔁 Restore JSON<input type="file" accept=".json" style={{display:"none"}} onChange={importJSON}/></label>
                </div>
                {storedBks.length>0 && (
                  <div>
                    <div style={{fontSize:11,color:TX_M,marginBottom:8,textTransform:"uppercase",letterSpacing:1,fontWeight:"600"}}>Stored Backups</div>
                    {storedBks.map((k,i) => {
                      const ts=parseInt(k.replace("bdms-backup-","")); const dt=new Date(ts).toLocaleString("en-IN");
                      return (
                        <div key={k} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 13px",borderRadius:9,background:i===0?GRN_L:APP_BG,border:`1px solid ${i===0?"rgba(5,150,105,0.2)":BDR}`,marginBottom:6}}>
                          <span style={{fontSize:11,color:i===0?"rgb(6 78 59)":TX_M,fontFamily:MF}}>{i===0?"🟢":"⚪"} {dt}</span>
                          <button style={{...g2Btn,padding:"5px 11px",fontSize:11}} onClick={() => dlStoredBk(k)}>↓ Download</button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>
            {/* Data management */}
            <div style={{...card,marginBottom:18}}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`}}><div style={{fontSize:14,fontWeight:"700",color:MAR,fontFamily:AF}}>📊 Data Management</div></div>
              <div style={{padding:22}}>
                <label style={{...g2Btn,display:"flex",alignItems:"center",gap:8,cursor:"pointer",marginBottom:14}}>📤 Import from Excel<input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={importExcel}/></label>
                <div style={{borderTop:`1px solid ${BDR}`,paddingTop:16}}>
                  <div style={{fontSize:12,color:RED,marginBottom:10,fontWeight:"700"}}>⚠️ Danger Zone</div>
                  {clearStep===0 && <button style={dBtn} onClick={() => setClearStep(1)}>🗑 Clear All Records</button>}
                  {clearStep===1 && (
                    <div style={{background:RED_L,border:"1px solid rgba(220,38,38,0.2)",borderRadius:10,padding:16}}>
                      <div style={{fontSize:13,color:MAR,marginBottom:12}}>⚠️ Sure? This will delete <strong>{donors.length}</strong> records permanently.</div>
                      <div style={{display:"flex",gap:10}}>
                        <button style={dBtn} onClick={() => setClearStep(2)}>Yes, delete all</button>
                        <button style={g2Btn} onClick={() => setClearStep(0)}>Cancel</button>
                      </div>
                    </div>
                  )}
                  {clearStep===2 && (
                    <div style={{background:"#FEE2E2",border:`2px solid ${RED}40`,borderRadius:10,padding:16}}>
                      <div style={{fontSize:14,color:MAR,marginBottom:6,fontWeight:"700"}}>🚨 FINAL CONFIRMATION</div>
                      <div style={{fontSize:12,color:RED,marginBottom:14}}>All {donors.length} records will be permanently deleted. CANNOT be undone!</div>
                      <div style={{display:"flex",gap:10}}>
                        <button style={{...dBtn,background:`linear-gradient(135deg,${MAR},${MAR_M})`,color:"#fff",border:"none"}} onClick={() => { saveDonors([]); setClearStep(0); showToast("All records deleted.","info"); }}>🗑 DELETE ALL NOW</button>
                        <button style={g2Btn} onClick={() => setClearStep(0)}>Cancel</button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            </div>
            {/* Storage health */}
            <div style={{...card,marginBottom:18}}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`}}><div style={{fontSize:14,fontWeight:"700",color:MAR,fontFamily:AF}}>🔒 Storage & Persistence</div></div>
              <div style={{padding:22}}>
                <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginBottom:16}}>
                  {[{l:"Donor Records",c:`${donors.length} records`,icon:"👥"},{l:"Gallery",c:`${gallery.length} images`,icon:"🖼"},{l:"Settings",c:"All saved",icon:"⚙️"}].map(item => (
                    <div key={item.l} style={{background:GRN_L,border:"1px solid rgba(5,150,105,0.18)",borderRadius:10,padding:"14px 12px",textAlign:"center"}}>
                      <div style={{fontSize:20,marginBottom:5}}>{item.icon}</div>
                      <div style={{fontSize:11,fontWeight:"700",color:GRN}}>{item.l}</div>
                      <div style={{fontSize:10,color:TX_S,marginTop:3}}>{item.c}</div>
                      <div style={{marginTop:5,fontSize:10,color:GRN}}>✅ Stored</div>
                    </div>
                  ))}
                </div>
                <div style={{background:APP_BG,borderRadius:9,padding:12,fontFamily:MF,fontSize:11,lineHeight:2,border:`1px solid ${BDR}`}}>
                  {restLog.map((l,i) => <div key={i} style={{color:l.startsWith("✅")?GRN:l.startsWith("⚠")?RED:TX_M}}>{l}</div>)}
                </div>
              </div>
            </div>
            {/* App info */}
            <div style={card}>
              <div style={{padding:"16px 22px",borderBottom:`1px solid ${BDR}`}}><div style={{fontSize:14,fontWeight:"700",color:MAR,fontFamily:AF}}>ℹ️ Application Info</div></div>
              <div style={{padding:22,fontSize:12,color:TX_S,lineHeight:2.1}}>
                <div><strong style={{color:MAR}}>Organisation:</strong> ANIRUDDHA'S ACADEMY OF DISASTER MANAGEMENT</div>
                <div><strong style={{color:MAR}}>Upasana Kendra:</strong> Sadguru Shree Aniruddha Upasana Kendra, Wai</div>
                <div><strong style={{color:MAR}}>Camp:</strong> Mega Blood Donation Camp</div>
                <div><strong style={{color:MAR}}>Records:</strong> {donors.length} · Next Sr.No: {nextSrNo}</div>
                <div><strong style={{color:MAR}}>Excel File:</strong> AADM_Shraddhavan_Donors_Data.xlsx</div>
                <div><strong style={{color:MAR}}>Storage:</strong> localStorage (persists across restarts)</div>
                <div style={{marginTop:14,paddingTop:12,borderTop:`1px solid ${BDR}`,fontSize:11,color:TX_M}}>Crafted with ♥ by Swanandsinh Ashok Mahamuni</div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* MODALS */}
      {showImgEd && <ImageEditorModal onClose={() => setShowImgEd(false)} onSave={d => { const p={...d,id:Date.now(),x:30+Math.random()*80,y:30+Math.random()*60}; saveGallery([...gallery,p]); setShowImgEd(false); showToast(`"${d.title}" saved!`); }} />}
      {bgEdMode && <BgEditorModal mode={bgEdMode} currentSrc={bgEdMode==="wallpaper"?bg:bannerImg} currentOpacity={bgEdMode==="wallpaper"?bgOpacity:bannerOp} onClose={() => setBgEdMode(null)} onApply={(img,op) => {
        if (bgEdMode==="wallpaper") { setBg(img); setBgOpacity(op); if(img) sSet(SK.BG,img); else sDel(SK.BG); sSet(SK.OPACITY,String(op)); showToast(img?"Wallpaper applied!":"Wallpaper removed","info"); }
        else { setBannerImg(img); setBannerOp(op); if(img) sSet(SK.BANNER_IMG,img); else sDel(SK.BANNER_IMG); sSet(SK.BANNER_OPACITY,String(op)); showToast(img?"Banner applied!":"Banner removed","info"); }
        setBgEdMode(null);
      }} />}

      {/* Delete confirm */}
      {delConfirm && (
        <div style={{position:"fixed",inset:0,background:"rgba(44,10,4,0.45)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:9999,backdropFilter:"blur(6px)"}}>
          <div style={{...card,padding:28,maxWidth:340,textAlign:"center"}}>
            <div style={{fontSize:38,marginBottom:10}}>⚠️</div>
            <div style={{fontSize:15,fontWeight:"700",color:MAR,marginBottom:7,fontFamily:AF}}>Delete this record?</div>
            <div style={{fontSize:12,color:TX_M,marginBottom:20}}>This action cannot be undone.</div>
            <div style={{display:"flex",gap:10,justifyContent:"center"}}>
              <button style={pBtn} onClick={() => handleDelete(delConfirm)}>Yes, Delete</button>
              <button style={g2Btn} onClick={() => setDelConfirm(null)}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {/* Toast */}
      {toast && (
        <div style={{position:"fixed",bottom:24,right:24,padding:"13px 20px",borderRadius:12,background:toast.type==="error"?MAR:toast.type==="info"?"#1e3a8a":"#065f46",border:"1px solid rgba(255,255,255,0.15)",color:"#fff",fontSize:13,boxShadow:`0 12px 40px rgba(44,10,4,0.3)`,zIndex:9999,maxWidth:340,fontWeight:"600",animation:"slideIn 0.3s ease"}}>
          {toast.type==="error"?"⚠️ ":toast.type==="info"?"ℹ️ ":"✅ "}{toast.msg}
        </div>
      )}

      {/* Backup notification */}
      {bkNotif && (
        <div style={{position:"fixed",bottom:24,left:24,padding:"14px 20px",borderRadius:12,background:GRN,color:"#fff",fontSize:13,zIndex:9999,maxWidth:380,animation:"backupSlide 0.4s ease",boxShadow:`0 0 30px rgba(5,150,105,0.4), 0 12px 40px rgba(0,0,0,0.2)`}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:5}}>
            <span style={{fontSize:18}}>💾</span>
            <strong style={{fontSize:14}}>Auto-Backup Complete</strong>
          </div>
          <div style={{fontSize:11,opacity:0.9,fontFamily:MF}}>{bkNotif.fname}</div>
          <div style={{fontSize:10,opacity:0.7,marginTop:2}}>Saved at {bkNotif.time}</div>
          <button onClick={() => setBkNotif(null)} style={{position:"absolute",top:8,right:8,background:"none",border:"none",color:"rgba(255,255,255,0.6)",cursor:"pointer",fontSize:13}}>✕</button>
        </div>
      )}

      <style>{`
        @keyframes slideIn { from{opacity:0;transform:translateY(10px)} to{opacity:1;transform:translateY(0)} }
        @keyframes backupSlide { from{opacity:0;transform:translateX(-40px)} to{opacity:1;transform:translateX(0)} }
        input:focus, select:focus { border-color:${SAF}!important; box-shadow:0 0 0 3px ${SAF_XL}!important; outline:none; }
        select option { background:#fff; color:${TX_P}; }
        ::-webkit-scrollbar { width:5px; height:5px; }
        ::-webkit-scrollbar-track { background:${APP_BG}; }
        ::-webkit-scrollbar-thumb { background:${SAF_L}; border-radius:3px; }
        input[type=range] { cursor:pointer; }
      `}</style>
    </div>
  );
}
