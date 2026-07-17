import { useState, useEffect, useCallback } from "react";
import { dbGet, dbSet } from "./supabase";
import * as XLSX from "xlsx";
import ShareCard from "./ShareCard";
import ChatTab from "./ChatTab";

// ─── DATOS ───────────────────────────────────────────────────────────────────
const GR = {
  A:["México 🇲🇽","Sudáfrica 🇿🇦","Corea del Sur 🇰🇷","Chequia 🇨🇿"],
  B:["Canadá 🇨🇦","Suiza 🇨🇭","Qatar 🇶🇦","Bosnia 🇧🇦"],
  C:["Brasil 🇧🇷","Marruecos 🇲🇦","Haití 🇭🇹","Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿"],
  D:["EE.UU. 🇺🇸","Paraguay 🇵🇾","Australia 🇦🇺","Turquía 🇹🇷"],
  E:["Alemania 🇩🇪","Curazao 🇨🇼","Costa de Marfil 🇨🇮","Ecuador 🇪🇨"],
  F:["Países Bajos 🇳🇱","Japón 🇯🇵","Túnez 🇹🇳","Suecia 🇸🇪"],
  G:["Bélgica 🇧🇪","Egipto 🇪🇬","Irán 🇮🇷","Nueva Zelanda 🇳🇿"],
  H:["España 🇪🇸","Cabo Verde 🇨🇻","Arabia Saudí 🇸🇦","Uruguay 🇺🇾"],
  I:["Francia 🇫🇷","Senegal 🇸🇳","Noruega 🇳🇴","Iraq 🇮🇶"],
  J:["Argentina 🇦🇷","Argelia 🇩🇿","Austria 🇦🇹","Jordania 🇯🇴"],
  K:["Portugal 🇵🇹","Colombia 🇨🇴","Uzbekistán 🇺🇿","RD Congo 🇨🇩"],
  L:["Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿","Croacia 🇭🇷","Ghana 🇬🇭","Panamá 🇵🇦"],
};
const PM = [
  {id:"a1",l:"México 🇲🇽",v:"Sudáfrica 🇿🇦",g:"A",d:"2026-06-11",j:1},
  {id:"a2",l:"Corea del Sur 🇰🇷",v:"Chequia 🇨🇿",g:"A",d:"2026-06-11",j:1},
  {id:"a3",l:"Chequia 🇨🇿",v:"Sudáfrica 🇿🇦",g:"A",d:"2026-06-18",j:2},
  {id:"a4",l:"México 🇲🇽",v:"Corea del Sur 🇰🇷",g:"A",d:"2026-06-18",j:2},
  {id:"a5",l:"Chequia 🇨🇿",v:"México 🇲🇽",g:"A",d:"2026-06-24",j:3},
  {id:"a6",l:"Sudáfrica 🇿🇦",v:"Corea del Sur 🇰🇷",g:"A",d:"2026-06-24",j:3},
  {id:"b1",l:"Canadá 🇨🇦",v:"Bosnia 🇧🇦",g:"B",d:"2026-06-12",j:1},
  {id:"b2",l:"Qatar 🇶🇦",v:"Suiza 🇨🇭",g:"B",d:"2026-06-13",j:1},
  {id:"b3",l:"Suiza 🇨🇭",v:"Bosnia 🇧🇦",g:"B",d:"2026-06-18",j:2},
  {id:"b4",l:"Canadá 🇨🇦",v:"Qatar 🇶🇦",g:"B",d:"2026-06-18",j:2},
  {id:"b5",l:"Suiza 🇨🇭",v:"Canadá 🇨🇦",g:"B",d:"2026-06-24",j:3},
  {id:"b6",l:"Bosnia 🇧🇦",v:"Qatar 🇶🇦",g:"B",d:"2026-06-24",j:3},
  {id:"c1",l:"Haití 🇭🇹",v:"Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿",g:"C",d:"2026-06-13",j:1},
  {id:"c2",l:"Brasil 🇧🇷",v:"Marruecos 🇲🇦",g:"C",d:"2026-06-13",j:1},
  {id:"c3",l:"Brasil 🇧🇷",v:"Haití 🇭🇹",g:"C",d:"2026-06-19",j:2},
  {id:"c4",l:"Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿",v:"Marruecos 🇲🇦",g:"C",d:"2026-06-19",j:2},
  {id:"c5",l:"Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿",v:"Brasil 🇧🇷",g:"C",d:"2026-06-24",j:3},
  {id:"c6",l:"Marruecos 🇲🇦",v:"Haití 🇭🇹",g:"C",d:"2026-06-24",j:3},
  {id:"d1",l:"EE.UU. 🇺🇸",v:"Paraguay 🇵🇾",g:"D",d:"2026-06-12",j:1},
  {id:"d2",l:"Australia 🇦🇺",v:"Turquía 🇹🇷",g:"D",d:"2026-06-14",j:1},
  {id:"d3",l:"Turquía 🇹🇷",v:"Paraguay 🇵🇾",g:"D",d:"2026-06-19",j:2},
  {id:"d4",l:"EE.UU. 🇺🇸",v:"Australia 🇦🇺",g:"D",d:"2026-06-19",j:2},
  {id:"d5",l:"Turquía 🇹🇷",v:"EE.UU. 🇺🇸",g:"D",d:"2026-06-25",j:3},
  {id:"d6",l:"Paraguay 🇵🇾",v:"Australia 🇦🇺",g:"D",d:"2026-06-25",j:3},
  {id:"e1",l:"Costa de Marfil 🇨🇮",v:"Ecuador 🇪🇨",g:"E",d:"2026-06-14",j:1},
  {id:"e2",l:"Alemania 🇩🇪",v:"Curazao 🇨🇼",g:"E",d:"2026-06-14",j:1},
  {id:"e3",l:"Alemania 🇩🇪",v:"Costa de Marfil 🇨🇮",g:"E",d:"2026-06-20",j:2},
  {id:"e4",l:"Ecuador 🇪🇨",v:"Curazao 🇨🇼",g:"E",d:"2026-06-20",j:2},
  {id:"e5",l:"Curazao 🇨🇼",v:"Costa de Marfil 🇨🇮",g:"E",d:"2026-06-25",j:3},
  {id:"e6",l:"Ecuador 🇪🇨",v:"Alemania 🇩🇪",g:"E",d:"2026-06-25",j:3},
  {id:"f1",l:"Países Bajos 🇳🇱",v:"Japón 🇯🇵",g:"F",d:"2026-06-14",j:1},
  {id:"f2",l:"Suecia 🇸🇪",v:"Túnez 🇹🇳",g:"F",d:"2026-06-14",j:1},
  {id:"f3",l:"Países Bajos 🇳🇱",v:"Suecia 🇸🇪",g:"F",d:"2026-06-20",j:2},
  {id:"f4",l:"Túnez 🇹🇳",v:"Japón 🇯🇵",g:"F",d:"2026-06-21",j:2},
  {id:"f5",l:"Japón 🇯🇵",v:"Suecia 🇸🇪",g:"F",d:"2026-06-25",j:3},
  {id:"f6",l:"Túnez 🇹🇳",v:"Países Bajos 🇳🇱",g:"F",d:"2026-06-25",j:3},
  {id:"g1",l:"Irán 🇮🇷",v:"Nueva Zelanda 🇳🇿",g:"G",d:"2026-06-15",j:1},
  {id:"g2",l:"Bélgica 🇧🇪",v:"Egipto 🇪🇬",g:"G",d:"2026-06-15",j:1},
  {id:"g3",l:"Bélgica 🇧🇪",v:"Irán 🇮🇷",g:"G",d:"2026-06-21",j:2},
  {id:"g4",l:"Egipto 🇪🇬",v:"Nueva Zelanda 🇳🇿",g:"G",d:"2026-06-21",j:2},
  {id:"g5",l:"Irán 🇮🇷",v:"Egipto 🇪🇬",g:"G",d:"2026-06-26",j:3},
  {id:"g6",l:"Nueva Zelanda 🇳🇿",v:"Bélgica 🇧🇪",g:"G",d:"2026-06-26",j:3},
  {id:"h1",l:"Arabia Saudí 🇸🇦",v:"Uruguay 🇺🇾",g:"H",d:"2026-06-15",j:1},
  {id:"h2",l:"España 🇪🇸",v:"Cabo Verde 🇨🇻",g:"H",d:"2026-06-15",j:1},
  {id:"h3",l:"Uruguay 🇺🇾",v:"Cabo Verde 🇨🇻",g:"H",d:"2026-06-21",j:2},
  {id:"h4",l:"España 🇪🇸",v:"Arabia Saudí 🇸🇦",g:"H",d:"2026-06-21",j:2},
  {id:"h5",l:"España 🇪🇸",v:"Uruguay 🇺🇾",g:"H",d:"2026-06-26",j:3},
  {id:"h6",l:"Cabo Verde 🇨🇻",v:"Arabia Saudí 🇸🇦",g:"H",d:"2026-06-26",j:3},
  {id:"i1",l:"Francia 🇫🇷",v:"Senegal 🇸🇳",g:"I",d:"2026-06-16",j:1},
  {id:"i2",l:"Iraq 🇮🇶",v:"Noruega 🇳🇴",g:"I",d:"2026-06-16",j:1},
  {id:"i3",l:"Noruega 🇳🇴",v:"Senegal 🇸🇳",g:"I",d:"2026-06-22",j:2},
  {id:"i4",l:"Francia 🇫🇷",v:"Iraq 🇮🇶",g:"I",d:"2026-06-22",j:2},
  {id:"i5",l:"Francia 🇫🇷",v:"Noruega 🇳🇴",g:"I",d:"2026-06-26",j:3},
  {id:"i6",l:"Senegal 🇸🇳",v:"Iraq 🇮🇶",g:"I",d:"2026-06-26",j:3},
  {id:"j1",l:"Argentina 🇦🇷",v:"Argelia 🇩🇿",g:"J",d:"2026-06-16",j:1},
  {id:"j2",l:"Austria 🇦🇹",v:"Jordania 🇯🇴",g:"J",d:"2026-06-17",j:1},
  {id:"j3",l:"Argentina 🇦🇷",v:"Austria 🇦🇹",g:"J",d:"2026-06-22",j:2},
  {id:"j4",l:"Argelia 🇩🇿",v:"Jordania 🇯🇴",g:"J",d:"2026-06-22",j:2},
  {id:"j5",l:"Argentina 🇦🇷",v:"Jordania 🇯🇴",g:"J",d:"2026-06-26",j:3},
  {id:"j6",l:"Argelia 🇩🇿",v:"Austria 🇦🇹",g:"J",d:"2026-06-26",j:3},
  {id:"k1",l:"Portugal 🇵🇹",v:"RD Congo 🇨🇩",g:"K",d:"2026-06-17",j:1},
  {id:"k2",l:"Uzbekistán 🇺🇿",v:"Colombia 🇨🇴",g:"K",d:"2026-06-17",j:1},
  {id:"k3",l:"Portugal 🇵🇹",v:"Uzbekistán 🇺🇿",g:"K",d:"2026-06-23",j:2},
  {id:"k4",l:"Colombia 🇨🇴",v:"RD Congo 🇨🇩",g:"K",d:"2026-06-23",j:2},
  {id:"k5",l:"Colombia 🇨🇴",v:"Portugal 🇵🇹",g:"K",d:"2026-06-27",j:3},
  {id:"k6",l:"RD Congo 🇨🇩",v:"Uzbekistán 🇺🇿",g:"K",d:"2026-06-27",j:3},
  {id:"l1",l:"Ghana 🇬🇭",v:"Panamá 🇵🇦",g:"L",d:"2026-06-17",j:1},
  {id:"l2",l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"Croacia 🇭🇷",g:"L",d:"2026-06-17",j:1},
  {id:"l3",l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"Ghana 🇬🇭",g:"L",d:"2026-06-23",j:2},
  {id:"l4",l:"Croacia 🇭🇷",v:"Panamá 🇵🇦",g:"L",d:"2026-06-23",j:2},
  {id:"l5",l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"Panamá 🇵🇦",g:"L",d:"2026-06-27",j:3},
  {id:"l6",l:"Croacia 🇭🇷",v:"Ghana 🇬🇭",g:"L",d:"2026-06-27",j:3},
];

const MULT={grupos:1,dieciseisavos:1,octavos:1.5,cuartos:2,semifinales:3,tercerpuesto:3,final:5};
const PHL={grupos:"Grupos",dieciseisavos:"1/16",octavos:"Octavos",cuartos:"Cuartos",semifinales:"Semis",tercerpuesto:"3er Puesto",final:"Final"};
const CUOTA=20;
const PIN="4805";
const PRIMER_PARTIDO=new Date("2026-06-11T21:00:00+02:00");

// ─── SCORING ─────────────────────────────────────────────────────────────────
function mPts(pred,res,ph,isFet){
  if(!pred||!res)return 0;
  const m=MULT[ph]||1,rl=+res.l,rv=+res.v,pl=+pred.l,pv=+pred.v;
  if([rl,rv,pl,pv].some(isNaN))return 0;
  let b=0;
  const exacto=pl===rl&&pv===rv;
  if(exacto)b+=5;
  else{const rw=rl>rv?"L":rv>rl?"V":"E",pw=pl>pv?"L":pv>pl?"V":"E";if(rw===pw)b+=2;}
  if(pl===rl)b+=1;if(pv===rv)b+=1;
  let pts=Math.round(b*m*10)/10;
  if(isFet){if(exacto)pts=Math.round(pts*5*10)/10;else if(b>=2)pts=Math.round(pts*3*10)/10;}
  return pts;
}
function prePts(p,r){
  let x=0;if(!r||!p)return x;
  if(p.campeon&&p.campeon===r.campeon)x+=30;
  if(p.subcampeon&&p.subcampeon===r.subcampeon)x+=20;
  if(p.tercero&&p.tercero===r.tercero)x+=15;
  if(p.goleador&&p.goleador===r.goleador)x+=20;
  if(p.mvp&&p.mvp===r.mvp)x+=15;
  const rs=(r.semis||[]).filter(Boolean);
  (p.semis||[]).filter(Boolean).forEach(t=>{if(rs.includes(t))x+=10;});
  return x;
}
function spcPts(p,r){
  let x=0;if(!r||!p)return x;
  if(p.expulsado&&p.expulsado===r.expulsado)x+=10;
  if(p.hattrick&&p.hattrick===r.hattrick)x+=10;
  if(p.revelacion&&p.revelacion===r.revelacion)x+=15;
  if(p.goleada&&p.goleada===r.goleada)x+=10;
  return x;
}
function totPts(part,allM,rPre,rSpc){
  let t=0;
  t+=prePts(part.pre||{},rPre);
  t+=spcPts(part.spc||{},rSpc);
  const fet=part.fetiche||null;
  const penaltis=part.penaltis||{};
  (allM||[]).forEach(m=>{
    t+=mPts((part.mp||{})[m.id],m.result,m.ph||"grupos",m.id===fet);
    // Bono penaltis: +3 si acertó que iba a penaltis y el admin lo confirmó
    if(penaltis[m.id]&&m.fueAPenaltis)t+=3;
  });
  return Math.round(t*10)/10;
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────
function fd(s){if(!s)return"";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;}
function isPast(d){return d?new Date()>new Date(d+"T23:59:59"):false;}

function TeamSel({val,onChange,placeholder,disabled}){
  return(
    <select style={{...S.sel,opacity:disabled?.6:1}} value={val||""} onChange={e=>onChange(e.target.value)} disabled={disabled}>
      <option value="">{placeholder||"— Seleccionar —"}</option>
      {Object.entries(GR).map(([g,ts])=>(
        <optgroup key={g} label={"Grupo "+g}>
          {ts.map(t=><option key={t} value={t}>{t}</option>)}
        </optgroup>
      ))}
    </select>
  );
}

function Chip({label,done,total}){
  const pct=total>0?Math.round(done/total*100):0;
  const col=pct===100?"#4ade80":pct>0?"#fbbf24":"#2a2a2a";
  return(
    <div style={{background:"rgba(0,0,0,0.3)",borderRadius:8,padding:"6px 8px",border:"1px solid "+col+"44"}}>
      <div style={{fontSize:10,color:"#4a5568",marginBottom:2}}>{label}</div>
      <div style={{fontSize:12,fontWeight:700,color:col}}>{done}/{total}</div>
      <div style={{height:3,background:"#111",borderRadius:2,marginTop:3}}>
        <div style={{height:"100%",width:pct+"%",background:col,borderRadius:2}}/>
      </div>
    </div>
  );
}

function Pos({n}){
  const bg=n===1?"linear-gradient(135deg,#e8b923,#f59e0b)":n===2?"linear-gradient(135deg,#e2e8f0,#94a3b8)":n===3?"linear-gradient(135deg,#cd7f32,#92400e)":"rgba(255,255,255,0.08)";
  const col=n<=2?"#000":"#fff";
  return <div style={{width:30,height:30,borderRadius:"50%",background:bg,color:col,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:13,flexShrink:0}}>{n}</div>;
}

function useCountdown(target){
  const[t,setT]=useState(Math.max(0,target-Date.now()));
  useEffect(()=>{if(t<=0)return;const i=setInterval(()=>setT(Math.max(0,target-Date.now())),1000);return()=>clearInterval(i);},[target,t]);
  const s=Math.floor(t/1000);
  return{days:Math.floor(s/86400),hrs:Math.floor((s%86400)/3600),mins:Math.floor((s%3600)/60),secs:s%60,done:t<=0};
}

function CountdownWidget(){
  const{days,hrs,mins,secs,done}=useCountdown(PRIMER_PARTIDO.getTime());
  const pad=n=>String(n).padStart(2,"0");
  if(done)return(
    <div style={{background:"linear-gradient(135deg,#052e16,#065f46)",border:"1px solid #10b98144",borderRadius:14,padding:"14px 16px",marginBottom:14,textAlign:"center"}}>
      <div style={{fontSize:28,marginBottom:4}}>⚽</div>
      <div style={{color:"#10b981",fontWeight:800,fontSize:15,letterSpacing:1}}>¡EL MUNDIAL HA COMENZADO!</div>
    </div>
  );
  return(
    <div style={{background:"linear-gradient(135deg,#0a0f1e,#0d1a0d)",border:"1px solid rgba(232,185,35,0.2)",borderRadius:14,padding:"14px 16px",marginBottom:14}}>
      <div style={{textAlign:"center",fontSize:11,color:"#4a5568",marginBottom:10,letterSpacing:2,textTransform:"uppercase",fontWeight:600}}>⚽ México vs Sudáfrica · 11 Jun · 21:00h</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6}}>
        {[[days,"DÍAS"],[hrs,"HORAS"],[mins,"MIN"],[pad(secs),"SEG"]].map(([v,l])=>(
          <div key={l} style={{background:"rgba(0,0,0,0.5)",borderRadius:10,padding:"10px 4px",textAlign:"center",border:"1px solid rgba(232,185,35,0.15)"}}>
            <div style={{fontWeight:900,fontSize:26,background:"linear-gradient(180deg,#fff5c0,#e8b923)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",lineHeight:1}}>{v}</div>
            <div style={{fontSize:9,color:"#4a5568",marginTop:3,letterSpacing:2,fontWeight:600}}>{l}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function MiPosicion({scores,myName,bote}){
  const idx=scores.findIndex(p=>p.name===myName);
  if(idx===-1)return null;
  const p=scores[idx],pos=idx+1,next=scores[idx-1];
  const diff=next?Math.round((next.tot-p.tot)*10)/10:null;
  const posCol=pos===1?"#e8b923":pos<=3?"#10b981":"#94a3b8";
  const premio=pos<=3?Math.round(bote*[.6,.25,.15][pos-1]):0;
  return(
    <div style={{background:"linear-gradient(135deg,rgba(232,185,35,0.1),rgba(200,16,46,0.08))",border:"1px solid rgba(232,185,35,0.25)",borderRadius:14,padding:"12px 14px",marginBottom:14}}>
      <div style={{fontSize:10,color:"#4a5568",letterSpacing:2,fontWeight:700,textTransform:"uppercase",marginBottom:8}}>📍 Tu posición</div>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{width:44,height:44,borderRadius:"50%",background:pos===1?"linear-gradient(135deg,#e8b923,#f59e0b)":pos<=3?"linear-gradient(135deg,#10b981,#059669)":"rgba(255,255,255,0.08)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:20,color:pos<=2?"#000":"#fff",flexShrink:0}}>
          {pos}
        </div>
        <div style={{flex:1}}>
          <div style={{fontWeight:800,fontSize:17,color:posCol}}>{p.tot} <span style={{fontSize:12,color:"#4a5568",fontWeight:400}}>pts</span></div>
          {diff!==null&&diff>0?<div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>A <span style={{color:"#ef4444",fontWeight:700}}>{diff} pts</span> de {next.name.split(" ")[0]}</div>
          :<div style={{fontSize:12,color:"#10b981",marginTop:2}}>🏆 ¡Líder de la porra!</div>}
        </div>
        {premio>0&&<div style={{textAlign:"center",background:"rgba(16,185,129,0.1)",border:"1px solid rgba(16,185,129,0.2)",borderRadius:10,padding:"6px 10px"}}>
          <div style={{fontSize:11,color:"#4a5568"}}>Premio est.</div>
          <div style={{fontSize:18,fontWeight:900,color:"#10b981"}}>+{premio}€</div>
        </div>}
      </div>
    </div>
  );
}

// ─── EXPORT EXCEL ────────────────────────────────────────────────────────────
function exportXLSX(parts,allM,rPre,rSpc){
  const wb=XLSX.utils.book_new();
  const sorted=[...parts].sort((a,b)=>totPts(b,allM,rPre,rSpc)-totPts(a,allM,rPre,rSpc));
  const bote=parts.length*20;
  const PHL_F={grupos:"Fase de Grupos",octavos:"Octavos",cuartos:"Cuartos",semifinales:"Semifinales",final:"Final"};
  const rankRows=[["#","Jugador","Puntos Totales","Pre-Mundial","Partidos","Retos","Premio €","Fetiche"]];
  sorted.forEach((p,i)=>{
    const fetM=allM.find(m=>m.id===p.fetiche);
    const pPre=prePts(p.pre||{},rPre),pSpc=spcPts(p.spc||{},rSpc);
    let pM=0;allM.forEach(m=>{pM+=mPts((p.mp||{})[m.id],m.result,m.ph||"grupos",m.id===p.fetiche);});
    const premio=i===0?Math.round(bote*.6):i===1?Math.round(bote*.25):i===2?Math.round(bote*.15):0;
    rankRows.push([i+1,p.name,totPts(p,allM,rPre,rSpc),pPre,Math.round(pM*10)/10,pSpc,premio||"-",fetM?`${fetM.l} vs ${fetM.v}`:"-"]);
  });
  const wsR=XLSX.utils.aoa_to_sheet(rankRows);
  wsR["!cols"]=[{wch:4},{wch:18},{wch:14},{wch:14},{wch:12},{wch:10},{wch:10},{wch:28}];
  XLSX.utils.book_append_sheet(wb,wsR,"🏆 Ranking");
  sorted.forEach(p=>{
    const rows=[],fetiche=p.fetiche||null;
    const tot=totPts(p,allM,rPre,rSpc);
    rows.push([`PRONÓSTICOS DE ${p.name.toUpperCase()}`]);
    rows.push([`Total: ${tot} pts`,`Registrado: ${p.registeredAt?new Date(p.registeredAt).toLocaleDateString("es-ES"):"-"}`]);
    rows.push([]);
    rows.push(["── PRE-MUNDIAL ──────────────────────"]);
    rows.push(["Campo","Tu pronóstico","Resultado real","Puntos"]);
    [["Campeón del mundo (30pts)",p.pre?.campeon,rPre?.campeon,30],["Subcampeón (20pts)",p.pre?.subcampeon,rPre?.subcampeon,20],["3er puesto (15pts)",p.pre?.tercero,rPre?.tercero,15],["Máx. goleador (20pts)",p.pre?.goleador,rPre?.goleador,20],["MVP (15pts)",p.pre?.mvp,rPre?.mvp,15]].forEach(([campo,pred,real,max])=>{
      rows.push([campo,pred||"-",real||"(pend.)",rPre&&pred&&real&&pred===real?max:rPre?0:"-"]);
    });
    (p.pre?.semis||[]).forEach((s,i)=>{
      const hit=s&&(rPre?.semis||[]).includes(s);
      rows.push([`Semifinalista ${i+1} (10pts)`,s||"-","-",rPre?hit?10:0:"-"]);
    });
    rows.push([]);
    rows.push(["── PARTIDO FETICHE ──────────────────"]);
    rows.push(["Partido","Tu pronóstico","Resultado real","Puntos","Mult."]);
    const fetM=allM.find(m=>m.id===fetiche);
    if(fetM){
      const pred=(p.mp||{})[fetM.id],res=fetM.result;
      rows.push([`${fetM.l} vs ${fetM.v}`,pred?`${pred.l}-${pred.v}`:"-",res?`${res.l}-${res.v}`:"(pend.)",res&&pred?mPts(pred,res,fetM.ph||"grupos",true):"-","×5/×3"]);
    }else{rows.push(["No seleccionado","-","-","-","-"]);}
    rows.push([]);
    rows.push(["── RETOS ESPECIALES ─────────────────"]);
    rows.push(["Reto","Tu pronóstico","Resultado real","Puntos"]);
    [["Primer expulsado (10pts)",p.spc?.expulsado,rSpc?.expulsado,10],["Primer hat-trick (10pts)",p.spc?.hattrick,rSpc?.hattrick,10],["Equipo revelación (15pts)",p.spc?.revelacion,rSpc?.revelacion,15],["Mayor goleada (10pts)",p.spc?.goleada,rSpc?.goleada,10]].forEach(([reto,pred,real,max])=>{
      rows.push([reto,pred||"-",real||"(pend.)",rSpc&&pred&&real&&pred===real?max:rSpc?0:"-"]);
    });
    rows.push([]);
    rows.push(["── PARTIDOS ─────────────────────────"]);
    rows.push(["Fecha","Fase","Partido","Tu pronóstico","Resultado real","Puntos","Fetiche"]);
    allM.forEach(m=>{
      const pred=(p.mp||{})[m.id],res=m.result,isFet=m.id===fetiche;
      rows.push([m.d||m.date||"-",PHL_F[m.ph]||m.ph,`${m.l} vs ${m.v}`,pred?`${pred.l}-${pred.v}`:"-",res?`${res.l}-${res.v}`:"(pend.)",res&&pred?mPts(pred,res,m.ph||"grupos",isFet):"-",isFet?"⭐":""]);
    });
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:12},{wch:18},{wch:28},{wch:14},{wch:14},{wch:8},{wch:8}];
    XLSX.utils.book_append_sheet(wb,ws,p.name.replace(/[\\/*?:[\]]/g,"").substring(0,28));
  });
  XLSX.writeFile(wb,"porra-mundial-2026.xlsx");
}

// ─── CLASIFICACION ───────────────────────────────────────────────────────────
function calcClasif(teams,matches){
  const t={};
  teams.forEach(x=>{t[x]={pj:0,g:0,e:0,p:0,gf:0,gc:0,pts:0};});
  matches.filter(m=>m.result).forEach(m=>{
    const rl=+m.result.l,rv=+m.result.v,tl=t[m.l],tv=t[m.v];
    if(!tl||!tv||isNaN(rl)||isNaN(rv))return;
    tl.pj++;tv.pj++;tl.gf+=rl;tl.gc+=rv;tv.gf+=rv;tv.gc+=rl;
    if(rl>rv){tl.g++;tl.pts+=3;tv.p++;}
    else if(rv>rl){tv.g++;tv.pts+=3;tl.p++;}
    else{tl.e++;tv.e++;tl.pts++;tv.pts++;}
  });
  return Object.entries(t).map(([n,s])=>({n,...s,dif:s.gf-s.gc})).sort((a,b)=>b.pts-a.pts||b.dif-a.dif||b.gf-a.gf);
}

function TablaGrupo({grupo,allM}){
  const teams=GR[grupo]||[],mG=allM.filter(m=>m.g===grupo);
  if(!mG.some(m=>m.result))return(
    <div style={{background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"10px 12px",textAlign:"center",fontSize:12,color:"#2d3748",marginBottom:14}}>Sin resultados aún en el Grupo {grupo}</div>
  );
  const tabla=calcClasif(teams,mG);
  return(
    <div style={{marginBottom:16}}>
      <div style={{fontSize:11,color:"#e8b923",fontWeight:700,letterSpacing:1,marginBottom:8,textTransform:"uppercase"}}>📊 Clasificación Grupo {grupo}</div>
      <div style={{background:"rgba(0,0,0,0.4)",borderRadius:12,overflow:"hidden",border:"1px solid rgba(255,255,255,0.07)"}}>
        <div style={{display:"grid",gridTemplateColumns:"20px 1fr 26px 26px 26px 26px 26px 26px 34px 34px",padding:"7px 10px",background:"rgba(232,185,35,0.08)",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
          {["#","Equipo","PJ","G","E","P","GF","GC","DIF","PTS"].map(c=><div key={c} style={{fontSize:10,fontWeight:700,color:"#4a5568",textAlign:c==="Equipo"?"left":"center"}}>{c}</div>)}
        </div>
        {tabla.map((r,i)=>{
          const isTop=i<2,difCol=r.dif>0?"#10b981":r.dif<0?"#ef4444":"#4a5568";
          const posCol=i===0?"#e8b923":i===1?"#10b981":"#2d3748";
          return(
            <div key={r.n} style={{display:"grid",gridTemplateColumns:"20px 1fr 26px 26px 26px 26px 26px 26px 34px 34px",padding:"8px 10px",background:i===0?"rgba(232,185,35,0.06)":i===1?"rgba(16,185,129,0.04)":"transparent",borderBottom:i<tabla.length-1?"1px solid rgba(255,255,255,0.04)":"none",alignItems:"center"}}>
              <div style={{fontSize:11,fontWeight:700,color:posCol,textAlign:"center"}}>{i+1}</div>
              <div style={{fontSize:12,fontWeight:isTop?700:400,color:isTop?"#e2e8f0":"#94a3b8",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",paddingRight:4}}>
                {r.n.split(" ").slice(0,-1).join(" ")||r.n}{isTop&&<span style={{marginLeft:4,fontSize:9,color:posCol}}>●</span>}
              </div>
              {[r.pj,r.g,r.e,r.p,r.gf,r.gc].map((v,j)=><div key={j} style={{fontSize:11,textAlign:"center",color:"#4a5568"}}>{v}</div>)}
              <div style={{fontSize:11,textAlign:"center",color:difCol,fontWeight:600}}>{r.dif>0?"+"+r.dif:r.dif}</div>
              <div style={{fontSize:13,textAlign:"center",fontWeight:900,color:isTop?"#e8b923":"#e2e8f0"}}>{r.pts}</div>
            </div>
          );
        })}
      </div>
      <div style={{display:"flex",gap:12,marginTop:5,paddingLeft:4}}>
        <div style={{fontSize:10,color:"#e8b923"}}>● 1º clasif.</div>
        <div style={{fontSize:10,color:"#10b981"}}>● 2º clasif.</div>
      </div>
    </div>
  );
}


// ─── ANUNCIO 11 JUNIO ─────────────────────────────────────────────────────────
function MundialAnnouncement({onClose}){
  const past=[1998,2002,2006,2010,2014,2018,2022,2026];
  const future=[2030,2034,2038,2042,2046,2050,2054,2058,2062,2066,2070,2074,2078];
  return(
    <div style={{position:"fixed",inset:0,zIndex:9999,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",padding:"16px",backdropFilter:"blur(4px)"}}>
      <div style={{background:"#f5f0e8",borderRadius:20,padding:"32px 24px",maxWidth:380,width:"100%",textAlign:"center",position:"relative",boxShadow:"0 20px 60px rgba(0,0,0,0.5)",maxHeight:"90vh",overflowY:"auto"}}>
        {/* Trophy */}
        <div style={{fontSize:40,marginBottom:8}}>🏆</div>
        <div style={{width:40,height:2,background:"#c8a84b",margin:"0 auto 24px"}}/>

        {/* Llevas X mundiales */}
        <div style={{fontFamily:"'Georgia',serif",fontSize:20,fontWeight:900,color:"#1a1a1a",textTransform:"uppercase",letterSpacing:2,lineHeight:1.2,marginBottom:6}}>LLEVAS</div>
        <div style={{fontSize:64,fontWeight:900,color:"#1a5c2e",lineHeight:1,fontFamily:"'Georgia',serif",marginBottom:4}}>8</div>
        <div style={{fontSize:22,fontWeight:900,color:"#1a5c2e",textTransform:"uppercase",letterSpacing:2,lineHeight:1.2,marginBottom:4}}>MUNDIALES</div>
        <div style={{fontSize:18,fontWeight:900,color:"#1a1a1a",textTransform:"uppercase",letterSpacing:2,marginBottom:20}}>A TUS ESPALDAS</div>

        {/* Camisetas */}
        <div style={{display:"flex",justifyContent:"center",gap:6,flexWrap:"wrap",marginBottom:6}}>
          {past.map((y,i)=>(
            <div key={y} style={{textAlign:"center"}}>
              <div style={{fontSize:i===0?22:18,filter:i===0?"none":"grayscale(0)"}}>{i===0?"⭐":"👕"}</div>
              <div style={{fontSize:9,color:"#555",marginTop:2}}>{y}</div>
            </div>
          ))}
        </div>

        {/* Divider con balón */}
        <div style={{display:"flex",alignItems:"center",gap:8,margin:"20px 0"}}>
          <div style={{flex:1,height:1,background:"#c8a84b"}}/>
          <div style={{fontSize:20}}>⚽</div>
          <div style={{flex:1,height:1,background:"#c8a84b"}}/>
        </div>

        {/* Te quedan */}
        <div style={{fontFamily:"'Georgia',serif",fontSize:16,fontWeight:700,color:"#1a1a1a",textTransform:"uppercase",letterSpacing:2,marginBottom:4}}>TE QUEDAN</div>
        <div style={{fontSize:72,fontWeight:900,color:"#1a5c2e",lineHeight:1,fontFamily:"'Georgia',serif",marginBottom:4}}>13</div>
        <div style={{fontSize:18,fontWeight:900,color:"#1a5c2e",textTransform:"uppercase",letterSpacing:2,marginBottom:16}}>POR DELANTE</div>

        {/* Años futuros */}
        <div style={{fontSize:12,color:"#555",lineHeight:2,marginBottom:20}}>
          {[future.slice(0,5),future.slice(5,10),future.slice(10)].map((row,i)=>(
            <div key={i}>{row.map((y,j)=><span key={y}>{y}{j<row.length-1?" · ":""}</span>)}</div>
          ))}
        </div>

        {/* Divider estrella */}
        <div style={{display:"flex",alignItems:"center",gap:8,margin:"16px 0"}}>
          <div style={{flex:1,height:1,background:"#c8a84b"}}/>
          <div style={{fontSize:14,color:"#c8a84b"}}>★</div>
          <div style={{flex:1,height:1,background:"#c8a84b"}}/>
        </div>

        {/* Mensaje final */}
        <div style={{fontSize:15,fontWeight:900,color:"#1a1a1a",textTransform:"uppercase",letterSpacing:1,lineHeight:1.6,marginBottom:4}}>DISFRUTA EL CAMINO.</div>
        <div style={{fontSize:15,fontWeight:900,color:"#1a5c2e",textTransform:"uppercase",letterSpacing:1,marginBottom:20}}>LO MEJOR AÚN ESTÁ POR VENIR.</div>
        <div style={{fontSize:18,color:"#555",marginBottom:24}}>🌐</div>

        {/* Botón cerrar arriba izq */}
        <button onClick={onClose} style={{position:"absolute",top:16,left:16,background:"#1a5c2e",border:"none",borderRadius:20,padding:"8px 16px",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:6,boxShadow:"0 2px 8px rgba(0,0,0,0.2)"}}>
          ← Ir a la app
        </button>
      </div>
    </div>
  );
}

// ─── APP ─────────────────────────────────────────────────────────────────────

// ─── ESTADÍSTICAS DE JORNADA ──────────────────────────────────────────────────
function jornadaCompleta(j, allM){
  const matchesJ = allM.filter(m => m.j === j && m.ph === "grupos");
  if (matchesJ.length === 0) return false;
  return matchesJ.every(m => m.result);
}

function calcStatsJornada(j, allM, parts){
  const matchesJ = allM.filter(m => m.j === j && m.ph === "grupos" && m.result);

  const ptsJornada = parts.map(p => {
    let pts = 0;
    matchesJ.forEach(m => { pts += mPts((p.mp||{})[m.id], m.result, "grupos", m.id === p.fetiche); });
    return { name: p.name, pts: Math.round(pts*10)/10, fetiche: p.fetiche, fetM: matchesJ.find(m=>m.id===p.fetiche) };
  }).sort((a,b) => b.pts - a.pts);

  let masAcertado = null, masFallado = null, mayorSorpresa = null, mayorGoleada = null;
  let maxAciertos = -1, minAciertos = Infinity, maxGoles = -1;

  matchesJ.forEach(m => {
    let aciertosExactos = 0, aciertosGanador = 0, totalConPred = 0;
    parts.forEach(p => {
      const pred = (p.mp||{})[m.id];
      if (!pred || pred.l === "" || pred.v === "") return;
      totalConPred++;
      const pl = +pred.l, pv = +pred.v, rl = +m.result.l, rv = +m.result.v;
      if (pl === rl && pv === rv) aciertosExactos++;
      const rw = rl > rv ? "L" : rv > rl ? "V" : "E";
      const pw = pl > pv ? "L" : pv > pl ? "V" : "E";
      if (rw === pw) aciertosGanador++;
    });
    if (totalConPred === 0) return;
    if (aciertosExactos > maxAciertos) { maxAciertos = aciertosExactos; masAcertado = { ...m, aciertos: aciertosExactos, total: totalConPred }; }
    if (aciertosGanador < minAciertos) { minAciertos = aciertosGanador; masFallado = { ...m, aciertos: aciertosGanador, total: totalConPred }; }
    const goles = (+m.result.l) + (+m.result.v);
    if (goles > maxGoles) { maxGoles = goles; mayorGoleada = m; }
    const pctGanador = aciertosGanador / totalConPred;
    if (!mayorSorpresa || pctGanador < mayorSorpresa.pct) {
      mayorSorpresa = { ...m, pct: pctGanador, aciertos: aciertosGanador, total: totalConPred };
    }
  });

  return { ptsJornada, masAcertado, masFallado, mayorGoleada, mayorSorpresa, matchesJ };
}

const PULLAS = [
  "menudo arranque... a remontar desde ya 😅",
  "última posición de la jornada, pero el Mundial es largo 🫠",
  "ni de coña con esa jornada... a recuperar 🔥",
  "esto pinta a temporada larga para ti 😬",
];
const PIROPOS = [
  "arrasando desde el primer minuto 🔥",
  "menudo arranque, vas con todo 🚀",
  "el resto ya puede empezar a preocuparse 😏",
  "liderato desde la J1, qué nivel 👑",
];

function StatsJornadaView({jornada, allM, parts, onClose}){
  const stats = calcStatsJornada(jornada, allM, parts);
  if (!stats.ptsJornada.length) return null;
  const lider = stats.ptsJornada[0];
  const ultimo = stats.ptsJornada[stats.ptsJornada.length - 1];
  const piropo = PIROPOS[jornada % PIROPOS.length];
  const pulla = PULLAS[jornada % PULLAS.length];

  return(
    <div style={{position:"fixed",inset:0,zIndex:9998,background:"rgba(0,0,0,0.85)",display:"flex",flexDirection:"column",alignItems:"center",overflowY:"auto",padding:"16px"}}>
      <div style={{width:"100%",maxWidth:480,background:"linear-gradient(160deg,#0a1628,#0d1f0d 60%,#1a0a00)",borderRadius:18,overflow:"hidden",border:"1px solid rgba(232,185,35,0.2)"}}>
        <button onClick={onClose} style={{position:"sticky",top:0,zIndex:10,background:"rgba(0,0,0,0.6)",border:"none",borderRadius:0,padding:"10px 16px",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",width:"100%",textAlign:"left",display:"flex",alignItems:"center",gap:8}}>
          ← Volver a la app
        </button>

        <div style={{padding:"20px 18px"}}>
          <div style={{textAlign:"center",marginBottom:20}}>
            <div style={{fontSize:36}}>📊</div>
            <div style={{fontSize:22,fontWeight:900,background:"linear-gradient(135deg,#e8b923,#fff5c0)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>RESUMEN JORNADA {jornada}</div>
            <div style={{fontSize:12,color:"#4a5568",marginTop:4}}>Fase de grupos · {stats.matchesJ.length} partidos jugados</div>
          </div>

          <div style={{background:"linear-gradient(135deg,rgba(232,185,35,0.15),rgba(232,185,35,0.05))",border:"1px solid rgba(232,185,35,0.3)",borderRadius:14,padding:"16px",marginBottom:12}}>
            <div style={{fontSize:10,color:"#e8b923",letterSpacing:2,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>🥇 Rey de la jornada</div>
            <div style={{fontWeight:900,fontSize:20,color:"#fff"}}>{lider.name}</div>
            <div style={{fontSize:13,color:"#fff5c0",marginTop:2}}>{lider.pts} pts en esta jornada · {piropo}</div>
          </div>

          {ultimo.name !== lider.name && (
            <div style={{background:"linear-gradient(135deg,rgba(220,38,38,0.15),rgba(220,38,38,0.05))",border:"1px dashed rgba(220,38,38,0.3)",borderRadius:14,padding:"16px",marginBottom:16}}>
              <div style={{fontSize:10,color:"#fca5a5",letterSpacing:2,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>😅 Farolillo de la jornada</div>
              <div style={{fontWeight:900,fontSize:18,color:"#fff"}}>{ultimo.name}</div>
              <div style={{fontSize:13,color:"#fecaca",marginTop:2}}>{ultimo.pts} pts en esta jornada · {pulla}</div>
            </div>
          )}

          {stats.masAcertado && (
            <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #10b981",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
              <div style={{fontSize:10,color:"#10b981",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>🎯 Partido más acertado</div>
              <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{stats.masAcertado.l} {stats.masAcertado.result.l}-{stats.masAcertado.result.v} {stats.masAcertado.v}</div>
              <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>{stats.masAcertado.aciertos}/{stats.masAcertado.total} acertaron el resultado exacto</div>
            </div>
          )}

          {stats.mayorSorpresa && stats.mayorSorpresa.pct < 0.4 && (
            <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #f472b6",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
              <div style={{fontSize:10,color:"#f472b6",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>😱 Sorpresa de la jornada</div>
              <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{stats.mayorSorpresa.l} {stats.mayorSorpresa.result.l}-{stats.mayorSorpresa.result.v} {stats.mayorSorpresa.v}</div>
              <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>Solo {stats.mayorSorpresa.aciertos}/{stats.mayorSorpresa.total} acertaron quién ganaba</div>
            </div>
          )}

          {stats.mayorGoleada && (
            <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #fb923c",borderRadius:10,padding:"12px 14px",marginBottom:16}}>
              <div style={{fontSize:10,color:"#fb923c",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>🔥 Partido con más goles</div>
              <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{stats.mayorGoleada.l} {stats.mayorGoleada.result.l}-{stats.mayorGoleada.result.v} {stats.mayorGoleada.v}</div>
            </div>
          )}

          <div style={{fontSize:11,color:"#4a5568",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:8}}>📋 Puntos de todos en la J{jornada}</div>
          {stats.ptsJornada.map((p, i) => (
            <div key={p.name} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:"rgba(0,0,0,0.25)",borderRadius:8,marginBottom:4}}>
              <div style={{width:22,fontSize:11,color:"#4a5568",fontWeight:700}}>{i+1}</div>
              <div style={{flex:1,fontSize:13,color:"#e2e8f0"}}>{p.name}</div>
              <div style={{fontWeight:800,fontSize:14,color:i===0?"#e8b923":"#10b981"}}>{p.pts} pts</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── PESTAÑA FIJA DE JORNADAS ─────────────────────────────────────────────────
function JornadasTab({jornadasCompletas, allM, parts}){
  const[sel,setSel]=useState(jornadasCompletas[jornadasCompletas.length-1]);
  useEffect(()=>{
    if(!jornadasCompletas.includes(sel)&&jornadasCompletas.length>0){
      setSel(jornadasCompletas[jornadasCompletas.length-1]);
    }
  },[jornadasCompletas,sel]);

  if(jornadasCompletas.length===0)return(
    <div style={S.content}>
      <div style={S.empty}>
        <div style={{fontSize:36}}>📈</div>
        <div style={{marginTop:8}}>Aún no hay jornadas completas</div>
        <div style={{fontSize:12,color:"#1a1a1a",marginTop:4}}>Aparecerán aquí en cuanto termine la J1</div>
      </div>
    </div>
  );

  return(
    <div>
      <div style={{display:"flex",gap:6,padding:"12px 16px",background:"rgba(0,0,0,0.3)",borderBottom:"1px solid rgba(255,255,255,0.05)"}}>
        {jornadasCompletas.map(j=>(
          <button key={j} style={{...S.chip,...(sel===j?S.chipA:{}),flex:1,textAlign:"center"}} onClick={()=>setSel(j)}>Jornada {j}</button>
        ))}
      </div>
      <div style={{padding:"16px"}}>
        <StatsJornadaInline jornada={sel} allM={allM} parts={parts}/>
      </div>
    </div>
  );
}

// Versión "inline" de las estadísticas (sin overlay, para vivir dentro de la pestaña)
function StatsJornadaInline({jornada, allM, parts}){
  const stats=calcStatsJornada(jornada,allM,parts);
  if(!stats.ptsJornada.length)return <div style={S.empty}>Sin datos para esta jornada</div>;
  const lider=stats.ptsJornada[0];
  const ultimo=stats.ptsJornada[stats.ptsJornada.length-1];
  const piropo=PIROPOS[jornada%PIROPOS.length];
  const pulla=PULLAS[jornada%PULLAS.length];
  return(
    <div>
      <div style={{background:"linear-gradient(135deg,rgba(232,185,35,0.15),rgba(232,185,35,0.05))",border:"1px solid rgba(232,185,35,0.3)",borderRadius:14,padding:"16px",marginBottom:12}}>
        <div style={{fontSize:10,color:"#e8b923",letterSpacing:2,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>🥇 Rey de la jornada</div>
        <div style={{fontWeight:900,fontSize:20,color:"#fff"}}>{lider.name}</div>
        <div style={{fontSize:13,color:"#fff5c0",marginTop:2}}>{lider.pts} pts en esta jornada · {piropo}</div>
      </div>

      {ultimo.name!==lider.name&&(
        <div style={{background:"linear-gradient(135deg,rgba(220,38,38,0.15),rgba(220,38,38,0.05))",border:"1px dashed rgba(220,38,38,0.3)",borderRadius:14,padding:"16px",marginBottom:16}}>
          <div style={{fontSize:10,color:"#fca5a5",letterSpacing:2,fontWeight:700,textTransform:"uppercase",marginBottom:6}}>😅 Farolillo de la jornada</div>
          <div style={{fontWeight:900,fontSize:18,color:"#fff"}}>{ultimo.name}</div>
          <div style={{fontSize:13,color:"#fecaca",marginTop:2}}>{ultimo.pts} pts en esta jornada · {pulla}</div>
        </div>
      )}

      {stats.masAcertado&&(
        <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #10b981",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{fontSize:10,color:"#10b981",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>🎯 Partido más acertado</div>
          <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{stats.masAcertado.l} {stats.masAcertado.result.l}-{stats.masAcertado.result.v} {stats.masAcertado.v}</div>
          <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>{stats.masAcertado.aciertos}/{stats.masAcertado.total} acertaron el resultado exacto</div>
        </div>
      )}

      {stats.mayorSorpresa&&stats.mayorSorpresa.pct<0.4&&(
        <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #f472b6",borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{fontSize:10,color:"#f472b6",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>😱 Sorpresa de la jornada</div>
          <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{stats.mayorSorpresa.l} {stats.mayorSorpresa.result.l}-{stats.mayorSorpresa.result.v} {stats.mayorSorpresa.v}</div>
          <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>Solo {stats.mayorSorpresa.aciertos}/{stats.mayorSorpresa.total} acertaron quién ganaba</div>
        </div>
      )}

      {stats.mayorGoleada&&(
        <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #fb923c",borderRadius:10,padding:"12px 14px",marginBottom:16}}>
          <div style={{fontSize:10,color:"#fb923c",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:4}}>🔥 Partido con más goles</div>
          <div style={{fontSize:14,fontWeight:700,color:"#fff"}}>{stats.mayorGoleada.l} {stats.mayorGoleada.result.l}-{stats.mayorGoleada.result.v} {stats.mayorGoleada.v}</div>
        </div>
      )}

      <div style={{fontSize:11,color:"#4a5568",fontWeight:700,letterSpacing:1,textTransform:"uppercase",marginBottom:8}}>📋 Puntos de todos en la J{jornada}</div>
      {stats.ptsJornada.map((p,i)=>(
        <div key={p.name} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",background:"rgba(0,0,0,0.25)",borderRadius:8,marginBottom:4}}>
          <div style={{width:22,fontSize:11,color:"#4a5568",fontWeight:700}}>{i+1}</div>
          <div style={{flex:1,fontSize:13,color:"#e2e8f0"}}>{p.name}</div>
          <div style={{fontWeight:800,fontSize:14,color:i===0?"#e8b923":"#10b981"}}>{p.pts} pts</div>
        </div>
      ))}
    </div>
  );
}



// ─── ELIMINATORIA TAB ─────────────────────────────────────────────────────────
function EliminatoriaTab({db, upDb, allM}){
  const[selName,setSelName]=useState("");
  const[view,setView]=useState("landing"); // landing | editing
  const[saved,setSaved]=useState(false);
  const save=()=>{setSaved(true);setTimeout(()=>setSaved(false),2000);};
  const parts=db.parts||[];
  const elimM=allM.filter(m=>m.ph!=="grupos");
  const p=parts.find(x=>x.name===selName);
  const fases=["dieciseisavos","octavos","cuartos","semifinales","tercerpuesto","final"];

  const upM=async(id,side,v)=>{
    if(!p)return;
    const updated=parts.map(x=>x.id!==p.id?x:{...x,mp:{...(x.mp||{}),[id]:{...((x.mp||{})[id]||{}),[side]:v}}});
    await upDb("parts",updated);
  };

  const setPenalti=async(id,ph)=>{
    if(!p)return;
    const current=p.penaltis||{};
    const newP={};
    Object.entries(current).forEach(([k,v])=>{if(v!==ph)newP[k]=v;});
    if(current[id]){await upDb("parts",parts.map(x=>x.id!==p.id?x:{...x,penaltis:newP}));return;}
    newP[id]=ph;
    await upDb("parts",parts.map(x=>x.id!==p.id?x:{...x,penaltis:newP}));
  };

  if(elimM.length===0)return(
    <div style={{padding:"40px 16px",textAlign:"center"}}>
      <div style={{fontSize:40,marginBottom:12}}>🏟</div>
      <div style={{color:"#e2e8f0",fontWeight:700,fontSize:16}}>Aún no hay partidos de eliminatoria</div>
      <div style={{color:"#4a5568",fontSize:13,marginTop:6}}>El admin los añade desde Panel → ➕ Partidos</div>
    </div>
  );

  if(view==="landing")return(
    <div style={{padding:"24px 16px"}}>
      <div style={{textAlign:"center",marginBottom:24}}>
        <div style={{fontSize:44}}>🏟</div>
        <div style={{fontSize:20,fontWeight:900,color:"#e8b923",marginTop:8}}>ELIMINATORIAS</div>
        <div style={{fontSize:12,color:"#4a5568",marginTop:4}}>{elimM.length} partidos · Pronostica y marca penaltis</div>
      </div>
      <div style={{background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:14,padding:"20px",marginBottom:16}}>
        <div style={{color:"#94a3b8",fontSize:13,fontWeight:600,marginBottom:10}}>¿Quién eres?</div>
        <select value={selName} onChange={e=>setSelName(e.target.value)}
          style={{width:"100%",padding:"10px 12px",background:"rgba(5,10,25,0.9)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,color:"#fff",fontSize:14,marginBottom:14,boxSizing:"border-box"}}>
          <option value="">— Selecciona tu nombre —</option>
          {parts.map(x=><option key={x.id} value={x.name}>{x.name}</option>)}
        </select>
        <button onClick={()=>selName&&setView("editing")}
          style={{width:"100%",padding:"13px",background:selName?"linear-gradient(135deg,#c8102e,#a00d25)":"rgba(255,255,255,0.1)",border:"none",borderRadius:12,color:"#fff",fontWeight:700,fontSize:14,cursor:selName?"pointer":"not-allowed",opacity:selName?1:0.5}}>
          ✏️ Meter mis pronósticos →
        </button>
      </div>
      <div style={{background:"rgba(99,102,241,0.1)",border:"1px solid rgba(99,102,241,0.25)",borderRadius:10,padding:"10px 14px",color:"#818cf8",fontSize:12}}>
        🥅 Marca un partido por ronda que crees que va a penaltis → +3 pts si aciertas
      </div>
    </div>
  );

  // Editing view
  const porFase=fases.map(ph=>({ph,ms:elimM.filter(m=>m.ph===ph)})).filter(x=>x.ms.length>0);
  return(
    <div style={{paddingBottom:80}}>
      <div style={{padding:"10px 16px",background:"rgba(0,0,0,0.5)",borderBottom:"1px solid rgba(255,255,255,0.06)",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:11,color:"#4a5568"}}>Pronósticos de</div>
          <div style={{fontWeight:800,fontSize:16,color:"#e8b923"}}>{selName}</div>
        </div>
        <div style={{display:"flex",gap:8}}>
          <button onClick={save} style={{background:saved?"linear-gradient(135deg,#065f46,#047857)":"linear-gradient(135deg,#c8102e,#a00d25)",border:"none",borderRadius:10,padding:"8px 16px",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>
            {saved?"✅ Guardado":"💾 Guardar"}
          </button>
          <button onClick={()=>setView("landing")} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:8,padding:"7px 14px",color:"#a0aec0",cursor:"pointer",fontSize:13}}>
            ← Volver
          </button>
        </div>
      </div>
      <div style={{padding:"14px 16px"}}>
        <div style={{background:"rgba(99,102,241,0.1)",border:"1px solid rgba(99,102,241,0.25)",borderRadius:10,padding:"9px 12px",color:"#818cf8",fontSize:12,marginBottom:16}}>
          🥅 Un partido por ronda que crees que va a penaltis → +3 pts si aciertas
        </div>
        {porFase.map(({ph,ms})=>{
          const penRonda=p&&Object.entries(p.penaltis||{}).find(([k,v])=>v===ph&&ms.find(m=>m.id===k));
          return(
            <div key={ph} style={{marginBottom:20}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                <span style={{fontSize:12,color:"#60a5fa",fontWeight:700,textTransform:"uppercase",letterSpacing:1}}>{PHL[ph]}</span>
                <span style={{fontSize:10,color:"#1e3a5f"}}>×{MULT[ph]}</span>
                {penRonda&&<span style={{fontSize:10,background:"rgba(99,102,241,0.2)",color:"#818cf8",border:"1px solid rgba(99,102,241,0.3)",borderRadius:8,padding:"1px 7px"}}>🥅 marcado</span>}
              </div>
              {ms.map(m=>{
                const pred=(p&&(p.mp||{})[m.id])||{};
                const res=m.result;
                const isPen=p&&!!(p.penaltis||{})[m.id];
                const pts=res&&pred.l!==""&&pred.v!==""?mPts(pred,res,m.ph,false):null;
                return(
                  <div key={m.id} style={{padding:"12px",background:"rgba(255,255,255,0.04)",border:"1px solid "+(isPen?"rgba(99,102,241,0.5)":res?"rgba(16,185,129,0.2)":"rgba(255,255,255,0.06)"),borderRadius:10,marginBottom:8}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                      <span style={{fontSize:11,color:"#4a5568"}}>{fd(m.date||"")}</span>
                      {!res&&<button onClick={()=>setPenalti(m.id,m.ph)}
                        style={{background:isPen?"rgba(99,102,241,0.3)":"rgba(0,0,0,0.3)",border:"1px solid "+(isPen?"#818cf8":"rgba(255,255,255,0.1)"),borderRadius:8,padding:"4px 10px",color:isPen?"#818cf8":"#4a5568",fontSize:12,cursor:"pointer",fontWeight:isPen?700:400}}>
                        {isPen?"🥅 Penaltis ✓":"🥅 ¿Penaltis?"}
                      </button>}
                      {res&&m.fueAPenaltis&&<span style={{fontSize:11,color:isPen?"#818cf8":"#4a5568"}}>🥅{isPen?" +3pts ✓":" (no apostaste)"}</span>}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 44px 14px 44px 1fr",alignItems:"center",gap:4}}>
                      <div style={{textAlign:"right",color:"#e2e8f0",fontSize:13,fontWeight:600}}>{m.l}</div>
                      <input type="number" min={0} max={20} value={pred.l||""} onChange={e=>upM(m.id,"l",e.target.value)} disabled={!!res}
                        style={{background:"rgba(5,10,25,0.9)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 2px",color:"#e8b923",fontSize:16,fontWeight:900,width:44,textAlign:"center",outline:"none",opacity:res?.5:1}}
                        placeholder="?"/>
                      <span style={{textAlign:"center",color:"rgba(255,255,255,0.1)"}}>–</span>
                      <input type="number" min={0} max={20} value={pred.v||""} onChange={e=>upM(m.id,"v",e.target.value)} disabled={!!res}
                        style={{background:"rgba(5,10,25,0.9)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 2px",color:"#e8b923",fontSize:16,fontWeight:900,width:44,textAlign:"center",outline:"none",opacity:res?.5:1}}
                        placeholder="?"/>
                      <div style={{color:"#e2e8f0",fontSize:13,fontWeight:600}}>{m.v}</div>
                    </div>
                    {res&&<div style={{textAlign:"center",fontSize:11,marginTop:6}}>
                      <span style={{color:"#10b981"}}>Real: {res.l}–{res.v}</span>
                      {pts!==null&&<span style={{color:"#10b981",marginLeft:8,fontWeight:700}}>+{pts}pts</span>}
                    </div>}
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
      {/* Botón flotante fijo abajo */}
      <div style={{position:"fixed",bottom:0,left:0,right:0,padding:"10px 16px 20px",background:"rgba(8,8,16,0.95)",borderTop:"1px solid rgba(255,255,255,0.08)",backdropFilter:"blur(10px)",zIndex:100}}>
        <button onClick={save} style={{width:"100%",padding:"13px",background:saved?"linear-gradient(135deg,#065f46,#047857)":"linear-gradient(135deg,#c8102e,#a00d25)",border:"none",borderRadius:12,color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer",maxWidth:568,display:"block",margin:"0 auto"}}>
          {saved?"✅ ¡Guardado!":"💾 Confirmar y guardar pronósticos"}
        </button>
      </div>
    </div>
  );
}


// ─── PRONOS TAB ───────────────────────────────────────────────────────────────
function PronosTab({db,scores,closed,allM,onRegister,onPredict,onShare}){
  return(
    <div>
      <div style={{...S.regCard,...(closed?{opacity:.6}:{})}}>
        <div><div style={{fontWeight:700,color:"#fff"}}>¿Aún no estás apuntado?</div><div style={{fontSize:12,color:"rgba(255,255,255,0.4)",marginTop:2}}>{CUOTA}€ · {db.deadline&&!isPast(db.deadline)?"Hasta "+fd(db.deadline):"Sin límite"}</div></div>
        <button style={{...S.btnGold,...(closed?{opacity:.4,cursor:"not-allowed"}:{})}} onClick={!closed?onRegister:undefined}>{closed?"Cerrado":"Unirse →"}</button>
      </div>
      {(db.parts||[]).length===0?<div style={S.empty}><div style={{fontSize:36}}>🌍</div><div style={{marginTop:8,color:"#2d3748"}}>¡Sé el primero en apuntarte!</div></div>:
        (db.parts||[]).map(p=>{
          const sc=scores.find(x=>x.id===p.id);
          const mp=p.mp||{},mD=PM.filter(m=>mp[m.id]&&mp[m.id].l!==""&&mp[m.id].v!=="").length;
          const pD=[p.pre?.campeon,p.pre?.subcampeon,p.pre?.tercero,p.pre?.goleador,p.pre?.mvp].filter(Boolean).length;
          const sD=[p.spc?.expulsado,p.spc?.hattrick,p.spc?.revelacion,p.spc?.goleada].filter(Boolean).length;
          const fetM=allM.find(m=>m.id===p.fetiche);
          return(
            <div key={p.id} style={S.card}>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:10}}>
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                    <span style={{fontWeight:700,fontSize:15,color:"#fff"}}>{p.name}</span>
                    {p.locked&&<span style={{fontSize:10,background:"rgba(13,26,13,0.8)",color:"#4ade80",border:"1px solid rgba(16,185,129,0.2)",borderRadius:10,padding:"1px 7px"}}>🔒</span>}
                  </div>
                  <div style={{fontWeight:800,fontSize:18,color:"#10b981",marginTop:2}}>{sc?.tot||0}<span style={{fontSize:11,color:"#2d3748",fontWeight:400}}> pts</span></div>
                  {fetM&&<div style={{fontSize:11,color:"#e8b923",marginTop:2}}>⭐ {fetM.l.split(" ")[0]} vs {fetM.v.split(" ")[0]}</div>}
                </div>
                <div style={{display:"flex",gap:6,flexDirection:"column",alignItems:"flex-end"}}>
                  {!p.locked&&!closed&&<button style={S.btnGold} onClick={()=>onPredict(p.id)}>✏️ Editar</button>}
                  {p.locked&&<button style={{...S.btnGhost,fontSize:11,padding:"4px 8px"}} onClick={()=>onPredict(p.id)}>👁 Ver pronósticos</button>}
                  {p.locked&&<button style={{...S.btnGhost,fontSize:11,padding:"4px 8px"}} onClick={()=>onShare(p)}>📤 Compartir</button>}
                </div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6}}>
                <Chip label="Pre-Mundial" done={pD} total={5}/>
                <Chip label="Partidos" done={mD} total={72}/>
                <Chip label="Retos" done={sD} total={4}/>
              </div>
            </div>
          );
        })
      }
    </div>
  );
}

// ─── REGISTER VIEW ────────────────────────────────────────────────────────────
function RegisterView({db,upDb,closed,saveDraft,loadDraft,deleteDraft,onDone,onBack}){
  const[step,setStep]=useState("form");
  const[name,setName]=useState("");
  const[err,setErr]=useState("");
  const[saving,setSaving]=useState(false);
  const[hasDraft,setHasDraft]=useState(false);
  const[draftLoading,setDraftLoading]=useState(false);
  const[draft,setDraft]=useState({pre:{campeon:"",subcampeon:"",tercero:"",goleador:"",mvp:"",semis:["","","",""]},spc:{expulsado:"",hattrick:"",revelacion:"",goleada:""},mp:{},fetiche:null,penaltis:{}});
  const[sec,setSec]=useState("pre");
  const[gr,setGr]=useState("A");
  useEffect(()=>{if(name.trim()&&step==="predict"){const t=setTimeout(()=>saveDraft(name,draft),1500);return()=>clearTimeout(t);}},[draft,name,step,saveDraft]);
  const checkDraft=async n=>{if(!n.trim())return;setDraftLoading(true);const d=await loadDraft(n);setHasDraft(!!d&&d._draftName);setDraftLoading(false);};
  const resumeDraft=async()=>{setDraftLoading(true);const d=await loadDraft(name);if(d){const{_draftName,_savedAt,...rest}=d;setDraft(rest);setStep("predict");}setDraftLoading(false);};
  const next=()=>{const n=name.trim();if(!n){setErr("Escribe tu nombre");return;}if((db.parts||[]).find(p=>p.name.toLowerCase()===n.toLowerCase())){setErr("Ese nombre ya existe");return;}setErr("");setStep("predict");};
  const confirm=async()=>{setSaving(true);const np={id:Date.now().toString(),name:name.trim(),pre:draft.pre,spc:draft.spc,mp:draft.mp,fetiche:draft.fetiche,locked:true,registeredAt:new Date().toISOString()};await upDb("parts",[...(db.parts||[]),np]);await deleteDraft(name);setSaving(false);onDone(name.trim());};
  const upPre=(k,v)=>setDraft(d=>({...d,pre:{...d.pre,[k]:v}}));
  const upSemi=(i,v)=>setDraft(d=>{const s=[...(d.pre.semis||["","","",""])];s[i]=v;return{...d,pre:{...d.pre,semis:s}};});
  const upSpc=(k,v)=>setDraft(d=>({...d,spc:{...d.spc,[k]:v}}));
  const upM=(id,side,v)=>setDraft(d=>({...d,mp:{...d.mp,[id]:{...(d.mp[id]||{}),[side]:v}}}));
  const setFetiche=id=>setDraft(d=>({...d,fetiche:d.fetiche===id?null:id}));
  const setPenalti=(id,ph)=>setDraft(d=>{
    const current=d.penaltis||{};
    // Only one per ronda — remove others in same phase first
    const filtered=Object.fromEntries(Object.entries(current).filter(([k])=>{
      const m=PM.find(x=>x.id===k)||(d.mp||{})[k];
      return true; // extraM partidos: allow
    }));
    // Remove all others in same ph
    const newP={};
    Object.entries(filtered).forEach(([k,v])=>{
      // We don't have easy access to ph here, so we store ph in the value
      if(v!==ph)newP[k]=v;
    });
    if(current[id])return{...d,penaltis:newP}; // toggle off
    newP[id]=ph; // mark with ph so we can filter by ronda
    return{...d,penaltis:newP};
  });
  if(closed)return(<div style={S.app}><div style={S.hdr}><div style={S.hdrRow}><button style={S.backBtn} onClick={onBack}>◀</button><div style={{color:"#e8b923",fontWeight:800}}>Registro</div><div/></div></div><div style={{...S.content,textAlign:"center",paddingTop:60}}><div style={{fontSize:48}}>🔒</div><div style={{color:"#ff6b35",fontWeight:700,fontSize:18,marginTop:12}}>Pronósticos cerrados</div></div></div>);
  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={step==="form"?onBack:()=>setStep("form")}>◀</button>
          <div style={{textAlign:"center"}}><div style={{fontWeight:800,fontSize:16,color:"#e8b923"}}>{step==="form"?"🎟 Nueva inscripción":"🎯 Tus pronósticos"}</div>{step==="predict"&&<div style={{fontSize:11,color:"#4a5568"}}>{name}</div>}</div>
          {step==="predict"?<div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:2}}><button style={{...S.btnGold,...(saving?{opacity:.5}:{})}} onClick={!saving?confirm:undefined}>{saving?"…":"✅ Confirmar"}</button><div style={{fontSize:9,color:"#4a5568"}}>💾 guardado auto</div></div>:<div/>}
        </div>
      </div>
      {step==="form"&&(
        <div style={S.content}>
          <div style={{background:"linear-gradient(135deg,#0a2a0a,#062006)",border:"1px solid rgba(74,222,128,0.27)",borderRadius:14,padding:"18px",marginBottom:20,textAlign:"center"}}>
            <div style={{fontSize:40,fontWeight:900,color:"#4ade80"}}>{CUOTA}€</div>
            <div style={{color:"#86efac",fontSize:14,marginTop:4}}>Inscripción a la porra</div>
            <div style={{color:"#2d3748",fontSize:12,marginTop:6}}>Una vez confirmado, tus pronósticos quedan bloqueados definitivamente</div>
          </div>
          <div style={{marginBottom:14}}>
            <div style={S.fl}>Tu nombre</div>
            <input style={{...S.inp,...(err?{borderColor:"#ef4444"}:{})}} placeholder="¿Cómo te llamas?" value={name} onChange={e=>{setName(e.target.value);setErr("");setHasDraft(false);}} onBlur={e=>checkDraft(e.target.value)} onKeyDown={e=>e.key==="Enter"&&next()}/>
            {err&&<div style={{color:"#ef4444",fontSize:12,marginTop:4}}>⚠ {err}</div>}
          </div>
          {hasDraft&&(
            <div style={{background:"rgba(26,20,0,0.8)",border:"1px solid rgba(232,185,35,0.3)",borderRadius:10,padding:"12px",marginBottom:12}}>
              <div style={{color:"#e8b923",fontWeight:700,fontSize:13,marginBottom:6}}>💾 Tienes un borrador guardado</div>
              <div style={{color:"#4a5568",fontSize:12,marginBottom:10}}>Puedes continuar donde lo dejaste o empezar de cero</div>
              <div style={{display:"flex",gap:8}}><button style={{...S.btnGold,flex:1}} onClick={resumeDraft}>▶ Continuar</button><button style={{...S.btnAct,flex:1}} onClick={()=>{setHasDraft(false);next();}}>🗑 De cero</button></div>
            </div>
          )}
          {!hasDraft&&<button style={{...S.btnGold,width:"100%",padding:"14px",fontSize:15,borderRadius:12}} onClick={next}>{draftLoading?"Buscando borrador…":"Continuar →"}</button>}
        </div>
      )}
      {step==="predict"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,padding:"10px 16px",background:"rgba(0,0,0,0.3)",borderBottom:"1px solid rgba(255,255,255,0.05)"}}>
            <Chip label="Pre-Mundial" done={[draft.pre.campeon,draft.pre.subcampeon,draft.pre.tercero,draft.pre.goleador,draft.pre.mvp].filter(Boolean).length} total={5}/>
            <Chip label="Partidos" done={PM.filter(m=>draft.mp[m.id]&&draft.mp[m.id].l!==""&&draft.mp[m.id].v!=="").length} total={72}/>
            <Chip label="Fetiche" done={draft.fetiche?1:0} total={1}/>
          </div>
          <div style={S.secNav}>
            {[["pre","🏆 Pre"],["grupos","⚽ Grupos"],["elim","🏟 Elim."],["retos","⭐ Retos"]].map(([id,l])=>(
              <button key={id} style={{...S.snBtn,...(sec===id?S.snA:{})}} onClick={()=>setSec(id)}>{l}</button>
            ))}
          </div>
          <div style={S.content}>
            {sec==="pre"&&<PreSec pre={draft.pre} upPre={upPre} upSemi={upSemi} locked={false}/>}
            {sec==="grupos"&&<GruposSec mp={draft.mp} gr={gr} setGr={setGr} upM={upM} grupoM={PM.filter(m=>m.g===gr)} results={{}} locked={false} fetiche={draft.fetiche} setFetiche={setFetiche}/>}
            {sec==="elim"&&<ElimSec mp={draft.mp} upM={upM} elimM={[]} locked={false} fetiche={draft.fetiche} setFetiche={setFetiche} penaltis={draft.penaltis||{}} setPenalti={setPenalti}/>}
            {sec==="retos"&&<RetosSec spc={draft.spc} upSpc={upSpc} locked={false}/>}
            <div style={{height:10}}/>
            {!draft.fetiche&&<div style={{background:"rgba(26,20,0,0.8)",border:"1px solid rgba(232,185,35,0.3)",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#e8b923",marginBottom:10}}>⭐ ¡No olvides marcar tu partido fetiche!</div>}
            <button style={{...S.btnGold,width:"100%",padding:"14px",fontSize:15,borderRadius:12,...(saving?{opacity:.5}:{})}} onClick={!saving?confirm:undefined}>{saving?"Guardando…":"✅ Confirmar y bloquear"}</button>
            <div style={{color:"#2d3748",fontSize:11,textAlign:"center",marginTop:5}}>⚠ Una vez confirmado no podrás modificarlos</div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── PREDICT VIEW ─────────────────────────────────────────────────────────────
function PredictView({p,allM,db,onSave,onBack}){
  const[data,setData]=useState({...p,pre:p.pre||{campeon:"",subcampeon:"",tercero:"",goleador:"",mvp:"",semis:["","","",""]},spc:p.spc||{expulsado:"",hattrick:"",revelacion:"",goleada:""},mp:p.mp||{},fetiche:p.fetiche||null});
  const[sec,setSec]=useState("pre");
  const[gr,setGr]=useState("A");
  const[saving,setSaving]=useState(false);
  const locked=p.locked||false;
  const elimM=allM.filter(m=>m.ph!=="grupos"),grupoM=allM.filter(m=>m.g===gr);
  const upPre=(k,v)=>setData(d=>({...d,pre:{...d.pre,[k]:v}}));
  const upSemi=(i,v)=>setData(d=>{const s=[...(d.pre.semis||["","","",""])];s[i]=v;return{...d,pre:{...d.pre,semis:s}};});
  const upSpc=(k,v)=>setData(d=>({...d,spc:{...d.spc,[k]:v}}));
  const upM=(id,side,v)=>setData(d=>({...d,mp:{...d.mp,[id]:{...(d.mp[id]||{}),[side]:v}}}));
  const setFetiche=id=>setData(d=>({...d,fetiche:d.fetiche===id?null:id}));
  const setPenalti=(id,ph)=>setData(d=>{
    const current=d.penaltis||{};
    const newP={};
    Object.entries(current).forEach(([k,v])=>{if(v!==ph)newP[k]=v;});
    if(current[id])return{...d,penaltis:newP};
    newP[id]=ph;
    return{...d,penaltis:newP};
  });
  const mp=data.mp||{};
  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={onBack}>◀</button>
          <div style={{fontWeight:800,fontSize:15,color:"#e8b923"}}>✏️ {p.name}</div>
          {!locked?<button style={{...S.btnGold,...(saving?{opacity:.5}:{})}} onClick={async()=>{setSaving(true);await onSave(data);setSaving(false);}}>{saving?"…":"💾 Confirmar"}</button>
          :<div style={{fontSize:11,color:"#10b981"}}>🔒 Bloqueado</div>}
        </div>
      </div>
      {locked&&<div style={{background:"rgba(13,26,0,0.8)",borderBottom:"1px solid rgba(16,185,129,0.2)",padding:"6px 16px",fontSize:12,color:"#4ade80"}}>🔒 Solo lectura</div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,padding:"10px 16px",background:"rgba(0,0,0,0.3)",borderBottom:"1px solid rgba(255,255,255,0.05)"}}>
        <Chip label="Pre-Mundial" done={[data.pre?.campeon,data.pre?.subcampeon,data.pre?.tercero,data.pre?.goleador,data.pre?.mvp].filter(Boolean).length} total={5}/>
        <Chip label="Partidos" done={PM.filter(m=>mp[m.id]&&mp[m.id].l!==""&&mp[m.id].v!=="").length} total={72}/>
        <Chip label="Fetiche" done={data.fetiche?1:0} total={1}/>
      </div>
      <div style={S.secNav}>
        {[["pre","🏆 Pre"],["grupos","⚽ Grupos"],["elim","🏟 Elim."],["retos","⭐ Retos"]].map(([id,l])=>(
          <button key={id} style={{...S.snBtn,...(sec===id?S.snA:{})}} onClick={()=>setSec(id)}>{l}</button>
        ))}
      </div>
      <div style={S.content}>
        {sec==="pre"&&<PreSec pre={data.pre} upPre={upPre} upSemi={upSemi} locked={locked&&!!db.rPre}/>}
        {sec==="grupos"&&<GruposSec mp={mp} gr={gr} setGr={setGr} upM={upM} grupoM={grupoM} results={db.results||{}} locked={locked} fetiche={data.fetiche} setFetiche={setFetiche}/>}
        {sec==="elim"&&<ElimSec mp={mp} upM={upM} elimM={elimM} locked={locked} fetiche={data.fetiche} setFetiche={setFetiche} penaltis={data.penaltis||{}} setPenalti={locked?null:setPenalti}/>}
        {sec==="retos"&&<RetosSec spc={data.spc} upSpc={upSpc} locked={locked}/>}
        {!locked&&<><div style={{height:10}}/><button style={{...S.btnGold,width:"100%",padding:"14px",fontSize:15}} onClick={async()=>{setSaving(true);await onSave(data);setSaving(false);}}>💾 Confirmar y bloquear</button><div style={{color:"#2d3748",fontSize:11,textAlign:"center",marginTop:5}}>⚠ No podrás modificarlos después</div></>}
      </div>
    </div>
  );
}

// ─── PREDICTION SECTIONS ──────────────────────────────────────────────────────
function PreSec({pre,upPre,upSemi,locked}){
  return(
    <div>
      {locked&&<div style={{background:"rgba(13,26,0,0.8)",border:"1px solid rgba(16,185,129,0.2)",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#4ade80",marginBottom:12}}>🔒 Solo lectura</div>}
      <div style={S.block}>
        <div style={S.blockT}>🏆 Podio final</div>
        {[["campeon","Campeón del mundo","30 pts"],["subcampeon","Subcampeón","20 pts"],["tercero","Tercer puesto","15 pts"]].map(([k,l,pts])=>(
          <div key={k} style={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:13,color:"#e2e8f0"}}>{l}</span><span style={{fontSize:11,color:"#e8b923"}}>{pts}</span></div>
            <TeamSel val={pre[k]||""} onChange={v=>upPre(k,v)} disabled={locked}/>
          </div>
        ))}
      </div>
      <div style={S.block}>
        <div style={S.blockT}>👟 Premios individuales</div>
        {[["goleador","Máximo goleador","20 pts"],["mvp","Mejor jugador (MVP)","15 pts"]].map(([k,l,pts])=>(
          <div key={k} style={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:13,color:"#e2e8f0"}}>{l}</span><span style={{fontSize:11,color:"#e8b923"}}>{pts}</span></div>
            <input style={{...S.inp,opacity:locked?.6:1}} value={pre[k]||""} onChange={e=>!locked&&upPre(k,e.target.value)} disabled={locked} placeholder="Nombre del jugador…"/>
          </div>
        ))}
      </div>
      <div style={S.block}>
        <div style={S.blockT}>4️⃣ Semifinalistas <span style={{fontSize:11,color:"#e8b923",fontWeight:400}}>(10 pts c/u)</span></div>
        {[0,1,2,3].map(i=>(
          <div key={i} style={{marginBottom:8}}>
            <div style={{fontSize:12,color:"#4a5568",marginBottom:3}}>Semifinalista {i+1}</div>
            <TeamSel val={(pre.semis||[])[i]||""} onChange={v=>upSemi(i,v)} disabled={locked}/>
          </div>
        ))}
      </div>
    </div>
  );
}

function GruposSec({mp,gr,setGr,upM,grupoM,results,locked,fetiche,setFetiche}){
  const gs=Object.keys(GR);
  return(
    <div>
      <div style={{fontSize:12,color:"#4a5568",marginBottom:4}}>Pronostica el marcador · ×1 en grupos</div>
      <div style={{fontSize:11,color:"#e8b923",marginBottom:8}}>⭐ Toca la estrella para marcar tu partido fetiche (solo uno)</div>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:10}}>
        {gs.map(g=>{
          const total=PM.filter(m=>m.g===g).length,done=PM.filter(m=>m.g===g&&mp[m.id]&&mp[m.id].l!==""&&mp[m.id].v!=="").length;
          const hasFet=PM.filter(m=>m.g===g).some(m=>m.id===fetiche);
          return(
            <button key={g} style={{...S.chip,...(gr===g?S.chipA:{}),position:"relative",...(hasFet?{borderColor:"rgba(232,185,35,0.4)"}:{})}} onClick={()=>setGr(g)}>
              Gr.{g}
              {done===total&&<span style={{position:"absolute",top:-5,right:-5,background:"#4ade80",color:"#000",borderRadius:"50%",width:13,height:13,fontSize:9,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900}}>✓</span>}
            </button>
          );
        })}
      </div>
      <div style={{display:"flex",gap:4,marginBottom:10,flexWrap:"wrap"}}>
        {(GR[gr]||[]).map(t=><span key={t} style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:10,padding:"2px 7px",fontSize:11,color:"#4a5568"}}>{t}</span>)}
      </div>
      {grupoM.map(m=>{
        const pred=mp[m.id]||{},res=results[m.id]||m.result,pts=res?mPts(pred,res,"grupos",m.id===fetiche):null,isFet=m.id===fetiche;
        return(
          <div key={m.id} style={{...S.mCard,borderColor:isFet?"rgba(232,185,35,0.5)":res?"rgba(16,185,129,0.25)":(pred.l!==""&&pred.v!=="")?"rgba(232,185,35,0.2)":"rgba(255,255,255,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
              <span style={{fontSize:11,color:"#2d3748"}}>{fd(m.d)}</span>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {isFet&&<span style={{fontSize:10,color:"#e8b923",fontWeight:700}}>×5/×3</span>}
                {!locked&&<button onClick={()=>setFetiche(m.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,padding:0,opacity:isFet?1:.25,filter:isFet?"none":"grayscale(1)"}}>⭐</button>}
                {locked&&isFet&&<span style={{fontSize:16}}>⭐</span>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 44px 14px 44px 1fr",alignItems:"center",gap:4}}>
              <div style={{textAlign:"right",fontWeight:600,color:"#e2e8f0",fontSize:12,lineHeight:1.3}}>{m.l}</div>
              <input type="number" min={0} max={20} value={pred.l||""} onChange={e=>!locked&&upM(m.id,"l",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1,...(isFet?{borderColor:"rgba(232,185,35,0.5)"}:{})}} disabled={locked} placeholder="?"/>
              <span style={{textAlign:"center",color:"rgba(255,255,255,0.1)",fontWeight:700}}>–</span>
              <input type="number" min={0} max={20} value={pred.v||""} onChange={e=>!locked&&upM(m.id,"v",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1,...(isFet?{borderColor:"rgba(232,185,35,0.5)"}:{})}} disabled={locked} placeholder="?"/>
              <div style={{fontWeight:600,color:"#e2e8f0",fontSize:12,lineHeight:1.3}}>{m.v}</div>
            </div>
            {res&&<div style={{textAlign:"center",fontSize:11,marginTop:6}}><span style={{color:"#10b981"}}>Real: {res.l}–{res.v}</span>{pts!==null&&<span style={{color:isFet?"#e8b923":"#10b981",marginLeft:8,fontWeight:700}}>{isFet?"⭐":""} +{pts}pts</span>}{m.fueAPenaltis&&<span style={{color:"#818cf8",marginLeft:6,fontWeight:700}}>🥅{isPen?" +3pts":" (no apostaste)"}</span>}</div>}
          </div>
        );
      })}
    </div>
  );
}

function ElimSec({mp,upM,elimM,locked,fetiche,setFetiche,penaltis,setPenalti}){
  if(!elimM||elimM.length===0)return(
    <div style={{...S.empty,paddingTop:50}}>
      <div style={{fontSize:36}}>⏳</div>
      <div style={{color:"#2d3748",marginTop:8,fontWeight:600}}>Los partidos de eliminatoria aparecerán aquí</div>
    </div>
  );
  const rondas=["dieciseisavos","octavos","cuartos","semifinales"];
  return(
    <div>
      {setPenalti&&<div style={{fontSize:11,color:"#818cf8",marginBottom:10,background:"rgba(99,102,241,0.08)",border:"1px solid rgba(99,102,241,0.2)",borderRadius:8,padding:"7px 12px"}}>🥅 Marca un partido por ronda que crees que irá a penaltis. Si aciertas, +3 pts extra.</div>}
      {elimM.map(m=>{
        const pred=mp[m.id]||{},res=m.result,pts=res?mPts(pred,res,m.ph,m.id===fetiche):null,isFet=m.id===fetiche,isPen=penaltis&&!!penaltis[m.id];
        return(
          <div key={m.id} style={{...S.mCard,borderColor:isFet?"rgba(232,185,35,0.5)":isPen?"rgba(129,140,248,0.4)":res?"rgba(16,185,129,0.25)":"rgba(255,255,255,0.06)"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
              <div><span style={{fontSize:11,color:"#2d3748"}}>{fd(m.date||"")}</span><span style={{fontSize:11,color:"#60a5fa",fontWeight:700,marginLeft:8}}>{PHL[m.ph]} ·×{MULT[m.ph]}</span></div>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {isFet&&<span style={{fontSize:10,color:"#e8b923",fontWeight:700}}>×5/×3</span>}
                {!locked&&<button onClick={()=>setFetiche(m.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,padding:0,opacity:isFet?1:.25,filter:isFet?"none":"grayscale(1)"}}>⭐</button>}
                {locked&&isFet&&<span>⭐</span>}
                {setPenalti&&!locked&&<button onClick={()=>setPenalti(m.id,m.ph)} title="Marcar como que va a penaltis (+3pts si aciertas)" style={{background:isPen?"rgba(99,102,241,0.2)":"none",border:isPen?"1px solid #818cf8":"none",borderRadius:6,cursor:"pointer",fontSize:16,padding:"2px 5px",opacity:isPen?1:.3}}>🥅</button>}
                {setPenalti&&locked&&isPen&&<span style={{fontSize:16}}>🥅</span>}
                {m.fueAPenaltis&&isPen&&<span style={{fontSize:10,color:"#818cf8",fontWeight:700}}>+3</span>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 44px 14px 44px 1fr",alignItems:"center",gap:4}}>
              <div style={{textAlign:"right",fontWeight:600,color:"#e2e8f0",fontSize:12}}>{m.l}</div>
              <input type="number" min={0} max={20} value={pred.l||""} onChange={e=>!locked&&upM(m.id,"l",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1}} disabled={locked} placeholder="?"/>
              <span style={{textAlign:"center",color:"rgba(255,255,255,0.1)",fontWeight:700}}>–</span>
              <input type="number" min={0} max={20} value={pred.v||""} onChange={e=>!locked&&upM(m.id,"v",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1}} disabled={locked} placeholder="?"/>
              <div style={{fontWeight:600,color:"#e2e8f0",fontSize:12}}>{m.v}</div>
            </div>
            {res&&<div style={{textAlign:"center",fontSize:11,marginTop:6}}><span style={{color:"#10b981"}}>Real: {res.l}–{res.v}</span>{pts!==null&&<span style={{color:isFet?"#e8b923":"#10b981",marginLeft:8,fontWeight:700}}>{isFet?"⭐":""} +{pts}pts</span>}{m.fueAPenaltis&&<span style={{color:"#818cf8",marginLeft:6,fontWeight:700}}>🥅{isPen?" +3pts":" (no apostaste)"}</span>}</div>}
          </div>
        );
      })}
    </div>
  );
}

function RetosSec({spc,upSpc,locked}){
  return(
    <div>
      {[{k:"expulsado",t:"🟥 Primer expulsado",p:"10 pts",d:"Primer jugador en ver roja",ph:"Nombre del jugador"},{k:"hattrick",t:"🎩 Primer hat-trick",p:"10 pts",d:"Primero en marcar 3 goles en un partido",ph:"Nombre del jugador"}].map(r=>(
        <div key={r.k} style={S.block}>
          <div style={{display:"flex",justifyContent:"space-between"}}><span style={S.blockT}>{r.t}</span><span style={{fontSize:11,color:"#e8b923"}}>{r.p}</span></div>
          <div style={{fontSize:11,color:"#4a5568",marginBottom:8,marginTop:3}}>{r.d}</div>
          <input style={{...S.inp,opacity:locked?.6:1}} value={spc[r.k]||""} onChange={e=>!locked&&upSpc(r.k,e.target.value)} disabled={locked} placeholder={r.ph}/>
        </div>
      ))}
      <div style={S.block}>
        <div style={{display:"flex",justifyContent:"space-between"}}><span style={S.blockT}>🚀 Equipo revelación</span><span style={{fontSize:11,color:"#e8b923"}}>15 pts</span></div>
        <div style={{fontSize:11,color:"#4a5568",marginBottom:8,marginTop:3}}>El que llegue más lejos de lo esperado</div>
        <TeamSel val={spc.revelacion||""} onChange={v=>upSpc("revelacion",v)} disabled={locked}/>
      </div>
      <div style={S.block}>
        <div style={{display:"flex",justifyContent:"space-between"}}><span style={S.blockT}>🎯 Mayor goleada</span><span style={{fontSize:11,color:"#e8b923"}}>10 pts</span></div>
        <div style={{fontSize:11,color:"#4a5568",marginBottom:8,marginTop:3}}>Resultado exacto — ej: Brasil 7-0 Haití</div>
        <input style={{...S.inp,opacity:locked?.6:1}} value={spc.goleada||""} onChange={e=>!locked&&upSpc("goleada",e.target.value)} disabled={locked} placeholder="Ej: Brasil 7-0 Haití"/>
      </div>
    </div>
  );
}

// ─── PARTIDOS TAB ─────────────────────────────────────────────────────────────
function PartidosTab({allM}){
  const[gf,setGf]=useState("A");
  const gs=Object.keys(GR);
  const isElim=gf==="ELIM",shown=isElim?allM.filter(m=>m.ph!=="grupos"):allM.filter(m=>m.g===gf);
  return(
    <div>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:12}}>
        {gs.map(g=>{const hasR=allM.filter(m=>m.g===g).some(m=>m.result);return<button key={g} style={{...S.chip,...(gf===g?S.chipA:{}),...(hasR?{borderColor:"rgba(16,185,129,0.3)"}:{})}} onClick={()=>setGf(g)}>Gr.{g}</button>;})}
        {allM.some(m=>m.ph!=="grupos")&&<button style={{...S.chip,...(isElim?S.chipA:{})}} onClick={()=>setGf("ELIM")}>Elim.</button>}
      </div>
      {!isElim&&<TablaGrupo grupo={gf} allM={allM}/>}
      <div style={{fontSize:11,color:"#4a5568",fontWeight:700,letterSpacing:1,marginBottom:8,textTransform:"uppercase"}}>⚽ Partidos</div>
      {shown.length===0?<div style={S.empty}>Sin partidos todavía</div>:shown.map(m=>(
        <div key={m.id} style={{...S.mCard,borderColor:m.result?"rgba(16,185,129,0.25)":"rgba(255,255,255,0.06)"}}>
          <div style={{fontSize:11,color:"#2d3748",marginBottom:4}}>{fd(m.d||m.date||"")}{m.ph!=="grupos"?" · "+PHL[m.ph]:""}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr auto 1fr",alignItems:"center",gap:8}}>
            <div style={{textAlign:"right",fontWeight:600,color:"#e2e8f0",fontSize:13}}>{m.l}</div>
            {m.result?<div style={{textAlign:"center",fontWeight:900,color:"#10b981",fontSize:22,minWidth:60,background:"rgba(16,185,129,0.08)",borderRadius:8,padding:"2px 6px"}}>{m.result.l}–{m.result.v}</div>
            :<div style={{textAlign:"center",color:"rgba(255,255,255,0.08)",fontSize:16,minWidth:60}}>vs</div>}
            <div style={{fontWeight:600,color:"#e2e8f0",fontSize:13}}>{m.v}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ─── REGLAS TAB ───────────────────────────────────────────────────────────────
function ReglasTab(){
  const B=[
    {t:"1️⃣ Pre-Mundial",c:"#e8b923",items:[["Campeón","30pts"],["Subcampeón","20pts"],["3er puesto","15pts"],["Máx. goleador","20pts"],["MVP","15pts"],["4 Semifinalistas","10pts c/u"]]},
    {t:"2️⃣ Por partido",c:"#4ade80",items:[["Resultado exacto","5pts"],["Ganador/empate","2pts"],["Goles de un equipo","1pt c/u"]]},
    {t:"⭐ Partido Fetiche",c:"#e8b923",items:[["Si aciertas resultado exacto","pts ×5"],["Si aciertas 1X2","pts ×3"],["Solo 1 por participante","Elige bien"]]},
    {t:"3️⃣ Multiplicadores fase",c:"#60a5fa",items:[["Grupos","×1"],["Octavos","×1.5"],["Cuartos","×2"],["Semis","×3"],["Final","×5"]]},
    {t:"4️⃣ Retos especiales",c:"#f472b6",items:[["Primer expulsado","10pts"],["Primer hat-trick","10pts"],["Equipo revelación","15pts"],["Mayor goleada","10pts"]]},
    {t:"💰 Premios",c:"#fb923c",items:[["1er puesto","60% del bote"],["2º puesto","25% del bote"],["3er puesto","15% del bote"],["Farolillo rojo","¡Ronda épica!"]]},
  ];
  return(
    <div>
      {B.map((b,i)=>(
        <div key={i} style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid "+b.c,borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{fontWeight:700,fontSize:13,color:b.c,marginBottom:8}}>{b.t}</div>
          {b.items.map(([l,p])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:13}}>
              <span style={{color:"#94a3b8"}}>{l}</span><span style={{fontWeight:700,color:b.c}}>{p}</span>
            </div>
          ))}
        </div>
      ))}
      <div style={{background:"rgba(5,26,5,0.8)",border:"1px solid rgba(16,185,129,0.2)",borderRadius:10,padding:"12px",fontSize:13,color:"#86efac",lineHeight:1.7}}>
        💡 <strong>Ejemplo semis:</strong> Pronosticas 2-1, real 3-1 → ganador(2) + goles visitante(1) = 3 × 3 = <strong style={{color:"#e8b923"}}>9pts</strong>
      </div>
    </div>
  );
}

// ─── COMPARE VIEW ─────────────────────────────────────────────────────────────
function CompareView({p,allM,scores,db,bote,onBack}){
  const[gr,setGr]=useState("A");
  const playedM=allM.filter(m=>m.result),grupoPlayed=playedM.filter(m=>m.g===gr);
  const mp=p.mp||{},fetiche=p.fetiche||null;
  const pp=prePts(p.pre||{},db.rPre),sp=spcPts(p.spc||{},db.rSpc);
  let matchT=0;playedM.forEach(m=>{matchT+=mPts(mp[m.id],m.result,m.ph||"grupos",m.id===fetiche);});
  matchT=Math.round(matchT*10)/10;
  const pos=scores.findIndex(x=>x.id===p.id)+1,gs=Object.keys(GR);
  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={onBack}>◀</button>
          <div style={{textAlign:"center"}}><div style={{fontWeight:800,fontSize:15,color:"#e8b923"}}>📊 {p.name}</div><div style={{fontSize:11,color:"#4a5568"}}>Desglose de puntuación</div></div>
          <div/>
        </div>
      </div>
      <div style={S.content}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
          {[{l:"Posición",v:"#"+pos,c:"#e8b923"},{l:"Total",v:(scores.find(x=>x.id===p.id)?.tot||0)+" pts",c:"#10b981"},{l:"Pre-Mundial",v:pp+" pts",c:"#60a5fa"},{l:"Partidos",v:matchT+" pts",c:"#a78bfa"},{l:"Retos",v:sp+" pts",c:"#f472b6"},{l:"Premio est.",v:pos<=3?Math.round(bote*[.6,.25,.15][pos-1])+"€":"—",c:"#10b981"}].map(x=>(
            <div key={x.l} style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:10,padding:"10px 12px"}}>
              <div style={{fontSize:11,color:"#4a5568"}}>{x.l}</div>
              <div style={{fontWeight:800,fontSize:17,color:x.c,marginTop:2}}>{x.v}</div>
            </div>
          ))}
        </div>
        {fetiche&&<div style={{background:"rgba(26,20,0,0.8)",border:"1px solid rgba(232,185,35,0.3)",borderRadius:10,padding:"10px 12px",marginBottom:12,fontSize:12,color:"#e8b923"}}>⭐ Fetiche: {(allM.find(m=>m.id===fetiche)||{l:"?",v:"?"}).l} vs {(allM.find(m=>m.id===fetiche)||{l:"?",v:"?"}).v}</div>}
        {db.rPre&&(
          <div style={{background:"rgba(0,0,0,0.3)",borderLeft:"3px solid #60a5fa",borderRadius:10,padding:"12px",marginBottom:12}}>
            <div style={{fontWeight:700,color:"#60a5fa",fontSize:13,marginBottom:8}}>🏆 Pre-Mundial — {pp} pts</div>
            {[["Campeón",p.pre?.campeon,db.rPre?.campeon,30],["Subcampeón",p.pre?.subcampeon,db.rPre?.subcampeon,20],["3er puesto",p.pre?.tercero,db.rPre?.tercero,15],["Goleador",p.pre?.goleador,db.rPre?.goleador,20],["MVP",p.pre?.mvp,db.rPre?.mvp,15]].map(([l,pred,real,max])=>{
              const hit=pred&&real&&pred===real;
              return(<div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid rgba(255,255,255,0.04)",fontSize:12}}><span style={{color:"#4a5568"}}>{l}</span><span style={{color:hit?"#10b981":pred?"#2d3748":"#1a1a1a"}}>{pred||"—"}{hit?" ✓ +"+max+"pts":real&&pred?" ✗ ("+real+")":""}</span></div>);
            })}
          </div>
        )}
        {playedM.length>0&&(
          <div>
            <div style={{fontWeight:700,color:"#a78bfa",fontSize:13,marginBottom:8}}>⚽ Partidos — {matchT} pts</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:10}}>
              {gs.map(g=><button key={g} style={{...S.chip,...(gr===g?S.chipA:{})}} onClick={()=>setGr(g)}>Gr.{g}</button>)}
            </div>
            {grupoPlayed.length===0?<div style={{color:"#2d3748",textAlign:"center",padding:"16px",fontSize:13}}>Sin partidos jugados aquí</div>:
              grupoPlayed.map(m=>{
                const pred=mp[m.id],isFet=m.id===fetiche,pts=mPts(pred,m.result,m.ph||"grupos",isFet);
                return(
                  <div key={m.id} style={{background:"rgba(0,0,0,0.3)",border:"1px solid "+(pts>0?"rgba(16,185,129,0.15)":"rgba(255,255,255,0.04)"),borderRadius:8,padding:"10px",marginBottom:6}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                      <span style={{fontSize:11,color:"#2d3748"}}>{fd(m.d||"")} {isFet&&<span style={{color:"#e8b923"}}>⭐</span>}</span>
                      <span style={{fontSize:12,fontWeight:700,color:pts>0?"#10b981":"#2d3748"}}>+{pts}pts</span>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr auto 1fr",gap:8,alignItems:"center"}}>
                      <div style={{textAlign:"right",fontSize:12,color:"#94a3b8"}}>{m.l}</div>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontWeight:900,color:"#10b981",fontSize:15}}>{m.result.l}–{m.result.v}</div>
                        {pred?<div style={{fontSize:11,color:pts>0?"#e8b923":"#2d3748"}}>Pred: {pred.l}–{pred.v}</div>:<div style={{fontSize:11,color:"#1a1a1a"}}>Sin pred.</div>}
                      </div>
                      <div style={{fontSize:12,color:"#94a3b8"}}>{m.v}</div>
                    </div>
                  </div>
                );
              })
            }
          </div>
        )}
      </div>
    </div>
  );
}

// ─── ADMIN VIEW ───────────────────────────────────────────────────────────────
function AdminView({db,upDb,allM,onBack}){
  const[tab,setTab]=useState("res");
  const[gf,setGf]=useState("A");
  const[eid,setEid]=useState(null);
  const[er,setEr]=useState({l:"",v:""});
  const[pf,setPf]=useState(db.rPre||{campeon:"",subcampeon:"",tercero:"",goleador:"",mvp:"",semis:["","","",""]});
  const[sf,setSf]=useState(db.rSpc||{expulsado:"",hattrick:"",revelacion:"",goleada:""});
  const[nm,setNm]=useState({l:"",v:"",ph:"dieciseisavos",date:""});
  const[saving,setSaving]=useState(false);
  const gs=Object.keys(GR);
  const isElim=gf==="ELIM",shown=isElim?allM.filter(m=>m.ph!=="grupos"):allM.filter(m=>m.g===gf);
  const saveRes=async id=>{setSaving(true);const newRes={...(db.results||{}),[id]:er};await upDb("results",newRes);const nEM=(db.extraM||[]).map(m=>m.id===id?{...m,result:er}:m);await upDb("extraM",nEM);setEid(null);setSaving(false);};
  const togglePenaltis=async id=>{const nEM=(db.extraM||[]).map(m=>m.id===id?{...m,fueAPenaltis:!m.fueAPenaltis}:m);await upDb("extraM",nEM);};
  const resetRes=async id=>{const r={...(db.results||{})};delete r[id];await upDb("results",r);};
  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={onBack}>◀</button>
          <div style={{fontWeight:900,fontSize:15,color:"#e8b923"}}>⚙️ Panel Admin</div>
          <div style={{fontSize:11,color:"#4a5568"}}>{(db.parts||[]).length} participantes</div>
        </div>
      </div>
      <div style={S.tabs}>
        {[["res","⚽ Res."],["pre","🏆 Global"],["spc","⭐ Retos"],["ext","➕ Partidos"],["cfg","⚙️ Config"]].map(([id,l])=>(
          <button key={id} style={{...S.tab,...(tab===id?S.tabA:{})}} onClick={()=>setTab(id)}>{l}</button>
        ))}
      </div>
      <div style={S.content}>
        {tab==="res"&&(
          <div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:10}}>
              {gs.map(g=><button key={g} style={{...S.chip,...(gf===g?S.chipA:{})}} onClick={()=>setGf(g)}>Gr.{g}</button>)}
              <button style={{...S.chip,...(isElim?S.chipA:{})}} onClick={()=>setGf("ELIM")}>Elim.</button>
            </div>
            {shown.map(m=>{
              const res=(db.results||{})[m.id]||m.result;
              return(
                <div key={m.id} style={{...S.mCard,display:"flex",flexDirection:"row",alignItems:"center",gap:8}}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontWeight:600,color:"#e2e8f0",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.l} vs {m.v}</div>
                    <div style={{fontSize:10,color:"#2d3748"}}>{fd(m.d||m.date||"")}{m.ph!=="grupos"?" · "+PHL[m.ph]:""}</div>
                    {res&&<div style={{display:"flex",alignItems:"center",gap:8,marginTop:2}}><span style={{color:"#10b981",fontWeight:700,fontSize:11}}>✅ {res.l}–{res.v}</span><button style={{background:"#7f1d1d",border:"none",borderRadius:4,padding:"1px 6px",color:"#fff",fontSize:10,cursor:"pointer"}} onClick={()=>resetRes(m.id)}>Reset</button></div>}
                  </div>
                  {m.ph!=="grupos"&&<button style={{background:m.fueAPenaltis?"rgba(99,102,241,0.3)":"rgba(0,0,0,0.3)",border:"1px solid "+( m.fueAPenaltis?"#818cf8":"rgba(255,255,255,0.1)"),borderRadius:6,padding:"2px 7px",color:m.fueAPenaltis?"#818cf8":"#4a5568",fontSize:11,cursor:"pointer",marginRight:4}} onClick={()=>togglePenaltis(m.id)}>🥅{m.fueAPenaltis?" Penaltis ✓":" ¿Penaltis?"}</button>}
                  {eid===m.id?(
                    <div style={{display:"flex",gap:3,alignItems:"center",flexShrink:0}}>
                      <input type="number" min={0} max={20} value={er.l} onChange={e=>setEr({...er,l:e.target.value})} style={{...S.scoreI,width:34}}/>
                      <span style={{color:"#fff",fontSize:12}}>–</span>
                      <input type="number" min={0} max={20} value={er.v} onChange={e=>setEr({...er,v:e.target.value})} style={{...S.scoreI,width:34}}/>
                      <button style={{...S.btnSm,...(saving?{opacity:.5}:{})}} onClick={()=>!saving&&saveRes(m.id)}>✅</button>
                      <button style={{...S.btnSm,background:"#222"}} onClick={()=>setEid(null)}>✕</button>
                    </div>
                  ):(
                    <button style={S.btnAct} onClick={()=>{setEid(m.id);setEr((db.results||{})[m.id]||{l:"",v:""});}}>✏️</button>
                  )}
                </div>
              );
            })}
          </div>
        )}
        {tab==="pre"&&(
          <div>
            {[["campeon","🏆 Campeón"],["subcampeon","🥈 Subcampeón"],["tercero","🥉 3er puesto"]].map(([k,l])=>(
              <div key={k} style={{marginBottom:10}}><div style={S.fl}>{l}</div><TeamSel val={pf[k]||""} onChange={v=>setPf({...pf,[k]:v})}/></div>
            ))}
            <div style={{marginBottom:10}}><div style={S.fl}>👟 Goleador</div><input style={S.inp} value={pf.goleador||""} onChange={e=>setPf({...pf,goleador:e.target.value})}/></div>
            <div style={{marginBottom:10}}><div style={S.fl}>⭐ MVP</div><input style={S.inp} value={pf.mvp||""} onChange={e=>setPf({...pf,mvp:e.target.value})}/></div>
            <div style={{marginBottom:10}}>
              <div style={S.fl}>4️⃣ Semifinalistas</div>
              {[0,1,2,3].map(i=>(
                <div key={i} style={{marginBottom:4}}><TeamSel val={(pf.semis||[])[i]||""} onChange={v=>{const s=[...(pf.semis||["","","",""])];s[i]=v;setPf({...pf,semis:s});}}/></div>
              ))}
            </div>
            <button style={{...S.btnGold,width:"100%"}} onClick={async()=>{setSaving(true);await upDb("rPre",pf);setSaving(false);}}>💾 Guardar</button>
            {db.rPre&&<button style={{...S.btnAct,width:"100%",marginTop:8,background:"#7f1d1d",padding:"9px",display:"flex",justifyContent:"center"}} onClick={async()=>upDb("rPre",null)}>🔓 Desbloquear preds.</button>}
          </div>
        )}
        {tab==="spc"&&(
          <div>
            {[["expulsado","🟥 Primer expulsado"],["hattrick","🎩 Primer hat-trick"],["goleada","🎯 Mayor goleada"]].map(([k,l])=>(
              <div key={k} style={{marginBottom:10}}><div style={S.fl}>{l}</div><input style={S.inp} value={sf[k]||""} onChange={e=>setSf({...sf,[k]:e.target.value})}/></div>
            ))}
            <div style={{marginBottom:10}}><div style={S.fl}>🚀 Equipo revelación</div><TeamSel val={sf.revelacion||""} onChange={v=>setSf({...sf,revelacion:v})}/></div>
            <button style={{...S.btnGold,width:"100%"}} onClick={async()=>{setSaving(true);await upDb("rSpc",sf);setSaving(false);}}>💾 Guardar retos</button>
          </div>
        )}
        {tab==="ext"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
              <TeamSel val={nm.l} onChange={v=>setNm({...nm,l:v})} placeholder="Local"/>
              <TeamSel val={nm.v} onChange={v=>setNm({...nm,v:v})} placeholder="Visitante"/>
              <select style={S.sel} value={nm.ph} onChange={e=>setNm({...nm,ph:e.target.value})}>
                {["dieciseisavos","octavos","cuartos","semifinales","tercerpuesto","final"].map(p=><option key={p} value={p}>{PHL[p]} ·×{MULT[p]}</option>)}
              </select>
              <input style={S.inp} type="date" value={nm.date} onChange={e=>setNm({...nm,date:e.target.value})}/>
            </div>
            <button style={{...S.btnGold,width:"100%"}} onClick={async()=>{if(!nm.l||!nm.v)return;await upDb("extraM",[...(db.extraM||[]),{...nm,id:"x_"+Date.now(),result:null,g:null}]);setNm({l:"",v:"",ph:"dieciseisavos",date:""});}}>➕ Añadir partido</button>
            <div style={{marginTop:12,background:"rgba(0,0,0,0.3)",border:"1px solid rgba(232,185,35,0.2)",borderRadius:10,padding:"12px"}}>
              <div style={{fontSize:12,color:"#e8b923",fontWeight:700,marginBottom:6}}>⚡ Cargar dieciseisavos del Mundial 2026</div>
              <div style={{fontSize:11,color:"#4a5568",marginBottom:10}}>Añade los 16 partidos de golpe</div>
              <button style={{...S.btnGold,width:"100%",background:"linear-gradient(135deg,#065f46,#047857)"}} onClick={async()=>{
                if(!window.confirm("¿Añadir los 16 partidos de dieciseisavos?"))return;
                const d16=[
                  {l:"Alemania 🇩🇪",v:"Paraguay 🇵🇾",date:"2026-06-28"},{l:"Francia 🇫🇷",v:"Suecia 🇸🇪",date:"2026-06-28"},
                  {l:"Sudáfrica 🇿🇦",v:"Canadá 🇨🇦",date:"2026-06-29"},{l:"Países Bajos 🇳🇱",v:"Marruecos 🇲🇦",date:"2026-06-29"},
                  {l:"Portugal 🇵🇹",v:"Croacia 🇭🇷",date:"2026-06-30"},{l:"España 🇪🇸",v:"Austria 🇦🇹",date:"2026-06-30"},
                  {l:"EE.UU. 🇺🇸",v:"Bosnia 🇧🇦",date:"2026-07-01"},{l:"Bélgica 🇧🇪",v:"Senegal 🇸🇳",date:"2026-07-01"},
                  {l:"Brasil 🇧🇷",v:"Japón 🇯🇵",date:"2026-07-02"},{l:"Costa de Marfil 🇨🇮",v:"Noruega 🇳🇴",date:"2026-07-02"},
                  {l:"México 🇲🇽",v:"Ecuador 🇪🇨",date:"2026-07-03"},{l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"RD Congo 🇨🇩",date:"2026-07-03"},
                  {l:"Argentina 🇦🇷",v:"Curazao 🇨🇼",date:"2026-07-04"},{l:"Australia 🇦🇺",v:"Egipto 🇪🇬",date:"2026-07-04"},
                  {l:"Suiza 🇨🇭",v:"Argelia 🇩🇿",date:"2026-07-05"},{l:"Colombia 🇨🇴",v:"Ghana 🇬🇭",date:"2026-07-05"},
                ].map((m,i)=>({...m,ph:"dieciseisavos",id:"d16_"+(Date.now()+i),result:null,g:null}));
                await upDb("extraM",[...(db.extraM||[]),...d16]);
              }}>⚡ Cargar los 16 partidos</button>
            </div>
            {(db.extraM||[]).map(m=>(
              <div key={m.id} style={{...S.mCard,display:"flex",flexDirection:"row",alignItems:"center",gap:8,marginTop:6}}>
                <div style={{flex:1}}><div style={{fontWeight:600,color:"#e2e8f0",fontSize:12}}>{m.l} vs {m.v}</div><div style={{fontSize:10,color:"#4a5568"}}>{PHL[m.ph]} · {fd(m.date||"")}</div></div>
                <button style={{...S.btnAct,background:"#7f1d1d"}} onClick={async()=>upDb("extraM",(db.extraM||[]).filter(x=>x.id!==m.id))}>🗑️</button>
              </div>
            ))}
          </div>
        )}
        {tab==="cfg"&&(
          <div>
            <div style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:700,color:"#e2e8f0",fontSize:13,marginBottom:4}}>📅 Fecha límite</div>
              <input style={S.inp} type="date" value={db.deadline||""} onChange={async e=>upDb("deadline",e.target.value)}/>
              {db.deadline&&<div style={{fontSize:12,color:isPast(db.deadline)?"#ef4444":"#10b981",marginTop:6}}>{isPast(db.deadline)?"🔒 CERRADO":"⏰ Abierto hasta "+fd(db.deadline)}</div>}
              {db.deadline&&<button style={{...S.btnAct,marginTop:8,background:"#7f1d1d"}} onClick={async()=>upDb("deadline","")}>Quitar límite</button>}
            </div>
            <div style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(13,185,129,0.15)",borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:700,color:"#10b981",fontSize:13,marginBottom:8}}>📥 Exportar datos</div>
              <div style={{fontSize:12,color:"#4a5568",marginBottom:10}}>Descarga un Excel con todos los pronósticos y puntuaciones</div>
              <button style={{...S.btnGold,width:"100%",background:"linear-gradient(135deg,#065f46,#047857)"}} onClick={()=>exportXLSX(db.parts||[],allM,db.rPre,db.rSpc)}>📥 Descargar Excel completo</button>
            </div>
            <div style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:700,color:"#e2e8f0",fontSize:13,marginBottom:10}}>👤 Participantes</div>
              {(db.parts||[]).length===0?<div style={{color:"#2d3748",fontSize:13}}>Sin participantes</div>:(db.parts||[]).map(p=>(
                <div key={p.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:"1px solid rgba(255,255,255,0.04)"}}>
                  <div style={{flex:1,color:"#94a3b8",fontSize:13}}>{p.name}{p.locked&&<span style={{marginLeft:6,fontSize:10,color:"#10b981"}}>🔒</span>}</div>
                  {p.locked&&<button style={{...S.btnAct,fontSize:10,padding:"3px 7px"}} onClick={async()=>{if(!window.confirm("¿Desbloquear a "+p.name+"?"))return;await upDb("parts",(db.parts||[]).map(x=>x.id===p.id?{...x,locked:false}:x));}}>🔓</button>}
                  <button style={{...S.btnAct,background:"#7f1d1d",fontSize:10,padding:"3px 7px"}} onClick={async()=>{if(!window.confirm("¿Eliminar a "+p.name+"?"))return;await upDb("parts",(db.parts||[]).filter(x=>x.id!==p.id));}}>🗑️</button>
                </div>
              ))}
            </div>
            <div style={{background:"rgba(0,0,0,0.3)",border:"1px solid rgba(127,29,29,0.4)",borderRadius:12,padding:"14px"}}>
              <div style={{fontWeight:700,color:"#ff6b35",fontSize:13,marginBottom:8}}>⚠ Zona peligrosa</div>
              <button style={{...S.btnAct,background:"#7f1d1d",width:"100%",padding:"10px"}} onClick={async()=>{if(!window.confirm("¿Borrar TODOS los datos? No se puede deshacer."))return;await Promise.all(["parts","extraM","results","rPre","rSpc","deadline"].map(k=>upDb(k,["parts","extraM"].includes(k)?[]:k==="results"?{}:null)));}}>🗑️ Borrar todos los datos</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const S={
  app:{minHeight:"100vh",background:"linear-gradient(160deg,#0a1628 0%,#0d1f0d 50%,#1a0a00 100%)",fontFamily:"'Segoe UI',system-ui,sans-serif",color:"#fff",maxWidth:600,margin:"0 auto",paddingTop:"env(safe-area-inset-top)",paddingBottom:"env(safe-area-inset-bottom)"},
  hdr:{background:"linear-gradient(180deg,#0a1628,#091220)",borderBottom:"2px solid transparent",borderImage:"linear-gradient(90deg,#c8102e,#e8b923,#c8102e) 1",padding:"12px 16px",paddingTop:"calc(12px + env(safe-area-inset-top))",boxShadow:"0 4px 20px rgba(0,0,0,0.5)"},
  hdrRow:{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,marginBottom:6},
  t1:{fontSize:21,fontWeight:900,background:"linear-gradient(135deg,#e8b923,#fff5c0,#e8b923)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",letterSpacing:2,lineHeight:1},
  t2:{fontSize:11,color:"#4a5568",letterSpacing:4,textTransform:"uppercase"},
  bote:{background:"linear-gradient(135deg,#064e3b,#065f46)",border:"1px solid rgba(16,185,129,0.4)",borderRadius:20,padding:"4px 12px",fontSize:13,color:"#10b981",fontWeight:700,boxShadow:"0 0 12px rgba(16,185,129,0.2)"},
  bannerR:{background:"linear-gradient(135deg,#450a0a,#7f1d1d)",border:"1px solid rgba(220,38,38,0.3)",borderRadius:8,padding:"6px 12px",fontSize:12,color:"#fca5a5",marginBottom:6},
  bannerG:{background:"linear-gradient(135deg,#052e16,#064e3b)",border:"1px solid rgba(22,163,74,0.3)",borderRadius:8,padding:"5px 12px",fontSize:12,color:"#86efac",marginBottom:6},
  tabs:{display:"flex",background:"rgba(0,0,0,0.4)",borderBottom:"1px solid rgba(232,185,35,0.15)",backdropFilter:"blur(10px)"},
  tab:{flex:1,padding:"11px 1px",background:"transparent",border:"none",color:"#4a5568",cursor:"pointer",fontSize:10,fontWeight:700,borderBottom:"2px solid transparent",transition:"all .2s",letterSpacing:.3},
  tabA:{color:"#e8b923",borderBottomColor:"#e8b923",background:"rgba(232,185,35,0.06)"},
  content:{padding:"14px 16px",paddingBottom:"calc(70px + env(safe-area-inset-bottom))"},
  secNav:{display:"grid",gridTemplateColumns:"repeat(4,1fr)",background:"rgba(0,0,0,0.5)",borderBottom:"1px solid rgba(255,255,255,0.06)"},
  snBtn:{padding:"10px 4px",background:"transparent",border:"none",color:"#4a5568",cursor:"pointer",borderBottom:"2px solid transparent",textAlign:"center",fontSize:11,fontWeight:700,transition:"all .2s"},
  snA:{color:"#e8b923",borderBottomColor:"#e8b923",background:"rgba(232,185,35,0.05)"},
  row:{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",background:"linear-gradient(135deg,rgba(15,20,40,0.9),rgba(10,25,15,0.9))",borderRadius:12,marginBottom:6,border:"1px solid rgba(255,255,255,0.07)",boxShadow:"0 2px 10px rgba(0,0,0,0.3)"},
  card:{background:"linear-gradient(135deg,rgba(12,18,38,0.95),rgba(10,20,12,0.95))",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"14px",marginBottom:10,boxShadow:"0 4px 16px rgba(0,0,0,0.3)"},
  regCard:{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,background:"linear-gradient(135deg,#052e16,#064e3b)",border:"1px solid rgba(16,185,129,0.2)",borderRadius:14,padding:"16px",marginBottom:16,boxShadow:"0 4px 20px rgba(16,185,129,0.15)"},
  mCard:{padding:"12px 14px",background:"linear-gradient(135deg,rgba(12,18,38,0.9),rgba(10,20,12,0.9))",borderRadius:10,marginBottom:6,border:"1px solid"},
  block:{background:"rgba(5,10,25,0.8)",border:"1px solid rgba(255,255,255,0.06)",borderRadius:12,padding:"14px",marginBottom:12},
  blockT:{fontWeight:700,fontSize:13,color:"#e2e8f0",marginBottom:10},
  empty:{color:"#2d3748",textAlign:"center",padding:"48px 16px",fontSize:14,lineHeight:1.8},
  inp:{background:"rgba(5,10,25,0.8)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:10,padding:"9px 13px",color:"#fff",fontSize:14,width:"100%",boxSizing:"border-box",outline:"none"},
  sel:{background:"rgba(5,10,25,0.8)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:10,padding:"9px 11px",color:"#fff",fontSize:13,width:"100%",boxSizing:"border-box",outline:"none"},
  scoreI:{background:"rgba(5,10,25,0.9)",border:"1px solid rgba(255,255,255,0.15)",borderRadius:10,padding:"8px 2px",color:"#e8b923",fontSize:16,fontWeight:900,width:46,textAlign:"center",outline:"none",boxSizing:"border-box"},
  pinInp:{background:"rgba(5,10,25,0.8)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:8,padding:"7px 11px",color:"#fff",fontSize:14,width:95,outline:"none"},
  btnGold:{background:"linear-gradient(135deg,#c8102e,#a00d25)",border:"none",borderRadius:10,padding:"10px 18px",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer",whiteSpace:"nowrap",boxShadow:"0 4px 14px rgba(200,16,46,0.4)"},
  btnSm:{background:"linear-gradient(135deg,#1e3a5f,#1a3255)",border:"none",borderRadius:8,padding:"6px 12px",color:"#fff",fontSize:12,cursor:"pointer"},
  btnGhost:{background:"rgba(255,255,255,0.04)",border:"1px solid rgba(255,255,255,0.12)",borderRadius:8,padding:"5px 12px",color:"#718096",fontSize:12,cursor:"pointer"},
  btnAct:{background:"rgba(15,30,60,0.8)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:8,padding:"7px 12px",color:"#a0aec0",fontSize:12,cursor:"pointer",whiteSpace:"nowrap"},
  backBtn:{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:10,padding:"7px 14px",color:"#a0aec0",cursor:"pointer",fontSize:13},
  chip:{background:"rgba(10,15,30,0.7)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:16,padding:"5px 11px",color:"#4a5568",fontSize:12,cursor:"pointer",position:"relative"},
  chipA:{background:"rgba(16,185,129,0.12)",borderColor:"rgba(16,185,129,0.4)",color:"#10b981"},
  fl:{fontSize:13,color:"#4a5568",marginBottom:5},
};


// ─── POPUP MESSI FINAL ────────────────────────────────────────────────────────
function MessiFinalPopup({onClose}){
  return(
    <div style={{position:"fixed",inset:0,zIndex:9999,background:"rgba(0,0,0,0.85)",display:"flex",alignItems:"center",justifyContent:"center",padding:"16px"}}>
      <div style={{position:"relative",maxWidth:440,width:"100%",borderRadius:18,overflow:"hidden",boxShadow:"0 20px 60px rgba(0,0,0,0.8)"}}>
        {/* Botón cerrar arriba izquierda */}
        <button onClick={onClose} style={{position:"absolute",top:12,left:12,zIndex:10,background:"rgba(0,0,0,0.6)",backdropFilter:"blur(8px)",border:"1px solid rgba(255,255,255,0.2)",borderRadius:20,padding:"7px 14px",color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer",display:"flex",alignItems:"center",gap:6}}>
          ← Ir a la app
        </button>
        {/* Imagen */}
        <img src="/images/messi_final.jpg" alt="We Meet Again" style={{width:"100%",display:"block"}}/>
      </div>
    </div>
  );
}

// ─── APP ─────────────────────────────────────────────────────────────────────
export default function App(){
  const[db,setDb]=useState({parts:[],extraM:[],results:{},rPre:null,rSpc:null,deadline:"",loaded:false});
  const[view,setView]=useState("welcome");
  const[tab,setTab]=useState("ranking");
  const[admin,setAdmin]=useState(false);
  const[pin,setPin]=useState("");
  const[showPin,setShowPin]=useState(false);
  const[selId,setSelId]=useState(null);
  const[shareP,setShareP]=useState(null);
  const[prevRanks,setPrevRanks]=useState({});
  const[myName,setMyName]=useState(()=>sessionStorage.getItem("myName")||null);
  const[showAnnouncement,setShowAnnouncement]=useState(()=>{
    const today=new Date();
    const isJune11=today.getFullYear()===2026&&today.getMonth()===5&&today.getDate()===11;
    const seenToday=sessionStorage.getItem("ann_seen_"+today.toDateString());
    return isJune11&&!seenToday;
  });
  const[showMessiPopup,setShowMessiPopup]=useState(true);

  const load=useCallback(async()=>{
    const[parts,extraM,results,rPre,rSpc,deadline]=await Promise.all([
      dbGet("parts"),dbGet("extraM"),dbGet("results"),dbGet("rPre"),dbGet("rSpc"),dbGet("deadline")
    ]);
    const dl=deadline||"";
    const newDb={parts:parts||[],extraM:extraM||[],results:results||{},rPre:rPre||null,rSpc:rSpc||null,deadline:dl,loaded:true};
    setDb(prev=>{
      if(prev.loaded&&prev.parts&&prev.parts.length>0){
        const allMprev=[...PM.map(m=>({...m,ph:"grupos",result:(prev.results||{})[m.id]||null})),...(prev.extraM||[])];
        const prevSorted=[...(prev.parts||[])].map(p=>({...p,tot:totPts(p,allMprev,prev.rPre,prev.rSpc)})).sort((a,b)=>b.tot-a.tot);
        const ranks={};prevSorted.forEach((p,i)=>{ranks[p.id]=i+1;});
        setPrevRanks(ranks);
      }
      return newDb;
    });
  },[]);

  useEffect(()=>{load();},[load]);
  useEffect(()=>{if(db.loaded&&view==="welcome"&&sessionStorage.getItem("seen"))setView("main");},[db.loaded,view]);
  useEffect(()=>{const t=setInterval(load,20000);return()=>clearInterval(t);},[load]);

  const upDb=async(key,val)=>{await dbSet(key,val);setDb(d=>({...d,[key]:val}));};
  const saveDraft=useCallback(async(name,draftData)=>{await dbSet("draft_"+name.trim().toLowerCase().replace(/\s+/g,"_"),{...draftData,_draftName:name,_savedAt:new Date().toISOString()});},[]);
  const loadDraft=useCallback(async(name)=>await dbGet("draft_"+name.trim().toLowerCase().replace(/\s+/g,"_")),[]);
  const deleteDraft=useCallback(async(name)=>await dbSet("draft_"+name.trim().toLowerCase().replace(/\s+/g,"_"),null),[]);

  const allM=[...PM.map(m=>({...m,ph:"grupos",result:(db.results||{})[m.id]||null})),...(db.extraM||[]).map(m=>({...m,result:m.result||(db.results||{})[m.id]||null}))];
  const closed=isPast(db.deadline);
  const scores=(db.parts||[]).map(p=>({...p,tot:totPts(p,allM,db.rPre,db.rSpc)})).sort((a,b)=>b.tot-a.tot);
  const bote=(db.parts||[]).length*CUOTA;
  const allLocked=(db.parts||[]).length>0&&(db.parts||[]).every(x=>x.locked);
  const jornadasCompletas=[1,2,3].filter(j=>jornadaCompleta(j,allM));

  if(!db.loaded)return(
    <div style={{...S.app,display:"flex",alignItems:"center",justifyContent:"center",height:"100vh"}}>
      <div style={{textAlign:"center"}}><div style={{fontSize:48}}>⚽</div><div style={{color:"#fbbf24",marginTop:8,fontSize:14}}>Cargando…</div></div>
    </div>
  );

  if(view==="welcome")return(
    <div style={{...S.app,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:"100vh",padding:"24px",background:"radial-gradient(ellipse at 50% 0%,#0a1628 0%,#0d1a0d 40%,#1a0800 100%)"}}>
      {showMessiPopup&&<MessiFinalPopup onClose={()=>setShowMessiPopup(false)}/>}
      <div style={{fontSize:64,marginBottom:8}}>🏆</div>
      <div style={{fontSize:32,fontWeight:900,background:"linear-gradient(135deg,#e8b923,#fff5c0,#e8b923)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",letterSpacing:3,marginBottom:4}}>SÚPER PORRA</div>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20}}>
        <div style={{height:1,width:40,background:"linear-gradient(90deg,transparent,#c8102e)"}}/>
        <div style={{fontSize:12,color:"#c8102e",letterSpacing:5,fontWeight:700}}>MUNDIAL 2026</div>
        <div style={{height:1,width:40,background:"linear-gradient(90deg,#c8102e,transparent)"}}/>
      </div>
      <CountdownWidget/>
      <div style={{width:"100%",maxWidth:400,background:"linear-gradient(135deg,rgba(10,15,30,0.95),rgba(10,20,12,0.95))",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"18px",marginBottom:24}}>
        <div style={{color:"#e8b923",fontWeight:700,fontSize:13,marginBottom:12}}>📋 Cómo funciona</div>
        {[["🎟","Inscripción","20€ — una vez confirmado, bloqueado para siempre"],["⚽","72 partidos","Pronostica todos los de fase de grupos"],["⭐","Partido fetiche","Si aciertas exacto ×5, si aciertas 1X2 ×3"],["📈","Multiplicadores","×1 grupos · ×1 dieciseisavos · hasta ×5 final"],["💰","Premios","60% · 25% · 15% del bote total"]].map(([i,t,d])=>(
          <div key={t} style={{display:"flex",gap:10,marginBottom:10}}>
            <span style={{fontSize:16,flexShrink:0}}>{i}</span>
            <div><div style={{color:"#fff",fontWeight:600,fontSize:13}}>{t}</div><div style={{color:"#4a5568",fontSize:12}}>{d}</div></div>
          </div>
        ))}
      </div>
      <button style={{...S.btnGold,width:"100%",maxWidth:400,padding:"16px",fontSize:15,borderRadius:14,letterSpacing:1}} onClick={()=>{sessionStorage.setItem("seen","1");setView("main");}}>
        ¡Entrar a la porra! →
      </button>
    </div>
  );

  if(view==="register")return <RegisterView db={db} upDb={upDb} closed={closed} saveDraft={saveDraft} loadDraft={loadDraft} deleteDraft={deleteDraft} onDone={name=>{if(name){setMyName(name);sessionStorage.setItem("myName",name);}setView("main");}} onBack={()=>setView("main")}/>;
  if(view==="predict"&&selId){
    const p=(db.parts||[]).find(x=>x.id===selId);
    if(!p){setView("main");return null;}
    return <PredictView p={p} allM={allM} db={db} onSave={async upd=>{await upDb("parts",(db.parts||[]).map(x=>x.id===upd.id?{...upd,locked:true}:x));setView("main");setTab("ranking");}} onBack={()=>setView("main")}/>;
  }
  if(view==="admin")return <AdminView db={db} upDb={upDb} allM={allM} onBack={()=>setView("main")}/>;
  if(view==="compare"&&selId){
    const p=(db.parts||[]).find(x=>x.id===selId);
    if(!p){setView("main");return null;}
    return <CompareView p={p} allM={allM} scores={scores} db={db} bote={bote} onBack={()=>setView("main")}/>;
  }

  return(
    <div style={S.app}>
      {showMessiPopup&&<MessiFinalPopup onClose={()=>{setShowMessiPopup(false);}}/>}
      {showAnnouncement&&<MundialAnnouncement onClose={()=>{sessionStorage.setItem("ann_seen_"+new Date().toDateString(),"1");setShowAnnouncement(false);}}/>}
      {shareP&&<ShareCard participant={shareP} allM={allM} onClose={()=>setShareP(null)}/>}
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:22}}>⚽</span>
            <div><div style={S.t1}>SÚPER PORRA</div><div style={S.t2}>MUNDIAL 2026</div></div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={S.bote}>💰 {bote}€</div>
            <div style={{fontSize:10,color:"#4a5568",marginTop:1}}>{bote/CUOTA||0} participante{bote/CUOTA!==1?"s":""}</div>
          </div>
        </div>
        {closed&&<div style={S.bannerR}>🔒 Pronósticos cerrados{db.deadline?" · "+fd(db.deadline):""}</div>}
        {db.deadline&&!closed&&<div style={S.bannerG}>⏰ Plazo: <strong>{fd(db.deadline)}</strong></div>}
        <div style={{borderTop:"1px solid rgba(255,255,255,0.05)",paddingTop:8,display:"flex",justifyContent:"flex-end"}}>
          {admin?(
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <span style={{fontSize:11,color:"#4a5568",background:"rgba(255,255,255,0.05)",padding:"2px 8px",borderRadius:10}}>⚙️ Admin</span>
              <button style={S.btnSm} onClick={()=>setView("admin")}>Panel</button>
              <button style={{...S.btnSm,background:"#333"}} onClick={()=>setAdmin(false)}>Salir</button>
            </div>
          ):showPin?(
            <div style={{display:"flex",gap:5,alignItems:"center"}}>
              <input type="password" placeholder="PIN" value={pin} onChange={e=>setPin(e.target.value)} onKeyDown={e=>{if(e.key==="Enter"){if(pin===PIN){setAdmin(true);setShowPin(false);setPin("");}else alert("PIN incorrecto");}}} style={S.pinInp}/>
              <button style={S.btnSm} onClick={()=>{if(pin===PIN){setAdmin(true);setShowPin(false);setPin("");}else alert("PIN incorrecto");}}>OK</button>
              <button style={{...S.btnSm,background:"#222"}} onClick={()=>setShowPin(false)}>✕</button>
            </div>
          ):(
            <button style={S.btnGhost} onClick={()=>setShowPin(true)}>🔒 Admin</button>
          )}
        </div>
      </div>

      <div style={S.tabs}>
        {[["ranking","📊 Ranking"],["pronos","🎯 Pronos"],["partidos","⚽ Partidos"],...(allM.some(m=>m.ph!=="grupos")?[["elim","🏟 Elim."]]:[]),...(jornadasCompletas.length>0?[["jornadas","📈 Jornadas"]]:[]),["chat","💬 Chat"],["reglas","📋 Reglas"]].map(([id,l])=>(
          <button key={id} style={{...S.tab,...(tab===id?S.tabA:{})}} onClick={()=>setTab(id)}>{l}</button>
        ))}
      </div>

      {tab!=="chat"&&tab!=="jornadas"&&tab!=="elim"&&<div style={S.content}>
        {tab==="ranking"&&(
          <div>
            <CountdownWidget/>
            {myName&&scores.length>0&&<MiPosicion scores={scores} myName={myName} bote={bote}/>}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
              {[{pct:.6,l:"1er Premio",col:"#e8b923",ico:"🥇",bg:"linear-gradient(135deg,#422006,#78350f)",sh:"0 0 20px rgba(232,185,35,0.25)"},{pct:.25,l:"2º Premio",col:"#cbd5e1",ico:"🥈",bg:"linear-gradient(135deg,#1a1a2e,#16213e)",sh:"none"},{pct:.15,l:"3er Premio",col:"#cd7f32",ico:"🥉",bg:"linear-gradient(135deg,#1c0a00,#3b1f00)",sh:"none"}].map((p,i)=>(
                <div key={i} style={{background:p.bg,border:"1px solid "+p.col+"44",borderRadius:12,padding:"12px 8px",textAlign:"center",boxShadow:p.sh}}>
                  <div style={{fontSize:22}}>{p.ico}</div>
                  <div style={{color:p.col,fontWeight:900,fontSize:20,marginTop:2}}>{Math.round(bote*p.pct)}€</div>
                  <div style={{fontSize:10,color:"rgba(255,255,255,0.4)",marginTop:2}}>{p.l}</div>
                </div>
              ))}
            </div>
            {!allLocked&&scores.length>0&&scores.length>1&&<div style={{background:"rgba(37,99,235,0.1)",border:"1px solid rgba(37,99,235,0.25)",borderRadius:10,padding:"8px 14px",fontSize:12,color:"#93c5fd",marginBottom:10,textAlign:"center"}}>🔒 Pronósticos privados hasta que todos confirmen</div>}
            {scores.length===0?<div style={S.empty}><div style={{fontSize:36}}>📊</div><div style={{marginTop:8}}>El ranking aparecerá cuando haya participantes</div></div>:
              scores.map((p,i)=>{
                const prevPos=prevRanks[p.id],curr=i+1,moved=prevPos&&prevPos!==curr?prevPos-curr:0;
                const isMe=myName&&p.name===myName;
                return(
                  <div key={p.id} style={{...S.row,...(i===scores.length-1&&scores.length>1?{borderColor:"rgba(127,29,29,0.3)",background:"rgba(20,5,5,0.5)"}:{}),...(isMe?{border:"1px solid rgba(232,185,35,0.4)",background:"rgba(232,185,35,0.06)"}:{})}}>
                    <Pos n={curr}/>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",alignItems:"center",gap:6}}>
                        <span style={{fontWeight:700,color:isMe?"#e8b923":"#fff"}}>{p.name}</span>
                        {isMe&&<span style={{fontSize:9,color:"#e8b923",background:"rgba(232,185,35,0.1)",border:"1px solid rgba(232,185,35,0.3)",borderRadius:8,padding:"1px 5px"}}>TÚ</span>}
                        {moved!==0&&<span style={{fontSize:11,fontWeight:700,color:moved>0?"#10b981":"#ef4444"}}>{moved>0?"▲"+moved:"▼"+Math.abs(moved)}</span>}
                      </div>
                      <div style={{display:"flex",gap:6,marginTop:2,flexWrap:"wrap"}}>
                        {allLocked&&p.pre?.campeon&&<div style={{fontSize:11,color:"#2d3748"}}>🏆 {p.pre.campeon.split(" ")[0]}</div>}
                        {p.fetiche&&<div style={{fontSize:11,color:"#e8b923"}}>⭐ {(allM.find(m=>m.id===p.fetiche)||{l:"?"}).l?.split(" ")[0]}</div>}
                      </div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      <div style={{fontWeight:900,fontSize:20,color:i===0?"#e8b923":i<3?"#10b981":"#bbb"}}>{p.tot}<span style={{fontSize:11,color:"#2d3748",fontWeight:400}}> pts</span></div>
                      {i<3&&bote>0&&<div style={{fontSize:11,color:"#10b981"}}>+{Math.round(bote*[.6,.25,.15][i])}€</div>}
                    </div>
                    <div style={{display:"flex",flexDirection:"column",gap:4,marginLeft:4}}>
                      <button style={{...S.btnGhost,fontSize:11,padding:"3px 7px"}} onClick={()=>{setSelId(p.id);setView("compare");}}>📊</button>
                      <button style={{...S.btnGhost,fontSize:11,padding:"3px 7px"}} onClick={()=>setShareP(p)}>📤</button>
                    </div>
                  </div>
                );
              })
            }
            {scores.length>1&&<div style={{background:"linear-gradient(135deg,rgba(127,29,29,0.3),rgba(100,20,20,0.2))",border:"1px dashed rgba(220,38,38,0.4)",borderRadius:12,padding:"12px 14px",marginTop:10,display:"flex",gap:10,alignItems:"center"}}>
              <span style={{fontSize:20}}>😎</span><div><div style={{fontWeight:700,color:"#ff6b35"}}>Farolillo Rojo</div><div style={{fontSize:12,color:"#4a5568"}}>¡Ronda épica para {scores[scores.length-1].name}!</div></div>
            </div>}
          </div>
        )}
        {tab==="pronos"&&<PronosTab db={db} scores={scores} closed={closed} allM={allM} onRegister={()=>setView("register")} onPredict={id=>{setSelId(id);setView("predict");setMyName((db.parts||[]).find(x=>x.id===id)?.name||myName);}} onShare={p=>setShareP(p)}/>}
        {tab==="partidos"&&<PartidosTab allM={allM}/>}
        {tab==="reglas"&&<ReglasTab/>}
      </div>}
      {tab==="elim"&&<EliminatoriaTab db={db} upDb={upDb} allM={allM}/>}
      {tab==="jornadas"&&<JornadasTab jornadasCompletas={jornadasCompletas} allM={allM} parts={db.parts||[]}/>}
      {tab==="chat"&&<ChatTab myName={myName} parts={db.parts||[]}/>}
    </div>
  );
}
