import { useState, useEffect, useCallback } from "react";
import { dbGet, dbSet } from "./supabase";
import * as XLSX from "xlsx";


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
  {id:"a1",l:"México 🇲🇽",v:"Sudáfrica 🇿🇦",g:"A",d:"2026-06-11"},
  {id:"a2",l:"Corea del Sur 🇰🇷",v:"Chequia 🇨🇿",g:"A",d:"2026-06-11"},
  {id:"a3",l:"Chequia 🇨🇿",v:"Sudáfrica 🇿🇦",g:"A",d:"2026-06-18"},
  {id:"a4",l:"México 🇲🇽",v:"Corea del Sur 🇰🇷",g:"A",d:"2026-06-18"},
  {id:"a5",l:"Chequia 🇨🇿",v:"México 🇲🇽",g:"A",d:"2026-06-24"},
  {id:"a6",l:"Sudáfrica 🇿🇦",v:"Corea del Sur 🇰🇷",g:"A",d:"2026-06-24"},
  {id:"b1",l:"Canadá 🇨🇦",v:"Bosnia 🇧🇦",g:"B",d:"2026-06-12"},
  {id:"b2",l:"Qatar 🇶🇦",v:"Suiza 🇨🇭",g:"B",d:"2026-06-13"},
  {id:"b3",l:"Suiza 🇨🇭",v:"Bosnia 🇧🇦",g:"B",d:"2026-06-18"},
  {id:"b4",l:"Canadá 🇨🇦",v:"Qatar 🇶🇦",g:"B",d:"2026-06-18"},
  {id:"b5",l:"Suiza 🇨🇭",v:"Canadá 🇨🇦",g:"B",d:"2026-06-24"},
  {id:"b6",l:"Bosnia 🇧🇦",v:"Qatar 🇶🇦",g:"B",d:"2026-06-24"},
  {id:"c1",l:"Haití 🇭🇹",v:"Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿",g:"C",d:"2026-06-13"},
  {id:"c2",l:"Brasil 🇧🇷",v:"Marruecos 🇲🇦",g:"C",d:"2026-06-13"},
  {id:"c3",l:"Brasil 🇧🇷",v:"Haití 🇭🇹",g:"C",d:"2026-06-19"},
  {id:"c4",l:"Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿",v:"Marruecos 🇲🇦",g:"C",d:"2026-06-19"},
  {id:"c5",l:"Escocia 🏴󠁧󠁢󠁳󠁣󠁴󠁿",v:"Brasil 🇧🇷",g:"C",d:"2026-06-24"},
  {id:"c6",l:"Marruecos 🇲🇦",v:"Haití 🇭🇹",g:"C",d:"2026-06-24"},
  {id:"d1",l:"EE.UU. 🇺🇸",v:"Paraguay 🇵🇾",g:"D",d:"2026-06-12"},
  {id:"d2",l:"Australia 🇦🇺",v:"Turquía 🇹🇷",g:"D",d:"2026-06-14"},
  {id:"d3",l:"Turquía 🇹🇷",v:"Paraguay 🇵🇾",g:"D",d:"2026-06-19"},
  {id:"d4",l:"EE.UU. 🇺🇸",v:"Australia 🇦🇺",g:"D",d:"2026-06-19"},
  {id:"d5",l:"Turquía 🇹🇷",v:"EE.UU. 🇺🇸",g:"D",d:"2026-06-25"},
  {id:"d6",l:"Paraguay 🇵🇾",v:"Australia 🇦🇺",g:"D",d:"2026-06-25"},
  {id:"e1",l:"Costa de Marfil 🇨🇮",v:"Ecuador 🇪🇨",g:"E",d:"2026-06-14"},
  {id:"e2",l:"Alemania 🇩🇪",v:"Curazao 🇨🇼",g:"E",d:"2026-06-14"},
  {id:"e3",l:"Alemania 🇩🇪",v:"Costa de Marfil 🇨🇮",g:"E",d:"2026-06-20"},
  {id:"e4",l:"Ecuador 🇪🇨",v:"Curazao 🇨🇼",g:"E",d:"2026-06-20"},
  {id:"e5",l:"Curazao 🇨🇼",v:"Costa de Marfil 🇨🇮",g:"E",d:"2026-06-25"},
  {id:"e6",l:"Ecuador 🇪🇨",v:"Alemania 🇩🇪",g:"E",d:"2026-06-25"},
  {id:"f1",l:"Países Bajos 🇳🇱",v:"Japón 🇯🇵",g:"F",d:"2026-06-14"},
  {id:"f2",l:"Suecia 🇸🇪",v:"Túnez 🇹🇳",g:"F",d:"2026-06-14"},
  {id:"f3",l:"Países Bajos 🇳🇱",v:"Suecia 🇸🇪",g:"F",d:"2026-06-20"},
  {id:"f4",l:"Túnez 🇹🇳",v:"Japón 🇯🇵",g:"F",d:"2026-06-21"},
  {id:"f5",l:"Japón 🇯🇵",v:"Suecia 🇸🇪",g:"F",d:"2026-06-25"},
  {id:"f6",l:"Túnez 🇹🇳",v:"Países Bajos 🇳🇱",g:"F",d:"2026-06-25"},
  {id:"g1",l:"Irán 🇮🇷",v:"Nueva Zelanda 🇳🇿",g:"G",d:"2026-06-15"},
  {id:"g2",l:"Bélgica 🇧🇪",v:"Egipto 🇪🇬",g:"G",d:"2026-06-15"},
  {id:"g3",l:"Bélgica 🇧🇪",v:"Irán 🇮🇷",g:"G",d:"2026-06-21"},
  {id:"g4",l:"Egipto 🇪🇬",v:"Nueva Zelanda 🇳🇿",g:"G",d:"2026-06-21"},
  {id:"g5",l:"Irán 🇮🇷",v:"Egipto 🇪🇬",g:"G",d:"2026-06-26"},
  {id:"g6",l:"Nueva Zelanda 🇳🇿",v:"Bélgica 🇧🇪",g:"G",d:"2026-06-26"},
  {id:"h1",l:"Arabia Saudí 🇸🇦",v:"Uruguay 🇺🇾",g:"H",d:"2026-06-15"},
  {id:"h2",l:"España 🇪🇸",v:"Cabo Verde 🇨🇻",g:"H",d:"2026-06-15"},
  {id:"h3",l:"Uruguay 🇺🇾",v:"Cabo Verde 🇨🇻",g:"H",d:"2026-06-21"},
  {id:"h4",l:"España 🇪🇸",v:"Arabia Saudí 🇸🇦",g:"H",d:"2026-06-21"},
  {id:"h5",l:"España 🇪🇸",v:"Uruguay 🇺🇾",g:"H",d:"2026-06-26"},
  {id:"h6",l:"Cabo Verde 🇨🇻",v:"Arabia Saudí 🇸🇦",g:"H",d:"2026-06-26"},
  {id:"i1",l:"Francia 🇫🇷",v:"Senegal 🇸🇳",g:"I",d:"2026-06-16"},
  {id:"i2",l:"Iraq 🇮🇶",v:"Noruega 🇳🇴",g:"I",d:"2026-06-16"},
  {id:"i3",l:"Noruega 🇳🇴",v:"Senegal 🇸🇳",g:"I",d:"2026-06-22"},
  {id:"i4",l:"Francia 🇫🇷",v:"Iraq 🇮🇶",g:"I",d:"2026-06-22"},
  {id:"i5",l:"Francia 🇫🇷",v:"Noruega 🇳🇴",g:"I",d:"2026-06-26"},
  {id:"i6",l:"Senegal 🇸🇳",v:"Iraq 🇮🇶",g:"I",d:"2026-06-26"},
  {id:"j1",l:"Argentina 🇦🇷",v:"Argelia 🇩🇿",g:"J",d:"2026-06-16"},
  {id:"j2",l:"Austria 🇦🇹",v:"Jordania 🇯🇴",g:"J",d:"2026-06-17"},
  {id:"j3",l:"Argentina 🇦🇷",v:"Austria 🇦🇹",g:"J",d:"2026-06-22"},
  {id:"j4",l:"Argelia 🇩🇿",v:"Jordania 🇯🇴",g:"J",d:"2026-06-22"},
  {id:"j5",l:"Argentina 🇦🇷",v:"Jordania 🇯🇴",g:"J",d:"2026-06-26"},
  {id:"j6",l:"Argelia 🇩🇿",v:"Austria 🇦🇹",g:"J",d:"2026-06-26"},
  {id:"k1",l:"Portugal 🇵🇹",v:"RD Congo 🇨🇩",g:"K",d:"2026-06-17"},
  {id:"k2",l:"Uzbekistán 🇺🇿",v:"Colombia 🇨🇴",g:"K",d:"2026-06-17"},
  {id:"k3",l:"Portugal 🇵🇹",v:"Uzbekistán 🇺🇿",g:"K",d:"2026-06-23"},
  {id:"k4",l:"Colombia 🇨🇴",v:"RD Congo 🇨🇩",g:"K",d:"2026-06-23"},
  {id:"k5",l:"Colombia 🇨🇴",v:"Portugal 🇵🇹",g:"K",d:"2026-06-27"},
  {id:"k6",l:"RD Congo 🇨🇩",v:"Uzbekistán 🇺🇿",g:"K",d:"2026-06-27"},
  {id:"l1",l:"Ghana 🇬🇭",v:"Panamá 🇵🇦",g:"L",d:"2026-06-17"},
  {id:"l2",l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"Croacia 🇭🇷",g:"L",d:"2026-06-17"},
  {id:"l3",l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"Ghana 🇬🇭",g:"L",d:"2026-06-23"},
  {id:"l4",l:"Croacia 🇭🇷",v:"Panamá 🇵🇦",g:"L",d:"2026-06-23"},
  {id:"l5",l:"Inglaterra 🏴󠁧󠁢󠁥󠁮󠁧󠁿",v:"Panamá 🇵🇦",g:"L",d:"2026-06-27"},
  {id:"l6",l:"Croacia 🇭🇷",v:"Ghana 🇬🇭",g:"L",d:"2026-06-27"},
];

const MULT={grupos:1,octavos:1.5,cuartos:2,semifinales:3,final:5};
const PHL={grupos:"Grupos",octavos:"Octavos",cuartos:"Cuartos",semifinales:"Semis",final:"Final"};
const CUOTA=20;
const PIN="4805";
const PRIMER_PARTIDO=new Date("2026-06-11T21:00:00+02:00"); // Madrid time

// ── SCORING ───────────────────────────────────────────────────────────────────
// fetiche: si acierta exacto x3, si acierta 1X2 x2, si falla x1 (normal)
function mPts(pred,res,ph,isFetiche){
  if(!pred||!res)return 0;
  const m=MULT[ph]||1,rl=+res.l,rv=+res.v,pl=+pred.l,pv=+pred.v;
  if([rl,rv,pl,pv].some(isNaN))return 0;
  let b=0;
  const exacto=pl===rl&&pv===rv;
  if(exacto)b+=5;
  else{const rw=rl>rv?"L":rv>rl?"V":"E",pw=pl>pv?"L":pv>pl?"V":"E";if(rw===pw)b+=2;}
  if(pl===rl)b+=1;if(pv===rv)b+=1;
  let pts=Math.round(b*m*10)/10;
  if(isFetiche){
    if(exacto)pts=Math.round(pts*5*10)/10;
    else if(b>=2)pts=Math.round(pts*3*10)/10;
  }
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
  const fetiche=part.fetiche||null;
  (allM||[]).forEach(m=>{t+=mPts((part.mp||{})[m.id],m.result,m.ph||"grupos",m.id===fetiche);});
  return Math.round(t*10)/10;
}

// ── STORAGE ───────────────────────────────────────────────────────────────────


function fd(s){if(!s)return"";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;}
function isPast(d){return d?new Date()>new Date(d+"T23:59:59"):false;}

// ── COUNTDOWN ────────────────────────────────────────────────────────────────
function useCountdown(target){
  const[t,setT]=useState(Math.max(0,target-Date.now()));
  useEffect(()=>{
    if(t<=0)return;
    const i=setInterval(()=>setT(Math.max(0,target-Date.now())),1000);
    return()=>clearInterval(i);
  },[target,t]);
  const s=Math.floor(t/1000);
  const days=Math.floor(s/86400);
  const hrs=Math.floor((s%86400)/3600);
  const mins=Math.floor((s%3600)/60);
  const secs=s%60;
  return{days,hrs,mins,secs,done:t<=0};
}

// ── EXPORT ───────────────────────────────────────────────────────────────────
function exportCSV(parts,allM,rPre,rSpc){
  const wb=XLSX.utils.book_new();
  const sorted=[...parts].sort((a,b)=>totPts(b,allM,rPre,rSpc)-totPts(a,allM,rPre,rSpc));
  const PHL_FULL={grupos:"Fase de Grupos",octavos:"Octavos",cuartos:"Cuartos",semifinales:"Semifinales",final:"Final"};

  // ── HOJA RANKING GENERAL ──────────────────────────────────────────────────
  const rankRows=[["#","Jugador","Puntos Totales","Puntos Pre-Mundial","Puntos Partidos","Puntos Retos","Premio estimado (€)","Fetiche"]];
  const bote=parts.length*20;
  sorted.forEach((p,i)=>{
    const fetM=allM.find(m=>m.id===p.fetiche);
    const pPre=prePts(p.pre||{},rPre);
    const pSpc=spcPts(p.spc||{},rSpc);
    let pM=0;allM.forEach(m=>{pM+=mPts((p.mp||{})[m.id],m.result,m.ph||"grupos",m.id===p.fetiche);});
    const premio=i===0?Math.round(bote*0.6):i===1?Math.round(bote*0.25):i===2?Math.round(bote*0.15):0;
    rankRows.push([i+1,p.name,totPts(p,allM,rPre,rSpc),pPre,Math.round(pM*10)/10,pSpc,premio||"-",fetM?`${fetM.l} vs ${fetM.v}`:"-"]);
  });
  const wsRank=XLSX.utils.aoa_to_sheet(rankRows);
  wsRank["!cols"]=[{wch:4},{wch:18},{wch:14},{wch:16},{wch:15},{wch:13},{wch:20},{wch:30}];
  XLSX.utils.book_append_sheet(wb,wsRank,"🏆 Ranking");

  // ── HOJA POR JUGADOR ──────────────────────────────────────────────────────
  sorted.forEach(p=>{
    const rows=[];
    const tot=totPts(p,allM,rPre,rSpc);
    const fetM=allM.find(m=>m.id===p.fetiche);

    // Header jugador
    rows.push([`PRONÓSTICOS DE ${p.name.toUpperCase()}`]);
    rows.push([`Total: ${tot} pts`,`Registrado: ${p.registeredAt?new Date(p.registeredAt).toLocaleDateString("es-ES"):"-"}`]);
    rows.push([]);

    // PRE-MUNDIAL
    rows.push(["── PRE-MUNDIAL ──────────────────────────────"]);
    rows.push(["Campo","Tu pronóstico","Resultado real","Puntos"]);
    const preItems=[
      ["Campeón del mundo (30pts)",p.pre?.campeon,rPre?.campeon,p.pre?.campeon&&rPre?.campeon&&p.pre.campeon===rPre.campeon?30:0],
      ["Subcampeón (20pts)",p.pre?.subcampeon,rPre?.subcampeon,p.pre?.subcampeon&&rPre?.subcampeon&&p.pre.subcampeon===rPre.subcampeon?20:0],
      ["3er puesto (15pts)",p.pre?.tercero,rPre?.tercero,p.pre?.tercero&&rPre?.tercero&&p.pre.tercero===rPre.tercero?15:0],
      ["Máx. goleador (20pts)",p.pre?.goleador,rPre?.goleador,p.pre?.goleador&&rPre?.goleador&&p.pre.goleador===rPre.goleador?20:0],
      ["MVP (15pts)",p.pre?.mvp,rPre?.mvp,p.pre?.mvp&&rPre?.mvp&&p.pre.mvp===rPre.mvp?15:0],
    ];
    (p.pre?.semis||[]).forEach((s,i)=>{
      const hit=s&&(rPre?.semis||[]).includes(s);
      preItems.push([`Semifinalista ${i+1} (10pts)`,s,"-",hit?10:0]);
    });
    preItems.forEach(([campo,pred,real,pts])=>{
      rows.push([campo,pred||"-",real||"(pendiente)",rPre?pts:"-"]);
    });
    rows.push([]);

    // PARTIDO FETICHE
    rows.push(["── PARTIDO FETICHE ──────────────────────────"]);
    rows.push(["Partido elegido","Tu pronóstico","Resultado real","Puntos","Multiplicador"]);
    if(fetM){
      const pred=(p.mp||{})[fetM.id];
      const res=fetM.result;
      const pts=res&&pred?mPts(pred,res,fetM.ph||"grupos",true):"-";
      rows.push([`${fetM.l} vs ${fetM.v}`,pred?`${pred.l}-${pred.v}`:"-",res?`${res.l}-${res.v}`:"(pendiente)",res?pts:"-","×5 exacto / ×3 1X2"]);
    } else {
      rows.push(["No seleccionado","-","-","-","-"]);
    }
    rows.push([]);

    // RETOS ESPECIALES
    rows.push(["── RETOS ESPECIALES ─────────────────────────"]);
    rows.push(["Reto","Tu pronóstico","Resultado real","Puntos"]);
    const retoItems=[
      ["Primer expulsado (10pts)",p.spc?.expulsado,rSpc?.expulsado,p.spc?.expulsado&&rSpc?.expulsado&&p.spc.expulsado===rSpc.expulsado?10:0],
      ["Primer hat-trick (10pts)",p.spc?.hattrick,rSpc?.hattrick,p.spc?.hattrick&&rSpc?.hattrick&&p.spc.hattrick===rSpc.hattrick?10:0],
      ["Equipo revelación (15pts)",p.spc?.revelacion,rSpc?.revelacion,p.spc?.revelacion&&rSpc?.revelacion&&p.spc.revelacion===rSpc.revelacion?15:0],
      ["Mayor goleada (10pts)",p.spc?.goleada,rSpc?.goleada,p.spc?.goleada&&rSpc?.goleada&&p.spc.goleada===rSpc.goleada?10:0],
    ];
    retoItems.forEach(([reto,pred,real,pts])=>{
      rows.push([reto,pred||"-",real||"(pendiente)",rSpc?pts:"-"]);
    });
    rows.push([]);

    // PARTIDOS
    rows.push(["── PRONÓSTICOS DE PARTIDOS ──────────────────"]);
    rows.push(["Fecha","Fase","Partido","Tu pronóstico","Resultado real","Puntos","Fetiche"]);
    const grupos=["A","B","C","D","E","F","G","H","I","J","K","L"];
    const fases=["grupos","octavos","cuartos","semifinales","final"];
    fases.forEach(fase=>{
      const matchesFase=fase==="grupos"
        ?grupos.flatMap(g=>allM.filter(m=>m.g===g&&m.ph==="grupos"))
        :allM.filter(m=>m.ph===fase);
      if(matchesFase.length===0)return;
      matchesFase.forEach(m=>{
        const pred=(p.mp||{})[m.id];
        const res=m.result;
        const isFet=m.id===p.fetiche;
        const pts=res&&pred?mPts(pred,res,m.ph||"grupos",isFet):"-";
        rows.push([
          m.d||m.date||"-",
          PHL_FULL[m.ph]||m.ph,
          `${m.l} vs ${m.v}`,
          pred?`${pred.l}-${pred.v}`:"-",
          res?`${res.l}-${res.v}`:"(pendiente)",
          res?pts:"-",
          isFet?"⭐ SÍ":""
        ]);
      });
    });

    // Nombre de hoja max 31 chars, sin caracteres especiales
    const sheetName=p.name.replace(/[\/*?:\[\]]/g,"").substring(0,28);
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"]=[{wch:12},{wch:18},{wch:30},{wch:16},{wch:16},{wch:8},{wch:10}];
    XLSX.utils.book_append_sheet(wb,ws,sheetName);
  });

  // Descargar
  XLSX.writeFile(wb,"porra-mundial-2026.xlsx");
}

// ── HELPERS UI ────────────────────────────────────────────────────────────────
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
    <div style={{background:"#0a0a10",borderRadius:8,padding:"6px 8px",border:"1px solid "+col+"44"}}>
      <div style={{fontSize:10,color:"#555",marginBottom:2}}>{label}</div>
      <div style={{fontSize:12,fontWeight:700,color:col}}>{done}/{total}</div>
      <div style={{height:3,background:"#111",borderRadius:2,marginTop:3}}>
        <div style={{height:"100%",width:pct+"%",background:col,borderRadius:2}}/>
      </div>
    </div>
  );
}
function Pos({n}){
  const bg=n===1?"linear-gradient(135deg,#FFD700,#f59e0b)":n===2?"linear-gradient(135deg,#e2e8f0,#94a3b8)":n===3?"linear-gradient(135deg,#cd7f32,#92400e)":"#1a1a1a";
  const col=n<=2?"#000":"#fff";
  return <div style={{width:30,height:30,borderRadius:"50%",background:bg,color:col,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:13,flexShrink:0}}>{n}</div>;
}

// ── COUNTDOWN WIDGET ─────────────────────────────────────────────────────────
function CountdownWidget(){
  const{days,hrs,mins,secs,done}=useCountdown(PRIMER_PARTIDO.getTime());
  if(done)return(
    <div style={{background:"linear-gradient(135deg,#052e16,#065f46)",border:"1px solid #10b98144",borderRadius:14,padding:"14px 16px",marginBottom:14,textAlign:"center",boxShadow:"0 0 20px rgba(16,185,129,0.15)"}}>
      <div style={{fontSize:28,marginBottom:4}}>⚽</div>
      <div style={{color:"#10b981",fontWeight:800,fontSize:15,letterSpacing:1}}>¡EL MUNDIAL HA COMENZADO!</div>
    </div>
  );
  const pad=n=>String(n).padStart(2,"0");
  return(
    <div style={{background:"linear-gradient(135deg,#0a0f1e,#0d1a0d)",border:"1px solid rgba(232,185,35,0.2)",borderRadius:14,padding:"14px 16px",marginBottom:14,boxShadow:"0 4px 20px rgba(0,0,0,0.4)"}}>
      <div style={{textAlign:"center",fontSize:11,color:"#4a5568",marginBottom:10,letterSpacing:2,textTransform:"uppercase",fontWeight:600}}>⚽ México vs Sudáfrica · 11 Jun · 21:00h</div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6}}>
        {[[days,"DÍAS"],[hrs,"HORAS"],[mins,"MIN"],[pad(secs),"SEG"]].map(([v,l])=>(
          <div key={l} style={{background:"rgba(0,0,0,0.5)",borderRadius:10,padding:"10px 4px",textAlign:"center",border:"1px solid rgba(232,185,35,0.15)",boxShadow:"inset 0 1px 0 rgba(255,255,255,0.05)"}}>
            <div style={{fontWeight:900,fontSize:26,background:"linear-gradient(180deg,#fff5c0,#e8b923)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",lineHeight:1}}>{v}</div>
            <div style={{fontSize:9,color:"#4a5568",marginTop:3,letterSpacing:2,fontWeight:600}}>{l}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── MI POSICIÓN ──────────────────────────────────────────────────────────────
function MiPosicion({scores,myName,bote}){
  const idx=scores.findIndex(p=>p.name===myName);
  if(idx===-1)return null;
  const p=scores[idx];
  const pos=idx+1;
  const next=scores[idx-1];
  const diff=next?Math.round((next.tot-p.tot)*10)/10:null;
  const posCol=pos===1?"#e8b923":pos===2?"#10b981":pos===3?"#10b981":"#94a3b8";
  const premio=pos<=3?Math.round(bote*[.6,.25,.15][pos-1]):0;
  return(
    <div style={{background:"linear-gradient(135deg,rgba(232,185,35,0.1),rgba(200,16,46,0.08))",border:"1px solid rgba(232,185,35,0.25)",borderRadius:14,padding:"12px 14px",marginBottom:14,boxShadow:"0 4px 16px rgba(0,0,0,0.3)"}}>
      <div style={{fontSize:10,color:"#4a5568",letterSpacing:2,fontWeight:700,textTransform:"uppercase",marginBottom:8}}>📍 Tu posición</div>
      <div style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{width:44,height:44,borderRadius:"50%",background:pos===1?"linear-gradient(135deg,#e8b923,#f59e0b)":pos<=3?"linear-gradient(135deg,#10b981,#059669)":"rgba(255,255,255,0.08)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:20,color:pos<=2?"#000":"#fff",flexShrink:0,boxShadow:pos===1?"0 0 16px rgba(232,185,35,0.4)":"none"}}>
          {pos}
        </div>
        <div style={{flex:1}}>
          <div style={{fontWeight:800,fontSize:17,color:posCol}}>{p.tot} <span style={{fontSize:12,color:"#4a5568",fontWeight:400}}>pts</span></div>
          {diff!==null&&diff>0?(
            <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>A <span style={{color:"#ef4444",fontWeight:700}}>{diff} pts</span> del {pos-1}º ({next.name.split(" ")[0]})</div>
          ):(
            <div style={{fontSize:12,color:"#10b981",marginTop:2}}>🏆 ¡Líder de la porra!</div>
          )}
        </div>
        {premio>0&&(
          <div style={{textAlign:"center",background:"rgba(16,185,129,0.1)",border:"1px solid rgba(16,185,129,0.2)",borderRadius:10,padding:"6px 10px"}}>
            <div style={{fontSize:11,color:"#4a5568"}}>Premio est.</div>
            <div style={{fontSize:18,fontWeight:900,color:"#10b981"}}>+{premio}€</div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── APP ───────────────────────────────────────────────────────────────────────
export default function App(){
  const[db,setDb]=useState({parts:[],extraM:[],results:{},rPre:null,rSpc:null,deadline:"",loaded:false});
  const[view,setView]=useState("welcome");
  const[tab,setTab]=useState("ranking");
  const[admin,setAdmin]=useState(false);
  const[pin,setPin]=useState("");
  const[showPin,setShowPin]=useState(false);
  const[selId,setSelId]=useState(null);
  const[prevRanks,setPrevRanks]=useState({}); // id -> posicion anterior
  const[myName,setMyName]=useState(()=>sessionStorage.getItem("myName")||null);

  const load=useCallback(async()=>{
    const[parts,extraM,results,rPre,rSpc,deadline]=await Promise.all([
      dbGet("parts"),dbGet("extraM"),dbGet("results"),dbGet("rPre"),dbGet("rSpc"),dbGet("deadline")
    ]);
    const dl=deadline||"";
    const newDb={parts:parts||[],extraM:extraM||[],results:results||{},rPre:rPre||null,rSpc:rSpc||null,deadline:dl,loaded:true};
    setDb(prev=>{
      // Track rank changes: compare previous scores with new ones
      if(prev.loaded&&prev.parts&&prev.parts.length>0){
        const allMprev=[...PM.map(m=>({...m,ph:"grupos",result:(prev.results||{})[m.id]||null})),...(prev.extraM||[])];
        const prevSorted=[...(prev.parts||[])].map(p=>({...p,tot:totPts(p,allMprev,prev.rPre,prev.rSpc)})).sort((a,b)=>b.tot-a.tot);
        const ranks={};
        prevSorted.forEach((p,i)=>{ranks[p.id]=i+1;});
        setPrevRanks(ranks);
      }
      return newDb;
    });
  },[]);

  useEffect(()=>{load();},[load]);
  useEffect(()=>{
    if(db.loaded&&view==="welcome"&&sessionStorage.getItem("seen"))setView("main");
  },[db.loaded,view]);
  useEffect(()=>{const t=setInterval(load,20000);return()=>clearInterval(t);},[load]);

  const upDb=async(key,val)=>{
    await dbSet(key,val);
    setDb(d=>({...d,[key]:val}));
  };
  // ── BORRADOR AUTO-SAVE ──────────────────────────────────────────────────────
  const saveDraft = useCallback(async(name, draftData) => {
    const key = "draft_" + name.trim().toLowerCase().replace(/\s+/g,"_");
    await dbSet(key, {...draftData, _draftName: name, _savedAt: new Date().toISOString()});
  }, []);

  const loadDraft = useCallback(async(name) => {
    const key = "draft_" + name.trim().toLowerCase().replace(/\s+/g,"_");
    return await dbGet(key);
  }, []);

  const deleteDraft = useCallback(async(name) => {
    const key = "draft_" + name.trim().toLowerCase().replace(/\s+/g,"_");
    await dbSet(key, null);
  }, []);



  const allM=[...PM.map(m=>({...m,ph:"grupos",result:(db.results||{})[m.id]||null})),...(db.extraM||[])];
  const closed=isPast(db.deadline);
  const scores=(db.parts||[]).map(p=>({...p,tot:totPts(p,allM,db.rPre,db.rSpc)})).sort((a,b)=>b.tot-a.tot);
  const bote=(db.parts||[]).length*CUOTA;
  const allLocked=(db.parts||[]).length>0&&(db.parts||[]).every(x=>x.locked);

  if(!db.loaded)return(
    <div style={{...S.app,display:"flex",alignItems:"center",justifyContent:"center",height:"100vh"}}>
      <div style={{textAlign:"center"}}><div style={{fontSize:48}}>⚽</div><div style={{color:"#fbbf24",marginTop:8,fontSize:14}}>Cargando…</div></div>
    </div>
  );

  if(view==="welcome")return(
    <div style={{...S.app,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:"100vh",padding:"24px",background:"radial-gradient(ellipse at 50% 0%,#0a1628 0%,#0d1a0d 40%,#1a0800 100%)"}}>
      <div style={{fontSize:64,marginBottom:8}}>🏆</div>
      <div style={{fontSize:32,fontWeight:900,background:"linear-gradient(135deg,#e8b923,#fff5c0,#e8b923)",WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",letterSpacing:3,marginBottom:4}}>SÚPER PORRA</div>
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20}}>
        <div style={{height:1,width:40,background:"linear-gradient(90deg,transparent,#c8102e)"}}/>
        <div style={{fontSize:12,color:"#c8102e",letterSpacing:5,fontWeight:700}}>MUNDIAL 2026</div>
        <div style={{height:1,width:40,background:"linear-gradient(90deg,#c8102e,transparent)"}}/>
      </div>
      <CountdownWidget/>
      <div style={{width:"100%",maxWidth:400,background:"linear-gradient(135deg,rgba(10,15,30,0.95),rgba(10,20,12,0.95))",border:"1px solid rgba(255,255,255,0.07)",borderRadius:14,padding:"18px",marginBottom:24,backdropFilter:"blur(10px)",boxShadow:"0 8px 32px rgba(0,0,0,0.4)"}}>
        <div style={{color:"#fbbf24",fontWeight:700,fontSize:13,marginBottom:12}}>📋 Cómo funciona</div>
        {[["🎟","Inscripción","20€ — una vez confirmado, bloqueado para siempre"],["⚽","72 partidos","Pronostica todos los partidos de la fase de grupos"],["⭐","Partido fetiche","Márcalo: si aciertas exacto ×5, si aciertas 1X2 ×3"],["📈","Multiplicadores","×1 grupos · ×1.5 octavos · hasta ×5 final"],["💰","Premios","60% · 25% · 15% del bote total"]].map(([i,t,d])=>(
          <div key={t} style={{display:"flex",gap:10,marginBottom:10}}>
            <span style={{fontSize:16,flexShrink:0}}>{i}</span>
            <div><div style={{color:"#fff",fontWeight:600,fontSize:13}}>{t}</div><div style={{color:"#555",fontSize:12}}>{d}</div></div>
          </div>
        ))}
      </div>
      <button style={{...S.btnGold,width:"100%",maxWidth:400,padding:"16px",fontSize:15,borderRadius:14,letterSpacing:1,background:"linear-gradient(135deg,#c8102e,#e8b923,#c8102e)",backgroundSize:"200% 100%"}}
        onClick={()=>{sessionStorage.setItem("seen","1");setView("main");}}>
        ¡Entrar a la porra! →
      </button>
    </div>
  );

  if(view==="register")return <RegisterView db={db} upDb={upDb} closed={closed} saveDraft={saveDraft} loadDraft={loadDraft} deleteDraft={deleteDraft} onDone={(name)=>{if(name){setMyName(name);sessionStorage.setItem("myName",name);}setView("main");}} onBack={()=>setView("main")}/>;
  if(view==="predict"&&selId){
    const p=(db.parts||[]).find(x=>x.id===selId);
    if(!p){setView("main");return null;}
    return <PredictView p={p} allM={allM} db={db} onSave={async upd=>{
      await upDb("parts",(db.parts||[]).map(x=>x.id===upd.id?{...upd,locked:true}:x));
      setView("main");setTab("ranking");
    }} onBack={()=>setView("main")}/>;
  }
  if(view==="admin")return <AdminView db={db} upDb={upDb} allM={allM} onBack={()=>setView("main")}/>;
  if(view==="compare"&&selId){
    const p=(db.parts||[]).find(x=>x.id===selId);
    if(!p){setView("main");return null;}
    return <CompareView p={p} allM={allM} scores={scores} db={db} bote={bote} onBack={()=>setView("main")}/>;
  }

  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:22}}>⚽</span>
            <div><div style={S.t1}>SÚPER PORRA</div><div style={S.t2}>MUNDIAL 2026</div></div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={S.bote}>💰 {bote}€</div>
            <div style={{fontSize:10,color:"#444",marginTop:1}}>{bote/CUOTA||0} participante{bote/CUOTA!==1?"s":""}</div>
          </div>
        </div>
        {closed&&<div style={S.bannerR}>🔒 Pronósticos cerrados{db.deadline?" · "+fd(db.deadline):""}</div>}
        {db.deadline&&!closed&&<div style={S.bannerG}>⏰ Plazo: <strong>{fd(db.deadline)}</strong></div>}
        <div style={{borderTop:"1px solid #1a1a2a",paddingTop:8,display:"flex",justifyContent:"flex-end"}}>
          {admin?(
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <span style={{fontSize:11,color:"#888",background:"#1a1a1a",padding:"2px 8px",borderRadius:10}}>⚙️ Admin</span>
              <button style={S.btnSm} onClick={()=>setView("admin")}>Panel</button>
              <button style={{...S.btnSm,background:"#333"}} onClick={()=>setAdmin(false)}>Salir</button>
            </div>
          ):showPin?(
            <div style={{display:"flex",gap:5,alignItems:"center"}}>
              <input type="password" placeholder="PIN" value={pin} onChange={e=>setPin(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter"){if(pin===PIN){setAdmin(true);setShowPin(false);setPin("");}else alert("PIN incorrecto");}}}
                style={S.pinInp}/>
              <button style={S.btnSm} onClick={()=>{if(pin===PIN){setAdmin(true);setShowPin(false);setPin("");}else alert("PIN incorrecto");}}>OK</button>
              <button style={{...S.btnSm,background:"#222"}} onClick={()=>setShowPin(false)}>✕</button>
            </div>
          ):(
            <button style={S.btnGhost} onClick={()=>setShowPin(true)}>🔒 Admin</button>
          )}
        </div>
      </div>

      <div style={S.tabs}>
        {[["ranking","📊 Ranking"],["pronos","🎯 Pronos"],["partidos","⚽ Partidos"],["reglas","📋 Reglas"]].map(([id,l])=>(
          <button key={id} style={{...S.tab,...(tab===id?S.tabA:{})}} onClick={()=>setTab(id)}>{l}</button>
        ))}
      </div>

      <div style={S.content}>
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
            {!allLocked&&scores.length>0&&<div style={{background:"rgba(37,99,235,0.1)",border:"1px solid rgba(37,99,235,0.25)",borderRadius:10,padding:"8px 14px",fontSize:12,color:"#93c5fd",marginBottom:10,textAlign:"center",backdropFilter:"blur(4px)"}}>🔒 Pronósticos privados hasta que todos confirmen</div>}
            {scores.length===0?<div style={S.empty}>El ranking aparecerá cuando haya participantes</div>:
              scores.map((p,i)=>(
                <div key={p.id} style={{...S.row,...(i===scores.length-1&&scores.length>1?{borderColor:"#7f1d1d33",background:"#140a0a"}:{})}}>
                  <Pos n={i+1}/>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:700,color:"#fff"}}>{p.name}</div>
                    <div style={{display:"flex",gap:6,marginTop:2,flexWrap:"wrap"}}>
                      {allLocked&&p.pre?.campeon&&<div style={{fontSize:11,color:"#444"}}>🏆 {p.pre.campeon.split(" ")[0]}</div>}
                      {p.fetiche&&<div style={{fontSize:11,color:"#fbbf24"}}>⭐ {(allM.find(m=>m.id===p.fetiche)||{l:"?"}).l?.split(" ")[0]}</div>}
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:900,fontSize:20,color:i===0?"#fbbf24":i<3?"#4ade80":"#bbb"}}>{p.tot}<span style={{fontSize:11,color:"#333",fontWeight:400}}> pts</span></div>
                    {i<3&&bote>0&&<div style={{fontSize:11,color:"#4ade80"}}>+{Math.round(bote*[.6,.25,.15][i])}€</div>}
                  </div>
                  {allLocked&&<button style={{...S.btnGhost,fontSize:11,padding:"3px 7px",marginLeft:4}} onClick={()=>{setSelId(p.id);setView("compare");}}>📊</button>}
                </div>
              ))
            }
            {scores.length>1&&<div style={{background:"#140a0a",border:"1px dashed #7f1d1d",borderRadius:10,padding:"10px 14px",marginTop:8,display:"flex",gap:10,alignItems:"center"}}>
              <span style={{fontSize:20}}>😎</span><div><div style={{fontWeight:700,color:"#ff6b35"}}>Farolillo Rojo</div><div style={{fontSize:12,color:"#555"}}>¡Ronda épica para {scores[scores.length-1].name}!</div></div>
            </div>}
          </div>
        )}

        {tab==="pronos"&&(
          <div>
            <div style={{...S.regCard,...(closed?{opacity:.6}:{})}}>
              <div>
                <div style={{fontWeight:700,color:"#fff"}}>¿Aún no estás apuntado?</div>
                <div style={{fontSize:12,color:"#555",marginTop:2}}>{CUOTA}€ · {db.deadline&&!closed?"Hasta "+fd(db.deadline):"Sin límite"}</div>
              </div>
              <button style={{...S.btnGold,...(closed?{opacity:.4,cursor:"not-allowed"}:{})}} onClick={!closed?()=>setView("register"):undefined}>{closed?"Cerrado":"Unirse →"}</button>
            </div>
            {(db.parts||[]).length===0?<div style={S.empty}>¡Sé el primero en apuntarte!</div>:
              (db.parts||[]).map(p=>{
                const sc=scores.find(x=>x.id===p.id);
                const mp=p.mp||{};
                const mD=PM.filter(m=>mp[m.id]&&mp[m.id].l!==""&&mp[m.id].v!=="").length;
                const pD=[p.pre?.campeon,p.pre?.subcampeon,p.pre?.tercero,p.pre?.goleador,p.pre?.mvp].filter(Boolean).length;
                const sD=[p.spc?.expulsado,p.spc?.hattrick,p.spc?.revelacion,p.spc?.goleada].filter(Boolean).length;
                const feticheM=allM.find(m=>m.id===p.fetiche);
                return(
                  <div key={p.id} style={S.card}>
                    <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:10}}>
                      <div>
                        <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                          <span style={{fontWeight:700,fontSize:15,color:"#fff"}}>{p.name}</span>
                          {p.locked&&<span style={{fontSize:10,background:"#0d1a0d",color:"#4ade80",border:"1px solid #4ade8033",borderRadius:10,padding:"1px 7px"}}>🔒</span>}
                        </div>
                        <div style={{fontWeight:800,fontSize:18,color:"#4ade80",marginTop:2}}>{sc?.tot||0}<span style={{fontSize:11,color:"#333",fontWeight:400}}> pts</span></div>
                        {feticheM&&<div style={{fontSize:11,color:"#fbbf24",marginTop:2}}>⭐ Fetiche: {feticheM.l.split(" ")[0]} vs {feticheM.v.split(" ")[0]}</div>}
                      </div>
                      {!p.locked&&!closed&&<button style={S.btnGold} onClick={()=>{setSelId(p.id);setView("predict");setMyName(p.name);sessionStorage.setItem("myName",p.name);}}>✏️ Editar</button>}
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
        )}
        {tab==="partidos"&&<PartidosTab allM={allM}/>}
        {tab==="reglas"&&<ReglasTab/>}
      </div>
    </div>
  );
}

// ── REGISTER ──────────────────────────────────────────────────────────────────
function RegisterView({db,upDb,closed,saveDraft,loadDraft,deleteDraft,onDone,onBack}){
  const[step,setStep]=useState("form");
  const[name,setName]=useState("");
  const[err,setErr]=useState("");
  const[saving,setSaving]=useState(false);
  const[hasDraft,setHasDraft]=useState(false);
  const[draftLoading,setDraftLoading]=useState(false);
  const[draft,setDraft]=useState({pre:{campeon:"",subcampeon:"",tercero:"",goleador:"",mvp:"",semis:["","","",""]},spc:{expulsado:"",hattrick:"",revelacion:"",goleada:""},mp:{},fetiche:null});
  const[sec,setSec]=useState("pre");
  const[gr,setGr]=useState("A");

  // Auto-save draft every time draft changes (after name is set)
  useEffect(()=>{
    if(name.trim()&&step==="predict"){
      const t=setTimeout(()=>saveDraft(name,draft),1500);
      return()=>clearTimeout(t);
    }
  },[draft,name,step,saveDraft]);

  const checkDraft=async(n)=>{
    if(!n.trim())return;
    setDraftLoading(true);
    const d=await loadDraft(n);
    setHasDraft(!!d&&d._draftName);
    setDraftLoading(false);
  };

  const resumeDraft=async()=>{
    setDraftLoading(true);
    const d=await loadDraft(name);
    if(d){
      const{_draftName,_savedAt,...rest}=d;
      setDraft(rest);
      setStep("predict");
    }
    setDraftLoading(false);
  };

  const next=()=>{
    const n=name.trim();
    if(!n){setErr("Escribe tu nombre");return;}
    if((db.parts||[]).find(p=>p.name.toLowerCase()===n.toLowerCase())){setErr("Ese nombre ya existe");return;}
    setErr("");setStep("predict");
  };
  const confirm=async()=>{
    setSaving(true);
    const np={id:Date.now().toString(),name:name.trim(),pre:draft.pre,spc:draft.spc,mp:draft.mp,fetiche:draft.fetiche,locked:true,registeredAt:new Date().toISOString()};
    await upDb("parts",[...(db.parts||[]),np]);
    await deleteDraft(name);
    setSaving(false);onDone(name.trim());
  };
  const upPre=(k,v)=>setDraft(d=>({...d,pre:{...d.pre,[k]:v}}));
  const upSemi=(i,v)=>setDraft(d=>{const s=[...(d.pre.semis||["","","",""])];s[i]=v;return{...d,pre:{...d.pre,semis:s}};});
  const upSpc=(k,v)=>setDraft(d=>({...d,spc:{...d.spc,[k]:v}}));
  const upM=(id,side,v)=>setDraft(d=>({...d,mp:{...d.mp,[id]:{...(d.mp[id]||{}),[side]:v}}}));
  const setFetiche=(id)=>setDraft(d=>({...d,fetiche:d.fetiche===id?null:id}));
  const grupoM=PM.filter(m=>m.g===gr);

  if(closed)return(
    <div style={S.app}><div style={S.hdr}><div style={S.hdrRow}><button style={S.backBtn} onClick={onBack}>◀</button><div style={{color:"#fbbf24",fontWeight:800}}>Registro</div><div/></div></div>
    <div style={{...S.content,textAlign:"center",paddingTop:60}}><div style={{fontSize:48}}>🔒</div><div style={{color:"#ff6b35",fontWeight:700,fontSize:18,marginTop:12}}>Pronósticos cerrados</div></div></div>
  );

  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={step==="form"?onBack:()=>setStep("form")}>◀</button>
          <div style={{textAlign:"center"}}><div style={{fontWeight:800,fontSize:16,color:"#fbbf24"}}>{step==="form"?"🎟 Nueva inscripción":"🎯 Tus pronósticos"}</div>{step==="predict"&&<div style={{fontSize:11,color:"#666"}}>{name}</div>}</div>
          {step==="predict"?(
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:2}}>
              <button style={{...S.btnGold,...(saving?{opacity:.5}:{})}} onClick={!saving?confirm:undefined}>{saving?"…":"✅ Confirmar"}</button>
              <div style={{fontSize:9,color:"#555"}}>💾 guardado auto</div>
            </div>
          ):<div/>}
        </div>
      </div>
      {step==="form"&&(
        <div style={S.content}>
          <div style={{background:"linear-gradient(135deg,#0a2a0a,#062006)",border:"1px solid #4ade8044",borderRadius:14,padding:"18px",marginBottom:20,textAlign:"center"}}>
            <div style={{fontSize:40,fontWeight:900,color:"#4ade80"}}>{CUOTA}€</div>
            <div style={{color:"#86efac",fontSize:14,marginTop:4}}>Inscripción a la porra</div>
            <div style={{color:"#444",fontSize:12,marginTop:6}}>Una vez confirmado, tus pronósticos quedan bloqueados definitivamente</div>
          </div>
          <div style={{marginBottom:14}}>
            <div style={S.fl}>Tu nombre</div>
            <input style={{...S.inp,...(err?{borderColor:"#ef4444"}:{})}} placeholder="¿Cómo te llamas?" value={name}
              onChange={e=>{setName(e.target.value);setErr("");setHasDraft(false);}}
              onBlur={e=>checkDraft(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&next()}/>
            {err&&<div style={{color:"#ef4444",fontSize:12,marginTop:4}}>⚠ {err}</div>}
          </div>
          {hasDraft&&(
            <div style={{background:"#1a1400",border:"1px solid #fbbf2466",borderRadius:10,padding:"12px",marginBottom:12}}>
              <div style={{color:"#fbbf24",fontWeight:700,fontSize:13,marginBottom:6}}>💾 Tienes un borrador guardado</div>
              <div style={{color:"#888",fontSize:12,marginBottom:10}}>Puedes continuar donde lo dejaste o empezar de cero</div>
              <div style={{display:"flex",gap:8}}>
                <button style={{...S.btnGold,flex:1}} onClick={resumeDraft}>▶ Continuar borrador</button>
                <button style={{...S.btnAct,flex:1}} onClick={()=>{setHasDraft(false);next();}}>🗑 Empezar de cero</button>
              </div>
            </div>
          )}
          {!hasDraft&&<button style={{...S.btnGold,width:"100%",padding:"14px",fontSize:15,borderRadius:12}} onClick={next}>{draftLoading?"Buscando borrador…":"Continuar →"}</button>}
        </div>
      )}
      {step==="predict"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,padding:"10px 16px",background:"#080810",borderBottom:"1px solid #141420"}}>
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
            {sec==="grupos"&&<GruposSec mp={draft.mp} gr={gr} setGr={setGr} upM={upM} grupoM={grupoM} results={{}} locked={false} fetiche={draft.fetiche} setFetiche={setFetiche}/>}
            {sec==="elim"&&<ElimSec mp={draft.mp} upM={upM} elimM={[]} locked={false} fetiche={draft.fetiche} setFetiche={setFetiche}/>}
            {sec==="retos"&&<RetosSec spc={draft.spc} upSpc={upSpc} locked={false}/>}
            <div style={{height:10}}/>
            {!draft.fetiche&&<div style={{background:"#1a1a00",border:"1px solid #fbbf2444",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#fbbf24",marginBottom:10}}>⭐ ¡No olvides marcar tu partido fetiche! (estrella junto a cualquier partido)</div>}
            <button style={{...S.btnGold,width:"100%",padding:"14px",fontSize:15,borderRadius:12,...(saving?{opacity:.5}:{})}} onClick={!saving?confirm:undefined}>{saving?"Guardando…":"✅ Confirmar y bloquear"}</button>
            <div style={{color:"#333",fontSize:11,textAlign:"center",marginTop:5}}>⚠ Una vez confirmado no podrás modificarlos</div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── PREDICT VIEW ──────────────────────────────────────────────────────────────
function PredictView({p,allM,db,onSave,onBack}){
  const[data,setData]=useState({...p,pre:p.pre||{campeon:"",subcampeon:"",tercero:"",goleador:"",mvp:"",semis:["","","",""]},spc:p.spc||{expulsado:"",hattrick:"",revelacion:"",goleada:""},mp:p.mp||{},fetiche:p.fetiche||null});
  const[sec,setSec]=useState("pre");
  const[gr,setGr]=useState("A");
  const[saving,setSaving]=useState(false);
  const locked=p.locked||false;
  const elimM=allM.filter(m=>m.ph!=="grupos");
  const grupoM=allM.filter(m=>m.g===gr);
  const upPre=(k,v)=>setData(d=>({...d,pre:{...d.pre,[k]:v}}));
  const upSemi=(i,v)=>setData(d=>{const s=[...(d.pre.semis||["","","",""])];s[i]=v;return{...d,pre:{...d.pre,semis:s}};});
  const upSpc=(k,v)=>setData(d=>({...d,spc:{...d.spc,[k]:v}}));
  const upM=(id,side,v)=>setData(d=>({...d,mp:{...d.mp,[id]:{...(d.mp[id]||{}),[side]:v}}}));
  const setFetiche=(id)=>setData(d=>({...d,fetiche:d.fetiche===id?null:id}));
  const mp=data.mp||{};
  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={onBack}>◀</button>
          <div style={{fontWeight:800,fontSize:15,color:"#fbbf24"}}>✏️ {p.name}</div>
          {!locked?<button style={{...S.btnGold,...(saving?{opacity:.5}:{})}} onClick={async()=>{setSaving(true);await onSave(data);setSaving(false);}}>{saving?"…":"💾 Confirmar"}</button>
          :<div style={{fontSize:11,color:"#4ade80"}}>🔒 Bloqueado</div>}
        </div>
      </div>
      {locked&&<div style={{background:"#0d1a00",borderBottom:"1px solid #14532d",padding:"6px 16px",fontSize:12,color:"#4ade80"}}>🔒 Solo lectura</div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,padding:"10px 16px",background:"#080810",borderBottom:"1px solid #141420"}}>
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
        {sec==="elim"&&<ElimSec mp={mp} upM={upM} elimM={elimM} locked={locked} fetiche={data.fetiche} setFetiche={setFetiche}/>}
        {sec==="retos"&&<RetosSec spc={data.spc} upSpc={upSpc} locked={locked}/>}
        {!locked&&<><div style={{height:10}}/><button style={{...S.btnGold,width:"100%",padding:"14px",fontSize:15}} onClick={async()=>{setSaving(true);await onSave(data);setSaving(false);}}>💾 Confirmar y bloquear</button><div style={{color:"#333",fontSize:11,textAlign:"center",marginTop:5}}>⚠ No podrás modificarlos después</div></>}
      </div>
    </div>
  );
}

// ── PRE SECTION ───────────────────────────────────────────────────────────────
function PreSec({pre,upPre,upSemi,locked}){
  return(
    <div>
      {locked&&<div style={{background:"#0d1a00",border:"1px solid #14532d",borderRadius:8,padding:"8px 12px",fontSize:12,color:"#4ade80",marginBottom:12}}>🔒 Solo lectura</div>}
      <div style={S.block}>
        <div style={S.blockT}>🏆 Podio final</div>
        {[["campeon","Campeón del mundo","30 pts"],["subcampeon","Subcampeón","20 pts"],["tercero","Tercer puesto","15 pts"]].map(([k,l,pts])=>(
          <div key={k} style={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:13,color:"#ccc"}}>{l}</span><span style={{fontSize:11,color:"#fbbf24"}}>{pts}</span></div>
            <TeamSel val={pre[k]||""} onChange={v=>upPre(k,v)} disabled={locked}/>
          </div>
        ))}
      </div>
      <div style={S.block}>
        <div style={S.blockT}>👟 Premios individuales</div>
        {[["goleador","Máximo goleador","20 pts"],["mvp","Mejor jugador (MVP)","15 pts"]].map(([k,l,pts])=>(
          <div key={k} style={{marginBottom:10}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:13,color:"#ccc"}}>{l}</span><span style={{fontSize:11,color:"#fbbf24"}}>{pts}</span></div>
            <input style={{...S.inp,opacity:locked?.6:1}} value={pre[k]||""} onChange={e=>!locked&&upPre(k,e.target.value)} disabled={locked} placeholder="Nombre del jugador…"/>
          </div>
        ))}
      </div>
      <div style={S.block}>
        <div style={S.blockT}>4️⃣ Semifinalistas <span style={{fontSize:11,color:"#fbbf24",fontWeight:400}}>(10 pts c/u)</span></div>
        {[0,1,2,3].map(i=>(
          <div key={i} style={{marginBottom:8}}>
            <div style={{fontSize:12,color:"#555",marginBottom:3}}>Semifinalista {i+1}</div>
            <TeamSel val={(pre.semis||[])[i]||""} onChange={v=>upSemi(i,v)} disabled={locked}/>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── GRUPOS SECTION ────────────────────────────────────────────────────────────
function GruposSec({mp,gr,setGr,upM,grupoM,results,locked,fetiche,setFetiche}){
  const gs=Object.keys(GR);
  return(
    <div>
      <div style={{fontSize:12,color:"#444",marginBottom:4}}>Pronostica el marcador · ×1 en grupos</div>
      <div style={{fontSize:11,color:"#fbbf24",marginBottom:8}}>⭐ Toca la estrella para marcar tu partido fetiche (solo uno)</div>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:10}}>
        {gs.map(g=>{
          const total=PM.filter(m=>m.g===g).length;
          const done=PM.filter(m=>m.g===g&&mp[m.id]&&mp[m.id].l!==""&&mp[m.id].v!=="").length;
          const hasFetiche=PM.filter(m=>m.g===g).some(m=>m.id===fetiche);
          return(
            <button key={g} style={{...S.chip,...(gr===g?S.chipA:{}),position:"relative",...(hasFetiche?{borderColor:"#fbbf2466"}:{})}} onClick={()=>setGr(g)}>
              Gr.{g}
              {done===total&&<span style={{position:"absolute",top:-5,right:-5,background:"#4ade80",color:"#000",borderRadius:"50%",width:13,height:13,fontSize:9,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900}}>✓</span>}
            </button>
          );
        })}
      </div>
      <div style={{display:"flex",gap:4,marginBottom:10,flexWrap:"wrap"}}>
        {(GR[gr]||[]).map(t=><span key={t} style={{background:"#0d0d14",border:"1px solid #1a2a1a",borderRadius:10,padding:"2px 7px",fontSize:11,color:"#777"}}>{t}</span>)}
      </div>
      {grupoM.map(m=>{
        const pred=mp[m.id]||{};
        const res=results[m.id]||m.result;
        const pts=res?mPts(pred,res,"grupos",m.id===fetiche):null;
        const filled=pred.l!==""&&pred.v!=="";
        const isFet=m.id===fetiche;
        return(
          <div key={m.id} style={{...S.mCard,borderColor:isFet?"#fbbf24":res?"#4ade8033":filled?"#fbbf2422":"#141420"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
              <span style={{fontSize:11,color:"#333"}}>{fd(m.d)}</span>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {isFet&&<span style={{fontSize:10,color:"#fbbf24",fontWeight:700}}>×5/×3</span>}
                {!locked&&<button onClick={()=>setFetiche(m.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,padding:0,opacity:isFet?1:.3,filter:isFet?"none":"grayscale(1)"}}>⭐</button>}
                {locked&&isFet&&<span style={{fontSize:16}}>⭐</span>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 42px 14px 42px 1fr",alignItems:"center",gap:4}}>
              <div style={{textAlign:"right",fontWeight:600,color:"#ddd",fontSize:12,lineHeight:1.3}}>{m.l}</div>
              <input type="number" min={0} max={20} value={pred.l||""} onChange={e=>!locked&&upM(m.id,"l",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1,...(isFet?{borderColor:"#fbbf24"}:{})}} disabled={locked} placeholder="?"/>
              <span style={{textAlign:"center",color:"#1e1e1e",fontWeight:700}}>–</span>
              <input type="number" min={0} max={20} value={pred.v||""} onChange={e=>!locked&&upM(m.id,"v",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1,...(isFet?{borderColor:"#fbbf24"}:{})}} disabled={locked} placeholder="?"/>
              <div style={{fontWeight:600,color:"#ddd",fontSize:12,lineHeight:1.3}}>{m.v}</div>
            </div>
            {res&&<div style={{textAlign:"center",fontSize:11,marginTop:6}}><span style={{color:"#4ade80"}}>Real: {res.l}–{res.v}</span>{pts!==null&&<span style={{color:isFet?"#fbbf24":"#4ade80",marginLeft:8,fontWeight:700}}>{isFet?"⭐":""} +{pts}pts</span>}</div>}
          </div>
        );
      })}
    </div>
  );
}

// ── ELIM SECTION ──────────────────────────────────────────────────────────────
function ElimSec({mp,upM,elimM,locked,fetiche,setFetiche}){
  if(!elimM||elimM.length===0)return(
    <div style={{...S.empty,paddingTop:50}}>
      <div style={{fontSize:36}}>⏳</div>
      <div style={{color:"#444",marginTop:8,fontWeight:600}}>Los partidos de eliminatoria aparecerán aquí</div>
      <div style={{fontSize:12,color:"#2a2a2a",marginTop:4}}>Octavos · Cuartos · Semis · Final</div>
    </div>
  );
  return(
    <div>
      <div style={{fontSize:11,color:"#fbbf24",marginBottom:8}}>⭐ Toca la estrella para marcar tu partido fetiche (solo uno)</div>
      {elimM.map(m=>{
        const pred=mp[m.id]||{};
        const res=m.result;
        const pts=res?mPts(pred,res,m.ph,m.id===fetiche):null;
        const isFet=m.id===fetiche;
        return(
          <div key={m.id} style={{...S.mCard,borderColor:isFet?"#fbbf24":res?"#4ade8033":"#141420"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:6,alignItems:"center"}}>
              <div><span style={{fontSize:11,color:"#444"}}>{fd(m.date||"")}</span><span style={{fontSize:11,color:"#60a5fa",fontWeight:700,marginLeft:8}}>{PHL[m.ph]} ·×{MULT[m.ph]}</span></div>
              <div style={{display:"flex",alignItems:"center",gap:6}}>
                {isFet&&<span style={{fontSize:10,color:"#fbbf24",fontWeight:700}}>×5/×3</span>}
                {!locked&&<button onClick={()=>setFetiche(m.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,padding:0,opacity:isFet?1:.3,filter:isFet?"none":"grayscale(1)"}}>⭐</button>}
                {locked&&isFet&&<span style={{fontSize:16}}>⭐</span>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 42px 14px 42px 1fr",alignItems:"center",gap:4}}>
              <div style={{textAlign:"right",fontWeight:600,color:"#ddd",fontSize:12}}>{m.l}</div>
              <input type="number" min={0} max={20} value={pred.l||""} onChange={e=>!locked&&upM(m.id,"l",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1,...(isFet?{borderColor:"#fbbf24"}:{})}} disabled={locked} placeholder="?"/>
              <span style={{textAlign:"center",color:"#1e1e1e",fontWeight:700}}>–</span>
              <input type="number" min={0} max={20} value={pred.v||""} onChange={e=>!locked&&upM(m.id,"v",e.target.value)} style={{...S.scoreI,opacity:locked?.5:1,...(isFet?{borderColor:"#fbbf24"}:{})}} disabled={locked} placeholder="?"/>
              <div style={{fontWeight:600,color:"#ddd",fontSize:12}}>{m.v}</div>
            </div>
            {res&&<div style={{textAlign:"center",fontSize:11,marginTop:6}}><span style={{color:"#4ade80"}}>Real: {res.l}–{res.v}</span>{pts!==null&&<span style={{color:isFet?"#fbbf24":"#4ade80",marginLeft:8,fontWeight:700}}>{isFet?"⭐":""} +{pts}pts</span>}</div>}
          </div>
        );
      })}
    </div>
  );
}

// ── RETOS SECTION ─────────────────────────────────────────────────────────────
function RetosSec({spc,upSpc,locked}){
  return(
    <div>
      {[{k:"expulsado",t:"🟥 Primer expulsado",p:"10 pts",d:"Primer jugador en ver roja",ph:"Nombre del jugador"},{k:"hattrick",t:"🎩 Primer hat-trick",p:"10 pts",d:"Primero en marcar 3 goles en un partido",ph:"Nombre del jugador"}].map(r=>(
        <div key={r.k} style={S.block}>
          <div style={{display:"flex",justifyContent:"space-between"}}><span style={S.blockT}>{r.t}</span><span style={{fontSize:11,color:"#fbbf24"}}>{r.p}</span></div>
          <div style={{fontSize:11,color:"#444",marginBottom:8,marginTop:3}}>{r.d}</div>
          <input style={{...S.inp,opacity:locked?.6:1}} value={spc[r.k]||""} onChange={e=>!locked&&upSpc(r.k,e.target.value)} disabled={locked} placeholder={r.ph}/>
        </div>
      ))}
      <div style={S.block}>
        <div style={{display:"flex",justifyContent:"space-between"}}><span style={S.blockT}>🚀 Equipo revelación</span><span style={{fontSize:11,color:"#fbbf24"}}>15 pts</span></div>
        <div style={{fontSize:11,color:"#444",marginBottom:8,marginTop:3}}>El que llegue más lejos de lo esperado</div>
        <TeamSel val={spc.revelacion||""} onChange={v=>upSpc("revelacion",v)} disabled={locked}/>
      </div>
      <div style={S.block}>
        <div style={{display:"flex",justifyContent:"space-between"}}><span style={S.blockT}>🎯 Mayor goleada</span><span style={{fontSize:11,color:"#fbbf24"}}>10 pts</span></div>
        <div style={{fontSize:11,color:"#444",marginBottom:8,marginTop:3}}>Resultado exacto — ej: Brasil 7-0 Haití</div>
        <input style={{...S.inp,opacity:locked?.6:1}} value={spc.goleada||""} onChange={e=>!locked&&upSpc("goleada",e.target.value)} disabled={locked} placeholder="Ej: Brasil 7-0 Haití"/>
      </div>
    </div>
  );
}

// ── PARTIDOS TAB ──────────────────────────────────────────────────────────────
// ── CLASIFICACIÓN DE GRUPO ───────────────────────────────────────────────────
function calcClasificacion(teams, matchesGrupo){
  const tabla={};
  teams.forEach(t=>{tabla[t]={pj:0,g:0,e:0,p:0,gf:0,gc:0,pts:0};});
  matchesGrupo.filter(m=>m.result).forEach(m=>{
    const rl=+m.result.l, rv=+m.result.v;
    if(isNaN(rl)||isNaN(rv))return;
    const tl=tabla[m.l], tv=tabla[m.v];
    if(!tl||!tv)return;
    tl.pj++;tv.pj++;
    tl.gf+=rl;tl.gc+=rv;
    tv.gf+=rv;tv.gc+=rl;
    if(rl>rv){tl.g++;tl.pts+=3;tv.p++;}
    else if(rv>rl){tv.g++;tv.pts+=3;tl.p++;}
    else{tl.e++;tv.e++;tl.pts++;tv.pts++;}
  });
  return Object.entries(tabla)
    .map(([name,s])=>({name,dif:s.gf-s.gc,...s}))
    .sort((a,b)=>b.pts-a.pts||b.dif-a.dif||b.gf-a.gf);
}

function TablaClasificacion({grupo, allM}){
  const teams=GR[grupo]||[];
  const matchesG=allM.filter(m=>m.g===grupo);
  const hayResultados=matchesG.some(m=>m.result);
  if(!hayResultados)return(
    <div style={{background:"rgba(0,0,0,0.3)",borderRadius:10,padding:"12px",textAlign:"center",fontSize:12,color:"#2d3748",marginBottom:14}}>
      Sin resultados aún en el Grupo {grupo}
    </div>
  );
  const tabla=calcClasificacion(teams,matchesG);
  const cols=["#","Equipo","PJ","G","E","P","GF","GC","DIF","PTS"];
  return(
    <div style={{marginBottom:16}}>
      <div style={{fontSize:11,color:"#e8b923",fontWeight:700,letterSpacing:1,marginBottom:8,textTransform:"uppercase"}}>📊 Clasificación Grupo {grupo}</div>
      <div style={{background:"rgba(0,0,0,0.4)",borderRadius:12,overflow:"hidden",border:"1px solid rgba(255,255,255,0.07)"}}>
        {/* Header */}
        <div style={{display:"grid",gridTemplateColumns:"20px 1fr 28px 28px 28px 28px 28px 28px 36px 36px",gap:0,padding:"7px 10px",background:"rgba(232,185,35,0.08)",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
          {cols.map(c=><div key={c} style={{fontSize:10,fontWeight:700,color:"#4a5568",textAlign:c==="Equipo"?"left":"center",letterSpacing:0.5}}>{c}</div>)}
        </div>
        {/* Rows */}
        {tabla.map((row,i)=>{
          const isTop2=i<2;
          const bg=i===0?"rgba(232,185,35,0.08)":i===1?"rgba(16,185,129,0.06)":"transparent";
          const posCol=i===0?"#e8b923":i===1?"#10b981":i===2?"#e8b923":"#2d3748";
          const difCol=row.dif>0?"#10b981":row.dif<0?"#ef4444":"#4a5568";
          return(
            <div key={row.name} style={{display:"grid",gridTemplateColumns:"20px 1fr 28px 28px 28px 28px 28px 28px 36px 36px",gap:0,padding:"8px 10px",background:bg,borderBottom:i<tabla.length-1?"1px solid rgba(255,255,255,0.04)":"none",alignItems:"center"}}>
              <div style={{fontSize:11,fontWeight:700,color:posCol,textAlign:"center"}}>{i+1}</div>
              <div style={{fontSize:12,fontWeight:isTop2?700:400,color:isTop2?"#e2e8f0":"#94a3b8",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",paddingRight:4}}>
                {row.name.split(" ").slice(0,-1).join(" ")||row.name}
                {isTop2&&<span style={{marginLeft:4,fontSize:9,color:i===0?"#e8b923":"#10b981"}}>●</span>}
              </div>
              {[row.pj,row.g,row.e,row.p,row.gf,row.gc].map((v,j)=>(
                <div key={j} style={{fontSize:11,textAlign:"center",color:"#4a5568"}}>{v}</div>
              ))}
              <div style={{fontSize:11,textAlign:"center",color:difCol,fontWeight:600}}>{row.dif>0?"+"+row.dif:row.dif}</div>
              <div style={{fontSize:13,textAlign:"center",fontWeight:900,color:isTop2?"#e8b923":"#e2e8f0"}}>{row.pts}</div>
            </div>
          );
        })}
      </div>
      <div style={{display:"flex",gap:12,marginTop:6,paddingLeft:4}}>
        <div style={{fontSize:10,color:"#e8b923"}}>● Clasificado</div>
        <div style={{fontSize:10,color:"#10b981"}}>● En zona de clasificación</div>
      </div>
    </div>
  );
}

function PartidosTab({allM}){
  const[gf,setGf]=useState("A");
  const gs=Object.keys(GR);
  const isElim=gf==="ELIM";
  const shown=isElim?allM.filter(m=>m.ph!=="grupos"):allM.filter(m=>m.g===gf);
  return(
    <div>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:12}}>
        {gs.map(g=>{
          const hasR=allM.filter(m=>m.g===g).some(m=>m.result);
          return<button key={g} style={{...S.chip,...(gf===g?S.chipA:{}),...(hasR?{borderColor:"rgba(16,185,129,0.3)"}:{})}} onClick={()=>setGf(g)}>Gr.{g}</button>;
        })}
        {allM.some(m=>m.ph!=="grupos")&&<button style={{...S.chip,...(isElim?S.chipA:{})}} onClick={()=>setGf("ELIM")}>Elim.</button>}
      </div>

      {/* Tabla de clasificación — solo en grupos */}
      {!isElim&&<TablaClasificacion grupo={gf} allM={allM}/>}

      {/* Partidos */}
      <div style={{fontSize:11,color:"#4a5568",fontWeight:700,letterSpacing:1,marginBottom:8,textTransform:"uppercase"}}>⚽ Partidos</div>
      {shown.length===0?<div style={S.empty}>Sin partidos todavía</div>:shown.map(m=>(
        <div key={m.id} style={{...S.mCard,borderColor:m.result?"rgba(16,185,129,0.25)":"rgba(255,255,255,0.06)"}}>
          <div style={{fontSize:11,color:"#2d3748",marginBottom:4}}>{fd(m.d||m.date||"")}{m.ph!=="grupos"?" · "+PHL[m.ph]:""}</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr auto 1fr",alignItems:"center",gap:8}}>
            <div style={{textAlign:"right",fontWeight:600,color:"#e2e8f0",fontSize:13}}>{m.l}</div>
            {m.result?(
              <div style={{textAlign:"center",fontWeight:900,color:"#10b981",fontSize:22,minWidth:60,background:"rgba(16,185,129,0.08)",borderRadius:8,padding:"2px 6px"}}>{m.result.l}–{m.result.v}</div>
            ):(
              <div style={{textAlign:"center",color:"rgba(255,255,255,0.1)",fontSize:16,minWidth:60}}>vs</div>
            )}
            <div style={{fontWeight:600,color:"#e2e8f0",fontSize:13}}>{m.v}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ── REGLAS TAB ────────────────────────────────────────────────────────────────
function ReglasTab(){
  const B=[
    {t:"1️⃣ Pre-Mundial",c:"#fbbf24",items:[["Campeón","30pts"],["Subcampeón","20pts"],["3er puesto","15pts"],["Máx. goleador","20pts"],["MVP","15pts"],["4 Semifinalistas","10pts c/u"]]},
    {t:"2️⃣ Por partido",c:"#4ade80",items:[["Resultado exacto","5pts"],["Ganador/empate","2pts"],["Goles de un equipo","1pt c/u"]]},
    {t:"⭐ Partido Fetiche",c:"#fbbf24",items:[["Si aciertas resultado exacto","pts ×5"],["Si aciertas 1X2","pts ×3"],["Solo 1 por participante","Elige bien"]]},
    {t:"3️⃣ Multiplicadores fase",c:"#60a5fa",items:[["Grupos","×1"],["Octavos","×1.5"],["Cuartos","×2"],["Semis","×3"],["Final","×5"]]},
    {t:"4️⃣ Retos especiales",c:"#f472b6",items:[["Primer expulsado","10pts"],["Primer hat-trick","10pts"],["Equipo revelación","15pts"],["Mayor goleada","10pts"]]},
    {t:"💰 Premios",c:"#fb923c",items:[["1er puesto","60% del bote"],["2º puesto","25% del bote"],["3er puesto","15% del bote"],["Farolillo rojo","¡Ronda épica!"]]},
  ];
  return(
    <div>
      {B.map((b,i)=>(
        <div key={i} style={{background:"#0d0d14",borderLeft:"3px solid "+b.c,borderRadius:10,padding:"12px 14px",marginBottom:10}}>
          <div style={{fontWeight:700,fontSize:13,color:b.c,marginBottom:8}}>{b.t}</div>
          {b.items.map(([l,p])=>(
            <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid #0a0a10",fontSize:13}}>
              <span style={{color:"#aaa"}}>{l}</span><span style={{fontWeight:700,color:b.c}}>{p}</span>
            </div>
          ))}
        </div>
      ))}
      <div style={{background:"#0a1a0a",border:"1px solid #14532d",borderRadius:10,padding:"12px",fontSize:13,color:"#86efac",lineHeight:1.7}}>
        💡 <strong>Ejemplo semis:</strong> Pronosticas 2-1, real 3-1 → ganador(2) + goles visitante(1) = 3 × 3 = <strong style={{color:"#fbbf24"}}>9pts</strong>
      </div>
    </div>
  );
}

// ── COMPARE VIEW ──────────────────────────────────────────────────────────────
function CompareView({p,allM,scores,db,bote,onBack}){
  const[gr,setGr]=useState("A");
  const playedM=allM.filter(m=>m.result);
  const grupoPlayed=playedM.filter(m=>m.g===gr);
  const mp=p.mp||{};
  const fetiche=p.fetiche||null;
  const pp=prePts(p.pre||{},db.rPre);
  const sp=spcPts(p.spc||{},db.rSpc);
  let matchT=0;playedM.forEach(m=>{matchT+=mPts(mp[m.id],m.result,m.ph||"grupos",m.id===fetiche);});
  matchT=Math.round(matchT*10)/10;
  const pos=scores.findIndex(x=>x.id===p.id)+1;
  const gs=Object.keys(GR);
  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={onBack}>◀</button>
          <div style={{textAlign:"center"}}><div style={{fontWeight:800,fontSize:15,color:"#fbbf24"}}>📊 {p.name}</div><div style={{fontSize:11,color:"#555"}}>Desglose de puntuación</div></div>
          <div/>
        </div>
      </div>
      <div style={S.content}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
          {[{l:"Posición",v:"#"+pos,c:"#fbbf24"},{l:"Total",v:(scores.find(x=>x.id===p.id)?.tot||0)+" pts",c:"#4ade80"},{l:"Pre-Mundial",v:pp+" pts",c:"#60a5fa"},{l:"Partidos",v:matchT+" pts",c:"#a78bfa"},{l:"Retos",v:sp+" pts",c:"#f472b6"},{l:"Premio est.",v:pos<=3?Math.round(bote*[.6,.25,.15][pos-1])+"€":"—",c:"#4ade80"}].map(x=>(
            <div key={x.l} style={{background:"#0d0d14",border:"1px solid #141420",borderRadius:10,padding:"10px 12px"}}>
              <div style={{fontSize:11,color:"#444"}}>{x.l}</div>
              <div style={{fontWeight:800,fontSize:17,color:x.c,marginTop:2}}>{x.v}</div>
            </div>
          ))}
        </div>
        {fetiche&&<div style={{background:"#1a1400",border:"1px solid #fbbf2444",borderRadius:10,padding:"10px 12px",marginBottom:12,fontSize:12,color:"#fbbf24"}}>⭐ Partido fetiche: {(allM.find(m=>m.id===fetiche)||{l:"?",v:"?"}).l} vs {(allM.find(m=>m.id===fetiche)||{l:"?",v:"?"}).v}</div>}
        {db.rPre&&(
          <div style={{background:"#0d0d14",borderLeft:"3px solid #60a5fa",borderRadius:10,padding:"12px",marginBottom:12}}>
            <div style={{fontWeight:700,color:"#60a5fa",fontSize:13,marginBottom:8}}>🏆 Pre-Mundial — {pp} pts</div>
            {[["Campeón",p.pre?.campeon,db.rPre?.campeon,30],["Subcampeón",p.pre?.subcampeon,db.rPre?.subcampeon,20],["3er puesto",p.pre?.tercero,db.rPre?.tercero,15],["Goleador",p.pre?.goleador,db.rPre?.goleador,20],["MVP",p.pre?.mvp,db.rPre?.mvp,15]].map(([l,pred,real,max])=>{
              const hit=pred&&real&&pred===real;
              return(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid #0a0a10",fontSize:12}}>
                  <span style={{color:"#666"}}>{l}</span>
                  <span style={{color:hit?"#4ade80":pred?"#444":"#222"}}>{pred||"—"}{hit?" ✓ +"+max+"pts":real&&pred?" ✗ ("+real+")":""}</span>
                </div>
              );
            })}
          </div>
        )}
        {playedM.length>0&&(
          <div>
            <div style={{fontWeight:700,color:"#a78bfa",fontSize:13,marginBottom:8}}>⚽ Partidos — {matchT} pts</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:10}}>
              {gs.map(g=><button key={g} style={{...S.chip,...(gr===g?S.chipA:{})}} onClick={()=>setGr(g)}>Gr.{g}</button>)}
            </div>
            {grupoPlayed.length===0?<div style={{color:"#333",textAlign:"center",padding:"16px",fontSize:13}}>Sin partidos jugados aquí</div>:
              grupoPlayed.map(m=>{
                const pred=mp[m.id];
                const isFet=m.id===fetiche;
                const pts=mPts(pred,m.result,m.ph||"grupos",isFet);
                return(
                  <div key={m.id} style={{background:"#0d0d14",border:"1px solid "+(pts>0?"#4ade8022":"#141420"),borderRadius:8,padding:"10px",marginBottom:6}}>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                      <span style={{fontSize:11,color:"#333"}}>{fd(m.d||"")} {isFet&&<span style={{color:"#fbbf24"}}>⭐</span>}</span>
                      <span style={{fontSize:12,fontWeight:700,color:pts>0?"#4ade80":"#444"}}>+{pts}pts</span>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr auto 1fr",gap:8,alignItems:"center"}}>
                      <div style={{textAlign:"right",fontSize:12,color:"#aaa"}}>{m.l}</div>
                      <div style={{textAlign:"center"}}>
                        <div style={{fontWeight:900,color:"#4ade80",fontSize:15}}>{m.result.l}–{m.result.v}</div>
                        {pred?<div style={{fontSize:11,color:pts>0?"#fbbf24":"#444"}}>Pred: {pred.l}–{pred.v}</div>:<div style={{fontSize:11,color:"#222"}}>Sin pred.</div>}
                      </div>
                      <div style={{fontSize:12,color:"#aaa"}}>{m.v}</div>
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

// ── ADMIN VIEW ────────────────────────────────────────────────────────────────
function AdminView({db,upDb,allM,onBack}){
  const[tab,setTab]=useState("res");
  const[gf,setGf]=useState("A");
  const[eid,setEid]=useState(null);
  const[er,setEr]=useState({l:"",v:""});
  const[pf,setPf]=useState(db.rPre||{campeon:"",subcampeon:"",tercero:"",goleador:"",mvp:"",semis:["","","",""]});
  const[sf,setSf]=useState(db.rSpc||{expulsado:"",hattrick:"",revelacion:"",goleada:""});
  const[nm,setNm]=useState({l:"",v:"",ph:"octavos",date:""});
  const[saving,setSaving]=useState(false);
  const gs=Object.keys(GR);
  const isElim=gf==="ELIM";
  const shown=isElim?allM.filter(m=>m.ph!=="grupos"):allM.filter(m=>m.g===gf);

  const saveRes=async id=>{
    setSaving(true);
    const newRes={...(db.results||{}),[id]:er};
    await upDb("results",newRes);
    const newEM=(db.extraM||[]).map(m=>m.id===id?{...m,result:er}:m);
    await upDb("extraM",newEM);
    setEid(null);setSaving(false);
  };

  const resetRes=async id=>{
    const newRes={...(db.results||{})};
    delete newRes[id];
    await upDb("results",newRes);
    const newEM=(db.extraM||[]).map(m=>m.id===id?{...m,result:null}:m);
    await upDb("extraM",newEM);
  };

  return(
    <div style={S.app}>
      <div style={S.hdr}>
        <div style={S.hdrRow}>
          <button style={S.backBtn} onClick={onBack}>◀</button>
          <div style={{fontWeight:900,fontSize:15,color:"#fbbf24"}}>⚙️ Panel Admin</div>
          <div style={{fontSize:11,color:"#444"}}>{(db.parts||[]).length} participantes</div>
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
                    <div style={{fontWeight:600,color:"#ddd",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.l} vs {m.v}</div>
                    <div style={{fontSize:10,color:"#333"}}>{fd(m.d||m.date||"")}{m.ph!=="grupos"?" · "+PHL[m.ph]:""}</div>
                    {res&&<div style={{display:"flex",alignItems:"center",gap:8,marginTop:2}}>
                      <span style={{color:"#4ade80",fontWeight:700,fontSize:11}}>✅ {res.l}–{res.v}</span>
                      <button style={{background:"#7f1d1d",border:"none",borderRadius:4,padding:"1px 6px",color:"#fff",fontSize:10,cursor:"pointer"}} onClick={()=>resetRes(m.id)}>Reset</button>
                    </div>}
                  </div>
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
                {["octavos","cuartos","semifinales","final"].map(p=><option key={p} value={p}>{PHL[p]} ·×{MULT[p]}</option>)}
              </select>
              <input style={S.inp} type="date" value={nm.date} onChange={e=>setNm({...nm,date:e.target.value})}/>
            </div>
            <button style={{...S.btnGold,width:"100%"}} onClick={async()=>{
              if(!nm.l||!nm.v)return;
              await upDb("extraM",[...(db.extraM||[]),{...nm,id:"x_"+Date.now(),result:null,g:null}]);
              setNm({l:"",v:"",ph:"octavos",date:""});
            }}>➕ Añadir partido</button>
            {(db.extraM||[]).map(m=>(
              <div key={m.id} style={{...S.mCard,display:"flex",flexDirection:"row",alignItems:"center",gap:8,marginTop:6}}>
                <div style={{flex:1}}><div style={{fontWeight:600,color:"#ddd",fontSize:12}}>{m.l} vs {m.v}</div><div style={{fontSize:10,color:"#555"}}>{PHL[m.ph]} · {fd(m.date||"")}</div></div>
                <button style={{...S.btnAct,background:"#7f1d1d"}} onClick={async()=>upDb("extraM",(db.extraM||[]).filter(x=>x.id!==m.id))}>🗑️</button>
              </div>
            ))}
          </div>
        )}

        {tab==="cfg"&&(
          <div>
            <div style={{background:"#0d0d14",border:"1px solid #1e1e2e",borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:700,color:"#ddd",fontSize:13,marginBottom:4}}>📅 Fecha límite</div>
              <input style={S.inp} type="date" value={db.deadline||""} onChange={async e=>upDb("deadline",e.target.value)}/>
              {db.deadline&&<div style={{fontSize:12,color:isPast(db.deadline)?"#ef4444":"#4ade80",marginTop:6}}>{isPast(db.deadline)?"🔒 CERRADO":"⏰ Abierto hasta "+fd(db.deadline)}</div>}
              {db.deadline&&<button style={{...S.btnAct,marginTop:8,background:"#7f1d1d"}} onClick={async()=>upDb("deadline","")}>Quitar límite</button>}
            </div>

            <div style={{background:"#0d0d14",border:"1px solid #1e2e0e",borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:700,color:"#4ade80",fontSize:13,marginBottom:8}}>📥 Exportar datos</div>
              <div style={{fontSize:12,color:"#555",marginBottom:10}}>Descarga un CSV con todos los pronósticos y puntuaciones</div>
              <button style={{...S.btnGold,width:"100%",background:"linear-gradient(135deg,#065f46,#047857)"}} onClick={()=>exportCSV(db.parts||[],allM,db.rPre,db.rSpc)}>
                📥 Descargar Excel completo
              </button>
            </div>

            <div style={{background:"#0d0d14",border:"1px solid #1e1e2e",borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontWeight:700,color:"#ddd",fontSize:13,marginBottom:10}}>👤 Participantes</div>
              {(db.parts||[]).length===0?<div style={{color:"#333",fontSize:13}}>Sin participantes</div>:(db.parts||[]).map(p=>(
                <div key={p.id} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 0",borderBottom:"1px solid #0a0a10"}}>
                  <div style={{flex:1,color:"#ccc",fontSize:13}}>{p.name}{p.locked&&<span style={{marginLeft:6,fontSize:10,color:"#4ade80"}}>🔒</span>}</div>
                  {p.locked&&<button style={{...S.btnAct,fontSize:10,padding:"3px 7px"}} onClick={async()=>{if(!window.confirm("¿Desbloquear a "+p.name+"?"))return;await upDb("parts",(db.parts||[]).map(x=>x.id===p.id?{...x,locked:false}:x));}}>🔓</button>}
                  <button style={{...S.btnAct,background:"#7f1d1d",fontSize:10,padding:"3px 7px"}} onClick={async()=>{if(!window.confirm("¿Eliminar a "+p.name+"?"))return;await upDb("parts",(db.parts||[]).filter(x=>x.id!==p.id));}}>🗑️</button>
                </div>
              ))}
            </div>

            <div style={{background:"#0d0d14",border:"1px solid #3a0a0a",borderRadius:12,padding:"14px"}}>
              <div style={{fontWeight:700,color:"#ff6b35",fontSize:13,marginBottom:8}}>⚠ Zona peligrosa</div>
              <button style={{...S.btnAct,background:"#7f1d1d",width:"100%",padding:"10px"}} onClick={async()=>{
                if(!window.confirm("¿Borrar TODOS los datos? No se puede deshacer."))return;
                await Promise.all(["parts","extraM","results","rPre","rSpc","deadline"].map(k=>upDb(k,["parts","extraM"].includes(k)?[]:k==="results"?{}:null)));
              }}>🗑️ Borrar todos los datos</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const S={
  // ── LAYOUT ────────────────────────────────────────────────────────────────
  app:{
    minHeight:"100vh",
    background:"linear-gradient(160deg,#0a1628 0%,#0d1f0d 50%,#1a0a00 100%)",
    fontFamily:"'Segoe UI',system-ui,sans-serif",
    color:"#fff",
    maxWidth:600,
    margin:"0 auto",
    position:"relative",
  },

  // ── HEADER ────────────────────────────────────────────────────────────────
  hdr:{
    background:"linear-gradient(180deg,#0a1628 0%,#091220 100%)",
    borderBottom:"2px solid transparent",
    borderImage:"linear-gradient(90deg,#c8102e,#e8b923,#c8102e) 1",
    padding:"12px 16px",
    boxShadow:"0 4px 20px rgba(0,0,0,0.5)",
  },
  hdrRow:{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,marginBottom:6},
  t1:{
    fontSize:21,fontWeight:900,
    background:"linear-gradient(135deg,#e8b923,#fff5c0,#e8b923)",
    WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent",
    letterSpacing:2,lineHeight:1,
    textShadow:"none",
  },
  t2:{fontSize:11,color:"#4a5568",letterSpacing:4,textTransform:"uppercase"},
  bote:{
    background:"linear-gradient(135deg,#064e3b,#065f46)",
    border:"1px solid #10b98166",
    borderRadius:20,padding:"4px 12px",
    fontSize:13,color:"#10b981",fontWeight:700,
    boxShadow:"0 0 12px rgba(16,185,129,0.2)",
  },

  // ── BANNERS ───────────────────────────────────────────────────────────────
  bannerR:{
    background:"linear-gradient(135deg,#450a0a,#7f1d1d)",
    border:"1px solid #dc262644",
    borderRadius:8,padding:"6px 12px",fontSize:12,
    color:"#fca5a5",marginBottom:6,
    boxShadow:"0 2px 8px rgba(220,38,38,0.2)",
  },
  bannerG:{
    background:"linear-gradient(135deg,#052e16,#064e3b)",
    border:"1px solid #16a34a44",
    borderRadius:8,padding:"5px 12px",fontSize:12,
    color:"#86efac",marginBottom:6,
  },

  // ── TABS ──────────────────────────────────────────────────────────────────
  tabs:{
    display:"flex",
    background:"rgba(0,0,0,0.4)",
    borderBottom:"1px solid rgba(232,185,35,0.15)",
    backdropFilter:"blur(10px)",
  },
  tab:{
    flex:1,padding:"12px 2px",
    background:"transparent",border:"none",
    color:"#4a5568",cursor:"pointer",fontSize:11,fontWeight:700,
    borderBottom:"2px solid transparent",
    transition:"all 0.2s",letterSpacing:0.5,
  },
  tabA:{
    color:"#e8b923",
    borderBottomColor:"#e8b923",
    background:"rgba(232,185,35,0.06)",
  },
  content:{padding:"14px 16px",paddingBottom:70},

  // ── SECTION NAV ───────────────────────────────────────────────────────────
  secNav:{
    display:"grid",gridTemplateColumns:"repeat(4,1fr)",
    background:"rgba(0,0,0,0.5)",
    borderBottom:"1px solid rgba(255,255,255,0.06)",
  },
  snBtn:{
    padding:"10px 4px",background:"transparent",border:"none",
    color:"#4a5568",cursor:"pointer",
    borderBottom:"2px solid transparent",
    textAlign:"center",fontSize:11,fontWeight:700,
    transition:"all 0.2s",
  },
  snA:{color:"#e8b923",borderBottomColor:"#e8b923",background:"rgba(232,185,35,0.05)"},

  // ── CARDS ─────────────────────────────────────────────────────────────────
  row:{
    display:"flex",alignItems:"center",gap:12,
    padding:"12px 14px",
    background:"linear-gradient(135deg,rgba(15,20,40,0.9),rgba(10,25,15,0.9))",
    borderRadius:12,marginBottom:6,
    border:"1px solid rgba(255,255,255,0.07)",
    boxShadow:"0 2px 10px rgba(0,0,0,0.3)",
    backdropFilter:"blur(4px)",
  },
  card:{
    background:"linear-gradient(135deg,rgba(12,18,38,0.95),rgba(10,20,12,0.95))",
    border:"1px solid rgba(255,255,255,0.07)",
    borderRadius:14,padding:"14px",marginBottom:10,
    boxShadow:"0 4px 16px rgba(0,0,0,0.3)",
  },
  regCard:{
    display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,
    background:"linear-gradient(135deg,#052e16,#064e3b)",
    border:"1px solid #10b98133",
    borderRadius:14,padding:"16px",marginBottom:16,
    boxShadow:"0 4px 20px rgba(16,185,129,0.15)",
  },
  mCard:{
    padding:"12px 14px",
    background:"linear-gradient(135deg,rgba(12,18,38,0.9),rgba(10,20,12,0.9))",
    borderRadius:10,marginBottom:6,
    border:"1px solid",
    boxShadow:"0 2px 8px rgba(0,0,0,0.2)",
  },
  block:{
    background:"rgba(5,10,25,0.8)",
    border:"1px solid rgba(255,255,255,0.06)",
    borderRadius:12,padding:"14px",marginBottom:12,
    backdropFilter:"blur(4px)",
  },
  blockT:{fontWeight:700,fontSize:13,color:"#e2e8f0",marginBottom:10},

  // ── PRIZES ────────────────────────────────────────────────────────────────
  prizeCard1:{
    background:"linear-gradient(135deg,#422006,#78350f)",
    border:"1px solid #e8b92366",
    borderRadius:12,padding:"12px 8px",textAlign:"center",
    boxShadow:"0 0 20px rgba(232,185,35,0.3)",
  },
  prizeCard2:{
    background:"linear-gradient(135deg,#1a1a2e,#16213e)",
    border:"1px solid #94a3b866",
    borderRadius:12,padding:"12px 8px",textAlign:"center",
  },
  prizeCard3:{
    background:"linear-gradient(135deg,#1c0a00,#3b1f00)",
    border:"1px solid #cd7f3266",
    borderRadius:12,padding:"12px 8px",textAlign:"center",
  },

  // ── EMPTY ─────────────────────────────────────────────────────────────────
  empty:{color:"#2d3748",textAlign:"center",padding:"48px 16px",fontSize:14,lineHeight:1.8},

  // ── FORMS ─────────────────────────────────────────────────────────────────
  inp:{
    background:"rgba(5,10,25,0.8)",
    border:"1px solid rgba(255,255,255,0.1)",
    borderRadius:10,padding:"9px 13px",
    color:"#fff",fontSize:14,
    width:"100%",boxSizing:"border-box",outline:"none",
    transition:"border-color 0.2s",
  },
  sel:{
    background:"rgba(5,10,25,0.8)",
    border:"1px solid rgba(255,255,255,0.1)",
    borderRadius:10,padding:"9px 11px",
    color:"#fff",fontSize:13,
    width:"100%",boxSizing:"border-box",outline:"none",
  },
  scoreI:{
    background:"rgba(5,10,25,0.9)",
    border:"1px solid rgba(255,255,255,0.15)",
    borderRadius:10,padding:"8px 2px",
    color:"#e8b923",fontSize:16,fontWeight:900,
    width:46,textAlign:"center",outline:"none",boxSizing:"border-box",
  },
  pinInp:{
    background:"rgba(5,10,25,0.8)",
    border:"1px solid rgba(255,255,255,0.1)",
    borderRadius:8,padding:"7px 11px",
    color:"#fff",fontSize:14,width:95,outline:"none",
  },

  // ── BUTTONS ───────────────────────────────────────────────────────────────
  btnGold:{
    background:"linear-gradient(135deg,#c8102e,#a00d25)",
    border:"none",borderRadius:10,
    padding:"10px 18px",color:"#fff",
    fontWeight:700,fontSize:13,cursor:"pointer",
    whiteSpace:"nowrap",
    boxShadow:"0 4px 14px rgba(200,16,46,0.4)",
    transition:"all 0.2s",
  },
  btnSm:{
    background:"linear-gradient(135deg,#1e3a5f,#1a3255)",
    border:"none",borderRadius:8,
    padding:"6px 12px",color:"#fff",fontSize:12,cursor:"pointer",
  },
  btnGhost:{
    background:"rgba(255,255,255,0.04)",
    border:"1px solid rgba(255,255,255,0.12)",
    borderRadius:8,padding:"5px 12px",
    color:"#718096",fontSize:12,cursor:"pointer",
    transition:"all 0.2s",
  },
  btnAct:{
    background:"rgba(15,30,60,0.8)",
    border:"1px solid rgba(255,255,255,0.08)",
    borderRadius:8,padding:"7px 12px",
    color:"#a0aec0",fontSize:12,cursor:"pointer",whiteSpace:"nowrap",
  },
  backBtn:{
    background:"rgba(255,255,255,0.06)",
    border:"1px solid rgba(255,255,255,0.1)",
    borderRadius:10,padding:"7px 14px",
    color:"#a0aec0",cursor:"pointer",fontSize:13,
  },

  // ── CHIPS ─────────────────────────────────────────────────────────────────
  chip:{
    background:"rgba(10,15,30,0.7)",
    border:"1px solid rgba(255,255,255,0.08)",
    borderRadius:16,padding:"5px 11px",
    color:"#4a5568",fontSize:12,cursor:"pointer",
    position:"relative",transition:"all 0.15s",
  },
  chipA:{
    background:"rgba(16,185,129,0.12)",
    borderColor:"#10b98166",color:"#10b981",
  },
  fl:{fontSize:13,color:"#4a5568",marginBottom:5},
};
