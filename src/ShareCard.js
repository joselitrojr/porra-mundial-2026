import { useState, useRef, useEffect } from "react";

const TEMPLATES = [
  { id:"fumada", titulo:"Esta es mi fumada 🚬", sub:"Mi Mundial, mis reglas 👑", img:"/images/fumada.jpg", accent:"#e8b923" },
  { id:"haaland", titulo:"Venid pa'cá, llorones 😂", sub:"Haaland manda. El resto llora.", img:"/images/haaland.jpg", accent:"#60a5fa" },
  { id:"ajedrez", titulo:"Jaque mate al Mundial ♟️", sub:"Una leyenda solo puede quedar en pie", img:"/images/ajedrez.jpg", accent:"#c084fc" },
  { id:"brasil", titulo:"La llamada de la Canarinha 🦇🇧🇷", sub:"La noche llama. Brasil responde.", img:"/images/brasil.jpg", accent:"#4ade80" },
];

export default function ShareCard({ participant, allM, onClose }) {
  const [sel, setSel] = useState(0);
  const [generating, setGenerating] = useState(false);
  const [preview, setPreview] = useState(null);
  const canvasRef = useRef();
  const tpl = TEMPLATES[sel];
  const p = participant;
  const fetM = allM.find(m => m.id === p.fetiche);

  const generate = () => {
    setGenerating(true);
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    const W = 800, H = 1000;
    canvas.width = W; canvas.height = H;

    const img = new window.Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      // ── Background illustration (left half) ──
      ctx.drawImage(img, 0, 0, W * 0.45, H);

      // ── Dark gradient overlay on left ──
      const gradL = ctx.createLinearGradient(W * 0.3, 0, W * 0.45, 0);
      gradL.addColorStop(0, "rgba(8,8,16,0)");
      gradL.addColorStop(1, "rgba(8,8,16,0.98)");
      ctx.fillStyle = gradL;
      ctx.fillRect(0, 0, W * 0.45, H);

      // ── Right panel background ──
      ctx.fillStyle = "rgba(8,8,16,0.97)";
      ctx.fillRect(W * 0.45, 0, W * 0.55, H);

      // ── Top border accent ──
      const gradTop = ctx.createLinearGradient(0, 0, W, 0);
      gradTop.addColorStop(0, "#c8102e");
      gradTop.addColorStop(0.5, tpl.accent);
      gradTop.addColorStop(1, "#c8102e");
      ctx.fillStyle = gradTop;
      ctx.fillRect(0, 0, W, 4);

      // ── Logo area ──
      const rx = W * 0.47;
      ctx.fillStyle = tpl.accent;
      ctx.font = "bold 18px system-ui";
      ctx.fillText("⚽ SÚPER PORRA", rx, 44);
      ctx.fillStyle = "rgba(255,255,255,0.4)";
      ctx.font = "11px system-ui";
      ctx.letterSpacing = "3px";
      ctx.fillText("MUNDIAL 2026", rx, 62);

      // ── Divider ──
      ctx.fillStyle = "rgba(255,255,255,0.08)";
      ctx.fillRect(rx, 72, W * 0.5, 1);

      // ── Title ──
      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 22px system-ui";
      ctx.fillText(tpl.titulo, rx, 104);
      ctx.fillStyle = "rgba(255,255,255,0.5)";
      ctx.font = "13px system-ui";
      ctx.fillText(`${p.name} · Mundial 2026`, rx, 124);

      // ── Section helper ──
      const sectionLabel = (label, y) => {
        ctx.fillStyle = "rgba(255,255,255,0.3)";
        ctx.font = "bold 10px system-ui";
        ctx.fillText(label.toUpperCase(), rx, y);
      };

      // ── Box helper ──
      const box = (x, y, w, h, accent) => {
        ctx.strokeStyle = accent || "rgba(255,255,255,0.1)";
        ctx.lineWidth = 1;
        roundRect(ctx, x, y, w, h, 8);
        ctx.fillStyle = "rgba(0,0,0,0.3)";
        ctx.fill();
        ctx.stroke();
      };

      // ── PODIO ──
      sectionLabel("Podio final", 158);
      const podioY = 166;
      const bw = (W * 0.5 - 24) / 3;
      [
        { ico:"🏆", label:"Campeón", val:p.pre?.campeon||"-", col:tpl.accent },
        { ico:"🥈", label:"Subcampeón", val:p.pre?.subcampeon||"-", col:"#94a3b8" },
        { ico:"🥉", label:"3er puesto", val:p.pre?.tercero||"-", col:"#cd7f32" },
      ].forEach((item, i) => {
        const bx = rx + i * (bw + 8);
        box(bx, podioY, bw, 68, item.col + "44");
        ctx.font = "16px system-ui"; ctx.fillStyle = "#fff";
        ctx.fillText(item.ico, bx + 8, podioY + 20);
        ctx.font = "9px system-ui"; ctx.fillStyle = "rgba(255,255,255,0.4)";
        ctx.fillText(item.label, bx + 8, podioY + 36);
        ctx.font = "bold 12px system-ui"; ctx.fillStyle = item.col;
        const shortVal = (item.val||"-").replace(/\s\S+$/, "");
        ctx.fillText(shortVal, bx + 8, podioY + 52);
      });

      // ── SEMIFINALISTAS ──
      sectionLabel("Mis 4 semifinalistas", 252);
      const semis = p.pre?.semis || ["","","",""];
      const sw = (W * 0.5 - 16) / 2;
      [0,1,2,3].forEach(i => {
        const sx = rx + (i % 2) * (sw + 8);
        const sy = 260 + Math.floor(i / 2) * 44;
        box(sx, sy, sw, 36);
        ctx.font = "9px system-ui"; ctx.fillStyle = "rgba(255,255,255,0.3)";
        ctx.fillText(`Semi ${i+1}`, sx + 8, sy + 13);
        ctx.font = "bold 12px system-ui"; ctx.fillStyle = "#e2e8f0";
        const s = (semis[i]||"-").replace(/\s\S+$/, "");
        ctx.fillText(s, sx + 8, sy + 28);
      });

      // ── FETICHE ──
      sectionLabel("Partido fetiche ⭐", 362);
      box(rx, 370, W * 0.5 - 8, 52, tpl.accent + "66");
      ctx.font = "9px system-ui"; ctx.fillStyle = tpl.accent;
      ctx.fillText("×5 SI ACIERTO EXACTO · ×3 SI ACIERTO 1X2", rx + 8, 386);
      ctx.font = "bold 14px system-ui"; ctx.fillStyle = "#fff5c0";
      const fetLabel = fetM ? `${fetM.l.replace(/\s\S+$/,"")} vs ${fetM.v.replace(/\s\S+$/,"")}` : "Sin seleccionar";
      ctx.fillText("⭐ " + fetLabel, rx + 8, 408);

      // ── RETOS ──
      sectionLabel("Retos especiales", 442);
      const retos = [
        { ico:"🟥", label:"Primer expulsado", val:p.spc?.expulsado||"-" },
        { ico:"🎩", label:"Primer hat-trick", val:p.spc?.hattrick||"-" },
        { ico:"🚀", label:"Revelación", val:(p.spc?.revelacion||"-").replace(/\s\S+$/,"") },
        { ico:"🎯", label:"Mayor goleada", val:p.spc?.goleada||"-" },
      ];
      retos.forEach((r, i) => {
        const rx2 = rx + (i % 2) * (sw + 8);
        const ry = 450 + Math.floor(i / 2) * 50;
        box(rx2, ry, sw, 42);
        ctx.font = "11px system-ui"; ctx.fillStyle = "#fff";
        ctx.fillText(r.ico, rx2 + 8, ry + 16);
        ctx.font = "9px system-ui"; ctx.fillStyle = "rgba(255,255,255,0.35)";
        ctx.fillText(r.label, rx2 + 24, ry + 16);
        ctx.font = "bold 12px system-ui"; ctx.fillStyle = "#e2e8f0";
        ctx.fillText(r.val.substring(0, 18), rx2 + 8, ry + 33);
      });

      // ── FOOTER ──
      ctx.fillStyle = "rgba(200,16,46,0.2)";
      ctx.fillRect(0, H - 44, W, 44);
      ctx.fillStyle = "rgba(255,255,255,0.3)";
      ctx.font = "10px system-ui";
      ctx.fillText("porra-mundial-2026-ruby.vercel.app", rx, H - 16);
      ctx.fillStyle = tpl.accent;
      ctx.font = "bold 11px system-ui";
      ctx.fillText("LA PORRA MÁS FUMADA 👑", W / 2 - 60, H - 16);
      ctx.fillStyle = "#c8102e";
      ctx.font = "bold 10px system-ui";
      ctx.fillText("MUNDIAL 2026 🏆", W - 110, H - 16);

      const url = canvas.toDataURL("image/jpeg", 0.9);
      setPreview(url);
      setGenerating(false);
    };
    img.src = tpl.img;
  };

  const download = () => {
    const a = document.createElement("a");
    a.href = preview;
    a.download = `fumada-${p.name.replace(/\s/g,"_")}.jpg`;
    a.click();
  };

  const share = async () => {
    if (!preview) return;
    try {
      const blob = await (await fetch(preview)).blob();
      const file = new File([blob], "mi-fumada.jpg", { type: "image/jpeg" });
      if (navigator.share && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: "Mi fumada del Mundial 2026" });
      } else {
        download();
      }
    } catch { download(); }
  };

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.85)", zIndex:1000, display:"flex", flexDirection:"column", alignItems:"center", overflowY:"auto", padding:"16px" }}>
      <canvas ref={canvasRef} style={{ display:"none" }}/>
      <div style={{ width:"100%", maxWidth:480, background:"linear-gradient(160deg,#0a1628,#0d1f0d)", borderRadius:16, overflow:"hidden" }}>
        {/* Header */}
        <div style={{ padding:"14px 16px", borderBottom:"1px solid rgba(255,255,255,0.08)", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <div style={{ fontWeight:800, fontSize:16, color:"#e8b923" }}>📤 Compartir fumada</div>
          <button onClick={onClose} style={{ background:"rgba(255,255,255,0.08)", border:"none", borderRadius:8, padding:"6px 12px", color:"#fff", cursor:"pointer", fontSize:13 }}>✕</button>
        </div>

        {/* Template selector */}
        <div style={{ padding:"14px 16px" }}>
          <div style={{ fontSize:11, color:"#4a5568", letterSpacing:2, fontWeight:700, textTransform:"uppercase", marginBottom:10 }}>Elige tu plantilla</div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:16 }}>
            {TEMPLATES.map((t, i) => (
              <button key={t.id} onClick={() => { setSel(i); setPreview(null); }}
                style={{ background: sel===i ? `rgba(${t.accent === "#e8b923" ? "232,185,35" : t.accent === "#60a5fa" ? "96,165,250" : t.accent === "#c084fc" ? "192,132,252" : "74,222,128"},0.15)` : "rgba(0,0,0,0.3)",
                  border: `1px solid ${sel===i ? t.accent : "rgba(255,255,255,0.08)"}`,
                  borderRadius:10, padding:"8px", cursor:"pointer", textAlign:"left" }}>
                <img src={t.img} alt={t.titulo} style={{ width:"100%", height:70, objectFit:"cover", borderRadius:6, marginBottom:6 }}/>
                <div style={{ fontSize:10, color: sel===i ? t.accent : "#94a3b8", fontWeight:600, lineHeight:1.3 }}>{t.titulo}</div>
              </button>
            ))}
          </div>

          {/* Preview */}
          {preview ? (
            <div style={{ marginBottom:14 }}>
              <img src={preview} alt="preview" style={{ width:"100%", borderRadius:12, border:"1px solid rgba(255,255,255,0.1)" }}/>
            </div>
          ) : (
            <div style={{ background:"rgba(0,0,0,0.3)", borderRadius:12, height:120, display:"flex", alignItems:"center", justifyContent:"center", marginBottom:14, border:"1px solid rgba(255,255,255,0.06)" }}>
              <div style={{ textAlign:"center", color:"#2d3748", fontSize:13 }}>
                <div style={{ fontSize:24, marginBottom:4 }}>🖼</div>
                Genera la imagen para previsualizarla
              </div>
            </div>
          )}

          {/* Buttons */}
          <div style={{ display:"flex", gap:8 }}>
            <button onClick={generate} disabled={generating}
              style={{ flex:1, background:"linear-gradient(135deg,#b45309,#92400e)", border:"none", borderRadius:10, padding:"12px", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer", opacity:generating?0.6:1 }}>
              {generating ? "Generando…" : preview ? "🔄 Regenerar" : "✨ Generar imagen"}
            </button>
            {preview && (
              <button onClick={share}
                style={{ flex:1, background:"linear-gradient(135deg,#c8102e,#a00d25)", border:"none", borderRadius:10, padding:"12px", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer" }}>
                📤 Compartir
              </button>
            )}
          </div>
          {preview && (
            <button onClick={download} style={{ width:"100%", marginTop:8, background:"rgba(255,255,255,0.06)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"10px", color:"#94a3b8", fontSize:13, cursor:"pointer" }}>
              ⬇️ Descargar
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

// Canvas helper
function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}
