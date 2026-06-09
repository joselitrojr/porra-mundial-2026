import { useState, useEffect, useRef, useCallback } from "react";
import { dbGet, dbSet } from "./supabase";

export default function ChatTab({ myName, parts }) {
  const [msgs, setMsgs] = useState([]);
  const [text, setText] = useState("");
  const [sending, setSending] = useState(false);
  const [loading, setLoading] = useState(true);
  const [myNameLocal, setMyNameLocal] = useState(myName || "");
  const [nameSet, setNameSet] = useState(!!myName);
  const bottomRef = useRef();

  const loadMsgs = useCallback(async () => {
    const data = await dbGet("chat");
    if (data) setMsgs(data);
    setLoading(false);
  }, []);

  useEffect(() => { loadMsgs(); }, [loadMsgs]);
  useEffect(() => { const t = setInterval(loadMsgs, 8000); return () => clearInterval(t); }, [loadMsgs]);
  useEffect(() => { bottomRef.current?.scrollIntoView({ behavior: "smooth" }); }, [msgs]);

  const send = async () => {
    if (!text.trim() || !myNameLocal.trim() || sending) return;
    setSending(true);
    const newMsg = {
      id: Date.now().toString(),
      name: myNameLocal.trim(),
      text: text.trim(),
      ts: new Date().toISOString(),
    };
    const updated = [...msgs, newMsg].slice(-200); // max 200 msgs
    await dbSet("chat", updated);
    setMsgs(updated);
    setText("");
    setSending(false);
    sessionStorage.setItem("myName", myNameLocal.trim());
  };

  const fmt = (iso) => {
    const d = new Date(iso);
    return d.toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
  };

  const fmtDay = (iso) => {
    const d = new Date(iso);
    const today = new Date();
    const diff = today.setHours(0,0,0,0) - d.setHours(0,0,0,0);
    if (diff === 0) return "Hoy";
    if (diff === 86400000) return "Ayer";
    return new Date(iso).toLocaleDateString("es-ES", { day:"numeric", month:"short" });
  };

  // Group messages by day
  const grouped = [];
  let lastDay = null;
  msgs.forEach(m => {
    const day = new Date(m.ts).toDateString();
    if (day !== lastDay) { grouped.push({ type:"day", label: fmtDay(m.ts) }); lastDay = day; }
    grouped.push({ type:"msg", ...m });
  });

  const avatarColor = (name) => {
    const colors = ["#e8b923","#10b981","#60a5fa","#f472b6","#fb923c","#a78bfa","#34d399","#f87171"];
    let hash = 0;
    for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
    return colors[Math.abs(hash) % colors.length];
  };

  if (!nameSet) return (
    <div style={{ padding:"24px 16px" }}>
      <div style={{ background:"rgba(0,0,0,0.4)", border:"1px solid rgba(255,255,255,0.08)", borderRadius:14, padding:"20px", textAlign:"center" }}>
        <div style={{ fontSize:32, marginBottom:10 }}>💬</div>
        <div style={{ fontWeight:700, fontSize:16, color:"#fff", marginBottom:6 }}>Únete al chat</div>
        <div style={{ fontSize:13, color:"#4a5568", marginBottom:16 }}>¿Con qué nombre quieres aparecer?</div>
        <input
          style={{ background:"rgba(5,10,25,0.8)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:10, padding:"10px 14px", color:"#fff", fontSize:14, width:"100%", boxSizing:"border-box", outline:"none", marginBottom:12 }}
          placeholder="Tu nombre…"
          value={myNameLocal}
          onChange={e => setMyNameLocal(e.target.value)}
          onKeyDown={e => e.key === "Enter" && myNameLocal.trim() && setNameSet(true)}
        />
        <button
          onClick={() => myNameLocal.trim() && setNameSet(true)}
          style={{ background:"linear-gradient(135deg,#b45309,#92400e)", border:"none", borderRadius:10, padding:"11px 24px", color:"#fff", fontWeight:700, fontSize:14, cursor:"pointer", width:"100%" }}>
          Entrar al chat →
        </button>
      </div>
    </div>
  );

  return (
    <div style={{ display:"flex", flexDirection:"column", height:"calc(100vh - 160px)" }}>
      {/* Messages */}
      <div style={{ flex:1, overflowY:"auto", padding:"12px 16px", display:"flex", flexDirection:"column", gap:2 }}>
        {loading ? (
          <div style={{ textAlign:"center", color:"#2d3748", padding:"40px 0", fontSize:13 }}>Cargando mensajes…</div>
        ) : msgs.length === 0 ? (
          <div style={{ textAlign:"center", color:"#2d3748", padding:"48px 16px", fontSize:14, lineHeight:1.8 }}>
            <div style={{ fontSize:32, marginBottom:8 }}>💬</div>
            <div>Sé el primero en escribir algo</div>
            <div style={{ fontSize:12, color:"#1a1a1a", marginTop:4 }}>¡El Mundial empieza el 11 de junio!</div>
          </div>
        ) : (
          grouped.map((item, i) => {
            if (item.type === "day") return (
              <div key={i} style={{ textAlign:"center", margin:"10px 0 6px", fontSize:11, color:"#2d3748", fontWeight:600 }}>
                <span style={{ background:"rgba(255,255,255,0.05)", padding:"3px 12px", borderRadius:10 }}>{item.label}</span>
              </div>
            );
            const isMe = item.name === myNameLocal;
            const col = avatarColor(item.name);
            return (
              <div key={item.id} style={{ display:"flex", gap:8, alignItems:"flex-end", flexDirection: isMe ? "row-reverse" : "row", marginBottom:4 }}>
                {!isMe && (
                  <div style={{ width:28, height:28, borderRadius:"50%", background:col, display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700, color:"#000", flexShrink:0 }}>
                    {item.name.charAt(0).toUpperCase()}
                  </div>
                )}
                <div style={{ maxWidth:"75%" }}>
                  {!isMe && <div style={{ fontSize:11, color:col, fontWeight:600, marginBottom:3, marginLeft:4 }}>{item.name}</div>}
                  <div style={{
                    background: isMe ? "linear-gradient(135deg,#b45309,#92400e)" : "rgba(255,255,255,0.08)",
                    borderRadius: isMe ? "14px 14px 4px 14px" : "14px 14px 14px 4px",
                    padding:"9px 13px",
                    fontSize:14,
                    color:"#fff",
                    lineHeight:1.5,
                    wordBreak:"break-word",
                  }}>
                    {item.text}
                  </div>
                  <div style={{ fontSize:10, color:"#2d3748", marginTop:3, textAlign: isMe ? "right" : "left", paddingLeft:4 }}>{fmt(item.ts)}</div>
                </div>
              </div>
            );
          })
        )}
        <div ref={bottomRef}/>
      </div>

      {/* Input */}
      <div style={{ padding:"10px 14px", borderTop:"1px solid rgba(255,255,255,0.07)", background:"rgba(0,0,0,0.4)", display:"flex", gap:8, alignItems:"center" }}>
        <div style={{ width:28, height:28, borderRadius:"50%", background:avatarColor(myNameLocal), display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700, color:"#000", flexShrink:0 }}>
          {myNameLocal.charAt(0).toUpperCase()}
        </div>
        <input
          style={{ flex:1, background:"rgba(5,10,25,0.8)", border:"1px solid rgba(255,255,255,0.1)", borderRadius:20, padding:"9px 14px", color:"#fff", fontSize:14, outline:"none" }}
          placeholder="Escribe algo…"
          value={text}
          onChange={e => setText(e.target.value)}
          onKeyDown={e => e.key === "Enter" && !e.shiftKey && send()}
        />
        <button onClick={send} disabled={!text.trim() || sending}
          style={{ background:"linear-gradient(135deg,#b45309,#92400e)", border:"none", borderRadius:"50%", width:38, height:38, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", opacity:!text.trim()||sending?0.4:1, fontSize:16, flexShrink:0 }}>
          ➤
        </button>
      </div>
    </div>
  );
}
