import { useState, useRef } from "react";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

const VERMELHO = "#C41E3A";
const PRETO = "#111111";
const LOGO_B64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAqOklEQVR42u2debyd073/3+t5nj3vfcYcCSGJoYgQoTEllIiZVpCm1FBFtJQfYigtyu1t9XVv70Vd92rQcEvVUEVUJTWWBheVNiEiFVESMpx5z8+wfn+sZ08n+5ycKck5sT6v18ax936e9az9/azvsL7r+xWfjR8v0dDQqApDT4GGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhiaIhoYmiIaGJoiGhgYAlp6CKhCi9ALwPJByeI15c0DK0ksT5IuoTw0lYLkcMpdTxDAMCAQQoZD6b9cdWqTwxyRzOXCczUdmwwDTRFgWBALq7+GyeGiCDAJME5lMgutijhmDtcsuiEQCmcngfvopzsqVkMsjamuGxipqGGDbeMkkRk0N1tixGNtth1FTA8EgYhA1iXRdZDqN19KCt24d7oYNyHQaYZqIaFQRRkpFFk2QbdCcArzWVoIHHkj8ggsITp2KUVtbEpB8HnvJElL33Uf2qacQsZj63tYiiWEgs1mM2lpqLrmE8DHHYI4bp7TcZobX1o77z4/JL15M/tVXyb/1Fu7atYhwWJFlGyOK+Gz8ePlFJ4hMJolfdhmJyy8vmQ7lWsIoxTLSDz1E+/XXbxFh7JbQto1oaKDxgQewdttcy/gIBXOuC9w1a8g++yzpRx/FXroUEQoporjuNmF6fbEJYpp4bW3U3nADsQsvLNnTBV+kq+B5HlgW6QceoP266xB1dVveJzFNvNZW6m67jejMmch8XvkFm9tBr0bCMtLIfJ7MU0+Ruusu7GXLlAY2zaHls2mC9FHQ2tuJHHcc9XPnKgfXNHsWMinVD25ZtH73u2SeeUI_g";

const CATEGORIAS = ["Chaves", "Carteira", "Celular", "Casaco", "Mochila", "Óculos", "Outros"];

// Função Utilitária para Redimensionar e Comprimir Imagens no Cliente
const compressImage = (base64Str, maxWidth = 800, maxHeight = 800, quality = 0.7) => {
  return new Promise((resolve) => {
    const img = new Image();
    img.src = base64Str;
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let width = img.width;
      let height = img.height;

      if (width > height) {
        if (width > maxWidth) {
          height = Math.round((height * maxWidth) / width);
          width = maxWidth;
        }
      } else {
        if (height > maxHeight) {
          width = Math.round((width * maxHeight) / height);
          height = maxHeight;
        }
      }

      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, width, height);

      // Comprime para JPEG com a qualidade de 70% (reduz de 5MB-10MB para 50KB-100KB)
      const compressedBase64 = canvas.toDataURL('image/jpeg', quality);
      resolve(compressedBase64);
    };
  });
};

function Login({ onLogin, usuarios }) {
  const [u, setU] = useState("");
  const [p, setP] = useState("");
  const sub = (e) => {
    e.preventDefault();
    const cadastrado = usuarios.find(x => x.usuario === u && x.senha === p && !x.suspenso);
    if (cadastrado) onLogin(cadastrado);
    else alert("Usuário ou senha incorretos (ou conta suspensa)!");
  };
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", minHeight: "100vh", background: PRETO, padding: 20 }}>
      <form onSubmit={sub} style={{ background: "#fff", padding: 30, borderRadius: 16, width: "100%", maxWidth: 360, boxShadow: "0 10px 25px rgba(0,0,0,0.3)", display: "flex", flexDirection: "column", alignItems: "center" }}>
        <img src={LOGO_B64} style={{ width: 80, height: 80, marginBottom: 16, objectFit: "contain" }} alt="CRF" />
        <div style={{ color: VERMELHO, fontWeight: "900", fontSize: 20, letterSpacing: 1, marginBottom: 4 }}>CR FLAMENGO</div>
        <div style={{ fontSize: 13, color: "#666", fontWeight: "bold", marginBottom: 24 }}>ACHADOS E PERDIDOS</div>
        <input value={u} onChange={e => setU(e.target.value)} placeholder="Usuário" required style={{ width: "100%", padding: 12, marginBottom: 12, borderRadius: 8, border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box" }} />
        <input type="password" value={p} onChange={e => setP(e.target.value)} placeholder="Senha" required style={{ width: "100%", padding: 12, marginBottom: 20, borderRadius: 8, border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box" }} />
        <button type="submit" style={{ width: "100%", padding: 14, background: VERMELHO, color: "#fff", border: "none", borderRadius: 8, fontSize: 15, fontWeight: "bold", cursor: "pointer" }}>Entrar</button>
      </form>
    </div>
  );
}

function Home({ usuario, onTela, totalDisponivel }) {
  return (
    <div style={{ padding: 16, maxWidth: 500, margin: "0 auto", paddingBottom: 80 }}>
      <div style={{ background: "#fff", padding: 16, borderRadius: 14, boxShadow: "0 2px 8px rgba(0,0,0,0.05)", marginBottom: 16, display: "flex", alignItems: "center", gap: 12 }}>
        <img src={LOGO_B64} style={{ width: 50, height: 50, objectFit: "contain" }} alt="CRF" />
        <div>
          <div style={{ fontSize: 14, color: "#666" }}>Olá, <strong style={{ color: PRETO }}>{usuario.nome}</strong></div>
          <div style={{ fontSize: 11, background: VERMELHO, color: "#fff", padding: "2px 6px", borderRadius: 4, display: "inline-block", marginTop: 4, fontWeight: "bold", uppercase: "true" }}>{usuario.perfil}</div>
        </div>
      </div>
      <div style={{ background: "linear-gradient(135deg, " + VERMELHO + ", #901428)", color: "#fff", padding: 20, borderRadius: 16, boxShadow: "0 4px 12px rgba(196,30,58,0.3)", marginBottom: 20, textAlign: "center" }}>
        <div style={{ fontSize: 14, opacity: 0.9, fontWeight: "500" }}>Itens sob custódia atual</div>
        <div style={{ fontSize: 48, fontWeight: "900", margin: "4px 0" }}>{totalDisponivel}</div>
        <div style={{ fontSize: 12, opacity: 0.8 }}>Sala de Segurança · Gávea</div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 16 }}>
        <button onClick={() => onTela("cadastrar")} style={{ background: "#fff", border: "1px solid #eee", padding: 20, borderRadius: 14, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 8, boxShadow: "0 2px 6px rgba(0,0,0,0.02)" }}>
          <span style={{ fontSize: 24 }}>➕</span>
          <span style={{ fontSize: 13, fontWeight: "bold", color: PRETO }}>Novo cadastro</span>
        </button>
        <button onClick={() => onTela("consultar")} style={{ background: "#fff", border: "1px solid #eee", padding: 20, borderRadius: 14, cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 8, boxShadow: "0 2px 6px rgba(0,0,0,0.02)" }}>
          <span style={{ fontSize: 24 }}>🔍</span>
          <span style={{ fontSize: 13, fontWeight: "bold", color: PRETO }}>Consultar itens</span>
        </button>
      </div>
      <button onClick={() => onTela("relatorios")} style={{ width: "100%", background: PRETO, color: "#fff", border: "none", padding: 16, borderRadius: 14, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, fontWeight: "bold", fontSize: 14, boxShadow: "0 4px 10px rgba(0,0,0,0.1)" }}>
        <span>📊</span> Painel Gerencial & Relatórios
      </button>
    </div>
  );
}

function Cadastrar({ onSalvar, onVoltar }) {
  const [categoria, setCategoria] = useState("");
  const [descricao, setDescricao] = useState("");
  const [local, setLocal] = useState("");
  const [foto, setFoto] = useState(null);
  const [etapa, setEtapa] = useState("form");
  const [iaCarregando, setIaCarregando] = useState(false);
  const [itemCadastrado, setItemCad] = useState(null);
  const fileRef = useRef(null);

  const analisarIA = async (base64, mimeType) => {
    setIaCarregando(true);
    try {
      const response = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-api-key": "SUA_API_KEY_AQUI",
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true"
        },
        body: JSON.stringify({
          model: "claude-3-5-sonnet-20241022",
          max_tokens: 150,
          messages: [{
            role: "user",
            content: [
              { type: "image", source: { type: "base64", media_type: mimeType || "image/jpeg", data: base64 } },
              { type: "text", text: "Analise a imagem e identifique o objeto. Responda SOMENTE com JSON puro: {\"categoria\":\"CATEGORIA\",\"descricao\":\"DESCRICAO\"}. Categorias: " + CATEGORIAS.join(", ") + ". Descricao em português, max 80 chars." }
            ]
          }]
        })
      });
      const data = await response.json();
      const txt = data.content[0].text;
      const obj = JSON.parse(txt.substring(txt.indexOf("{"), txt.lastIndexOf("}") + 1));
      if (obj.categoria && CATEGORIAS.includes(obj.categoria)) setCategoria(obj.categoria);
      if (obj.descricao) setDescricao(obj.descricao);
    } catch (e) {
      console.error("Erro na IA:", e);
    } finally {
      setIaCarregando(false);
    }
  };

  const onFoto = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const r = new FileReader();
    r.onload = async (ev) => { 
      // COMPRESSÃO EM TEMPO REAL ATIVA AQUI
      const compressed = await compressImage(ev.target.result); 
      setFoto(compressed); 
      await analisarIA(compressed.split(",")[1], f.type); 
    };
    r.readAsDataURL(f);
  };

  const cadastrar = () => {
    if (!categoria || !descricao || !local) { alert("Preencha todos os campos!"); return; }
    const novo = { id: "A" + String(Math.floor(Math.random() * 900) + 100), categoria, descricao, local, data: new Date().toISOString().split("T")[0], status: "disponivel", foto };
    onSalvar(novo); setItemCad(novo); setEtapa("etiqueta");
  };

  if (etapa === "camera") return (
    <div style={{ position: "fixed", inset: 0, background: "#000", zIndex: 100, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "space-between", padding: "40px 20px" }}>
      <div style={{ color: "#fff", fontSize: 14, fontWeight: "bold" }}>Alinhe o objeto na câmera</div>
      <div style={{ width: "100%", maxWidth: 320, height: 320, border: "2px dashed #fff", borderRadius: 20, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <span style={{ color: "#fff", fontSize: 12, opacity: 0.5 }}>Visualização da Câmera</span>
      </div>
      <div style={{ display: "flex", gap: 20, width: "100%", maxWidth: 320 }}>
        <button onClick={() => setEtapa("form")} style={{ flex: 1, padding: 14, background: "rgba(255,255,255,0.2)", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: "bold", cursor: "pointer" }}>Cancelar</button>
        <button onClick={() => { setFoto("data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="); setEtapa("form"); }} style={{ flex: 1, padding: 14, background: VERMELHO, color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: "bold", cursor: "pointer" }}>Tirar Foto</button>
      </div>
    </div>
  );

  if (etapa === "etiqueta" && itemCadastrado) return (
    <div style={{ padding: 16, maxWidth: 400, margin: "0 auto", textAlign: "center" }}>
      <div style={{ background: "#fff", padding: 24, borderRadius: 16, boxShadow: "0 4px 15px rgba(0,0,0,0.05)", marginBottom: 16, display: "flex", flexDirection: "column", alignItems: "center" }}>
        <div style={{ color: "#4CAF50", fontSize: 40, marginBottom: 8 }}>✅</div>
        <div style={{ fontWeight: "bold", fontSize: 18, color: PRETO, marginBottom: 4 }}>Item cadastrado!</div>
        <div style={{ fontSize: 13, color: "#666", marginBottom: 20 }}>Cole a etiqueta física no objeto.</div>
        <div style={{ border: "2px solid " + PRETO, padding: 16, borderRadius: 12, width: "100%", boxSizing: "border-box", background: "#fff" }}>
          <div style={{ fontSize: 12, color: "#666", fontWeight: "bold", letterSpacing: 1 }}>CR FLAMENGO · A&P</div>
          <div style={{ fontSize: 32, fontMone: "true", fontWeight: "900", color: VERMELHO, margin: "8px 0" }}>{itemCadastrado.id}</div>
          <div style={{ fontSize: 13, fontWeight: "bold", color: PRETO, marginBottom: 4 }}>{itemCadastrado.categoria}</div>
          <div style={{ fontSize: 12, color: "#555" }}>{itemCadastrado.descricao}</div>
          <div style={{ fontSize: 10, color: "#999", marginTop: 12 }}>{itemCadastrado.data} · Sala Seg.</div>
        </div>
      </div>
      <button onClick={onVoltar} style={{ width: "100%", padding: 14, background: PRETO, color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: "bold", cursor: "pointer" }}>Voltar ao início</button>
    </div>
  );

  return (
    <div style={{ padding: 16, maxWidth: 450, margin: "0 auto" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
        <button onClick={onVoltar} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", padding: 4 }}>⬅️</button>
        <h2 style={{ fontSize: 18, fontWeight: "bold", color: PRETO }}>Novo cadastro</h2>
      </div>
      <div style={{ background: "#fff", padding: 16, borderRadius: 16, boxShadow: "0 2px 8px rgba(0,0,0,0.04)", display: "flex", flexDirection: "column", gap: 14 }}>
        <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={onFoto} />
        {foto ? (
          <div style={{ position: "relative" }}>
            <img src={foto} style={{ width: "100%", height: 180, objectFit: "cover", borderRadius: 12 }} alt="Objeto" />
            <button onClick={() => setFoto(null)} style={{ position: "absolute", top: 8, right: 8, background: "rgba(0,0,0,0.6)", color: "#fff", border: "none", width: 28, height: 28, borderRadius: 14, cursor: "pointer", fontWeight: "bold" }}>X</button>
            <div style={{ fontSize: 11, color: "#4CAF50", fontWeight: "bold", mt: 4 }}>✓ Foto compactada automaticamente (ideal para nuvem).</div>
          </div>
        ) : (
          <div onClick={() => fileRef.current?.click()} style={{ width: "100%", height: 140, background: "#f9f9f9", border: "2px dashed #ddd", borderRadius: 12, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 6, cursor: "pointer" }}>
            <span style={{ fontSize: 28 }}>📷</span>
            <span style={{ fontSize: 13, color: "#666", fontWeight: "bold" }}>Tirar foto do objeto</span>
            <span style={{ fontSize: 10, color: "#999" }}>Compactação inteligente ativa</span>
          </div>
        )}
        {iaCarregando && <div style={{ background: "#FFF8E1", color: "#B78103", padding: 10, borderRadius: 8, fontSize: 12, fontWeight: "bold", textAlign: "center" }}>🤖 Claude analisando imagem e preenchendo campos...</div>}
        <div>
          <label style={{ fontSize: 12, fontWeight: "bold", color: "#555" }}>Categoria</label>
          <select value={categoria} onChange={e => setCategoria(e.target.value)} style={{ width: "100%", padding: 12, marginTop: 4, borderRadius: 8, border: "1px solid #ddd", fontSize: 14, background: "#fff" }}>
            <option value="">Selecione...</option>
            {CATEGORIAS.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: 12, fontWeight: "bold", color: "#555" }}>Descrição resumida</label>
          <input value={descricao} onChange={e => setDescricao(e.target.value)} placeholder="Ex: Capinha preta, arranhado no canto" style={{ width: "100%", padding: 12, marginTop: 4, borderRadius: 8, border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box" }} />
        </div>
        <div>
          <label style={{ fontSize: 12, fontWeight: "bold", color: "#555" }}>Local onde foi encontrado</label>
          <input value={local} onChange={e => setLocal(e.target.value)} placeholder="Ex: Arquibancada Oeste, Setor 3" style={{ width: "100%", padding: 12, marginTop: 4, borderRadius: 8, border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box" }} />
        </div>
        <button onClick={cadastrar} style={{ width: "100%", padding: 14, background: VERMELHO, color: "#fff", border: "none", borderRadius: 10, fontSize: 15, fontWeight: "bold", cursor: "pointer", marginTop: 6 }}>Concluir Cadastro</button>
      </div>
    </div>
  );
}

function Consultar({ itens, onVerItem, onVoltar }) {
  const [busca, setBusca] = useState("");
  const [cat, setCat] = useState("");
  const [st, setSt] = useState("disponivel");
  const filtrados = itens.filter(i => {
    if (i.excluido) return false;
    const mB = i.id.toLowerCase().includes(busca.toLowerCase()) || i.descricao.toLowerCase().includes(busca.toLowerCase()) || i.local.toLowerCase().includes(busca.toLowerCase());
    const mC = !cat || i.categoria === cat;
    const mS = !st || i.status === st;
    return mB && mC && mS;
  });
  return (
    <div style={{ padding: 16, maxWidth: 500, margin: "0 auto" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <button onClick={onVoltar} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", padding: 4 }}>⬅️</button>
        <h2 style={{ fontSize: 18, fontWeight: "bold", color: PRETO }}>Consultar itens</h2>
      </div>
      <input value={busca} onChange={e => setBusca(e.target.value)} placeholder="🔍 Buscar por ID, descrição ou local..." style={{ width: "100%", padding: 12, marginBottom: 10, borderRadius: 10, border: "1px solid #ddd", fontSize: 14, boxSizing: "border-box", background: "#fff" }} />
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <select value={cat} onChange={e => setCat(e.target.value)} style={{ flex: 1, padding: 10, borderRadius: 8, border: "1px solid #ddd", fontSize: 13, background: "#fff" }}>
          <option value="">Todas categorias</option>
          {CATEGORIAS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
        <select value={st} onChange={e => setSt(e.target.value)} style={{ flex: 1, padding: 10, borderRadius: 8, border: "1px solid #ddd", fontSize: 13, background: "#fff" }}>
          <option value="">Todos status</option>
          <option value="disponivel">Disponível</option>
          <option value="entregue">Entregue</option>
        </select>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {filtrados.map(item => (
          <div key={item.id} onClick={() => onVerItem(item)} style={{ background: "#fff", padding: 12, borderRadius: 12, border: "1px solid #eee", cursor: "pointer", display: "flex", alignItems: "center", gap: 12, boxShadow: "0 2px 4px rgba(0,0,0,0.01)" }}>
            {item.foto ? <img src={item.foto} style={{ width: 50, height: 50, borderRadius: 8, objectFit: "cover" }} alt="" /> : <div style={{ width: 50, height: 50, borderRadius: 8, background: "#f5f5f5", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>📦</div>}
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: "900", color: VERMELHO, fontSize: 14, fontFamily: "monospace" }}>{item.id}</span>
                <span style={{ fontSize: 11, color: item.status === "disponivel" ? "#B78103" : "#4CAF50", fontWeight: "bold" }}>{item.status === "disponivel" ? "⏳ Na sala" : "✅ Entregue"}</span>
              </div>
              <div style={{ fontWeight: "bold", fontSize: 13, color: PRETO, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", marginTop: 2 }}>{item.categoria} · <span style={{ fontWeight: "normal", color: "#555" }}>{item.descricao}</span></div>
              <div style={{ fontSize: 11, color: "#888", display: "flex", justifyContent: "space-between", marginTop: 4 }}><span>📍 {item.local}</span><span>📅 {item.data}</span></div>
            </div>
          </div>
        ))}
        {filtrados.length === 0 && <div style={{ textAlign: "center", color: "#999", padding: 40, fontSize: 14 }}>Nenhum item encontrado.</div>}
      </div>
    </div>
  );
}

function Ficha({ item, onVoltar, onEntregue, onExcluir, perfilLogado }) {
  const [modal, setModal] = useState(false);
  const [canal, setCanal] = useState("email");
  const [nome, setNome] = useState("");
  const [sobrenome, setSobrenome] = useState("");
  const [contato, setContato] = useState("");
  const [fotoAs, setFotoAs] = useState(null);
  const fileRef = useRef(null);

  return (
    <div style={{ padding: 16, maxWidth: 450, margin: "0 auto", paddingBottom: 100 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
        <button onClick={onVoltar} style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", padding: 4 }}>⬅️</button>
        <h2 style={{ fontSize: 18, fontWeight: "bold", color: PRETO }}>Detalhes do Item</h2>
      </div>

      <div style={{ background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 4px 12px rgba(0,0,0,0.04)" }}>
        {item.foto ? <img src={item.foto} style={{ width: "100%", maxHeight: 220, objectFit: "cover" }} alt="" /> : <div style={{ height: 140, background: "#f5f5f5", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 40 }}>📦</div>}
        <div style={{ padding: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <span style={{ fontSize: 24, fontWeight: "900", color: VERMELHO, fontFamily: "monospace" }}>{item.id}</span>
            <span style={{ background: item.status === "disponivel" ? "#FFF8E1" : "#E8F5E9", color: item.status === "disponivel" ? "#B78103" : "#4CAF50", padding: "4px 10px", borderRadius: 6, fontSize: 12, fontWeight: "bold" }}>{item.status === "disponivel" ? "⏳ Sob custódia" : "✅ Entregue"}</span>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8, fontSize: 14, color: "#333" }}>
            <div><strong>Categoria:</strong> {item.categoria}</div>
            <div><strong>Descrição:</strong> {item.descricao}</div>
            <div><strong>Encontrado em:</strong> {item.local}</div>
            <div><strong>Data de entrada:</strong> {item.data}</div>
          </div>

          {item.status === "entregue" && item.comprovante && (
            <div style={{ marginTop: 16, padding: 12, background: "#f9f9f9", borderRadius: 10, border: "1px dashed #ddd", fontSize: 13 }}>
              <div style={{ fontWeight: "bold", color: "#4CAF50", marginBottom: 6 }}>📋 Dados da Devolução:</div>
              <div><strong>Retirado por:</strong> {item.comprovante.nome}</div>
              <div><strong>Canal de aviso:</strong> {item.comprovante.canal} ({item.comprovante.contato})</div>
              <div><strong>Data de entrega:</strong> {item.entregueEm || item.data}</div>
              {item.fotoAssinatura && (
                <div style={{ marginTop: 8 }}>
                  <div style={{ fontWeight: "bold", fontSize: 11, color: "#666", mb: 4 }}>Foto do sócio comprovante (Otimizada):</div>
                  <img src={item.fotoAssinatura} style={{ width: "100%", borderRadius: 10, maxHeight: 180, objectFit: "cover", marginBottom: 14 }} alt="retirada" />
                </div>
              )}
            </div>
          )}
        </div>
      </div>

      {item.status === "disponivel" && (
        <div style={{ marginTop: 16, display: "flex", flexDirection: "column", gap: 8 }}>
          <button onClick={() => setModal(true)} style={{ width: "100%", padding: 14, background: "#4CAF50", color: "#fff", border: "none", borderRadius: 12, fontSize: 15, fontWeight: "bold", cursor: "pointer", boxShadow: "0 4px 10px rgba(76,175,80,0.2)" }}>🤝 Registrar Devolução ao Sócio</button>
          {perfilLogado === "admin" && (
            <button onClick={() => { if (confirm("Deseja mesmo excluir?")) onExcluir(item.id); }} style={{ width: "100%", padding: 10, background: "none", border: "1px solid " + VERMELHO, color: VERMELHO, borderRadius: 12, fontSize: 13, fontWeight: "bold", cursor: "pointer" }}>🗑️ Excluir registro</button>
          )}
        </div>
      )}

      {modal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", backdropBlur: "2px", zIndex: 200, display: "flex", alignItems: "flex-end" }}>
          <div style={{ background: "#fff", borderRadius: "20px 20px 0 0", padding: 20, width: "100%", boxSizing: "border-box", maxHeight: "90vh", overflowY: "auto" }}>
            <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }}
              onChange={e => { 
                const f = e.target.files?.[0]; 
                if (f) { 
                  const r = new FileReader(); 
                  r.onload = async ev => {
                    // COMPRESSÃO EM TEMPO REAL ATIVA NA FOTO DO SÓCIO
                    const compressed = await compressImage(ev.target.result);
                    setFotoAs(compressed);
                  }; 
                  r.readAsDataURL(f); 
                } 
              }} 
            />
            <div style={{ fontWeight: "bold", fontSize: 16, marginBottom: 14 }}>📋 Confirmar devolução</div>
            <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
              <div style={{ flex: 1 }}>
                <label style={{ fontSize: 12, fontWeight: "bold" }}>Nome *</label>
                <input value={nome} onChange={e => setNome(e.target.value)} placeholder="Nome" style={{ width: "100%", padding: 10, marginTop: 4, borderRadius: 8, border: "1px solid " + (nome ? "#4CAF50" : "#ddd"), fontSize: 14, boxSizing: "border-box" }} />
              </div>
              <div style={{ flex: 1 }}>
                <label style={{ fontSize: 12, fontWeight: "bold" }}>Sobrenome *</label>
                <input value={sobrenome} onChange={e => setSobrenome(e.target.value)} placeholder="Sobrenome" style={{ width: "100%", padding: 10, marginTop: 4, borderRadius: 8, border: "1px solid " + (sobrenome ? "#4CAF50" : "#ddd"), fontSize: 14, boxSizing: "border-box" }} />
              </div>
            </div>
            <div style={{ marginBottom: 12 }}>
              <label style={{ fontSize: 12, fontWeight: "bold" }}>Canal de notificação de recibo</label>
              <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
                <button type="button" onClick={() => setCanal("email")} style={{ flex: 1, padding: 10, background: canal === "email" ? PRETO : "#f5f5f5", color: canal === "email" ? "#fff" : "#333", border: "none", borderRadius: 8, fontSize: 13, fontWeight: "bold", cursor: "pointer" }}>📧 E-mail</button>
                <button type="button" onClick={() => setCanal("whatsapp")} style={{ flex: 1, padding: 10, background: canal === "whatsapp" ? "#25D366" : "#f5f5f5", color: canal === "whatsapp" ? "#fff" : "#333", border: "none", borderRadius: 8, fontSize: 13, fontWeight: "bold", cursor: "pointer" }}>💬 WhatsApp</button>
              </div>
            </div>
            <div style={{ marginBottom: 14 }}>
              <label style={{ fontSize: 12, fontWeight: "bold" }}>{canal === "email" ? "Endereço de E-mail *" : "Número com DDD *"}</label>
              <input value={contato} onChange={e => setContato(e.target.value)} placeholder={canal === "email" ? "socio@email.com" : "(21) 99999-9999"} style={{ width: "100%", padding: 10, marginTop: 4, borderRadius: 8, border: "1px solid " + (contato ? "#4CAF50" : "#ddd"), fontSize: 14, boxSizing: "border-box" }} />
            </div>
            {fotoAs ? (
              <div style={{ position: "relative" }}>
                <img src={fotoAs} style={{ width: "100%", height: 120, objectFit: "cover", borderRadius: 8, marginBottom: 10 }} alt="assinatura" />
                <button type="button" onClick={() => setFotoAs(null)} style={{ position: "absolute", top: 6, right: 6, background: "rgba(0,0,0,0.6)", color: "#fff", border: "none", width: 24, height: 24, borderRadius: 12, cursor: "pointer" }}>X</button>
              </div>
            ) : (
              <button onClick={() => fileRef.current?.click()} style={{ width: "100%", padding: 12, background: "#f5f5f5", border: "2px dashed " + (fotoAs ? "#4CAF50" : "#ccc"), borderRadius: 8, cursor: "pointer", fontSize: 13, marginBottom: 10 }}>📷 Foto do sócio com o objeto *</button>
            )}
            {(!nome || !sobrenome || !fotoAs) && (
              <div style={{ background: "#FFEBEE", color: VERMELHO, padding: 8, borderRadius: 6, fontSize: 11, fontWeight: "bold", marginBottom: 10, textAlign: "center" }}>
                ⚠️ Obrigatório: {[!nome && "Nome", !sobrenome && "Sobrenome", !fotoAs && "Foto"].filter(Boolean).join(" · ")}
              </div>
            )}
            <button onClick={() => {
              if (!nome || !sobrenome || !contato) { alert("⚠️ Preencha os campos obrigatórios!"); return; }
              if (!fotoAs) { alert("⚠️ A foto é obrigatória!"); return; }
              onEntregue(item.id, { canal, contato, nome: (nome + " " + sobrenome).trim() }, fotoAs);
            }} style={{ width: "100%", padding: 14, background: (!nome || !sobrenome || !fotoAs) ? "#ccc" : VERMELHO, color: "#fff", border: "none", borderRadius: 12, fontSize: 15, fontWeight: "bold", cursor: "pointer", marginBottom: 8 }}>✅ Confirmar e enviar</button>
            <button onClick={() => setModal(false)} style={{ width: "100%", padding: 10, background: "none", border: "none", color: "#666", fontSize: 13, cursor: "pointer" }}>Voltar</button>
          </div>
        </div>
      )}
    </div>
  );
}

function Relatorios({ itens, usuarios, perfilLogado, onCadastrarUsuario, onEditarUsuario, onRestaurar }) {
  const [abaRel, setAbaRel] = useState("graficos");
  const [abaSubUser, setAbaSubUser] = useState("lista");
  const [uNome, setUNome] = useState("");
  const [uUser, setUUser] = useState("");
  const [uSenha, setUSenha] = useState("");
  const [uPerfil, setUPerfil] = useState("operador");
  const [userEditando, setUserEditando] = useState(null);

  const dadosCat = CATEGORIAS.map(c => {
    return { name: c, value: itens.filter(i => i.categoria === c && !i.excluido).length };
  }).filter(d => d.value > 0);

  const CORES = ["#C41E3A", "#111111", "#4CAF50", "#FFCC00", "#2196F3", "#9C27B0", "#FF9800"];

  const exportarExcel = () => {
    const limpos = itens.map(i => ({
      ID: i.id, Categoria: i.categoria, Descricao: i.descricao, Local: i.local, Data: i.data, Status: i.status, Excluido: i.excluido ? "Sim" : "Não"
    }));
    const ws = XLSX.utils.json_to_sheet(limpos);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Itens");
    XLSX.writeFile(wb, "AP_Flamengo_" + new Date().toISOString().split("T")[0] + ".xlsx");
  };

  const criarUser = (e) => {
    e.preventDefault();
    if (!uNome || !uUser || !uSenha) return;
    if (usuarios.some(x => x.usuario === uUser)) { alert("Esse usuário já existe!"); return; }
    onCadastrarUsuario({ nome: uNome, usuario: uUser, senha: uSenha, perfil: uPerfil, suspenso: false });
    setUNome(""); setUUser(""); setUSenha(""); setUPerfil("operador"); setAbaSubUser("lista");
  };

  return (
    <div style={{ padding: 16, maxWidth: 600, margin: "0 auto", paddingBottom: 100 }}>
      <div style={{ display: "flex", gap: 6, marginBottom: 16, background: "#eee", padding: 4, borderRadius: 10 }}>
        <button onClick={() => setAbaRel("graficos")} style={{ flex: 1, padding: 10, background: abaRel === "graficos" ? "#fff" : "none", border: "none", borderRadius: 8, fontSize: 13, fontWeight: "bold", cursor: "pointer", color: PRETO }}>📊 Métricas</button>
        <button onClick={() => setAbaRel("usuarios")} style={{ flex: 1, padding: 10, background: abaRel === "usuarios" ? "#fff" : "none", border: "none", borderRadius: 8, fontSize: 13, fontWeight: "bold", cursor: "pointer", color: PRETO }}>👥 Equipe</button>
        {perfilLogado === "admin" && <button onClick={() => setAbaRel("lixeira")} style={{ flex: 1, padding: 10, background: abaRel === "lixeira" ? "#fff" : "none", border: "none", borderRadius: 8, fontSize: 13, fontWeight: "bold", cursor: "pointer", color: PRETO }}>🗑️ Lixeira</button>}
      </div>

      {abaRel === "graficos" && (
        <div style={{ background: "#fff", padding: 16, borderRadius: 16, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
          <div style={{ fontWeight: "bold", fontSize: 15, marginBottom: 12, color: PRETO }}>Distribuição por Categoria</div>
          <div style={{ height: 200, width: "100%" }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={dadosCat} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={65} label>
                  {dadosCat.map((entry, index) => <Cell key={`cell-${index}`} fill={CORES[index % CORES.length]} />)}
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <button onClick={exportarExcel} style={{ width: "100%", padding: 14, background: "#4CAF50", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: "bold", cursor: "pointer", marginTop: 16 }}>📥 Exportar Base de Dados (Excel)</button>
        </div>
      )}

      {abaRel === "usuarios" && (
        <div style={{ background: "#fff", padding: 16, borderRadius: 16, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
          {perfilLogado === "admin" && (
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <button onClick={() => { setAbaSubUser("lista"); setUserEditando(null); }} style={{ padding: "6px 12px", background: abaSubUser === "lista" ? PRETO : "#f5f5f5", color: abaSubUser === "lista" ? "#fff" : "#333", border: "none", borderRadius: 6, fontSize: 12, fontWeight: "bold", cursor: "pointer" }}>Lista</button>
              <button onClick={() => setAbaSubUser("novo")} style={{ padding: "6px 12px", background: abaSubUser === "novo" ? PRETO : "#f5f5f5", color: abaSubUser === "novo" ? "#fff" : "#333", border: "none", borderRadius: 6, fontSize: 12, fontWeight: "bold", cursor: "pointer" }}>➕ Novo Operador</button>
            </div>
          )}

          {abaSubUser === "lista" && (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {usuarios.map(x => (
                <div key={x.usuario} style={{ padding: 12, border: "1px solid #eee", borderRadius: 10, display: "flex", justifyContent: "space-between", alignItems: "center", opacity: x.suspenso ? 0.5 : 1 }}>
                  <div>
                    <div style={{ fontWeight: "bold", fontSize: 14 }}>{x.nome} <span style={{ fontSize: 10, color: "#666", background: "#eee", padding: "1px 4px", borderRadius: 3 }}>{x.perfil}</span></div>
                    <div style={{ fontSize: 12, color: "#777" }}>u: {x.usuario} {x.suspenso && <strong style={{ color: VERMELHO }}>· SUSPENSO</strong>}</div>
                  </div>
                  {perfilLogado === "admin" && x.usuario !== "master" && (
                    <button onClick={() => { setUserEditando(x); setAbaSubUser("editar"); setUNome(x.nome); setUPerfil(x.perfil); }} style={{ padding: "6px 10px", background: "#f5f5f5", border: "1px solid #ddd", borderRadius: 6, fontSize: 11, cursor: "pointer" }}>⚙️ Gerenciar</button>
                  )}
                </div>
              ))}
            </div>
          )}

          {abaSubUser === "novo" && (
            <form onSubmit={criarUser} style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <input value={uNome} onChange={e => setUNome(e.target.value)} placeholder="Nome completo" required style={{ padding: 10, borderRadius: 6, border: "1px solid #ddd" }} />
              <input value={uUser} onChange={e => setUUser(e.target.value)} placeholder="Nome de usuário (login)" required style={{ padding: 10, borderRadius: 6, border: "1px solid #ddd" }} />
              <input type="password" value={uSenha} onChange={e => setUSenha(e.target.value)} placeholder="Senha de acesso" required style={{ padding: 10, borderRadius: 6, border: "1px solid #ddd" }} />
              <select value={uPerfil} onChange={e => setUPerfil(e.target.value)} style={{ padding: 10, borderRadius: 6, border: "1px solid #ddd", background: "#fff" }}>
                <option value="operador">Operador (Apenas Cadastro/Consulta)</option>
                <option value="admin">Administrador (Total)</option>
              </select>
              <button type="submit" style={{ padding: 12, background: VERMELHO, color: "#fff", border: "none", borderRadius: 8, fontWeight: "bold", cursor: "pointer" }}>Salvar Usuário</button>
            </form>
          )}

          {abaSubUser === "editar" && userEditando && (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <div style={{ fontWeight: "bold", fontSize: 14 }}>Modificar: {userEditando.usuario}</div>
              <input value={uNome} onChange={e => setUNome(e.target.value)} placeholder="Nome completo" style={{ padding: 10, borderRadius: 6, border: "1px solid #ddd" }} />
              <select value={uPerfil} onChange={e => setUPerfil(e.target.value)} style={{ padding: 10, borderRadius: 6, border: "1px solid #ddd", background: "#fff" }}>
                <option value="operador">Operador</option>
                <option value="admin">Administrador</option>
              </select>
              <div style={{ display: "flex", gap: 8 }}>
                <button onClick={() => { onEditarUsuario({ ...userEditando, nome: uNome, perfil: uPerfil }); setAbaSubUser("lista"); }} style={{ flex: 2, padding: 10, background: PRETO, color: "#fff", border: "none", borderRadius: 6, fontWeight: "bold", cursor: "pointer" }}>Atualizar dados</button>
                <button onClick={() => { onEditarUsuario({ ...userEditando, suspenso: !userEditando.suspenso }); setAbaSubUser("lista"); }} style={{ flex: 1, padding: 10, background: VERMELHO, color: "#fff", border: "none", borderRadius: 6, fontWeight: "bold", cursor: "pointer" }}>{userEditando.suspenso ? "Ativar" : "Suspender"}</button>
              </div>
            </div>
          )}
        </div>
      )}

      {abaRel === "lixeira" && (
        <div style={{ background: "#fff", padding: 16, borderRadius: 16, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
          <div style={{ fontWeight: "bold", fontSize: 14, marginBottom: 10, color: "#666" }}>Registros Removidos</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {itens.filter(i => i.excluido).map(i => (
              <div key={i.id} style={{ padding: 10, border: "1px solid #ffcdd2", background: "#ffebee", borderRadius: 8, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontWeight: "bold", fontSize: 13 }}>{i.id} - {i.categoria}</div>
                  <div style={{ fontSize: 11, color: "#c62828" }}>Removido por: {i.excluidoPor} ({i.excluidoEm})</div>
                </div>
                <button onClick={() => onRestaurar(i.id)} style={{ padding: "4px 8px", background: "#fff", border: "1px solid #c62828", color: "#c62828", borderRadius: 4, fontSize: 11, cursor: "pointer", fontWeight: "bold" }}>Restaurar</button>
              </div>
            ))}
            {itens.filter(i => i.excluido).length === 0 && <div style={{ textAlign: "center", color: "#999", padding: 20, fontSize: 13 }}>Lixeira vazia.</div>}
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [tela, setTela] = useState("login");
  const [usuarioLogado, setUsuarioLogado] = useState(null);
  const [itemSel, setItemSel] = useState(null);

  const [usuarios, setUsuarios] = useState([
    { nome: "Master Administrador", usuario: "master", senha: "123", perfil: "admin", suspenso: false },
    { nome: "Segurança Portaria", usuario: "fla01", senha: "123", perfil: "operador", suspenso: false }
  ]);

  const [itens, setItens] = useState([
    { id: "A102", categoria: "Carteira", descricao: "Couro preta com documento de sócio", local: "Gávea - Piscina", data: "2024-03-10", status: "disponivel" },
    { id: "A543", categoria: "Celular", descricao: "iPhone 12 azul, tela trincada", local: "Maracanã - Setor Norte", data: "2024-03-12", status: "disponivel" }
  ]);

  const loginEfetuado = (u) => { setUsuarioLogado(u); setTela("home"); };
  const voltar = () => { setItemSel(null); setTela("consultar"); };

  const totalDisp = itens.filter(i => i.status === "disponivel" && !i.excluido).length;

  return (
    <div style={{ minHeight: "100vh", background: "#f5f5f5", display: "flex", flexDirection: "column" }}>
      {tela !== "login" && (
        <header style={{ background: VERMELHO, color: "#fff", padding: "14px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", boxShadow: "0 2px 10px rgba(0,0,0,0.15)", position: "sticky", top: 0, zIndex: 50 }}>
          <div onClick={() => setTela("home")} style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }}>
            <div style={{ background: PRETO, width: 32, height: 32, borderRadius: 16, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: "900", fontSize: 13, border: "1px solid #ffcc00", color: "#ffcc00" }}>CRF</div>
            <span style={{ fontWeight: "900", fontSize: 15, letterSpacing: 0.5 }}>A&P FLAMENGO</span>
          </div>
          <button onClick={() => { setTela("login"); setUsuarioLogado(null); }} style={{ background: "none", border: "none", color: "#fff", fontSize: 13, cursor: "pointer", opacity: 0.8, fontWeight: "bold" }}>Sair ↩</button>
        </header>
      )}

      <div style={{ flex: 1 }}>
        {tela === "login" && <Login usuarios={usuarios} onLogin={loginEfetuado} />}
        {tela === "home" && usuarioLogado && <Home usuario={usuarioLogado} totalDisponivel={totalDisp} onTela={setTela} />}
        {tela === "cadastrar" && <Cadastrar onSalvar={n => setItens(p => [n, ...p])} onVoltar={() => setTela("home")} />}
        {tela === "consultar" && <Consultar itens={itens} onVoltar={() => setTela("home")} onVerItem={i => { setItemSel(i); setTela("ficha"); }} />}
        {tela === "ficha" && itemSel && (
          <Ficha 
            item={itens.find(i => i.id === itemSel.id) || itemSel} 
            onVoltar={voltar}
            perfilLogado={usuarioLogado?.perfil}
            onExcluir={(id) => { setItens(p => p.map(i => i.id === id ? { ...i, excluido: true, excluidoEm: new Date().toLocaleDateString("pt-BR"), excluidoPor: usuarioLogado?.nome } : i)); voltar(); }}
            onEntregue={(id, comp, fotoAs) => {
              setItens(p => p.map(i => i.id === id ? { ...i, status: "entregue", comprovante: comp, fotoAssinatura: fotoAs, entregueEm: new Date().toLocaleDateString("pt-BR") } : i));
              alert(`✅ Recibo leve de devolução enviado por ${comp.canal === "email" ? "E-mail" : "WhatsApp"}!`);
              voltar();
            }} 
          />
        )}
        {tela === "relatorios" && (
          <Relatorios 
            itens={itens} 
            usuarios={usuarios} 
            perfilLogado={usuarioLogado?.perfil}
            onCadastrarUsuario={u => setUsuarios(p => [...p, u])}
            onEditarUsuario={u => setUsuarios(p => p.map(x => x.usuario === u.usuario ? u : x))}
            onRestaurar={(id) => setItens(p => p.map(i => i.id === id ? { ...i, excluido: false } : i))}
          />
        )}
      </div>

      <footer style={{ background: PRETO, color: "#555", textAlign: "center", padding: 12, fontSize: 10, fontWeight: "bold", borderTop: "1px solid #222" }}>
        © {new Date().getFullYear()} C.R. FLAMENGO · TI & SEGURANÇA PATRIMONIAL
      </footer>
    </div>
  );
}
