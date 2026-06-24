import { useState, useRef } from "react";
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";

const VERMELHO = "#C41E3A";
const PRETO = "#111111";
const LOGO_B64 = "data:image/png;base64,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";

const CATEGORIAS = ["Brinquedos e Lazer","Calçados","Documentos e Carteirinhas","Eletrônicos e Acessórios","Equipamentos Esportivos","Garrafas e Copos","Higiene e Uso Pessoal","Roupas de Banho","Roupas Secas"];
const CORES_GRAFICO = ["#C41E3A","#333","#E63950","#555","#FF6B6B","#777","#AA1133","#999","#880022"];

const USUARIOS_INICIAIS = [
  { usuario:"mirna", senha:"1234", nome:"Mirna", email:"mirna@flamengo.com.br", perfil:"admin", ativo:true, criado:"2026-06-01" },
];

const MOCK = [
  { id:"A001", categoria:"Eletrônicos e Acessórios", descricao:"Celular Samsung preto", local:"Setor Norte", data:"2026-06-15", status:"disponivel", foto:null },
  { id:"A002", categoria:"Documentos e Carteirinhas", descricao:"Carteira de identidade", local:"Bilheteria 3", data:"2026-06-16", status:"entregue", foto:null, comprovante:{canal:"whatsapp",contato:"+5521999990000",nome:"João Silva"} },
  { id:"A003", categoria:"Garrafas e Copos", descricao:"Garrafa térmica vermelha", local:"Setor Sul", data:"2026-06-17", status:"disponivel", foto:null },
  { id:"A004", categoria:"Calçados", descricao:"Tênis branco número 42", local:"Vestiário", data:"2026-06-18", status:"disponivel", foto:null },
  { id:"A005", categoria:"Roupas Secas", descricao:"Camiseta do Flamengo", local:"Setor Leste", data:"2026-06-18", status:"entregue", foto:null, comprovante:{canal:"email",contato:"socio@email.com",nome:"Maria Santos"} },
];

function normalizar(str) {
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9 ]/g,"");
}

function exportarExcel(itens) {
  const dados = itens.map(i => ({
    "ID": i.id, "Descrição": i.descricao, "Categoria": i.categoria,
    "Local": i.local, "Data Achado": i.data,
    "Status": i.status==="disponivel"?"Disponível":"Entregue",
    "Recebido por": i.comprovante?.nome||"", "Canal": i.comprovante?.canal||"",
    "Contato": i.comprovante?.contato||"",
    "Data Entrega": i.status==="entregue"?new Date().toLocaleDateString("pt-BR"):"",
  }));
  const ws = XLSX.utils.json_to_sheet(dados);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "A&P Flamengo");
  XLSX.writeFile(wb, "AP_Flamengo_"+new Date().toISOString().split("T")[0]+".xlsx");
}

function Logo({ size=44 }) {
  return (
    <div style={{width:size,height:size,borderRadius:"50%",overflow:"hidden",flexShrink:0,border:"2px solid rgba(255,255,255,0.5)"}}>
      <img src={LOGO_B64} alt="A&P" style={{width:"100%",height:"100%",objectFit:"cover"}} />
    </div>
  );
}

function Login({ onLogin, usuarios, onCadastrarUsuario }) {
  const [tela, setTela] = useState("login");
  const [usuario, setUsuario] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [ver, setVer] = useState(false);
  const [novoNome, setNovoNome] = useState("");
  const [novoUser, setNovoUser] = useState("");
  const [novoEmail, setNovoEmail] = useState("");
  const [novaSenha, setNovaSenha] = useState("");
  const [novaSenha2, setNovaSenha2] = useState("");
  const [novoPerfil, setNovoPerfil] = useState("operador");

  const entrar = () => {
    const u = usuarios.find(x => x.usuario.toLowerCase()===usuario.toLowerCase() && x.senha===senha && x.ativo);
    if (u) { onLogin(u); }
    else { setErro("Usuário ou senha incorretos."); setTimeout(()=>setErro(""),3000); }
  };

  const cadastrar = () => {
    if(!novoNome||!novoUser||!novoEmail||!novaSenha){ alert("Preencha todos os campos!"); return; }
    if(novaSenha!==novaSenha2){ alert("As senhas não coincidem!"); return; }
    if(usuarios.find(x=>x.usuario.toLowerCase()===novoUser.toLowerCase())){ alert("Usuário já existe!"); return; }
    onCadastrarUsuario({usuario:novoUser,senha:novaSenha,nome:novoNome,email:novoEmail,perfil:novoPerfil,ativo:true,criado:new Date().toISOString().split("T")[0]});
    alert("✅ Usuário cadastrado!"); setTela("login");
    setNovoNome(""); setNovoUser(""); setNovoEmail(""); setNovaSenha(""); setNovaSenha2("");
  };

  const inp = (val, set, ph, type="text", extra={}) => (
    <input value={val} onChange={e=>set(e.target.value)} placeholder={ph} type={type}
      style={{width:"100%",padding:"11px 14px",marginTop:6,marginBottom:12,borderRadius:10,border:"1.5px solid #ddd",fontSize:14,boxSizing:"border-box"}} {...extra} />
  );

  return (
    <div style={{minHeight:"100vh",background:"#fff",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{width:110,height:110,borderRadius:"50%",overflow:"hidden",border:"3px solid "+VERMELHO,marginBottom:16}}>
        <img src={LOGO_B64} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="logo" />
      </div>
      <div style={{color:"#666",fontSize:14,marginBottom:28}}>Gestão de Achados & Perdidos</div>
      <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:380,boxSizing:"border-box",boxShadow:"0 4px 24px rgba(0,0,0,0.1)"}}>
        {tela==="login" ? (
          <>
            <div style={{fontWeight:"bold",fontSize:18,color:PRETO,marginBottom:18}}>Entrar</div>
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Usuário</label>
            {inp(usuario, setUsuario, "Digite seu usuário","text",{autoCapitalize:"none"})}
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Senha</label>
            <div style={{position:"relative",marginTop:6,marginBottom:16}}>
              <input value={senha} onChange={e=>setSenha(e.target.value)} type={ver?"text":"password"}
                onKeyDown={e=>e.key==="Enter"&&entrar()} placeholder="Digite sua senha"
                style={{width:"100%",padding:"11px 44px 11px 14px",borderRadius:10,border:"1.5px solid #ddd",fontSize:14,boxSizing:"border-box"}} />
              <button onClick={()=>setVer(!ver)} style={{position:"absolute",right:12,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#aaa"}}>{ver?"🙈":"👁"}</button>
            </div>
            {erro && <div style={{background:"#FFEBEE",border:"1px solid #FFCDD2",borderRadius:8,padding:10,marginBottom:12,fontSize:13,color:"#C62828",textAlign:"center"}}>⚠️ {erro}</div>}
            <button onClick={entrar} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:16,fontWeight:"bold",cursor:"pointer",marginBottom:10}}>Entrar</button>
          </>
        ) : (
          <>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:18}}>
              <button onClick={()=>setTela("login")} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:VERMELHO}}>←</button>
              <div style={{fontWeight:"bold",fontSize:18,color:PRETO}}>Novo usuário</div>
            </div>
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Nome completo</label>
            {inp(novoNome,setNovoNome,"Ex: João Silva")}
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Usuário</label>
            {inp(novoUser,setNovoUser,"Ex: joao.silva","text",{autoCapitalize:"none"})}
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>E-mail</label>
            {inp(novoEmail,setNovoEmail,"email@exemplo.com","email")}
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Perfil</label>
            <select value={novoPerfil} onChange={e=>setNovoPerfil(e.target.value)}
              style={{width:"100%",padding:"11px 14px",marginTop:6,marginBottom:12,borderRadius:10,border:"1.5px solid #ddd",fontSize:14,background:"#fff"}}>
              <option value="operador">Operador — acesso padrão</option>
              <option value="admin">Administrador — acesso total</option>
            </select>
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Senha</label>
            {inp(novaSenha,setNovaSenha,"Crie uma senha","password")}
            <label style={{fontSize:13,fontWeight:"600",color:"#555"}}>Confirmar senha</label>
            {inp(novaSenha2,setNovaSenha2,"Repita a senha","password")}
            <button onClick={cadastrar} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>✅ Cadastrar</button>
          </>
        )}
      </div>
      <div style={{color:"#aaa",fontSize:11,marginTop:24}}>© 2026 Clube de Regatas do Flamengo</div>
    </div>
  );
}

function Header({ tela, onVoltar, onReset, nomeLogado, onPerfil }) {
  return (
    <div style={{background:VERMELHO,padding:"14px 16px",display:"flex",alignItems:"center",gap:12,position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 8px rgba(0,0,0,0.2)"}}>
      {tela!=="home"
        ? <button onClick={onVoltar} style={{background:"rgba(255,255,255,0.2)",border:"none",color:"#fff",fontSize:16,cursor:"pointer",borderRadius:"50%",width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center"}}>←</button>
        : <button onClick={onPerfil} style={{background:"none",border:"none",cursor:"pointer",padding:0}}><Logo size={40} /></button>
      }
      <div style={{flex:1}}>
        <div style={{fontWeight:"bold",fontSize:17,color:"#fff"}}>A&P Flamengo</div>
        <div style={{fontSize:12,color:"rgba(255,255,255,0.8)"}}>Logado como {nomeLogado}</div>
      </div>
      {tela==="home"
        ? <button onClick={onReset} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",fontSize:16,cursor:"pointer",borderRadius:"50%",width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center"}}>↺</button>
        : <button onClick={onPerfil} style={{background:"none",border:"none",cursor:"pointer",padding:0}}><Logo size={36} /></button>
      }
    </div>
  );
}

function Home({ itens, onNav }) {
  const disp = itens.filter(i=>i.status==="disponivel").length;
  const entr = itens.filter(i=>i.status==="entregue").length;
  return (
    <div style={{background:"#f5f5f5",minHeight:"100vh"}}>
      <div style={{background:"#fff",display:"flex",borderBottom:"1px solid #eee"}}>
        {[{l:"Total",v:itens.length,c:PRETO},{l:"Disponíveis",v:disp,c:"#4CAF50"},{l:"Entregues",v:entr,c:"#2196F3"}].map((s,i)=>(
          <div key={s.l} style={{flex:1,padding:"16px 8px",textAlign:"center",borderRight:i<2?"1px solid #eee":"none"}}>
            <div style={{fontSize:26,fontWeight:"700",color:s.c}}>{s.v}</div>
            <div style={{fontSize:12,color:"#888"}}>{s.l}</div>
          </div>
        ))}
      </div>
      <div style={{background:"#fff",textAlign:"center",padding:"24px 16px 20px",marginBottom:10}}>
        <div style={{display:"inline-block",width:100,height:100,borderRadius:"50%",overflow:"hidden",border:"3px solid "+VERMELHO,marginBottom:12}}>
          <img src={LOGO_B64} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="logo" />
        </div>
        <div style={{fontSize:22,fontWeight:"bold",color:PRETO}}>Bem-vindo!</div>
        <div style={{color:"#888",fontSize:14,marginTop:4}}>Gestão de Achados & Perdidos</div>
      </div>
      <div style={{background:"#fff",overflow:"hidden"}}>
        {[
          {label:"Cadastrar achados",desc:"Registrar novo item encontrado",tela:"cadastrar",icon:"➕"},
          {label:"Consultar perdidos",desc:"Buscar e filtrar o inventário",tela:"consultar",icon:"🔍"},
          {label:"Relatórios",desc:"Exportar dados e comprovantes",tela:"relatorios",icon:"📊"},
        ].map((b,i,arr)=>(
          <button key={b.tela} onClick={()=>onNav(b.tela)} style={{width:"100%",padding:"16px",background:"#fff",border:"none",borderBottom:i<arr.length-1?"1px solid #f0f0f0":"none",cursor:"pointer",display:"flex",alignItems:"center",gap:14,textAlign:"left"}}>
            <div style={{width:44,height:44,borderRadius:12,background:"#fff0f0",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{b.icon}</div>
            <div style={{flex:1}}>
              <div style={{fontWeight:"600",color:PRETO,fontSize:15}}>{b.label}</div>
              <div style={{color:"#999",fontSize:12,marginTop:2}}>{b.desc}</div>
            </div>
            <div style={{color:"#ccc",fontSize:20}}>›</div>
          </button>
        ))}
      </div>
    </div>
  );
}

function Cadastrar({ onSalvar, onVoltar }) {
  const [etapa, setEtapa] = useState("camera");
  const [foto, setFoto] = useState(null);
  const [categoria, setCategoria] = useState("");
  const [descricao, setDescricao] = useState("");
  const [local, setLocal] = useState("");
  const [iaOk, setIaOk] = useState(false);
  const [erroIA, setErroIA] = useState(null);
  const [itemCad, setItemCad] = useState(null);
  const fileRef = useRef();

  const analisarIA = async (base64, mediaType) => {
    setEtapa("analisando"); setErroIA(null);
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method:"POST",
        headers:{"Content-Type":"application/json","anthropic-dangerous-direct-browser-access":"true"},
        body: JSON.stringify({
          model:"claude-sonnet-4-6", max_tokens:300,
          messages:[{role:"user",content:[
            {type:"image",source:{type:"base64",media_type:mediaType||"image/jpeg",data:base64}},
            {type:"text",text:"Analise a imagem e identifique o objeto. Responda SOMENTE com JSON puro: {\"categoria\":\"CATEGORIA\",\"descricao\":\"DESCRICAO\"}. Categorias: "+CATEGORIAS.join(", ")+". Descricao em português, max 80 chars."}
          ]}]
        })
      });
      if (!res.ok) throw new Error("HTTP "+res.status);
      const d = await res.json();
      const txt = (d.content?.[0]?.text||"").replace(/```json|```/g,"").trim();
      const m = txt.match(/\{[\s\S]*\}/);
      if (!m) throw new Error("JSON inválido");
      const j = JSON.parse(m[0]);
      const cat = CATEGORIAS.find(c=>c===j.categoria)||CATEGORIAS.find(c=>normalizar(c).includes(normalizar(j.categoria||"")));
      setCategoria(cat||""); setDescricao(j.descricao||""); setIaOk(true);
    } catch(e) { setErroIA("IA indisponível. Preencha manualmente."); }
    setEtapa("form");
  };

  const onFoto = (e) => {
    const f = e.target.files?.[0]; if(!f) return;
    const r = new FileReader();
    r.onload = async (ev) => { setFoto(ev.target.result); await analisarIA(ev.target.result.split(",")[1], f.type); };
    r.readAsDataURL(f);
  };

  const cadastrar = () => {
    if(!categoria||!descricao||!local){ alert("Preencha todos os campos!"); return; }
    const novo = {id:"A"+String(Math.floor(Math.random()*900)+100),categoria,descricao,local,data:new Date().toISOString().split("T")[0],status:"disponivel",foto};
    onSalvar(novo); setItemCad(novo); setEtapa("etiqueta");
  };

  if (etapa==="camera") return (
    <div style={{padding:16}}>
      <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={onFoto} />
      <div style={{background:"#000",borderRadius:16,height:240,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:12,marginBottom:16}}>
        <div style={{fontSize:56}}>📷</div>
        <div style={{color:"#aaa",fontSize:13}}>Toque no botão para fotografar</div>
      </div>
      <button onClick={()=>fileRef.current?.click()} style={{width:"100%",padding:16,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:16,fontWeight:"bold",cursor:"pointer"}}>📷 Abrir câmera</button>
    </div>
  );

  if (etapa==="analisando") return (
    <div style={{padding:16}}>
      {foto && <div style={{position:"relative"}}>
        <img src={foto} style={{width:"100%",borderRadius:12,maxHeight:240,objectFit:"cover"}} alt="foto" />
        <div style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.65)",borderRadius:12,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:10}}>
          <div style={{fontSize:28}}>🔍</div>
          <div style={{color:"#fff",fontWeight:"bold",fontSize:14}}>IA identificando o objeto…</div>
        </div>
      </div>}
    </div>
  );

  if (etapa==="form") return (
    <div style={{padding:16}}>
      <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{display:"none"}} onChange={onFoto} />
      {foto && <>
        <img src={foto} style={{width:"100%",borderRadius:12,maxHeight:180,objectFit:"cover",marginBottom:8}} alt="foto" />
        <button onClick={()=>fileRef.current?.click()} style={{width:"100%",marginBottom:12,padding:10,background:"#f5f5f5",border:"1px solid #ddd",borderRadius:8,cursor:"pointer",fontSize:13}}>🔄 Refazer foto</button>
      </>}
      {erroIA && <div style={{background:"#FFF3CD",border:"1px solid #FFC107",borderRadius:8,padding:10,marginBottom:12,fontSize:12,color:"#856404"}}>⚠️ {erroIA}</div>}
      {iaOk && <div style={{background:"#E8F5E9",border:"1px solid #4CAF50",borderRadius:8,padding:10,marginBottom:12,fontSize:13,color:"#2E7D32"}}>✨ Campos preenchidos pela IA</div>}
      <label style={{fontSize:13,fontWeight:"bold"}}>Categoria</label>
      <select value={categoria} onChange={e=>setCategoria(e.target.value)} style={{width:"100%",padding:10,marginTop:4,marginBottom:12,borderRadius:8,border:"2px solid "+(iaOk&&categoria?"#4CAF50":"#ddd"),fontSize:14,background:"#fff"}}>
        <option value="">Selecione...</option>
        {CATEGORIAS.map(c=><option key={c} value={c}>{c}</option>)}
      </select>
      <label style={{fontSize:13,fontWeight:"bold"}}>Detalhes do material</label>
      <textarea value={descricao} onChange={e=>setDescricao(e.target.value)} rows={3}
        style={{width:"100%",padding:10,marginTop:4,marginBottom:12,borderRadius:8,border:"2px solid "+(iaOk&&descricao?"#4CAF50":"#ddd"),fontSize:14,resize:"none",boxSizing:"border-box"}} placeholder="Descreva o objeto..." />
      <label style={{fontSize:13,fontWeight:"bold"}}>Local encontrado</label>
      <input value={local} onChange={e=>setLocal(e.target.value)}
        style={{width:"100%",padding:10,marginTop:4,marginBottom:20,borderRadius:8,border:"2px solid #ddd",fontSize:14,boxSizing:"border-box"}} placeholder="Ex: Setor Norte, Cadeira 12" />
      <button onClick={cadastrar} style={{width:"100%",padding:16,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:16,fontWeight:"bold",cursor:"pointer"}}>✅ Cadastrar item</button>
    </div>
  );

  return (
    <div style={{padding:16,textAlign:"center"}}>
      <div style={{fontSize:50,marginBottom:8}}>✅</div>
      <div style={{fontSize:20,fontWeight:"bold",color:"#4CAF50",marginBottom:4}}>Item cadastrado!</div>
      <div style={{color:"#666",marginBottom:20}}>ID: <strong>{itemCad?.id}</strong></div>
      <div style={{background:"#f5f5f5",borderRadius:12,padding:16,marginBottom:20,border:"1px solid #ddd",textAlign:"left"}}>
        <div style={{fontWeight:"bold",marginBottom:8,textAlign:"center"}}>🏷️ Etiqueta</div>
        <div style={{fontFamily:"monospace",fontSize:12,lineHeight:1.7}}>
          A&P FLAMENGO | ID: {itemCad?.id}<br/>Cat: {itemCad?.categoria}<br/>{itemCad?.descricao}<br/>Local: {itemCad?.local} | {itemCad?.data}
        </div>
        <div style={{textAlign:"center",fontSize:32,marginTop:8}}>▦</div>
      </div>
      <select style={{width:"100%",padding:10,marginBottom:12,borderRadius:8,border:"1px solid #ddd",fontSize:14}}>
        <option>Zebra ZD220</option><option>Zebra ZQ520</option><option>Brother QL-820</option>
      </select>
      <button onClick={onVoltar} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:8}}>🖨️ Imprimir etiqueta</button>
      <button onClick={onVoltar} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:14,cursor:"pointer"}}>Voltar sem imprimir</button>
    </div>
  );
}

function Ficha({ item, onVoltar, onEntregue, onExcluir, perfilLogado }) {
  const [modal, setModal] = useState(false);
  const [canal, setCanal] = useState("email");
  const [contato, setContato] = useState("");
  const [nome, setNome] = useState("");
  const [sobrenome, setSobrenome] = useState("");
  const [fotoAs, setFotoAs] = useState(null);
  const [lightbox, setLightbox] = useState(false);
  const [verComp, setVerComp] = useState(false);
  const [confirmarExclusao, setConfirmarExclusao] = useState(false);
  const fileRef = useRef();

  return (
    <div style={{padding:16}}>
      {lightbox && (
        <div onClick={()=>setLightbox(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.95)",zIndex:500,display:"flex",alignItems:"center",justifyContent:"center"}}>
          {item.foto
            ? <img src={item.foto} style={{maxWidth:"95%",maxHeight:"90vh",borderRadius:12,objectFit:"contain"}} alt="foto" />
            : <div style={{textAlign:"center"}}><div style={{fontSize:100}}>📦</div><div style={{color:"#aaa",marginTop:8}}>Sem foto</div></div>
          }
          <button onClick={()=>setLightbox(false)} style={{position:"absolute",top:20,right:20,background:"rgba(255,255,255,0.2)",border:"none",color:"#fff",fontSize:20,borderRadius:"50%",width:40,height:40,cursor:"pointer"}}>✕</button>
        </div>
      )}
      <div onClick={()=>setLightbox(true)} style={{width:"100%",height:200,borderRadius:16,background:"#f0f0f0",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:16,overflow:"hidden",cursor:"zoom-in",position:"relative"}}>
        {item.foto ? <img src={item.foto} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="item" /> : <span style={{fontSize:60}}>📦</span>}
        <div style={{position:"absolute",bottom:8,right:8,background:"rgba(0,0,0,0.45)",color:"#fff",fontSize:11,padding:"3px 10px",borderRadius:20}}>🔍 Ampliar</div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <div style={{fontSize:17,fontWeight:"bold",flex:1,marginRight:8}}>{item.descricao}</div>
        <span style={{background:item.status==="disponivel"?"#4CAF50":"#2196F3",color:"#fff",fontSize:12,padding:"4px 10px",borderRadius:20,fontWeight:"bold"}}>
          {item.status==="disponivel"?"Disponível":"Entregue"}
        </span>
      </div>
      {[["ID",item.id],["Categoria",item.categoria],["Local",item.local],["Data",item.data]].map(([l,v])=>(
        <div key={l} style={{background:"#f9f9f9",borderRadius:8,padding:12,marginBottom:8,display:"flex",justifyContent:"space-between"}}>
          <span style={{color:"#888",fontSize:13}}>{l}</span>
          <span style={{fontWeight:"bold",fontSize:13}}>{v}</span>
        </div>
      ))}
      {item.status==="disponivel" && (
        <button onClick={()=>setModal(true)} style={{width:"100%",marginTop:12,padding:16,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>✋ Marcar como entregue</button>
      )}
      {item.status==="entregue" && (
        <button onClick={()=>setVerComp(true)} style={{width:"100%",marginTop:12,padding:16,background:"#2196F3",color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>🎟️ Ver comprovante</button>
      )}
      {perfilLogado==="admin" && (
        <button onClick={()=>setConfirmarExclusao(true)} style={{width:"100%",marginTop:10,padding:14,background:"#fff",color:"#C62828",border:"2px solid #C62828",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>
          🗑️ Excluir item
        </button>
      )}

      {/* Modal confirmação exclusão */}
      {confirmarExclusao && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:24}}>
          <div style={{background:"#fff",borderRadius:20,padding:24,width:"100%",maxWidth:360,boxSizing:"border-box"}}>
            <div style={{textAlign:"center",marginBottom:16}}>
              <div style={{fontSize:48,marginBottom:8}}>⚠️</div>
              <div style={{fontWeight:"bold",fontSize:18,color:PRETO,marginBottom:8}}>Excluir item?</div>
              <div style={{fontSize:14,color:"#666",lineHeight:1.5}}>
                O item <strong>"{item.descricao}"</strong> será excluído permanentemente. Esta ação não pode ser desfeita.
              </div>
            </div>
            <button onClick={()=>{ onExcluir(item.id); onVoltar(); }} style={{width:"100%",padding:14,background:"#C62828",color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:10}}>
              🗑️ Sim, excluir
            </button>
            <button onClick={()=>setConfirmarExclusao(false)} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:15,cursor:"pointer"}}>
              Cancelar
            </button>
          </div>
        </div>
      )}

      {verComp && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:300,display:"flex",alignItems:"flex-end"}}>
          <div style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:20,width:"100%",boxSizing:"border-box",maxHeight:"88vh",overflowY:"auto"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontWeight:"bold",fontSize:16}}>🎟️ Comprovante</div>
              <button onClick={()=>setVerComp(false)} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:"#888"}}>✕</button>
            </div>
            <div style={{textAlign:"center",marginBottom:16}}>
              <div style={{display:"inline-block",width:70,height:70,borderRadius:"50%",overflow:"hidden",border:"2px solid "+VERMELHO}}>
                <img src={LOGO_B64} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="logo" />
              </div>
              <div style={{fontWeight:"bold",fontSize:15,marginTop:6,color:VERMELHO}}>A&P Flamengo</div>
              <div style={{fontSize:11,color:"#888"}}>Comprovante de Devolução</div>
            </div>
            {item.fotoAssinatura
              ? <img src={item.fotoAssinatura} style={{width:"100%",borderRadius:10,maxHeight:180,objectFit:"cover",marginBottom:14}} alt="retirada" />
              : <div style={{background:"#f5f5f5",borderRadius:10,padding:12,marginBottom:14,textAlign:"center",color:"#aaa",fontSize:12}}>📷 Foto de retirada não registrada</div>
            }
            <div style={{background:"#f9f9f9",borderRadius:10,padding:12,marginBottom:14}}>
              {[["Item",item.descricao],["ID",item.id],["Local",item.local],["Data achado",item.data],["Data entrega",new Date().toLocaleDateString("pt-BR")],["Recebido por",item.comprovante?.nome||"—"],["Canal",item.comprovante?.canal==="email"?"✉️ E-mail":"💬 WhatsApp"],["Contato",item.comprovante?.contato||"—"]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #eee"}}>
                  <span style={{color:"#888",fontSize:12}}>{l}</span>
                  <span style={{fontWeight:"600",fontSize:12,textAlign:"right",maxWidth:"60%"}}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{background:"#f5f5f5",borderRadius:8,padding:10,fontSize:10,color:"#999",lineHeight:1.5,marginBottom:16}}>
              🔒 Dados tratados conforme LGPD (Lei nº 13.709/2018).
            </div>
            <button onClick={()=>setVerComp(false)} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>Fechar</button>
          </div>
        </div>
      )}

      {modal && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:200,display:"flex",alignItems:"flex-end"}}>
          <div style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:20,width:"100%",boxSizing:"border-box",maxHeight:"90vh",overflowY:"auto"}}>
            <input ref={fileRef} type="file" accept="image/*" capture="environment" style={{display:"none"}}
              onChange={e=>{const f=e.target.files?.[0];if(f){const r=new FileReader();r.onload=ev=>setFotoAs(ev.target.result);r.readAsDataURL(f);}}} />
            <div style={{fontWeight:"bold",fontSize:16,marginBottom:14}}>📋 Confirmar devolução</div>
            <div style={{display:"flex",gap:8,marginBottom:12}}>
              <div style={{flex:1}}>
                <label style={{fontSize:12,fontWeight:"bold"}}>Nome *</label>
                <input value={nome} onChange={e=>setNome(e.target.value)} placeholder="Nome"
                  style={{width:"100%",padding:10,marginTop:4,borderRadius:8,border:"1px solid "+(nome?"#4CAF50":"#ddd"),fontSize:14,boxSizing:"border-box"}} />
              </div>
              <div style={{flex:1}}>
                <label style={{fontSize:12,fontWeight:"bold"}}>Sobrenome *</label>
                <input value={sobrenome} onChange={e=>setSobrenome(e.target.value)} placeholder="Sobrenome"
                  style={{width:"100%",padding:10,marginTop:4,borderRadius:8,border:"1px solid "+(sobrenome?"#4CAF50":"#ddd"),fontSize:14,boxSizing:"border-box"}} />
              </div>
            </div>
            {fotoAs
              ? <img src={fotoAs} style={{width:"100%",height:120,objectFit:"cover",borderRadius:8,marginBottom:10}} alt="assinatura" />
              : <button onClick={()=>fileRef.current?.click()} style={{width:"100%",padding:12,background:"#f5f5f5",border:"2px dashed "+(fotoAs?"#4CAF50":"#ccc"),borderRadius:8,cursor:"pointer",fontSize:13,marginBottom:10}}>📷 Foto do sócio com o objeto *</button>
            }
            {(!nome||!sobrenome||!fotoAs) && (
              <div style={{background:"#FFF3CD",border:"1px solid #FFC107",borderRadius:8,padding:10,marginBottom:10,fontSize:12,color:"#856404"}}>
                ⚠️ Obrigatório: {[!nome&&"Nome",!sobrenome&&"Sobrenome",!fotoAs&&"Foto"].filter(Boolean).join(" · ")}
              </div>
            )}
            <div style={{display:"flex",gap:8,marginBottom:10}}>
              {["email","whatsapp"].map(c=>(
                <button key={c} onClick={()=>{setCanal(c);setContato(c==="whatsapp"?"+5521":"");}}
                  style={{flex:1,padding:10,borderRadius:8,cursor:"pointer",background:canal===c?VERMELHO:"#f5f5f5",color:canal===c?"#fff":PRETO,border:"1px solid "+(canal===c?VERMELHO:"#ddd"),fontWeight:"bold",fontSize:13}}>
                  {c==="email"?"✉️ E-mail":"💬 WhatsApp"}
                </button>
              ))}
            </div>
            <input value={contato} onChange={e=>setContato(e.target.value)}
              placeholder={canal==="email"?"email@socio.com":"+5521 99999-0000"}
              style={{width:"100%",padding:10,borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box",marginBottom:10}} />
            <div style={{background:"#f0f8f0",border:"1px solid #ccc",borderRadius:8,padding:10,fontSize:11,color:"#555",marginBottom:12,lineHeight:1.6}}>
              {nome?"Olá, "+nome+"! 👋":"Olá! 👋"} Comprovante de retirada: <strong>{item.descricao}</strong>. Data: {new Date().toLocaleDateString("pt-BR")}. Que bom que encontrou! 🏆<br/>
              <em>Dados tratados conforme LGPD (Lei 13.709/2018).</em>
            </div>
            <button onClick={()=>{
              if(!nome){alert("⚠️ Informe o nome!");return;}
              if(!sobrenome){alert("⚠️ Informe o sobrenome!");return;}
              if(!fotoAs){alert("⚠️ A foto é obrigatória!");return;}
              if(!contato){alert("⚠️ Informe o contato!");return;}
              onEntregue(item.id,{canal,contato,nome:(nome+" "+sobrenome).trim()},fotoAs);
              setModal(false); onVoltar();
            }} style={{width:"100%",padding:14,background:(!nome||!sobrenome||!fotoAs)?"#ccc":VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:8}}>✅ Confirmar e enviar</button>
            <button onClick={()=>setModal(false)} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:14,cursor:"pointer"}}>Cancelar</button>
          </div>
        </div>
      )}
    </div>
  );
}

function Consultar({ itens, onVerItem }) {
  const [busca, setBusca] = useState("");
  const [aba, setAba] = useState("todos");
  const [dIni, setDIni] = useState("");
  const [dFim, setDFim] = useState("");

  const filtrados = itens.filter(i => {
    const b = normalizar(busca);
    return !i.excluido &&
      (!busca||normalizar(i.descricao).includes(b)||normalizar(i.categoria).includes(b)||normalizar(i.local).includes(b))
      && (aba==="todos"||i.status===aba) && (!dIni||i.data>=dIni) && (!dFim||i.data<=dFim);
  });

  return (
    <div style={{padding:16}}>
      <input value={busca} onChange={e=>setBusca(e.target.value)} placeholder="🔍 Buscar (aceita sem acento)..."
        style={{width:"100%",padding:12,borderRadius:10,border:"1px solid #ddd",fontSize:14,marginBottom:10,boxSizing:"border-box"}} />
      <div style={{display:"flex",gap:8,marginBottom:10}}>
        <div style={{flex:1}}><div style={{fontSize:11,color:"#888"}}>De</div>
          <input type="date" value={dIni} onChange={e=>setDIni(e.target.value)} style={{width:"100%",padding:8,borderRadius:8,border:"1px solid #ddd",fontSize:12,boxSizing:"border-box"}} /></div>
        <div style={{flex:1}}><div style={{fontSize:11,color:"#888"}}>Até</div>
          <input type="date" value={dFim} onChange={e=>setDFim(e.target.value)} style={{width:"100%",padding:8,borderRadius:8,border:"1px solid #ddd",fontSize:12,boxSizing:"border-box"}} /></div>
      </div>
      <div style={{display:"flex",gap:6,marginBottom:14}}>
        {[{l:"Todos",v:"todos"},{l:"Disponíveis",v:"disponivel"},{l:"Entregues",v:"entregue"}].map(a=>(
          <button key={a.v} onClick={()=>setAba(a.v)} style={{flex:1,padding:"8px 4px",borderRadius:8,cursor:"pointer",fontSize:12,background:aba===a.v?VERMELHO:"#f5f5f5",color:aba===a.v?"#fff":PRETO,border:"1px solid "+(aba===a.v?VERMELHO:"#ddd"),fontWeight:aba===a.v?"bold":"normal"}}>{a.l}</button>
        ))}
      </div>
      {filtrados.length===0
        ? <div style={{textAlign:"center",color:"#aaa",padding:40}}>Nenhum item encontrado</div>
        : filtrados.map(item=>(
          <div key={item.id} onClick={()=>onVerItem(item)} style={{background:"#fff",borderRadius:12,padding:12,marginBottom:10,boxShadow:"0 2px 6px rgba(0,0,0,0.08)",cursor:"pointer",display:"flex",gap:12,alignItems:"center",border:"1px solid #f0f0f0"}}>
            <div style={{width:50,height:50,borderRadius:10,background:"#f5f5f5",display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",flexShrink:0}}>
              {item.foto?<img src={item.foto} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="" />:<span style={{fontSize:24}}>📦</span>}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:"bold",fontSize:14,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.descricao}</div>
              <div style={{fontSize:12,color:"#888"}}>{item.categoria} · {item.local}</div>
              <div style={{fontSize:11,color:"#aaa"}}>{item.data} · {item.id}</div>
            </div>
            <span style={{background:item.status==="disponivel"?"#E8F5E9":"#E3F2FD",color:item.status==="disponivel"?"#2E7D32":"#1565C0",fontSize:11,padding:"3px 8px",borderRadius:20,fontWeight:"bold",whiteSpace:"nowrap"}}>
              {item.status==="disponivel"?"Disponível":"Entregue"}
            </span>
          </div>
        ))
      }
    </div>
  );
}

function Relatorios({ itens, usuarios, onCadastrarUsuario, onEditarUsuario, perfilLogado, onRestaurar }) {
  const [aba, setAba] = useState(null);
  const [modal, setModal] = useState(false);
  const [email, setEmail] = useState("");
  const [compSel, setCompSel] = useState(null);
  const [usuarioSel, setUsuarioSel] = useState(null);
  const [editSenha, setEditSenha] = useState("");
  const [novoUsuarioModal, setNovoUsuarioModal] = useState(false);
  const [novoNome, setNovoNome] = useState("");
  const [novoUser, setNovoUser] = useState("");
  const [novoEmail, setNovoEmail] = useState("");
  const [novoPerfil, setNovoPerfil] = useState("operador");
  const [novaSenha, setNovaSenha] = useState("");
  const [novaSenha2, setNovaSenha2] = useState("");

  const comprovantes = itens.filter(i=>i.status==="entregue" && !i.excluido);
  const excluidos = itens.filter(i=>i.excluido);

  const Voltar = () => (
    <button onClick={()=>{setAba(null);setUsuarioSel(null);}} style={{display:"flex",alignItems:"center",gap:6,background:"none",border:"none",color:VERMELHO,fontWeight:"bold",fontSize:14,cursor:"pointer",marginBottom:16,padding:0}}>← Voltar</button>
  );

  if (!aba) return (
    <div style={{padding:16}}>
      <div style={{color:"#888",fontSize:13,marginBottom:16}}>Selecione uma opção:</div>
      {[
        {icon:"🔍",label:"Achados",desc:"Itens disponíveis no inventário",val:"achados",count:itens.filter(i=>i.status==="disponivel"&&!i.excluido).length,cor:"#4CAF50"},
        {icon:"✅",label:"Retirados",desc:"Itens já entregues aos donos",val:"retirados",count:itens.filter(i=>i.status==="entregue"&&!i.excluido).length,cor:"#2196F3"},
        {icon:"🎟️",label:"Comprovantes",desc:"Histórico de comprovantes enviados",val:"comprovantes",count:comprovantes.length,cor:VERMELHO},
        {icon:"👥",label:"Usuários",desc:"Gestão de usuários do sistema",val:"usuarios",count:usuarios.length,cor:"#9C27B0"},
        ...(perfilLogado==="admin"?[{icon:"🗑️",label:"Excluídos",desc:"Itens removidos do inventário",val:"excluidos",count:excluidos.length,cor:"#757575"}]:[]),
      ].map(c=>(
        <button key={c.val} onClick={()=>setAba(c.val)} style={{width:"100%",marginBottom:12,padding:18,background:"#fff",border:"1px solid #f0f0f0",borderRadius:14,cursor:"pointer",display:"flex",alignItems:"center",gap:16,boxShadow:"0 2px 8px rgba(0,0,0,0.07)",textAlign:"left"}}>
          <div style={{width:52,height:52,borderRadius:14,background:c.cor+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:26,flexShrink:0}}>{c.icon}</div>
          <div style={{flex:1}}>
            <div style={{fontWeight:"bold",fontSize:16,color:PRETO}}>{c.label}</div>
            <div style={{fontSize:12,color:"#999",marginTop:2}}>{c.desc}</div>
          </div>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:24,fontWeight:"bold",color:c.cor}}>{c.count}</div>
            <div style={{fontSize:11,color:"#aaa"}}>{c.val==="usuarios"?"cadastros":"itens"}</div>
          </div>
        </button>
      ))}
    </div>
  );

  if (aba==="achados"||aba==="retirados") {
    const itensFilt = aba==="achados"?itens.filter(i=>i.status==="disponivel"&&!i.excluido):itens.filter(i=>i.status==="entregue"&&!i.excluido);
    return (
      <div style={{padding:16}}>
        <Voltar />
        <div style={{fontWeight:"bold",fontSize:16,marginBottom:4}}>{aba==="achados"?"🔍 Achados":"✅ Retirados"}</div>
        <div style={{fontSize:12,color:"#888",marginBottom:14}}>{itensFilt.length} item(s)</div>
        <button onClick={()=>exportarExcel(itensFilt)} style={{width:"100%",padding:14,background:"#217346",color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:10}}>📊 Baixar planilha Excel</button>
        <button onClick={()=>setModal(true)} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:16}}>✉️ Enviar por e-mail</button>
        {itensFilt.map(item=>(
          <div key={item.id} style={{background:"#fff",borderRadius:12,padding:12,marginBottom:10,boxShadow:"0 2px 6px rgba(0,0,0,0.08)",border:"1px solid #f0f0f0",display:"flex",gap:12,alignItems:"center"}}>
            <div style={{width:44,height:44,borderRadius:10,background:"#f5f5f5",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,overflow:"hidden"}}>
              {item.foto?<img src={item.foto} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="" />:<span style={{fontSize:22}}>📦</span>}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:"bold",fontSize:14,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{item.descricao}</div>
              <div style={{fontSize:12,color:"#888"}}>{item.categoria} · {item.local}</div>
              <div style={{fontSize:11,color:"#aaa"}}>{item.data} · {item.id}</div>
            </div>
          </div>
        ))}
        {modal && (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:200,display:"flex",alignItems:"flex-end"}}>
            <div style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:20,width:"100%",boxSizing:"border-box"}}>
              <div style={{fontWeight:"bold",fontSize:16,marginBottom:14}}>✉️ Enviar por e-mail</div>
              <input value={email} onChange={e=>setEmail(e.target.value)} placeholder="email@flamengo.com.br" type="email"
                style={{width:"100%",padding:12,borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box",marginBottom:16}} />
              <button onClick={()=>{if(!email){alert("Informe o e-mail!");return;}exportarExcel(itensFilt);alert("Download iniciado!\nEm produção será enviado para: "+email);setModal(false);}}
                style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:8}}>✉️ Enviar e baixar</button>
              <button onClick={()=>setModal(false)} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:14,cursor:"pointer"}}>Cancelar</button>
            </div>
          </div>
        )}
      </div>
    );
  }

  if (aba==="excluidos") {
    return (
      <div style={{padding:16}}>
        <Voltar />
        <div style={{fontWeight:"bold",fontSize:16,marginBottom:4}}>🗑️ Itens Excluídos</div>
        <div style={{fontSize:12,color:"#888",marginBottom:14}}>{excluidos.length} item(s) excluído(s)</div>
        {excluidos.length===0
          ? <div style={{textAlign:"center",color:"#aaa",padding:40}}>Nenhum item excluído</div>
          : excluidos.map(item=>(
            <div key={item.id} style={{background:"#fff",borderRadius:12,padding:14,marginBottom:10,boxShadow:"0 2px 6px rgba(0,0,0,0.08)",border:"1px solid #f0f0f0",opacity:0.8}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
                <div style={{fontWeight:"bold",fontSize:14,color:"#555"}}>{item.descricao}</div>
                <span style={{background:"#f5f5f5",color:"#888",fontSize:11,padding:"3px 8px",borderRadius:20,fontWeight:"bold"}}>Excluído</span>
              </div>
              <div style={{fontSize:12,color:"#aaa",marginBottom:4}}>{item.categoria} · {item.local}</div>
              <div style={{fontSize:11,color:"#bbb"}}>ID: {item.id} · Cadastrado: {item.data}</div>
              {item.excluidoEm && <div style={{fontSize:11,color:"#bbb"}}>Excluído em: {item.excluidoEm} por {item.excluidoPor}</div>}
              <button onClick={()=>{ onRestaurar(item.id); }} style={{width:"100%",marginTop:10,padding:10,background:"#E8F5E9",color:"#2E7D32",border:"1px solid #A5D6A7",borderRadius:8,fontSize:13,fontWeight:"bold",cursor:"pointer"}}>
                ↩️ Restaurar item
              </button>
            </div>
          ))
        }
      </div>
    );
  }

  if (aba==="usuarios") {
    const perfilCor = {admin:VERMELHO,operador:"#4CAF50"};
    const perfilLabel = {admin:"Administrador",operador:"Operador"};
    return (
      <div style={{padding:16}}>
        <Voltar />
        <div style={{fontWeight:"bold",fontSize:16,marginBottom:4}}>👥 Usuários</div>
        <div style={{fontSize:12,color:"#888",marginBottom:14}}>{usuarios.length} cadastro(s)</div>
        {perfilLogado==="admin" && (
          <button onClick={()=>setNovoUsuarioModal(true)} style={{width:"100%",marginBottom:16,padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>➕ Novo usuário</button>
        )}
        {usuarios.map(u=>(
          <div key={u.usuario} style={{background:"#fff",borderRadius:12,padding:14,marginBottom:10,boxShadow:"0 2px 6px rgba(0,0,0,0.08)",border:"1px solid #f0f0f0"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}>
              <div style={{fontWeight:"bold",fontSize:15}}>{u.nome}</div>
              <span style={{background:u.ativo?"#E8F5E9":"#FFEBEE",color:u.ativo?"#2E7D32":"#C62828",fontSize:11,padding:"3px 10px",borderRadius:20,fontWeight:"bold"}}>{u.ativo?"Ativo":"Inativo"}</span>
            </div>
            <div style={{fontSize:12,color:"#888",marginBottom:6}}>@{u.usuario} · {u.email||"—"} · Desde {u.criado}</div>
            <span style={{background:perfilCor[u.perfil]+"20",color:perfilCor[u.perfil],fontSize:11,padding:"2px 10px",borderRadius:20,fontWeight:"bold"}}>{perfilLabel[u.perfil]||u.perfil}</span>
            {perfilLogado==="admin" && (
              <div style={{display:"flex",gap:8,marginTop:10}}>
                <button onClick={()=>{setUsuarioSel(u);setEditSenha("");}} style={{flex:1,padding:8,background:"#f5f5f5",border:"1px solid #ddd",borderRadius:8,fontSize:12,cursor:"pointer"}}>✏️ Editar</button>
                <button onClick={()=>{if(u.usuario==="mirna"){alert("Não é possível desativar o admin principal!");return;}onEditarUsuario({...u,ativo:!u.ativo});}}
                  style={{flex:1,padding:8,background:u.ativo?"#FFF3F3":"#F3FFF3",border:"1px solid "+(u.ativo?"#ffcccc":"#ccffcc"),borderRadius:8,fontSize:12,cursor:"pointer",color:u.ativo?"#C62828":"#2E7D32",fontWeight:"bold"}}>
                  {u.ativo?"🚫 Desativar":"✅ Ativar"}
                </button>
              </div>
            )}
          </div>
        ))}
        {novoUsuarioModal && (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:300,display:"flex",alignItems:"flex-end"}}>
            <div style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:20,width:"100%",boxSizing:"border-box",maxHeight:"90vh",overflowY:"auto"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
                <div style={{fontWeight:"bold",fontSize:16}}>➕ Novo usuário</div>
                <button onClick={()=>setNovoUsuarioModal(false)} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:"#888"}}>✕</button>
              </div>
              {[["Nome completo",novoNome,setNovoNome,"Ex: João Silva","text"],["Usuário",novoUser,setNovoUser,"Ex: joao.silva","text"],["E-mail",novoEmail,setNovoEmail,"email@exemplo.com","email"],["Senha",novaSenha,setNovaSenha,"Crie uma senha","password"],["Confirmar senha",novaSenha2,setNovaSenha2,"Repita a senha","password"]].map(([l,v,s,ph,t])=>(
                <div key={l}>
                  <label style={{fontSize:13,fontWeight:"bold"}}>{l}</label>
                  <input value={v} onChange={e=>s(e.target.value)} placeholder={ph} type={t} autoCapitalize={t==="text"&&l==="Usuário"?"none":"sentences"}
                    style={{width:"100%",padding:10,marginTop:4,marginBottom:12,borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box"}} />
                </div>
              ))}
              <label style={{fontSize:13,fontWeight:"bold"}}>Perfil</label>
              <select value={novoPerfil} onChange={e=>setNovoPerfil(e.target.value)}
                style={{width:"100%",padding:10,marginTop:4,marginBottom:16,borderRadius:8,border:"1px solid #ddd",fontSize:14,background:"#fff"}}>
                <option value="operador">Operador — acesso padrão</option>
                <option value="admin">Administrador — acesso total</option>
              </select>
              <button onClick={()=>{
                if(!novoNome||!novoUser||!novoEmail||!novaSenha){alert("Preencha todos os campos!");return;}
                if(novaSenha!==novaSenha2){alert("As senhas não coincidem!");return;}
                if(usuarios.find(u=>u.usuario.toLowerCase()===novoUser.toLowerCase())){alert("Usuário já existe!");return;}
                onCadastrarUsuario({usuario:novoUser,senha:novaSenha,nome:novoNome,email:novoEmail,perfil:novoPerfil,ativo:true,criado:new Date().toISOString().split("T")[0]});
                setNovoUsuarioModal(false);setNovoNome("");setNovoUser("");setNovoEmail("");setNovaSenha("");setNovaSenha2("");
                alert("✅ Usuário cadastrado!");
              }} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:8}}>✅ Cadastrar</button>
              <button onClick={()=>setNovoUsuarioModal(false)} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:14,cursor:"pointer"}}>Cancelar</button>
            </div>
          </div>
        )}
        {usuarioSel && (
          <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:300,display:"flex",alignItems:"flex-end"}}>
            <div style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:20,width:"100%",boxSizing:"border-box"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:16}}>
                <div style={{fontWeight:"bold",fontSize:16}}>✏️ Editar: {usuarioSel.nome}</div>
                <button onClick={()=>setUsuarioSel(null)} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:"#888"}}>✕</button>
              </div>
              <label style={{fontSize:13,fontWeight:"bold"}}>Perfil</label>
              <select value={usuarioSel.perfil} onChange={e=>setUsuarioSel({...usuarioSel,perfil:e.target.value})}
                style={{width:"100%",padding:10,marginTop:4,marginBottom:12,borderRadius:8,border:"1px solid #ddd",fontSize:14,background:"#fff"}}>
                <option value="operador">Operador — acesso padrão</option>
                <option value="admin">Administrador — acesso total</option>
              </select>
              <label style={{fontSize:13,fontWeight:"bold"}}>Nova senha (deixe vazio para manter)</label>
              <input value={editSenha} onChange={e=>setEditSenha(e.target.value)} type="password" placeholder="Nova senha"
                style={{width:"100%",padding:10,marginTop:4,borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box",marginBottom:16}} />
              <button onClick={()=>{onEditarUsuario({...usuarioSel,...(editSenha?{senha:editSenha}:{})});setUsuarioSel(null);alert("✅ Usuário atualizado!");}}
                style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:8}}>Salvar</button>
              <button onClick={()=>setUsuarioSel(null)} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:14,cursor:"pointer"}}>Cancelar</button>
            </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={{padding:16}}>
      <Voltar />
      <div style={{fontWeight:"bold",fontSize:16,marginBottom:14}}>🎟️ Comprovantes</div>
      {comprovantes.length===0
        ? <div style={{textAlign:"center",color:"#aaa",padding:40}}>Nenhum comprovante enviado ainda</div>
        : comprovantes.map(item=>(
          <div key={item.id} style={{background:"#fff",borderRadius:12,padding:12,marginBottom:10,boxShadow:"0 2px 6px rgba(0,0,0,0.08)",border:"1px solid #f0f0f0"}}>
            <div style={{fontWeight:"bold",fontSize:14}}>{item.descricao}</div>
            <div style={{fontSize:12,color:"#888",marginBottom:8}}>ID: {item.id} · {item.data}{item.comprovante&&<span> · {item.comprovante.canal==="email"?"✉️":"💬"} {item.comprovante.contato}</span>}</div>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setCompSel(item)} style={{flex:1,padding:8,background:"#f5f5f5",border:"1px solid #ddd",borderRadius:8,fontSize:12,cursor:"pointer"}}>👁 Visualizar</button>
              <button onClick={()=>alert("Reenviado para "+(item.comprovante?.contato||"—")+"!")} style={{flex:1,padding:8,background:"#f5f5f5",border:"1px solid #ddd",borderRadius:8,fontSize:12,cursor:"pointer"}}>🔄 Reenviar</button>
            </div>
          </div>
        ))
      }
      {compSel && (
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:300,display:"flex",alignItems:"flex-end"}}>
          <div style={{background:"#fff",borderRadius:"20px 20px 0 0",padding:20,width:"100%",boxSizing:"border-box",maxHeight:"88vh",overflowY:"auto"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontWeight:"bold",fontSize:16}}>🎟️ Comprovante</div>
              <button onClick={()=>setCompSel(null)} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:"#888"}}>✕</button>
            </div>
            <div style={{textAlign:"center",marginBottom:16}}>
              <div style={{display:"inline-block",width:70,height:70,borderRadius:"50%",overflow:"hidden",border:"2px solid "+VERMELHO}}>
                <img src={LOGO_B64} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="logo" />
              </div>
              <div style={{fontWeight:"bold",fontSize:15,marginTop:6,color:VERMELHO}}>A&P Flamengo</div>
              <div style={{fontSize:11,color:"#888"}}>Comprovante de Devolução</div>
            </div>
            {compSel.fotoAssinatura
              ? <img src={compSel.fotoAssinatura} style={{width:"100%",borderRadius:10,maxHeight:180,objectFit:"cover",marginBottom:14}} alt="retirada" />
              : <div style={{background:"#f5f5f5",borderRadius:10,padding:12,marginBottom:14,textAlign:"center",color:"#aaa",fontSize:12}}>📷 Foto não registrada</div>
            }
            <div style={{background:"#f9f9f9",borderRadius:10,padding:12,marginBottom:14}}>
              {[["Item",compSel.descricao],["ID",compSel.id],["Local",compSel.local],["Data achado",compSel.data],["Data entrega",new Date().toLocaleDateString("pt-BR")],["Recebido por",compSel.comprovante?.nome||"—"],["Canal",compSel.comprovante?.canal==="email"?"✉️ E-mail":"💬 WhatsApp"],["Contato",compSel.comprovante?.contato||"—"]].map(([l,v])=>(
                <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #eee"}}>
                  <span style={{color:"#888",fontSize:12}}>{l}</span>
                  <span style={{fontWeight:"600",fontSize:12,textAlign:"right",maxWidth:"60%"}}>{v}</span>
                </div>
              ))}
            </div>
            <div style={{background:"#f5f5f5",borderRadius:8,padding:10,fontSize:10,color:"#999",lineHeight:1.5,marginBottom:16}}>
              🔒 Dados tratados conforme LGPD (Lei nº 13.709/2018).
            </div>
            <button onClick={()=>setCompSel(null)} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer"}}>Fechar</button>
          </div>
        </div>
      )}
    </div>
  );
}

function Perfil({ usuario, onVoltar, onSalvar }) {
  const [novaSenha, setNovaSenha] = useState("");
  const [novaSenha2, setNovaSenha2] = useState("");
  const [email, setEmail] = useState(usuario.email||"");
  const [ver, setVer] = useState(false);
  const [sucesso, setSucesso] = useState(false);

  const salvar = () => {
    if(novaSenha && novaSenha!==novaSenha2){ alert("As senhas não coincidem!"); return; }
    onSalvar({...usuario,email,...(novaSenha?{senha:novaSenha}:{})});
    setSucesso(true); setTimeout(()=>setSucesso(false),3000);
    setNovaSenha(""); setNovaSenha2("");
  };

  return (
    <div style={{padding:16}}>
      <div style={{textAlign:"center",marginBottom:24}}>
        <div style={{display:"inline-block",width:90,height:90,borderRadius:"50%",overflow:"hidden",border:"3px solid "+VERMELHO,marginBottom:12}}>
          <img src={LOGO_B64} style={{width:"100%",height:"100%",objectFit:"cover"}} alt="perfil" />
        </div>
        <div style={{fontWeight:"bold",fontSize:20,color:PRETO}}>{usuario.nome}</div>
        <div style={{fontSize:12,color:"#888",marginTop:2}}>@{usuario.usuario}</div>
        <span style={{background:usuario.perfil==="admin"?VERMELHO+"20":"#E8F5E9",color:usuario.perfil==="admin"?VERMELHO:"#2E7D32",fontSize:12,padding:"3px 12px",borderRadius:20,fontWeight:"bold",marginTop:6,display:"inline-block"}}>
          {usuario.perfil==="admin"?"Administrador":"Operador"}
        </span>
      </div>
      <div style={{background:"#fff",borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 2px 6px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:"bold",fontSize:14,marginBottom:12}}>Informações da conta</div>
        {[["👤 Nome",usuario.nome],["🔑 Usuário","@"+usuario.usuario],["📅 Desde",usuario.criado]].map(([l,v])=>(
          <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:"1px solid #f5f5f5"}}>
            <span style={{color:"#888",fontSize:13}}>{l}</span>
            <span style={{fontWeight:"600",fontSize:13}}>{v}</span>
          </div>
        ))}
      </div>
      <div style={{background:"#fff",borderRadius:14,padding:16,marginBottom:16,boxShadow:"0 2px 6px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:"bold",fontSize:14,marginBottom:12}}>📧 E-mail de recuperação</div>
        <input value={email} onChange={e=>setEmail(e.target.value)} type="email" placeholder="seu@email.com"
          style={{width:"100%",padding:10,borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box"}} />
        <div style={{fontSize:11,color:"#aaa",marginTop:6}}>Usado para recuperação de senha</div>
      </div>
      <div style={{background:"#fff",borderRadius:14,padding:16,marginBottom:20,boxShadow:"0 2px 6px rgba(0,0,0,0.06)"}}>
        <div style={{fontWeight:"bold",fontSize:14,marginBottom:12}}>🔒 Alterar senha</div>
        <label style={{fontSize:12,color:"#666"}}>Nova senha</label>
        <div style={{position:"relative",marginTop:4,marginBottom:10}}>
          <input value={novaSenha} onChange={e=>setNovaSenha(e.target.value)} type={ver?"text":"password"} placeholder="Digite nova senha"
            style={{width:"100%",padding:"10px 40px 10px 10px",borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box"}} />
          <button onClick={()=>setVer(!ver)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#aaa"}}>{ver?"🙈":"👁"}</button>
        </div>
        <label style={{fontSize:12,color:"#666"}}>Confirmar nova senha</label>
        <input value={novaSenha2} onChange={e=>setNovaSenha2(e.target.value)} type="password" placeholder="Repita a senha"
          style={{width:"100%",padding:10,marginTop:4,borderRadius:8,border:"1px solid #ddd",fontSize:14,boxSizing:"border-box"}} />
        <div style={{fontSize:11,color:"#aaa",marginTop:6}}>Deixe em branco para manter a senha atual</div>
      </div>
      {sucesso && <div style={{background:"#E8F5E9",border:"1px solid #4CAF50",borderRadius:8,padding:10,marginBottom:12,fontSize:13,color:"#2E7D32",textAlign:"center"}}>✅ Perfil atualizado!</div>}
      <button onClick={salvar} style={{width:"100%",padding:14,background:VERMELHO,color:"#fff",border:"none",borderRadius:12,fontSize:15,fontWeight:"bold",cursor:"pointer",marginBottom:10}}>💾 Salvar alterações</button>
      <button onClick={onVoltar} style={{width:"100%",padding:14,background:"#f5f5f5",color:PRETO,border:"1px solid #ddd",borderRadius:12,fontSize:14,cursor:"pointer"}}>Voltar</button>
    </div>
  );
}

export default function App() {
  const [logado, setLogado] = useState(false);
  const [usuarioLogado, setUsuarioLogado] = useState(null);
  const [tela, setTela] = useState("home");
  const [itens, setItens] = useState(MOCK);
  const [itemSel, setItemSel] = useState(null);
  const [usuarios, setUsuarios] = useState(USUARIOS_INICIAIS);

  if (!logado) return (
    <Login
      onLogin={(u)=>{ setUsuarioLogado(u); setLogado(true); }}
      usuarios={usuarios}
      onCadastrarUsuario={u=>setUsuarios(p=>[...p,u])}
    />
  );

  const voltar = () => { setItemSel(null); setTela(tela==="ficha"?"consultar":"home"); };
  const reset = () => { setItens(MOCK); setTela("home"); setItemSel(null); };
  const titulos = { cadastrar:"Cadastrar Achado", consultar:"Consultar Perdidos", relatorios:"Relatórios", ficha:itemSel?.descricao, perfil:"Meu Perfil" };

  return (
    <div style={{width:"100%",minHeight:"100vh",background:"#f5f5f5",fontFamily:"'Segoe UI',sans-serif",boxSizing:"border-box",overflowX:"hidden"}}>
      <Header tela={tela} onVoltar={voltar} onReset={reset} nomeLogado={usuarioLogado?.nome||""} onPerfil={()=>setTela("perfil")} />
      {tela!=="home" && titulos[tela] && (
        <div style={{background:VERMELHO,color:"#fff",padding:"10px 16px",fontSize:15,fontWeight:"bold"}}>{titulos[tela]}</div>
      )}
      {tela==="home" && <Home itens={itens} onNav={setTela} />}
      {tela==="cadastrar" && <Cadastrar onSalvar={n=>setItens(p=>[n,...p])} onVoltar={()=>setTela("home")} />}
      {tela==="consultar" && <Consultar itens={itens} onVerItem={i=>{setItemSel(i);setTela("ficha");}} />}
      {tela==="relatorios" && (
        <Relatorios itens={itens} usuarios={usuarios} perfilLogado={usuarioLogado?.perfil}
          onCadastrarUsuario={u=>setUsuarios(p=>[...p,u])}
          onEditarUsuario={u=>setUsuarios(p=>p.map(x=>x.usuario===u.usuario?u:x))}
          onRestaurar={(id)=>setItens(p=>p.map(i=>i.id===id?{...i,excluido:false,excluidoEm:null,excluidoPor:null}:i))}
        />
      )}
      {tela==="ficha" && itemSel && <Ficha item={itens.find(i=>i.id===itemSel.id)||itemSel} onVoltar={voltar}
        perfilLogado={usuarioLogado?.perfil}
        onExcluir={(id)=>{ 
          if(usuarioLogado?.perfil!=="admin"){ alert("Apenas administradores podem excluir itens!"); return; }
          setItens(p=>p.map(i=>i.id===id?{...i,excluido:true,excluidoEm:new Date().toLocaleDateString("pt-BR"),excluidoPor:usuarioLogado?.nome}:i)); 
          voltar(); 
        }}
        onEntregue={(id,comp,fotoAs)=>setItens(p=>p.map(i=>i.id===id?{...i,status:"entregue",comprovante:comp,fotoAssinatura:fotoAs}:i))} />}
      {tela==="perfil" && usuarioLogado && <Perfil usuario={usuarioLogado}
        onVoltar={()=>setTela("home")}
        onSalvar={(dados)=>{setUsuarios(p=>p.map(u=>u.usuario===dados.usuario?dados:u));setUsuarioLogado(dados);}} />}
    </div>
  );
}
