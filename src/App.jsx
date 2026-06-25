import React, { useState } from 'react';
// Caso seu olho seja um ícone do react-icons, pode importar aqui. 
// Fiz com emoji de olho (👁️) para garantir que não quebre por falta de dependência, igual ao print.

export default function App() {
  const [usuario, setUsuario] = useState('');
  const [senha, setSenha] = useState('');
  const [mostrarSenha, setMostrarSenha] = useState(false);

  const handleLogin = (e) => {
    e.preventDefault();
    
    // CASO SUA VALIDAÇÃO SEJA LOCAL, AJUSTE AQUI:
    // Exemplo: if (usuario === 'admin' && senha === '123') { ... }
    
    console.log('Tentando logar com:', { usuario, senha });
    
    // Se o erro persistir, verifique se as credenciais digitadas batem 
    // com o que o seu back-end ou validação local estão esperando.
    alert("Usuário ou senha incorretos (ou conta suspensa)!"); 
  };

  return (
    <div className="min-h-screen bg-[#111111] flex flex-col items-center justify-between p-6 font-sans selection:bg-red-500 selection:text-white">
      
      {/* Topo / Logo e Título Externo */}
      <div className="flex flex-col items-center mt-8 w-full max-w-md">
        {/* Logo Redondo - Altere o src para o caminho correto se necessário */}
        <div className="w-32 h-32 bg-[#E11A2B] rounded-full flex items-center justify-center shadow-lg border-4 border-white mb-4 overflow-hidden">
          {/* Se tiver o arquivo local ex: /logo-ap.png, mude aqui */}
          <img 
            src="/logo-ap.png" 
            alt="A&P Flamengo" 
            className="w-full h-full object-cover"
            onError={(e) => {
              // Fallback caso a imagem suma de novo, para não quebrar o layout
              e.target.style.display = 'none';
              e.target.parentNode.innerHTML = '<span className="text-white font-bold text-xl">A&P<br/><span className="text-xs">FLAMENGO</span></span>';
            }}
          />
        </div>
        <p className="text-gray-400 text-sm tracking-wide text-center">
          Gestão de Achados & Perdidos
        </p>
      </div>

      {/* Card de Login */}
      <div className="bg-white w-full max-w-md rounded-3xl p-8 shadow-2xl my-auto border border-gray-100">
        <h2 className="text-xl font-bold text-gray-800 mb-6 text-left">
          Entrar
        </h2>

        <form onSubmit={handleLogin} className="space-y-5">
          {/* Campo Usuário */}
          <div className="flex flex-col space-y-1">
            <label className="text-sm font-semibold text-gray-600 text-left">
              Usuário
            </label>
            <input
              type="text"
              value={usuario}
              onChange={(e) => setUsuario(e.target.value)}
              placeholder="Digite seu usuário"
              className="w-full px-4 py-3 rounded-xl border border-gray-300 focus:outline-none focus:border-red-500 text-gray-800 bg-white placeholder-gray-400 transition-colors"
              required
            />
          </div>

          {/* Campo Senha */}
          <div className="flex flex-col space-y-1 relative">
            <label className="text-sm font-semibold text-gray-600 text-left">
              Senha
            </label>
            <div className="relative w-full">
              <input
                type={mostrarSenha ? "text" : "password"}
                value={senha}
                onChange={(e) => setSenha(e.target.value)}
                placeholder="Digite sua senha"
                className="w-full px-4 py-3 rounded-xl border border-gray-300 focus:outline-none focus:border-red-500 text-gray-800 bg-white placeholder-gray-400 transition-colors pr-12"
                required
              />
              {/* Botão do Olho para Mostrar/Esconder Senha */}
              <button
                type="button"
                onClick={() => setMostrarSenha(!mostrarSenha)}
                className="absolute right-4 top-1/2 -translate-y-1/2 text-xl focus:outline-none hover:scale-110 transition-transform"
                title={mostrarSenha ? "Ocultar senha" : "Mostrar senha"}
              >
                👁️
              </button>
            </div>
          </div>

          {/* Botão Entrar */}
          <button
            type="submit"
            className="w-full bg-[#C21E34] hover:bg-[#A6182A] text-white font-bold py-3.5 rounded-xl transition-colors shadow-md mt-2 focus:outline-none focus:ring-2 focus:ring-red-400"
          >
            Entrar
          </button>
        </form>
      </div>

      {/* Rodapé */}
      <div className="text-center text-xs text-gray-600 mt-auto tracking-wide">
        © 2026 Clube de Regatas do Flamengo
      </div>

    </div>
  );
}
