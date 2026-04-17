import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import styles from './MenuChamado.module.scss';

export interface IMenuChamadosProps {
  departamento: 'TI' | 'Marketing' | 'Frotas' | 'Facilities'; 
  emailUsuario: string;
}

export const MenuChamados: React.FC<IMenuChamadosProps> = (props) => {
  // === ESTADOS DO COMPONENTE ===
  const [isNotificacaoOpen, setIsNotificacaoOpen] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [unreadTicketsCount, setUnreadTicketsCount] = useState(0);
  
  const [meusChamados, setMeusChamados] = useState<any[]>([]);
  const [loadingChamados, setLoadingChamados] = useState(false);
  
  const [expandedTicketIndex, setExpandedTicketIndex] = useState<number | null>(null);
  const [comentariosDoChamado, setComentariosDoChamado] = useState<any[]>([]);
  const [loadingHistorico, setLoadingHistorico] = useState(false);
  
  const [novoComentarioChamado, setNovoComentarioChamado] = useState("");
  const [enviandoComentarioChamado, setEnviandoComentarioChamado] = useState(false);

  const [arquivoAnexo, setArquivoAnexo] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const rawEmail = props.emailUsuario || "";
  const userEmail = rawEmail.toLowerCase().trim();

// === EFEITO: RODA AO CARREGAR A PÁGINA (BACKGROUND) ===
  useEffect(() => {
    buscarChamadosEmBackground();

    // Cria um ouvinte global com TRAVA DE SEGURANÇA
    const handleOpenTickets = (e: any) => {
      // Se o botão enviou um departamento, e NÃO FOR o meu departamento, eu ignoro e não abro!
      if (e.detail && e.detail !== props.departamento) {
        return;
      }
      abrirModalMeusChamados();
    };

    window.addEventListener('abrirMeusChamadosGrunner', handleOpenTickets);

    return () => {
      window.removeEventListener('abrirMeusChamadosGrunner', handleOpenTickets);
    };
  }, [userEmail, props.departamento]);

  // === FUNÇÕES ===
  const buscarChamadosEmBackground = async () => {
    // Adicionado quebrador de cache
    const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/meus-chamados?email=${userEmail}&t=${Date.now()}`;
    try {
      const response = await fetch(apiUrl, { cache: 'no-store' });
      const data = await response.json();
      if (data.sucesso && Array.isArray(data.chamados)) {
        setMeusChamados(data.chamados);
        recalcularNotificacoes(data.chamados);
      }
    } catch (error) {
      console.error("Erro ao buscar chamados no background", error);
    }
  };

  const recalcularNotificacoes = (chamadosList: any[]) => {
    let unreadCount = 0;
    chamadosList.forEach((ticket: any) => {
      const lastSeen = localStorage.getItem(`grunner_visto_${ticket.id}`);
      const isEscondido = localStorage.getItem(`grunner_escondido_${ticket.id}`) === "true";
      const isEncerrado = ticket.status.toLowerCase().includes('encerrado') || ticket.status.toLowerCase().includes('conclu');
      
      if (isEscondido && isEncerrado) return; 
      
      // TIMESTAMP EM VEZ DA DATA FORMATADA
      const dataClickUp = parseInt(ticket.timestampAtualizacao || '0');
      const dataLida = parseInt(lastSeen || '0');
      
      if (dataClickUp > dataLida) {
        unreadCount++;
      }
    });
    setUnreadTicketsCount(unreadCount);
  };

  const abrirModalMeusChamados = async () => {
    setIsModalOpen(true);
    setIsNotificacaoOpen(false);
    setLoadingChamados(true);
    setExpandedTicketIndex(null);
    setComentariosDoChamado([]);
    
    // Adicionado quebrador de cache
    const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/meus-chamados?email=${userEmail}&t=${Date.now()}`;
    try {
      const response = await fetch(apiUrl, { cache: 'no-store' });
      const data = await response.json();
      const lista = data.sucesso && Array.isArray(data.chamados) ? data.chamados : [];
      setMeusChamados(lista);
      recalcularNotificacoes(lista);
    } catch (error) {
      console.error("Erro ao abrir modal", error);
    } finally {
      setLoadingChamados(false);
    }
  };

  const toggleDetalhesChamado = async (index: number, idChamado: string) => {
    const ticket = meusChamados[index];
    
    if (expandedTicketIndex === index) {
      setExpandedTicketIndex(null);
      setComentariosDoChamado([]);
      return;
    }

    if (ticket.timestampAtualizacao) {
      localStorage.setItem(`grunner_visto_${idChamado}`, ticket.timestampAtualizacao);
    }

    setExpandedTicketIndex(index);
    setLoadingHistorico(true);
    setComentariosDoChamado([]);
    recalcularNotificacoes(meusChamados);

    try {
      // Adicionado quebrador de cache
      const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/comentarios?idChamado=${idChamado}&t=${Date.now()}`;
      const response = await fetch(apiUrl, { cache: 'no-store' });
      const data = await response.json();
      if (data.sucesso) {
        setComentariosDoChamado(data.comentarios);
      }
    } catch (error) {
      console.error("Erro ao carregar chat:", error);
    } finally {
      setLoadingHistorico(false);
    }
  };

  const dispensarChamado = (idChamado: string) => {
    if (window.confirm("Deseja ocultar este chamado da sua lista?")) {
      localStorage.setItem(`grunner_escondido_${idChamado}`, "true");
      setExpandedTicketIndex(null);
      recalcularNotificacoes(meusChamados);
      // Força a atualização local removendo da lista visual
      setMeusChamados(meusChamados.filter(t => t.id !== idChamado));
    }
  };

  const enviarComentarioChamado = async (idChamado: string) => {
    // Só bloqueia se NÃO houver texto E NÃO houver anexo
    if (!novoComentarioChamado.trim() && !arquivoAnexo) return;

    setEnviandoComentarioChamado(true);
    
    // Função mágica para converter o ficheiro em texto (Base64)
    const toBase64 = (file: File) => new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = error => reject(error);
    });

    let anexoBase64 = null;
    let nomeArquivo = null;

    if (arquivoAnexo) {
      anexoBase64 = await toBase64(arquivoAnexo);
      nomeArquivo = arquivoAnexo.name;
    }

    const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/comentar`;

    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          idChamado: idChamado,
          comentario: novoComentarioChamado,
          email: userEmail,
          anexo: anexoBase64,      
          nomeArquivo: nomeArquivo  
        })
      });

      const result = await response.json();
      if (result.sucesso) {
        setNovoComentarioChamado("");
        setArquivoAnexo(null); // Limpa o clipe de papel após o envio
        
        // Recarrega o chat com quebrador de cache
        const chatResp = await fetch(`https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/comentarios?idChamado=${idChamado}&t=${Date.now()}`, { cache: 'no-store' });
        const chatData = await chatResp.json();
        if (chatData.sucesso) setComentariosDoChamado(chatData.comentarios);
      } else {
        alert("Ocorreu um erro ao enviar: " + result.mensagem);
      }
    } catch (error) {
      alert("Erro de comunicação com o servidor.");
    } finally {
      setEnviandoComentarioChamado(false);
    }
  };

  // === RENDERIZAÇÃO ===
  return (
    <>
      {/* 1. O ENVELOPE (SININHO) */}
      <div className={styles.notificationContainer}>
        <button 
          className={styles.notificationBtn} 
          onClick={() => setIsNotificacaoOpen(!isNotificacaoOpen)}
          title={`Mensagens não lidas de ${props.departamento}`}
        >
          <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.8} stroke="currentColor" style={{ width: '22px', height: '22px' }}>
            <path strokeLinecap="round" strokeLinejoin="round" d="M21.75 6.75v10.5a2.25 2.25 0 0 1-2.25 2.25h-15a2.25 2.25 0 0 1-2.25-2.25V6.75m19.5 0A2.25 2.25 0 0 0 19.5 4.5h-15a2.25 2.25 0 0 0-2.25 2.25m19.5 0v.243a2.25 2.25 0 0 1-1.07 1.916l-7.5 4.615a2.25 2.25 0 0 1-2.36 0L3.32 8.91a2.25 2.25 0 0 1-1.07-1.916V6.75" />
          </svg>
          {unreadTicketsCount > 0 && (
            <span className={styles.notificationBadge}>{unreadTicketsCount}</span>
          )}
        </button>

        {isNotificacaoOpen && (
          <div className={styles.notificationDropdown}>
            <div className={styles.notifHeader}>
               <h4>Mensagens de {props.departamento}</h4>
            </div>
          <div className={styles.notifBody}>
               {unreadTicketsCount > 0 ? (
                  <div className={styles.notifItem} onClick={abrirModalMeusChamados}>
                     <div className={styles.notifIcon}>💬</div>
                     <div className={styles.notifText}>
                        <p>Você tem <strong>{unreadTicketsCount}</strong> chamado(s) com novas mensagens.</p>
                        <span>Clique para visualizar ➔</span>
                     </div>
                  </div>
               ) : (
                  <div style={{ textAlign: 'center' }}>
                    <p className={styles.notifEmpty}>Tudo limpo! Nenhuma mensagem nova por aqui.</p>
                    <button 
                      onClick={abrirModalMeusChamados}
                      style={{ width: '100%', padding: '15px', background: '#F8FAFC', border: 'none', borderTop: '1px solid #E5E7EB', cursor: 'pointer', color: '#2E5C31', fontWeight: 'bold', fontSize: '14px', transition: '0.2s' }}
                      onMouseEnter={(e) => e.currentTarget.style.background = '#e2e8f0'}
                      onMouseLeave={(e) => e.currentTarget.style.background = '#F8FAFC'}
                    >
                      🎫 Abrir Meus Chamados
                    </button>
                  </div>
               )}
            </div>
          </div>
        )}
      </div>

      {/* 2. O MODAL DE MEUS CHAMADOS */}
      {isModalOpen && (
        <div className={styles.modalOverlay}>
          <div className={styles.modalContent} style={{ width: '750px', maxHeight: '85vh', maxWidth: '95%' }}>
            <header className={styles.modalHeader}>
              <h3>🎫 Meus Chamados - {props.departamento}</h3>
              <button className={styles.closeBtn} onClick={() => setIsModalOpen(false)}>✕</button>
            </header>
            
            <div className={styles.commentsList} style={{ padding: '25px', backgroundColor: '#F8FAFC' }}>
              {loadingChamados ? (
                <div style={{ textAlign: 'center', padding: '40px 0', color: '#6B7280' }}>
                  <p style={{ fontSize: '16px', fontWeight: 'bold' }}>📡 Conectando ao painel...</p>
                  <p style={{ fontSize: '13px' }}>Buscando seus chamados em andamento.</p>
                </div>
              ) : meusChamados.length > 0 ? (
                <div className={styles.ticketsGrid}>
                  
                  {meusChamados.map((ticket: any, index: number) => {
                    const isExpanded = expandedTicketIndex === index;
                    const lastSeen = localStorage.getItem(`grunner_visto_${ticket.id}`);
                    const isEscondido = localStorage.getItem(`grunner_escondido_${ticket.id}`) === "true";
                    const isUnread = ticket.timestampAtualizacao && parseInt(ticket.timestampAtualizacao) > parseInt(lastSeen || '0');
                    const isEncerrado = ticket.status.toLowerCase().includes('encerrado') || ticket.status.toLowerCase().includes('conclu');

                    if (isEscondido && isEncerrado) return null;
                    
                    return (
                      <div key={index} className={styles.ticketCard} style={{ opacity: isEncerrado && !isUnread ? 0.7 : 1 }}>
                        <div className={styles.ticketHeader}>
                          <h4 style={{ display: 'flex', alignItems: 'center', gap: '8px', margin: 0, fontSize: '16px', color: '#171E0D' }}>
                            {ticket.titulo || "Chamado sem título"}
                            {isUnread && <span className={styles.unreadBadge}>🔴 Novo</span>}
                          </h4>
                          <span className={styles.ticketStatus} style={{ backgroundColor: ticket.corStatus || '#A6CE39' }}>
                            {ticket.status || "Sem Status"}
                          </span>
                        </div>
                        
                        <div className={styles.ticketBody}>
                          <p><strong>Filas/Área:</strong> {ticket.area || props.departamento}</p>
                          <p><strong>Criado em:</strong> {ticket.dataCriacao}</p>
                        </div>

                        {isExpanded && (
                          <div className={styles.ticketExpandedArea}>
                            <div className={styles.ticketDetailsBox}>
                              <h5>Descrição do Chamado:</h5>
                              <p>{ticket.descricao ? ticket.descricao : "Nenhuma descrição fornecida."}</p>

                              {ticket.motivoPausa && (
                                <div className={styles.ticketCustomField}>
                                  <strong>⏸️ Motivo da Pausa:</strong>
                                  <p>{ticket.motivoPausa}</p>
                                </div>
                              )}

                              {ticket.comentarioEncerramento && (
                                <div className={styles.ticketCustomField}>
                                  <strong>✅ Comentário de Encerramento:</strong>
                                  <p>{ticket.comentarioEncerramento}</p>
                                </div>
                              )}
                            </div>

                            <div className={styles.chatHistoryArea}>
                              <h5>Histórico de Mensagens:</h5>
                              {loadingHistorico ? (
                                <p style={{ fontSize: '13px', color: '#6B7280', fontStyle: 'italic' }}>Carregando conversas do ClickUp...</p>
                              ) : comentariosDoChamado.length === 0 ? (
                                <p style={{ fontSize: '13px', color: '#9CA3AF', fontStyle: 'italic' }}>Nenhuma mensagem trocada neste chamado ainda.</p>
                              ) : (
                                <div className={styles.chatContainer}>
                                  {comentariosDoChamado.map((c: any, i: number) => (
                                    <div key={i} className={`${styles.chatBubble} ${c.isIntranet ? styles.chatUser : styles.chatIT}`}>
                                      <span className={styles.chatAuthor}>{c.autor} • {c.data}</span>
                                      <p>{c.texto}</p>
                                      
                                      {/* 👇 RENDERIZAÇÃO DOS ANEXOS 👇 */}
                                      {c.anexos && c.anexos.length > 0 && (
                                        <div style={{ marginTop: '10px', display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                                          {c.anexos.map((url: string, anidx: number) => (
                                            <a 
                                              key={anidx} 
                                              href={url} 
                                              target="_blank" 
                                              rel="noopener noreferrer"
                                              style={{ 
                                                padding: '6px 12px', 
                                                backgroundColor: c.isIntranet ? '#A6CE39' : '#E2E8F0', 
                                                color: '#171E0D', 
                                                borderRadius: '6px', 
                                                fontSize: '11px', 
                                                fontWeight: 'bold', 
                                                textDecoration: 'none',
                                                display: 'flex',
                                                alignItems: 'center',
                                                gap: '5px'
                                              }}
                                            >
                                              📎 Ver Anexo
                                            </a>
                                          ))}
                                        </div>
                                      )}
                                    </div>
                                  ))}
                                </div>
                              )}
                            </div>

                            <div className={styles.ticketReplyArea}>
                              <h5>Responder:</h5>
                              <textarea 
                                className={styles.ticketTextarea}
                                placeholder="Digite a sua resposta ou adicione mais informações..."
                                value={novoComentarioChamado}
                                onChange={(e) => setNovoComentarioChamado(e.target.value)}
                                disabled={enviandoComentarioChamado}
                              />
                              
                              {/* BARRA DE AÇÕES (ANEXO + ENVIAR) */}
                              <div className={styles.replyActions}>
                                <div className={styles.attachSection}>
                                  {/* Input invisível que abre a janela de ficheiros */}
                                  <input
                                    type="file"
                                    ref={fileInputRef}
                                    style={{ display: 'none' }}
                                    onChange={(e) => setArquivoAnexo(e.target.files ? e.target.files[0] : null)}
                                  />
                                  <button
                                    className={styles.attachBtn}
                                    onClick={() => fileInputRef.current?.click()}
                                    disabled={enviandoComentarioChamado}
                                    title="Anexar print ou ficheiro"
                                  >
                                    📎 Anexar
                                  </button>
                                  
                                  {/* Pílula com o nome do ficheiro selecionado */}
                                  {arquivoAnexo && (
                                    <span className={styles.attachmentPreview} title={arquivoAnexo.name}>
                                      {arquivoAnexo.name}
                                      <button onClick={() => setArquivoAnexo(null)}>✕</button>
                                    </span>
                                  )}
                                </div>

                                <button 
                                  className={styles.btnReply}
                                  onClick={() => enviarComentarioChamado(ticket.id)}
                                  disabled={enviandoComentarioChamado || (!novoComentarioChamado.trim() && !arquivoAnexo)}
                                >
                                  {enviandoComentarioChamado ? "⏳ A enviar..." : "Enviar ➔"}
                                </button>
                              </div>
                            </div>
                          </div>
                        )}

                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '10px' }}>
                          {isEncerrado && (
                            <button className={styles.btnDismiss} onClick={() => dispensarChamado(ticket.id)}>
                              🗑️ Ocultar
                            </button>
                          )}
                          <button className={styles.btnToggleDetails} onClick={() => toggleDetalhesChamado(index, ticket.id)}>
                            {isExpanded ? "↑ Ocultar detalhes" : "↓ Ver detalhes"}
                          </button>
                        </div>
                      </div>
                    );
                  })}

                </div>
              ) : (
                <div style={{ textAlign: 'center', padding: '40px 0', color: '#6B7280' }}>
                  <p style={{ fontSize: '16px', fontWeight: 'bold' }}>✅ Tudo limpo por aqui!</p>
                  <p style={{ fontSize: '13px' }}>Você não tem chamados abertos atrelados ao seu e-mail.</p>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </>
  );
};