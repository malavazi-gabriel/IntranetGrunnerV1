import * as React from 'react';
import styles from './PainelAtivosGrunner.module.scss';
import { IPainelAtivosGrunnerProps } from './IPainelAtivosGrunnerProps';
import { SharePointService } from '../services/SharePointService';

import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

const logoGrunner = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo-grunner.png";
const logoCompleta = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo.png";
const homeUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Inicio.aspx?env=Embedded";
const atalhosUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/centraldeatalhos.aspx?env=Embedded";

interface IEquipamentoCarrinho {
  tipo: string; fabricante: string; modelo: string; serie: string;
  imei: string; patrimonioFin: string; especificacoes: string;
  observacoes: string; codigoGerado?: string; 
}

interface IPainelState {
  abaAtiva: 'consulta' | 'cadastro';
  isMobileMenuOpen: boolean;
  isSalvando: boolean;
  
  novoNome: string; 
  novoEmailResponsavel: string; 
  novoDepartamento: string;
  novoTipo: string; novoFabricante: string; novoModelo: string;
  novoSerie: string; novoImei: string; novoPatrimonioFin: string;
  novaEspecificacao: string; novaObservacao: string; 
  carrinho: IEquipamentoCarrinho[];

  usuariosSugeridos: any[];
  mostrarSugestoes: boolean;

  ativosSalvos: any[];
  termoBusca: string;
  carregandoConsulta: boolean;

  ativoSendoEditado: any | null;
  editNome: string; editEmail: string; editDepartamento: string;
  editTipo: string; editFabricante: string; editModelo: string;
  editSerie: string; editImei: string; editPatrimonioFin: string;
  editEspecificacao: string; editObservacao: string;

  itensSelecionados: number[];
  mostrarModalTransferenciaLote: boolean;
}

export default class PainelAtivosGrunner extends React.Component<IPainelAtivosGrunnerProps, IPainelState> {
  private _service: SharePointService;
  private footerObserver?: MutationObserver;

  constructor(props: IPainelAtivosGrunnerProps) {
    super(props);
    this._service = new SharePointService(this.props.context);
    this.state = { 
      abaAtiva: 'cadastro', 
      isMobileMenuOpen: false, 
      isSalvando: false,
      novoNome: '', novoEmailResponsavel: '', novoDepartamento: '', novoTipo: 'Notebook', novoFabricante: '', novoModelo: '', novoSerie: '', novoImei: '', novoPatrimonioFin: '', novaEspecificacao: '', novaObservacao: '', carrinho: [],
      usuariosSugeridos: [], mostrarSugestoes: false,
      ativosSalvos: [], termoBusca: '', carregandoConsulta: false,
      ativoSendoEditado: null, editNome: '', editEmail: '', editDepartamento: '', editTipo: '', editFabricante: '', editModelo: '', editSerie: '', editImei: '', editPatrimonioFin: '', editEspecificacao: '', editObservacao: '',
      itensSelecionados: [], mostrarModalTransferenciaLote: false
    };
  }

  private shouldHideSharePointChrome = (): boolean => {
    const search = window.location.search.toLowerCase();
    return (search.includes('env=embedded') || search.includes('mode=embed')) && !search.includes('mode=edit');
  }

  private applyChromeFixes = (): void => {
    const selectors = ['#sp-appBar', '#sp-page-footer', '[data-automation-id="page-bottom-actions"]', '.CommentsWrapper'];
    document.querySelectorAll(selectors.join(',')).forEach((el) => { (el as HTMLElement).style.display = 'none'; });
  }

  public componentDidMount(): void {
    if (this.shouldHideSharePointChrome()) {
      this.applyChromeFixes();
      this.footerObserver = new MutationObserver(() => this.applyChromeFixes());
      if (document.body) this.footerObserver.observe(document.body, { childList: true, subtree: true });
    }
  }

  private carregarAtivosParaConsulta = async () => {
    this.setState({ carregandoConsulta: true, abaAtiva: 'consulta', itensSelecionados: [] });
    try {
      const dados = await this._service.getTodosAtivos();
      this.setState({ ativosSalvos: dados, carregandoConsulta: false });
    } catch (error) {
      console.error(error);
      alert("Erro ao carregar a lista de equipamentos. Verifique a consola.");
      this.setState({ carregandoConsulta: false });
    }
  }

  private toggleSelecao = (id: number) => {
    const selecionados = [...this.state.itensSelecionados];
    const index = selecionados.indexOf(id);
    if (index > -1) {
      selecionados.splice(index, 1);
    } else {
      selecionados.push(id);
    }
    this.setState({ itensSelecionados: selecionados });
  }

  private fecharModal = () => {
    this.setState({ ativoSendoEditado: null, mostrarModalTransferenciaLote: false, mostrarSugestoes: false });
  }

  private salvarTransferenciaLote = async (gerarTermo: boolean) => {
    if (!this.state.editNome) { alert("O Responsável não pode estar vazio."); return; }
    
    this.setState({ isSalvando: true });
    
    try {
      const ativosSelecionadosCompletos = this.state.ativosSalvos.filter(a => this.state.itensSelecionados.includes(a.id));

      for (const ativo of ativosSelecionadosCompletos) {
        await this._service.transferirAtivo(ativo.id, this.state.editNome, this.state.editDepartamento, this.state.editEmail, this.state.editObservacao);
      }

      if (gerarTermo) {
        const itensParaWord = ativosSelecionadosCompletos.filter(a => a.tipo === 'Notebook' || a.tipo === 'Celular / Smartphone');
        
        if (itensParaWord.length > 0) {
          const content = await this._service.getTemplateTermo();
          const zip = new PizZip(content);
          const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

          const equipamentosParaWord = itensParaWord.map((item, index) => {
            const letraItem = String.fromCharCode(97 + index); 
            return {
                letra: letraItem,
                quantidade_tipo: `1 (um) ${item.tipo}`,
                fabricante_modelo: `${item.fabricante} ${item.modelo}`,
                especificacoes: item.especificacoes ? ` - ${item.especificacoes}` : '',
                numero_serie: item.serie || item.imei || "N/A",
                patrimonio: item.patrimonio, 
                patrimonio_fin: item.patrimonioFin || "N/A",
                observacoes: this.state.editObservacao ? this.state.editObservacao : "Sem observações adicionais"
            };
          });

          doc.render({ nome: this.state.editNome, mês: new Date().toLocaleDateString('pt-BR', { month: 'long' }), equipamentos: equipamentosParaWord });
          const blob = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
          saveAs(blob, `Termo_Transferencia_${this.state.editNome.split(' ')[0]}.docx`);
          
          alert(`Sucesso! ${ativosSelecionadosCompletos.length} iten(s) transferido(s) e Termo gerado com os equipamentos principais!`);
        } else {
          alert(`Transferência guardada! Nenhum Termo foi gerado, pois não selecionou nenhum Notebook ou Telemóvel.`);
        }
      } else {
        alert(`Sucesso! ${ativosSelecionadosCompletos.length} equipamento(s) transferido(s).`);
      }

      this.setState({ mostrarModalTransferenciaLote: false, itensSelecionados: [], isSalvando: false });
      this.carregarAtivosParaConsulta(); 
    } catch (error) {
      console.error(error);
      alert("Erro ao transferir equipamentos.");
      this.setState({ isSalvando: false });
    }
  }

  private abrirModalEdicao = (ativo: any) => {
    this.setState({
      ativoSendoEditado: ativo, editNome: ativo.responsavel, editEmail: ativo.emailResponsavel || '', editDepartamento: ativo.departamento,
      editTipo: ativo.tipo, editFabricante: ativo.fabricante, editModelo: ativo.modelo, editSerie: ativo.serie !== "-" ? ativo.serie : "",
      editImei: ativo.tipo === 'Celular / Smartphone' ? (ativo.serie !== "-" ? ativo.serie : "") : "", editPatrimonioFin: ativo.patrimonioFin !== "-" ? ativo.patrimonioFin : "",
      editEspecificacao: ativo.especificacoes, editObservacao: ativo.observacoes,
    });
  }

  private salvarEdicaoIndividual = async (gerarTermo: boolean) => {
    if (!this.state.editNome) { alert("O Responsável não pode estar vazio."); return; }
    this.setState({ isSalvando: true });
    
    const dadosAtualizados = {
      nome: this.state.editNome, departamento: this.state.editDepartamento, tipo: this.state.editTipo,
      fabricante: this.state.editFabricante, modelo: this.state.editModelo, serie: this.state.editSerie,
      imei: this.state.editImei, patrimonioFin: this.state.editPatrimonioFin, especificacao: this.state.editEspecificacao, observacao: this.state.editObservacao
    };

    try {
      await this._service.atualizarAtivo(this.state.ativoSendoEditado.id, dadosAtualizados, this.state.editEmail);
      
      if (gerarTermo && (this.state.editTipo === 'Notebook' || this.state.editTipo === 'Celular / Smartphone')) {
        const content = await this._service.getTemplateTermo();
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        const equipamentosParaWord = [{
            letra: 'a', quantidade_tipo: `1 (um) ${this.state.editTipo}`, fabricante_modelo: `${this.state.editFabricante} ${this.state.editModelo}`,
            especificacoes: this.state.editEspecificacao ? ` - ${this.state.editEspecificacao}` : '', numero_serie: this.state.editSerie || this.state.editImei || "N/A",
            patrimonio: this.state.ativoSendoEditado.patrimonio, patrimonio_fin: this.state.editPatrimonioFin || "N/A", observacoes: this.state.editObservacao ? this.state.editObservacao : "Sem observações adicionais"
        }];

        doc.render({ nome: this.state.editNome, mês: new Date().toLocaleDateString('pt-BR', { month: 'long' }), equipamentos: equipamentosParaWord });
        const blob = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        saveAs(blob, `Termo_Responsabilidade_${this.state.editNome.split(' ')[0]}.docx`);
        
        alert("Transferência guardada e Termo gerado com sucesso!");
      } else {
        alert("Equipamento atualizado com sucesso!");
      }

      this.setState({ ativoSendoEditado: null, isSalvando: false });
      this.carregarAtivosParaConsulta(); 
    } catch (error) {
      console.error(error); alert("Erro ao atualizar o equipamento."); this.setState({ isSalvando: false });
    }
  }

  private adicionarAoCarrinho = () => {
    if (!this.state.novoTipo || !this.state.novoFabricante || !this.state.novoModelo) { alert("Preencha os dados básicos do equipamento."); return; }
    const novoItem: IEquipamentoCarrinho = {
      tipo: this.state.novoTipo, fabricante: this.state.novoFabricante, modelo: this.state.novoModelo, serie: this.state.novoSerie, imei: this.state.novoImei, patrimonioFin: this.state.novoPatrimonioFin, especificacoes: this.state.novaEspecificacao, observacoes: this.state.novaObservacao
    };
    this.setState({ carrinho: [...this.state.carrinho, novoItem], novoTipo: 'Notebook', novoFabricante: '', novoModelo: '', novoSerie: '', novoImei: '', novoPatrimonioFin: '', novaEspecificacao: '', novaObservacao: '' });
  }

  private removerDoCarrinho = (index: number) => {
    const novoCarrinho = [...this.state.carrinho]; novoCarrinho.splice(index, 1);
    this.setState({ carrinho: novoCarrinho });
  }

  // --- ATUALIZADO: Agora recebe o parâmetro gerarTermo ---
  private salvarEGerarTermo = async (gerarTermo: boolean) => {
    if (!this.state.novoNome || this.state.carrinho.length === 0) { alert("Preencha o Responsável e adicione pelo menos um equipamento!"); return; }
    this.setState({ isSalvando: true });

    try {
      const carrinhoProcessado = [...this.state.carrinho];
      for (let i = 0; i < carrinhoProcessado.length; i++) {
        const resultado = await this._service.salvarNovoAtivo(carrinhoProcessado[i], this.state.novoNome, this.state.novoDepartamento, this.state.novoEmailResponsavel);
        carrinhoProcessado[i].codigoGerado = resultado.codigo; 
      }
      
      if (gerarTermo) {
        const itensParaWord = carrinhoProcessado.filter(item => item.tipo === 'Notebook' || item.tipo === 'Celular / Smartphone');

        if (itensParaWord.length > 0) {
          const content = await this._service.getTemplateTermo();
          const zip = new PizZip(content);
          const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

          const equipamentosParaWord = itensParaWord.map((item, index) => {
            const letraItem = String.fromCharCode(97 + index); 
            return {
                letra: letraItem, quantidade_tipo: `1 (um) ${item.tipo}`, fabricante_modelo: `${item.fabricante} ${item.modelo}`,
                especificacoes: item.especificacoes ? ` - ${item.especificacoes}` : '', numero_serie: item.serie || item.imei || "N/A",
                patrimonio: item.codigoGerado, patrimonio_fin: item.patrimonioFin || "N/A", observacoes: item.observacoes ? item.observacoes : "Sem observações adicionais"
            };
          });

          doc.render({ nome: this.state.novoNome, mês: new Date().toLocaleDateString('pt-BR', { month: 'long' }), equipamentos: equipamentosParaWord });
          const blob = doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
          saveAs(blob, `Termo_Responsabilidade_${this.state.novoNome.split(' ')[0]}.docx`);
          
          alert(`Sucesso! ${carrinhoProcessado.length} equipamento(s) guardado(s) e Termo gerado!`);
        } else {
          alert(`Sucesso! ${carrinhoProcessado.length} equipamento(s) guardado(s). Nenhum termo foi gerado (sem Notebooks/Telemóveis).`);
        }
      } else {
        // Se a opção for NÃO gerar termo (apenas salvar)
        alert(`Sucesso! ${carrinhoProcessado.length} equipamento(s) cadastrado(s) diretamente no sistema.`);
      }
      
      this.setState({ isSalvando: false, carrinho: [], novoNome: '', novoEmailResponsavel: '', novoDepartamento: '' });
    } catch (error) {
      console.error(error); alert("Erro ao processar o salvamento."); this.setState({ isSalvando: false });
    }
  };

  private getIconeEquipamento = (tipo: string) => {
    if (tipo.includes('Notebook')) return '💻';
    if (tipo.includes('Celular') || tipo.includes('Smartphone')) return '📱';
    if (tipo.includes('Monitor')) return '🖥️';
    return '🖱️';
  }

  public render(): React.ReactElement<IPainelAtivosGrunnerProps> {
    const userEmail = this.props.context.pageContext.user.email;
    const nomeUsuario = this.props.userDisplayName?.split(' ')[0] || 'Colaborador';
    const dataAtual = new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' });

    const ativosFiltrados = this.state.ativosSalvos.filter(ativo => {
      const termo = this.state.termoBusca.toLowerCase();
      return (
        ativo.responsavel.toLowerCase().includes(termo) ||
        ativo.patrimonio.toLowerCase().includes(termo) ||
        ativo.serie.toLowerCase().includes(termo) ||
        ativo.modelo.toLowerCase().includes(termo)
      );
    });

    return (
      <div className={styles.container}>
        <div className={styles.mobileHeaderBar}>
          <button className={styles.hamburgerBtn} onClick={() => this.setState({ isMobileMenuOpen: true })}>☰ Menu grunnertec</button>
        </div>

        <aside className={`${styles.sidebar} ${this.state.isMobileMenuOpen ? styles.open : ''}`}>
          <div className={styles.logoArea}>
            <img src={logoGrunner} alt="Logo" className={styles.logoSemente} />
            <h2 style={{ whiteSpace: 'nowrap' }}>Intranet grunnertec</h2>
          </div>
          <div className={styles.navGroup}>
            <h3>Navegação</h3>
            <a href={homeUrl}>🏠 Painel Inicial</a>
            <a href={atalhosUrl}>🖥️ Central de Atalhos</a>
          </div>
          <div className={styles.navGroup}>
            <h3>Serviços e Chamados</h3>
              <a href="#" className={styles.active}>💻 Gestão de Ativos (TI)</a>
              <a href="https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5" target="_blank">🖥️ Chamado de TI</a>
          </div>
        </aside>

        <div className={styles.contentArea}>
          <header className={styles.header}>
            <div className={styles.headerLeft}>
              <img src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${userEmail}`} className={styles.userAvatar} />
              <div className={styles.headerText}>
                <h1>Painel de Ativos, {nomeUsuario}!</h1>
                <p>Gestão centralizada do inventário de TI grunnertec.</p>
                <span className={styles.dateBadge}>📅 {dataAtual}</span>
              </div>
            </div>
            <img src={logoCompleta} className={styles.logoCentral} />
          </header>

          <main className={styles.grid}>
            <div className={styles.card}>
              
              <div className={styles.tabsContainer}>
                <button className={this.state.abaAtiva === 'consulta' ? styles.tabActive : styles.tab} onClick={this.carregarAtivosParaConsulta}>🔍 Consulta de Ativos</button>
                <button className={this.state.abaAtiva === 'cadastro' ? styles.tabActive : styles.tab} onClick={() => this.setState({ abaAtiva: 'cadastro' })}>➕ Novo Cadastro</button>
              </div>

              {this.state.abaAtiva === 'cadastro' && (
                <div>
                   <div style={{ backgroundColor: '#f8fafc', padding: '20px', borderRadius: '12px', marginBottom: '25px', border: '1px solid #e2e8f0' }}>
                    <h3 style={{ marginTop: 0, fontSize: '16px', color: '#2E5C31' }}>👤 Dados do Responsável</h3>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                      <div className={styles.inputGroup} style={{ position: 'relative' }}>
                        <label>Responsável (Busca no AD)</label>
                        <input 
                          type="text" value={this.state.novoNome} placeholder="Digite o nome (Ex: Gabriel)" autoComplete="off"
                          onChange={async (e) => {
                            const texto = e.target.value;
                            this.setState({ novoNome: texto });
                            if (texto.length >= 3) {
                              const resultados = await this._service.buscarUsuariosAD(texto);
                              this.setState({ usuariosSugeridos: resultados, mostrarSugestoes: true });
                            } else {
                              this.setState({ mostrarSugestoes: false, novoEmailResponsavel: '' });
                            }
                          }}
                        />
                        {this.state.mostrarSugestoes && this.state.usuariosSugeridos.length > 0 && (
                          <ul style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '0', margin: '5px 0 0 0', listStyle: 'none', zIndex: 10, boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                            {this.state.usuariosSugeridos.map((user, idx) => (
                              <li key={idx} style={{ padding: '10px 15px', cursor: 'pointer', borderBottom: '1px solid #f1f5f9' }} onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f1f5f9'} onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                                onClick={async () => {
                                  this.setState({ novoNome: user.nome, novoEmailResponsavel: user.email, mostrarSugestoes: false });
                                  if (user.email) {
                                    const depto = await this._service.getDepartamentoUsuario(user.email);
                                    if (depto) this.setState({ novoDepartamento: depto });
                                  }
                                }}
                              >
                                <strong>{user.nome}</strong> <br/><span style={{ fontSize: '11px', color: '#64748b' }}>{user.email}</span>
                              </li>
                            ))}
                          </ul>
                        )}
                      </div>
                      <div className={styles.inputGroup}><label>Departamento</label><input value={this.state.novoDepartamento} onChange={(e) => this.setState({ novoDepartamento: e.target.value })} placeholder="Auto-preenchido ou digite manualmente" /></div>
                    </div>
                  </div>

                  <h3 style={{ marginTop: 0, fontSize: '16px', color: '#2E5C31' }}>💻 Adicionar Equipamento à Lista</h3>
                  <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                    <div className={styles.inputGroup}><label>Tipo de Ativo</label><select value={this.state.novoTipo} onChange={(e) => this.setState({ novoTipo: e.target.value })}><option value="Notebook">Notebook</option><option value="Celular / Smartphone">Celular / Smartphone</option><option value="Monitor">Monitor</option><option value="Periférico">Periférico</option></select></div>
                    <div className={styles.inputGroup}><label>Fabricante</label><input value={this.state.novoFabricante} onChange={(e) => this.setState({ novoFabricante: e.target.value })} placeholder="Ex: Dell, Samsung" /></div>
                    <div className={styles.inputGroup}><label>Modelo Exato</label><input value={this.state.novoModelo} onChange={(e) => this.setState({ novoModelo: e.target.value })} placeholder="Ex: Latitude 3420" /></div>
                    <div className={styles.inputGroup}><label>Número de Série</label><input value={this.state.novoSerie} onChange={(e) => this.setState({ novoSerie: e.target.value })} placeholder="Serial Number" /></div>
                    <div className={styles.inputGroup}><label>IMEI (Celulares)</label><input value={this.state.novoImei} onChange={(e) => this.setState({ novoImei: e.target.value })} disabled={this.state.novoTipo !== 'Celular / Smartphone'} placeholder={this.state.novoTipo === 'Celular / Smartphone' ? "Apenas números" : "Bloqueado"} /></div>
                    <div className={styles.inputGroup}><label>Patrimônio Financeiro</label><input value={this.state.novoPatrimonioFin} onChange={(e) => this.setState({ novoPatrimonioFin: e.target.value })} placeholder="Cód. Financeiro (opcional)" /></div>
                    <div className={styles.inputGroup} style={{ gridColumn: 'span 2' }}><label>Especificações Técnicas</label><input value={this.state.novaEspecificacao} onChange={(e) => this.setState({ novaEspecificacao: e.target.value })} placeholder="Ex: i5, 8GB RAM, 256GB SSD" /></div>
                    <div className={styles.inputGroup} style={{ gridColumn: 'span 2' }}><label>Observações / Acessórios Adicionais</label><input value={this.state.novaObservacao} onChange={(e) => this.setState({ novaObservacao: e.target.value })} placeholder="Ex: Acompanha carregador e mouse. Tela riscada." /></div>
                    
                    <div style={{ gridColumn: 'span 2', display: 'flex', justifyContent: 'flex-start' }}>
                      <button onClick={this.adicionarAoCarrinho} style={{ background: '#f1f5f9', border: '1px solid #cbd5e1', padding: '10px 20px', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: '#334155' }}>➕ Adicionar este equipamento à lista</button>
                    </div>
                  </div>

                  {this.state.carrinho.length > 0 && (
                    <div style={{ marginTop: '35px', borderTop: '2px solid #eef1f3', paddingTop: '25px' }}>
                      <h3 style={{ marginTop: 0, fontSize: '16px', color: '#171E0D' }}>📋 Equipamentos na Lista ({this.state.carrinho.length})</h3>
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginBottom: '25px' }}>
                        {this.state.carrinho.map((item, idx) => (
                          <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', padding: '15px', background: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px', boxShadow: '0 2px 5px rgba(0,0,0,0.02)' }}>
                            <div>
                              <strong>{item.tipo} {item.fabricante} {item.modelo}</strong>
                              <p style={{ margin: '5px 0 0 0', fontSize: '13px', color: '#64748b' }}>Série/IMEI: {item.serie || item.imei} | Spec: {item.especificacoes}</p>
                              <p style={{ margin: '5px 0 0 0', fontSize: '12px', color: '#f59e0b', fontStyle: 'italic' }}>Obs: {item.observacoes || "Nenhuma"}</p>
                            </div>
                            <button onClick={() => this.removerDoCarrinho(idx)} style={{ background: 'none', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: '18px' }} title="Remover">🗑️</button>
                          </div>
                        ))}
                      </div>
                      
                      {/* NOVOS BOTÕES DE SALVAMENTO (COM E SEM TERMO) */}
                      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '15px' }}>
                        <button 
                          onClick={() => this.salvarEGerarTermo(false)} 
                          disabled={this.state.isSalvando} 
                          style={{ padding: '12px 25px', background: '#f1f5f9', border: '1px solid #cbd5e1', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: '#475569', transition: 'all 0.2s' }}
                          onMouseEnter={(e) => { e.currentTarget.style.background = '#e2e8f0'; }}
                          onMouseLeave={(e) => { e.currentTarget.style.background = '#f1f5f9'; }}
                        >
                          {this.state.isSalvando ? 'Aguarde...' : `💾 Apenas Salvar no Banco (${this.state.carrinho.length})`}
                        </button>

                        <button 
                          className={styles.btnPrimary} 
                          disabled={this.state.isSalvando} 
                          onClick={() => this.salvarEGerarTermo(true)}
                        >
                          {this.state.isSalvando ? '🚀 Processando...' : `📄 Salvar Tudo e Gerar Termo (${this.state.carrinho.length})`}
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {this.state.abaAtiva === 'consulta' && (
                <div>
                  
                  {this.state.itensSelecionados.length > 0 && (
                    <div style={{ background: '#e0f2fe', border: '1px solid #7dd3fc', padding: '15px 20px', borderRadius: '10px', marginBottom: '25px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', boxShadow: '0 4px 15px rgba(2, 132, 199, 0.1)' }}>
                      <span style={{ fontWeight: 'bold', color: '#0369a1', fontSize: '15px' }}>📦 {this.state.itensSelecionados.length} equipamento(s) selecionado(s) para transferência</span>
                      <button 
                        onClick={() => this.setState({ mostrarModalTransferenciaLote: true, editNome: '', editEmail: '', editDepartamento: '', editObservacao: '' })} 
                        style={{ background: '#0284c7', color: 'white', border: 'none', padding: '12px 24px', borderRadius: '8px', fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 2px 8px rgba(2, 132, 199, 0.3)', transition: 'background 0.2s' }}
                        onMouseEnter={(e) => e.currentTarget.style.background = '#0369a1'}
                        onMouseLeave={(e) => e.currentTarget.style.background = '#0284c7'}
                      >
                        🔄 Transferir em Lote
                      </button>
                    </div>
                  )}

                  <div style={{ marginBottom: '25px', display: 'flex', gap: '10px' }}>
                    <input 
                      type="text" 
                      placeholder="Pesquise por Nome, Património (Ex: N0340), Série, Modelo..." 
                      value={this.state.termoBusca}
                      onChange={(e) => this.setState({ termoBusca: e.target.value })}
                      style={{ flex: 1, padding: '12px 15px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '15px' }}
                    />
                    <button onClick={this.carregarAtivosParaConsulta} style={{ background: '#2E5C31', color: 'white', border: 'none', padding: '0 20px', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold' }}>
                      Atualizar Lista
                    </button>
                  </div>

                  {this.state.carregandoConsulta ? (
                    <div style={{ textAlign: 'center', padding: '40px', color: '#64748b' }}>⏳ A carregar banco de dados...</div>
                  ) : (
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '20px' }}>
                      {ativosFiltrados.length > 0 ? (
                        ativosFiltrados.map(ativo => (
                          <div key={ativo.id} style={{ background: this.state.itensSelecionados.includes(ativo.id) ? '#f0fdf4' : '#ffffff', border: '1px solid', borderColor: this.state.itensSelecionados.includes(ativo.id) ? '#86efac' : '#e2e8f0', borderRadius: '12px', padding: '20px', boxShadow: '0 4px 6px rgba(0,0,0,0.02)', transition: 'all 0.2s', cursor: 'default' }}>
                            
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '15px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                <input 
                                  type="checkbox" 
                                  checked={this.state.itensSelecionados.includes(ativo.id)} 
                                  onChange={() => this.toggleSelecao(ativo.id)} 
                                  style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#2E5C31' }}
                                />
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                  <span style={{ background: '#dcfce7', color: '#b45309', padding: '4px 10px', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', display: 'inline-block' }}>
                                    TI: {ativo.patrimonio}
                                  </span>
                                  {ativo.patrimonioFin && ativo.patrimonioFin !== "-" && (
                                    <span style={{ background: '#f1f5f9', color: '#475569', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', display: 'inline-block' }}>
                                      FIN: {ativo.patrimonioFin}
                                    </span>
                                  )}
                                </div>
                              </div>
                              <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                                <button 
                                  onClick={() => this.abrirModalEdicao(ativo)} 
                                  style={{ background: 'transparent', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '6px 10px', cursor: 'pointer', fontSize: '12px', color: '#475569', fontWeight: 'bold' }}
                                >
                                  ✏️
                                </button>
                                <span style={{ fontSize: '24px', background: '#f8fafc', padding: '8px', borderRadius: '50%' }}>
                                  {this.getIconeEquipamento(ativo.tipo)}
                                </span>
                              </div>
                            </div>

                            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '15px', opacity: ativo.responsavel.includes('Estoque') ? 0.6 : 1 }}>
                              {ativo.emailResponsavel ? (
                                <img src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${ativo.emailResponsavel}`} alt={ativo.responsavel} style={{ width: '42px', height: '42px', borderRadius: '50%', objectFit: 'cover', border: '2px solid #e2e8f0' }} onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling && ((e.currentTarget.nextElementSibling as HTMLElement).style.display = 'flex'); }} />
                              ) : null}
                              <div style={{ width: '42px', height: '42px', borderRadius: '50%', background: '#f1f5f9', display: ativo.emailResponsavel ? 'none' : 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '20px' }}>👤</div>
                              <div>
                                <h4 style={{ margin: '0 0 3px 0', color: '#0f172a', fontSize: '15px', fontWeight: 'bold' }}>{ativo.responsavel}</h4>
                                <p style={{ margin: '0', color: '#64748b', fontSize: '12px' }}>🏢 {ativo.departamento || "Sem departamento"}</p>
                              </div>
                            </div>
                            
                            <hr style={{ border: '0', borderTop: '1px dashed #e2e8f0', margin: '15px 0' }}/>

                            <div style={{ fontSize: '13px', color: '#334155', lineHeight: '1.6' }}>
                              <p style={{ margin: '0' }}><strong>Equipamento:</strong> {ativo.tipo} {ativo.fabricante} {ativo.modelo}</p>
                              <p style={{ margin: '0' }}><strong>SN/IMEI:</strong> {ativo.serie}</p>
                              {ativo.especificacoes && <p style={{ margin: '0' }}><strong>Spec:</strong> {ativo.especificacoes}</p>}
                              <p style={{ margin: '0', color: '#94a3b8', fontSize: '11px', marginTop: '8px' }}>Cadastrado em {ativo.dataCriacao}</p>
                              
                              {ativo.observacoes && ativo.observacoes !== "Sem observações adicionais" && ativo.observacoes !== "Sem observações" && (
                                 <div style={{ marginTop: '12px', padding: '10px', background: '#fffbeb', borderLeft: '3px solid #f59e0b', borderRadius: '4px', color: '#92400e', fontSize: '12px', fontStyle: 'italic' }}>
                                   ⚠️ {ativo.observacoes}
                                 </div>
                              )}
                            </div>

                          </div>
                        ))
                      ) : (
                        <div style={{ gridColumn: '1 / -1', textAlign: 'center', padding: '40px', color: '#94a3b8', background: '#f8fafc', borderRadius: '12px' }}>
                          Nenhum equipamento encontrado.
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

            </div>
          </main>
        </div>

        {/* MODAL DE EDIÇÃO INDIVIDUAL */}
        {this.state.ativoSendoEditado && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, backdropFilter: 'blur(3px)' }}>
            <div style={{ background: 'white', padding: '35px', borderRadius: '12px', width: '90%', maxWidth: '800px', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 20px 40px rgba(0,0,0,0.2)' }}>
              
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '25px', borderBottom: '1px solid #e2e8f0', paddingBottom: '15px' }}>
                <h2 style={{ margin: 0, color: '#2E5C31', fontSize: '20px' }}>✏️ Editando Transferência: {this.state.ativoSendoEditado.patrimonio}</h2>
                <button onClick={this.fecharModal} style={{ background: 'none', border: 'none', fontSize: '20px', cursor: 'pointer', color: '#64748b' }}>❌</button>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                <div className={styles.inputGroup} style={{ position: 'relative' }}>
                  <label style={{ color: '#b45309' }}>Responsável (Digite 'Estoque' para devolver)</label>
                  <input type="text" value={this.state.editNome} placeholder="Ex: Estoque TI, ou nome da pessoa" autoComplete="off"
                    onChange={async (e) => {
                      const texto = e.target.value;
                      this.setState({ editNome: texto });
                      if (texto.length >= 3 && texto.toLowerCase() !== "estoque") {
                        const resultados = await this._service.buscarUsuariosAD(texto);
                        this.setState({ usuariosSugeridos: resultados, mostrarSugestoes: true });
                      } else {
                        this.setState({ mostrarSugestoes: false, editEmail: '' }); 
                      }
                    }}
                  />
                  {this.state.mostrarSugestoes && this.state.usuariosSugeridos.length > 0 && (
                    <ul style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '0', margin: '5px 0 0 0', listStyle: 'none', zIndex: 10, boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                      {this.state.usuariosSugeridos.map((user, idx) => (
                        <li key={idx} style={{ padding: '10px 15px', cursor: 'pointer', borderBottom: '1px solid #f1f5f9' }} onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f1f5f9'} onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                          onClick={async () => {
                            this.setState({ editNome: user.nome, editEmail: user.email, mostrarSugestoes: false });
                            if (user.email) {
                              const depto = await this._service.getDepartamentoUsuario(user.email);
                              if (depto) this.setState({ editDepartamento: depto });
                            }
                          }}
                        >
                          <strong>{user.nome}</strong> <br/><span style={{ fontSize: '11px', color: '#64748b' }}>{user.email}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
                
                <div className={styles.inputGroup}><label>Departamento</label><input value={this.state.editDepartamento} onChange={(e) => this.setState({ editDepartamento: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>Tipo de Ativo</label><select value={this.state.editTipo} onChange={(e) => this.setState({ editTipo: e.target.value })}><option value="Notebook">Notebook</option><option value="Celular / Smartphone">Celular / Smartphone</option><option value="Monitor">Monitor</option><option value="Periférico">Periférico</option></select></div>
                <div className={styles.inputGroup}><label>Fabricante</label><input value={this.state.editFabricante} onChange={(e) => this.setState({ editFabricante: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>Modelo Exato</label><input value={this.state.editModelo} onChange={(e) => this.setState({ editModelo: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>Número de Série</label><input value={this.state.editSerie} onChange={(e) => this.setState({ editSerie: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>IMEI</label><input value={this.state.editImei} onChange={(e) => this.setState({ editImei: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>Património Fin.</label><input value={this.state.editPatrimonioFin} onChange={(e) => this.setState({ editPatrimonioFin: e.target.value })} /></div>
                <div className={styles.inputGroup} style={{ gridColumn: 'span 2' }}><label>Especificações</label><input value={this.state.editEspecificacao} onChange={(e) => this.setState({ editEspecificacao: e.target.value })} /></div>
                <div className={styles.inputGroup} style={{ gridColumn: 'span 2' }}><label>Observações / Status atual</label><input value={this.state.editObservacao} onChange={(e) => this.setState({ editObservacao: e.target.value })} placeholder="Ex: Devolvido para o estoque. Aguarda formatação." style={{ borderColor: '#f59e0b', background: '#fffbeb' }} /></div>
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '15px', marginTop: '30px' }}>
                <button onClick={this.fecharModal} style={{ padding: '12px 25px', background: 'transparent', border: '1px solid #94a3b8', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: '#475569' }}>Cancelar</button>
                
                {(this.state.editTipo === 'Notebook' || this.state.editTipo === 'Celular / Smartphone') && (
                  <button onClick={() => this.salvarEdicaoIndividual(true)} disabled={this.state.isSalvando} style={{ padding: '12px 20px', background: '#0284c7', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: 'white', boxShadow: '0 4px 10px rgba(2, 132, 199, 0.3)' }}>
                    {this.state.isSalvando ? 'Aguarde...' : '📄 Salvar e Gerar Termo'}
                  </button>
                )}

                <button onClick={() => this.salvarEdicaoIndividual(false)} disabled={this.state.isSalvando} style={{ padding: '12px 30px', background: '#2E5C31', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: 'white', boxShadow: '0 4px 10px rgba(46, 92, 49, 0.3)' }}>
                  {this.state.isSalvando ? 'A guardar...' : '💾 Apenas Salvar'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* MODAL DE TRANSFERÊNCIA EM LOTE */}
        {this.state.mostrarModalTransferenciaLote && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, backdropFilter: 'blur(3px)' }}>
            <div style={{ background: 'white', padding: '35px', borderRadius: '12px', width: '90%', maxWidth: '600px', boxShadow: '0 20px 40px rgba(0,0,0,0.2)' }}>
              
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '25px', borderBottom: '1px solid #e2e8f0', paddingBottom: '15px' }}>
                <h2 style={{ margin: 0, color: '#0369a1', fontSize: '20px' }}>🔄 Transferir {this.state.itensSelecionados.length} Itens</h2>
                <button onClick={this.fecharModal} style={{ background: 'none', border: 'none', fontSize: '20px', cursor: 'pointer', color: '#64748b' }}>❌</button>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div className={styles.inputGroup} style={{ position: 'relative' }}>
                  <label style={{ color: '#0369a1' }}>Novo Responsável (para todos os {this.state.itensSelecionados.length} itens)</label>
                  <input type="text" value={this.state.editNome} placeholder="Ex: Gabriel Henrique..." autoComplete="off"
                    onChange={async (e) => {
                      const texto = e.target.value;
                      this.setState({ editNome: texto });
                      if (texto.length >= 3 && texto.toLowerCase() !== "estoque") {
                        const resultados = await this._service.buscarUsuariosAD(texto);
                        this.setState({ usuariosSugeridos: resultados, mostrarSugestoes: true });
                      } else {
                        this.setState({ mostrarSugestoes: false, editEmail: '' }); 
                      }
                    }}
                  />
                  {this.state.mostrarSugestoes && this.state.usuariosSugeridos.length > 0 && (
                    <ul style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '0', margin: '5px 0 0 0', listStyle: 'none', zIndex: 10, boxShadow: '0 4px 6px rgba(0,0,0,0.1)' }}>
                      {this.state.usuariosSugeridos.map((user, idx) => (
                        <li key={idx} style={{ padding: '10px 15px', cursor: 'pointer', borderBottom: '1px solid #f1f5f9' }} onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f1f5f9'} onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                          onClick={async () => {
                            this.setState({ editNome: user.nome, editEmail: user.email, mostrarSugestoes: false });
                            if (user.email) {
                              const depto = await this._service.getDepartamentoUsuario(user.email);
                              if (depto) this.setState({ editDepartamento: depto });
                            }
                          }}
                        >
                          <strong>{user.nome}</strong> <br/><span style={{ fontSize: '11px', color: '#64748b' }}>{user.email}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
                
                <div className={styles.inputGroup}><label>Departamento</label><input value={this.state.editDepartamento} onChange={(e) => this.setState({ editDepartamento: e.target.value })} /></div>
                
                <div className={styles.inputGroup}>
                  <label>Observação da Transferência (aplicada a todos os itens)</label>
                  <input value={this.state.editObservacao} onChange={(e) => this.setState({ editObservacao: e.target.value })} placeholder="Ex: Equipamentos entregues ao novo funcionário." />
                </div>
              </div>

              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '15px', marginTop: '30px' }}>
                <button onClick={this.fecharModal} style={{ padding: '12px 25px', background: 'transparent', border: '1px solid #94a3b8', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: '#475569' }}>Cancelar</button>
                
                <button onClick={() => this.salvarTransferenciaLote(true)} disabled={this.state.isSalvando} style={{ padding: '12px 20px', background: '#0284c7', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: 'white', boxShadow: '0 4px 10px rgba(2, 132, 199, 0.3)' }}>
                  {this.state.isSalvando ? 'Aguarde...' : '📄 Transferir e Gerar Termo'}
                </button>

                <button onClick={() => this.salvarTransferenciaLote(false)} disabled={this.state.isSalvando} style={{ padding: '12px 30px', background: '#2E5C31', border: 'none', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', color: 'white', boxShadow: '0 4px 10px rgba(46, 92, 49, 0.3)' }}>
                  {this.state.isSalvando ? 'A guardar...' : '💾 Apenas Transferir'}
                </button>
              </div>
            </div>
          </div>
        )}

      </div>
    );
  }
}