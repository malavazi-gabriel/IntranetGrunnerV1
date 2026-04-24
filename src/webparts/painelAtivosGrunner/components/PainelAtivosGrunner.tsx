import * as React from 'react';
import styles from './PainelAtivosGrunner.module.scss';
import { IPainelAtivosGrunnerProps } from './IPainelAtivosGrunnerProps';
import { SharePointService } from '../services/SharePointService';
import { MenuChamados } from '../../../shared/components/MenuChamado/MenuChamados';

import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';

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
  abaAtiva: 'consulta' | 'cadastro' | 'acessos'; 
  isMobileMenuOpen: boolean;
  isMenuTIOpen: boolean;
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

  filtroTipo: string;
  filtroDepartamento: string;
  ativoAuditoria: any | null;
  dadosAuditoria: any[];
  carregandoAuditoria: boolean;

  // Variáveis de Acesso Dinâmico
  verificandoAcessos: boolean;
  isTI: boolean;
  isVisualizador: boolean;
  isAdminView: boolean;

  // Variáveis da Aba de Gerenciamento de Acessos
  listaAcessos: any[];
  novoNomeAcesso: string;
  novoEmailAcesso: string;
  novoNivelAcesso: string;
  isIframeModalOpen: boolean;
  iframeUrl: string;
  iframeTitle: string;
}

export default class PainelAtivosGrunner extends React.Component<IPainelAtivosGrunnerProps, IPainelState> {
  private _service: SharePointService;
  private footerObserver?: MutationObserver;

  constructor(props: IPainelAtivosGrunnerProps) {
    super(props);
    this._service = new SharePointService(this.props.context);
    this.state = { 
      abaAtiva: 'consulta', 
      isMobileMenuOpen: false, 
      isMenuTIOpen: false,
      isSalvando: false,
      novoNome: '', novoEmailResponsavel: '', novoDepartamento: '', novoTipo: 'Notebook', novoFabricante: '', novoModelo: '', novoSerie: '', novoImei: '', novoPatrimonioFin: '', novaEspecificacao: '', novaObservacao: '', carrinho: [],
      usuariosSugeridos: [], mostrarSugestoes: false,
      ativosSalvos: [], termoBusca: '', carregandoConsulta: false,
      ativoSendoEditado: null, editNome: '', editEmail: '', editDepartamento: '', editTipo: '', editFabricante: '', editModelo: '', editSerie: '', editImei: '', editPatrimonioFin: '', editEspecificacao: '', editObservacao: '',
      itensSelecionados: [], mostrarModalTransferenciaLote: false,
      filtroTipo: '', filtroDepartamento: '', ativoAuditoria: null, dadosAuditoria: [], carregandoAuditoria: false,
      
      verificandoAcessos: true,
      isTI: false,
      isVisualizador: false,
      isAdminView: false,

      // Estados iniciais do Painel de Acessos
      listaAcessos: [],
      novoNomeAcesso: '',
      novoEmailAcesso: '',
      novoNivelAcesso: 'Visualizador',
      isIframeModalOpen: false,
      iframeUrl: '',
      iframeTitle: ''
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
    
    this.inicializarAplicacao();
  }

  private abrirModalFormulario = (url: string, titulo: string, e: React.MouseEvent) => {
    e.preventDefault(); 
    this.setState({ 
      isIframeModalOpen: true, 
      iframeUrl: url, 
      iframeTitle: titulo 
    });
  }

  private inicializarAplicacao = async () => {
    const userEmailLogado = this.props.context.pageContext.user.email;
    const acessos = await this._service.verificarAcessoUsuario(userEmailLogado);
    
    this.setState({
      isTI: acessos.isTI,
      isVisualizador: acessos.isVisualizador,
      isAdminView: acessos.isTI || acessos.isVisualizador,
      verificandoAcessos: false
    }, () => {
      this.carregarAtivosParaConsulta();
    });
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

  // --- MÉTODOS DA TELA DE ACESSOS ---
  private carregarAcessosUI = async () => {
    try {
      const acessos = await this._service.getTodosAcessos();
      this.setState({ listaAcessos: acessos });
    } catch (error) {
      console.error(error);
    }
  }

  private salvarNovoAcesso = async () => {
    if (!this.state.novoEmailAcesso || !this.state.novoNivelAcesso) {
      alert("Preencha ao menos o e-mail e o nível de acesso.");
      return;
    }
    this.setState({ isSalvando: true });
    try {
      await this._service.adicionarAcesso(this.state.novoNomeAcesso, this.state.novoEmailAcesso, this.state.novoNivelAcesso);
      this.setState({ novoNomeAcesso: '', novoEmailAcesso: '', novoNivelAcesso: 'Visualizador', isSalvando: false });
      this.carregarAcessosUI();
      alert("Acesso concedido com sucesso!");
    } catch (error) {
      console.error(error);
      alert("Erro ao salvar o acesso.");
      this.setState({ isSalvando: false });
    }
  }

  private deletarAcesso = async (id: number) => {
    if (confirm("Tem certeza que deseja remover as permissões deste usuário?")) {
      try {
        await this._service.removerAcesso(id);
        this.carregarAcessosUI();
      } catch (error) {
        console.error(error);
        alert("Erro ao remover o acesso.");
      }
    }
  }
  // ------------------------------------

  private getAtivosFiltrados = () => {
    return this.state.ativosSalvos.filter(ativo => {
      const termo = this.state.termoBusca.toLowerCase();
      const res = ativo.responsavel ? ativo.responsavel.toLowerCase() : "";
      const pat = ativo.patrimonio ? ativo.patrimonio.toLowerCase() : "";
      const ser = ativo.serie ? ativo.serie.toLowerCase() : "";
      const mod = ativo.modelo ? ativo.modelo.toLowerCase() : "";
      const tip = ativo.tipo ? ativo.tipo.toLowerCase() : "";

      const passaBusca = res.includes(termo) || pat.includes(termo) || ser.includes(termo) || mod.includes(termo);
      
      let passaTipo = true;
      if (this.state.filtroTipo) {
        const filtroNormalizado = this.state.filtroTipo.split('/')[0].trim().toLowerCase(); 
        passaTipo = tip.includes(filtroNormalizado);
      }

      const passaDepto = this.state.filtroDepartamento ? ativo.departamento === this.state.filtroDepartamento : true;
      
      return passaBusca && passaTipo && passaDepto;
    });
  }

  private exportarParaExcel = () => {
    const listaFiltrada = this.getAtivosFiltrados();
    if (listaFiltrada.length === 0) { alert("Não há dados para exportar com os filtros atuais."); return; }

    const dadosMapeados = listaFiltrada.map(ativo => ({
      "Patrimônio TI": ativo.patrimonio,
      "Patrimônio Financeiro": ativo.patrimonioFin !== "-" ? ativo.patrimonioFin : "",
      "Responsável": ativo.responsavel,
      "Departamento": ativo.departamento,
      "Tipo de Equipamento": ativo.tipo,
      "Fabricante": ativo.fabricante,
      "Modelo": ativo.modelo,
      "Nº de Série / IMEI": ativo.serie,
      "Especificações": ativo.especificacoes,
      "Observações": ativo.observacoes,
      "Data de Cadastro": ativo.dataCriacao
    }));

    const worksheet = XLSX.utils.json_to_sheet(dadosMapeados);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Inventário Grunnertec");
    
    const wscols = [{wch:15}, {wch:15}, {wch:30}, {wch:20}, {wch:15}, {wch:15}, {wch:25}, {wch:25}, {wch:35}, {wch:30}, {wch:15}];
    worksheet['!cols'] = wscols;

    XLSX.writeFile(workbook, `Inventario_Grunnertec_${new Date().toLocaleDateString('pt-BR').replace(/\//g, '-')}.xlsx`);
  }

  private abrirAuditoria = async (ativo: any) => {
    this.setState({ ativoAuditoria: ativo, carregandoAuditoria: true, dadosAuditoria: [] });
    const historico = await this._service.getHistoricoAtivo(ativo.id);
    this.setState({ dadosAuditoria: historico, carregandoAuditoria: false });
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
    this.setState({ ativoSendoEditado: null, mostrarModalTransferenciaLote: false, mostrarSugestoes: false, ativoAuditoria: null });
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
        const itensParaWord = ativosSelecionadosCompletos.filter(a => a.tipo !== 'Periférico');
        
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
          alert(`Transferência guardada! Nenhum Termo foi gerado, pois selecionou apenas periféricos.`);
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
      
      if (gerarTermo && this.state.editTipo !== 'Periférico') {
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
        const itensParaWord = carrinhoProcessado.filter(item => item.tipo !== 'Periférico');

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
          alert(`Sucesso! ${carrinhoProcessado.length} equipamento(s) guardado(s). "Sucesso! Nenhum termo foi gerado (apenas Periféricos selecionados).`);
        }
      } else {
        alert(`Sucesso! ${carrinhoProcessado.length} equipamento(s) cadastrado(s) diretamente no sistema.`);
      }
      
      this.setState({ isSalvando: false, carrinho: [], novoNome: '', novoEmailResponsavel: '', novoDepartamento: '', abaAtiva: 'consulta' });
      this.carregarAtivosParaConsulta();
    } catch (error) {
      console.error(error); alert("Erro ao processar o salvamento."); this.setState({ isSalvando: false });
    }
  };

  private getIconeEquipamento = (tipo: string) => {
    if (tipo.includes('Notebook')) return '💻';
    if (tipo.includes('Desktop')) return '🖥️';
    if (tipo.includes('Celular') || tipo.includes('Smartphone')) return '📱';
    if (tipo.includes('Tablet')) return '📋';
    if (tipo.includes('Monitor')) return '📺';
    return '🖱️';
  }

  public render(): React.ReactElement<IPainelAtivosGrunnerProps> {
    if (this.state.verificandoAcessos) {
      return (
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh', background: '#F8FAFC', flexDirection: 'column', gap: '20px' }}>
          <img src={logoGrunner} alt="Carregando" style={{ width: '80px', animation: 'pulse 1.5s infinite' }} />
          <h2 style={{ color: '#2E5C31', fontFamily: 'Segoe UI' }}>A carregar permissões do sistema...</h2>
        </div>
      );
    }

    const { isTI, isAdminView } = this.state;
    const userEmailLogado = this.props.context.pageContext.user.email.toLowerCase();
    const nomeUsuario = this.props.userDisplayName?.split(' ')[0] || 'Colaborador';
    const dataAtual = new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' });

    const ativosParaExibir = isAdminView 
      ? this.getAtivosFiltrados() 
      : this.state.ativosSalvos.filter(a => a.emailResponsavel && a.emailResponsavel.toLowerCase() === userEmailLogado);
      
    // DASHBOARD MATEMÁTICA
    const totalEquipamentos = this.state.ativosSalvos.length;
    const qtdEstoque = this.state.ativosSalvos.filter(a => a.responsavel && a.responsavel.toLowerCase().includes('estoque')).length;
    const qtdManutencao = this.state.ativosSalvos.filter(a => a.responsavel && (a.responsavel.toLowerCase().includes('manuten') || a.responsavel.toLowerCase().includes('conserto') || a.responsavel.toLowerCase().includes('bancada'))).length;
    
    const qtdAtivos = totalEquipamentos - qtdEstoque - qtdManutencao;

    const qtdNotebooks = this.state.ativosSalvos.filter(a => a.tipo && a.tipo.toLowerCase().includes('notebook')).length;
    const qtdDesktops = this.state.ativosSalvos.filter(a => a.tipo && a.tipo.toLowerCase().includes('desktop')).length;
    const qtdCelulares = this.state.ativosSalvos.filter(a => a.tipo && (a.tipo.toLowerCase().includes('celular') || a.tipo.toLowerCase().includes('smartphone'))).length;
    const qtdMonitores = this.state.ativosSalvos.filter(a => a.tipo && a.tipo.toLowerCase().includes('monitor')).length;
    const qtdPerifericos = this.state.ativosSalvos.filter(a => a.tipo && a.tipo.toLowerCase().includes('perif')).length;
    const qtdTablets = this.state.ativosSalvos.filter(a => a.tipo && a.tipo.toLowerCase().includes('tablet')).length;

    const qtdOutros = totalEquipamentos - (qtdNotebooks + qtdDesktops + qtdCelulares + qtdMonitores + qtdPerifericos + qtdTablets);

    const departamentosUnicos = Array.from(new Set(this.state.ativosSalvos.map(a => a.departamento).filter(d => d && d.trim() !== ""))).sort();

    return (
      <div className={styles.container}>
        <div className={styles.mobileHeaderBar}>
          <button className={styles.hamburgerBtn} onClick={() => this.setState({ isMobileMenuOpen: true })}>☰ Menu grunnertec</button>
        </div>

<aside className={`${styles.sidebar} ${this.state.isMobileMenuOpen ? styles.open : ''}`}>
          <div className={styles.logoArea}>
            <img src={logoGrunner} alt="Logo Semente" className={styles.logoSemente} />
            <h2>Intranet Grunner</h2>
          </div>

          <div className={styles.navGroup}>
            <h3>Navegação</h3>
            <a href={homeUrl}>🏠 Painel Inicial</a>
            <a href={atalhosUrl}>🖥️ Central de Atalhos</a>
          </div>
          
          <div className={styles.navGroup}>
            <h3>Serviços e Chamados</h3>
            
            {/* BOTÃO PRINCIPAL DE TI (ACORDEÃO) */}
            <a 
              href="#" 
              className={`${styles.menuToggle} ${this.state.isMenuTIOpen ? styles.active : ''}`}
              onClick={(e) => { e.preventDefault(); this.setState({ isMenuTIOpen: !this.state.isMenuTIOpen }); }}
            >
              <span style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>💻 Tecnologia (TI)</span>
              <span style={{ fontSize: '10px', opacity: 0.8 }}>{this.state.isMenuTIOpen ? '▲' : '▼'}</span>
            </a>

            {/* SUB-ITENS DE TI (Só aparecem se estiver aberto) */}
            {this.state.isMenuTIOpen && (
              <div className={styles.navSubGroup}>
                <a href="#" className={styles.active}>🖥️ {isAdminView ? 'Gestão de Ativos' : 'Meus Equipamentos'}</a>
                <a href="#" onClick={(e) => this.abrirModalFormulario("https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5", "➕ Abrir Novo Chamado", e)}>➕ Abrir Novo Chamado</a>
                <a href="#" onClick={(e) => { e.preventDefault(); window.dispatchEvent(new CustomEvent('abrirMeusChamadosGrunner', { detail: 'TI' })); }}>🎫 Meus Chamados</a>
              </div>
            )}

            {/* RESTANTE DOS DEPARTAMENTOS */}
            <a href="#" onClick={(e) => this.abrirModalFormulario("https://grunnerteccombr.sharepoint.com/sites/Marketing/_layouts/15/listforms.aspx?cid=MTQ1MjlmMzEtNjk2Ni00MTI2LWJhNzItMzE1MTc0NDU2YTE4&nav=MGIwZDdiNzMtODQwNi00MDhiLTk5ZDEtNGE5NWNlYzljNDg3&env=Embedded", "📢 Solicitação - Marketing", e)}>📢 Marketing</a>
            <a href="#" onClick={(e) => this.abrirModalFormulario("https://grunnerteccombr.sharepoint.com/sites/GPS/_layouts/15/listforms.aspx?cid=ZWFlMDE1MWUtOTFlMS00MmJiLWFiNzEtOWM0NGVkZTVkMTdh&nav=ZGJmNmMxZGMtNjU5Zi00ZTUxLThjMTctZmFhODY5YTQ3NjBi&env=Embedded", "🚗 Solicitação - Frotas", e)}>🚗 Frotas</a>
            <a href="#" onClick={(e) => this.abrirModalFormulario("https://forms.monday.com/forms/embed/2a2a29caa20e7e1517cc397586af97eb?r=use1", "🛠️ Solicitação - Facilities", e)}>🛠️ Facilities</a>
          </div>
          <div className={styles.navGroup}>
            <h3>Institucional</h3>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">🏛️ Nossa História</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">📖 Políticas da Empresa</a>
          </div>
        </aside>

        <div className={styles.contentArea}>
          <header className={styles.header}>
          <MenuChamados 
             departamento="TI" 
             emailUsuario={this.props.context.pageContext.user.email} 
              />
                      <div className={styles.headerLeft}>
                        <img 
                          src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${userEmailLogado}`} 
                          className={styles.userAvatar} 
                          alt="Perfil"
                          onError={(e) => { e.currentTarget.style.display = 'none'; }}
                        />
                        <div className={styles.headerText}>
                          <h1>{isAdminView ? 'Painel de Ativos' : 'Meus Equipamentos'}, {nomeUsuario}!</h1>
                          <p>{isAdminView ? 'Gestão centralizada do inventário de TI Grunner.' : 'Confira os equipamentos registrados sob sua responsabilidade.'}</p>
                          <span className={styles.dateBadge}>📅 {dataAtual.charAt(0).toUpperCase() + dataAtual.slice(1)}</span>
                        </div>
                      </div>
                      <img src={logoCompleta} className={styles.logoCentral} alt="Grunner" />
              </header>
          <main className={styles.grid}>
            <div className={styles.card}>
              
              {isAdminView && (
                <div className={styles.tabsContainer}>
                  <button className={this.state.abaAtiva === 'consulta' ? styles.tabActive : styles.tab} onClick={this.carregarAtivosParaConsulta}>🔍 Consulta de Ativos</button>
                  {isTI && (
                    <>
                      <button className={this.state.abaAtiva === 'cadastro' ? styles.tabActive : styles.tab} onClick={() => this.setState({ abaAtiva: 'cadastro' })}>➕ Novo Cadastro</button>
                      <button className={this.state.abaAtiva === 'acessos' ? styles.tabActive : styles.tab} onClick={() => { this.setState({ abaAtiva: 'acessos' }); this.carregarAcessosUI(); }}>⚙️ Gerir Acessos</button>
                    </>
                  )}
                </div>
              )}

              {/* === NOVA ABA: GERENCIAR ACESSOS === */}
              {this.state.abaAtiva === 'acessos' && isTI && (
                <div style={{ padding: '10px' }}>
                  <h2 style={{ color: '#2E5C31', marginTop: 0, marginBottom: '10px' }}>⚙️ Gerenciamento de Acessos</h2>
                  <p style={{ color: '#64748b', marginBottom: '25px', fontSize: '14px' }}>Adicione ou remova as permissões para os usuários do painel. Colaboradores comuns que não estiverem na lista terão acesso restrito apenas aos seus próprios equipamentos.</p>

                  <div style={{ background: '#f8fafc', padding: '25px', borderRadius: '12px', marginBottom: '35px', border: '1px solid #e2e8f0' }}>
                    <h3 style={{ marginTop: 0, fontSize: '15px', color: '#0f172a', display: 'flex', alignItems: 'center', gap: '8px' }}>➕ Conceder Novo Acesso</h3>
                    <div style={{ display: 'flex', gap: '15px', alignItems: 'flex-end', flexWrap: 'wrap' }}>
                      <div className={styles.inputGroup} style={{ flex: '1 1 200px' }}>
                        <label>Nome do Usuário</label>
                        <input value={this.state.novoNomeAcesso} onChange={e => this.setState({ novoNomeAcesso: e.target.value })} placeholder="Ex: Maria Souza" />
                      </div>
                      <div className={styles.inputGroup} style={{ flex: '1 1 250px' }}>
                        <label>E-mail da Empresa</label>
                        <input value={this.state.novoEmailAcesso} onChange={e => this.setState({ novoEmailAcesso: e.target.value })} placeholder="exemplo@grunnertec.com.br" />
                      </div>
                      <div className={styles.inputGroup} style={{ flex: '1 1 150px' }}>
                        <label>Nível de Permissão</label>
                        <select value={this.state.novoNivelAcesso} onChange={e => this.setState({ novoNivelAcesso: e.target.value })}>
                          <option value="Visualizador">Apenas Visualização (Gerência/RH)</option>
                          <option value="TI">Equipe de TI (Acesso Total)</option>
                        </select>
                      </div>
                      <button onClick={this.salvarNovoAcesso} disabled={this.state.isSalvando} style={{ background: '#2E5C31', color: 'white', border: 'none', padding: '14px 30px', borderRadius: '10px', cursor: 'pointer', fontWeight: 'bold' }}>{this.state.isSalvando ? 'Salvando...' : 'Gravar Acesso'}</button>
                    </div>
                  </div>

                  <h3 style={{ fontSize: '15px', color: '#0f172a', marginBottom: '15px', display: 'flex', alignItems: 'center', gap: '8px' }}>📋 Usuários Atuais com Privilégios</h3>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    {this.state.listaAcessos.map((acesso, idx) => (
                      <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'white', padding: '15px 25px', borderRadius: '10px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                        <div>
                          <strong style={{ fontSize: '15px', color: '#0f172a' }}>{acesso.nome || "Usuário do Sistema"}</strong>
                          <p style={{ margin: '5px 0 0 0', fontSize: '13px', color: '#64748b' }}>{acesso.email}</p>
                        </div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                          <span style={{ background: acesso.nivel === 'TI' ? '#dcfce7' : '#e0f2fe', color: acesso.nivel === 'TI' ? '#166534' : '#0369a1', padding: '6px 12px', borderRadius: '20px', fontSize: '12px', fontWeight: 'bold' }}>Nível: {acesso.nivel}</span>
                          <button onClick={() => this.deletarAcesso(acesso.id)} style={{ background: 'none', border: 'none', color: '#ef4444', fontSize: '20px', cursor: 'pointer', padding: '5px' }} title="Remover acesso">🗑️</button>
                        </div>
                      </div>
                    ))}
                    {this.state.listaAcessos.length === 0 && <p style={{ color: '#94a3b8', textAlign: 'center', padding: '20px' }}>A lista de acessos está vazia.</p>}
                  </div>
                </div>
              )}

              {this.state.abaAtiva === 'cadastro' && isTI && (
                <div>
                   <div style={{ backgroundColor: '#f8fafc', padding: '20px', borderRadius: '12px', marginBottom: '25px', border: '1px solid #e2e8f0' }}>
                    <h3 style={{ marginTop: 0, fontSize: '16px', color: '#2E5C31' }}>👤 Dados do Responsável</h3>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                      
                      <div className={styles.inputGroup} style={{ position: 'relative' }}>
                        <label>Responsável (Busca no AD)</label>
                        <input 
                          type="text" value={this.state.novoNome} placeholder="Digite o nome..." autoComplete="off"
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
                    <div className={styles.inputGroup}><label>Tipo de Ativo</label><select value={this.state.novoTipo} onChange={(e) => this.setState({ novoTipo: e.target.value })}><option value="Notebook">Notebook</option><option value="Desktop">Desktop</option><option value="Celular / Smartphone">Celular / Smartphone</option><option value="Tablet">Tablet</option><option value="Monitor">Monitor</option><option value="Periférico">Periférico</option></select></div>
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
                  
                  {/* DASHBOARD DIVIDIDO EM DUAS LINHAS (Apenas para TI) */}
                  {isTI && this.state.ativosSalvos.length > 0 && (
                    <div style={{ marginBottom: '30px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                      
                      {/* LINHA 1: SAÚDE DO INVENTÁRIO */}
                      <div>
                        <h3 style={{ margin: '0 0 10px 0', fontSize: '14px', color: '#475569', textTransform: 'uppercase', letterSpacing: '0.5px' }}>📊 Saúde do Inventário</h3>
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '15px' }}>
                          <div style={{ background: '#f8fafc', padding: '15px 20px', borderRadius: '12px', borderLeft: '4px solid #3b82f6', boxShadow: '0 2px 4px rgba(0,0,0,0.03)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>📦 Total Geral</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '28px', fontWeight: 'bold', color: '#0f172a' }}>{totalEquipamentos}</p>
                          </div>
                          <div style={{ background: '#f8fafc', padding: '15px 20px', borderRadius: '12px', borderLeft: '4px solid #10b981', boxShadow: '0 2px 4px rgba(0,0,0,0.03)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>✅ Equipamentos em Uso</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '28px', fontWeight: 'bold', color: '#0f172a' }}>{qtdAtivos}</p>
                          </div>
                          <div style={{ background: '#f8fafc', padding: '15px 20px', borderRadius: '12px', borderLeft: '4px solid #f59e0b', boxShadow: '0 2px 4px rgba(0,0,0,0.03)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>🏢 No Estoque</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '28px', fontWeight: 'bold', color: '#0f172a' }}>{qtdEstoque}</p>
                          </div>
                          <div style={{ background: '#f8fafc', padding: '15px 20px', borderRadius: '12px', borderLeft: '4px solid #ef4444', boxShadow: '0 2px 4px rgba(0,0,0,0.03)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>🛠️ Em Manutenção</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '28px', fontWeight: 'bold', color: '#0f172a' }}>{qtdManutencao}</p>
                          </div>
                        </div>
                      </div>

                      {/* LINHA 2: PRINCIPAIS CATEGORIAS */}
                      <div>
                        <h3 style={{ margin: '0 0 10px 0', fontSize: '14px', color: '#475569', textTransform: 'uppercase', letterSpacing: '0.5px' }}>💻 Equipamentos</h3>
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: '15px' }}>
                          <div style={{ background: '#ffffff', padding: '15px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>💻 Notebooks</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#334155' }}>{qtdNotebooks}</p>
                          </div>
                          <div style={{ background: '#ffffff', padding: '15px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>🖥️ Desktops</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#334155' }}>{qtdDesktops}</p>
                          </div>
                          <div style={{ background: '#ffffff', padding: '15px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>📱 Celulares</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#334155' }}>{qtdCelulares}</p>
                          </div>
                          <div style={{ background: '#ffffff', padding: '15px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>📋 Tablets</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#334155' }}>{qtdTablets}</p>
                          </div>
                          <div style={{ background: '#ffffff', padding: '15px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>📺 Monitores</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#334155' }}>{qtdMonitores}</p>
                          </div>
                          <div style={{ background: '#ffffff', padding: '15px', borderRadius: '12px', border: '1px solid #e2e8f0', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                            <h4 style={{ margin: 0, color: '#64748b', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>🖱️ Periféricos</h4>
                            <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#334155' }}>{qtdPerifericos}</p>
                          </div>
                          
                          {/* NOVA CAIXINHA: OUTROS (Só aparece se tiver algum perdido) */}
                          {qtdOutros > 0 && (
                            <div style={{ background: '#fffbeb', padding: '15px', borderRadius: '12px', border: '1px solid #fde68a', boxShadow: '0 2px 4px rgba(0,0,0,0.02)' }}>
                              <h4 style={{ margin: 0, color: '#92400e', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>❓ Outros / Sem Tipo</h4>
                              <p style={{ margin: '5px 0 0 0', fontSize: '24px', fontWeight: 'bold', color: '#78350f' }}>{qtdOutros}</p>
                            </div>
                          )}
                        </div>
                      </div>

                    </div>
                  )}

                  {/* BLOCO TRANSFERENCIA LOTE (Apenas TI) */}
                  {isTI && this.state.itensSelecionados.length > 0 && (
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

                  {/* BARRA DE FILTROS AVANÇADOS (TI e Visualizadores) */}
                  {isAdminView && (
                    <div style={{ marginBottom: '25px', display: 'flex', gap: '10px', flexWrap: 'wrap', background: '#f1f5f9', padding: '15px', borderRadius: '12px', alignItems: 'center' }}>
                      <input 
                        type="text" placeholder="Pesquise (Nome, TI, Série...)" value={this.state.termoBusca} onChange={(e) => this.setState({ termoBusca: e.target.value })}
                        style={{ flex: '1 1 250px', padding: '12px 15px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '14px' }}
                      />
                      
                      <select value={this.state.filtroTipo} onChange={(e) => this.setState({ filtroTipo: e.target.value })} style={{ padding: '12px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '14px', background: 'white', flex: '1 1 150px' }}>
                        <option value="">Todos os Tipos</option>
                        <option value="Notebook">Notebooks</option>
                        <option value="Desktop">Desktops</option>
                        <option value="Celular / Smartphone">Celulares</option>
                        <option value="Tablet">Tablets</option>
                        <option value="Monitor">Monitores</option>
                        <option value="Periférico">Periféricos</option>
                      </select>

                      <select value={this.state.filtroDepartamento} onChange={(e) => this.setState({ filtroDepartamento: e.target.value })} style={{ padding: '12px', borderRadius: '8px', border: '1px solid #cbd5e1', fontSize: '14px', background: 'white', flex: '1 1 180px' }}>
                        <option value="">Todos os Departamentos</option>
                        {departamentosUnicos.map((dep, idx) => (
                          <option key={idx} value={dep as string}>{dep}</option>
                        ))}
                      </select>

                      <button onClick={this.carregarAtivosParaConsulta} style={{ background: '#2E5C31', color: 'white', border: 'none', padding: '12px 20px', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', gap: '8px' }}>🔄 Atualizar</button>
                      <button onClick={this.exportarParaExcel} style={{ background: '#10b981', color: 'white', border: 'none', padding: '12px 20px', borderRadius: '8px', cursor: 'pointer', fontWeight: 'bold', display: 'flex', gap: '8px' }}>📥 Exportar Excel</button>
                    </div>
                  )}

                  {this.state.carregandoConsulta ? (
                    <div style={{ textAlign: 'center', padding: '40px', color: '#64748b' }}>⏳ A carregar banco de dados...</div>
                  ) : (
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(320px, 1fr))', gap: '20px' }}>
                      {ativosParaExibir.length > 0 ? (
                        ativosParaExibir.map(ativo => (
                          <div key={ativo.id} style={{ background: this.state.itensSelecionados.includes(ativo.id) ? '#f0fdf4' : '#ffffff', border: '1px solid', borderColor: this.state.itensSelecionados.includes(ativo.id) ? '#86efac' : '#e2e8f0', borderRadius: '12px', padding: '20px', boxShadow: '0 4px 6px rgba(0,0,0,0.02)', transition: 'all 0.2s', cursor: 'default' }}>
                            
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '15px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                                {isTI && (
                                  <input 
                                    type="checkbox" 
                                    checked={this.state.itensSelecionados.includes(ativo.id)} 
                                    onChange={() => this.toggleSelecao(ativo.id)} 
                                    style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#2E5C31' }}
                                  />
                                )}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                  <span style={{ background: '#dcfce7', color: '#b45309', padding: '4px 10px', borderRadius: '6px', fontWeight: 'bold', fontSize: '13px', display: 'inline-block', width: 'fit-content' }}>
                                    TI: {ativo.patrimonio}
                                  </span>
                                  {ativo.patrimonioFin && ativo.patrimonioFin !== "-" && (
                                    <span style={{ background: '#f1f5f9', color: '#475569', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', display: 'inline-block', width: 'fit-content' }}>
                                      FIN: {ativo.patrimonioFin}
                                    </span>
                                  )}
                                </div>
                              </div>
                              <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                                {isAdminView && (
                                  <button onClick={() => this.abrirAuditoria(ativo)} title="Ver Histórico" style={{ background: 'transparent', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '6px 8px', cursor: 'pointer', fontSize: '12px', color: '#475569' }}>🕵️‍♂️</button>
                                )}
                                {isTI && (
                                  <button onClick={() => this.abrirModalEdicao(ativo)} title="Editar" style={{ background: 'transparent', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '6px 8px', cursor: 'pointer', fontSize: '12px', color: '#475569' }}>✏️</button>
                                )}
                                <span style={{ fontSize: '20px', background: '#f8fafc', padding: '6px', borderRadius: '50%' }}>
                                  {this.getIconeEquipamento(ativo.tipo)}
                                </span>
                              </div>
                            </div>

                            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '15px', opacity: ativo.responsavel.includes('Estoque') ? 0.6 : 1 }}>
                              {ativo.emailResponsavel ? (
                                <img 
                                  src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${ativo.emailResponsavel}`} 
                                  alt={ativo.responsavel}
                                  style={{ width: '42px', height: '42px', borderRadius: '50%', objectFit: 'cover', border: '2px solid #e2e8f0' }}
                                  onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling && ((e.currentTarget.nextElementSibling as HTMLElement).style.display = 'flex'); }}
                                />
                              ) : null}
                              <div style={{ width: '42px', height: '42px', borderRadius: '50%', background: '#f1f5f9', display: ativo.emailResponsavel ? 'none' : 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '20px' }}>
                                👤
                              </div>
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
                          {isAdminView ? 'Nenhum equipamento encontrado com estes filtros.' : 'Nenhum equipamento localizado sob sua responsabilidade.'}
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}

            </div>
          </main>
        </div>

        {/* === MODAL DE HISTÓRICO / AUDITORIA === */}
        {this.state.ativoAuditoria && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, backdropFilter: 'blur(3px)' }}>
            <div style={{ background: 'white', padding: '35px', borderRadius: '12px', width: '90%', maxWidth: '600px', maxHeight: '80vh', overflowY: 'auto', boxShadow: '0 20px 40px rgba(0,0,0,0.2)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '25px', borderBottom: '1px solid #e2e8f0', paddingBottom: '15px' }}>
                <h2 style={{ margin: 0, color: '#0f172a', fontSize: '20px' }}>🕵️‍♂️ Histórico: {this.state.ativoAuditoria.patrimonio}</h2>
                <button onClick={this.fecharModal} style={{ background: 'none', border: 'none', fontSize: '20px', cursor: 'pointer', color: '#64748b' }}>❌</button>
              </div>

              {this.state.carregandoAuditoria ? (
                <div style={{ textAlign: 'center', padding: '30px', color: '#64748b' }}>⏳ A carregar histórico do SharePoint...</div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
                  {this.state.dadosAuditoria.length > 0 ? (
                    this.state.dadosAuditoria.map((versao, idx) => (
                      <div key={idx} style={{ background: '#f8fafc', borderLeft: '4px solid #3b82f6', padding: '15px', borderRadius: '6px' }}>
                        <p style={{ margin: '0 0 5px 0', fontSize: '12px', color: '#64748b' }}><strong>Versão {versao.versao}</strong> • Modificado por {versao.modificadoPor} em {versao.data}</p>
                        <p style={{ margin: '0 0 5px 0', fontSize: '14px', color: '#0f172a' }}><strong>Responsável:</strong> {versao.responsavel}</p>
                        <p style={{ margin: '0', fontSize: '13px', color: '#475569' }}><strong>Obs:</strong> {versao.observacao}</p>
                      </div>
                    ))
                  ) : (
                    <div style={{ textAlign: 'center', padding: '20px', color: '#94a3b8' }}>Não foi possível carregar as versões. Certifique-se de que o "Histórico de Versões" está ativado nas Configurações da Lista no SharePoint.</div>
                  )}
                </div>
              )}
            </div>
          </div>
        )}

        {/* MODAL DE EDIÇÃO INDIVIDUAL */}
        {this.state.ativoSendoEditado && (
          <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 9999, backdropFilter: 'blur(3px)' }}>
            <div style={{ background: 'white', padding: '35px', borderRadius: '12px', width: '90%', maxWidth: '800px', maxHeight: '90vh', overflowY: 'auto', boxShadow: '0 20px 40px rgba(0,0,0,0.2)' }}>
              
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '25px', borderBottom: '1px solid #e2e8f0', paddingBottom: '15px' }}>
                <h2 style={{ margin: 0, color: '#2E5C31', fontSize: '20px' }}>✏️ Editando Transferência: {this.state.ativoSendoEditado.patrimonio}</h2>
                <button onClick={this.fecharModal} style={{ background: 'none', border: 'none', fontSize: '20px', cursor: 'pointer', color: '#64748b' }}>❌</button>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                
                {/* 1. RESPONSÁVEL (Ocupa as duas colunas - span 2) */}
                <div className={styles.inputGroup} style={{ position: 'relative', gridColumn: 'span 2' }}>
                  <label style={{ color: '#2E5C31', fontWeight: '900', fontSize: '14px' }}>👤 Novo Responsável / Status (Busca ou clique nas tags)</label>
                  <input type="text" value={this.state.editNome} placeholder="Ex: Nome da pessoa..." autoComplete="off"
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
                    style={{ border: '2px solid #A6CE39' }} /* Destaque sutil na borda */
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
                  {/* AS FLAGS COM MAIS ESPAÇO HORIZONTAL */}
                  <div style={{ display: 'flex', gap: '10px', marginTop: '8px', flexWrap: 'wrap' }}>
                    <button type="button" onClick={() => this.setState({ editNome: 'Estoque TI', editEmail: '', editDepartamento: 'TI', mostrarSugestoes: false })} style={{ background: '#f8fafc', border: '1px solid #cbd5e1', padding: '8px 16px', borderRadius: '16px', fontSize: '12px', cursor: 'pointer', color: '#475569', fontWeight: 'bold', transition: '0.2s' }} onMouseEnter={e => e.currentTarget.style.background = '#e2e8f0'} onMouseLeave={e => e.currentTarget.style.background = '#f8fafc'}>📦 Devolver ao Estoque</button>
                    <button type="button" onClick={() => this.setState({ editNome: 'Em Manutenção', editEmail: '', editDepartamento: 'TI', mostrarSugestoes: false })} style={{ background: '#fff1f2', border: '1px solid #fecdd3', padding: '8px 16px', borderRadius: '16px', fontSize: '12px', cursor: 'pointer', color: '#be123c', fontWeight: 'bold', transition: '0.2s' }} onMouseEnter={e => e.currentTarget.style.background = '#ffe4e6'} onMouseLeave={e => e.currentTarget.style.background = '#fff1f2'}>🛠️ P/ Manutenção</button>
                    <button type="button" onClick={() => this.setState({ editNome: 'Sucata / Descarte', editEmail: '', editDepartamento: 'TI', mostrarSugestoes: false })} style={{ background: '#f1f5f9', border: '1px solid #e2e8f0', padding: '8px 16px', borderRadius: '16px', fontSize: '12px', cursor: 'pointer', color: '#64748b', fontWeight: 'bold', transition: '0.2s' }} onMouseEnter={e => e.currentTarget.style.background = '#e2e8f0'} onMouseLeave={e => e.currentTarget.style.background = '#f1f5f9'}>🗑️ Sucata/Descarte</button>
                  </div>
                </div>
                
                {/* 2. DADOS ADMINISTRATIVOS */}
                <div className={styles.inputGroup}><label>Departamento</label><input value={this.state.editDepartamento} onChange={(e) => this.setState({ editDepartamento: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>Patrimônio Fin.</label><input value={this.state.editPatrimonioFin} onChange={(e) => this.setState({ editPatrimonioFin: e.target.value })} /></div>
                
                {/* 3. CLASSIFICAÇÃO */}
                <div className={styles.inputGroup}><label>Tipo de Ativo</label><select value={this.state.editTipo} onChange={(e) => this.setState({ editTipo: e.target.value })}><option value="Notebook">Notebook</option><option value="Desktop">Desktop</option><option value="Celular / Smartphone">Celular / Smartphone</option><option value="Tablet">Tablet</option><option value="Monitor">Monitor</option><option value="Periférico">Periférico</option></select></div>
                <div className={styles.inputGroup}><label>Fabricante</label><input value={this.state.editFabricante} onChange={(e) => this.setState({ editFabricante: e.target.value })} /></div>
                
                {/* 4. ESPECIFICAÇÕES */}
                <div className={styles.inputGroup}><label>Modelo Exato</label><input value={this.state.editModelo} onChange={(e) => this.setState({ editModelo: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>Especificações</label><input value={this.state.editEspecificacao} onChange={(e) => this.setState({ editEspecificacao: e.target.value })} /></div>
                
                {/* 5. IDENTIFICADORES (Com bloqueio inteligente do IMEI) */}
                <div className={styles.inputGroup}><label>Número de Série</label><input value={this.state.editSerie} onChange={(e) => this.setState({ editSerie: e.target.value })} /></div>
                <div className={styles.inputGroup}><label>IMEI (Celulares)</label><input value={this.state.editImei} onChange={(e) => this.setState({ editImei: e.target.value })} disabled={this.state.editTipo !== 'Celular / Smartphone'} placeholder={this.state.editTipo === 'Celular / Smartphone' ? "" : "Bloqueado"} /></div>
                
                {/* 6. OBSERVAÇÕES GERAIS */}
                <div className={styles.inputGroup} style={{ gridColumn: 'span 2' }}>
                  <label>Observações / Status atual</label>
                  <input value={this.state.editObservacao} onChange={(e) => this.setState({ editObservacao: e.target.value })} placeholder="Ex: Devolvido para o estoque. Aguarda formatação." style={{ borderColor: '#f59e0b', background: '#fffbeb' }} />
                </div>
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

        {/* ==============================================
            MODAL DE IFRAME PARA FORMULÁRIOS EXTERNOS
        ============================================== */}
        {this.state.isIframeModalOpen && (
          <div className={styles.modalOverlay} style={{ zIndex: 99999 }}>
            <div className={styles.modalContent} style={{ width: '900px', height: '85vh', maxWidth: '95%', display: 'flex', flexDirection: 'column' }}>
              <header className={styles.modalHeader} style={{ padding: '20px 30px' }}>
                <h3 style={{ margin: 0, color: '#171E0D', fontSize: '20px' }}>{this.state.iframeTitle}</h3>
                <button className={styles.closeBtn} onClick={() => this.setState({ isIframeModalOpen: false })}>✕</button>
              </header>
              <iframe 
                 src={this.state.iframeUrl} 
                 style={{ flex: 1, width: '100%', border: 'none', background: '#F8FAFC' }}
                 title={this.state.iframeTitle} 
              />
            </div>
          </div>
        )}

      </div>
    );
  }
}
