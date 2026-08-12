import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import type { IPoliticasGrunnerProps } from './IPoliticasGrunnerProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MenuChamados } from '../../../shared/components/MenuChamado/MenuChamados';
import FormularioSGQ from './FormularioSGQ';
import FormularioMapeamento from './FormularioMapeamento';
import FormularioProcedimento from './FormularioProcedimento';

// URLs de navegação
const logoGrunner = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo-grunner.png";
const homeUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Inicio.aspx?env=Embedded";
const historiaUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded";
const politicasUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded";
const atalhosUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/centraldeatalhos.aspx?env=Embedded";

export interface IPoliticaDocumento {
  Id: number;
  UniqueId?: string;
  FileLeafRef: string;
  FileRef: string;
  Area?: string;
  CodigoDocumento?: string;
  TipoDocumento?: string;
  TipoProcessoDocumento?: string;
  DocumentoControlado?: boolean;
  NumeroRevisao?: string;
  DataUltimaRevisao?: string;
  DataProximaRevisao?: string;
  StatusDocumento?: string;
  StatusCalculado?: string;
  ObservacaoRevisao?: string;
  PeriodicidadeRevisaoMeses?: number;
  UltimoAvisoRevisao?: string;
  DiasAvisoRevisao?: number;
  PermiteImpressaoControlada?: boolean;
  ExibirNaIntranet?: boolean;
  ResponsavelRevisao?: { Title?: string; EMail?: string };
  AprovadorQualidade?: { Title?: string; EMail?: string };
}

interface IPoliticasGrunnerState {
  areaAtiva: string;
  todosDocumentos: IPoliticaDocumento[];
  loading: boolean;
  termoBusca: string;
  filtroTipoProcesso: string;
  isMobileMenuOpen: boolean;
  isMenuTIOpen: boolean;
  isQualidadeUser: boolean;
  modoGestaoQualidade: boolean;
  documentoSelecionado?: IPoliticaDocumento | null;
  iframeDocumentoUrl: string | null;
  salvandoDocumento: boolean;
  filtroStatusAdmin: string;
  editFormData: Partial<IPoliticaDocumento>;
  isMenuProcedimentosOpen: boolean;
  activeModalTab: 'metadados' | 'historico';
  documentHistory: IDocumentVersion[];
  isLoadingHistory: boolean;
  isCreateModalOpen: boolean;
  selectedNewDocType: string;
}

export interface IDocumentVersion {
  VersionLabel: string;
  Modified: string;
  Editor: string;
  CheckInComment?: string;
}

export default class PoliticasGrunner extends React.Component<IPoliticasGrunnerProps, IPoliticasGrunnerState> {
  private areas = ['Institucional', 'TI', 'Sistemas', 'Marketing', 'RH', 'Compras'];
  private footerObserver?: MutationObserver;

  constructor(props: IPoliticasGrunnerProps) {
    super(props);

    this.state = {
      areaAtiva: 'Institucional',
      todosDocumentos: [],
      loading: true,
      isMenuProcedimentosOpen: true,
      termoBusca: '',
      isMobileMenuOpen: false,
      isMenuTIOpen: false,
      isQualidadeUser: false,
      modoGestaoQualidade: false,
      documentoSelecionado: null,
      salvandoDocumento: false,
      filtroStatusAdmin: 'Todos',
      editFormData: {},
      iframeDocumentoUrl: null,
      filtroTipoProcesso: '',
      activeModalTab: 'metadados',
      documentHistory: [],
      isLoadingHistory: false,
      isCreateModalOpen: false,
      selectedNewDocType: ''
    };
  }

  // Dicionário Estático: Mapeamento de Tipos de Documento para URLs de Formulários
  private formUrls: { [key: string]: string } = {
    'MAPEAMENTO DE PROCESSO': 'https://forms.office.com/r/exemplo-mapeamento',
    'PROCEDIMENTO': 'https://forms.office.com/r/exemplo-procedimento',
    'PROCEDIMENTO OPERACIONAL PADRÃO': 'https://forms.office.com/r/exemplo-pop',
    'INSTRUÇÃO DE TRABALHO': 'https://forms.office.com/r/exemplo-instrucao',
    'FORMULÁRIO': 'https://forms.office.com/r/exemplo-formulario',
    'MANUAL': 'https://forms.office.com/r/exemplo-manual',
    'POLÍTICA': 'https://forms.office.com/r/exemplo-politica'
  };

  // Ocultação padrão do SharePoint (Mantido)
  private shouldHideSharePointChrome = (): boolean => {
    const search = window.location.search.toLowerCase();
    const isEditMode = search.includes('mode=edit');
    const isEmbedded = search.includes('env=embedded') || search.includes('mode=embed');
    const forceAdmin = search.includes('admin=1');
    return isEmbedded && !isEditMode && !forceAdmin;
  }

  private collapseElement = (element: HTMLElement | null): void => {
    if (!element) return;
    element.style.setProperty('display', 'none', 'important');
    // ... restante das propriedades de ocultação originais ...
    element.style.setProperty('visibility', 'hidden', 'important');
    element.style.setProperty('height', '0', 'important');
    element.style.setProperty('margin', '0', 'important');
    element.style.setProperty('padding', '0', 'important');
    element.style.setProperty('overflow', 'hidden', 'important');
  }

  private hideSharePointFooter = (): void => {
    const selectors = ['[data-automation-id="page-bottom-actions"]', '#sp-page-footer', '.CommentsWrapper', '[data-sp-feature-tag="Comments"]'];
    document.querySelectorAll(selectors.join(',')).forEach((node) => {
      const el = node as HTMLElement;
      this.collapseElement(el);
      this.collapseElement(el.parentElement);
    });
  }

  private hideSharePointAppBar = (): void => {
    document.querySelectorAll('#sp-appBar, [data-automation-id="sp-appBar"], div[class^="appBar_"]').forEach((node) => {
      this.collapseElement(node as HTMLElement);
    });
  }

  private fixSharePointCanvasSpacing = (): void => {
    const applyFullBleed = (element: HTMLElement | null): void => {
      if (!element) return;
      element.style.setProperty('margin', '0', 'important');
      element.style.setProperty('padding', '0', 'important');
      element.style.setProperty('max-width', '100%', 'important');
      element.style.setProperty('width', '100%', 'important');
    };
    applyFullBleed(document.documentElement);
    applyFullBleed(document.body);
    document.querySelectorAll('.CanvasZone, .CanvasSection, #spPageCanvasContent').forEach(node => applyFullBleed(node as HTMLElement));
  }

  private bloquearAtalhos = (e: KeyboardEvent): void => {
    if ((e.ctrlKey || e.metaKey) && (e.key.toLowerCase() === 'p' || e.key.toLowerCase() === 's')) {
      e.preventDefault();
      e.stopPropagation();
      alert('Ação bloqueada: A impressão e o download não são permitidos para documentos controlados.');
    }
  }

  // Governança: Bloqueio do botão direito (menu de contexto)
  private bloquearBotaoDireito = (e: MouseEvent): void => {
    e.preventDefault();
  }

  public componentDidMount(): void {
    this.verificarAcessoQualidade();
    this.buscarTodosDocumentos();

    window.addEventListener('keydown', this.bloquearAtalhos, true);
    window.addEventListener('contextmenu', this.bloquearBotaoDireito, true);

    if (this.shouldHideSharePointChrome()) {
      const applyFixes = (): void => {
        this.hideSharePointFooter();
        this.hideSharePointAppBar();
        this.fixSharePointCanvasSpacing();
      };
      applyFixes();
      window.setTimeout(applyFixes, 500);
      window.setTimeout(applyFixes, 1500);

      this.footerObserver = new MutationObserver(() => applyFixes());
      if (document.body) this.footerObserver.observe(document.body, { childList: true, subtree: true });
    }
  }

  public componentWillUnmount(): void {
    if (this.footerObserver) this.footerObserver.disconnect();

    window.removeEventListener('keydown', this.bloquearAtalhos, true);
    window.removeEventListener('contextmenu', this.bloquearBotaoDireito, true);
  }

  // Governança: Verificação de Permissão
  private verificarAcessoQualidade = async (): Promise<void> => {
    try {
      const emailUsuario = this.props.context.pageContext.user.email;
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/sitegroups/getbyname('Qualidade - Gestão de Documentos')/users?$select=Email`;
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      const isQualidadeUser = data.value?.some((user: any) => user.Email?.toLowerCase() === emailUsuario.toLowerCase());
      this.setState({ isQualidadeUser });
    } catch (error) {
      console.warn('Usuário não pertence ao grupo de Qualidade ou grupo não existe.');
    }
  }

  // Governança: Busca com Metadados Completos
  private buscarTodosDocumentos = async (): Promise<void> => {
    this.setState({ loading: true });
    try {
      const select = 'Id,UniqueId,FileLeafRef,FileRef,Area,CodigoDocumento,TipoDocumento,NumeroRevisao,DataUltimaRevisao,DataProximaRevisao,StatusDocumento,ObservacaoRevisao,PeriodicidadeRevisaoMeses,UltimoAvisoRevisao,DiasAvisoRevisao,PermiteImpressaoControlada,ExibirNaIntranet,ResponsavelRevisao/Title,ResponsavelRevisao/EMail,AprovadorQualidade/Title,AprovadorQualidade/EMail,TipoProcessoDocumento,DocumentoControlado';
      const expand = 'ResponsavelRevisao,AprovadorQualidade';
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PoliticasGrunner')/items?$select=${select}&$expand=${expand}&$top=5000`;

      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data && data.value) {
        const documentosProcessados = data.value.map((doc: any) => {
          return {
            ...doc,
            StatusCalculado: this.calcularStatusDocumento(doc)
          };
        });
        this.setState({ todosDocumentos: documentosProcessados, loading: false });
      } else {
        this.setState({ loading: false });
      }
    } catch (error) {
      console.error("Erro ao buscar documentos:", error);
      this.setState({ loading: false });
    }
  }

  // Governança: Busca de Histórico de Versões do Documento
  private buscarHistoricoDocumento = async (itemId: number): Promise<void> => {
    this.setState({ isLoadingHistory: true, documentHistory: [] });

    try {
      // O SharePoint guarda a data de modificação da versão no campo 'Created' do endpoint /versions
      const select = 'VersionLabel,Created,CheckInComment,Editor/Title';
      const expand = 'Editor';
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PoliticasGrunner')/items(${itemId})/versions?$select=${select}&$expand=${expand}`;

      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`Erro na requisição: ${response.statusText}`);
      }

      const data = await response.json();

      if (data && data.value) {
        // Mapeia o retorno sujo da API para a nossa interface limpa
        const historicoFormatado: IDocumentVersion[] = data.value.map((versao: any) => ({
          VersionLabel: versao.VersionLabel,
          Modified: versao.Created,
          Editor: versao.Editor ? versao.Editor.Title : 'Sistema',
          CheckInComment: versao.CheckInComment || ''
        }));

        this.setState({
          documentHistory: historicoFormatado,
          isLoadingHistory: false
        });
      } else {
        this.setState({ isLoadingHistory: false });
      }
    } catch (error) {
      console.error("Erro ao buscar o histórico de versões:", error);
      this.setState({ isLoadingHistory: false });
    }
  }

  // Governança: Lógica de Status
  private calcularStatusDocumento = (doc: any): string => {
    if (doc.StatusDocumento === 'Obsoleto') return 'Obsoleto';
    if (doc.StatusDocumento === 'Arquivado') return 'Arquivado';
    if (doc.StatusDocumento === 'Em revisão') return 'Em revisão';

    if (doc.DataProximaRevisao) {
      const dataVencimento = new Date(doc.DataProximaRevisao);
      const hoje = new Date();
      const diasRestantes = Math.ceil((dataVencimento.getTime() - hoje.getTime()) / (1000 * 3600 * 24));

      if (diasRestantes < 0) return 'Vencido';
      if (diasRestantes <= 30) return 'Vence em breve';
    }
    return 'Vigente';
  }

  // Governança: Atualizar Documento (Admin)
  private getUserIdByEmail = async (email: string): Promise<number | null> => {
    if (!email) return null;
    try {
      const response = await this.props.context.spHttpClient.post(
        `${this.props.context.pageContext.web.absoluteUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-type': 'application/json;odata=verbose' }, body: JSON.stringify({ logonName: email }) }
      );
      const user = await response.json();
      return user.Id || user.d?.Id;
    } catch (e) {
      console.error('Erro ao resolver usuário', e);
      return null;
    }
  }

  private salvarEdicaoDocumento = async (): Promise<void> => {
    this.setState({ salvandoDocumento: true });
    const { documentoSelecionado, editFormData } = this.state;
    if (!documentoSelecionado) return;

    try {
      const payload: any = {
        CodigoDocumento: editFormData.CodigoDocumento || null,
        TipoProcessoDocumento: editFormData.TipoProcessoDocumento || null,
        DocumentoControlado: editFormData.DocumentoControlado,
        TipoDocumento: editFormData.TipoDocumento || null,
        NumeroRevisao: editFormData.NumeroRevisao || null,
        DataUltimaRevisao: editFormData.DataUltimaRevisao || null,
        DataProximaRevisao: editFormData.DataProximaRevisao || null,
        StatusDocumento: editFormData.StatusDocumento || null,
        ObservacaoRevisao: editFormData.ObservacaoRevisao || null,
        PeriodicidadeRevisaoMeses: editFormData.PeriodicidadeRevisaoMeses || null,
        PermiteImpressaoControlada: editFormData.PermiteImpressaoControlada,
        ExibirNaIntranet: editFormData.ExibirNaIntranet,
        Area: editFormData.Area || null
      };

      // Resolução de usuários (Pessoa ou Grupo)
      if (editFormData.ResponsavelRevisao?.EMail) {
        payload.ResponsavelRevisaoId = await this.getUserIdByEmail(editFormData.ResponsavelRevisao.EMail);
      } else if (editFormData.ResponsavelRevisao?.EMail === '') {
        // Enviar -1 força a API do SharePoint a limpar um campo de Lookup/Pessoa
        payload.ResponsavelRevisaoId = -1;
      }

      if (editFormData.AprovadorQualidade?.EMail) {
        payload.AprovadorQualidadeId = await this.getUserIdByEmail(editFormData.AprovadorQualidade.EMail);
      } else if (editFormData.AprovadorQualidade?.EMail === '') {
        payload.AprovadorQualidadeId = -1;
      }

      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PoliticasGrunner')/items(${documentoSelecionado.Id})`;
      await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      });

      this.setState({
        documentoSelecionado: null,
        salvandoDocumento: false,
        activeModalTab: 'metadados',
        documentHistory: []
      });
      this.buscarTodosDocumentos();
    } catch (error) {
      console.error("Erro ao salvar:", error);
      this.setState({ salvandoDocumento: false });
    }
  }

  private formatDate = (dateStr?: string): string => {
    if (!dateStr) return 'Não cadastrado';

    const apenasData = dateStr.split('T')[0];
    const [ano, mes, dia] = apenasData.split('-');

    return `${dia}/${mes}/${ano}`;
  }

  private getStatusClass = (status?: string): string => {
    switch (status) {
      case 'Vigente': return styles.statusVigente;
      case 'Vence em breve': return styles.statusAtencao;
      case 'Vencido': return styles.statusVencido;
      case 'Em revisão': return styles.statusRevisao;
      case 'Arquivado': return styles.statusArquivado;
      default: return styles.statusBadge;
    }
  }

  private exportarParaCSV = (documentos: IPoliticaDocumento[]): void => {
    const cabecalho = ['Código', 'Nome', 'Área', 'Tipo', 'Controlado', 'Revisão', 'Vencimento', 'Status'];
    const linhas = documentos.map(doc => [
      doc.CodigoDocumento || '-',
      doc.FileLeafRef || '-',
      doc.Area || '-',
      doc.TipoProcessoDocumento || '-',
      doc.DocumentoControlado ? 'Sim' : 'Não',
      doc.NumeroRevisao || '-',
      this.formatDate(doc.DataProximaRevisao),
      doc.StatusCalculado || '-'
    ]);

    const conteudoCSV = [cabecalho, ...linhas].map(e => e.join(';')).join('\n');
    const blob = new Blob(["\ufeff", conteudoCSV], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'Relatorio_Documentos.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  public render(): React.ReactElement<IPoliticasGrunnerProps> {
    const { areaAtiva, todosDocumentos, termoBusca, loading, isQualidadeUser, modoGestaoQualidade } = this.state;

    // Métricas
    const total = todosDocumentos.length;
    const vigentes = todosDocumentos.filter(d => d.StatusCalculado === 'Vigente').length;
    const atencao = todosDocumentos.filter(d => d.StatusCalculado === 'Vence em breve').length;
    const vencidos = todosDocumentos.filter(d => d.StatusCalculado === 'Vencido').length;
    const revisao = todosDocumentos.filter(d => d.StatusCalculado === 'Em revisão').length;

    // Filtros de exibição
    let documentosExibidos = todosDocumentos;

    if (!modoGestaoQualidade) {
      // Regra Pública: Não exibir arquivados, VENCIDOS, Em revisão, nem OBSOLETOS, e exibir apenas marcados para Intranet
      documentosExibidos = documentosExibidos.filter(d =>
        d.StatusCalculado !== 'Arquivado' &&
        d.StatusCalculado !== 'Vencido' &&
        d.StatusCalculado !== 'Em revisão' &&
        d.StatusCalculado !== 'Obsoleto' && // <-- NOVA REGRA AQUI
        (d.ExibirNaIntranet !== false)
      );

      if (termoBusca.trim().length > 0) {
        documentosExibidos = documentosExibidos.filter(doc =>
          doc.FileLeafRef?.toLowerCase().includes(termoBusca.toLowerCase()) ||
          doc.CodigoDocumento?.toLowerCase().includes(termoBusca.toLowerCase())
        );
      } else {
        documentosExibidos = documentosExibidos.filter(doc => doc.Area === areaAtiva);
      }
    } else {
      // Gestão de Qualidade: Filtros Administrativos
      if (this.state.filtroStatusAdmin !== 'Todos') {
        documentosExibidos = documentosExibidos.filter(d => d.StatusCalculado === this.state.filtroStatusAdmin);
      }
      if (termoBusca.trim().length > 0) {
        documentosExibidos = documentosExibidos.filter(doc =>
          doc.FileLeafRef?.toLowerCase().includes(termoBusca.toLowerCase()) ||
          doc.CodigoDocumento?.toLowerCase().includes(termoBusca.toLowerCase())
        );
      }
    }

    // NOVO: Aplica o filtro de Tipo de Processo para ambas as visões (Pública e Qualidade)
    if (this.state.filtroTipoProcesso) {
      documentosExibidos = documentosExibidos.filter(doc => doc.TipoProcessoDocumento === this.state.filtroTipoProcesso);
    }

    return (
      <div className={styles.container}>
        {this.shouldHideSharePointChrome() && (
          <style dangerouslySetInnerHTML={{ __html: `... ocultações do sharepoint originais mantidas no seu código ...` }} />
        )}

        <div className={styles.mobileHeaderBar}>
          <button className={styles.hamburgerBtn} onClick={() => this.setState({ isMobileMenuOpen: true })}>☰ Menu Grunner</button>
        </div>

        {this.state.isMobileMenuOpen && <div className={styles.mobileOverlayBackdrop} onClick={() => this.setState({ isMobileMenuOpen: false })} />}

        <aside className={`${styles.sidebar} ${this.state.isMobileMenuOpen ? styles.open : ''}`}>
          <button
            className={styles.closeMenuBtn}
            onClick={() => this.setState({ isMobileMenuOpen: false })}
          >
            ✕
          </button>

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
              className={`${styles.menuToggle} ${this.state.isMenuTIOpen ? styles.active : ''}`}
              onClick={(e) => { e.preventDefault(); this.setState({ isMenuTIOpen: !this.state.isMenuTIOpen }); }}
            >
              <span style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>💻 Tecnologia (TI)</span>
              <span style={{ fontSize: '10px', opacity: 0.8 }}>{this.state.isMenuTIOpen ? '▲' : '▼'}</span>
            </a>

            {/* SUB-ITENS DE TI */}
            {this.state.isMenuTIOpen && (
              <div className={styles.navSubGroup}>
                <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/GerenciamentoDeAtivos.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">🖥️ Gestão de Ativos</a>
                <a href="https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5" target="_blank" rel="noopener noreferrer">➕ Abrir Novo Chamado</a>
                <a href="#" onClick={(e) => { e.preventDefault(); window.dispatchEvent(new CustomEvent('abrirMeusChamadosGrunner', { detail: 'TI' })); }}>🎫 Meus Chamados</a>
              </div>
            )}

            {/* RESTANTE DOS DEPARTAMENTOS */}
            <a href="https://grunnerteccombr.sharepoint.com/sites/Marketing/_layouts/15/listforms.aspx?cid=MTQ1MjlmMzEtNjk2Ni00MTI2LWJhNzItMzE1MTc0NDU2YTE4&nav=MGIwZDdiNzMtODQwNi00MDhiLTk5ZDEtNGE5NWNlYzljNDg3" target="_blank" rel="noopener noreferrer" data-interception="off">📢 Marketing</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/GPS/_layouts/15/listforms.aspx?cid=ZWFlMDE1MWUtOTFlMS00MmJiLWFiNzEtOWM0NGVkZTVkMTdh&nav=ZGJmNmMxZGMtNjU5Zi00ZTUxLThjMTctZmFhODY5YTQ3NjBi" target="_blank" rel="noopener noreferrer" data-interception="off">🚗 Frotas</a>
            <a href="https://grunnerteccombr.sharepoint.com/:l:/s/Facilities/JADJeN1a-IAVRIrzsns79wBEAS_s9zB21POwKXunqjUuK5Y?nav=MDk0ODE1N2QtZWE0Ny00ZDhjLWFhYjItMGVlNmIwMWIzNTY4" target="_blank" rel="noopener noreferrer">🛠️ Facilities</a>
          </div>

          <div className={styles.navGroup}>
            <h3>Institucional</h3>
            <a href={historiaUrl} target="_blank" rel="noopener noreferrer">🏛️ Nossa História</a>
            <a href="https://grunnertec.com.br/assets/PDFs/codigoconduta.pdf" target="_blank" rel="noopener noreferrer">⚖️ Código de Conduta</a>
            <a href="https://grunner.canaldeouvidoria.com.br/" target="_blank" rel="noopener noreferrer">🗣️ Canal de Ética</a>

            {/* =========================================================
                LÓGICA DO MENU PROCEDIMENTOS (QUALIDADE VS NORMAL)
            ========================================================= */}
            {this.state.isQualidadeUser ? (
              <>
                {/* MENU ACORDEÃO (Apenas para quem tem acesso à Qualidade) */}
                <a
                  className={`${styles.menuToggle} ${this.state.isMenuProcedimentosOpen ? styles.active : ''}`}
                  onClick={(e) => { e.preventDefault(); this.setState({ isMenuProcedimentosOpen: !this.state.isMenuProcedimentosOpen }); }}
                >
                  <span style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>📖 Procedimentos</span>
                  <span style={{ fontSize: '10px', opacity: 0.8 }}>{this.state.isMenuProcedimentosOpen ? '▲' : '▼'}</span>
                </a>

                {/* SUB-ITENS DE PROCEDIMENTOS */}
                {this.state.isMenuProcedimentosOpen && (
                  <div className={styles.navSubGroup}>
                    <a href={politicasUrl} className={!this.state.modoGestaoQualidade ? styles.active : ''}>
                      📖 Todos os Documentos
                    </a>
                    <a
                      href="#"
                      className={this.state.modoGestaoQualidade ? styles.active : ''}
                      onClick={(e) => { e.preventDefault(); this.setState({ modoGestaoQualidade: !this.state.modoGestaoQualidade }); }}
                    >
                      ⚙️ Gestão da Qualidade
                    </a>
                  </div>
                )}
              </>
            ) : (
              /* LINK DIRETO (Para colaborador normal - Sem setinha e sem submenu) */
              <a href={politicasUrl} className={!this.state.modoGestaoQualidade ? styles.active : ''}>
                📖 Procedimentos
              </a>
            )}
          </div>

        </aside>
        <div className={styles.contentArea}>
          <header className={styles.pageHeader}>
            <MenuChamados departamento="TI" emailUsuario={this.props.context.pageContext.user.email} />
            <div className={styles.headerText}>
              <h1>{modoGestaoQualidade ? '⚙️ Gestão de Documentos da Qualidade' : '📖 Políticas e Diretrizes Grunner'}</h1>
              <p>{modoGestaoQualidade ? 'Controle de revisões, vencimentos e auditoria de procedimentos.' : 'Acesse os documentos normativos, manuais e procedimentos de cada área da empresa.'}</p>
            </div>
          </header>

          {/* PAINEL DE MÉTRICAS */}
          <div className={styles.metricsPanel}>
            <div
              className={`${styles.metricCard} ${modoGestaoQualidade ? styles.clickableCard : ''} ${modoGestaoQualidade && this.state.filtroStatusAdmin === 'Todos' ? styles.metricActive : ''}`}
              onClick={() => modoGestaoQualidade && this.setState({ filtroStatusAdmin: 'Todos' })}
            >
              <div className={styles.metricLabel}>Total</div>
              <div className={styles.metricValue}>{total}</div>
            </div>

            <div
              className={`${styles.metricCard} ${styles.metricVigente} ${modoGestaoQualidade ? styles.clickableCard : ''} ${modoGestaoQualidade && this.state.filtroStatusAdmin === 'Vigente' ? styles.metricActive : ''}`}
              onClick={() => modoGestaoQualidade && this.setState({ filtroStatusAdmin: 'Vigente' })}
            >
              <div className={styles.metricLabel}>Vigentes</div>
              <div className={styles.metricValue}>{vigentes}</div>
            </div>

            <div
              className={`${styles.metricCard} ${styles.metricAtencao} ${modoGestaoQualidade ? styles.clickableCard : ''} ${modoGestaoQualidade && this.state.filtroStatusAdmin === 'Vence em breve' ? styles.metricActive : ''}`}
              onClick={() => modoGestaoQualidade && this.setState({ filtroStatusAdmin: 'Vence em breve' })}
            >
              <div className={styles.metricLabel}>Vence em breve</div>
              <div className={styles.metricValue}>{atencao}</div>
            </div>

            <div
              className={`${styles.metricCard} ${styles.metricVencido} ${modoGestaoQualidade ? styles.clickableCard : ''} ${modoGestaoQualidade && this.state.filtroStatusAdmin === 'Vencido' ? styles.metricActive : ''}`}
              onClick={() => modoGestaoQualidade && this.setState({ filtroStatusAdmin: 'Vencido' })}
            >
              <div className={styles.metricLabel}>Vencidos</div>
              <div className={styles.metricValue}>{vencidos}</div>
            </div>

            <div
              className={`${styles.metricCard} ${styles.metricRevisao} ${modoGestaoQualidade ? styles.clickableCard : ''} ${modoGestaoQualidade && this.state.filtroStatusAdmin === 'Em revisão' ? styles.metricActive : ''}`}
              onClick={() => modoGestaoQualidade && this.setState({ filtroStatusAdmin: 'Em revisão' })}
            >
              <div className={styles.metricLabel}>Em Revisão</div>
              <div className={styles.metricValue}>{revisao}</div>
            </div>
          </div>

          <div className={styles.searchContainer}>
            <input type="text" placeholder="🔍 Buscar por nome ou código..." value={termoBusca} onChange={(e) => this.setState({ termoBusca: e.target.value })} className={styles.searchInput} />
          </div>

          <div className={styles.filtersRow}>
            <select
              value={this.state.filtroTipoProcesso}
              onChange={(e) => this.setState({ filtroTipoProcesso: e.target.value })}
            >
              <option value="">Todos os Tipos de Documento</option>
              <option value="MAPEAMENTO DE PROCESSO">Mapeamento de Processo</option>
              <option value="PROCEDIMENTO">Procedimento</option>
              <option value="PROCEDIMENTO OPERACIONAL PADRÃO">Procedimento Operacional Padrão (POP)</option>
              <option value="INSTRUÇÃO DE TRABALHO">Instrução de Trabalho</option>
              <option value="FORMULÁRIO">Formulário</option>
              <option value="MANUAL">Manual</option>
              <option value="POLÍTICA">Política</option>
            </select>

            {modoGestaoQualidade && (
              <button className={styles.exportButton} onClick={() => this.exportarParaCSV(documentosExibidos)}>
                📊 Exportar Excel/CSV
              </button>
            )}

            {/*BOTÃO CRIAR DOCUMENTO (Apenas Visão Pública) */}
            {!modoGestaoQualidade && (
              <button
                className={styles.createButton}
                onClick={() => this.setState({ isCreateModalOpen: true })}
              >
                ➕ Solicitar Novo Documento
              </button>
            )}
          </div>

          {!modoGestaoQualidade && (
            <nav className={`${styles.tabsContainer} ${termoBusca.length > 0 ? styles.tabsDisabled : ''}`}>
              {this.areas.map((area) => (
                <button key={area} className={areaAtiva === area && termoBusca.length === 0 ? styles.tabActive : styles.tab} onClick={() => this.setState({ areaAtiva: area, termoBusca: '' })}>
                  {area}
                </button>
              ))}
            </nav>
          )}

          {modoGestaoQualidade && (
            <div className={styles.adminHeaderControls}>
              <div className={styles.filterStatusContainer}>
                {['Todos', 'Vigente', 'Vence em breve', 'Vencido', 'Em revisão', 'Arquivado'].map(status => (
                  <button key={status} className={this.state.filtroStatusAdmin === status ? styles.filterStatusActive : styles.filterStatusButton} onClick={() => this.setState({ filtroStatusAdmin: status })}>
                    {status}
                  </button>
                ))}
              </div>

              <a
                href={`${this.props.context.pageContext.web.absoluteUrl}/PoliticasGrunner/Forms/AllItems.aspx`}
                target="_blank"
                rel="noopener noreferrer"
                className={styles.uploadButton}
              >
                ➕ Carregar Novos Documentos
              </a>
            </div>
          )}

          <main className={styles.documentsArea}>
            {loading ? (
              <div className={styles.loadingState}><div className={styles.spinner}></div><p>Carregando documentos...</p></div>
            ) : !modoGestaoQualidade ? (

              // VISÃO PÚBLICA (Cards)
              <div className={styles.documentGrid}>
                {documentosExibidos.map((doc, index) => {
                  const extensao = doc.FileLeafRef ? doc.FileLeafRef.split('.').pop()?.toLowerCase() : '';
                  const isPdf = extensao === 'pdf';
                  return (
                    <div key={index} className={styles.documentCard}>
                      {/* CABEÇALHO DO CARD (Ícone e Status) */}
                      <div className={styles.cardHeader}>
                        <div className={styles.headerTop}>
                          <div className={isPdf ? styles.iconPdf : styles.iconDoc}>{isPdf ? 'PDF' : 'DOC'}</div>
                          <span className={`${styles.statusBadge} ${this.getStatusClass(doc.StatusCalculado)}`}>
                            {doc.StatusCalculado}
                          </span>
                        </div>
                        <h3 className={styles.docTitle} title={doc.FileLeafRef.replace(`.${extensao}`, '')}>
                          {doc.FileLeafRef.replace(`.${extensao}`, '')}
                        </h3>
                      </div>

                      {/* CORPO DO CARD (Área, Tipo, Controle e Código) */}
                      <div className={styles.cardBody}>
                        {/* NOVO: Badge de Documento Controlado */}
                        <span className={`${styles.badgeControlado} ${doc.DocumentoControlado ? styles.isControlado : styles.isNaoControlado}`}>
                          {doc.DocumentoControlado ? '🛡️ Controlado' : '📄 Não Controlado'}
                        </span>

                        <span className={styles.areaBadge}>
                          {doc.Area || 'Geral'} {doc.TipoProcessoDocumento ? `• ${doc.TipoProcessoDocumento}` : (doc.TipoDocumento ? `• ${doc.TipoDocumento}` : '')}
                        </span>
                        <span className={styles.docCode}>
                          {doc.CodigoDocumento ? `Código: ${doc.CodigoDocumento}` : <span className={styles.emptyCode}>Sem código</span>}
                        </span>
                      </div>

                      {/* RODAPÉ DO CARD (Revisão, Vencimento e Botão Iframe) */}
                      <div className={styles.cardFooter}>
                        <div className={styles.revisionInfo}>
                          <span className={styles.revText}>Rev. {doc.NumeroRevisao || '00'}</span>
                          <span className={styles.venceText}>Vence: {this.formatDate(doc.DataProximaRevisao)}</span>
                        </div>
                        {/* NOVO: Usa o visualizador universal e seguro da Microsoft */}
                        <a
                          onClick={(e) => {
                            e.preventDefault();

                            const siteUrl = this.props.context.pageContext.web.absoluteUrl;

                            // Usa o visualizador universal da Microsoft e força a ocultação de barras do Office
                            const urlIframe = `${siteUrl}/_layouts/15/embed.aspx?UniqueId=${doc.UniqueId}&wdHideRibbon=True&wdHideHeaders=True`;

                            this.setState({ iframeDocumentoUrl: urlIframe });
                          }}
                          className={styles.openButton}
                        >
                          Abrir documento
                        </a>
                      </div>
                    </div>
                  );
                })}
              </div>

            ) : (
              // VISÃO ADMINISTRATIVA (Tabela)
              <div className={styles.adminTableWrapper}>
                <table className={styles.adminTable}>
                  <thead>
                    <tr>
                      <th>Código</th><th>Nome</th><th>Área</th><th>Rev</th><th>Vencimento</th><th>Status</th><th>Responsável</th><th>Ações</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documentosExibidos.map((doc, idx) => (
                      <tr key={idx}>
                        <td>{doc.CodigoDocumento || '-'}</td>
                        <td>{doc.FileLeafRef}</td>
                        <td>{doc.Area}</td>
                        <td>{doc.NumeroRevisao || '-'}</td>
                        <td>{this.formatDate(doc.DataProximaRevisao)}</td>
                        <td><span className={`${styles.statusBadge} ${this.getStatusClass(doc.StatusCalculado)}`}>{doc.StatusCalculado}</span></td>
                        <td>{doc.ResponsavelRevisao?.Title || '-'}</td>
                        <td className={styles.adminActions}>
                          <button onClick={() => this.setState({ documentoSelecionado: doc, editFormData: { ...doc, ResponsavelRevisao: { EMail: doc.ResponsavelRevisao?.EMail }, AprovadorQualidade: { EMail: doc.AprovadorQualidade?.EMail } } })} className={styles.editButton}>✏️ Editar</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </main>
        </div>

        {/* MODAL DE EDIÇÃO ADMINISTRATIVA */}
        {this.state.documentoSelecionado && (
          <div className={styles.editModalBackdrop}>
            <div className={styles.editModal}>

              <div className={styles.editModalHeader}>
                <h2>{this.state.documentoSelecionado.FileLeafRef}</h2>
                <button onClick={() => this.setState({ documentoSelecionado: null, activeModalTab: 'metadados', documentHistory: [] })} className={styles.closeModal}>✕</button>
              </div>

              {/* NOVAS ABAS DE NAVEGAÇÃO INTERNA */}
              <div className={styles.modalTabs}>
                <button
                  className={`${styles.modalTab} ${this.state.activeModalTab === 'metadados' ? styles.modalTabActive : ''}`}
                  onClick={() => this.setState({ activeModalTab: 'metadados' })}
                >
                  📝 Editar Metadados
                </button>
                <button
                  className={`${styles.modalTab} ${this.state.activeModalTab === 'historico' ? styles.modalTabActive : ''}`}
                  onClick={() => {
                    this.setState({ activeModalTab: 'historico' });

                    // O "this.state.documentoSelecionado &&" acalma o TypeScript
                    if (this.state.documentoSelecionado && (!this.state.documentHistory || this.state.documentHistory.length === 0)) {
                      this.buscarHistoricoDocumento(this.state.documentoSelecionado.Id);
                    }
                  }}
                >
                  🕒 Histórico de Revisões
                </button>
              </div>

              <div className={styles.editModalBody}>

                {/* CONTEÚDO DA ABA 1: METADADOS */}
                {this.state.activeModalTab === 'metadados' && (
                  <div className={styles.formGrid}>

                    <div className={styles.formGroup}>
                      <label>Código do Documento</label>
                      <input type="text" value={this.state.editFormData.CodigoDocumento || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, CodigoDocumento: e.target.value } })} />
                    </div>

                    <div className={styles.formGroup}>
                      <label>Número da Revisão</label>
                      <input type="text" value={this.state.editFormData.NumeroRevisao || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, NumeroRevisao: e.target.value } })} />
                    </div>

                    <div className={styles.formGroup}>
                      <label>Data Última Revisão</label>
                      <input type="date" value={this.state.editFormData.DataUltimaRevisao ? this.state.editFormData.DataUltimaRevisao.split('T')[0] : ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, DataUltimaRevisao: e.target.value } })} />
                    </div>

                    <div className={styles.formGroup}>
                      <label>Data Próxima Revisão</label>
                      <input type="date" value={this.state.editFormData.DataProximaRevisao ? this.state.editFormData.DataProximaRevisao.split('T')[0] : ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, DataProximaRevisao: e.target.value } })} />
                    </div>

                    <div className={styles.formGroup}>
                      <label>Tipo de Processo/Documento</label>
                      <select value={this.state.editFormData.TipoProcessoDocumento || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, TipoProcessoDocumento: e.target.value } })}>
                        <option value="">Selecione...</option>
                        <option value="MAPEAMENTO DE PROCESSO">Mapeamento de Processo</option>
                        <option value="PROCEDIMENTO">Procedimento</option>
                        <option value="PROCEDIMENTO OPERACIONAL PADRÃO">Procedimento Operacional Padrão (POP)</option>
                        <option value="INSTRUÇÃO DE TRABALHO">Instrução de Trabalho</option>
                        <option value="FORMULÁRIO">Formulário</option>
                        <option value="MANUAL">Manual</option>
                        <option value="POLÍTICA">Política</option>
                      </select>
                    </div>

                    <div className={styles.formGroup}>
                      <label>Documento Controlado?</label>
                      <select value={this.state.editFormData.DocumentoControlado ? 'sim' : 'nao'} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, DocumentoControlado: e.target.value === 'sim' } })}>
                        <option value="nao">Não - Não Controlado</option>
                        <option value="sim">Sim - Controlado</option>
                      </select>
                    </div>

                    <div className={styles.formGroup}>
                      <label>Status (Sobrescreve regra de data se Arquivado/Em revisão)</label>
                      <select value={this.state.editFormData.StatusDocumento || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, StatusDocumento: e.target.value } })}>
                        <option value="">Automático (Pela Data)</option>
                        <option value="Em revisão">Em revisão</option>
                        <option value="Arquivado">Arquivado</option>
                        <option value="Obsoleto">Obsoleto</option>
                      </select>
                    </div>

                    <div className={styles.formGroup}>
                      <label>E-mail do Responsável</label>
                      <input type="email" placeholder="email@grunner.com.br" value={this.state.editFormData.ResponsavelRevisao?.EMail || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, ResponsavelRevisao: { EMail: e.target.value } } })} />
                    </div>

                  </div>
                )}

                {/* CONTEÚDO DA ABA 2: HISTÓRICO */}
                {this.state.activeModalTab === 'historico' && (
                  <div className={styles.timelineContainer}>
                    {this.state.isLoadingHistory ? (
                      <div className={styles.loadingState}><div className={styles.spinner}></div><p>Buscando histórico...</p></div>
                    ) : this.state.documentHistory && this.state.documentHistory.length > 0 ? (
                      this.state.documentHistory.map((versao, idx) => (
                        <div key={idx} className={styles.timelineItem}>
                          <div className={styles.timelineHeader}>
                            <span className={styles.timelineVersion}>Versão {versao.VersionLabel}</span>
                            <span className={styles.timelineDate}>{this.formatDate(versao.Modified)}</span>
                          </div>
                          <div className={styles.timelineUser}>
                            Modificado por <strong>{versao.Editor}</strong>
                          </div>
                          <div className={styles.timelineComment}>
                            {versao.CheckInComment || 'Revisão aprovada sem comentários adicionais.'}
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className={styles.emptyState}><p>Nenhum histórico de versão encontrado para este documento.</p></div>
                    )}
                  </div>
                )}

              </div>

              {/* FOOTER DO MODAL */}
              <div className={styles.editModalFooter}>
                <button className={styles.cancelBtn} onClick={() => this.setState({ documentoSelecionado: null, activeModalTab: 'metadados', documentHistory: [] })}>Cancelar</button>
                {this.state.activeModalTab === 'metadados' && (
                  <button className={styles.saveBtn} onClick={this.salvarEdicaoDocumento} disabled={this.state.salvandoDocumento}>
                    {this.state.salvandoDocumento ? 'Salvando...' : 'Salvar Alterações'}
                  </button>
                )}
              </div>

            </div>
          </div>
        )}

        {/* MODAL DO IFRAME */}
        {this.state.iframeDocumentoUrl && (
          <div className={(styles as any).iframeModalBackdrop} onClick={() => this.setState({ iframeDocumentoUrl: null })}>
            <div className={(styles as any).iframeModalHeader}>
              <button className={(styles as any).closeIframeBtn} onClick={() => this.setState({ iframeDocumentoUrl: null })}>✕ Fechar Documento</button>
            </div>
            <div
              className={(styles as any).iframeContainer}
              onClick={(e) => e.stopPropagation()}
              onContextMenu={(e) => e.preventDefault()}
            >
              <iframe src={this.state.iframeDocumentoUrl} title="Document Viewer" />
            </div>
          </div>
        )}

        {/* =========================================================
        PASSO 1: MODAL DE ESCOLHA DO TIPO DE DOCUMENTO
    ========================================================= */}
        {this.state.isCreateModalOpen && (
          <div className={styles.editModalBackdrop}>
            <div className={styles.editModal}>

              <div className={styles.editModalHeader}>
                <h2>Criar Novo Documento</h2>
                <button onClick={() => this.setState({ isCreateModalOpen: false, selectedNewDocType: '' })} className={styles.closeModal}>✕</button>
              </div>

              <div className={styles.editModalBody}>
                <p style={{ fontSize: '14px', color: '#6B7280', marginBottom: '20px' }}>
                  Selecione o tipo de documento que deseja criar.
                </p>
                <div className={styles.formGroup}>
                  <label>Tipo de Processo/Documento</label>
                  <select
                    value={this.state.selectedNewDocType}
                    onChange={(e) => this.setState({ selectedNewDocType: e.target.value })}
                  >
                    <option value="">Selecione...</option>
                    <option value="MAPEAMENTO DE PROCESSO">Mapeamento de Processo</option>
                    <option value="PROCEDIMENTO">Procedimento</option>
                    <option value="PROCEDIMENTO OPERACIONAL PADRÃO">Procedimento Operacional Padrão (POP)</option>
                    <option value="INSTRUÇÃO DE TRABALHO">Instrução de Trabalho</option>
                    <option value="FORMULÁRIO">Formulário</option>
                    <option value="MANUAL">Manual</option>
                    <option value="POLÍTICA">Política</option>
                  </select>
                </div>
              </div>

              <div className={styles.editModalFooter}>
                <button className={styles.cancelBtn} onClick={() => this.setState({ isCreateModalOpen: false, selectedNewDocType: '' })}>Cancelar</button>
                <button
                  className={styles.saveBtn}
                  disabled={!this.state.selectedNewDocType}
                  onClick={() => {
                    // Removemos o redirecionamento antigo do Forms.
                    // Agora apenas fechamos este modal, mantendo o tipo selecionado no estado.
                    this.setState({ isCreateModalOpen: false });
                  }}
                  style={{ opacity: !this.state.selectedNewDocType ? 0.5 : 1, cursor: !this.state.selectedNewDocType ? 'not-allowed' : 'pointer' }}
                >
                  Continuar para o Formulário ➔
                </button>
              </div>

            </div>
          </div>
        )}

{/* =========================================================
            PASSO 2: ROTEAMENTO DOS FORMULÁRIOS
        ========================================================= */}

        {/* 1. Formulário de MAPEAMENTO */}
        {!this.state.isCreateModalOpen && this.state.selectedNewDocType === 'MAPEAMENTO DE PROCESSO' && (
          <FormularioMapeamento
            tipoDocumento={this.state.selectedNewDocType}
            usuarioEmail={this.props.context.pageContext.user.email}
            spContext={this.props.context}
            onFechar={() => this.setState({ selectedNewDocType: '' })}
            onSucesso={() => {
              this.setState({ selectedNewDocType: '' });
              this.buscarTodosDocumentos();
            }}
          />
        )}

        {/* 2. Formulário de PROCEDIMENTO (Apenas Ele) */}
        {!this.state.isCreateModalOpen && this.state.selectedNewDocType === 'PROCEDIMENTO' && (
          <FormularioProcedimento
            tipoDocumento={this.state.selectedNewDocType}
            usuarioEmail={this.props.context.pageContext.user.email}
            spContext={this.props.context}
            onFechar={() => this.setState({ selectedNewDocType: '' })}
            onSucesso={() => {
              this.setState({ selectedNewDocType: '' });
              this.buscarTodosDocumentos();
            }}
          />
        )}

        {/* 3. Formulário PADRÃO SGQ (Apenas Instrução de Trabalho) */}
        {!this.state.isCreateModalOpen && this.state.selectedNewDocType === 'INSTRUÇÃO DE TRABALHO' && (
          <FormularioSGQ
            tipoDocumento={this.state.selectedNewDocType}
            usuarioEmail={this.props.context.pageContext.user.email}
            spContext={this.props.context}
            onFechar={() => this.setState({ selectedNewDocType: '' })}
            onSucesso={() => {
              this.setState({ selectedNewDocType: '' });
              this.buscarTodosDocumentos();
            }}
          />
        )}

        {/* 4. AVISO DE DESENVOLVIMENTO (Bloqueia POP, Política, Formulário e Manual) */}
        {!this.state.isCreateModalOpen && this.state.selectedNewDocType && !['MAPEAMENTO DE PROCESSO', 'PROCEDIMENTO', 'INSTRUÇÃO DE TRABALHO'].includes(this.state.selectedNewDocType) && (
          <div className={styles.editModalBackdrop}>
            <div className={styles.editModal} style={{ width: '90%', maxWidth: '450px', textAlign: 'center', padding: '40px 30px', borderTop: '6px solid #A6CE39' }}>
              
              <div style={{ fontSize: '50px', marginBottom: '15px' }}>🚀</div>
              
              <h2 style={{ color: '#1C2510', margin: '0 0 15px 0', fontSize: '22px', fontWeight: '800' }}>
                Módulo em Desenvolvimento
              </h2>
              
              <p style={{ color: '#6B7280', fontSize: '15px', lineHeight: '1.6', marginBottom: '30px' }}>
                O gerador automatizado para <strong>{this.state.selectedNewDocType}</strong> está sendo construído pela nossa equipe de Tecnologia para garantir a melhor experiência possível.
                <br/><br/>
                No momento, os modelos já liberados para uso são: <strong>Instrução de Trabalho, Mapeamento de Processo e Procedimento</strong>.
              </p>
              
              <button
                onClick={() => this.setState({ selectedNewDocType: '', isCreateModalOpen: true })}
                style={{ backgroundColor: '#A6CE39', color: '#1C2510', border: 'none', padding: '12px 25px', borderRadius: '8px', fontWeight: '700', cursor: 'pointer', fontSize: '14px', transition: 'all 0.2s', boxShadow: '0 4px 10px rgba(166, 206, 57, 0.2)' }}
                onMouseEnter={(e) => e.currentTarget.style.transform = 'translateY(-2px)'}
                onMouseLeave={(e) => e.currentTarget.style.transform = 'translateY(0)'}
              >
                ⬅️ Voltar e escolher outro modelo
              </button>
              
              <button 
                onClick={() => this.setState({ selectedNewDocType: '' })}
                style={{ background: 'none', border: 'none', color: '#9CA3AF', marginTop: '15px', cursor: 'pointer', fontSize: '13px', textDecoration: 'underline' }}
              >
                Cancelar e fechar
              </button>
              
            </div>
          </div>
        )}
      </div>
    );
  }
}