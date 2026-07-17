import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import type { IPoliticasGrunnerProps } from './IPoliticasGrunnerProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MenuChamados } from '../../../shared/components/MenuChamado/MenuChamados';

// URLs de navegação
const logoGrunner = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo-grunner.png";
const homeUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Inicio.aspx?env=Embedded";
const historiaUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded";
const politicasUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded";
const atalhosUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/centraldeatalhos.aspx?env=Embedded";

export interface IPoliticaDocumento {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  Area?: string;
  CodigoDocumento?: string;
  TipoDocumento?: string;
  NumeroRevisao?: string;
  DataUltimaRevisao?: string;
  DataProximaRevisao?: string;
  StatusDocumento?: string;
  StatusCalculado?: string; // Usado no frontend
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
  isMobileMenuOpen: boolean;
  isMenuTIOpen: boolean;
  isQualidadeUser: boolean;
  modoGestaoQualidade: boolean;
  documentoSelecionado?: IPoliticaDocumento | null;
  salvandoDocumento: boolean;
  filtroStatusAdmin: string;
  editFormData: Partial<IPoliticaDocumento>;
  isMenuProcedimentosOpen: boolean;
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
      editFormData: {}
    };
  }

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

  public componentDidMount(): void {
    this.verificarAcessoQualidade();
    this.buscarTodosDocumentos();

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
      const select = 'Id,FileLeafRef,FileRef,Area,CodigoDocumento,TipoDocumento,NumeroRevisao,DataUltimaRevisao,DataProximaRevisao,StatusDocumento,ObservacaoRevisao,PeriodicidadeRevisaoMeses,UltimoAvisoRevisao,DiasAvisoRevisao,PermiteImpressaoControlada,ExibirNaIntranet,ResponsavelRevisao/Title,ResponsavelRevisao/EMail,AprovadorQualidade/Title,AprovadorQualidade/EMail';
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

  // Governança: Lógica de Status
  private calcularStatusDocumento = (doc: any): string => {
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
        CodigoDocumento: editFormData.CodigoDocumento,
        TipoDocumento: editFormData.TipoDocumento,
        NumeroRevisao: editFormData.NumeroRevisao,
        DataUltimaRevisao: editFormData.DataUltimaRevisao,
        DataProximaRevisao: editFormData.DataProximaRevisao,
        StatusDocumento: editFormData.StatusDocumento,
        ObservacaoRevisao: editFormData.ObservacaoRevisao,
        PeriodicidadeRevisaoMeses: editFormData.PeriodicidadeRevisaoMeses,
        PermiteImpressaoControlada: editFormData.PermiteImpressaoControlada,
        ExibirNaIntranet: editFormData.ExibirNaIntranet,
        Area: editFormData.Area
      };

 // Resolução de usuários (Pessoa ou Grupo)
      if (editFormData.ResponsavelRevisao?.EMail) {
        payload.ResponsavelRevisaoId = await this.getUserIdByEmail(editFormData.ResponsavelRevisao.EMail);
      } else if (editFormData.ResponsavelRevisao?.EMail === '') {
        // Se o campo for limpo no React, força o valor nulo no SharePoint
        payload.ResponsavelRevisaoId = null; 
      }

      if (editFormData.AprovadorQualidade?.EMail) {
        payload.AprovadorQualidadeId = await this.getUserIdByEmail(editFormData.AprovadorQualidade.EMail);
      } else if (editFormData.AprovadorQualidade?.EMail === '') {
        payload.AprovadorQualidadeId = null; 
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

      this.setState({ documentoSelecionado: null, salvandoDocumento: false });
      this.buscarTodosDocumentos(); // Recarrega os dados
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
      // Regra Pública: Não exibir arquivados nem VENCIDOS, e exibir apenas marcados para Intranet
      documentosExibidos = documentosExibidos.filter(d => 
        d.StatusCalculado !== 'Arquivado' && 
        d.StatusCalculado !== 'Vencido' &&
        d.StatusCalculado !== 'Em revisão' && 
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

    return (
      <div className={styles.container}>
        {this.shouldHideSharePointChrome() && (
          <style dangerouslySetInnerHTML={{__html: `... ocultações do sharepoint originais mantidas no seu código ...`}} />
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
              href="#"
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
            
            {/* BOTÃO PRINCIPAL DE PROCEDIMENTOS (ACORDEÃO) */}
            <a
              href="#"
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
                
                {/* BOTÃO DA GESTÃO DE QUALIDADE (Aparece só para quem tem acesso) */}
                {this.state.isQualidadeUser && (
                  <a 
                    href="#" 
                    className={this.state.modoGestaoQualidade ? styles.active : ''} 
                    onClick={(e) => { e.preventDefault(); this.setState({ modoGestaoQualidade: !this.state.modoGestaoQualidade }); }}
                  >
                    ⚙️ Gestão da Qualidade
                  </a>
                )}
              </div>
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
            <div className={styles.metricCard}>
              <div className={styles.metricLabel}>Total</div>
              <div className={styles.metricValue}>{total}</div>
            </div>
            <div className={`${styles.metricCard} ${styles.metricVigente}`}>
              <div className={styles.metricLabel}>Vigentes</div>
              <div className={styles.metricValue}>{vigentes}</div>
            </div>
            <div className={`${styles.metricCard} ${styles.metricAtencao}`}>
              <div className={styles.metricLabel}>Vence em breve</div>
              <div className={styles.metricValue}>{atencao}</div>
            </div>
            <div className={`${styles.metricCard} ${styles.metricVencido}`}>
              <div className={styles.metricLabel}>Vencidos</div>
              <div className={styles.metricValue}>{vencidos}</div>
            </div>
            <div className={`${styles.metricCard} ${styles.metricRevisao}`}>
              <div className={styles.metricLabel}>Em Revisão</div>
              <div className={styles.metricValue}>{revisao}</div>
            </div>
          </div>

          <div className={styles.searchContainer}>
            <input type="text" placeholder="🔍 Buscar por nome ou código..." value={termoBusca} onChange={(e) => this.setState({ termoBusca: e.target.value })} className={styles.searchInput} />
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

                      {/* CORPO DO CARD (Área, Tipo e Código) */}
                      <div className={styles.cardBody}>
                        <span className={styles.areaBadge}>
                          {doc.Area || 'Geral'} {doc.TipoDocumento ? `• ${doc.TipoDocumento}` : ''}
                        </span>
                        <span className={styles.docCode}>
                          {doc.CodigoDocumento ? `Código: ${doc.CodigoDocumento}` : <span className={styles.emptyCode}>Sem código</span>}
                        </span>
                      </div>

                      {/* RODAPÉ DO CARD (Revisão, Vencimento e Botão) */}
                      <div className={styles.cardFooter}>
                        <div className={styles.revisionInfo}>
                          <span className={styles.revText}>Rev. {doc.NumeroRevisao || '00'}</span>
                          <span className={styles.venceText}>Vence: {this.formatDate(doc.DataProximaRevisao)}</span>
                        </div>
                        <a href={`${doc.FileRef}?web=1`} target="_blank" rel="noopener noreferrer" className={styles.openButton}>
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
                <h2>Editar Metadados: {this.state.documentoSelecionado.FileLeafRef}</h2>
                <button onClick={() => this.setState({ documentoSelecionado: null })} className={styles.closeModal}>✕</button>
              </div>
              <div className={styles.editModalBody}>
              <div className={styles.formGrid}>
                  <div className={styles.formGroup}>
                    <label>Código do Documento</label>
                    <input type="text" value={this.state.editFormData.CodigoDocumento || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, CodigoDocumento: e.target.value } })} />
                  </div>
                  <div className={styles.formGroup}>
                    <label>Número da Revisão</label>
                    <input type="text" value={this.state.editFormData.NumeroRevisao || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, NumeroRevisao: e.target.value } })} />
                  </div>
                  
                  {/* BLOCOS DE DATA CORRIGIDOS */}
                  <div className={styles.formGroup}>
                    <label>Data Última Revisão</label>
                    <input 
                      type="date" 
                      value={this.state.editFormData.DataUltimaRevisao ? this.state.editFormData.DataUltimaRevisao.split('T')[0] : ''} 
                      onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, DataUltimaRevisao: e.target.value ? `${e.target.value}T12:00:00Z` : null as any } })} 
                    />
                  </div>

                  <div className={styles.formGroup}>
                    <label>Data Próxima Revisão</label>
                    <input 
                      type="date" 
                      value={this.state.editFormData.DataProximaRevisao ? this.state.editFormData.DataProximaRevisao.split('T')[0] : ''} 
                      onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, DataProximaRevisao: e.target.value ? `${e.target.value}T12:00:00Z` : null as any } })} 
                    />
                  </div>

                  <div className={styles.formGroup}>
                    <label>Status (Sobrescreve regra de data se Arquivado/Em revisão)</label>
                    <select value={this.state.editFormData.StatusDocumento || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, StatusDocumento: e.target.value } })}>
                      <option value="">Automático (Pela Data)</option>
                      <option value="Em revisão">Em revisão</option>
                      <option value="Arquivado">Arquivado</option>
                    </select>
                  </div>
                  <div className={styles.formGroup}>
                    <label>E-mail do Responsável</label>
                    <input type="email" placeholder="email@grunner.com.br" value={this.state.editFormData.ResponsavelRevisao?.EMail || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, ResponsavelRevisao: { EMail: e.target.value } } })} />
                  </div>
                </div>
              </div>
              <div className={styles.editModalFooter}>
                <button className={styles.cancelBtn} onClick={() => this.setState({ documentoSelecionado: null })}>Cancelar</button>
                <button className={styles.saveBtn} onClick={this.salvarEdicaoDocumento} disabled={this.state.salvandoDocumento}>
                  {this.state.salvandoDocumento ? 'Salvando...' : 'Salvar Alterações'}
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }
}