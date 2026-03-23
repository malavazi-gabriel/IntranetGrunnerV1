import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import type { IPoliticasGrunnerProps } from './IPoliticasGrunnerProps';
import { SPHttpClient } from '@microsoft/sp-http';

// URLs de navegação
const logoGrunner = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo-grunner.png";
const homeUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Inicio.aspx?env=Embedded";
const historiaUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded";
const politicasUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded";
const atalhosUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/CentralAtalhos.aspx?env=Embedded";

interface IPoliticasGrunnerState {
  areaAtiva: string;
  todosDocumentos: any[];
  loading: boolean;
  termoBusca: string;
  isMobileMenuOpen: boolean;
}

export default class PoliticasGrunner extends React.Component<IPoliticasGrunnerProps, IPoliticasGrunnerState> {
  
  private areas = ['Institucional', 'TI', 'Marketing', 'RH', 'Operacional'];
  private footerObserver?: MutationObserver;

  constructor(props: IPoliticasGrunnerProps) {
    super(props);
    this.state = {
      areaAtiva: 'Institucional', 
      todosDocumentos: [],
      loading: true,
      termoBusca: '',
      isMobileMenuOpen: false
    };
  }

  // Lógica para esconder o visual padrão do SharePoint
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
    element.style.setProperty('visibility', 'hidden', 'important');
    element.style.setProperty('height', '0', 'important');
    element.style.setProperty('min-height', '0', 'important');
    element.style.setProperty('max-height', '0', 'important');
    element.style.setProperty('margin', '0', 'important');
    element.style.setProperty('padding', '0', 'important');
    element.style.setProperty('overflow', 'hidden', 'important');
    element.style.setProperty('opacity', '0', 'important');
    element.style.setProperty('pointer-events', 'none', 'important');
  }

  private hideSharePointFooter = (): void => {
    const selectors = [
      '[data-automation-id="page-bottom-actions"]',
      '[data-automation-id="page-bottom-bar"]',
      '#sp-page-footer',
      '[data-automation-id="socialBar"]',
      '.CommentsWrapper',
      '[id*="Page_CommentsWrapper"]',
      '[id^="Page_CommentsWrapper"]',
      '[data-sp-feature-tag="Comments"]'
    ];
    const elements = document.querySelectorAll(selectors.join(','));
    elements.forEach((node) => {
      this.collapseElement(node as HTMLElement);
    });
  }

  private hideSharePointAppBar = (): void => {
    const selectors = [
      '#sp-appBar',
      '[data-automation-id="sp-appBar"]',
      'div[class^="appBar_"]',
      'div[class*="sp-appBar"]'
    ];
    const elements = document.querySelectorAll(selectors.join(','));
    elements.forEach((node) => {
      this.collapseElement(node as HTMLElement);
    });
  }

  private fixSharePointCanvasSpacing = (): void => {
    const selectors = [
      '#workbenchPageContent',
      '#spPageCanvasContent',
      '.SPCanvas-canvas',
      '.CanvasZone',
      '.CanvasSection',
      '.ControlZone',
      'div[data-automation-id="CanvasZone"] > div'
    ];
    const elements = document.querySelectorAll(selectors.join(','));
    elements.forEach((node) => {
      const el = node as HTMLElement;
      el.style.setProperty('margin-left', '0', 'important');
      el.style.setProperty('padding-left', '0', 'important');
      el.style.setProperty('max-width', '100%', 'important');
      el.style.setProperty('width', '100%', 'important');
    });
    document.body?.style.setProperty('overflow-x', 'hidden', 'important');
  }

  public componentDidMount(): void {
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

      this.footerObserver = new MutationObserver(() => { applyFixes(); });
      if (document.body) {
        this.footerObserver.observe(document.body, { childList: true, subtree: true });
      }
    }
  }

  public componentWillUnmount(): void {
    if (this.footerObserver) {
      this.footerObserver.disconnect();
    }
  }

  private buscarTodosDocumentos = async (): Promise<void> => {
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PoliticasGrunner')/items?$select=FileLeafRef,FileRef,Area&$top=5000`;
      
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data && data.value) {
        this.setState({ todosDocumentos: data.value, loading: false });
      } else {
        this.setState({ loading: false });
      }
    } catch (error) {
      console.error("Erro ao buscar documentos:", error);
      this.setState({ loading: false });
    }
  }

  public render(): React.ReactElement<IPoliticasGrunnerProps> {
    const { areaAtiva, todosDocumentos, termoBusca, loading } = this.state;

    let documentosExibidos = [];
    const isBuscando = termoBusca.trim().length > 0;

    if (isBuscando) {
      documentosExibidos = todosDocumentos.filter(doc => 
        doc.FileLeafRef && doc.FileLeafRef.toLowerCase().includes(termoBusca.toLowerCase())
      );
    } else {
      documentosExibidos = todosDocumentos.filter(doc => 
        doc.Area === areaAtiva
      );
    }

    return (
      <div className={styles.container}>
        {this.shouldHideSharePointChrome() && (
          <style>{`
            [data-automation-id="page-bottom-actions"],
            [data-automation-id="page-bottom-bar"],
            #sp-page-footer,
            [data-automation-id="socialBar"],
            .CommentsWrapper,
            [id*="Page_CommentsWrapper"],
            [id^="Page_CommentsWrapper"],
            [data-sp-feature-tag="Comments"],
            #sp-appBar,
            [data-automation-id="sp-appBar"],
            div[class^="appBar_"],
            div[class*="sp-appBar"],
            #SuiteNavWrapper,
            #spSiteHeader,
            #spCommandBar,
            div[class^="commandBarWrapper_"],
            div[data-automation-id="pageHeader"] {
              display: none !important;
              visibility: hidden !important;
              height: 0 !important;
              min-height: 0 !important;
              max-height: 0 !important;
              margin: 0 !important;
              padding: 0 !important;
              overflow: hidden !important;
              opacity: 0 !important;
              pointer-events: none !important;
            }
            #workbenchPageContent,
            #spPageCanvasContent,
            .SPCanvas-canvas,
            .CanvasZone,
            .CanvasSection,
            .ControlZone,
            div[data-automation-id="CanvasZone"] > div {
              margin-left: 0 !important;
              padding-left: 0 !important;
              max-width: 100% !important;
              width: 100% !important;
            }
            body { overflow-x: hidden !important; }
          `}</style>
        )}

        {/* BARRA MOBILE E OVERLAY */}
        <div className={styles.mobileHeaderBar}>
          <button
            className={styles.hamburgerBtn}
            onClick={() => this.setState({ isMobileMenuOpen: true })}
          >
            ☰ Menu Grunner
          </button>
        </div>

        {this.state.isMobileMenuOpen && (
          <div
            className={styles.mobileOverlayBackdrop}
            onClick={() => this.setState({ isMobileMenuOpen: false })}
          />
        )}

        {/* SIDEBAR */}
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
            <a href="https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5" target="_blank" rel="noopener noreferrer">🖥️ TI</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/Marketing/_layouts/15/listforms.aspx?cid=MTQ1MjlmMzEtNjk2Ni00MTI2LWJhNzItMzE1MTc0NDU2YTE4&nav=MGIwZDdiNzMtODQwNi00MDhiLTk5ZDEtNGE5NWNlYzljNDg3" target="_blank" rel="noopener noreferrer">📢 Marketing</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/GPS/_layouts/15/listforms.aspx?cid=ZWFlMDE1MWUtOTFlMS00MmJiLWFiNzEtOWM0NGVkZTVkMTdh&nav=ZGJmNmMxZGMtNjU5Zi00ZTUxLThjMTctZmFhODY5YTQ3NjBi" target="_blank" rel="noopener noreferrer">🚗 Frotas</a>
            <a href="https://forms.monday.com/forms/2a2a29caa20e7e1517cc397586af97eb?r=use1" target="_blank" rel="noopener noreferrer">🛠️ Facilities</a>
          </div>

          <div className={styles.navGroup}>
            <h3>Institucional</h3>
            <a href={historiaUrl} target="_blank" rel="noopener noreferrer">🏛️ Nossa História</a>
            <a href={politicasUrl} className={styles.active}>📖 Políticas da Empresa</a>
          </div>
        </aside>

        {/* CONTEÚDO PRINCIPAL (Ex-Container) */}
        <div className={styles.contentArea}>
          <header className={styles.pageHeader}>
            <div className={styles.headerText}>
              <h1>📖 Políticas e Diretrizes Grunner</h1>
              <p>Acesse os documentos normativos, manuais e procedimentos de cada área da empresa.</p>
            </div>
          </header>

          <div className={styles.searchContainer}>
            <input 
              type="text" 
              placeholder="🔍 Buscar qualquer política, manual ou termo..."
              value={termoBusca}
              onChange={(e) => this.setState({ termoBusca: e.target.value })}
              className={styles.searchInput}
            />
          </div>

          <nav className={`${styles.tabsContainer} ${isBuscando ? styles.tabsDisabled : ''}`}>
            {this.areas.map((area) => (
              <button
                key={area}
                className={areaAtiva === area && !isBuscando ? styles.tabActive : styles.tab}
                onClick={() => this.setState({ areaAtiva: area, termoBusca: '' })}
              >
                {area}
              </button>
            ))}
          </nav>

          <main className={styles.documentsArea}>
            {loading ? (
              <div className={styles.loadingState}>
                <div className={styles.spinner}></div>
                <p>Carregando biblioteca de políticas...</p>
              </div>
            ) : documentosExibidos.length > 0 ? (
              <div className={styles.documentGrid}>
                {documentosExibidos.map((doc, index) => {
                  const extensao = doc.FileLeafRef ? doc.FileLeafRef.split('.').pop()?.toLowerCase() : '';
                  const isPdf = extensao === 'pdf';
                  const areaDoc = doc.Area ? doc.Area : 'Geral';
                  
                  return (
                    <a key={index} href={doc.FileRef} target="_blank" rel="noopener noreferrer" className={styles.documentCard}>
                      <div className={isPdf ? styles.iconPdf : styles.iconDoc}>
                        {isPdf ? 'PDF' : 'DOC'}
                      </div>
                      <div className={styles.docInfo}>
                        <h3>{doc.FileLeafRef.replace(`.${extensao}`, '')}</h3>
                        <div className={styles.docMeta}>
                          <span className={styles.areaBadge}>{areaDoc}</span>
                          <span className={styles.clickText}>Abrir arquivo</span>
                        </div>
                      </div>
                    </a>
                  );
                })}
              </div>
            ) : (
              <div className={styles.emptyState}>
                {isBuscando ? (
                  <p>Nenhum documento encontrado na empresa inteira para "<strong>{termoBusca}</strong>".</p>
                ) : (
                  <p>📭 Nenhum documento cadastrado para a área de <strong>{areaAtiva}</strong> no momento.</p>
                )}
              </div>
            )}
          </main>
        </div>

      </div>
    );
  }
}