import * as React from 'react';
import styles from './CentralAtalhosGrunner.module.scss';
import type { ICentralAtalhosGrunnerProps } from './ICentralAtalhosGrunnerProps';
import { SPHttpClient } from '@microsoft/sp-http';

const logoGrunner = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo-grunner.png";
const logoCompleta = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo.png";
const homeUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Inicio.aspx?env=Embedded";
const historiaUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded";
const politicasUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded";
const atalhosUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/CentralAtalhos.aspx?env=Embedded";

interface ILinkUtil {
  ID: number;
  Title: string;
  Descricao?: string;
  Categoria?: string;
  Icone?: string;
  LinkURL?: any;
  Ordem?: number;
  Ativo?: boolean | number | string;
  Destaque?: boolean | number | string;
}

interface ICentralAtalhosGrunnerState {
  todosLinks: ILinkUtil[];
  loading: boolean;
  termoBusca: string;
  categoriaAtiva: string;
  isMobileMenuOpen: boolean;
}

export default class CentralAtalhosGrunner extends React.Component<ICentralAtalhosGrunnerProps, ICentralAtalhosGrunnerState> {
  private footerObserver?: MutationObserver;

  constructor(props: ICentralAtalhosGrunnerProps) {
    super(props);

    this.state = {
      todosLinks: [],
      loading: true,
      termoBusca: '',
      categoriaAtiva: 'Todos',
      isMobileMenuOpen: false
    };
  }

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

    // AQUI ESTAVA O BUG: Removida a linha que apagava os "avós" da página e bloqueava o clique!
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
    this.buscarLinks();

    if (this.shouldHideSharePointChrome()) {
      const applyFixes = (): void => {
        this.hideSharePointFooter();
        this.hideSharePointAppBar();
        this.fixSharePointCanvasSpacing();
      };

      applyFixes();
      window.setTimeout(applyFixes, 500);
      window.setTimeout(applyFixes, 1500);

      this.footerObserver = new MutationObserver(() => {
        applyFixes();
      });

      if (document.body) {
        this.footerObserver.observe(document.body, {
          childList: true,
          subtree: true
        });
      }
    }
  }

  public componentWillUnmount(): void {
    if (this.footerObserver) {
      this.footerObserver.disconnect();
    }
  }

  private buscarLinks = async (): Promise<void> => {
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('LinksUteisGrunner')/items?$select=ID,Title,Descricao,Categoria,Icone,LinkURL,Ordem,Ativo,Destaque&$top=5000&$orderby=Ordem asc`;

      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data?.value) {
        this.setState({
          todosLinks: data.value,
          loading: false
        });
      } else {
        this.setState({ loading: false });
      }
    } catch (error) {
      console.error('Erro ao buscar links úteis:', error);
      this.setState({ loading: false });
    }
  }

  private isEnabled = (value: unknown): boolean => {
    if (value === undefined || value === null || value === '') return true;
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value === 1;

    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      return normalized === 'true' || normalized === '1' || normalized === 'sim' || normalized === 'yes';
    }

    return Boolean(value);
  }

  private isTruthy = (value: unknown): boolean => {
    if (value === undefined || value === null || value === '') return false;
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value === 1;

    if (typeof value === 'string') {
      const normalized = value.trim().toLowerCase();
      return normalized === 'true' || normalized === '1' || normalized === 'sim' || normalized === 'yes';
    }

    return Boolean(value);
  }

  private normalizeCategory = (categoria?: string): string => {
    if (!categoria || !categoria.trim()) return 'Outros';
    return categoria.trim();
  }

  private removerAcentos = (texto: string): string => {
    if (!texto) return '';
    return texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  }

  private resolveLinkUrl = (linkValue: any): string => {
    if (!linkValue) return '#';

    if (typeof linkValue === 'string') {
      const trimmed = linkValue.trim();

      if (/^https?:\/\//i.test(trimmed) && trimmed.includes(',')) {
        return trimmed.split(',')[0].trim();
      }

      return trimmed;
    }

    if (typeof linkValue === 'object') {
      if (linkValue.Url) return linkValue.Url;
      if (linkValue.url) return linkValue.url;
      if (linkValue.href) return linkValue.href;
    }

    return '#';
  }

  private resolveIcon = (link: ILinkUtil): string => {
    if (link.Icone && link.Icone.trim()) return link.Icone.trim();

    const categoria = this.normalizeCategory(link.Categoria).toLowerCase();

    if (categoria.includes('ti')) return '💻';
    if (categoria.includes('marketing')) return '📣';
    if (categoria.includes('rh')) return '👥';
    if (categoria.includes('oper')) return '⚙️';
    if (categoria.includes('facilities')) return '🛠️';
    if (categoria.includes('frotas')) return '🚗';
    if (categoria.includes('institucional')) return '🏛️';
    if (categoria.includes('comercial')) return '🤝';
    if (categoria.includes('finance')) return '💰';

    return '🔗';
  }

  private sortByOrder = (a: ILinkUtil, b: ILinkUtil): number => {
    const ordemA = typeof a.Ordem === 'number' ? a.Ordem : 9999;
    const ordemB = typeof b.Ordem === 'number' ? b.Ordem : 9999;

    if (ordemA !== ordemB) return ordemA - ordemB;

    return (a.Title || '').localeCompare(b.Title || '');
  }

  public render(): React.ReactElement<ICentralAtalhosGrunnerProps> {
    const userEmail = this.props.context?.pageContext?.user?.email || '';
    const nomeUsuario = this.props.userDisplayName?.split(' ')[0] || 'Colaborador';
    const { todosLinks, loading, termoBusca, categoriaAtiva } = this.state;

    const linksAtivos = todosLinks
      .filter((link) => this.isEnabled(link.Ativo))
      .filter((link) => this.resolveLinkUrl(link.LinkURL) !== '#')
      .sort(this.sortByOrder);

    const categorias = [
      'Todos',
      ...Array.from(new Set(linksAtivos.map((link) => this.normalizeCategory(link.Categoria))))
    ];

    const linksDestaque = linksAtivos
      .filter((link) => this.isTruthy(link.Destaque))
      .slice(0, 4);

    const termoLimpo = this.removerAcentos(termoBusca).trim();

    const linksFiltrados = linksAtivos.filter((link) => {
      const titulo = this.removerAcentos(link.Title || '');
      const descricao = this.removerAcentos(link.Descricao || '');
      const categoria = this.removerAcentos(this.normalizeCategory(link.Categoria));

      const passouBusca =
        !termoLimpo ||
        titulo.includes(termoLimpo) ||
        descricao.includes(termoLimpo) ||
        categoria.includes(termoLimpo);

      const passouCategoria =
        categoriaAtiva === 'Todos' ||
        this.normalizeCategory(link.Categoria) === categoriaAtiva;

      return passouBusca && passouCategoria;
    });

    return (
      <div className={styles.container}>
        {/* AQUI ESTÁ A BLINDAGEM COMPLETA DAS BARRAS SUPERIORES E INFERIORES */}
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

            body {
              overflow-x: hidden !important;
            }
          `}</style>
        )}

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
            <a href={atalhosUrl} className={styles.active}>🖥️ Central de Atalhos</a>
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
            <a href={politicasUrl} target="_blank" rel="noopener noreferrer">📖 Políticas da Empresa</a>
          </div>
        </aside>

        <div className={styles.contentArea}>
          <header className={styles.header}>
            <div className={styles.headerLeft}>
              <img
                src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${userEmail}`}
                alt="Perfil"
                className={styles.userAvatar}
                onError={(e) => { e.currentTarget.style.display = 'none'; }}
              />
              <div className={styles.headerText}>
                <h1>Olá, {nomeUsuario}!</h1>
                <p>Bem-vindo à Central de Atalhos Grunner • seu desktop corporativo para acessos rápidos</p>
                <span className={styles.dateBadge}>🖥️ Links úteis, sistemas e páginas importantes em um só lugar</span>
              </div>
            </div>
            <img src={logoCompleta} className={styles.logoCentral} alt="Grunner" />
          </header>

          <main className={styles.mainContent}>
            <section className={styles.heroSection}>
              <a href={homeUrl} className={styles.backBtn}>← Voltar para a Intranet</a>

              <div className={styles.heroContent}>
                <h1>Central de Atalhos Grunner</h1>
                <p>
                  Um desktop corporativo com os links mais usados pelos times.
                  Ideal para acelerar o dia a dia e facilitar a vida de novos colaboradores.
                </p>
              </div>
            </section>

            <section className={styles.toolbarSection}>
              <div className={styles.searchContainer}>
                <input
                  type="text"
                  placeholder="🔍 Buscar sistema, área, processo ou nome do atalho..."
                  value={termoBusca}
                  onChange={(e) => this.setState({ termoBusca: e.target.value })}
                  className={styles.searchInput}
                />
              </div>

              <nav className={styles.tabsContainer}>
                {categorias.map((categoria) => (
                  <button
                    key={categoria}
                    className={categoriaAtiva === categoria ? styles.tabActive : styles.tab}
                    onClick={() => this.setState({ categoriaAtiva: categoria })}
                  >
                    {categoria}
                  </button>
                ))}
              </nav>
            </section>

            {linksDestaque.length > 0 && !termoBusca.trim() && categoriaAtiva === 'Todos' && (
              <section className={styles.featuredSection}>
                <div className={styles.sectionHeader}>
                  <h2>⭐ Mais acessados</h2>
                  <p>Os atalhos que normalmente ajudam mais no início da jornada.</p>
                </div>

                <div className={styles.featuredGrid}>
                  {linksDestaque.map((link) => (
                    <a
                      key={link.ID}
                      href={this.resolveLinkUrl(link.LinkURL)}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={styles.featuredCard}
                    >
                      <div className={styles.featuredIcon}>{this.resolveIcon(link)}</div>
                      <div className={styles.featuredInfo}>
                        <h3>{link.Title}</h3>
                        <span>{this.normalizeCategory(link.Categoria)}</span>
                      </div>
                    </a>
                  ))}
                </div>
              </section>
            )}

            <section className={styles.desktopSection}>
              <div className={styles.sectionHeader}>
                <h2>🗂️ Área de trabalho</h2>
                <p>
                  {loading
                    ? 'Carregando atalhos...'
                    : `${linksFiltrados.length} atalho(s) encontrado(s)`}
                </p>
              </div>

              {loading ? (
                <div className={styles.loadingState}>
                  <div className={styles.spinner}></div>
                  <p>Montando o desktop da equipe...</p>
                </div>
              ) : linksFiltrados.length > 0 ? (
                <div className={styles.desktopSurface}>
                  <div className={styles.desktopGrid}>
                    {linksFiltrados.map((link) => (
                      <a
                        key={link.ID}
                        href={this.resolveLinkUrl(link.LinkURL)}
                        target="_blank"
                        rel="noopener noreferrer"
                        className={styles.shortcutCard}
                        title={link.Descricao || link.Title}
                      >
                        <div className={styles.shortcutIcon}>{this.resolveIcon(link)}</div>
                        <div className={styles.shortcutLabel}>{link.Title}</div>
                        <div className={styles.shortcutMeta}>{this.normalizeCategory(link.Categoria)}</div>
                        {link.Descricao && (
                          <p className={styles.shortcutDescription}>{link.Descricao}</p>
                        )}
                      </a>
                    ))}
                  </div>
                </div>
              ) : (
                <div className={styles.emptyState}>
                  <div className={styles.emptyIcon}>🧭</div>
                  <h3>Nenhum atalho encontrado</h3>
                  <p>Tente mudar a busca ou a categoria selecionada.</p>
                </div>
              )}
            </section>
          </main>
        </div>
      </div>
    );
  }
}