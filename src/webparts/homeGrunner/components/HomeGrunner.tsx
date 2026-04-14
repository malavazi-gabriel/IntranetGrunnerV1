import * as React from 'react';
import styles from './HomeGrunner.module.scss';
import { IHomeGrunnerProps } from './IHomeGrunnerProps';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

const logoGrunner = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo-grunner.png";
const logoCompleta = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SiteAssets/Logos/logo.png";

interface IHomeGrunnerState {
  isChamadoModalOpen: boolean;
  noticiasReais: any[];
  aniversariantesReais: any[];
  eventosReais: any[];
  loading: boolean;
  isModalOpen: boolean;
  currentNoticiaId: number | null;
  novoComentario: string;
  comentariosDaNoticia: any[];
  loadingComentarios: boolean;
  todasCurtidas: any[];
  todosComentarios: any[];
  isMobileMenuOpen: boolean;
  expandedNoticiaId: number | null;
  limiteNoticias: number;
  mostrarTodosAniversariantes: boolean;
}

export default class HomeGrunner extends React.Component<IHomeGrunnerProps, IHomeGrunnerState> {
  private footerObserver?: MutationObserver;

  constructor(props: IHomeGrunnerProps) {
    super(props);
    this.state = {
      isChamadoModalOpen: false,
      noticiasReais: [],
      aniversariantesReais: [],
      eventosReais: [],
      loading: true,
      isModalOpen: false,
      currentNoticiaId: null,
      novoComentario: "",
      comentariosDaNoticia: [],
      loadingComentarios: false,
      todasCurtidas: [],
      todosComentarios: [],
      isMobileMenuOpen: false,
      expandedNoticiaId: null,
      limiteNoticias: 7,
      mostrarTodosAniversariantes: false
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
    elements.forEach((node) => {
      const el = node as HTMLElement;
      const parent = el.parentElement as HTMLElement | null;
      const grandParent = parent?.parentElement as HTMLElement | null;
      this.collapseElement(el);
      this.collapseElement(parent);
      this.collapseElement(grandParent);
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
      const el = node as HTMLElement;
      this.collapseElement(el);
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
    this.carregarDadosIniciais();

    const urlParams = new URLSearchParams(window.location.search);
    const noticiaIdParam = urlParams.get('noticiaId');
    
    if (noticiaIdParam) {
      this.setState({ expandedNoticiaId: parseInt(noticiaIdParam, 10) });
    }

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
        this.footerObserver.observe(document.body, { childList: true, subtree: true });
      }
    }
  }

  public componentWillUnmount(): void {
    if (this.footerObserver) {
      this.footerObserver.disconnect();
    }
  }

  private carregarDadosIniciais = async () => {
    await Promise.all([
      this.buscarNoticias(),
      this.buscarAniversariantes(),
      this.buscarEventos(),
      this.buscarEngajamento()
    ]);
    this.setState({ loading: false });
  }

  private buscarNoticias = async () => {
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('NoticiasGrunner')/items?$select=ID,Title,Resumo,ImagemURL,LinkNoticia,ConteudoNoticia,Attachments,AttachmentFiles/ServerRelativeUrl&$expand=AttachmentFiles&$top=${this.state.limiteNoticias}&$orderby=Created desc`;
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      if (data?.value) this.setState({ noticiasReais: data.value });
    } catch (e) {
      console.error("Erro ao buscar notícias:", e);
    }
  }

  private carregarMaisNoticias = () => {
    this.setState((prevState) => ({
      limiteNoticias: prevState.limiteNoticias + 3
    }), this.buscarNoticias); 
  }

  private buscarEngajamento = async () => {
    try {
      const urlCurtidas = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('CurtidasGrunner')/items`;
      const urlComentarios = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ComentariosGrunner')/items`;
      const [respCurtidas, respComentarios] = await Promise.all([
        this.props.context.spHttpClient.get(urlCurtidas, SPHttpClient.configurations.v1),
        this.props.context.spHttpClient.get(urlComentarios, SPHttpClient.configurations.v1)
      ]);
      const dataCurtidas = await respCurtidas.json();
      const dataComentarios = await respComentarios.json();

      this.setState({
        todasCurtidas: dataCurtidas?.value || [],
        todosComentarios: dataComentarios?.value || []
      });
    } catch (e) {
      console.error("Erro ao buscar engajamento:", e);
    }
  }

  private buscarAniversariantes = async () => {
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('AniversariantesGrunner')/items?$select=Title,Dia,Setor,Email&$top=100`;
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      if (data?.value) this.setState({ aniversariantesReais: data.value });
    } catch (e) {
      console.error("Erro ao buscar aniversariantes:", e);
    }
  }

  private buscarEventos = async () => {
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EventosGrunner')/items?$select=Title,Dia,Mes,Local,ImagemTema&$top=3&$orderby=Created desc`;
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      if (data?.value) this.setState({ eventosReais: data.value });
    } catch (e) {
      console.error("Erro ao buscar eventos:", e);
    }
  }

  private isAniversarianteDaSemana = (diaStr: string): boolean => {
    const dia = parseInt(diaStr, 10);
    if (isNaN(dia)) return false;

    const hoje = new Date();
    const diasDaSemana: number[] = []; 
    
    const domingo = new Date(hoje);
    domingo.setDate(hoje.getDate() - hoje.getDay());

    for (let i = 0; i < 7; i++) {
      const dataDaSemana = new Date(domingo);
      dataDaSemana.setDate(domingo.getDate() + i);
      diasDaSemana.push(dataDaSemana.getDate());
    }

    return diasDaSemana.indexOf(dia) !== -1;
  }

  private handleLike = async (noticiaId: number) => {
    const userEmail = this.props.context.pageContext.user.email;
    const userName = this.props.userDisplayName;

    const likeExistente = this.state.todasCurtidas.find(
      c => c.NoticiaID === noticiaId.toString() && c.UsuarioEmail === userEmail
    );

    try {
      if (likeExistente) {
        const urlDelete = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('CurtidasGrunner')/items(${likeExistente.ID})`;
        await this.props.context.spHttpClient.post(urlDelete, SPHttpClient.configurations.v1, {
          headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
        });
      } else {
        const urlPost = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('CurtidasGrunner')/items`;
        const body = JSON.stringify({
          Title: `Like-${noticiaId}`,
          NoticiaID: noticiaId.toString(),
          UsuarioEmail: userEmail,
          UsuarioNome: userName
        });
        await this.props.context.spHttpClient.post(urlPost, SPHttpClient.configurations.v1, { body: body });
      }
      this.buscarEngajamento();
    } catch (e) {
      console.error("Erro ao processar curtida:", e);
    }
  }

  private getTextQuemCurtiu = (noticiaId: number) => {
    const curtidas = this.state.todasCurtidas.filter(c => c.NoticiaID === noticiaId.toString());
    if (curtidas.length === 0) return "Seja o primeiro a curtir!";
    const nomes = curtidas.map(c => c.UsuarioNome || c.UsuarioEmail.split('@')[0]);
    return `Curtido por:\n${nomes.join('\n')}`;
  }

  private openCommentModal = (id: number) => {
    this.setState({ isModalOpen: true, currentNoticiaId: id, novoComentario: "" });
    this.buscarComentarios(id);
  }

  private buscarComentarios = async (noticiaId: number) => {
    this.setState({ loadingComentarios: true, comentariosDaNoticia: [] });
    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ComentariosGrunner')/items?$filter=NoticiaID eq '${noticiaId}'&$orderby=Created desc`;
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data?.value) {
        this.setState({ comentariosDaNoticia: data.value, loadingComentarios: false });
      } else {
        this.setState({ loadingComentarios: false });
      }
    } catch (e) {
      console.error("Erro ao buscar comentários:", e);
      this.setState({ loadingComentarios: false });
    }
  }

  private enviarComentario = async () => {
    if (!this.state.novoComentario || !this.state.currentNoticiaId) return;

    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ComentariosGrunner')/items`;
    const body = JSON.stringify({
      Title: `Comentário-${this.state.currentNoticiaId}`,
      NoticiaID: this.state.currentNoticiaId.toString(),
      Comentario: this.state.novoComentario,
      Autor: this.props.userDisplayName
    });

    const options: ISPHttpClientOptions = { body: body };

    try {
      await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
      this.setState({ novoComentario: "" });
      this.buscarComentarios(this.state.currentNoticiaId);
      this.buscarEngajamento();
    } catch (e) {
      console.error("Erro ao enviar comentário:", e);
    }
  }

  private getLikesCount = (noticiaId: number) => {
    return this.state.todasCurtidas.filter(c => c.NoticiaID === noticiaId.toString()).length;
  }

  private getCommentsCount = (noticiaId: number) => {
    return this.state.todosComentarios.filter(c => c.NoticiaID === noticiaId.toString()).length;
  }

  private userAlreadyLiked = (noticiaId: number) => {
    const userEmail = this.props.context.pageContext.user.email;
    return this.state.todasCurtidas.some(c => c.NoticiaID === noticiaId.toString() && c.UsuarioEmail === userEmail);
  }

  private noticiaTemConteudo = (noticia: any): boolean => {
    const conteudo = (noticia?.ConteudoNoticia || '').toString().trim();
    return conteudo.length > 0;
  }

  private handleReadMore = (noticia: any): void => {
    if (!noticia) return;

    if (this.noticiaTemConteudo(noticia)) {
      this.setState((prevState) => ({
        expandedNoticiaId: prevState.expandedNoticiaId === noticia.ID ? null : noticia.ID
      }));
      return;
    }

    if (noticia?.LinkNoticia) {
      window.open(noticia.LinkNoticia, '_blank');
    }
  }

  private getImagemNoticia = (noticia: any): string => {
    if (noticia.Attachments && noticia.AttachmentFiles && noticia.AttachmentFiles.length > 0) {
      return noticia.AttachmentFiles[0].ServerRelativeUrl;
    }
    return noticia.ImagemURL || '';
  }

  private renderExpandedMainNews = (noticia: any): React.ReactNode => {
    if (!noticia || this.state.expandedNoticiaId !== noticia.ID || !this.noticiaTemConteudo(noticia)) {
      return null;
    }

    return (
      <div className={styles.expandedArticleWrapper}>
        <div dangerouslySetInnerHTML={{ __html: noticia.ConteudoNoticia }} />

        {noticia.LinkNoticia && (
          <div style={{ marginTop: '35px', display: 'flex', justifyContent: 'flex-start' }}>
            <button
              className={styles.btnPrimary}
              onClick={() => window.open(noticia.LinkNoticia, '_blank')}
            >
              Abrir Link Original ➔
            </button>
          </div>
        )}
      </div>
    );
  }

  private renderExpandedSubNewsCard = (noticia: any): React.ReactNode => {
    if (!noticia || this.state.expandedNoticiaId !== noticia.ID || !this.noticiaTemConteudo(noticia)) {
      return null;
    }

    const imagemExibicao = this.getImagemNoticia(noticia);

    return (
      <div style={{ width: '100%', display: 'flex', flexDirection: 'column' }}>
        <div className={styles.heroBanner} style={{ marginBottom: 0, borderRadius: '20px 20px 0 0' }}>
          <div className={styles.heroImage} style={{ backgroundImage: `url('${imagemExibicao}')` }} />
          <div className={styles.heroOverlay}>
            <span className={styles.badge}>Matéria em Leitura</span>
            <h2 className={styles.heroTitle}>{noticia.Title}</h2>
            {noticia.Resumo && (
              <p className={styles.heroResumo}>{noticia.Resumo}</p>
            )}

            <div className={styles.interactions}>
              <button
                className={styles.actionBtn}
                onClick={(e) => { e.stopPropagation(); this.handleLike(noticia.ID); }}
                title={this.getTextQuemCurtiu(noticia.ID)}
              >
                {this.userAlreadyLiked(noticia.ID) ? '❤️' : '🤍'} {this.getLikesCount(noticia.ID)} Curtidas
              </button>

              <button
                className={styles.actionBtn}
                onClick={(e) => { e.stopPropagation(); this.openCommentModal(noticia.ID); }}
              >
                💬 {this.getCommentsCount(noticia.ID)} Comentários
              </button>

              <button
                className={styles.actionBtn}
                style={{ marginLeft: 'auto', background: 'rgba(255,0,0,0.2)' }}
                onClick={() => this.handleReadMore(noticia)}
              >
                ✕ Fechar Matéria
              </button>
            </div>
          </div>
        </div>

        {this.renderExpandedMainNews(noticia)}
      </div>
    );
  }

  public render(): React.ReactElement<IHomeGrunnerProps> {
    const nomeUsuario = this.props.userDisplayName?.split(' ')[0] || 'Colaborador';
    const noticiaDestaque = this.state.noticiasReais[0];
    const outrasNoticias = this.state.noticiasReais.slice(1);

    const userEmail = this.props.context.pageContext.user.email;
    const dataAtual = new Date().toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' });

    return (
      <div className={styles.container}>
        {this.shouldHideSharePointChrome() && (
          <style>{`
            [data-automation-id="page-bottom-actions"], [data-automation-id="page-bottom-bar"], #sp-page-footer, [data-automation-id="socialBar"], .CommentsWrapper, [id*="Page_CommentsWrapper"], [id^="Page_CommentsWrapper"], [data-sp-feature-tag="Comments"], #sp-appBar, [data-automation-id="sp-appBar"], div[class^="appBar_"], div[class*="sp-appBar"] { display: none !important; visibility: hidden !important; height: 0 !important; min-height: 0 !important; max-height: 0 !important; margin: 0 !important; padding: 0 !important; overflow: hidden !important; opacity: 0 !important; pointer-events: none !important; }
            #workbenchPageContent, #spPageCanvasContent, .SPCanvas-canvas, .CanvasZone, .CanvasSection, .ControlZone, div[data-automation-id="CanvasZone"] > div { margin-left: 0 !important; padding-left: 0 !important; max-width: 100% !important; width: 100% !important; }
            body { overflow-x: hidden !important; }
          `}</style>
        )}

        <div className={styles.mobileHeaderBar}>
          <button className={styles.hamburgerBtn} onClick={() => this.setState({ isMobileMenuOpen: true })}>☰ Menu Grunner</button>
        </div>

        {this.state.isMobileMenuOpen && (
          <div className={styles.mobileOverlayBackdrop} onClick={() => this.setState({ isMobileMenuOpen: false })} />
        )}

        <aside className={`${styles.sidebar} ${this.state.isMobileMenuOpen ? styles.open : ''}`}>
          <button className={styles.closeMenuBtn} onClick={() => this.setState({ isMobileMenuOpen: false })}>✕</button>
          <div className={styles.logoArea}>
            <img src={logoGrunner} alt="Logo Semente" className={styles.logoSemente} />
            <h2>Intranet Grunner</h2>
          </div>
          <div className={styles.navGroup}>
            <h3>Navegação</h3>
            <a href="#" className={styles.active}>🏠 Painel Inicial</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/centraldeatalhos.aspx?env=Embedded">🖥️ Central de Atalhos</a>
          </div>
          <div className={styles.navGroup}>
            <h3>Serviços e Chamados</h3>
              <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/GerenciamentoDeAtivos.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">💻 Gestão de Ativos (TI)</a>
              <a href="#" onClick={(e) => { e.preventDefault(); this.setState({ isChamadoModalOpen: true }); }}>🖥️ Chamado de TI</a>
              <a href="https://grunnerteccombr.sharepoint.com/sites/Marketing/_layouts/15/listforms.aspx?cid=MTQ1MjlmMzEtNjk2Ni00MTI2LWJhNzItMzE1MTc0NDU2YTE4&nav=MGIwZDdiNzMtODQwNi00MDhiLTk5ZDEtNGE5NWNlYzljNDg3" target="_blank" rel="noopener noreferrer" data-interception="off">📢 Marketing</a>
              <a href="https://grunnerteccombr.sharepoint.com/sites/GPS/_layouts/15/listforms.aspx?cid=ZWFlMDE1MWUtOTFlMS00MmJiLWFiNzEtOWM0NGVkZTVkMTdh&nav=ZGJmNmMxZGMtNjU5Zi00ZTUxLThjMTctZmFhODY5YTQ3NjBi" target="_blank" rel="noopener noreferrer" data-interception="off">🚗 Frotas</a>
              <a href="https://forms.monday.com/forms/2a2a29caa20e7e1517cc397586af97eb?r=use1" target="_blank" rel="noopener noreferrer">🛠️ Facilities</a>
          </div>
          <div className={styles.navGroup}>
            <h3>Institucional</h3>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">🏛️ Nossa História</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">📖 Políticas da Empresa</a>
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
                <p>Bem-vindo à Intranet Grunner • O seu ecossistema agro e tecnológico</p>
                <span className={styles.dateBadge}>📅 {dataAtual.charAt(0).toUpperCase() + dataAtual.slice(1)}</span>
              </div>
            </div>
            <img src={logoCompleta} className={styles.logoCentral} alt="Grunner" />
          </header>

          <main className={styles.grid}>
            <section className={styles.newsSection}>
              {/* NOTÍCIA DE DESTAQUE */}
              {noticiaDestaque && (
                <div 
                  className={styles.heroBanner}
                  style={this.state.expandedNoticiaId === noticiaDestaque.ID ? { marginBottom: 0, borderRadius: '20px 20px 0 0' } : {}}
                >
                  <div className={styles.heroImage} style={{ backgroundImage: `url('${this.getImagemNoticia(noticiaDestaque)}')` }} />
                  <div className={styles.heroOverlay}>
                    <span className={styles.badge}>Destaque Operacional</span>
                    <h2 className={styles.heroTitle}>{noticiaDestaque.Title}</h2>
                    <p className={styles.heroResumo}>{noticiaDestaque.Resumo}</p>

                    <div className={styles.interactions}>
                      <button
                        className={styles.actionBtn}
                        onClick={(e) => { e.stopPropagation(); this.handleLike(noticiaDestaque.ID); }}
                        title={this.getTextQuemCurtiu(noticiaDestaque.ID)}
                      >
                        {this.userAlreadyLiked(noticiaDestaque.ID) ? '❤️' : '🤍'} {this.getLikesCount(noticiaDestaque.ID)} Curtidas
                      </button>

                      <button
                        className={styles.actionBtn}
                        onClick={(e) => { e.stopPropagation(); this.openCommentModal(noticiaDestaque.ID); }}
                      >
                        💬 {this.getCommentsCount(noticiaDestaque.ID)} Comentários
                      </button>

                      <button
                        className={styles.readMoreBtn}
                        onClick={() => this.handleReadMore(noticiaDestaque)}
                      >
                        {this.noticiaTemConteudo(noticiaDestaque)
                          ? this.state.expandedNoticiaId === noticiaDestaque.ID
                            ? '✕ Fechar Matéria'
                            : 'Ler Matéria ➔'
                          : 'Abrir Link ➔'}
                      </button>
                    </div>
                  </div>
                </div>
              )}

              {this.renderExpandedMainNews(noticiaDestaque)}

              {/* LISTA DE NOTÍCIAS MENORES */}
              <div className={styles.subNewsGrid}>
                {outrasNoticias.map((noticia, i) => {
                  const isExpanded = this.state.expandedNoticiaId === noticia.ID && this.noticiaTemConteudo(noticia);

                  return (
                    <div key={i} style={isExpanded ? { gridColumn: '1 / -1' } : undefined}>
                      {isExpanded ? (
                        this.renderExpandedSubNewsCard(noticia)
                      ) : (
                        <div className={styles.cardNewsSmall} style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
                          <div
                            className={styles.smallNewsImg}
                            style={{ backgroundImage: `url('${this.getImagemNoticia(noticia)}')` }}
                            onClick={() => this.noticiaTemConteudo(noticia) ? this.handleReadMore(noticia) : window.open(noticia.LinkNoticia, '_blank')}
                          />

                          <div className={styles.smallNewsContent} style={{ display: 'flex', flexDirection: 'column', flexGrow: 1, padding: '24px' }}>
                            <h3 
                              style={{ margin: '0 0 15px 0', cursor: 'pointer', lineHeight: 1.4 }} 
                              onClick={() => this.noticiaTemConteudo(noticia) ? this.handleReadMore(noticia) : window.open(noticia.LinkNoticia, '_blank')}
                            >
                              {noticia.Title}
                            </h3>

                            <div className={styles.smallInteractions} style={{ display: 'flex', gap: '15px', marginTop: 'auto', paddingTop: '15px', borderTop: '1px solid #F3F4F6', fontSize: '14px', marginBottom: '15px' }}>
                              <span
                                style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
                                onClick={(e) => { e.stopPropagation(); this.handleLike(noticia.ID); }}
                                title={this.getTextQuemCurtiu(noticia.ID)}
                              >
                                {this.userAlreadyLiked(noticia.ID) ? '❤️' : '🤍'} <small>{this.getLikesCount(noticia.ID)}</small>
                              </span>

                              <span 
                                style={{ cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
                                onClick={(e) => { e.stopPropagation(); this.openCommentModal(noticia.ID); }}
                              >
                                💬 <small>{this.getCommentsCount(noticia.ID)}</small>
                              </span>
                            </div>

                            <div>
                              <button
                                onClick={() => this.handleReadMore(noticia)}
                                style={{ width: '100%', backgroundColor: '#2E5C31', color: 'white', border: 'none', padding: '12px', borderRadius: '8px', fontWeight: 'bold', fontSize: '14px', cursor: 'pointer' }}
                              >
                                {this.noticiaTemConteudo(noticia) ? 'Ler Matéria ➔' : 'Abrir Link ➔'}
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>

              {/* BOTÃO CARREGAR MAIS */}
              {this.state.noticiasReais.length >= this.state.limiteNoticias && (
                <div style={{ display: 'flex', justifyContent: 'center', marginTop: '30px', width: '100%' }}>
                  <button className={styles.btnSecondaryOutline} onClick={this.carregarMaisNoticias} style={{ maxWidth: '300px' }}>
                    Carregar mais notícias ↓
                  </button>
                </div>
              )}

            </section>

            <aside className={styles.widgetsSection}>
              <div className={styles.card}>
                <h2>Datas importantes</h2>
                <div className={styles.eventList}>
                  {this.state.eventosReais.length > 0 ? this.state.eventosReais.map((evento, i) => {
                    const urlImagem = evento.ImagemTema ? (evento.ImagemTema.Url || evento.ImagemTema) : null;
                    const estiloDoQuadrado = urlImagem
                      ? { backgroundImage: `linear-gradient(rgba(255, 255, 255, 0.40), rgba(255, 255, 255, 0.40)), url('${urlImagem}')`, backgroundSize: 'cover', backgroundPosition: 'center' }
                      : {};

                    return (
                      <div key={i} className={styles.eventItem}>
                        <div className={styles.eventDate} style={estiloDoQuadrado}>
                          <span className={styles.eventDay}>{evento.Dia}</span>
                          <span className={styles.eventMonth}>{evento.Mes}</span>
                        </div>
                        <div className={styles.eventInfo}>
                          <div className={styles.eventTitle}>{evento.Title}</div>
                          <div className={styles.eventLocal}>📍 {evento.Local}</div>
                        </div>
                      </div>
                    );
                  }) : <p>Nenhum evento agendado.</p>}
                </div>
              </div>

              <div className={styles.card}>
                <h2>Aniversariantes do mês</h2>
                
                {(() => {
                  if (this.state.aniversariantesReais.length === 0) {
                    return <p>Nenhum aniversariante neste mês.</p>;
                  }

                  const aniversariantesOrdenados = [...this.state.aniversariantesReais].sort((a, b) => {
                    return parseInt(a.Dia, 10) - parseInt(b.Dia, 10);
                  });

                  const aniversariantesDaSemana = aniversariantesOrdenados.filter(n => this.isAniversarianteDaSemana(n.Dia));
                  const aniversariantesRestantes = aniversariantesOrdenados.filter(n => !this.isAniversarianteDaSemana(n.Dia));

                  const diaDeHoje = new Date().getDate();

                  return (
                    <>
                      <div className={styles.teamList}>
                        
                        {aniversariantesDaSemana.length > 0 ? (
                          aniversariantesDaSemana.map((niver, i) => {
                            const isHoje = parseInt(niver.Dia, 10) === diaDeHoje;

                            return (
                              <div key={`semana-${i}`} className={styles.teamItem} style={{
                                backgroundColor: isHoje ? '#ffffff' : '#F4FAEB',
                                border: isHoje ? '2px solid #A6CE39' : '1px solid #A6CE39', 
                                padding: '12px 16px',
                                borderRadius: '12px',
                                boxShadow: isHoje ? '0 6px 16px rgba(166, 206, 57, 0.4)' : '0 4px 12px rgba(166, 206, 57, 0.2)', 
                                borderBottom: isHoje ? '2px solid #A6CE39' : 'none',
                                marginBottom: '16px'
                              }}>
                                {niver.Email ? (
                                  <img src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${niver.Email}`} alt={niver.Title} className={styles.teamAvatar} />
                                ) : (
                                  <div className={styles.teamAvatarPlaceholder}>🎉</div>
                                )}
                                <div className={styles.teamInfo}>
                                  <div className={styles.teamName}>{niver.Title}</div>
                                  <div className={styles.teamDetail}>{niver.Setor} • Dia {niver.Dia}</div>
                                </div>
                                
                                <div style={{ 
                                  marginLeft: 'auto', 
                                  background: isHoje ? '#A6CE39' : '#2E5C31', 
                                  color: isHoje ? '#171E0D' : '#ffffff', 
                                  padding: '6px 12px', 
                                  borderRadius: '20px', 
                                  fontSize: '11px', 
                                  fontWeight: '900', 
                                  boxShadow: isHoje ? '0 4px 10px rgba(166, 206, 57, 0.5)' : '0 2px 6px rgba(46, 92, 49, 0.4)',
                                  whiteSpace: 'nowrap' 
                                }}>
                                  {isHoje ? 'É Hoje! 🎂' : 'Nesta semana 🎈'}
                                </div>
                              </div>
                            );
                          })
                        ) : (
                          <p style={{ fontSize: '14px', color: '#6B7280', fontStyle: 'italic', marginBottom: '15px' }}>Nenhum aniversariante para esta semana.</p>
                        )}

                        {this.state.mostrarTodosAniversariantes && (
                          aniversariantesRestantes.map((niver, i) => (
                            <div key={`resto-${i}`} className={styles.teamItem}>
                              {niver.Email ? (
                                <img src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${niver.Email}`} alt={niver.Title} className={styles.teamAvatar} />
                              ) : (
                                <div className={styles.teamAvatarPlaceholder}>🎉</div>
                              )}
                              <div className={styles.teamInfo}>
                                <div className={styles.teamName}>{niver.Title}</div>
                                <div className={styles.teamDetail}>{niver.Setor} • Dia {niver.Dia}</div>
                              </div>
                            </div>
                          ))
                        )}
                      </div>

                      {aniversariantesRestantes.length > 0 && (
                        <button 
                          onClick={() => this.setState({ mostrarTodosAniversariantes: !this.state.mostrarTodosAniversariantes })}
                          style={{ 
                            width: '100%', 
                            background: 'transparent', 
                            border: '1px dashed #cbd5e1', 
                            padding: '8px', 
                            borderRadius: '8px', 
                            color: '#64748b', 
                            cursor: 'pointer', 
                            fontWeight: 'bold', 
                            fontSize: '12px', 
                            marginTop: '15px', 
                            transition: '0.2s' 
                          }}
                          onMouseOver={(e) => e.currentTarget.style.backgroundColor = '#f8fafc'}
                          onMouseOut={(e) => e.currentTarget.style.backgroundColor = 'transparent'}
                        >
                          {this.state.mostrarTodosAniversariantes ? '↑ Ocultar restantes' : `↓ Ver outros ${aniversariantesRestantes.length} do mês`}
                        </button>
                      )}
                    </>
                  );
                })()}
              </div>
            </aside>
          </main>
        </div>

        {/* MODAL DE COMENTÁRIOS */}
        {this.state.isModalOpen && (
          <div className={styles.modalOverlay}>
            <div className={styles.modalContent}>
              <header className={styles.modalHeader}>
                <h3>Comentários da Publicação</h3>
                <button className={styles.closeBtn} onClick={() => this.setState({ isModalOpen: false })}>✕</button>
              </header>
              <div className={styles.commentsList}>
                {this.state.loadingComentarios ? (
                  <p className={styles.loadingText}>Carregando conversas...</p>
                ) : this.state.comentariosDaNoticia.length > 0 ? (
                  this.state.comentariosDaNoticia.map((item, idx) => (
                    <div key={idx} className={styles.commentBubble}>
                      <strong>{item.Autor}</strong>
                      <p>{item.Comentario}</p>
                    </div>
                  ))
                ) : (
                  <p className={styles.noComments}>Ninguém comentou ainda. Seja o primeiro a puxar assunto!</p>
                )}
              </div>
              <div className={styles.newCommentArea}>
                <textarea
                  placeholder="Escreva algo para a equipe..."
                  value={this.state.novoComentario}
                  onChange={(e) => this.setState({ novoComentario: e.target.value })}
                  style={{ width: '100%', minHeight: '80px', padding: '10px', borderRadius: '8px', border: '1px solid #d1d5db' }}
                />
                <div style={{ display: 'flex', gap: '10px', marginTop: '8px', marginBottom: '12px' }}>
                  {['👍', '❤️', '👏', '🚀', '🎉', '💡', '😂', '👀'].map(emoji => (
                    <span
                      key={emoji}
                      style={{ cursor: 'pointer', fontSize: '20px', transition: 'transform 0.2s' }}
                      onClick={() => this.setState({ novoComentario: this.state.novoComentario + emoji })}
                      onMouseOver={(e) => e.currentTarget.style.transform = 'scale(1.2)'}
                      onMouseOut={(e) => e.currentTarget.style.transform = 'scale(1)'}
                      title={`Adicionar ${emoji}`}
                    >
                      {emoji}
                    </span>
                  ))}
                </div>
                <button className={styles.sendBtn} onClick={this.enviarComentario}>Enviar Comentário</button>
              </div>
            </div>
          </div>
        )}

        {/* === MODAL DO CLICKUP (CHAMADO DE TI) === */}
        {this.state.isChamadoModalOpen && (
          <div className={styles.modalOverlay}>
            <div className={styles.modalContent} style={{ width: '800px', height: '85vh', maxWidth: '95%' }}>
              <header className={styles.modalHeader}>
                <h3>🖥️ Abertura de Chamado - TI</h3>
                <button className={styles.closeBtn} onClick={() => this.setState({ isChamadoModalOpen: false })}>✕</button>
              </header>
              <div style={{ flex: 1, padding: 0, overflow: 'hidden' }}>
                <iframe 
                  src="https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5" 
                  width="100%" 
                  height="100%" 
                  style={{ border: 'none', display: 'block' }}
                  title="Formulário de Chamado de TI"
                />
              </div>
            </div>
          </div>
        )}

      </div>
    );
  }
}