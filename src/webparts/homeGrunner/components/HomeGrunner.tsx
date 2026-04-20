import * as React from 'react';
import styles from './HomeGrunner.module.scss';
import { IHomeGrunnerProps } from './IHomeGrunnerProps';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { MenuChamados } from '../../../shared/components/MenuChamado/MenuChamados';
import { MSGraphClientV3 } from '@microsoft/sp-http';

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
  
  isTiMenuOpen: boolean;
  isMeusChamadosModalOpen: boolean;
  meusChamados: any[];
  loadingChamados: boolean;
  
  expandedTicketIndex: number | null;
  novoComentarioChamado: string;
  enviandoComentarioChamado: boolean;
  
  comentariosDoChamado: any[];
  loadingHistorico: boolean;

  // === ESTADOS DA NOTIFICAÇÃO (NOVOS) ===
  unreadTicketsCount: number;
  isNotificacaoOpen: boolean;
  
  // AS 3 VARIÁVEIS NOVAS DO IFRAME 
  isIframeModalOpen: boolean;
  iframeUrl: string;
  iframeTitle: string;

  filtroCelebracao: 'todos' | 'nascimento' | 'empresa';
  loadingCelebracoes: boolean;
  
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
      mostrarTodosAniversariantes: false,
      
      isTiMenuOpen: false,
      isMeusChamadosModalOpen: false,
      meusChamados: [],
      loadingChamados: false,
      expandedTicketIndex: null,
      novoComentarioChamado: "",
      enviandoComentarioChamado: false,

      comentariosDoChamado: [],
      loadingHistorico: false,

      // INICIALIZANDO AS NOTIFICAÇÕES
      unreadTicketsCount: 0,
      isNotificacaoOpen: false,

      // INICIALIZANDO O IFRAME
      isIframeModalOpen: false,
      iframeUrl: '',
      iframeTitle: '',

      filtroCelebracao: 'todos',
      loadingCelebracoes: true
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

  private abrirModalFormulario = (url: string, titulo: string, e: React.MouseEvent) => {
    e.preventDefault(); 
    this.setState({ 
      isIframeModalOpen: true, 
      iframeUrl: url, 
      iframeTitle: titulo 
    });
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
      // this.buscarAniversariantes()
      this.buscarCelebracoesDoGraph(),
      this.buscarEventos(),
      this.buscarEngajamento(),
      this.buscarChamadosEmBackground()
    ]);
    this.setState({ loading: false });
  }

  // ==== NOVO MOTOR DE BUSCA: ENTRA ID ====
  private buscarCelebracoesDoGraph = async () => {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");
      
      const response = await client.api('/users')
        .version('v1.0')
        .select('displayName,mail,jobTitle,onPremisesExtensionAttributes')
        .filter('accountEnabled eq true')
        .get();

      const hoje = new Date();
      const mesAtual = hoje.getMonth() + 1;

      const celebracoesMap = response.value.reduce((acc: any[], user: any) => {
        const attrs = user.onPremisesExtensionAttributes;
        
        // 1. Processa Aniversário de Vida (extensionAttribute1: DD/MM)
        if (attrs?.extensionAttribute1) {
          const [dia, mes] = attrs.extensionAttribute1.split('/');
          if (parseInt(mes) === mesAtual) {
            acc.push({
              Title: user.displayName,
              Dia: dia,
              Setor: user.jobTitle || "Grunner",
              Email: user.mail,
              Tipo: 'nascimento'
            });
          }
        }

        // 2. Processa Tempo de Empresa (extensionAttribute10: DD/MM/YYYY)
        if (attrs?.extensionAttribute10) {
          const [dia, mes, ano] = attrs.extensionAttribute10.split('/');
          if (parseInt(mes) === mesAtual) {
            acc.push({
              Title: user.displayName,
              Dia: dia,
              Setor: user.jobTitle || "Grunner",
              Email: user.mail,
              Tipo: 'empresa',
              Anos: hoje.getFullYear() - parseInt(ano)
            });
          }
        }
        return acc;
      }, []);

      this.setState({ 
        aniversariantesReais: celebracoesMap.sort((a: any, b: any) => parseInt(a.Dia) - parseInt(b.Dia)),
        loadingCelebracoes: false 
      });

    } catch (error) {
      console.error("Erro ao buscar dados do Entra ID:", error);
      this.setState({ loadingCelebracoes: false });
    }
  }

  // ==== NOVA FUNÇÃO: BUSCAR CHAMADOS SILENCIOSAMENTE PARA O BANNER ====
  private buscarChamadosEmBackground = async () => {
    const rawEmail = this.props.context.pageContext.user.email || "";
    const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/meus-chamados?email=${rawEmail.toLowerCase().trim()}`;

    try {
      const response = await fetch(apiUrl);
      const data = await response.json();
      
      if (data.sucesso && Array.isArray(data.chamados)) {
        this.setState({ meusChamados: data.chamados }, this.recalcularNotificacoes);
      }
    } catch (error) {
      console.error("Erro ao buscar chamados no background", error);
    }
  }

  // ==== FUNÇÃO: RECALCULAR A MATEMÁTICA DO SININHO ====
  private recalcularNotificacoes = () => {
    let unreadCount = 0;
    this.state.meusChamados.forEach((ticket: any) => {
      const lastSeen = localStorage.getItem(`grunner_visto_${ticket.id}`);
      const isEscondido = localStorage.getItem(`grunner_escondido_${ticket.id}`) === "true";
      const isEncerrado = ticket.status.toLowerCase().includes('encerrado') || ticket.status.toLowerCase().includes('conclu');
      
      if (isEscondido && isEncerrado) return; // Se escondeu e tá fechado, ignora
      
      const dataClickUp = parseInt(ticket.dataAtualizacao || '0');
      const dataLida = parseInt(lastSeen || '0');
      
      if (dataClickUp > dataLida) {
        unreadCount++;
      }
    });

    this.setState({ unreadTicketsCount: unreadCount });
  }

  private abrirModalMeusChamados = async () => {
    this.setState({ 
      isMeusChamadosModalOpen: true, 
      isNotificacaoOpen: false,
      loadingChamados: true, 
      meusChamados: [],
      expandedTicketIndex: null,
      novoComentarioChamado: "",
      comentariosDoChamado: []
    });
    
    const rawEmail = this.props.context.pageContext.user.email || "";
    const userEmail = rawEmail.toLowerCase().trim();

    const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/meus-chamados?email=${userEmail}`;

    try {
      const response = await fetch(apiUrl);
      const data = await response.json();
      
      this.setState({ 
        meusChamados: data.sucesso && Array.isArray(data.chamados) ? data.chamados : [], 
        loadingChamados: false 
      });
    } catch (error) {
      this.setState({ loadingChamados: false, meusChamados: [] });
    }
  }

  // ==== FUNÇÃO ATUALIZADA: ABRIR DETALHES E MARCAR COMO LIDO ====
  private toggleDetalhesChamado = async (index: number, idChamado: string) => {
    const ticket = this.state.meusChamados[index];
    
    if (this.state.expandedTicketIndex === index) {
      this.setState({ expandedTicketIndex: null, comentariosDoChamado: [] });
      return;
    }

    // Salva a data da última atualização vista no navegador (para apagar a bolinha vermelha)
    if (ticket.dataAtualizacao) {
      localStorage.setItem(`grunner_visto_${idChamado}`, ticket.dataAtualizacao);
    }

    this.setState({ 
      expandedTicketIndex: index, 
      loadingHistorico: true,
      comentariosDoChamado: []
    }, this.recalcularNotificacoes);

    this.carregarHistoricoDoChamado(idChamado);
  }

  // ==== NOVA FUNÇÃO: OCULTAR CHAMADO ENCERRADO ====
  private dispensarChamado = (idChamado: string) => {
    if (window.confirm("Deseja ocultar este chamado da sua lista?")) {
      localStorage.setItem(`grunner_escondido_${idChamado}`, "true");
      this.setState({ expandedTicketIndex: null }, this.recalcularNotificacoes);
      this.setState({ expandedTicketIndex: null }); // Fecha a sanfona
      this.forceUpdate(); // Força a tela a desenhar de novo para o chamado sumir
    }
  }

  private carregarHistoricoDoChamado = async (idChamado: string) => {
    try {
      const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/comentarios?idChamado=${idChamado}`;
      const response = await fetch(apiUrl);
      const data = await response.json();

      if (data.sucesso) {
        this.setState({ comentariosDoChamado: data.comentarios, loadingHistorico: false });
      } else {
        this.setState({ loadingHistorico: false });
      }
    } catch (error) {
      console.error("Erro ao carregar chat:", error);
      this.setState({ loadingHistorico: false });
    }
  }

  private enviarComentarioChamado = async (idChamado: string) => {
    if (!this.state.novoComentarioChamado.trim()) return;

    this.setState({ enviandoComentarioChamado: true });
    
    const rawEmail = this.props.context.pageContext.user.email || "";
    const userEmail = rawEmail.toLowerCase().trim();
    const apiUrl = `https://bw4oogog00scckw0wgo08cww.82.25.70.48.sslip.io/api/clickup/comentar`;

    try {
      const response = await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          idChamado: idChamado,
          comentario: this.state.novoComentarioChamado,
          email: userEmail
        })
      });

      const result = await response.json();

      if (result.sucesso) {
        this.setState({ novoComentarioChamado: "", enviandoComentarioChamado: false });
        this.carregarHistoricoDoChamado(idChamado);
      } else {
        alert("Ocorreu um erro ao enviar: " + result.mensagem);
        this.setState({ enviandoComentarioChamado: false });
      }
    } catch (error) {
      alert("Erro de comunicação com o servidor.");
      this.setState({ enviandoComentarioChamado: false });
    }
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
              
              <div className={styles.accordionGroup}>
                <button
                  className={`${styles.accordionToggle} ${this.state.isTiMenuOpen ? styles.open : ''}`}
                  onClick={() => this.setState({ isTiMenuOpen: !this.state.isTiMenuOpen })}
                >
                  <span style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>💻 Tecnologia (TI)</span>
                  <span className={styles.chevron}>▼</span>
                </button>
                
                {this.state.isTiMenuOpen && (
                  <div className={styles.accordionContent}>
                    <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/GerenciamentoDeAtivos.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">🖥️ Gestão de Ativos</a>
                    <a href="#" onClick={(e) => this.abrirModalFormulario("https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5", "➕ Abrir Novo Chamado", e)}>➕ Abrir Novo Chamado</a>
                    <a href="#" onClick={(e) => { e.preventDefault(); window.dispatchEvent(new CustomEvent('abrirMeusChamadosGrunner', { detail: 'TI' })); }}>🎫 Meus Chamados</a>
                  </div>
                )}
              </div>

              {/* RESTANTE DOS DEPARTAMENTOS A USAR O MODAL */}
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
            {/* NOSSO NOVO COMPONENTE COMPARTILHADO */}
              <MenuChamados 
                departamento="TI" 
                emailUsuario={userEmail} 
              />
            <div className={styles.headerRight}>
               <img src={logoCompleta} className={styles.logoCentral} alt="Grunner" />
            </div>
          </header>

          <main className={styles.grid}>
            <section className={styles.newsSection}>
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
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
                  <h2 style={{ margin: 0 }}>Celebrações do Mês</h2>
                  
                  {/* Botões de Filtro de UX */}
                  <div style={{ display: 'flex', gap: '4px' }}>
                    {(['todos', 'nascimento', 'empresa'] as const).map(f => (
                      <button
                        key={f}
                        onClick={() => this.setState({ filtroCelebracao: f })}
                        style={{
                          padding: '4px 8px',
                          borderRadius: '12px',
                          fontSize: '10px',
                          fontWeight: 'bold',
                          cursor: 'pointer',
                          border: '1px solid #2E5C31',
                          backgroundColor: this.state.filtroCelebracao === f ? '#2E5C31' : 'transparent',
                          color: this.state.filtroCelebracao === f ? '#fff' : '#2E5C31',
                          transition: '0.2s'
                        }}
                      >
                        {f === 'todos' ? 'Todos' : f === 'nascimento' ? 'Bday' : 'Casa'}
                      </button>
                    ))}
                  </div>
                </div>

                <div className={styles.teamList}>
                  {this.state.aniversariantesReais
                    .filter(c => this.state.filtroCelebracao === 'todos' || c.Tipo === this.state.filtroCelebracao)
                    .map((niver, i) => (
                      <div key={i} className={styles.teamItem} style={{ borderLeft: niver.Tipo === 'empresa' ? '4px solid #2E5C31' : 'none', paddingLeft: '8px' }}>
                        {niver.Email ? (
                          <img 
                            src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${niver.Email}`} 
                            className={styles.teamAvatar} 
                          />
                        ) : (
                          <div className={styles.teamAvatarPlaceholder}>🎉</div>
                        )}
                        <div className={styles.teamInfo}>
                          <div className={styles.teamName}>{niver.Title}</div>
                          <div className={styles.teamDetail}>{niver.Setor} • Dia {niver.Dia}</div>
                        </div>
                        
                        <div style={{ 
                          marginLeft: 'auto', 
                          background: niver.Tipo === 'empresa' ? '#2E5C31' : '#A6CE39', 
                          color: niver.Tipo === 'empresa' ? '#ffffff' : '#171E0D', 
                          padding: '4px 10px', 
                          borderRadius: '20px', 
                          fontSize: '10px', 
                          fontWeight: '900' 
                        }}>
                          {niver.Tipo === 'empresa' ? `${niver.Anos} Anos 🚜` : 'Bday 🎂'}
                        </div>
                      </div>
                  ))}
                </div>
              </div>
            </aside>
          </main>
        </div>

        {/* MODAL DE COMENTÁRIOS DE NOTÍCIAS */}
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
{/* ==============================================
            MODAL UNIVERSAL DE FORMULÁRIOS EXTERNOS
 ============================================== */}
        {this.state.isIframeModalOpen && (
          <div className={styles.modalOverlay}>
            <div className={styles.modalContent} style={{ width: '900px', height: '85vh', maxWidth: '95%', display: 'flex', flexDirection: 'column' }}>
              <header className={styles.modalHeader}>
                <h3>{this.state.iframeTitle}</h3>
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