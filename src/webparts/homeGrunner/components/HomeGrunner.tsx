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

  isMarketingUser: boolean;

  // === ESTADOS DA NOTIFICAÇÃO ===
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

      isMarketingUser: false,

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
    // 1º OBRIGATÓRIO: Descobre se é o Marketing antes de buscar as notícias
    await this.verificarSeMarketing();

    // 2º Busca as informações já sabendo quem é o usuário
    await Promise.all([
      this.buscarNoticias(),
      this.buscarCelebracoesDoGraph(),
      this.buscarEventos(),
      this.buscarEngajamento(),
      this.buscarChamadosEmBackground()
    ]);
    this.setState({ loading: false });
  }

/*   // TRAVA DE TESTE: Bloqueia para a empresa, libera só para a sua máquina
  private verificarSeMarketing = async () => {
    try {
      // Pega o e-mail de quem está acessando a página agora
      const emailLogado = this.props.context.pageContext.user.email.toLowerCase().trim();

      // O seu e-mail exato para liberar o teste na produção
      const meuEmailDeTeste = "malavazi.gabriel@grunnertec.com.br"; 

      // TRAVA DE SEGURANÇA: Só libera se o e-mail for exatamente o seu
      const isMarketing = (emailLogado === meuEmailDeTeste);

      this.setState({ isMarketingUser: isMarketing });
    } catch (error) {
      console.error("Erro ao verificar o e-mail do usuário:", error);
    }
  } */

    // TRAVA INTELIGENTE: Libera para você e para qualquer pessoa do Marketing
  private verificarSeMarketing = async () => {
    try {
      // Pega o e-mail de quem está acessando a página agora
      const emailLogado = this.props.context.pageContext.user.email.toLowerCase().trim();
      const meuEmailDeTeste = "malavazi.gabriel@grunnertec.com.br"; 

      // 1. Você sempre tem acesso garantido
      let isMarketing = (emailLogado === meuEmailDeTeste);

      // 2. Se não for você, vamos perguntar pro Entra ID (Azure AD) o departamento da pessoa
      if (!isMarketing) {
        const client = await this.props.context.msGraphClientFactory.getClient("3");
        
        // Busca os dados do usuário atual (me)
        const userProfile = await client.api('/me').select('department,jobTitle').get();

        const departamento = userProfile.department ? userProfile.department.toLowerCase() : "";
        const cargo = userProfile.jobTitle ? userProfile.jobTitle.toLowerCase() : "";

        // 3. Valida se a palavra "marketing" está no departamento ou cargo
        isMarketing = departamento.includes("marketing") || cargo.includes("marketing");
      }

      this.setState({ isMarketingUser: isMarketing });
    } catch (error) {
      console.error("Erro ao verificar o departamento do usuário:", error);
      // Fallback: se der qualquer erro na API, por segurança, bloqueia a visualização de rascunhos
      this.setState({ isMarketingUser: false });
    }
  }

// =======================================================================
  // FUNÇÃO DE IMPRESSÃO - BANNER FULL WIDTH COM MOLDURA (SEM CORTES)
  // =======================================================================
  private imprimirCartaz = (noticia: any): void => {
    if (!noticia) {
      console.error("Nenhuma notícia selecionada para impressão.");
      return;
    }

    const imagemURL = this.getImagemNoticia(noticia);
    const dataPublicacao = noticia.Created ? new Date(noticia.Created).toLocaleDateString('pt-BR') : new Date().toLocaleDateString('pt-BR');
    const logoPrint = logoCompleta || "";

    const conteudoCompleto = noticia.ConteudoNoticia || noticia.Resumo || '<p>Conteúdo não disponível.</p>';

    const printWindow = window.open('', '_blank', 'width=900,height=1200');
    if (!printWindow) {
      alert("Por favor, permita pop-ups para visualizar o cartaz de impressão.");
      return;
    }

    printWindow.document.write(`
      <!DOCTYPE html>
      <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
        <title>Cartaz Grunner - ${noticia.Title}</title>
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #fff; 
            color: #171E0D; 
            -webkit-print-color-adjust: exact !important; 
            print-color-adjust: exact !important; 
          }
          
          @page { size: A4 portrait; margin: 0; }
          
          .page-container {
            width: 210mm;
            margin: 0 auto;
            padding: 15mm;
            background: #fff;
            display: flex;
            flex-direction: column;
            min-height: 297mm;
          }

          .top-header {
            background-color: #2E5C31;
            border-radius: 12px;
            padding: 15px 25px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 25px;
          }

          .top-header img { 
            height: 35px; 
            width: auto; 
            filter: brightness(0) invert(1); 
          }

          .badge-oficial {
            background-color: #A6CE39;
            color: #171E0D;
            font-weight: 800;
            font-size: 13px;
            padding: 8px 18px;
            border-radius: 30px;
            text-transform: uppercase;
          }

          .title-section { margin-bottom: 20px; }
          
          .title-section h1 {
            font-size: 22pt;
            font-weight: 900;
            color: #2E5C31;
            text-transform: uppercase;
            line-height: 1.2;
            margin-bottom: 12px;
          }

          .date-location {
            font-size: 10pt;
            color: #6B7280;
            font-weight: 700;
            text-transform: uppercase;
            display: flex;
            align-items: center;
            gap: 12px;
          }

          .date-location::after {
            content: "";
            flex: 1;
            height: 1px;
            background-color: #E5E7EB;
          }

          /* O CORAÇÃO DO ENQUADRAMENTO (MOLDURA LARGA E ESTICADA) */
          .image-wrapper {
            width: 100%;
            margin-bottom: 25px;
          }

          .image-container {
            background-color: #F8FAFC; 
            padding: 10px; 
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(0, 0, 0, 0.12); 
            border: 1px solid #E5E7EB;
            width: 100%; /* Obriga o quadro cinza a usar 100% da largura da folha */
          }
          
          .image-container img {
            width: 100%; /* Obriga a imagem a esticar junto com o quadro */
            height: auto; /* Deixa a altura crescer naturalmente para não distorcer */
            border-radius: 8px; 
            display: block;
          }

          /* ÁREA DE CONTEÚDO */
          .content-area {
            font-size: 11.5pt;
            line-height: 1.6;
            color: #374151;
            margin-bottom: 30px;
          }

          .content-area p { margin-bottom: 15px; }
          .content-area ul, .content-area ol { margin-bottom: 15px; padding-left: 25px; }
          .content-area li { margin-bottom: 8px; }
          .content-area strong { color: #171E0D; }
          
          .content-area img { 
            max-width: 100%; 
            height: auto; 
            display: block; 
            margin: 10px auto; 
            border-radius: 8px; 
          }

          .footer-box {
            background-color: #F8FAFC;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            padding: 14px;
            text-align: center;
            font-size: 9pt;
            font-weight: 800;
            color: #9CA3AF;
            text-transform: uppercase;
            margin-top: auto; 
          }
        </style>
      </head>
      <body>
        <div class="page-container">
          
          <div class="top-header">
            <img src="${logoPrint}" alt="Grunner" />
            <div class="badge-oficial">Comunicado Oficial</div>
          </div>

          <div class="title-section">
            <h1>${noticia.Title}</h1>
            <div class="date-location">Lençóis Paulista, ${dataPublicacao}</div>
          </div>
          
          <div class="image-wrapper">
            <div class="image-container">
              <img src="${imagemURL}" alt="Arte da Notícia" />
            </div>
          </div>
          
          <div class="content-area">
            ${conteudoCompleto}
          </div>

          <div class="footer-box">
            Documento de Comunicação Interna - Intranet Grunner
          </div>

        </div>

        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
              setTimeout(function() { window.close(); }, 500);
            }, 300);
          }
        </script>
      </body>
      </html>
    `);

    printWindow.document.close();
  }
 
  // ==== NOVO MOTOR DE BUSCA: ENTRA ID ====
  private buscarCelebracoesDoGraph = async () => {
    try {
      const client: MSGraphClientV3 = await this.props.context.msGraphClientFactory.getClient("3");

      const response = await client.api('/users')
        .version('v1.0')
        .select('displayName,mail,jobTitle,onPremisesExtensionAttributes')
        .filter('accountEnabled eq true')
        .top(999)
        .get();

      // Congela a data de hoje sem horas para a matemática ser perfeita
      const hoje = new Date();
      hoje.setHours(0, 0, 0, 0);

      // FUNÇÃO DO RADAR: Calcula quantos dias faltam para a data
      const calcularDiasFaltantes = (dia: number, mes: number): number => {
        const anoAtual = hoje.getFullYear();
        let dataComemoracao = new Date(anoAtual, mes - 1, dia);

        // Se a data já passou este ano, a próxima será só ano que vem
        if (dataComemoracao < hoje) {
          dataComemoracao.setFullYear(anoAtual + 1);
        }

        // Converte a diferença de milissegundos para dias
        const diffTime = dataComemoracao.getTime() - hoje.getTime();
        return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      };

      const celebracoesMap = response.value.reduce((acc: any[], user: any) => {
        const attrs = user.onPremisesExtensionAttributes;

        // 1. Processa Aniversário de Vida (extensionAttribute1: DD/MM)
        if (attrs?.extensionAttribute1) {
          const [diaStr, mesStr] = attrs.extensionAttribute1.split('/');
          const diasFaltantes = calcularDiasFaltantes(parseInt(diaStr), parseInt(mesStr));

          // Pega se for hoje (0) ou até os próximos 30 dias
          if (diasFaltantes >= 0 && diasFaltantes <= 30) {
            acc.push({
              Title: user.displayName,
              Dia: diaStr,
              Mes: mesStr,
              Setor: user.jobTitle || "Grunner",
              Email: user.mail,
              Tipo: 'nascimento',
              DiasFaltantes: diasFaltantes // <-- Essa é a nossa nova arma secreta
            });
          }
        }

        // 2. Processa Tempo de Empresa (extensionAttribute10: DD/MM/YYYY)
        if (attrs?.extensionAttribute10) {
          const [diaStr, mesStr, anoStr] = attrs.extensionAttribute10.split('/');
          const diasFaltantes = calcularDiasFaltantes(parseInt(diaStr), parseInt(mesStr));

          if (diasFaltantes >= 0 && diasFaltantes <= 30) {
            // Calcula a idade de empresa baseada no ano em que a comemoração vai cair
            const anoDaCelebracao = diasFaltantes > 0 && parseInt(mesStr) < hoje.getMonth() + 1 ? hoje.getFullYear() + 1 : hoje.getFullYear();

            acc.push({
              Title: user.displayName,
              Dia: diaStr,
              Mes: mesStr,
              Setor: user.jobTitle || "Grunner",
              Email: user.mail,
              Tipo: 'empresa',
              Anos: anoDaCelebracao - parseInt(anoStr),
              DiasFaltantes: diasFaltantes
            });
          }
        }
        return acc;
      }, []);

      // Ordena do mais próximo (hoje) para o mais distante (daqui a 30 dias)
      this.setState({
        aniversariantesReais: celebracoesMap.sort((a: any, b: any) => a.DiasFaltantes - b.DiasFaltantes),
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
      // MÁGICA DE PROTEÇÃO: 
      // Se for marketing: traz TUDO (rascunhos e publicados).
      // Se não for marketing: traz apenas os 'Publicado' OU os antigos que estão em branco (null).
      const filtroStatus = this.state.isMarketingUser 
        ? "" 
        : "&$filter=(StatusNoticia eq 'Publicado' or StatusNoticia eq null)";

      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('NoticiasGrunner')/items?$select=ID,Title,Resumo,ImagemURL,VideoURL,LinkNoticia,ConteudoNoticia,StatusNoticia,Attachments,AttachmentFiles/ServerRelativeUrl&$expand=AttachmentFiles&$top=${this.state.limiteNoticias}&$orderby=Created desc${filtroStatus}`;
      
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
    const urlCurtidas = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('CurtidasGrunner')/items?$select=ID,NoticiaID,UsuarioEmail,UsuarioNome&$top=5000`;
    
    const urlComentarios = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ComentariosGrunner')/items?$select=ID,NoticiaID,Autor,Comentario,Created&$top=5000`;
    
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
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EventosGrunner')/items?$select=Title,Dia,Mes,Local,ImagemTema&$top=20&$orderby=Created desc`;
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
  const userEmail = this.props.context.pageContext.user.email.toLowerCase();
  const userName = this.props.userDisplayName;

  const likeExistente = this.state.todasCurtidas.find(
    c => c.NoticiaID === noticiaId.toString() && c.UsuarioEmail.toLowerCase() === userEmail
  );

  const totalAtualNaTela = this.getLikesCount(noticiaId);
  const novoTotalCurtidas = likeExistente ? (totalAtualNaTela - 1) : (totalAtualNaTela + 1);

  try {
    // 1. Grava no Histórico
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

    // 2. Atualiza a Notícia (MERGE)
    const urlUpdateNoticia = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('NoticiasGrunner')/items(${noticiaId})`;
    await this.props.context.spHttpClient.post(urlUpdateNoticia, SPHttpClient.configurations.v1, {
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify({
        TotalCurtidas: novoTotalCurtidas > 0 ? novoTotalCurtidas : 0
      })
    });

    // 3. Recarrega os dados na tela
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

  const noticiaId = this.state.currentNoticiaId;
  const totalAtualNaTela = this.getCommentsCount(noticiaId);
  const novoTotalComentarios = totalAtualNaTela + 1;

  try {
    // 1. Grava no Histórico
    const urlPost = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ComentariosGrunner')/items`;
    const body = JSON.stringify({
      Title: `Comentário-${noticiaId}`,
      NoticiaID: noticiaId.toString(),
      Comentario: this.state.novoComentario,
      Autor: this.props.userDisplayName
    });
    await this.props.context.spHttpClient.post(urlPost, SPHttpClient.configurations.v1, { body: body });

    // 2. Atualiza a Notícia (MERGE)
    const urlUpdateNoticia = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('NoticiasGrunner')/items(${noticiaId})`;
    await this.props.context.spHttpClient.post(urlUpdateNoticia, SPHttpClient.configurations.v1, {
      headers: {
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*'
      },
      body: JSON.stringify({
        TotalComentarios: novoTotalComentarios
      })
    });

    // 3. Atualiza a tela e limpa o modal
    this.setState({ novoComentario: "" });
    this.buscarComentarios(noticiaId);
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
      
{/* 🚀 BLOCO INTELIGENTE: Verifica se é YouTube ou Vídeo Direto (MP4) */}
      {noticia.VideoURL && (
        <div style={{ marginBottom: '30px' }}>
          {noticia.VideoURL.includes('youtube.com') || noticia.VideoURL.includes('youtu.be') ? (
            /* Renderiza iFrame se for YouTube */
            <iframe 
              width="100%" 
              height="450" 
              src={noticia.VideoURL} 
              title="Vídeo da Matéria" 
              frameBorder="0" 
              allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" 
              allowFullScreen
              style={{ borderRadius: '12px', boxShadow: '0 10px 25px rgba(0,0,0,0.1)' }}
            />
          ) : (
            /* Renderiza Player Nativo se for vídeo interno (SharePoint/MP4) */
            <video 
              width="100%" 
              controls 
              style={{ borderRadius: '12px', boxShadow: '0 10px 25px rgba(0,0,0,0.1)', backgroundColor: '#000' }}
            >
              <source src={noticia.VideoURL} type="video/mp4" />
              Seu navegador não suporta a exibição deste vídeo.
            </video>
          )}
        </div>
      )}

      {/* Texto original da matéria */}
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
            
            {/* AVISO DE RASCUNHO EXCLUSIVO DO MARKETING */}
            {noticia.StatusNoticia === 'Rascunho' && (
              <span 
                className={styles.draftBadge}
                title="Rascunho (Invisível para a empresa)"
              >
                👁️
              </span>
            )}

            <h2 className={styles.heroTitle}>{noticia.Title}</h2>

            <div className={styles.interactions}>
              
              {/* BOTÃO CURTIDA SOLTO (GHOST BUTTON) */}
              <button
                className={styles.actionIconBtn}
                onClick={(e) => { e.stopPropagation(); this.handleLike(noticia.ID); }}
                title={this.getTextQuemCurtiu(noticia.ID)}
              >
                {this.userAlreadyLiked(noticia.ID) ? '❤️' : '🤍'} 
                <span>{this.getLikesCount(noticia.ID)}</span>
              </button>

              {/* BOTÃO COMENTÁRIO SOLTO (GHOST BUTTON) */}
              <button
                className={styles.actionIconBtn}
                onClick={(e) => { e.stopPropagation(); this.openCommentModal(noticia.ID); }}
                title="Ver Comentários"
              >
                💬 <span>{this.getCommentsCount(noticia.ID)}</span>
              </button>

              {this.state.isMarketingUser && (
                <button
                  className={styles.actionBtnLight}
                  onClick={(e) => { 
                    e.stopPropagation(); 
                    this.imprimirCartaz(noticia); 
                  }}
                >
                  🖨️ Print
                </button>
              )}

              <button
                className={styles.readMoreBtn}
                style={{ backgroundColor: '#DC2626', boxShadow: '0 4px 14px rgba(220, 38, 38, 0.3)' }}
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

    const celebracoesFiltradas = this.state.aniversariantesReais
      .filter(c => this.state.filtroCelebracao === 'todos' || c.Tipo === this.state.filtroCelebracao);

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

            <a href="#" onClick={(e) => this.abrirModalFormulario("https://grunnerteccombr.sharepoint.com/sites/Marketing/_layouts/15/listforms.aspx?cid=MTQ1MjlmMzEtNjk2Ni00MTI2LWJhNzItMzE1MTc0NDU2YTE4&nav=MGIwZDdiNzMtODQwNi00MDhiLTk5ZDEtNGE5NWNlYzljNDg3&env=Embedded", "📢 Solicitação - Marketing", e)}>📢 Marketing</a>
            <a href="#" onClick={(e) => this.abrirModalFormulario("https://grunnerteccombr.sharepoint.com/sites/GPS/_layouts/15/listforms.aspx?cid=ZWFlMDE1MWUtOTFlMS00MmJiLWFiNzEtOWM0NGVkZTVkMTdh&nav=ZGJmNmMxZGMtNjU5Zi00ZTUxLThjMTctZmFhODY5YTQ3NjBi&env=Embedded", "🚗 Solicitação - Frotas", e)}>🚗 Frotas</a>
            <a href="#" onClick={(e) => this.abrirModalFormulario("https://forms.monday.com/forms/embed/2a2a29caa20e7e1517cc397586af97eb?r=use1", "🛠️ Solicitação - Facilities", e)}>🛠️ Facilities</a>
          </div>
          <div className={styles.navGroup}>
            <h3>Institucional</h3>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Historia.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">🏛️ Nossa História</a>
            <a href="https://grunnertec.com.br/assets/PDFs/codigoconduta.pdf" target="_blank" rel="noopener noreferrer">⚖️ Código de Conduta</a>
            <a href="https://grunner.canaldeouvidoria.com.br/" target="_blank" rel="noopener noreferrer">🗣️ Canal de Ética</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Pol%C3%ADticas-da-Empresa.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">📖 Procedimentos</a>
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
                    
                    {/* AVISO DE RASCUNHO */}
                    {noticiaDestaque.StatusNoticia === 'Rascunho' && (
                      <span 
                        className={styles.draftBadge}
                        title="Rascunho (Apenas o Marketing consegue ver. Mude o status para 'Publicado' no SharePoint para liberar para a empresa)."
                      >
                        👁️
                      </span>
                    )}

                    <h2 className={styles.heroTitle}>{noticiaDestaque.Title}</h2>

                    <div className={styles.interactions}>
                      
                      {/* BOTÃO CURTIDA (GHOST BUTTON) NO BANNER PRINCIPAL */}
                      <button
                        className={styles.actionIconBtn}
                        onClick={(e) => { e.stopPropagation(); this.handleLike(noticiaDestaque.ID); }}
                        title={this.getTextQuemCurtiu(noticiaDestaque.ID)}
                      >
                        {this.userAlreadyLiked(noticiaDestaque.ID) ? '❤️' : '🤍'} 
                        <span>{this.getLikesCount(noticiaDestaque.ID)}</span>
                      </button>

                      {/* BOTÃO COMENTÁRIO (GHOST BUTTON) NO BANNER PRINCIPAL */}
                      <button
                        className={styles.actionIconBtn}
                        onClick={(e) => { e.stopPropagation(); this.openCommentModal(noticiaDestaque.ID); }}
                        title="Ver Comentários"
                      >
                        💬 <span>{this.getCommentsCount(noticiaDestaque.ID)}</span>
                      </button>

                      {this.state.isMarketingUser && this.state.expandedNoticiaId === noticiaDestaque.ID && (
                        <button
                          className={styles.actionBtnLight}
                          onClick={(e) => { 
                            e.stopPropagation(); 
                            this.imprimirCartaz(noticiaDestaque); 
                          }}
                        >
                          🖨️ Print
                        </button>
                      )}

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
                            
                            {noticia.StatusNoticia === 'Rascunho' && (
                              <span 
                                className={styles.draftBadge}
                                title="Rascunho (Invisível para a empresa)"
                                style={{ alignSelf: 'flex-start', marginBottom: '10px' }}
                              >
                                👁️
                              </span>
                            )}

                            <h3
                              style={{ margin: '0 0 10px 0', cursor: 'pointer', lineHeight: 1.4 }}
                              onClick={() => this.noticiaTemConteudo(noticia) ? this.handleReadMore(noticia) : window.open(noticia.LinkNoticia, '_blank')}
                            >
                              {noticia.Title}
                            </h3>

                            {noticia.Resumo && (
                              <p style={{
                                margin: '0 0 15px 0',
                                fontSize: '13px',
                                color: '#6B7280',
                                lineHeight: 1.5,
                                display: '-webkit-box',
                                WebkitLineClamp: 3,
                                WebkitBoxOrient: 'vertical',
                                overflow: 'hidden'
                              }}>
                                {noticia.Resumo}
                              </p>
                            )}

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
                      ? {
                        backgroundImage: `linear-gradient(rgba(255, 255, 255, 0.40), rgba(255, 255, 255, 0.40)), url('${urlImagem}')`,
                        backgroundSize: 'cover',
                        backgroundPosition: 'center'
                      }
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

              <div className={`${styles.card} ${styles.celebrationsCard}`}>
                <div className={styles.celebrationsHeader}>
                  <div className={styles.celebrationsHeading}>
                    <h2 className={styles.celebrationsTitle}>🎉 Celebrações</h2>
                    <p className={styles.celebrationsSubtitle}>Aniversários e tempo de casa</p>
                  </div>

                  <div className={styles.celebrationsFilters}>
                    {(['todos', 'nascimento', 'empresa'] as const).map(f => {
                      const ativo = this.state.filtroCelebracao === f;

                      return (
                        <button
                          key={f}
                          type="button"
                          onClick={() => this.setState({ filtroCelebracao: f })}
                          title={f === 'todos' ? 'Todos' : f === 'nascimento' ? 'Aniversários' : 'Tempo de empresa'}
                          className={`${styles.celebrationFilterBtn} ${ativo ? styles.celebrationFilterBtnActive : ''}`}
                        >
                          {f === 'todos' ? 'Todos' : f === 'nascimento' ? 'Aniv.' : 'Emp.'}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className={styles.celebrationsList}>
                  {this.state.loadingCelebracoes ? (
                    <div className={styles.celebrationEmpty}>
                      Carregando celebrações...
                    </div>
                  ) : celebracoesFiltradas.length > 0 ? (
                    celebracoesFiltradas.map((niver, i) => {
                      const isHoje = niver.DiasFaltantes === 0;
                      const isEmpresa = niver.Tipo === 'empresa';

                      const badgeClass = isEmpresa
                        ? styles.celebrationBadgeEmpresa
                        : isHoje
                          ? styles.celebrationBadgeToday
                          : styles.celebrationBadgeBirthday;

                      const badgeText = isEmpresa
                        ? (niver.Anos === 0 ? 'Novo' : `${niver.Anos} ano${niver.Anos > 1 ? 's' : ''}`)
                        : (isHoje ? 'Hoje' : 'Aniv.');

                      const iniciais = String(niver.Title || '?')
                        .split(' ')
                        .filter(Boolean)
                        .slice(0, 2)
                        .map(parte => parte.charAt(0))
                        .join('')
                        .toUpperCase();

                      return (
                        <div
                          key={`${niver.Email || niver.Title}-${niver.Tipo}-${i}`}
                          className={`${styles.celebrationItem} ${isHoje ? styles.celebrationItemToday : ''}`}
                        >
                          {niver.Email ? (
                            <img
                              src={`${this.props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${niver.Email}`}
                              alt={niver.Title}
                              className={styles.celebrationAvatar}
                            />
                          ) : (
                            <div className={styles.celebrationAvatarPlaceholder}>
                              {iniciais}
                            </div>
                          )}

                          <div className={styles.celebrationInfo}>
                            <div className={styles.celebrationName}>{niver.Title}</div>
                            <div className={styles.celebrationDetail}>
                              <span>{niver.Setor || 'Grunner'}</span>
                              <span className={styles.celebrationDetailDot}>•</span>
                              <span>{`Dia ${niver.Dia}/${niver.Mes}`}</span>
                            </div>
                          </div>

                          <div className={`${styles.celebrationBadge} ${badgeClass}`}>
                            {badgeText}
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className={styles.celebrationEmpty}>
                      Nenhuma celebração para este filtro.
                    </div>
                  )}
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
        
        {/* MODAL UNIVERSAL DE FORMULÁRIOS EXTERNOS */}
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