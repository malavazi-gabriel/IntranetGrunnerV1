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
  ProcessoExtinto?: boolean;

  // CAMPOS PARA O FLUXO DE RASCUNHO E APROVAÇÃO
  StatusDaRevisao?: string;
  AprovadoresEmailsText?: string;
  AprovadoresdoDocumento?: any[];
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
  visaoAtual: 'oficiais' | 'rascunhos' | 'obsoletos';
  documentosRascunho: any[];
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
      selectedNewDocType: '',
      visaoAtual: 'oficiais', // Inicia na visão oficial
      documentosRascunho: []
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
      const select = 'Id,UniqueId,FileLeafRef,FileRef,Area,CodigoDocumento,TipoDocumento,NumeroRevisao,DataUltimaRevisao,DataProximaRevisao,StatusDocumento,ObservacaoRevisao,PeriodicidadeRevisaoMeses,UltimoAvisoRevisao,DiasAvisoRevisao,PermiteImpressaoControlada,ExibirNaIntranet,ResponsavelRevisao/Title,ResponsavelRevisao/EMail,AprovadorQualidade/Title,AprovadorQualidade/EMail,TipoProcessoDocumento,DocumentoControlado,AprovadoresdoDocumento/Title,AprovadoresdoDocumento/EMail,ProcessoExtinto';
      const expand = 'ResponsavelRevisao,AprovadorQualidade,AprovadoresdoDocumento';
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

  private buscarRascunhos = async (): Promise<void> => {
    this.setState({ loading: true });
    try {
      // Adicionada a coluna Avaliador no Select
      const select = 'Id,FileLeafRef,StatusdaRevisao,AprovadoresdoDocumento/EMail,MotivoRejeicao,Avaliador';
      const expand = 'AprovadoresdoDocumento';
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('RascunhosSGQ')/items?$select=${select}&$expand=${expand}`;

      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      this.setState({ documentosRascunho: data.value || [], loading: false });
    } catch (error) {
      console.error("Erro ao buscar rascunhos:", error);
      this.setState({ loading: false });
    }
  }

  private buscarHistoricoDocumento = async (itemId: number): Promise<void> => {
    this.setState({ isLoadingHistory: true, documentHistory: [] });

    try {
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PoliticasGrunner')/items(${itemId})/versions`;

      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`Erro na requisição: ${response.statusText}`);
      }

      const data = await response.json();

      if (data && data.value) {
        const historicoFormatado: IDocumentVersion[] = data.value.map((versao: any) => {

          let nomeEditor = 'Sistema';
          if (versao.Editor) {
            nomeEditor = versao.Editor.LookupValue || versao.Editor.Title || versao.Editor.Email || 'Sistema';
          }

          // A MÁGICA ACONTECE AQUI: Ele tenta puxar a sua ObservacaoRevisao primeiro.
          const observacaoFinal = versao.ObservacaoRevisao || versao.CheckInComment || '';

          return {
            VersionLabel: versao.VersionLabel,
            Modified: versao.Created,
            Editor: nomeEditor,
            CheckInComment: observacaoFinal
          };
        });

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

  private gerarCodigoAutomatico = (): void => {
    const { editFormData, todosDocumentos } = this.state;

    // O sistema precisa saber a Área e o Tipo para montar o código
    if (!editFormData.TipoProcessoDocumento || !editFormData.Area) {
      alert("⚠️ Selecione a 'Área' e o 'Tipo de Processo/Documento' primeiro para gerar o código.");
      return;
    }

    // Dicionários extraídos da sua planilha oficial
    const siglaTipo: { [key: string]: string } = {
      'DESENHO TECNICO': 'DET',
      'ESPECIFICACAO TECNICA': 'EST',
      'FORMULARIO': 'FOR',
      'INSTRUCAO DE TRABALHO': 'IT',
      'MANUAL': 'MAN',
      'MAPEAMENTO DE PROCESSO': 'MAP',
      'POLITICA': 'POL',
      'PROCEDIMENTO OPERACIONAL PADRAO': 'POP',
      'PROCEDIMENTO': 'PRO'
    };

    const siglaArea: { [key: string]: string } = {
      'ASSISTENCIA TECNICA': 'AST', 'ATENDIMENTO': 'ATD', 'COMERCIAL': 'COM',
      'COMPLIANCE': 'CPL', 'CONTROLADORIA': 'CTL', 'DEPARTAMENTO PESSOAL': 'DP',
      'ESCRITORIO DE GERENCIAMENTO DE PROJETOS': 'EGP', 'ENGENHARIA DE PROCESSO': 'ENG',
      'FACILITIES': 'FAC', 'FINANCEIRO': 'FIN', 'FISCAL': 'FIS', 'FROTA LEVE': 'FLE',
      'JURIDICO': 'JUR', 'LOGISTICA': 'LOG', 'MEIO AMBIENTE': 'MA', 'MARKETING': 'MKT',
      'PLANEJAMENTO, PROGRAMACAO E CONTROLE DA PRODUCAO': 'PCP', 'PESQUISA E DESENVOLVIMENTO': 'PED',
      'PERFORMANCE': 'PER', 'PRODUCAO': 'PRD', 'QUALIDADE': 'QUA', 'RECURSOS HUMANOS': 'RH',
      'SUCESSO DO CLIENTE': 'SDC', 'SEGURANCA DO TRABALHO': 'SEG',
      'SUPRIMENTOS - COMPRAS / MATERIAIS': 'SUP', 'TECNOLOGIA DA INFORMACAO': 'TI',
      'USINAGEM': 'USI', 'VENDA DE PECAS': 'VPE',
      // Variações comuns para garantir a identificação
      'RH': 'RH', 'TI': 'TI', 'SISTEMAS': 'TI', 'COMPRAS': 'SUP', 'INSTITUCIONAL': 'INST'
    };

    // Função interna para remover acentos para não dar falha no Match
    const removerAcentos = (str: string) => str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    const tipo = removerAcentos(editFormData.TipoProcessoDocumento.toUpperCase().trim());
    const area = removerAcentos(editFormData.Area.toUpperCase().trim());

    // Se por acaso inventarem uma área nova que não tá na lista, ele pega as 3 primeiras letras
    const prefixoTipo = siglaTipo[tipo] || 'DOC';
    const prefixoArea = siglaArea[area] || area.substring(0, 3).toUpperCase();

    const baseCode = `${prefixoTipo}.${prefixoArea}.`;

    // Rastreia o maior número já usado para não duplicar
    let maiorNumero = 0;
    todosDocumentos.forEach(doc => {
      if (doc.CodigoDocumento && doc.CodigoDocumento.startsWith(baseCode)) {
        const partes = doc.CodigoDocumento.split('.');
        if (partes.length >= 3) {
          const numFinal = parseInt(partes[partes.length - 1], 10);
          if (!isNaN(numFinal) && numFinal > maiorNumero) {
            maiorNumero = numFinal;
          }
        }
      }
    });

    const proximoNumero = (maiorNumero + 1).toString().padStart(3, '0'); // Garante os 3 dígitos (ex: 001)

    this.setState({
      editFormData: {
        ...editFormData,
        CodigoDocumento: `${baseCode}${proximoNumero}`
      }
    });
  }

  private salvarEdicaoDocumento = async (): Promise<void> => {
    const { documentoSelecionado, editFormData, visaoAtual } = this.state;
    if (!documentoSelecionado) return;

    this.setState({ salvandoDocumento: true });

    const listaDestino = visaoAtual === 'rascunhos' ? 'RascunhosSGQ' : 'PoliticasGrunner';
    const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listaDestino}')/items(${documentoSelecionado.Id})`;

    try {
      let payload: any = {};

      // 1. SE FOR RASCUNHO: Envia APENAS os campos que existem na lista de Rascunhos
      if (visaoAtual === 'rascunhos') {
        payload.StatusdaRevisao = editFormData.StatusDaRevisao || 'Rascunho';

        if (editFormData.AprovadoresEmailsText && editFormData.AprovadoresEmailsText.trim().length > 0) {
          const emails = editFormData.AprovadoresEmailsText.split(/[,;]/).map(e => e.trim()).filter(e => e.length > 0);

          const arrayDeIdsDosAprovadores = await Promise.all(
            emails.map(async (email: string) => {
              return await this.getUserIdByEmail(email);
            })
          );

          // Correção do Array: Em 'odata=nometadata', envia-se a lista diretamente
          payload.AprovadoresdoDocumentoId = arrayDeIdsDosAprovadores.filter(id => id !== null);
        } else {
          payload.AprovadoresdoDocumentoId = [];
        }
      }
      // 2. SE FOR OFICIAL: Envia todos os metadados completos
      else {
        payload = {
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
          Area: editFormData.Area || null,
          ProcessoExtinto: editFormData.ProcessoExtinto !== undefined ? editFormData.ProcessoExtinto : false
        };

        if (editFormData.ResponsavelRevisao?.EMail) {
          payload.ResponsavelRevisaoId = await this.getUserIdByEmail(editFormData.ResponsavelRevisao.EMail);
        } else if (editFormData.ResponsavelRevisao?.EMail === '') {
          payload.ResponsavelRevisaoId = -1;
        }

        if (editFormData.AprovadorQualidade?.EMail) {
          payload.AprovadorQualidadeId = await this.getUserIdByEmail(editFormData.AprovadorQualidade.EMail);
        } else if (editFormData.AprovadorQualidade?.EMail === '') {
          payload.AprovadorQualidadeId = -1;
        }
      }

      const response = await this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      });

      // Se o SharePoint rejeitar por algum motivo, forçamos um erro para ver no console
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText);
      }

      this.setState({
        documentoSelecionado: null,
        salvandoDocumento: false,
        activeModalTab: 'metadados',
        documentHistory: []
      });

      if (visaoAtual === 'rascunhos') {
        this.buscarRascunhos();
      } else {
        this.buscarTodosDocumentos();
      }

    } catch (error) {
      console.error("Erro detalhado ao salvar no SharePoint:", error);
      alert("Erro ao salvar! Pressione F12 e veja a aba 'Console' para ver o motivo.");
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
    // Cabeçalho limpo, alinhado com a interface do sistema
    const cabecalho = [
      'Código',
      'Nome do Documento',
      'Área',
      'Tipo',
      'Controlado?',
      'Última Revisão',
      'Vencimento',
      'Status',
      'Responsável',
      'Aprovadores',
      'Link de Acesso'
    ];

    const linhas = documentos.map(doc => {
      // Extrai os Aprovadores
      const aprovadoresStr = doc.AprovadoresdoDocumento && doc.AprovadoresdoDocumento.length > 0
        ? doc.AprovadoresdoDocumento.map((ap: any) => ap.Title || ap.EMail).join(', ')
        : '-';

      // Extrai o Responsável
      const responsavelStr = doc.ResponsavelRevisao ? (doc.ResponsavelRevisao.Title || doc.ResponsavelRevisao.EMail || '-') : '-';

      // Cria o link direto
      const linkDireto = doc.FileRef ? `https://grunnerteccombr.sharepoint.com${doc.FileRef}` : '-';

      return [
        doc.CodigoDocumento || '-',
        doc.FileLeafRef || '-',
        doc.Area || '-',
        doc.TipoProcessoDocumento || '-',
        doc.DocumentoControlado ? 'Sim' : 'Não',
        this.formatDate(doc.DataUltimaRevisao),
        this.formatDate(doc.DataProximaRevisao),
        doc.StatusCalculado || '-',
        responsavelStr,
        aprovadoresStr,
        linkDireto
      ];
    });

    const conteudoCSV = [cabecalho, ...linhas].map(e => e.join(';')).join('\n');

    const blob = new Blob(["\ufeff", conteudoCSV], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'Relatorio_Documentos_SGQ.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }

  public render(): React.ReactElement<IPoliticasGrunnerProps> {
    const { areaAtiva, todosDocumentos, termoBusca, loading, isQualidadeUser, modoGestaoQualidade, visaoAtual, documentosRascunho } = this.state;

    // Métricas (Sempre conta da base oficial)
    const total = todosDocumentos.length;
    const vigentes = todosDocumentos.filter(d => d.StatusCalculado === 'Vigente').length;
    const atencao = todosDocumentos.filter(d => d.StatusCalculado === 'Vence em breve').length;
    const vencidos = todosDocumentos.filter(d => d.StatusCalculado === 'Vencido').length;
    const revisao = todosDocumentos.filter(d => d.StatusCalculado === 'Em revisão').length;

    // Filtros de exibição da visão oficial
    let documentosExibidos = todosDocumentos;

    if (!modoGestaoQualidade) {
      // Regra Pública
      documentosExibidos = documentosExibidos.filter(d =>
        d.StatusCalculado !== 'Arquivado' &&
        d.StatusCalculado !== 'Vencido' &&
        d.StatusCalculado !== 'Em revisão' &&
        d.StatusCalculado !== 'Obsoleto' &&
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
      // Gestão de Qualidade
      if (this.state.filtroStatusAdmin !== 'Todos' && visaoAtual === 'oficiais') {
        documentosExibidos = documentosExibidos.filter(d => d.StatusCalculado === this.state.filtroStatusAdmin);
      }
      if (termoBusca.trim().length > 0) {
        documentosExibidos = documentosExibidos.filter(doc =>
          doc.FileLeafRef?.toLowerCase().includes(termoBusca.toLowerCase()) ||
          doc.CodigoDocumento?.toLowerCase().includes(termoBusca.toLowerCase())
        );
      }
    }

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

            <a
              className={`${styles.menuToggle} ${this.state.isMenuTIOpen ? styles.active : ''}`}
              onClick={(e) => { e.preventDefault(); this.setState({ isMenuTIOpen: !this.state.isMenuTIOpen }); }}
            >
              <span style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>💻 Tecnologia (TI)</span>
              <span style={{ fontSize: '10px', opacity: 0.8 }}>{this.state.isMenuTIOpen ? '▲' : '▼'}</span>
            </a>

            {this.state.isMenuTIOpen && (
              <div className={styles.navSubGroup}>
                <a href="https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/GerenciamentoDeAtivos.aspx?env=Embedded" target="_blank" rel="noopener noreferrer">🖥️ Gestão de Ativos</a>
                <a href="https://forms.clickup.com/9007063382/f/8cdtrap-43393/OCRETZOXI4CU88XQA5" target="_blank" rel="noopener noreferrer">➕ Abrir Novo Chamado</a>
                <a href="#" onClick={(e) => { e.preventDefault(); window.dispatchEvent(new CustomEvent('abrirMeusChamadosGrunner', { detail: 'TI' })); }}>🎫 Meus Chamados</a>
              </div>
            )}

            <a href="https://grunnerteccombr.sharepoint.com/sites/Marketing/_layouts/15/listforms.aspx?cid=MTQ1MjlmMzEtNjk2Ni00MTI2LWJhNzItMzE1MTc0NDU2YTE4&nav=MGIwZDdiNzMtODQwNi00MDhiLTk5ZDEtNGE5NWNlYzljNDg3" target="_blank" rel="noopener noreferrer" data-interception="off">📢 Marketing</a>
            <a href="https://grunnerteccombr.sharepoint.com/sites/GPS/_layouts/15/listforms.aspx?cid=ZWFlMDE1MWUtOTFlMS00MmJiLWFiNzEtOWM0NGVkZTVkMTdh&nav=ZGJmNmMxZGMtNjU5Zi00ZTUxLThjMTctZmFhODY5YTQ3NjBi" target="_blank" rel="noopener noreferrer" data-interception="off">🚗 Frotas</a>
            <a href="https://grunnerteccombr.sharepoint.com/:l:/s/Facilities/JADJeN1a-IAVRIrzsns79wBEAS_s9zB21POwKXunqjUuK5Y?nav=MDk0ODE1N2QtZWE0Ny00ZDhjLWFhYjItMGVlNmIwMWIzNTY4" target="_blank" rel="noopener noreferrer">🛠️ Facilities</a>
          </div>

          <div className={styles.navGroup}>
            <h3>Institucional</h3>
            <a href={historiaUrl} target="_blank" rel="noopener noreferrer">🏛️ Nossa História</a>
            <a href="https://grunnertec.com.br/assets/PDFs/codigoconduta.pdf" target="_blank" rel="noopener noreferrer">⚖️ Código de Conduta</a>
            <a href="https://grunner.canaldeouvidoria.com.br/" target="_blank" rel="noopener noreferrer">🗣️ Canal de Ética</a>

            {this.state.isQualidadeUser ? (
              <>
                <a
                  className={`${styles.menuToggle} ${this.state.isMenuProcedimentosOpen ? styles.active : ''}`}
                  onClick={(e) => { e.preventDefault(); this.setState({ isMenuProcedimentosOpen: !this.state.isMenuProcedimentosOpen }); }}
                >
                  <span style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>📖 Procedimentos</span>
                  <span style={{ fontSize: '10px', opacity: 0.8 }}>{this.state.isMenuProcedimentosOpen ? '▲' : '▼'}</span>
                </a>

                {this.state.isMenuProcedimentosOpen && (
                  <div className={styles.navSubGroup}>
                    <a href={politicasUrl} className={!this.state.modoGestaoQualidade ? styles.active : ''}>
                      📖 Todos os Documentos
                    </a>
                    <a
                      href="#"
                      className={this.state.modoGestaoQualidade ? styles.active : ''}
                      onClick={(e) => { e.preventDefault(); this.setState({ modoGestaoQualidade: true, visaoAtual: 'oficiais' }); }}
                    >
                      ⚙️ Gestão da Qualidade
                    </a>
                  </div>
                )}
              </>
            ) : (
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
              <p>{modoGestaoQualidade ? 'Controle de revisões, rascunhos pendentes e auditoria.' : 'Acesse os documentos normativos, manuais e procedimentos de cada área da empresa.'}</p>
            </div>
          </header>

          {/* PAINEL DE MÉTRICAS */}
          {modoGestaoQualidade && visaoAtual === 'oficiais' && (
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

              {/* NOVO CARTÃO DE RASCUNHOS */}
              <div
                className={`${styles.metricCard} ${modoGestaoQualidade ? styles.clickableCard : ''}`}
                style={{ borderTop: '4px solid #8B5CF6' }}
                onClick={() => {
                  this.setState({ visaoAtual: 'rascunhos' });
                  if (this.state.documentosRascunho.length === 0) this.buscarRascunhos();
                }}
              >
                <div className={styles.metricLabel}>Rascunhos</div>
                <div className={styles.metricValue}>{this.state.documentosRascunho ? this.state.documentosRascunho.length : 0}</div>
              </div>

            </div>
          )}

          {/* NAVEGAÇÃO ENTRE OFICIAIS, RASCUNHOS E OBSOLETOS (Apenas Qualidade) */}
          {modoGestaoQualidade && (
            <div className={styles.tabsContainer} style={{ marginBottom: '20px' }}>
              <button
                className={visaoAtual === 'oficiais' ? styles.tabActive : styles.tab}
                onClick={() => this.setState({ visaoAtual: 'oficiais' })}>
                📂 Documentos Oficiais
              </button>
              <button
                className={visaoAtual === 'rascunhos' ? styles.tabActive : styles.tab}
                onClick={() => {
                  this.setState({ visaoAtual: 'rascunhos' });
                  if (documentosRascunho.length === 0) this.buscarRascunhos();
                }}>
                📥 Caixa de Entrada (Rascunhos)
              </button>
              <button
                className={visaoAtual === 'obsoletos' ? styles.tabActive : styles.tab}
                onClick={() => this.setState({ visaoAtual: 'obsoletos' })}>
                🗄️ Arquivo Morto (Obsoletos)
              </button>
            </div>
          )}

          {(visaoAtual === 'oficiais' || visaoAtual === 'obsoletos') && (
            <>
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
            </>
          )}

          {modoGestaoQualidade && visaoAtual === 'oficiais' && (
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
            ) : visaoAtual === 'rascunhos' ? (

              // ================= VISÃO DE RASCUNHOS (CAIXA DE ENTRADA) =================
              <div className={styles.adminTableWrapper}>
                <table className={styles.adminTable}>
                  <thead>
                    <tr>
                      <th>Nome do Rascunho</th>
                      <th>Status da Revisão</th>
                      <th>Feedback da Avaliação</th> {/* Título mais claro */}
                      <th>Gestores Atribuídos</th>
                      <th>Ações</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documentosRascunho.length === 0 ? (
                      <tr><td colSpan={5} style={{ textAlign: 'center' }}>Nenhum rascunho pendente no momento.</td></tr>
                    ) : (
                      documentosRascunho.map((rasc, idx) => {
                        const emailsStr = rasc.AprovadoresdoDocumento
                          ? rasc.AprovadoresdoDocumento.map((ap: any) => ap.EMail).join('; ')
                          : '';

                        // Lógica de cores para o Status
                        let badgeColor = '#E5E7EB';
                        let textColor = '#374151';

                        if (rasc.StatusdaRevisao === 'Aguardando Gestores') {
                          badgeColor = '#FEF08A';
                          textColor = '#854D0E';
                        } else if (rasc.StatusdaRevisao === 'Rejeitado') {
                          badgeColor = '#FECACA';
                          textColor = '#991B1B';
                        } else if (rasc.StatusdaRevisao === 'Aprovado') {
                          badgeColor = '#DEF7EC';
                          textColor = '#03543F';
                        }

                        return (
                          <tr key={idx}>
                            <td style={{ fontWeight: '500' }}>{rasc.FileLeafRef}</td>

                            <td>
                              <span className={styles.statusBadge} style={{ backgroundColor: badgeColor, color: textColor, fontWeight: 'bold' }}>
                                {rasc.StatusdaRevisao || 'Rascunho'}
                              </span>
                            </td>

                            {/* Coluna Centralizada de Feedback com o Nome de quem avaliou */}
                            <td style={{ maxWidth: '280px', whiteSpace: 'normal', lineHeight: '1.4' }}>
                              {rasc.StatusdaRevisao === 'Rejeitado' && rasc.MotivoRejeicao ? (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                  <span style={{ color: '#991B1B', fontSize: '13px', fontWeight: '600' }}>
                                    ⚠️ {rasc.MotivoRejeicao}
                                  </span>
                                  {rasc.Avaliador && (
                                    <span style={{ fontSize: '11px', color: '#6B7280' }}>👤 Rejeitado por: <strong>{rasc.Avaliador}</strong></span>
                                  )}
                                </div>
                              ) : rasc.StatusdaRevisao === 'Aprovado' && rasc.Avaliador ? (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                  <span style={{ color: '#03543F', fontSize: '13px', fontWeight: '600' }}>
                                    ✅ Documento aceito
                                  </span>
                                  <span style={{ fontSize: '11px', color: '#6B7280' }}>👤 Aprovado por: <strong>{rasc.Avaliador}</strong></span>
                                </div>
                              ) : (
                                <span style={{ color: '#9CA3AF' }}>-</span>
                              )}
                            </td>

                            {/* Coluna de Gestores mantida para saber quem mais está envolvido */}
                            <td style={{ fontSize: '12px', color: '#4B5563' }}>{emailsStr || '-'}</td>

                            <td className={styles.adminActions}>
                              <button onClick={() => this.setState({
                                documentoSelecionado: rasc,
                                editFormData: {
                                  StatusDaRevisao: rasc.StatusdaRevisao || 'Rascunho',
                                  AprovadoresEmailsText: emailsStr
                                }
                              })} className={styles.editButton}>✏️ Gerenciar Aprovação</button>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>

            ) : visaoAtual === 'obsoletos' ? (

              // ================= VISÃO DE ARQUIVO MORTO / OBSOLETOS =================
              <div className={styles.adminTableWrapper}>
                <table className={styles.adminTable}>
                  <thead>
                    <tr>
                      <th>Código</th>
                      <th>Nome do Documento</th>
                      <th>Área</th>
                      <th>Tipo</th>
                      <th>Observação / Destino</th>
                      <th>Ações</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documentosExibidos.filter(d => d.StatusCalculado === 'Obsoleto' || d.StatusCalculado === 'Arquivado').length === 0 ? (
                      <tr><td colSpan={6} style={{ textAlign: 'center' }}>Nenhum documento obsoleto ou arquivado no momento.</td></tr>
                    ) : (
                      documentosExibidos.filter(d => d.StatusCalculado === 'Obsoleto' || d.StatusCalculado === 'Arquivado').map((doc, idx) => (
                        <tr key={idx} style={{ backgroundColor: '#F9FAFB', opacity: 0.85 }}>
                          <td style={{ color: '#6B7280' }}>{doc.CodigoDocumento || '-'}</td>
                          <td style={{ textDecoration: 'line-through', fontWeight: '500', color: '#4B5563' }}>{doc.FileLeafRef}</td>
                          <td style={{ color: '#6B7280' }}>{doc.Area}</td>
                          <td style={{ color: '#6B7280' }}>{doc.TipoProcessoDocumento || doc.TipoDocumento || '-'}</td>
                          <td style={{ maxWidth: '300px', whiteSpace: 'normal', color: '#991B1B', fontSize: '12px', lineHeight: '1.4' }}>
                            {doc.ObservacaoRevisao ? `📌 ${doc.ObservacaoRevisao}` : <span style={{ color: '#9CA3AF' }}>Sem destino registrado</span>}
                          </td>
                          <td className={styles.adminActions}>
                            <button onClick={() => this.setState({ documentoSelecionado: doc, editFormData: { ...doc, ResponsavelRevisao: { EMail: doc.ResponsavelRevisao?.EMail }, AprovadorQualidade: { EMail: doc.AprovadorQualidade?.EMail } } })} className={styles.editButton}>🔍 Consultar</button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>

            ) : !modoGestaoQualidade ? (

              // ================= VISÃO PÚBLICA (Cards) =================
              <div className={styles.documentGrid}>
                {documentosExibidos.map((doc, index) => {
                  const extensao = doc.FileLeafRef ? doc.FileLeafRef.split('.').pop()?.toLowerCase() : '';
                  const isPdf = extensao === 'pdf';
                  return (
                    <div key={index} className={styles.documentCard}>
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

                      <div className={styles.cardBody}>
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

                      <div className={styles.cardFooter}>
                        <div className={styles.revisionInfo}>
                          <span className={styles.revText}>Rev. {doc.NumeroRevisao || '00'}</span>
                          <span className={styles.venceText}>Vence: {this.formatDate(doc.DataProximaRevisao)}</span>
                        </div>
                        <a
                          onClick={(e) => {
                            e.preventDefault();
                            const siteUrl = this.props.context.pageContext.web.absoluteUrl;
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
              // ================= VISÃO ADMINISTRATIVA OFICIAL (Tabela) =================
              <div className={styles.adminTableWrapper}>
                <table className={styles.adminTable}>
                  <thead>
                    <tr>
                      <th>Código</th>
                      <th>Nome</th>
                      <th>Área</th>
                      <th>Rev</th>
                      <th>Última Revisão</th> {/* Logo após a Rev */}
                      <th>Status</th>
                      <th>Vencimento</th>      {/* Movido para perto do status */}
                      <th>Responsável</th>
                      <th>Aprovadores</th>
                      <th>Ações</th>
                    </tr>
                  </thead>
                  <tbody>
                    {documentosExibidos.map((doc, idx) => {
                      const aprovadoresOficiaisStr = doc.AprovadoresdoDocumento && doc.AprovadoresdoDocumento.length > 0
                        ? doc.AprovadoresdoDocumento.map((ap: any) => ap.Title || ap.EMail).join('; ')
                        : '-';

                      return (
                        <tr key={idx}>
                          <td>{doc.CodigoDocumento || '-'}</td>
                          <td>{doc.FileLeafRef}</td>
                          <td>{doc.Area}</td>
                          <td>{doc.NumeroRevisao || '-'}</td>
                          <td>{this.formatDate(doc.DataUltimaRevisao)}</td> {/* Última Revisão */}
                          <td><span className={`${styles.statusBadge} ${this.getStatusClass(doc.StatusCalculado)}`}>{doc.StatusCalculado}</span></td>
                          <td>{this.formatDate(doc.DataProximaRevisao)}</td>     {/* Vencimento */}
                          <td>{doc.ResponsavelRevisao?.Title || '-'}</td>
                          <td style={{ fontSize: '12px', color: '#4B5563' }}>{aprovadoresOficiaisStr}</td>
                          <td className={styles.adminActions}>
                            <button onClick={() => this.setState({ documentoSelecionado: doc, editFormData: { ...doc, ResponsavelRevisao: { EMail: doc.ResponsavelRevisao?.EMail }, AprovadorQualidade: { EMail: doc.AprovadorQualidade?.EMail } } })} className={styles.editButton}>✏️ Editar</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </main>
        </div>

        {/* ================= MODAL DE EDIÇÃO (DINÂMICO PARA OFICIAL OU RASCUNHO) ================= */}
        {this.state.documentoSelecionado && (
          <div className={styles.editModalBackdrop}>
            <div className={styles.editModal}>

              <div className={styles.editModalHeader}>
                <h2>{this.state.documentoSelecionado.FileLeafRef}</h2>
                <button onClick={() => this.setState({ documentoSelecionado: null, activeModalTab: 'metadados', documentHistory: [] })} className={styles.closeModal}>✕</button>
              </div>

              {/* SÓ MOSTRA ABAS SE FOR DOCUMENTO OFICIAL OU OBSOLETO */}
              {(visaoAtual === 'oficiais' || visaoAtual === 'obsoletos') && (
                <div className={styles.modalTabs}>
                  <button
                    className={`${styles.modalTab} ${this.state.activeModalTab === 'metadados' ? styles.modalTabActive : ''}`}
                    onClick={() => this.setState({ activeModalTab: 'metadados' })}
                  >
                    {visaoAtual === 'obsoletos' ? '🔍 Consultar Metadados' : '📝 Editar Metadados'}
                  </button>
                  <button
                    className={`${styles.modalTab} ${this.state.activeModalTab === 'historico' ? styles.modalTabActive : ''}`}
                    onClick={() => {
                      this.setState({ activeModalTab: 'historico' });
                      if (this.state.documentoSelecionado && (!this.state.documentHistory || this.state.documentHistory.length === 0)) {
                        this.buscarHistoricoDocumento(this.state.documentoSelecionado.Id);
                      }
                    }}
                  >
                    🕒 Histórico de Revisões
                  </button>
                </div>
              )}

              <div className={styles.editModalBody}>

                {/* CONTEÚDO PARA RASCUNHOS */}
                {visaoAtual === 'rascunhos' && (
                  <div className={styles.formGrid}>
                    <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                      <label>Status da Revisão</label>
                      <select
                        value={this.state.editFormData.StatusDaRevisao || 'Rascunho'}
                        onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, StatusDaRevisao: e.target.value } })}
                      >
                        <option value="Rascunho">Rascunho</option>
                        <option value="Aguardando Gestores">Aguardando Gestores</option>
                        <option value="Aprovado">Aprovado</option>
                        <option value="Rejeitado">Rejeitado</option>
                      </select>
                    </div>

                    <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                      <label>E-mails dos Aprovadores (separe por vírgula)</label>
                      <textarea
                        rows={3}
                        placeholder="gestor1@grunner.com.br, gestor2@grunner.com.br"
                        value={this.state.editFormData.AprovadoresEmailsText || ''}
                        onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, AprovadoresEmailsText: e.target.value } })}
                        style={{ padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px', fontFamily: 'inherit', resize: 'vertical' }}
                      />
                    </div>
                  </div>
                )}

                {/* CONTEÚDO PARA METADADOS OFICIAIS E OBSOLETOS */}
                {(visaoAtual === 'oficiais' || visaoAtual === 'obsoletos') && this.state.activeModalTab === 'metadados' && (
                  <>
                    {visaoAtual === 'oficiais' ? (
                      // ================== TELA DE EDIÇÃO PADRÃO (DOCUMENTOS VIGENTES) ==================
                      <div className={styles.formGrid}>
                        <div className={styles.formGroup}>
                          <label>Área Responsável</label>
                          <select value={this.state.editFormData.Area || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, Area: e.target.value } })}>
                            <option value="">Selecione...</option>
                            <option value="Assistência técnica">Assistência técnica</option>
                            <option value="Atendimento">Atendimento</option>
                            <option value="Comercial">Comercial</option>
                            <option value="Compliance">Compliance</option>
                            <option value="Controladoria">Controladoria</option>
                            <option value="Departamento pessoal">Departamento pessoal</option>
                            <option value="Escritório de gerenciamento de projetos">Escritório de gerenciamento de projetos</option>
                            <option value="Engenharia de processo">Engenharia de processo</option>
                            <option value="Facilities">Facilities</option>
                            <option value="Financeiro">Financeiro</option>
                            <option value="Fiscal">Fiscal</option>
                            <option value="Frota leve">Frota leve</option>
                            <option value="Jurídico">Jurídico</option>
                            <option value="Logística">Logística</option>
                            <option value="Meio Ambiente">Meio Ambiente</option>
                            <option value="Marketing">Marketing</option>
                            <option value="Planejamento, programação e controle da produção">Planejamento, programação e controle da produção</option>
                            <option value="Pesquisa e Desenvolvimento">Pesquisa e Desenvolvimento</option>
                            <option value="Performance">Performance</option>
                            <option value="Produção">Produção</option>
                            <option value="Qualidade">Qualidade</option>
                            <option value="Recursos humanos">Recursos humanos</option>
                            <option value="Sucesso do cliente">Sucesso do cliente</option>
                            <option value="Segurança do trabalho">Segurança do trabalho</option>
                            <option value="Suprimentos - compras / materiais">Suprimentos - compras / materiais</option>
                            <option value="Tecnologia da Informação">Tecnologia da Informação</option>
                            <option value="Usinagem">Usinagem</option>
                            <option value="Venda de peças">Venda de peças</option>
                          </select>
                        </div>

                        <div className={styles.formGroup}>
                          <label>Tipo de Processo/Documento</label>
                          <select value={this.state.editFormData.TipoProcessoDocumento || ''} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, TipoProcessoDocumento: e.target.value } })}>
                            <option value="">Selecione...</option>
                            <option value="DESENHO TECNICO">Desenho Técnico</option>
                            <option value="ESPECIFICACAO TECNICA">Especificação Técnica</option>
                            <option value="FORMULARIO">Formulário</option>
                            <option value="INSTRUCAO DE TRABALHO">Instrução de Trabalho</option>
                            <option value="MANUAL">Manual</option>
                            <option value="MAPEAMENTO DE PROCESSO">Mapeamento de Processo</option>
                            <option value="POLITICA">Política</option>
                            <option value="PROCEDIMENTO OPERACIONAL PADRAO">Procedimento Operacional Padrão (POP)</option>
                            <option value="PROCEDIMENTO">Procedimento</option>
                          </select>
                        </div>

                        <div className={styles.formGroup} style={{ gridColumn: '1 / -1' }}>
                          <label>Código do Documento</label>
                          <div style={{ display: 'flex', gap: '10px' }}>
                            <input
                              type="text"
                              style={{ flex: 1, padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px' }}
                              value={this.state.editFormData.CodigoDocumento || ''}
                              onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, CodigoDocumento: e.target.value } })}
                            />
                            <button
                              type="button"
                              onClick={this.gerarCodigoAutomatico}
                              style={{ backgroundColor: '#A6CE39', color: '#1C2510', border: 'none', padding: '0 15px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold' }}>
                              🪄 Gerar Código
                            </button>
                          </div>
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
                          <label>Documento Controlado?</label>
                          <select value={this.state.editFormData.DocumentoControlado ? 'sim' : 'nao'} onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, DocumentoControlado: e.target.value === 'sim' } })}>
                            <option value="nao">Não - Não Controlado</option>
                            <option value="sim">Sim - Controlado</option>
                          </select>
                        </div>

                        <div className={styles.formGroup}>
                          <label>Status</label>
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

                        <div className={styles.formGroup} style={{ gridColumn: '1 / -1' }}>
                          <label>Observações internas (Uso da Qualidade)</label>
                          <textarea
                            rows={3}
                            placeholder="Digite lembretes, motivos de alteração ou notas para a próxima revisão..."
                            value={this.state.editFormData.ObservacaoRevisao || ''}
                            onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, ObservacaoRevisao: e.target.value } })}
                            style={{ padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px', fontFamily: 'inherit', resize: 'vertical', width: '100%', boxSizing: 'border-box' }}
                          />
                        </div>
                      </div>

                    ) : (
                      // ================== TELA ESPECÍFICA PARA DOCUMENTOS OBSOLETOS ==================
                      <div className={styles.formGrid}>

                        <div style={{ gridColumn: '1 / -1', backgroundColor: '#F3F4F6', padding: '15px', borderRadius: '8px', borderLeft: '4px solid #4B5563', marginBottom: '5px' }}>
                          <h3 style={{ margin: '0 0 10px 0', color: '#374151', fontSize: '16px' }}>🗄️ Painel de Documento Obsoleto</h3>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px', fontSize: '13px', color: '#4B5563' }}>
                            <p style={{ margin: 0 }}><strong>Código:</strong> {this.state.editFormData.CodigoDocumento || '-'}</p>
                            <p style={{ margin: 0 }}><strong>Área:</strong> {this.state.editFormData.Area || '-'}</p>
                            <p style={{ margin: 0 }}><strong>Última Revisão:</strong> {this.state.editFormData.NumeroRevisao || '-'}</p>
                            <p style={{ margin: 0 }}><strong>Tipo:</strong> {this.state.editFormData.TipoProcessoDocumento || '-'}</p>
                          </div>
                        </div>

                        <div className={styles.formGroup} style={{ gridColumn: '1 / -1' }}>
                          <label style={{ fontWeight: 'bold' }}>Status Atual do Arquivo</label>
                          <select
                            value={this.state.editFormData.StatusDocumento || 'Obsoleto'}
                            onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, StatusDocumento: e.target.value } })}
                            style={{ padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px' }}
                          >
                            <option value="Obsoleto">Obsoleto</option>
                            <option value="Arquivado">Arquivado</option>
                            <option value="Em revisão">🔄 Restaurar para Revisão (Voltar à vida)</option>
                          </select>
                        </div>

                        {/* --- NOVA FLAG DE EXTINÇÃO (Fundo Amarelo) --- */}
                        <div className={styles.formGroup} style={{ gridColumn: '1 / -1', backgroundColor: '#FFFBEB', padding: '10px', borderRadius: '6px', border: '1px solid #FDE68A' }}>
                          <label style={{ fontWeight: 'bold', color: '#92400E' }}>O processo atrelado a este documento foi extinto da empresa?</label>
                          <select
                            value={this.state.editFormData.ProcessoExtinto ? 'sim' : 'nao'}
                            onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, ProcessoExtinto: e.target.value === 'sim' } })}
                            style={{ padding: '8px', border: '1px solid #D1D5DB', borderRadius: '4px', width: '100%', marginTop: '5px' }}
                          >
                            <option value="nao">Não - O processo continua (Foi substituído por outro documento)</option>
                            <option value="sim">Sim - O processo foi 100% extinto / descontinuado</option>
                          </select>
                        </div>

                        <div className={styles.formGroup} style={{ gridColumn: '1 / -1' }}>
                          <label style={{ fontWeight: 'bold', color: '#991B1B' }}>Destino / Motivo da Obsolescência</label>
                          <textarea
                            rows={5}
                            placeholder="Ex: Documento descontinuado. O processo foi agrupado e substituído pelo documento POP.ATD.005..."
                            value={this.state.editFormData.ObservacaoRevisao || ''}
                            onChange={(e) => this.setState({ editFormData: { ...this.state.editFormData, ObservacaoRevisao: e.target.value } })}
                            style={{ padding: '12px', border: '1px solid #FCA5A5', borderRadius: '6px', fontFamily: 'inherit', resize: 'vertical', width: '100%', boxSizing: 'border-box', backgroundColor: '#FEF2F2' }}
                          />
                        </div>

                      </div>
                    )}
                  </>
                )}
                
                {/* CONTEÚDO PARA HISTÓRICO OFICIAL E OBSOLETOS */}
                {(visaoAtual === 'oficiais' || visaoAtual === 'obsoletos') && this.state.activeModalTab === 'historico' && (
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

              {/* FOOTER DO MODAL (BOTÕES) CORRIGIDO */}
              <div className={styles.editModalFooter}>
                <button className={styles.cancelBtn} onClick={() => this.setState({ documentoSelecionado: null, activeModalTab: 'metadados', documentHistory: [] })}>Fechar</button>

                {/* O Botão de Salvar aparece para tudo, mas com textos dinâmicos! */}
                {(visaoAtual === 'rascunhos' || this.state.activeModalTab === 'metadados') && (
                  <button className={styles.saveBtn} onClick={this.salvarEdicaoDocumento} disabled={this.state.salvandoDocumento}>
                    {this.state.salvandoDocumento ? 'Salvando...' :
                      (visaoAtual === 'rascunhos' ? 'Atualizar Rascunho' :
                        (visaoAtual === 'obsoletos' ? 'Salvar Destino' : 'Salvar Alterações'))}
                  </button>
                )}
              </div>

              {/* FOOTER DO MODAL (BOTÕES) CORRIGIDO */}
              <div className={styles.editModalFooter}>
                <button className={styles.cancelBtn} onClick={() => this.setState({ documentoSelecionado: null, activeModalTab: 'metadados', documentHistory: [] })}>Fechar</button>

                {/* O Botão de Salvar agora só aparece para Rascunhos ou Oficiais, NUNCA para Obsoletos */}
                {(visaoAtual === 'rascunhos' || (visaoAtual === 'oficiais' && this.state.activeModalTab === 'metadados')) && (
                  <button className={styles.saveBtn} onClick={this.salvarEdicaoDocumento} disabled={this.state.salvandoDocumento}>
                    {this.state.salvandoDocumento ? 'Salvando...' : (visaoAtual === 'rascunhos' ? 'Atualizar Rascunho' : 'Salvar Alterações')}
                  </button>
                )}
              </div>

            </div>
          </div>
        )}

        {/* MODAL DO IFRAME E RESTO DO CÓDIGO (CRIAR DOCUMENTO) MANTIDO... */}
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

        {/* MODAL DE CRIAÇÃO MANTIDO */}
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
                  onClick={() => this.setState({ isCreateModalOpen: false })}
                  style={{ opacity: !this.state.selectedNewDocType ? 0.5 : 1, cursor: !this.state.selectedNewDocType ? 'not-allowed' : 'pointer' }}
                >
                  Continuar para o Formulário ➔
                </button>
              </div>
            </div>
          </div>
        )}

        {/* FORMULÁRIOS MANTIDOS */}
        {!this.state.isCreateModalOpen && this.state.selectedNewDocType === 'MAPEAMENTO DE PROCESSO' && (
          <FormularioMapeamento tipoDocumento={this.state.selectedNewDocType} usuarioEmail={this.props.context.pageContext.user.email} spContext={this.props.context} onFechar={() => this.setState({ selectedNewDocType: '' })} onSucesso={() => { this.setState({ selectedNewDocType: '' }); this.buscarTodosDocumentos(); }} />
        )}

        {!this.state.isCreateModalOpen && this.state.selectedNewDocType === 'PROCEDIMENTO' && (
          <FormularioProcedimento tipoDocumento={this.state.selectedNewDocType} usuarioEmail={this.props.context.pageContext.user.email} spContext={this.props.context} onFechar={() => this.setState({ selectedNewDocType: '' })} onSucesso={() => { this.setState({ selectedNewDocType: '' }); this.buscarTodosDocumentos(); }} />
        )}

        {!this.state.isCreateModalOpen && this.state.selectedNewDocType === 'INSTRUÇÃO DE TRABALHO' && (
          <FormularioSGQ tipoDocumento={this.state.selectedNewDocType} usuarioEmail={this.props.context.pageContext.user.email} spContext={this.props.context} onFechar={() => this.setState({ selectedNewDocType: '' })} onSucesso={() => { this.setState({ selectedNewDocType: '' }); this.buscarTodosDocumentos(); }} />
        )}

        {!this.state.isCreateModalOpen && this.state.selectedNewDocType && !['MAPEAMENTO DE PROCESSO', 'PROCEDIMENTO', 'INSTRUÇÃO DE TRABALHO'].includes(this.state.selectedNewDocType) && (
          <div className={styles.editModalBackdrop}>
            <div className={styles.editModal} style={{ width: '90%', maxWidth: '450px', textAlign: 'center', padding: '40px 30px', borderTop: '6px solid #A6CE39' }}>
              <div style={{ fontSize: '50px', marginBottom: '15px' }}>🚀</div>
              <h2 style={{ color: '#1C2510', margin: '0 0 15px 0', fontSize: '22px', fontWeight: '800' }}>
                Módulo em Desenvolvimento
              </h2>
              <p style={{ color: '#6B7280', fontSize: '15px', lineHeight: '1.6', marginBottom: '30px' }}>
                O gerador automatizado para <strong>{this.state.selectedNewDocType}</strong> está sendo construído pela nossa equipe de Tecnologia para garantir a melhor experiência possível.
                <br /><br />
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