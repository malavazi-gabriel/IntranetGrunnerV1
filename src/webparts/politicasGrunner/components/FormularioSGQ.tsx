import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { SPHttpClient } from '@microsoft/sp-http'; // Adicionado para as requisições HTTP

interface IFormularioSGQProps {
  tipoDocumento: string;
  usuarioEmail: string;
  spContext: any; // Contexto do SharePoint passado pelo componente pai
  onFechar: () => void;
  onSucesso: () => void;
}

interface IFormularioSGQState {
  nomeProcesso: string;
  nomeDocumento: string;
  elaborador: string;
  objetivo: string;
  aplicacao: string;
  docReferencia: string;
  definicoes: string;
  papeisResponsabilidades: string;
  preRequisitos: string;
  sequenciaExecutiva: string;
  registrosComplementares: string;
  enviando: boolean;
}

export default class FormularioSGQ extends React.Component<IFormularioSGQProps, IFormularioSGQState> {
  constructor(props: IFormularioSGQProps) {
    super(props);
    this.state = {
      nomeProcesso: '',
      nomeDocumento: '',
      elaborador: props.usuarioEmail || '',
      objetivo: '',
      aplicacao: '',
      docReferencia: '',
      definicoes: '',
      papeisResponsabilidades: '',
      preRequisitos: '',
      sequenciaExecutiva: '',
      registrosComplementares: '',
      enviando: false
    };
  }

  // ==========================================
  // FUNÇÕES DE COMUNICAÇÃO COM O SHAREPOINT
  // ==========================================

  // 1. Busca o arquivo template .docx em branco no SharePoint
  private getTemplateSgq = async (): Promise<ArrayBuffer> => {
    // URL atualizada com o caminho correto da biblioteca e nome do arquivo
    const caminhoRelativo = "/sites/IntranetGrunner/Templates_SGQ/Template - Instrução de Trabalho.docx";
    const urlTemplate = `${this.props.spContext.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${caminhoRelativo}')/$value`;

    const response = await this.props.spContext.spHttpClient.get(urlTemplate, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(`Erro ao baixar template: ${response.statusText}`);
    }

    return await response.arrayBuffer();
  }

  // 2. Faz o upload do documento finalizado para a pasta de Rascunhos
  private uploadRascunhoSgq = async (blob: Blob, nomeArquivo: string): Promise<void> => {
    // Apontando para a biblioteca RascunhosSGQ
    const urlUpload = `${this.props.spContext.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('RascunhosSGQ')/RootFolder/Files/add(url='${nomeArquivo}', overwrite=true)`;

    const response = await this.props.spContext.spHttpClient.post(urlUpload, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata'
      },
      body: blob
    });

    if (!response.ok) {
      throw new Error(`Erro no upload: ${response.statusText}`);
    }
  }

  // 3. Envia um e-mail nativo pelo SharePoint para notificar a Qualidade
  private notificarQualidade = async (nomeDocumento: string, nomeArquivo: string): Promise<void> => {
    const urlEmail = `${this.props.spContext.pageContext.web.absoluteUrl}/_api/SP.Utilities.Utility.SendEmail`;
    const emailDestino = "malavazi.gabriel@grunnertec.com.br";

    // 1. Corpo da mensagem (HTML limpo)
    const corpoHtml = `
      <div style="font-family: Arial, sans-serif; color: #333;">
        <h2>Novo Rascunho de Documento SGQ</h2>
        <p>Olá equipe da Qualidade,</p>
        <p>Um novo rascunho foi gerado pelo sistema e está aguardando a análise de vocês.</p>
        <ul>
          <li><strong>Tipo:</strong> ${this.props.tipoDocumento}</li>
          <li><strong>Processo:</strong> ${this.state.nomeProcesso}</li>
          <li><strong>Documento:</strong> ${nomeDocumento}</li>
          <li><strong>Elaborador:</strong> ${this.state.elaborador}</li>
        </ul>
        <p>O arquivo <b>${nomeArquivo}</b> já está disponível na biblioteca de <strong>RascunhosSGQ</strong> para revisão e aprovação oficial.</p>
      </div>
    `.replace(/\r?\n|\r/g, "");

    // 2. NOVO FORMATO: OData=nometadata (Padrão moderno do SPFx)
    // Tudo é passado diretamente, sem objetos complexos ou metadados de versão.
    const emailProps = {
      properties: {
        To: [emailDestino],
        Subject: `Novo Rascunho SGQ para Análise: ${nomeDocumento}`,
        Body: corpoHtml,
        AdditionalHeaders: {
          "content-type": "text/html"
        }
      }
    };

    try {
      const response = await this.props.spContext.spHttpClient.post(urlEmail, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: JSON.stringify(emailProps)
      });

      if (!response.ok) {
        const erroDetalhado = await response.text();
        throw new Error(`Servidor retornou status ${response.status}: ${erroDetalhado}`);
      }

      console.log("E-mail disparado com sucesso pela API nativa!");

    } catch (error) {
      console.error("Erro ao enviar o e-mail de notificação:", error);
      alert("Atenção: O documento foi salvo nos rascunhos, mas o aviso por e-mail falhou. Consulte o console (F12) para detalhes.");
    }
  }

  // ==========================================
  // PROCESSAMENTO DO DOCUMENTO E ENVIO
  // ==========================================

  private gerarDocumentoSgqEmMemoria = async (): Promise<Blob> => {
    try {
      const content = await this.getTemplateSgq();

      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      const dataAtual = new Date().toLocaleDateString('pt-BR');

      doc.render({
        tipoDocumento: this.props.tipoDocumento,
        it_NumeroDoc: "GR.XXX.IT.00X",
        it_DataCabecalho: dataAtual,
        it_NomeProcesso: this.state.nomeProcesso,
        it_NomeDocumento: this.state.nomeDocumento,
        it_DataEmissao: dataAtual,
        it_Elaborador: this.state.elaborador,
        it_Objetivo: this.state.objetivo,
        it_Aplicacao: this.state.aplicacao || 'N/A',
        it_DocReferencia: this.state.docReferencia || 'N/A',
        it_Definicoes: this.state.definicoes || 'N/A',
        it_Papeis: this.state.papeisResponsabilidades || 'N/A',
        it_PreRequisitos: this.state.preRequisitos || 'N/A',
        it_Sequencia: this.state.sequenciaExecutiva || 'N/A',
        it_Registros: this.state.registrosComplementares || 'N/A'
      });

      const blob = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });

      return blob;

    } catch (error) {
      console.error("Erro ao gerar o Word em memória:", error);
      throw error;
    }
  }

  private enviarParaAprovacao = async (): Promise<void> => {
    const { nomeProcesso, nomeDocumento, objetivo } = this.state;

    if (!nomeProcesso || !nomeDocumento || !objetivo) {
      alert("Por favor, preencha os campos obrigatórios (Nome do Processo, Nome do Documento e Objetivo).");
      return;
    }

    this.setState({ enviando: true });

    try {
      const documentoBlob = await this.gerarDocumentoSgqEmMemoria();
      const nomeArquivo = `Rascunho_${this.props.tipoDocumento}_${nomeDocumento.replace(/\s+/g, '_')}.docx`;

      // 1. Faz o upload do documento no SharePoint
      await this.uploadRascunhoSgq(documentoBlob, nomeArquivo);

      alert("Solicitação enviada para a Qualidade com sucesso! O documento foi gerado e salvo nos rascunhos.");
      this.setState({ enviando: false });
      this.props.onSucesso();

    } catch (erro) {
      console.error("Erro ao gerar/enviar o documento:", erro);
      alert("Houve um erro ao processar o documento. Verifique as URLs do SharePoint no código e o console para mais detalhes.");
      this.setState({ enviando: false });
    }
  }

  public render(): React.ReactElement<IFormularioSGQProps> {
    const { tipoDocumento, onFechar } = this.props;
    const {
      nomeProcesso, nomeDocumento, elaborador, objetivo, aplicacao,
      docReferencia, definicoes, papeisResponsabilidades, preRequisitos,
      sequenciaExecutiva, registrosComplementares, enviando
    } = this.state;

    const textareaStyle = { padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px', fontFamily: 'inherit', resize: 'vertical' as 'vertical', minHeight: '80px' };

    return (
      <div className={styles.editModalBackdrop}>
        <div className={styles.editModal} style={{ maxHeight: '90vh', overflowY: 'auto' }}>

          <div className={styles.editModalHeader}>
            <h2>Solicitar Novo: {tipoDocumento}</h2>
            <button onClick={onFechar} className={styles.closeModal}>✕</button>
          </div>

          <div className={styles.editModalBody}>
            <div className={styles.formGrid}>
              <div className={styles.formGroup}>
                <label>Nome do Processo *</label>
                <input type="text" placeholder="Ex: Gestão da Qualidade" value={nomeProcesso} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this.setState({ nomeProcesso: e.target.value })} />
              </div>
              <div className={styles.formGroup}>
                <label>Nome do Documento *</label>
                <input type="text" placeholder="Ex: POP de Auditoria Interna" value={nomeDocumento} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this.setState({ nomeDocumento: e.target.value })} />
              </div>
              <div className={styles.formGroup}>
                <label>Elaborador (E-mail ou Nome)</label>
                <input type="text" value={elaborador} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this.setState({ elaborador: e.target.value })} />
              </div>
              <div className={styles.formGroup}>
                <label>Documento de Referência</label>
                <input type="text" placeholder="Ex: ISO 9001" value={docReferencia} onChange={(e: React.ChangeEvent<HTMLInputElement>) => this.setState({ docReferencia: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Objetivo *</label>
                <textarea rows={2} style={textareaStyle} placeholder="Descreva o objetivo principal..." value={objetivo} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ objetivo: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Aplicação / Escopo</label>
                <textarea rows={2} style={textareaStyle} placeholder="Ex: Aplica-se a todos os colaboradores da matriz..." value={aplicacao} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ aplicacao: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Definições</label>
                <textarea rows={3} style={textareaStyle} placeholder="Termos e siglas utilizados no documento..." value={definicoes} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ definicoes: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Papéis e Responsabilidades</label>
                <textarea rows={3} style={textareaStyle} placeholder="Quem faz o que neste processo..." value={papeisResponsabilidades} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ papeisResponsabilidades: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Pré-requisitos</label>
                <textarea rows={2} style={textareaStyle} placeholder="O que é necessário antes de iniciar o processo..." value={preRequisitos} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ preRequisitos: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Sequência Executiva</label>
                <textarea rows={6} style={textareaStyle} placeholder="Passo a passo detalhado do processo..." value={sequenciaExecutiva} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ sequenciaExecutiva: e.target.value })} />
              </div>
              <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                <label>Registros e Documentos Complementares</label>
                <textarea rows={2} style={textareaStyle} placeholder="Ex: GR.XX.FOR.00X - Formulário de Vistoria..." value={registrosComplementares} onChange={(e: React.ChangeEvent<HTMLTextAreaElement>) => this.setState({ registrosComplementares: e.target.value })} />
              </div>
            </div>
          </div>

          <div className={styles.editModalFooter}>
            <button className={styles.cancelBtn} onClick={onFechar}>Cancelar</button>
            <button className={styles.saveBtn} onClick={this.enviarParaAprovacao} disabled={enviando} style={{ opacity: enviando ? 0.6 : 1 }}>
              {enviando ? 'Gerando documento...' : 'Enviar para Aprovação ➔'}
            </button>
          </div>
        </div>
      </div>
    );
  }
}