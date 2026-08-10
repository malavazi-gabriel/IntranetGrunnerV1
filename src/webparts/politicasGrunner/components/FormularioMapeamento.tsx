import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { SPHttpClient } from '@microsoft/sp-http';

interface IFormularioMapeamentoProps {
    tipoDocumento: string;
    usuarioEmail: string;
    spContext: any;
    onFechar: () => void;
    onSucesso: () => void;
}

interface IFormularioMapeamentoState {
    nomeDocumento: string;
    elaborador: string;
    objetivo: string;
    mp_PartesInteressadas: string;
    mp_Fornecedores: string;
    mp_Executores: string;
    mp_Unidades: string;
    mp_Clientes: string;
    mp_Entradas: string;
    mp_Processos: string;
    mp_Saidas: string;
    mp_Requisitos: string;
    mp_Recursos: string;
    mp_Indicadores: string;
    enviando: boolean;
}

export default class FormularioMapeamento extends React.Component<IFormularioMapeamentoProps, IFormularioMapeamentoState> {
    constructor(props: IFormularioMapeamentoProps) {
        super(props);
        this.state = {
            nomeDocumento: '',
            elaborador: props.usuarioEmail || '',
            objetivo: '',
            mp_PartesInteressadas: '', mp_Fornecedores: '', mp_Executores: '', mp_Unidades: '', mp_Clientes: '',
            mp_Entradas: '', mp_Processos: '', mp_Saidas: '', mp_Requisitos: '', mp_Recursos: '', mp_Indicadores: '',
            enviando: false
        };
    }

    private getTemplateSgq = async (): Promise<ArrayBuffer> => {
        // Aponta direto para o Template de Mapeamento
        const caminhoRelativo = "/sites/IntranetGrunner/Templates_SGQ/Template - Mapeamento de Processo.docx";
        const urlTemplate = `${this.props.spContext.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${caminhoRelativo}')/$value`;

        const response = await this.props.spContext.spHttpClient.get(urlTemplate, SPHttpClient.configurations.v1);
        if (!response.ok) throw new Error(`Erro ao baixar template: ${response.statusText}`);
        return await response.arrayBuffer();
    }

    private uploadRascunhoSgq = async (blob: Blob, nomeArquivo: string): Promise<void> => {
        const urlUpload = `${this.props.spContext.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('RascunhosSGQ')/RootFolder/Files/add(url='${nomeArquivo}', overwrite=true)`;
        const response = await this.props.spContext.spHttpClient.post(urlUpload, SPHttpClient.configurations.v1, {
            headers: { 'Accept': 'application/json;odata=nometadata', 'Content-type': 'application/json;odata=nometadata' },
            body: blob
        });
        if (!response.ok) throw new Error(`Erro no upload: ${response.statusText}`);
    }

    private gerarDocumentoSgqEmMemoria = async (): Promise<Blob> => {
        try {
            const content = await this.getTemplateSgq();
            const zip = new PizZip(content);
            const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            // 1. Pega o texto do textarea e fatia separando por quebras de linha (Enter)
            // O .filter remove linhas vazias caso o usuário dê enters a mais
            const listaProcessos = (this.state.mp_Processos || '')
                .split('\n')
                .map(p => p.trim())
                .filter(p => p.length > 0);

            doc.render({
                mp_NomeProcesso: this.state.nomeDocumento || 'N/A',
                mp_Responsavel: this.state.elaborador || 'N/A',
                mp_Objetivo: this.state.objetivo || 'N/A',
                mp_PartesInteressadas: this.state.mp_PartesInteressadas || 'N/A',
                mp_Fornecedores: this.state.mp_Fornecedores || 'N/A',
                mp_Executores: this.state.mp_Executores || 'N/A',
                mp_Unidades: this.state.mp_Unidades || 'N/A',
                mp_Clientes: this.state.mp_Clientes || 'N/A',
                mp_Entradas: this.state.mp_Entradas || 'N/A',
                mp_Saidas: this.state.mp_Saidas || 'N/A',
                mp_Requisitos: this.state.mp_Requisitos || 'N/A',
                mp_Recursos: this.state.mp_Recursos || 'N/A',
                mp_Indicadores: this.state.mp_Indicadores || 'N/A',

                // 2. Distribui cada processo fatiado para a sua respectiva caixinha no Word
                // Se não houver texto para aquela caixa, ele envia uma string vazia para limpar a caixa
                proc1: listaProcessos[0] || '',
                proc2: listaProcessos[1] || '',
                proc3: listaProcessos[2] || '',
                proc4: listaProcessos[3] || '',
                proc5: listaProcessos[4] || '',
                proc6: listaProcessos[5] || ''
            });

            return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        } catch (error) {
            console.error("Erro ao gerar Word em memória:", error);
            throw error;
        }
    }

    private enviarParaAprovacao = async (): Promise<void> => {
        if (!this.state.nomeDocumento || !this.state.objetivo) {
            alert("Por favor, preencha o Nome do Macroprocesso e o Objetivo.");
            return;
        }
        this.setState({ enviando: true });
        try {
            const documentoBlob = await this.gerarDocumentoSgqEmMemoria();
            const nomeArquivo = `Rascunho_MAPEAMENTO_${this.state.nomeDocumento.replace(/\s+/g, '_')}.docx`;
            await this.uploadRascunhoSgq(documentoBlob, nomeArquivo);

            alert("Solicitação enviada para a Qualidade com sucesso!");
            this.setState({ enviando: false });
            this.props.onSucesso();
        } catch (erro) {
            console.error("Erro ao gerar/enviar o documento:", erro);
            alert("Houve um erro ao processar o documento. Verifique o console.");
            this.setState({ enviando: false });
        }
    }

    public render(): React.ReactElement<IFormularioMapeamentoProps> {
        const { onFechar } = this.props;
        const textareaStyle = { padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px', fontFamily: 'inherit', resize: 'vertical' as 'vertical', minHeight: '80px' };

        return (
            <div className={styles.editModalBackdrop}>
                <div className={styles.editModal} style={{ maxHeight: '90vh', overflowY: 'auto', width: '90%', maxWidth: '800px' }}>

                    <div className={styles.editModalHeader}>
                        <h2>Mapeamento de Processo</h2>
                        <button onClick={onFechar} className={styles.closeModal}>✕</button>
                    </div>

                    <div className={styles.editModalBody}>
                        <div className={styles.formGrid}>

                            {/* --- SEÇÃO 1: INFORMAÇÕES GERAIS --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>📌 1. Informações Gerais</h3>
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Nome do Macroprocesso *</label>
                                <input type="text" placeholder="Ex: Gestão Financeira" value={this.state.nomeDocumento} onChange={(e) => this.setState({ nomeDocumento: e.target.value })} />
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Responsável</label>
                                <input type="text" value={this.state.elaborador} onChange={(e) => this.setState({ elaborador: e.target.value })} />
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Objetivo do Macroprocesso *</label>
                                <textarea rows={2} style={textareaStyle} value={this.state.objetivo} onChange={(e) => this.setState({ objetivo: e.target.value })} />
                            </div>

                            {/* --- SEÇÃO 2: ATORES E ENVOLVIDOS --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px', marginTop: '15px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>👥 2. Atores e Envolvidos</h3>
                            </div>

                            <div className={styles.formGroup}><label>Partes Interessadas</label><textarea rows={3} style={textareaStyle} value={this.state.mp_PartesInteressadas} onChange={(e) => this.setState({ mp_PartesInteressadas: e.target.value })} /></div>
                            <div className={styles.formGroup}><label>Fornecedores</label><textarea rows={3} style={textareaStyle} value={this.state.mp_Fornecedores} onChange={(e) => this.setState({ mp_Fornecedores: e.target.value })} /></div>
                            <div className={styles.formGroup}><label>Executores</label><textarea rows={3} style={textareaStyle} value={this.state.mp_Executores} onChange={(e) => this.setState({ mp_Executores: e.target.value })} /></div>
                            <div className={styles.formGroup}><label>Unidades Envolvidas</label><textarea rows={3} style={textareaStyle} value={this.state.mp_Unidades} onChange={(e) => this.setState({ mp_Unidades: e.target.value })} /></div>
                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}><label>Clientes</label><textarea rows={2} style={textareaStyle} value={this.state.mp_Clientes} onChange={(e) => this.setState({ mp_Clientes: e.target.value })} /></div>

                            {/* --- SEÇÃO 3: FLUXO OPERACIONAL --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px', marginTop: '15px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>🔄 3. Fluxo Operacional</h3>
                            </div>

                            <div className={styles.formGroup}><label>Entradas</label><textarea rows={4} style={textareaStyle} value={this.state.mp_Entradas} onChange={(e) => this.setState({ mp_Entradas: e.target.value })} /></div>
                            <div className={styles.formGroup}><label>Saídas</label><textarea rows={4} style={textareaStyle} value={this.state.mp_Saidas} onChange={(e) => this.setState({ mp_Saidas: e.target.value })} /></div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Processos</label>
                                <textarea
                                    rows={5}
                                    style={{ ...textareaStyle, borderColor: '#A6CE39', backgroundColor: '#F9FAFB' }}
                                    placeholder="Digite um processo por linha (até 6 processos para preencher os quadros)...&#10;1. Aprovação do orçamento&#10;2. Emissão da nota fiscal&#10;3. Pagamento do fornecedor"
                                    value={this.state.mp_Processos}
                                    onChange={(e) => this.setState({ mp_Processos: e.target.value })}
                                />
                            </div>

                            {/* --- SEÇÃO 4: CONTROLE E REQUISITOS --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px', marginTop: '15px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>📊 4. Controle e Gestão</h3>
                            </div>

                            <div className={styles.formGroup}><label>Requisitos Aplicáveis / Métodos</label><textarea rows={3} style={textareaStyle} value={this.state.mp_Requisitos} onChange={(e) => this.setState({ mp_Requisitos: e.target.value })} /></div>
                            <div className={styles.formGroup}><label>Recursos e Sistemas</label><textarea rows={3} style={textareaStyle} value={this.state.mp_Recursos} onChange={(e) => this.setState({ mp_Recursos: e.target.value })} /></div>
                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}><label>Indicadores</label><textarea rows={3} style={textareaStyle} value={this.state.mp_Indicadores} onChange={(e) => this.setState({ mp_Indicadores: e.target.value })} /></div>

                        </div>
                    </div>

                    <div className={styles.editModalFooter}>
                        <button className={styles.cancelBtn} onClick={onFechar}>Cancelar</button>
                        <button className={styles.saveBtn} onClick={this.enviarParaAprovacao} disabled={this.state.enviando} style={{ opacity: this.state.enviando ? 0.6 : 1 }}>
                            {this.state.enviando ? 'Gerando documento...' : 'Salvar Rascunho ➔'}
                        </button>
                    </div>
                </div>
            </div>
        );
    }
}