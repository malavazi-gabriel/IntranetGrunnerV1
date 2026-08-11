import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { SPHttpClient } from '@microsoft/sp-http';

interface IFormularioProcedimentoProps {
    tipoDocumento: string;
    usuarioEmail: string;
    spContext: any;
    onFechar: () => void;
    onSucesso: () => void;
}

interface IPapelResponsabilidade {
    nomePapel: string;
    listaResponsabilidades: string;
}

interface IEtapaDescricao {
    tituloEtapa: string;
    detalhesEtapa: string;
}

interface IFormularioProcedimentoState {
    nomeDocumento: string;
    elaborador: string;
    pr_Objetivo: string;
    pr_Aplicacao: string;
    pr_Referencia: string;
    pr_Definicoes: string;
    papeis: IPapelResponsabilidade[];
    pr_Fluxograma: string;
    descricoes: IEtapaDescricao[]; // Novo array dinâmico para a Descrição
    enviando: boolean;
}

export default class FormularioProcedimento extends React.Component<IFormularioProcedimentoProps, IFormularioProcedimentoState> {
    constructor(props: IFormularioProcedimentoProps) {
        super(props);
        this.state = {
            nomeDocumento: '',
            elaborador: props.usuarioEmail || '',
            pr_Objetivo: '',
            pr_Aplicacao: '',
            pr_Referencia: '',
            pr_Definicoes: '',
            papeis: [{ nomePapel: '', listaResponsabilidades: '' }],
            pr_Fluxograma: 'Não aplicável.',
            descricoes: [{ tituloEtapa: '', detalhesEtapa: '' }], // Inicia com 1 etapa em branco
            enviando: false
        };
    }

    private getTemplateSgq = async (): Promise<ArrayBuffer> => {
        const caminhoRelativo = "/sites/IntranetGrunner/Templates_SGQ/Template - Procedimento.docx";
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

            // Processa o array de Papéis
            const papeisParaWord = this.state.papeis.map(papel => {
                return {
                    papel_nome: papel.nomePapel || 'Sem Nome',
                    resp_list: (papel.listaResponsabilidades || '')
                        .split('\n')
                        .map(r => r.trim())
                        .filter(r => r.length > 0)
                        .map(r => { return { desc: r }; })
                };
            });

            // Processa o array de Descrições
            const descricoesParaWord = this.state.descricoes.map(etapa => {
                return {
                    desc_titulo: etapa.tituloEtapa || 'Sem Título',
                    desc_items: (etapa.detalhesEtapa || '')
                        .split('\n')
                        .map(r => r.trim())
                        .filter(r => r.length > 0)
                        .map(r => { return { item: r }; })
                };
            });

            doc.render({
                pr_NomeDocumento: this.state.nomeDocumento || 'N/A',
                pr_Elaborador: this.state.elaborador || 'N/A',
                pr_Objetivo: this.state.pr_Objetivo || 'N/A',
                pr_Aplicacao: this.state.pr_Aplicacao || 'N/A',
                pr_Referencia: this.state.pr_Referencia || 'N/A',
                pr_Definicoes: this.state.pr_Definicoes || 'N/A',
                pr_Papeis: papeisParaWord,
                pr_Fluxograma: this.state.pr_Fluxograma || 'N/A',
                pr_DescricaoList: descricoesParaWord,
                rev_data: new Date().toLocaleDateString('pt-BR'),
                rev_elaborador: this.state.elaborador || 'N/A'
            });

            return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        } catch (error) {
            console.error("Erro ao gerar Word em memória:", error);
            throw error;
        }
    }

    private enviarParaAprovacao = async (): Promise<void> => {
        if (!this.state.nomeDocumento || !this.state.pr_Objetivo) {
            alert("Por favor, preencha o Nome do Documento e o Objetivo.");
            return;
        }
        this.setState({ enviando: true });
        try {
            const documentoBlob = await this.gerarDocumentoSgqEmMemoria();
            const sigla = this.props.tipoDocumento === 'POLÍTICA' ? 'POL' : 'PROC';
            const nomeArquivo = `Rascunho_${sigla}_${this.state.nomeDocumento.replace(/\s+/g, '_')}.docx`;
            
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

    // Gerenciamento do Estado dos Papéis
    private atualizarPapel = (index: number, campo: 'nomePapel' | 'listaResponsabilidades', valor: string) => {
        const novosPapeis = [...this.state.papeis];
        novosPapeis[index][campo] = valor;
        this.setState({ papeis: novosPapeis });
    }

    private adicionarPapel = (e: React.MouseEvent) => {
        e.preventDefault();
        this.setState({ papeis: [...this.state.papeis, { nomePapel: '', listaResponsabilidades: '' }] });
    }

    private removerPapel = (index: number, e: React.MouseEvent) => {
        e.preventDefault();
        const novosPapeis = [...this.state.papeis];
        novosPapeis.splice(index, 1);
        this.setState({ papeis: novosPapeis });
    }

    // Gerenciamento do Estado das Descrições (Etapas)
    private atualizarDescricao = (index: number, campo: 'tituloEtapa' | 'detalhesEtapa', valor: string) => {
        const novasDescricoes = [...this.state.descricoes];
        novasDescricoes[index][campo] = valor;
        this.setState({ descricoes: novasDescricoes });
    }

    private adicionarDescricao = (e: React.MouseEvent) => {
        e.preventDefault();
        this.setState({ descricoes: [...this.state.descricoes, { tituloEtapa: '', detalhesEtapa: '' }] });
    }

    private removerDescricao = (index: number, e: React.MouseEvent) => {
        e.preventDefault();
        const novasDescricoes = [...this.state.descricoes];
        novasDescricoes.splice(index, 1);
        this.setState({ descricoes: novasDescricoes });
    }

    public render(): React.ReactElement<IFormularioProcedimentoProps> {
        const { onFechar, tipoDocumento } = this.props;
        const textareaStyle = { padding: '10px', border: '1px solid #D1D5DB', borderRadius: '6px', fontFamily: 'inherit', resize: 'vertical' as 'vertical', minHeight: '80px' };

        return (
            <div className={styles.editModalBackdrop}>
                <div className={styles.editModal} style={{ maxHeight: '90vh', overflowY: 'auto', width: '90%', maxWidth: '800px' }}>

                    <div className={styles.editModalHeader}>
                        <h2>Solicitar Novo: {tipoDocumento}</h2>
                        <button onClick={onFechar} className={styles.closeModal}>✕</button>
                    </div>

                    <div className={styles.editModalBody}>
                        <div className={styles.formGrid}>

                            {/* --- SEÇÃO 1: INFORMAÇÕES GERAIS --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>📌 1. Informações Gerais</h3>
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Nome do Documento *</label>
                                <input type="text" placeholder="Ex: Procedimento de Auditoria Interna" value={this.state.nomeDocumento} onChange={(e) => this.setState({ nomeDocumento: e.target.value })} />
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Responsável / Elaborador</label>
                                <input type="text" value={this.state.elaborador} onChange={(e) => this.setState({ elaborador: e.target.value })} />
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Objetivo *</label>
                                <textarea rows={2} style={textareaStyle} value={this.state.pr_Objetivo} onChange={(e) => this.setState({ pr_Objetivo: e.target.value })} />
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Aplicação / Escopo</label>
                                <textarea rows={2} style={textareaStyle} value={this.state.pr_Aplicacao} onChange={(e) => this.setState({ pr_Aplicacao: e.target.value })} />
                            </div>

                            {/* --- SEÇÃO 2: ESTRUTURA E RESPONSABILIDADES --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px', marginTop: '15px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>👥 2. Estrutura e Responsabilidades</h3>
                            </div>

                            <div className={styles.formGroup}><label>Documentos de Referência</label><textarea rows={3} style={textareaStyle} value={this.state.pr_Referencia} onChange={(e) => this.setState({ pr_Referencia: e.target.value })} /></div>
                            <div className={styles.formGroup}><label>Definições e Siglas</label><textarea rows={3} style={textareaStyle} value={this.state.pr_Definicoes} onChange={(e) => this.setState({ pr_Definicoes: e.target.value })} /></div>
                            
                            {/* BLOCO DINÂMICO DE PAPÉIS */}
                            <div style={{ gridColumn: 'span 2', backgroundColor: '#F9FAFB', padding: '15px', borderRadius: '8px', border: '1px solid #E5E7EB' }}>
                                <label style={{ fontWeight: 'bold', display: 'block', marginBottom: '15px', color: '#374151' }}>Papéis e Responsabilidades</label>
                                
                                {this.state.papeis.map((papel, index) => (
                                    <div key={index} style={{ marginBottom: '20px', paddingBottom: '15px', borderBottom: index !== this.state.papeis.length - 1 ? '1px dashed #D1D5DB' : 'none' }}>
                                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                                            <span style={{ fontSize: '13px', fontWeight: 'bold', color: '#6B7280' }}>Cargo / Papel #{index + 1}</span>
                                            {this.state.papeis.length > 1 && (
                                                <button onClick={(e) => this.removerPapel(index, e)} style={{ background: 'none', border: 'none', color: '#EF4444', cursor: 'pointer', fontSize: '12px' }}>🗑️ Remover</button>
                                            )}
                                        </div>
                                        
                                        <div className={styles.formGroup} style={{ marginBottom: '10px' }}>
                                            <input 
                                                type="text" 
                                                placeholder="Ex: Analista de Infraestrutura" 
                                                value={papel.nomePapel} 
                                                onChange={(e) => this.atualizarPapel(index, 'nomePapel', e.target.value)} 
                                            />
                                        </div>
                                        
                                        <div className={styles.formGroup}>
                                            <textarea 
                                                rows={4} 
                                                style={textareaStyle} 
                                                placeholder="Liste as responsabilidades (dê 'Enter' para criar os bullets no Word)...&#10;Garantir o funcionamento da rede&#10;Configurar ativos" 
                                                value={papel.listaResponsabilidades} 
                                                onChange={(e) => this.atualizarPapel(index, 'listaResponsabilidades', e.target.value)} 
                                            />
                                        </div>
                                    </div>
                                ))}

                                <button onClick={this.adicionarPapel} style={{ background: '#E5E7EB', color: '#374151', border: 'none', padding: '8px 15px', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}>
                                    ➕ Adicionar Novo Papel
                                </button>
                            </div>

                            {/* --- SEÇÃO 3: DESENVOLVIMENTO E DESCRIÇÃO --- */}
                            <div style={{ gridColumn: 'span 2', borderBottom: '2px solid #E5E7EB', paddingBottom: '8px', marginBottom: '10px', marginTop: '15px' }}>
                                <h3 style={{ margin: 0, color: '#1C2510', fontSize: '15px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>🔄 3. Desenvolvimento</h3>
                            </div>

                            <div className={styles.formGroup} style={{ gridColumn: 'span 2' }}>
                                <label>Fluxograma (Link, Referência ou "Não aplicável")</label>
                                <input type="text" value={this.state.pr_Fluxograma} onChange={(e) => this.setState({ pr_Fluxograma: e.target.value })} />
                            </div>

                            {/* BLOCO DINÂMICO DE DESCRIÇÃO (ETAPAS) */}
                            <div style={{ gridColumn: 'span 2', backgroundColor: '#F9FAFB', padding: '15px', borderRadius: '8px', border: '1px solid #E5E7EB' }}>
                                <label style={{ fontWeight: 'bold', display: 'block', marginBottom: '15px', color: '#374151' }}>Descrição Completa do Procedimento</label>
                                
                                {this.state.descricoes.map((etapa, index) => (
                                    <div key={index} style={{ marginBottom: '20px', paddingBottom: '15px', borderBottom: index !== this.state.descricoes.length - 1 ? '1px dashed #D1D5DB' : 'none' }}>
                                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '10px' }}>
                                            <span style={{ fontSize: '13px', fontWeight: 'bold', color: '#6B7280' }}>Etapa #{index + 1}</span>
                                            {this.state.descricoes.length > 1 && (
                                                <button onClick={(e) => this.removerDescricao(index, e)} style={{ background: 'none', border: 'none', color: '#EF4444', cursor: 'pointer', fontSize: '12px' }}>🗑️ Remover</button>
                                            )}
                                        </div>
                                        
                                        <div className={styles.formGroup} style={{ marginBottom: '10px' }}>
                                            <input 
                                                type="text" 
                                                placeholder={`Ex: 7.${index + 1} Avaliação de Sistemas`} 
                                                value={etapa.tituloEtapa} 
                                                onChange={(e) => this.atualizarDescricao(index, 'tituloEtapa', e.target.value)} 
                                            />
                                        </div>
                                        
                                        <div className={styles.formGroup}>
                                            <textarea 
                                                rows={5} 
                                                style={textareaStyle} 
                                                placeholder="Liste os passos detalhados desta etapa (dê 'Enter' para criar os bullets no Word)...&#10;Verificar os logs de sistema diariamente.&#10;Reportar falhas ao gestor responsável." 
                                                value={etapa.detalhesEtapa} 
                                                onChange={(e) => this.atualizarDescricao(index, 'detalhesEtapa', e.target.value)} 
                                            />
                                        </div>
                                    </div>
                                ))}

                                <button onClick={this.adicionarDescricao} style={{ background: '#E5E7EB', color: '#374151', border: 'none', padding: '8px 15px', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: 'bold' }}>
                                    ➕ Adicionar Nova Etapa
                                </button>
                            </div>

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