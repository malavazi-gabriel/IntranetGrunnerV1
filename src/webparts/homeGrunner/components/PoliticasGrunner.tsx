import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';

// Interface das props (ajuste conforme o nome real da sua prop)
export interface IPoliticasGrunnerProps {
  description: string;
  context: any; 
}

interface IPoliticasGrunnerState {
  areaAtiva: string;
  documentos: any[];
  loading: boolean;
}

export default class PoliticasGrunner extends React.Component<IPoliticasGrunnerProps, IPoliticasGrunnerState> {
  
  // As áreas que terão abas
  private areas = ['Institucional', 'TI', 'Marketing', 'RH', 'Operacional'];

  constructor(props: IPoliticasGrunnerProps) {
    super(props);
    this.state = {
      areaAtiva: 'Institucional', // Aba padrão ao abrir a página
      documentos: [],
      loading: true
    };
  }

  public componentDidMount() {
    this.buscarDocumentos(this.state.areaAtiva);
  }

  // Função que busca os documentos na biblioteca do SharePoint filtrando pela área
  private buscarDocumentos = async (area: string) => {
    this.setState({ areaAtiva: area, loading: true, documentos: [] });

    try {
      // ATENÇÃO: É necessário ter uma Biblioteca de Documentos chamada 'PoliticasGrunner' com uma coluna Choice chamada 'Area'
      const url = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PoliticasGrunner')/items?$select=FileLeafRef,ServerRelativeUrl,Area&$filter=Area eq '${area}'`;
      const response = await this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();

      if (data && data.value) {
        this.setState({ documentos: data.value, loading: false });
      } else {
        this.setState({ loading: false });
      }
    } catch (error) {
      console.error(`Erro ao buscar políticas de ${area}:`, error);
      this.setState({ loading: false });
    }
  }

  public render(): React.ReactElement<IPoliticasGrunnerProps> {
    return (
      <div className={styles.container}>
        
        {/* --- CABEÇALHO DA PÁGINA --- */}
        <header className={styles.pageHeader}>
          <div className={styles.headerText}>
            <h1>📖 Políticas e Diretrizes Grunner</h1>
            <p>Acesse os documentos normativos, manuais e procedimentos de cada área da empresa.</p>
          </div>
        </header>

        {/* --- MENU DE ABAS (TABS) --- */}
        <nav className={styles.tabsContainer}>
          {this.areas.map((area) => (
            <button
              key={area}
              className={this.state.areaAtiva === area ? styles.tabActive : styles.tab}
              onClick={() => this.buscarDocumentos(area)}
            >
              {area}
            </button>
          ))}
        </nav>

        {/* --- ÁREA DOS DOCUMENTOS --- */}
        <main className={styles.documentsArea}>
          {this.state.loading ? (
            <div className={styles.loadingState}>
              <div className={styles.spinner}></div>
              <p>Buscando documentos de {this.state.areaAtiva}...</p>
            </div>
          ) : this.state.documentos.length > 0 ? (
            <div className={styles.documentGrid}>
              {this.state.documentos.map((doc, index) => {
                // Tenta descobrir a extensão do arquivo para mostrar o ícone certo
                const extensao = doc.FileLeafRef.split('.').pop()?.toLowerCase();
                const isPdf = extensao === 'pdf';
                
                return (
                  <a key={index} href={doc.ServerRelativeUrl} target="_blank" rel="noopener noreferrer" className={styles.documentCard}>
                    <div className={isPdf ? styles.iconPdf : styles.iconDoc}>
                      {isPdf ? 'PDF' : 'DOC'}
                    </div>
                    <div className={styles.docInfo}>
                      <h3>{doc.FileLeafRef.replace(`.${extensao}`, '')}</h3>
                      <span>Visualizar arquivo</span>
                    </div>
                  </a>
                );
              })}
            </div>
          ) : (
            <div className={styles.emptyState}>
              <p>📭 Nenhum documento encontrado para a área de <strong>{this.state.areaAtiva}</strong> no momento.</p>
            </div>
          )}
        </main>

      </div>
    );
  }
}