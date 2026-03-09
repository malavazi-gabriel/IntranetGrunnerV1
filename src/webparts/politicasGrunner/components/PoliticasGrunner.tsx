import * as React from 'react';
import styles from './PoliticasGrunner.module.scss';
import type { IPoliticasGrunnerProps } from './IPoliticasGrunnerProps';
import { SPHttpClient } from '@microsoft/sp-http';

interface IPoliticasGrunnerState {
  areaAtiva: string;
  todosDocumentos: any[];
  loading: boolean;
  termoBusca: string;
}

export default class PoliticasGrunner extends React.Component<IPoliticasGrunnerProps, IPoliticasGrunnerState> {
  
  private areas = ['Institucional', 'TI', 'Marketing', 'RH', 'Operacional'];

  constructor(props: IPoliticasGrunnerProps) {
    super(props);
    this.state = {
      areaAtiva: 'Institucional', 
      todosDocumentos: [],
      loading: true,
      termoBusca: ''
    };
  }

  public componentDidMount(): void {
    this.buscarTodosDocumentos();
  }

  private buscarTodosDocumentos = async (): Promise<void> => {
    try {
      // CORREÇÃO: Removemos o $expand e o /Title. Coluna do tipo 'Opção' retorna o texto direto!
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
      // Busca Global
      documentosExibidos = todosDocumentos.filter(doc => 
        doc.FileLeafRef && doc.FileLeafRef.toLowerCase().includes(termoBusca.toLowerCase())
      );
    } else {
      // CORREÇÃO: Lemos doc.Area diretamente (antes estava doc.Area.Title)
      documentosExibidos = todosDocumentos.filter(doc => 
        doc.Area === areaAtiva
      );
    }

    return (
      <div className={styles.container}>
        
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
                // CORREÇÃO: Lemos a área direto da propriedade
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
    );
  }
}