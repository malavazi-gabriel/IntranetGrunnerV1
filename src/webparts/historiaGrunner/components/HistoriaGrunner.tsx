import * as React from 'react';
import styles from './HistoriaGrunner.module.scss';
import { IHistoriaGrunnerProps } from './IHistoriaGrunnerProps';

const homeUrl = "https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/SitePages/Inicio.aspx?env=Embedded";

export default class HistoriaGrunner extends React.Component<IHistoriaGrunnerProps, {}> {
  private footerObserver?: MutationObserver;

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
      this.collapseElement(node as HTMLElement);
    });
  }

  private fixSharePointCanvasSpacing = (): void => {
    const applyFullBleed = (element: HTMLElement | null): void => {
      if (!element) return;

      element.style.setProperty('margin', '0', 'important');
      element.style.setProperty('padding', '0', 'important');
      element.style.setProperty('left', '0', 'important');
      element.style.setProperty('right', '0', 'important');
      element.style.setProperty('max-width', '100%', 'important');
      element.style.setProperty('width', '100%', 'important');
      element.style.setProperty('box-sizing', 'border-box', 'important');
      element.style.setProperty('background', 'transparent', 'important');
    };

    applyFullBleed(document.documentElement as unknown as HTMLElement);
    applyFullBleed(document.body);

    document.documentElement.style.setProperty('overflow-x', 'hidden', 'important');
    document.body?.style.setProperty('overflow-x', 'hidden', 'important');
    document.documentElement.style.setProperty('background', '#f3f4f6', 'important');
    document.body?.style.setProperty('background', '#f3f4f6', 'important');

    const selectors = [
      '#spPageChromeAppDiv',
      '[data-automation-id="contentScrollRegion"]',
      '#workbenchPageContent',
      '#spPageCanvasContent',
      '.SPCanvas-canvas',
      'div[data-automation-id="Canvas"]',
      'div[data-automation-id="CanvasZone"]',
      'div[data-automation-id="CanvasZone"] > div',
      '.CanvasZone',
      '.CanvasSection',
      '.ControlZone',
      'div[class*="CanvasComponent"]'
    ];

    const elements = document.querySelectorAll(selectors.join(','));

    elements.forEach((node) => {
      applyFullBleed(node as HTMLElement);
    });
  }

  public componentDidMount(): void {
    if (this.shouldHideSharePointChrome()) {
      const applyFixes = (): void => {
        this.hideSharePointFooter();
        this.hideSharePointAppBar();
        this.fixSharePointCanvasSpacing();
      };

      applyFixes();
      window.setTimeout(applyFixes, 500);
      window.setTimeout(applyFixes, 1500);
      window.setTimeout(applyFixes, 3000);

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

  public render(): React.ReactElement<IHistoriaGrunnerProps> {
    return (
      <div className={styles.container}>
        {this.shouldHideSharePointChrome() && (
          <style>{`
            [data-automation-id="page-bottom-actions"],
            [data-automation-id="page-bottom-bar"],
            #sp-page-footer,
            .CommentsWrapper,
            [data-automation-id="socialBar"],
            div[class*="socialBar_"],
            div[class*="footer_"],
            div[class*="pageBottomBar_"],
            #sp-appBar,
            [data-automation-id="sp-appBar"],
            div[class^="appBar_"],
            div[class*="sp-appBar"],
            #SuiteNavWrapper,
            #spSiteHeader,
            #spCommandBar,
            div[data-automation-id="pageHeader"],
            div[class^="commandBarWrapper_"] {
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

            html,
            body {
              margin: 0 !important;
              padding: 0 !important;
              overflow-x: hidden !important;
              background: #f3f4f6 !important;
            }

            #spPageChromeAppDiv,
            [data-automation-id="contentScrollRegion"],
            #workbenchPageContent,
            #spPageCanvasContent,
            .SPCanvas-canvas,
            div[data-automation-id="Canvas"],
            div[data-automation-id="CanvasZone"],
            div[data-automation-id="CanvasZone"] > div,
            .CanvasZone,
            .CanvasSection,
            .ControlZone,
            div[class*="CanvasComponent"] {
              margin: 0 !important;
              padding: 0 !important;
              left: 0 !important;
              right: 0 !important;
              max-width: 100% !important;
              width: 100% !important;
              box-sizing: border-box !important;
              background: transparent !important;
            }
          `}</style>
        )}

        <div className={styles.heroSection}>
          <a href={homeUrl} className={styles.backBtn}>
            ← Voltar para a Intranet
          </a>
          <div className={styles.heroContent}>
            <h1>Nossa História</h1>
            <p>Inovando com velocidade para atender às necessidades reais do campo.</p>
          </div>
        </div>

        <div className={styles.mvvContainer}>
          <div className={styles.mvvCard}>
            <div className={styles.icon}>🎯</div>
            <h3>Nossa Missão</h3>
            <p>Desenvolver soluções inovadoras, aumentando a produtividade e a sustentabilidade no campo através da tecnologia.</p>
          </div>
          <div className={styles.mvvCard}>
            <div className={styles.icon}>👁️</div>
            <h3>Nossa Visão</h3>
            <p>Ser a principal referência em tecnologia agrícola, liderando a transformação e o cuidado com a lavoura.</p>
          </div>
          <div className={styles.mvvCard}>
            <div className={styles.icon}>💎</div>
            <h3>Nossos Valores</h3>
            <p>Inovação constante, foco no produtor, sustentabilidade (redução de CO2) e excelência operacional.</p>
          </div>
        </div>

        <div className={styles.storyBlock}>
          <h2>A Origem da Inovação</h2>
          <p>
            Na nova era tecnológica, as transformações trazidas por grandes inovações são quase sempre muito rápidas. Com esse conceito de inovar com velocidade que atendesse às necessidades dos produtores de cana-de-açúcar os irmãos <strong>Henrique e Mateus Belei</strong>, tradicionais produtores de cana de Lençóis Paulista, no interior do estado de São Paulo, criaram a Grunner, no ano de 2018.
          </p>
          <p>
            Incomodados com o chamado 'pisoteio' nas linhas de cana, onde o trator, literalmente, 'pisa' na planta, eles resolveram adaptar um caminhão para executar a operação de colheita e aplicação de insumos. A estratégia funcionou. Além de aumentar a produtividade da fazenda, a ideia original – um caminhão autônomo que executa as operações sem pisotear as linhas – revelou-se fundamental para reduzir custos e combater as emissões de CO2 no processo agrícola.
          </p>
          <p>
            A invenção conquistou produtores de diversas regiões brasileiras, consolidando a Grunner como uma companhia dotada de alta capacidade inovadora na produção de tecnologia para o campo, com foco no aumento de produtividade da operação e maior cuidado com a lavoura – pelo controle de tráfego e menor compactação do solo.
          </p>
        </div>

        <div className={styles.historySection}>
          <h2 className={styles.sectionTitle}>Nossa Linha do Tempo</h2>

          <div className={styles.timeline}>
            <div className={styles.timelineItem}>
              <div className={styles.timelineDot}></div>
              <div className={styles.timelineContent}>
                <span className={styles.year}>2018</span>
                <h3>O Início em Lençóis Paulista</h3>
                <p>Os irmãos Henrique e Mateus Belei criam a Grunner com a missão de adaptar caminhões para acabar com o 'pisoteio' nas linhas de cana-de-açúcar.</p>
              </div>
            </div>

            <div className={styles.timelineItem}>
              <div className={styles.timelineDot}></div>
              <div className={styles.timelineContent}>
                <span className={styles.year}>2018</span>
                <h3>A Parceria Exclusiva</h3>
                <p>No mesmo ano, é firmada a parceria com a Mercedes-Benz. Os caminhões alemães recebem o protocolo de tecnologia que deu origem às Smart Machines.</p>
              </div>
            </div>

            <div className={styles.timelineItem}>
              <div className={styles.timelineDot}></div>
              <div className={styles.timelineContent}>
                <span className={styles.year}>2021</span>
                <h3>No Top 10 da Mercedes</h3>
                <p>Com as engenharias conectadas, o sucesso da união faz com que as máquinas da Grunner sejam incluídas no ranking dos dez principais produtos da história da marca alemã.</p>
              </div>
            </div>

            <div className={styles.timelineItem}>
              <div className={styles.timelineDot}></div>
              <div className={styles.timelineContent}>
                <span className={styles.year}>Hoje</span>
                <h3>Consolidação Nacional</h3>
                <p>A Grunner conquista produtores de diversas regiões brasileiras, focada no controle de tráfego, menor compactação do solo e sustentabilidade.</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}