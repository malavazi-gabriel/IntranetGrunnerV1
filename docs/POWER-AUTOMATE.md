# Power Automate: inventário, qualidade e custo

## Base analisada

Foram lidos diretamente cinco pacotes exportados do Power Automate em 27/08/2026. A medição considera gatilhos, ações, ramificações, loops, conectores e dependências; os endereços individuais existentes nas definições não são reproduzidos neste documento.

## Inventário dos fluxos

| ID | Fluxo | Gatilho | Ações | Complexidade | Horas de reconstrução |
|---|---|---|---:|---:|---:|
| PA01 | Alerta de Notícias | Item do SharePoint criado ou modificado, consulta a cada minuto | 18 | 4/5 | 48 h |
| PA02 | Gestão ISO — alerta de vencimento | Recorrência diária às 07:00, horário de Brasília | 8 | 3/5 | 32 h |
| PA03 | Resumo Semanal da Intranet | Recorrência semanal | 16 | 4/5 | 44 h |
| PA04 | Notificação de novo rascunho SGQ | Arquivo criado em `RascunhosSGQ` | 1 | 2/5 | 12 h |
| PA05 | Ciclo de Vida de Documentos | Arquivo alterado e status “Aguardando Gestores” | 15 | 5/5 | 64 h |
| — | Trabalho compartilhado | Empacotamento, conexões por ambiente, testes integrados, documentação, publicação e monitoramento mínimo | — | — | 40 h |
| **Total** | **5 fluxos, 58 ações declaradas** |  |  |  | **240 h** |

## O que cada automação entrega

### PA01 — Alerta de Notícias

- aguarda a consolidação do item publicado;
- consulta notícia e anexos no SharePoint;
- trata imagem/anexo em loop;
- decide o envio por três condições;
- envia e-mail normal ou por caixa compartilhada;
- marca o item como enviado e inicializa contadores.

### PA02 — Alerta de vencimento ISO

- executa diariamente;
- calcula data atual e horizonte de 30 dias;
- filtra documentos não arquivados e fora de revisão;
- percorre os documentos elegíveis;
- envia alertas diferentes conforme proximidade ou vencimento.

### PA03 — Resumo semanal

- consulta publicações da semana;
- monta conteúdo HTML agregado;
- percorre notícias e anexos;
- inclui imagens quando disponíveis;
- envia o resumo por caixa compartilhada.

### PA04 — Novo rascunho SGQ

- detecta arquivo novo em `RascunhosSGQ`;
- envia notificação com autor, nome e link do documento.

### PA05 — Ciclo de vida documental

- coleta os aprovadores configurados no documento;
- inicia aprovação que aguarda todos os gestores;
- no aceite, move o arquivo para `PoliticasGrunner`, atualiza metadados, reúne comentários e notifica;
- na rejeição, registra status, motivo e avaliador, reúne comentários e notifica o responsável.

## Conectores e dependências

| Conector | Uso |
|---|---|
| SharePoint Online | Gatilhos, consultas, anexos, atualização, movimentação e metadados |
| Microsoft 365 Outlook | E-mails e caixa de correio compartilhada |
| Approvals | Aprovação de todos os gestores no ciclo documental |

Não foram observados conectores premium nos pacotes. Licenciamento continua excluído do custo porque depende do contrato Microsoft 365 da empresa.

## Precificação pelos mesmos cenários da apresentação

As horas são uma estimativa de reconstrução profissional. O cenário médio aplica fator 1,10 e o grande fator 1,35, mantendo os valores-hora já adotados na análise.

| Fluxo | Horas | Empresa pequena | Empresa média | Empresa grande |
|---|---:|---:|---:|---:|
| PA01 — Alerta de Notícias | 48 | R$ 6.000 | R$ 10.824 | R$ 22.032 |
| PA02 — Vencimento ISO | 32 | R$ 4.000 | R$ 7.216 | R$ 14.688 |
| PA03 — Resumo semanal | 44 | R$ 5.500 | R$ 9.922 | R$ 20.196 |
| PA04 — Novo rascunho | 12 | R$ 1.500 | R$ 2.706 | R$ 5.508 |
| PA05 — Ciclo documental | 64 | R$ 8.000 | R$ 14.432 | R$ 29.376 |
| Trabalho compartilhado | 40 | R$ 5.000 | R$ 9.020 | R$ 18.360 |
| **Total Power Automate** | **240** | **R$ 30.000** | **R$ 54.120** | **R$ 110.160** |

## Avaliação técnica dos fluxos

**Maturidade estimada: 2,5/5.** A automação funcional é relevante e o ciclo documental é complexo, mas ainda faltam controles para continuidade empresarial.

Pontos positivos:

- cinco entregas reais e conectadas aos processos da intranet;
- tratamento de anexos, ramificações e aprovações múltiplas;
- uso de caixa compartilhada e metadados SharePoint;
- pacotes exportáveis disponíveis.

Lacunas encontradas:

- todos os fluxos estão exportados como não gerenciados;
- a assinatura de alerta de falha aparece desativada na propriedade operacional dos cinco pacotes;
- não há escopos padronizados de tentativa, captura e encerramento;
- existem destinatários individuais e URLs fixas nas definições;
- conexões estão embutidas e precisam ser parametrizadas por ambiente;
- duas definições apresentam `contentVersion` sem versão válida;
- o e-mail do ramo aprovado do ciclo documental ainda possui assunto de teste;
- o status gravado após aprovação deve ser validado com a Qualidade;
- não há evidência de testes automatizados, pipeline, métricas ou revisão periódica.

## Requisitos para 5/5

- [ ] Levar os fluxos para uma Solution com connection references e environment variables.
- [ ] Substituir destinatários individuais por grupos/configuração.
- [ ] Padronizar scopes de sucesso, falha e timeout.
- [ ] Habilitar alertas e painel de execuções/erros.
- [ ] Definir proprietários e co-proprietários de cada fluxo.
- [ ] Criar casos de teste para aprovação, rejeição, ausência de anexos e falha de conexão.
- [ ] Validar assunto, estados e destinatários com os donos dos processos.
- [ ] Exportar versão gerenciada por release e guardar o pacote no repositório de artefatos.
- [ ] Documentar rollback e executar restauração em ambiente de teste.

