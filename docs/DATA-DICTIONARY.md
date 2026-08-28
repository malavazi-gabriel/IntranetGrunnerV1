# Dicionário de dados e integrações

Este dicionário foi derivado do código-fonte. Ele descreve os nomes internos que a aplicação espera, mas **não substitui** a exportação do esquema real do SharePoint. Tipo, obrigatoriedade, valor padrão, índice e unicidade devem ser confirmados no ambiente.

## Notícias — `NoticiasGrunner`

| Campo interno | Finalidade observada |
|---|---|
| `ID` | Identificador da notícia |
| `Title` | Título |
| `Resumo` | Resumo de exibição |
| `ImagemURL` | Imagem de capa |
| `VideoURL` | Vídeo associado |
| `LinkNoticia` | Link externo |
| `ConteudoNoticia` | Conteúdo completo |
| `StatusNoticia` | Estado de publicação |
| `Attachments` / `AttachmentFiles` | Anexos |
| `TotalCurtidas` | Total apresentado de curtidas |
| `TotalComentarios` | Total apresentado de comentários |

## Curtidas — `CurtidasGrunner`

| Campo interno | Finalidade observada |
|---|---|
| `ID` | Identificador |
| `Title` | Título técnico/legado |
| `NoticiaID` | Referência lógica à notícia |
| `UsuarioEmail` | Identidade do usuário |
| `UsuarioNome` | Nome para exibição |

## Comentários — `ComentariosGrunner`

| Campo interno | Finalidade observada |
|---|---|
| `ID` | Identificador |
| `Title` | Título técnico/legado |
| `NoticiaID` | Referência lógica à notícia |
| `Autor` | Autor do comentário |
| `Comentario` | Texto informado |
| `Created` | Data de criação do SharePoint |

## Eventos — `EventosGrunner`

| Campo interno | Finalidade observada |
|---|---|
| `Title` | Nome do evento |
| `Dia` | Dia do mês |
| `Mes` | Mês |
| `Local` | Local do evento |
| `ImagemTema` | Imagem de apresentação |

## Aniversariantes — `AniversariantesGrunner`

Fonte de fallback; a fonte principal observada é o Microsoft Graph.

| Campo interno | Finalidade observada |
|---|---|
| `Title` | Nome |
| `Dia` | Data/dia do aniversário |
| `Setor` | Departamento |
| `Email` | Identidade corporativa |

## Links úteis — `LinksUteisGrunner`

| Campo interno | Finalidade observada |
|---|---|
| `ID` | Identificador |
| `Title` | Nome do atalho |
| `Descricao` | Texto explicativo |
| `Categoria` | Agrupamento |
| `Icone` | Ícone |
| `LinkURL` | Destino |
| `Ordem` | Ordenação |
| `Ativo` | Exibição habilitada |

## Ativos — `Ativos de TI`

Alguns campos usam nomes internos gerados pelo SharePoint. Não renomear nem recriar sem atualizar o código.

| Campo interno | Nome funcional observado |
|---|---|
| `Title` | Nome/título do ativo |
| `field_4` | Identificador do ativo |
| `field_5` | Número do ativo financeiro |
| `field_9` | IMEI |
| `field_10` | Especificações |
| `Responsavel_AD` | Responsável |
| `field_1` | Departamento |
| `field_2` | Tipo |
| `field_3` | Prefixo |
| `field_6` | Fabricante |
| `field_7` | Modelo |
| `field_8` | Número de série |
| `field_11` | Observações |

## Acessos do painel — `Acessos_Painel_Ativos`

| Campo interno | Finalidade observada | Valores esperados |
|---|---|---|
| `ID` | Identificador | Número do SharePoint |
| `Title` | Nome/título | Texto |
| `Email` | Usuário autorizado | E-mail corporativo |
| `NivelAcesso` | Perfil funcional | `TI` ou `Visualizador` |

## Políticas publicadas — `PoliticasGrunner`

| Campo interno | Finalidade observada |
|---|---|
| `Id` / `UniqueId` | Identificadores |
| `FileLeafRef` / `FileRef` | Nome e caminho do arquivo |
| `Area` | Área responsável |
| `CodigoDocumento` | Código de controle |
| `TipoDocumento` | Tipo documental |
| `NumeroRevisao` | Revisão vigente |
| `DataUltimaRevisao` | Última revisão |
| `DataProximaRevisao` | Próxima revisão prevista |
| `StatusDocumento` | Estado do documento |
| `ObservacaoRevisao` | Observação de revisão |
| `PeriodicidadeRevisaoMeses` | Ciclo de revisão |
| `UltimoAvisoRevisao` | Último aviso emitido |
| `DiasAvisoRevisao` | Antecedência de aviso |
| `PermiteImpressaoControlada` | Regra de impressão |
| `ExibirNaIntranet` | Visibilidade |
| `ResponsavelRevisao` | Responsável funcional |
| `AprovadorQualidade` | Aprovador da Qualidade |
| `TipoProcessoDocumento` | Classificação de processo |
| `DocumentoControlado` | Indicador de controle |
| `AprovadoresdoDocumento` | Aprovadores |
| `ProcessoExtinto` | Indicador de processo encerrado |

## Rascunhos — `RascunhosSGQ`

| Campo interno | Finalidade observada |
|---|---|
| `FileLeafRef` | Nome do documento |
| `StatusdaRevisao` | Etapa/resultado da revisão |
| `AprovadoresdoDocumento` | Pessoas aprovadoras |
| `MotivoRejeicao` | Justificativa da rejeição |
| `Avaliador` | Pessoa que avaliou |

## Microsoft Graph

| Propriedade | Uso |
|---|---|
| `displayName` | Nome do colaborador |
| `mail` | Identidade e contato |
| `jobTitle` | Cargo e regra funcional |
| `department` | Departamento e regra funcional |
| `accountEnabled` | Filtrar contas ativas |
| `onPremisesExtensionAttributes.extensionAttribute1` | Aniversário `DD/MM` |
| `onPremisesExtensionAttributes.extensionAttribute10` | Data de empresa `DD/MM/AAAA` |

## Serviço de chamados

O contrato de payload não está formalizado neste repositório. Antes de atribuir nota máxima de continuidade:

- versionar um OpenAPI/JSON Schema do serviço;
- documentar autenticação, limites, timeouts e códigos de erro;
- identificar o repositório e o proprietário do backend;
- incluir exemplos sem dados pessoais reais;
- definir SLA e comportamento da intranet quando o serviço estiver indisponível.

## Procedimento de confirmação

1. Exportar campos e configurações de cada lista/biblioteca no ambiente produtivo.
2. Comparar nome interno, tipo, obrigatoriedade, índice, unicidade e padrão.
3. Corrigir este documento ou abrir uma mudança de código.
4. Anexar o relatório da comparação à evidência de release.
5. Repetir após qualquer alteração estrutural no SharePoint.

