# Inventário de ambiente e dependências

Este documento registra tudo o que precisa existir fora do repositório para a Intranet Grunner funcionar. Os itens marcados como **PENDENTE** exigem confirmação do responsável pelo Microsoft 365 ou pela infraestrutura.

## Identificação

| Item | Valor conhecido |
|---|---|
| Tenant SharePoint | `grunnerteccombr.sharepoint.com` |
| Site principal | `https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner` |
| Solução SPFx | `intranet-grunnerv-1-client-side-solution` |
| ID da solução | `46cfb39b-a6aa-4103-9a23-9aaf88165a6a` |
| Versão do pacote analisada | `1.0.0.14` |
| Pacote gerado | `sharepoint/solution/intranet-grunnerv-1.sppkg` |
| Feature ID | `1837c82e-f12e-407b-ad12-08be7f921362` |
| App Catalog usado | **PENDENTE** |
| URL do repositório remoto | **PENDENTE** |
| Branch protegida de produção | **PENDENTE** |

## Páginas SharePoint

Validar em produção se os web parts abaixo estão publicados nas páginas indicadas.

| Página | Web part principal | ID do componente |
|---|---|---|
| `Inicio.aspx` | Home Grunner | `6a418d3b-a153-479f-ada1-c113ff679dfe` |
| `centraldeatalhos.aspx` | Central de Atalhos | `ea7d911e-2241-4676-a37f-e4f14c58d75e` |
| `Historia.aspx` | História | `c138e08d-5b9f-42de-a59d-f283c5ef84b4` |
| `Políticas-da-Empresa.aspx` | Políticas Grunner | `076d85df-9a04-4af6-97cc-35cf923c58d2` |
| `GerenciamentoDeAtivos.aspx` | Painel de Ativos | `008ad6e2-11b0-4604-9327-097db2c16440` |

Os manifests também declaram suporte a Teams Personal App, Teams Tab e SharePoint Full Page. Esse suporte deve ser testado antes de ser declarado como canal oficialmente atendido.

## Microsoft Graph

| Item | Uso | Requisito |
|---|---|---|
| `/me` | Perfil, cargo e departamento do usuário | Usuário autenticado |
| `/users` | Diretório e aniversariantes | Permissão delegada `User.Read.All` |
| `onPremisesExtensionAttributes.extensionAttribute1` | Aniversário no formato `DD/MM` | Sincronização correta do AD |
| `onPremisesExtensionAttributes.extensionAttribute10` | Data de empresa no formato `DD/MM/AAAA` | Sincronização correta do AD |

O pedido de permissão `User.Read.All` consta em `config/package-solution.json`. Confirmar a aprovação no centro de administração do SharePoint e revisar anualmente se o escopo continua necessário.

## Listas e bibliotecas SharePoint

Foram identificados **12 recursos**: oito listas, duas bibliotecas operacionais e duas bibliotecas de modelos.

| Recurso | Tipo esperado | Função | Dono |
|---|---|---|---|
| `NoticiasGrunner` | Lista | Notícias e comunicados | **PENDENTE** |
| `CurtidasGrunner` | Lista | Curtidas de notícias | **PENDENTE** |
| `ComentariosGrunner` | Lista | Comentários de notícias | **PENDENTE** |
| `AniversariantesGrunner` | Lista | Fallback de aniversariantes | **PENDENTE** |
| `EventosGrunner` | Lista | Eventos corporativos | **PENDENTE** |
| `LinksUteisGrunner` | Lista | Atalhos corporativos | **PENDENTE** |
| `PoliticasGrunner` | Biblioteca | Documentos publicados do SGQ | Qualidade — confirmar |
| `RascunhosSGQ` | Biblioteca | Fluxo de revisão do SGQ | Qualidade — confirmar |
| `Ativos de TI` | Lista | Inventário de ativos | TI — confirmar |
| `Acessos_Painel_Ativos` | Lista | Perfis de acesso ao painel | TI — confirmar |
| `Templates_SGQ` | Biblioteca | Três modelos de documentos SGQ | Qualidade — confirmar |
| `Modelos_TI` | Biblioteca | Modelo de termo de ativo | TI — confirmar |

O esquema observado no código está em [DATA-DICTIONARY.md](DATA-DICTIONARY.md). A confirmação definitiva deve ser feita exportando as configurações reais das listas.

## Grupos e regras de acesso

| Regra | Uso | Fonte |
|---|---|---|
| Grupo `Qualidade - Gestão de Documentos` | Funções administrativas do SGQ | Grupo do site SharePoint |
| Níveis `TI` e `Visualizador` | Painel de ativos | Lista `Acessos_Painel_Ativos` |
| Departamento/cargo com termo “marketing” | Administração de conteúdo da Home | Microsoft Graph `/me` |
| Exceção por e-mail no código | Administração de conteúdo | **Remover e substituir por grupo** |

As permissões efetivas das listas e bibliotecas continuam sendo controladas pelo SharePoint. A interface não é uma barreira de segurança.

## Arquivos modelo

| Caminho relativo ao tenant | Uso |
|---|---|
| `/sites/IntranetGrunner/Modelos_TI/Molde_Termo_Grunner.docx` | Termo de responsabilidade de ativo |
| `/sites/IntranetGrunner/Templates_SGQ/Template - Procedimento.docx` | Documento de procedimento |
| `/sites/IntranetGrunner/Templates_SGQ/Template - Mapeamento de Processo.docx` | Mapeamento de processo |
| `/sites/IntranetGrunner/Templates_SGQ/Template - Instrução de Trabalho.docx` | Instrução de trabalho |

Esses arquivos precisam ser incluídos no backup e testados após restauração.

## Serviço externo de chamados

Base atual: `https://admin.grunnertec.com.br/api/clickup`.

| Rota | Método observado | Função | Contrato/SLA |
|---|---|---|---|
| `/meus-chamados?email=...` | GET | Consultar chamados do usuário | **PENDENTE** |
| `/comentarios?idChamado=...` | GET | Consultar conversa | **PENDENTE** |
| `/comentar` | POST | Adicionar comentário | **PENDENTE** |

O backend não está neste repositório. Registrar repositório, proprietário, segredo/configuração, ambiente, SLA, monitoramento e procedimento de contingência.

## Power Automate

Cinco pacotes foram analisados: alerta de notícias, alerta de vencimento ISO, resumo semanal, notificação de novo rascunho e ciclo de vida documental. O inventário detalhado, riscos e custo estão em [POWER-AUTOMATE.md](POWER-AUTOMATE.md).

## Parâmetros de continuidade pendentes

| Parâmetro | Valor aprovado | Aprovador | Data |
|---|---|---|---|
| RTO — tempo máximo para restaurar | **PENDENTE** |  |  |
| RPO — perda máxima de dados | **PENDENTE** |  |  |
| Janela de manutenção | **PENDENTE** |  |  |
| Retenção de logs | **PENDENTE** |  |  |
| Retenção por lista/biblioteca | **PENDENTE** |  |  |
| Contatos de escalonamento | Ver [OWNERS.md](OWNERS.md) |  |  |

## Validação trimestral

- [ ] URLs e páginas continuam válidas.
- [ ] App Catalog, pacote ativo e versão foram conferidos.
- [ ] Permissões Graph foram revisadas.
- [ ] Listas, bibliotecas, campos e grupos foram comparados com o dicionário.
- [ ] Modelos Word foram abertos e usados em uma geração de teste.
- [ ] As três rotas de chamados responderam em cenário de teste.
- [ ] Donos, suplentes, RTO, RPO e contatos estão preenchidos.
