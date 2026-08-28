# ADR-0001 — SharePoint Framework e Microsoft 365

- Status: aceito como estado atual; revisar antes de migração relevante
- Data do registro: 27/08/2026
- Decisores originais: não identificados no repositório
- Responsável pela revisão: **PENDENTE**

## Contexto

A Intranet Grunner é uma ferramenta interna para comunicação, documentos do SGQ, atalhos, pessoas, chamados e ativos. A organização já utiliza SharePoint Online e identidades Microsoft 365. O sistema analisado foi implementado como cinco web parts SPFx em um único pacote.

## Decisão observada

Manter a experiência dentro do SharePoint Online, usando:

- SharePoint Framework como hospedagem e ciclo de implantação;
- React e TypeScript no frontend;
- PnPjs/REST para listas, bibliotecas, grupos e arquivos;
- Microsoft Graph para perfil e diretório;
- serviço externo para integração com ClickUp;
- documentos Word como modelos operacionais de SGQ e ativos.

Esta ADR documenta uma decisão já materializada no código; não afirma que houve um processo formal anterior.

## Consequências positivas

- Autenticação corporativa e contexto do usuário já disponíveis.
- Conteúdo e documentos próximos dos processos existentes no SharePoint.
- Implantação centralizada pelo App Catalog.
- Uso de grupos, permissões e histórico do ecossistema Microsoft 365.

## Custos e riscos

- Forte dependência de nomes internos, URLs, grupos e estrutura do tenant.
- Permissões Graph e ACLs SharePoint exigem governança fora do código.
- Atualizações SPFx/Node dependem da matriz de suporte da Microsoft.
- Um único pacote reúne módulos com donos e criticidades diferentes.
- O backend de chamados tem ciclo de vida próprio e não está neste repositório.
- Lógica de autorização no frontend não substitui segurança do dado.

## Regras derivadas

- Configuração por ambiente deve substituir URLs e identidades hardcoded.
- Mudança estrutural em SharePoint exige atualização do inventário/dicionário e plano de migração.
- Permissões devem seguir menor privilégio e ser testadas no destino.
- Releases precisam de pacote versionado, hash, evidência e rollback.
- Integrações externas precisam de contrato, proprietário, SLA e contingência.

## Alternativas a avaliar no futuro

- Separar módulos críticos em soluções SPFx independentes.
- Usar API intermediária para operações que exigem autorização e auditoria no servidor.
- Migrar conteúdo/integrações específicas para Power Platform ou serviços dedicados quando houver benefício comprovado.
- Adotar aplicação independente somente se requisitos de experiência, disponibilidade ou segurança superarem a vantagem de integração nativa.

## Gatilhos para revisão

- Fim de suporte da versão SPFx/Node.
- Aumento relevante de volume, criticidade ou indisponibilidade.
- Requisito de acesso externo/móvel não atendido pelo SharePoint.
- Necessidade de segregar ciclos de release ou donos por módulo.
- Incidente de segurança ligado às limitações da arquitetura atual.

