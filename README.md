# Intranet Grunner

Intranet corporativa desenvolvida com SharePoint Framework para centralizar comunicação interna, pessoas e eventos, chamados, atalhos, documentos do SGQ e gestão de ativos.

## Estado atual

- Plataforma: SharePoint Online e Microsoft 365
- SharePoint Framework: 1.22.2
- React: 17.0.1
- TypeScript: 5.8
- Node.js: 22.14 ou superior e inferior a 23
- Web parts: 5
- Pacote SharePoint declarado: 1.0.0.14
- Permissão Microsoft Graph solicitada: `User.Read.All`
- Situação do build em 27/08/2026: falha conhecida, consulte [Problemas conhecidos](docs/TROUBLESHOOTING.md)

## Funcionalidades

| Área | Entrega principal |
|---|---|
| Home | Notícias, rascunhos, vídeos, curtidas, comentários, aniversários, tempo de empresa e eventos |
| Chamados | Consulta, notificações, histórico e comentários integrados ao serviço administrativo e ClickUp |
| Atalhos | Catálogo de links e acesso a solicitações corporativas |
| História | Conteúdo institucional da empresa |
| Políticas e SGQ | Busca, histórico, geração de Word, rascunhos, aprovação, revisão e obsolescência |
| Ativos | Cadastro, transferência, termos de responsabilidade, auditoria e exportação Excel |
| Acessos | Perfis de TI, visualizador, Qualidade e Marketing |
| Automações | Alertas de notícias, resumo semanal, vencimento ISO, novos rascunhos e ciclo de aprovação documental |

## Comece por aqui

1. Leia o [índice da documentação](docs/INDEX.md).
2. Confira os [pré-requisitos e preparação local](docs/DEVELOPMENT.md).
3. Entenda a [arquitetura](docs/ARCHITECTURE.md).
4. Valide as [dependências do ambiente](docs/ENVIRONMENT-INVENTORY.md).
5. Consulte o [inventário e custo dos fluxos Power Automate](docs/POWER-AUTOMATE.md).
6. Para publicar, siga o [procedimento de implantação](docs/DEPLOYMENT.md).
7. Para incidentes, use o [manual operacional](docs/OPERATIONS-RUNBOOK.md) e a [solução de problemas](docs/TROUBLESHOOTING.md).

## Comandos

```powershell
npm ci
npm start
npm run build
```

O comando `npm run build` não deve ser usado para uma publicação enquanto as falhas registradas em [Problemas conhecidos](docs/TROUBLESHOOTING.md#build-falha-com-módulos-sass-ausentes) não forem corrigidas e o checklist de liberação não estiver aprovado.

## Estrutura principal

```text
config/                         Configuração do build e pacote SharePoint
src/shared/                     Componentes reutilizados entre páginas
src/webparts/homeGrunner/       Página inicial e comunicação
src/webparts/centralAtalhos.../ Central de atalhos
src/webparts/historiaGrunner/   História institucional
src/webparts/politicasGrunner/  Políticas e gestão do SGQ
src/webparts/painelAtivos.../   Gestão de ativos
sharepoint/solution/            Pacote SharePoint gerado localmente
teams/                          Ícones de integração com Teams
docs/                           Arquitetura, operação e continuidade
```

## Regras de manutenção

- Nenhuma publicação sem build aprovado.
- Toda mudança funcional deve atualizar a documentação relacionada no mesmo commit.
- Alterações de lista, biblioteca, grupo, permissão ou template devem atualizar `docs/ENVIRONMENT-INVENTORY.md` e `docs/DATA-DICTIONARY.md`.
- Alterações arquiteturais devem criar ou atualizar um ADR em `docs/decisions/`.
- Credenciais, tokens e segredos não podem ser armazenados no repositório.
- Endereços e e-mails individuais fixos devem ser substituídos por configuração antes de uma classificação técnica 5/5.

## Responsabilidade

Os responsáveis técnicos, funcionais e de operação devem ser preenchidos em [OWNERS.md](docs/OWNERS.md). A documentação não substitui as permissões reais do SharePoint ou a aprovação dos donos dos processos.
