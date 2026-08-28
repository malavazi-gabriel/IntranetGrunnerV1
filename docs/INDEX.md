# Índice da documentação

## Leitura por objetivo

### Entender o sistema

- [Arquitetura](ARCHITECTURE.md)
- [Inventário do ambiente](ENVIRONMENT-INVENTORY.md)
- [Dicionário de dados](DATA-DICTIONARY.md)
- [Power Automate: inventário, qualidade e custo](POWER-AUTOMATE.md)
- [Decisão de usar SharePoint Framework](decisions/ADR-0001-SPFX-E-MICROSOFT-365.md)

### Desenvolver

- [Preparação do ambiente de desenvolvimento](DEVELOPMENT.md)
- [Problemas conhecidos e solução de problemas](TROUBLESHOOTING.md)
- [Checklist de liberação](checklists/RELEASE.md)

### Publicar e operar

- [Implantação e rollback](DEPLOYMENT.md)
- [Manual operacional](OPERATIONS-RUNBOOK.md)
- [Recuperação de desastre](DISASTER-RECOVERY.md)
- [Segurança e acessos](SECURITY-ACCESS.md)

### Governança e continuidade

- [LGPD e governança de dados](LGPD-DATA-GOVERNANCE.md)
- [Responsáveis](OWNERS.md)
- [Continuidade e transferência de conhecimento](CONTINUITY-HANDOVER.md)
- [Evidências para nota 5/5](CONTINUITY-EVIDENCE.md)
- [Checklist de validação por outra pessoa](checklists/HANDOVER-VALIDATION.md)

## Estado dos documentos

| Documento | Fonte principal | Situação inicial |
|---|---|---|
| Arquitetura | Código e configuração | Preenchido com fatos do repositório |
| Desenvolvimento | `package.json`, `.yo-rc.json`, `config/serve.json` | Preenchido, build com falha conhecida |
| Implantação | Configuração SPFx | Procedimento base preenchido, catálogo real pendente |
| Ambiente | Código, URLs e nomes de recursos | Preenchido, permissões reais pendentes |
| Dados | Consultas e cargas presentes no código | Preenchido, tipos e retenção pendentes de validação |
| Power Automate | Cinco pacotes exportados em 27/08/2026 | Medido; implantação gerenciada e monitoramento pendentes |
| Operação | Fluxos implementados e falhas tratadas | Procedimentos base preenchidos |
| Segurança | Grupos, listas e Graph no código | Modelo lógico preenchido, ACL real pendente |
| LGPD | Campos processados pelo código | Inventário inicial, validação jurídica pendente |
| Recuperação | Git, pacote e SharePoint | Plano criado, exercício ainda necessário |
| Continuidade | Estado do repositório | Plano e critérios criados, validação independente necessária |

## Regra de atualização

O responsável por uma alteração deve atualizar o documento correspondente no mesmo pull request ou commit. Se não houver impacto documental, o revisor deve registrar essa conclusão no checklist de liberação.
