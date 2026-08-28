# LGPD e governança de dados

Este documento é um inventário técnico inicial, não um parecer jurídico. Base legal, retenção e comunicação de incidente devem ser aprovadas pelo encarregado/DPO e pelos donos de negócio.

## Dados pessoais observados

| Categoria | Exemplos | Fonte | Uso funcional | Retenção aprovada |
|---|---|---|---|---|
| Identificação corporativa | Nome, e-mail | Microsoft Graph/SharePoint | Perfil, autoria e acesso | **PENDENTE** |
| Vínculo profissional | Cargo, departamento, data de empresa | Microsoft Graph | Diretório, celebrações e regra funcional | **PENDENTE** |
| Data de aniversário | Dia e mês | Graph/lista de fallback | Aniversariantes | **PENDENTE** |
| Interações internas | Curtidas e comentários | SharePoint | Engajamento em notícias | **PENDENTE** |
| Suporte | E-mail, chamado e comentários | Backend ClickUp | Atendimento ao colaborador | **PENDENTE** |
| Responsabilidade patrimonial | Responsável e ativo associado | SharePoint | Gestão de ativos | **PENDENTE** |
| Aprovação documental | Responsáveis, avaliadores e aprovadores | SharePoint | Governança do SGQ | **PENDENTE** |

## Decisões que precisam de aprovação

Para cada categoria, registrar:

- controlador, operador e dono interno;
- finalidade específica e base legal;
- origem e compartilhamentos;
- usuários/grupos autorizados;
- prazo e forma de descarte;
- necessidade de relatório de impacto;
- procedimento para acesso, correção, oposição ou eliminação aplicável.

## Regras de implementação

- Coletar apenas campos necessários à função exibida.
- Não usar e-mail hardcoded como autorização.
- Não registrar respostas completas do Graph ou chamados no console de produção.
- Não usar dados reais em documentação, testes ou capturas de evidência.
- Sanitizar conteúdo editável antes de renderizar HTML.
- Restringir listas e bibliotecas com menor privilégio.
- Aplicar retenção e descarte também a anexos, versões e lixeira do SharePoint.
- Formalizar o tratamento realizado pelo backend de chamados, que está fora deste repositório.

## Atendimento ao titular

1. Receber a solicitação pelo canal corporativo definido pelo DPO.
2. Validar identidade e escopo sem coletar dados excessivos.
3. Consultar Graph, listas/bibliotecas SharePoint e backend de chamados.
4. Encaminhar aos donos responsáveis, preservando rastreabilidade.
5. Responder dentro do prazo aplicável e registrar a conclusão.
6. Executar correção/eliminação somente com autorização e respeitando obrigações de retenção.

## Incidente envolvendo dados pessoais

1. Conter o acesso ou integração afetada.
2. Preservar logs e evidências com acesso restrito.
3. Identificar dados, titulares, volume, período e consequências.
4. Acionar segurança, DPO, jurídico e direção conforme a matriz corporativa.
5. Avaliar notificações obrigatórias sem prometer prazo por conta própria.
6. Documentar correção, monitoramento e prevenção.

## Evidências para maturidade 5/5

- [ ] Inventário acima aprovado por DPO/dono dos dados.
- [ ] Base legal e retenção preenchidas para todas as categorias.
- [ ] Regra de descarte implementada e testada.
- [ ] Contrato do backend de chamados inclui privacidade e segurança.
- [ ] Revisão de acessos vigente.
- [ ] Simulado de solicitação de titular executado.
- [ ] Simulado de incidente executado e ações concluídas.

