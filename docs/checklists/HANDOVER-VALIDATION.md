# Validação independente de continuidade

Este checklist deve ser executado por alguém que não seja o mantenedor principal. Orientações dadas durante o teste devem ser registradas como lacunas da documentação.

## Identificação

```text
Data:
Participante:
Função/equipe:
Mantenedor observado:
Ambiente de teste:
Versão/commit:
Tempo total:
```

## Parte 1 — descoberta

- [ ] Explicou o objetivo e os seis blocos funcionais sem ajuda.
- [ ] Localizou os cinco web parts e as páginas.
- [ ] Identificou SharePoint, Graph e backend de chamados como dependências.
- [ ] Localizou donos, suplentes e canais de escalonamento.
- [ ] Identificou dados pessoais e controles de acesso.

## Parte 2 — desenvolvimento

- [ ] Preparou uma máquina/pasta limpa com a versão correta de Node.
- [ ] Executou `npm ci` sem instrução adicional.
- [ ] Iniciou o ambiente local e localizou o workbench.
- [ ] Executou build limpo.
- [ ] Interpretou corretamente uma falha de build simulada ou real.

## Parte 3 — ambiente e dados

- [ ] Conferiu uma lista real contra o dicionário de dados.
- [ ] Localizou os grupos/perfis de acesso.
- [ ] Validou um modelo Word.
- [ ] Identificou a permissão Graph e sua justificativa.
- [ ] Validou o contrato e o dono do backend de chamados.

## Parte 4 — release e operação

- [ ] Gerou pacote versionado e hash em ambiente de teste.
- [ ] Publicou ou simulou a publicação pelo procedimento oficial.
- [ ] Executou smoke test das cinco páginas.
- [ ] Registrou e classificou um incidente simulado.
- [ ] Executou ou explicou rollback com o artefato anterior.

## Parte 5 — recuperação

- [ ] Restaurou uma lista/arquivo de teste com metadados.
- [ ] Restaurou/configurou um web part em página de teste.
- [ ] Testou um acesso autorizado e um negado.
- [ ] Registrou os tempos e comparou com RTO/RPO.
- [ ] Acionou o escalonamento correto sem depender do autor.

## Resultado

| Critério | Resultado | Evidência/lacuna |
|---|---|---|
| Compreensão | Aprovado / Reprovado |  |
| Desenvolvimento | Aprovado / Reprovado |  |
| Ambiente e dados | Aprovado / Reprovado |  |
| Release e operação | Aprovado / Reprovado |  |
| Recuperação | Aprovado / Reprovado |  |

```text
Resultado geral:
Orientações que o autor precisou fornecer:
Ações corretivas:
Responsáveis e prazos:
Nova data de validação:
Aceite do participante:
Aceite do responsável técnico:
```

A maturidade 5/5 requer aprovação geral, restauração dentro dos objetivos aprovados e ausência de dependência operacional do mantenedor principal.

