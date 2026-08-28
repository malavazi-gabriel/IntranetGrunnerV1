# Evidência e nota de documentação/continuidade

## Escala objetiva de 0 a 5

| Nota | Evidência mínima |
|---|---|
| 0 | Sem documentação utilizável ou conhecimento exclusivamente informal. |
| 1 | Histórico e dependências básicas existem, mas não há guia operacional nem transferência. |
| 2 | README e arquitetura parcial; instalação/operação ainda dependem do autor. |
| 3 | Pacote documental completo no repositório, porém ainda não validado por terceiro ou com build/ambiente pendentes. |
| 4 | Documentação validada por outro técnico; build, publicação e operação reproduzíveis; donos e rotinas definidos. |
| 5 | Continuidade comprovada: suplência treinada, restauração testada dentro de RTO/RPO, revisão periódica e evidência auditável. |

## Situação após a criação deste pacote

**Estimativa documental: 3,5/5**, condicionada à incorporação no repositório.

O conteúdo necessário foi estruturado, mas não é honesto atribuir 5/5 enquanto:

- o build oficial ainda falhar;
- nomes de donos, suplentes e contatos estiverem pendentes;
- RTO, RPO, retenção e SLA não estiverem aprovados;
- inventário e campos não tiverem sido confrontados com produção;
- uma segunda pessoa não tiver publicado e operado o sistema;
- um exercício de restauração não tiver sido concluído.

## Evidências obrigatórias para 5/5

| Controle | Evidência | Resultado | Data | Responsável/aprovador |
|---|---|---|---|---|
| Documentação versionada | Pull request com pacote completo | **PENDENTE** |  |  |
| Build reproduzível | Log limpo em máquina/pipeline novo | **PENDENTE** |  |  |
| Testes automatizados | Pipeline e relatório de cobertura mínima aprovada | **PENDENTE** |  |  |
| Ambiente inventariado | Comparação assinada com produção | **PENDENTE** |  |  |
| Donos e suplentes | `OWNERS.md` preenchido e revisado | **PENDENTE** |  |  |
| Segurança | Revisão de acesso e testes negativos | **PENDENTE** |  |  |
| Privacidade | Inventário/base legal/retenção aprovados | **PENDENTE** |  |  |
| Release | Publicação e rollback em ambiente de teste | **PENDENTE** |  |  |
| Operação | Incidente simulado resolvido pelo suplente | **PENDENTE** |  |  |
| Recuperação | Exercício dentro de RTO/RPO | **PENDENTE** |  |  |
| Continuidade | Checklist de handover aprovado por terceiro | **PENDENTE** |  |  |
| Recorrência | Agenda trimestral e última revisão registradas | **PENDENTE** |  |  |

## Regra para alteração da nota na apresentação

A nota só deve subir após anexar evidência nesta matriz. Criar documentos aumenta a prontidão, mas não prova execução. Para 5/5, todos os controles acima precisam estar aprovados ou uma exceção formal deve registrar risco, dono e prazo.

## Reavaliação

- Frequência: trimestral e após mudança relevante.
- Avaliador: pessoa diferente do mantenedor principal.
- Saída: nota, evidências, lacunas, ações, responsáveis e prazos.
- Destino: repositório ou sistema corporativo de evidências, sem dados pessoais desnecessários.

