# Plano de continuidade e transferência de conhecimento

## Resultado esperado

Uma pessoa tecnicamente qualificada, sem participação prévia no projeto, deve conseguir compreender, instalar, compilar, publicar, operar, diagnosticar e restaurar o sistema usando apenas o repositório, acessos oficialmente concedidos e esta documentação.

## Pacote mínimo de handover

- [README](../README.md) e índice de documentação.
- Arquitetura, integrações e decisões registradas.
- Inventário do ambiente e dicionário confirmados contra produção.
- Runbook, troubleshooting, segurança e LGPD aprovados.
- Processo de desenvolvimento, release, rollback e recuperação.
- Donos e suplentes preenchidos.
- Backlog de riscos e dívida técnica acessível.
- Artefato aprovado, hash e histórico de releases.

## Roteiro de transferência

### Sessão 1 — contexto e arquitetura

- Objetivo de negócio e módulos.
- Limites entre SPFx, SharePoint, Graph e backend de chamados.
- Fluxos de dados, perfis e riscos conhecidos.
- Navegação pelo código e pelos documentos.

### Sessão 2 — desenvolvimento e qualidade

- Preparação de máquina limpa.
- `npm ci`, servidor local e build.
- Correção de defeito simples com revisão.
- Gates de teste, segurança e documentação.

### Sessão 3 — operação e incidente

- Rotinas do runbook.
- Diagnóstico de falha simulada de Graph, lista ou chamados.
- Comunicação, escalonamento e registro de causa.

### Sessão 4 — implantação e recuperação

- Geração e validação do pacote.
- Publicação em ambiente de teste.
- Rollback e restauração simulada.
- Aceite dos responsáveis.

## Critério de aprovação do handover

O handover passa somente se o novo mantenedor, sem orientação passo a passo do autor:

1. localizar os cinco módulos e suas dependências;
2. preparar o ambiente e obter build limpo;
3. publicar em ambiente de teste;
4. resolver ou diagnosticar um incidente simulado;
5. explicar os controles de acesso e dados pessoais;
6. restaurar ao menos um componente de teste;
7. encontrar responsáveis e escalar corretamente;
8. registrar uma pequena alteração com documentação e evidências.

Falhas geram ações com prazo; não são convertidas em aprovação verbal.

## Redução do risco de pessoa-chave

- No mínimo dois mantenedores habilitados para código e implantação.
- No mínimo dois administradores capazes de recuperar o site e o App Catalog.
- Segredos e acessos de emergência sob processo corporativo, não em posse individual.
- Mudanças relevantes revisadas por outra pessoa.
- Exercício trimestral alternando o executor.
- Documentação atualizada no mesmo pull request da mudança.

## Saída de mantenedor

1. Transferir tarefas, decisões, acessos e contatos pendentes.
2. Validar que o suplente consegue executar o checklist de handover.
3. Revogar acessos individuais conforme a política.
4. Trocar segredos somente quando o processo de segurança indicar.
5. Atualizar [OWNERS.md](OWNERS.md) e registrar a revisão.

## Registro de sessão

```text
Data:
Tema:
Instrutor:
Participante:
Ambiente:
Exercícios executados:
Resultado:
Dúvidas/lacunas:
Ações, responsáveis e prazos:
Aceite do participante:
Aceite do responsável técnico:
```

