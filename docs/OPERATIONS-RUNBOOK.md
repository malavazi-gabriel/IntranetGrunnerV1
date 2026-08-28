# Runbook de operação

Este é o roteiro de suporte para manter a Intranet Grunner disponível e orientar a primeira resposta a incidentes. Ele deve ser usado junto com [TROUBLESHOOTING.md](TROUBLESHOOTING.md) e [DISASTER-RECOVERY.md](DISASTER-RECOVERY.md).

## Rotina operacional

### Diariamente ou após alerta

- Abrir as cinco páginas principais em uma sessão corporativa comum.
- Confirmar carregamento sem erro visível e sem falhas críticas no console do navegador.
- Na Home, verificar notícias, perfil, eventos e chamados.
- No SGQ, abrir pelo menos um documento publicado.
- No painel de ativos, consultar um item com um usuário autorizado.
- Registrar incidente quando uma dependência externa estiver indisponível.

### Semanalmente

- Verificar mensagens de erro e chamados recorrentes por módulo.
- Testar uma consulta ao Microsoft Graph e as três rotas de chamados.
- Conferir se documentos próximos do vencimento seguem o fluxo da Qualidade.
- Verificar alterações recentes em grupos, acessos do painel e administradores de conteúdo.

### Mensalmente

- Conferir a versão implantada no App Catalog com a versão registrada na release.
- Revisar o crescimento das listas e bibliotecas e os limites do SharePoint.
- Testar a abertura dos quatro modelos Word.
- Revisar dependências com vulnerabilidades e atualizações suportadas.
- Confirmar se os donos e suplentes de [OWNERS.md](OWNERS.md) continuam válidos.

### Trimestralmente

- Executar revisão de acessos e guardar a evidência.
- Validar o inventário de ambiente e o dicionário de dados contra a produção.
- Realizar o teste de restauração descrito em [DISASTER-RECOVERY.md](DISASTER-RECOVERY.md).
- Pedir que uma pessoa fora da manutenção habitual execute o checklist de handover.

## Matriz de triagem

| Sintoma | Verificação inicial | Equipe primária | Escalonar para |
|---|---|---|---|
| Toda a intranet indisponível | Saúde Microsoft 365, site e pacote | TI/M365 | Microsoft/fornecedor M365 |
| Apenas uma página falha | Web part, console e recursos daquela página | Desenvolvimento | Dono funcional |
| Perfil/aniversariantes falham | Graph, consentimento e atributos do AD | Identidade/TI | Desenvolvimento |
| Notícias/eventos/atalhos não aparecem | Lista, filtros, permissões e status | Marketing/Comunicação | SharePoint/Desenvolvimento |
| Políticas não aparecem | Biblioteca, metadados, grupo Qualidade | Qualidade | SharePoint/Desenvolvimento |
| Ativos não aparecem | Lista, campos internos e acesso | TI | SharePoint/Desenvolvimento |
| Chamados não carregam | API externa, rede e resposta HTTP | Dono do backend ClickUp | Desenvolvimento/Infra |
| Documento Word não é gerado | Modelo, caminho e campos | TI ou Qualidade | Desenvolvimento |

## Classificação e resposta

| Severidade | Critério | Início da resposta | Comunicação |
|---|---|---|---|
| S1 crítica | Indisponibilidade geral ou risco de dados/segurança | Imediato, conforme SLA aprovado | Diretoria, TI e donos afetados |
| S2 alta | Módulo crítico indisponível sem contorno | Conforme SLA aprovado | Dono funcional e TI |
| S3 média | Falha parcial com contorno | Próximo horário útil | Dono funcional |
| S4 baixa | Dúvida, melhoria ou defeito cosmético | Backlog | Solicitante |

Os tempos exatos dependem do SLA ainda **PENDENTE** no inventário. Não prometer prazo até ele ser aprovado.

## Registro mínimo de incidente

```text
ID:
Data/hora e fuso:
Relator e contato:
Página/módulo:
Usuários afetados:
Severidade:
Sintoma e mensagem exata:
URL e versão implantada:
Passos para reproduzir:
Evidências sem dados pessoais:
Dependência envolvida:
Ação de contenção:
Causa raiz:
Correção e validação:
Responsável e aprovador do encerramento:
```

## Procedimento de publicação emergencial

1. Confirmar severidade e autorização do responsável de release.
2. Reproduzir e registrar a causa; não corrigir diretamente em produção.
3. Criar mudança mínima em branch identificada.
4. Executar build, análise estática e testes disponíveis.
5. Atualizar versão do pacote e notas de release.
6. Gerar o `.sppkg` e guardar hash do artefato.
7. Implantar seguindo [DEPLOYMENT.md](DEPLOYMENT.md).
8. Executar smoke test das cinco páginas.
9. Registrar quem aprovou, quem implantou e o resultado.
10. Programar revisão pós-incidente e correção estrutural.

## Encerramento de incidente

O incidente só é encerrado quando:

- o serviço foi restaurado e validado por usuário representativo;
- causa, impacto e linha do tempo foram registrados;
- ações preventivas têm responsável e prazo;
- houve revisão de segurança/LGPD quando aplicável;
- documentação e monitoramento foram atualizados no mesmo ciclo.

