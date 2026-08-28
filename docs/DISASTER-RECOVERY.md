# Recuperação de desastre

## Objetivo

Restaurar a Intranet Grunner de forma previsível, preservando código, pacote, configuração, conteúdo, permissões e integrações. Os valores de RTO e RPO permanecem **PENDENTES** de aprovação em [ENVIRONMENT-INVENTORY.md](ENVIRONMENT-INVENTORY.md).

## O que precisa ser recuperável

| Camada | Itens | Fonte/backup | Validação |
|---|---|---|---|
| Código | Repositório, tags e histórico | Git remoto — **PENDENTE** | Clone limpo e hash conferido |
| Artefato | `.sppkg` por release | Repositório de artefatos — **PENDENTE** | Hash e versão |
| Solução | Registro no App Catalog e permissões Graph | Microsoft 365 | Web parts disponíveis |
| Páginas | Cinco páginas e configuração dos web parts | SharePoint | Smoke test |
| Dados | 12 listas/bibliotecas conhecidas | Política de backup M365 — **PENDENTE** | Amostragem e contagens |
| Documentos | Arquivos e histórico de versões do SGQ | SharePoint | Documento e metadados |
| Modelos | Quatro modelos Word | SharePoint + cópia controlada | Geração de documento |
| Acessos | Grupos, ACLs e lista de perfis de ativos | Microsoft 365/SharePoint | Teste positivo e negativo |
| Identidade | Permissões Graph e atributos do AD | Entra ID/AD | `/me` e `/users` |
| Chamados | Backend, configuração e dados ClickUp | Equipe do serviço — **PENDENTE** | Três rotas funcionais |
| Automações | Cinco fluxos Power Automate, conexões e variáveis | Pacote por release — **PENDENTE** | Importação e execução controlada |

## Pré-requisitos

- Um titular e um suplente com acesso ao Git, Microsoft 365 e artefatos.
- Credenciais de emergência controladas pela política corporativa, nunca neste repositório.
- Exportação atualizada do esquema das listas/bibliotecas.
- Registro da versão produtiva e do hash do pacote.
- Backups e retenções aprovados pelos donos dos dados.
- Ambiente/local de teste para validar antes da publicação.

## Cenário A — rollback de aplicação

Use quando dados e SharePoint estão íntegros, mas a nova versão do frontend falhou.

1. Interromper novas implantações e registrar o incidente.
2. Identificar a última versão aprovada e seu hash.
3. Reimplantar o `.sppkg` anterior conforme política do App Catalog.
4. Confirmar permissões Graph e associação ao site.
5. Executar o smoke test das cinco páginas.
6. Obter aceite dos donos afetados e comunicar restauração.
7. Manter a versão defeituosa isolada para análise, sem apagar evidências.

## Cenário B — lista, biblioteca ou arquivo excluído/corrompido

1. Bloquear mudanças no recurso afetado.
2. Identificar horário, escopo, usuário e última versão íntegra.
3. Usar versionamento/lixeira/restauração do Microsoft 365 conforme a política vigente.
4. Restaurar esquema, itens, arquivos, metadados, versões e ACLs necessários.
5. Comparar contagens e amostra de registros com a evidência anterior.
6. Validar no módulo consumidor com perfis autorizado e não autorizado.
7. Avaliar impacto de privacidade e registrar perda dentro do RPO aprovado.

## Cenário C — reconstrução do site

1. Criar/recuperar o site no tenant correto.
2. Restaurar as 12 listas/bibliotecas com nomes internos compatíveis.
3. Restaurar os quatro modelos Word e grupos/ACLs.
4. Aprovar as permissões Graph necessárias.
5. Implantar o último pacote aprovado no App Catalog.
6. Recriar as cinco páginas e adicionar os respectivos web parts.
7. Validar o backend de chamados e os atributos de identidade.
8. Executar todos os casos do checklist de release.
9. Obter aceite de TI, Comunicação/Marketing e Qualidade.

## Cenário D — backend de chamados indisponível

1. Confirmar que a Intranet e o Microsoft 365 continuam disponíveis.
2. Acionar o proprietário do serviço ClickUp documentado em [OWNERS.md](OWNERS.md).
3. Comunicar indisponibilidade apenas do módulo e informar o canal alternativo aprovado.
4. Não implantar credencial, proxy ou endpoint temporário sem revisão de segurança.
5. Após retorno, testar consulta, conversa e comentário antes de encerrar.

## Teste trimestral de recuperação

O teste deve ocorrer em ambiente seguro e não destrutivo.

- [ ] Participante que não mantém o sistema recebeu apenas este repositório e acessos aprovados.
- [ ] Clone limpo, instalação e build foram executados.
- [ ] Pacote e hash foram localizados.
- [ ] Uma lista de teste com campos representativos foi restaurada.
- [ ] Um documento e um modelo foram restaurados com metadados.
- [ ] Um web part foi publicado em página de teste.
- [ ] Testes de acesso positivo e negativo passaram.
- [ ] Tempo total foi comparado com o RTO.
- [ ] Perda simulada foi comparada com o RPO.
- [ ] Lacunas geraram ações com responsável e prazo.

## Registro do exercício

```text
Data e ambiente:
Cenário:
Participantes:
Versão/hash:
RTO/RPO aprovados:
Horário de início e fim:
Itens restaurados:
Testes executados:
Resultado:
Lacunas:
Ações, responsáveis e prazos:
Aprovação do dono técnico:
Aprovação do dono de negócio:
```
