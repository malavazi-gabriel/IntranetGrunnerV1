# Checklist de release

Preencher para cada implantação. Um item não atendido exige bloqueio ou exceção formal aprovada com risco, dono e prazo.

## Identificação

```text
Versão:
Commit/tag:
Ambiente:
Solicitação/mudança:
Responsável técnico:
Aprovador de negócio:
Data/janela:
Plano de rollback:
```

## Antes do pacote

- [ ] Escopo e critérios de aceite aprovados.
- [ ] Código revisado por outra pessoa.
- [ ] `npm ci` executado em ambiente limpo.
- [ ] `npm run build` concluído sem erro.
- [ ] Testes automatizados passaram e relatório foi guardado.
- [ ] Testes manuais dos módulos afetados passaram.
- [ ] Dependências/vulnerabilidades foram verificadas e classificadas.
- [ ] Nenhum segredo, token ou dado pessoal foi incluído.
- [ ] Alterações de HTML rico foram testadas contra XSS.
- [ ] Documentação e ADRs foram atualizados quando aplicável.

## Configuração e dados

- [ ] Alterações em listas/bibliotecas têm script ou roteiro de migração e reversão.
- [ ] Nomes internos e permissões foram validados em homologação.
- [ ] Alterações Graph foram aprovadas por administrador autorizado.
- [ ] Grupos e perfis passaram por testes positivo e negativo.
- [ ] Modelos Word afetados foram versionados e testados.
- [ ] Backend de chamados foi validado quando impactado.
- [ ] Backup/versão anterior e artefato de rollback estão disponíveis.

## Geração e implantação

- [ ] Versão do `package-solution.json` foi incrementada.
- [ ] `.sppkg` foi gerado pelo processo oficial.
- [ ] Hash SHA-256 do pacote foi registrado.
- [ ] Pacote e notas de release foram armazenados no local aprovado.
- [ ] App Catalog, tenant, site e janela foram confirmados.
- [ ] Implantação foi realizada por pessoa autorizada.
- [ ] Aprovações de API foram conferidas.

## Smoke test

- [ ] `Inicio.aspx`: conteúdo, perfil, eventos e chamados.
- [ ] `centraldeatalhos.aspx`: links e solicitações.
- [ ] `Historia.aspx`: conteúdo institucional.
- [ ] `Políticas-da-Empresa.aspx`: consulta, documento e perfil Qualidade.
- [ ] `GerenciamentoDeAtivos.aspx`: consulta, acesso e geração de termo.
- [ ] Usuário comum não recebeu funções administrativas.
- [ ] Console/rede sem nova falha crítica.
- [ ] Donos funcionais afetados deram aceite.

## Encerramento

```text
Resultado:
Versão confirmada em produção:
Hash do pacote:
Evidências:
Incidentes/desvios:
Rollback necessário?:
Quem implantou:
Quem validou tecnicamente:
Quem aprovou funcionalmente:
Data/hora de encerramento:
```

