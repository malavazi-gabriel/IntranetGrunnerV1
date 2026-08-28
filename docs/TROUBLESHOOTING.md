# Solução de problemas

Antes de alterar o sistema, registre URL, usuário de teste, horário, versão do pacote, mensagem exata, resposta HTTP e passos de reprodução. Não inclua tokens nem dados pessoais nas evidências.

## Build falha com módulos Sass ausentes

**Sintoma:** erros como `Can't resolve ...module.scss` a partir de arquivos gerados em `lib`.

**Estado observado no levantamento:** `npm run build` não concluiu; foram encontrados erros de lint e erros de resolução Sass.

**Verificações:**

1. Executar `npm ci` com Node da faixa declarada no `package.json`.
2. Executar `npm run clean` e depois `npm run build`.
3. Confirmar se o pipeline copia/processa `.module.scss` na versão atual do toolchain SPFx.
4. Tratar os erros de lint como defeitos do build, sem desabilitar regras globalmente apenas para publicar.
5. Registrar a correção e adicionar o build limpo como gate de release.

## Página em branco ou web part com erro

1. Confirmar se apenas uma página ou todo o site foi afetado.
2. Verificar console e rede do navegador.
3. Conferir se a versão correta está implantada no App Catalog.
4. Validar a página e o ID do web part em [ENVIRONMENT-INVENTORY.md](ENVIRONMENT-INVENTORY.md).
5. Verificar permissões e existência das listas usadas pelo módulo.
6. Se começou após release, executar o rollback de [DEPLOYMENT.md](DEPLOYMENT.md).

## Perfil, diretório ou aniversariantes não carregam

1. Verificar a resposta de `/me` e `/users` no contexto do usuário.
2. Confirmar aprovação de `User.Read.All`.
3. Confirmar `accountEnabled` e os atributos `extensionAttribute1`/`extensionAttribute10` no AD.
4. Validar formatos `DD/MM` e `DD/MM/AAAA`.
5. Conferir se a lista `AniversariantesGrunner` está disponível para o fallback esperado.
6. Não aumentar permissão Graph sem revisão de segurança.

## Notícias, eventos ou links úteis não aparecem

1. Verificar existência e ACL das listas `NoticiasGrunner`, `EventosGrunner` e `LinksUteisGrunner`.
2. Conferir nomes internos no [DATA-DICTIONARY.md](DATA-DICTIONARY.md).
3. Validar status, indicador `Ativo`, ordem e URLs.
4. Testar com colaborador e com administrador de conteúdo.
5. Confirmar que conteúdo HTML não foi bloqueado/sanitizado incorretamente.

## Usuário de conteúdo não recebe ações administrativas

1. Consultar cargo e departamento retornados por `/me`.
2. Verificar a regra atual e a exceção por e-mail no código.
3. Como correção estrutural, migrar a autorização para grupo corporativo e ACL de serviço.
4. Nunca conceder acesso apenas adicionando nova exceção hardcoded.

## Políticas/SGQ não carregam ou não podem ser gerenciadas

1. Conferir bibliotecas `PoliticasGrunner` e `RascunhosSGQ`.
2. Conferir nomes internos e metadados obrigatórios.
3. Verificar participação no grupo `Qualidade - Gestão de Documentos`.
4. Testar leitura e escrita diretamente no SharePoint com o mesmo usuário.
5. Abrir os três modelos em `/Templates_SGQ` e testar geração.

## Painel de ativos vazio ou acesso incorreto

1. Conferir lista `Ativos de TI` e campos `field_*` do dicionário.
2. Conferir e-mail e `NivelAcesso` em `Acessos_Painel_Ativos`.
3. Validar as ACLs reais da lista; a tela não substitui segurança.
4. Abrir o modelo `/Modelos_TI/Molde_Termo_Grunner.docx`.
5. Testar um usuário sem perfil, um `Visualizador` e um `TI`.

## Termo de ativo ou documento SGQ não é gerado

1. Abrir o arquivo modelo diretamente no SharePoint.
2. Confirmar o caminho exato e permissões.
3. Validar se os placeholders esperados continuam presentes.
4. Testar um documento mínimo e registrar a mensagem da biblioteca de geração.
5. Restaurar a última versão aprovada do modelo, não um arquivo local desconhecido.

## Chamados ou comentários não carregam

1. Verificar DNS/TLS e resposta de `https://admin.grunnertec.com.br`.
2. Testar as rotas documentadas no inventário com conta de teste autorizada.
3. Registrar código HTTP, tempo e corpo sem dados pessoais.
4. Verificar autenticação/autorização no backend; não corrigir apenas no frontend.
5. Acionar o dono do backend, pois o serviço não está neste repositório.
6. Comunicar indisponibilidade parcial e orientar o canal alternativo aprovado.

## Pacote não aparece após implantação

1. Confirmar incremento da versão no `package-solution.json`.
2. Confirmar upload, implantação e escopo no App Catalog correto.
3. Verificar erros de aprovação de API.
4. Confirmar associação da solução ao site e publicação da página.
5. Limpar cache somente depois de confirmar servidor e versão.

## Quando escalar

Escalar imediatamente quando houver suspeita de exposição de dados, concessão indevida, corrupção/exclusão relevante ou indisponibilidade geral. Para demais casos, usar a matriz de [OPERATIONS-RUNBOOK.md](OPERATIONS-RUNBOOK.md).

