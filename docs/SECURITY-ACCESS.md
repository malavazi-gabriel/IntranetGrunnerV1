# Segurança e controle de acesso

## Princípio central

A ocultação de botão ou página no React melhora a experiência, mas não autoriza nem protege dados. Toda operação deve ser restringida no SharePoint, Microsoft Graph ou backend que efetivamente armazena/processa o dado.

## Perfis observados

| Perfil | Capacidade funcional | Fonte de autorização | Dono da concessão |
|---|---|---|---|
| Colaborador | Consultar conteúdo geral e seus chamados | Conta corporativa e ACLs SharePoint | **PENDENTE** |
| Marketing/Comunicação | Administrar notícias, eventos e atalhos | Hoje: cargo/departamento e exceção por e-mail | **PENDENTE** |
| Qualidade | Administrar documentos e revisões SGQ | Grupo `Qualidade - Gestão de Documentos` | **PENDENTE** |
| Visualizador de ativos | Consultar ativos permitidos | Lista `Acessos_Painel_Ativos` | TI — confirmar |
| TI de ativos | Administrar ativos e acessos | Lista `Acessos_Painel_Ativos` | TI — confirmar |
| Administrador M365 | App Catalog, permissões e site | Funções administrativas Microsoft 365 | **PENDENTE** |

## Fragilidades a eliminar

- Regra administrativa baseada em texto de cargo/departamento pode conceder ou retirar acesso por variação cadastral.
- Exceção de usuário por e-mail está codificada no frontend e deve migrar para grupo gerenciado.
- O frontend consulta membros de grupo e listas de acesso, mas a proteção real depende das ACLs dos recursos.
- URLs de serviços estão fixas no código, sem configuração por ambiente.
- Conteúdo rico utiliza renderização HTML; qualquer origem editável deve ser sanitizada e testada contra XSS.
- O contrato de autenticação/autorização do backend de chamados não está documentado aqui.

## Modelo-alvo

1. Criar grupos corporativos por função: conteúdo, Qualidade, ativos visualização e ativos administração.
2. Aplicar menor privilégio diretamente às listas, bibliotecas e APIs.
3. Remover e-mails hardcoded e decisões baseadas em cargo/departamento.
4. Manter no código apenas a leitura do perfil/grupo para experiência; negar no serviço para segurança.
5. Separar configuração de desenvolvimento, homologação e produção.
6. Registrar todas as concessões, remoções e revisões de acesso.

## Microsoft Graph

- Permissão declarada: `User.Read.All`.
- Uso observado: `/me` e `/users`, incluindo atributos sincronizados do AD.
- Aprovação deve ser feita por administrador autorizado e vinculada a justificativa de negócio.
- Revisar anualmente a necessidade do escopo e registrar o resultado.
- Não registrar payloads completos de usuários em logs.

## Processo de acesso

### Concessão

1. Solicitação identifica usuário, perfil, justificativa, período e gestor.
2. Dono do dado aprova; TI executa a inclusão no grupo/lista adequada.
3. Outra pessoa valida o perfil efetivo com usuário de teste.
4. Evidência é anexada ao chamado de acesso.

### Alteração ou desligamento

1. RH/gestor informa a mudança.
2. TI remove acessos incompatíveis no mesmo prazo da política corporativa.
3. Tokens/sessões são revogados quando houver risco.
4. A remoção é testada e registrada.

### Revisão trimestral

- Exportar membros dos grupos e da lista `Acessos_Painel_Ativos`.
- Comparar com vínculo, função e necessidade atual.
- Remover órfãos e excessos.
- Obter aceite dos donos de cada dado.
- Guardar data, executor, aprovadores, itens removidos e exceções.

## Checklist técnico de release

- [ ] Sem segredo, token ou dado pessoal real no repositório e no pacote.
- [ ] Dependências verificadas e vulnerabilidades classificadas.
- [ ] Conteúdo HTML sanitizado e casos de XSS testados.
- [ ] Operações de escrita negadas para usuários sem permissão no serviço de destino.
- [ ] Permissões Graph e SharePoint não foram ampliadas sem aprovação.
- [ ] Logs não expõem e-mail, payload de usuário ou conteúdo sensível desnecessário.
- [ ] Backend de chamados valida identidade e autorização no servidor.
- [ ] Evidência de revisão de acesso vigente.
- [ ] Plano de reversão testável.

## Resposta a evento de segurança

1. Preservar evidências e limitar acesso ao caso.
2. Conter credenciais, grupos, pacote ou integração afetada.
3. Acionar os responsáveis de segurança e privacidade definidos em [OWNERS.md](OWNERS.md).
4. Avaliar dados e titulares afetados conforme [LGPD-DATA-GOVERNANCE.md](LGPD-DATA-GOVERNANCE.md).
5. Corrigir, validar, restaurar e monitorar.
6. Registrar causa, impacto, decisões e prevenção.

