# Implantação e rollback

## Visão geral

A solução gera um pacote SharePoint Framework com o nome:

```text
sharepoint/solution/intranet-grunnerv-1.sppkg
```

O pacote declara:

- Solução: `intranet-grunnerv-1-client-side-solution`
- Identificador: `46cfb39b-a6aa-4103-9a23-9aaf88165a6a`
- Versão observada: `1.0.0.14`
- Implantação disponível para todos os sites: `skipFeatureDeployment: true`
- Permissão Graph: `User.Read.All`

## Pré-condições

- Build de produção aprovado com código de saída 0
- Checklist de liberação preenchido
- Versão do pacote atualizada
- Backup do pacote anterior disponível em local controlado
- Responsável por SharePoint e aprovador da mudança identificados em `OWNERS.md`
- Janela de mudança aprovada
- Permissão Graph revisada e aprovada pelo administrador do tenant

## Gerar o pacote

```powershell
npm ci
npm run build
```

Confirme a existência e a data do arquivo:

```powershell
Get-Item .\sharepoint\solution\intranet-grunnerv-1.sppkg
```

Não publique pacote produzido por um build que terminou com erro.

## Versionamento

Antes do build de uma nova liberação:

1. Atualize `solution.version` em `config/package-solution.json`.
2. Atualize `features[0].version` para o mesmo valor.
3. Registre o conteúdo da versão no changelog da equipe.
4. Crie uma tag Git após a validação em produção.

O processo atual não possui tags. Implantação e código podem divergir sem essa etapa.

## Publicação no App Catalog

1. Acesse o App Catalog definido pelo administrador do tenant.
2. Registre o pacote atualmente publicado e a versão.
3. Faça upload do novo `.sppkg`.
4. Confirme a substituição da solução existente.
5. Selecione implantação ampla somente se esse for o padrão aprovado pela empresa.
6. Revise as permissões solicitadas.
7. No SharePoint Admin Center, confirme a aprovação de `Microsoft Graph / User.Read.All`.
8. Aguarde a propagação.

O endereço exato do App Catalog está `PENDENTE DE VALIDAÇÃO` e deve ser registrado em `ENVIRONMENT-INVENTORY.md` sem incluir credenciais.

## Teste após publicação

Valide no mínimo:

1. Home abre sem tela em branco.
2. Notícias publicadas aparecem para colaborador comum.
3. Rascunhos aparecem somente para o perfil de Marketing previsto.
4. Celebrações carregam pelo Graph.
5. Chamados carregam para um usuário de teste autorizado.
6. Central de atalhos carrega links ativos.
7. História abre corretamente.
8. Políticas carregam, pesquisam e abrem documento.
9. Perfil de Qualidade acessa a área de gestão.
10. Perfil de TI consulta ativos e histórico.
11. Colaborador comum vê somente seus próprios ativos.
12. Geração de Word e exportação Excel funcionam.
13. Navegação funciona em desktop e celular.

Registre executor, data, ambiente e resultado no checklist de liberação.

## Rollback

1. Interrompa novas alterações no ambiente.
2. Registre horário, sintomas e versão afetada.
3. Recupere o `.sppkg` da última versão aprovada no repositório de artefatos.
4. Faça upload do pacote anterior no App Catalog.
5. Confirme a substituição.
6. Execute os testes críticos da seção anterior.
7. Registre a causa e abra uma ação corretiva.

Não use como única estratégia a pasta local `Bk anterior`. O pacote anterior deve ficar em armazenamento corporativo controlado e associado a uma tag Git.

## Evidência necessária para 5/5

- Uma implantação executada por pessoa que não seja a autora original, usando apenas este documento.
- Um rollback ensaiado e registrado.
- Pacote publicado rastreável até commit e tag.
- Tempo real de implantação e rollback registrado.

