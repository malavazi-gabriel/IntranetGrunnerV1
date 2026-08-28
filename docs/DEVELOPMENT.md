# Ambiente de desenvolvimento

## Pré-requisitos

- Windows, macOS ou Linux com suporte ao toolchain SPFx
- Node.js `>=22.14.0 <23.0.0`
- Versão observada na criação do projeto: 22.16.0
- npm compatível com a versão do Node
- Acesso ao tenant Microsoft 365 da Grunner
- Acesso ao site `https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner`
- Permissão para usar o Workbench do SharePoint

## Preparação

```powershell
git clone <URL-DO-REPOSITORIO>
Set-Location IntranetGrunnerV1
node --version
npm --version
npm ci
```

`<URL-DO-REPOSITORIO>` deve ser preenchida em `docs/OWNERS.md` pelo responsável técnico. Use `npm ci` para respeitar o arquivo de lock.

## Execução local

```powershell
npm start
```

O projeto está configurado para abrir:

```text
https://grunnerteccombr.sharepoint.com/sites/IntranetGrunner/_layouts/workbench.aspx
```

O servidor local usa HTTPS na porta 4321. A primeira execução pode solicitar confiança no certificado de desenvolvimento, dependendo da estação.

## Build de produção

```powershell
npm run build
```

Esse comando executa:

1. Limpeza das saídas geradas
2. Sass e tipos de estilo
3. TypeScript
4. ESLint
5. Webpack
6. Empacotamento SharePoint

### Estado conhecido em 27/08/2026

O build falhou com:

- 169 avisos de lint
- 58 erros de lint
- 9 erros Webpack por módulos Sass não encontrados na saída `lib`
- Total final informado: 67 erros

Não considere o projeto liberável até que `npm run build` termine com código de saída 0. Consulte [Problemas conhecidos](TROUBLESHOOTING.md#build-de-produção).

## Convenções de código

- Código de aplicação em TypeScript e React.
- Estilos em arquivos `.module.scss`.
- Arquivos `.module.scss.ts` são gerados e não devem ser versionados.
- Componentes compartilhados ficam em `src/shared`.
- Acesso a dados novo deve ficar em arquivos de serviço, não dentro de `render`.
- Toda `Promise` deve ser aguardada ou tratada explicitamente.
- Evite `any`; crie interfaces para respostas SharePoint, Graph e serviços externos.
- Não grave e-mail individual, segredo, token ou URL de ambiente diretamente no componente.

## Antes de abrir uma mudança

1. Confirme que não há alteração não relacionada no diretório de trabalho.
2. Identifique listas, colunas e integrações afetadas.
3. Atualize os documentos correspondentes.
4. Execute o checklist em [RELEASE.md](checklists/RELEASE.md).

