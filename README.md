# FinançaPro – Receitas & Despesas com Google Sheets

Aplicação em **arquivo único** (`index.html`) + função serverless da Vercel (`api/sheets.js`) que faz a ponte segura com um **Google Apps Script** ligado à planilha.

```
Navegador (index.html) ──► /api/sheets (Vercel) ──► Apps Script (Web App) ──► Google Sheets
```

## Estrutura do repositório GitHub

```
financas-sheets/
├── index.html            # Front-end completo (HTML + CSS + JS)
├── api/
│   └── sheets.js         # Proxy serverless (guarda o token)
├── apps-script/
│   └── Code.gs           # Código a colar no Apps Script da planilha (referência)
└── README.md
```
Nenhum `package.json` é necessário (Vercel detecta `api/` automaticamente, Node 18+).

## 1. Planilha no Google Sheets
1. Crie uma planilha nova em https://sheets.google.com (ex.: "Finanças").
2. Renomeie a aba para **Lancamentos** (ou deixe que o script crie).
3. Cabeçalho na linha 1 (A→H):
   `Data | Descrição | Valor | Tipo Despesa | Qt Parcelas | Forma Pagto | Tipo Receita | ID`
   - A coluna **ID** é técnica (gerada automaticamente) e é obrigatória para editar/excluir.
   - Formate a coluna A como **Texto simples** (Formatar > Número > Texto simples) para manter DD/MM/AAAA.
   - Regra de natureza: linha com **Tipo Receita** preenchido = receita; com **Tipo Despesa** = despesa.

## 2. Apps Script (backend)
1. Na planilha: **Extensões > Apps Script**.
2. Apague o conteúdo e cole o arquivo `apps-script/Code.gs`. Salve.
3. **Configurações do projeto (⚙) > Propriedades do script > Adicionar**: `API_TOKEN` = uma senha longa aleatória (ex.: 40 caracteres).
4. **Implantar > Nova implantação > Tipo: App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
5. Autorize as permissões e copie a **URL do App da Web** (termina em `/exec`).
6. Sempre que alterar o código: **Implantar > Gerenciar implantações > Editar > Nova versão**.

## 3. Variáveis de ambiente na Vercel
Project Settings → Environment Variables (Production, Preview e Development):

| Variável          | Obrigatória | Valor                                              |
|-------------------|-------------|----------------------------------------------------|
| `APPS_SCRIPT_URL` | Sim         | URL `/exec` do App da Web                          |
| `API_TOKEN`       | Sim         | Mesmo valor da propriedade `API_TOKEN` do script   |
| `ACCESS_KEY`      | Opcional    | Senha que o operador digita no app para acessar    |

## 4. Implantação
1. Crie o repositório no GitHub e envie os arquivos acima.
2. Na Vercel: **Add New > Project > Import** o repositório. Framework Preset: **Other**. Sem build command.
3. Cadastre as variáveis e clique **Deploy** (redeploy após alterar variáveis).
4. Acesse o domínio, clique em **Carregar dados da planilha**.

## Funcionalidades
- Carrossel animado (Ken Burns) na página inicial e botão de carregamento.
- Dashboard: KPIs, gráfico de fluxo mensal, distribuição por forma de pagamento.
- Filtros: busca, período, atalhos (mês, mês anterior, 90 dias, ano, futuros), natureza, classificação, forma de pagto; ordenação e paginação.
- CRUD completo; despesa **Fixa** com N parcelas gera N linhas mensais "Descrição (i/N)"; exclusão de parcelamento inteiro ou só da parcela.
- PDF profissional com pré-visualização; Excel estilizado (cabeçalho, moeda, autofiltro, totais e aba "Resumo Mensal").
- Modo claro/escuro (memoriza a escolha) e **modo demonstração** automático se a API não responder.
