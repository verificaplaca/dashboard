# Contexto do Projeto — Verifica Placa Dashboard

## Diretório de trabalho
O terminal sempre estará em `/Users/danielmacedo/Documents/GitHub/verificaplaca/`.
Nunca incluir `cd` para esse caminho nem caminhos absolutos nos comandos.

---

## Git

**Push normal:**
```bash
git add <arquivo>
git commit -m "mensagem"
git push
```

**Quando o push for rejeitado (remote tem commits novos):**
```bash
git pull --rebase origin main
git push
```

**Resolver conflito durante rebase:**
```bash
# 1. Edite o arquivo e resolva os marcadores <<<<<<< / ======= / >>>>>>>
git add dashboard.html
git rebase --continue
git push
```

**Remover lock travado:**
```bash
rm .git/HEAD.lock
```

---

## SSH / Autenticação

Remote configurado com alias SSH:
```
git remote: github-verificaplaca:verificaplaca/dashboard.git
```

Entrada no `~/.ssh/config`:
```
Host github-verificaplaca
  HostName github.com
  User git
  AddKeysToAgent yes
  UseKeychain yes
  IdentityFile ~/.ssh/id_ed25519_git_verificaplaca
```

Para não pedir passphrase:
```bash
ssh-add --apple-use-keychain ~/.ssh/id_ed25519_git_verificaplaca
```

---

## Arquivos principais

| Arquivo | Descrição |
|---|---|
| `dashboard.html` | Dashboard principal (único arquivo, tudo inline) |
| `index.html` | Landing page |
| `favicon-verifica-placa-2.jpg` | Favicon da dash |
| `gads-intraday-script.js` | Script Google Ads intraday |
| `pagarme-unified.js` | Integração Pagar.me |

---

## Convenções do dashboard

- Todo o código (HTML, CSS, JS) está em `dashboard.html` — arquivo único.
- Dados vêm do Supabase via `supaGet()`.
- KPI cards são definidos em arrays `row1`, `row2`, `row3` e renderizados por `renderRow()`.
- Cada card tem: `id`, `label`, `value`, `sub` (legenda), `spark` (mini gráfico), `delta`, `badge`, `color`.

**Cálculos principais:**
```js
// Lucro Bruto
pft = receita − custo_ads − custo_bureau

// Lucro Líquido (imposto sobre receita bruta)
netPft = rev * 0.92
```
