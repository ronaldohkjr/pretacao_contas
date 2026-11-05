# Como publicar a versão online

## Antes de tudo: o que o servidor precisa

- **Memória**: pelo menos **1 GB** (no teste, a leitura das notas chegou a 450 MB e a planilha em PDF usa o
  LibreOffice, mais ~200 MB).
- **Disco permanente** para a pasta `/dados` (usuários, CPF/banco dos técnicos, notas e relatórios).
  Servidor que apaga o disco ao reiniciar **perde tudo**.
- **HTTPS** (cadeado no navegador): obrigatório, porque trafegam senhas e dados pessoais.

> ⚠️ O **Render grátis não serve**: tem 512 MB de memória, "dorme" após 15 min e apaga o disco ao reiniciar.

| Opção | Custo | Esforço | Observação |
|---|---|---|---|
| **A. Oracle Cloud "Always Free"** | grátis | médio (uma vez) | VM com 4 núcleos / 24 GB, disco próprio. Recomendado. |
| B. Railway (plano Hobby) | ~US$ 5/mês | baixo | Publica direto do GitHub, com volume para os dados. |
| C. Servidor/VM da empresa | — | TI | Mesmo passo a passo da opção A. |

---

## 1. Colocar o código no GitHub (repositório privado)

1. Crie uma conta em https://github.com (se não tiver) e um repositório **privado**, ex.: `gerador-nf-online`.
2. Envie esta pasta para o repositório (pelo GitHub Desktop, ou na linha de comando):
   ```bash
   git init -b main
   git add .
   git commit -m "Versão online do Gerador de Prestação de Contas"
   git remote add origin https://github.com/SEU-USUARIO/gerador-nf-online.git
   git push -u origin main
   ```
   A pasta `dados/` **não vai** para o GitHub (está no `.gitignore`).
3. Na aba **Actions** do repositório, o fluxo "Montar, testar e publicar" roda sozinho: monta a imagem Docker,
   testa (servidor, planilha em PDF e OCR) e publica em `ghcr.io/SEU-USUARIO/gerador-nf-online`.
   Ele roda de novo a cada atualização. Para lançar uma versão nova: altere `VERSAO.txt` e envie (push).

---

## Opção A — Oracle Cloud "Always Free" (grátis)

1. Crie a conta em https://www.oracle.com/cloud/free/ (pede cartão só para confirmar identidade; o "Always Free"
   não cobra).
2. Crie uma instância **Ubuntu 22.04**, forma **VM.Standard.A1.Flex** (Ampere), com **2 OCPU e 12 GB** (cabe no
   grátis). Guarde a chave SSH que ele oferece.
3. Em *Virtual Cloud Network → Security List*, libere a entrada das portas **80** e **443** (TCP).
4. Entre no servidor por SSH e instale o Docker:
   ```bash
   curl -fsSL https://get.docker.com | sudo sh
   sudo usermod -aG docker $USER && newgrp docker
   sudo iptables -I INPUT -p tcp -m multiport --dports 80,443 -j ACCEPT && sudo netfilter-persistent save
   ```
5. Dê acesso à imagem privada: no GitHub, *Settings → Developer settings → Personal access tokens (classic)*,
   crie um token só com **read:packages** e, no servidor:
   ```bash
   echo SEU_TOKEN | docker login ghcr.io -u SEU-USUARIO --password-stdin
   ```
6. Copie a pasta `servidor/` deste projeto para o servidor (ex.: `~/gerador`), crie o `.env` a partir do
   `.env.exemplo` e preencha `GITHUB_USUARIO` e `DOMINIO`. Sem domínio próprio, use o IP público com
   **sslip.io** (ex.: IP `129.146.10.20` → `129-146-10-20.sslip.io`): o certificado HTTPS sai sozinho.
7. Suba:
   ```bash
   cd ~/gerador && docker compose up -d
   ```
8. Abra `https://SEU-DOMINIO` → crie o **administrador** → botão **Usuários** → cadastre os técnicos.

O **Watchtower** (no mesmo docker-compose) confere a cada hora se saiu imagem nova no GitHub e atualiza sozinho.

**Backup** (recomendado, semanal):
```bash
docker run --rm -v gerador_dados:/dados -v $PWD:/bkp alpine tar czf /bkp/dados-$(date +%F).tgz -C /dados .
```

---

## Opção B — Railway (~US$ 5/mês)

1. Em https://railway.app, entre com o GitHub e assine o plano **Hobby**.
2. *New Project → Deploy from GitHub repo* → escolha o repositório (ele usa o `Dockerfile`).
3. No serviço: *Settings → Volumes → Add Volume*, caminho **`/dados`**.
4. *Variables*: `RAILWAY_RUN_UID=0` (para poder gravar no volume), `URL_PUBLICA=https://<endereço que o Railway
   gerar>` e, se for usar, as variáveis da Microsoft (abaixo).
5. *Settings → Networking → Generate Domain*. Abra o endereço → crie o administrador → cadastre os técnicos.

---

## Rascunho no Outlook com os anexos (opcional, precisa da TI)

Sem isso o sistema funciona: o botão abre um e-mail com destinatário, assunto e texto e baixa os PDFs para o
técnico anexar. Para o rascunho já sair **com os PDFs anexados** na caixa do técnico, a **TI** registra o
aplicativo no Microsoft Entra (uns 15 minutos):

1. **Entra admin center → App registrations → New registration**
   - Nome: `Gerador de Prestação de Contas`
   - Tipo de conta: *Accounts in this organizational directory only*
   - Redirect URI (**Web**): `https://SEU-DOMINIO/ms/retorno`
2. **API permissions → Add → Microsoft Graph → Delegated**: `offline_access`, `User.Read`, `Mail.ReadWrite`
   → **Grant admin consent**. (O sistema só **cria rascunhos**; quem envia é o técnico, no Outlook.)
3. **Certificates & secrets → New client secret** (anote o valor; vence conforme o prazo escolhido).
4. No servidor, preencha `MS_CLIENT_ID` (Application ID), `MS_CLIENT_SECRET` (o valor do segredo) e `MS_TENANT`
   (Directory/tenant ID) e reinicie (`docker compose up -d`).

Na primeira vez, o técnico clica em **Criar e-mail no Outlook**, entra com a conta da empresa e autoriza; nas
próximas, o rascunho é criado direto. Ele pode desconectar a qualquer momento.

---

## Segurança e LGPD (resumo para a TI)

- Senhas guardadas só como hash **PBKDF2-SHA256** (240 mil iterações, com sal); sessões em cookie `HttpOnly`,
  `SameSite=Lax`, `Secure` sob HTTPS; o arquivo de sessões guarda só o hash do token.
- Login bloqueado por 15 min após 8 erros do mesmo IP (ou 20 de qualquer IP) para o mesmo usuário.
- Cada técnico só acessa a própria pasta (`/dados/usuarios/<login>`); nomes de arquivo são limpos (sem `../`).
- Pedidos que alteram dados exigem um cabeçalho que só a própria página envia (proteção contra CSRF).
- Dados pessoais (nome, CPF, banco, assinatura) ficam só no disco do servidor, nunca no GitHub.
- Viagem enviada e arquivada tem as fotos originais apagadas (fica o PDF das notas e a planilha).
