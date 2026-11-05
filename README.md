# Gerador de Prestação de Contas — versão online

Sistema web para os técnicos montarem o relatório de despesas de viagem a partir das fotos e PDFs das notas:
lê cada nota (OCR), aplica a política de gastos, gera o **PDF das notas**, a **planilha** (Excel e PDF) e o
**e-mail** para o aprovador, e guarda o **histórico** de cada técnico.

- Cada técnico entra com **usuário e senha** e vê só as próprias viagens.
- O **administrador** cadastra os técnicos, redefine senhas e desativa quem saiu (botão **Usuários**).
- **E-mail**: com a conta Microsoft configurada (opcional), o botão cria o rascunho na caixa do técnico com os
  PDFs anexados; sem ela, abre um e-mail com destinatário, assunto e texto e baixa os PDFs para anexar.

Como publicar: veja **[PUBLICAR.md](PUBLICAR.md)**.

## Rodar no próprio computador (teste)

```bash
pip install -r requirements.txt
pip install --no-deps rapidocr_onnxruntime
python app/servidor.py
```

Abra http://localhost:8080. No primeiro acesso o sistema pede para criar o **administrador**.
Os dados ficam na pasta `dados/` (ou na pasta da variável `DADOS_DIR`). Sem LibreOffice, a planilha em PDF
usa o Excel do Windows; no servidor (Docker) ela usa o LibreOffice.

## Variáveis de ambiente

| Variável | Para quê | Padrão |
|---|---|---|
| `PORT` | porta do servidor | `8080` |
| `DADOS_DIR` | pasta dos usuários, notas e relatórios (tem que ser um disco que não se apaga) | `./dados` |
| `URL_PUBLICA` | endereço público, ex.: `https://gerador.exemplo.com` (usado no login da Microsoft) | — |
| `MS_CLIENT_ID`, `MS_CLIENT_SECRET`, `MS_TENANT` | rascunho no Outlook pela conta Microsoft (opcional) | — |

## Estrutura

```
app/servidor.py       servidor web (rotas, geração do relatório, política, histórico)
app/contas.py         usuários, senhas (PBKDF2), sessões e bloqueio de tentativas
app/microsoft.py      rascunho no Outlook (Microsoft Graph)
app/leitor.py         leitura automática das notas (texto do PDF ou OCR RapidOCR)
app/processamento.py  tratamento das fotos (efeito "scanner")
app/index.html        página principal · app/login.html  tela de entrada
modelo/               planilha-modelo (sem dados pessoais)
servidor/             docker-compose + HTTPS (Caddy) para servidor próprio
.github/workflows/    monta, testa e publica a imagem Docker a cada atualização
```
