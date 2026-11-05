# Gerador de Prestação de Contas de Viagem — versão ONLINE (Agres)

Você vai gerar o relatório de prestação de contas do técnico a partir de imagens de notas
fiscais (fotos .jpg/.png), prints de extrato de pedágio e PDFs (NF-e/NFS-e/recibos).

## Dados do técnico (config.json)
Cada técnico tem a sua pasta `DADOS_DIR/usuarios/<login>/` (config.json, viagem.json, historico.json,
entradas/, saidas/, arquivo/, assinatura.png). O `config.json` é preenchido pela tela **"Meus dados"**:
`nome`, `nome_curto`, `cpf`, `centro_custo`, `banco`, `agencia`, `conta`, `aprovador`, `email_para`, `email_cc`,
`corpo_email`, `limite_hospedagem_diaria`, `limite_refeicao_dia`. A assinatura (opcional) fica em `./assinatura.png`.
- Leia sempre esses dados do `config.json`. **Nunca** escreva dados pessoais no código, no
  modelo ou neste arquivo.
- Se o `config.json` não existir ou estiver incompleto, peça ao técnico para abrir o programa
  e preencher "Meus dados" (ou pergunte os dados e grave no `config.json`).

## Entradas
- Pasta `./entradas/` com as imagens e PDFs das notas.
- A planilha-MODELO `./modelo/DESPESAS_MODELO.xlsx` (sempre usar como base; nunca criar do zero).
  O modelo vem com os dados pessoais em branco.
- No início, o técnico informa: CLIENTE, CIDADE de atendimento, UF e FINALIDADE.
- As DATAS da viagem saem das notas: data inicial = nota mais antiga, data final = nota mais recente
  (entre as notas incluídas no reembolso; hospedagem conta pela data de entrada).
  Se faltar CLIENTE, CIDADE ou FINALIDADE, **pergunte antes** de nomear/gerar.
  Em qualquer dúvida de leitura de uma nota (valor, data, categoria) ou de lançamento, **pergunte**.

## Saídas (em ./saidas/)
1. `AAAAMMDD_NOTAS_<NOME_CURTO>_<CLIENTE>_<CIDADE>_<UF>.pdf`  (AAAAMMDD = 1º dia da viagem)
2. `AAAAMMDD_DESPESAS_<NOME_CURTO>_<CLIENTE>_<CIDADE>_<UF>.xlsx` e a mesma planilha em PDF
   (`AAAAMMDD_DESPESAS_<NOME_CURTO>_<CLIENTE>_<CIDADE>_<UF>.pdf`, exportada pelo Excel do computador, 1 página)
3. O texto do e-mail (abaixo), impresso no final.
> `<NOME_CURTO>` = `nome_curto` do config.json. CLIENTE/CIDADE com espaços → use `_`
> (ex.: ENCRUZILHADA_DO_SUL), sem acentos, em maiúsculas.

## Classificação de cada arquivo
- Foto de cupom/recibo térmico (jpg/png de câmera) → `process_photo()` (modo "cupom").
- Print/screenshot de extrato de pedágio (imagem digital, fundo branco) → `process_screenshot()` (modo "print": realce leve, SEM recorte).
- Documento A4 fotografado (NFS-e de hotel em folha) → `process_document()` (modo "documento": Otsu, página inteira).
- Foto em que o recorte corta cabeçalho/valor → `process_full()` (modo "inteira").
- PDF (NF-e/NFS-e/recibo) → renderizar a(s) página(s) a 150 dpi (pymupdf) e incluir como está.

O código de tratamento está em `app/processamento.py` (use exatamente essa lógica).
Regra de qualidade: a saída deve ficar como no app **Adobe Scan**: nota INTEIRA,
sem cortar bordas, fundo branco, alto contraste, letras nítidas. Se um recorte cortar
cabeçalho/valor, reprocesse aquela imagem com `process_document()` ou `process_full()`.
Sempre revise visualmente as páginas críticas (hotéis e cupons com valor).

## Montagem do PDF (uma nota por página, A4)
Margem de 20 pt, imagem em tons de cinza, lado maior limitado a 1700 px, centralizada e
escalada para caber na página (ver `gerar()` em `app/servidor.py`).

## Ordenação das páginas do PDF
1. Por **dia** (crescente).
2. Dentro do dia, por categoria: **1) hospedagem, 2) abastecimento, 3) pedágio,
   4) alimentação, 5) material de uso e consumo, 6) outros**.
3. Dentro da categoria, por horário.
4. **Hospedagem (NFS-e)**: usar a **data de ENTRADA da diária** para definir o dia
   (não a data de emissão).
5. "Pagamento de violação" em praça de pedágio conta como **pedágio**.

## Planilha — mapa EXATO de células do modelo (openpyxl)
Abra o modelo, preencha e salve com o novo nome.
- Dados do técnico (do config.json): `B8`=nome, `I8`=CPF, `B10`=centro de custo,
  `B12`=banco, `K12`=agência, `P12`=conta. Assinatura (se houver `assinatura.png`) no campo
  Assinatura (G64:K68).
- `E10` = Data de Saída Realizada | `I10` = Data de Chegada Realizada
- `B14` = Motivo = `"Atendimento <CLIENTE> <CIDADE>/<UF>"`
- Linha 19 = cabeçalho de datas por dia. Colunas de dia (merge de 2 colunas cada):
  `D, F, H, J, L, N, P, R` (até 8 dias) e `T` = coluna Total. Dia1->D, dia2->F, ...
- Linhas de VALOR por categoria (escreva na coluna do dia):
  - Hospedagem -> **21**
  - Refeições -> **24**
  - Pedágio -> **27**
  - Frete -> **30**
  - Combustível -> **33**
  - Material de Uso e Consumo -> **36**
  - Locações de Veículos -> **39**
  - Transporte/táxi -> **42**
  - Outras Despesas -> **45**
  - (linha **47** = totais; já tem fórmulas `SUM`)
  Ex.: hospedagem do dia 3 = `H21`; combustível do dia 1 = `D33`.
  Pode usar fórmula para somar várias notas do mesmo dia/categoria, ex.: `=19+57.31+53`.
- Detalhe de **Outras Despesas** (linhas 51-53): `B`=data, `D`=descrição, `X`=valor;
  `X54` soma automática. (ex.: Starlink, lavagem de veículo, estacionamento)
- Quilometragem (56-61): deixar vazio; `S60`=0,8 (taxa) permanece.
- ACERTO: `K64`=adiantamento (deixar vazio), `O64`=`=T47+X61` (total), `S64`=saldo. `C66` = data da geração (a data ao lado da assinatura do técnico; formato dd/mm/yy).
- **Formato de data**: aplicar `dd/mm/yy` em TODA célula cujo number_format tenha data.
> openpyxl não calcula fórmulas; o Excel recalcula ao abrir.

## Regras de lançamento
- **Combustível**: lançar o **VALOR PAGO** (líquido de desconto).
- **Lavagem de veículo** (Lava Car etc.) -> **Outras Despesas**.
- **Compra de ferramenta/equipamento** (NF-e de produto) -> **Material de Uso e Consumo**.
- **Starlink/internet**: confirmar o valor com o técnico.
- Item cuja data difere dos dias de viagem: **perguntar** se cria coluna de novo dia
  ou se lança em outro relatório.

## Política de gastos (limites por dia)
- Limites no config.json: `limite_hospedagem_diaria` (padrão R$ 120 por diária; uma nota de hotel de
  N diárias tem limite 120 × N no dia da entrada) e `limite_refeicao_dia` (padrão R$ 90 por dia). 0 = sem limite.
- Dia/categoria acima do limite: **perguntar** ao técnico o motivo.
  - Com motivo: reembolsa o valor cheio e o motivo entra no e-mail, na seção
    "Justificativa dos gastos acima da política" (`- dd/mm – Categoria: R$ total (limite R$ X): motivo`).
  - Sem motivo: o reembolso é **ajustado ao limite automaticamente**, reduzindo as notas daquele
    dia/categoria a partir da de maior valor. O PDF mantém as notas originais, com um aviso em vermelho no
    topo da página de cada nota ajustada ("REEMBOLSO AJUSTADO À POLÍTICA: será reembolsado R$ X desta nota de
    R$ Y (...)"), e a foto um pouco menor para não ser coberta; a planilha usa o valor ajustado; o e-mail ganha a seção "Valores ajustados ao limite da política" (`- dd/mm – Categoria: R$ total → R$ limite`).
- No programa a decisão fica em `viagem.json` → `excessos` (`"AAAA-MM-DD|Categoria": {"modo": "justificar", "motivo": "..."}`)
  e as diárias de cada nota de hotel em `linhas[].diarias`.

## Texto do e-mail (imprimir no final)
Assunto: `Relatório de Despesas atendimento <CLIENTE> <CIDADE>/<UF>`
Corpo: `corpo_email` do config.json, trocando `{APROVADOR}` (aprovador do config.json),
`{CLIENTE}`, `{CIDADE}`, `{UF}`, `{FINALIDADE}`, `{DATA_INICIAL}`, `{DATA_FINAL}` (dd/mm/aaaa)
e `{NOME}`. Texto padrão:
```
{APROVADOR},
Segue em anexo para aprovação, planilha de despesas referente a viagem a {CIDADE}, atendimento a {CLIENTE}. Para {FINALIDADE}.

Entre os dias {DATA_INICIAL} a {DATA_FINAL}
```

## Programa (versão online)
- Servidor web multiusuário: `app/servidor.py` (porta `PORT`, dados em `DADOS_DIR`), login por usuário e senha
  (`app/contas.py`: PBKDF2, sessões em cookie, bloqueio de tentativas). O primeiro acesso cria o administrador;
  o administrador cadastra os técnicos (botão Usuários). Senha provisória: o técnico troca no primeiro acesso.
- Cada requisição usa as pastas do técnico logado (`P.usar(login)` em `servidor.py`, objeto `P`).
- Planilha em PDF pelo LibreOffice (`soffice --headless --convert-to pdf`; no Windows sem LibreOffice, Excel).
- E-mail: com `MS_CLIENT_ID/MS_CLIENT_SECRET/MS_TENANT/URL_PUBLICA`, cria o rascunho na caixa do técnico pela
  conta Microsoft com os PDFs anexados (`app/microsoft.py`, Microsoft Graph, permissão Mail.ReadWrite); sem isso,
  devolve um mailto (para/assunto/texto) e o zip para anexar à mão. Nunca envia e-mail.
- Sem Windows: "Abrir pasta" virou download (zip por viagem no Histórico); viagem enviada e arquivada tem as fotos
  apagadas (não há Lixeira); o cache de cada técnico fica em `DADOS_DIR/cache/<login>`.
- Publicação: Docker (`Dockerfile`), GitHub Actions monta, testa e publica a imagem; `servidor/` tem o
  docker-compose com HTTPS (Caddy) e atualização automática (Watchtower). Ver `PUBLICAR.md`.
- Mínimo de 1 GB de memória (OCR ~450 MB + LibreOffice) e disco permanente para `DADOS_DIR`.
- Dados pessoais dos técnicos nunca vão para o GitHub (`dados/` no `.gitignore`).

## Fluxo
1. Ler o `config.json` do técnico (pasta dele em `DADOS_DIR/usuarios/<login>/`).
2. Listar os arquivos de `./entradas/`, ler data/valor/estabelecimento/categoria de cada um.
3. Mostrar um resumo por dia/categoria com os totais e pedir CLIENTE/CIDADE/UF/FINALIDADE (se faltarem) e confirmar qualquer dúvida.
4. Processar imagens, montar o PDF na ordem, preencher a planilha, gerar o e-mail.
5. Salvar tudo em `./saidas/` com os nomes padronizados e mostrar os totais.
