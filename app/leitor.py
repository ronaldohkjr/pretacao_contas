# Leitura automática das notas: texto (PDF direto ou OCR nas fotos) -> categoria, valor, data, hora, descrição.
# É uma SUGESTÃO: o programa marca os campos lidos para o técnico conferir.
import re, threading, unicodedata
from collections import Counter
from pathlib import Path

import numpy as np
from PIL import Image, ImageOps

_ocr = None
_ocr_lock = threading.Lock()


def _motor():
    global _ocr
    if _ocr is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr = RapidOCR()
    return _ocr


def _sem_acento(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode().upper()


def texto(path, imagem_tratada=None):
    """Linhas de texto da nota, com cache em disco (o OCR é lento)."""
    import json, os, tempfile, hashlib
    path = Path(path)
    st = path.stat()
    cache = Path(tempfile.gettempdir()) / "gerador_nf" / ("ocr_" + hashlib.md5(f"{path.name}|{st.st_size}|{st.st_mtime}|{imagem_tratada}".encode()).hexdigest() + ".json")
    if cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    linhas = _texto(path, imagem_tratada)
    cache.parent.mkdir(parents=True, exist_ok=True); cache.write_text(json.dumps(linhas, ensure_ascii=False), encoding="utf-8")
    return linhas


def _texto(path, imagem_tratada=None):
    if path.suffix.lower() == ".pdf":
        import pymupdf
        with pymupdf.open(path) as d:
            linhas = [l for pg in d for l in pg.get_text().splitlines()]
        if sum(len(l) for l in linhas) > 40:
            return linhas
        im = Image.open(imagem_tratada) if imagem_tratada else None  # PDF escaneado: OCR
    else:
        im = Image.open(imagem_tratada) if imagem_tratada else ImageOps.exif_transpose(Image.open(path))
    if im is None:
        return []
    im = im.convert("RGB")
    if max(im.size) > 1800:
        r = 1800 / max(im.size); im = im.resize((int(im.size[0] * r), int(im.size[1] * r)))
    with _ocr_lock:
        res, _ = _motor()(np.array(im))
    if not res:
        return []
    # junta caixas na mesma altura numa linha só (esquerda -> direita)
    caixas = sorted(((b[0][1] + b[2][1]) / 2, b[0][0], t, b[2][1] - b[0][1]) for b, t, _ in res)
    linhas, atual, y0, h0 = [], [], None, None
    for y, x, t, h in caixas:
        if y0 is not None and abs(y - y0) > max(h, h0) * 0.6:
            linhas.append(" ".join(t for _, t in sorted(atual))); atual = []
        if not atual: y0, h0 = y, h
        atual.append((x, t))
    if atual: linhas.append(" ".join(t for _, t in sorted(atual)))
    return linhas


RE_VALOR = re.compile(r"(?<![\d.,])(\d{1,3}(?:\.\d{3})+,\d{2}|\d{1,5}[.,]\d{2})(?!\d)")
RE_DATA = re.compile(r"(?<!\d)(\d{1,2})[/.](\d{1,2})[/.](20\d{2}|\d{2}(?!\d))")  # OCR às vezes gruda a hora no ano
RE_HORA = re.compile(r"(?<!\d)([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?(?!\d)")

CATEGORIAS = [  # (categoria, palavras) — na ordem de prioridade
    ("Outras Despesas", ["ESTACIONAMENTO", "LAVA CAR", "LAVACAR", "LAVA JATO", "LAVAGEM", "STARLINK"]),
    ("Hospedagem", ["HOTEL", "HOSPEDAGEM", "POUSADA", "DIARIA", "HOSPEDE", "APTO"]),
    ("Pedágio", ["PEDAGIO", "PRACA:", "PRACA DE PEDAGIO", "CONCESSIONARIA", "ARRECADADOR", "AUTOBAN", "ECOVIA", "ROTA SERTANEJA"]),
    ("Combustível", ["ETANOL", "GASOLINA", "DIESEL", "COMBUSTIV", "AUTO POSTO", "POSTO "]),
    ("Transporte/táxi", ["UBER", "99POP", "99 TAXI", "TAXI", "CABIFY"]),
    ("Material de Uso e Consumo", ["FERRAG", "FERRAMENTA", "ABRACADEIRA", "DISCO CORTE", "FITA ISOLANTE", "PARAFUSO",
                                   "MATERIAIS ELETRICOS", "MAT CONSTRUCAO"]),
    ("Refeições", ["RESTAURANTE", "REFEIC", "LANCH", "PADARIA", "CAFE", "AGUA", "COCA", "REFRI", "BUFFET", "GRILL",
                   "PIZZ", "ESFIHA", "EMPADA", "SANDUICHE", "MARMITA", "PRATO", "ALMOCO", "JANTA", "CHURRASC",
                   "DELICIAS", "BAR ", "SUCO", "DOOG", "CARNINHA", "PALMITO", "ESPRESSO"]),
]
# palavras-chave do total, comparadas SEM espaços e tolerando trocas comuns do OCR (C/G, I/L/1)
CHAVES_VALOR = [r"VALORPAGO", r"VALORAPAGAR", r"TOTALAPAGAR", r"VALORLIQU[IL1]DODA(NOTA|NFS)", r"VALORTOTALDOSSERVI[CG]OS",
                r"VALORDOSSERVI[CG]OS", r"VALORTOTAL", r"TOTALR\$", r"CREDITOAVISTA", r"DEBITOAVISTA", r"VALORLIQU[IL1]DO",
                r"TOTAL", r"CREDITO", r"DEBITO", r"VALOR"]
IGNORAR_VALOR = ["TRIB", "IMPOSTO", "ISS", "IBPT", "DESCONTO", "TROCO", "UNIT", "ALIQ", "BASEDECALC", "PIS", "COFINS",
                 "LEI", "IBS", "CBS", "DEDUC", "RETENC", "RETENG", "TAXADESERVICO", "GANHOS", "MOTORISTA", "CUSTOFIXO",
                 "PRECODAVIAGEM", "QTD", "VLUN", "VL.UN"]


def _junto(l): return l.replace(" ", "")


def _num(s):
    s = s.strip()
    if "," in s: s = s.replace(".", "").replace(",", ".")
    return round(float(s), 2)


def _valores(linha):
    return [_num(m) for m in RE_VALOR.findall(linha.replace("R$", " ").replace("RS ", " "))]


def categoria(T):
    tudo = " " + " ".join(T) + " "
    for cat, palavras in CATEGORIAS:
        if any(p in tudo for p in palavras):
            return cat
    return ""


def valor(T):
    J = [_junto(l).replace("R$", "RS") for l in T]
    for chave in CHAVES_VALOR:
        rx = re.compile(chave.replace(r"R\$", "RS"))
        for i, l in enumerate(J):
            m = rx.search(l)
            if not m: continue
            depois = l[m.end():]
            if any(x in l[:m.start()] for x in ("IMPOSTO", "ISS", "TRIB")) and not depois: continue
            vs = [v for v in _valores(depois) if 0 < v < 20000]
            if vs: return vs[0]
            for j in range(i + 1, min(i + 3, len(J))):  # valor na linha de baixo (OCR quebra linhas)
                if any(x in J[j] for x in IGNORAR_VALOR): break
                vs = [v for v in _valores(J[j]) if 0 < v < 20000]
                if vs: return vs[0]
    # sem palavra-chave: o valor que mais se repete na nota (o total costuma aparecer várias vezes)
    vs = [v for l, lj in zip(T, J) if not any(x in lj for x in IGNORAR_VALOR) for v in _valores(l) if 0.5 < v < 20000]
    if not vs: return None
    c = Counter(vs).most_common()
    melhor = max(n for _, n in c)
    return max(v for v, n in c if n == melhor) if melhor > 1 else None


def pedagio_extrato(T):
    """Valores das passagens de um extrato de pedágio: linhas com data + PASSAGEM/PEDAGIO + valor no fim.
    Só vale com 2 ou mais passagens (um cupom de praça de pedágio usa o valor normal)."""
    out = []
    for l in T:
        if RE_DATA.search(l) and any(x in l for x in ("PASSAGEM", "PEDAGIO", "VIOLACAO")):
            vs = _valores(l)
            if vs and 0 < vs[-1] < 500: out.append(vs[-1])
    return out if len(out) >= 2 else []


def data_hora(T, ano_ref=None):
    datas = []
    for i, l in enumerate(T):
        for d, m, a in RE_DATA.findall(l):
            a = int(a); a = a + 2000 if a < 100 else a
            d, m = int(d), int(m)
            if 1 <= d <= 31 and 1 <= m <= 12 and 2020 <= a <= 2040:
                datas.append((i, f"{a:04d}-{m:02d}-{d:02d}", l))
    data = ""
    ano = Counter(d[:4] for _, d, _ in datas).most_common(1)[0][0] if datas else None
    for l in T:  # hospedagem: data de ENTRADA ("ABERTURA: 01/10/2026", "HOSPEDAGEM DO DIA 30/09 A 01/10")
        m = re.search(r"(?:ABERTURA|ENTRADA|CHECK-?IN|HOSPEDAGEMDODIA|DIARIADODIA|PERIODODE)[:.]?(\d{1,2})[/.](\d{1,2})(?:[/.](20\d{2}))?", _junto(l))
        if m and (m.group(3) or ano):
            d_, m_ = int(m.group(1)), int(m.group(2))
            if 1 <= d_ <= 31 and 1 <= m_ <= 12:
                data = f"{m.group(3) or ano}-{m_:02d}-{d_:02d}"; break
    if not data and datas:
        data = Counter(d for _, d, _ in datas).most_common(1)[0][0]
    hora = ""
    for i, d, l in datas:
        if d == data:
            h = RE_HORA.search(l[l.find(str(int(d[8:]))):] if l else "") or (RE_HORA.search(T[i + 1]) if i + 1 < len(T) else None)
            if h: hora = f"{int(h.group(1)):02d}:{h.group(2)}"; break
    return data, hora


def descricao(T_original):
    for l in T_original[:8]:
        u = _sem_acento(l)
        if len(re.sub(r"[^A-Z]", "", u)) >= 4 and not any(x in u for x in
                ("CNPJ", "DOCUMENTO", "DANFE", "NOTA FISCAL", "CUPOM", "VIA ", "IE:", "I.E", "CONSUMIDOR", "OBRIGADO", "CREDITO", "DEMONSTRATIVO")):
            return l.strip()[:60]
    return ""


def _data_do_nome(nome):
    m = re.search(r"(20\d{2})(\d{2})(\d{2})[_-]\d{4}", nome)
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else ""


def ler(path, imagem_tratada=None):
    import datetime as dt
    T0 = texto(path, imagem_tratada)
    T = [_sem_acento(l) for l in T0]
    data, hora = data_hora(T)
    foto = _data_do_nome(Path(path).name)
    if data and foto:  # data lida tem que ser coerente com a data da foto (OCR erra ano/mês)
        dias = (dt.date.fromisoformat(foto) - dt.date.fromisoformat(data)).days
        if not (0 <= dias <= 31):
            data, hora = "", ""
    cat = categoria(T)
    v = valor(T)
    passagens = pedagio_extrato(T)
    if passagens:
        cat, v = "Pedágio", round(sum(passagens), 2)
    if cat == "Transporte/táxi":  # recibo Uber: "Total R$ xx" é o cobrado
        for l in T:
            if l.startswith("TOTAL") and _valores(l): v = _valores(l)[-1]; break
    obs = ""
    if cat == "Outras Despesas":  # sugestão para a observação do relatório
        tudo = " ".join(T)
        for chave, txt in (("ESTACIONAMENTO", "Estacionamento"), ("STARLINK", "Starlink - internet (viagem)"),
                           ("LAVA", "Lavagem do veículo")):
            if chave in tudo:
                obs = txt + (f" - {descricao(T0)}" if chave != "STARLINK" and descricao(T0) else ""); break
    return {"categoria": cat, "valor": f"{v:.2f}" if v else "", "data": data, "hora": hora,
            "descricao": descricao(T0), "obs": obs, "cidades": cidades(T0), "linhas_lidas": len(T0)}


# ---------- cidade / UF onde a nota foi emitida (para sugerir a cidade do atendimento) ----------
UFS = ("AC AL AP AM BA CE DF ES GO MA MT MS MG PA PB PR PE PI RJ RN RS RO RR SC SP SE TO").split()
_RE_CIDADE = re.compile(r"([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'.]*(?:\s+[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ'.]*){0,4})\s*[-/,|]\s*(" + "|".join(UFS) + r")(?![A-Za-zÀ-ÿ])")
_NAO_CIDADE = {"SETOR", "CENTRO", "CENTRAL", "RUA", "R", "AV", "AVENIDA", "BAIRRO", "RODOVIA", "KM", "QUADRA", "QD", "LOTE",
               "LT", "CEP", "RESIDENCIAL", "RESID", "FONE", "TEL", "CNPJ", "SN", "S/N", "JARDIM", "VILA", "PARQUE", "ZONA",
               "RURAL", "URBANO", "PERIMETRO", "LTDA", "ME", "EIRELI", "SA", "S.A", "ENDERECO", "MUNICIPIO", "CIDADE",
               "LOCAL", "PRESTACAO", "UF", "ESTADO", "EMITIDA", "EM", "INCIDENCIA", "NORTE", "SUL", "LESTE", "OESTE",
               "FERROVIARIO", "AEROPORTO", "ALAMEDA", "AL", "PRACA", "ESTRADA", "MORADA", "NOVA", "LIBERDADE", "DEVIDO",
               "IMPOSTO", "SERVICO", "PRESTADO", "PAIS", "SIGLA", "GOV", "BR", "HTTP", "WWW"}


_PREFIXOS = {"RIO", "SAO", "SANTA", "SANTO", "PORTO", "CAMPO", "CAMPOS", "PONTA", "FOZ", "NOVA", "NOVO", "BOM", "BELO",
             "TRES", "PASSO", "BARRA", "LUIS", "ALTO", "SERRA", "MONTE", "POUSO", "VARZEA", "JOSE", "APARECIDA"}


def cidades(T0):
    out = []
    T0 = [re.sub(r"(?<=[A-Za-z])0|0(?=[A-Za-z])", "O", l) for l in T0]  # OCR troca O por zero ("RI0 VERDE")
    for n, l in enumerate(T0):
        for m in _RE_CIDADE.finditer(l):
            palavras = m.group(1).split()
            if m.start(1) == 0 and n > 0 and T0[n - 1].split():  # cidade quebrada em duas linhas
                ant = [w for w in re.split(r"[\s\-/,|]+", T0[n - 1]) if w]
                while ant and _sem_acento(ant[-1]).strip(".,") in _PREFIXOS | {"DE", "DOS", "DAS", "DA", "DO"}:
                    palavras.insert(0, ant.pop().strip(".,"))
            palavras = palavras[-4:]
            # tira palavras de endereço do começo ("SETOR CENTRAL, CROMÍNIA" -> "CROMÍNIA")
            while palavras and (_sem_acento(palavras[0]).strip(".") in _NAO_CIDADE or any(c.isdigit() for c in palavras[0])):
                palavras = palavras[1:]
            if not palavras: continue
            nome = " ".join(palavras).strip(" .'")
            chave = _sem_acento(nome)
            if len(chave.replace(" ", "")) < 3 or any(_sem_acento(p).strip(".") in _NAO_CIDADE for p in palavras[-1:]): continue
            if len(palavras) > 4 or any(len(p) > 16 for p in palavras): continue
            minus = {"de", "da", "do", "das", "dos", "e"}
            nome = " ".join(p.lower() if p.lower() in minus and i else p.capitalize() for i, p in enumerate(nome.lower().split()))
            out.append({"cidade": nome, "uf": m.group(2)})
    vistos, unicos = set(), []
    for c in out:
        k = (_sem_acento(c["cidade"]), c["uf"])
        if k not in vistos: vistos.add(k); unicos.append(c)
    return unicos
