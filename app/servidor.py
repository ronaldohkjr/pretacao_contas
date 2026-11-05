# Servidor do Gerador de Prestação de Contas — versão online (vários técnicos, com login).
# Cada técnico tem a sua pasta em DADOS/usuarios/<login>/ (viagem, notas, relatórios, histórico, "Meus dados").
import datetime as dt, hashlib, http.cookies, http.server, io, json, os, re, shutil, socketserver, subprocess, sys
import tempfile, threading, traceback, unicodedata, urllib.parse, zipfile
from pathlib import Path

import openpyxl
from PIL import Image, ImageOps
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

sys.path.insert(0, str(Path(__file__).resolve().parent))
import processamento as proc
import contas as _contas
import microsoft as ms

PORTA = int(os.environ.get("PORT", "8080"))
APP = Path(__file__).resolve().parent
RAIZ = APP.parent
DADOS = Path(os.environ.get("DADOS_DIR") or RAIZ / "dados")   # no servidor: um disco que não se apaga ao reiniciar
MODELO = RAIZ / "modelo" / "DESPESAS_MODELO.xlsx"
VERSAO = (RAIZ / "VERSAO.txt").read_text(encoding="utf-8").strip() if (RAIZ / "VERSAO.txt").exists() else ""
MAX_UPLOAD = 30 * 1024 * 1024   # por arquivo
CONTAS = _contas.Contas(DADOS)


class _Usuario(threading.local):
    """Pastas do técnico da requisição atual (cada requisição roda na sua thread)."""
    def usar(self, login):
        self.login = login
        self.BASE = DADOS / "usuarios" / login
        self.ENT, self.SAI, self.ARQ = self.BASE / "entradas", self.BASE / "saidas", self.BASE / "arquivo"
        self.ESTADO, self.CONFIG = self.BASE / "viagem.json", self.BASE / "config.json"
        self.ASSINATURA, self.HISTORICO = self.BASE / "assinatura.png", self.BASE / "historico.json"
        self.MS_TOKEN = self.BASE / "microsoft.json"
        self.TMP = DADOS / "cache" / login                         # cache (imagens tratadas): pode ser apagado
        for pasta in (self.ENT, self.SAI, self.TMP): pasta.mkdir(parents=True, exist_ok=True)


P = _Usuario()

CORPO_PADRAO = ("{APROVADOR},\nSegue em anexo para aprovação, planilha de despesas referente a viagem a {CIDADE}, "
                "atendimento a {CLIENTE}. Para {FINALIDADE}.\n\nEntre os dias {DATA_INICIAL} a {DATA_FINAL}")
CONFIG_PADRAO = {"nome": "", "nome_curto": "", "cpf": "", "centro_custo": "", "banco": "", "agencia": "", "conta": "",
                 "aprovador": "Adir", "email_para": "", "email_cc": "", "corpo_email": CORPO_PADRAO,
                 "limite_hospedagem_diaria": 120.0, "limite_refeicao_dia": 90.0}
LIMITES = {"Hospedagem": "limite_hospedagem_diaria", "Refeições": "limite_refeicao_dia"}  # categoria -> campo do config
CONFIG_OBRIGATORIOS = {"nome": "Nome completo", "nome_curto": "Nome curto", "cpf": "CPF", "centro_custo": "Centro de custo",
                       "banco": "Banco", "agencia": "Agência", "conta": "Conta", "aprovador": "Aprovador"}
CATEGORIAS = {  # categoria -> (linha na planilha, ordem no PDF)
    "Hospedagem": (21, 1), "Combustível": (33, 2), "Pedágio": (27, 3), "Refeições": (24, 4),
    "Material de Uso e Consumo": (36, 5), "Frete": (30, 6), "Locações de Veículos": (39, 6),
    "Transporte/táxi": (42, 6), "Outras Despesas": (45, 6)}
COLS_DIA = ["D", "F", "H", "J", "L", "N", "P", "R"]
EXT_OK = (".jpg", ".jpeg", ".png", ".pdf")
EXT_HEIC = (".heic", ".heif")
LIMITE_PROC = threading.Semaphore(2)  # no máximo 2 fotos sendo tratadas ao mesmo tempo


# ---------- estado ----------
def novo_id(): return dt.datetime.now().strftime("%Y%m%d%H%M%S%f")


def ler_estado():
    e = json.loads(P.ESTADO.read_text(encoding="utf-8-sig")) if P.ESTADO.exists() else         {"cliente": "", "cidade": "", "uf": "", "finalidade": "", "inicio": "", "fim": "", "linhas": []}
    if not e.get("id"):
        e["id"] = novo_id(); P.ESTADO.write_text(json.dumps(e, ensure_ascii=False, indent=1), encoding="utf-8")
    return e


def aba_desatualizada(e):
    """Outra aba começou uma nova viagem: não deixa esta aba antiga gravar por cima."""
    return e.get("id") != ler_estado().get("id")


def gravar(arq, texto):
    """Grava o arquivo; se o OneDrive/antivírus estiver com ele preso, tenta de novo por alguns segundos."""
    import time
    for i in range(10):
        try:
            arq.write_text(texto, encoding="utf-8"); return
        except PermissionError:
            if i == 9: raise
            time.sleep(0.5)


def salvar_estado(e):
    gravar(P.ESTADO, json.dumps(e, ensure_ascii=False, indent=1))


def arquivos_entrada():
    return sorted(f.name for f in P.ENT.iterdir() if f.is_file() and f.suffix.lower() in EXT_OK)


# ---------- Meus dados (config.json) ----------
def ler_config():
    c = dict(CONFIG_PADRAO)
    if P.CONFIG.exists():
        c.update(json.loads(P.CONFIG.read_text(encoding="utf-8-sig")))
    if not str(c.get("corpo_email", "")).strip(): c["corpo_email"] = CORPO_PADRAO
    for k in LIMITES.values():
        try: c[k] = float(c.get(k) or 0)
        except (TypeError, ValueError): c[k] = CONFIG_PADRAO[k]
    return c


def politica_limites(c=None):
    """Limite por dia de cada categoria da política (hospedagem: por diária). 0 = sem limite."""
    c = c or ler_config()
    return {cat: c[k] for cat, k in LIMITES.items()}


def faltando_config(c):
    return [rot for k, rot in CONFIG_OBRIGATORIOS.items() if not str(c.get(k, "")).strip()]


def cpf_valido(cpf):
    d = [int(x) for x in re.sub(r"\D", "", cpf)]
    if len(d) != 11 or len(set(d)) == 1: return False
    for n in (9, 10):
        if (sum(v * (n + 1 - i) for i, v in enumerate(d[:n])) * 10 % 11) % 10 != d[n]: return False
    return True


def salvar_config(novo):
    c = ler_config()
    for k in CONFIG_PADRAO:
        if k in novo: c[k] = novo[k] if k in LIMITES.values() else str(novo[k]).strip()
    c["nome_curto"] = slug(c["nome_curto"])
    cpf = re.sub(r"\D", "", c["cpf"])
    if c["cpf"] and not cpf_valido(cpf): return {"ok": False, "erros": ["CPF inválido: confira os números."]}
    if len(cpf) == 11: c["cpf"] = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    for k in ("email_para", "email_cc"):  # vários e-mails: separados por ; (como no Outlook)
        c[k] = "; ".join(x.strip() for x in re.split(r"[;,]", c[k]) if x.strip())
        if any(not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", x) for x in c[k].split("; ") if x):
            return {"ok": False, "erros": [f"E-mail inválido em {'Para' if k == 'email_para' else 'Cópia'}: confira o endereço."]}
    for k, rot in (("limite_hospedagem_diaria", "Limite de hospedagem: digite um valor, ex.: 120"),
                   ("limite_refeicao_dia", "Limite de refeição: digite um valor, ex.: 90")):
        try: c[k] = round(float(str(c[k] or 0).replace(",", ".")), 2)
        except ValueError: return {"ok": False, "erros": [rot]}
    P.CONFIG.write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")
    return {"ok": True, "config": c, "faltando": faltando_config(c)}


def salvar_assinatura(dados):
    """Guarda a assinatura como PNG com fundo transparente, recortada nas bordas."""
    import numpy as np
    im = ImageOps.exif_transpose(Image.open(io.BytesIO(dados))).convert("RGBA"); im.thumbnail((1600, 1600))
    a = np.array(im); a[(a[..., :3] > 200).all(-1), 3] = 0  # fundo claro -> transparente
    im = Image.fromarray(a); caixa = im.getbbox()
    if not caixa: raise ValueError("A imagem da assinatura parece estar em branco.")
    im = im.crop(caixa); im.thumbnail((900, 400)); im.save(P.ASSINATURA)


def numero_ou_texto(v):
    """Centro de custo/agência: número na planilha (como no modelo), a não ser que comece com zero."""
    v = str(v).strip()
    return int(v) if v.isdigit() and not v.startswith("0") else v


# ---------- imagens ----------
def processado(nome, modo):
    """Páginas tratadas, com cache em disco (por arquivo + data de modificação + modo)."""
    src = P.ENT / nome
    chave = hashlib.md5(f"{nome}|{src.stat().st_mtime}|{modo}".encode()).hexdigest()
    pags = sorted(P.TMP.glob(chave + "_*.png"))
    if not pags:
        with LIMITE_PROC:
            pags = sorted(P.TMP.glob(chave + "_*.png"))
            if not pags:
                for i, im in enumerate(proc.paginas(src, modo)):
                    im.convert("L").save(P.TMP / f"{chave}_{i:02d}.png")
                pags = sorted(P.TMP.glob(chave + "_*.png"))
    return pags


def preview_jpg(nome, modo, largura=700):
    im = Image.open(processado(nome, modo)[0]).convert("L")
    im.thumbnail((largura, largura * 2))
    b = io.BytesIO(); im.save(b, "JPEG", quality=80); return b.getvalue()


# ---------- utilidades ----------
def slug(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return "_".join(s.upper().split())


def data(s): return dt.date.fromisoformat(s)


def br(v): return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def aplicar_periodo(e):
    """Data inicial e final = nota mais antiga e mais recente entre as incluídas."""
    ds = sorted(l["data"] for l in linhas_incluidas(e) if l.get("data"))
    if ds: e["inicio"], e["fim"] = ds[0], ds[-1]
    return e


def dias_viagem(e):
    ini, fim = data(e["inicio"]), data(e["fim"])
    return [ini + dt.timedelta(days=i) for i in range((fim - ini).days + 1)]


def validar(e):
    erros = []
    falta = faltando_config(ler_config())
    if falta: erros.append("Preencha seus dados em \"Meus dados\" (botão no topo): " + ", ".join(falta) + ".")
    for k, rot in (("cliente", "Cliente"), ("cidade", "Cidade"), ("uf", "UF"), ("finalidade", "Finalidade"),
                   ("inicio", "Data das notas"), ("fim", "Data das notas")):
        if not str(e.get(k, "")).strip(): erros.append(f"Preencha: {rot}.")
    if erros: return erros
    dias = dias_viagem(e)
    if len(dias) > 8: erros.append(f"As notas vão de {dias[0]:%d/%m} a {dias[-1]:%d/%m} ({len(dias)} dias); a planilha aceita no máximo 8. Confira as datas ou desmarque notas de outra viagem.")
    if len(dias) < 1: erros.append("Data final antes da inicial.")
    for i, l in enumerate(linhas_incluidas(e), 1):
        nome = l.get("arquivo") or f"linha {i}"
        if l.get("categoria") not in CATEGORIAS: erros.append(f"{nome}: escolha a categoria.")
        if str(l.get("valor", "")).strip() == "" or float(l["valor"]) < 0: erros.append(f"{nome}: informe o valor (0 = só anexar ao PDF).")
        if not l.get("data"): erros.append(f"{nome}: informe a data.")
    for l in linhas_incluidas(e):
        if l.get("categoria") == "Outras Despesas" and not str(l.get("obs", "")).strip():
            erros.append(f"{l.get('arquivo') or 'nota'}: Outras Despesas — escreva o que foi (vai na observação do relatório).")
    if sum(1 for l in linhas_incluidas(e) if l.get("categoria") == "Outras Despesas" and float(l.get("valor") or 0) > 0) > 3:
        erros.append("Outras Despesas: a planilha só tem 3 linhas de detalhe.")
    return erros


def linhas_incluidas(e): return [l for l in e.get("linhas", []) if l.get("incluir", True)]


def ordenar(linhas):
    return sorted(linhas, key=lambda l: (l["data"], CATEGORIAS[l["categoria"]][1], l.get("hora") or "99:99"))


def _diarias(l):
    try: return max(1, int(float(l.get("diarias") or 1)))
    except (TypeError, ValueError): return 1


def politica(e):
    """Dias/categorias acima do limite da política e o valor reembolsado de cada nota.
    Decisão por dia (e["excessos"]["AAAA-MM-DD|Categoria"]): justificar com motivo (vai no e-mail);
    sem motivo, o reembolso é ajustado ao limite automaticamente, tirando primeiro da nota de maior valor."""
    dec = e.get("excessos") or {}
    lims = politica_limites()
    inc = [l for l in linhas_incluidas(e) if str(l.get("valor", "")).strip()]
    reemb = {id(l): round(float(l["valor"]), 2) for l in inc}
    grupos = {}
    for l in inc:
        if lims.get(l.get("categoria")) and l.get("data") and float(l["valor"]) > 0:
            grupos.setdefault((l["data"], l["categoria"]), []).append(l)
    excessos = []
    for (d, cat), ls in sorted(grupos.items()):
        total = round(sum(float(l["valor"]) for l in ls), 2)
        diarias = sum(_diarias(l) for l in ls) if cat == "Hospedagem" else 1
        limite = round(lims[cat] * diarias, 2)
        if total <= limite + 0.004: continue
        chave = f"{d}|{cat}"; dd = dec.get(chave) or {}
        motivo = str(dd.get("motivo", "")).strip()
        modo = "justificar" if dd.get("modo") == "justificar" and motivo else "ajustar"
        ajustes = []
        if modo == "ajustar":
            falta = round(total - limite, 2)
            for l in sorted(ls, key=lambda x: -float(x["valor"])):
                if falta <= 0: break
                tira = min(falta, reemb[id(l)])
                reemb[id(l)] = round(reemb[id(l)] - tira, 2); falta = round(falta - tira, 2)
                ajustes.append({"arquivo": l.get("arquivo"), "descricao": l.get("descricao", ""),
                                "de": round(float(l["valor"]), 2), "para": reemb[id(l)]})
        excessos.append({"chave": chave, "data": d, "categoria": cat, "total": total, "limite": limite,
                         "diarias": diarias, "excesso": round(total - limite, 2), "modo": modo,
                         "escolha": dd.get("modo") or "ajustar", "motivo": motivo, "ajustes": ajustes})
    return excessos, reemb


def resumo(e):
    dias = [d.isoformat() for d in dias_viagem(e)]
    tab = {d: {} for d in dias}
    for l in linhas_incluidas(e):
        if l.get("data") in tab and l.get("categoria") in CATEGORIAS and l.get("valor"):
            tab[l["data"]][l["categoria"]] = round(tab[l["data"]].get(l["categoria"], 0) + float(l["valor"]), 2)
    alertas = []
    for x in politica(e)[0]:
        if x["modo"] == "ajustar":
            alertas.append(f"{data(x['data']):%d/%m} — {x['categoria']}: R$ {br(x['total'])} ajustado ao limite de R$ {br(x['limite'])}.")
        else:
            alertas.append(f"{data(x['data']):%d/%m} — {x['categoria']}: R$ {br(x['total'])} acima do limite, justificado no e-mail.")
    return tab, alertas


# ---------- geração ----------
def gerar(e):
    aplicar_periodo(e)
    erros = validar(e)
    if erros: return {"ok": False, "erros": erros}
    dias = dias_viagem(e); cfg = ler_config()
    base = f"{dias[0]:%Y%m%d}_%s_{slug(cfg['nome_curto'])}_{slug(e['cliente'])}_{slug(e['cidade'])}_{slug(e['uf'])}"
    linhas = ordenar(linhas_incluidas(e))
    excessos, reemb = politica(e)

    # Aviso na página da nota cujo reembolso foi ajustado à política (o valor que entra não é o total da nota)
    avisos = {}
    for x in excessos:
        lim_txt = (f"limite R$ {br(x['limite'])} para {x['diarias']} diárias" if x["diarias"] > 1 else
                   f"limite R$ {br(x['limite'])} por {'diária' if x['categoria'] == 'Hospedagem' else 'dia'}")
        for a in x["ajustes"]:
            if a.get("arquivo"):
                avisos.setdefault(a["arquivo"], []).append(
                    f"REEMBOLSO AJUSTADO À POLÍTICA: será reembolsado R$ {br(a['para'])} desta nota de R$ {br(a['de'])} "
                    f"({x['categoria']} em {data(x['data']):%d/%m}: total do dia R$ {br(x['total'])}, {lim_txt}).")

    # PDF: uma nota por página, na ordem; arquivo repetido entra uma vez só
    from reportlab.lib.utils import simpleSplit
    pdf = P.SAI / (base % "NOTAS" + ".pdf")
    PW, PH = A4; m = 20
    c = canvas.Canvas(str(pdf), pagesize=A4); vistos = set(); n_pag = 0
    for l in linhas:
        nome = l.get("arquivo")
        if not nome or nome in vistos or not (P.ENT / nome).exists(): continue
        vistos.add(nome)
        for k, p in enumerate(processado(nome, l.get("modo", "cupom"))):
            im = Image.open(p).convert("L")
            if max(im.size) > 1700:
                r = 1700 / max(im.size); im = im.resize((int(im.size[0] * r), int(im.size[1] * r)))
            j = P.TMP / f"pag_{n_pag:03d}.jpg"; im.save(j, quality=88)
            txt = [t for a in avisos.get(nome, []) if k == 0 for t in simpleSplit(a, "Helvetica-Bold", 10.5, PW - 2 * m - 20)]
            faixa = 16 * len(txt) + 18 if txt else 0  # espaço reservado no topo para o aviso
            if txt:
                c.setStrokeColorRGB(0.75, 0.1, 0.1); c.setFillColorRGB(1, 0.94, 0.94); c.setLineWidth(1.5)
                c.rect(m, PH - m - faixa + 6, PW - 2 * m, faixa - 6, fill=1)
                c.setFillColorRGB(0.7, 0.05, 0.05); c.setFont("Helvetica-Bold", 10.5)
                for n, t in enumerate(txt): c.drawString(m + 10, PH - m - 16 - 16 * n, t)
            iw, ih = im.size; s = min((PW - 2 * m) / iw, (PH - 2 * m - faixa) / ih); dw, dh = iw * s, ih * s
            c.drawImage(str(j), (PW - dw) / 2, (PH - faixa - dh) / 2, width=dw, height=dh); c.showPage(); n_pag += 1
    c.save()

    # Planilha a partir do modelo
    wb = openpyxl.load_workbook(MODELO); ws = wb.active
    ws["B8"] = cfg["nome"]; ws["I8"] = cfg["cpf"]; ws["B10"] = numero_ou_texto(cfg["centro_custo"])
    ws["B12"] = cfg["banco"]; ws["K12"] = numero_ou_texto(cfg["agencia"]); ws["P12"] = cfg["conta"]
    if P.ASSINATURA.exists():  # assinatura do técnico no campo "Assinatura" (G64:K68), como no modelo original
        from openpyxl.drawing.image import Image as XLImage
        img = XLImage(str(P.ASSINATURA)); r = min(178 / img.width, 74 / img.height)
        img.width, img.height = int(img.width * r), int(img.height * r)
        ws.add_image(img, "G65")
    for col, d in zip(COLS_DIA, dias):
        ws[col + "19"] = dt.datetime.combine(d, dt.time())
    ws["E10"] = dt.datetime.combine(dias[0], dt.time()); ws["I10"] = dt.datetime.combine(dias[-1], dt.time())
    ws["B14"] = f"Atendimento {e['cliente']} {e['cidade']}/{e['uf']}"
    soma = {}
    for l in linhas:
        if float(l["valor"]) == 0: continue
        cel = COLS_DIA[dias.index(data(l["data"]))] + str(CATEGORIAS[l["categoria"]][0])
        soma.setdefault(cel, []).append(reemb.get(id(l), round(float(l["valor"]), 2)))  # valor ajustado à política
    for cel, vs in soma.items():
        ws[cel] = vs[0] if len(vs) == 1 else "=" + "+".join(f"{v:g}" for v in vs)
    outras = [l for l in linhas if l["categoria"] == "Outras Despesas" and float(l["valor"]) > 0]
    for i, l in enumerate(outras):
        r = 51 + i
        ws[f"B{r}"] = dt.datetime.combine(data(l["data"]), dt.time())
        ws[f"D{r}"] = str(l.get("obs", "")).strip()
        ws[f"X{r}"] = reemb.get(id(l), round(float(l["valor"]), 2))
    for row in ws.iter_rows():
        for cell in row:
            f = (cell.number_format or "").lower()
            if "yy" in f and ("mm" in f or "dd" in f or "m" in f): cell.number_format = "dd/mm/yy"
    ws["C66"] = dt.datetime.combine(dt.date.today(), dt.time())  # data ao lado da assinatura = dia da geração
    ws["C66"].number_format = "dd/mm/yy"
    xlsx = P.SAI / (base % "DESPESAS" + ".xlsx"); wb.save(xlsx)
    xlsx_pdf = xlsx.with_suffix(".pdf"); aviso_pdf = planilha_em_pdf(xlsx, xlsx_pdf)  # planilha também em PDF

    total_notas = round(sum(float(l["valor"]) for l in linhas), 2)
    total = round(sum(reemb.get(id(l), float(l["valor"])) for l in linhas), 2)
    assunto = f"Relatório de Despesas atendimento {e['cliente']} {e['cidade']}/{e['uf']}"
    corpo = cfg["corpo_email"]
    for k, v in {"APROVADOR": cfg["aprovador"], "CIDADE": e["cidade"], "UF": e["uf"], "CLIENTE": e["cliente"],
                 "FINALIDADE": e["finalidade"], "DATA_INICIAL": f"{dias[0]:%d/%m/%Y}", "DATA_FINAL": f"{dias[-1]:%d/%m/%Y}",
                 "NOME": cfg["nome"]}.items():
        corpo = corpo.replace("{" + k + "}", str(v))
    def lim(x): return f"limite R$ {br(x['limite'])}" + (f" para {x['diarias']} diárias" if x["diarias"] > 1 else "")
    just = [x for x in excessos if x["modo"] == "justificar"]
    if just:
        corpo += "\n\nJustificativa dos gastos acima da política:\n" + "\n".join(
            f"- {data(x['data']):%d/%m} – {x['categoria']}: R$ {br(x['total'])} ({lim(x)}): {x['motivo']}" for x in just)
    ajust = [x for x in excessos if x["modo"] == "ajustar"]
    if ajust:
        corpo += "\n\nValores ajustados ao limite da política:\n" + "\n".join(
            f"- {data(x['data']):%d/%m} – {x['categoria']}: R$ {br(x['total'])} → R$ {br(x['limite'])}" for x in ajust)
    tab, alertas = resumo(e)
    if aviso_pdf: alertas = alertas + [f"A planilha em PDF não foi gerada: {aviso_pdf}. A planilha .xlsx está normal."]
    r = {"ok": True, "pdf": pdf.name, "xlsx": xlsx.name, "xlsx_pdf": "" if aviso_pdf else xlsx_pdf.name, "paginas": n_pag,
         "pdf_mb": round(pdf.stat().st_size / 1e6, 1), "total": total, "total_notas": total_notas,
         "assunto": assunto, "corpo": corpo, "alertas": alertas}
    registrar_geracao(e, r)
    return r


# ---------- histórico de relatórios gerados / enviados ----------
_hist_lock = threading.Lock()


def ler_hist():
    return json.loads(P.HISTORICO.read_text(encoding="utf-8-sig")) if P.HISTORICO.exists() else []


def salvar_hist(h):
    gravar(P.HISTORICO, json.dumps(h, ensure_ascii=False, indent=1))


def agora(): return dt.datetime.now().isoformat(timespec="seconds")


def registrar_geracao(e, r):
    with _hist_lock:
        h = ler_hist()
        it = next((x for x in h if x["id"] == e["id"]), None)
        if it is None:
            it = {"id": e["id"], "enviado": False, "enviado_em": "", "geracoes": 0}; h.append(it)
        it.update(cliente=e["cliente"], cidade=e["cidade"], uf=e["uf"], finalidade=e["finalidade"],
                  inicio=e["inicio"], fim=e["fim"], total=r["total"], notas=len(linhas_incluidas(e)),
                  pdf=r["pdf"], xlsx=r["xlsx"], pasta="saidas", gerado_em=agora(), geracoes=it.get("geracoes", 0) + 1)
        salvar_hist(h)


def marcar_enviado(id_, enviado):
    with _hist_lock:
        h = ler_hist()
        it = next((x for x in h if x["id"] == id_), None)
        if it is None: raise KeyError("viagem não está no histórico")
        it["enviado"] = bool(enviado); it["enviado_em"] = agora() if enviado else ""
        enxugar(it)
        salvar_hist(h); return it


def historico():
    h = ler_hist()
    for it in h:  # gerou de novo depois de marcar como enviado: avisa
        it["gerado_depois_do_envio"] = bool(it.get("enviado") and it.get("gerado_em", "") > it.get("enviado_em", ""))
    return sorted(h, key=lambda x: x.get("inicio") or "", reverse=True)


# ---------- limpeza: cache temporário e viagens já enviadas ----------
def limpar_cache(minutos, pasta=None):
    """Apaga do cache temporário (imagens tratadas, texto lido, páginas do PDF) o que não foi mexido há mais de
    `minutos`. É só cache: o programa refaz o que precisar."""
    import time
    limite = time.time() - minutos * 60
    for f in list((pasta or P.TMP).iterdir()):
        try:
            if f.is_file() and f.stat().st_mtime < limite: f.unlink()
        except OSError:
            pass


def apagar_arquivos(arquivos):
    """Apaga os arquivos (no servidor não há Lixeira). Devolve os nomes que não deu para apagar."""
    falhas = []
    for f in arquivos:
        try: Path(f).unlink()
        except OSError: falhas.append(Path(f).name)
    return falhas


def enxugar(it):
    """Viagem enviada e arquivada: as fotos/PDFs originais das notas são apagados e ficam o PDF das notas,
    a planilha e o viagem.json (o PDF já tem todas as notas). Só mexe se o PDF e a planilha estiverem lá."""
    if not it.get("enviado") or it.get("enxugado"): return
    pasta = P.BASE / it.get("pasta", "saidas")
    if pasta.parent != P.ARQ or not pasta.is_dir(): return  # ainda não foi arquivada
    sai, ent = pasta / "saidas", pasta / "entradas"
    if not (any(sai.glob("*_NOTAS_*.pdf")) and any(sai.glob("*_DESPESAS_*.xlsx"))): return
    fotos = [f for f in ent.iterdir() if f.is_file()] if ent.is_dir() else []
    tamanho = sum(f.stat().st_size for f in fotos)
    if apagar_arquivos(fotos): return  # algum arquivo preso: tenta de novo na próxima vez
    try: ent.rmdir()
    except OSError: pass
    it.update(enxugado=True, enxugado_em=agora(), liberado_mb=round(tamanho / 1e6, 1))



def planilha_em_pdf(xlsx, pdf):
    """Converte a planilha em PDF. No servidor, pelo LibreOffice (recalcula as fórmulas);
    num Windows sem LibreOffice (teste local), pelo Excel. Devolve "" se gerou, ou o motivo."""
    pdf.unlink(missing_ok=True)
    soffice = shutil.which("soffice") or shutil.which("libreoffice") or \
        next((c for c in (r"C:\\Program Files\\LibreOffice\\program\\soffice.exe",) if Path(c).exists()), None)
    if soffice:
        perfil = P.TMP / "lo_perfil"  # perfil próprio: duas conversões ao mesmo tempo não se atrapalham
        try:
            subprocess.run([soffice, f"-env:UserInstallation={perfil.as_uri()}", "--headless", "--norestore",
                            "--convert-to", "pdf", "--outdir", str(pdf.parent), str(xlsx)],
                           capture_output=True, timeout=180)
        except Exception as ex:
            return f"não consegui chamar o LibreOffice ({ex})"
        return "" if pdf.exists() else "o LibreOffice não conseguiu converter a planilha"
    if os.name == "nt":
        ps = ("$ErrorActionPreference='Stop'; $xl = New-Object -ComObject Excel.Application;"
              "$xl.Visible = $false; $xl.DisplayAlerts = $false;"
              "try { $wb = $xl.Workbooks.Open($env:GNF_XLSX, 0, $true); $wb.ExportAsFixedFormat(0, $env:GNF_PDF); $wb.Close($false) }"
              "finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }")
        try:
            subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps], capture_output=True, timeout=180,
                           env=dict(os.environ, GNF_XLSX=str(xlsx), GNF_PDF=str(pdf)), creationflags=0x08000000)
        except Exception as ex:
            return f"não consegui chamar o Excel ({ex})"
        return "" if pdf.exists() else "o Excel não está instalado ou deu erro ao exportar"
    return "o LibreOffice não está instalado no servidor"


def anexos_do_email(c):
    anexos = [P.SAI / Path(c[k]).name for k in ("pdf", "xlsx_pdf") if c.get(k)]
    if not c.get("xlsx_pdf") and c.get("xlsx"): anexos.append(P.SAI / Path(c["xlsx"]).name)  # sem PDF da planilha: vai o .xlsx
    return anexos


def abrir_email(c):
    """Cria o rascunho (não envia) na caixa do técnico pela conta Microsoft, com os PDFs anexados.
    Sem a conta Microsoft configurada no servidor: devolve um link mailto (para/assunto/texto) e o zip para anexar à mão."""
    import html
    cfg = ler_config()
    anexos = anexos_do_email(c)
    faltam = [f.name for f in anexos if not f.is_file()]
    if faltam: return {"ok": False, "erro": "Não encontrei: " + ", ".join(faltam) + ". Gere o relatório de novo."}
    if ms.configurado():
        if not ms.conectado(P.MS_TOKEN): return {"ok": False, "conectar": True}
        corpo_html = ('<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' +
                      html.escape(c.get("corpo", "")).replace("\n", "<br>") + "</div>")
        try:
            link = ms.criar_rascunho(P.MS_TOKEN, cfg.get("email_para", ""), cfg.get("email_cc", ""), c.get("assunto", ""),
                                     corpo_html, anexos)
        except RuntimeError as ex:
            if "401" in str(ex) or "invalid_grant" in str(ex):  # permissão expirou/revogada: conectar de novo
                ms.desconectar(P.MS_TOKEN); return {"ok": False, "conectar": True}
            raise
        return {"ok": True, "como": "outlook", "link": link, "para": cfg.get("email_para", "")}
    q = urllib.parse.urlencode({"subject": c.get("assunto", ""), "body": c.get("corpo", ""),
                                **({"cc": cfg["email_cc"]} if cfg.get("email_cc") else {})}, quote_via=urllib.parse.quote)
    return {"ok": True, "como": "mailto", "para": cfg.get("email_para", ""),
            "mailto": "mailto:" + urllib.parse.quote(cfg.get("email_para", ""), safe="@;") + "?" + q}


# ---------- nova viagem ----------
def arquivar(e):
    aplicar_periodo(e)
    vazio = {"id": novo_id(), "cliente": "", "cidade": "", "uf": "", "finalidade": "", "inicio": "", "fim": "", "linhas": []}
    if not any(f.is_file() for pasta in (P.ENT, P.SAI) for f in pasta.iterdir()) and not linhas_incluidas(e):
        salvar_estado(vazio); return {"ok": True, "pasta": "", "falhas": []}  # nada para guardar: só esvazia
    nome = f"{(e.get('inicio') or dt.date.today().isoformat()).replace('-', '')}_{slug(e.get('cliente') or 'VIAGEM')}"
    dest = P.ARQ / nome; (dest / "entradas").mkdir(parents=True, exist_ok=True); (dest / "saidas").mkdir(exist_ok=True)
    falhas = []
    for pasta, sub in ((P.ENT, "entradas"), (P.SAI, "saidas")):
        for f in list(pasta.iterdir()):
            if not f.is_file(): continue
            try:
                shutil.copy2(f, dest / sub / f.name); f.unlink()
            except Exception:
                falhas.append(f.name)
    if P.ESTADO.exists(): shutil.copy2(P.ESTADO, dest / "viagem.json")
    with _hist_lock:  # o histórico passa a apontar para a pasta arquivada
        h = ler_hist(); it = next((x for x in h if x["id"] == e.get("id")), None)
        if it: it["pasta"] = str(dest.relative_to(P.BASE)); enxugar(it)
        elif linhas_incluidas(e):
            h.append({"id": e.get("id"), "cliente": e.get("cliente", ""), "cidade": e.get("cidade", ""), "uf": e.get("uf", ""),
                      "inicio": e.get("inicio", ""), "fim": e.get("fim", ""), "notas": len(linhas_incluidas(e)),
                      "pasta": str(dest.relative_to(P.BASE)), "gerado_em": "", "enviado": False, "enviado_em": "", "geracoes": 0})
        salvar_hist(h)
    salvar_estado({"id": novo_id(), "cliente": "", "cidade": "", "uf": "", "finalidade": "", "inicio": "", "fim": "",
                   "linhas": [l for l in e.get("linhas", []) if l.get("arquivo") in falhas]})
    limpar_cache(10)  # as notas desta viagem saíram do programa: cache delas não serve mais
    return {"ok": True, "pasta": str(dest.relative_to(P.BASE)), "falhas": falhas,
            "enxugado": bool(it and it.get("enxugado")), "liberado_mb": (it or {}).get("liberado_mb", 0)}



def zip_da_viagem(id_):
    """Relatórios (PDF/planilha) de uma viagem do histórico, num zip para baixar."""
    it = next(x for x in ler_hist() if x["id"] == id_)
    pasta = P.BASE / it.get("pasta", "saidas")
    if (pasta / "saidas").is_dir(): pasta = pasta / "saidas"
    if pasta.resolve().parent != P.BASE.resolve() and P.BASE.resolve() not in pasta.resolve().parents:
        raise ValueError("pasta inválida")
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(pasta.iterdir()):
            if f.is_file() and f.suffix.lower() in (".pdf", ".xlsx"): z.write(f, f.name)
    nome = f"{(it.get('inicio') or '').replace('-', '')}_RELATORIO_{slug(it.get('cliente') or 'VIAGEM')}.zip"
    return nome, b.getvalue()


# ---------- HTTP ----------
PUBLICAS = {"/login", "/api/login", "/api/primeiro_acesso", "/api/ping"}


class H(http.server.BaseHTTPRequestHandler):
    server_version = "Gerador"
    sys_version = ""

    def log_message(self, fmt, *a):
        if " 5" in (fmt % a): sys.stderr.write("%s %s\n" % (self.address_string(), fmt % a))

    # ---------- respostas ----------
    def _cabecalhos_seguranca(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Cache-Control", "no-store")

    def enviar(self, corpo, tipo="application/json; charset=utf-8", status=200, extra=None):
        if isinstance(corpo, (dict, list)): corpo = json.dumps(corpo, ensure_ascii=False).encode()
        elif isinstance(corpo, str): corpo = corpo.encode()
        self.send_response(status); self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(corpo))); self._cabecalhos_seguranca()
        for k, v in (extra or {}).items(): self.send_header(k, v)
        self.end_headers(); self.wfile.write(corpo)

    def baixar(self, nome, dados, tipo):
        self.enviar(dados, tipo, extra={"Content-Disposition": "attachment; filename*=UTF-8''" + urllib.parse.quote(nome)})

    def redirecionar(self, url, extra=None):
        self.send_response(303); self.send_header("Location", url); self._cabecalhos_seguranca()
        for k, v in (extra or {}).items(): self.send_header(k, v)
        self.send_header("Content-Length", "0"); self.end_headers()

    def corpo_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n > 2 * 1024 * 1024: raise ValueError("pedido grande demais")
        return json.loads(self.rfile.read(n) or b"{}")

    # ---------- sessão ----------
    def ip(self):
        return (self.headers.get("X-Forwarded-For") or self.client_address[0]).split(",")[0].strip()

    def https(self):
        return self.headers.get("X-Forwarded-Proto", "") == "https" or ms.URL_PUBLICA.startswith("https://")

    def cookie_sessao(self, token, idade):
        c = f"sessao={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={idade}"
        return c + ("; Secure" if self.https() else "")

    def token(self):
        try: return http.cookies.SimpleCookie(self.headers.get("Cookie", "")).get("sessao").value
        except AttributeError: return ""

    def autenticar(self, caminho, metodo):
        """Confere a sessão. Devolve a sessão (e prepara as pastas do técnico) ou None (já respondeu)."""
        s = CONTAS.sessao(self.token())
        if s is None:
            if caminho.startswith("/api/"): self.enviar({"erro": "SESSAO"}, status=401)
            else: self.redirecionar("/login")
            return None
        # pedidos que alteram dados só valem vindos da própria página (proteção contra CSRF)
        if metodo == "POST" and self.headers.get("X-Gerador") != "1":
            self.enviar({"erro": "pedido recusado"}, status=403); return None
        if s["trocar_senha"] and caminho not in ("/", "/api/eu", "/api/senha", "/api/logout"):
            self.enviar({"erro": "TROCAR_SENHA"}, status=403); return None
        P.usar(s["login"])
        return s

    # ---------- GET ----------
    def do_GET(self):
        u = urllib.parse.urlparse(self.path); q = dict(urllib.parse.parse_qsl(u.query))
        try:
            if u.path == "/api/ping": return self.enviar({"ok": True})
            if u.path == "/login":
                if CONTAS.sessao(self.token()): return self.redirecionar("/")
                return self.enviar((APP / "login.html").read_bytes(), "text/html; charset=utf-8")
            s = self.autenticar(u.path, "GET")
            if s is None: return
            if u.path == "/":
                return self.enviar((APP / "index.html").read_bytes(), "text/html; charset=utf-8")
            if u.path == "/api/eu":
                return self.enviar({**s, "microsoft": ms.configurado(), "ms_conta": ms.conectado(P.MS_TOKEN) if ms.configurado() else ""})
            if u.path == "/api/estado":
                return self.enviar({"estado": ler_estado(), "arquivos": arquivos_entrada(),
                                    "categorias": list(CATEGORIAS), "politica": politica_limites()})
            if u.path == "/api/config":
                c = ler_config()
                return self.enviar({"config": c, "faltando": faltando_config(c), "existe": P.CONFIG.exists(),
                                    "assinatura": P.ASSINATURA.exists(), "versao": VERSAO, "corpo_padrao": CORPO_PADRAO})
            if u.path == "/api/assinatura":
                return self.enviar(P.ASSINATURA.read_bytes(), "image/png")
            if u.path == "/api/preview":
                return self.enviar(preview_jpg(Path(q["arquivo"]).name, q.get("modo", "cupom"), min(int(q.get("w", 700)), 1600)), "image/jpeg")
            if u.path == "/api/historico":
                return self.enviar(historico())
            if u.path == "/api/historico/baixar":
                nome, dados = zip_da_viagem(q["id"]); return self.baixar(nome, dados, "application/zip")
            if u.path == "/api/ultimo":
                arq = P.BASE / "ultimo.json"
                return self.enviar(json.loads(arq.read_text(encoding="utf-8")) if arq.exists() else {})
            if u.path == "/api/baixar_zip":  # PDF + planilha (+ planilha em PDF) num .zip só
                fs = [P.SAI / Path(q.get(k, "")).name for k in ("pdf", "xlsx")]
                if q.get("xlsx_pdf") and (P.SAI / Path(q["xlsx_pdf"]).name).is_file(): fs.append(P.SAI / Path(q["xlsx_pdf"]).name)
                if not all(f.is_file() for f in fs[:2]): return self.enviar({"erro": "arquivo não encontrado — gere de novo"}, status=404)
                b = io.BytesIO()
                with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
                    for f in fs: z.write(f, f.name)
                return self.baixar(fs[1].stem.replace("_DESPESAS_", "_RELATORIO_") + ".zip", b.getvalue(), "application/zip")
            if u.path == "/api/baixar":  # PDF/planilha gerados (só arquivos da pasta saidas do técnico)
                f = P.SAI / Path(q["arquivo"]).name
                if not f.is_file(): return self.enviar({"erro": "arquivo não encontrado — gere de novo"}, status=404)
                tipo = "application/pdf" if f.suffix.lower() == ".pdf" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                return self.baixar(f.name, f.read_bytes(), tipo)
            if u.path == "/api/original":
                f = P.ENT / Path(q["arquivo"]).name
                tipo = "application/pdf" if f.suffix.lower() == ".pdf" else "image/png" if f.suffix.lower() == ".png" else "image/jpeg"
                return self.enviar(f.read_bytes(), tipo)
            if u.path == "/ms/conectar":
                if not ms.configurado(): return self.enviar("Conta Microsoft não configurada no servidor.", "text/plain; charset=utf-8", 404)
                return self.redirecionar(ms.url_login(s["login"]))
            if u.path == "/ms/retorno":
                if q.get("error"): return self.redirecionar("/?ms=erro")
                ms.concluir_login(s["login"], q.get("state", ""), q.get("code", ""), P.MS_TOKEN)
                return self.redirecionar("/?ms=ok")
            if u.path == "/api/usuarios" and s["admin"]:
                return self.enviar(CONTAS.lista())
            self.enviar({"erro": "não encontrado"}, status=404)
        except (FileNotFoundError, StopIteration):
            self.enviar({"erro": "não encontrado"}, status=404)
        except Exception as ex:
            traceback.print_exc(); self.enviar({"erro": str(ex).replace(str(DADOS), "…")}, status=500)  # sem caminhos internos

    # ---------- POST ----------
    def do_POST(self):
        u = urllib.parse.urlparse(self.path); q = dict(urllib.parse.parse_qsl(u.query))
        try:
            if u.path in ("/api/login", "/api/primeiro_acesso"):
                if self.headers.get("X-Gerador") != "1": return self.enviar({"erro": "pedido recusado"}, status=403)
                c = self.corpo_json()
                if u.path == "/api/primeiro_acesso":  # só funciona enquanto não existe nenhum usuário
                    if CONTAS.tem_usuarios(): return self.enviar({"erro": "O sistema já tem administrador."}, status=403)
                    erro = CONTAS.criar(c.get("login", ""), c.get("nome", ""), c.get("senha", ""), admin=True, trocar_senha=False)
                    if erro: return self.enviar({"erro": erro}, status=400)
                token, erro = CONTAS.entrar(c.get("login", ""), c.get("senha", ""), self.ip())
                if not token: return self.enviar({"erro": erro}, status=401)
                return self.enviar({"ok": True}, extra={"Set-Cookie": self.cookie_sessao(token, _contas.DURACAO_SESSAO)})
            s = self.autenticar(u.path, "POST")
            if s is None: return
            if u.path == "/api/logout":
                CONTAS.sair(self.token())
                return self.enviar({"ok": True}, extra={"Set-Cookie": self.cookie_sessao("", 0)})
            if u.path == "/api/senha":  # o próprio técnico troca a senha
                c = self.corpo_json()
                if not CONTAS.entrar(s["login"], c.get("atual", ""), self.ip())[0]:
                    return self.enviar({"erro": "A senha atual não confere."}, status=400)
                erro = CONTAS.definir_senha(s["login"], c.get("nova", ""))
                return self.enviar({"erro": erro} if erro else {"ok": True}, status=400 if erro else 200)
            if u.path.startswith("/api/usuarios"):  # administração (só administrador)
                if not s["admin"]: return self.enviar({"erro": "só o administrador"}, status=403)
                c = self.corpo_json()
                if u.path == "/api/usuarios/novo":
                    erro = CONTAS.criar(c.get("login", ""), c.get("nome", ""), c.get("senha", ""), admin=bool(c.get("admin")))
                elif u.path == "/api/usuarios/senha":
                    erro = CONTAS.definir_senha(c.get("login", ""), c.get("senha", ""), trocar_senha=True)
                elif u.path == "/api/usuarios/alterar":
                    if c.get("login") == s["login"] and c.get("ativo") is False: erro = "Você não pode desativar a si mesmo."
                    else: erro = CONTAS.alterar(c.get("login", ""), **{k: c[k] for k in ("nome", "admin", "ativo") if k in c})
                else: erro = "ação desconhecida"
                return self.enviar({"erro": erro} if erro else {"ok": True, "usuarios": CONTAS.lista()}, status=400 if erro else 200)
            if u.path == "/api/upload":
                n = int(self.headers.get("Content-Length") or 0)
                if n > MAX_UPLOAD: return self.enviar({"erro": f"arquivo maior que {MAX_UPLOAD // 1024 // 1024} MB"}, status=413)
                dados = self.rfile.read(n)
                nome = Path(q["nome"]).name; ext = Path(nome).suffix.lower()
                if ext not in EXT_OK + EXT_HEIC:
                    return self.enviar({"erro": f"{nome}: tipo não aceito (use .jpg, .png, .pdf ou .heic)"}, status=400)
                if len(dados) == 0 or len(dados) != n:
                    return self.enviar({"erro": f"{nome}: o arquivo chegou vazio ou incompleto, envie de novo"}, status=400)
                if ext in EXT_HEIC:  # foto de iPhone: converte para JPG
                    import pillow_heif; pillow_heif.register_heif_opener()
                    b = io.BytesIO(); Image.open(io.BytesIO(dados)).convert("RGB").save(b, "JPEG", quality=92)
                    dados, nome = b.getvalue(), Path(nome).stem + ".jpg"
                for f in P.ENT.iterdir():  # mesmo arquivo já enviado antes: não duplica
                    if f.is_file() and f.stat().st_size == len(dados) and f.read_bytes() == dados:
                        return self.enviar({"ok": True, "arquivo": f.name, "duplicado": True})
                destino, k = P.ENT / nome, 1
                while destino.exists():
                    destino = P.ENT / f"{Path(nome).stem}_{k}{Path(nome).suffix}"; k += 1
                parcial = destino.with_name(destino.name + ".parcial")
                parcial.write_bytes(dados); parcial.replace(destino)
                return self.enviar({"ok": True, "arquivo": destino.name})
            if u.path in ("/api/estado", "/api/gerar", "/api/nova_viagem"):
                e = self.corpo_json()
                if aba_desatualizada(e): return self.enviar({"erro": "DESATUALIZADA"}, status=409)
                self._e = e
            if u.path == "/api/estado":
                salvar_estado(self._e); return self.enviar({"ok": True})
            if u.path == "/api/ler":  # leitura automática (OCR / texto do PDF)
                import leitor
                return self.enviar(leitor.ler(P.ENT / Path(q["arquivo"]).name))
            if u.path == "/api/gerar":
                salvar_estado(self._e); r = gerar(self._e)
                if r.get("ok"): (P.BASE / "ultimo.json").write_text(json.dumps(r, ensure_ascii=False), encoding="utf-8")
                return self.enviar(r)
            if u.path == "/api/config":
                return self.enviar(salvar_config(self.corpo_json()))
            if u.path == "/api/email":
                return self.enviar(abrir_email(self.corpo_json()))
            if u.path == "/api/ms/desconectar":
                ms.desconectar(P.MS_TOKEN); return self.enviar({"ok": True})
            if u.path == "/api/politica":
                e = self.corpo_json(); aplicar_periodo(e); x, reemb = politica(e)
                return self.enviar({"excessos": x, "reembolso": [reemb.get(id(l)) for l in e.get("linhas", [])]})
            if u.path == "/api/assinatura":
                n = int(self.headers.get("Content-Length") or 0)
                if n > MAX_UPLOAD: return self.enviar({"erro": "imagem grande demais"}, status=413)
                dados = self.rfile.read(n)
                if q.get("apagar"):
                    P.ASSINATURA.unlink(missing_ok=True); return self.enviar({"ok": True})
                try: salvar_assinatura(dados)
                except Exception as ex: return self.enviar({"erro": f"Não consegui usar essa imagem: {ex}"}, status=400)
                return self.enviar({"ok": True})
            if u.path == "/api/historico/enviado":
                c = self.corpo_json(); return self.enviar(marcar_enviado(c["id"], c["enviado"]))
            if u.path == "/api/nova_viagem":
                (P.BASE / "ultimo.json").unlink(missing_ok=True)
                return self.enviar(arquivar(self._e))
            self.enviar({"erro": "não encontrado"}, status=404)
        except (FileNotFoundError, StopIteration):
            self.enviar({"erro": "não encontrado"}, status=404)
        except Exception as ex:
            traceback.print_exc(); self.enviar({"erro": str(ex).replace(str(DADOS), "…")}, status=500)  # sem caminhos internos


class Servidor(socketserver.ThreadingMixIn, http.server.HTTPServer):
    daemon_threads = True


def limpar_caches_antigos():
    for pasta in (DADOS / "cache").glob("*"):
        if pasta.is_dir(): limpar_cache(7 * 24 * 60, pasta)


if __name__ == "__main__":
    host = os.environ.get("HOST", "0.0.0.0")
    srv = Servidor((host, PORTA), H)
    print(f"{dt.datetime.now():%d/%m %H:%M} Gerador online em http://{host}:{PORTA}  (dados em {DADOS})", flush=True)
    if not CONTAS.tem_usuarios(): print("Nenhum usuário ainda: abra /login para criar o administrador.", flush=True)
    print("Conta Microsoft (rascunho no Outlook): " + ("configurada" if ms.configurado() else "não configurada — o botão abre um e-mail sem anexo"), flush=True)
    threading.Thread(target=limpar_caches_antigos, daemon=True).start()
    srv.serve_forever()
