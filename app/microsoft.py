# Rascunho no Outlook pela conta Microsoft do técnico (Microsoft Graph).
# Só funciona se a TI registrar o aplicativo no Microsoft Entra e o servidor tiver as variáveis:
#   MS_CLIENT_ID, MS_CLIENT_SECRET, MS_TENANT (id do tenant ou "organizations") e URL_PUBLICA (ex.: https://gerador.exemplo.com)
# Permissões delegadas: offline_access, User.Read, Mail.ReadWrite (cria o rascunho; o sistema nunca envia e-mail).
import base64, json, os, secrets, time, urllib.error, urllib.parse, urllib.request
from pathlib import Path

CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "")
TENANT = os.environ.get("MS_TENANT", "organizations")
URL_PUBLICA = os.environ.get("URL_PUBLICA", "").rstrip("/")
ESCOPOS = "offline_access User.Read Mail.ReadWrite"
GRAPH = "https://graph.microsoft.com/v1.0"
_estados = {}  # state do login -> (login do técnico, expira)


def configurado(): return bool(CLIENT_ID and CLIENT_SECRET and URL_PUBLICA)


def _retorno(): return URL_PUBLICA + "/ms/retorno"


def _pedir(url, dados=None, metodo=None, token=None, cabecalhos=None, json_corpo=None):
    h = dict(cabecalhos or {})
    corpo = None
    if json_corpo is not None: corpo = json.dumps(json_corpo).encode(); h["Content-Type"] = "application/json"
    elif dados is not None: corpo = urllib.parse.urlencode(dados).encode(); h["Content-Type"] = "application/x-www-form-urlencoded"
    if token: h["Authorization"] = "Bearer " + token
    req = urllib.request.Request(url, data=corpo, headers=h, method=metodo or ("POST" if corpo is not None else "GET"))
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            txt = r.read(); return json.loads(txt) if txt else {}
    except urllib.error.HTTPError as ex:
        detalhe = ex.read().decode("utf-8", "ignore")[:400]
        raise RuntimeError(f"Microsoft respondeu {ex.code}: {detalhe}") from None


# ---------- login (código de autorização) ----------
def url_login(login):
    estado = secrets.token_urlsafe(24)
    agora = time.time()
    for k in [k for k, v in _estados.items() if v[1] < agora]: _estados.pop(k)
    _estados[estado] = (login, agora + 600)
    q = {"client_id": CLIENT_ID, "response_type": "code", "redirect_uri": _retorno(), "response_mode": "query",
         "scope": ESCOPOS, "state": estado, "prompt": "select_account"}
    return f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize?" + urllib.parse.urlencode(q)


def concluir_login(login, estado, codigo, arq_token: Path):
    dono = _estados.pop(estado, None)
    if not dono or dono[0] != login or dono[1] < time.time(): raise RuntimeError("Login da Microsoft expirou. Tente de novo.")
    t = _pedir(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token", dados={
        "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET, "grant_type": "authorization_code",
        "code": codigo, "redirect_uri": _retorno(), "scope": ESCOPOS})
    eu = _pedir(GRAPH + "/me?$select=mail,userPrincipalName", token=t["access_token"])
    _salvar(arq_token, t, eu.get("mail") or eu.get("userPrincipalName", ""))


def _salvar(arq, t, email=None):
    antigo = json.loads(arq.read_text(encoding="utf-8")) if arq.exists() else {}
    arq.write_text(json.dumps({"refresh_token": t.get("refresh_token") or antigo.get("refresh_token"),
                               "access_token": t["access_token"], "expira": time.time() + int(t.get("expires_in", 3600)) - 120,
                               "email": email or antigo.get("email", "")}), encoding="utf-8")


def conectado(arq_token: Path):
    if not arq_token.exists(): return ""
    return json.loads(arq_token.read_text(encoding="utf-8")).get("email") or "conectado"


def desconectar(arq_token: Path): arq_token.unlink(missing_ok=True)


def _token(arq_token: Path):
    d = json.loads(arq_token.read_text(encoding="utf-8"))
    if d["expira"] > time.time(): return d["access_token"]
    t = _pedir(f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token", dados={
        "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET, "grant_type": "refresh_token",
        "refresh_token": d["refresh_token"], "scope": ESCOPOS})
    _salvar(arq_token, t); return t["access_token"]


# ---------- rascunho ----------
def criar_rascunho(arq_token: Path, para, cc, assunto, corpo_html, anexos):
    """Cria o rascunho na caixa do técnico (não envia) e devolve o link para abrir no Outlook na web."""
    tk = _token(arq_token)
    def end(s): return [{"emailAddress": {"address": x.strip()}} for x in (s or "").split(";") if x.strip()]
    msg = _pedir(GRAPH + "/me/messages", token=tk, json_corpo={
        "subject": assunto, "body": {"contentType": "HTML", "content": corpo_html},
        "toRecipients": end(para), "ccRecipients": end(cc)})
    for f in anexos:
        dados = Path(f).read_bytes()
        if len(dados) < 3 * 1024 * 1024:   # anexo pequeno: vai direto
            _pedir(f"{GRAPH}/me/messages/{msg['id']}/attachments", token=tk, json_corpo={
                "@odata.type": "#microsoft.graph.fileAttachment", "name": Path(f).name,
                "contentType": "application/pdf" if Path(f).suffix.lower() == ".pdf" else "application/octet-stream",
                "contentBytes": base64.b64encode(dados).decode()})
        else:                               # anexo grande (PDF das notas): sessão de envio em partes
            s = _pedir(f"{GRAPH}/me/messages/{msg['id']}/attachments/createUploadSession", token=tk, json_corpo={
                "AttachmentItem": {"attachmentType": "file", "name": Path(f).name, "size": len(dados)}})
            parte = 320 * 1024 * 12        # ~3,75 MB (múltiplo de 320 KB)
            for ini in range(0, len(dados), parte):
                pedaco = dados[ini:ini + parte]
                req = urllib.request.Request(s["uploadUrl"], data=pedaco, method="PUT", headers={
                    "Content-Length": str(len(pedaco)), "Content-Type": "application/octet-stream",
                    "Content-Range": f"bytes {ini}-{ini + len(pedaco) - 1}/{len(dados)}"})
                urllib.request.urlopen(req, timeout=120).read()
    return msg.get("webLink", "")
