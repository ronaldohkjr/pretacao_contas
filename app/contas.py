# Contas dos técnicos: usuários, senhas (PBKDF2), sessões e bloqueio de tentativas.
# Tudo fica em DADOS/usuarios.json e DADOS/sessoes.json (nunca a senha em texto: só o hash com sal).
import hashlib, hmac, json, re, secrets, threading, time
from pathlib import Path

ITERACOES = 240_000
DURACAO_SESSAO = 30 * 24 * 3600      # 30 dias
TENTATIVAS, JANELA = 8, 15 * 60      # 8 senhas erradas em 15 min bloqueiam o login por 15 min
_lock = threading.Lock()
_falhas = {}                          # (ip, login) -> [instantes das falhas]


class Contas:
    def __init__(self, dados: Path):
        self.arq_usuarios = dados / "usuarios.json"
        self.arq_sessoes = dados / "sessoes.json"
        dados.mkdir(parents=True, exist_ok=True)

    # ---------- arquivos ----------
    def _ler(self, arq):
        return json.loads(arq.read_text(encoding="utf-8")) if arq.exists() else {}

    def _gravar(self, arq, d):
        tmp = arq.with_suffix(".tmp"); tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8"); tmp.replace(arq)

    def usuarios(self): return self._ler(self.arq_usuarios)

    def tem_usuarios(self): return bool(self.usuarios())

    # ---------- senhas ----------
    @staticmethod
    def _hash(senha, sal):
        return hashlib.pbkdf2_hmac("sha256", senha.encode(), bytes.fromhex(sal), ITERACOES).hex()

    @staticmethod
    def validar_senha(senha):
        if len(senha) < 8: return "A senha precisa ter pelo menos 8 caracteres."
        if not (re.search(r"[A-Za-z]", senha) and re.search(r"\d", senha)): return "Use letras e números na senha."
        return ""

    @staticmethod
    def validar_login(login):
        return "" if re.fullmatch(r"[a-z0-9._-]{3,30}", login) else \
            "Usuário: 3 a 30 caracteres, só letras minúsculas, números, ponto, hífen ou _ (ex.: joao.silva)."

    # ---------- cadastro (administrador) ----------
    def criar(self, login, nome, senha, admin=False, trocar_senha=True):
        login = login.strip().lower()
        erro = self.validar_login(login) or self.validar_senha(senha)
        if erro: return erro
        with _lock:
            us = self.usuarios()
            if login in us: return "Já existe um usuário com esse login."
            sal = secrets.token_hex(16)
            us[login] = {"nome": nome.strip() or login, "sal": sal, "hash": self._hash(senha, sal), "admin": bool(admin),
                         "ativo": True, "trocar_senha": bool(trocar_senha), "criado": time.strftime("%Y-%m-%dT%H:%M:%S")}
            self._gravar(self.arq_usuarios, us)
        return ""

    def definir_senha(self, login, senha, trocar_senha=False):
        erro = self.validar_senha(senha)
        if erro: return erro
        with _lock:
            us = self.usuarios()
            if login not in us: return "Usuário não encontrado."
            sal = secrets.token_hex(16)
            us[login].update(sal=sal, hash=self._hash(senha, sal), trocar_senha=bool(trocar_senha))
            self._gravar(self.arq_usuarios, us)
        if trocar_senha: self.encerrar_sessoes(login)  # senha redefinida pelo administrador: derruba as sessões abertas
        return ""

    def alterar(self, login, **campos):
        with _lock:
            us = self.usuarios()
            if login not in us: return "Usuário não encontrado."
            for k in ("nome", "admin", "ativo"):
                if k in campos: us[login][k] = campos[k]
            if not any(u["admin"] and u["ativo"] for u in us.values()): return "Precisa ficar pelo menos um administrador ativo."
            self._gravar(self.arq_usuarios, us)
        if campos.get("ativo") is False: self.encerrar_sessoes(login)
        return ""

    def lista(self):
        return [{"login": k, "nome": u["nome"], "admin": u["admin"], "ativo": u["ativo"], "criado": u.get("criado", ""),
                 "ultimo_acesso": u.get("ultimo_acesso", "")} for k, u in sorted(self.usuarios().items())]

    # ---------- login / sessões ----------
    def bloqueado(self, ip, login):
        """8 erros do mesmo IP, ou 20 de qualquer IP (protege mesmo se o IP for forjado), em 15 min."""
        agora = time.time()
        for chave, limite in (((ip, login), TENTATIVAS), (("*", login), TENTATIVAS * 2 + 4)):
            f = [t for t in _falhas.get(chave, []) if agora - t < JANELA]
            _falhas[chave] = f
            if len(f) >= limite: return True
        return False

    def entrar(self, login, senha, ip):
        """Devolve (token, "") se a senha confere, ou (None, motivo)."""
        login = (login or "").strip().lower()
        if self.bloqueado(ip, login): return None, "Muitas tentativas erradas. Espere 15 minutos e tente de novo."
        u = self.usuarios().get(login)
        ok = bool(u) and hmac.compare_digest(self._hash(senha or "", u["sal"]), u["hash"])
        if not u: self._hash(senha or "", "00" * 16)  # mesmo tempo de resposta com usuário inexistente
        if not ok:
            for chave in ((ip, login), ("*", login)): _falhas.setdefault(chave, []).append(time.time())
            return None, "Usuário ou senha incorretos."
        if not u["ativo"]: return None, "Este usuário está desativado. Fale com o administrador."
        _falhas.pop((ip, login), None)
        token = secrets.token_urlsafe(32)
        with _lock:
            ss = self._ler(self.arq_sessoes); agora = time.time()
            ss = {k: v for k, v in ss.items() if v["expira"] > agora}  # limpa as vencidas
            ss[self._chave(token)] = {"login": login, "expira": agora + DURACAO_SESSAO}
            self._gravar(self.arq_sessoes, ss)
            us = self.usuarios(); us[login]["ultimo_acesso"] = time.strftime("%Y-%m-%dT%H:%M:%S"); self._gravar(self.arq_usuarios, us)
        return token, ""

    @staticmethod
    def _chave(token):  # guarda só o hash do token: quem ler o arquivo não consegue usar a sessão
        return hashlib.sha256(token.encode()).hexdigest()

    def sessao(self, token):
        """Devolve {"login", "nome", "admin", "trocar_senha"} da sessão válida, ou None."""
        if not token: return None
        s = self._ler(self.arq_sessoes).get(self._chave(token))
        if not s or s["expira"] < time.time(): return None
        u = self.usuarios().get(s["login"])
        if not u or not u["ativo"]: return None
        return {"login": s["login"], "nome": u["nome"], "admin": u["admin"], "trocar_senha": u.get("trocar_senha", False)}

    def sair(self, token):
        with _lock:
            ss = self._ler(self.arq_sessoes); ss.pop(self._chave(token or ""), None); self._gravar(self.arq_sessoes, ss)

    def encerrar_sessoes(self, login):
        with _lock:
            ss = self._ler(self.arq_sessoes)
            self._gravar(self.arq_sessoes, {k: v for k, v in ss.items() if v["login"] != login})
