# =====================================================
# BOT DICOR FINAL RESOLVIDO - MESAS + PROCURADOS + CATALOGO HTML
# Feito para GTA RP / personagens ficticios
# =====================================================

import os
import re
import json
import html
import hmac
import secrets
import string
import time
import asyncio
import datetime
import unicodedata
from pathlib import Path
from typing import Optional, Dict, Any, List

import discord
from discord import app_commands
from discord.ext import commands, tasks
from discord.ui import View, Button, Modal, TextInput
from dotenv import load_dotenv
from aiohttp import web

# =====================================================
# CONFIGURACAO
# =====================================================

load_dotenv()

DISCORD_TOKEN = os.getenv("DISCORD_TOKEN", "").strip()
GUILD_ID = int(os.getenv("GUILD_ID", "0") or 0)

# Procurados
PROCURADOS_CHANNEL_ID = int(os.getenv("PROCURADOS_CHANNEL_ID", "0") or 0)
HISTORICO_PROCURADOS_ID = int(os.getenv("HISTORICO_PROCURADOS_ID", "1490200536207855857") or 1490200536207855857)
LOGS_CHANNEL_ID = int(os.getenv("LOGS_CHANNEL_ID", "0") or 0)
PROCURADOS_TEMP_CATEGORY_ID = int(os.getenv("PROCURADOS_TEMP_CATEGORY_ID", "0") or 0)

# Boletins
BOLETINS_CHANNEL_ID = int(os.getenv("BOLETINS_CHANNEL_ID", "1490200514837745754") or 1490200514837745754)
BOLETIM_TEMP_CATEGORY_ID = int(os.getenv("BOLETIM_TEMP_CATEGORY_ID", "0") or 0)

# Mesas
CATEGORIA_MESAS_ABERTAS_ID = int(os.getenv("CATEGORIA_MESAS_ABERTAS_ID", "1490200456855552192") or 1490200456855552192)
CATEGORIA_MESAS_FECHADAS_ID = int(os.getenv("CATEGORIA_MESAS_FECHADAS_ID", "1515165416815722586") or 1515165416815722586)
BACKUP_CHANNEL_ID = int(os.getenv("BACKUP_CHANNEL_ID", "1515165673276440677") or 1515165673276440677)

# Catalogo
CATALOG_PUBLIC_URL = os.getenv("CATALOG_PUBLIC_URL", "http://127.0.0.1:8000/").strip()
CATALOG_ADMIN_PASSWORD = os.getenv("CATALOG_ADMIN_PASSWORD", "").strip()
PORT = int(os.getenv("PORT", "8000") or 8000)


def _ids_env(nome: str) -> List[int]:
    ids: List[int] = []
    for parte in os.getenv(nome, "").split(","):
        parte = parte.strip()
        if parte.isdigit():
            ids.append(int(parte))
    return ids

CARGOS_ADMIN_IDS = _ids_env("CARGOS_ADMIN_IDS")
CARGOS_EQUIPE_IDS = _ids_env("CARGOS_EQUIPE_IDS") or CARGOS_ADMIN_IDS

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
PUBLIC_DIR = BASE_DIR / "public"
UPLOADS_DIR = PUBLIC_DIR / "uploads"
BACKUP_DIR = BASE_DIR / "backups"

CATALOGO_JSON = DATA_DIR / "procurados.json"
CATALOGO_HTML = PUBLIC_DIR / "index.html"
MESAS_JSON = DATA_DIR / "mesas.json"
BOLETINS_JSON = DATA_DIR / "boletins.json"
BOLETINS_PENDENTES_JSON = DATA_DIR / "boletins_pendentes.json"
USUARIOS_RG_JSON = DATA_DIR / "usuarios_rg.json"
BOLETINS_CONTADOR_JSON = DATA_DIR / "boletins_contador.json"
CATALOG_ADMIN_JSON = DATA_DIR / "catalog_admin.json"
ORGANIZACOES_JSON = DATA_DIR / "organizacoes.json"
HISTORICO_ORGANIZACOES_JSON = DATA_DIR / "historico_organizacoes.json"

for pasta in [DATA_DIR, PUBLIC_DIR, UPLOADS_DIR, BACKUP_DIR]:
    pasta.mkdir(exist_ok=True)

# =====================================================
# BOT
# =====================================================

intents = discord.Intents.default()
intents.guilds = True
intents.members = True
intents.messages = True
intents.message_content = True

bot = commands.Bot(command_prefix="!", intents=intents)

cadastros_pendentes: Dict[int, Dict[str, Any]] = {}
boletins_pendentes: Dict[int, Dict[str, Any]] = {}
catalogo_tentativas_senha: Dict[str, List[float]] = {}

# =====================================================
# FUNCOES GERAIS
# =====================================================

def agora_br() -> str:
    tz = datetime.timezone(datetime.timedelta(hours=-3))
    return datetime.datetime.now(tz).strftime("%d/%m/%Y %H:%M")


def data_caso() -> str:
    tz = datetime.timezone(datetime.timedelta(hours=-3))
    return datetime.datetime.now(tz).strftime("%Y%m%d-%H%M%S")


def slugify(texto: str) -> str:
    texto = str(texto or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[^a-z0-9]+", "-", texto)
    texto = texto.strip("-")
    return texto[:70] or "item"


def limpar_rg(rg: str) -> str:
    return re.sub(r"\s+", "", str(rg or "").lower())


def escape(texto: Any) -> str:
    return html.escape(str(texto or ""))


def cortar_discord(texto: str, limite: int = 1900) -> str:
    texto = str(texto or "")
    if len(texto) <= limite:
        return texto
    return texto[:limite - 40].rstrip() + "\n\n... texto cortado pelo limite do Discord."


def carregar_json(caminho: Path, padrao):
    if not caminho.exists():
        return padrao
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return padrao


def salvar_json(caminho: Path, dados) -> None:
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)


def garantir_senha_catalogo() -> str:
    """Carrega a senha do Railway ou gera uma senha local persistente."""
    global CATALOG_ADMIN_PASSWORD

    if CATALOG_ADMIN_PASSWORD:
        return CATALOG_ADMIN_PASSWORD

    configuracao = carregar_json(CATALOG_ADMIN_JSON, {})
    if isinstance(configuracao, dict):
        senha_salva = str(configuracao.get("senha", "") or "").strip()
        if senha_salva:
            CATALOG_ADMIN_PASSWORD = senha_salva
            return CATALOG_ADMIN_PASSWORD

    alfabeto = string.ascii_letters + string.digits
    CATALOG_ADMIN_PASSWORD = "DICOR-" + "".join(
        secrets.choice(alfabeto) for _ in range(14)
    )
    salvar_json(CATALOG_ADMIN_JSON, {"senha": CATALOG_ADMIN_PASSWORD})
    print("ATENÇÃO: CATALOG_ADMIN_PASSWORD não foi configurada no Railway.")
    print(f"SENHA TEMPORÁRIA DO CATÁLOGO: {CATALOG_ADMIN_PASSWORD}")
    print(
        "Cadastre essa senha na variável CATALOG_ADMIN_PASSWORD "
        "para ela não mudar em novos deploys."
    )
    return CATALOG_ADMIN_PASSWORD


def senha_catalogo_valida(senha: str) -> bool:
    senha_ativa = garantir_senha_catalogo()
    return bool(senha_ativa) and hmac.compare_digest(
        str(senha or ""),
        senha_ativa,
    )


def catalogo_limite_tentativas(ip: str) -> bool:
    """Permite até 8 tentativas erradas por IP a cada 10 minutos."""
    agora = time.monotonic()
    janela = 600.0
    tentativas = [
        instante
        for instante in catalogo_tentativas_senha.get(ip, [])
        if agora - instante < janela
    ]
    catalogo_tentativas_senha[ip] = tentativas
    return len(tentativas) < 8


def registrar_falha_senha_catalogo(ip: str) -> None:
    catalogo_tentativas_senha.setdefault(ip, []).append(time.monotonic())


def carregar_procurados() -> List[Dict[str, Any]]:
    dados = carregar_json(CATALOGO_JSON, [])
    return dados if isinstance(dados, list) else []


def salvar_procurados(lista: List[Dict[str, Any]]) -> None:
    salvar_json(CATALOGO_JSON, lista)


def carregar_mesas() -> List[Dict[str, Any]]:
    dados = carregar_json(MESAS_JSON, [])
    return dados if isinstance(dados, list) else []


def salvar_mesas(lista: List[Dict[str, Any]]) -> None:
    salvar_json(MESAS_JSON, lista)


def procurar_por_rg(rg: str) -> Optional[Dict[str, Any]]:
    alvo = limpar_rg(rg)
    for p in carregar_procurados():
        if limpar_rg(p.get("rg", "")) == alvo:
            return p
    return None


def remover_duplicados(lista: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    vistos = set()
    saida = []
    for item in lista:
        rg = limpar_rg(item.get("rg", ""))
        chave = rg or item.get("id") or item.get("caso")
        if chave in vistos:
            continue
        vistos.add(chave)
        saida.append(item)
    return saida


async def enviar_log(texto: str) -> None:
    if not LOGS_CHANNEL_ID:
        return
    canal = bot.get_channel(LOGS_CHANNEL_ID)
    if canal:
        try:
            await canal.send(texto[:1900])
        except Exception:
            pass


def usuario_tem_admin(member: discord.Member) -> bool:
    if not CARGOS_ADMIN_IDS:
        return True
    cargos = {role.id for role in member.roles}
    return any(cargo in cargos for cargo in CARGOS_ADMIN_IDS)


def cargos_equipe_permissoes(guild: discord.Guild) -> Dict[Any, discord.PermissionOverwrite]:
    overwrites: Dict[Any, discord.PermissionOverwrite] = {
        guild.default_role: discord.PermissionOverwrite(view_channel=False)
    }
    for cargo_id in set(CARGOS_EQUIPE_IDS + CARGOS_ADMIN_IDS):
        cargo = guild.get_role(cargo_id)
        if cargo:
            overwrites[cargo] = discord.PermissionOverwrite(
                view_channel=True,
                send_messages=True,
                attach_files=True,
                read_message_history=True,
                create_public_threads=True,
                send_messages_in_threads=True,
                manage_threads=True,
            )
    if guild.me:
        overwrites[guild.me] = discord.PermissionOverwrite(
            view_channel=True,
            send_messages=True,
            manage_channels=True,
            attach_files=True,
            read_message_history=True,
            create_public_threads=True,
            send_messages_in_threads=True,
            manage_threads=True,
        )
    return overwrites

# =====================================================
# CATALOGO HTML COM FOTO EM TELA CHEIA
# =====================================================

def valor_crimes_registro(registro: Dict[str, Any]) -> str:
    """Aceita o nome atual e também chaves usadas por versões antigas."""
    for chave in ("crimes", "crimes_cometidos", "crimes_imputados", "crime"):
        valor = str(registro.get(chave, "") or "").strip()
        if valor and valor.lower() not in {
            "não informado",
            "nao informado",
            "não identificado na mensagem antiga",
            "nao identificado na mensagem antiga",
        }:
            return valor
    return "Não informado"


def registro_tem_valor_util(valor: Any) -> bool:
    texto = str(valor or "").strip().lower()
    return bool(texto) and texto not in {
        "não informado",
        "nao informado",
        "não identificado na mensagem antiga",
        "nao identificado na mensagem antiga",
        "importado automaticamente do histórico do canal.",
    }


def gerar_catalogo_html() -> None:
    procurados = carregar_procurados()

    # Corrige registros criados por versões antigas.
    alterou = False
    for registro in procurados:
        crimes = valor_crimes_registro(registro)
        if registro.get("crimes") != crimes:
            registro["crimes"] = crimes
            alterou = True
    if alterou:
        salvar_procurados(procurados)

    visiveis = [
        p for p in procurados
        if str(p.get("status", "A PROCURAR") or "A PROCURAR").upper() != "APAGADO"
    ]
    ativos = [
        p for p in visiveis
        if str(p.get("status", "A PROCURAR") or "A PROCURAR").upper() != "RETIRADO"
    ]
    retirados = [
        p for p in visiveis
        if str(p.get("status", "") or "").upper() == "RETIRADO"
    ]

    def img_html(src: str, alt: str, vazio: str) -> str:
        if not src:
            return f'<div class="placeholder">{vazio}</div>'
        src_e = escape(src)
        return (
            f'<img class="zoom-img" src="{src_e}" alt="{escape(alt)}" '
            f'onclick="abrirFoto(this.src, this.alt)">'
        )

    def card_html(registro: Dict[str, Any]) -> str:
        status = str(registro.get("status", "A PROCURAR") or "A PROCURAR").upper()
        classe_status = "retirado" if status == "RETIRADO" else "ativo"
        foto_ind = registro.get("foto_individuo", "") or ""
        foto_rg = registro.get("foto_rg", "") or ""
        link_msg = registro.get("mensagem_url", "") or ""
        crimes = valor_crimes_registro(registro)
        registro_id_js = json.dumps(
            str(
                registro.get("id")
                or registro.get("caso")
                or registro.get("rg")
                or ""
            ),
            ensure_ascii=False,
        )
        nome_js = json.dumps(
            str(registro.get("nome") or "Sem nome"),
            ensure_ascii=False,
        )

        return f"""
        <article class="card {classe_status}">
            <div class="card-head">
                <div><span>Nº DO CASO</span><b>{escape(registro.get('caso'))}</b></div>
                <div><span>DATA</span><b>{escape(registro.get('data'))}</b></div>
                <div class="status">{escape(status)}</div>
            </div>

            <div class="card-grid">
                <div class="photos">
                    <div class="photo big">
                        <div class="photo-title">Foto do Indivíduo</div>
                        {img_html(foto_ind, 'Foto do indivíduo', 'Sem foto')}
                    </div>
                    <div class="photo small">
                        <div class="photo-title">Foto do RG</div>
                        {img_html(foto_rg, 'Foto do RG', 'Sem RG')}
                    </div>
                </div>

                <div class="info">
                    <div class="linha"><b>Nome</b><p>{escape(registro.get('nome'))}</p></div>
                    <div class="linha"><b>RG</b><p>{escape(registro.get('rg'))}</p></div>
                    <div class="box crimes"><b>Crimes Cometidos</b><p>{escape(crimes)}</p></div>
                    <div class="box destaque"><b>Último Avistamento</b><p>{escape(registro.get('ultimo_avistamento'))}</p></div>
                    <div class="box"><b>Informações</b><p>{escape(registro.get('informacoes'))}</p></div>
                    {f'<a class="msg" href="{escape(link_msg)}" target="_blank" rel="noopener noreferrer">Abrir mensagem no Discord</a>' if link_msg else ''}
                    {f'<div class="motivo"><b>Motivo da retirada:</b> {escape(registro.get("motivo_retirada"))}</div>' if registro.get('motivo_retirada') else ''}
                    <button class="apagar-registro" type="button" onclick='abrirExclusao({registro_id_js}, {nome_js})'>
                        🗑️ Apagar permanentemente
                    </button>
                </div>
            </div>
        </article>
        """

    cards_ativos = "".join(card_html(p) for p in ativos)
    cards_retirados = "".join(card_html(p) for p in retirados)

    if not cards_ativos:
        cards_ativos = """
        <div class="vazio">
            <h2>Nenhum procurado ativo.</h2>
            <p>Os novos procurados aparecerão nesta aba automaticamente.</p>
        </div>
        """

    if not cards_retirados:
        cards_retirados = """
        <div class="vazio">
            <h2>Nenhum procurado retirado.</h2>
            <p>Os registros retirados aparecerão nesta aba.</p>
        </div>
        """

    html_final = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catálogo de Procurados - DICOR</title>
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            background:
                radial-gradient(circle at top, rgba(230,169,55,.12), transparent 35%),
                linear-gradient(135deg, #07111f, #02060d 70%);
            color: #f7f7f7;
        }}
        header {{
            border-bottom: 2px solid #d99a2b;
            padding: 28px 18px 22px;
            text-align: center;
            background: rgba(0,0,0,.38);
            position: sticky;
            top: 0;
            z-index: 10;
            backdrop-filter: blur(7px);
        }}
        header h1 {{ margin:0; color:#f0b64d; font-size:clamp(30px,5vw,62px); letter-spacing:2px; text-transform:uppercase; }}
        header p {{ margin:8px 0 0; color:#ddd; }}
        .stats {{ display:flex; gap:14px; justify-content:center; flex-wrap:wrap; margin-top:18px; }}
        .stat {{ border:1px solid rgba(240,182,77,.55); background:rgba(8,19,35,.9); padding:10px 18px; border-radius:12px; min-width:150px; }}
        .stat span {{ display:block; font-size:12px; color:#bfc6d1; text-transform:uppercase; }}
        .stat b {{ font-size:24px; color:#fff; }}
        .abas {{ display:flex; justify-content:center; gap:10px; flex-wrap:wrap; margin-top:18px; }}
        .aba-btn {{
            border:1px solid rgba(240,182,77,.65);
            background:#081323;
            color:#f7f7f7;
            padding:11px 18px;
            border-radius:12px;
            font-weight:800;
            cursor:pointer;
            transition:.15s ease;
        }}
        .aba-btn:hover {{ transform:translateY(-1px); border-color:#f0b64d; }}
        .aba-btn.ativa {{ background:#f0b64d; color:#07111f; }}
        .aba-conteudo {{ display:none; }}
        .aba-conteudo.ativa {{ display:block; }}
        .grade {{ width:min(1380px,96%); margin:24px auto 50px; display:grid; grid-template-columns:repeat(auto-fit,minmax(520px,1fr)); gap:22px; }}
        .card {{ border:1px solid rgba(240,182,77,.55); background:linear-gradient(180deg,rgba(8,19,35,.97),rgba(3,8,16,.97)); box-shadow:0 0 0 1px rgba(255,255,255,.05),0 18px 50px rgba(0,0,0,.32); border-radius:18px; overflow:hidden; }}
        .card.retirado {{ opacity:.82; filter:grayscale(.25); }}
        .card-head {{ display:grid; grid-template-columns:1fr 1fr auto; gap:12px; padding:14px; border-bottom:1px solid rgba(240,182,77,.32); align-items:center; }}
        .card-head span {{ display:block; font-size:11px; color:#f0b64d; text-transform:uppercase; font-weight:bold; }}
        .card-head b {{ display:block; margin-top:4px; }}
        .status {{ color:#101820; background:#f0b64d; font-weight:900; padding:8px 12px; border-radius:10px; white-space:nowrap; }}
        .retirado .status {{ background:#c94a4a; color:#fff; }}
        .card-grid {{ display:grid; grid-template-columns:210px 1fr; gap:16px; padding:16px; }}
        .photos {{ display:grid; gap:14px; }}
        .photo {{ border:1px solid rgba(240,182,77,.55); background:#081323; border-radius:14px; padding:10px; }}
        .photo-title {{ text-align:center; color:#f0b64d; font-weight:800; margin-bottom:8px; text-transform:uppercase; font-size:13px; }}
        .photo img {{ width:100%; height:240px; object-fit:cover; border-radius:10px; background:#ddd; display:block; }}
        .photo.small img {{ height:125px; object-fit:cover; }}
        .zoom-img {{ cursor:zoom-in; transition:transform .15s ease,filter .15s ease; }}
        .zoom-img:hover {{ transform:scale(1.02); filter:brightness(1.08); }}
        .placeholder {{ height:240px; border:1px dashed #98a2b3; display:grid; place-items:center; color:#98a2b3; border-radius:10px; }}
        .photo.small .placeholder {{ height:125px; }}
        .info {{ display:grid; gap:10px; }}
        .linha,.box {{ background:#f6f7f9; color:#07111f; border-radius:12px; padding:10px 12px; border-left:5px solid #f0b64d; }}
        .linha b,.box b {{ display:block; color:#07111f; margin-bottom:3px; }}
        .linha p,.box p {{ margin:0; white-space:pre-wrap; line-height:1.4; overflow-wrap:anywhere; }}
        .box.crimes {{ min-height:90px; }}
        .destaque {{ background:#fff4dc; }}
        .msg {{ display:inline-block; color:#f0b64d; text-decoration:none; font-weight:bold; margin-top:4px; }}
        .motivo {{ background:rgba(201,74,74,.18); color:#ffd7d7; border:1px solid rgba(201,74,74,.4); padding:10px; border-radius:10px; }}
        .apagar-registro {{
            border:1px solid rgba(222,75,75,.72);
            background:rgba(138,29,29,.25);
            color:#ffdede;
            padding:10px 12px;
            border-radius:10px;
            font-weight:800;
            cursor:pointer;
            transition:.15s ease;
            justify-self:start;
        }}
        .apagar-registro:hover {{
            background:#a92f2f;
            color:#fff;
            transform:translateY(-1px);
        }}
        .vazio {{ grid-column:1/-1; text-align:center; padding:60px 20px; border:1px dashed rgba(240,182,77,.6); border-radius:18px; background:rgba(8,19,35,.7); }}
        #modalExclusao {{
            display:none;
            position:fixed;
            inset:0;
            z-index:10000;
            background:rgba(0,0,0,.86);
            align-items:center;
            justify-content:center;
            padding:20px;
        }}
        #modalExclusao.ativo {{ display:flex; }}
        .modal-exclusao-caixa {{
            width:min(460px,96vw);
            background:#081323;
            border:1px solid rgba(240,182,77,.65);
            border-radius:16px;
            padding:22px;
            box-shadow:0 20px 80px rgba(0,0,0,.55);
        }}
        .modal-exclusao-caixa h2 {{
            margin:0 0 8px;
            color:#f0b64d;
        }}
        .modal-exclusao-caixa p {{
            color:#d8dde6;
            line-height:1.45;
        }}
        .modal-exclusao-caixa input {{
            width:100%;
            background:#02060d;
            color:#fff;
            border:1px solid #586274;
            border-radius:10px;
            padding:12px;
            font-size:16px;
            margin:8px 0 12px;
        }}
        .modal-acoes {{
            display:flex;
            gap:10px;
            flex-wrap:wrap;
        }}
        .modal-acoes button {{
            border:0;
            padding:11px 14px;
            border-radius:10px;
            font-weight:800;
            cursor:pointer;
        }}
        .confirmar-exclusao {{ background:#c83d3d; color:#fff; }}
        .cancelar-exclusao {{ background:#2a3546; color:#fff; }}
        #mensagemExclusao {{
            min-height:22px;
            margin-top:10px;
            color:#ffd078;
        }}
        footer {{ text-align:center; color:#bfc6d1; border-top:1px solid rgba(240,182,77,.25); padding:18px; }}
        #lightbox {{ display:none; position:fixed; inset:0; background:rgba(0,0,0,.94); z-index:9999; align-items:center; justify-content:center; padding:22px; }}
        #lightbox.ativo {{ display:flex; }}
        #lightbox img {{ max-width:96vw; max-height:92vh; object-fit:contain; border-radius:12px; box-shadow:0 0 30px rgba(0,0,0,.7); }}
        #fecharLightbox {{ position:fixed; top:18px; right:22px; background:#f0b64d; color:#06111f; border:0; padding:10px 15px; border-radius:10px; font-weight:900; cursor:pointer; }}
        #legendaLightbox {{ position:fixed; bottom:16px; left:50%; transform:translateX(-50%); color:#fff; background:rgba(0,0,0,.65); padding:8px 12px; border-radius:10px; }}
        @media (max-width:720px) {{
            .grade {{ grid-template-columns:1fr; }}
            .card-grid {{ grid-template-columns:1fr; }}
            .card-head {{ grid-template-columns:1fr; }}
            .photo img {{ height:330px; }}
        }}
    </style>
</head>
<body>
    <header>
        <h1>Catálogo de Procurados</h1>
        <p>DICOR • Atualizado automaticamente pelo bot • Clique nas fotos para abrir em tela cheia</p>
        <div class="stats">
            <div class="stat"><span>Registros totais</span><b>{len(visiveis)}</b></div>
            <div class="stat"><span>Ativos</span><b>{len(ativos)}</b></div>
            <div class="stat"><span>Retirados</span><b>{len(retirados)}</b></div>
            <div class="stat"><span>Última atualização</span><b style="font-size:16px">{agora_br()}</b></div>
        </div>
        <div class="abas">
            <button id="btn-ativos" class="aba-btn ativa" onclick="mostrarAba('ativos')">🚨 Procurados Ativos ({len(ativos)})</button>
            <button id="btn-retirados" class="aba-btn" onclick="mostrarAba('retirados')">✅ Retirados ({len(retirados)})</button>
        </div>
    </header>

    <section id="aba-ativos" class="aba-conteudo ativa">
        <main class="grade">{cards_ativos}</main>
    </section>

    <section id="aba-retirados" class="aba-conteudo">
        <main class="grade">{cards_retirados}</main>
    </section>

    <footer>Uso exclusivo para GTA RP / sistema interno autorizado.</footer>

    <div id="modalExclusao" onclick="fecharExclusao(event)">
        <div class="modal-exclusao-caixa" onclick="event.stopPropagation()">
            <h2>🗑️ Apagar registro</h2>
            <p id="textoExclusao">
                Essa ação remove o procurado de forma permanente das abas de ativos e retirados.
            </p>
            <input
                id="senhaExclusao"
                type="password"
                autocomplete="current-password"
                placeholder="Senha administrativa do catálogo"
            >
            <div class="modal-acoes">
                <button class="confirmar-exclusao" type="button" onclick="confirmarExclusao()">
                    Apagar permanentemente
                </button>
                <button class="cancelar-exclusao" type="button" onclick="fecharExclusao()">
                    Cancelar
                </button>
            </div>
            <div id="mensagemExclusao"></div>
        </div>
    </div>

    <div id="lightbox" onclick="fecharFoto()">
        <button id="fecharLightbox" onclick="fecharFoto()">Fechar ✕</button>
        <img id="lightboxImg" src="" alt="Foto em tela cheia">
        <div id="legendaLightbox"></div>
    </div>

    <script>
        let registroExclusaoId = '';
        let registroExclusaoNome = '';

        function abrirExclusao(id, nome) {{
            registroExclusaoId = id || '';
            registroExclusaoNome = nome || 'Sem nome';
            document.getElementById('textoExclusao').textContent =
                'Apagar permanentemente "' + registroExclusaoNome +
                '"? Ele não aparecerá mais em nenhuma aba do catálogo.';
            document.getElementById('senhaExclusao').value = '';
            document.getElementById('mensagemExclusao').textContent = '';
            document.getElementById('modalExclusao').classList.add('ativo');
            setTimeout(
                () => document.getElementById('senhaExclusao').focus(),
                50
            );
        }}

        function fecharExclusao(event) {{
            if (
                event
                && event.target
                && event.target.id !== 'modalExclusao'
            ) {{
                return;
            }}
            document.getElementById('modalExclusao').classList.remove('ativo');
            registroExclusaoId = '';
        }}

        async function confirmarExclusao() {{
            const senha = document.getElementById('senhaExclusao').value;
            const mensagem = document.getElementById('mensagemExclusao');

            if (!senha) {{
                mensagem.textContent = 'Digite a senha administrativa.';
                return;
            }}

            mensagem.textContent = 'Apagando registro...';
            try {{
                const resposta = await fetch('/api/catalogo/apagar', {{
                    method: 'POST',
                    headers: {{'Content-Type': 'application/json'}},
                    body: JSON.stringify({{
                        id: registroExclusaoId,
                        senha: senha
                    }})
                }});
                const dados = await resposta.json();

                if (!resposta.ok || !dados.ok) {{
                    mensagem.textContent =
                        dados.erro || 'Não foi possível apagar o registro.';
                    return;
                }}

                mensagem.textContent =
                    'Registro apagado. Atualizando o catálogo...';
                window.location.reload();
            }} catch (erro) {{
                mensagem.textContent =
                    'Falha ao conectar ao servidor do catálogo.';
            }}
        }}

        function mostrarAba(nome) {{
            document.querySelectorAll('.aba-conteudo').forEach(el => el.classList.remove('ativa'));
            document.querySelectorAll('.aba-btn').forEach(el => el.classList.remove('ativa'));
            document.getElementById('aba-' + nome).classList.add('ativa');
            document.getElementById('btn-' + nome).classList.add('ativa');
            window.scrollTo({{ top: 0, behavior: 'smooth' }});
        }}
        function abrirFoto(src, alt) {{
            document.getElementById('lightboxImg').src = src;
            document.getElementById('legendaLightbox').textContent = alt || '';
            document.getElementById('lightbox').classList.add('ativo');
            document.body.style.overflow = 'hidden';
        }}
        function fecharFoto() {{
            document.getElementById('lightbox').classList.remove('ativo');
            document.getElementById('lightboxImg').src = '';
            document.getElementById('legendaLightbox').textContent = '';
            document.body.style.overflow = '';
        }}
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'Escape') {{
                fecharFoto();
                document
                    .getElementById('modalExclusao')
                    .classList.remove('ativo');
            }}
        }});
        document.getElementById('lightboxImg').addEventListener('click', function(e) {{ e.stopPropagation(); }});
        document.getElementById('fecharLightbox').addEventListener('click', function(e) {{ e.stopPropagation(); }});
    </script>
</body>
</html>
    """
    with open(CATALOGO_HTML, "w", encoding="utf-8") as f:
        f.write(html_final)

async def salvar_anexo_publico(anexo: discord.Attachment, prefixo: str) -> str:
    ext = Path(anexo.filename).suffix.lower()
    if ext not in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
        ext = ".png"
    nome = f"{data_caso()}-{slugify(prefixo)}-{slugify(Path(anexo.filename).stem)}{ext}"
    caminho = UPLOADS_DIR / nome
    await anexo.save(str(caminho))
    return f"uploads/{nome}"


async def salvar_procurado_catalogo(registro: Dict[str, Any]) -> None:
    registro["crimes"] = valor_crimes_registro(registro)
    lista = carregar_procurados()
    lista.append(registro)
    lista = remover_duplicados(lista)
    salvar_procurados(lista)
    gerar_catalogo_html()

# =====================================================
# SERVIDOR HTML
# =====================================================

async def pagina_inicial(request):
    gerar_catalogo_html()
    return web.FileResponse(CATALOGO_HTML)


async def obter_canal_por_id(canal_id: int):
    if not canal_id:
        return None

    canal = bot.get_channel(canal_id)
    if canal is not None:
        return canal

    try:
        return await bot.fetch_channel(canal_id)
    except Exception:
        return None


async def apagar_mensagens_discord_do_registro(
    registro: Dict[str, Any],
) -> None:
    """Apaga o post ativo e/ou o post arquivado relacionado ao registro."""
    alvos = [
        (PROCURADOS_CHANNEL_ID, registro.get("mensagem_id")),
        (
            HISTORICO_PROCURADOS_ID,
            registro.get("mensagem_arquivada_id"),
        ),
    ]

    for canal_id, mensagem_id in alvos:
        if not canal_id or not mensagem_id:
            continue

        canal = await obter_canal_por_id(int(canal_id))
        if canal is None or not hasattr(canal, "fetch_message"):
            continue

        try:
            mensagem = await canal.fetch_message(int(mensagem_id))
            await mensagem.delete()
        except discord.NotFound:
            pass
        except Exception as erro:
            await enviar_log(
                f"⚠️ Não foi possível apagar a mensagem `{mensagem_id}` "
                f"do registro `{registro.get('id')}`: {erro}"
            )


def apagar_arquivos_locais_do_registro(
    registro: Dict[str, Any],
) -> None:
    """Remove as fotos locais pertencentes ao registro apagado."""
    for chave in ("foto_individuo", "foto_rg"):
        caminho_relativo = str(registro.get(chave, "") or "").strip()

        if (
            not caminho_relativo
            or not caminho_relativo.startswith("uploads/")
        ):
            continue

        caminho = PUBLIC_DIR / caminho_relativo
        try:
            caminho_resolvido = caminho.resolve()
            uploads_resolvido = UPLOADS_DIR.resolve()

            if (
                uploads_resolvido in caminho_resolvido.parents
                and caminho_resolvido.exists()
            ):
                caminho_resolvido.unlink()
        except Exception:
            pass


async def api_apagar_procurado(
    request: web.Request,
) -> web.Response:
    ip = request.remote or "desconhecido"

    if not catalogo_limite_tentativas(ip):
        return web.json_response(
            {
                "ok": False,
                "erro": "Muitas tentativas. Aguarde 10 minutos.",
            },
            status=429,
        )

    try:
        dados = await request.json()
    except Exception:
        return web.json_response(
            {"ok": False, "erro": "Requisição inválida."},
            status=400,
        )

    senha = str(dados.get("senha", "") or "")
    registro_id = str(dados.get("id", "") or "").strip()

    if not senha_catalogo_valida(senha):
        registrar_falha_senha_catalogo(ip)
        return web.json_response(
            {"ok": False, "erro": "Senha incorreta."},
            status=401,
        )

    if not registro_id:
        return web.json_response(
            {"ok": False, "erro": "Registro não informado."},
            status=400,
        )

    lista = carregar_procurados()
    encontrado = None
    nova_lista: List[Dict[str, Any]] = []

    for registro in lista:
        identificador = str(
            registro.get("id")
            or registro.get("caso")
            or registro.get("rg")
            or ""
        )

        if encontrado is None and identificador == registro_id:
            encontrado = registro
            continue

        nova_lista.append(registro)

    if encontrado is None:
        return web.json_response(
            {
                "ok": False,
                "erro": "O procurado já foi apagado ou não existe.",
            },
            status=404,
        )

    # Remove fisicamente do JSON. Assim não aparece em nenhuma aba.
    salvar_procurados(nova_lista)
    gerar_catalogo_html()

    # Limpa, quando possível, os posts e as fotos relacionados.
    await apagar_mensagens_discord_do_registro(encontrado)
    apagar_arquivos_locais_do_registro(encontrado)

    catalogo_tentativas_senha.pop(ip, None)

    await enviar_log(
        f"🗑️ **Procurado apagado permanentemente pelo catálogo**\n"
        f"Nome: {encontrado.get('nome')}\n"
        f"RG: {encontrado.get('rg')}\n"
        f"Registro: {registro_id}"
    )

    return web.json_response(
        {
            "ok": True,
            "mensagem": "Registro apagado permanentemente.",
        }
    )


async def start_web_server():
    gerar_catalogo_html()
    app = web.Application()
    app.router.add_get("/", pagina_inicial)
    app.router.add_get("/index.html", pagina_inicial)
    app.router.add_post(
        "/api/catalogo/apagar",
        api_apagar_procurado,
    )
    app.router.add_static(
        "/uploads/",
        path=str(UPLOADS_DIR),
        show_index=False,
    )
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    print(f"Catalogo rodando na porta {PORT}")

# =====================================================
# PROCURADOS
# =====================================================

def criar_texto_procurado(registro: Dict[str, Any]) -> str:
    return f"""
🚨 **MANDADO DE PRISÃO E PROCURAÇÃO INVESTIGATIVA** 🚨

A Polícia Polícia Federal de Capital Morada, por intermédio da **Divisão de Investigações Criminais (DICOR)**, informa que o indivíduo abaixo encontra-se oficialmente procurado pelas autoridades competentes.

As investigações apontam seu envolvimento em atividades criminosas, havendo mandado ativo para sua localização, abordagem e condução para os procedimentos cabíveis.

📍 **ÚLTIMO AVISTAMENTO:** {registro.get('ultimo_avistamento', 'Não informado')}

⚠️ **CRIMES IMPUTADOS:**
{registro.get('crimes', 'Não informado')}

━━━━━━━━━━━━━━━━━━━━━━━

🆔 **IDENTIFICAÇÃO DO PROCURADO**

👤 **Nome:** {registro.get('nome', 'Não informado')}
🆔 **RG:** {registro.get('rg', 'Não informado')}

━━━━━━━━━━━━━━━━━━━━━━━

📞 Qualquer informação sobre o paradeiro deste indivíduo deverá ser repassada imediatamente a um agente da Polícia Federal ou da DICOR.

🔒 O sigilo do denunciante será integralmente preservado.

🔹 Polícia Polícia Federal de Capital Morada
🔹 Divisão de Investigações Criminais (DICOR)
""".strip()


def criar_embed_procurado(registro: Dict[str, Any]) -> discord.Embed:
    cor = discord.Color.red() if registro.get("status") == "A PROCURAR" else discord.Color.dark_gray()
    embed = discord.Embed(description=criar_texto_procurado(registro), color=cor)
    embed.set_footer(text=f"Caso: {registro.get('caso')} • Cadastro: {registro.get('data')}")
    return embed


async def postar_procurado_oficial(registro: Dict[str, Any]) -> Optional[discord.Message]:
    canal = bot.get_channel(PROCURADOS_CHANNEL_ID)
    if not canal:
        return None

    arquivos = []
    for chave, nome in [("foto_individuo", "foto_individuo"), ("foto_rg", "foto_rg")]:
        rel = registro.get(chave, "")
        if rel:
            caminho = PUBLIC_DIR / rel
            if caminho.exists():
                arquivos.append(discord.File(str(caminho), filename=f"{nome}_{registro.get('rg','')}.png"))

    # Mensagem normal, sem embed, para ficar igual ao modelo antigo.
    # As imagens seguem anexadas no mesmo envio e aparecem embaixo da mensagem.
    texto = cortar_discord(criar_texto_procurado(registro), 1900)
    msg = await canal.send(content=texto, files=arquivos)
    return msg


def arquivos_locais_procurado(
    registro: Dict[str, Any],
) -> List[discord.File]:
    arquivos: List[discord.File] = []

    for chave, nome_base in (
        ("foto_individuo", "foto_individuo"),
        ("foto_rg", "foto_rg"),
    ):
        caminho_relativo = str(registro.get(chave, "") or "").strip()
        if not caminho_relativo:
            continue

        caminho = PUBLIC_DIR / caminho_relativo
        if not caminho.exists():
            continue

        extensao = caminho.suffix.lower() or ".png"
        arquivos.append(
            discord.File(
                str(caminho),
                filename=(
                    f"{nome_base}_{registro.get('rg', '')}{extensao}"
                ),
            )
        )

    return arquivos


async def arquivar_procurado_discord(
    registro: Dict[str, Any],
    motivo: str,
    retirado_por: str,
) -> discord.Message:
    """
    Publica o procurado completo no histórico.
    O post ativo só é apagado depois do arquivamento confirmado.
    """
    canal_historico = await obter_canal_por_id(
        HISTORICO_PROCURADOS_ID
    )

    if (
        canal_historico is None
        or not hasattr(canal_historico, "send")
    ):
        raise RuntimeError(
            f"Canal de histórico `{HISTORICO_PROCURADOS_ID}` "
            "não encontrado."
        )

    mensagem_original = None
    if registro.get("mensagem_id") and PROCURADOS_CHANNEL_ID:
        canal_ativo = await obter_canal_por_id(
            PROCURADOS_CHANNEL_ID
        )

        if (
            canal_ativo is not None
            and hasattr(canal_ativo, "fetch_message")
        ):
            try:
                mensagem_original = await canal_ativo.fetch_message(
                    int(registro["mensagem_id"])
                )
            except Exception:
                mensagem_original = None

    arquivos: List[discord.File] = []

    if mensagem_original is not None:
        for anexo in mensagem_original.attachments[:10]:
            try:
                arquivos.append(
                    await anexo.to_file(use_cached=True)
                )
            except Exception:
                pass

    if not arquivos:
        arquivos = arquivos_locais_procurado(registro)

    texto_arquivado = cortar_discord(
        criar_texto_procurado(registro)
        + "\n\n━━━━━━━━━━━━━━━━━━━━━━━"
        + "\n📁 **STATUS:** ARQUIVADO / RETIRADO"
        + f"\n📌 **Motivo:** {motivo}"
        + f"\n👮 **Retirado por:** {retirado_por}"
        + f"\n🕒 **Data da retirada:** {agora_br()}",
        1900,
    )

    mensagem_arquivada = await canal_historico.send(
        content=texto_arquivado,
        files=arquivos,
    )

    if mensagem_original is not None:
        try:
            await mensagem_original.delete(
                reason=(
                    "Procurado retirado e arquivado no histórico"
                )
            )
        except Exception as erro:
            await enviar_log(
                "⚠️ Procurado arquivado, mas não foi possível "
                f"apagar a mensagem ativa "
                f"`{registro.get('mensagem_id')}`: {erro}"
            )

    return mensagem_arquivada


class NovoProcuradoModal(Modal, title="Cadastrar Novo Procurado"):
    nome = TextInput(label="Nome", placeholder="Nome do procurado", max_length=100)
    rg = TextInput(label="RG", placeholder="RG do procurado", max_length=50)
    crimes = TextInput(label="Crimes Cometidos", placeholder="Ex: Art. 8.3\nFormação de Quadrilha", style=discord.TextStyle.paragraph, max_length=1000)
    ultimo = TextInput(label="Último Avistamento", placeholder="Ex: Caixa d'água", style=discord.TextStyle.paragraph, max_length=600)
    informacoes = TextInput(label="Informações", placeholder="Detalhes extras da investigação", style=discord.TextStyle.paragraph, max_length=1200, required=False)

    async def on_submit(self, interaction: discord.Interaction):
        guild = interaction.guild
        if guild is None:
            await interaction.response.send_message("❌ Use isso dentro de um servidor.", ephemeral=True)
            return

        categoria = guild.get_channel(PROCURADOS_TEMP_CATEGORY_ID) if PROCURADOS_TEMP_CATEGORY_ID else None
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(view_channel=False),
            interaction.user: discord.PermissionOverwrite(view_channel=True, send_messages=True, attach_files=True, read_message_history=True),
        }
        if guild.me:
            overwrites[guild.me] = discord.PermissionOverwrite(view_channel=True, send_messages=True, manage_channels=True, read_message_history=True, attach_files=True)
        for cargo_id in set(CARGOS_ADMIN_IDS + CARGOS_EQUIPE_IDS):
            cargo = guild.get_role(cargo_id)
            if cargo:
                overwrites[cargo] = discord.PermissionOverwrite(view_channel=True, send_messages=True, read_message_history=True, attach_files=True)

        nome_canal = f"📸・procurado-{slugify(str(self.nome.value))}"
        canal = await guild.create_text_channel(name=nome_canal, category=categoria, overwrites=overwrites)

        cadastros_pendentes[canal.id] = {
            "nome": str(self.nome.value),
            "rg": str(self.rg.value),
            "crimes": str(self.crimes.value),
            "ultimo_avistamento": str(self.ultimo.value),
            "informacoes": str(self.informacoes.value or ""),
            "autor_id": interaction.user.id,
            "autor_nome": str(interaction.user),
        }

        await canal.send(
            f"🚨 **Cadastro de Procurado — DICOR**\n\n"
            f"👤 **Nome:** {self.nome.value}\n"
            f"🪪 **RG:** {self.rg.value}\n\n"
            f"📸 Envie aqui **2 imagens obrigatórias**:\n"
            f"**1. Foto do indivíduo**\n"
            f"**2. Foto do RG**\n\n"
            f"Depois clique em **✅ Finalizar Cadastro**.",
            view=FinalizarProcuradoView()
        )

        await enviar_log(f"📁 Canal provisório de procurado criado por {interaction.user.mention}: {canal.mention}")
        await interaction.response.send_message(f"✅ Canal provisório criado: {canal.mention}", ephemeral=True)


class FinalizarProcuradoView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="Finalizar Cadastro", emoji="✅", style=discord.ButtonStyle.green, custom_id="dic_finalizar_procurado")
    async def finalizar(self, interaction: discord.Interaction, button: Button):
        canal = interaction.channel
        if not isinstance(canal, discord.TextChannel):
            await interaction.response.send_message("❌ Canal inválido.", ephemeral=True)
            return

        dados = cadastros_pendentes.get(canal.id)
        if not dados:
            await interaction.response.send_message("❌ Não encontrei os dados desse cadastro. Crie novamente pelo painel.", ephemeral=True)
            return

        if interaction.user.id != dados.get("autor_id") and isinstance(interaction.user, discord.Member) and not usuario_tem_admin(interaction.user):
            await interaction.response.send_message("❌ Apenas quem criou ou a equipe pode finalizar.", ephemeral=True)
            return

        anexos: List[discord.Attachment] = []
        async for msg in canal.history(limit=100, oldest_first=True):
            for a in msg.attachments:
                if (a.content_type and a.content_type.startswith("image/")) or Path(a.filename).suffix.lower() in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
                    anexos.append(a)

        if len(anexos) < 2:
            await interaction.response.send_message("❌ Envie as 2 imagens: foto do indivíduo e foto do RG.", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True)
        foto_ind = await salvar_anexo_publico(anexos[0], f"individuo-{dados['rg']}")
        foto_rg = await salvar_anexo_publico(anexos[1], f"rg-{dados['rg']}")

        registro = {
            "id": data_caso(),
            "caso": f"DICOR-{data_caso()}",
            "data": agora_br(),
            "status": "A PROCURAR",
            "nome": dados["nome"],
            "rg": dados["rg"],
            "crimes": dados["crimes"],
            "ultimo_avistamento": dados["ultimo_avistamento"],
            "informacoes": dados["informacoes"],
            "foto_individuo": foto_ind,
            "foto_rg": foto_rg,
            "autor_id": dados["autor_id"],
            "autor_nome": dados["autor_nome"],
            "mensagem_id": None,
            "mensagem_url": None,
        }

        msg = await postar_procurado_oficial(registro)
        if msg:
            registro["mensagem_id"] = msg.id
            registro["mensagem_url"] = msg.jump_url

        await salvar_procurado_catalogo(registro)
        cadastros_pendentes.pop(canal.id, None)

        await enviar_log(f"✅ Procurado cadastrado: {registro['nome']} | RG {registro['rg']} | Catálogo: {CATALOG_PUBLIC_URL}")
        await interaction.followup.send(f"✅ Procurado cadastrado e enviado ao catálogo: {CATALOG_PUBLIC_URL}", ephemeral=True)
        await canal.send("✅ Cadastro finalizado. Este canal será apagado em 5 segundos.")
        await asyncio.sleep(5)
        try:
            await canal.delete(reason="Cadastro de procurado finalizado")
        except Exception:
            pass

    @discord.ui.button(label="Cancelar", emoji="❌", style=discord.ButtonStyle.red, custom_id="dic_cancelar_procurado")
    async def cancelar(self, interaction: discord.Interaction, button: Button):
        canal = interaction.channel
        if isinstance(canal, discord.TextChannel):
            cadastros_pendentes.pop(canal.id, None)
            await interaction.response.send_message("Cadastro cancelado. Canal será apagado.", ephemeral=True)
            await asyncio.sleep(3)
            try:
                await canal.delete(reason="Cadastro de procurado cancelado")
            except Exception:
                pass


class RetirarProcuradoModal(Modal, title="Retirar Procurado"):
    rg = TextInput(label="RG", placeholder="RG do procurado", max_length=50)
    motivo = TextInput(label="Motivo", placeholder="Motivo da retirada", style=discord.TextStyle.paragraph, max_length=800, required=False)

    async def on_submit(self, interaction: discord.Interaction):
        await retirar_procurado(interaction, str(self.rg.value), str(self.motivo.value or "Não informado"))


class PainelProcuradosView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="Novo Procurado", emoji="➕", style=discord.ButtonStyle.danger, custom_id="dic_novo_procurado")
    async def novo(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(NovoProcuradoModal())

    @discord.ui.button(label="Lista de Procurados", emoji="📋", style=discord.ButtonStyle.blurple, custom_id="dic_lista_procurados")
    async def lista(self, interaction: discord.Interaction, button: Button):
        procurados = carregar_procurados()
        ativos = [p for p in procurados if p.get("status", "A PROCURAR") == "A PROCURAR"]
        if not ativos:
            texto = "Nenhum procurado ativo cadastrado."
        else:
            linhas = [f"• **{p.get('nome','Sem nome')}** — RG: `{p.get('rg','')}`" for p in ativos[:20]]
            texto = "\n".join(linhas)
            if len(ativos) > 20:
                texto += f"\n... e mais {len(ativos)-20} no catálogo."
        await interaction.response.send_message(f"📋 **Procurados ativos:**\n{texto}\n\n🔗 {CATALOG_PUBLIC_URL}", ephemeral=True)

    @discord.ui.button(label="Retirar Procurado", emoji="❌", style=discord.ButtonStyle.gray, custom_id="dic_retirar_procurado")
    async def retirar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(RetirarProcuradoModal())

    @discord.ui.button(label="Abrir Catálogo", emoji="📄", style=discord.ButtonStyle.green, custom_id="dic_abrir_catalogo")
    async def abrir_catalogo(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_message(f"📄 **Catálogo de Procurados:**\n{CATALOG_PUBLIC_URL}", ephemeral=True)

    @discord.ui.button(label="Sincronizar Antigos", emoji="🔄", style=discord.ButtonStyle.secondary, custom_id="dic_sync_catalogo")
    async def sync_old(self, interaction: discord.Interaction, button: Button):
        await sincronizar_catalogo_core(interaction)


async def retirar_procurado(
    interaction: discord.Interaction,
    rg: str,
    motivo: str,
):
    if not interaction.response.is_done():
        await interaction.response.defer(
            ephemeral=True,
            thinking=True,
        )

    lista = carregar_procurados()
    alvo = limpar_rg(rg)
    encontrado = None

    for registro in lista:
        if limpar_rg(registro.get("rg", "")) == alvo:
            encontrado = registro
            break

    if not encontrado:
        await interaction.followup.send(
            "❌ Não encontrei procurado com esse RG no catálogo.",
            ephemeral=True,
        )
        return

    if (
        str(encontrado.get("status", "") or "").upper()
        == "RETIRADO"
        and encontrado.get("mensagem_arquivada_id")
    ):
        await interaction.followup.send(
            "⚠️ Esse procurado já foi retirado e está arquivado.",
            ephemeral=True,
        )
        return

    try:
        mensagem_arquivada = await arquivar_procurado_discord(
            encontrado,
            motivo,
            interaction.user.mention,
        )
    except Exception as erro:
        await enviar_log(
            f"❌ Falha ao arquivar procurado "
            f"{encontrado.get('nome')} "
            f"(RG {encontrado.get('rg')}): {erro}"
        )
        await interaction.followup.send(
            "❌ Não foi possível enviar o procurado para o "
            "canal de arquivados. Nada foi removido. "
            "Confira o ID e as permissões do canal "
            f"`{HISTORICO_PROCURADOS_ID}`.",
            ephemeral=True,
        )
        return

    encontrado["status"] = "RETIRADO"
    encontrado["motivo_retirada"] = motivo
    encontrado["data_retirada"] = agora_br()
    encontrado["retirado_por"] = str(interaction.user)
    encontrado["mensagem_original_id"] = encontrado.get(
        "mensagem_id"
    )
    encontrado["mensagem_original_url"] = encontrado.get(
        "mensagem_url"
    )
    encontrado["mensagem_arquivada_id"] = (
        mensagem_arquivada.id
    )
    encontrado["mensagem_arquivada_url"] = (
        mensagem_arquivada.jump_url
    )
    # No catálogo, o link passa a abrir o post arquivado.
    encontrado["mensagem_url"] = mensagem_arquivada.jump_url

    salvar_procurados(lista)
    gerar_catalogo_html()

    await enviar_log(
        f"📁 **Procurado retirado e arquivado**\n"
        f"Nome: {encontrado.get('nome')}\n"
        f"RG: {encontrado.get('rg')}\n"
        f"Motivo: {motivo}\n"
        f"Arquivo: {mensagem_arquivada.jump_url}"
    )

    await interaction.followup.send(
        "✅ Procurado retirado, movido para o canal de "
        "arquivados e catálogo atualizado.\n"
        f"📁 {mensagem_arquivada.jump_url}\n"
        f"🔗 {CATALOG_PUBLIC_URL}",
        ephemeral=True,
    )

def limpar_markdown_extracao(texto: str) -> str:
    texto = str(texto or "").replace("**", "").replace("__", "")
    return texto.replace("`", "")


def extrair_bloco_rotulado(texto: str, rotulo: str, proximos: List[str]) -> str:
    fim = "|".join(proximos)
    padrao = rf"(?:{rotulo})\s*:\s*(.*?)(?=\n\s*(?:{fim})\s*:|\n\s*━+|\Z)"
    m = re.search(padrao, texto, flags=re.I | re.S)
    if not m:
        return ""
    valor = re.sub(r"\n{3,}", "\n\n", m.group(1).strip())
    return valor[:1800]


def extrair_por_labels(texto: str) -> Dict[str, str]:
    texto_limpo = limpar_markdown_extracao(texto)
    dados: Dict[str, str] = {}

    m_nome = re.search(r"(?:nome|indiv[ií]duo)\s*[:\-]\s*([^\n]+)", texto_limpo, flags=re.I)
    m_rg = re.search(r"(?:rg|passaporte)\s*[:\-]\s*([^\n]+)", texto_limpo, flags=re.I)
    if m_nome:
        dados["nome"] = m_nome.group(1).strip()[:200]
    if m_rg:
        dados["rg"] = m_rg.group(1).strip()[:100]

    proximos_comuns = [
        r"(?:📍\s*)?(?:último avistamento|ultimo avistamento|avistamento)",
        r"(?:⚠️\s*)?(?:crimes cometidos|crimes imputados|crimes|crime)",
        r"(?:ℹ️\s*)?(?:informações|informacoes|observações|observacoes|detalhes)",
        r"(?:🆔\s*)?(?:identificação do procurado|identificacao do procurado)",
        r"(?:👤\s*)?nome",
        r"(?:🆔\s*)?rg",
    ]

    crimes = extrair_bloco_rotulado(
        texto_limpo,
        r"(?:⚠️\s*)?(?:crimes cometidos|crimes imputados|crimes|crime)",
        proximos_comuns,
    )
    ultimo = extrair_bloco_rotulado(
        texto_limpo,
        r"(?:📍\s*)?(?:último avistamento|ultimo avistamento|avistamento)",
        proximos_comuns,
    )
    informacoes = extrair_bloco_rotulado(
        texto_limpo,
        r"(?:ℹ️\s*)?(?:informações|informacoes|observações|observacoes|detalhes)",
        proximos_comuns,
    )

    if crimes:
        dados["crimes"] = crimes
    if ultimo:
        dados["ultimo_avistamento"] = ultimo
    if informacoes:
        dados["informacoes"] = informacoes
    return dados


async def importar_mensagem_antiga(msg: discord.Message) -> Optional[Dict[str, Any]]:
    texto_total = msg.content or ""
    for embed in msg.embeds:
        if embed.title:
            texto_total += "\n" + embed.title
        if embed.description:
            texto_total += "\n" + embed.description
        for field in embed.fields:
            texto_total += f"\n{field.name}: {field.value}"

    dados = extrair_por_labels(texto_total)
    if not dados.get("rg") or not dados.get("nome"):
        return None

    existente = procurar_por_rg(dados.get("rg", ""))

    imagens = []
    for anexo in msg.attachments:
        if (
            anexo.content_type and anexo.content_type.startswith("image/")
        ) or Path(anexo.filename).suffix.lower() in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
            imagens.append(anexo)

    foto_ind = str(existente.get("foto_individuo", "") if existente else "")
    foto_rg = str(existente.get("foto_rg", "") if existente else "")
    if not foto_ind and len(imagens) >= 1:
        foto_ind = await salvar_anexo_publico(imagens[0], f"sync-ind-{dados.get('rg', '')}")
    if not foto_rg and len(imagens) >= 2:
        foto_rg = await salvar_anexo_publico(imagens[1], f"sync-rg-{dados.get('rg', '')}")

    return {
        "id": f"SYNC-{msg.id}",
        "caso": f"SYNC-{msg.id}",
        "data": msg.created_at.strftime("%d/%m/%Y %H:%M"),
        "status": str(existente.get("status", "A PROCURAR") if existente else "A PROCURAR"),
        "nome": dados.get("nome", ""),
        "rg": dados.get("rg", ""),
        "crimes": dados.get("crimes", "Não identificado na mensagem antiga"),
        "ultimo_avistamento": dados.get("ultimo_avistamento", "Não identificado na mensagem antiga"),
        "informacoes": dados.get("informacoes", "Importado automaticamente do histórico do canal."),
        "foto_individuo": foto_ind,
        "foto_rg": foto_rg,
        "autor_id": None,
        "autor_nome": "Sincronização antiga",
        "mensagem_id": msg.id,
        "mensagem_url": msg.jump_url,
    }

async def sincronizar_catalogo_core(interaction: discord.Interaction):
    if not PROCURADOS_CHANNEL_ID:
        await interaction.response.send_message("❌ Configure PROCURADOS_CHANNEL_ID.", ephemeral=True)
        return

    canal = bot.get_channel(PROCURADOS_CHANNEL_ID)
    if not isinstance(canal, discord.TextChannel):
        await interaction.response.send_message("❌ Não encontrei o canal de procurados.", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)
    importados = 0
    atualizados = 0
    analisados = 0
    lista = carregar_procurados()

    async for msg in canal.history(limit=1000, oldest_first=True):
        analisados += 1
        registro = await importar_mensagem_antiga(msg)
        if not registro:
            continue

        alvo = limpar_rg(registro.get("rg", ""))
        existente = next(
            (p for p in lista if limpar_rg(p.get("rg", "")) == alvo),
            None,
        )

        if existente is None:
            registro["crimes"] = valor_crimes_registro(registro)
            lista.append(registro)
            importados += 1
            continue

        mudou = False
        for campo in (
            "nome",
            "crimes",
            "ultimo_avistamento",
            "informacoes",
            "foto_individuo",
            "foto_rg",
            "mensagem_id",
            "mensagem_url",
        ):
            novo_valor = registro.get(campo)
            valor_atual = existente.get(campo)
            if registro_tem_valor_util(novo_valor) and not registro_tem_valor_util(valor_atual):
                existente[campo] = novo_valor
                mudou = True

        crimes_novos = valor_crimes_registro(registro)
        if registro_tem_valor_util(crimes_novos) and not registro_tem_valor_util(existente.get("crimes")):
            existente["crimes"] = crimes_novos
            mudou = True

        if mudou:
            atualizados += 1

    lista = remover_duplicados(lista)
    salvar_procurados(lista)
    gerar_catalogo_html()

    await interaction.followup.send(
        f"✅ Sincronização finalizada.\n"
        f"Mensagens analisadas: `{analisados}`\n"
        f"Procurados importados: `{importados}`\n"
        f"Registros corrigidos/atualizados: `{atualizados}`\n"
        f"Catálogo: {CATALOG_PUBLIC_URL}",
        ephemeral=True,
    )

# =====================================================
# MESAS
# =====================================================

TOPICOS_MESA = [
    "📌 Painel",
    "🖼️ Fotos líderes",
    "👥 Fotos dos membros",
    "📻 Rádio",
    "📍 Localização",
    "⚖️ Crimes da comunidade",
    "📦 Baú de líder",
    "📦 Baú de membros",
    "🌾 Rota de farm",
    "⚙️ Rota de produção",
    "🧪 Ingredientes base e produtos",
    "🕵️ Informante",
    "💬 Chat",
]



async def criar_topicos_reais_mesa(canal: discord.TextChannel) -> List[int]:
    """Cria cada item da mesa como tópico/thread público e já aberto."""
    topicos_ids: List[int] = []

    await canal.send(
        "🧵 **Tópicos da investigação**\n"
        "Abra o tópico desejado para adicionar textos, imagens e provas."
    )

    for topico in TOPICOS_MESA:
        mensagem = await canal.send(f"🧵 **{topico}**")
        try:
            try:
                thread = await mensagem.create_thread(
                    name=topico[:100],
                    auto_archive_duration=10080,
                    reason="Tópico automático da mesa de investigação",
                )
            except discord.HTTPException:
                thread = await mensagem.create_thread(
                    name=topico[:100],
                    auto_archive_duration=1440,
                    reason="Tópico automático da mesa de investigação",
                )

            await thread.send(
                f"## {topico}\n"
                "> Envie aqui todas as informações, imagens e provas relacionadas a este tópico."
            )
            topicos_ids.append(thread.id)
            await asyncio.sleep(0.25)
        except Exception as erro:
            await mensagem.reply(
                f"⚠️ Não foi possível abrir este tópico automaticamente: `{erro}`"
            )
            await enviar_log(
                f"⚠️ Erro ao criar tópico `{topico}` no canal {canal.id}: {erro}"
            )

    return topicos_ids

def registrar_mesa(dados: Dict[str, Any]) -> None:
    mesas = carregar_mesas()
    mesas = [m for m in mesas if m.get("canal_id") != dados.get("canal_id")]
    mesas.append(dados)
    salvar_mesas(mesas)


def buscar_mesa_por_canal(canal_id: int) -> Optional[Dict[str, Any]]:
    for mesa in carregar_mesas():
        if int(mesa.get("canal_id", 0)) == canal_id:
            return mesa
    return None


class CriarMesaModal(Modal, title="Criar Mesa de Investigação"):
    apelido = TextInput(label="Apelido do agente", placeholder="Ex: Baiano", max_length=50)
    familia = TextInput(label="Nome da família/organização", placeholder="Ex: Olimpo", max_length=80)
    observacao = TextInput(label="Observação inicial", placeholder="Opcional", style=discord.TextStyle.paragraph, required=False, max_length=600)

    async def on_submit(self, interaction: discord.Interaction):
        # Defere rápido para o Discord não mostrar "Esta interação falhou"
        # enquanto o bot cria o canal e envia todos os tópicos.
        await criar_mesa_core(interaction, str(self.apelido.value), str(self.familia.value), str(self.observacao.value or ""))

    async def on_error(self, interaction: discord.Interaction, error: Exception) -> None:
        await enviar_log(f"❌ Erro no modal de criar mesa: {error}")
        try:
            if interaction.response.is_done():
                await interaction.followup.send("❌ Deu erro ao criar a mesa. Veja os logs do bot/Railway.", ephemeral=True)
            else:
                await interaction.response.send_message("❌ Deu erro ao criar a mesa. Veja os logs do bot/Railway.", ephemeral=True)
        except Exception:
            pass


class FecharMesaView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="Fechar Mesa", emoji="🔒", style=discord.ButtonStyle.red, custom_id="dic_fechar_mesa_botao")
    async def fechar(self, interaction: discord.Interaction, button: Button):
        await fechar_mesa_core(interaction, motivo="Fechada pelo botão")


class PainelMesasView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(label="Criar Mesa", emoji="➕", style=discord.ButtonStyle.green, custom_id="dic_criar_mesa")
    async def criar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(CriarMesaModal())

    @discord.ui.button(label="Reabrir Mesa", emoji="🔓", style=discord.ButtonStyle.blurple, custom_id="dic_reabrir_mesa")
    async def reabrir(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_message("Use `/reabrirmesa canal:#canal` para escolher a mesa que será reaberta.", ephemeral=True)

    @discord.ui.button(label="Histórico", emoji="📂", style=discord.ButtonStyle.gray, custom_id="dic_historico_mesas")
    async def historico(self, interaction: discord.Interaction, button: Button):
        await historico_mesa_core(interaction)

    @discord.ui.button(label="Backup", emoji="📦", style=discord.ButtonStyle.gray, custom_id="dic_backup_mesas")
    async def backup(self, interaction: discord.Interaction, button: Button):
        await backup_core(interaction)

    @discord.ui.button(label="Estatísticas", emoji="📊", style=discord.ButtonStyle.gray, custom_id="dic_stats_mesas")
    async def stats(self, interaction: discord.Interaction, button: Button):
        await estatisticas_core(interaction)


    async def on_error(self, interaction: discord.Interaction, error: Exception, item) -> None:
        await enviar_log(f"❌ Erro em botão do painel de mesas: {error}")
        try:
            if interaction.response.is_done():
                await interaction.followup.send("❌ O botão deu erro. Tente usar o comando direto ou veja os logs.", ephemeral=True)
            else:
                await interaction.response.send_message("❌ O botão deu erro. Tente usar o comando direto ou veja os logs.", ephemeral=True)
        except Exception:
            pass


async def responder_interacao(interaction: discord.Interaction, mensagem: str, ephemeral: bool = True):
    if interaction.response.is_done():
        await interaction.followup.send(mensagem, ephemeral=ephemeral)
    else:
        await interaction.response.send_message(mensagem, ephemeral=ephemeral)


async def criar_mesa_core(interaction: discord.Interaction, apelido: str, familia: str, observacao: str = ""):
    guild = interaction.guild
    if guild is None:
        await responder_interacao(interaction, "❌ Use dentro de um servidor.", ephemeral=True)
        return

    if not interaction.response.is_done():
        await interaction.response.defer(ephemeral=True, thinking=True)

    categoria = guild.get_channel(CATEGORIA_MESAS_ABERTAS_ID) if CATEGORIA_MESAS_ABERTAS_ID else None
    overwrites = cargos_equipe_permissoes(guild)
    overwrites[interaction.user] = discord.PermissionOverwrite(
        view_channel=True,
        send_messages=True,
        attach_files=True,
        read_message_history=True,
        send_messages_in_threads=True,
    )

    nome_canal = f"🕵️┃{slugify(apelido)}-{slugify(familia)}"
    canal = await guild.create_text_channel(
        name=nome_canal,
        category=categoria,
        overwrites=overwrites,
    )

    await canal.send(
        f"🕵️ **Mesa de Investigação — {familia}**\n"
        f"👮 **Agente:** {interaction.user.mention}\n"
        f"📛 **Apelido:** {apelido}\n"
        f"🏷️ **Organização/Família:** {familia}\n"
        f"🕒 **Criada em:** {agora_br()}\n"
        f"📝 **Observação:** {observacao or 'Nenhuma'}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🧵 **Use os tópicos abertos abaixo para organizar a investigação.**"
    )

    topicos_ids = await criar_topicos_reais_mesa(canal)
    await canal.send(
        "🔒 Para encerrar esta mesa, clique no botão abaixo.",
        view=FecharMesaView(),
    )

    registrar_mesa({
        "canal_id": canal.id,
        "nome_canal": canal.name,
        "apelido": apelido,
        "familia": familia,
        "autor_id": interaction.user.id,
        "autor_nome": str(interaction.user),
        "status": "ABERTA",
        "criada_em": agora_br(),
        "fechada_em": None,
        "topicos_ids": topicos_ids,
    })

    await enviar_log(
        f"➕ Mesa criada: {canal.mention} | Família: {familia} | Por: {interaction.user.mention}"
    )
    await responder_interacao(interaction, f"✅ Mesa criada: {canal.mention}", ephemeral=True)

async def fechar_mesa_core(interaction: discord.Interaction, motivo: str = "Fechada"):
    canal = interaction.channel
    if not isinstance(canal, discord.TextChannel):
        await interaction.response.send_message("❌ Use dentro do canal da mesa.", ephemeral=True)
        return

    mesa = buscar_mesa_por_canal(canal.id)
    if not mesa:
        await interaction.response.send_message("❌ Esse canal não parece ser uma mesa cadastrada.", ephemeral=True)
        return

    guild = interaction.guild
    categoria_fechada = guild.get_channel(CATEGORIA_MESAS_FECHADAS_ID) if guild and CATEGORIA_MESAS_FECHADAS_ID else None
    try:
        if categoria_fechada:
            await canal.edit(category=categoria_fechada, name=canal.name if canal.name.startswith("🔒") else f"🔒・{canal.name}")
        else:
            await canal.edit(name=canal.name if canal.name.startswith("🔒") else f"🔒・{canal.name}")
    except Exception:
        pass

    mesas = carregar_mesas()
    for m in mesas:
        if int(m.get("canal_id", 0)) == canal.id:
            m["status"] = "FECHADA"
            m["fechada_em"] = agora_br()
            m["motivo_fechamento"] = motivo
    salvar_mesas(mesas)

    await enviar_log(f"🔒 Mesa fechada: {canal.mention} | Por: {interaction.user.mention}")
    await interaction.response.send_message("🔒 Mesa fechada com sucesso.", ephemeral=True)


async def reabrir_mesa_core(interaction: discord.Interaction, canal: discord.TextChannel):
    mesa = buscar_mesa_por_canal(canal.id)
    if not mesa:
        await interaction.response.send_message("❌ Essa mesa não foi encontrada no histórico.", ephemeral=True)
        return
    guild = interaction.guild
    categoria_aberta = guild.get_channel(CATEGORIA_MESAS_ABERTAS_ID) if guild and CATEGORIA_MESAS_ABERTAS_ID else None
    try:
        novo_nome = canal.name.replace("🔒・", "").replace("🔒-", "")
        await canal.edit(category=categoria_aberta, name=novo_nome)
    except Exception:
        pass
    mesas = carregar_mesas()
    for m in mesas:
        if int(m.get("canal_id", 0)) == canal.id:
            m["status"] = "ABERTA"
            m["reaberta_em"] = agora_br()
    salvar_mesas(mesas)
    await enviar_log(f"🔓 Mesa reaberta: {canal.mention} | Por: {interaction.user.mention}")
    await interaction.response.send_message(f"🔓 Mesa reaberta: {canal.mention}", ephemeral=True)


async def historico_mesa_core(interaction: discord.Interaction):
    mesas = carregar_mesas()
    if not mesas:
        await interaction.response.send_message("📂 Nenhuma mesa registrada ainda.", ephemeral=True)
        return
    linhas = []
    for m in mesas[-25:]:
        linhas.append(f"• **{m.get('familia','Sem nome')}** | `{m.get('status')}` | <#{m.get('canal_id')}> | {m.get('criada_em')}")
    await interaction.response.send_message("📂 **Histórico de Mesas:**\n" + "\n".join(linhas), ephemeral=True)


async def backup_core(interaction: discord.Interaction):
    guild = interaction.guild
    if guild is None:
        await interaction.response.send_message("❌ Use dentro de um servidor.", ephemeral=True)
        return
    await interaction.response.defer(ephemeral=True)
    arquivo = BACKUP_DIR / f"backup-{data_caso()}.json"
    dados = {
        "data": agora_br(),
        "servidor": guild.name,
        "procurados": carregar_procurados(),
        "mesas": carregar_mesas(),
        "organizacoes": carregar_organizacoes(),
        "historico_organizacoes": carregar_historico_organizacoes(),
    }
    salvar_json(arquivo, dados)
    canal_backup = bot.get_channel(BACKUP_CHANNEL_ID) if BACKUP_CHANNEL_ID else None
    if canal_backup:
        try:
            await canal_backup.send(f"📦 **Backup Executado**\nServidor: {guild.name}\nData: {agora_br()}", file=discord.File(str(arquivo)))
        except Exception:
            pass
    await enviar_log(f"📦 Backup executado: {arquivo.name}")
    await interaction.followup.send(f"✅ Backup criado: `{arquivo.name}`", ephemeral=True)


async def estatisticas_core(interaction: discord.Interaction):
    procurados = carregar_procurados()
    mesas = carregar_mesas()
    organizacoes = carregar_organizacoes()
    ativos = len([p for p in procurados if p.get("status", "A PROCURAR") == "A PROCURAR"])
    retirados = len([p for p in procurados if p.get("status") == "RETIRADO"])
    abertas = len([m for m in mesas if m.get("status") == "ABERTA"])
    fechadas = len([m for m in mesas if m.get("status") == "FECHADA"])
    organizacoes_editadas = len([
        org for org in organizacoes
        if int(org.get("versao", 1) or 1) > 1
    ])
    await interaction.response.send_message(
        f"📊 **Estatísticas DICOR**\n\n"
        f"🚨 Procurados totais: `{len(procurados)}`\n"
        f"🟢 Procurados ativos: `{ativos}`\n"
        f"🔴 Procurados retirados: `{retirados}`\n\n"
        f"🕵️ Mesas totais: `{len(mesas)}`\n"
        f"🟢 Mesas abertas: `{abertas}`\n"
        f"🔒 Mesas fechadas: `{fechadas}`\n\n"
        f"🏴 Organizações fixas: `{len(organizacoes)}`\n"
        f"✏️ Organizações já editadas: `{organizacoes_editadas}`\n\n"
        f"📄 Catálogo: {CATALOG_PUBLIC_URL}",
        ephemeral=True,
    )




# =====================================================
# BOLETINS DE OCORRENCIA
# =====================================================

def agora_datetime_br() -> datetime.datetime:
    tz = datetime.timezone(datetime.timedelta(hours=-3))
    return datetime.datetime.now(tz)


def data_atual_br() -> str:
    return agora_datetime_br().strftime("%d/%m/%Y")


def horario_atual_br() -> str:
    return agora_datetime_br().strftime("%Hh%M")


def carregar_boletins() -> List[Dict[str, Any]]:
    dados = carregar_json(BOLETINS_JSON, [])
    return dados if isinstance(dados, list) else []


def salvar_boletins(lista: List[Dict[str, Any]]) -> None:
    salvar_json(BOLETINS_JSON, lista)


def carregar_usuarios_rg() -> Dict[str, str]:
    dados = carregar_json(USUARIOS_RG_JSON, {})
    return dados if isinstance(dados, dict) else {}


def salvar_usuarios_rg(dados: Dict[str, str]) -> None:
    salvar_json(USUARIOS_RG_JSON, dados)


def salvar_boletins_pendentes() -> None:
    salvar_json(
        BOLETINS_PENDENTES_JSON,
        {str(canal_id): dados for canal_id, dados in boletins_pendentes.items()},
    )


def carregar_boletins_pendentes_memoria() -> None:
    boletins_pendentes.clear()
    dados = carregar_json(BOLETINS_PENDENTES_JSON, {})
    if not isinstance(dados, dict):
        return
    for canal_id, rascunho in dados.items():
        try:
            boletins_pendentes[int(canal_id)] = rascunho
        except (TypeError, ValueError):
            continue


def usuario_tem_equipe(member: discord.Member) -> bool:
    if member.guild_permissions.administrator:
        return True
    permitidos = set(CARGOS_ADMIN_IDS + CARGOS_EQUIPE_IDS)
    if not permitidos:
        return False
    return any(cargo.id in permitidos for cargo in member.roles)


def obter_rg_usuario(usuario_id: int) -> str:
    return str(carregar_usuarios_rg().get(str(usuario_id), "") or "").strip()


def vincular_rg_usuario(usuario_id: int, rg: str) -> None:
    dados = carregar_usuarios_rg()
    dados[str(usuario_id)] = str(rg).strip()
    salvar_usuarios_rg(dados)


def remover_rg_usuario(usuario_id: int) -> bool:
    dados = carregar_usuarios_rg()
    chave = str(usuario_id)
    if chave not in dados:
        return False
    dados.pop(chave, None)
    salvar_usuarios_rg(dados)
    return True


def gerar_numero_boletim() -> str:
    agora = agora_datetime_br()
    chave_data = agora.strftime("%Y%m%d")
    contador = carregar_json(BOLETINS_CONTADOR_JSON, {})
    if not isinstance(contador, dict):
        contador = {}

    if contador.get("data") == chave_data:
        ultimo = int(contador.get("ultimo", 0) or 0) + 1
    else:
        ultimo = 1

    salvar_json(
        BOLETINS_CONTADOR_JSON,
        {"data": chave_data, "ultimo": ultimo},
    )
    return f"BO-DICOR-{chave_data}-{ultimo:03d}"


def numero_curto_boletim(numero: str) -> str:
    parte = str(numero or "").rsplit("-", 1)[-1]
    return parte.zfill(3) if parte.isdigit() else parte


def buscar_boletim_numero(numero: str) -> Optional[Dict[str, Any]]:
    alvo = str(numero or "").strip().upper()
    for boletim in carregar_boletins():
        if str(boletim.get("numero", "")).strip().upper() == alvo:
            return boletim
    return None


def obter_rascunho_boletim(
    interaction: discord.Interaction,
    exigir_dono: bool = True,
) -> Optional[Dict[str, Any]]:
    canal = interaction.channel
    if not isinstance(canal, discord.TextChannel):
        return None

    dados = boletins_pendentes.get(canal.id)
    if not dados:
        return None

    if exigir_dono and interaction.user.id != int(dados.get("comunicante_id", 0)):
        if not isinstance(interaction.user, discord.Member) or not usuario_tem_equipe(interaction.user):
            return None
    return dados


def dividir_texto_discord(texto: str, limite: int = 1900) -> List[str]:
    texto = str(texto or "").strip()
    if not texto:
        return [""]
    if len(texto) <= limite:
        return [texto]

    partes: List[str] = []
    atual = ""
    for bloco in texto.split("\n"):
        linha = bloco + "\n"
        if len(linha) > limite:
            if atual.strip():
                partes.append(atual.rstrip())
                atual = ""
            restante = linha
            while len(restante) > limite:
                corte = restante.rfind(" ", 0, limite)
                if corte < limite // 2:
                    corte = limite
                partes.append(restante[:corte].rstrip())
                restante = restante[corte:].lstrip()
            atual = restante
            continue

        if len(atual) + len(linha) > limite:
            partes.append(atual.rstrip())
            atual = linha
        else:
            atual += linha

    if atual.strip():
        partes.append(atual.rstrip())
    return partes


def resumir_relato_sem_inventar(relato: str, limite: int = 430) -> str:
    texto = re.sub(r"\s+", " ", str(relato or "")).strip()
    if not texto:
        return ""

    sentencas = re.split(r"(?<=[.!?])\s+", texto)
    selecionadas: List[str] = []
    total = 0
    for sentenca in sentencas:
        sentenca = sentenca.strip()
        if not sentenca:
            continue
        if selecionadas and total + len(sentenca) + 1 > limite:
            break
        if not selecionadas and len(sentenca) > limite:
            selecionadas.append(sentenca[:limite].rstrip(" ,;:") + "...")
            break
        selecionadas.append(sentenca)
        total += len(sentenca) + 1
        if len(selecionadas) >= 2:
            break

    return " ".join(selecionadas).strip()


def gerar_conclusao_automatica(
    relato: str,
    tipo_identificacao: str,
    dados_individuo: Optional[Dict[str, Any]] = None,
    dados_veiculo: Optional[Dict[str, Any]] = None,
) -> str:
    resumo = resumir_relato_sem_inventar(relato)
    tipo = str(tipo_identificacao or "").lower()

    if tipo == "veiculo":
        encerramento = (
            "Os dados do veículo e as provas coletadas foram registrados para "
            "análise e continuidade das diligências investigativas."
        )
    elif tipo == "individuo":
        encerramento = (
            "As informações do indivíduo e as provas coletadas foram registradas "
            "para continuidade das diligências investigativas."
        )
    else:
        encerramento = (
            "Os dados do indivíduo, do veículo e as provas coletadas foram "
            "registrados para análise e continuidade das investigações."
        )

    if resumo:
        return f"Conforme o relato, {resumo[0].lower() + resumo[1:] if len(resumo) > 1 else resumo.lower()} {encerramento}"
    return encerramento


def formatar_secao_individuo(dados: Dict[str, Any]) -> str:
    if not dados:
        return ""
    return (
        "👤 **INDIVÍDUO IDENTIFICADO**\n\n"
        f"**Nome:** {dados.get('nome') or 'Não identificado'}\n"
        f"**RG:** {dados.get('rg') or 'Não identificado'}\n"
        f"**Características:** {dados.get('caracteristicas') or 'Não identificado'}\n"
        f"**Participação no fato:** {dados.get('participacao') or 'Não identificado'}\n"
        f"**Outras informações:** {dados.get('outras_informacoes') or 'Não identificado'}"
    )


def formatar_secao_veiculo(dados: Dict[str, Any]) -> str:
    if not dados:
        return ""
    return (
        "🚘 **VEÍCULO IDENTIFICADO**\n\n"
        f"**Modelo:** {dados.get('modelo') or 'Não identificado'}\n"
        f"**Placa:** {dados.get('placa') or 'Não identificado'}\n"
        f"**Telefone:** {dados.get('telefone') or 'Não identificado'}\n"
        f"**Proprietário:** {dados.get('proprietario') or 'Não identificado'}\n"
        f"**RG do proprietário:** {dados.get('rg_proprietario') or 'Não identificado'}"
    )


def formatar_boletim(dados: Dict[str, Any], previa: bool = False) -> str:
    secoes: List[str] = []
    tipo = str(dados.get("tipo_identificacao", "") or "").lower()

    if tipo in {"individuo", "ambos"}:
        secao = formatar_secao_individuo(dados.get("dados_individuo", {}))
        if secao:
            secoes.append(secao)

    if tipo in {"veiculo", "ambos"}:
        secao = formatar_secao_veiculo(dados.get("dados_veiculo", {}))
        if secao:
            secoes.append(secao)

    identificacoes = "\n\n━━━━━━━━━━━━━━━━━━━━━━━\n\n".join(secoes)
    if identificacoes:
        identificacoes = (
            "\n\n━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            + identificacoes
            + "\n\n━━━━━━━━━━━━━━━━━━━━━━━\n"
        )
    else:
        identificacoes = "\n\n━━━━━━━━━━━━━━━━━━━━━━━\n"

    provas_texto = (
        "As provas encontram-se anexadas abaixo."
        if not previa
        else "As provas enviadas neste canal serão anexadas na publicação."
    )

    numero = str(dados.get("numero", ""))
    comunicante = str(
        dados.get("comunicante_mention")
        or f"<@{dados.get('comunicante_id', 0)}>"
    )

    return (
        f"📋 **BOLETIM DE OCORRÊNCIA — {numero_curto_boletim(numero)}**\n\n"
        f"**Comunicante:** {comunicante}\n"
        f"**RG:** {dados.get('rg') or 'Não informado'}\n\n"
        f"**Data do fato:** {dados.get('data_fato') or data_atual_br()}\n"
        f"**Horário aproximado:** {dados.get('horario_fato') or horario_atual_br()}\n"
        f"**Local:** {dados.get('local') or 'Não informado'}\n\n"
        f"**RELATO:**\n\n"
        f"{dados.get('relato') or 'Não informado'}"
        f"{identificacoes}\n"
        f"📎 **PROVAS:**\n\n"
        f"{provas_texto}\n\n"
        f"**CONCLUSÃO:**\n\n"
        f"{dados.get('conclusao') or 'A conclusão ainda não foi gerada.'}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"**Número do boletim:** {numero}\n"
        f"**Registrado por:** {comunicante}\n"
        f"**Data do registro:** {data_atual_br()} — {horario_atual_br()}"
    )


async def enviar_previa_boletim(
    canal: discord.TextChannel,
    dados: Dict[str, Any],
) -> None:
    texto = "🔎 **PRÉVIA DO BOLETIM**\n\n" + formatar_boletim(dados, previa=True)
    partes = dividir_texto_discord(texto)
    mensagens_ids: List[int] = []
    for indice, parte in enumerate(partes):
        view = PreviaBoletimView() if indice == len(partes) - 1 else None
        mensagem = await canal.send(parte, view=view)
        mensagens_ids.append(mensagem.id)

    dados["preview_message_ids"] = mensagens_ids
    dados["etapa"] = "PREVIA"
    boletins_pendentes[canal.id] = dados
    salvar_boletins_pendentes()


async def enviar_etapa_provas(
    interaction: discord.Interaction,
    dados: Dict[str, Any],
) -> None:
    canal = interaction.channel
    if not isinstance(canal, discord.TextChannel):
        await responder_interacao(interaction, "❌ Canal inválido.", ephemeral=True)
        return

    dados["etapa"] = "PROVAS"
    boletins_pendentes[canal.id] = dados
    salvar_boletins_pendentes()

    mensagem = (
        "📎 **ETAPA DE PROVAS**\n\n"
        "Envie neste canal todas as fotos, vídeos, prints, documentos ou outros "
        "arquivos relacionados ao boletim.\n\n"
        "Quando terminar, clique em **Finalizar Provas**."
    )
    if interaction.response.is_done():
        await canal.send(mensagem, view=ProvasBoletimView())
    else:
        await interaction.response.send_message(
            mensagem,
            view=ProvasBoletimView(),
        )


async def depois_identificacao(
    interaction: discord.Interaction,
    dados: Dict[str, Any],
) -> None:
    canal = interaction.channel
    if not isinstance(canal, discord.TextChannel):
        await responder_interacao(interaction, "❌ Canal inválido.", ephemeral=True)
        return

    retorno = dados.pop("retorno_identificacao", "")
    dados["conclusao"] = gerar_conclusao_automatica(
        dados.get("relato", ""),
        dados.get("tipo_identificacao", ""),
        dados.get("dados_individuo", {}),
        dados.get("dados_veiculo", {}),
    )
    boletins_pendentes[canal.id] = dados
    salvar_boletins_pendentes()

    if retorno == "PREVIA":
        if not interaction.response.is_done():
            await interaction.response.defer(ephemeral=True)
        await enviar_previa_boletim(canal, dados)
        await interaction.followup.send("✅ Identificação atualizada.", ephemeral=True)
        return

    await enviar_etapa_provas(interaction, dados)


async def cancelar_boletim_core(interaction: discord.Interaction) -> None:
    canal = interaction.channel
    if not isinstance(canal, discord.TextChannel):
        await responder_interacao(interaction, "❌ Canal inválido.", ephemeral=True)
        return

    dados = obter_rascunho_boletim(interaction)
    if not dados:
        await responder_interacao(
            interaction,
            "❌ Você não pode cancelar este boletim ou ele não foi encontrado.",
            ephemeral=True,
        )
        return

    boletins_pendentes.pop(canal.id, None)
    salvar_boletins_pendentes()
    await responder_interacao(
        interaction,
        "❌ Boletim cancelado. O canal será apagado em 3 segundos.",
        ephemeral=True,
    )
    await asyncio.sleep(3)
    try:
        await canal.delete(reason="Boletim cancelado")
    except Exception:
        pass


async def coletar_anexos_boletim(
    canal: discord.TextChannel,
) -> List[discord.Attachment]:
    anexos: List[discord.Attachment] = []
    async for mensagem in canal.history(limit=None, oldest_first=True):
        if mensagem.author.bot:
            continue
        anexos.extend(mensagem.attachments)
    return anexos


async def enviar_anexos_boletim(
    canal_oficial: discord.abc.Messageable,
    anexos: List[discord.Attachment],
    numero: str,
) -> List[Dict[str, Any]]:
    resultados: List[Dict[str, Any]] = []

    for indice in range(0, len(anexos), 5):
        lote = anexos[indice:indice + 5]
        arquivos: List[discord.File] = []
        originais: List[discord.Attachment] = []
        for anexo in lote:
            try:
                arquivos.append(await anexo.to_file())
                originais.append(anexo)
            except Exception as erro:
                await enviar_log(
                    f"⚠️ Não foi possível preparar o anexo `{anexo.filename}` "
                    f"do boletim {numero}: {erro}"
                )

        if not arquivos:
            continue

        try:
            mensagem = await canal_oficial.send(
                content=f"📎 **Provas — {numero}**",
                files=arquivos,
            )
            for anexo_publicado in mensagem.attachments:
                resultados.append({
                    "id": anexo_publicado.id,
                    "nome": anexo_publicado.filename,
                    "url": anexo_publicado.url,
                    "mensagem_id": mensagem.id,
                    "mensagem_url": mensagem.jump_url,
                })
        except Exception as erro:
            await enviar_log(
                f"⚠️ Falha ao publicar um lote de provas do boletim {numero}: {erro}"
            )
            # Tenta novamente um arquivo por mensagem.
            for original in originais:
                try:
                    arquivo = await original.to_file()
                    mensagem = await canal_oficial.send(
                        content=f"📎 **Prova — {numero}**",
                        file=arquivo,
                    )
                    for anexo_publicado in mensagem.attachments:
                        resultados.append({
                            "id": anexo_publicado.id,
                            "nome": anexo_publicado.filename,
                            "url": anexo_publicado.url,
                            "mensagem_id": mensagem.id,
                            "mensagem_url": mensagem.jump_url,
                        })
                except Exception as erro_individual:
                    await enviar_log(
                        f"⚠️ Falha ao publicar `{original.filename}` no boletim "
                        f"{numero}: {erro_individual}"
                    )

    return resultados


async def publicar_boletim_core(interaction: discord.Interaction) -> None:
    canal_temp = interaction.channel
    if not isinstance(canal_temp, discord.TextChannel):
        await responder_interacao(interaction, "❌ Canal inválido.", ephemeral=True)
        return

    dados = obter_rascunho_boletim(interaction)
    if not dados:
        await responder_interacao(
            interaction,
            "❌ Boletim não encontrado ou você não tem permissão.",
            ephemeral=True,
        )
        return

    if dados.get("publicando"):
        await responder_interacao(
            interaction,
            "⏳ Este boletim já está sendo publicado.",
            ephemeral=True,
        )
        return

    if buscar_boletim_numero(str(dados.get("numero", ""))):
        await responder_interacao(
            interaction,
            "⚠️ Este boletim já foi publicado.",
            ephemeral=True,
        )
        return

    dados["publicando"] = True
    boletins_pendentes[canal_temp.id] = dados
    salvar_boletins_pendentes()
    await interaction.response.defer(ephemeral=True, thinking=True)

    try:
        canal_oficial = bot.get_channel(BOLETINS_CHANNEL_ID)
        if canal_oficial is None:
            canal_oficial = await bot.fetch_channel(BOLETINS_CHANNEL_ID)
        if not isinstance(canal_oficial, discord.TextChannel):
            raise RuntimeError("O canal oficial de boletins não foi encontrado.")

        partes = dividir_texto_discord(formatar_boletim(dados, previa=False))
        mensagens_publicadas: List[discord.Message] = []
        for parte in partes:
            mensagens_publicadas.append(await canal_oficial.send(parte))

        mensagem_principal = mensagens_publicadas[0]
        anexos_origem = await coletar_anexos_boletim(canal_temp)
        anexos_publicados = await enviar_anexos_boletim(
            canal_oficial,
            anexos_origem,
            str(dados.get("numero", "")),
        )

        registro = dict(dados)
        registro.pop("preview_message_ids", None)
        registro.pop("publicando", None)
        registro["status"] = "PUBLICADO"
        registro["mensagem_id"] = mensagem_principal.id
        registro["mensagem_url"] = mensagem_principal.jump_url
        registro["mensagens_oficiais"] = [
            {"id": mensagem.id, "url": mensagem.jump_url}
            for mensagem in mensagens_publicadas
        ]
        registro["anexos"] = anexos_publicados
        registro["data_criacao"] = agora_br()
        registro["canal_provisorio_id"] = canal_temp.id

        boletins = carregar_boletins()
        if not any(
            str(item.get("numero", "")).upper()
            == str(registro.get("numero", "")).upper()
            for item in boletins
        ):
            boletins.append(registro)
            salvar_boletins(boletins)

        boletins_pendentes.pop(canal_temp.id, None)
        salvar_boletins_pendentes()

        await enviar_log(
            f"📋 Boletim publicado: {registro.get('numero')} | "
            f"Comunicante: <@{registro.get('comunicante_id')}> | "
            f"Mensagem: {mensagem_principal.jump_url}"
        )
        await interaction.followup.send(
            f"✅ Boletim publicado com sucesso: {mensagem_principal.jump_url}",
            ephemeral=True,
        )
        await canal_temp.send("✅ Boletim publicado com sucesso. Este canal será apagado em 5 segundos.")
        await asyncio.sleep(5)
        await canal_temp.delete(reason="Boletim publicado com sucesso")
    except Exception as erro:
        dados["publicando"] = False
        boletins_pendentes[canal_temp.id] = dados
        salvar_boletins_pendentes()
        await enviar_log(f"❌ Erro ao publicar boletim {dados.get('numero')}: {erro}")
        await interaction.followup.send(
            f"❌ Não foi possível publicar o boletim: `{erro}`\n"
            "O canal e o rascunho foram mantidos.",
            ephemeral=True,
        )


async def abrir_boletim_core(interaction: discord.Interaction) -> None:
    guild = interaction.guild
    if guild is None:
        await responder_interacao(interaction, "❌ Use dentro de um servidor.", ephemeral=True)
        return

    if not interaction.response.is_done():
        await interaction.response.defer(ephemeral=True, thinking=True)

    numero = gerar_numero_boletim()
    rg = obter_rg_usuario(interaction.user.id)

    categoria: Optional[discord.CategoryChannel] = None
    if BOLETIM_TEMP_CATEGORY_ID:
        categoria_encontrada = guild.get_channel(BOLETIM_TEMP_CATEGORY_ID)
        if isinstance(categoria_encontrada, discord.CategoryChannel):
            categoria = categoria_encontrada
    if categoria is None and isinstance(interaction.channel, discord.TextChannel):
        categoria = interaction.channel.category

    overwrites: Dict[Any, discord.PermissionOverwrite] = {
        guild.default_role: discord.PermissionOverwrite(view_channel=False),
        interaction.user: discord.PermissionOverwrite(
            view_channel=True,
            send_messages=True,
            attach_files=True,
            read_message_history=True,
        ),
    }
    if guild.me:
        overwrites[guild.me] = discord.PermissionOverwrite(
            view_channel=True,
            send_messages=True,
            attach_files=True,
            read_message_history=True,
            manage_channels=True,
        )
    for cargo_id in set(CARGOS_ADMIN_IDS + CARGOS_EQUIPE_IDS):
        cargo = guild.get_role(cargo_id)
        if cargo:
            overwrites[cargo] = discord.PermissionOverwrite(
                view_channel=True,
                send_messages=True,
                attach_files=True,
                read_message_history=True,
            )

    nome_usuario = slugify(
        getattr(interaction.user, "display_name", interaction.user.name)
    )[:35]
    canal = await guild.create_text_channel(
        name=f"📋・boletim-{nome_usuario}-{numero_curto_boletim(numero)}",
        category=categoria,
        overwrites=overwrites,
        reason=f"Canal provisório do boletim {numero}",
    )

    dados = {
        "numero": numero,
        "comunicante_id": interaction.user.id,
        "comunicante_mention": interaction.user.mention,
        "comunicante_nome": str(interaction.user),
        "rg": rg,
        "data_fato": data_atual_br(),
        "horario_fato": horario_atual_br(),
        "local": "",
        "relato": "",
        "tipo_identificacao": "",
        "dados_individuo": {},
        "dados_veiculo": {},
        "conclusao": "",
        "etapa": "RG" if not rg else "DADOS_INICIAIS",
        "status": "RASCUNHO",
        "criado_em": agora_br(),
        "canal_id": canal.id,
    }
    boletins_pendentes[canal.id] = dados
    salvar_boletins_pendentes()

    await canal.send(
        f"📋 **Boletim de Ocorrência — {numero}**\n\n"
        f"**Comunicante:** {interaction.user.mention}\n"
        f"**RG:** {rg or 'Ainda não vinculado'}\n"
        f"**Data:** {dados['data_fato']}\n"
        f"**Horário:** {dados['horario_fato']}\n\n"
        "O bot vai ajudar no preenchimento do boletim por etapas."
    )

    if rg:
        await canal.send(
            "Clique abaixo para informar o local e escrever o relato.",
            view=IniciarBoletimView(),
        )
    else:
        await canal.send(
            "Seu Discord ainda não possui RG vinculado. Informe o RG para continuar.",
            view=InformarRGBoletimView(),
        )

    await enviar_log(
        f"📁 Canal provisório do boletim {numero} criado por "
        f"{interaction.user.mention}: {canal.mention}"
    )
    await interaction.followup.send(
        f"✅ Canal provisório criado: {canal.mention}",
        ephemeral=True,
    )


def texto_consulta_boletim(boletim: Dict[str, Any]) -> str:
    comunicante = boletim.get("comunicante_mention") or f"<@{boletim.get('comunicante_id', 0)}>"
    return (
        f"📋 **{boletim.get('numero', 'Boletim')}**\n"
        f"**Comunicante:** {comunicante}\n"
        f"**Data:** {boletim.get('data_fato', 'Não informado')}\n"
        f"**Local:** {boletim.get('local', 'Não informado')}\n"
        f"**Status:** {boletim.get('status', 'Não informado')}\n"
        f"**Mensagem:** {boletim.get('mensagem_url', 'Sem link')}"
    )


class InformarRGBoletimModal(Modal, title="Informar RG"):
    rg = TextInput(
        label="RG do comunicante",
        placeholder="Ex: 6027",
        max_length=50,
    )

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message(
                "❌ Rascunho não encontrado ou sem permissão.",
                ephemeral=True,
            )
            return

        valor = str(self.rg.value).strip()
        dados["rg"] = valor
        dados["etapa"] = "DADOS_INICIAIS"
        vincular_rg_usuario(interaction.user.id, valor)
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()

        await interaction.response.send_message(
            f"✅ RG `{valor}` vinculado. Agora preencha o local e o relato.",
            view=IniciarBoletimView(),
        )


class DadosIniciaisBoletimModal(Modal, title="Local e Relato"):
    local = TextInput(
        label="Local do fato",
        placeholder="Ex: Porto, Caixa d'água, Ilha...",
        max_length=200,
    )
    relato = TextInput(
        label="Relato completo",
        placeholder="Descreva toda a ocorrência...",
        style=discord.TextStyle.paragraph,
        max_length=4000,
    )

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message(
                "❌ Rascunho não encontrado ou sem permissão.",
                ephemeral=True,
            )
            return

        dados["local"] = str(self.local.value).strip()
        dados["relato"] = str(self.relato.value).strip()
        dados["etapa"] = "TIPO_IDENTIFICACAO"
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()

        await interaction.response.send_message(
            "✅ Local e relato registrados.\n\n"
            "Agora escolha o que foi identificado:",
            view=TipoIdentificacaoBoletimView(),
        )


class VeiculoBoletimModal(Modal, title="Veículo Identificado"):
    modelo = TextInput(label="Modelo", placeholder="Não identificado", max_length=100)
    placa = TextInput(label="Placa", placeholder="Não identificado", max_length=50)
    telefone = TextInput(label="Telefone", placeholder="Não identificado", max_length=100)
    proprietario = TextInput(label="Proprietário", placeholder="Não identificado", max_length=150)
    rg_proprietario = TextInput(label="RG do proprietário", placeholder="Não identificado", max_length=80)

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message(
                "❌ Rascunho não encontrado ou sem permissão.",
                ephemeral=True,
            )
            return

        dados["dados_veiculo"] = {
            "modelo": str(self.modelo.value).strip(),
            "placa": str(self.placa.value).strip(),
            "telefone": str(self.telefone.value).strip(),
            "proprietario": str(self.proprietario.value).strip(),
            "rg_proprietario": str(self.rg_proprietario.value).strip(),
        }
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await depois_identificacao(interaction, dados)


class IndividuoBoletimModal(Modal, title="Indivíduo Identificado"):
    nome = TextInput(label="Nome", placeholder="Nome do indivíduo", max_length=150)
    rg = TextInput(label="RG", placeholder="Não identificado", max_length=80, required=False)
    caracteristicas = TextInput(
        label="Características",
        placeholder="Não identificado",
        style=discord.TextStyle.paragraph,
        max_length=800,
        required=False,
    )
    participacao = TextInput(
        label="Participação no fato",
        placeholder="Não identificado",
        style=discord.TextStyle.paragraph,
        max_length=800,
        required=False,
    )
    outras_informacoes = TextInput(
        label="Outras informações",
        placeholder="Não identificado",
        style=discord.TextStyle.paragraph,
        max_length=800,
        required=False,
    )

    def __init__(self, continuar_veiculo: bool = False):
        super().__init__()
        self.continuar_veiculo = continuar_veiculo

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message(
                "❌ Rascunho não encontrado ou sem permissão.",
                ephemeral=True,
            )
            return

        dados["dados_individuo"] = {
            "nome": str(self.nome.value).strip(),
            "rg": str(self.rg.value or "").strip(),
            "caracteristicas": str(self.caracteristicas.value or "").strip(),
            "participacao": str(self.participacao.value or "").strip(),
            "outras_informacoes": str(self.outras_informacoes.value or "").strip(),
        }
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()

        if self.continuar_veiculo:
            await interaction.response.send_message(
                "✅ Indivíduo registrado. Agora clique abaixo para preencher o veículo.",
                view=ContinuarVeiculoBoletimView(),
            )
            return

        await depois_identificacao(interaction, dados)


class EditarLocalBoletimModal(Modal, title="Editar Local"):
    local = TextInput(label="Novo local", max_length=200)

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["local"] = str(self.local.value).strip()
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.defer(ephemeral=True)
        await enviar_previa_boletim(interaction.channel, dados)
        await interaction.followup.send("✅ Local atualizado.", ephemeral=True)


class EditarRelatoBoletimModal(Modal, title="Editar Relato"):
    relato = TextInput(
        label="Novo relato",
        style=discord.TextStyle.paragraph,
        max_length=4000,
    )

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["relato"] = str(self.relato.value).strip()
        dados["conclusao"] = gerar_conclusao_automatica(
            dados["relato"],
            dados.get("tipo_identificacao", ""),
            dados.get("dados_individuo", {}),
            dados.get("dados_veiculo", {}),
        )
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.defer(ephemeral=True)
        await enviar_previa_boletim(interaction.channel, dados)
        await interaction.followup.send(
            "✅ Relato atualizado e conclusão refeita.",
            ephemeral=True,
        )


class EditarConclusaoBoletimModal(Modal, title="Editar Conclusão"):
    conclusao = TextInput(
        label="Conclusão",
        style=discord.TextStyle.paragraph,
        max_length=2000,
    )

    async def on_submit(self, interaction: discord.Interaction):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["conclusao"] = str(self.conclusao.value).strip()
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.defer(ephemeral=True)
        await enviar_previa_boletim(interaction.channel, dados)
        await interaction.followup.send("✅ Conclusão atualizada.", ephemeral=True)


class ConsultarBoletimModal(Modal, title="Consultar Boletim"):
    numero = TextInput(
        label="Número do boletim",
        placeholder="BO-DICOR-20260702-001",
        max_length=40,
    )

    async def on_submit(self, interaction: discord.Interaction):
        boletim = buscar_boletim_numero(str(self.numero.value))
        if not boletim:
            await interaction.response.send_message(
                "❌ Boletim não encontrado.",
                ephemeral=True,
            )
            return
        await interaction.response.send_message(
            texto_consulta_boletim(boletim),
            ephemeral=True,
        )


class PainelBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Abrir Boletim",
        emoji="📝",
        style=discord.ButtonStyle.green,
        custom_id="dic_boletim_abrir",
    )
    async def abrir(self, interaction: discord.Interaction, button: Button):
        await abrir_boletim_core(interaction)

    @discord.ui.button(
        label="Meus Boletins",
        emoji="📂",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_boletim_meus",
    )
    async def meus(self, interaction: discord.Interaction, button: Button):
        boletins = [
            boletim
            for boletim in carregar_boletins()
            if int(boletim.get("comunicante_id", 0)) == interaction.user.id
        ][-20:]
        if not boletins:
            await interaction.response.send_message(
                "📂 Você ainda não possui boletins publicados.",
                ephemeral=True,
            )
            return
        linhas = [
            f"• **{b.get('numero')}** | {b.get('data_fato')} | "
            f"{b.get('local')} | {b.get('mensagem_url', 'Sem link')}"
            for b in reversed(boletins)
        ]
        await interaction.response.send_message(
            "📂 **Meus últimos boletins:**\n\n" + "\n".join(linhas),
            ephemeral=True,
        )

    @discord.ui.button(
        label="Consultar Boletim",
        emoji="🔎",
        style=discord.ButtonStyle.gray,
        custom_id="dic_boletim_consultar",
    )
    async def consultar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(ConsultarBoletimModal())

    @discord.ui.button(
        label="Histórico",
        emoji="🗃️",
        style=discord.ButtonStyle.gray,
        custom_id="dic_boletim_historico",
    )
    async def historico(self, interaction: discord.Interaction, button: Button):
        if not isinstance(interaction.user, discord.Member) or not usuario_tem_equipe(interaction.user):
            await interaction.response.send_message(
                "❌ Apenas a equipe autorizada pode consultar o histórico.",
                ephemeral=True,
            )
            return
        boletins = carregar_boletins()[-25:]
        if not boletins:
            await interaction.response.send_message(
                "🗃️ Nenhum boletim publicado.",
                ephemeral=True,
            )
            return
        linhas = []
        for b in reversed(boletins):
            comunicante = b.get("comunicante_mention") or f"<@{b.get('comunicante_id', 0)}>"
            linhas.append(
                f"• **{b.get('numero')}** | {comunicante} | "
                f"{b.get('data_fato')} | {b.get('local')} | "
                f"{b.get('mensagem_url', 'Sem link')}"
            )
        await interaction.response.send_message(
            "🗃️ **Últimos boletins:**\n\n" + "\n".join(linhas),
            ephemeral=True,
        )

    async def on_error(self, interaction: discord.Interaction, error: Exception, item) -> None:
        await enviar_log(f"❌ Erro no painel de boletins: {error}")
        try:
            await responder_interacao(
                interaction,
                "❌ Ocorreu um erro no painel de boletins.",
                ephemeral=True,
            )
        except Exception:
            pass


class InformarRGBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Informar RG",
        emoji="🪪",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_boletim_informar_rg",
    )
    async def informar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(InformarRGBoletimModal())


class IniciarBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Preencher Local e Relato",
        emoji="📝",
        style=discord.ButtonStyle.green,
        custom_id="dic_boletim_dados_iniciais",
    )
    async def iniciar(self, interaction: discord.Interaction, button: Button):
        if not obter_rascunho_boletim(interaction):
            await interaction.response.send_message(
                "❌ Rascunho não encontrado ou sem permissão.",
                ephemeral=True,
            )
            return
        await interaction.response.send_modal(DadosIniciaisBoletimModal())


class TipoIdentificacaoBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Veículo",
        emoji="🚘",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_boletim_tipo_veiculo",
    )
    async def veiculo(self, interaction: discord.Interaction, button: Button):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["tipo_identificacao"] = "veiculo"
        dados["dados_individuo"] = {}
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.send_modal(VeiculoBoletimModal())

    @discord.ui.button(
        label="Indivíduo",
        emoji="👤",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_boletim_tipo_individuo",
    )
    async def individuo(self, interaction: discord.Interaction, button: Button):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["tipo_identificacao"] = "individuo"
        dados["dados_veiculo"] = {}
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.send_modal(IndividuoBoletimModal())

    @discord.ui.button(
        label="Ambos",
        emoji="🔄",
        style=discord.ButtonStyle.green,
        custom_id="dic_boletim_tipo_ambos",
    )
    async def ambos(self, interaction: discord.Interaction, button: Button):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["tipo_identificacao"] = "ambos"
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.send_modal(
            IndividuoBoletimModal(continuar_veiculo=True)
        )


class ContinuarVeiculoBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Preencher Veículo",
        emoji="🚘",
        style=discord.ButtonStyle.green,
        custom_id="dic_boletim_continuar_veiculo",
    )
    async def continuar(self, interaction: discord.Interaction, button: Button):
        if not obter_rascunho_boletim(interaction):
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        await interaction.response.send_modal(VeiculoBoletimModal())


class ProvasBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Finalizar Provas",
        emoji="✅",
        style=discord.ButtonStyle.green,
        custom_id="dic_boletim_finalizar_provas",
    )
    async def finalizar(self, interaction: discord.Interaction, button: Button):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return

        dados["conclusao"] = gerar_conclusao_automatica(
            dados.get("relato", ""),
            dados.get("tipo_identificacao", ""),
            dados.get("dados_individuo", {}),
            dados.get("dados_veiculo", {}),
        )
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()

        await interaction.response.defer(ephemeral=True)
        await enviar_previa_boletim(interaction.channel, dados)
        anexos = await coletar_anexos_boletim(interaction.channel)
        await interaction.followup.send(
            f"✅ Prévia gerada. Provas encontradas: `{len(anexos)}`.",
            ephemeral=True,
        )

    @discord.ui.button(
        label="Voltar",
        emoji="↩️",
        style=discord.ButtonStyle.gray,
        custom_id="dic_boletim_voltar_identificacao",
    )
    async def voltar(self, interaction: discord.Interaction, button: Button):
        if not obter_rascunho_boletim(interaction):
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        await interaction.response.send_message(
            "Escolha novamente o tipo de identificação:",
            view=TipoIdentificacaoBoletimView(),
        )

    @discord.ui.button(
        label="Cancelar Boletim",
        emoji="❌",
        style=discord.ButtonStyle.red,
        custom_id="dic_boletim_cancelar_provas",
    )
    async def cancelar(self, interaction: discord.Interaction, button: Button):
        await cancelar_boletim_core(interaction)


class PreviaBoletimView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Publicar Boletim",
        emoji="✅",
        style=discord.ButtonStyle.green,
        custom_id="dic_boletim_publicar",
        row=0,
    )
    async def publicar(self, interaction: discord.Interaction, button: Button):
        await publicar_boletim_core(interaction)

    @discord.ui.button(
        label="Editar Local",
        emoji="✏️",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_boletim_editar_local",
        row=0,
    )
    async def editar_local(self, interaction: discord.Interaction, button: Button):
        if not obter_rascunho_boletim(interaction):
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        await interaction.response.send_modal(EditarLocalBoletimModal())

    @discord.ui.button(
        label="Editar Relato",
        emoji="✏️",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_boletim_editar_relato",
        row=0,
    )
    async def editar_relato(self, interaction: discord.Interaction, button: Button):
        if not obter_rascunho_boletim(interaction):
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        await interaction.response.send_modal(EditarRelatoBoletimModal())

    @discord.ui.button(
        label="Editar Identificação",
        emoji="✏️",
        style=discord.ButtonStyle.gray,
        custom_id="dic_boletim_editar_identificacao",
        row=1,
    )
    async def editar_identificacao(self, interaction: discord.Interaction, button: Button):
        dados = obter_rascunho_boletim(interaction)
        if not dados:
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        dados["retorno_identificacao"] = "PREVIA"
        boletins_pendentes[interaction.channel.id] = dados
        salvar_boletins_pendentes()
        await interaction.response.send_message(
            "Escolha novamente o tipo de identificação:",
            view=TipoIdentificacaoBoletimView(),
        )

    @discord.ui.button(
        label="Editar Conclusão",
        emoji="✏️",
        style=discord.ButtonStyle.gray,
        custom_id="dic_boletim_editar_conclusao",
        row=1,
    )
    async def editar_conclusao(self, interaction: discord.Interaction, button: Button):
        if not obter_rascunho_boletim(interaction):
            await interaction.response.send_message("❌ Rascunho não encontrado.", ephemeral=True)
            return
        await interaction.response.send_modal(EditarConclusaoBoletimModal())

    @discord.ui.button(
        label="Cancelar",
        emoji="❌",
        style=discord.ButtonStyle.red,
        custom_id="dic_boletim_cancelar_previa",
        row=1,
    )
    async def cancelar(self, interaction: discord.Interaction, button: Button):
        await cancelar_boletim_core(interaction)


def embed_painel_boletim_padrao() -> discord.Embed:
    return discord.Embed(
        title="📋 Sistema de Boletins - DICOR",
        description=(
            "Utilize os botões abaixo para registrar e consultar boletins.\n\n"
            "📝 **Abrir Boletim** — Cria um canal provisório privado.\n"
            "📂 **Meus Boletins** — Mostra seus últimos registros.\n"
            "🔎 **Consultar Boletim** — Consulta pelo número.\n"
            "🗃️ **Histórico** — Exibe os boletins recentes para a equipe."
        ),
        color=discord.Color.blue(),
    )



# =====================================================
# TABELA COMPARTILHADA DE 56 ORGANIZACOES
# =====================================================

TOTAL_ORGANIZACOES = 56
ORGANIZACOES_POR_PAGINA = 8

# Dados importados da planilha compartilhada. Apenas os campos solicitados
# são mantidos no sistema: Nome, Zona de risco, Status, Produto,
# Possui informante, Líder, Características e Histórico operacional.
ORGANIZACOES_INICIAIS: List[Dict[str, Any]] = [{'id': 1,
  'nome': '4 IRMÃOS / 4M',
  'zona_risco': '1',
  'status': 'Já Pacificada',
  'produto': 'H.',
  'informante': 'Não',
  'lider': 'Bruno Diniz 18130',
  'caracteristicas': 'Não informado',
  'historico': 'Invadimos e n tinha ngm'},
 {'id': 2,
  'nome': 'NOVA HOLANDA',
  'zona_risco': '2',
  'status': 'Investigar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Lucas Leal 5766',
  'caracteristicas': 'Os de Vermelho [milan]/Caixa dagua',
  'historico': 'ja incursamos varias vezes, chorão e faz feio.'},
 {'id': 3,
  'nome': 'ZR03',
  'zona_risco': '3',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 4,
  'nome': 'Não informado',
  'zona_risco': '4',
  'status': 'Conhecendo',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'FAVELA DA PRAIA',
  'historico': 'Sem registros.'},
 {'id': 5,
  'nome': 'LOS MALDITOS',
  'zona_risco': '5',
  'status': 'Investigando',
  'produto': 'Lança',
  'informante': 'Não',
  'lider': 'Vinicius Thugnine 28280',
  'caracteristicas': 'Não informado',
  'historico': 'Tel Vini: 305-133'},
 {'id': 6,
  'nome': 'ZR06',
  'zona_risco': '6',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'PLAYBOY',
  'historico': 'Sem registros.'},
 {'id': 7,
  'nome': 'ELIPA',
  'zona_risco': '7',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 8,
  'nome': 'ITALIA',
  'zona_risco': '8',
  'status': 'Conhecendo',
  'produto': 'Ticket | Placa | Nitro',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 9,
  'nome': 'ZR09',
  'zona_risco': '9',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'VINHEDOS',
  'historico': 'Sem registros.'},
 {'id': 10,
  'nome': 'ZR10',
  'zona_risco': '10',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'BAHAMAS',
  'historico': 'Sem registros.'},
 {'id': 11,
  'nome': 'BOLA MAIS UM',
  'zona_risco': '11',
  'status': 'Investigando',
  'produto': 'Balinha',
  'informante': 'Não',
  'lider': 'Caio Henrique 27365',
  'caracteristicas': 'FAZENDINHA',
  'historico': 'Sem registros.'},
 {'id': 12,
  'nome': 'ZR12',
  'zona_risco': '12',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'VANILLA',
  'historico': 'Sem registros.'},
 {'id': 13,
  'nome': 'ZR13',
  'zona_risco': '13',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'ANTIGA BROOCLIN',
  'historico': 'Sem registros.'},
 {'id': 14,
  'nome': 'ROYALT',
  'zona_risco': '14',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Vice Ragnar',
  'caracteristicas': 'Identidade Preto com Rosa. Carros Rosa',
  'historico': 'Sem registros.'},
 {'id': 15,
  'nome': 'MORRO DA SERPENTE',
  'zona_risco': '15',
  'status': 'Conhecendo',
  'produto': 'Armas',
  'informante': 'Não',
  'lider': 'Vulgo Portuga',
  'caracteristicas': 'Identidade "escama" preto com verde.',
  'historico': 'Sem registros.'},
 {'id': 16,
  'nome': 'PORTUGAL',
  'zona_risco': '16',
  'status': 'Conhecendo',
  'produto': 'Armas',
  'informante': 'Não',
  'lider': '02 Vulgo Bibi Perigosa',
  'caracteristicas': 'Identidade banca com dourado.',
  'historico': 'Tel 02: 675-403'},
 {'id': 17,
  'nome': 'ZR17',
  'zona_risco': '17',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'BOATE',
  'historico': 'Sem registros.'},
 {'id': 18,
  'nome': 'ZR18',
  'zona_risco': '18',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'PEGAR INFORMES HISTÓRICO COM CIVIL - PIKACHU DO BOPE OU CONSTANTINO DA CIVIL',
  'historico': 'Sem registros.'},
 {'id': 19,
  'nome': 'BALLAS',
  'zona_risco': '19',
  'status': 'Fora do Radar',
  'produto': 'Desmanche | CAPOFOL',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Informação proviniente de interrogatório durante prisional'},
 {'id': 20,
  'nome': 'FILHOS DA ANARQUIA',
  'zona_risco': '20',
  'status': 'Observando',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Identidade "motociclista" preto com vermelho.',
  'historico': 'Local com suspeita de venda de droga, seguido de fuga para o interior do "bar".'},
 {'id': 21,
  'nome': 'ZR21',
  'zona_risco': '21',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'MANSÃO SCARFACE/PODEROSO CHEFÃO',
  'historico': 'Sem registros.'},
 {'id': 22,
  'nome': 'ZR22',
  'zona_risco': '22',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'MUSEU DE ARTE',
  'historico': 'Sem registros.'},
 {'id': 23,
  'nome': 'ZR23',
  'zona_risco': '23',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'GALAX',
  'historico': 'Sem registros.'},
 {'id': 24,
  'nome': 'DINASTIA',
  'zona_risco': '24',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'HOTEL CALISTO',
  'historico': 'Sem registros.'},
 {'id': 25,
  'nome': 'BLACK DRAGONS',
  'zona_risco': '25',
  'status': 'Conhecendo',
  'produto': 'Skunk',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Identidade Cammo Cinza com Preto',
  'historico': 'Sem registros.'},
 {'id': 26,
  'nome': 'ZR26',
  'zona_risco': '26',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'PETROLÍFERA',
  'historico': 'Sem registros.'},
 {'id': 27,
  'nome': 'ZR27',
  'zona_risco': '27',
  'status': 'Investigar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'LIFE INVADER',
  'historico': 'Sem registros.'},
 {'id': 28,
  'nome': 'MORRO DO MINEIRO',
  'zona_risco': '28',
  'status': 'Observando',
  'produto': 'Desmanche | Lança',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'OBSERVATÓRIO',
  'historico': 'Sem registros.'},
 {'id': 29,
  'nome': 'ZR29',
  'zona_risco': '29',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'SUCATINGA',
  'historico': 'Sem registros.'},
 {'id': 30,
  'nome': 'MEDELIN',
  'zona_risco': '30',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'CEMITÉRIO',
  'historico': 'Sem registros.'},
 {'id': 31,
  'nome': 'ZR31',
  'zona_risco': '31',
  'status': 'Investigar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 32,
  'nome': 'ZR32',
  'zona_risco': '32',
  'status': 'Investigar',
  'produto': 'Desconhecido',
  'informante': 'Sim',
  'lider': 'Não informado',
  'caracteristicas': 'BUNKER DO FAROL',
  'historico': 'INFORMANTE M.B. - DO CORVO'},
 {'id': 33,
  'nome': 'ZR33',
  'zona_risco': '33',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'BOLA DA BENNYS',
  'historico': 'Sem registros.'},
 {'id': 34,
  'nome': 'VERA CRUZ',
  'zona_risco': '34',
  'status': 'Conhecendo',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'JOSE PARAÍBA 29063',
  'caracteristicas': 'ANTIGO CAMPINHO',
  'historico': 'Sem registros.'},
 {'id': 35,
  'nome': 'ZR35',
  'zona_risco': '35',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 36,
  'nome': 'ANONIMOUS',
  'zona_risco': '36',
  'status': 'Observando',
  'produto': 'Attachs',
  'informante': 'Não',
  'lider': 'Evon Zayon',
  'caracteristicas': 'MORRO DO DINO',
  'historico': 'Sem registros.'},
 {'id': 37,
  'nome': 'ZR37',
  'zona_risco': '37',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 38,
  'nome': 'ZR38',
  'zona_risco': '38',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 39,
  'nome': 'ZR39',
  'zona_risco': '39',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'YAKUZA',
  'historico': 'Sem registros.'},
 {'id': 40,
  'nome': 'ZR40',
  'zona_risco': '40',
  'status': 'Investigando',
  'produto': 'Rapé',
  'informante': 'Sim',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': '[INFO JOAQUIM] Elemento entrando na família. Lá tem Rapé.'},
 {'id': 41,
  'nome': 'ZR41',
  'zona_risco': '41',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 42,
  'nome': 'ZR42',
  'zona_risco': '42',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'PROXIMO A RED LINE',
  'historico': 'Sem registros.'},
 {'id': 43,
  'nome': 'ZR43',
  'zona_risco': '43',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 44,
  'nome': 'ZR44',
  'zona_risco': '44',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 45,
  'nome': 'ZR45',
  'zona_risco': '45',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 46,
  'nome': 'ZR46',
  'zona_risco': '46',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'DE FRENTE PRA LIFE',
  'historico': 'Sem registros.'},
 {'id': 47,
  'nome': 'ZR47',
  'zona_risco': '47',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 48,
  'nome': '????',
  'zona_risco': '48',
  'status': 'Observando',
  'produto': 'Desmanche',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'MANICÔMIO',
  'historico': 'Sem registros.'},
 {'id': 49,
  'nome': 'ELEMENTS',
  'zona_risco': '49',
  'status': 'Investigando',
  'produto': 'Master Pick | Flipper mk2',
  'informante': 'Não',
  'lider': 'Tomas Biel [01/gerente]',
  'caracteristicas': 'Não informado',
  'historico': 'Tel do Gerente Tomas: 764-013 // ofereceu masterpick'},
 {'id': 50,
  'nome': 'ZR50',
  'zona_risco': '50',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 51,
  'nome': 'ZR51',
  'zona_risco': '51',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 52,
  'nome': 'ZR52',
  'zona_risco': '52',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 53,
  'nome': 'ZR53',
  'zona_risco': '53',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 54,
  'nome': 'ZR54',
  'zona_risco': '54',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 55,
  'nome': 'ZR55',
  'zona_risco': '55',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'},
 {'id': 56,
  'nome': 'ZR56',
  'zona_risco': '56',
  'status': 'Fora do Radar',
  'produto': 'Desconhecido',
  'informante': 'Não',
  'lider': 'Não informado',
  'caracteristicas': 'Não informado',
  'historico': 'Sem registros.'}]
ORGANIZACOES_INICIAIS_POR_ID = {int(item["id"]): item for item in ORGANIZACOES_INICIAIS}


def modelo_organizacao(numero: int) -> Dict[str, Any]:
    inicial = dict(ORGANIZACOES_INICIAIS_POR_ID.get(numero, {}))
    return {
        "id": numero,
        "nome": str(inicial.get("nome", f"Organização {numero:02d}")),
        "zona_risco": str(inicial.get("zona_risco", numero)),
        "status": str(inicial.get("status", "Não informado")),
        "produto": str(inicial.get("produto", "Desconhecido")),
        "informante": str(inicial.get("informante", "Não")),
        "lider": str(inicial.get("lider", "Não informado")),
        "caracteristicas": str(inicial.get("caracteristicas", "Não informado")),
        "historico": str(inicial.get("historico", "Sem registros.")),
        "versao": 1,
        "ultima_edicao": None,
        "editado_por": None,
        "editado_por_id": None,
        "origem": "Planilha Google",
    }


def registro_organizacao_foi_editado(registro: Dict[str, Any]) -> bool:
    return bool(
        registro.get("editado_por_id")
        or registro.get("ultima_edicao")
        or int(registro.get("versao", 1) or 1) > 1
    )


def normalizar_organizacao(registro: Dict[str, Any], numero: int) -> Dict[str, Any]:
    base = modelo_organizacao(numero)

    # Mantém edições realizadas pelo Discord. Registros antigos que nunca
    # foram editados são atualizados com os dados reais da planilha.
    if isinstance(registro, dict) and registro_organizacao_foi_editado(registro):
        for chave in base:
            if chave in registro:
                base[chave] = registro[chave]

    base["id"] = numero
    try:
        base["versao"] = max(1, int(base.get("versao", 1) or 1))
    except (TypeError, ValueError):
        base["versao"] = 1
    return base


def garantir_56_organizacoes() -> List[Dict[str, Any]]:
    dados = carregar_json(ORGANIZACOES_JSON, [])
    existentes: Dict[int, Dict[str, Any]] = {}

    if isinstance(dados, list):
        for registro in dados:
            if not isinstance(registro, dict):
                continue
            try:
                numero = int(registro.get("id", 0) or 0)
            except (TypeError, ValueError):
                continue
            if 1 <= numero <= TOTAL_ORGANIZACOES and numero not in existentes:
                existentes[numero] = registro

    organizacoes = [
        normalizar_organizacao(existentes.get(numero, {}), numero)
        for numero in range(1, TOTAL_ORGANIZACOES + 1)
    ]
    salvar_json(ORGANIZACOES_JSON, organizacoes)
    return organizacoes


def carregar_organizacoes() -> List[Dict[str, Any]]:
    return garantir_56_organizacoes()


def salvar_organizacoes(lista: List[Dict[str, Any]]) -> None:
    por_id: Dict[int, Dict[str, Any]] = {}
    for registro in lista:
        if not isinstance(registro, dict):
            continue
        try:
            numero = int(registro.get("id", 0) or 0)
        except (TypeError, ValueError):
            continue
        if 1 <= numero <= TOTAL_ORGANIZACOES:
            por_id[numero] = registro

    organizacoes = [
        normalizar_organizacao(por_id.get(numero, {}), numero)
        for numero in range(1, TOTAL_ORGANIZACOES + 1)
    ]
    salvar_json(ORGANIZACOES_JSON, organizacoes)


def carregar_historico_organizacoes() -> List[Dict[str, Any]]:
    dados = carregar_json(HISTORICO_ORGANIZACOES_JSON, [])
    return dados if isinstance(dados, list) else []


def obter_organizacao_por_id(numero: int) -> Optional[Dict[str, Any]]:
    for organizacao in carregar_organizacoes():
        if int(organizacao.get("id", 0) or 0) == numero:
            return organizacao
    return None


def texto_normalizado_busca(texto: str) -> str:
    valor = unicodedata.normalize("NFKD", str(texto or ""))
    valor = valor.encode("ascii", "ignore").decode("ascii").lower().strip()
    return re.sub(r"\s+", " ", valor)


def pesquisar_organizacoes(termo: str) -> List[Dict[str, Any]]:
    termo_limpo = texto_normalizado_busca(termo)
    if not termo_limpo:
        return []

    if termo_limpo.isdigit():
        numero = int(termo_limpo)
        organizacao = obter_organizacao_por_id(numero)
        return [organizacao] if organizacao else []

    organizacoes = carregar_organizacoes()
    exatas = [
        org for org in organizacoes
        if texto_normalizado_busca(org.get("nome", "")) == termo_limpo
    ]
    if exatas:
        return exatas

    return [
        org for org in organizacoes
        if termo_limpo in texto_normalizado_busca(org.get("nome", ""))
    ]


def valor_org(registro: Dict[str, Any], chave: str, padrao: str = "Não informado") -> str:
    valor = str(registro.get(chave, "") or "").strip()
    return valor or padrao


def cortar_campo_org(texto: str, limite: int) -> str:
    valor = str(texto or "").strip()
    if len(valor) <= limite:
        return valor
    return valor[: limite - 3].rstrip() + "..."


def formatar_ficha_organizacao(organizacao: Dict[str, Any]) -> str:
    numero = int(organizacao.get("id", 0) or 0)
    ultima = valor_org(organizacao, "ultima_edicao", "Nunca editada")
    editor = valor_org(organizacao, "editado_por", "Nenhum")

    texto = (
        f"🏴 **ORGANIZAÇÃO {numero:02d} — {valor_org(organizacao, 'nome')}**\n\n"
        f"⚠️ **Zona de risco:** {valor_org(organizacao, 'zona_risco')}\n"
        f"📊 **Status:** {valor_org(organizacao, 'status')}\n"
        f"📦 **Produto:** {valor_org(organizacao, 'produto')}\n"
        f"🕵️ **Possui informante:** {valor_org(organizacao, 'informante')}\n"
        f"👑 **Líder:** {valor_org(organizacao, 'lider')}\n\n"
        f"🔎 **Características:**\n"
        f"{cortar_campo_org(valor_org(organizacao, 'caracteristicas'), 620)}\n\n"
        f"📂 **Histórico operacional:**\n"
        f"{cortar_campo_org(valor_org(organizacao, 'historico', 'Sem registros.'), 620)}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"✏️ **Última edição:** {ultima}\n"
        f"👤 **Editado por:** {editor}\n"
        f"🔢 **Versão:** {int(organizacao.get('versao', 1) or 1)}"
    )
    return cortar_discord(texto, 1900)


def formatar_lista_organizacoes(pagina: int = 0) -> str:
    organizacoes = carregar_organizacoes()
    total_paginas = max(
        1,
        (len(organizacoes) + ORGANIZACOES_POR_PAGINA - 1)
        // ORGANIZACOES_POR_PAGINA,
    )
    pagina = max(0, min(pagina, total_paginas - 1))
    inicio = pagina * ORGANIZACOES_POR_PAGINA
    fim = inicio + ORGANIZACOES_POR_PAGINA

    linhas = [
        "📋 **TABELA DE ORGANIZAÇÕES — DICOR**",
        f"Página `{pagina + 1}/{total_paginas}` • Total: `{TOTAL_ORGANIZACOES}`",
        "",
    ]

    for org in organizacoes[inicio:fim]:
        linhas.append(
            f"**{int(org.get('id', 0)):02d}. {cortar_campo_org(valor_org(org, 'nome'), 35)}**\n"
            f"> ⚠️ {cortar_campo_org(valor_org(org, 'zona_risco'), 18)} | "
            f"📊 {cortar_campo_org(valor_org(org, 'status'), 24)} | "
            f"📦 {cortar_campo_org(valor_org(org, 'produto'), 22)}\n"
            f"> 🕵️ Informante: {cortar_campo_org(valor_org(org, 'informante'), 12)} | "
            f"👑 {cortar_campo_org(valor_org(org, 'lider'), 32)}"
        )

    return cortar_discord("\n".join(linhas), 1900)


def formatar_resultados_organizacoes(resultados: List[Dict[str, Any]]) -> str:
    linhas = ["🔎 **Resultados encontrados:**", ""]
    for org in resultados[:15]:
        linhas.append(
            f"• `{int(org.get('id', 0)):02d}` — **{valor_org(org, 'nome')}** "
            f"| {valor_org(org, 'status')}"
        )
    if len(resultados) > 15:
        linhas.append(f"\n... e mais `{len(resultados) - 15}` resultado(s).")
    linhas.append("\nUse o número exato no botão **Editar Organização**.")
    return cortar_discord("\n".join(linhas), 1900)


async def registrar_edicao_organizacao(
    organizacao_antes: Dict[str, Any],
    organizacao_depois: Dict[str, Any],
    usuario: Any,
    alteracoes: Dict[str, Dict[str, str]],
) -> None:
    historico = carregar_historico_organizacoes()
    historico.append({
        "organizacao_id": int(organizacao_depois.get("id", 0) or 0),
        "nome": valor_org(organizacao_depois, "nome"),
        "data": agora_br(),
        "usuario_id": usuario.id,
        "usuario": str(usuario),
        "versao_anterior": int(organizacao_antes.get("versao", 1) or 1),
        "versao_nova": int(organizacao_depois.get("versao", 1) or 1),
        "alteracoes": alteracoes,
    })
    salvar_json(HISTORICO_ORGANIZACOES_JSON, historico)

    campos = ", ".join(alteracoes.keys())
    await enviar_log(
        f"✏️ **Organização editada**\n"
        f"Organização: `{int(organizacao_depois.get('id', 0)):02d}` — "
        f"{valor_org(organizacao_depois, 'nome')}\n"
        f"Editado por: <@{usuario.id}>\n"
        f"Campos alterados: {campos}\n"
        f"Versão: {int(organizacao_depois.get('versao', 1) or 1)}"
    )


async def aplicar_edicao_organizacao(
    numero: int,
    versao_esperada: int,
    novos_valores: Dict[str, str],
    usuario: Any,
) -> tuple[bool, str, Optional[Dict[str, Any]]]:
    organizacoes = carregar_organizacoes()
    indice = next(
        (
            i for i, org in enumerate(organizacoes)
            if int(org.get("id", 0) or 0) == numero
        ),
        None,
    )
    if indice is None:
        return False, "❌ Organização não encontrada.", None

    atual = organizacoes[indice]
    versao_atual = int(atual.get("versao", 1) or 1)
    if versao_atual != versao_esperada:
        return (
            False,
            "⚠️ Essa organização foi editada por outra pessoa enquanto você preenchia. "
            "Abra a edição novamente para não sobrescrever dados mais recentes.",
            atual,
        )

    antes = dict(atual)
    alteracoes: Dict[str, Dict[str, str]] = {}

    for chave, valor in novos_valores.items():
        if chave not in {
            "nome",
            "zona_risco",
            "status",
            "produto",
            "informante",
            "lider",
            "caracteristicas",
            "historico",
        }:
            continue

        novo = str(valor or "").strip() or "Não informado"
        antigo = str(atual.get(chave, "") or "").strip() or "Não informado"
        if novo != antigo:
            alteracoes[chave] = {"antes": antigo, "depois": novo}
            atual[chave] = novo

    if not alteracoes:
        return False, "ℹ️ Nenhuma informação foi alterada.", atual

    atual["versao"] = versao_atual + 1
    atual["ultima_edicao"] = agora_br()
    atual["editado_por"] = str(usuario)
    atual["editado_por_id"] = usuario.id
    organizacoes[indice] = atual
    salvar_organizacoes(organizacoes)
    await registrar_edicao_organizacao(antes, atual, usuario, alteracoes)
    return True, "✅ Organização atualizada e alteração salva.", atual


class PaginacaoOrganizacoesView(View):
    def __init__(self, pagina: int = 0):
        super().__init__(timeout=300)
        total_paginas = max(
            1,
            (TOTAL_ORGANIZACOES + ORGANIZACOES_POR_PAGINA - 1)
            // ORGANIZACOES_POR_PAGINA,
        )
        self.pagina = max(0, min(pagina, total_paginas - 1))
        self.total_paginas = total_paginas
        for item in self.children:
            if isinstance(item, Button) and item.label == "Anterior":
                item.disabled = self.pagina <= 0
            elif isinstance(item, Button) and item.label == "Próxima":
                item.disabled = self.pagina >= self.total_paginas - 1

    @discord.ui.button(label="Anterior", emoji="⬅️", style=discord.ButtonStyle.gray)
    async def anterior(self, interaction: discord.Interaction, button: Button):
        nova_pagina = max(0, self.pagina - 1)
        await interaction.response.edit_message(
            content=formatar_lista_organizacoes(nova_pagina),
            view=PaginacaoOrganizacoesView(nova_pagina),
        )

    @discord.ui.button(label="Próxima", emoji="➡️", style=discord.ButtonStyle.gray)
    async def proxima(self, interaction: discord.Interaction, button: Button):
        nova_pagina = min(self.total_paginas - 1, self.pagina + 1)
        await interaction.response.edit_message(
            content=formatar_lista_organizacoes(nova_pagina),
            view=PaginacaoOrganizacoesView(nova_pagina),
        )


class EditarOrganizacaoBasicoModal(Modal):
    def __init__(self, organizacao: Dict[str, Any]):
        numero = int(organizacao.get("id", 0) or 0)
        super().__init__(title=f"Editar organização {numero:02d}")
        self.numero = numero
        self.versao = int(organizacao.get("versao", 1) or 1)

        self.nome = TextInput(
            label="Nome",
            default=valor_org(organizacao, "nome"),
            max_length=100,
        )
        self.zona_risco = TextInput(
            label="Zona de risco",
            default=valor_org(organizacao, "zona_risco"),
            placeholder="Ex: Baixa, Média, Alta ou Crítica",
            max_length=80,
        )
        self.status = TextInput(
            label="Status",
            default=valor_org(organizacao, "status"),
            placeholder="Ex: Ativa, Monitoramento, Em investigação",
            max_length=100,
        )
        self.produto = TextInput(
            label="Produto",
            default=valor_org(organizacao, "produto"),
            max_length=150,
        )
        self.informante = TextInput(
            label="Possui informante?",
            default=valor_org(organizacao, "informante"),
            placeholder="Sim, Não ou Não identificado",
            max_length=60,
        )

        for item in (
            self.nome,
            self.zona_risco,
            self.status,
            self.produto,
            self.informante,
        ):
            self.add_item(item)

    async def on_submit(self, interaction: discord.Interaction):
        sucesso, mensagem, organizacao = await aplicar_edicao_organizacao(
            self.numero,
            self.versao,
            {
                "nome": str(self.nome.value),
                "zona_risco": str(self.zona_risco.value),
                "status": str(self.status.value),
                "produto": str(self.produto.value),
                "informante": str(self.informante.value),
            },
            interaction.user,
        )
        conteudo = mensagem
        view = None
        if organizacao:
            conteudo += "\n\n" + formatar_ficha_organizacao(organizacao)
            view = OrganizacaoAcoesView(
                int(organizacao.get("id", 0)),
                int(organizacao.get("versao", 1) or 1),
            )
        await interaction.response.send_message(
            cortar_discord(conteudo, 1900),
            view=view,
            ephemeral=True,
        )


class EditarOrganizacaoDetalhesModal(Modal):
    def __init__(self, organizacao: Dict[str, Any]):
        numero = int(organizacao.get("id", 0) or 0)
        super().__init__(title=f"Detalhes da organização {numero:02d}")
        self.numero = numero
        self.versao = int(organizacao.get("versao", 1) or 1)

        self.lider = TextInput(
            label="Líder",
            default=valor_org(organizacao, "lider"),
            max_length=150,
        )
        self.caracteristicas = TextInput(
            label="Características",
            default=valor_org(organizacao, "caracteristicas"),
            style=discord.TextStyle.paragraph,
            max_length=1800,
        )
        self.historico = TextInput(
            label="Histórico operacional",
            default=valor_org(organizacao, "historico", "Sem registros."),
            placeholder="Registre ocorrências, datas e informações relevantes.",
            style=discord.TextStyle.paragraph,
            max_length=3000,
        )

        self.add_item(self.lider)
        self.add_item(self.caracteristicas)
        self.add_item(self.historico)

    async def on_submit(self, interaction: discord.Interaction):
        sucesso, mensagem, organizacao = await aplicar_edicao_organizacao(
            self.numero,
            self.versao,
            {
                "lider": str(self.lider.value),
                "caracteristicas": str(self.caracteristicas.value),
                "historico": str(self.historico.value),
            },
            interaction.user,
        )
        conteudo = mensagem
        view = None
        if organizacao:
            conteudo += "\n\n" + formatar_ficha_organizacao(organizacao)
            view = OrganizacaoAcoesView(
                int(organizacao.get("id", 0)),
                int(organizacao.get("versao", 1) or 1),
            )
        await interaction.response.send_message(
            cortar_discord(conteudo, 1900),
            view=view,
            ephemeral=True,
        )


class OrganizacaoAcoesView(View):
    def __init__(self, numero: int, versao: int):
        super().__init__(timeout=300)
        self.numero = numero
        self.versao = versao

    @discord.ui.button(label="Editar Dados", emoji="✏️", style=discord.ButtonStyle.blurple)
    async def editar_dados(self, interaction: discord.Interaction, button: Button):
        organizacao = obter_organizacao_por_id(self.numero)
        if not organizacao:
            await interaction.response.send_message("❌ Organização não encontrada.", ephemeral=True)
            return
        await interaction.response.send_modal(EditarOrganizacaoBasicoModal(organizacao))

    @discord.ui.button(label="Editar Detalhes", emoji="📝", style=discord.ButtonStyle.green)
    async def editar_detalhes(self, interaction: discord.Interaction, button: Button):
        organizacao = obter_organizacao_por_id(self.numero)
        if not organizacao:
            await interaction.response.send_message("❌ Organização não encontrada.", ephemeral=True)
            return
        await interaction.response.send_modal(EditarOrganizacaoDetalhesModal(organizacao))

    @discord.ui.button(label="Atualizar Ficha", emoji="🔄", style=discord.ButtonStyle.gray)
    async def atualizar(self, interaction: discord.Interaction, button: Button):
        organizacao = obter_organizacao_por_id(self.numero)
        if not organizacao:
            await interaction.response.send_message("❌ Organização não encontrada.", ephemeral=True)
            return
        await interaction.response.edit_message(
            content=formatar_ficha_organizacao(organizacao),
            view=OrganizacaoAcoesView(
                self.numero,
                int(organizacao.get("versao", 1) or 1),
            ),
        )


class PesquisarOrganizacaoModal(Modal, title="Pesquisar Organização"):
    termo = TextInput(
        label="Nome ou número",
        placeholder="Ex: 12 ou Medelim",
        max_length=100,
    )

    async def on_submit(self, interaction: discord.Interaction):
        resultados = pesquisar_organizacoes(str(self.termo.value))
        if not resultados:
            await interaction.response.send_message(
                "❌ Nenhuma organização encontrada.",
                ephemeral=True,
            )
            return

        if len(resultados) == 1:
            organizacao = resultados[0]
            await interaction.response.send_message(
                formatar_ficha_organizacao(organizacao),
                view=OrganizacaoAcoesView(
                    int(organizacao.get("id", 0)),
                    int(organizacao.get("versao", 1) or 1),
                ),
                ephemeral=True,
            )
            return

        await interaction.response.send_message(
            formatar_resultados_organizacoes(resultados),
            ephemeral=True,
        )


class SelecionarOrganizacaoEditarModal(Modal, title="Editar Organização"):
    termo = TextInput(
        label="Nome ou número",
        placeholder="Ex: 12 ou Medelim",
        max_length=100,
    )

    async def on_submit(self, interaction: discord.Interaction):
        resultados = pesquisar_organizacoes(str(self.termo.value))
        if not resultados:
            await interaction.response.send_message(
                "❌ Nenhuma organização encontrada.",
                ephemeral=True,
            )
            return

        if len(resultados) > 1:
            await interaction.response.send_message(
                formatar_resultados_organizacoes(resultados),
                ephemeral=True,
            )
            return

        organizacao = resultados[0]
        await interaction.response.send_message(
            formatar_ficha_organizacao(organizacao),
            view=OrganizacaoAcoesView(
                int(organizacao.get("id", 0)),
                int(organizacao.get("versao", 1) or 1),
            ),
            ephemeral=True,
        )


class PainelOrganizacoesView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Ver Organizações",
        emoji="📋",
        style=discord.ButtonStyle.blurple,
        custom_id="dic_org_ver_organizacoes",
    )
    async def ver_organizacoes(self, interaction: discord.Interaction, button: Button):
        garantir_56_organizacoes()
        await interaction.response.send_message(
            formatar_lista_organizacoes(0),
            view=PaginacaoOrganizacoesView(0),
            ephemeral=True,
        )

    @discord.ui.button(
        label="Pesquisar Organização",
        emoji="🔎",
        style=discord.ButtonStyle.gray,
        custom_id="dic_org_pesquisar",
    )
    async def pesquisar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(PesquisarOrganizacaoModal())

    @discord.ui.button(
        label="Editar Organização",
        emoji="✏️",
        style=discord.ButtonStyle.green,
        custom_id="dic_org_editar",
    )
    async def editar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(SelecionarOrganizacaoEditarModal())

    async def on_error(self, interaction: discord.Interaction, error: Exception, item) -> None:
        await enviar_log(f"❌ Erro no painel de organizações: {error}")
        try:
            if interaction.response.is_done():
                await interaction.followup.send(
                    "❌ Ocorreu um erro no sistema de organizações.",
                    ephemeral=True,
                )
            else:
                await interaction.response.send_message(
                    "❌ Ocorreu um erro no sistema de organizações.",
                    ephemeral=True,
                )
        except Exception:
            pass


def embed_painel_organizacoes_padrao() -> discord.Embed:
    return discord.Embed(
        title="🏴 Tabela de Organizações - DICOR",
        description=(
            "As **56 organizações** são fixas e compartilhadas. "
            "Todos podem consultar e editar qualquer ficha.\n\n"
            "📋 **Ver Organizações** — Abre a tabela completa por páginas.\n"
            "🔎 **Pesquisar Organização** — Busca pelo nome ou número.\n"
            "✏️ **Editar Organização** — Atualiza os dados da ficha.\n\n"
            "🔒 Todas as alterações ficam salvas com usuário, data, "
            "valores anteriores e valores novos. Nenhuma organização pode ser apagada."
        ),
        color=discord.Color.dark_blue(),
    )

# =====================================================
# COMANDOS DE TEXTO (!) - PLANO B PARA QUANDO O DISCORD NÃO MOSTRAR SLASH
# =====================================================

def embed_painel_procurados_padrao() -> discord.Embed:
    return discord.Embed(
        title="🚨 Sistema de Procurados - DICOR",
        description=(
            "Utilize os botões abaixo para gerenciar procurados.\n\n"
            "➕ **Novo Procurado** — Cadastrar um novo procurado.\n"
            "📋 **Lista de Procurados** — Ver procurados ativos.\n"
            "❌ **Retirar Procurado** — Retirar um procurado pelo RG.\n"
            "📄 **Abrir Catálogo** — Abrir o catálogo HTML.\n"
            "🔄 **Sincronizar Antigos** — Importar procurados antigos do canal oficial."
        ),
        color=discord.Color.red(),
    )


def embed_painel_mesas_padrao() -> discord.Embed:
    return discord.Embed(
        title="🕵️ Sistema de Mesas - DICOR",
        description=(
            "Utilize os botões abaixo para gerenciar as mesas.\n\n"
            "➕ **Criar Mesa** — Abre uma nova mesa de investigação.\n"
            "🔓 **Reabrir Mesa** — Reabre uma mesa fechada.\n"
            "📂 **Histórico** — Mostra mesas registradas.\n"
            "📦 **Backup** — Gera backup dos dados.\n"
            "📊 **Estatísticas** — Mostra números do sistema."
        ),
        color=discord.Color.gold(),
    )


async def criar_mesa_por_texto(ctx: commands.Context, apelido: str, familia: str, observacao: str = ""):
    guild = ctx.guild
    if guild is None:
        await ctx.reply("❌ Use dentro de um servidor.")
        return

    categoria = guild.get_channel(CATEGORIA_MESAS_ABERTAS_ID) if CATEGORIA_MESAS_ABERTAS_ID else None
    overwrites = cargos_equipe_permissoes(guild)
    overwrites[ctx.author] = discord.PermissionOverwrite(
        view_channel=True,
        send_messages=True,
        attach_files=True,
        read_message_history=True,
        send_messages_in_threads=True,
    )

    nome_canal = f"🕵️┃{slugify(apelido)}-{slugify(familia)}"
    canal = await guild.create_text_channel(
        name=nome_canal,
        category=categoria,
        overwrites=overwrites,
    )

    await canal.send(
        f"🕵️ **Mesa de Investigação — {familia}**\n"
        f"👮 **Agente:** {ctx.author.mention}\n"
        f"📛 **Apelido:** {apelido}\n"
        f"🏷️ **Organização/Família:** {familia}\n"
        f"🕒 **Criada em:** {agora_br()}\n"
        f"📝 **Observação:** {observacao or 'Nenhuma'}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🧵 **Use os tópicos abertos abaixo para organizar a investigação.**"
    )

    topicos_ids = await criar_topicos_reais_mesa(canal)
    await canal.send(
        "🔒 Para encerrar esta mesa, clique no botão abaixo.",
        view=FecharMesaView(),
    )

    registrar_mesa({
        "canal_id": canal.id,
        "nome_canal": canal.name,
        "apelido": apelido,
        "familia": familia,
        "autor_id": ctx.author.id,
        "autor_nome": str(ctx.author),
        "status": "ABERTA",
        "criada_em": agora_br(),
        "fechada_em": None,
        "topicos_ids": topicos_ids,
    })

    await enviar_log(
        f"➕ Mesa criada por comando de texto: {canal.mention} | Família: {familia} | Por: {ctx.author.mention}"
    )
    await ctx.reply(f"✅ Mesa criada: {canal.mention}")

async def fechar_mesa_por_texto(ctx: commands.Context, motivo: str = "Fechada por comando de texto"):
    canal = ctx.channel
    if not isinstance(canal, discord.TextChannel):
        await ctx.reply("❌ Use dentro do canal da mesa.")
        return

    mesa = buscar_mesa_por_canal(canal.id)
    if not mesa:
        await ctx.reply("❌ Esse canal não parece ser uma mesa cadastrada.")
        return

    guild = ctx.guild
    categoria_fechada = guild.get_channel(CATEGORIA_MESAS_FECHADAS_ID) if guild and CATEGORIA_MESAS_FECHADAS_ID else None
    try:
        novo_nome = canal.name if canal.name.startswith("🔒") else f"🔒・{canal.name}"
        if categoria_fechada:
            await canal.edit(category=categoria_fechada, name=novo_nome)
        else:
            await canal.edit(name=novo_nome)
    except Exception:
        pass

    mesas = carregar_mesas()
    for m in mesas:
        if int(m.get("canal_id", 0)) == canal.id:
            m["status"] = "FECHADA"
            m["fechada_em"] = agora_br()
            m["motivo_fechamento"] = motivo
    salvar_mesas(mesas)

    await enviar_log(f"🔒 Mesa fechada por comando de texto: {canal.mention} | Motivo: {motivo}")
    await ctx.reply("🔒 Mesa fechada e movida para a categoria de mesas fechadas.")



@bot.command(name="painelorganizacoes")
async def cmd_painelorganizacoes(ctx: commands.Context):
    """Envia o painel compartilhado das 56 organizações."""
    garantir_56_organizacoes()
    await ctx.send(
        embed=embed_painel_organizacoes_padrao(),
        view=PainelOrganizacoesView(),
    )


@bot.command(name="painelboletim")
async def cmd_painelboletim(ctx: commands.Context):
    """Abre o painel de boletins pelo comando !painelboletim."""
    if not isinstance(ctx.author, discord.Member) or not usuario_tem_equipe(ctx.author):
        await ctx.reply("❌ Apenas a equipe autorizada pode enviar este painel.")
        return
    await ctx.send(embed=embed_painel_boletim_padrao(), view=PainelBoletimView())


@bot.command(name="consultarboletim")
async def cmd_consultarboletim(ctx: commands.Context, *, numero: str = ""):
    """Consulta um boletim pelo número."""
    if not numero.strip():
        await ctx.reply("❌ Use assim: `!consultarboletim BO-DICOR-20260702-001`")
        return
    boletim = buscar_boletim_numero(numero)
    if not boletim:
        await ctx.reply("❌ Boletim não encontrado.")
        return
    await ctx.reply(texto_consulta_boletim(boletim))


@bot.command(name="painelmesas")
@commands.has_permissions(administrator=True)
async def cmd_painelmesas(ctx: commands.Context):
    """Abre o painel de mesas pelo comando !painelmesas."""
    await ctx.send(embed=embed_painel_mesas_padrao(), view=PainelMesasView())


@bot.command(name="painelprocurados")
@commands.has_permissions(administrator=True)
async def cmd_painelprocurados(ctx: commands.Context):
    """Abre o painel de procurados pelo comando !painelprocurados."""
    await ctx.send(embed=embed_painel_procurados_padrao(), view=PainelProcuradosView())


@bot.command(name="criarmesa")
@commands.has_permissions(administrator=True)
async def cmd_criarmesa(ctx: commands.Context, *, texto: str = ""):
    """Cria mesa direto por texto. Ex: !criarmesa Baiano | Olimpo | Observação"""
    partes = [p.strip() for p in texto.split("|")]
    if len(partes) < 2 or not partes[0] or not partes[1]:
        await ctx.reply("❌ Use assim: `!criarmesa Apelido | Família | Observação opcional`\nExemplo: `!criarmesa Baiano | Olimpo | Mesa inicial`")
        return
    apelido = partes[0]
    familia = partes[1]
    observacao = partes[2] if len(partes) >= 3 else ""
    await criar_mesa_por_texto(ctx, apelido, familia, observacao)


@bot.command(name="fecharmesa")
@commands.has_permissions(administrator=True)
async def cmd_fecharmesa(ctx: commands.Context, *, motivo: str = "Fechada por comando de texto"):
    await fechar_mesa_por_texto(ctx, motivo)


@bot.command(name="catalogo")
async def cmd_catalogo(ctx: commands.Context):
    await ctx.reply(f"📄 **Catálogo de Procurados:**\n{CATALOG_PUBLIC_URL}")


@bot.command(name="comandos")
async def cmd_comandos(ctx: commands.Context):
    await ctx.reply(
        "📌 **Comandos disponíveis:**\n\n"
        "**Slash:** `/painelprocurados`, `/painelmesas`, `/painelboletim`, `/painelorganizacoes`, `/criarmesa`, `/fecharmesa`, `/backup`, `/estatisticas`, `/consultarboletim`\n"
        "**Plano B por texto:** `!painelprocurados`, `!painelmesas`, `!painelboletim`, `!painelorganizacoes`, `!criarmesa Apelido | Família | Observação`, `!fecharmesa`, `!catalogo`, `!consultarboletim NUMERO`"
    )


@bot.command(name="synccomandos")
@commands.has_permissions(administrator=True)
async def cmd_synccomandos(ctx: commands.Context):
    """Força sincronização dos slash commands no servidor atual."""
    if ctx.guild is None:
        await ctx.reply("Use dentro do servidor.")
        return
    guild_obj = discord.Object(id=ctx.guild.id)
    try:
        bot.tree.copy_global_to(guild=guild_obj)
        comandos = await bot.tree.sync(guild=guild_obj)
        await ctx.reply(f"✅ Slash commands sincronizados neste servidor: `{len(comandos)}`. Agora aperte `Ctrl + R` no Discord.")
    except Exception as e:
        await ctx.reply(f"❌ Erro ao sincronizar: `{e}`")


# =====================================================
# SLASH COMMANDS - TODOS
# =====================================================


@bot.tree.command(name="painelorganizacoes", description="Envia a tabela compartilhada das 56 organizações.")
async def painelorganizacoes(interaction: discord.Interaction):
    garantir_56_organizacoes()
    await interaction.response.send_message(
        embed=embed_painel_organizacoes_padrao(),
        view=PainelOrganizacoesView(),
    )


@bot.tree.command(name="painelboletim", description="Envia o painel de boletins de ocorrência.")
async def painelboletim(interaction: discord.Interaction):
    if not isinstance(interaction.user, discord.Member) or not usuario_tem_equipe(interaction.user):
        await interaction.response.send_message(
            "❌ Apenas a equipe autorizada pode enviar este painel.",
            ephemeral=True,
        )
        return
    await interaction.response.send_message(
        embed=embed_painel_boletim_padrao(),
        view=PainelBoletimView(),
    )


@bot.tree.command(name="vincularrg", description="Vincula um RG a um usuário do Discord.")
@app_commands.describe(usuario="Usuário que receberá o RG", rg="RG do usuário")
@app_commands.checks.has_permissions(administrator=True)
async def vincularrg(
    interaction: discord.Interaction,
    usuario: discord.Member,
    rg: str,
):
    vincular_rg_usuario(usuario.id, rg)
    await interaction.response.send_message(
        f"✅ RG `{rg}` vinculado a {usuario.mention}.",
        ephemeral=True,
    )


@bot.tree.command(name="removervinculorg", description="Remove o RG vinculado de um usuário.")
@app_commands.describe(usuario="Usuário que terá o RG removido")
@app_commands.checks.has_permissions(administrator=True)
async def removervinculorg(
    interaction: discord.Interaction,
    usuario: discord.Member,
):
    removido = remover_rg_usuario(usuario.id)
    mensagem = (
        f"✅ RG de {usuario.mention} removido."
        if removido
        else f"⚠️ {usuario.mention} não possuía RG vinculado."
    )
    await interaction.response.send_message(mensagem, ephemeral=True)


@bot.tree.command(name="consultarboletim", description="Consulta um boletim pelo número.")
@app_commands.describe(numero="Ex: BO-DICOR-20260702-001")
async def consultarboletim(
    interaction: discord.Interaction,
    numero: str,
):
    boletim = buscar_boletim_numero(numero)
    if not boletim:
        await interaction.response.send_message(
            "❌ Boletim não encontrado.",
            ephemeral=True,
        )
        return
    await interaction.response.send_message(
        texto_consulta_boletim(boletim),
        ephemeral=True,
    )


@bot.tree.command(name="painelprocurados", description="Envia o painel de gerenciamento de procurados.")
@app_commands.checks.has_permissions(administrator=True)
async def painelprocurados(interaction: discord.Interaction):
    await interaction.response.send_message(embed=embed_painel_procurados_padrao(), view=PainelProcuradosView())


@bot.tree.command(name="painelmesas", description="Envia o painel de gerenciamento de mesas.")
@app_commands.checks.has_permissions(administrator=True)
async def painelmesas(interaction: discord.Interaction):
    await interaction.response.send_message(embed=embed_painel_mesas_padrao(), view=PainelMesasView())


@bot.tree.command(name="catalogo", description="Mostra o link do catálogo de procurados.")
async def catalogo(interaction: discord.Interaction):
    await interaction.response.send_message(f"📄 **Catálogo de Procurados:**\n{CATALOG_PUBLIC_URL}", ephemeral=True)


@bot.tree.command(name="retirarprocurado", description="Retira um procurado pelo RG.")
@app_commands.describe(rg="RG do procurado", motivo="Motivo da retirada")
async def retirarprocurado(interaction: discord.Interaction, rg: str, motivo: str = "Não informado"):
    await retirar_procurado(interaction, rg, motivo)


@bot.tree.command(name="sincronizarcatalogo", description="Importa procurados antigos do canal oficial para o catálogo.")
@app_commands.checks.has_permissions(administrator=True)
async def sincronizarcatalogo(interaction: discord.Interaction):
    await sincronizar_catalogo_core(interaction)


@bot.tree.command(name="regerarcatalogo", description="Recria o HTML do catálogo.")
@app_commands.checks.has_permissions(administrator=True)
async def regerarcatalogo(interaction: discord.Interaction):
    gerar_catalogo_html()
    await interaction.response.send_message(f"✅ Catálogo recriado: {CATALOG_PUBLIC_URL}", ephemeral=True)


@bot.tree.command(name="criarmesa", description="Cria uma mesa de investigação.")
@app_commands.describe(apelido="Apelido do agente", familia="Nome da família/organização", observacao="Observação opcional")
async def criarmesa(interaction: discord.Interaction, apelido: str, familia: str, observacao: str = ""):
    await criar_mesa_core(interaction, apelido, familia, observacao)


@bot.tree.command(name="fecharmesa", description="Fecha a mesa do canal atual.")
@app_commands.describe(motivo="Motivo do fechamento")
async def fecharmesa(interaction: discord.Interaction, motivo: str = "Fechada por comando"):
    await fechar_mesa_core(interaction, motivo)


@bot.tree.command(name="reabrirmesa", description="Reabre uma mesa fechada.")
@app_commands.describe(canal="Canal da mesa")
async def reabrirmesa(interaction: discord.Interaction, canal: discord.TextChannel):
    await reabrir_mesa_core(interaction, canal)


@bot.tree.command(name="historicomesa", description="Mostra histórico das mesas.")
async def historicomesa(interaction: discord.Interaction):
    await historico_mesa_core(interaction)


@bot.tree.command(name="backup", description="Gera backup dos dados do bot.")
@app_commands.checks.has_permissions(administrator=True)
async def backup(interaction: discord.Interaction):
    await backup_core(interaction)


@bot.tree.command(name="estatisticas", description="Mostra estatísticas do sistema.")
async def estatisticas(interaction: discord.Interaction):
    await estatisticas_core(interaction)

# =====================================================
# EVENTOS
# =====================================================

comandos_ja_sincronizados = False

@bot.event
async def on_ready():
    global comandos_ja_sincronizados
    
    # 1. ATIVA A PERSISTÊNCIA DOS BOTÕES DE TODOS OS PAINÉIS
    try:
        # Painel Novo de Relatórios
        bot.add_view(RelatoriosPainelView())
        
        # Painel Antigo de Boletins (Para resolver o "Esta interação falhou")
        bot.add_view(PainelBoletimView())
        bot.add_view(IniciarBoletimView())
        bot.add_view(InformarRGBoletimView())
        bot.add_view(TipoIdentificacaoBoletimView())
        bot.add_view(ContinuarVeiculoBoletimView())
        bot.add_view(ProvasBoletimView())
        bot.add_view(PreviaBoletimView())
        
        # Outros Painéis do Sistema
        bot.add_view(PainelProcuradosView())
        bot.add_view(FinalizarProcuradoView())
        bot.add_view(PainelMesasView())
        bot.add_view(FecharMesaView())
        bot.add_view(PainelOrganizacoesView())
        
        print("✅ Todas as persistências de botões (Relatórios, Boletins e Sistemas) foram carregadas!")
    except Exception as e:
        print(f"Aviso ao carregar persistência dos painéis: {e}")

    # Lógica de sincronização do bot
    try:
        if GUILD_ID > 0:
            guild = discord.Object(id=GUILD_ID)
            bot.tree.copy_global_to(guild=guild)
            comandos_servidor = await bot.tree.sync(guild=guild)
            
            bot.tree.clear_commands(guild=None)
            comandos_globais = await bot.tree.sync()

            print(f"Comandos sincronizados no servidor: {len(comandos_servidor)}")
            print(f"Comandos globais antigos limpos: {len(comandos_globais)}")
            print("Comandos ativos:", ", ".join(f"/{cmd.name}" for cmd in comandos_servidor))
        else:
            comandos_globais = await bot.tree.sync()
            print(f"Comandos sincronizados globalmente: {len(comandos_globais)}")

        comandos_ja_sincronizados = True
    except Exception as e:
        print(f"Erro ao sincronizar comandos: {e}")

    print('VERSAO COMPLETA: mesas | procurados | catalogo | boletins | organizacoes | relatorios')
    print(f"Bot online como {bot.user}")

# =====================================================
# SISTEMA DE PAINEL DE RELATÓRIOS OPERACIONAIS (COMPLETO)
# =====================================================

RELATORIOS_CONTADOR_JSON = DATA_DIR / "relatorios_contador.json"

# Canais informados para publicação automática
CANAIS_RELATORIOS = {
    "tocaia": 1490200477248520333,
    "olb": 1490200479995789374,
    "pericia_externa": 1490200524367200297,
    "diario": 1490200525340278854
}

def obter_proximo_numero_relatorio(tipo_relatorio: str) -> str:
    contadores = carregar_json(RELATORIOS_CONTADOR_JSON, {})
    if not isinstance(contadores, dict):
        contadores = {}
    
    proximo = contadores.get(tipo_relatorio, 0) + 1
    contadores[tipo_relatorio] = proximo
    salvar_json(RELATORIOS_CONTADOR_JSON, contadores)
    return f"{proximo:03d}"

class RelatoriosPainelView(View):
    def __init__(self):
        super().__init__(timeout=None)

    async def criar_canal_temporario_operacional(self, interaction: discord.Interaction, nome: str) -> Optional[discord.TextChannel]:
        guild = interaction.guild
        if not guild:
            return None
        
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(view_channel=False),
            interaction.user: discord.PermissionOverwrite(
                view_channel=True,
                send_messages=True,
                attach_files=True,       # Permite anexar arquivos/fotos diretamente
                embed_links=True,        # Permite enviar links com preview/imagens externas
                read_message_history=True
            )
        }
        
        try:
            if 'cargos_equipe_permissoes' in globals():
                overwrites = cargos_equipe_permissoes(guild)
                overwrites[interaction.user] = discord.PermissionOverwrite(
                    view_channel=True, send_messages=True, attach_files=True, embed_links=True, read_message_history=True
                )
        except Exception:
            pass

        cat_id = globals().get("PROCURADOS_TEMP_CATEGORY_ID") or globals().get("BOLETIM_TEMP_CATEGORY_ID") or 0
        categoria = guild.get_channel(cat_id) if cat_id else None
        
        nome_canal = f"relatorio-{nome}-{interaction.user.name}".lower().replace(" ", "-")
        
        try:
            canal = await guild.create_text_channel(
                name=nome_canal,
                category=categoria,
                overwrites=overwrites,
                reason="Painel de Relatório Operacional Temporário"
            )
            return canal
        except Exception as e:
            print(f"Erro crítico ao criar canal de relatório: {e}")
            return None

    @discord.ui.button(label="👀 TOCAIA", style=discord.ButtonStyle.secondary, custom_id="rel_btn_tocaia")
    async def btn_tocaia(self, interaction: discord.Interaction, button: Button):
        await interaction.response.defer(ephemeral=True)
        canal = await self.criar_canal_temporario_operacional(interaction, "tocaia")
        if not canal:
            return await interaction.followup.send("❌ Não foi possível criar o canal temporário.", ephemeral=True)
        
        await interaction.followup.send(f"✅ Canal temporário criado: {canal.mention}", ephemeral=True)
        await canal.send(f"👋 {interaction.user.mention}, clica no botão abaixo para abrir o formulário de Tocaia e usa este canal para enviar links ou anexar fotos se necessário.", view=IniciarFormularioRelatorioView("tocaia"))

    @discord.ui.button(label="🚔 OLB", style=discord.ButtonStyle.secondary, custom_id="rel_btn_olb")
    async def btn_olb(self, interaction: discord.Interaction, button: Button):
        await interaction.response.defer(ephemeral=True)
        canal = await self.criar_canal_temporario_operacional(interaction, "olb")
        if not canal:
            return await interaction.followup.send("❌ Não foi possível criar o canal temporário.", ephemeral=True)
        
        await interaction.followup.send(f"✅ Canal temporário criado: {canal.mention}", ephemeral=True)
        await canal.send(f"👋 {interaction.user.mention}, clica no botão abaixo para abrir o formulário de OLB e usa este canal para enviar links ou anexar fotos se necessário.", view=IniciarFormularioRelatorioView("olb"))

    @discord.ui.button(label="🔬 PERÍCIA EXTERNA", style=discord.ButtonStyle.secondary, custom_id="rel_btn_pericia_ext")
    async def btn_pericia_ext(self, interaction: discord.Interaction, button: Button):
        await interaction.response.defer(ephemeral=True)
        canal = await self.criar_canal_temporario_operacional(interaction, "pericia_externa")
        if not canal:
            return await interaction.followup.send("❌ Não foi possível criar o canal temporário.", ephemeral=True)
        
        await interaction.followup.send(f"✅ Canal temporário criado: {canal.mention}", ephemeral=True)
        await canal.send(f"👋 {interaction.user.mention}, clica no botão abaixo para abrir o formulário de Perícia Externa e usa este canal para enviar links ou anexar fotos se necessário.", view=IniciarFormularioRelatorioView("pericia_externa"))

    @discord.ui.button(label="📑 RELATÓRIO DIÁRIO", style=discord.ButtonStyle.secondary, custom_id="rel_btn_diario")
    async def btn_diario(self, interaction: discord.Interaction, button: Button):
        await interaction.response.defer(ephemeral=True)
        canal = await self.criar_canal_temporario_operacional(interaction, "diario")
        if not canal:
            return await interaction.followup.send("❌ Não foi possível criar o canal temporário.", ephemeral=True)
        
        await interaction.followup.send(f"✅ Canal temporário criado: {canal.mention}", ephemeral=True)
        await canal.send(f"👋 {interaction.user.mention}, clica no botão abaixo para abrir o formulário de Relatório Diário e usa este canal para enviar links ou anexar fotos se necessário.", view=IniciarFormularioRelatorioView("diario"))


class IniciarFormularioRelatorioView(View):
    def __init__(self, tipo: str):
        super().__init__(timeout=None)
        self.tipo = tipo

    @discord.ui.button(label="✍️ Preencher Formulário", style=discord.ButtonStyle.primary, custom_id="rel_btn_preencher")
    async def preencher(self, interaction: discord.Interaction, button: Button):
        if self.tipo == "tocaia":
            await interaction.response.send_modal(TocaiaModal())
        elif self.tipo == "olb":
            await interaction.response.send_modal(OlbModal())
        elif self.tipo == "pericia_externa":
            await interaction.response.send_modal(PericiaExternaModal())
        elif self.tipo == "diario":
            await interaction.response.send_modal(RelatorioDiarioModal())


async def finalizar_e_postar_relatorio(interaction: discord.Interaction, tipo: str, texto_conteudo: str):
    canal_id = CANAIS_RELATORIOS.get(tipo)
    canal_destino = interaction.guild.get_channel(canal_id) if interaction.guild and canal_id else None
    
    if canal_destino:
        await canal_destino.send(texto_conteudo)
    
    canal_atual = interaction.channel
    if isinstance(canal_atual, discord.TextChannel):
        await canal_atual.send("✅ Relatório enviado com sucesso! Este canal será apagado em 5 segundos...")
        await asyncio.sleep(5)
        try:
            await canal_atual.delete(reason="Relatório operacional concluído.")
        except Exception:
            pass


class TocaiaModal(Modal, title="Relatório de Tocaia"):
    local = TextInput(label="Local", placeholder="Local de interesse observado", max_length=150)
    tempo = TextInput(label="Tempo de tocaia", placeholder="Duração da vigilância", max_length=100)
    infos = TextInput(label="Informações obtidas", style=discord.TextStyle.paragraph, placeholder="Detalhes colhidos durante a tocaia", max_length=1000)

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        num = obter_proximo_numero_relatorio("tocaia")
        data_hora = agora_br()
        
        texto = (
            f"**RELATÓRIO DE TOCAIA Nº {num}**\n\n"
            f"**RESPONSÁVEL:**\n{interaction.user.mention}\n\n"
            f"**LOCAL:**\n{self.local.value}\n\n"
            f"**TEMPO DE TOCAIA:**\n{self.tempo.value}\n\n"
            f"**INFORMAÇÕES OBTIDAS:**\n{self.infos.value}\n\n"
            f"**DATA:** {data_hora.split()[0]} | **HORÁRIO:** {data_hora.split()[1]}"
        )
        await finalizar_e_postar_relatorio(interaction, "tocaia", texto)


class OlbModal(Modal, title="Relatório de OLB"):
    dicors = TextInput(label="DICORs envolvidos", placeholder="Agentes participantes da operação", max_length=200)
    relato = TextInput(label="Relatório da emboscada", style=discord.TextStyle.paragraph, placeholder="Como ocorreu a operação", max_length=1000)
    itens = TextInput(label="Itens ilegais apreendidos", style=discord.TextStyle.paragraph, placeholder="Lista de itens retidos", max_length=500)
    pericia = TextInput(label="Houve perícia?", placeholder="Sim / Não", max_length=150)
    prejuizo = TextInput(label="Prejuízo estimado", placeholder="Valor aproximado do impacto", max_length=100)

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        num = obter_proximo_numero_relatorio("olb")
        data_hora = agora_br()
        
        texto = (
            f"**RELATÓRIO DE OLB Nº {num}**\n\n"
            f"**RESPONSÁVEL:**\n{interaction.user.mention}\n\n"
            f"**DICORs ENVOLVIDOS:**\n{self.dicors.value}\n\n"
            f"**RELATÓRIO DA EMBOSCADA:**\n{self.relato.value}\n\n"
            f"**ITENS ILEGAIS APREENDIDOS:**\n{self.itens.value}\n\n"
            f"**HOUVE PERÍCIA:**\n{self.pericia.value}\n\n"
            f"**PREJUÍZO ESTIMADO:**\n≈ {self.prejuizo.value}\n\n"
            f"**DATA:** {data_hora.split()[0]} | **HORÁRIO:** {data_hora.split()[1]}"
        )
        await finalizar_e_postar_relatorio(interaction, "olb", texto)


class PericiaExternaModal(Modal, title="Perícia Externa"):
    codigo = TextInput(label="Código da ocorrência", placeholder="Número de referência", max_length=100)
    local = TextInput(label="Local", placeholder="Local da perícia", max_length=150)
    suspeito = TextInput(label="Suspeito", placeholder="Nome ou descrição", max_length=150, required=False)
    conclusao = TextInput(label="Conclusão", style=discord.TextStyle.paragraph, placeholder="Resultados das análises", max_length=1000)
    fotos = TextInput(label="Deseja anexar fotos?", placeholder="Sim / Não", max_length=200, required=False)

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        num = obter_proximo_numero_relatorio("pericia_externa")
        data_hora = agora_br()
        
        texto = (
            f"**RELATÓRIO DE PERÍCIA EXTERNA Nº {num}**\n\n"
            f"**RESPONSÁVEL:**\n{interaction.user.mention}\n\n"
            f"**ANÁLISE DO LOCAL:**\nCódigo da Ocorrência: {self.codigo.value}\n\n"
            f"**LOCAL:**\n{self.local.value}\n\n"
            f"**REGISTRO FOTOGRÁFICO:**\n{self.fotos.value or 'Em anexo'}\n\n"
            f"**SUSPEITO:**\n{self.suspeito.value or 'Não identificado'}\n\n"
            f"**CONCLUSÃO:**\n{self.conclusao.value}\n\n"
            f"**DATA:** {data_hora.split()[0]} | **HORÁRIO:** {data_hora.split()[1]}"
        )
        await finalizar_e_postar_relatorio(interaction, "pericia_externa", texto)


class RelatorioDiarioModal(Modal, title="Relatório Diário de Perícia"):
    material = TextInput(label="Material apreendido", style=discord.TextStyle.paragraph, placeholder="Descrição detalhada", max_length=400)
    local = TextInput(label="Local", placeholder="Local dos fatos", max_length=150)
    proprietario = TextInput(label="Proprietário", placeholder="Nome completo", max_length=150, required=False)
    telefone = TextInput(label="Telefone", placeholder="Contato", max_length=50, required=False)
    rg_doc = TextInput(label="RG / Placa / Modelo", placeholder="Documentos e Veículo", max_length=150, required=False)
    relato = TextInput(label="Relato dos fatos", style=discord.TextStyle.paragraph, placeholder="Resumo dos acontecimentos", max_length=1000)

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        num = obter_proximo_numero_relatorio("diario")
        data_hora = agora_br()
        
        texto = (
            f"**RELATÓRIO DE PERÍCIA INVESTIGATIVA Nº {num}**\n\n"
            f"**PERITO RESPONSÁVEL:**\n{interaction.user.mention}\n\n"
            f"**MATERIAL APREENDIDO:**\n{self.material.value}\n\n"
            f"**LOCAL:**\n{self.local.value}\n\n"
            f"**PROPRIETÁRIO:**\n{self.proprietario.value or 'Não informado'}\n\n"
            f"**TELEFONE:**\n{self.telefone.value or 'Não informado'}\n\n"
            f"**RG / PLACA / MODELO:**\n{self.rg_doc.value or 'Não informado'}\n\n"
            f"**RELATO DOS FATOS:**\n{self.relato.value}\n\n"
            f"**DATA:** {data_hora.split()[0]} | **HORÁRIO:** {data_hora.split()[1]}"
        )
        await finalizar_e_postar_relatorio(interaction, "diario", texto)


@bot.tree.command(name="painel-relatorios", description="Envia o painel com os Relatórios Operacionais.")
async def painel_relatorios(interaction: discord.Interaction):
    if not usuario_tem_admin(interaction.user):
        return await interaction.response.send_message("❌ Apenas membros autorizados podem usar este comando.", ephemeral=True)
        
    embed = discord.Embed(
        title="📑 PAINEL DE RELATÓRIOS OPERACIONAIS",
        description=(
            "Selecione uma das opções abaixo para iniciar um procedimento operacional.\n\n"
            "👀 **TOCAIA:** Registrar vigilâncias em locais de interesse.\n"
            "🚔 **OLB:** Registro de emboscadas e operações planejadas.\n"
            "🔬 **PERÍCIA EXTERNA:** Análises e levantamentos periciais externos.\n"
            "📑 **RELATÓRIO DIÁRIO:** Consolidação de perícias investigativas realizadas."
        ),
        color=discord.Color.blurple()
    )
    embed.set_footer(text="DICOR • Sistema Operacional Seguro")
    await interaction.response.send_message(embed=embed, view=RelatoriosPainelView())


# =====================================================
# MAIN
# =====================================================

async def main():
    carregar_boletins_pendentes_memoria()
    garantir_56_organizacoes()
    garantir_senha_catalogo()
    if not DISCORD_TOKEN or DISCORD_TOKEN == "COLE_O_TOKEN_DO_BOT_AQUI":
        print("ERRO: Coloque o token correto nas variáveis do Railway ou no arquivo .env.")
        return
    await start_web_server()
    await bot.start(DISCORD_TOKEN)

if __name__ == "__main__":
    asyncio.run(main())
