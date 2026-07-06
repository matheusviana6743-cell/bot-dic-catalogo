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
# CATALOGO HTML MODIFICADO
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
    
    # Separação Correta por status
    ativos = [
        p for p in visiveis
        if str(p.get("status", "A PROCURAR") or "A PROCURAR").upper() not in ["RETIRADO", "PEGO", "ARQUIVADO"]
    ]
    retirados = [
        p for p in visiveis
        if str(p.get("status", "") or "").upper() in ["RETIRADO", "PEGO", "ARQUIVADO"]
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
        classe_status = "retirado" if status in ["RETIRADO", "PEGO", "ARQUIVADO"] else "ativo"
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
            <h2>Nenhum procurado arquivado.</h2>
            <p>Os registros pegos ou retirados aparecerão nesta aba.</p>
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
            <div class="stat"><span>Arquivados</span><b>{len(retirados)}</b></div>
            <div class="stat"><span>Última atualização</span><b style="font-size:16px">{agora_br()}</b></div>
        </div>
        <div class="abas">
            <button id="btn-ativos" class="aba-btn ativa" onclick="mostrarAba('ativos')">🚨 Procurados Ativos ({len(ativos)})</button>
            <button id="btn-retirados" class="aba-btn" onclick="mostrarAba('retirados')">📦 Arquivados / Pegos ({len(retirados)})</button>
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

A Polícia Polícia Federal de Capital Morada, por intermédio da **Divisão de Investigações Criminais (DICOR)**, informa que o indivíduo abaixo encontra-se oficialmente procurado pelas autoridades competentes. As investigações apontam seu envolvimento em atividades criminosas, havendo mandado ativo para sua localização, abordagem e condução para os procedimentos cabíveis.

📍 **ÚLTIMO AVISTAMENTO:** {registro.get('ultimo_avistamento', 'Não informado')}
⚠️ **CRIMES IMPUTADOS:** {registro.get('crimes', 'Não informado')}

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

    texto = cortar_discord(criar_texto_procurado(registro), 1900)
    msg = await canal.send(content=texto, files=arquivos)
    return msg


def arquivos_locais_procurado(
    registro: Dict[str, Any],
) -> List[discord.File]:
    arquivos: List[discord.File] = []
    for chave, nome_base in (
        {"foto_individuo", "foto_individuo"},
        {"foto_rg", "foto_rg"},
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
    """ Publica o procurado completo no histórico. O post ativo só é apagado depois do arquivamento confirmado. """
    canal_historico = await obter_canal_por_id(
        HISTORICO_PROCURADOS_ID
    )
    if (
        canal_historico is None or not hasattr(canal_historico, "send")
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
            canal_ativo is not None and hasattr(canal_ativo, "fetch_message")
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
    ultimo_avistamento = TextInput(label="Último Avistamento", placeholder="Local e hora aproximados", max_length=200)
    informacoes = TextInput(label="Informações Adicionais", placeholder="Características físicas, carros usados, etc.", style=discord.TextStyle.paragraph, max_length=1000, required=False)

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        caso_id = data_caso()
        
        registro = {
            "id": caso_id,
            "caso": caso_id,
            "data": agora_br(),
            "status": "A PROCURAR",
            "nome": self.nome.value.strip(),
            "rg": self.rg.value.strip(),
            "crimes": self.crimes.value.strip(),
            "ultimo_avistamento": self.ultimo_avistamento.value.strip(),
            "informacoes": self.informacoes.value.strip() or "Nenhuma",
            "foto_individuo": "",
            "foto_rg": "",
            "mensagem_id": "",
            "mensagem_url": "",
            "motivo_retirada": ""
        }
        
        cadastros_pendentes[interaction.user.id] = registro
        
        embed = discord.Embed(
            title="Cadastro Iniciado",
            description=f"Procurado **{registro['nome']}** pré-registrado.\n\nAgora, por favor, envie as fotos usando os botões abaixo.",
            color=discord.Color.orange()
        )
        await interaction.followup.send(embed=embed, view=FotosProcuradoView(interaction.user.id), ephemeral=True)


class FotosProcuradoView(View):
    def __init__(self, usuario_id: int):
        super().__init__(timeout=600)
        self.usuario_id = usuario_id

    @discord.ui.button(label="📸 Foto do Indivíduo", style=discord.ButtonStyle.primary)
    async def foto_individuo(self, interaction: discord.Interaction, button: Button):
        if interaction.user.id != self.usuario_id:
            await interaction.response.send_message("Você não iniciou este cadastro.", ephemeral=True)
            return
        await interaction.response.send_modal(UploadFotoModal(self.usuario_id, "foto_individuo"))

    @discord.ui.button(label="🪪 Foto do RG", style=discord.ButtonStyle.primary)
    async def foto_rg(self, interaction: discord.Interaction, button: Button):
        if interaction.user.id != self.usuario_id:
            await interaction.response.send_message("Você não iniciou este cadastro.", ephemeral=True)
            return
        await interaction.response.send_modal(UploadFotoModal(self.usuario_id, "foto_rg"))

    @discord.ui.button(label="✅ Finalizar e Postar", style=discord.ButtonStyle.success)
    async def finalizar(self, interaction: discord.Interaction, button: Button):
        if interaction.user.id != self.usuario_id:
            await interaction.response.send_message("Você não iniciou este cadastro.", ephemeral=True)
            return
            
        registro = cadastros_pendentes.get(self.usuario_id)
        if not registro:
            await interaction.response.send_message("Nenhum cadastro pendente encontrado.", ephemeral=True)
            return
            
        await interaction.response.defer(ephemeral=True)
        
        msg_oficial = await postar_procurado_oficial(registro)
        if msg_oficial:
            registro["mensagem_id"] = str(msg_oficial.id)
            registro["mensagem_url"] = msg_oficial.jump_url
            
        await salvar_procurado_catalogo(registro)
        cadastros_pendentes.pop(self.usuario_id, None)
        
        await interaction.followup.send("✅ Procurado postado com sucesso no Discord e integrado ao Catálogo!", ephemeral=True)
        await enviar_log(f"👮 **Novo procurado adicionado por {interaction.user}**\nNome: {registro['nome']}\nRG: {registro['rg']}")


class UploadFotoModal(Modal, title="Inserir Link da Foto"):
    url = TextInput(label="URL da Imagem", placeholder="Cole o link da foto aqui (Discord, Imgur, etc.)", max_length=500)

    def __init__(self, usuario_id: int, tipo: str):
        super().__init__()
        self.usuario_id = usuario_id
        self.tipo = tipo

    async def on_submit(self, interaction: discord.Interaction):
        registro = cadastros_pendentes.get(self.usuario_id)
        if not registro:
            await interaction.response.send_message("Sessão expirada. Comece novamente.", ephemeral=True)
            return
            
        registro[self.tipo] = self.url.value.strip()
        label = "Foto do Indivíduo" if self.tipo == "foto_individuo" else "Foto do RG"
        await interaction.response.send_message(f"✅ URL para **{label}** armazenada com sucesso!", ephemeral=True)


class RetirarProcuradoModal(Modal, title="Retirar Indivíduo da Procura"):
    motivo = TextInput(label="Motivo da Retirada", placeholder="Ex: Preso em ação / Mandado expirado / Pego pela equipe", style=discord.TextStyle.paragraph, max_length=500)

    def __init__(self, registro: Dict[str, Any]):
        super().__init__()
        self.registro = registro

    async def on_submit(self, interaction: discord.Interaction):
        await interaction.response.defer(ephemeral=True)
        
        motivo_txt = self.motivo.value.strip()
        self.registro["status"] = "RETIRADO"
        self.registro["motivo_retirada"] = motivo_txt

        try:
            msg_arq = await arquivar_procurado_discord(self.registro, motivo_txt, str(interaction.user))
            self.registro["mensagem_arquivada_id"] = str(msg_arq.id)
        except Exception as e:
            print(f"Erro ao arquivar post no Discord: {e}")

        # Atualiza o banco do catálogo
        lista = carregar_procurados()
        for i, p in enumerate(lista):
            if limpar_rg(p.get("rg", "")) == limpar_rg(self.registro.get("rg", "")):
                lista[i] = self.registro
                break
        salvar_procurados(lista)
        gerar_catalogo_html()

        await interaction.followup.send(f"✅ Indivíduo **{self.registro['nome']}** foi marcado como RETIRADO e movido para o histórico.", ephemeral=True)
        await enviar_log(f"📦 **Procurado retirado por {interaction.user}**\nNome: {self.registro['nome']}\nMotivo: {motivo_txt}")


@bot.tree.command(name="procurado-cadastrar", description="Cadastra um novo indivíduo na lista de procurados")
async def procurado_cadastrar(interaction: discord.Interaction):
    if not usuario_tem_admin(interaction.user):
        await interaction.response.send_message("Apenas membros autorizados podem usar este comando.", ephemeral=True)
        return
    await interaction.response.send_modal(NovoProcuradoModal())


@bot.tree.command(name="procurado-retirar", description="Remove um indivíduo da lista de procurados ativos por RG")
@app_commands.describe(rg="RG do indivíduo a ser retirado")
async def procurado_retirar(interaction: discord.Interaction, rg: str):
    if not usuario_tem_admin(interaction.user):
        await interaction.response.send_message("Apenas membros autorizados podem usar este comando.", ephemeral=True)
        return
        
    registro = procurar_por_rg(rg)
    if not registro:
        await interaction.response.send_message(f"Nenhum procurado ativo localizado com o RG: `{rg}`", ephemeral=True)
        return
        
    if registro.get("status") == "RETIRADO":
        await interaction.response.send_message("Este indivíduo já foi retirado anteriormente.", ephemeral=True)
        return
        
    await interaction.response.send_modal(RetirarProcuradoModal(registro))


@bot.tree.command(name="procurado-buscar", description="Busca informações de um procurado pelo RG")
@app_commands.describe(rg="RG para consulta")
async def procurado_buscar(interaction: discord.Interaction, rg: str):
    await interaction.response.defer(ephemeral=True)
    registro = procurar_por_rg(rg)
    if not registro:
        await interaction.followup.send("Nenhum registro encontrado para este RG.", ephemeral=True)
        return
        
    embed = criar_embed_procurado(registro)
    await interaction.followup.send(embed=embed, ephemeral=True)


@bot.tree.command(name="catalogo-url", description="Obtém a URL pública do catálogo de procurados")
async def catalogo_url(interaction: discord.Interaction):
    senha = garantir_senha_catalogo()
    embed = discord.Embed(
        title="🖥️ Catálogo de Procurados DICOR",
        description=f"Acesse o catálogo online em tempo real:\n🔗 **{CATALOG_PUBLIC_URL}**\n\n🔑 *Senha administrativa para exclusões:* `{senha}`",
        color=discord.Color.blue()
    )
    await interaction.response.send_message(embed=embed, ephemeral=True)

# =====================================================
# SINCRO E EVENTOS ON_READY
# =====================================================

@bot.event
async def on_ready():
    print(f"Logado com sucesso como {bot.user}")
    try:
        if GUILD_ID > 0:
            servidor = discord.Object(id=GUILD_ID)
            bot.tree.copy_global_to(guild=servidor)
            comandos_servidor = await bot.tree.sync(guild=servidor)
            bot.tree.clear_commands(guild=None)
            comandos_globais = await bot.tree.sync()
            print(f"Comandos sincronizados no servidor: {len(comandos_servidor)}")
        else:
            comandos_globais = await bot.tree.sync()
            print(f"Comandos sincronizados globalmente: {len(comandos_globais)}")
    except Exception as e:
        print(f"Erro ao sincronizar comandos: {e}")
    
    await start_web_server()


async def main():
    garantir_senha_catalogo()
    if not DISCORD_TOKEN:
        print("ERRO: Coloque o token correto.")
        return
    await bot.start(DISCORD_TOKEN)

if __name__ == "__main__":
    asyncio.run(main())
