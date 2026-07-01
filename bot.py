# =====================================================
# BOT DIC - PROCURADOS + CATALOGO HTML AUTOMATICO
# Feito para GTA RP / personagens ficticios
# =====================================================

import os
import re
import json
import html
import asyncio
import datetime
import unicodedata
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple

import discord
from discord import app_commands
from discord.ext import commands
from discord.ui import View, Button, Modal, TextInput
from dotenv import load_dotenv
from aiohttp import web

# =====================================================
# CONFIGURACAO
# =====================================================

load_dotenv()

DISCORD_TOKEN = os.getenv("DISCORD_TOKEN", "").strip()
GUILD_ID = int(os.getenv("GUILD_ID", "0") or 0)
PROCURADOS_CHANNEL_ID = int(os.getenv("PROCURADOS_CHANNEL_ID", "0") or 0)
HISTORICO_PROCURADOS_ID = int(os.getenv("HISTORICO_PROCURADOS_ID", "0") or 0)
LOGS_CHANNEL_ID = int(os.getenv("LOGS_CHANNEL_ID", "0") or 0)
PROCURADOS_TEMP_CATEGORY_ID = int(os.getenv("PROCURADOS_TEMP_CATEGORY_ID", "0") or 0)
CATALOG_PUBLIC_URL = os.getenv("CATALOG_PUBLIC_URL", "http://127.0.0.1:8000/").strip()
PORT = int(os.getenv("PORT", "8000") or 8000)

CARGOS_ADMIN_IDS = []
for parte in os.getenv("CARGOS_ADMIN_IDS", "").split(","):
    parte = parte.strip()
    if parte.isdigit():
        CARGOS_ADMIN_IDS.append(int(parte))

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
PUBLIC_DIR = BASE_DIR / "public"
UPLOADS_DIR = PUBLIC_DIR / "uploads"
CATALOGO_JSON = DATA_DIR / "procurados.json"
CATALOGO_HTML = PUBLIC_DIR / "index.html"

DATA_DIR.mkdir(exist_ok=True)
PUBLIC_DIR.mkdir(exist_ok=True)
UPLOADS_DIR.mkdir(exist_ok=True)

# =====================================================
# BOT
# =====================================================

intents = discord.Intents.default()
intents.guilds = True
intents.members = True
intents.messages = True
intents.message_content = True  # Ative tambem no Developer Portal se for sincronizar mensagens antigas.

bot = commands.Bot(command_prefix="!", intents=intents)

# Canais provisórios em aberto: canal_id -> dados do procurado
cadastros_pendentes: Dict[int, Dict[str, Any]] = {}

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
    return texto[:50] or "procurado"


def limpar_rg(rg: str) -> str:
    return re.sub(r"\s+", "", str(rg or "").lower())


def escape(texto: Any) -> str:
    return html.escape(str(texto or ""))


def normalizar_label(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode("ascii")
    return texto.lower().strip()


def carregar_procurados() -> List[Dict[str, Any]]:
    if not CATALOGO_JSON.exists():
        return []
    try:
        with open(CATALOGO_JSON, "r", encoding="utf-8") as f:
            dados = json.load(f)
        if isinstance(dados, list):
            return dados
        return []
    except Exception:
        return []


def salvar_procurados(lista: List[Dict[str, Any]]) -> None:
    with open(CATALOGO_JSON, "w", encoding="utf-8") as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)


def procurar_por_rg(rg: str) -> Optional[Dict[str, Any]]:
    alvo = limpar_rg(rg)
    for p in carregar_procurados():
        if limpar_rg(p.get("rg", "")) == alvo:
            return p
    return None


def substituir_por_rg(registro: Dict[str, Any]) -> None:
    lista = carregar_procurados()
    rg = limpar_rg(registro.get("rg", ""))
    novo = []
    achou = False
    for item in lista:
        if limpar_rg(item.get("rg", "")) == rg:
            novo.append(registro)
            achou = True
        else:
            novo.append(item)
    if not achou:
        novo.append(registro)
    salvar_procurados(novo)
    gerar_catalogo_html()


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

# =====================================================
# CATALOGO HTML
# =====================================================

def gerar_catalogo_html() -> None:
    procurados = carregar_procurados()
    ativos = [p for p in procurados if p.get("status", "A PROCURAR") == "A PROCURAR"]
    retirados = [p for p in procurados if p.get("status") == "RETIRADO"]

    cards = ""
    for p in procurados:
        status = p.get("status", "A PROCURAR")
        classe_status = "retirado" if status == "RETIRADO" else "ativo"
        foto_ind = p.get("foto_individuo", "") or ""
        foto_rg = p.get("foto_rg", "") or ""
        link_msg = p.get("mensagem_url", "")

        cards += f"""
        <article class="card {classe_status}">
            <div class="card-head">
                <div><span>Nº DO CASO</span><b>{escape(p.get('caso'))}</b></div>
                <div><span>DATA</span><b>{escape(p.get('data'))}</b></div>
                <div class="status">{escape(status)}</div>
            </div>

            <div class="card-grid">
                <div class="photos">
                    <div class="photo big">
                        <div class="photo-title">Foto do Indivíduo</div>
                        {f'<img src="{escape(foto_ind)}" alt="Foto do indivíduo">' if foto_ind else '<div class="placeholder">Sem foto</div>'}
                    </div>
                    <div class="photo small">
                        <div class="photo-title">Foto do RG</div>
                        {f'<img src="{escape(foto_rg)}" alt="Foto do RG">' if foto_rg else '<div class="placeholder">Sem RG</div>'}
                    </div>
                </div>

                <div class="info">
                    <div class="linha"><b>Nome</b><p>{escape(p.get('nome'))}</p></div>
                    <div class="linha"><b>RG</b><p>{escape(p.get('rg'))}</p></div>
                    <div class="box"><b>Crimes Cometidos</b><p>{escape(p.get('crimes'))}</p></div>
                    <div class="box destaque"><b>Último Avistamento</b><p>{escape(p.get('ultimo_avistamento'))}</p></div>
                    <div class="box"><b>Informações</b><p>{escape(p.get('informacoes'))}</p></div>
                    {f'<a class="msg" href="{escape(link_msg)}" target="_blank">Abrir mensagem no Discord</a>' if link_msg else ''}
                    {f'<div class="motivo"><b>Motivo da retirada:</b> {escape(p.get("motivo_retirada"))}</div>' if p.get('motivo_retirada') else ''}
                </div>
            </div>
        </article>
        """

    if not cards:
        cards = """
        <div class="vazio">
            <h2>Nenhum procurado cadastrado ainda.</h2>
            <p>Quando o bot finalizar um cadastro, ele aparecerá aqui automaticamente.</p>
        </div>
        """

    html_final = f"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Catálogo de Procurados - DIC</title>
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
            padding: 28px 18px;
            text-align: center;
            background: rgba(0,0,0,.38);
            position: sticky;
            top: 0;
            z-index: 10;
            backdrop-filter: blur(7px);
        }}
        header h1 {{
            margin: 0;
            color: #f0b64d;
            font-size: clamp(30px, 5vw, 62px);
            letter-spacing: 2px;
            text-transform: uppercase;
        }}
        header p {{ margin: 8px 0 0; color: #ddd; }}
        .stats {{
            display: flex;
            gap: 14px;
            justify-content: center;
            flex-wrap: wrap;
            margin-top: 18px;
        }}
        .stat {{
            border: 1px solid rgba(240,182,77,.55);
            background: rgba(8,19,35,.9);
            padding: 10px 18px;
            border-radius: 12px;
            min-width: 150px;
        }}
        .stat span {{ display: block; font-size: 12px; color: #bfc6d1; text-transform: uppercase; }}
        .stat b {{ font-size: 24px; color: #fff; }}
        main {{
            width: min(1380px, 96%);
            margin: 24px auto 50px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(520px, 1fr));
            gap: 22px;
        }}
        .card {{
            border: 1px solid rgba(240,182,77,.55);
            background: linear-gradient(180deg, rgba(8,19,35,.97), rgba(3,8,16,.97));
            box-shadow: 0 0 0 1px rgba(255,255,255,.05), 0 18px 50px rgba(0,0,0,.32);
            border-radius: 18px;
            overflow: hidden;
        }}
        .card.retirado {{ opacity: .65; filter: grayscale(.35); }}
        .card-head {{
            display: grid;
            grid-template-columns: 1fr 1fr auto;
            gap: 12px;
            padding: 14px;
            border-bottom: 1px solid rgba(240,182,77,.32);
            align-items: center;
        }}
        .card-head span {{
            display: block;
            font-size: 11px;
            color: #f0b64d;
            text-transform: uppercase;
            font-weight: bold;
        }}
        .card-head b {{ display: block; margin-top: 4px; }}
        .status {{
            color: #101820;
            background: #f0b64d;
            font-weight: 900;
            padding: 8px 12px;
            border-radius: 10px;
            white-space: nowrap;
        }}
        .retirado .status {{ background: #c94a4a; color: #fff; }}
        .card-grid {{
            display: grid;
            grid-template-columns: 210px 1fr;
            gap: 16px;
            padding: 16px;
        }}
        .photos {{ display: grid; gap: 14px; }}
        .photo {{
            border: 1px solid rgba(240,182,77,.55);
            background: #081323;
            border-radius: 14px;
            padding: 10px;
        }}
        .photo-title {{
            text-align: center;
            color: #f0b64d;
            font-weight: 800;
            margin-bottom: 8px;
            text-transform: uppercase;
            font-size: 13px;
        }}
        .photo img {{
            width: 100%;
            height: 240px;
            object-fit: cover;
            border-radius: 10px;
            background: #ddd;
            display: block;
        }}
        .photo.small img {{ height: 125px; object-fit: cover; }}
        .placeholder {{
            height: 240px;
            border: 1px dashed #98a2b3;
            display: grid;
            place-items: center;
            color: #98a2b3;
            border-radius: 10px;
        }}
        .photo.small .placeholder {{ height: 125px; }}
        .info {{ display: grid; gap: 10px; }}
        .linha, .box {{
            background: #f6f7f9;
            color: #07111f;
            border-radius: 12px;
            padding: 10px 12px;
            border-left: 5px solid #f0b64d;
        }}
        .linha b, .box b {{ display: block; color: #07111f; margin-bottom: 3px; }}
        .linha p, .box p {{ margin: 0; white-space: pre-wrap; line-height: 1.35; }}
        .destaque {{ background: #fff4dc; }}
        .msg {{
            display: inline-block;
            color: #f0b64d;
            text-decoration: none;
            font-weight: bold;
            margin-top: 4px;
        }}
        .motivo {{
            background: rgba(201,74,74,.18);
            color: #ffd7d7;
            border: 1px solid rgba(201,74,74,.4);
            padding: 10px;
            border-radius: 10px;
        }}
        .vazio {{
            grid-column: 1 / -1;
            text-align: center;
            padding: 60px 20px;
            border: 1px dashed rgba(240,182,77,.6);
            border-radius: 18px;
            background: rgba(8,19,35,.7);
        }}
        footer {{
            text-align: center;
            color: #bfc6d1;
            border-top: 1px solid rgba(240,182,77,.25);
            padding: 18px;
        }}
        @media (max-width: 720px) {{
            main {{ grid-template-columns: 1fr; }}
            .card-grid {{ grid-template-columns: 1fr; }}
            .card-head {{ grid-template-columns: 1fr; }}
            .photo img {{ height: 330px; }}
        }}
    </style>
</head>
<body>
    <header>
        <h1>Catálogo de Procurados</h1>
        <p>DIC • Atualizado automaticamente pelo bot</p>
        <div class="stats">
            <div class="stat"><span>Registros totais</span><b>{len(procurados)}</b></div>
            <div class="stat"><span>Ativos</span><b>{len(ativos)}</b></div>
            <div class="stat"><span>Retirados</span><b>{len(retirados)}</b></div>
            <div class="stat"><span>Última atualização</span><b style="font-size:16px">{agora_br()}</b></div>
        </div>
    </header>
    <main>
        {cards}
    </main>
    <footer>
        Uso exclusivo para GTA RP / sistema interno autorizado. Não use com dados reais de pessoas.
    </footer>
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
    lista = carregar_procurados()
    lista.append(registro)
    lista = remover_duplicados(lista)
    salvar_procurados(lista)
    gerar_catalogo_html()

# =====================================================
# SERVIDOR HTML
# =====================================================

async def start_web_server():
    gerar_catalogo_html()
    app = web.Application()
    app.router.add_static("/", path=str(PUBLIC_DIR), show_index=False)
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    print(f"Catalogo rodando em http://127.0.0.1:{PORT}/")

# =====================================================
# EMBEDS / POSTAGEM
# =====================================================

def criar_embed_procurado(registro: Dict[str, Any]) -> discord.Embed:
    cor = discord.Color.red() if registro.get("status") == "A PROCURAR" else discord.Color.dark_gray()
    embed = discord.Embed(
        title="🚨 Mandado de Busca e Apreensão/Prisão",
        description="**Sistema de Procurados - DIC**",
        color=cor,
    )
    embed.add_field(name="👤 Nome", value=registro.get("nome", "Não informado"), inline=True)
    embed.add_field(name="🪪 RG", value=registro.get("rg", "Não informado"), inline=True)
    embed.add_field(name="📌 Status", value=registro.get("status", "A PROCURAR"), inline=True)
    embed.add_field(name="⚖️ Crimes Cometidos", value=registro.get("crimes", "Não informado"), inline=False)
    embed.add_field(name="📍 Último Avistamento", value=registro.get("ultimo_avistamento", "Não informado"), inline=False)
    embed.add_field(name="ℹ️ Informações", value=registro.get("informacoes", "Não informado"), inline=False)
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

    embed = criar_embed_procurado(registro)
    msg = await canal.send(embed=embed, files=arquivos)
    return msg

# =====================================================
# PAINEL E CADASTRO POR CANAL PROVISORIO
# =====================================================

class NovoProcuradoModal(Modal, title="Cadastrar Novo Procurado"):
    nome = TextInput(label="Nome", placeholder="Nome do procurado", max_length=100)
    rg = TextInput(label="RG", placeholder="RG do procurado", max_length=50)
    crimes = TextInput(label="Crimes Cometidos", placeholder="Ex: Tráfico, fuga, porte ilegal...", style=discord.TextStyle.paragraph, max_length=1000)
    ultimo = TextInput(label="Último Avistamento", placeholder="Ex: Porto, favela, bairro, data...", style=discord.TextStyle.paragraph, max_length=600)
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
            guild.me: discord.PermissionOverwrite(view_channel=True, send_messages=True, manage_channels=True, read_message_history=True, attach_files=True),
        }

        for cargo_id in CARGOS_ADMIN_IDS:
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
            f"🚨 **Cadastro de Procurado — DIC**\n\n"
            f"👤 **Nome:** {self.nome.value}\n"
            f"🪪 **RG:** {self.rg.value}\n\n"
            f"📸 Envie aqui **2 imagens obrigatórias**:\n"
            f"**1. Foto do indivíduo**\n"
            f"**2. Foto do RG**\n\n"
            f"Depois clique em **✅ Finalizar Cadastro**.",
            view=FinalizarProcuradoView()
        )

        await enviar_log(
            f"📁 **Canal provisório de procurado criado**\n"
            f"Responsável: {interaction.user.mention}\n"
            f"Nome: {self.nome.value}\n"
            f"RG: {self.rg.value}\n"
            f"Canal: {canal.mention}"
        )

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
                if a.content_type and a.content_type.startswith("image/"):
                    anexos.append(a)
                elif Path(a.filename).suffix.lower() in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
                    anexos.append(a)

        if len(anexos) < 2:
            await interaction.response.send_message("❌ Envie as 2 imagens: foto do indivíduo e foto do RG.", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True)

        foto_ind = await salvar_anexo_publico(anexos[0], f"individuo-{dados['rg']}")
        foto_rg = await salvar_anexo_publico(anexos[1], f"rg-{dados['rg']}")

        registro = {
            "id": data_caso(),
            "caso": f"DIC-{data_caso()}",
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
        await enviar_log(
            f"✅ **Procurado cadastrado e enviado ao catálogo**\n"
            f"Nome: {registro['nome']}\n"
            f"RG: {registro['rg']}\n"
            f"Catálogo: {CATALOG_PUBLIC_URL}"
        )

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
            linhas = []
            for p in ativos[:20]:
                linhas.append(f"• **{p.get('nome','Sem nome')}** — RG: `{p.get('rg','')}`")
            texto = "\n".join(linhas)
            if len(ativos) > 20:
                texto += f"\n... e mais {len(ativos)-20} no catálogo."
        await interaction.response.send_message(f"📋 **Procurados ativos:**\n{texto}\n\n🔗 {CATALOG_PUBLIC_URL}", ephemeral=True)

    @discord.ui.button(label="Retirar Procurado", emoji="❌", style=discord.ButtonStyle.gray, custom_id="dic_retirar_procurado")
    async def retirar(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(RetirarProcuradoModal())


async def retirar_procurado(interaction: discord.Interaction, rg: str, motivo: str):
    lista = carregar_procurados()
    alvo = limpar_rg(rg)
    encontrado = None
    for p in lista:
        if limpar_rg(p.get("rg", "")) == alvo:
            encontrado = p
            break

    if not encontrado:
        await interaction.response.send_message("❌ Não encontrei procurado com esse RG no catálogo.", ephemeral=True)
        return

    encontrado["status"] = "RETIRADO"
    encontrado["motivo_retirada"] = motivo
    encontrado["data_retirada"] = agora_br()
    encontrado["retirado_por"] = str(interaction.user)

    salvar_procurados(lista)
    gerar_catalogo_html()

    # Tenta editar a mensagem oficial no canal de procurados
    if encontrado.get("mensagem_id") and PROCURADOS_CHANNEL_ID:
        canal = bot.get_channel(PROCURADOS_CHANNEL_ID)
        if canal:
            try:
                msg = await canal.fetch_message(int(encontrado["mensagem_id"]))
                embed = criar_embed_procurado(encontrado)
                embed.add_field(name="❌ Motivo da Retirada", value=motivo, inline=False)
                await msg.edit(embed=embed)
            except Exception:
                pass

    # Envia para histórico se existir
    if HISTORICO_PROCURADOS_ID:
        historico = bot.get_channel(HISTORICO_PROCURADOS_ID)
        if historico:
            try:
                await historico.send(
                    f"❌ **Procurado Retirado**\n"
                    f"👤 Nome: {encontrado.get('nome')}\n"
                    f"🪪 RG: {encontrado.get('rg')}\n"
                    f"📌 Motivo: {motivo}\n"
                    f"👮 Retirado por: {interaction.user.mention}"
                )
            except Exception:
                pass

    await enviar_log(f"❌ Procurado retirado: {encontrado.get('nome')} | RG {encontrado.get('rg')} | Motivo: {motivo}")
    await interaction.response.send_message(f"✅ Procurado retirado do status ativo e catálogo atualizado.\n🔗 {CATALOG_PUBLIC_URL}", ephemeral=True)

# =====================================================
# SINCRONIZAR PROCURADOS ANTIGOS
# =====================================================

def extrair_por_labels(texto: str) -> Dict[str, str]:
    dados = {}
    padroes = {
        "nome": r"(?:nome|indiv[ií]duo)\s*[:\-]\s*(.+)",
        "rg": r"(?:rg|passaporte)\s*[:\-]\s*(.+)",
        "crimes": r"(?:crimes cometidos|crimes imputados|crimes|crime)\s*[:\-]\s*(.+)",
        "ultimo_avistamento": r"(?:[úu]ltimo avistamento|avistamento)\s*[:\-]\s*(.+)",
        "informacoes": r"(?:informa[cç][õo]es|observa[cç][õo]es|detalhes)\s*[:\-]\s*(.+)",
    }
    for chave, padrao in padroes.items():
        m = re.search(padrao, texto, flags=re.I)
        if m:
            dados[chave] = m.group(1).strip()[:1000]
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

    # Se nao achou RG e Nome, ignora
    if not dados.get("rg") or not dados.get("nome"):
        return None

    # Evita duplicar
    if procurar_por_rg(dados.get("rg", "")):
        return None

    imagens = []
    for a in msg.attachments:
        if (a.content_type and a.content_type.startswith("image/")) or Path(a.filename).suffix.lower() in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
            imagens.append(a)

    foto_ind = ""
    foto_rg = ""
    if len(imagens) >= 1:
        foto_ind = await salvar_anexo_publico(imagens[0], f"sync-ind-{dados.get('rg','')}")
    if len(imagens) >= 2:
        foto_rg = await salvar_anexo_publico(imagens[1], f"sync-rg-{dados.get('rg','')}")

    registro = {
        "id": f"SYNC-{msg.id}",
        "caso": f"SYNC-{msg.id}",
        "data": msg.created_at.strftime("%d/%m/%Y %H:%M"),
        "status": "A PROCURAR",
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
    return registro

# =====================================================
# SLASH COMMANDS
# =====================================================

@bot.tree.command(name="painelprocurados", description="Envia o painel de gerenciamento de procurados.")
@app_commands.checks.has_permissions(administrator=True)
async def painelprocurados(interaction: discord.Interaction):
    embed = discord.Embed(
        title="🚨 Sistema de Procurados - DIC",
        description=(
            "Utilize os botões abaixo para gerenciar procurados.\n\n"
            "➕ **Novo Procurado** — Cadastrar um novo procurado.\n"
            "📋 **Lista de Procurados** — Ver procurados ativos.\n"
            "❌ **Retirar Procurado** — Retirar um procurado pelo RG."
        ),
        color=discord.Color.red(),
    )
    await interaction.response.send_message(embed=embed, view=PainelProcuradosView())


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
    if not PROCURADOS_CHANNEL_ID:
        await interaction.response.send_message("❌ Configure PROCURADOS_CHANNEL_ID no .env.", ephemeral=True)
        return

    canal = bot.get_channel(PROCURADOS_CHANNEL_ID)
    if not isinstance(canal, discord.TextChannel):
        await interaction.response.send_message("❌ Não encontrei o canal de procurados.", ephemeral=True)
        return

    await interaction.response.defer(ephemeral=True)

    importados = 0
    analisados = 0
    lista = carregar_procurados()

    async for msg in canal.history(limit=1000, oldest_first=True):
        analisados += 1
        registro = await importar_mensagem_antiga(msg)
        if registro:
            lista.append(registro)
            importados += 1

    lista = remover_duplicados(lista)
    salvar_procurados(lista)
    gerar_catalogo_html()

    await interaction.followup.send(
        f"✅ Sincronização finalizada.\n"
        f"Mensagens analisadas: `{analisados}`\n"
        f"Procurados importados: `{importados}`\n"
        f"Catálogo: {CATALOG_PUBLIC_URL}",
        ephemeral=True,
    )


@bot.tree.command(name="regerarcatalogo", description="Recria o HTML do catálogo a partir do JSON salvo.")
@app_commands.checks.has_permissions(administrator=True)
async def regerarcatalogo(interaction: discord.Interaction):
    gerar_catalogo_html()
    await interaction.response.send_message(f"✅ Catálogo recriado: {CATALOG_PUBLIC_URL}", ephemeral=True)

# =====================================================
# EVENTOS
# =====================================================

@bot.event
async def on_ready():
    bot.add_view(PainelProcuradosView())
    bot.add_view(FinalizarProcuradoView())

    try:
        if GUILD_ID:
            guild_obj = discord.Object(id=GUILD_ID)
            bot.tree.copy_global_to(guild=guild_obj)
            await bot.tree.sync(guild=guild_obj)
        else:
            await bot.tree.sync()
        print("Comandos sincronizados!")
    except Exception as e:
        print(f"Erro ao sincronizar comandos: {e}")

    print(f"Bot online como {bot.user}")

# =====================================================
# MAIN
# =====================================================

async def main():
    if not DISCORD_TOKEN or DISCORD_TOKEN == "COLE_O_TOKEN_DO_BOT_AQUI":
        print("ERRO: Coloque o token correto no arquivo .env em DISCORD_TOKEN.")
        return

    if not PROCURADOS_CHANNEL_ID:
        print("AVISO: PROCURADOS_CHANNEL_ID esta 0. Configure o canal oficial no .env.")

    await start_web_server()
    await bot.start(DISCORD_TOKEN)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Bot desligado.")
