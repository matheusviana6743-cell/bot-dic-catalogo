# =====================================================
# BOT DIC FINAL RESOLVIDO - MESAS + PROCURADOS + CATALOGO HTML
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
HISTORICO_PROCURADOS_ID = int(os.getenv("HISTORICO_PROCURADOS_ID", "0") or 0)
LOGS_CHANNEL_ID = int(os.getenv("LOGS_CHANNEL_ID", "0") or 0)
PROCURADOS_TEMP_CATEGORY_ID = int(os.getenv("PROCURADOS_TEMP_CATEGORY_ID", "0") or 0)

# Mesas
CATEGORIA_MESAS_ABERTAS_ID = int(os.getenv("CATEGORIA_MESAS_ABERTAS_ID", "1490200456855552192") or 1490200456855552192)
CATEGORIA_MESAS_FECHADAS_ID = int(os.getenv("CATEGORIA_MESAS_FECHADAS_ID", "1515165416815722586") or 1515165416815722586)
BACKUP_CHANNEL_ID = int(os.getenv("BACKUP_CHANNEL_ID", "1515165673276440677") or 1515165673276440677)

# Catalogo
CATALOG_PUBLIC_URL = os.getenv("CATALOG_PUBLIC_URL", "http://127.0.0.1:8000/").strip()
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
            )
    if guild.me:
        overwrites[guild.me] = discord.PermissionOverwrite(
            view_channel=True,
            send_messages=True,
            manage_channels=True,
            attach_files=True,
            read_message_history=True,
        )
    return overwrites

# =====================================================
# CATALOGO HTML COM FOTO EM TELA CHEIA
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

        def img_html(src: str, alt: str, vazio: str):
            if not src:
                return f'<div class="placeholder">{vazio}</div>'
            src_e = escape(src)
            return f'<img class="zoom-img" src="{src_e}" alt="{escape(alt)}" onclick="abrirFoto(\'{src_e}\')">'

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
                        {img_html(foto_ind, 'Foto do indivíduo', 'Sem foto')}
                    </div>
                    <div class="photo small">
                        <div class="photo-title">Foto do RG</div>
                        {img_html(foto_rg, 'Foto do RG', 'Sem RG')}
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
        .stats {{ display:flex; gap:14px; justify-content:center; flex-wrap:wrap; margin-top:18px; }}
        .stat {{ border:1px solid rgba(240,182,77,.55); background:rgba(8,19,35,.9); padding:10px 18px; border-radius:12px; min-width:150px; }}
        .stat span {{ display:block; font-size:12px; color:#bfc6d1; text-transform:uppercase; }}
        .stat b {{ font-size:24px; color:#fff; }}
        main {{ width:min(1380px, 96%); margin:24px auto 50px; display:grid; grid-template-columns:repeat(auto-fit, minmax(520px, 1fr)); gap:22px; }}
        .card {{ border:1px solid rgba(240,182,77,.55); background:linear-gradient(180deg, rgba(8,19,35,.97), rgba(3,8,16,.97)); box-shadow:0 0 0 1px rgba(255,255,255,.05), 0 18px 50px rgba(0,0,0,.32); border-radius:18px; overflow:hidden; }}
        .card.retirado {{ opacity:.65; filter:grayscale(.35); }}
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
        .zoom-img {{ cursor: zoom-in; transition: transform .15s ease, filter .15s ease; }}
        .zoom-img:hover {{ transform: scale(1.02); filter: brightness(1.08); }}
        .placeholder {{ height:240px; border:1px dashed #98a2b3; display:grid; place-items:center; color:#98a2b3; border-radius:10px; }}
        .photo.small .placeholder {{ height:125px; }}
        .info {{ display:grid; gap:10px; }}
        .linha, .box {{ background:#f6f7f9; color:#07111f; border-radius:12px; padding:10px 12px; border-left:5px solid #f0b64d; }}
        .linha b, .box b {{ display:block; color:#07111f; margin-bottom:3px; }}
        .linha p, .box p {{ margin:0; white-space:pre-wrap; line-height:1.35; }}
        .destaque {{ background:#fff4dc; }}
        .msg {{ display:inline-block; color:#f0b64d; text-decoration:none; font-weight:bold; margin-top:4px; }}
        .motivo {{ background:rgba(201,74,74,.18); color:#ffd7d7; border:1px solid rgba(201,74,74,.4); padding:10px; border-radius:10px; }}
        .vazio {{ grid-column:1 / -1; text-align:center; padding:60px 20px; border:1px dashed rgba(240,182,77,.6); border-radius:18px; background:rgba(8,19,35,.7); }}
        footer {{ text-align:center; color:#bfc6d1; border-top:1px solid rgba(240,182,77,.25); padding:18px; }}
        #lightbox {{ display:none; position:fixed; inset:0; background:rgba(0,0,0,.92); z-index:9999; align-items:center; justify-content:center; padding:22px; }}
        #lightbox.ativo {{ display:flex; }}
        #lightbox img {{ max-width:96vw; max-height:92vh; border-radius:12px; box-shadow:0 0 30px rgba(0,0,0,.7); }}
        #fecharLightbox {{ position:fixed; top:18px; right:22px; background:#f0b64d; color:#06111f; border:0; padding:10px 15px; border-radius:10px; font-weight:900; cursor:pointer; }}
        @media (max-width:720px) {{ main {{ grid-template-columns:1fr; }} .card-grid {{ grid-template-columns:1fr; }} .card-head {{ grid-template-columns:1fr; }} .photo img {{ height:330px; }} }}
    </style>
</head>
<body>
    <header>
        <h1>Catálogo de Procurados</h1>
        <p>DIC • Atualizado automaticamente pelo bot • Clique nas fotos para abrir em tela cheia</p>
        <div class="stats">
            <div class="stat"><span>Registros totais</span><b>{len(procurados)}</b></div>
            <div class="stat"><span>Ativos</span><b>{len(ativos)}</b></div>
            <div class="stat"><span>Retirados</span><b>{len(retirados)}</b></div>
            <div class="stat"><span>Última atualização</span><b style="font-size:16px">{agora_br()}</b></div>
        </div>
    </header>
    <main>{cards}</main>
    <footer>Uso exclusivo para GTA RP / sistema interno autorizado. Não use com dados reais de pessoas.</footer>

    <div id="lightbox" onclick="fecharFoto()">
        <button id="fecharLightbox" onclick="fecharFoto()">Fechar ✕</button>
        <img id="lightboxImg" src="" alt="Foto em tela cheia">
    </div>

    <script>
        function abrirFoto(src) {{
            document.getElementById('lightboxImg').src = src;
            document.getElementById('lightbox').classList.add('ativo');
        }}
        function fecharFoto() {{
            document.getElementById('lightbox').classList.remove('ativo');
            document.getElementById('lightboxImg').src = '';
        }}
        document.addEventListener('keydown', function(e) {{
            if (e.key === 'Escape') fecharFoto();
        }});
        document.getElementById('lightboxImg').addEventListener('click', function(e) {{ e.stopPropagation(); }});
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


async def start_web_server():
    gerar_catalogo_html()
    app = web.Application()
    app.router.add_get("/", pagina_inicial)
    app.router.add_get("/index.html", pagina_inicial)
    app.router.add_static("/uploads/", path=str(UPLOADS_DIR), show_index=False)
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

A Polícia DENARC de Capital Morada, por intermédio da **Divisão de Investigações Criminais (DIC)**, informa que o indivíduo abaixo encontra-se oficialmente procurado pelas autoridades competentes.

As investigações apontam seu envolvimento em atividades criminosas, havendo mandado ativo para sua localização, abordagem e condução para os procedimentos cabíveis.

📍 **ÚLTIMO AVISTAMENTO:** {registro.get('ultimo_avistamento', 'Não informado')}

⚠️ **CRIMES IMPUTADOS:**
{registro.get('crimes', 'Não informado')}

━━━━━━━━━━━━━━━━━━━━━━━

🆔 **IDENTIFICAÇÃO DO PROCURADO**

👤 **Nome:** {registro.get('nome', 'Não informado')}
🆔 **RG:** {registro.get('rg', 'Não informado')}

━━━━━━━━━━━━━━━━━━━━━━━

📞 Qualquer informação sobre o paradeiro deste indivíduo deverá ser repassada imediatamente a um agente da DENARC ou da DIC.

🔒 O sigilo do denunciante será integralmente preservado.

🔹 Polícia DENARC de Capital Morada
🔹 Divisão de Investigações Criminais (DIC)
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
            f"🚨 **Cadastro de Procurado — DIC**\n\n"
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

    if encontrado.get("mensagem_id") and PROCURADOS_CHANNEL_ID:
        canal = bot.get_channel(PROCURADOS_CHANNEL_ID)
        if canal:
            try:
                msg = await canal.fetch_message(int(encontrado["mensagem_id"]))
                texto_retirado = cortar_discord(
                    criar_texto_procurado(encontrado) + f"\n\n❌ **STATUS:** RETIRADO\n📌 **Motivo:** {motivo}",
                    1900,
                )
                await msg.edit(content=texto_retirado)
            except Exception:
                pass

    if HISTORICO_PROCURADOS_ID:
        historico = bot.get_channel(HISTORICO_PROCURADOS_ID)
        if historico:
            try:
                await historico.send(f"❌ **Procurado Retirado**\n👤 Nome: {encontrado.get('nome')}\n🪪 RG: {encontrado.get('rg')}\n📌 Motivo: {motivo}\n👮 Retirado por: {interaction.user.mention}")
            except Exception:
                pass

    await enviar_log(f"❌ Procurado retirado: {encontrado.get('nome')} | RG {encontrado.get('rg')} | Motivo: {motivo}")
    await interaction.response.send_message(f"✅ Procurado retirado e catálogo atualizado.\n🔗 {CATALOG_PUBLIC_URL}", ephemeral=True)


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
    if not dados.get("rg") or not dados.get("nome"):
        return None
    if procurar_por_rg(dados.get("rg", "")):
        return None

    imagens = []
    for a in msg.attachments:
        if (a.content_type and a.content_type.startswith("image/")) or Path(a.filename).suffix.lower() in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
            imagens.append(a)

    foto_ind = await salvar_anexo_publico(imagens[0], f"sync-ind-{dados.get('rg','')}") if len(imagens) >= 1 else ""
    foto_rg = await salvar_anexo_publico(imagens[1], f"sync-rg-{dados.get('rg','')}") if len(imagens) >= 2 else ""

    return {
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
    await interaction.followup.send(f"✅ Sincronização finalizada.\nMensagens analisadas: `{analisados}`\nProcurados importados: `{importados}`\nCatálogo: {CATALOG_PUBLIC_URL}", ephemeral=True)

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

    # Criar o canal + mandar todos os tópicos pode passar de 3 segundos.
    # Por isso o bot responde/defer primeiro e só depois cria a mesa.
    if not interaction.response.is_done():
        await interaction.response.defer(ephemeral=True, thinking=True)

    categoria = guild.get_channel(CATEGORIA_MESAS_ABERTAS_ID) if CATEGORIA_MESAS_ABERTAS_ID else None
    overwrites = cargos_equipe_permissoes(guild)
    overwrites[interaction.user] = discord.PermissionOverwrite(view_channel=True, send_messages=True, attach_files=True, read_message_history=True)

    nome_canal = f"🕵️┃{slugify(apelido)}-{slugify(familia)}"
    canal = await guild.create_text_channel(name=nome_canal, category=categoria, overwrites=overwrites)

    # Mesa em mensagens normais, com todos os tópicos já abertos no canal.
    await canal.send(
        f"🕵️ **Mesa de Investigação — {familia}**\n"
        f"👮 **Agente:** {interaction.user.mention}\n"
        f"📛 **Apelido:** {apelido}\n"
        f"🏷️ **Organização/Família:** {familia}\n"
        f"🕒 **Criada em:** {agora_br()}\n"
        f"📝 **Observação:** {observacao or 'Nenhuma'}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📌 **Preencha as informações diretamente abaixo de cada tópico.**"
    )

    for topico in TOPICOS_MESA:
        await canal.send(
            f"## {topico}\n"
            f"> Preencha as informações deste tópico abaixo.\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━"
        )

    await canal.send("🔒 Para encerrar esta mesa, clique no botão abaixo.", view=FecharMesaView())

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
    })

    await enviar_log(f"➕ Mesa criada: {canal.mention} | Família: {familia} | Por: {interaction.user.mention}")
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
    ativos = len([p for p in procurados if p.get("status", "A PROCURAR") == "A PROCURAR"])
    retirados = len([p for p in procurados if p.get("status") == "RETIRADO"])
    abertas = len([m for m in mesas if m.get("status") == "ABERTA"])
    fechadas = len([m for m in mesas if m.get("status") == "FECHADA"])
    await interaction.response.send_message(
        f"📊 **Estatísticas DIC**\n\n"
        f"🚨 Procurados totais: `{len(procurados)}`\n"
        f"🟢 Procurados ativos: `{ativos}`\n"
        f"🔴 Procurados retirados: `{retirados}`\n\n"
        f"🕵️ Mesas totais: `{len(mesas)}`\n"
        f"🟢 Mesas abertas: `{abertas}`\n"
        f"🔒 Mesas fechadas: `{fechadas}`\n\n"
        f"📄 Catálogo: {CATALOG_PUBLIC_URL}",
        ephemeral=True,
    )



# =====================================================
# COMANDOS DE TEXTO (!) - PLANO B PARA QUANDO O DISCORD NÃO MOSTRAR SLASH
# =====================================================

def embed_painel_procurados_padrao() -> discord.Embed:
    return discord.Embed(
        title="🚨 Sistema de Procurados - DIC",
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
        title="🕵️ Sistema de Mesas - DIC",
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
    )

    nome_canal = f"🕵️┃{slugify(apelido)}-{slugify(familia)}"
    canal = await guild.create_text_channel(name=nome_canal, category=categoria, overwrites=overwrites)

    await canal.send(
        f"🕵️ **Mesa de Investigação — {familia}**\n"
        f"👮 **Agente:** {ctx.author.mention}\n"
        f"📛 **Apelido:** {apelido}\n"
        f"🏷️ **Organização/Família:** {familia}\n"
        f"🕒 **Criada em:** {agora_br()}\n"
        f"📝 **Observação:** {observacao or 'Nenhuma'}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📌 **Preencha as informações diretamente abaixo de cada tópico.**"
    )

    for topico in TOPICOS_MESA:
        await canal.send(
            f"## {topico}\n"
            f"> Preencha as informações deste tópico abaixo.\n\n"
            f"━━━━━━━━━━━━━━━━━━━━━━━"
        )

    await canal.send("🔒 Para encerrar esta mesa, clique no botão abaixo.", view=FecharMesaView())

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
    })

    await enviar_log(f"➕ Mesa criada por comando de texto: {canal.mention} | Família: {familia} | Por: {ctx.author.mention}")
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
        "**Slash:** `/painelprocurados`, `/painelmesas`, `/criarmesa`, `/fecharmesa`, `/backup`, `/estatisticas`\n"
        "**Plano B por texto:** `!painelprocurados`, `!painelmesas`, `!criarmesa Apelido | Família | Observação`, `!fecharmesa`, `!catalogo`"
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

    bot.add_view(PainelProcuradosView())
    bot.add_view(FinalizarProcuradoView())
    bot.add_view(PainelMesasView())
    bot.add_view(FecharMesaView())

    if comandos_ja_sincronizados:
        print(f"Bot online como {bot.user}")
        return

    try:
        if GUILD_ID:
            guild_obj = discord.Object(id=GUILD_ID)

            # 1) Copia todos os comandos definidos no código para o servidor correto.
            bot.tree.copy_global_to(guild=guild_obj)
            comandos_servidor = await bot.tree.sync(guild=guild_obj)

            # 2) Limpa comandos globais antigos para tirar duplicados do Discord.
            bot.tree.clear_commands(guild=None)
            comandos_globais = await bot.tree.sync()

            print(f"Comandos sincronizados no servidor: {len(comandos_servidor)}")
            print(f"Comandos globais antigos limpos: {len(comandos_globais)}")
            print("Comandos ativos:", ", ".join(f"/{cmd.name}" for cmd in comandos_servidor))
        else:
            comandos_globais = await bot.tree.sync()
            print(f"Comandos sincronizados globalmente: {len(comandos_globais)}")
            print("AVISO: GUILD_ID está 0. Os comandos podem demorar para aparecer no Discord.")

        comandos_ja_sincronizados = True
    except Exception as e:
        print(f"Erro ao sincronizar comandos: {e}")

    print('VERSAO FINAL RESOLVIDA: procurados-texto-normal | mesas-categorias-fixas | catalogo-lightbox')
    print(f"Bot online como {bot.user}")

# =====================================================
# MAIN
# =====================================================

async def main():
    if not DISCORD_TOKEN or DISCORD_TOKEN == "COLE_O_TOKEN_DO_BOT_AQUI":
        print("ERRO: Coloque o token correto nas variáveis do Railway ou no arquivo .env.")
        return
    await start_web_server()
    await bot.start(DISCORD_TOKEN)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("Bot desligado.")
