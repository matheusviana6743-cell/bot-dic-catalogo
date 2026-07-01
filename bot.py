import os
import re
import html
import json
import uuid
import shutil
import asyncio
import datetime
import unicodedata
from pathlib import Path
from zoneinfo import ZoneInfo

import discord
from discord import app_commands
from discord.ext import commands
from discord.ui import Modal, TextInput, View, Button
from aiohttp import web
from dotenv import load_dotenv


# =====================================================
# CONFIGURAÇÃO
# =====================================================

load_dotenv()

DISCORD_TOKEN = os.getenv("DISCORD_TOKEN", "").strip()
GUILD_ID = int(os.getenv("GUILD_ID", "0") or 0)
PROCURADOS_CHANNEL_ID = int(os.getenv("PROCURADOS_CHANNEL_ID", "0") or 0)
CATALOG_PUBLIC_URL = os.getenv("CATALOG_PUBLIC_URL", "http://127.0.0.1:8000/").strip()
PORT = int(os.getenv("PORT", "8000") or 8000)

BASE_DIR = Path(__file__).parent
PUBLIC_DIR = BASE_DIR / "public"
UPLOADS_DIR = PUBLIC_DIR / "uploads"
DATA_DIR = BASE_DIR / "data"
CATALOGO_HTML = PUBLIC_DIR / "index.html"
CATALOGO_JSON = DATA_DIR / "procurados.json"

PUBLIC_DIR.mkdir(exist_ok=True)
UPLOADS_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

FUSO_BR = ZoneInfo("America/Sao_Paulo")


# =====================================================
# FUNÇÕES BÁSICAS
# =====================================================

def agora_br():
    return datetime.datetime.now(FUSO_BR)


def data_hora_br():
    return agora_br().strftime("%d/%m/%Y %H:%M")


def data_caso():
    return agora_br().strftime("%Y%m%d-%H%M%S")


def somente_numero(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def slug(texto: str) -> str:
    texto = str(texto or "")
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r"[^a-zA-Z0-9_-]+", "-", texto).strip("-").lower()
    return texto or "sem-nome"


def esc(texto) -> str:
    return html.escape(str(texto or ""))


def esc_br(texto) -> str:
    return esc(texto).replace("\n", "<br>")


def carregar_procurados():
    if not CATALOGO_JSON.exists():
        return []
    try:
        with open(CATALOGO_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def salvar_procurados(lista):
    with open(CATALOGO_JSON, "w", encoding="utf-8") as f:
        json.dump(lista, f, ensure_ascii=False, indent=2)


def procurar_por_rg(rg: str):
    rg_limpo = somente_numero(rg)
    for p in carregar_procurados():
        if somente_numero(p.get("rg", "")) == rg_limpo:
            return p
    return None


def upsert_procurado(dados: dict):
    lista = carregar_procurados()
    rg_limpo = somente_numero(dados.get("rg", ""))

    atualizado = False
    for i, p in enumerate(lista):
        if somente_numero(p.get("rg", "")) == rg_limpo and rg_limpo:
            lista[i].update(dados)
            atualizado = True
            break

    if not atualizado:
        lista.append(dados)

    salvar_procurados(lista)
    gerar_catalogo_html()


def marcar_retirado(rg: str, motivo: str, responsavel: str):
    lista = carregar_procurados()
    rg_limpo = somente_numero(rg)
    achou = False

    for p in lista:
        if somente_numero(p.get("rg", "")) == rg_limpo:
            p["status"] = "RETIRADO"
            p["motivo_retirada"] = motivo
            p["retirado_por"] = responsavel
            p["data_retirada"] = data_hora_br()
            achou = True
            break

    salvar_procurados(lista)
    gerar_catalogo_html()
    return achou


def eh_imagem(attachment: discord.Attachment):
    nome = attachment.filename.lower()
    return nome.endswith((".png", ".jpg", ".jpeg", ".webp", ".gif"))


async def salvar_attachment(attachment: discord.Attachment, prefixo: str, rg: str):
    extensao = os.path.splitext(attachment.filename)[1].lower()
    if extensao not in [".png", ".jpg", ".jpeg", ".webp", ".gif"]:
        extensao = ".png"

    nome_arquivo = f"{slug(rg)}-{prefixo}-{uuid.uuid4().hex[:8]}{extensao}"
    caminho = UPLOADS_DIR / nome_arquivo

    await attachment.save(caminho)
    return f"uploads/{nome_arquivo}", caminho


# =====================================================
# CATÁLOGO HTML
# =====================================================

def gerar_catalogo_html():
    procurados = carregar_procurados()
    ativos = [p for p in procurados if p.get("status", "A PROCURAR") == "A PROCURAR"]
    retirados = [p for p in procurados if p.get("status") == "RETIRADO"]

    cards = ""

    if not procurados:
        cards = """
        <div class="vazio">
            <h2>Nenhum procurado cadastrado ainda.</h2>
            <p>Quando o bot cadastrar um procurado, ele aparecerá automaticamente aqui.</p>
        </div>
        """

    for p in procurados:
        status = p.get("status", "A PROCURAR")
        status_class = "retirado" if status == "RETIRADO" else "ativo"

        foto_ind = p.get("foto_individuo") or ""
        foto_rg = p.get("foto_rg") or ""

        if foto_ind:
            foto_ind_html = f'<img src="{esc(foto_ind)}" alt="Foto do indivíduo">'
        else:
            foto_ind_html = '<div class="placeholder">SEM FOTO</div>'

        if foto_rg:
            foto_rg_html = f'<img src="{esc(foto_rg)}" alt="Foto do RG">'
        else:
            foto_rg_html = '<div class="placeholder">SEM RG</div>'

        motivo = ""
        if status == "RETIRADO":
            motivo = f"""
            <div class="campo retirada">
                <b>Motivo da Retirada</b>
                <p>{esc_br(p.get("motivo_retirada", ""))}</p>
                <small>Retirado por {esc(p.get("retirado_por", ""))} em {esc(p.get("data_retirada", ""))}</small>
            </div>
            """

        link_msg = ""
        if p.get("discord_jump_url"):
            link_msg = f'<a class="link-discord" href="{esc(p.get("discord_jump_url"))}" target="_blank">Abrir postagem no Discord</a>'

        cards += f"""
        <section class="card">
            <div class="card-head">
                <div>
                    <span>Nº DO CASO</span>
                    <strong>{esc(p.get("caso", ""))}</strong>
                </div>
                <div>
                    <span>CADASTRO</span>
                    <strong>{esc(p.get("data_cadastro", ""))}</strong>
                </div>
                <div class="status {status_class}">
                    {esc(status)}
                </div>
            </div>

            <div class="card-body">
                <div class="fotos">
                    <div class="foto">
                        <h3>Foto do Indivíduo</h3>
                        {foto_ind_html}
                    </div>

                    <div class="foto rg">
                        <h3>Foto do RG</h3>
                        {foto_rg_html}
                    </div>
                </div>

                <div class="dados">
                    <div class="linha">
                        <b>Nome</b>
                        <p>{esc(p.get("nome", ""))}</p>
                    </div>

                    <div class="linha">
                        <b>RG</b>
                        <p>{esc(p.get("rg", ""))}</p>
                    </div>

                    <div class="campo">
                        <b>Crimes Cometidos</b>
                        <p>{esc_br(p.get("crimes_cometidos", ""))}</p>
                    </div>

                    <div class="campo destaque">
                        <b>Último Avistamento</b>
                        <p>{esc_br(p.get("ultimo_avistamento", ""))}</p>
                    </div>

                    <div class="campo">
                        <b>Informações</b>
                        <p>{esc_br(p.get("informacoes", ""))}</p>
                    </div>

                    {motivo}
                    {link_msg}
                </div>
            </div>
        </section>
        """

    html_final = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Catálogo de Procurados - DIC</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<style>
* {{
    box-sizing: border-box;
}}

body {{
    margin: 0;
    background: radial-gradient(circle at top, #10213b 0%, #050b14 55%, #02050a 100%);
    color: #fff;
    font-family: Arial, Helvetica, sans-serif;
}}

header {{
    padding: 34px 18px 20px;
    text-align: center;
    border-bottom: 2px solid #c8952d;
    background: linear-gradient(90deg, #050b14, #0c1b31, #050b14);
}}

header h1 {{
    margin: 0;
    font-size: clamp(30px, 5vw, 64px);
    letter-spacing: 2px;
    color: #f3c45d;
    text-transform: uppercase;
}}

header p {{
    margin: 8px 0 0;
    color: #d9d9d9;
    font-size: 16px;
}}

.stats {{
    display: flex;
    justify-content: center;
    gap: 15px;
    flex-wrap: wrap;
    padding: 16px;
    background: rgba(0,0,0,.30);
    border-bottom: 1px solid rgba(200,149,45,.45);
}}

.stat {{
    min-width: 170px;
    padding: 12px 16px;
    border: 1px solid rgba(200,149,45,.55);
    border-radius: 12px;
    background: rgba(7,17,31,.85);
}}

.stat small {{
    display: block;
    color: #f3c45d;
    font-weight: 700;
    text-transform: uppercase;
    font-size: 12px;
}}

.stat strong {{
    display: block;
    margin-top: 6px;
    font-size: 22px;
}}

main {{
    max-width: 1250px;
    margin: 0 auto;
    padding: 24px 14px 40px;
}}

.card {{
    margin-bottom: 24px;
    border: 1px solid rgba(200,149,45,.72);
    border-radius: 16px;
    overflow: hidden;
    background: rgba(7,17,31,.90);
    box-shadow: 0 12px 28px rgba(0,0,0,.35);
}}

.card-head {{
    display: grid;
    grid-template-columns: 1fr 1fr auto;
    gap: 10px;
    padding: 12px 14px;
    background: #07111f;
    border-bottom: 1px solid rgba(200,149,45,.55);
    align-items: center;
}}

.card-head span {{
    display: block;
    color: #f3c45d;
    font-size: 11px;
    font-weight: 700;
}}

.card-head strong {{
    font-size: 14px;
}}

.status {{
    padding: 8px 12px;
    border-radius: 999px;
    font-weight: 800;
    font-size: 13px;
    border: 1px solid #f3c45d;
    color: #f3c45d;
}}

.status.retirado {{
    color: #ff7070;
    border-color: #ff7070;
}}

.card-body {{
    display: grid;
    grid-template-columns: 310px 1fr;
    gap: 16px;
    padding: 16px;
}}

.fotos {{
    display: grid;
    gap: 14px;
}}

.foto {{
    border: 1px solid rgba(200,149,45,.70);
    border-radius: 12px;
    padding: 10px;
    background: #030812;
}}

.foto h3 {{
    margin: 0 0 8px;
    color: #f3c45d;
    text-align: center;
    font-size: 14px;
    text-transform: uppercase;
}}

.foto img {{
    width: 100%;
    max-height: 320px;
    object-fit: cover;
    border-radius: 9px;
    border: 1px solid rgba(255,255,255,.15);
    background: #111;
}}

.foto.rg img {{
    max-height: 185px;
    object-fit: contain;
    background: #fff;
}}

.placeholder {{
    height: 180px;
    display: grid;
    place-items: center;
    color: #888;
    border: 1px dashed #777;
    border-radius: 9px;
}}

.dados {{
    display: grid;
    gap: 10px;
}}

.linha, .campo {{
    padding: 12px;
    border: 1px solid rgba(255,255,255,.12);
    border-radius: 10px;
    background: rgba(255,255,255,.055);
}}

.linha b, .campo b {{
    display: block;
    color: #f3c45d;
    margin-bottom: 5px;
}}

.linha p, .campo p {{
    margin: 0;
    line-height: 1.45;
    white-space: normal;
}}

.destaque {{
    border-color: rgba(243,196,93,.65);
    background: rgba(243,196,93,.08);
}}

.retirada {{
    border-color: rgba(255,112,112,.60);
    background: rgba(255,112,112,.08);
}}

.retirada b {{
    color: #ff9090;
}}

.link-discord {{
    display: inline-block;
    width: fit-content;
    padding: 10px 12px;
    border-radius: 9px;
    background: #5865f2;
    color: white;
    text-decoration: none;
    font-weight: 700;
}}

.vazio {{
    text-align: center;
    padding: 60px 15px;
    border: 1px dashed rgba(200,149,45,.6);
    border-radius: 14px;
    background: rgba(7,17,31,.75);
}}

footer {{
    text-align: center;
    padding: 22px;
    color: #bdbdbd;
    border-top: 1px solid rgba(200,149,45,.45);
    background: #050b14;
}}

@media (max-width: 800px) {{
    .card-body {{
        grid-template-columns: 1fr;
    }}

    .card-head {{
        grid-template-columns: 1fr;
    }}

    .foto img {{
        max-height: 420px;
    }}
}}
</style>
</head>

<body>
<header>
    <h1>Catálogo de Procurados</h1>
    <p>Documento automático da DIC • Atualizado pelo bot</p>
</header>

<section class="stats">
    <div class="stat">
        <small>Registros totais</small>
        <strong>{len(procurados)}</strong>
    </div>

    <div class="stat">
        <small>A procurar</small>
        <strong>{len(ativos)}</strong>
    </div>

    <div class="stat">
        <small>Retirados</small>
        <strong>{len(retirados)}</strong>
    </div>

    <div class="stat">
        <small>Última atualização</small>
        <strong>{esc(data_hora_br())}</strong>
    </div>
</section>

<main>
    {cards}
</main>

<footer>
    🔹 Polícia DENARC de Capital Morada • Divisão de Investigações Criminais (DIC)
</footer>
</body>
</html>
"""

    with open(CATALOGO_HTML, "w", encoding="utf-8") as f:
        f.write(html_final)


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

    print(f"Catálogo rodando na porta {PORT}")


# =====================================================
# MENSAGEM DO PROCURADO NO DISCORD
# =====================================================

def montar_descricao_mandado(nome, rg, crimes_cometidos, ultimo_avistamento, informacoes=""):
    extra_info = ""
    if informacoes and informacoes.strip() and informacoes.strip() != "-":
        extra_info = f"""
━━━━━━━━━━━━━━━━━━━━━━━

📌 **INFORMAÇÕES COMPLEMENTARES**

{informacoes}
"""

    return f"""🚨 **MANDADO DE PRISÃO E PROCURAÇÃO INVESTIGATIVA** 🚨

A Polícia DENARC de Capital Morada, por intermédio da **Divisão de Investigações Criminais (DIC)**, informa que o indivíduo abaixo encontra-se oficialmente procurado pelas autoridades competentes.

As investigações apontam seu envolvimento em atividades criminosas, havendo mandado ativo para sua localização, abordagem e condução para os procedimentos cabíveis.

📍 **ÚLTIMO AVISTAMENTO:** {ultimo_avistamento}

⚠️ **CRIMES IMPUTADOS:**
{crimes_cometidos}

━━━━━━━━━━━━━━━━━━━━━━━

🆔 **IDENTIFICAÇÃO DO PROCURADO**

👤 **Nome:** {nome}
🆔 **RG:** {rg}
{extra_info}
━━━━━━━━━━━━━━━━━━━━━━━

📞 Qualquer informação sobre o paradeiro deste indivíduo deverá ser repassada imediatamente a um agente da DENARC ou da DIC.

🔒 O sigilo do denunciante será integralmente preservado.

🔹 Polícia DENARC de Capital Morada
🔹 Divisão de Investigações Criminais (DIC)"""


def montar_embed_mandado(nome, rg, crimes, ultimo, informacoes, caso):
    embed = discord.Embed(
        description=montar_descricao_mandado(
            nome=nome,
            rg=rg,
            crimes_cometidos=crimes,
            ultimo_avistamento=ultimo,
            informacoes=informacoes
        ),
        color=discord.Color.red()
    )
    embed.set_footer(text=f"Caso: {caso} • Cadastro: {data_hora_br()}")
    return embed


# =====================================================
# MODAIS E BOTÕES
# =====================================================

class NovoProcuradoModal(Modal, title="Cadastrar Novo Procurado"):
    nome = TextInput(
        label="Nome do procurado",
        placeholder="Ex: RODRIGUES ARTHUR",
        max_length=100
    )

    rg = TextInput(
        label="RG",
        placeholder="Ex: 24348",
        max_length=40
    )

    crimes_cometidos = TextInput(
        label="Crimes cometidos",
        placeholder="Ex: Art. 2.2.1 - Sequestro de Oficial\nArt. 8.3 - Formação de Quadrilha",
        style=discord.TextStyle.paragraph,
        max_length=1000
    )

    ultimo_avistamento = TextInput(
        label="Último avistamento",
        placeholder="Ex: Caixa d'água",
        style=discord.TextStyle.paragraph,
        max_length=500
    )

    informacoes = TextInput(
        label="Informações extras",
        placeholder="Ex: Suspeito de alta periculosidade...",
        style=discord.TextStyle.paragraph,
        required=False,
        max_length=1000
    )

    async def on_submit(self, interaction: discord.Interaction):
        guild = interaction.guild

        overwrites = {
            guild.default_role: discord.PermissionOverwrite(view_channel=False),
            interaction.user: discord.PermissionOverwrite(
                view_channel=True,
                send_messages=True,
                attach_files=True,
                read_message_history=True
            ),
            guild.me: discord.PermissionOverwrite(
                view_channel=True,
                send_messages=True,
                manage_channels=True,
                attach_files=True,
                read_message_history=True
            )
        }

        nome_canal = f"procurado-{slug(self.nome.value)}-{somente_numero(self.rg.value) or 'sem-rg'}"
        canal = await guild.create_text_channel(
            name=nome_canal[:90],
            overwrites=overwrites,
            reason=f"Cadastro provisório de procurado por {interaction.user}"
        )

        view = FinalizarProcuradoView(
            nome=self.nome.value,
            rg=self.rg.value,
            crimes=str(self.crimes_cometidos.value),
            ultimo_avistamento=str(self.ultimo_avistamento.value),
            informacoes=str(self.informacoes.value or ""),
            autor_id=interaction.user.id,
            autor_nome=str(interaction.user)
        )

        await canal.send(
            f"🚨 **Cadastro de Procurado — DIC**\n\n"
            f"👤 **Nome:** {self.nome.value}\n"
            f"🆔 **RG:** {self.rg.value}\n\n"
            f"📸 **Envie 2 imagens neste canal:**\n"
            f"1️⃣ **Foto do indivíduo**\n"
            f"2️⃣ **Foto do RG / identificação**\n\n"
            f"Depois de enviar as duas imagens, clique em **✅ Finalizar Cadastro**.\n\n"
            f"⚠️ Este canal é provisório e pode ser apagado depois.",
            view=view
        )

        await interaction.response.send_message(
            f"✅ Canal provisório criado: {canal.mention}\nEnvie as duas fotos lá e finalize o cadastro.",
            ephemeral=True
        )


class FinalizarProcuradoView(View):
    def __init__(self, nome, rg, crimes, ultimo_avistamento, informacoes, autor_id, autor_nome):
        super().__init__(timeout=None)
        self.nome = nome
        self.rg = rg
        self.crimes = crimes
        self.ultimo_avistamento = ultimo_avistamento
        self.informacoes = informacoes
        self.autor_id = autor_id
        self.autor_nome = autor_nome

    @discord.ui.button(
        label="Finalizar Cadastro",
        emoji="✅",
        style=discord.ButtonStyle.success
    )
    async def finalizar(self, interaction: discord.Interaction, button: Button):
        if interaction.user.id != self.autor_id and not interaction.user.guild_permissions.manage_guild:
            await interaction.response.send_message("❌ Apenas quem iniciou o cadastro ou a administração pode finalizar.", ephemeral=True)
            return

        await interaction.response.defer(ephemeral=True, thinking=True)

        anexos = []
        async for msg in interaction.channel.history(limit=100, oldest_first=True):
            for anexo in msg.attachments:
                if eh_imagem(anexo):
                    anexos.append(anexo)

        if len(anexos) < 2:
            await interaction.followup.send(
                "❌ Envie as duas imagens antes de finalizar: **foto do indivíduo** e **foto do RG**.",
                ephemeral=True
            )
            return

        foto_ind_url, foto_ind_path = await salvar_attachment(anexos[0], "individuo", self.rg)
        foto_rg_url, foto_rg_path = await salvar_attachment(anexos[1], "rg", self.rg)

        caso = f"DIC-{data_caso()}"

        canal_oficial = interaction.guild.get_channel(PROCURADOS_CHANNEL_ID)
        if not canal_oficial:
            await interaction.followup.send(
                "❌ Canal oficial de procurados não encontrado. Confira `PROCURADOS_CHANNEL_ID` no `.env`.",
                ephemeral=True
            )
            return

        embed = montar_embed_mandado(
            nome=self.nome,
            rg=self.rg,
            crimes=self.crimes,
            ultimo=self.ultimo_avistamento,
            informacoes=self.informacoes,
            caso=caso
        )

        arquivos = [
            discord.File(str(foto_ind_path), filename=f"foto_individuo_{slug(self.rg)}{foto_ind_path.suffix}"),
            discord.File(str(foto_rg_path), filename=f"foto_rg_{slug(self.rg)}{foto_rg_path.suffix}")
        ]

        mensagem = await canal_oficial.send(files=arquivos, embed=embed)

        dados = {
            "id": uuid.uuid4().hex,
            "caso": caso,
            "nome": self.nome,
            "rg": self.rg,
            "crimes_cometidos": self.crimes,
            "ultimo_avistamento": self.ultimo_avistamento,
            "informacoes": self.informacoes,
            "foto_individuo": foto_ind_url,
            "foto_rg": foto_rg_url,
            "status": "A PROCURAR",
            "data_cadastro": data_hora_br(),
            "autor_id": self.autor_id,
            "autor_nome": self.autor_nome,
            "discord_message_id": mensagem.id,
            "discord_jump_url": mensagem.jump_url
        }

        upsert_procurado(dados)

        await interaction.followup.send(
            f"✅ Procurado cadastrado com sucesso!\n"
            f"📌 Postagem: {mensagem.jump_url}\n"
            f"📄 Catálogo: {CATALOG_PUBLIC_URL}",
            ephemeral=True
        )

        try:
            await interaction.channel.send("✅ Cadastro finalizado. Este canal será apagado em 10 segundos.")
            await asyncio.sleep(10)
            await interaction.channel.delete(reason="Cadastro de procurado finalizado")
        except Exception:
            pass

    @discord.ui.button(
        label="Cancelar Cadastro",
        emoji="🗑️",
        style=discord.ButtonStyle.danger
    )
    async def cancelar(self, interaction: discord.Interaction, button: Button):
        if interaction.user.id != self.autor_id and not interaction.user.guild_permissions.manage_guild:
            await interaction.response.send_message("❌ Apenas quem iniciou o cadastro ou a administração pode cancelar.", ephemeral=True)
            return

        await interaction.response.send_message("🗑️ Cadastro cancelado. Canal será apagado em 5 segundos.", ephemeral=True)
        await asyncio.sleep(5)
        try:
            await interaction.channel.delete(reason="Cadastro de procurado cancelado")
        except Exception:
            pass


class RetirarProcuradoModal(Modal, title="Retirar Procurado"):
    rg = TextInput(
        label="RG do procurado",
        placeholder="Ex: 24348",
        max_length=40
    )

    motivo = TextInput(
        label="Motivo da retirada",
        placeholder="Ex: Capturado / Mandado encerrado / Erro no cadastro",
        style=discord.TextStyle.paragraph,
        max_length=500
    )

    async def on_submit(self, interaction: discord.Interaction):
        ok = marcar_retirado(self.rg.value, self.motivo.value, str(interaction.user))

        if ok:
            await interaction.response.send_message(
                f"✅ Procurado de RG `{self.rg.value}` foi marcado como **RETIRADO** no catálogo.\n{CATALOG_PUBLIC_URL}",
                ephemeral=True
            )
        else:
            await interaction.response.send_message(
                f"❌ Nenhum procurado encontrado com o RG `{self.rg.value}`.",
                ephemeral=True
            )


class PainelProcuradosView(View):
    def __init__(self):
        super().__init__(timeout=None)

    @discord.ui.button(
        label="Novo Procurado",
        emoji="➕",
        style=discord.ButtonStyle.danger,
        custom_id="dic_procurados_novo"
    )
    async def novo_procurado(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(NovoProcuradoModal())

    @discord.ui.button(
        label="Lista de Procurados",
        emoji="📋",
        style=discord.ButtonStyle.primary,
        custom_id="dic_procurados_lista"
    )
    async def lista_procurados(self, interaction: discord.Interaction, button: Button):
        procurados = carregar_procurados()
        ativos = [p for p in procurados if p.get("status", "A PROCURAR") == "A PROCURAR"]

        if not ativos:
            texto = "📋 Não há procurados ativos no momento."
        else:
            linhas = []
            for p in ativos[:20]:
                linhas.append(f"• **{p.get('nome', 'Sem nome')}** — RG `{p.get('rg', 'Sem RG')}`")
            texto = "\n".join(linhas)

        await interaction.response.send_message(
            f"📋 **Procurados Ativos:**\n{texto}\n\n📄 **Catálogo:** {CATALOG_PUBLIC_URL}",
            ephemeral=True
        )

    @discord.ui.button(
        label="Retirar Procurado",
        emoji="❌",
        style=discord.ButtonStyle.secondary,
        custom_id="dic_procurados_retirar"
    )
    async def retirar_procurado(self, interaction: discord.Interaction, button: Button):
        await interaction.response.send_modal(RetirarProcuradoModal())


# =====================================================
# BOT
# =====================================================

intents = discord.Intents.default()
intents.message_content = True
intents.members = True
intents.guilds = True


class DICBot(commands.Bot):
    def __init__(self):
        super().__init__(command_prefix="!", intents=intents)

    async def setup_hook(self):
        self.add_view(PainelProcuradosView())

        if GUILD_ID:
            guild_obj = discord.Object(id=GUILD_ID)
            self.tree.copy_global_to(guild=guild_obj)
            await self.tree.sync(guild=guild_obj)
            print("Comandos sincronizados no servidor.")
        else:
            await self.tree.sync()
            print("Comandos sincronizados globalmente.")


bot = DICBot()


@bot.event
async def on_ready():
    print(f"Bot online como {bot.user}")


# =====================================================
# COMANDOS
# =====================================================

@bot.tree.command(name="painelprocurados", description="Envia o painel de procurados da DIC.")
@app_commands.default_permissions(manage_guild=True)
async def painelprocurados(interaction: discord.Interaction):
    embed = discord.Embed(
        title="🚨 Sistema de Procurados - DIC",
        description=(
            "Utilize os botões abaixo para gerenciar procurados.\n\n"
            "➕ **Novo Procurado** — Cadastrar um novo procurado.\n"
            "📋 **Lista de Procurados** — Ver procurados ativos.\n"
            "❌ **Retirar Procurado** — Retirar um procurado pelo RG."
        ),
        color=discord.Color.red()
    )

    await interaction.response.send_message(embed=embed, view=PainelProcuradosView())


@bot.tree.command(name="catalogo", description="Mostra o link do catálogo de procurados.")
async def catalogo(interaction: discord.Interaction):
    gerar_catalogo_html()
    await interaction.response.send_message(
        f"📄 **Catálogo de Procurados - DIC:**\n{CATALOG_PUBLIC_URL}",
        ephemeral=True
    )


@bot.tree.command(name="regerarcatalogo", description="Regera o HTML do catálogo.")
@app_commands.default_permissions(manage_guild=True)
async def regerarcatalogo(interaction: discord.Interaction):
    gerar_catalogo_html()
    await interaction.response.send_message(
        f"✅ Catálogo regenerado.\n{CATALOG_PUBLIC_URL}",
        ephemeral=True
    )


@bot.tree.command(name="retirarprocurado", description="Retira um procurado pelo RG.")
@app_commands.describe(rg="RG do procurado", motivo="Motivo da retirada")
@app_commands.default_permissions(manage_guild=True)
async def retirarprocurado(interaction: discord.Interaction, rg: str, motivo: str):
    ok = marcar_retirado(rg, motivo, str(interaction.user))

    if ok:
        await interaction.response.send_message(
            f"✅ Procurado de RG `{rg}` marcado como **RETIRADO** no catálogo.\n{CATALOG_PUBLIC_URL}",
            ephemeral=True
        )
    else:
        await interaction.response.send_message(
            f"❌ Nenhum procurado encontrado com o RG `{rg}`.",
            ephemeral=True
        )


def extrair_linha(desc: str, campo: str):
    padrao = rf"{campo}:\*\*\s*(.+)"
    m = re.search(padrao, desc, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return ""


def extrair_bloco(desc: str, inicio: str, fim_regex: str):
    padrao = rf"{inicio}:\*\*\s*(.*?)(?:{fim_regex})"
    m = re.search(padrao, desc, flags=re.IGNORECASE | re.DOTALL)
    if m:
        return m.group(1).strip()
    return ""


@bot.tree.command(name="sincronizarcatalogo", description="Tenta puxar procurados antigos do canal oficial para o catálogo.")
@app_commands.describe(limite="Quantidade de mensagens para analisar. Ex: 200")
@app_commands.default_permissions(manage_guild=True)
async def sincronizarcatalogo(interaction: discord.Interaction, limite: int = 200):
    await interaction.response.defer(ephemeral=True, thinking=True)

    canal = interaction.guild.get_channel(PROCURADOS_CHANNEL_ID)
    if not canal:
        await interaction.followup.send("❌ Canal de procurados não encontrado.", ephemeral=True)
        return

    existentes = carregar_procurados()
    rgs_existentes = {somente_numero(p.get("rg", "")) for p in existentes}

    adicionados = 0
    ultimos_anexos = []

    async for msg in canal.history(limit=limite, oldest_first=True):
        imagens_msg = [a for a in msg.attachments if eh_imagem(a)]
        if imagens_msg:
            ultimos_anexos = imagens_msg

        for embed in msg.embeds:
            desc = embed.description or ""

            if "MANDADO" not in desc.upper() or "Nome:" not in desc:
                continue

            nome = extrair_linha(desc, "Nome")
            rg = extrair_linha(desc, "RG")
            ultimo = extrair_linha(desc, "ÚLTIMO AVISTAMENTO") or extrair_linha(desc, "Ultimo Avistamento")
            crimes = extrair_bloco(desc, "CRIMES IMPUTADOS", r"━━━━━━━━|🆔")
            informacoes = extrair_bloco(desc, "INFORMAÇÕES COMPLEMENTARES", r"━━━━━━━━|📞")

            if not nome or not rg:
                continue

            rg_limpo = somente_numero(rg)
            if rg_limpo in rgs_existentes:
                continue

            foto_ind_url = ""
            foto_rg_url = ""

            anexos_usar = imagens_msg or ultimos_anexos
            if len(anexos_usar) >= 1:
                try:
                    foto_ind_url, _ = await salvar_attachment(anexos_usar[0], "individuo", rg)
                except Exception:
                    pass

            if len(anexos_usar) >= 2:
                try:
                    foto_rg_url, _ = await salvar_attachment(anexos_usar[1], "rg", rg)
                except Exception:
                    pass

            dados = {
                "id": uuid.uuid4().hex,
                "caso": f"SYNC-{msg.id}",
                "nome": nome,
                "rg": rg,
                "crimes_cometidos": crimes,
                "ultimo_avistamento": ultimo,
                "informacoes": informacoes,
                "foto_individuo": foto_ind_url,
                "foto_rg": foto_rg_url,
                "status": "A PROCURAR",
                "data_cadastro": data_hora_br(),
                "autor_id": 0,
                "autor_nome": "Sincronização automática",
                "discord_message_id": msg.id,
                "discord_jump_url": msg.jump_url
            }

            upsert_procurado(dados)
            rgs_existentes.add(rg_limpo)
            adicionados += 1

    gerar_catalogo_html()
    await interaction.followup.send(
        f"✅ Sincronização finalizada.\n"
        f"📌 Novos registros adicionados: **{adicionados}**\n"
        f"📄 Catálogo: {CATALOG_PUBLIC_URL}",
        ephemeral=True
    )


# =====================================================
# INICIAR BOT + SITE
# =====================================================

async def main():
    if not DISCORD_TOKEN or DISCORD_TOKEN == "COLE_O_TOKEN_DO_BOT_AQUI":
        print("ERRO: coloque o DISCORD_TOKEN nas variáveis de ambiente ou no arquivo .env")
        return

    if not PROCURADOS_CHANNEL_ID:
        print("AVISO: PROCURADOS_CHANNEL_ID está vazio. O cadastro não vai conseguir postar no canal oficial.")

    await start_web_server()
    await bot.start(DISCORD_TOKEN)


if __name__ == "__main__":
    asyncio.run(main())
