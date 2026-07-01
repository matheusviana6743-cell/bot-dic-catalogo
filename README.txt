BOT DIC - PROCURADOS + CATALOGO HTML

1) Renomeie .env.example para .env
2) Preencha:
   DISCORD_TOKEN=
   GUILD_ID=
   PROCURADOS_CHANNEL_ID=
   CATALOG_PUBLIC_URL=http://127.0.0.1:8000/
   PORT=8000

3) Instale:
   python -m pip install -r requirements.txt

4) Ligue:
   python bot.py

5) Abra o catalogo:
   http://127.0.0.1:8000/

Comandos:
/painelprocurados - envia o painel com botoes
/catalogo - mostra o link do catalogo
/retirarprocurado - retira pelo RG
/sincronizarcatalogo - importa procurados antigos do canal oficial
/regerarcatalogo - recria o HTML

Fluxo:
+ Novo Procurado -> modal de texto -> canal provisorio para fotos -> Finalizar Cadastro -> posta no canal oficial e atualiza o catalogo.

IMPORTANTE:
Use apenas para GTA RP/personagens ficticios ou ambiente autorizado.
Nao mostre seu token em prints.
