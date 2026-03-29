import os
import json
import sys
from pathlib import Path
from datetime import datetime

# Força encoding UTF-8 no terminal Windows
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Carrega variáveis do .env
env_path = Path(__file__).parent / ".env"
if env_path.exists():
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            chave, valor = line.split("=", 1)
            os.environ[chave.strip()] = valor.strip()

import anthropic

# ──────────────────────────────────────────
# CONFIGURAÇÕES
# ──────────────────────────────────────────

# Grupo "manha": verificado às 13h (reuniões ocorrem de manhã)
PASTAS_MANHA = [
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo - Coordenação - Recordings",
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo-0👥 Enablement - Gravações",
]

# Grupo "tarde": verificado às 17h (reuniões ocorrem à tarde ou ao longo do dia)
PASTAS_TARDE = [
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo - Suporte Educacional - Recordings",
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo-1🗨 Eficiência Operacional - Gravações",
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Gravações",
]

EMAIL_DESTINO = "vasco.ginde@alura.com.br"
PROCESSADOS_FILE = Path(__file__).parent / "processados.json"
LOG_FILE = Path(__file__).parent / "log.txt"

# Mapeamento pasta → canal do Teams (sem webhook = não envia ao Teams)
WEBHOOKS = {
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo - Coordenação - Recordings": {
        "canal": "1🗨 _Geral Coordenação",
        "url": "https://fiapcom.webhook.office.com/webhookb2/18f3a6e6-bdce-4a7b-b991-8d3bf4e584fa@11dbbfe2-89b8-4549-be10-cec364e59551/IncomingWebhook/eff9ff1806fc44118aecfd557180d262/0f94b05d-4507-4d35-ab13-15127ef23434/V20ldbpTf-6s4m3UKtG9ouDpcyLCkqQ1RBxUJTMEbE3Z01",
    },
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo - Suporte Educacional - Recordings": {
        "canal": "1🗨 _Geral Suporte Educacional",
        "url": "https://fiapcom.webhook.office.com/webhookb2/f725dc17-2f01-4b82-ad9b-ca9ccfc97c85@11dbbfe2-89b8-4549-be10-cec364e59551/IncomingWebhook/12f44d48ce664a8ba38d6482faf3018e/0f94b05d-4507-4d35-ab13-15127ef23434/V2weh7FbgOCEWKaMSHC61rmjL8Q4_8MTfNtpIqDZ_9Tgc1",
    },
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo-0👥 Enablement - Gravações": {
        "canal": "0👥 Enablement",
        "url": "https://fiapcom.webhook.office.com/webhookb2/7fc03465-f1dc-40a5-942c-90acd0aa9727@11dbbfe2-89b8-4549-be10-cec364e59551/IncomingWebhook/6b40b5364e744e5b96c2b3e0dd85a3eb/0f94b05d-4507-4d35-ab13-15127ef23434/V2Ia_QwbacPJu87NrXPthQvxTTNlPwLXYlgO8HKnBRdBc1",
    },
    r"C:\Users\vasco\OneDrive - Fiap-Faculdade de Informática e Administração Paulista\Conteúdo-1🗨 Eficiência Operacional - Gravações": {
        "canal": "1🗨 Eficiência Operacional",
        "url": "https://fiapcom.webhook.office.com/webhookb2/7fc03465-f1dc-40a5-942c-90acd0aa9727@11dbbfe2-89b8-4549-be10-cec364e59551/IncomingWebhook/36446492cff84c1b8332c527a2556aa6/0f94b05d-4507-4d35-ab13-15127ef23434/V2VINd3H-RHsAS1KIc7smZpTpkriYIj50l9dsHxoGP7Wo1",
    },
}

PROMPT_ATA = """Você receberá uma transcrição automática do Teams ou notas brutas de uma reunião. Seu objetivo é transformar esse conteúdo em uma ata limpa, objetiva e padronizada para o time de Eficiência Operacional da Alura.

## Instruções

1. **Ignore ruídos de transcrição** — remova hesitações ("é...", "tipo assim"), repetições, falas incompletas e conversas paralelas sem relevância.

2. **Extraia as informações obrigatórias:**
   - Data, participantes identificados na transcrição e duração estimada
   - Tópicos discutidos (pauta real, não necessariamente a pauta planejada)
   - Decisões tomadas — cada uma com responsável e prazo quando mencionados
   - Próximos passos / ações — no formato: **ação** → responsável → prazo

3. **Se alguma informação não estiver na transcrição**, deixe o campo com `—` e adicione uma nota discreta ao final: _"Informações não identificadas na transcrição: [lista]"_

4. **Tom:** objetivo e profissional. Não parafraseie demais — preserve o sentido das decisões com precisão.

5. Se o arquivo de áudio contiver silêncio excessivo ou falas inaudíveis, registre como `[trecho inaudível]`.

## Formato de saída

# Ata de Reunião

**Data:** [data]
**Participantes:** [nomes identificados]
**Duração:** [estimada]

---

## Pauta

- [tópico 1]
- [tópico 2]

---

## Decisões

| Decisão | Responsável | Prazo |
|---|---|---|
| [decisão] | [nome] | [data ou "a definir"] |

---

## Próximos passos

- [ ] [ação] → [responsável] → [prazo]
- [ ] [ação] → [responsável] → [prazo]

---

_Ata gerada automaticamente a partir de transcrição. Revise antes de compartilhar._

## Comportamento esperado

- Se a transcrição for muito longa (+1h de reunião), priorize decisões e próximos passos — a pauta pode ser resumida em bullets curtos.
- Se houver ambiguidade sobre quem é o responsável por uma ação, registre como `[a confirmar]` em vez de inferir.
- Nunca invente datas ou prazos que não foram mencionados.
- Não pare de ler a transcrição se o arquivo for muito grande. Tome o tempo que for necessário. É muito importante que toda a transcrição seja lida."""


# ──────────────────────────────────────────
# UTILITÁRIOS
# ──────────────────────────────────────────

def log(mensagem: str):
    linha = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {mensagem}"
    print(linha)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(linha + "\n")


def carregar_processados() -> dict:
    if PROCESSADOS_FILE.exists():
        return json.loads(PROCESSADOS_FILE.read_text(encoding="utf-8"))
    return {}


def salvar_processados(processados: dict):
    PROCESSADOS_FILE.write_text(json.dumps(processados, indent=2, ensure_ascii=False), encoding="utf-8")


# ──────────────────────────────────────────
# TRANSCRIÇÃO (Whisper local)
# ──────────────────────────────────────────

def transcrever(caminho_mp4: str) -> str:
    try:
        from faster_whisper import WhisperModel
    except ImportError:
        log("ERRO: faster-whisper não instalado. Execute instalar.bat")
        sys.exit(1)

    log(f"  Transcrevendo: {Path(caminho_mp4).name} (pode demorar alguns minutos)...")
    model = WhisperModel("small", device="cpu", compute_type="int8")
    segments, info = model.transcribe(caminho_mp4, language="pt", beam_size=5)
    texto = " ".join([seg.text.strip() for seg in segments])
    log(f"  Transcrição concluída. Idioma detectado: {info.language} ({info.language_probability:.0%})")
    return texto


# ──────────────────────────────────────────
# GERAÇÃO DA ATA (Claude API)
# ──────────────────────────────────────────

def gerar_ata(transcricao: str, nome_reuniao: str) -> str:
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        log("ERRO: ANTHROPIC_API_KEY não encontrada no .env")
        sys.exit(1)

    log(f"  Gerando ata com Claude...")
    client = anthropic.Anthropic(api_key=api_key)

    mensagem = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": f"{PROMPT_ATA}\n\n---\n\nNome da reunião: {nome_reuniao}\n\nTranscrição:\n{transcricao}"
            }
        ]
    )
    return mensagem.content[0].text


# ──────────────────────────────────────────
# BUSCA DE PARTICIPANTES NO OUTLOOK
# ──────────────────────────────────────────

def extrair_nomes_da_ata(ata: str) -> list[str]:
    """Extrai a lista de nomes da linha Participantes da ata."""
    for linha in ata.splitlines():
        if linha.startswith("**Participantes:**"):
            trecho = linha.replace("**Participantes:**", "").strip()
            # Remove observações entre parênteses
            import re
            trecho = re.sub(r"\(.*?\)", "", trecho)
            # Separa por vírgula ou 'e'
            nomes = re.split(r",|\be\b", trecho)
            return [n.strip() for n in nomes if n.strip() and n.strip() != "—"]
    return []


def buscar_emails_nos_contatos(nomes: list[str]) -> tuple[list[dict], list[str]]:
    """Busca cada nome no arquivo contatos.json."""
    contatos_path = Path(__file__).parent / "contatos.json"
    if not contatos_path.exists():
        log("  Aviso: contatos.json não encontrado. Crie o arquivo em ata-automatica/contatos.json")
        return [], nomes

    contatos = json.loads(contatos_path.read_text(encoding="utf-8"))

    encontrados = []
    nao_encontrados = []

    for nome in nomes:
        if not nome:
            continue
        # Busca por correspondência parcial (case-insensitive)
        match = next(
            ({"nome": nome, "email": email}
             for chave, email in contatos.items()
             if nome.lower() in chave.lower() or chave.lower() in nome.lower()),
            None
        )
        if match:
            encontrados.append(match)
        else:
            nao_encontrados.append(nome)

    return encontrados, nao_encontrados


def montar_secao_participantes(ata: str) -> str:
    """Gera o bloco markdown com participantes e e-mails encontrados."""
    nomes = extrair_nomes_da_ata(ata)
    if not nomes:
        return ""

    log(f"  Buscando e-mails de {len(nomes)} participante(s) em contatos.json...")
    encontrados, nao_encontrados = buscar_emails_nos_contatos(nomes)

    linhas = ["\n---\n\n## Contatos dos Participantes\n"]

    if encontrados:
        linhas.append("| Nome | E-mail |")
        linhas.append("|---|---|")
        for p in encontrados:
            linhas.append(f"| {p['nome']} | {p['email']} |")

    if nao_encontrados:
        linhas.append(f"\n_Não encontrados no catálogo: {', '.join(nao_encontrados)}_")

    return "\n".join(linhas)


# ──────────────────────────────────────────
# ENVIO DE E-MAIL (via Outlook desktop)
# ──────────────────────────────────────────

def markdown_para_html(texto: str) -> str:
    import markdown
    corpo_html = markdown.markdown(texto, extensions=["tables", "nl2br"])
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{ font-family: Calibri, Arial, sans-serif; font-size: 14px; color: #1a1a1a; max-width: 800px; margin: 0 auto; padding: 24px; }}
  h1 {{ font-size: 22px; border-bottom: 2px solid #0078d4; padding-bottom: 8px; color: #0078d4; }}
  h2 {{ font-size: 16px; margin-top: 28px; color: #333; border-left: 4px solid #0078d4; padding-left: 10px; }}
  table {{ border-collapse: collapse; width: 100%; margin: 12px 0; }}
  th {{ background-color: #0078d4; color: white; padding: 8px 12px; text-align: left; }}
  td {{ border: 1px solid #ddd; padding: 8px 12px; }}
  tr:nth-child(even) {{ background-color: #f5f5f5; }}
  ul {{ padding-left: 20px; }}
  li {{ margin-bottom: 6px; }}
  hr {{ border: none; border-top: 1px solid #ddd; margin: 20px 0; }}
  em {{ color: #666; font-size: 12px; }}
  strong {{ color: #1a1a1a; }}
</style>
</head>
<body>
{corpo_html}
</body>
</html>"""


def enviar_email(assunto: str, corpo: str):
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_DESTINO
        mail.Subject = assunto
        mail.HTMLBody = markdown_para_html(corpo)
        mail.Send()
        log(f"  E-mail enviado para {EMAIL_DESTINO}")
    except Exception as e:
        log(f"  ERRO ao enviar e-mail via Outlook: {e}")
        # Fallback: salva a ata em arquivo
        pasta_saida = Path(__file__).parent / "atas_geradas"
        pasta_saida.mkdir(exist_ok=True)
        nome_arquivo = assunto.replace(":", "").replace("/", "-")[:80] + ".txt"
        (pasta_saida / nome_arquivo).write_text(corpo, encoding="utf-8")
        log(f"  Ata salva em: {pasta_saida / nome_arquivo}")


# ──────────────────────────────────────────
# ──────────────────────────────────────────
# ENVIO PARA O TEAMS
# ──────────────────────────────────────────

def enviar_teams(ata: str, nome_reuniao: str, pasta: str):
    import urllib.request
    import re

    webhook = WEBHOOKS.get(pasta)
    if not webhook:
        return  # pasta Gravações não tem canal vinculado

    canal = webhook["canal"]
    url = webhook["url"]

    # Extrai seções da ata para montar card resumido
    def extrair_secao(texto, titulo):
        padrao = rf"## {titulo}\n(.*?)(?=\n---|\Z)"
        match = re.search(padrao, texto, re.DOTALL)
        return match.group(1).strip() if match else ""

    decisoes = extrair_secao(ata, "Decisões")
    proximos = extrair_secao(ata, "Próximos passos")
    participantes_linha = next(
        (l for l in ata.splitlines() if l.startswith("**Participantes:**")), ""
    ).replace("**Participantes:**", "").strip()
    data_linha = next(
        (l for l in ata.splitlines() if l.startswith("**Data:**")), ""
    ).replace("**Data:**", "").strip()

    def tabela_para_bullets(tabela: str) -> str:
        """Converte tabela markdown em lista de bullets para o Teams."""
        linhas = [l.strip() for l in tabela.splitlines() if l.strip()]
        bullets = []
        for linha in linhas:
            # Ignora cabeçalho e separador
            if linha.startswith("|---") or linha.startswith("| ---") or linha.startswith("| Decisão"):
                continue
            if linha.startswith("|"):
                colunas = [c.strip() for c in linha.strip("|").split("|")]
                colunas = [c for c in colunas if c]
                if colunas:
                    bullets.append("- " + " → ".join(colunas))
        return "\n".join(bullets)

    corpo = f"📋 **Ata de Reunião — {nome_reuniao}**\n\n"
    if data_linha:
        corpo += f"📅 **Data:** {data_linha}\n\n"
    if participantes_linha:
        corpo += f"👥 **Participantes:** {participantes_linha}\n\n"
    if decisoes:
        corpo += f"---\n\n**✅ Decisões**\n\n{tabela_para_bullets(decisoes)}\n\n"
    if proximos:
        corpo += f"---\n\n**⏭ Próximos passos**\n\n{proximos}\n\n"
    corpo += "---\n\n_Ata gerada automaticamente. Revise antes de compartilhar._"

    payload = json.dumps({"text": corpo}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"})

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            if resp.status == 200:
                log(f"  Teams: mensagem enviada para o canal '{canal}'")
            else:
                log(f"  Teams: resposta inesperada ({resp.status})")
    except Exception as e:
        log(f"  ERRO ao enviar para o Teams (canal '{canal}'): {e}")


# LOOP PRINCIPAL
# ──────────────────────────────────────────

def main():
    grupo = sys.argv[1] if len(sys.argv) > 1 else "todos"

    if grupo == "manha":
        pastas = PASTAS_MANHA
    elif grupo == "tarde":
        pastas = PASTAS_TARDE
    else:
        pastas = PASTAS_MANHA + PASTAS_TARDE

    log("=" * 60)
    log(f"Iniciando verificação — grupo: {grupo} ({len(pastas)} pasta(s))")

    processados = carregar_processados()
    novos = 0

    for pasta in pastas:
        pasta_path = Path(pasta)
        if not pasta_path.exists():
            log(f"Pasta não encontrada (ignorando): {pasta_path.name}")
            continue

        arquivos_mp4 = list(pasta_path.glob("*.mp4"))
        log(f"Pasta '{pasta_path.name}': {len(arquivos_mp4)} arquivo(s) .mp4")

        for arquivo in arquivos_mp4:
            chave = str(arquivo)

            if chave in processados:
                continue

            log(f"Novo arquivo detectado: {arquivo.name}")
            novos += 1

            try:
                transcricao = transcrever(str(arquivo))

                if not transcricao.strip():
                    log("  AVISO: Transcrição vazia — pulando este arquivo.")
                    processados[chave] = {"status": "vazio", "data": datetime.now().isoformat()}
                    salvar_processados(processados)
                    continue

                ata = gerar_ata(transcricao, arquivo.stem)
                secao_contatos = montar_secao_participantes(ata)
                assunto = f"Ata de Reunião — {arquivo.stem}"
                enviar_email(assunto, ata + secao_contatos)
                enviar_teams(ata, arquivo.stem, pasta)

                processados[chave] = {"status": "ok", "data": datetime.now().isoformat()}
                salvar_processados(processados)

            except Exception as e:
                log(f"  ERRO ao processar {arquivo.name}: {e}")
                processados[chave] = {"status": "erro", "erro": str(e), "data": datetime.now().isoformat()}
                salvar_processados(processados)

    if novos == 0:
        log("Nenhum arquivo novo encontrado.")
    else:
        log(f"Processamento concluído. {novos} arquivo(s) novo(s) tratado(s).")

    log("=" * 60)


if __name__ == "__main__":
    main()
