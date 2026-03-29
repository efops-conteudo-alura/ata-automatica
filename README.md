# Gerador Automático de Atas de Reunião

Sistema que monitora pastas de gravações do Microsoft Teams, transcreve o áudio, gera uma ata formatada usando IA e envia por e-mail e para o canal do Teams correspondente.

---

## Como funciona

```
Gravação .mp4 nova na pasta
        ↓
Transcrição com Whisper (local, gratuito)
        ↓
Geração da ata com Claude API
        ↓
E-mail HTML → você
Teams webhook → canal do time
```

---

## Pré-requisitos

- Windows 10/11
- Python 3.10+
- Microsoft Outlook instalado e logado
- Chave de API da Anthropic ([console.anthropic.com](https://console.anthropic.com))

---

## Instalação

**1. Clone o repositório**
```bash
git clone https://github.com/seu-usuario/ata-automatica.git
cd ata-automatica
```

**2. Instale as dependências**

Dê dois cliques em `instalar.bat` ou rode no terminal:
```bash
pip install anthropic faster-whisper pywin32 python-dotenv markdown
```

**3. Configure o arquivo `.env`**

Copie o arquivo de exemplo e preencha com sua chave:
```bash
copy .env.example .env
```

Edite o `.env`:
```
ANTHROPIC_API_KEY=sk-ant-...sua chave aqui...
```

**4. Configure as pastas monitoradas**

Edite as variáveis `PASTAS_MANHA` e `PASTAS_TARDE` no `monitor.py` com os caminhos das suas pastas de gravações do Teams.

**5. Configure os webhooks do Teams**

Para cada canal que deve receber as atas:
1. Acesse o canal no Teams → `...` → **Fluxos de trabalho** → **Incoming Webhook**
2. Copie a URL gerada
3. Cole no dicionário `WEBHOOKS` no `monitor.py`

**6. Configure os contatos**

Edite o `contatos.json` com os nomes e e-mails do seu time:
```json
{
  "Nome Sobrenome": "email@empresa.com.br"
}
```

**7. Agende a execução automática**

Abra o PowerShell como Administrador e rode:
```powershell
.\agendar_tarefa.ps1
```

Isso cria duas tarefas no Agendador do Windows:
- **13h** — pastas com reuniões de manhã
- **17h** — pastas com reuniões à tarde

---

## Executar manualmente

```bash
# Verifica todas as pastas
python monitor.py

# Verifica só as pastas do grupo manhã
python monitor.py manha

# Verifica só as pastas do grupo tarde
python monitor.py tarde
```

---

## Estrutura dos arquivos

```
ata-automatica/
├── monitor.py           # Script principal
├── contatos.json        # Mapeamento nome → e-mail do time
├── agendar_tarefa.ps1   # Cria as tarefas agendadas no Windows
├── instalar.bat         # Instala as dependências Python
├── .env                 # Chave da API (não versionar!)
├── .env.example         # Modelo do .env
├── processados.json     # Registro de arquivos já processados (gerado automaticamente)
└── log.txt              # Log de execuções (gerado automaticamente)
```

---

## O que é enviado

### E-mail (só para você)
- Ata completa formatada em HTML
- Tabela de decisões com responsáveis e prazos
- Próximos passos
- Seção com e-mails dos participantes identificados

### Teams (canal do time)
- Resumo com data e participantes
- Decisões em formato de lista
- Próximos passos

---

## Observações

- Arquivos já processados ficam registrados em `processados.json` — a ata nunca é gerada duas vezes para o mesmo arquivo
- Gravações muito curtas ou silenciosas são puladas automaticamente
- Nomes na transcrição podem ter erros — o Whisper ocasionalmente erra nomes próprios. Revise antes de compartilhar a ata
- A pasta `Gravações` (1:1s pessoais) só envia e-mail, sem canal do Teams vinculado
- Em caso de falha no envio por e-mail, a ata é salva em `atas_geradas/`

---

## Dependências

| Biblioteca | Uso |
|---|---|
| `anthropic` | Geração da ata via Claude API |
| `faster-whisper` | Transcrição local do áudio |
| `pywin32` | Envio de e-mail via Outlook |
| `markdown` | Conversão do markdown para HTML no e-mail |
