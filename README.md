# Gerador Automático de Atas de Reunião

Sistema que monitora pastas de gravações do Microsoft Teams, transcreve o áudio, gera uma ata formatada usando IA e envia por e-mail e para o canal do Teams correspondente.

---

## Como funciona

```
Gravação .mp4 nova na pasta
        ↓
Aguarda download do OneDrive (se necessário)
        ↓
Transcrição com Whisper (local, gratuito)
        ↓
Resolve participantes:
  ├─ via participantes_reunioes.json (reuniões recorrentes)
  └─ via diálogo tkinter (reuniões não reconhecidas, timeout 10 min)
        ↓
Geração da ata com Claude API (claude-sonnet-4-6)
        ↓
E-mail HTML + transcrição em anexo → você (+ CC para participantes)
Teams webhook → canal do time
```

---

## Pré-requisitos

- Windows 10/11
- Python 3.10+
- Microsoft Outlook instalado e logado
- Chave de API da Anthropic ([console.anthropic.com](https://console.anthropic.com))
- GPU NVIDIA com CUDA (recomendado) **ou** CPU (ver seção abaixo)

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

**6. Configure os participantes e contatos** *(opcional, mas recomendado)*

Veja a [seção de Participantes e E-mails](#participantes-e-e-mails) abaixo.

**7. Agende a execução automática**

Abra o PowerShell como Administrador e rode:
```powershell
.\agendar_tarefa.ps1
```

Isso cria duas tarefas no Agendador do Windows:
- **GerarAtasReuniao_Manha** — roda às **13h** (pastas com reuniões de manhã)
- **GerarAtasReuniao_Tarde** — roda às **17h** (pastas com reuniões à tarde)

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

## Participantes e E-mails

O sistema resolve automaticamente os participantes de cada reunião em duas etapas:

### 1. `participantes_reunioes.json` — reuniões recorrentes

Mapeia padrões de nome de arquivo para listas fixas de participantes. A busca é parcial e ignora maiúsculas/minúsculas.

```json
{
  "Weekly Enablement": {
    "participantes": "Ana Silva, João Souza, Maria Lima",
    "emails_extra": []
  },
  "Weekly Financeiro": {
    "participantes": "Carlos Melo, Paula Ramos",
    "emails_extra": ["externo@empresa.com"]
  }
}
```

- **`participantes`** — nomes separados por vírgula; aparecem na ata e são usados para buscar e-mails em `contatos.json`
- **`emails_extra`** — e-mails adicionais de pessoas que não estão em `contatos.json` (ex.: convidados externos)
- Deixe `participantes` vazio (`""`) para que o diálogo apareça mesmo em reuniões reconhecidas

### 2. Diálogo interativo — reuniões não reconhecidas

Quando o nome do arquivo não bate com nenhuma entrada do `participantes_reunioes.json`, uma janela do Tkinter abre pedindo os nomes:

- Digite os nomes separados por vírgula e clique em **Confirmar**
- Clique em **Gerar sem nomes** para omitir o campo
- A janela fecha sozinha após **10 minutos** sem interação (a ata é gerada sem nomes)

### 3. `contatos.json` — catálogo de e-mails

Dicionário `"Nome Completo": "email@empresa.com"` usado para montar o CC do e-mail:

```json
{
  "Ana Silva": "ana.silva@empresa.com",
  "João Souza": "joao.souza@empresa.com"
}
```

A busca é parcial e case-insensitive — `"João"` encontra `"João Souza"`. Os participantes encontrados nesse arquivo são adicionados automaticamente no campo CC do e-mail gerado.

> **Atenção:** esse arquivo pode conter dados pessoais (e-mails). Avalie se deve incluí-lo no `.gitignore` conforme a política da sua organização.

---

## Configuração de GPU / CPU

O programa usa o modelo **Whisper large-v3** para transcrição. Por padrão, roda na GPU via CUDA, o que é muito mais rápido. Se você não tiver GPU NVIDIA, é necessário ajustar duas linhas no `monitor.py`.

### Com GPU NVIDIA (padrão)

Localize a função `transcrever` no `monitor.py` e confirme que está assim:

```python
_whisper_model = WhisperModel("large-v3", device="cuda", compute_type="float16")
```

Requer placa NVIDIA com suporte a CUDA. Uma RTX com 8GB+ de VRAM é o suficiente.

### Sem GPU (CPU)

Troque para:

```python
_whisper_model = WhisperModel("large-v3", device="cpu", compute_type="int8")
```

> **Atenção:** na CPU, a transcrição é significativamente mais lenta. Uma reunião de 30 minutos pode levar 10–20 minutos dependendo do processador. Se o tempo for crítico, considere usar um modelo menor como `medium` ou `small` — com perda de precisão, mas muito mais rápido.

### Modelos disponíveis

| Modelo | Precisão | Velocidade | VRAM mínima |
|---|---|---|---|
| `large-v3` | Melhor | Mais lento | ~5 GB |
| `medium` | Boa | Moderado | ~2 GB |
| `small` | Razoável | Rápido | ~1 GB |

Para trocar o modelo, altere `"large-v3"` na linha acima pelo nome desejado.

---

## Integração com OneDrive / SharePoint

As pastas monitoradas podem ser sincronizadas via OneDrive (pastas do SharePoint que aparecem no seu PC). Nesse caso, os arquivos exibem dois estados:

- ☁ **Nuvem** — o arquivo ainda não foi baixado para o disco local
- ✅ **Disponível** — o arquivo está no HD e pronto para leitura

O programa detecta automaticamente quando um arquivo está apenas na nuvem e aguarda o download pelo OneDrive antes de iniciar a transcrição (timeout de 10 minutos). Isso evita transcrições vazias ou corrompidas.

> Se o download não concluir em 10 minutos, o arquivo é pulado e registrado como erro no `log.txt`. Isso pode ocorrer com arquivos muito grandes em conexões lentas.

---

## Estrutura dos arquivos

```
ata-automatica/
├── monitor.py                   # Script principal
├── agendar_tarefa.ps1           # Cria as tarefas agendadas no Windows
├── instalar.bat                 # Instala as dependências Python
├── .env                         # Chave da API (não versionar!)
├── .env.example                 # Modelo do .env
├── contatos.json                # Catálogo nome → e-mail para CC automático
├── participantes_reunioes.json  # Mapeamento reunião recorrente → participantes
├── processados.json             # Registro de arquivos já processados (gerado automaticamente)
├── log.txt                      # Log de execuções (gerado automaticamente)
└── atas_geradas/                # Atas salvas localmente em caso de falha no e-mail
```

---

## O que é enviado

### E-mail (para você + CC dos participantes)
- Ata completa formatada em HTML com cabeçalho em azul
- Participantes encontrados em `contatos.json` são adicionados em CC automaticamente
- Próximos passos e encaminhamentos com prazos (quando mencionados)
- Transcrição bruta do Whisper em anexo (`.txt`)

### Teams (canal do time)
- Resumo com data e participantes
- Encaminhamentos em formato de lista
- Enviado somente para pastas com webhook configurado no dicionário `WEBHOOKS`

---

## Observações

- Arquivos já processados ficam registrados em `processados.json` — a ata nunca é gerada duas vezes para o mesmo arquivo
- Gravações muito curtas ou silenciosas são puladas automaticamente com `status: "vazio"` no registro
- A ata é gerada de forma **impessoal** — nomes de pessoas não aparecem em decisões, ações ou próximos passos, apenas no campo Participantes
- A pasta `Gravações` (1:1s pessoais) só envia e-mail, sem canal do Teams vinculado
- Em caso de falha no envio por e-mail, a ata é salva em `atas_geradas/`
- Em caso de crash silencioso durante a transcrição (comum em drivers Windows WDDM), verifique o `log.txt` para identificar em qual etapa parou

---

## Dependências

| Biblioteca | Uso |
|---|---|
| `anthropic` | Geração da ata via Claude API (modelo `claude-sonnet-4-6`) |
| `faster-whisper` | Transcrição local do áudio |
| `pywin32` | Envio de e-mail via Outlook e leitura de atributos de arquivo |
| `markdown` | Conversão do markdown para HTML no e-mail |
