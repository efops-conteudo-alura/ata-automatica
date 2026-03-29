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
Geração da ata com Claude API
        ↓
E-mail HTML + transcrição em anexo → você
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

**6. Agende a execução automática**

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

## Configuração de GPU / CPU

O programa usa o modelo **Whisper large-v3** para transcrição. Por padrão, roda na GPU via CUDA, o que é muito mais rápido. Se você não tiver GPU NVIDIA, é necessário ajustar duas linhas no `monitor.py`.

### Com GPU NVIDIA (padrão)

Localiza a função `transcrever` no `monitor.py` e confirme que está assim:

```python
_whisper_model = WhisperModel("large-v3", device="cuda", compute_type="float16")
```

Requer placa NVIDIA com suporte a CUDA. Uma RTX com 8GB+ de VRAM é o suficiente.

### Sem GPU (CPU)

Troque para:

```python
_whisper_model = WhisperModel("large-v3", device="cpu", compute_type="int8")
```

> **Atenção:** na CPU, a transcrição é significativamente mais lenta. Uma reunião de 30 minutos pode levar 10–20 minutos dependendo do processador. Se o tempo for crítico, considere usar um modelo menor como `medium` ou `small` no lugar de `large-v3` — com perda de precisão, mas muito mais rápido.

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
├── monitor.py           # Script principal
├── agendar_tarefa.ps1   # Cria as tarefas agendadas no Windows
├── instalar.bat         # Instala as dependências Python
├── .env                 # Chave da API (não versionar!)
├── .env.example         # Modelo do .env
├── processados.json     # Registro de arquivos já processados (gerado automaticamente)
├── log.txt              # Log de execuções (gerado automaticamente)
└── atas_geradas/        # Atas salvas localmente em caso de falha no e-mail
```

---

## O que é enviado

### E-mail (só para você)
- Ata completa formatada em HTML com cabeçalho em azul
- Tabela de decisões com prazos (quando mencionados)
- Próximos passos
- Transcrição bruta do Whisper em anexo (`.txt`)

### Teams (canal do time)
- Resumo com data
- Decisões em formato de lista
- Próximos passos

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
| `anthropic` | Geração da ata via Claude API |
| `faster-whisper` | Transcrição local do áudio |
| `pywin32` | Envio de e-mail via Outlook e leitura de atributos de arquivo |
| `markdown` | Conversão do markdown para HTML no e-mail |
