# Registra duas tarefas no Agendador do Windows:
#   - GerarAtasReuniao_Manha: roda às 13h (Coordenação + Enablement)
#   - GerarAtasReuniao_Tarde: roda às 17h (Suporte Educacional + EfOps + Gravações)
# Execute este script como Administrador uma unica vez

$pythonPath = (Get-Command python).Source
$scriptPath = "$PSScriptRoot\monitor.py"

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4)

# Tarefa da manhã — 13h
$actionManha = New-ScheduledTaskAction `
    -Execute $pythonPath `
    -Argument "`"$scriptPath`" manha" `
    -WorkingDirectory $PSScriptRoot

$triggerManha = New-ScheduledTaskTrigger -Daily -At "13:00"

Register-ScheduledTask `
    -TaskName "GerarAtasReuniao_Manha" `
    -Action $actionManha `
    -Trigger $triggerManha `
    -Settings $settings `
    -Description "Gera atas das gravacoes de Coordenacao e Enablement (reunioes de manha)" `
    -RunLevel Highest `
    -Force

# Tarefa da tarde — 17h
$actionTarde = New-ScheduledTaskAction `
    -Execute $pythonPath `
    -Argument "`"$scriptPath`" tarde" `
    -WorkingDirectory $PSScriptRoot

$triggerTarde = New-ScheduledTaskTrigger -Daily -At "17:00"

Register-ScheduledTask `
    -TaskName "GerarAtasReuniao_Tarde" `
    -Action $actionTarde `
    -Trigger $triggerTarde `
    -Settings $settings `
    -Description "Gera atas das gravacoes de Suporte Educacional, EfOps e Gravacoes pessoais" `
    -RunLevel Highest `
    -Force

# Remove a tarefa antiga unificada, se existir
Unregister-ScheduledTask -TaskName "GerarAtasReuniao" -Confirm:$false -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Tarefas agendadas com sucesso!" -ForegroundColor Green
Write-Host "  GerarAtasReuniao_Manha → todo dia as 13h (Coordenacao + Enablement)"
Write-Host "  GerarAtasReuniao_Tarde → todo dia as 17h (Suporte Educacional + EfOps + Gravacoes)"
