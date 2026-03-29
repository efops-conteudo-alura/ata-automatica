@echo off
echo ============================================
echo  Instalando dependencias do sistema de atas
echo ============================================
echo.

pip install anthropic faster-whisper pywin32 python-dotenv

echo.
echo ============================================
echo  Instalacao concluida!
echo  Execute: python monitor.py  para testar
echo ============================================
pause
