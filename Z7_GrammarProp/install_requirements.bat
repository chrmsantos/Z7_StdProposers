@echo off
echo Instalando dependencias do Python para o script do Gemini no Word...
python -m pip install --upgrade pip
pip install google-generativeai pywin32 python-dotenv
echo.
echo Instalacao concluida! Pressione qualquer tecla para fechar.
pause > nul
