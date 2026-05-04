@echo off
setlocal
echo Instalando dependencias do Python para o script do Gemini no Word...

where py >nul 2>&1
if %errorlevel%==0 (
	set "PY_CMD=py -3"
) else (
	set "PY_CMD=python"
)

echo Usando interpretador: %PY_CMD%
%PY_CMD% -m pip install --upgrade pip
if errorlevel 1 goto :error

%PY_CMD% -m pip install --upgrade google-generativeai pywin32 python-dotenv
if errorlevel 1 goto :error

%PY_CMD% -c "import win32com.client, win32crypt; print('pywin32 OK')"
if errorlevel 1 goto :error

echo.
echo Instalacao concluida! Pressione qualquer tecla para fechar.
pause > nul
exit /b 0

:error
echo.
echo Falha na instalacao das dependencias.
echo Feche esta janela, abra o Prompt de Comando e rode novamente para ver detalhes.
pause > nul
exit /b 1
