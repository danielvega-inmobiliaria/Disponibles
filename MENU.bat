@echo off
:inicio
cls
echo.
echo  =============================================
echo    DISPONIBLES - MENU PRINCIPAL
echo  =============================================
echo.
echo   1. Actualizar precios desde listas Excel
echo   2. Generar HTML (app movil)
echo   3. Abrir app en el browser
echo   4. Servir en red local (mismo WiFi)
echo   5. Sincronizar con listas nuevas
echo      (detecta altas, bajas y cambios de precio)
echo   6. Gestion de propiedades
echo      (alta / baja / reactivar)
echo   7. Ver historial de propiedades vendidas
echo   8. Publicar a GitHub
echo      (regenera HTML + sube cambios a la app)
echo.
echo   0. Instalar dependencias (primera vez)
echo   9. Salir
echo.
set /p op="  Elegir opcion [0-9]: "

if "%op%"=="0" goto instalar
if "%op%"=="1" goto actualizar
if "%op%"=="2" goto generar
if "%op%"=="3" goto abrir
if "%op%"=="4" goto servir
if "%op%"=="5" goto sincronizar
if "%op%"=="6" goto gestion
if "%op%"=="7" goto historial
if "%op%"=="8" goto publicar
if "%op%"=="9" goto fin
goto inicio

:instalar
cls
echo.
echo  Instalando dependencias Python...
echo.
pip install openpyxl pillow
echo.
echo  Listo. Ya podes usar las opciones 1 y 2.
echo.
pause
goto inicio

:actualizar
cls
echo.
echo  Actualizando precios...
echo.
python "%~dp0actualizar_disponibles.py"
echo.
pause
goto inicio

:generar
cls
echo.
echo  Generando HTML...
echo.
python "%~dp0generar_html.py"
echo.
pause
goto inicio

:abrir
start "" "%~dp0disponibles_mobile.html"
goto inicio

:servir
cls
echo.
echo  Buscando IP local...
echo.
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do set IP=%%a
set IP=%IP: =%
echo  ============================================
echo    Abre en el celular (mismo WiFi):
echo.
echo    http://%IP%:8080/disponibles_mobile.html
echo.
echo  ============================================
echo.
echo  Presiona Ctrl+C para detener.
echo.
cd /d "%~dp0"
python -m http.server 8080
goto inicio

:sincronizar
cls
echo.
echo  =============================================
echo    SINCRONIZAR CON LISTAS NUEVAS
echo  =============================================
echo.
echo  Compara las listas de precio M2 y Obring
echo  con las propiedades actuales en DISPONIBLES.
echo.
echo  Antes de ejecutar, asegurate de que las
echo  listas nuevas esten en sus carpetas:
echo    M2 DESARROLLOS\LISTAS DE PRECIO\
echo    OBRING\
echo.
pause
python "%~dp0sincronizar_disponibles.py"
echo.
pause
goto inicio

:gestion
cls
echo.
echo  =============================================
echo    GESTION DE PROPIEDADES
echo  =============================================
echo.
python "%~dp0baja_manual.py"
echo.
pause
goto inicio

:historial
cls
echo.
echo  =============================================
echo    HISTORIAL DE PROPIEDADES VENDIDAS
echo  =============================================
echo.
python "%~dp0baja_manual.py" --historial
echo.
pause
goto inicio

:publicar
cls
echo.
echo  =============================================
echo    PUBLICAR A GITHUB
echo  =============================================
echo.
echo  Paso 1: Regenerando HTML...
echo.
cd /d "%~dp0"
python "%~dp0generar_html.py"
echo.
echo  ============================================
echo  Paso 2: Subiendo a GitHub...
echo  ============================================
echo.
git add -A
git status --short
echo.
echo  Ingresa un mensaje para el commit (o Enter para mensaje automatico):
set /p msg="  > "
if "%msg%"=="" set msg=Actualizar disponibles %date% %time:~0,5%

git commit -m "%msg%"
if errorlevel 1 (
    echo.
    echo  AVISO: Git no encontro cambios nuevos para subir.
    echo  Si acabas de hacer un cambio, espera un momento y vuelve a intentar.
    echo.
    pause
    goto inicio
)

echo.
echo  Subiendo...
echo.
git push
if errorlevel 1 (
    echo.
    echo  ERROR al subir. Verifica tu conexion o credenciales de GitHub.
    echo.
    pause
    goto inicio
)

echo.
echo  ============================================
echo    OK - App publicada en GitHub Pages
echo.
echo    https://danielvega-inmobiliaria.github.io/Disponibles/
echo.
echo    Espera 1-2 minutos y recarga en el celu.
echo  ============================================
echo.
pause
goto inicio

:fin
exit
