@echo off
echo Baixando o Python 3...
curl -o python-installer.exe https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe
echo Instalando o Python 3...
python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
echo.
echo Python 3 foi instalado com sucesso!
echo Agora execute o arquivo 02_download_bibliotecas.bat 
echo.
pause