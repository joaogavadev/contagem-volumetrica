# =============================================================================
#  Contagem Volumetrica - preparacao do computador
#
#  Instala o Python (se ainda nao tiver) e confere se esta tudo pronto
#  para rodar o main.py. Nao precisa de admin: instala so para o usuario.
#
#  Como rodar, no PowerShell:
#     powershell -ExecutionPolicy Bypass -File instalar.ps1
# =============================================================================

$ErrorActionPreference = 'Stop'
$VersaoPython = '3.13.14'   # usada so no plano B, se nao houver winget

function Escrever($txt, $cor = 'Gray') { Write-Host $txt -ForegroundColor $cor }

function Atualizar-Path {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

function Python-Funciona {
    # Cuidado: num Windows limpo, "python" costuma ser um atalho falso que so
    # abre a Microsoft Store. Por isso testamos executando de verdade.
    try {
        $v = & python -c "import sys; print(sys.version_info.major)" 2>$null
        return ($LASTEXITCODE -eq 0 -and $v -eq '3')
    } catch { return $false }
}

Escrever ""
Escrever "== Contagem Volumetrica - preparacao ==" Cyan
Escrever ""

# --- 1. Python ---------------------------------------------------------------
Atualizar-Path

if (Python-Funciona) {
    $v = & python -c "import sys; print(sys.version.split()[0])"
    Escrever "[ok] Python $v ja instalado." Green
}
else {
    Escrever "[..] Python nao encontrado. Instalando..." Yellow

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Escrever "     usando o winget"
        winget install -e --id Python.Python.3.13 --scope user `
               --accept-package-agreements --accept-source-agreements
        Atualizar-Path
    }

    if (-not (Python-Funciona)) {
        Escrever "     baixando o instalador do python.org"
        $url = "https://www.python.org/ftp/python/$VersaoPython/python-$VersaoPython-amd64.exe"
        $exe = Join-Path $env:TEMP "python-$VersaoPython-amd64.exe"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $url -OutFile $exe

        Escrever "     instalando (pode demorar 1-2 minutos)"
        Start-Process $exe -Wait -ArgumentList @(
            '/quiet', 'InstallAllUsers=0', 'PrependPath=1',
            'Include_tcltk=1', 'Include_pip=1', 'Include_test=0'
        )
        Remove-Item $exe -ErrorAction SilentlyContinue
        Atualizar-Path
    }

    if (-not (Python-Funciona)) {
        Escrever "[X] Nao consegui instalar o Python automaticamente." Red
        Escrever "    Instale a mao em https://www.python.org/downloads/" Red
        Escrever "    e marque 'Add python.exe to PATH' na primeira tela." Red
        exit 1
    }

    $v = & python -c "import sys; print(sys.version.split()[0])"
    Escrever "[ok] Python $v instalado." Green
}

# --- 2. Bibliotecas ----------------------------------------------------------
# O main.py usa so a biblioteca padrao. Nao ha nada para instalar com pip -
# aqui a gente so confere se os modulos necessarios respondem.
Escrever ""
Escrever "[..] Conferindo os modulos necessarios..."

$teste = @'
import sys
faltando = []
for m in ("zipfile", "re", "shutil", "collections",
          "xml.etree.ElementTree", "tkinter"):
    try:
        __import__(m)
    except Exception:
        faltando.append(m)
if faltando:
    print("FALTA:" + ",".join(faltando))
    sys.exit(1)
print("OK")
'@

$r = $teste | & python -
if ($LASTEXITCODE -ne 0) {
    Escrever "[X] Faltam modulos: $r" Red
    Escrever "    Reinstale o Python marcando a opcao 'tcl/tk and IDLE'." Red
    exit 1
}
Escrever "[ok] Todos os modulos presentes (nenhuma instalacao com pip e necessaria)." Green

# --- 3. Pronto ---------------------------------------------------------------
Escrever ""
Escrever "== Tudo pronto ==" Cyan
Escrever ""
Escrever "Para rodar, na pasta do script:"
Escrever "   python main.py" White
Escrever ""
Escrever "Se o comando 'python' nao for reconhecido, feche e reabra o PowerShell." DarkGray
Escrever ""
