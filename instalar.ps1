# =============================================================================
#  Contagem Volumetrica - instalacao completa
#
#  Baixa o programa, instala o Python (se precisar) e cria o atalho na
#  Area de Trabalho. Nao precisa de administrador.
#
#  No PowerShell, uma linha so:
#     irm https://raw.githubusercontent.com/joaogavadev/contagem-volumetrica/main/instalar.ps1 | iex
#
#  Rodar de novo depois atualiza o programa para a ultima versao.
# =============================================================================

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$Repo         = 'joaogavadev/contagem-volumetrica'
$Ramo         = 'main'
$VersaoPython = '3.13.14'      # plano B, se a maquina nao tiver winget
$Destino      = Join-Path $env:LOCALAPPDATA 'ContagemVolumetrica'
$NomeAtalho   = 'Contagem Volumetrica'

function Escrever($txt, $cor = 'Gray') { Write-Host $txt -ForegroundColor $cor }

function Atualizar-Path {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' +
                [Environment]::GetEnvironmentVariable('Path', 'User')
}

function Python-Funciona {
    # Num Windows limpo, "python" costuma ser um atalho falso que so abre a
    # Microsoft Store. Por isso testamos executando de verdade.
    try {
        $v = & python -c "import sys; print(sys.version_info.major)" 2>$null
        return ($LASTEXITCODE -eq 0 -and $v -eq '3')
    } catch { return $false }
}

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Escrever ""
Escrever "===  Contagem Volumetrica  ===" Cyan
Escrever ""

# --- 1. Python ---------------------------------------------------------------
Atualizar-Path

if (Python-Funciona) {
    Escrever "[ok] Python $(& python -c 'import sys; print(sys.version.split()[0])') ja instalado." Green
} else {
    Escrever "[..] Instalando o Python..." Yellow

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        winget install -e --id Python.Python.3.13 --scope user `
               --accept-package-agreements --accept-source-agreements
        Atualizar-Path
    }

    if (-not (Python-Funciona)) {
        Escrever "     baixando do python.org (pode demorar alguns minutos)"
        $exe = Join-Path $env:TEMP "python-$VersaoPython-amd64.exe"
        Invoke-WebRequest "https://www.python.org/ftp/python/$VersaoPython/python-$VersaoPython-amd64.exe" -OutFile $exe
        Start-Process $exe -Wait -ArgumentList @(
            '/quiet', 'InstallAllUsers=0', 'PrependPath=1',
            'Include_tcltk=1', 'Include_pip=1', 'Include_test=0')
        Remove-Item $exe -ErrorAction SilentlyContinue
        Atualizar-Path
    }

    if (-not (Python-Funciona)) {
        Escrever "[X] Nao consegui instalar o Python." Red
        Escrever "    Instale a mao em https://www.python.org/downloads/" Red
        Escrever "    marcando 'Add python.exe to PATH' e 'tcl/tk and IDLE'." Red
        return
    }
    Escrever "[ok] Python instalado." Green
}

# --- 2. Programa -------------------------------------------------------------
Escrever ""
Escrever "[..] Baixando o programa..."

$zip  = Join-Path $env:TEMP 'contagem.zip'
$tmp  = Join-Path $env:TEMP 'contagem_tmp'
Invoke-WebRequest "https://github.com/$Repo/archive/refs/heads/$Ramo.zip" -OutFile $zip

if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
Expand-Archive $zip -DestinationPath $tmp -Force
Remove-Item $zip -Force

$origem = Get-ChildItem $tmp -Directory | Select-Object -First 1
if (-not (Test-Path $Destino)) { New-Item -ItemType Directory -Path $Destino -Force | Out-Null }
Copy-Item (Join-Path $origem.FullName '*') $Destino -Recurse -Force
Remove-Item $tmp -Recurse -Force

Escrever "[ok] Programa em $Destino" Green

# --- 3. Conferencia ----------------------------------------------------------
$teste = @'
import sys
falta = [m for m in ("zipfile","re","shutil","collections",
                     "xml.etree.ElementTree","tkinter")
         if not __import__("importlib").util.find_spec(m)]
print("FALTA:" + ",".join(falta) if falta else "OK")
sys.exit(1 if falta else 0)
'@
$r = $teste | & python -
if ($LASTEXITCODE -ne 0) {
    Escrever "[!] Faltam modulos: $r" Yellow
    Escrever "    Reinstale o Python marcando 'tcl/tk and IDLE'." Yellow
}

# --- 4. Atalho na Area de Trabalho -------------------------------------------
# pythonw abre a janela sem o console preto atras.
$pyw = (Get-Command pythonw -ErrorAction SilentlyContinue).Source
if (-not $pyw) { $pyw = (Get-Command python).Source }

$atalho = Join-Path ([Environment]::GetFolderPath('Desktop')) "$NomeAtalho.lnk"
$s = (New-Object -ComObject WScript.Shell).CreateShortcut($atalho)
$s.TargetPath       = $pyw
$s.Arguments        = '"' + (Join-Path $Destino 'main.py') + '"'
$s.WorkingDirectory = $Destino
$s.Description      = 'Conta o Word e preenche a planilha da contagem'
$s.Save()

Escrever "[ok] Atalho criado na Area de Trabalho." Green

# --- 5. Pronto ---------------------------------------------------------------
Escrever ""
Escrever "===  Tudo pronto  ===" Cyan
Escrever ""
Escrever "Abra pelo atalho '$NomeAtalho' na Area de Trabalho." White
Escrever "Para atualizar depois, rode esta mesma linha de novo." DarkGray
Escrever ""

Start-Process $pyw -ArgumentList ('"' + (Join-Path $Destino 'main.py') + '"') -WorkingDirectory $Destino
