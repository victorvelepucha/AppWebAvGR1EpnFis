<#
.SYNOPSIS
    Despliega la aplicacion PoliMusic (ASP clasico) en IIS.

.DESCRIPTION
    Realiza el despliegue completo y de forma idempotente:

      1. Comprueba privilegios de administrador (si no los tiene, se
         relanza a si mismo solicitando elevacion por UAC).
      2. Habilita las caracteristicas de IIS necesarias, incluida
         IIS-ASP (ASP clasico), que no viene activada por defecto.
      3. Crea un grupo de aplicaciones dedicado (sin codigo administrado).
      4. Copia la aplicacion a la carpeta de destino.
      5. Ajusta los permisos NTFS para la identidad del grupo de aplicaciones.
      6. Registra la aplicacion en IIS y configura el documento predeterminado
         y la pagina de error 404.
      7. Verifica el despliegue con peticiones HTTP reales.

.PARAMETER Destino
    Carpeta raiz de IIS. Por defecto C:\inetpub\wwwroot.
    La aplicacion se publica en una subcarpeta llamada PoliMusic.

.PARAMETER EnLaRaiz
    Publica la aplicacion directamente en la raiz del sitio (sin subcarpeta).

.PARAMETER Debug
    Activa scriptErrorSentToBrowser para ver los errores de ASP en el navegador.
    Usar solo durante la puesta a punto.

.PARAMETER SaltarCaracteristicas
    Omite la habilitacion de caracteristicas de Windows (util si ya estan
    instaladas o si no se dispone de conexion a Windows Update).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\desplegar.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\desplegar.ps1 -EnLaRaiz -Debug
#>
[CmdletBinding()]
param(
    [string] $Destino = "C:\inetpub\wwwroot",
    [switch] $EnLaRaiz,
    [switch] $Debug,
    [switch] $SaltarCaracteristicas
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
#  0. Elevacion
# ---------------------------------------------------------------------------
function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host ""
    Write-Host "  Se requieren privilegios de administrador." -ForegroundColor Yellow
    Write-Host "  Solicitando elevacion (aparecera el aviso de UAC)..." -ForegroundColor Yellow
    Write-Host ""
    $argsList = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`"",
        "-Destino", "`"$Destino`""
    )
    if ($EnLaRaiz) { $argsList += "-EnLaRaiz" }
    if ($Debug) { $argsList += "-Debug" }
    if ($SaltarCaracteristicas) { $argsList += "-SaltarCaracteristicas" }
    Start-Process -FilePath "powershell.exe" -ArgumentList $argsList -Verb RunAs -Wait
    exit $LASTEXITCODE
}

# ---------------------------------------------------------------------------
#  Utilidades de salida
# ---------------------------------------------------------------------------
function Paso  ($t) { Write-Host ""; Write-Host "==> $t" -ForegroundColor Cyan }
function Ok    ($t) { Write-Host "    [OK]   $t" -ForegroundColor Green }
function Info  ($t) { Write-Host "    [info] $t" -ForegroundColor Gray }
function Aviso ($t) { Write-Host "    [!]    $t" -ForegroundColor Yellow }
function Error ($t) { Write-Host "    [ERROR] $t" -ForegroundColor Red }

$origen  = $PSScriptRoot
$nombreApp = "PoliMusic"
if ($EnLaRaiz) {
    $raizVirtual = "/"
    $rutaFisica  = $Destino
} else {
    $raizVirtual = "/$nombreApp"
    $rutaFisica  = Join-Path $Destino $nombreApp
}
$poolNombre = "PoliMusicPool"
$urlBase    = "http://localhost" + ($(if ($EnLaRaiz) { "/" } else { "/$nombreApp/" }))

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Blue
Write-Host "   PoliMusic - Despliegue en IIS" -ForegroundColor Blue
Write-Host "  ============================================================" -ForegroundColor Blue
Write-Host ""
Write-Host "   Origen   : $origen"
Write-Host "   Destino  : $rutaFisica"
Write-Host "   URL      : $urlBase"

$fallos = 0

# ---------------------------------------------------------------------------
#  1. Caracteristicas de IIS
# ---------------------------------------------------------------------------
Paso "1/7  Habilitando caracteristicas de IIS"

$caracteristicas = @(
    "IIS-WebServerRole",
    "IIS-WebServer",
    "IIS-CommonHttpFeatures",
    "IIS-StaticContent",
    "IIS-DefaultDocument",
    "IIS-HttpErrors",
    "IIS-ApplicationDevelopment",
    "IIS-ISAPIExtensions",
    "IIS-ASP",                  # <-- ASP clasico
    "IIS-RequestFiltering",
    "IIS-HttpCompressionStatic",
    "IIS-HttpCompressionDynamic",
    "IIS-ManagementConsole"
)

if ($SaltarCaracteristicas) {
    Info "Omitido por parametro -SaltarCaracteristicas."
} else {
    $dism = Get-Command Enable-WindowsOptionalFeature -ErrorAction SilentlyContinue
    foreach ($c in $caracteristicas) {
        try {
            if ($dism) {
                $st = (Get-WindowsOptionalFeature -Online -FeatureName $c -ErrorAction Stop).State
                if ($st -eq "Enabled") { Info "$c ya estaba habilitada."; continue }
                Enable-WindowsOptionalFeature -Online -FeatureName $c -All -NoRestart -ErrorAction Stop | Out-Null
                Ok "$c habilitada."
            } else {
                $r = & dism.exe /online /enable-feature /featurename:$c /all /norestart 2>&1
                if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 3010) { Ok "$c habilitada." }
                else { Aviso "$c no se pudo habilitar (codigo $LASTEXITCODE)."; $fallos++ }
            }
        } catch {
            Aviso "$c : $($_.Exception.Message)"
            $fallos++
        }
    }
}

# Verificacion explicita de la caracteristica ASP clasico
$schemaAsp = Join-Path $env:WINDIR "System32\inetsrv\config\schema\ASP_schema.xml"
if (Test-Path $schemaAsp) {
    Ok "ASP clasico esta instalado (ASP_schema.xml presente)."
} else {
    Error "ASP clasico NO esta instalado: falta ASP_schema.xml."
    Error "Ejecuta:  dism /online /enable-feature /featurename:IIS-ASP /all"
    $fallos++
}

# ---------------------------------------------------------------------------
#  2. Servicios de IIS
# ---------------------------------------------------------------------------
Paso "2/7  Comprobando los servicios de IIS"
foreach ($sv in @("W3SVC", "WAS")) {
    $s = Get-Service -Name $sv -ErrorAction SilentlyContinue
    if (-not $s) { Error "El servicio $sv no existe."; $fallos++; continue }
    if ($s.Status -ne "Running") {
        Start-Service -Name $sv
        Ok "Servicio $sv iniciado."
    } else {
        Ok "Servicio $sv en ejecucion."
    }
}

# ---------------------------------------------------------------------------
#  3. Copia de archivos
# ---------------------------------------------------------------------------
Paso "3/7  Copiando la aplicacion"

if (-not (Test-Path $rutaFisica)) {
    New-Item -ItemType Directory -Path $rutaFisica -Force | Out-Null
    Ok "Carpeta creada: $rutaFisica"
}

# /MIR deja el destino identico al origen; se excluyen los archivos de trabajo
# (pruebas, vista previa, scripts y documentacion) para publicar solo la aplicacion.
$robocopyArgs = @(
    $origen, $rutaFisica, "/MIR", "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/R:1", "/W:1",
    "/XD", "_asp_tests", "_preview", "_capturas", ".workbuddy-ai", "__pycache__", "SongFilesAndImages_resources",
    "/XF", "*.py", "*.ps1", "*.bat", "*.sql", "*.md"
)
$rc = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -Wait -PassThru -NoNewWindow
if ($rc.ExitCode -ge 8) {
    Error "robocopy devolvio el codigo $($rc.ExitCode)."
    $fallos++
} else {
    $n = (Get-ChildItem -Path $rutaFisica -Recurse -File).Count
    Ok "Aplicacion copiada ($n archivos)."
}

# ---------------------------------------------------------------------------
#  4. Permisos
# ---------------------------------------------------------------------------
Paso "4/7  Ajustando permisos NTFS"

$identidad = "IIS AppPool\$poolNombre"
& icacls.exe $rutaFisica /grant "${identidad}:(OI)(CI)(RX)" /T /Q | Out-Null
if ($LASTEXITCODE -eq 0) { Ok "Lectura y ejecucion concedidas a $identidad." }
else { Aviso "No se pudieron asignar permisos a $identidad (codigo $LASTEXITCODE)." }

& icacls.exe $rutaFisica /grant "IIS_IUSRS:(OI)(CI)(RX)" /T /Q | Out-Null
if ($LASTEXITCODE -eq 0) { Ok "Lectura y ejecucion concedidas a IIS_IUSRS." }

# ---------------------------------------------------------------------------
#  5. Grupo de aplicaciones y aplicacion en IIS
# ---------------------------------------------------------------------------
Paso "5/7  Registrando la aplicacion en IIS"

Import-Module WebAdministration -ErrorAction SilentlyContinue

if (Test-Path "IIS:\AppPools\$poolNombre") {
    Info "El grupo de aplicaciones $poolNombre ya existe."
} else {
    New-WebAppPool -Name $poolNombre | Out-Null
    Ok "Grupo de aplicaciones $poolNombre creado."
}
Set-ItemProperty "IIS:\AppPools\$poolNombre" -Name managedRuntimeVersion -Value ""
Set-ItemProperty "IIS:\AppPools\$poolNombre" -Name enable32BitAppOnWin64 -Value $false
Set-ItemProperty "IIS:\AppPools\$poolNombre" -Name processModel.identityType -Value "ApplicationPoolIdentity"
Set-ItemProperty "IIS:\AppPools\$poolNombre" -Name recycling.periodicRestart.time -Value "00:00:00"
Ok "Grupo de aplicaciones configurado (sin codigo administrado, 64 bits)."

if ($EnLaRaiz) {
    $sitio = (Get-Website | Where-Object { $_.State -eq "Started" } | Select-Object -First 1)
    if ($sitio) {
        Set-ItemProperty "IIS:\Sites\$($sitio.Name)" -Name physicalPath -Value $rutaFisica
        Set-ItemProperty "IIS:\Sites\$($sitio.Name)" -Name applicationPool -Value $poolNombre
        Ok "Sitio '$($sitio.Name)' apuntando a $rutaFisica."
    } else {
        Error "No se encontro ningun sitio IIS iniciado."
        $fallos++
    }
} else {
    if (Test-Path "IIS:\Sites\Default Web Site\$nombreApp") {
        Info "La aplicacion $raizVirtual ya existe."
        Set-ItemProperty "IIS:\Sites\Default Web Site\$nombreApp" -Name physicalPath -Value $rutaFisica
        Set-ItemProperty "IIS:\Sites\Default Web Site\$nombreApp" -Name applicationPool -Value $poolNombre
    } else {
        New-WebApplication -Site "Default Web Site" -Name $nombreApp `
            -PhysicalPath $rutaFisica -ApplicationPool $poolNombre | Out-Null
        Ok "Aplicacion $raizVirtual creada en 'Default Web Site'."
    }
}

# Documento predeterminado
$webConfig = Join-Path $rutaFisica "web.config"
if (Test-Path $webConfig) { Ok "web.config presente." } else { Error "Falta web.config."; $fallos++ }

# Pagina de error 404 personalizada
try {
    $prefijo = $(if ($EnLaRaiz) { "" } else { "/$nombreApp" })
    & "$env:WINDIR\System32\inetsrv\appcmd.exe" set config `
        "$(if ($EnLaRaiz) { 'Default Web Site' } else { 'Default Web Site/PoliMusic' })" `
        -section:system.webServer/httpErrors `
        "/+[statusCode='404',subStatusCode='-1',path='$prefijo/error404.asp',responseMode='ExecuteURL']" | Out-Null
    Ok "Pagina de error 404 personalizada registrada."
} catch {
    Aviso "No se pudo registrar la pagina 404 personalizada."
}

# Modo depuracion
if ($Debug) {
    $contenido = Get-Content $webConfig -Raw
    $contenido = $contenido -replace 'scriptErrorSentToBrowser="false"', 'scriptErrorSentToBrowser="true"'
    Set-Content -Path $webConfig -Value $contenido -Encoding UTF8
    Ok "Modo depuracion activado (scriptErrorSentToBrowser=true)."
}

# ---------------------------------------------------------------------------
#  6. Reinicio de IIS
# ---------------------------------------------------------------------------
Paso "6/7  Reiniciando IIS"
& "$env:WINDIR\System32\inetsrv\appcmd.exe" recycle apppool /apppool.name:"$poolNombre" | Out-Null
& iisreset.exe /noforce | Out-Null
Ok "IIS reiniciado."

# ---------------------------------------------------------------------------
#  7. Verificacion
# ---------------------------------------------------------------------------
Paso "7/7  Verificando el despliegue"

Start-Sleep -Seconds 3

$paginas = @(
    @{ Url = $urlBase + "login.asp";  Debe = "Iniciar sesi" },
    @{ Url = $urlBase + "default.asp"; Debe = "login.asp" },
    @{ Url = $urlBase + "css/style.css"; Debe = "PoliMusic" },
    @{ Url = $urlBase + "assets/img/logo.svg"; Debe = "svg" }
)

foreach ($p in $paginas) {
    try {
        $r = Invoke-WebRequest -Uri $p.Url -UseBasicParsing -TimeoutSec 20
        $cuerpo = [string]$r.Content
        if ($cuerpo -like "*$($p.Debe)*" -or $r.StatusCode -eq 302) {
            Ok "$($p.Url)  ->  HTTP $($r.StatusCode)"
        } else {
            Aviso "$($p.Url) respondio $($r.StatusCode) pero el contenido no parece correcto."
            $fallos++
        }
    } catch {
        Error "$($p.Url)  ->  $($_.Exception.Message)"
        $fallos++
    }
}

# Comprobacion de la cadena de conexion a SQL Server
Paso "Extra  Comprobando la conexion con SQL Server"
try {
    $cn = New-Object System.Data.SqlClient.SqlConnection
    $cn.ConnectionString = "Server=localhost;Database=BDD_PoliMusic;User Id=usr_polimusic;Password=Politecnica1;Encrypt=true;TrustServerCertificate=true;"
    $cn.Open()
    $cmd = $cn.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) FROM TBL_SONG"
    $n = $cmd.ExecuteScalar()
    $cn.Close()
    Ok "Conexion correcta. Canciones en la base: $n"
} catch {
    Aviso "No se pudo validar la conexion: $($_.Exception.Message)"
}

# ---------------------------------------------------------------------------
#  Resumen
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Blue
if ($fallos -eq 0) {
    Write-Host "   DESPLIEGUE COMPLETADO CORRECTAMENTE" -ForegroundColor Green
} else {
    Write-Host "   DESPLIEGUE FINALIZADO CON $fallos ADVERTENCIA(S)" -ForegroundColor Yellow
}
Write-Host "  ============================================================" -ForegroundColor Blue
Write-Host ""
Write-Host "   Abre en el navegador:  $urlBase"
Write-Host ""
Write-Host "   Cuentas de prueba:"
Write-Host "     Administrador : userAdmin / 12345678"
Write-Host "     Usuario normal: epn.fis   / 12345678"
Write-Host ""
Write-Host "   Registro de errores de IIS:"
Write-Host "     %SystemDrive%\inetpub\logs\LogFiles\W3SVC1"
Write-Host ""
