# 1. PREPARACIÓN DE ENTORNO
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Elevar a administrador
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 2. PUENTE COM
$source = @"
using System;
using System.Runtime.InteropServices;
[ComVisible(true)]
public class ZebraBridge {
    public delegate void ActionString(string val);
    public event ActionString OnDpiSelected;
    public event ActionString OnInstallRequested;
    public event EventHandler OnCloseRequested;
    public event EventHandler OnStopSpooler;
    public event EventHandler OnStartSpooler;

    public void SetDPI(string dpi) { if (OnDpiSelected != null) OnDpiSelected(dpi); }
    public void Instalar(string opcion) { if (OnInstallRequested != null) OnInstallRequested(opcion); }
    public void Cerrar() { if (OnCloseRequested != null) OnCloseRequested(this, EventArgs.Empty); }
    public void StopSpooler() { if (OnStopSpooler != null) OnStopSpooler(this, EventArgs.Empty); }
    public void StartSpooler() { if (OnStartSpooler != null) OnStartSpooler(this, EventArgs.Empty); }
}
"@
Add-Type -TypeDefinition $source -ReferencedAssemblies "System.Windows.Forms"

# 3. VARIABLES
$Global:DPI_Seleccionado = 203
$CarpetaDRS = $PSScriptRoot

# 4a-0. DIÁLOGO DE IP SIEMPRE AL FRENTE
function Get-NetworkDialog {
    # Devuelve hashtable con IP — o $null si se cancela

    $dlgForm = New-Object System.Windows.Forms.Form
    $dlgForm.Text            = "Instalación Ethernet — Impresora Zebra (VLAN 102)"
    $dlgForm.Size            = New-Object System.Drawing.Size(400, 160)
    $dlgForm.StartPosition   = "CenterScreen"
    $dlgForm.FormBorderStyle = "FixedDialog"
    $dlgForm.MaximizeBox     = $false
    $dlgForm.MinimizeBox     = $false
    $dlgForm.TopMost         = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = "No se detectó Zebra por USB.`nIntroduce la IP de la impresora (VLAN 102):"
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size     = New-Object System.Drawing.Size(370, 36)
    $dlgForm.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text     = "192.168.102."
    $txt.Location = New-Object System.Drawing.Point(12, 54)
    $txt.Size     = New-Object System.Drawing.Size(360, 22)
    # Colocar cursor al final para que el operario escriba solo el último octeto
    $txt.Add_GotFocus({ $txt.SelectionStart = $txt.Text.Length })
    $dlgForm.Controls.Add($txt)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text         = "Aceptar"
    $btnOK.Location     = New-Object System.Drawing.Point(196, 86)
    $btnOK.Size         = New-Object System.Drawing.Size(84, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlgForm.AcceptButton = $btnOK
    $dlgForm.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = "Cancelar"
    $btnCancel.Location     = New-Object System.Drawing.Point(290, 86)
    $btnCancel.Size         = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlgForm.CancelButton   = $btnCancel
    $dlgForm.Controls.Add($btnCancel)

    $result = $dlgForm.ShowDialog()
    $dlgForm.Dispose()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

    return @{ IP = $txt.Text.Trim() }
}

# 4a. INSTALACIÓN AUTOMÁTICA DE DRIVER
function Confirm-Driver {
    param([string]$DriverName)

    # Ya instalado → OK
    if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) { return $true }

    $InfPath = Join-Path $PSScriptRoot "ZBRN.inf"

    # Intento 1: Add-PrinterDriver con el INF local (funciona si los .dll están en la misma carpeta)
    if (Test-Path $InfPath) {
        try {
            Add-PrinterDriver -Name $DriverName -InfPath $InfPath -ErrorAction Stop
            if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
                return $true
            }
        } catch { }
    }

    # Intento 2: Driver ya en caché de Windows (si la impresora estuvo conectada antes por USB)
    try {
        Add-PrinterDriver -Name $DriverName -ErrorAction Stop
        if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
            return $true
        }
    } catch { }

    # Intento 3: Descargar e instalar Zebra Setup Utilities en silencio
    try {
        $zsUrl  = "https://www.zebra.com/content/dam/zebra_dam/global/software/printer/zdesigner/windows/ZebraSetupUtilities.exe"
        $zsTemp = Join-Path $env:TEMP "ZebraSetupUtilities.exe"

        $dlg = [System.Windows.Forms.MessageBox]::Show(
            "El controlador '$DriverName' no está instalado.`n`n" +
            "¿Descargar e instalar Zebra Setup Utilities automáticamente?`n" +
            "(requiere conexión a internet, ~80 MB)",
            "Driver no encontrado",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($dlg -eq [System.Windows.Forms.DialogResult]::Yes) {
            [System.Windows.Forms.MessageBox]::Show(
                "Descargando drivers Zebra... Esto puede tardar unos minutos.`nEl instalador se abrirá al terminar.",
                "Descargando",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            $wc = New-Object System.Net.WebClient
            $wc.DownloadFile($zsUrl, $zsTemp)

            # Instalación silenciosa de ZSU (incluye todos los drivers ZDesigner)
            Start-Process -FilePath $zsTemp -ArgumentList "/S" -Wait
            Remove-Item $zsTemp -Force -ErrorAction SilentlyContinue

            if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) {
                return $true
            }

            # ZSU instalado pero driver concreto aún no en spooler → agregar
            $inf2 = Join-Path ${env:ProgramFiles(x86)} "Zebra Technologies\ZebraDesigner\Drivers\ZBRN.inf"
            if (-not (Test-Path $inf2)) {
                $inf2 = Join-Path $env:ProgramFiles "Zebra Technologies\ZebraDesigner\Drivers\ZBRN.inf"
            }
            if (Test-Path $inf2) {
                Add-PrinterDriver -Name $DriverName -InfPath $inf2 -ErrorAction SilentlyContinue
            }

            if (Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue) { return $true }
        }
    } catch {
        # La descarga falló o el usuario canceló
    }

    # Todo falló → guiar al usuario
    $res = [System.Windows.Forms.MessageBox]::Show(
        "No se pudo instalar el controlador '$DriverName' automáticamente.`n`n" +
        "Instala manualmente 'Zebra Setup Utilities' desde zebra.com y vuelve a ejecutar este script.`n`n" +
        "¿Abrir la página de descarga ahora?",
        "Driver no encontrado",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-Process "https://www.zebra.com/us/en/support-downloads/software/printer-software/zdesigner-driver.html"
    }
    return $false
}

# 4b. LÓGICA DE INSTALACIÓN (USB O ETHERNET MANUAL)
function Invoke-Instalacion {
    param($opcion)

    $perfiles = @{
        "1" = "ZDesigner Nissan Corto.drs"
        "2" = "ZDesigner Nissan Largo.drs"
        "3" = "ZDesigner MQB Bandera.drs"
        "4" = "ZDesigner VQ QR Brida.drs"
    }

    $nombres = @{
        "1" = "Nissan Corto"
        "2" = "Nissan Largo"
        "3" = "MQB Bandera"
        "4" = "VQ QR Brida"
    }

    if (-not $perfiles.ContainsKey($opcion)) { return }
    $ArchivoNombre   = $perfiles[$opcion]
    $NombreImpresora = "$($nombres[$opcion]) $($Global:DPI_Seleccionado)dpi"
    $Driver = if ($Global:DPI_Seleccionado -eq 300) { "ZDesigner ZT411-300dpi ZPL" } else { "ZDesigner ZT410R-203dpi ZPL" }

    try {
        # --- ASEGURAR QUE EL SPOOLER ESTÁ ACTIVO ---
        $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
        if ($spooler -and $spooler.Status -ne 'Running') {
            Start-Service -Name Spooler -ErrorAction Stop
            Start-Sleep -Seconds 2
        }

        # --- DETECCIÓN INTELIGENTE: DISPOSITIVO FÍSICO ZEBRA (fuente de verdad) ---
        # Re-enumerar dispositivos antes de buscar
        Start-Process -FilePath "pnputil.exe" -ArgumentList "/scan-devices" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        # Buscar dispositivo Zebra físico por USB en WMI — ignora otros USBs conectados
        $dispositivoZebra = Get-WmiObject Win32_PnPEntity -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -match "Zebra|ZDesigner|ZT410|ZT411" -and
                $_.DeviceID -match "^USB\\"
            } | Select-Object -First 1

        $Puerto = $null
        $Modo   = $null

        if ($dispositivoZebra) {
            # Zebra conectada por USB → buscar su puerto en el spooler
            $PuertoUSB = Get-PrinterPort -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -like "USB*" } |
                Sort-Object Name -Descending |
                Select-Object -First 1 -ExpandProperty Name

            if (-not $PuertoUSB) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Se detectó una impresora Zebra por USB pero Windows aún no ha creado el puerto.`n`nDesconecta y vuelve a conectar el cable USB, espera unos segundos y vuelve a intentarlo.",
                    "Puerto USB pendiente",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            $Puerto = $PuertoUSB
            $Modo   = "USB ($Puerto)"

            # Corregir automáticamente cualquier impresora ZDesigner que apunte a un puerto IP erróneo
            $impresorasConPuertoErroneo = Get-Printer -ErrorAction SilentlyContinue |
                Where-Object { $_.DriverName -like "ZDesigner*" -and $_.PortName -notlike "USB*" }

            foreach ($imp in $impresorasConPuertoErroneo) {
                Set-Printer -Name $imp.Name -PortName $Puerto -ErrorAction SilentlyContinue
            }

        } else {
            # No hay Zebra por USB → Ethernet, pedir configuración de red
            $RedConfig = Get-NetworkDialog

            if ($null -eq $RedConfig) {
                [System.Windows.Forms.MessageBox]::Show("Instalación cancelada.", "Aviso")
                return
            }

            $IP = $RedConfig.IP

            if ($IP -notmatch '^\d{1,3}(\.\d{1,3}){3}$') {
                [System.Windows.Forms.MessageBox]::Show("Dirección IP no válida: '$IP'.", "Error")
                return
            }

            $Puerto = "IP_$IP"
            $Modo   = "Ethernet (IP: $IP — VLAN 102)"

            if (-not (Get-PrinterPort -Name $Puerto -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $Puerto -PrinterHostAddress $IP -ErrorAction Stop
            }

        }

        # Verificación e instalación automática de driver
        if (-not (Confirm-Driver -DriverName $Driver)) { return }

        # Eliminar impresoras auto-creadas por Windows con nombre del driver (p.ej. "ZDesigner ZT410R-203dpi ZPL")
        Get-Printer | Where-Object { $_.DriverName -eq $Driver -and $_.Name -ne $NombreImpresora } | ForEach-Object {
            Remove-Printer -Name $_.Name -ErrorAction SilentlyContinue
        }

        # Reinstalar impresora con nombre personalizado
        if (Get-Printer -Name $NombreImpresora -ErrorAction SilentlyContinue) {
            Remove-Printer -Name $NombreImpresora
            Start-Sleep -Seconds 1
        }

        Add-Printer -Name $NombreImpresora -DriverName $Driver -PortName $Puerto -ErrorAction Stop

        # Enviar configuración DRS
        $RutaArchivo = Join-Path $CarpetaDRS $ArchivoNombre
        if (Test-Path $RutaArchivo) {
            $tempFile = [System.IO.Path]::GetTempFileName()
            try {
                [System.IO.File]::Copy($RutaArchivo, $tempFile, $true)
                $printerEsc = $NombreImpresora -replace '"', '\"'
                $tempEsc    = $tempFile        -replace '"', '\"'
                Start-Process cmd.exe -ArgumentList "/c copy /b `"$tempEsc`" `"\\.\$printerEsc`"" -WindowStyle Hidden -Wait
            } finally {
                if (Test-Path $tempFile) { Remove-Item $tempFile -Force }
            }
        }

        [System.Windows.Forms.MessageBox]::Show(
            "Éxito: '$NombreImpresora' lista en $Puerto.`nModo: $Modo",
            "Zebra Config",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# 5. MOTOR DE INTERFAZ
function New-Interfaz {
    param($Titulo, $Html, $Alto)

    $form = New-Object Windows.Forms.Form
    $form.Text = $Titulo
    $form.Size = New-Object Drawing.Size(420, $Alto)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.TopMost = $true

    $browser = New-Object Windows.Forms.WebBrowser
    $browser.Dock = "Fill"
    $browser.IsWebBrowserContextMenuEnabled = $false
    $browser.ScriptErrorsSuppressed = $true

    $bridge = New-Object ZebraBridge
    $bridge.add_OnDpiSelected({
        param($d)
        $Global:DPI_Seleccionado = [int]$d
        $form.Close()
    })
    $bridge.add_OnInstallRequested({
        param($o)
        Invoke-Instalacion -opcion $o
    })
    $bridge.add_OnCloseRequested({
        $form.Close()
    })
    $bridge.add_OnStopSpooler({
        try {
            Stop-Service -Name Spooler -Force -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show(
                "Servicio de cola de impresión detenido.",
                "Spooler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al detener Spooler: $($_.Exception.Message)", "Error")
        }
    })
    $bridge.add_OnStartSpooler({
        try {
            Start-Service -Name Spooler -ErrorAction Stop
            [System.Windows.Forms.MessageBox]::Show(
                "Servicio de cola de impresión iniciado.",
                "Spooler",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al iniciar Spooler: $($_.Exception.Message)", "Error")
        }
    })

    $browser.ObjectForScripting = $bridge
    $form.Controls.Add($browser)
    $browser.DocumentText = $Html
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# 6. HTML — SELECCIÓN DE DPI
$htmlDPI = @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta charset="utf-8">
<style>
  body { font-family:'Segoe UI',sans-serif; background:#1a1a2e; color:#eee;
         text-align:center; padding:24px 16px; margin:0; user-select:none; }
  h2   { color:#00d4ff; margin:0 0 24px; font-size:18px; letter-spacing:1px; }
  .btn { display:block; width:78%; margin:12px auto; padding:14px;
         font-size:15px; border:none; border-radius:8px; cursor:pointer;
         font-weight:600; transition:background .15s; }
  .b203 { background:#0066cc; color:#fff; }
  .b203:hover { background:#0088ff; }
  .b300 { background:#cc6600; color:#fff; }
  .b300:hover { background:#ff8800; }
</style>
</head>
<body>
  <h2>Selecciona la resoluci&oacute;n DPI</h2>
  <button class="btn b203" onclick="window.external.SetDPI('203')">203 DPI — ZT410R</button>
  <button class="btn b300" onclick="window.external.SetDPI('300')">300 DPI — ZT411</button>
</body>
</html>
"@

# Pantalla DPI primero
New-Interfaz -Titulo "Zebra Config - Resolucion" -Html $htmlDPI -Alto 220

# 7. HTML — PANEL PRINCIPAL (se construye DESPUÉS de elegir DPI)
$dpiActual = $Global:DPI_Seleccionado
$htmlMain = @"
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta charset="utf-8">
<style>
  body { font-family:'Segoe UI',sans-serif; background:#1a1a2e; color:#eee;
         text-align:center; padding:18px 12px; margin:0; user-select:none; }
  h2        { color:#00d4ff; margin:0 0 4px; font-size:17px; letter-spacing:1px; }
  .dpi-info { color:#aaa; font-size:12px; margin-bottom:14px; }
  .btn { display:block; width:82%; margin:7px auto; padding:11px;
         font-size:13px; border:none; border-radius:7px; cursor:pointer;
         font-weight:600; transition:background .15s; }
  .bi  { background:#006633; color:#fff; }  .bi:hover  { background:#008844; }
  .bs  { background:#882200; color:#fff; }  .bs:hover  { background:#bb3300; }
  .bst { background:#003388; color:#fff; }  .bst:hover { background:#0044bb; }
  .bc  { background:#333;    color:#ccc; }  .bc:hover  { background:#555; }
  hr   { border:none; border-top:1px solid #2a2a4a; margin:10px auto; width:80%; }
</style>
</head>
<body>
  <h2>Zebra Config</h2>
  <div class="dpi-info">DPI seleccionado: $dpiActual</div>

  <button class="btn bi" onclick="window.external.Instalar('1')">Nissan Corto</button>
  <button class="btn bi" onclick="window.external.Instalar('2')">Nissan Largo</button>
  <button class="btn bi" onclick="window.external.Instalar('3')">MQB Bandera</button>
  <button class="btn bi" onclick="window.external.Instalar('4')">VQ QR Brida</button>

  <hr>
  <button class="btn bs"  onclick="window.external.StopSpooler()">&#9632; Detener Spooler</button>
  <button class="btn bst" onclick="window.external.StartSpooler()">&#9654; Iniciar Spooler</button>
  <hr>
  <button class="btn bc"  onclick="window.external.Cerrar()">Cerrar</button>
</body>
</html>
"@

# 8. PANEL PRINCIPAL
New-Interfaz -Titulo "Zebra Config - $dpiActual DPI" -Html $htmlMain -Alto 460