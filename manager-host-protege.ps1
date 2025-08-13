# Requisitos:
# - PowerShell com módulo ActiveDirectory disponível
# - Permissão para consultar o AD

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configurações
$LogPath = "$ENV:USERPROFILE\Computadores_AD.txt"

function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Módulo ActiveDirectory não encontrado. Instale-o e tente novamente.",
            "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return $false
    }
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    return $true
}

function Ensure-LogFolder {
    $dir = Split-Path $LogPath
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

# Formulário
$form = New-Object System.Windows.Forms.Form
$form.Text = "Consulta de DN (Get-ADComputer)"
$form.Size = New-Object System.Drawing.Size(560, 260)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Label host
$lblHost = New-Object System.Windows.Forms.Label
$lblHost.Text = "Nome do computador:"
$lblHost.Location = New-Object System.Drawing.Point(15, 20)
$lblHost.AutoSize = $true
$form.Controls.Add($lblHost)

# TextBox host
$txtHost = New-Object System.Windows.Forms.TextBox
$txtHost.Location = New-Object System.Drawing.Point(15, 45)
$txtHost.Width = 400
$form.Controls.Add($txtHost)

# Botão consultar
$btnBuscar = New-Object System.Windows.Forms.Button
$btnBuscar.Text = "Buscar e salvar"
$btnBuscar.Location = New-Object System.Drawing.Point(430, 42)
$btnBuscar.Width = 100
$form.Controls.Add($btnBuscar)

# Label DN
$lblDN = New-Object System.Windows.Forms.Label
$lblDN.Text = "DistinguishedName:"
$lblDN.Location = New-Object System.Drawing.Point(15, 85)
$lblDN.AutoSize = $true
$form.Controls.Add($lblDN)

# TextBox DN (somente leitura)
$txtDN = New-Object System.Windows.Forms.TextBox
$txtDN.Location = New-Object System.Drawing.Point(15, 110)
$txtDN.Width = 515
$txtDN.Height = 60
$txtDN.Multiline = $true
$txtDN.ReadOnly = $true
$form.Controls.Add($txtDN)

# Label log
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log: $LogPath"
$lblLog.Location = New-Object System.Drawing.Point(15, 180)
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

# Ação do botão
$btnBuscar.Add_Click({
    $hostName = ($txtHost.Text).Trim()

    if ([string]::IsNullOrWhiteSpace($hostName)) {
        [System.Windows.Forms.MessageBox]::Show("Informe o nome do computador.", "Atenção",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    if (-not (Ensure-ADModule)) { return }

    try {
        $computer = Get-ADComputer -Identity $hostName -Properties DistinguishedName -ErrorAction Stop
        if ($null -ne $computer) {
            $dn = $computer.DistinguishedName
            $txtDN.Text = $dn

            Ensure-LogFolder
            $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            "$timestamp - $hostName - $dn" | Out-File -FilePath $LogPath -Append -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("DistinguishedName encontrado e salvo no log.",
                "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } else {
            $txtDN.Text = ""
            [System.Windows.Forms.MessageBox]::Show("Computador não encontrado no AD.",
                "Não encontrado", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    }
    catch {
        $txtDN.Text = ""
        [System.Windows.Forms.MessageBox]::Show("Erro ao consultar o AD: $($_.Exception.Message)",
            "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

# Tecla Enter ativa o botão
$form.AcceptButton = $btnBuscar

# Foco inicial
$form.Add_Shown({ $txtHost.Focus() })

# Exibe a janela
[void]$form.ShowDialog()
