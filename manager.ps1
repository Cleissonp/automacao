# Requisito: Garante que o módulo do Active Directory esteja disponível
#requires -Modules ActiveDirectory

# Carregar Assemblies necessários para criar a interface
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Reflection

# Criar a Janela do Formulário
$form = New-Object system.windows.forms.Form
$form.Text = "Ferramenta de Gerenciamento"
$form.Size = New-Object System.Drawing.Size(850, 650)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Criar TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$form.Controls.Add($tabControl)

# Criar TabPage para Gerenciamento de Contas (Aba 1)
$tabPageGerenciar = New-Object System.Windows.Forms.TabPage
$tabPageGerenciar.Text = "Gerenciar Contas"
$tabControl.Controls.Add($tabPageGerenciar)

# Criar TabPage para Remover Host (Aba 2)
$tabPageRemoverHost = New-Object System.Windows.Forms.TabPage
$tabPageRemoverHost.Text = "Remover Host"
$tabControl.Controls.Add($tabPageRemoverHost)

# Criar TabPage para Spooler de Impressão (Aba 3)
$tabPageSpooler = New-Object System.Windows.Forms.TabPage
$tabPageSpooler.Text = "Spooler de Impressão"
$tabControl.Controls.Add($tabPageSpooler)

# -----------------------------------------------------------------------------------
# Conteúdo da Aba "Gerenciar Contas"
# -----------------------------------------------------------------------------------
$startX_Gerenciar = 50
$startY_Gerenciar = 20
$spacingY_Gerenciar = 40

$labelUser_Gerenciar = New-Object system.windows.forms.Label
$labelUser_Gerenciar.Text = "Login do Usuário (ex: joao.silva):"
$labelUser_Gerenciar.AutoSize = $true
$labelUser_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$labelUser_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageGerenciar.Controls.Add($labelUser_Gerenciar)
$startY_Gerenciar += 30

$textboxUser_Gerenciar = New-Object system.windows.forms.TextBox
$textboxUser_Gerenciar.Size = New-Object System.Drawing.Size(450, 30)
$textboxUser_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$textboxUser_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageGerenciar.Controls.Add($textboxUser_Gerenciar)
$startY_Gerenciar += $spacingY_Gerenciar

$labelAD_Gerenciar = New-Object system.windows.forms.Label
$labelAD_Gerenciar.Text = "Servidor AD:"
$labelAD_Gerenciar.AutoSize = $true
$labelAD_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$labelAD_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageGerenciar.Controls.Add($labelAD_Gerenciar)
$startY_Gerenciar += 30

$comboBoxAD_Gerenciar = New-Object system.windows.forms.ComboBox
$comboBoxAD_Gerenciar.Size = New-Object System.Drawing.Size(450, 30)
$comboBoxAD_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$comboBoxAD_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$comboBoxAD_Gerenciar.Items.AddRange(@("vmad01-new", "hospital.unimedangra.com.br"))
$comboBoxAD_Gerenciar.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabPageGerenciar.Controls.Add($comboBoxAD_Gerenciar)
$startY_Gerenciar += $spacingY_Gerenciar

$labelAction_Gerenciar = New-Object system.windows.forms.Label
$labelAction_Gerenciar.Text = "Ação a Executar:"
$labelAction_Gerenciar.AutoSize = $true
$labelAction_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$labelAction_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageGerenciar.Controls.Add($labelAction_Gerenciar)
$startY_Gerenciar += 30

$comboBoxAction_Gerenciar = New-Object system.windows.forms.ComboBox
$comboBoxAction_Gerenciar.Size = New-Object System.Drawing.Size(450, 30)
$comboBoxAction_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$comboBoxAction_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$comboBoxAction_Gerenciar.Items.AddRange(@("Resetar Senha", "Desbloquear Usuário", "Desativar Usuário", "Apagar Usuário"))
$comboBoxAction_Gerenciar.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabPageGerenciar.Controls.Add($comboBoxAction_Gerenciar)
$startY_Gerenciar += $spacingY_Gerenciar + 10

$buttonExecute_Gerenciar = New-Object system.windows.forms.Button
$buttonExecute_Gerenciar.Size = New-Object System.Drawing.Size(150, 40)
$buttonExecute_Gerenciar.Text = "Executar Ação"
$buttonExecute_Gerenciar.Location = New-Object System.Drawing.Point($startX_Gerenciar, $startY_Gerenciar)
$buttonExecute_Gerenciar.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$buttonExecute_Gerenciar.ForeColor = [System.Drawing.Color]::White
$buttonExecute_Gerenciar.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonExecute_Gerenciar.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabPageGerenciar.Controls.Add($buttonExecute_Gerenciar)

# -----------------------------------------------------------------------------------
# Conteúdo da Aba "Remover Host"
# -----------------------------------------------------------------------------------
$startX_Remover = 50
$startY_Remover = 20
$spacingY_Remover = 40

$labelHost_Remover = New-Object system.windows.forms.Label
$labelHost_Remover.Text = "Nome do Host:"
$labelHost_Remover.AutoSize = $true
$labelHost_Remover.Location = New-Object System.Drawing.Point($startX_Remover, $startY_Remover)
$labelHost_Remover.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageRemoverHost.Controls.Add($labelHost_Remover)
$startY_Remover += 30

$textboxHost_Remover = New-Object system.windows.forms.TextBox
$textboxHost_Remover.Size = New-Object System.Drawing.Size(450, 30)
$textboxHost_Remover.Location = New-Object System.Drawing.Point($startX_Remover, $startY_Remover)
$textboxHost_Remover.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageRemoverHost.Controls.Add($textboxHost_Remover)
$startY_Remover += $spacingY_Remover

$labelAD_Remover = New-Object system.windows.forms.Label
$labelAD_Remover.Text = "Servidor AD:"
$labelAD_Remover.AutoSize = $true
$labelAD_Remover.Location = New-Object System.Drawing.Point($startX_Remover, $startY_Remover)
$labelAD_Remover.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageRemoverHost.Controls.Add($labelAD_Remover)
$startY_Remover += 30

$comboBoxAD_Remover = New-Object system.windows.forms.ComboBox
$comboBoxAD_Remover.Size = New-Object System.Drawing.Size(450, 30)
$comboBoxAD_Remover.Location = New-Object System.Drawing.Point($startX_Remover, $startY_Remover)
$comboBoxAD_Remover.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$comboBoxAD_Remover.Items.AddRange(@("vmad01-new", "hospital.unimedangra.com.br"))
$comboBoxAD_Remover.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$tabPageRemoverHost.Controls.Add($comboBoxAD_Remover)
$startY_Remover += $spacingY_Remover

$buttonExecute_Remover = New-Object system.windows.forms.Button
$buttonExecute_Remover.Size = New-Object System.Drawing.Size(150, 40)
$buttonExecute_Remover.Text = "Remover Host"
$buttonExecute_Remover.Location = New-Object System.Drawing.Point($startX_Remover, $startY_Remover)
$buttonExecute_Remover.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)
$buttonExecute_Remover.ForeColor = [System.Drawing.Color]::White
$buttonExecute_Remover.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonExecute_Remover.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabPageRemoverHost.Controls.Add($buttonExecute_Remover)

# -----------------------------------------------------------------------------------
# Conteúdo da Aba "Spooler de Impressão"
# -----------------------------------------------------------------------------------
$startX_Spooler = 20
$startY_Spooler = 20
$spacingY_Spooler = 40

$labelComputerName = New-Object system.windows.forms.Label
$labelComputerName.Text = "Servidor:"
$labelComputerName.AutoSize = $true
$labelComputerName.Location = New-Object System.Drawing.Point($startX_Spooler, $startY_Spooler + 5)
$labelComputerName.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageSpooler.Controls.Add($labelComputerName)

$textBoxComputerName = New-Object system.windows.forms.TextBox
$textBoxComputerName.Size = New-Object System.Drawing.Size(350, 30)
$textBoxComputerName.Location = New-Object System.Drawing.Point($startX_Spooler + 80, $startY_Spooler)
$textBoxComputerName.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$tabPageSpooler.Controls.Add($textBoxComputerName)

$buttonLoadServer = New-Object system.windows.forms.Button
$buttonLoadServer.Size = New-Object System.Drawing.Size(150, 40)
$buttonLoadServer.Text = "Carregar Servidor"
$buttonLoadServer.Location = New-Object System.Drawing.Point($textBoxComputerName.Left + $textBoxComputerName.Width + 10, $startY_Spooler - 5)
$buttonLoadServer.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
$buttonLoadServer.ForeColor = [System.Drawing.Color]::White
$buttonLoadServer.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonLoadServer.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabPageSpooler.Controls.Add($buttonLoadServer)

$startY_Spooler += $spacingY_Spooler + 10

$panelButtons = New-Object System.Windows.Forms.Panel
$panelButtons.Location = New-Object System.Drawing.Point($startX_Spooler, $startY_Spooler)
$panelButtons.Size = New-Object System.Drawing.Size(480, 50)
$tabPageSpooler.Controls.Add($panelButtons)

$buttonRestartSpooler = New-Object system.windows.forms.Button
$buttonRestartSpooler.Size = New-Object System.Drawing.Size(150, 40)
$buttonRestartSpooler.Text = "Reiniciar Spooler"
$buttonRestartSpooler.Location = New-Object System.Drawing.Point(0, 0)
$buttonRestartSpooler.BackColor = [System.Drawing.Color]::FromArgb(255, 165, 0)
$buttonRestartSpooler.ForeColor = [System.Drawing.Color]::White
$buttonRestartSpooler.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonRestartSpooler.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$panelButtons.Controls.Add($buttonRestartSpooler)

$buttonListQueue = New-Object system.windows.forms.Button
$buttonListQueue.Size = New-Object System.Drawing.Size(150, 40)
$buttonListQueue.Text = "Atualizar Fila"
$buttonListQueue.Location = New-Object System.Drawing.Point(170, 0)
$buttonListQueue.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
$buttonListQueue.ForeColor = [System.Drawing.Color]::White
$buttonListQueue.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$buttonListQueue.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$panelButtons.Controls.Add($buttonListQueue)

$startY_Spooler += 60

$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Location = New-Object System.Drawing.Point($startX_Spooler, $startY_Spooler)
$splitContainer.Size = New-Object System.Drawing.Size($form.ClientSize.Width - ($startX_Spooler * 2), $form.ClientSize.Height - $startY_Spooler - 50)
$splitContainer.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$splitContainer.SplitterDistance = 220
$splitContainer.Anchor = 'Top, Bottom, Left, Right'
$tabPageSpooler.Controls.Add($splitContainer)

$dataGridViewPrinters = New-Object System.Windows.Forms.DataGridView
$dataGridViewPrinters.Dock = "Fill"
$dataGridViewPrinters.AutoSizeColumnsMode = 'Fill'
$dataGridViewPrinters.ReadOnly = $true
$dataGridViewPrinters.AllowUserToAddRows = $false
$splitContainer.Panel1.Controls.Add($dataGridViewPrinters)

$dataGridViewQueue = New-Object System.Windows.Forms.DataGridView
$dataGridViewQueue.Dock = "Fill"
$dataGridViewQueue.AutoSizeColumnsMode = 'Fill'
$dataGridViewQueue.ReadOnly = $true
$dataGridViewQueue.AllowUserToAddRows = $false
$splitContainer.Panel2.Controls.Add($dataGridViewQueue)

$labelEmptyQueue = New-Object System.Windows.Forms.Label
$labelEmptyQueue.Text = "A fila de impressão está vazia ou o servidor não foi carregado."
$labelEmptyQueue.Dock = "Fill"
$labelEmptyQueue.TextAlign = "MiddleCenter"
$labelEmptyQueue.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Italic)
$labelEmptyQueue.Visible = $true
$splitContainer.Panel2.Controls.Add($labelEmptyQueue)
$labelEmptyQueue.BringToFront()

# -----------------------------------------------------------------------------------
# Lógica de Ações
# -----------------------------------------------------------------------------------
$logPath = ".\actions.log"

function Log-Action {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $env:USERNAME -> $message" | Out-File -Append -FilePath $logPath
}

function New-RandomPassword {
    param([int]$length = 12)
    $charSet = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789!@#$%^&*'
    -join ($charSet.ToCharArray() | Get-Random -Count $length)
}

$buttonExecute_Gerenciar.Add_Click({
    $usuario = $textboxUser_Gerenciar.Text.Trim()
    $action = $comboBoxAction_Gerenciar.SelectedItem
    $selectedADServer = $comboBoxAD_Gerenciar.SelectedItem

    if (-not $usuario) { [System.Windows.Forms.MessageBox]::Show("Por favor, insira o login do usuário.", "Erro", "OK", "Error"); return }
    if (-not $selectedADServer) { [System.Windows.Forms.MessageBox]::Show("Por favor, selecione um servidor AD.", "Erro", "OK", "Error"); return }
    if (-not $action) { [System.Windows.Forms.MessageBox]::Show("Por favor, selecione uma ação.", "Erro", "OK", "Error"); return }

    try {
        switch ($action) {
            "Resetar Senha" {
                $newPassword = New-RandomPassword
                $securePassword = ConvertTo-SecureString $newPassword -AsPlainText -Force
                
                Unlock-ADAccount -Identity $usuario -Server $selectedADServer
                Set-ADAccountPassword -Identity $usuario -NewPassword $securePassword -Reset -Server $selectedADServer
                Set-ADUser -Identity $usuario -ChangePasswordAtLogon $true -Server $selectedADServer

                $user = Get-ADUser $usuario -Properties Name -Server $selectedADServer
                
                # --- INÍCIO DA ALTERAÇÃO ---
                
                # 1. Definir um array com os formatos de mensagem. {0} é o nome, {1} é a senha.
                $messageFormats = @(
                    "A senha do colaborador(a) {0} foi redefinida. Nova senha temporária: {1}",
                    "Pronto! A nova senha para o acesso de {0} é: {1}",
                    "Senha resetada para o usuário {0}. A nova senha provisória é: {1}",
                    "Acesso de {0} atualizado. A senha temporária para o próximo login é: {1}",
                    "Sucesso! A nova senha de {0} é: {1}. O usuário deverá alterá-la no primeiro acesso."
                )

                # 2. Escolher um formato aleatoriamente
                $randomFormat = Get-Random -InputObject $messageFormats
                
                # 3. Criar a mensagem final usando o formato escolhido
                $clipboardMessage = $randomFormat -f $user.Name, $newPassword

                # 4. Criar a mensagem para a caixa de diálogo
                $dialogMessage = "Senha redefinida com sucesso!`n`n$clipboardMessage`n`n(A mensagem acima foi copiada para a área de transferência)"
                
                # 5. Exibir a caixa de diálogo e copiar para a área de transferência
                [System.Windows.Forms.MessageBox]::Show($dialogMessage, "Senha Redefinida", "OK", "Information")
                Set-Clipboard -Value $clipboardMessage

                # --- FIM DA ALTERAÇÃO ---
                
                Log-Action "Senha redefinida para $usuario no servidor $selectedADServer."
            }
            "Desbloquear Usuário" {
                Unlock-ADAccount -Identity $usuario -Server $selectedADServer
                $user = Get-ADUser $usuario -Properties Name -Server $selectedADServer
                $message = "A conta do colaborador(a) $($user.Name) foi desbloqueada."
                [System.Windows.Forms.MessageBox]::Show($message, "Usuário Desbloqueado", "OK", "Information")
                Log-Action "Usuário $usuario desbloqueado no servidor $selectedADServer."
            }
            "Desativar Usuário" {
                Disable-ADAccount -Identity $usuario -Server $selectedADServer
                $user = Get-ADUser $usuario -Properties Name -Server $selectedADServer
                $message = "A conta do usuário $($user.Name) foi desativada com sucesso."
                [System.Windows.Forms.MessageBox]::Show($message, "Usuário Desativado", "OK", "Information")
                Log-Action "Usuário $usuario desativado no servidor $selectedADServer."
            }
            "Apagar Usuário" {
                #$user = Get-ADUser -Identity $usuario -Properties whenChanged, Enabled, DisplayName, SamAccountName -Server $selectedADServer
                $user = Get-ADUser -Filter "displayname -like '*${usuario}*'" -Properties * -Server $selectedADServer
                if ($user.Enabled) {
                    [System.Windows.Forms.MessageBox]::Show("O usuário precisa estar desativado para ser apagado.", "Erro", "OK", "Error")
                    return
                }
                
                Set-Clipboard -Value $user.SamAccountName
                $confirmMsg = "Tem certeza que deseja APAGAR a conta $($user.DisplayName)?`n`nLogin: $($user.SamAccountName)`nÚltima modificação: $($user.whenChanged)`n`nEsta ação é IRREVERSÍVEL. Verifique as políticas internas antes de prosseguir."
                $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirmar Exclusão Permanente", "YesNo", "Warning")

                if ($confirm -eq 'Yes') {
                    Remove-ADUser -Identity $user.SamAccountName -Confirm:$false -Server $selectedADServer
                    $message = "A conta de usuário $($user.SamAccountName) foi excluída permanentemente."
                    [System.Windows.Forms.MessageBox]::Show($message, "Usuário Apagado", "OK", "Information")
                    Log-Action "Usuário $($user.SamAccountName) apagado no servidor $selectedADServer."
                }
            }
        }
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
         [System.Windows.Forms.MessageBox]::Show("Erro: O usuário '$usuario' não foi encontrado no servidor '$selectedADServer'.", "Usuário Não Encontrado", "OK", "Error")
         Log-Action "Erro: Falha ao executar '$action' para '$usuario' em '$selectedADServer'. Usuário não encontrado."
    } 
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro: $($_.Exception.Message)", "Erro Inesperado", "OK", "Error")
        Log-Action "Erro ao processar a função ($action) no servidor $$selectedADServer: $($_.Exception.Message)"
    }
})

$buttonExecute_Remover.Add_Click({
    $hostname = $textboxHost_Remover.Text.Trim()
    $selectedADServer = $comboBoxAD_Remover.SelectedItem

    if (-not $hostname) { [System.Windows.Forms.MessageBox]::Show("Por favor, insira o nome do host.", "Erro", "OK", "Error"); return }
    if (-not $selectedADServer) { [System.Windows.Forms.MessageBox]::Show("Por favor, selecione um servidor AD.", "Erro", "OK", "Error"); return }
    try {
        $computerObject = Get-ADComputer -Identity $hostname -Server $selectedADServer -Properties CanonicalName
        $confirmMessage = "Tem certeza que deseja remover o seguinte host do AD?`n`nNome: $hostname`nLocalização: $($computerObject.CanonicalName)`nServidor AD: $selectedADServer`n`nEsta ação não pode ser desfeita."
        $confirmResult = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirmar Remoção de Host", "YesNo", "Warning")

        if ($confirmResult -eq 'Yes') {
            Remove-ADComputer -Identity $hostname -Server $selectedADServer -Confirm:$false
            $message = "O host '$hostname' foi removido com sucesso do Active Directory."
            [System.Windows.Forms.MessageBox]::Show($message, "Host Removido", "OK", "Information")
            Log-Action "Host '$hostname' (Local: $($computerObject.CanonicalName)) removido do AD no servidor $selectedADServer."
        } else {
            [System.Windows.Forms.MessageBox]::Show("A remoção do host '$hostname' foi cancelada.", "Remoção Cancelada", "OK", "Information")
            Log-Action "Remoção do host '$hostname' cancelada pelo usuário."
        }
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        [System.Windows.Forms.MessageBox]::Show("Erro: O host '$hostname' não foi encontrado no Active Directory no servidor '$selectedADServer'.", "Erro - Host Não Encontrado", "OK", "Error")
        Log-Action "Falha ao remover host '$hostname' no servidor $$selectedADServer: Host não encontrado."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro inesperado: $($_.Exception.Message)", "Erro Inesperado", "OK", "Error")
        Log-Action "Erro ao remover host '$hostname' no servidor $$selectedADServer: $($_.Exception.Message)"
    }
})

function Get-PrintersCim {
    param($computerName)
    $printers = Get-CimInstance -ClassName Win32_Printer -ComputerName $computerName
    return $printers | Select-Object @{N='Nome'; E={$_.Name}}, @{N='Compartilhada'; E={$_.Shared}}, @{N='Padrão'; E={$_.Default}}, @{N='Status'; E={$_.PrinterStatus}}
}

function Get-PrintQueueCim {
    param($computerName)
    $printJobs = Get-CimInstance -ClassName Win32_PrintJob -ComputerName $computerName
    return $printJobs | Select-Object @{N='Nome do Documento'; E={$_.Document}}, @{N='Usuário'; E={$_.Owner}}, @{N='Status'; E={$_.JobStatus}}
}

$buttonLoadServer.Add_Click({
    $computerName = $textBoxComputerName.Text.Trim()
    if (-not $computerName) { [System.Windows.Forms.MessageBox]::Show("Por favor, insira o nome de um servidor.", "Aviso", "OK", "Warning"); return }
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $dataGridViewPrinters.DataSource = Get-PrintersCim -computerName $computerName
        $queueData = Get-PrintQueueCim -computerName $computerName
        $dataGridViewQueue.DataSource = $queueData
        $labelEmptyQueue.Visible = ($queueData.Count -eq 0)
        Log-Action "Dados de impressão carregados para o servidor '$computerName'."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Não foi possível conectar ao servidor '$computerName'. Verifique o nome e a conectividade de rede.`n`nErro: $($_.Exception.Message)", "Erro de Conexão", "OK", "Error")
        Log-Action "Erro ao carregar dados do servidor '$computerName': $($_.Exception.Message)"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$buttonRestartSpooler.Add_Click({
    $computerName = $textBoxComputerName.Text.Trim()
    if (-not $computerName) { [System.Windows.Forms.MessageBox]::Show("Insira o nome de um servidor.", "Aviso", "OK", "Warning"); return }
    try {
        Get-Service Spooler -ComputerName $computerName | Restart-Service -Force
        [System.Windows.Forms.MessageBox]::Show("Serviço de Spooler reiniciado com sucesso em '$computerName'.", "Spooler Reiniciado", "OK", "Information")
        Log-Action "Serviço de Spooler reiniciado em '$computerName'."
        $dataGridViewPrinters.DataSource = Get-PrintersCim -computerName $computerName
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao reiniciar o serviço de Spooler em '$computerName': $($_.Exception.Message)", "Erro", "OK", "Error")
        Log-Action "Erro ao reiniciar o serviço de Spooler em '$computerName': $($_.Exception.Message)"
    }
})

$buttonListQueue.Add_Click({
    $computerName = $textBoxComputerName.Text.Trim()
    if (-not $computerName) { [System.Windows.Forms.MessageBox]::Show("Insira o nome de um servidor e clique em 'Carregar Servidor'.", "Aviso", "OK", "Warning"); return }
    try {
        $queueData = Get-PrintQueueCim -computerName $computerName
        $dataGridViewQueue.DataSource = $queueData
        $labelEmptyQueue.Visible = ($queueData.Count -eq 0)
        Log-Action "Fila de impressão atualizada para '$computerName'."
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao listar a fila de impressão para '$computerName': $($_.Exception.Message)", "Erro", "OK", "Error")
        Log-Action "Erro ao listar a fila de impressão para '$computerName': $($_.Exception.Message)"
    }
})

# Executar o formulário
$form.ShowDialog()

# Limpar variáveis do formulário da memória após fechar
$form.Dispose()
