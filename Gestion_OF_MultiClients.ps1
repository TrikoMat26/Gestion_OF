<#
    Gestion_OF_MultiClients.ps1
    Outil autonome pour la gestion des Ordres de Fabrication (OF) et Numéros de Série (SN)
    Support multi-clients.
    (Version avec Interface Moderne)
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$GlobalRegistryPath = Join-Path $ScriptPath "Global_OF_Registry.json"

# ==========================================
# FONCTIONS DE BASE & DONNEES
# ==========================================

Function Get-GlobalRegistry {
    if (Test-Path -LiteralPath $GlobalRegistryPath) {
        try {
            $json = Get-Content -LiteralPath $GlobalRegistryPath -Raw -Encoding UTF8
            if (-not $json -or $json.Trim().Length -eq 0) { return [ordered]@{} }
            $obj = $json | ConvertFrom-Json
            $reg = [ordered]@{}
            foreach ($clientProp in $obj.PSObject.Properties) {
                # $clientProp.Name est le nom du client (ex: Wattsy)
                $clientReg = [ordered]@{}
                if ($clientProp.Value) {
                    foreach ($ofProp in $clientProp.Value.PSObject.Properties) {
                        $clientReg[$ofProp.Name] = @($ofProp.Value)
                    }
                }
                $reg[$clientProp.Name] = $clientReg
            }
            return $reg
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Erreur lecture Global_OF_Registry.json :`n$($_.Exception.Message)", "Erreur", "OK", "Error")
            return [ordered]@{}
        }
    }
    return [ordered]@{}
}

Function Save-GlobalRegistry($registry) {
    try {
        $sorted = [ordered]@{}
        foreach ($client in ($registry.Keys | Sort-Object)) {
            $sortedClient = [ordered]@{}
            foreach ($of in ($registry[$client].Keys | Sort-Object)) {
                $sortedClient[$of] = @($registry[$client][$of] | Sort-Object)
            }
            $sorted[$client] = $sortedClient
        }
        $json = $sorted | ConvertTo-Json -Depth 5
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($GlobalRegistryPath, $json, $utf8NoBom)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur sauvegarde Global_OF_Registry.json :`n$($_.Exception.Message)", "Erreur", "OK", "Error")
    }
}

Function Find-OFBySN_Global($registry, $sn) {
    # Retourne un objet @{ Client = "..."; OF = "..." }
    foreach ($client in $registry.Keys) {
        foreach ($of in $registry[$client].Keys) {
            if ($registry[$client][$of] -contains $sn) { 
                return @{ Client=$client; OF=$of }
            }
        }
    }
    return $null
}

Function Expand-SNRange($inputText) {
    if ([string]::IsNullOrWhiteSpace($inputText)) { return $null }
    $inputText = $inputText.Trim()

    if ($inputText.Contains(',')) {
        $parts = $inputText -split ','
        $results = @()
        foreach ($p in $parts) {
            $clean = $p.Trim()
            if ($clean -match '^\d+$') { $results += $clean }
            else { return $null }
        }
        return $results
    }
    
    if ($inputText -match '^\s*(\d+)\s*-\s*(\d+)\s*$') {
        $startNum = [int]$matches[1]
        $endNum = [int]$matches[2]
        if ($startNum -gt $endNum) { return $null }
        $width = $matches[1].Length
        $formatStr = "{0:D$width}"
        $result = @()
        for ($i = $startNum; $i -le $endNum; $i++) {
            $result += $formatStr -f $i
        }
        return $result
    }
    elseif ($inputText -match '^\d+$') {
        return @($inputText)
    }
    return $null
}

# ==========================================
# INTERFACE GRAPHIQUE (WINFORMS) MODERNE
# ==========================================

$global:currentRegistry = Get-GlobalRegistry

# --- PALETTE DE COULEURS ---
$ColorBg = [System.Drawing.ColorTranslator]::FromHtml("#F3F4F6") # Gris extrèmement clair
$ColorPanel = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF") # Blanc
$ColorPrimary = [System.Drawing.ColorTranslator]::FromHtml("#2563EB") # Bleu moderne (Actions positives)
$ColorWhite = [System.Drawing.ColorTranslator]::FromHtml("#FFFFFF")
$ColorDanger = [System.Drawing.ColorTranslator]::FromHtml("#DC2626") # Rouge (Suppressions)
$ColorText = [System.Drawing.ColorTranslator]::FromHtml("#1F2937") # Gris très foncé
$ColorSecondary = [System.Drawing.ColorTranslator]::FromHtml("#9CA3AF") # Gris bordures
$ColorSuccess = [System.Drawing.ColorTranslator]::FromHtml("#16A34A") # Vert (Succès)

# --- FENETRE PRINCIPALE ---
$Form = New-Object Windows.Forms.Form
$Form.Text = "Gestionnaire Central Multi-Clients des OF"
$Form.Size = New-Object Drawing.Size(1100, 700)
$Form.MinimumSize = New-Object Drawing.Size(900, 550)
$Form.StartPosition = "CenterScreen"
$Form.BackColor = $ColorBg
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)

# --- STATUS STRIP ---
$statusStrip = New-Object Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#E5E7EB")
$statusLabel = New-Object Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Prêt."
$statusLabel.ForeColor = $ColorPrimary
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$statusStrip.Items.Add($statusLabel) | Out-Null
$Form.Controls.Add($statusStrip)

# --- LAYOUT PRINCIPAL (3 COLONNES) ---
$mainTable = New-Object Windows.Forms.TableLayoutPanel
$mainTable.Dock = "Fill"
$mainTable.ColumnCount = 3
$mainTable.RowCount = 1
$mainTable.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 31)))
$mainTable.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 31)))
$mainTable.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 38)))
$mainTable.Padding = New-Object Windows.Forms.Padding(10, 10, 10, 35) # Espace en bas pour éviter le StatusBar
$Form.Controls.Add($mainTable)

Function Create-FlatButton ($text, $colorBg, $colorFg) {
    $btn = New-Object Windows.Forms.Button
    $btn.Text = $text
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = $colorBg
    $btn.ForeColor = $colorFg
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

# ------------------------------------------
# --- COLONNE 1 : CLIENTS                ---
# ------------------------------------------
$grpClient = New-Object Windows.Forms.GroupBox
$grpClient.Text = " 1. Clients "
$grpClient.Dock = "Fill"
$grpClient.BackColor = $ColorPanel
$grpClient.ForeColor = $ColorPrimary
$grpClient.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$grpClient.Margin = New-Object Windows.Forms.Padding(5)
$mainTable.Controls.Add($grpClient, 0, 0)

$g1Table = New-Object Windows.Forms.TableLayoutPanel
$g1Table.Dock = "Fill" ; $g1Table.RowCount = 5 ; $g1Table.ColumnCount = 1; $g1Table.Padding = New-Object Windows.Forms.Padding(10)
$g1Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100))) # ListBox
$g1Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # TXT
$g1Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # ADD
$g1Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # DEL
$g1Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # IMPORT
$grpClient.Controls.Add($g1Table)

$lstClient = New-Object Windows.Forms.ListBox; $lstClient.Dock = "Fill" ; $lstClient.Font = New-Object System.Drawing.Font("Segoe UI", 10.5); $lstClient.ItemHeight = 22; $lstClient.BorderStyle = "FixedSingle"; $lstClient.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 15); $lstClient.ForeColor = $ColorText
$txtNewClient = New-Object Windows.Forms.TextBox; $txtNewClient.Dock = "Fill" ; $txtNewClient.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtNewClient.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 5); $txtNewClient.BorderStyle = "FixedSingle"
$btnAddClient = Create-FlatButton "Ajouter Client" $ColorPrimary $ColorWhite ; $btnAddClient.Dock = "Fill" ; $btnAddClient.Height = 38 ; $btnAddClient.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 5)
$btnDelClient = Create-FlatButton "Supprimer Client" $ColorBg $ColorDanger ; $btnDelClient.FlatAppearance.BorderSize = 1 ; $btnDelClient.FlatAppearance.BorderColor = $ColorDanger ; $btnDelClient.Dock = "Fill" ; $btnDelClient.Height = 38 ; $btnDelClient.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 20)
$btnImport = Create-FlatButton "Importer Wattsy" $ColorBg $ColorText ; $btnImport.FlatAppearance.BorderSize = 1 ; $btnImport.FlatAppearance.BorderColor = $ColorSecondary ; $btnImport.Dock = "Fill" ; $btnImport.Height = 35 ; $btnImport.Margin = New-Object Windows.Forms.Padding(0)

$g1Table.Controls.Add($lstClient, 0, 0)
$g1Table.Controls.Add($txtNewClient, 0, 1)
$g1Table.Controls.Add($btnAddClient, 0, 2)
$g1Table.Controls.Add($btnDelClient, 0, 3)
$g1Table.Controls.Add($btnImport, 0, 4)


# ------------------------------------------
# --- COLONNE 2 : ORDRES FABRICATION     ---
# ------------------------------------------
$grpOF = New-Object Windows.Forms.GroupBox
$grpOF.Text = " 2. Ordres de Fabrication "
$grpOF.Dock = "Fill" 
$grpOF.BackColor = $ColorPanel 
$grpOF.ForeColor = $ColorPrimary 
$grpOF.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$grpOF.Margin = New-Object Windows.Forms.Padding(5)
$mainTable.Controls.Add($grpOF, 1, 0)

$g2Table = New-Object Windows.Forms.TableLayoutPanel
$g2Table.Dock = "Fill" ; $g2Table.RowCount = 5 ; $g2Table.ColumnCount = 1; $g2Table.Padding = New-Object Windows.Forms.Padding(10)
$g2Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # Search
$g2Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100))) # ListBox
$g2Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # TXT
$g2Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # ADD
$g2Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # DEL
$grpOF.Controls.Add($g2Table)

# Recherche intégrée
$searchFlow = New-Object Windows.Forms.TableLayoutPanel; $searchFlow.Dock = "Fill"; $searchFlow.Height = 35 ; $searchFlow.ColumnCount = 2 ; $searchFlow.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 15)
$searchFlow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 60)))
$searchFlow.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 40)))
$searchFlow.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100)))
$txtSearchSN = New-Object Windows.Forms.TextBox; $txtSearchSN.Dock = "Fill" ; $txtSearchSN.Font = New-Object System.Drawing.Font("Segoe UI", 10.5); $txtSearchSN.BorderStyle = "FixedSingle"; $txtSearchSN.Margin = New-Object Windows.Forms.Padding(0, 2, 5, 0)
$btnSearch = Create-FlatButton "Recherche" $ColorBg $ColorText; $btnSearch.FlatAppearance.BorderSize = 1 ; $btnSearch.FlatAppearance.BorderColor = $ColorSecondary ; $btnSearch.Dock = "Fill" ; $btnSearch.Height = 30 ; $btnSearch.Margin = New-Object Windows.Forms.Padding(0)
$searchFlow.Controls.Add($txtSearchSN, 0, 0); $searchFlow.Controls.Add($btnSearch, 1, 0)

$lstOF = New-Object Windows.Forms.ListBox; $lstOF.Dock = "Fill" ; $lstOF.Font = New-Object System.Drawing.Font("Segoe UI", 10.5); $lstOF.ItemHeight = 22; $lstOF.BorderStyle = "FixedSingle"; $lstOF.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 15); $lstOF.ForeColor = $ColorText
$txtNewOF = New-Object Windows.Forms.TextBox; $txtNewOF.Dock = "Fill" ; $txtNewOF.MaxLength = 10; $txtNewOF.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtNewOF.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 5); $txtNewOF.BorderStyle = "FixedSingle"
$btnAddOF = Create-FlatButton "Ajouter OF" $ColorPrimary $ColorWhite ; $btnAddOF.Dock = "Fill" ; $btnAddOF.Height = 38 ; $btnAddOF.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 5)
$btnDelOF = Create-FlatButton "Supprimer OF" $ColorBg $ColorDanger ; $btnDelOF.FlatAppearance.BorderSize = 1 ; $btnDelOF.FlatAppearance.BorderColor = $ColorDanger ; $btnDelOF.Dock = "Fill" ; $btnDelOF.Height = 38 ; $btnDelOF.Margin = New-Object Windows.Forms.Padding(0)

$g2Table.Controls.Add($searchFlow, 0, 0)
$g2Table.Controls.Add($lstOF, 0, 1)
$g2Table.Controls.Add($txtNewOF, 0, 2)
$g2Table.Controls.Add($btnAddOF, 0, 3)
$g2Table.Controls.Add($btnDelOF, 0, 4)


# ------------------------------------------
# --- COLONNE 3 : NUMEROS SERIE (SN)     ---
# ------------------------------------------
$grpSN = New-Object Windows.Forms.GroupBox
$grpSN.Text = " 3. Numéros de Série (SN) "
$grpSN.Dock = "Fill" 
$grpSN.BackColor = $ColorPanel 
$grpSN.ForeColor = $ColorPrimary 
$grpSN.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$grpSN.Margin = New-Object Windows.Forms.Padding(5)
$mainTable.Controls.Add($grpSN, 2, 0)

$g3Table = New-Object Windows.Forms.TableLayoutPanel
$g3Table.Dock = "Fill" ; $g3Table.RowCount = 4 ; $g3Table.ColumnCount = 1; $g3Table.Padding = New-Object Windows.Forms.Padding(10)
$g3Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100))) # ListBox
$g3Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # LBL
$g3Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # TXT
$g3Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize))) # FLOW BTN
$grpSN.Controls.Add($g3Table)

$lstSN = New-Object Windows.Forms.ListBox; $lstSN.Dock = "Fill" ; $lstSN.Font = New-Object System.Drawing.Font("Consolas", 11.5); $lstSN.ItemHeight = 22; $lstSN.BorderStyle = "FixedSingle"; $lstSN.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 15); $lstSN.ForeColor = $ColorText
$lblAddSN = New-Object Windows.Forms.Label; $lblAddSN.Text = "Saisie SN (ex: 043590 ou 043590-043600) :"; $lblAddSN.AutoSize = $true; $lblAddSN.Font = New-Object System.Drawing.Font("Segoe UI", 9.5) ; $lblAddSN.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 4); $lblAddSN.ForeColor = $ColorText
$txtNewSN = New-Object Windows.Forms.TextBox; $txtNewSN.Dock = "Fill" ; $txtNewSN.Font = New-Object System.Drawing.Font("Segoe UI", 11); $txtNewSN.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 10); $txtNewSN.BorderStyle = "FixedSingle"

$btnFlowSN = New-Object Windows.Forms.TableLayoutPanel; $btnFlowSN.Dock = "Fill" ; $btnFlowSN.Height = 38 ; $btnFlowSN.ColumnCount = 2 ; $btnFlowSN.Margin = New-Object Windows.Forms.Padding(0)
$btnFlowSN.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 50)))
$btnFlowSN.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 50)))
$btnFlowSN.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100)))

$btnAddSN = Create-FlatButton "Ajouter SN" $ColorPrimary $ColorWhite ; $btnAddSN.Dock = "Fill" ; $btnAddSN.Height = 38 ; $btnAddSN.Margin = New-Object Windows.Forms.Padding(0, 0, 5, 0)
$btnDelSN = Create-FlatButton "Supprimer SN" $ColorBg $ColorDanger ; $btnDelSN.FlatAppearance.BorderSize = 1 ; $btnDelSN.FlatAppearance.BorderColor = $ColorDanger ; $btnDelSN.Dock = "Fill" ; $btnDelSN.Height = 38 ; $btnDelSN.Margin = New-Object Windows.Forms.Padding(5, 0, 0, 0)

$btnFlowSN.Controls.Add($btnAddSN, 0, 0); $btnFlowSN.Controls.Add($btnDelSN, 1, 0)

$g3Table.Controls.Add($lstSN, 0, 0)
$g3Table.Controls.Add($lblAddSN, 0, 1)
$g3Table.Controls.Add($txtNewSN, 0, 2)
$g3Table.Controls.Add($btnFlowSN, 0, 3)

# ==========================================
# LOGIQUE EVENEMENTIELLE
# ==========================================

# --- HELPERS ---
$GetSelClient = { if ($lstClient.SelectedItem) { return ($lstClient.SelectedItem -split ' \[')[0] } return $null }
$GetSelOF     = { if ($lstOF.SelectedItem) { return ($lstOF.SelectedItem -split ' \(')[0] } return $null }

$RefreshSNList = {
    $selClient = & $GetSelClient
    $selOF = & $GetSelOF
    $lstSN.Items.Clear()
    if ($selClient -and $selOF) {
        $count = 0
        if ($global:currentRegistry[$selClient].Contains($selOF)) {
            foreach ($s in ($global:currentRegistry[$selClient][$selOF] | Sort-Object)) { [void]$lstSN.Items.Add($s); $count++ }
        }
        $statusLabel.Text = "$count SN dans l'OF $selOF (Client: $selClient)"
        $statusLabel.ForeColor = $ColorPrimary
        $txtNewOF.Text = ""
        $grpSN.Enabled = $true
    }
    else {
        $statusLabel.Text = "Sélectionnez un OF"
        $statusLabel.ForeColor = $ColorText
        $grpSN.Enabled = $false
    }
}

$RefreshOFList = {
    $selClient = & $GetSelClient
    $selOF = & $GetSelOF
    $lstOF.Items.Clear()
    if ($selClient) {
        foreach ($k in $global:currentRegistry[$selClient].Keys) {
            $cnt = 0
            if ($global:currentRegistry[$selClient][$k]) { $cnt = $global:currentRegistry[$selClient][$k].Count }
            [void]$lstOF.Items.Add("$k ($cnt)")
        }
        if ($selOF) {
            $idx = $lstOF.FindString($selOF)
            if ($idx -ge 0) { $lstOF.SelectedIndex = $idx }
        }
        $grpOF.Enabled = $true
        $txtNewClient.Text = ""
    }
    else {
        $grpOF.Enabled = $false
    }
    & $RefreshSNList
}

$RefreshClientList = {
    $selClient = & $GetSelClient
    $lstClient.Items.Clear()
    foreach ($k in $global:currentRegistry.Keys) {
        $cnt = $global:currentRegistry[$k].Keys.Count
        [void]$lstClient.Items.Add("$k [$cnt OF]")
    }
    if ($selClient) {
        $idx = $lstClient.FindString($selClient)
        if ($idx -ge 0) { $lstClient.SelectedIndex = $idx }
    }
    & $RefreshOFList
}

# --- EVENTS LISTBOX & ACTIONS RAPIDES ---

$lstClient.Add_SelectedIndexChanged({ & $RefreshOFList })
$lstOF.Add_SelectedIndexChanged({ & $RefreshSNList })

$lstClient.Add_DoubleClick({ 
    $sel = & $GetSelClient
    if ($sel) { $txtNewClient.Text = $sel ; $txtNewClient.Focus() ; $txtNewClient.Select($txtNewClient.Text.Length, 0) }
})

$lstOF.Add_DoubleClick({ 
    $sel = & $GetSelOF
    if ($sel) { $txtNewOF.Text = $sel ; $txtNewOF.Focus() ; $txtNewOF.Select($txtNewOF.Text.Length, 0) }
})

$lstSN.Add_DoubleClick({ 
    if ($lstSN.SelectedItem) { 
        $txtNewSN.Text = $lstSN.SelectedItem ; $txtNewSN.Focus() ; $txtNewSN.Select($txtNewSN.Text.Length, 0) 
    }
})

$txtNewClient.Add_KeyDown({ param($s, $e); if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $btnAddClient.PerformClick() } })
$txtNewOF.Add_KeyDown({ param($s, $e);     if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $btnAddOF.PerformClick() } })
$txtNewSN.Add_KeyDown({ param($s, $e);     if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $btnAddSN.PerformClick() } })
$txtSearchSN.Add_KeyDown({ param($s, $e);  if ($e.KeyCode -eq 'Enter') { $e.SuppressKeyPress = $true; $btnSearch.PerformClick() } })

# --- BOUTONS D'ACTION (CLIENTS) ---

$btnAddClient.Add_Click({
    $newClient = $txtNewClient.Text.Trim()
    if ($newClient) {
        if (-not $global:currentRegistry.Contains($newClient)) {
            $global:currentRegistry[$newClient] = [ordered]@{}
            Save-GlobalRegistry $global:currentRegistry
            & $RefreshClientList
            $idx = $lstClient.FindString($newClient)
            if ($idx -ge 0) { $lstClient.SelectedIndex = $idx }
        } else { [System.Windows.Forms.MessageBox]::Show("Le client '$newClient' existe déjà.", "Info") }
    }
})

$btnDelClient.Add_Click({
    $selClient = & $GetSelClient
    if ($selClient) {
        $res = [System.Windows.Forms.MessageBox]::Show("Supprimer le client $selClient ainsi que TOUS ses OF ?", "Attention", "YesNo", "Warning")
        if ($res -eq "Yes") {
            $global:currentRegistry.Remove($selClient)
            Save-GlobalRegistry $global:currentRegistry
            & $RefreshClientList
            $statusLabel.Text = "Client $selClient supprimé."
            $statusLabel.ForeColor = $ColorDanger
        }
    }
})

$btnImport.Add_Click({
    $oldRegPath = Join-Path $ScriptPath "OF_Registry.json"
    if (Test-Path $oldRegPath) {
        $json = Get-Content $oldRegPath -Raw -Encoding UTF8
        if ($json) {
            $obj = $json | ConvertFrom-Json
            if (-not $global:currentRegistry.Contains("Wattsy")) {
                $global:currentRegistry["Wattsy"] = [ordered]@{}
            }
            $count = 0
            foreach ($prop in $obj.PSObject.Properties) {
                $ofName = $prop.Name
                if (-not $global:currentRegistry["Wattsy"].Contains($ofName)) {
                    $global:currentRegistry["Wattsy"][$ofName] = @($prop.Value)
                    $count++
                } else {
                    $global:currentRegistry["Wattsy"][$ofName] = @( ($global:currentRegistry["Wattsy"][$ofName] + $prop.Value) | Select-Object -Unique )
                }
            }
            Save-GlobalRegistry $global:currentRegistry
            & $RefreshClientList
            [System.Windows.Forms.MessageBox]::Show("Importation réussie. $count OF ajoutés ou mis à jour sous le client 'Wattsy'.", "Succès")
            $statusLabel.Text = "$count OF importés."
            $statusLabel.ForeColor = $ColorSuccess
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("L'ancien fichier OF_Registry.json n'a pas été trouvé dans ce dossier.", "Erreur")
    }
})

# --- BOUTONS D'ACTION (OF) ---

$btnAddOF.Add_Click({
    $selClient = & $GetSelClient
    if ($selClient) {
        $newOF = $txtNewOF.Text.Trim()
        if ($newOF) {
            if (-not $global:currentRegistry[$selClient].Contains($newOF)) {
                $global:currentRegistry[$selClient][$newOF] = @()
                Save-GlobalRegistry $global:currentRegistry
                & $RefreshClientList 
                $idx = $lstOF.FindString($newOF)
                if ($idx -ge 0) { $lstOF.SelectedIndex = $idx }
            } else { [System.Windows.Forms.MessageBox]::Show("L'OF '$newOF' existe déjà pour ce client.", "Info") }
        }
    }
})

$btnDelOF.Add_Click({
    $selClient = & $GetSelClient
    $selOF = & $GetSelOF
    if ($selClient -and $selOF) {
        $res = [System.Windows.Forms.MessageBox]::Show("Supprimer l'OF $selOF du client $selClient ?", "Confirmation", "YesNo", "Warning")
        if ($res -eq "Yes") {
            $global:currentRegistry[$selClient].Remove($selOF)
            Save-GlobalRegistry $global:currentRegistry
            & $RefreshClientList
            $statusLabel.Text = "OF $selOF supprimé."
            $statusLabel.ForeColor = $ColorDanger
        }
    }
})

# --- BOUTONS D'ACTION (SN) ---

$btnAddSN.Add_Click({
    $selClient = & $GetSelClient
    $selOF = & $GetSelOF
    if ($selClient -and $selOF) {
        $inputSN = $txtNewSN.Text
        $snList = Expand-SNRange $inputSN
        if ($snList -and $snList.Count -gt 0) {
            # Vérifier conflits globaux
            $conflicts = @()
            foreach ($s in $snList) {
                $found = Find-OFBySN_Global $global:currentRegistry $s
                if ($found -and ($found.Client -ne $selClient -or $found.OF -ne $selOF)) {
                    $conflicts += "$s (dans OF $($found.OF) chez $($found.Client))"
                }
            }

            if ($conflicts.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Impossible, SN déjà assignés ailleurs :`n" + ($conflicts -join "`n"), "Conflit SN")
            }
            else {
                $currentList = $global:currentRegistry[$selClient][$selOF]
                $addedCount = 0
                foreach ($s in $snList) {
                    if ($currentList -notcontains $s) {
                        $global:currentRegistry[$selClient][$selOF] += $s
                        $addedCount++
                    }
                }
                if ($addedCount -gt 0) {
                    Save-GlobalRegistry $global:currentRegistry
                    $txtNewSN.Text = ""
                    & $RefreshOFList # Update Count dans le titre
                    $statusLabel.Text = "$addedCount SN ajouté(s) avec succès."
                    $statusLabel.ForeColor = $ColorSuccess
                } else { 
                    $statusLabel.Text = "Ces SN sont déjà dans l'OF." 
                    $statusLabel.ForeColor = $ColorPrimary
                }
            }
        } else { [System.Windows.Forms.MessageBox]::Show("Format invalide.`nUtilisez 'XXXXXX' ou 'XXXXXX-YYYYYY'.", "Erreur") }
    }
})

$btnDelSN.Add_Click({
    $selClient = & $GetSelClient
    $selOF = & $GetSelOF
    if ($selClient -and $selOF) {
        $inputSN = $txtNewSN.Text
        if (-not $inputSN -and $lstSN.SelectedItem) {
            $toDel = $lstSN.SelectedItem
            $global:currentRegistry[$selClient][$selOF] = @($global:currentRegistry[$selClient][$selOF] | Where-Object { $_ -ne $toDel })
            Save-GlobalRegistry $global:currentRegistry
            & $RefreshOFList
            $statusLabel.Text = "1 SN supprimé."
            $statusLabel.ForeColor = $ColorDanger
        } elseif ($inputSN) {
            $snList = Expand-SNRange $inputSN
            if ($snList) {
                $before = $global:currentRegistry[$selClient][$selOF].Count
                $global:currentRegistry[$selClient][$selOF] = @($global:currentRegistry[$selClient][$selOF] | Where-Object { $snList -notcontains $_ })
                $after = $global:currentRegistry[$selClient][$selOF].Count
                if ($before -ne $after) {
                    Save-GlobalRegistry $global:currentRegistry
                    $txtNewSN.Text = ""
                    & $RefreshOFList
                    $statusLabel.Text = "$($before - $after) SN supprimés."
                    $statusLabel.ForeColor = $ColorDanger
                } else { 
                    $statusLabel.Text = "Aucun SN correspondant trouvé pour suppression." 
                    $statusLabel.ForeColor = $ColorText
                }
            }
        } else { [System.Windows.Forms.MessageBox]::Show("Sélectionnez un SN ou saisissez une plage a supprimer.", "Info") }
    }
})

# --- RECHERCHE GLOBALE ---

$btnSearch.Add_Click({
    $snToFind = $txtSearchSN.Text.Trim()
    if ($snToFind) {
        $found = Find-OFBySN_Global $global:currentRegistry $snToFind
        if ($found) {
            # Select Client
            $idxC = $lstClient.FindString($found.Client)
            if ($idxC -ge 0) { $lstClient.SelectedIndex = $idxC }
            
            # Select OF
            $idxO = $lstOF.FindString($found.OF)
            if ($idxO -ge 0) { $lstOF.SelectedIndex = $idxO }
            
            # Select SN
            $idxS = $lstSN.FindString($snToFind)
            if ($idxS -ge 0) { $lstSN.SelectedIndex = $idxS; $lstSN.Focus() }
            
            $statusLabel.Text = "SN Trouvé : Client $($found.Client) / OF $($found.OF)"
            $statusLabel.ForeColor = $ColorSuccess
        }
        else {
            $statusLabel.Text = "SN $snToFind introuvable dans la base globale."
            $statusLabel.ForeColor = $ColorDanger
            [System.Windows.Forms.MessageBox]::Show("SN non trouvé dans la base globale.", "Résultat")
        }
    }
})

# INITIALISATION ET LANCEMENT
& $RefreshClientList
$Form.Add_Shown({ $Form.Activate() })
[void]$Form.ShowDialog()
