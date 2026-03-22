<#
    Gestion_OF_MultiClients.ps1
    Outil autonome pour la gestion des Ordres de Fabrication (OF) et Numéros de Série (SN)
    Support multi-clients.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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
# INTERFACE GRAPHIQUE (WINFORMS)
# ==========================================

$global:currentRegistry = Get-GlobalRegistry

$Form = New-Object Windows.Forms.Form
$Form.Text = "Gestionnaire Central Multi-Clients des OF"
$Form.Size = New-Object Drawing.Size(900, 550)
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedDialog"
$Form.MaximizeBox = $false
$Form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$mainTable = New-Object Windows.Forms.TableLayoutPanel
$mainTable.Dock = "Fill"
$mainTable.ColumnCount = 3
$mainTable.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 30)))
$mainTable.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 30)))
$mainTable.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent, 40)))
$Form.Controls.Add($mainTable)

# --- COL 1 : CLIENTS ---
$panelClient = New-Object Windows.Forms.FlowLayoutPanel
$panelClient.Dock = "Fill" ; $panelClient.FlowDirection = "TopDown" ; $panelClient.Padding = New-Object Windows.Forms.Padding(10)
$mainTable.Controls.Add($panelClient, 0, 0)

$lblClient = New-Object Windows.Forms.Label; $lblClient.Text = "Clients :"; $lblClient.AutoSize = $true; $lblClient.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lstClient = New-Object Windows.Forms.ListBox; $lstClient.Width = 220; $lstClient.Height = 320
$txtNewClient = New-Object Windows.Forms.TextBox; $txtNewClient.Width = 220
$btnAddClient = New-Object Windows.Forms.Button; $btnAddClient.Text = "Ajouter Client"; $btnAddClient.Width = 220
$btnDelClient = New-Object Windows.Forms.Button; $btnDelClient.Text = "Supprimer Client"; $btnDelClient.Width = 220
$btnImport = New-Object Windows.Forms.Button; $btnImport.Text = "Importer Wattsy (depuis V1)"; $btnImport.Width = 220 ; $btnImport.Margin = New-Object Windows.Forms.Padding(0, 15, 0, 0)

$panelClient.Controls.Add($lblClient); $panelClient.Controls.Add($lstClient); $panelClient.Controls.Add($txtNewClient)
$panelClient.Controls.Add($btnAddClient); $panelClient.Controls.Add($btnDelClient); $panelClient.Controls.Add($btnImport)

# --- COL 2 : OF ---
$panelOF = New-Object Windows.Forms.FlowLayoutPanel
$panelOF.Dock = "Fill" ; $panelOF.FlowDirection = "TopDown" ; $panelOF.Padding = New-Object Windows.Forms.Padding(10)
$mainTable.Controls.Add($panelOF, 1, 0)

$lblSearch = New-Object Windows.Forms.Label; $lblSearch.Text = "Recherche Globale via SN :"; $lblSearch.AutoSize = $true
$searchFlow = New-Object Windows.Forms.FlowLayoutPanel; $searchFlow.AutoSize = $true ; $searchFlow.Margin = New-Object Windows.Forms.Padding(0)
$txtSearchSN = New-Object Windows.Forms.TextBox; $txtSearchSN.Width = 130
$btnSearch = New-Object Windows.Forms.Button; $btnSearch.Text = "Chercher"; $btnSearch.Width = 80
$searchFlow.Controls.Add($txtSearchSN); $searchFlow.Controls.Add($btnSearch)

$sep1 = New-Object Windows.Forms.Label; $sep1.Text = "-----------------------------"; $sep1.AutoSize = $true
$lblOF = New-Object Windows.Forms.Label; $lblOF.Text = "OF du Client :"; $lblOF.AutoSize = $true; $lblOF.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lstOF = New-Object Windows.Forms.ListBox; $lstOF.Width = 220; $lstOF.Height = 250
$txtNewOF = New-Object Windows.Forms.TextBox; $txtNewOF.Width = 220; $txtNewOF.MaxLength = 10
$btnAddOF = New-Object Windows.Forms.Button; $btnAddOF.Text = "Ajouter OF"; $btnAddOF.Width = 220
$btnDelOF = New-Object Windows.Forms.Button; $btnDelOF.Text = "Supprimer OF"; $btnDelOF.Width = 220

$panelOF.Controls.Add($lblSearch); $panelOF.Controls.Add($searchFlow); $panelOF.Controls.Add($sep1); $panelOF.Controls.Add($lblOF)
$panelOF.Controls.Add($lstOF); $panelOF.Controls.Add($txtNewOF); $panelOF.Controls.Add($btnAddOF); $panelOF.Controls.Add($btnDelOF)

# --- COL 3 : SN ---
$panelSN = New-Object Windows.Forms.FlowLayoutPanel
$panelSN.Dock = "Fill" ; $panelSN.FlowDirection = "TopDown" ; $panelSN.Padding = New-Object Windows.Forms.Padding(10)
$mainTable.Controls.Add($panelSN, 2, 0)

$lblSNList = New-Object Windows.Forms.Label; $lblSNList.Text = "Numéros de Série (SN) :"; $lblSNList.AutoSize = $true; $lblSNList.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lstSN = New-Object Windows.Forms.ListBox; $lstSN.Width = 320; $lstSN.Height = 310
$lblAddSN = New-Object Windows.Forms.Label; $lblAddSN.Text = "Ajout SN (ex: 043590 ou 043590-043600) :"; $lblAddSN.AutoSize = $true
$txtNewSN = New-Object Windows.Forms.TextBox; $txtNewSN.Width = 320
$btnFlowSN = New-Object Windows.Forms.FlowLayoutPanel; $btnFlowSN.AutoSize = $true ; $btnFlowSN.Margin = New-Object Windows.Forms.Padding(0)
$btnAddSN = New-Object Windows.Forms.Button; $btnAddSN.Text = "Ajouter SN"; $btnAddSN.Width = 155
$btnDelSN = New-Object Windows.Forms.Button; $btnDelSN.Text = "Supprimer SN"; $btnDelSN.Width = 155
$btnFlowSN.Controls.Add($btnAddSN); $btnFlowSN.Controls.Add($btnDelSN)

$lblStatus = New-Object Windows.Forms.Label; $lblStatus.Text = "Prêt."; $lblStatus.AutoSize = $true; $lblStatus.ForeColor = "Blue" ; $lblStatus.Margin = New-Object Windows.Forms.Padding(0, 10, 0, 0)

$panelSN.Controls.Add($lblSNList); $panelSN.Controls.Add($lstSN); $panelSN.Controls.Add($lblAddSN); $panelSN.Controls.Add($txtNewSN)
$panelSN.Controls.Add($btnFlowSN); $panelSN.Controls.Add($lblStatus)

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
        $lblStatus.Text = "$count SN dans l'OF $selOF (Client: $selClient)"
        $txtNewOF.Text = ""
        $panelSN.Enabled = $true
    }
    else {
        $lblStatus.Text = "Sélectionnez un OF"
        $panelSN.Enabled = $false
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
        $panelOF.Enabled = $true
        $txtNewClient.Text = ""
    }
    else {
        $panelOF.Enabled = $false
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

# DOUBLE CLICS (Actions rapides demandées par l'utilisateur)
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

# ENTREES (ENTER KEYS)
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
                & $RefreshClientList # Met à jour le compteur du client
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
                    $lblStatus.Text = "$addedCount SN ajoutés."
                } else { $lblStatus.Text = "Ces SN sont déjà dans l'OF." }
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
                    $lblStatus.Text = "$($before - $after) SN supprimés."
                } else { $lblStatus.Text = "Aucun SN trouvé." }
            }
        } else { [System.Windows.Forms.MessageBox]::Show("Sélectionnez un SN ou saisissez une plage.", "Info") }
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
            
            $lblStatus.Text = "Trouvé chez $($found.Client), OF $($found.OF)"
        }
        else {
            $lblStatus.Text = "SN $snToFind non trouvé."
            [System.Windows.Forms.MessageBox]::Show("SN non trouvé dans la base globale.", "Résultat")
        }
    }
})

# INITIALISATION ET LANCEMENT
& $RefreshClientList
$Form.Add_Shown({ $Form.Activate() })
[void]$Form.ShowDialog()
