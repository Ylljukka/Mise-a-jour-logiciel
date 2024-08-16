# Script permettant de mettre à jour les logiciels via Winget
# La 2nd partie permet de mettre à jour les packages Python installés
# REYNAUD Thomas 26/07/2024


# Fonction pour rafraîchir la liste des logiciels à mettre à jour
function Refresh-SoftwareList {
    $panel.Controls.Clear()
    $updatable = Get-WinGetPackage -Source winget | Where-Object IsUpdateAvailable | Select-Object -ExpandProperty Id

    $global:checkboxes = @{}
    $yPosition = 10

    foreach ($logiciel in $updatable) {
        $checkbox = New-Object System.Windows.Forms.CheckBox
        $checkbox.Text = $logiciel
        $checkbox.Width = 350
        $checkbox.Location = New-Object System.Drawing.Point(10, $yPosition)
        $panel.Controls.Add($checkbox)
        $global:checkboxes[$logiciel] = $checkbox
        $yPosition += 20
    }
}

# Fonction de rafraichissement liste package Python
function Refresh-Python {
    param([int]$exist)
    $panel_python.Controls.Clear()

    # Si PIP est installé, on le met à jour
    if ($exist -eq 1) {
    $pip_version = python -m pip --version

    if ($pip_version -ne $null) {
    python -m pip install --upgrade pip
    }
    }

    # Puis on liste les packages à mettre à jour
    $pip_out_put = python -m pip list --outdated --format=json
    if ($pip_out_put -ne "") {
    $outdated_Packages = ConvertFrom-Json $pip_out_put


    $global:checkboxes_python = @{}
    $yPosition_python = 10

    foreach ($package in $outdated_Packages) {
        $checkbox_python = New-Object System.Windows.Forms.CheckBox
        $checkbox_python.Text = $package.name
        $checkbox_python.Width = 200
        $checkbox_python.Location = New-Object System.Drawing.Point(10, $yPosition_python)
        $panel_python.Controls.Add($checkbox_python)
        $global:checkboxes_python[$package.name] = $checkbox_python
        $yPosition_python += 20
    }
}
    }


# Fonction pour la mise à jour logiciel
function Update-Software {
    param (
        [string]$software
    )

    try {
        Update-WinGetPackage $software -ErrorAction Stop
        Start-sleep -seconds 1
        Add-Text "Mise à jour de $software terminée avec succès." -ForegroundColor Green
    } catch {
        $errorMessage = $_.Exception.Message
        Start-sleep -seconds 1
        Add-Text "Erreur lors de la mise à jour de $software : $errorMessage" -ForegroundColor Red
    }
}

# Fonction mise à jours des packages Python
function Update-Package {
    param (
        [string]$package
    )
    try{
        python -m pip install $package --upgrade
        Start-sleep -seconds 1
        Add-Text "Mise à jour du package $package terminée avec succès." -ForegroundColor Green
    } catch {
        $errorMessage = $_.Exception.Message
        Start-sleep -seconds 1
        Add-Text "Erreur lors de la mise à jour du package $package : $errorMessage" -ForegroundColor Red
    }

    }

# Fonction pour ajouter du texte à la boîte de texte
function Add-Text {
    param (
        [string]$text,
        [System.Drawing.Color]$ForegroundColor = [System.Drawing.Color]::Black
    )
    
    $RichTextBox.SelectionStart = $RichTextBox.TextLength
    $RichTextBox.SelectionLength = 0
    $RichTextBox.SelectionColor = $ForegroundColor
    $RichTextBox.AppendText("$text`r`n")
}

# Fonction pour ouvrir les notes de version
function Notes-Version {
    param([string]$chemin)
    $ouverture = Get-Content $chemin -Raw

     # Créer la fenêtre de visualisation de texte
     $form2 = New-Object System.Windows.Forms.Form
     $form2.Text = "Notes de version"
     $form2.Size = New-Object System.Drawing.Size(500, 200)
     
     # Créer une TextBox en lecture seule
     $textBox = New-Object System.Windows.Forms.TextBox
     $textBox.Multiline = $true
     $textBox.ReadOnly = $true
     $textBox.Dock = [System.Windows.Forms.DockStyle]::Fill
     $textBox.ScrollBars = "Both"
     $textBox.Text = $ouverture
     
     # Ajouter la TextBox à la fenêtre
     $form2.Controls.Add($textBox)
     
     # Afficher la fenêtre
     $form2.ShowDialog()
}

# La fenêtre principale du logiciel
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Threading
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$Form = New-Object System.Windows.Forms.Form
$Form.ClientSize = "700,600"
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $false
$Form.Text = "Mise à jour logiciels et packages V2.0.0"

#Les boutons
$Valider = New-Object System.Windows.Forms.Button
$Valider.Location = New-Object System.Drawing.Point(500,530)
$Valider.Width = 150
$Valider.Height = 50
$Valider.Text = "Mettre à jour"
$Form.controls.Add($Valider)

$Quitter = New-Object System.Windows.Forms.Button
$Quitter.Location = New-Object System.Drawing.Point(50,530)
$Quitter.Width = 150
$Quitter.Height = 50
$Quitter.Text = "Quitter"
$Form.controls.Add($Quitter)

#Le menu cascade
$Menu = New-Object System.Windows.Forms.MenuStrip
$Menu.Location = New-Object System.Drawing.Point(0,0)
$Menu.ShowItemToolTips = $True
$MenuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuFile.Text = "&Fichier"
$MenuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAbout.Text = "&A propos"
$MenuFileQuit = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuFileQuit.Text = "&Quitter"
$MenuAboutNotes = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAboutNotes.Text = "&Notes de version"
$MenuAboutDons = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAboutDons.Text = "&Dons"

# On rattache Quitter à fichier
$MenuFile.DropDownItems.Add($MenuFileQuit)
#On rattache les notes de version à A propos
$MenuAbout.DropDownItems.Add($MenuAboutNotes)
#On rattache l'onglet Dons à A propos
$MenuAbout.DropDownItems.Add($MenuAboutDons)
#On rattache le tous au menu cascade
$Menu.Items.AddRange(@($MenuFile,$MenuAbout))
# Ajouter le MenuStrip au formulaire
$Form.MainMenuStrip = $Menu
$Form.Controls.Add($Menu)

#Event menu cascade Quitter
$MenuFileQuit.Add_Click({
    $Form.Close()
    })

#Ouverture des notes de version
$MenuAboutNotes.Add_Click({
    try {
    Write-Host $PSCommandPath
    $scriptDir = Split-Path -Parent $PSCommandPath
    $chemin = Join-Path -Path $scriptDir -ChildPath "Notes de version.txt"
    if (Test-Path -Path $chemin){
    Write-Host $chemin
    Notes-Version -chemin $chemin
    }else{
        Add-Text "Fichier Notes de version.txt non présent dans le répertoire du script" -ForegroundColor Red
    }
    } finally {}
})

#Event menu cascade Dons
$MenuAboutDons.Add_Click({
    Start-Process "https://www.paypal.com/paypalme/Thomas6598"
    })

# Ajouter un contrôle de texte pour afficher les messages
$RichTextBox = New-Object System.Windows.Forms.RichTextBox
$RichTextBox.Location = New-Object System.Drawing.Point(300, 60)
$RichTextBox.Size = New-Object System.Drawing.Size(350, 460)
$RichTextBox.Multiline = $true
$RichTextBox.ScrollBars = 'Vertical'
$RichTextBox.ReadOnly = $true
$Form.Controls.Add($RichTextBox)


# Affichage des logiciels à mettre à jour

# Créer un panel pour contenir les cases à cocher
$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(50, 70)
$panel.Size = New-Object System.Drawing.Size(200, 150)
$Form.Controls.Add($panel)

$Group_box_logiciel = New-Object System.Windows.Forms.GroupBox
$Group_box_logiciel.Location = New-Object System.Drawing.Point(30,50)
$Group_box_logiciel.Width = 250
$Group_box_logiciel.Height = 250
$Group_box_logiciel.Text = "Liste des logiciels à mettre à jour" 

# Checker si module winget  est installé
$nom_module = "Microsoft.WinGet.Client"

# Récupérer le module
$installed_module = Get-Module -ListAvailable -Name $nom_module

# Si pas installé, on l'installe
if ($null -eq $installed_module) {
    Install-Module -Name Microsoft.WinGet.Client -Force

#Si installé on le met  à jour
}else{ 
    $module_version_instal = $module | Select-Object Version
    $available_Module = Find-Module -Name $nom_module -ErrorAction SilentlyContinue
    $source_microsoft = $available_Module | Select-Object Version

    if ($module_version_instal -lt $source_microsoft) {
        Update-Module -Name Microsoft.WinGet.Client -ErrorAction SilentlyContinue    
    }
}


# Affichage des packages Python
$panel_python = New-Object System.Windows.Forms.Panel
$panel_python.Location = New-Object System.Drawing.Point(50, 350)
$panel_python.Size = New-Object System.Drawing.Size(160, 150)
$Form.Controls.Add($panel_python)

$Group_box_package = New-Object System.Windows.Forms.GroupBox
$Group_box_package.Location = New-Object System.Drawing.Point(30,330)
$Group_box_package.Width = 250
$Group_box_package.Height = 190
$Group_box_package.Text = "Liste des packages Python à mettre à jour"

$Form.controls.AddRange(@($Group_box_logiciel,$Group_box_package))

# Initialisation de la liste des logiciels
Refresh-SoftwareList

# Vérifier si pip est installé
$PipInstalled = (Get-Command pip -ErrorAction SilentlyContinue) -ne $null

# Si oui on lance la fonction d'affichage des paques à MAJ
if ($PipInstalled) {
    Refresh-Python -exist 1
}

#Si on valide
$Valider.Add_Click({
    $selectedCount = 0
    $selectedCountPython = 0
    $checkboxKeys = @($checkboxes.Keys)
    $checkboxKeysPython = @($checkboxes_python.Keys)

    # Désactiver le bouton "Mettre à jour" pendant la mise à jour
    $Valider.Enabled = $false

    foreach ($software in $checkboxKeys) {
        if ($checkboxes[$software].Checked) {
            $selectedCount += 1
            if ($selectedCount -gt 0) {
                Add-Text "Mise à jour de $software en cours..."
                Start-sleep -seconds 1
                Update-Software $software
                $checkboxes[$software].Checked = $false
            }           
            
        }
        
                               
    }

    # Pour les packages Python on fait la MAJ
    foreach ($package in $checkboxKeysPython) {
        if ($checkboxes_python[$package].Checked) {
            $selectedCountPython += 1
            if ($selectedCountPython -gt 0) {
                Add-Text "Mise à jour du package $package en cours..."
                Start-sleep -seconds 1
                Update-Package $package
                $checkboxes_python[$package].Checked = $false

                }
            }
        }
    
    if ($selectedCount -gt 0) {
    Refresh-SoftwareList 
    }
    if ($selectedCountPython -gt 0) {
    Refresh-Python -exist 0
    }
            
    
    if ($selectedCount -eq 0 -and $selectedCountPython -eq 0) {
        Add-Text "Veuillez cocher au moins un applicatif ou un package avant de valider." -ForegroundColor Red
        }
    

    # Réactiver le bouton "Mettre à jour" après la fin des mises à jour
    $Valider.Enabled = $true
     
})      

#Fermeture de la fenêtre si bouton Quitter
$Quitter.Add_Click({
    $Form.Close()
})

$Form.ShowDialog()



