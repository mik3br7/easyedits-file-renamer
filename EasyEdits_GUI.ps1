Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. "$PSScriptRoot\EasyEdits.ps1"

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "EasyEdits"
$form.Size = New-Object System.Drawing.Size(1100, 865)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1100, 865)
$form.AllowDrop = $true

$title = New-Object System.Windows.Forms.Label
$title.Text = "EasyEdits"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(20, 15)
$title.AutoSize = $true
$form.Controls.Add($title)

$instructions = New-Object System.Windows.Forms.Label
$instructions.Text = "Open a folder manually with 'Browse' button, or drag and drop the folder to the landing area with the desired files for editing. Replace, Remove, or Sequantially Rename entire folders or specific checked files. 'Preview changes' button to see changes, verify desired changes and then 'Apply Rename' button."
$instructions.Location = New-Object System.Drawing.Point(22, 55)
$instructions.Size = New-Object System.Drawing.Size(1030, 36)
$form.Controls.Add($instructions)

function Add-FieldLabel {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y,
        [int]$Width = 120
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, 22)
    $form.Controls.Add($label)
    return $label
}

Add-FieldLabel -Text "Directory" -X 20 -Y 136 -Width 65 | Out-Null
$folderBox = New-Object System.Windows.Forms.TextBox
$folderBox.Location = New-Object System.Drawing.Point(95, 133)
$folderBox.Size = New-Object System.Drawing.Size(520, 23)
$folderBox.Text = (Get-Location).Path
$form.Controls.Add($folderBox)

$dropZonePanel = New-Object System.Windows.Forms.Panel
$dropZonePanel.Location = New-Object System.Drawing.Point(457, 96)
$dropZonePanel.Size = New-Object System.Drawing.Size(185, 34)
$dropZonePanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dropZonePanel.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$dropZonePanel.AllowDrop = $true
$form.Controls.Add($dropZonePanel)

$dropZoneLabel = New-Object System.Windows.Forms.Label
$dropZoneLabel.Text = "Drop Folder Here"
$dropZoneLabel.Location = New-Object System.Drawing.Point(0, 0)
$dropZoneLabel.Size = New-Object System.Drawing.Size(185, 34)
$dropZoneLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$dropZoneLabel.BackColor = [System.Drawing.Color]::Transparent
$dropZonePanel.Controls.Add($dropZoneLabel)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse"
$browseButton.Location = New-Object System.Drawing.Point(825, 131)
$browseButton.Size = New-Object System.Drawing.Size(85, 28)
$form.Controls.Add($browseButton)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point(920, 131)
$refreshButton.Size = New-Object System.Drawing.Size(75, 28)
$form.Controls.Add($refreshButton)

$applyButton = New-Object System.Windows.Forms.Button
$applyButton.Text = "Apply Rename"
$applyButton.Location = New-Object System.Drawing.Point(760, 52)
$applyButton.Size = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($applyButton)

$undoButton = New-Object System.Windows.Forms.Button
$undoButton.Text = "Undo Changes"
$undoButton.Location = New-Object System.Drawing.Point(760, 88)
$undoButton.Size = New-Object System.Drawing.Size(170, 30)
$undoButton.Enabled = $false
$form.Controls.Add($undoButton)

$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Check All"
$selectAllButton.Location = New-Object System.Drawing.Point(20, 410)
$selectAllButton.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($selectAllButton)

$clearChecksButton = New-Object System.Windows.Forms.Button
$clearChecksButton.Text = "Clear Checks"
$clearChecksButton.Location = New-Object System.Drawing.Point(130, 410)
$clearChecksButton.Size = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($clearChecksButton)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready."
$statusLabel.Location = New-Object System.Drawing.Point(260, 415)
$statusLabel.Size = New-Object System.Drawing.Size(685, 22)
$form.Controls.Add($statusLabel)

$toolTabs = New-Object System.Windows.Forms.TabControl
$toolTabs.Location = New-Object System.Drawing.Point(20, 205)
$toolTabs.Size = New-Object System.Drawing.Size(1055, 195)
$form.Controls.Add($toolTabs)

$replaceTab = New-Object System.Windows.Forms.TabPage
$replaceTab.Text = "Replace"
$toolTabs.TabPages.Add($replaceTab)

$removeTab = New-Object System.Windows.Forms.TabPage
$removeTab.Text = "Remove"
$toolTabs.TabPages.Add($removeTab)

$addTab = New-Object System.Windows.Forms.TabPage
$addTab.Text = "Add"
$toolTabs.TabPages.Add($addTab)

$sequenceTab = New-Object System.Windows.Forms.TabPage
$sequenceTab.Text = "Sequential Rename"
$toolTabs.TabPages.Add($sequenceTab)

$replaceTextLabel = New-Object System.Windows.Forms.Label
$replaceTextLabel.Text = "Filename Text To Be Replaced"
$replaceTextLabel.Location = New-Object System.Drawing.Point(20, 20)
$replaceTextLabel.Size = New-Object System.Drawing.Size(170, 22)
$replaceTab.Controls.Add($replaceTextLabel)

$replaceFindTextBox = New-Object System.Windows.Forms.TextBox
$replaceFindTextBox.Location = New-Object System.Drawing.Point(195, 18)
$replaceFindTextBox.Size = New-Object System.Drawing.Size(170, 23)
$replaceTab.Controls.Add($replaceFindTextBox)

$replaceWithTextLabel = New-Object System.Windows.Forms.Label
$replaceWithTextLabel.Text = "Filename Text Replaced With"
$replaceWithTextLabel.Location = New-Object System.Drawing.Point(390, 20)
$replaceWithTextLabel.Size = New-Object System.Drawing.Size(170, 22)
$replaceTab.Controls.Add($replaceWithTextLabel)

$replaceWithTextBox = New-Object System.Windows.Forms.TextBox
$replaceWithTextBox.Location = New-Object System.Drawing.Point(565, 18)
$replaceWithTextBox.Size = New-Object System.Drawing.Size(170, 23)
$replaceTab.Controls.Add($replaceWithTextBox)

$replaceSelectedButton = New-Object System.Windows.Forms.Button
$replaceSelectedButton.Text = "Preview Changes"
$replaceSelectedButton.Location = New-Object System.Drawing.Point(760, 16)
$replaceSelectedButton.Size = New-Object System.Drawing.Size(170, 30)
$replaceTab.Controls.Add($replaceSelectedButton)

$replaceHint = New-Object System.Windows.Forms.Label
$replaceHint.Text = "Use both fields together to swap matching text across checked file names."
$replaceHint.Location = New-Object System.Drawing.Point(20, 58)
$replaceHint.Size = New-Object System.Drawing.Size(700, 20)
$replaceTab.Controls.Add($replaceHint)

$replaceOriginalExample = New-Object System.Windows.Forms.Label
$replaceOriginalExample.Text = "Example: Original File Name - report_draft.txt"
$replaceOriginalExample.Location = New-Object System.Drawing.Point(20, 82)
$replaceOriginalExample.Size = New-Object System.Drawing.Size(700, 20)
$replaceOriginalExample.ForeColor = [System.Drawing.Color]::DimGray
$replaceTab.Controls.Add($replaceOriginalExample)

$replaceExample = New-Object System.Windows.Forms.Label
$replaceExample.Text = "Example: Filename Text To Be Replaced = 'draft' -> Filename Text Replaced With = 'final'"
$replaceExample.Location = New-Object System.Drawing.Point(20, 102)
$replaceExample.Size = New-Object System.Drawing.Size(700, 20)
$replaceExample.ForeColor = [System.Drawing.Color]::DimGray
$replaceTab.Controls.Add($replaceExample)

$replaceResultExample = New-Object System.Windows.Forms.Label
$replaceResultExample.Text = "Example: End Result - report_draft.txt -> report_final.txt"
$replaceResultExample.Location = New-Object System.Drawing.Point(20, 122)
$replaceResultExample.Size = New-Object System.Drawing.Size(700, 20)
$replaceResultExample.ForeColor = [System.Drawing.Color]::DimGray
$replaceTab.Controls.Add($replaceResultExample)

$removeLabel = New-Object System.Windows.Forms.Label
$removeLabel.Text = "Remove Text"
$removeLabel.Location = New-Object System.Drawing.Point(20, 20)
$removeLabel.Size = New-Object System.Drawing.Size(150, 22)
$removeTab.Controls.Add($removeLabel)

$removeTextBox = New-Object System.Windows.Forms.TextBox
$removeTextBox.Location = New-Object System.Drawing.Point(175, 18)
$removeTextBox.Size = New-Object System.Drawing.Size(220, 23)
$removeTab.Controls.Add($removeTextBox)

$removeSelectedButton = New-Object System.Windows.Forms.Button
$removeSelectedButton.Text = "Preview Changes"
$removeSelectedButton.Location = New-Object System.Drawing.Point(760, 16)
$removeSelectedButton.Size = New-Object System.Drawing.Size(170, 30)
$removeTab.Controls.Add($removeSelectedButton)

$removeHint = New-Object System.Windows.Forms.Label
$removeHint.Text = "This removes the typed text anywhere it appears in the checked file names."
$removeHint.Location = New-Object System.Drawing.Point(20, 58)
$removeHint.Size = New-Object System.Drawing.Size(700, 20)
$removeTab.Controls.Add($removeHint)

$removeOriginalExample = New-Object System.Windows.Forms.Label
$removeOriginalExample.Text = "Example: Original File Name - summer_EDIT_photo.jpg"
$removeOriginalExample.Location = New-Object System.Drawing.Point(20, 82)
$removeOriginalExample.Size = New-Object System.Drawing.Size(700, 20)
$removeOriginalExample.ForeColor = [System.Drawing.Color]::DimGray
$removeTab.Controls.Add($removeOriginalExample)

$removeExample = New-Object System.Windows.Forms.Label
$removeExample.Text = "Example: Remove Text = 'EDIT_'"
$removeExample.Location = New-Object System.Drawing.Point(20, 102)
$removeExample.Size = New-Object System.Drawing.Size(700, 20)
$removeExample.ForeColor = [System.Drawing.Color]::DimGray
$removeTab.Controls.Add($removeExample)

$removeResultExample = New-Object System.Windows.Forms.Label
$removeResultExample.Text = "Example: End Result - summer_EDIT_photo.jpg -> summer_photo.jpg"
$removeResultExample.Location = New-Object System.Drawing.Point(20, 122)
$removeResultExample.Size = New-Object System.Drawing.Size(700, 20)
$removeResultExample.ForeColor = [System.Drawing.Color]::DimGray
$removeTab.Controls.Add($removeResultExample)

$addTextLabel = New-Object System.Windows.Forms.Label
$addTextLabel.Text = "Text To Add"
$addTextLabel.Location = New-Object System.Drawing.Point(20, 20)
$addTextLabel.Size = New-Object System.Drawing.Size(120, 22)
$addTab.Controls.Add($addTextLabel)

$addTextBox = New-Object System.Windows.Forms.TextBox
$addTextBox.Location = New-Object System.Drawing.Point(145, 18)
$addTextBox.Size = New-Object System.Drawing.Size(180, 23)
$addTab.Controls.Add($addTextBox)

$addPositionLabel = New-Object System.Windows.Forms.Label
$addPositionLabel.Text = "Add Position"
$addPositionLabel.Location = New-Object System.Drawing.Point(345, 20)
$addPositionLabel.Size = New-Object System.Drawing.Size(85, 22)
$addTab.Controls.Add($addPositionLabel)

$addPositionDropdown = New-Object System.Windows.Forms.ComboBox
$addPositionDropdown.Location = New-Object System.Drawing.Point(435, 18)
$addPositionDropdown.Size = New-Object System.Drawing.Size(130, 23)
$addPositionDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$addPositionDropdown.Items.AddRange(@("Beginning", "End", "Before reference text", "After reference text"))
$addPositionDropdown.SelectedIndex = 0
$addTab.Controls.Add($addPositionDropdown)

$addAnchorLabel = New-Object System.Windows.Forms.Label
$addAnchorLabel.Text = "Reference Text"
$addAnchorLabel.Location = New-Object System.Drawing.Point(580, 20)
$addAnchorLabel.Size = New-Object System.Drawing.Size(90, 22)
$addTab.Controls.Add($addAnchorLabel)

$addAnchorTextBox = New-Object System.Windows.Forms.TextBox
$addAnchorTextBox.Location = New-Object System.Drawing.Point(675, 18)
$addAnchorTextBox.Size = New-Object System.Drawing.Size(70, 23)
$addTab.Controls.Add($addAnchorTextBox)

$addSelectedButton = New-Object System.Windows.Forms.Button
$addSelectedButton.Text = "Preview Changes"
$addSelectedButton.Location = New-Object System.Drawing.Point(760, 16)
$addSelectedButton.Size = New-Object System.Drawing.Size(170, 30)
$addTab.Controls.Add($addSelectedButton)

$addHint = New-Object System.Windows.Forms.Label
$addHint.Text = "This adds typed text to checked file names at the beginning, end, or beside matching reference text."
$addHint.Location = New-Object System.Drawing.Point(20, 58)
$addHint.Size = New-Object System.Drawing.Size(720, 20)
$addTab.Controls.Add($addHint)

$addOriginalExample = New-Object System.Windows.Forms.Label
$addOriginalExample.Text = "Example: Original File Name - report_final.txt"
$addOriginalExample.Location = New-Object System.Drawing.Point(20, 82)
$addOriginalExample.Size = New-Object System.Drawing.Size(720, 20)
$addOriginalExample.ForeColor = [System.Drawing.Color]::DimGray
$addTab.Controls.Add($addOriginalExample)

$addExample = New-Object System.Windows.Forms.Label
$addExample.Text = "Example: Text To Add = 'v2_', Add Position = 'Beginning'"
$addExample.Location = New-Object System.Drawing.Point(20, 102)
$addExample.Size = New-Object System.Drawing.Size(720, 20)
$addExample.ForeColor = [System.Drawing.Color]::DimGray
$addTab.Controls.Add($addExample)

$addResultExample = New-Object System.Windows.Forms.Label
$addResultExample.Text = "Example: End Result - report_final.txt -> v2_report_final.txt"
$addResultExample.Location = New-Object System.Drawing.Point(20, 122)
$addResultExample.Size = New-Object System.Drawing.Size(720, 20)
$addResultExample.ForeColor = [System.Drawing.Color]::DimGray
$addTab.Controls.Add($addResultExample)

$renameAllLabel = New-Object System.Windows.Forms.Label
$renameAllLabel.Text = "Rename Checked In Order As"
$renameAllLabel.Location = New-Object System.Drawing.Point(20, 18)
$renameAllLabel.Size = New-Object System.Drawing.Size(165, 22)
$sequenceTab.Controls.Add($renameAllLabel)

$renameAllBox = New-Object System.Windows.Forms.TextBox
$renameAllBox.Location = New-Object System.Drawing.Point(190, 16)
$renameAllBox.Size = New-Object System.Drawing.Size(220, 23)
$sequenceTab.Controls.Add($renameAllBox)

$renameAllButton = New-Object System.Windows.Forms.Button
$renameAllButton.Text = "Preview Changes"
$renameAllButton.Location = New-Object System.Drawing.Point(760, 16)
$renameAllButton.Size = New-Object System.Drawing.Size(170, 30)
$sequenceTab.Controls.Add($renameAllButton)

$startNumberLabel = New-Object System.Windows.Forms.Label
$startNumberLabel.Text = "Start Number"
$startNumberLabel.Location = New-Object System.Drawing.Point(20, 55)
$startNumberLabel.Size = New-Object System.Drawing.Size(85, 22)
$sequenceTab.Controls.Add($startNumberLabel)

$startNumberBox = New-Object System.Windows.Forms.TextBox
$startNumberBox.Location = New-Object System.Drawing.Point(105, 53)
$startNumberBox.Size = New-Object System.Drawing.Size(55, 23)
$startNumberBox.Text = "1"
$sequenceTab.Controls.Add($startNumberBox)

$paddingLabel = New-Object System.Windows.Forms.Label
$paddingLabel.Text = "Padding"
$paddingLabel.Location = New-Object System.Drawing.Point(185, 55)
$paddingLabel.Size = New-Object System.Drawing.Size(55, 22)
$sequenceTab.Controls.Add($paddingLabel)

$paddingDropdown = New-Object System.Windows.Forms.ComboBox
$paddingDropdown.Location = New-Object System.Drawing.Point(245, 53)
$paddingDropdown.Size = New-Object System.Drawing.Size(95, 23)
$paddingDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$paddingDropdown.Items.AddRange(@("No padding", "2 digits", "3 digits", "4 digits"))
$paddingDropdown.SelectedIndex = 0
$sequenceTab.Controls.Add($paddingDropdown)

$separatorLabel = New-Object System.Windows.Forms.Label
$separatorLabel.Text = "Separator"
$separatorLabel.Location = New-Object System.Drawing.Point(350, 55)
$separatorLabel.Size = New-Object System.Drawing.Size(60, 22)
$sequenceTab.Controls.Add($separatorLabel)

$separatorDropdown = New-Object System.Windows.Forms.ComboBox
$separatorDropdown.Location = New-Object System.Drawing.Point(415, 53)
$separatorDropdown.Size = New-Object System.Drawing.Size(95, 23)
$separatorDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$separatorDropdown.Items.AddRange(@("Space", "Dash (-)", "Dash ( - )", "Underscore (_)", "None"))
$separatorDropdown.SelectedIndex = 0
$sequenceTab.Controls.Add($separatorDropdown)

$positionLabel = New-Object System.Windows.Forms.Label
$positionLabel.Text = "Number Position"
$positionLabel.Location = New-Object System.Drawing.Point(525, 55)
$positionLabel.Size = New-Object System.Drawing.Size(95, 22)
$sequenceTab.Controls.Add($positionLabel)

$positionDropdown = New-Object System.Windows.Forms.ComboBox
$positionDropdown.Location = New-Object System.Drawing.Point(625, 53)
$positionDropdown.Size = New-Object System.Drawing.Size(125, 23)
$positionDropdown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$positionDropdown.Items.AddRange(@("After name", "After name (no space)", "Before name"))
$positionDropdown.SelectedIndex = 0
$sequenceTab.Controls.Add($positionDropdown)

$sequenceHint = New-Object System.Windows.Forms.Label
$sequenceHint.Text = "This renames checked files in order using one base name, then adds numbers with your chosen padding, separator, and position."
$sequenceHint.Location = New-Object System.Drawing.Point(20, 88)
$sequenceHint.Size = New-Object System.Drawing.Size(720, 20)
$sequenceTab.Controls.Add($sequenceHint)

$sequenceOriginalExample = New-Object System.Windows.Forms.Label
$sequenceOriginalExample.Text = "Example: Original File Name - IMG_4321.jpg, IMG_4322.jpg, IMG_4323.jpg"
$sequenceOriginalExample.Location = New-Object System.Drawing.Point(20, 108)
$sequenceOriginalExample.Size = New-Object System.Drawing.Size(720, 20)
$sequenceOriginalExample.ForeColor = [System.Drawing.Color]::DimGray
$sequenceTab.Controls.Add($sequenceOriginalExample)

$sequenceExample = New-Object System.Windows.Forms.Label
$sequenceExample.Text = "Example: Rename Checked In Order As = 'Vacation', Start Number = '1', Padding = '2 digits', Separator = 'Space', Number Position = 'After name'"
$sequenceExample.Location = New-Object System.Drawing.Point(20, 128)
$sequenceExample.Size = New-Object System.Drawing.Size(735, 20)
$sequenceExample.ForeColor = [System.Drawing.Color]::DimGray
$sequenceTab.Controls.Add($sequenceExample)

$sequenceResultExample = New-Object System.Windows.Forms.Label
$sequenceResultExample.Text = "Example: End Result - Vacation 01.jpg, Vacation 02.jpg, Vacation 03.jpg"
$sequenceResultExample.Location = New-Object System.Drawing.Point(20, 148)
$sequenceResultExample.Size = New-Object System.Drawing.Size(720, 20)
$sequenceResultExample.ForeColor = [System.Drawing.Color]::DimGray
$sequenceTab.Controls.Add($sequenceResultExample)

function Set-ApplyButtonLayout {
    $selectedTab = $toolTabs.SelectedTab

    if ($selectedTab -eq $sequenceTab) {
        if ($applyButton.Parent -ne $sequenceTab) {
            $sequenceTab.Controls.Add($applyButton)
        }
        if ($undoButton.Parent -ne $sequenceTab) {
            $sequenceTab.Controls.Add($undoButton)
        }

        $applyButton.Location = New-Object System.Drawing.Point(760, 52)
        $applyButton.Size = New-Object System.Drawing.Size(170, 30)
        $undoButton.Location = New-Object System.Drawing.Point(760, 88)
        $undoButton.Size = New-Object System.Drawing.Size(170, 30)
        $applyButton.BringToFront()
        $undoButton.BringToFront()
        return
    }

    if ($selectedTab -eq $addTab) {
        if ($applyButton.Parent -ne $addTab) {
            $addTab.Controls.Add($applyButton)
        }
        if ($undoButton.Parent -ne $addTab) {
            $addTab.Controls.Add($undoButton)
        }

        $applyButton.Location = New-Object System.Drawing.Point(760, 52)
        $applyButton.Size = New-Object System.Drawing.Size(170, 30)
        $undoButton.Location = New-Object System.Drawing.Point(760, 88)
        $undoButton.Size = New-Object System.Drawing.Size(170, 30)
        $applyButton.BringToFront()
        $undoButton.BringToFront()
        return
    }

    if ($selectedTab -eq $removeTab) {
        if ($applyButton.Parent -ne $removeTab) {
            $removeTab.Controls.Add($applyButton)
        }
        if ($undoButton.Parent -ne $removeTab) {
            $removeTab.Controls.Add($undoButton)
        }

        $applyButton.Location = New-Object System.Drawing.Point(760, 52)
        $applyButton.Size = New-Object System.Drawing.Size(170, 30)
        $undoButton.Location = New-Object System.Drawing.Point(760, 88)
        $undoButton.Size = New-Object System.Drawing.Size(170, 30)
        $applyButton.BringToFront()
        $undoButton.BringToFront()
        return
    }

    if ($applyButton.Parent -ne $replaceTab) {
        $replaceTab.Controls.Add($applyButton)
    }
    if ($undoButton.Parent -ne $replaceTab) {
        $replaceTab.Controls.Add($undoButton)
    }

    $applyButton.Location = New-Object System.Drawing.Point(760, 52)
    $applyButton.Size = New-Object System.Drawing.Size(170, 30)
    $undoButton.Location = New-Object System.Drawing.Point(760, 88)
    $undoButton.Size = New-Object System.Drawing.Size(170, 30)
    $applyButton.BringToFront()
    $undoButton.BringToFront()
}

function Update-UndoButtonState {
    $undoButton.Enabled = [bool](@($script:lastUndoPlan).Count -gt 0)
    $undoButton.BringToFront()
}

function Update-ManualUndoButtonState {
    $undoManualRenameButton.Enabled = [bool](@($script:lastManualUndoPlan).Count -gt 0)
    $undoManualRenameButton.BringToFront()
}

Set-ApplyButtonLayout

$fileListLabel = New-Object System.Windows.Forms.Label
$fileListLabel.Text = "Files And Folders"
$fileListLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$fileListLabel.Location = New-Object System.Drawing.Point(20, 450)
$fileListLabel.Size = New-Object System.Drawing.Size(130, 22)
$form.Controls.Add($fileListLabel)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 475)
$grid.Size = New-Object System.Drawing.Size(660, 360)
$grid.ReadOnly = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoGenerateColumns = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.BackgroundColor = [System.Drawing.Color]::White
$checkColumn = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$checkColumn.Name = "Selected"
$checkColumn.HeaderText = "Use"
$checkColumn.FillWeight = 12
$grid.Columns.Add($checkColumn) | Out-Null
$null = $grid.Columns.Add("CurrentName", "Current Name")
$null = $grid.Columns.Add("ProposedName", "Previewed Changes")
$null = $grid.Columns.Add("Status", "Status")
$grid.Columns[1].FillWeight = 34
$grid.Columns[2].FillWeight = 36
$grid.Columns[3].FillWeight = 18
$grid.Columns[1].ReadOnly = $true
$grid.Columns[2].ReadOnly = $true
$grid.Columns[3].ReadOnly = $true
$form.Controls.Add($grid)

$detailsLabel = New-Object System.Windows.Forms.Label
$detailsLabel.Text = "Selected File"
$detailsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$detailsLabel.Location = New-Object System.Drawing.Point(705, 450)
$detailsLabel.Size = New-Object System.Drawing.Size(120, 22)
$form.Controls.Add($detailsLabel)

Add-FieldLabel -Text "Original Name" -X 705 -Y 480 -Width 120 | Out-Null
$originalNameBox = New-Object System.Windows.Forms.TextBox
$originalNameBox.Location = New-Object System.Drawing.Point(705, 505)
$originalNameBox.Size = New-Object System.Drawing.Size(370, 23)
$originalNameBox.ReadOnly = $true
$form.Controls.Add($originalNameBox)

Add-FieldLabel -Text "Changed Text" -X 705 -Y 540 -Width 120 | Out-Null
$changedTextBox = New-Object System.Windows.Forms.RichTextBox
$changedTextBox.Location = New-Object System.Drawing.Point(705, 565)
$changedTextBox.Size = New-Object System.Drawing.Size(370, 150)
$changedTextBox.ReadOnly = $true
$changedTextBox.BackColor = [System.Drawing.Color]::White
$changedTextBox.BorderStyle = "FixedSingle"
$changedTextBox.Font = New-Object System.Drawing.Font("Consolas", 11)
$form.Controls.Add($changedTextBox)

Add-FieldLabel -Text "Single File Rename" -X 705 -Y 730 -Width 140 | Out-Null
$newNameBox = New-Object System.Windows.Forms.TextBox
$newNameBox.Location = New-Object System.Drawing.Point(705, 755)
$newNameBox.Size = New-Object System.Drawing.Size(370, 23)
$form.Controls.Add($newNameBox)

$applyManualRenameButton = New-Object System.Windows.Forms.Button
$applyManualRenameButton.Text = "Apply Single File Rename"
$applyManualRenameButton.Location = New-Object System.Drawing.Point(705, 785)
$applyManualRenameButton.Size = New-Object System.Drawing.Size(185, 28)
$form.Controls.Add($applyManualRenameButton)

$undoManualRenameButton = New-Object System.Windows.Forms.Button
$undoManualRenameButton.Text = "Undo Single File Rename"
$undoManualRenameButton.Location = New-Object System.Drawing.Point(900, 785)
$undoManualRenameButton.Size = New-Object System.Drawing.Size(175, 28)
$undoManualRenameButton.Enabled = $false
$form.Controls.Add($undoManualRenameButton)

$changeSummary = New-Object System.Windows.Forms.Label
$changeSummary.Text = "Select a file to view the rename details."
$changeSummary.Location = New-Object System.Drawing.Point(705, 818)
$changeSummary.Size = New-Object System.Drawing.Size(370, 55)
$form.Controls.Add($changeSummary)

$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$script:lastPlan = @()
$script:lastUndoPlan = @()
$script:lastManualUndoPlan = @()
$script:manualOverrides = @{}
$script:isUpdatingDetailBoxes = $false
$script:selectedSources = New-Object System.Collections.Generic.HashSet[string]
$script:settingsPath = Join-Path $PSScriptRoot "EasyEdits_settings.json"

function Get-SavedFolderPath {
    if (-not (Test-Path -LiteralPath $script:settingsPath)) {
        return $null
    }

    try {
        $settings = Get-Content -LiteralPath $script:settingsPath -Raw | ConvertFrom-Json
        if ($settings.lastFolder -and (Test-Path -LiteralPath $settings.lastFolder)) {
            return $settings.lastFolder
        }
    } catch {
        return $null
    }

    return $null
}

function Save-FolderPath {
    param(
        [string]$FolderPath
    )

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        return
    }

    try {
        @{ lastFolder = $FolderPath } | ConvertTo-Json | Set-Content -LiteralPath $script:settingsPath
    } catch {
        $statusLabel.Text = "Could not save the last folder path."
    }
}

function Get-CurrentPlan {
    $plan = Get-RenamePlan -Folder $folderBox.Text -IncludeDirectories $true
    if (-not $plan) {
        return @()
    }

    foreach ($item in $plan) {
        if ($script:manualOverrides.ContainsKey($item.Source)) {
            $override = $script:manualOverrides[$item.Source]
            $item.TargetName = $override.TargetName
            $item.Target = Join-Path (Split-Path -Parent $item.Source) $override.TargetName
            $item.WillChange = $item.SourceName -ne $override.TargetName
            if ($item.PSObject.Properties["ChangeDetails"]) {
                $item.ChangeDetails = $override.ChangeDetails
            } else {
                $item | Add-Member -NotePropertyName ChangeDetails -NotePropertyValue $override.ChangeDetails
            }
        }
    }

    Test-RenamePlan -Plan $plan
    return $plan
}

function Get-SelectedPlan {
    param(
        [object[]]$Plan
    )

    return @($Plan | Where-Object { $script:selectedSources.Contains($_.Source) })
}

function Get-UndoPlanFromAppliedPlan {
    param(
        [object[]]$Plan
    )

    $undoPlan = New-Object System.Collections.Generic.List[object]

    for ($i = $Plan.Count - 1; $i -ge 0; $i--) {
        $item = $Plan[$i]
        if (-not $item.WillChange) {
            continue
        }

        $undoPlan.Add([PSCustomObject]@{
            Source = $item.Target
            SourceName = $item.TargetName
            Target = $item.Source
            TargetName = $item.SourceName
            WillChange = $true
            IsDirectory = $item.IsDirectory
            Extension = $item.Extension
        })
    }

    return $undoPlan.ToArray()
}

function Set-AllChecks {
    param(
        [bool]$Checked
    )

    foreach ($row in $grid.Rows) {
        $row.Cells[0].Value = $Checked
        if ($row.Tag) {
            if ($Checked) {
                [void]$script:selectedSources.Add($row.Tag.Source)
            } else {
                [void]$script:selectedSources.Remove($row.Tag.Source)
            }
        }
    }
}

function Get-ChangeParts {
    param(
        [string]$SourceName,
        [string]$TargetName
    )

    $prefixLength = 0
    $maxPrefix = [Math]::Min($SourceName.Length, $TargetName.Length)
    while ($prefixLength -lt $maxPrefix -and $SourceName[$prefixLength] -eq $TargetName[$prefixLength]) {
        $prefixLength++
    }

    $suffixLength = 0
    $maxSuffix = [Math]::Min($SourceName.Length - $prefixLength, $TargetName.Length - $prefixLength)
    while ($suffixLength -lt $maxSuffix) {
        $sourceIndex = $SourceName.Length - 1 - $suffixLength
        $targetIndex = $TargetName.Length - 1 - $suffixLength
        if ($SourceName[$sourceIndex] -ne $TargetName[$targetIndex]) {
            break
        }
        $suffixLength++
    }

    $removedLength = [Math]::Max(0, $SourceName.Length - $prefixLength - $suffixLength)
    $addedLength = [Math]::Max(0, $TargetName.Length - $prefixLength - $suffixLength)

    [PSCustomObject]@{
        Removed = if ($removedLength -gt 0) { $SourceName.Substring($prefixLength, $removedLength) } else { "" }
        Added = if ($addedLength -gt 0) { $TargetName.Substring($prefixLength, $addedLength) } else { "" }
    }
}

function Format-VisibleText {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    return $Text.Replace(" ", "[space]").Replace("`t", "[tab]")
}

function Format-FileSize {
    param(
        $Length
    )

    if ($null -eq $Length) {
        return "Folder"
    }

    if ($Length -lt 1KB) {
        return "$Length bytes"
    }

    if ($Length -lt 1MB) {
        return ("{0:N1} KB" -f ($Length / 1KB))
    }

    if ($Length -lt 1GB) {
        return ("{0:N1} MB" -f ($Length / 1MB))
    }

    return ("{0:N1} GB" -f ($Length / 1GB))
}

function Get-SeparatorValue {
    param(
        [string]$SelectedItem
    )

    switch ($SelectedItem) {
        "Dash (-)" { return "-" }
        "Dash ( - )" { return " - " }
        "Underscore (_)" { return "_" }
        "None" { return "" }
        default { return " " }
    }
}

function Get-ItemStem {
    param(
        [object]$Item,
        [bool]$UseTargetName = $true
    )

    $name = if ($UseTargetName) { $Item.TargetName } else { $Item.SourceName }
    if ($Item.PSObject.Properties["IsDirectory"] -and $Item.IsDirectory) {
        return $name
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($name)
}

function Show-SelectionDetails {
    param(
        [object]$Item
    )

    if (-not $Item) {
        $script:isUpdatingDetailBoxes = $true
        $originalNameBox.Text = ""
        $newNameBox.Text = ""
        $script:isUpdatingDetailBoxes = $false
        $changedTextBox.Clear()
        $changeSummary.Text = "Select a file to view the rename details."
        return
    }

    $script:isUpdatingDetailBoxes = $true
    $originalNameBox.Text = $Item.SourceName
    $newNameBox.Text = $Item.TargetName
    $script:isUpdatingDetailBoxes = $false

    if ($Item.PSObject.Properties["ChangeDetails"] -and $Item.ChangeDetails) {
        $parts = $Item.ChangeDetails
    } else {
        $parts = Get-ChangeParts -SourceName $Item.SourceName -TargetName $Item.TargetName
    }
    $changedTextBox.Clear()
    $changedTextBox.SelectionColor = [System.Drawing.Color]::Black
    $changedTextBox.AppendText("Removed:`r`n")

    if ([string]::IsNullOrEmpty($parts.Removed)) {
        $changedTextBox.SelectionColor = [System.Drawing.Color]::DarkGray
        $changedTextBox.AppendText("(nothing removed)`r`n`r`n")
    } else {
        $changedTextBox.SelectionBackColor = [System.Drawing.Color]::MistyRose
        $changedTextBox.SelectionColor = [System.Drawing.Color]::DarkRed
        $changedTextBox.AppendText((Format-VisibleText -Text $parts.Removed) + "`r`n`r`n")
    }

    $changedTextBox.SelectionBackColor = [System.Drawing.Color]::White
    $changedTextBox.SelectionColor = [System.Drawing.Color]::Black
    $changedTextBox.AppendText("Added:`r`n")

    if ([string]::IsNullOrEmpty($parts.Added)) {
        $changedTextBox.SelectionColor = [System.Drawing.Color]::DarkGray
        $changedTextBox.AppendText("(nothing added)")
    } else {
        $changedTextBox.SelectionBackColor = [System.Drawing.Color]::Honeydew
        $changedTextBox.SelectionColor = [System.Drawing.Color]::DarkGreen
        $changedTextBox.AppendText((Format-VisibleText -Text $parts.Added))
    }

    $changedTextBox.SelectionBackColor = [System.Drawing.Color]::White
    $changedTextBox.SelectionLength = 0

    if ($Item.WillChange) {
        $changeSummary.Text = "Edit the Full New Name box, then click Apply Manual Rename to rename this selected file."
    } else {
        $changeSummary.Text = ""
    }
}

function Update-Grid {
    param(
        [object[]]$Plan
    )

    $grid.Rows.Clear()

    foreach ($item in $Plan) {
        $status = if ($item.WillChange) { "Will rename" } else { "No change" }
        $isChecked = $script:selectedSources.Contains($item.Source)
        $rowIndex = $grid.Rows.Add($isChecked, $item.SourceName, $item.TargetName, $status)
        $grid.Rows[$rowIndex].Tag = $item
    }

    if ($grid.Rows.Count -gt 0) {
        $grid.ClearSelection()
        $grid.Rows[0].Selected = $true
        Show-SelectionDetails -Item $grid.Rows[0].Tag
    } else {
        Show-SelectionDetails -Item $null
    }
}

function Update-Status {
    if (-not $script:lastPlan -or $script:lastPlan.Count -eq 0) {
        $statusLabel.Text = "No files found in this folder."
        return
    }

    $selectedPlan = Get-SelectedPlan -Plan $script:lastPlan
    $renameCount = (@($selectedPlan | Where-Object WillChange)).Count
    $selectedCount = $selectedPlan.Count
    $statusLabel.Text = "Loaded $($script:lastPlan.Count) file(s). $selectedCount checked, $renameCount ready to rename."
}

function Refresh-Preview {
    try {
        if (-not [string]::IsNullOrWhiteSpace($folderBox.Text) -and (Test-Path -LiteralPath $folderBox.Text)) {
            Save-FolderPath -FolderPath $folderBox.Text
        }

        $script:lastPlan = Get-CurrentPlan
        Update-Grid -Plan $script:lastPlan
        Update-Status
        Update-UndoButtonState
    } catch {
        $grid.Rows.Clear()
        Show-SelectionDetails -Item $null
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Preview Failed")
    }
}

function Load-FolderIntoUi {
    param(
        [string]$TargetFolder,
        [string]$StatusMessage = "Loaded folder."
    )

    if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
        return
    }

    $folderBox.Text = $TargetFolder
    $dialog.SelectedPath = $TargetFolder
    Save-FolderPath -FolderPath $TargetFolder
    $script:lastUndoPlan = @()
    $script:lastManualUndoPlan = @()
    Update-UndoButtonState
    Update-ManualUndoButtonState
    Reset-UiState
    Refresh-Preview
    $statusLabel.Text = $StatusMessage
}

function Reset-UiState {
    $script:manualOverrides = @{}
    $script:selectedSources = New-Object System.Collections.Generic.HashSet[string]
    $script:lastPlan = @()
    $script:isUpdatingDetailBoxes = $true
    $renameAllBox.Text = ""
    $startNumberBox.Text = "1"
    $paddingDropdown.SelectedIndex = 0
    $separatorDropdown.SelectedIndex = 0
    $positionDropdown.SelectedIndex = 0
    $replaceFindTextBox.Text = ""
    $replaceWithTextBox.Text = ""
    $removeTextBox.Text = ""
    $newNameBox.Text = ""
    $script:isUpdatingDetailBoxes = $false
    $grid.Rows.Clear()
    Show-SelectionDetails -Item $null
}

function Apply-TextTransformToChecked {
    param(
        [scriptblock]$Transform,
        [string]$RemovedText = "",
        [string]$AddedText = "",
        [string]$SuccessMessage
    )

    try {
        if ($script:lastPlan.Count -eq 0) {
            $statusLabel.Text = "No files are loaded."
            return
        }

        $selectedPlan = Get-SelectedPlan -Plan $script:lastPlan
        if ($selectedPlan.Count -eq 0) {
            $statusLabel.Text = "Check at least one file first."
            return
        }

        foreach ($item in $selectedPlan) {
            $updatedName = & $Transform $item.SourceName
            $changeDetails = [PSCustomObject]@{
                Removed = $RemovedText
                Added = $AddedText
            }
            $item.TargetName = $updatedName
            $item.Target = Join-Path (Split-Path -Parent $item.Source) $updatedName
            $item.WillChange = $item.SourceName -ne $updatedName
            $script:manualOverrides[$item.Source] = @{
                TargetName = $updatedName
                ChangeDetails = $changeDetails
            }
            if ($item.PSObject.Properties["ChangeDetails"]) {
                $item.ChangeDetails = $changeDetails
            } else {
                $item | Add-Member -NotePropertyName ChangeDetails -NotePropertyValue $changeDetails
            }
        }

        Test-RenamePlan -Plan $script:lastPlan
        Update-Grid -Plan $script:lastPlan
        Update-Status
        $statusLabel.Text = $SuccessMessage
    } catch {
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Batch Edit Failed")
    }
}

function Insert-TextIntoName {
    param(
        [string]$Name,
        [string]$TextToAdd,
        [string]$Position,
        [string]$ReferenceText
    )

    switch ($Position) {
        "Beginning" { return "$TextToAdd$Name" }
        "End" { return "$Name$TextToAdd" }
        "Before reference text" {
            if ([string]::IsNullOrEmpty($ReferenceText)) {
                throw "Type reference text first."
            }

            $index = $Name.IndexOf($ReferenceText)
            if ($index -lt 0) {
                throw "Reference text '$ReferenceText' was not found in one or more checked file names."
            }

            return $Name.Insert($index, $TextToAdd)
        }
        "After reference text" {
            if ([string]::IsNullOrEmpty($ReferenceText)) {
                throw "Type reference text first."
            }

            $index = $Name.IndexOf($ReferenceText)
            if ($index -lt 0) {
                throw "Reference text '$ReferenceText' was not found in one or more checked file names."
            }

            return $Name.Insert($index + $ReferenceText.Length, $TextToAdd)
        }
        default { return "$Name$TextToAdd" }
    }
}

function Apply-RenameAllSequence {
    try {
        if ($script:lastPlan.Count -eq 0) {
            $statusLabel.Text = "No files are loaded."
            return
        }

        $selectedPlan = Get-SelectedPlan -Plan $script:lastPlan
        if ($selectedPlan.Count -eq 0) {
            $statusLabel.Text = "Check at least one file first."
            return
        }

        $baseName = $renameAllBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($baseName)) {
            $statusLabel.Text = "Type a base name into Rename All As first."
            return
        }

        $startNumber = 0
        if (-not [int]::TryParse($startNumberBox.Text, [ref]$startNumber)) {
            $statusLabel.Text = "Start Number must be a whole number."
            return
        }

        $paddingLength = switch ($paddingDropdown.SelectedItem) {
            "2 digits" { 2 }
            "3 digits" { 3 }
            "4 digits" { 4 }
            default { 0 }
        }

        $separator = Get-SeparatorValue -SelectedItem ([string]$separatorDropdown.SelectedItem)

        $numberPosition = [string]$positionDropdown.SelectedItem

        for ($i = 0; $i -lt $selectedPlan.Count; $i++) {
            $item = $selectedPlan[$i]
            $currentNumber = $startNumber + $i
            $numberText = if ($paddingLength -gt 0) {
                $currentNumber.ToString("D$paddingLength")
            } else {
                $currentNumber.ToString()
            }
            $newStem = switch ($numberPosition) {
                "Before name" {
                    "$numberText$separator$baseName"
                }
                "After name (no space)" {
                    $afterNameSeparator = if ($separator -eq " ") { "" } else { $separator }
                    "$baseName$afterNameSeparator$numberText"
                }
                default {
                    "$baseName$separator$numberText"
                }
            }
            $extension = if ($item.PSObject.Properties["Extension"]) { $item.Extension } else { [System.IO.Path]::GetExtension($item.SourceName) }
            $updatedName = "$newStem$extension"
            $oldStem = if ($item.PSObject.Properties["IsDirectory"] -and $item.IsDirectory) {
                $item.SourceName
            } else {
                [System.IO.Path]::GetFileNameWithoutExtension($item.SourceName)
            }

            $item.TargetName = $updatedName
            $item.Target = Join-Path (Split-Path -Parent $item.Source) $updatedName
            $item.WillChange = $item.SourceName -ne $updatedName

            $changeDetails = [PSCustomObject]@{
                Removed = $oldStem
                Added = $newStem
            }

            if ($item.PSObject.Properties["ChangeDetails"]) {
                $item.ChangeDetails = $changeDetails
            } else {
                $item | Add-Member -NotePropertyName ChangeDetails -NotePropertyValue $changeDetails
            }

            $script:manualOverrides[$item.Source] = @{
                TargetName = $updatedName
                ChangeDetails = $changeDetails
            }
        }

        Test-RenamePlan -Plan $script:lastPlan
        Update-Grid -Plan $script:lastPlan
        Update-Status
        $statusLabel.Text = "Prepared sequential names for checked files."
    } catch {
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Rename All Failed")
    }
}

$browseButton.Add_Click({
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Load-FolderIntoUi -TargetFolder $dialog.SelectedPath -StatusMessage "Loaded folder from browse."
    }
})

$refreshButton.Add_Click({
    Reset-UiState
    Refresh-Preview
    $statusLabel.Text = "Reset all staged changes and reloaded files from disk."
})

$dragEnterHandler = {
    param($sender, $e)

    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        if ($paths -and $paths.Count -gt 0) {
            $firstPath = [string]$paths[0]
            if ((Test-Path -LiteralPath $firstPath -PathType Container) -or (Test-Path -LiteralPath $firstPath -PathType Leaf)) {
                $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
                return
            }
        }
    }

    $e.Effect = [System.Windows.Forms.DragDropEffects]::None
}

$form.Add_DragEnter($dragEnterHandler)
$dropZonePanel.Add_DragEnter($dragEnterHandler)
$dropZoneLabel.Add_DragEnter($dragEnterHandler)

$dragDropHandler = {
    param($sender, $e)

    if (-not $e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        return
    }

    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    if (-not $paths -or $paths.Count -eq 0) {
        return
    }

    $firstPath = [string]$paths[0]
    $targetFolder = $null

    if (Test-Path -LiteralPath $firstPath -PathType Container) {
        $targetFolder = $firstPath
    } elseif (Test-Path -LiteralPath $firstPath -PathType Leaf) {
        $targetFolder = Split-Path -Parent $firstPath
    }

    if (-not [string]::IsNullOrWhiteSpace($targetFolder)) {
        Load-FolderIntoUi -TargetFolder $targetFolder -StatusMessage "Loaded folder from drag and drop."
    }
}

$form.Add_DragDrop($dragDropHandler)
$dropZonePanel.Add_DragDrop($dragDropHandler)
$dropZoneLabel.Add_DragDrop($dragDropHandler)

$folderBox.Add_Leave({
    Refresh-Preview
})

$toolTabs.Add_SelectedIndexChanged({
    Set-ApplyButtonLayout
})

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -gt 0) {
        Show-SelectionDetails -Item $grid.SelectedRows[0].Tag
    }
})

$grid.Add_CurrentCellDirtyStateChanged({
    if ($grid.IsCurrentCellDirty) {
        $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})

$grid.Add_CellValueChanged({
    param($sender, $e)

    if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) {
        return
    }

    $row = $grid.Rows[$e.RowIndex]
    if (-not $row.Tag) {
        return
    }

    if ($row.Cells[0].Value -eq $true) {
        [void]$script:selectedSources.Add($row.Tag.Source)
    } else {
        [void]$script:selectedSources.Remove($row.Tag.Source)
    }

    Update-Status
})

$newNameBox.Add_TextChanged({
    if ($script:isUpdatingDetailBoxes -or $grid.SelectedRows.Count -eq 0) {
        return
    }

    $selectedItem = $grid.SelectedRows[0].Tag
    if (-not $selectedItem) {
        return
    }

    $editedName = $newNameBox.Text
    if ([string]::IsNullOrEmpty($editedName)) {
        $statusLabel.Text = "The new file name cannot be blank."
        return
    }

    $selectedItem.TargetName = $editedName
    $selectedItem.Target = Join-Path (Split-Path -Parent $selectedItem.Source) $editedName
    $selectedItem.WillChange = $selectedItem.SourceName -ne $editedName
    if ($selectedItem.PSObject.Properties["ChangeDetails"]) {
        $selectedItem.ChangeDetails = $null
    } else {
        $selectedItem | Add-Member -NotePropertyName ChangeDetails -NotePropertyValue $null
    }
    $script:manualOverrides[$selectedItem.Source] = @{
        TargetName = $editedName
        ChangeDetails = $null
    }

    $selectedRow = $grid.SelectedRows[0]
    $selectedRow.Cells[2].Value = $selectedItem.TargetName
    $selectedRow.Cells[3].Value = if ($selectedItem.WillChange) { "Will rename" } else { "No change" }

    Show-SelectionDetails -Item $selectedItem
    Update-Status
})

$applyManualRenameButton.Add_Click({
    try {
        if ($grid.SelectedRows.Count -eq 0) {
            $statusLabel.Text = "Select a file first."
            return
        }

        $selectedItem = $grid.SelectedRows[0].Tag
        if (-not $selectedItem) {
            $statusLabel.Text = "Select a file first."
            return
        }

        $editedName = $newNameBox.Text
        if ([string]::IsNullOrWhiteSpace($editedName)) {
            $statusLabel.Text = "The new file name cannot be blank."
            return
        }

        $manualPlanItem = [PSCustomObject]@{
            Source = $selectedItem.Source
            SourceName = $selectedItem.SourceName
            Target = Join-Path (Split-Path -Parent $selectedItem.Source) $editedName
            TargetName = $editedName
            WillChange = $selectedItem.SourceName -ne $editedName
            IsDirectory = $selectedItem.IsDirectory
            Extension = $selectedItem.Extension
        }

        if (-not $manualPlanItem.WillChange) {
            $statusLabel.Text = "Nothing to rename for this file."
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Rename the selected file to '$editedName'?",
            "Confirm Manual Rename",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = "Manual rename canceled."
            return
        }

        $changedCount = Invoke-RenamePlan -Plan @($manualPlanItem) -Apply
        $script:lastManualUndoPlan = Get-UndoPlanFromAppliedPlan -Plan @($manualPlanItem)
        Update-ManualUndoButtonState
        $null = $script:manualOverrides.Remove($selectedItem.Source)
        [void]$script:selectedSources.Remove($selectedItem.Source)

        Refresh-Preview
        $statusLabel.Text = "Manually renamed $changedCount file(s)."
        [System.Windows.Forms.MessageBox]::Show("Manually renamed $changedCount file(s).", "Manual Rename Finished")
    } catch {
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Manual Rename Failed")
    }
})

$undoManualRenameButton.Add_Click({
    try {
        if (-not $script:lastManualUndoPlan -or $script:lastManualUndoPlan.Count -eq 0) {
            $statusLabel.Text = "Nothing to undo for manual rename."
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Undo the last manual rename?",
            "Confirm Manual Undo",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = "Manual undo canceled."
            return
        }

        $undoneCount = Invoke-RenamePlan -Plan $script:lastManualUndoPlan -Apply
        $script:lastManualUndoPlan = @()
        Update-ManualUndoButtonState

        Refresh-Preview
        $statusLabel.Text = "Undid $undoneCount manual rename(s)."
        [System.Windows.Forms.MessageBox]::Show("Undid $undoneCount manual rename(s).", "Manual Undo Finished")
    } catch {
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Manual Undo Failed")
    }
})

$removeSelectedButton.Add_Click({
    $removeText = $removeTextBox.Text
    if ([string]::IsNullOrEmpty($removeText)) {
        $statusLabel.Text = "Type text into Remove Text first."
        return
    }

    Apply-TextTransformToChecked -Transform {
        param($name)
        $name.Replace($removeText, "")
    } -RemovedText $removeText -AddedText "" -SuccessMessage "Removed text from checked files."
})

$replaceSelectedButton.Add_Click({
    $findText = $replaceFindTextBox.Text
    if ([string]::IsNullOrEmpty($findText)) {
        $statusLabel.Text = "Type text into Replace Text first."
        return
    }

    Apply-TextTransformToChecked -Transform {
        param($name)
        $name.Replace($findText, $replaceWithTextBox.Text)
    } -RemovedText $findText -AddedText $replaceWithTextBox.Text -SuccessMessage "Replaced text in checked files."
})

$addSelectedButton.Add_Click({
    $textToAdd = $addTextBox.Text
    if ([string]::IsNullOrEmpty($textToAdd)) {
        $statusLabel.Text = "Type text into Text To Add first."
        return
    }

    $positionChoice = [string]$addPositionDropdown.SelectedItem
    $referenceText = $addAnchorTextBox.Text

    Apply-TextTransformToChecked -Transform {
        param($name)
        Insert-TextIntoName -Name $name -TextToAdd $textToAdd -Position $positionChoice -ReferenceText $referenceText
    } -RemovedText "" -AddedText $textToAdd -SuccessMessage "Added text to checked files."
})

$renameAllButton.Add_Click({
    Apply-RenameAllSequence
})

$selectAllButton.Add_Click({
    Set-AllChecks -Checked $true
    Update-Status
})

$clearChecksButton.Add_Click({
    Set-AllChecks -Checked $false
    Update-Status
})

$undoButton.Add_Click({
    try {
        if (-not $script:lastUndoPlan -or $script:lastUndoPlan.Count -eq 0) {
            $statusLabel.Text = "Nothing to undo."
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Undo the last applied rename batch?",
            "Confirm Undo",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = "Undo canceled."
            return
        }

        $undoneCount = Invoke-RenamePlan -Plan $script:lastUndoPlan -Apply
        $script:lastUndoPlan = @()
        Update-UndoButtonState

        Refresh-Preview
        $statusLabel.Text = "Undid $undoneCount file(s)."
        [System.Windows.Forms.MessageBox]::Show("Undid $undoneCount file(s).", "Undo Finished")
    } catch {
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Undo Failed")
    }
})

$applyButton.Add_Click({
    try {
        $script:lastPlan = Get-CurrentPlan
        Update-Grid -Plan $script:lastPlan

        $selectedPlan = Get-SelectedPlan -Plan $script:lastPlan
        $renameCount = (@($selectedPlan | Where-Object WillChange)).Count
        if ($selectedPlan.Count -eq 0) {
            $statusLabel.Text = "Check at least one file first."
            return
        }

        if ($renameCount -eq 0) {
            $statusLabel.Text = "Nothing to rename."
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Rename $renameCount checked file(s)?",
            "Confirm Rename",
            [System.Windows.Forms.MessageBoxButtons]::OKCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -ne [System.Windows.Forms.DialogResult]::OK) {
            $statusLabel.Text = "Rename canceled."
            return
        }

        $changedCount = Invoke-RenamePlan -Plan $selectedPlan -Apply
        $script:lastUndoPlan = Get-UndoPlanFromAppliedPlan -Plan $selectedPlan
        Update-UndoButtonState
        foreach ($item in $selectedPlan) {
            $null = $script:manualOverrides.Remove($item.Source)
            [void]$script:selectedSources.Remove($item.Source)
        }

        Refresh-Preview
        $statusLabel.Text = "Renamed $changedCount file(s)."
        [System.Windows.Forms.MessageBox]::Show("Renamed $changedCount file(s).", "Finished")
    } catch {
        $statusLabel.Text = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Rename Failed")
    }
})

$form.Add_Shown({
    $savedFolder = Get-SavedFolderPath
    if ($savedFolder) {
        $folderBox.Text = $savedFolder
        $dialog.SelectedPath = $savedFolder
    }

    Refresh-Preview
})

[void]$form.ShowDialog()
