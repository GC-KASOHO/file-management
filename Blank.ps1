# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer with Preview"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create main SplitContainer for left and right panels
$mainSplitContainer = New-Object System.Windows.Forms.SplitContainer
$mainSplitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainSplitContainer.SplitterDistance = 900
$form.Controls.Add($mainSplitContainer)

# Left Panel Container
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainSplitContainer.Panel1.Controls.Add($leftPanel)

# Create Quick Access panel
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Size = New-Object System.Drawing.Size(250, 220)
$quickAccessPanel.Location = New-Object System.Drawing.Point(10, 10)
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoSize = $false
$leftPanel.Controls.Add($quickAccessPanel)

# Create TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Location = New-Object System.Drawing.Point(10, 240)
$treeView.Size = New-Object System.Drawing.Size(250, 510)
$treeView.Scrollable = $true
$leftPanel.Controls.Add($treeView)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(270, 10)
$listView.Size = New-Object System.Drawing.Size(620, 740)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$leftPanel.Controls.Add($listView)

# Create Preview Panel (Right Side)
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewPanel.BackColor = [System.Drawing.Color]::White
$mainSplitContainer.Panel2.Controls.Add($previewPanel)

# Create Preview Controls
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$pictureBox.Visible = $false
$previewPanel.Controls.Add($pictureBox)

$textPreview = New-Object System.Windows.Forms.RichTextBox
$textPreview.ReadOnly = $true
$textPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$textPreview.Visible = $false
$textPreview.Font = New-Object System.Drawing.Font("Consolas", 10)
$previewPanel.Controls.Add($textPreview)

$noPreviewLabel = New-Object System.Windows.Forms.Label
$noPreviewLabel.Text = "No preview available"
$noPreviewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$noPreviewLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$noPreviewLabel.Visible = $true
$previewPanel.Controls.Add($noPreviewLabel)

# Add columns to ListView
$columns = @(
    @{Name="Name"; Width=250},
    @{Name="Type"; Width=100},
    @{Name="Size"; Width=100},
    @{Name="Modified"; Width=150}
)

foreach ($column in $columns) {
    $listView.Columns.Add($column.Name, $column.Width)
}

# Function to format file size
function Format-FileSize {
    param ([long]$size)
    if ($size -lt 1KB) { return "$size B" }
    elseif ($size -lt 1MB) { return "{0:N2} KB" -f ($size/1KB) }
    elseif ($size -lt 1GB) { return "{0:N2} MB" -f ($size/1MB) }
    elseif ($size -lt 1TB) { return "{0:N2} GB" -f ($size/1GB) }
    else { return "{0:N2} TB" -f ($size/1TB) }
}

# Function to show preview
function Show-Preview {
    param ([string]$filePath)
    
    # Reset all preview controls
    $pictureBox.Visible = $false
    $textPreview.Visible = $false
    $noPreviewLabel.Visible = $false
    
    if (-not (Test-Path $filePath)) {
        $noPreviewLabel.Visible = $true
        return
    }
    
    try {
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        switch -Regex ($extension) {
            # Image files
            '\.(jpg|jpeg|png|gif|bmp)$' {
                $image = [System.Drawing.Image]::FromFile($filePath)
                $pictureBox.Image = $image
                $pictureBox.Visible = $true
            }
            
            # Text files
            '\.(txt|log|xml|json|ps1|cmd|bat|cfg|ini|csv|html|htm)$' {
                $textPreview.Text = Get-Content -Path $filePath -Raw -ErrorAction Stop
                $textPreview.Visible = $true
            }
            
            # Other files
            default {
                $item = Get-Item $filePath
                $info = "File Properties:`r`n`r`n"
                $info += "Name: $($item.Name)`r`n"
                $info += "Type: $($item.Extension.TrimStart('.'))`r`n"
                $info += "Size: $(Format-FileSize $item.Length)`r`n"
                $info += "Created: $($item.CreationTime)`r`n"
                $info += "Modified: $($item.LastWriteTime)`r`n"
                $info += "Attributes: $($item.Attributes)`r`n"
                
                $textPreview.Text = $info
                $textPreview.Visible = $true
            }
        }
    }
    catch {
        $noPreviewLabel.Text = "Error loading preview: $($_.Exception.Message)"
        $noPreviewLabel.Visible = $true
    }
    finally {
        [System.GC]::Collect()
    }
}

# Function to populate ListView
function Populate-ListView {
    param ([string]$path)
    
    $listView.Items.Clear()
    
    try {
        $items = Get-ChildItem -Path $path -ErrorAction Stop
        
        foreach ($item in $items) {
            $listViewItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
            
            if ($item.PSIsContainer) {
                $type = "Folder"
                $size = ""
            } else {
                $type = if ($item.Extension) { $item.Extension.TrimStart(".").ToUpper() } else { "File" }
                $size = Format-FileSize $item.Length
            }
            
            $listViewItem.SubItems.Add($type)
            $listViewItem.SubItems.Add($size)
            $listViewItem.SubItems.Add($item.LastWriteTime.ToString("g"))
            $listViewItem.Tag = $item.FullName
            
            $listView.Items.Add($listViewItem)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error accessing path: $path`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to create Quick Access buttons
function Add-QuickAccessButton($text, $path) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Width = 240
    $button.Height = 30
    $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.Tag = $path
    
    $button.Add_Click({
        $buttonPath = $this.Tag
        if (Test-Path -Path $buttonPath) {
            Populate-ListView -path $buttonPath
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "Path does not exist: $buttonPath",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    
    $quickAccessPanel.Controls.Add($button)
}

# Add Quick Access buttons
$quickAccessButtons = @{
    "Desktop" = [Environment]::GetFolderPath("Desktop")
    "Downloads" = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    "Documents" = [Environment]::GetFolderPath("MyDocuments")
    "Pictures" = [Environment]::GetFolderPath("MyPictures")
    "Music" = [Environment]::GetFolderPath("MyMusic")
    "Videos" = [Environment]::GetFolderPath("MyVideos")
}

foreach ($button in $quickAccessButtons.GetEnumerator()) {
    Add-QuickAccessButton $button.Key $button.Value
}

# Function to populate TreeView
function Populate-TreeView {
    $treeView.Nodes.Clear()
    $thisPC = $treeView.Nodes.Add("This PC")
    
    Get-PSDrive -PSProvider FileSystem | ForEach-Object {
        $driveNode = $thisPC.Nodes.Add($_.Root)
        $driveNode.Tag = $_.Root
        try {
            Get-ChildItem -Path $_.Root -Directory -ErrorAction Stop | ForEach-Object {
                $subNode = $driveNode.Nodes.Add($_.Name)
                $subNode.Tag = $_.FullName
            }
        } catch {}
    }
    
    $thisPC.Expand()
}

# Event handlers
$treeView.add_AfterSelect({
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode.Tag) {
        Populate-ListView -path $selectedNode.Tag
    }
})

$listView.add_SelectedIndexChanged({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        $itemPath = $selectedItem.Tag
        
        if (-not (Test-Path -Path $itemPath -PathType Container)) {
            Show-Preview -filePath $itemPath
        } else {
            $noPreviewLabel.Text = "Folder: $($selectedItem.Text)`n`nContains: $(Get-ChildItem $itemPath | Measure-Object | Select-Object -ExpandProperty Count) items"
            $noPreviewLabel.Visible = $true
        }
    }
})

$listView.add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $itemPath = $selectedItem.Tag
        
        if (Test-Path -Path $itemPath -PathType Container) {
            Populate-ListView -path $itemPath
        } else {
            Start-Process $itemPath
        }
    }
})

# Initial population
Populate-TreeView
$desktopPath = [Environment]::GetFolderPath("Desktop")
Populate-ListView -path $desktopPath

# Show the form
[void]$form.ShowDialog()