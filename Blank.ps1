# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(1200, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create a splitter for dynamic preview panel
$splitter = New-Object System.Windows.Forms.Splitter
$splitter.Dock = [System.Windows.Forms.DockStyle]::Right
$splitter.Width = 5
$splitter.Visible = $false
$form.Controls.Add($splitter)

# Create Preview Panel with improved visibility control
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Size = New-Object System.Drawing.Size(390, 510)
$previewPanel.Location = New-Object System.Drawing.Point(780, 50)
$previewPanel.BackColor = [System.Drawing.Color]::White
$previewPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                       [System.Windows.Forms.AnchorStyles]::Right -bor `
                       [System.Windows.Forms.AnchorStyles]::Bottom
$previewPanel.Visible = $false
$form.Controls.Add($previewPanel)

# Create Preview Header
$previewHeader = New-Object System.Windows.Forms.Panel
$previewHeader.Height = 30
$previewHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$previewHeader.BackColor = [System.Drawing.Color]::WhiteSmoke
$previewPanel.Controls.Add($previewHeader)

# Create Close Button for Preview Panel
$closePreviewButton = New-Object System.Windows.Forms.Button
$closePreviewButton.Text = "×"
$closePreviewButton.Size = New-Object System.Drawing.Size(30, 30)
$closePreviewButton.Dock = [System.Windows.Forms.DockStyle]::Right
$closePreviewButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$closePreviewButton.Add_Click({
    Hide-PreviewPanel
})
$previewHeader.Controls.Add($closePreviewButton)

# Create Preview Title Label
$previewTitle = New-Object System.Windows.Forms.Label
$previewTitle.Text = "File Preview"
$previewTitle.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$previewTitle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$previewHeader.Controls.Add($previewTitle)

# Create Preview Content Panel
$previewContent = New-Object System.Windows.Forms.Panel
$previewContent.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewPanel.Controls.Add($previewContent)

# Create various preview controls
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.Visible = $false
$previewContent.Controls.Add($pictureBox)

$textPreview = New-Object System.Windows.Forms.RichTextBox
$textPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$textPreview.ReadOnly = $true
$textPreview.Font = New-Object System.Drawing.Font("Consolas", 10)
$textPreview.Visible = $false
$previewContent.Controls.Add($textPreview)

$webBrowser = New-Object System.Windows.Forms.WebBrowser
$webBrowser.Dock = [System.Windows.Forms.DockStyle]::Fill
$webBrowser.Visible = $false
$previewContent.Controls.Add($webBrowser)

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$previewLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$previewContent.Controls.Add($previewLabel)

# Create MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.Controls.Add($menuStrip)

# Initialize navigation history
$global:navigationHistory = New-Object System.Collections.ArrayList
$global:currentIndex = -1

# File Menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

$newWindow = New-Object System.Windows.Forms.ToolStripMenuItem
$newWindow.Text = "New Window"
$newWindow.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N
$newWindow.Add_Click({
    Start-Process powershell -ArgumentList "-File `"$PSCommandPath`""
})

$exit = New-Object System.Windows.Forms.ToolStripMenuItem
$exit.Text = "Exit"
$exit.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$exit.Add_Click({ $form.Close() })

$fileMenu.DropDownItems.AddRange(@($newWindow, $exit))

# Edit Menu
$editMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$editMenu.Text = "Edit"

$copy = New-Object System.Windows.Forms.ToolStripMenuItem
$copy.Text = "Copy"
$copy.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::C
$copy.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $paths = $listView.SelectedItems | ForEach-Object { $_.Tag }
        [System.Windows.Forms.Clipboard]::SetText(($paths -join "`r`n"))
    }
})

$paste = New-Object System.Windows.Forms.ToolStripMenuItem
$paste.Text = "Paste"
$paste.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::V
$paste.Add_Click({
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $paths = [System.Windows.Forms.Clipboard]::GetText() -split "`r`n"
        foreach ($path in $paths) {
            if (Test-Path $path) {
                $destination = Join-Path $global:currentPath (Split-Path $path -Leaf)
                Copy-Item -Path $path -Destination $destination -Recurse
            }
        }
        Populate-ListView $global:currentPath
    }
})

$delete = New-Object System.Windows.Forms.ToolStripMenuItem
$delete.Text = "Delete"
$delete.ShortcutKeys = [System.Windows.Forms.Keys]::Delete
$delete.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete the selected items?",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $listView.SelectedItems | ForEach-Object {
                Remove-Item $_.Tag -Recurse -Force
            }
            Populate-ListView $global:currentPath
        }
    }
})

$editMenu.DropDownItems.AddRange(@($copy, $paste, $delete))

# View Menu
$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "View"

$refresh = New-Object System.Windows.Forms.ToolStripMenuItem
$refresh.Text = "Refresh"
$refresh.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$refresh.Add_Click({
    Populate-ListView $global:currentPath
    Populate-TreeView
})

$viewMenu.DropDownItems.Add($refresh)

# Create Navigation Buttons
$btnBack = New-Object System.Windows.Forms.ToolStripMenuItem
$btnBack.Text = "←"
$btnBack.Enabled = $false
$btnBack.Add_Click({
    if ($global:currentIndex -gt 0) {
        $global:currentIndex--
        $previousPath = $global:navigationHistory[$global:currentIndex]
        Populate-ListView $previousPath
        Update-NavigationButtons
    }
})

$btnForward = New-Object System.Windows.Forms.ToolStripMenuItem
$btnForward.Text = "→"
$btnForward.Enabled = $false
$btnForward.Add_Click({
    if ($global:currentIndex -lt $global:navigationHistory.Count - 1) {
        $global:currentIndex++
        $nextPath = $global:navigationHistory[$global:currentIndex]
        Populate-ListView $nextPath
        Update-NavigationButtons
    }
})

$btnUp = New-Object System.Windows.Forms.ToolStripMenuItem
$btnUp.Text = "↑"
$btnUp.Add_Click({
    $parentPath = Split-Path $global:currentPath -Parent
    if ($parentPath) {
        Navigate-To $parentPath
    }
})

$btnDown = New-Object System.Windows.Forms.ToolStripMenuItem
$btnDown.Text = "↓"
$btnDown.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $selectedItem = $listView.SelectedItems[0]
        if ($selectedItem -and (Test-Path -Path $selectedItem.Tag -PathType Container)) {
            Navigate-To $selectedItem.Tag
        }
    }
})

# Add all items to MenuStrip in order
$menuStrip.Items.AddRange(@($fileMenu, $editMenu, $viewMenu, $btnBack, $btnForward, $btnUp, $btnDown))

# Create search box in MenuStrip
$searchBox = New-Object System.Windows.Forms.ToolStripTextBox
$searchBox.Size = New-Object System.Drawing.Size(200, 25)
$searchBox.Name = "SearchBox"
$searchBox.PlaceholderText = "Search..."

# Create search button in MenuStrip
$searchButton = New-Object System.Windows.Forms.ToolStripButton
$searchButton.Text = "Search"
$searchButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::Text

# Add spring to push search controls to right
$spring = New-Object System.Windows.Forms.ToolStripStatusLabel
$spring.Spring = $true

# Add search controls to MenuStrip
$menuStrip.Items.Add($spring)
$menuStrip.Items.Add($searchBox)
$menuStrip.Items.Add($searchButton)

# Create Quick Access panel
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Size = New-Object System.Drawing.Size(250, 220)
$quickAccessPanel.Location = New-Object System.Drawing.Point(10, 50)
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoSize = $false
$quickAccessPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                           [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($quickAccessPanel)

# Create TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 290)
$treeView.Location = New-Object System.Drawing.Point(10, 270)
$treeView.Scrollable = $true
$treeView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                   [System.Windows.Forms.AnchorStyles]::Left -bor `
                   [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($treeView)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(500, 510)
$listView.Location = New-Object System.Drawing.Point(270, 50)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                   [System.Windows.Forms.AnchorStyles]::Left -bor `
                   [System.Windows.Forms.AnchorStyles]::Right -bor `
                   [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($listView)

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

#search files function
function Search-Files {
    param ([string]$searchTerm)
    
    if ([string]::IsNullOrWhiteSpace($searchTerm)) {
        Populate-ListView $global:currentPath
        return
    }
    
    $listView.Items.Clear()
    
    try {
        $searchResults = Get-ChildItem -Path $global:currentPath -Recurse -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -like "*$searchTerm*" }
        
        foreach ($item in $searchResults) {
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
            "Error performing search: $($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

#search button click event
$searchButton.Add_Click({
    Search-Files $searchBox.Text
})

#search box key
$searchBox.Add_KeyPress({
    param($sender, $e)
    if ($e.KeyChar -eq [System.Windows.Forms.Keys]::Enter) {
        $e.Handled = $true
        Search-Files $searchBox.Text
    }
})

# Function to format file size
function Format-FileSize {
    param ([long]$size)
    if ($size -lt 1KB) { return "$size B" }
    elseif ($size -lt 1MB) { return "{0:N2} KB" -f ($size/1KB) }
    elseif ($size -lt 1GB) { return "{0:N2} MB" -f ($size/1MB) }
    elseif ($size -lt 1TB) { return "{0:N2} GB" -f ($size/1GB) }
    else { return "{0:N2} TB" -f ($size/1TB) }
}

# Function to update navigation buttons
function Update-NavigationButtons {
    $btnBack.Enabled = $global:currentIndex -gt 0
    $btnForward.Enabled = $global:currentIndex -lt ($global:navigationHistory.Count - 1)
    $btnUp.Enabled = (Split-Path $global:currentPath -Parent) -ne $null
    $btnDown.Enabled = ($listView.SelectedItems.Count -gt 0) -and 
                      (Test-Path -Path $listView.SelectedItems[0].Tag -PathType Container)
}

# Function to handle navigation
function Navigate-To {
    param ([string]$path)
    
    if (Test-Path $path) {
        $global:currentIndex++
        if ($global:currentIndex -lt $global:navigationHistory.Count) {
            $global:navigationHistory.RemoveRange($global:currentIndex, $global:navigationHistory.Count - $global:currentIndex)
        }
        [void]$global:navigationHistory.Add($path)
        Populate-ListView $path
        Update-NavigationButtons
    }
}

# Function to show preview panel
function Show-PreviewPanel {
    if (-not $previewPanel.Visible) {
        $previewPanel.Visible = $true
        $splitter.Visible = $true
        $listView.Width -= ($previewPanel.Width + $splitter.Width)
    }
}

# Function to hide preview panel
function Hide-PreviewPanel {
    if ($previewPanel.Visible) {
        $listView.Width += ($previewPanel.Width + $splitter.Width)
        $previewPanel.Visible = $false
        $splitter.Visible = $false
    }
}

# Function to clear preview
function Clear-Preview {
    $pictureBox.Image = $null
    $pictureBox.Visible = $false
    $textPreview.Clear()
    $textPreview.Visible = $false
    $webBrowser.Visible = $false
    $previewLabel.Visible = $true
    $previewLabel.Text = "Select a file to preview"
}

# Function to preview file with enhanced format support
function Show-FilePreview {
    param ([string]$filePath)
    
    Clear-Preview
    Show-PreviewPanel
    
    if (-not (Test-Path $filePath)) {
        return
    }
    
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    $previewTitle.Text = $fileName
    $previewLabel.Visible = $false
    
    switch -Regex ($extension) {
        # Image files
        '\.(jpg|jpeg|png|gif|bmp|ico|tiff)$' {
            try {
                $image = [System.Drawing.Image]::FromFile($filePath)
                $pictureBox.Image = $image
                $pictureBox.Visible = $true
                $previewLabel.Text = "Size: $($image.Width)x$($image.Height)"
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error loading image"
                $previewLabel.Visible = $true
            }
        }
        
        # Text files
        '\.(txt|log|ps1|cmd|bat|csv|json|xml|html|css|js|md|yml|yaml|ini|conf|cfg|reg)$' {
            try {
                $content = Get-Content -Path $filePath -Raw -ErrorAction Stop
                $textPreview.Text = $content
                $textPreview.Visible = $true
                
                # Syntax highlighting based on extension
                switch -Regex ($extension) {
                    '\.(ps1|cmd|bat)$' {
                        # PowerShell/Batch highlighting (basic)
                        $keywords = @('function', 'param', 'if', 'else', 'while', 'foreach', 'return', 'try', 'catch')
                        foreach ($keyword in $keywords) {
                            $textPreview.SelectionColor = [System.Drawing.Color]::Blue
                        }
                    }
                    '\.(json|xml|html|css)$' {
                        # Web format highlighting (basic)
                        $webBrowser.Navigate($filePath)
                        $webBrowser.Visible = $true
                        $textPreview.Visible = $false
                    }
                }
            }
            catch {
                $previewLabel.Text = "Error loading text file"
                $previewLabel.Visible = $true
            }
        }
        
        # Office documents and PDFs
        '\.(doc|docx|xls|xlsx|ppt|pptx|pdf)$' {
            $fileInfo = Get-Item $filePath
            $previewLabel.Text = @"
File Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $fileInfo.Length)
Created: $($fileInfo.CreationTime)
Modified: $($fileInfo.LastWriteTime)
"@
            $previewLabel.Visible = $true
        }
        
        # Audio files
        '\.(mp3|wav|wma|m4a|aac)$' {
            $previewLabel.Text = @"
Audio File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Video files
        '\.(mp4|avi|mkv|wmv|mov)$' {
            $previewLabel.Text = @"
Video File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Archive files
        '\.(zip|rar|7z|tar|gz)$' {
            try {
                $archive = Get-Item $filePath
                $previewLabel.Text = @"
Archive File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $archive.Length)
Created: $($archive.CreationTime)
Modified: $($archive.LastWriteTime)
"@
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error reading archive"
                $previewLabel.Visible = $true
            }
        }
        
        default {
            $previewLabel.Text = "Preview not available for this file type"
            $previewLabel.Visible = $true
        }
    }
}

# Function to populate the ListView
function Populate-ListView {
    param ([string]$path)
    
    $global:currentPath = $path
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
        
        Update-NavigationButtons
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
            Navigate-To -path $buttonPath
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
    return $button
}

# Add Quick Access buttons
$quickAccessButtons = @{
    "Desktop" = [Environment]::GetFolderPath("Desktop")
    "Downloads" = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    "Documents" = [Environment]::GetFolderPath("MyDocuments")
    "Music" = [Environment]::GetFolderPath("MyMusic")
    "Pictures" = [Environment]::GetFolderPath("MyPictures")
    "Videos" = [Environment]::GetFolderPath("MyVideos")
}

foreach ($button in $quickAccessButtons.GetEnumerator()) {
    Add-QuickAccessButton $button.Key $button.Value
}




# Function to populate the TreeView
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

# Event handler for TreeView node click
$treeView.add_AfterSelect({
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode.Tag) {
        Navigate-To -path $selectedNode.Tag
    }
})

# Event handler for ListView double-click
$listView.add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $itemPath = $selectedItem.Tag
        
        if (Test-Path -Path $itemPath -PathType Container) {
            Navigate-To -path $itemPath
        } else {
            Start-Process $itemPath
        }
    }
})

# Create Preview Panel
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Size = New-Object System.Drawing.Size(390, 510)
$previewPanel.Location = New-Object System.Drawing.Point(780, 50)
$previewPanel.BackColor = [System.Drawing.Color]::White
$previewPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                       [System.Windows.Forms.AnchorStyles]::Right -bor `
                       [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($previewPanel)

# Create Preview Controls
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Size = New-Object System.Drawing.Size(380, 380)
$pictureBox.Location = New-Object System.Drawing.Point(5, 5)
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.Visible = $false
$previewPanel.Controls.Add($pictureBox)

$textPreview = New-Object System.Windows.Forms.RichTextBox
$textPreview.Size = New-Object System.Drawing.Size(380, 480)
$textPreview.Location = New-Object System.Drawing.Point(5, 5)
$textPreview.ReadOnly = $true
$textPreview.Font = New-Object System.Drawing.Font("Consolas", 10)
$textPreview.Visible = $false
$previewPanel.Controls.Add($textPreview)

$mediaPlayer = New-Object System.Windows.Forms.Panel
$mediaPlayer.Size = New-Object System.Drawing.Size(380, 380)
$mediaPlayer.Location = New-Object System.Drawing.Point(5, 5)
$mediaPlayer.Visible = $false
$previewPanel.Controls.Add($mediaPlayer)

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Size = New-Object System.Drawing.Size(380, 40)
$previewLabel.Location = New-Object System.Drawing.Point(5, 5)
$previewLabel.Text = "Select a file to preview"
$previewLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$previewLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12)
$previewPanel.Controls.Add($previewLabel)

# Function to clear preview
function Clear-Preview {
    $pictureBox.Image = $null
    $pictureBox.Visible = $false
    $textPreview.Clear()
    $textPreview.Visible = $false
    $mediaPlayer.Visible = $false
    $previewLabel.Visible = $true
    $previewLabel.Text = "Select a file to preview"
}

function Show-FilePreview {
    param ([string]$filePath)
    
    Clear-Preview
    Show-PreviewPanel
    
    if (-not (Test-Path $filePath)) {
        return
    }
    
    $fileName = [System.IO.Path]::GetFileName($filePath)
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    $previewTitle.Text = $fileName
    $previewLabel.Visible = $false
    
    # Add folder preview at the beginning
    if (Test-Path -Path $filePath -PathType Container) {
        try {
            $folder = Get-Item $filePath
            $items = Get-ChildItem $filePath
            $fileCount = ($items | Where-Object { -not $_.PSIsContainer }).Count
            $folderCount = ($items | Where-Object { $_.PSIsContainer }).Count
            
            $previewLabel.Text = @"
Folder: $($folder.Name)
Created: $($folder.CreationTime)
Modified: $($folder.LastWriteTime)
Contains: $fileCount files, $folderCount folders
"@
            $previewLabel.Visible = $true
            return
        }
        catch {
            $previewLabel.Text = "Error reading folder"
            $previewLabel.Visible = $true
            return
        }
    }
    
    switch -Regex ($extension) {
        # Image files
        '\.(jpg|jpeg|png|gif|bmp|ico|tiff)$' {
            try {
                $image = [System.Drawing.Image]::FromFile($filePath)
                $pictureBox.Image = $image
                $pictureBox.Visible = $true
                $previewLabel.Text = "Size: $($image.Width)x$($image.Height)"
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error loading image"
                $previewLabel.Visible = $true
            }
        }
        
        # Text files
        '\.(txt|log|ps1|cmd|bat|csv|json|xml|html|css|js|md|yml|yaml|ini|conf|cfg|reg)$' {
            try {
                $content = Get-Content -Path $filePath -Raw -ErrorAction Stop
                $textPreview.Text = $content
                $textPreview.Visible = $true
                
                # Syntax highlighting based on extension
                switch -Regex ($extension) {
                    '\.(ps1|cmd|bat)$' {
                        # PowerShell/Batch highlighting (basic)
                        $keywords = @('function', 'param', 'if', 'else', 'while', 'foreach', 'return', 'try', 'catch')
                        foreach ($keyword in $keywords) {
                            $textPreview.SelectionColor = [System.Drawing.Color]::Blue
                        }
                    }
                    '\.(json|xml|html|css)$' {
                        # Web format highlighting (basic)
                        $webBrowser.Navigate($filePath)
                        $webBrowser.Visible = $true
                        $textPreview.Visible = $false
                    }
                }
            }
            catch {
                $previewLabel.Text = "Error loading text file"
                $previewLabel.Visible = $true
            }
        }
        
        # Office documents and PDFs
        '\.(doc|docx|xls|xlsx|ppt|pptx|pdf)$' {
            $fileInfo = Get-Item $filePath
            $previewLabel.Text = @"
File Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $fileInfo.Length)
Created: $($fileInfo.CreationTime)
Modified: $($fileInfo.LastWriteTime)
"@
            $previewLabel.Visible = $true
        }
        
        # Audio files
        '\.(mp3|wav|wma|m4a|aac)$' {
            $previewLabel.Text = @"
Audio File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Video files
        '\.(mp4|avi|mkv|wmv|mov)$' {
            $previewLabel.Text = @"
Video File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize (Get-Item $filePath).Length)
Double-click to play in default player
"@
            $previewLabel.Visible = $true
        }
        
        # Archive files
        '\.(zip|rar|7z|tar|gz)$' {
            try {
                $archive = Get-Item $filePath
                $previewLabel.Text = @"
Archive File
Type: $($extension.TrimStart('.').ToUpper())
Size: $(Format-FileSize $archive.Length)
Created: $($archive.CreationTime)
Modified: $($archive.LastWriteTime)
"@
                $previewLabel.Visible = $true
            }
            catch {
                $previewLabel.Text = "Error reading archive"
                $previewLabel.Visible = $true
            }
        }
        
        default {
            $previewLabel.Text = "Preview not available for this file type"
            $previewLabel.Visible = $true
        }
    }
}

# Event handler for ListView selection changed
$listView.add_SelectedIndexChanged({
    Update-NavigationButtons
})

# Event handler for ListView click
$listView.add_MouseClick({
    param($sender, $e)
    
    $item = $listView.GetItemAt($e.X, $e.Y)
    if ($item -ne $null) {
        $itemPath = $item.Tag
        Show-FilePreview -filePath $itemPath
    }
})

# Event handler for ListView selection changed
$listView.add_SelectedIndexChanged({
    Update-NavigationButtons
})

# Event handler for ListView click
$listView.add_MouseClick({
    param($sender, $e)
    
    $item = $listView.GetItemAt($e.X, $e.Y)
    if ($item -ne $null) {
        $itemPath = $item.Tag
        Show-FilePreview -filePath $itemPath
    }
})

# Event handler for ListView double-click
$listView.add_DoubleClick({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $itemPath = $selectedItem.Tag
        
        if (Test-Path -Path $itemPath -PathType Container) {
            Navigate-To -path $itemPath
            # Keep preview visible when navigating folders
        } else {
            Start-Process $itemPath
        }
    }
})

# Initialize navigation with desktop path
$desktopPath = [Environment]::GetFolderPath("Desktop")
[void]$global:navigationHistory.Add($desktopPath)
$global:currentIndex = 0

# Initial setup
Populate-TreeView
Populate-ListView $desktopPath

# Show the form
[void]$form.ShowDialog()