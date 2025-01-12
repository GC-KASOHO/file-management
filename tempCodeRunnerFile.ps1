# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.Controls.Add($menuStrip)

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

# Add menus to MenuStrip
$menuStrip.Items.AddRange(@($fileMenu, $editMenu, $viewMenu))

# Create Quick Access panel with increased height (adjusted for MenuStrip)
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Size = New-Object System.Drawing.Size(250, 220)
$quickAccessPanel.Location = New-Object System.Drawing.Point(10, 50)  # Adjusted Y position
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoSize = $false
$form.Controls.Add($quickAccessPanel)

# Create a TreeView to display directories (left side)
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 290)
$treeView.Location = New-Object System.Drawing.Point(10, 270)  # Adjusted Y position
$treeView.Scrollable = $true
$form.Controls.Add($treeView)

# Create a ListView to display files (right side)
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(600, 510)
$listView.Location = New-Object System.Drawing.Point(270, 50)  # Adjusted Y position
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$form.Controls.Add($listView)

# Rest of your original code remains exactly the same from here
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

# Function to populate the ListView
function Populate-ListView {
    param ([string]$path)
    
    $global:currentPath = $path
    $listView.Items.Clear()
    
    try {
        # Get all items in the directory
        $items = Get-ChildItem -Path $path -ErrorAction Stop
        
        foreach ($item in $items) {
            $listViewItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
            
            # Set item type and size
            if ($item.PSIsContainer) {
                $type = "Folder"
                $size = ""
            } else {
                $type = if ($item.Extension) { $item.Extension.TrimStart(".").ToUpper() } else { "File" }
                $size = Format-FileSize $item.Length
            }
            
            # Add subitems
            $listViewItem.SubItems.Add($type)
            $listViewItem.SubItems.Add($size)
            $listViewItem.SubItems.Add($item.LastWriteTime.ToString("g"))
            
            # Store the full path in the Tag property
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
    
    # Store the path in the button's Tag property
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
    return $button
}

# Add Quick Access buttons with paths using environment variables
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
    
    # Add "This PC" node
    $thisPC = $treeView.Nodes.Add("This PC")
    
    # Get all drives
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
        Populate-ListView -path $selectedNode.Tag
    }
})

# Event handler for ListView double-click
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

# Initial TreeView population
Populate-TreeView

# Initial ListView population (using Desktop path from environment variable)
$desktopPath = [Environment]::GetFolderPath("Desktop")
Populate-ListView -path $desktopPath

# Show the form
[void]$form.ShowDialog()