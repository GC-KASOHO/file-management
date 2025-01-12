# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create Quick Access panel with increased height
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Size = New-Object System.Drawing.Size(250, 220)
$quickAccessPanel.Location = New-Object System.Drawing.Point(10, 30)
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoSize = $false
$form.Controls.Add($quickAccessPanel)

# Create a TreeView to display directories (left side)
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 290)
$treeView.Location = New-Object System.Drawing.Point(10, 250)
$treeView.Scrollable = $true
$form.Controls.Add($treeView)

# Create a ListView to display files (right side)
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(600, 510)
$listView.Location = New-Object System.Drawing.Point(270, 30)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
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


# Create a ContextMenuStrip
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Add menu items
$menuItemOpen = $contextMenu.Items.Add("Open")
$menuItemCopy = $contextMenu.Items.Add("Copy")
$menuItemDelete = $contextMenu.Items.Add("Delete")
$menuItemProperties = $contextMenu.Items.Add("Properties")

# Event handler for context menu items
$menuItemOpen.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            Start-Process $itemPath
        }
    }
})

$menuItemCopy.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            [System.Windows.Forms.Clipboard]::SetText($itemPath)
            [System.Windows.Forms.MessageBox]::Show("Copied to clipboard: $itemPath")
        }
    }
})

$menuItemDelete.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
            if (Test-Path -Path $itemPath) {
                $confirmation = [System.Windows.Forms.MessageBox]::Show(
                    "Are you sure you want to delete this item: $itemPath?",
                    "Confirm Delete",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Remove-Item -Path $itemPath -Recurse -Force -ErrorAction SilentlyContinue
                    Populate-ListView -path $global:currentPath
                }
            }
        }
    })
    
    $menuItemProperties.Add_Click({
        if ($listView.SelectedItems.Count -gt 0) {
            $itemPath = $listView.SelectedItems[0].Tag
            if (Test-Path -Path $itemPath) {
                $properties = Get-Item -Path $itemPath
                
                # Determine the type
                $type = if ($properties.PSIsContainer) { 'Folder' } else { 'File' }
                
                # Determine the size
                $size = if ([System.IO.File]::Exists($itemPath)) {
                    Format-FileSize $properties.Length
                } else {
                    'N/A'
                }
                
                # Create the properties info string
                $propertiesInfo = "Name: $($properties.Name)`n" +
                                  "Path: $($properties.FullName)`n" +
                                  "Type: $type`n" +
                                  "Size: $size`n" +
                                  "Modified: $($properties.LastWriteTime.ToString('g'))"
                
                [System.Windows.Forms.MessageBox]::Show($propertiesInfo, "Properties", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    })
    
    # Associate the context menu with the ListView
    $listView.ContextMenuStrip = $contextMenu
    
    # Event handler for right-click on ListView
    $listView.add_MouseDown({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $hitTest = $listView.HitTest($_.Location)
            if ($hitTest.Item -ne $null) {
                $listView.SelectedItems.Clear()
                $hitTest.Item.Selected = $true
            } else {
                # If right-clicked on empty space, clear selection
                $listView.SelectedItems.Clear()
            }
        }
    })
    
    # Associate the context menu with the TreeView
    $treeView.ContextMenuStrip = $contextMenu
    
    # Event handler for right-click on TreeView
    $treeView.add_MouseDown({
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
            $hitTest = $treeView.HitTest($_.Location)
            if ($hitTest.Node -ne $null) {
                $treeView.SelectedNode = $hitTest.Node
                # Optionally, you can add context menu actions for folders here
            } else {
                # If right-clicked on empty space, clear selection
                $treeView.SelectedNode = $null
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