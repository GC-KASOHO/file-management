# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)  # Set minimum size

# Create TableLayoutPanel for better resizing
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.ColumnCount = 2
$tableLayoutPanel.RowCount = 2
$tableLayoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))

# Create MenuStrip
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.MainMenuStrip = $menuStrip

# File Menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

$newFolderItem = New-Object System.Windows.Forms.ToolStripMenuItem
$newFolderItem.Text = "New Folder"
$newFolderItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N
$newFolderItem.Add_Click({
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode) {
        $folderName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter folder name:", "New Folder", "New Folder")
        if ($folderName) {
            $path = Join-Path $selectedNode.Tag $folderName
            New-Item -Path $path -ItemType Directory
            $newNode = New-Object System.Windows.Forms.TreeNode
            $newNode.Text = $folderName
            $newNode.Tag = $path
            $selectedNode.Nodes.Add($newNode)
            $statusLabel.Text = "Created new folder: $folderName"
        }
    }
})

$refreshItem = New-Object System.Windows.Forms.ToolStripMenuItem
$refreshItem.Text = "Refresh"
$refreshItem.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$refreshItem.Add_Click({
    RefreshExplorer
    $statusLabel.Text = "View refreshed"
})

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitItem.Text = "Exit"
$exitItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$exitItem.Add_Click({ $form.Close() })

# Edit Menu
$editMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$editMenu.Text = "Edit"

$deleteItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteItem.Text = "Delete"
$deleteItem.ShortcutKeys = [System.Windows.Forms.Keys]::Delete
$deleteItem.Add_Click({
    if ($listBox.SelectedItem) {
        $selectedPath = Join-Path $currentPath $listBox.SelectedItem
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete $($listBox.SelectedItem)?",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq 'Yes') {
            Remove-Item $selectedPath -Force -Recurse
            RefreshExplorer
            $statusLabel.Text = "Deleted: $($listBox.SelectedItem)"
        }
    }
})

$renameItem = New-Object System.Windows.Forms.ToolStripMenuItem
$renameItem.Text = "Rename"
$renameItem.ShortcutKeys = [System.Windows.Forms.Keys]::F2
$renameItem.Add_Click({
    if ($listBox.SelectedItem) {
        $oldName = $listBox.SelectedItem
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new name:", "Rename", $oldName)
        if ($newName -and ($newName -ne $oldName)) {
            $oldPath = Join-Path $currentPath $oldName
            $newPath = Join-Path $currentPath $newName
            Rename-Item -Path $oldPath -NewName $newName
            RefreshExplorer
            $statusLabel.Text = "Renamed: $oldName to $newName"
        }
    }
})

# View Menu
$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "View"

$detailsViewItem = New-Object System.Windows.Forms.ToolStripMenuItem
$detailsViewItem.Text = "Details View"
$detailsViewItem.Add_Click({
    # TODO: Implement detailed view
    $statusLabel.Text = "Switched to details view"
})

$iconsViewItem = New-Object System.Windows.Forms.ToolStripMenuItem
$iconsViewItem.Text = "Icons View"
$iconsViewItem.Add_Click({
    # TODO: Implement icons view
    $statusLabel.Text = "Switched to icons view"
})

# Create TreeView and ListBox with Dock property
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# Create Quick Access panel with increased height
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Size = New-Object System.Drawing.Size(250, 220)
$quickAccessPanel.Location = New-Object System.Drawing.Point(10, 30)
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoSize = $false
$quickAccessPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($quickAccessPanel)

# Create This PC Button
$thisPCButton = New-Object System.Windows.Forms.Button
$thisPCButton.Text = "This PC ▶"
$thisPCButton.Width = 240
$thisPCButton.Height = 30
$thisPCButton.Location = New-Object System.Drawing.Point(13, 250)
$thisPCButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$thisPCButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$thisPCButton.BackColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($thisPCButton)

# Create a TreeView to display directories (left side)
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeView.Size = New-Object System.Drawing.Size(250, 260)
$treeView.Location = New-Object System.Drawing.Point(10, 280)
$treeView.Scrollable = $true
$treeView.Visible = $false
$treeView.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($treeView)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Dock = [System.Windows.Forms.DockStyle]::Fill

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusLabel.Text = "Ready"

# Add controls to TableLayoutPanel
$tableLayoutPanel.Controls.Add($treeView, 0, 0)
$tableLayoutPanel.Controls.Add($listBox, 1, 0)
$tableLayoutPanel.SetColumnSpan($statusLabel, 2)
$tableLayoutPanel.Controls.Add($statusLabel, 0, 1)

# Add menu items to their menus
$fileMenu.DropDownItems.AddRange(@($newFolderItem, $refreshItem, 
    (New-Object System.Windows.Forms.ToolStripSeparator), $exitItem))
$editMenu.DropDownItems.AddRange(@($deleteItem, $renameItem))
$viewMenu.DropDownItems.AddRange(@($detailsViewItem, $iconsViewItem))

# Add menus to menu strip
$menuStrip.Items.AddRange(@($fileMenu, $editMenu, $viewMenu))

# Function to refresh the explorer
function RefreshExplorer {
    if ($treeView.SelectedNode) {
        $currentPath = $treeView.SelectedNode.Tag
        $listBox.Items.Clear()
        Get-ChildItem -Path $currentPath | ForEach-Object {
            $listBox.Items.Add($_.Name)
        }
    }
}

# TreeView node selection event
$treeView.Add_AfterSelect({
    $currentPath = $this.SelectedNode.Tag
    RefreshExplorer
})

# Initialize root node (My Computer)
$rootNode = New-Object System.Windows.Forms.TreeNode
$rootNode.Text = "My Computer"
$rootNode.Tag = "\\"
$treeView.Nodes.Add($rootNode)

# Populate drives
Get-PSDrive -PSProvider FileSystem | ForEach-Object {
    $driveNode = New-Object System.Windows.Forms.TreeNode
    $driveNode.Text = $_.Name + ":\"
    $driveNode.Tag = $_.Root
    $rootNode.Nodes.Add($driveNode)
}

# Add MenuStrip and TableLayoutPanel to form
$form.Controls.Add($menuStrip)
$form.Controls.Add($tableLayoutPanel)
# Create a ListView to display files (right side)
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(600, 510)
$listView.Location = New-Object System.Drawing.Point(270, 30)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.BackColor = [System.Drawing.Color]::WhiteSmoke
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

# Create a timer for drive monitoring
$driveTimer = New-Object System.Windows.Forms.Timer
$driveTimer.Interval = 2000  # Check every 2 seconds
$script:currentDrives = @()

# Function to get current drives
function Get-CurrentDrives {
    return @(Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Root)
}

# Initialize current drives
$script:currentDrives = Get-CurrentDrives

# Timer tick event handler
$driveTimer.Add_Tick({
    $newDrives = Get-CurrentDrives
    
    # Check if drives have changed
    if (($newDrives.Count -ne $script:currentDrives.Count) -or 
        (Compare-Object -ReferenceObject $script:currentDrives -DifferenceObject $newDrives)) {
        
        $script:currentDrives = $newDrives
        
        # Only refresh if TreeView is visible
        if ($treeView.Visible) {
            Populate-TreeView
        }
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
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
    $button.BackColor = [System.Drawing.Color]::LightGray
    
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
    
    # Get all drives
    Get-PSDrive -PSProvider FileSystem | ForEach-Object {
        $driveNode = $treeView.Nodes.Add($_.Root)
        $driveNode.Tag = $_.Root
        try {
            Get-ChildItem -Path $_.Root -Directory -ErrorAction Stop | ForEach-Object {
                $subNode = $driveNode.Nodes.Add($_.Name)
                $subNode.Tag = $_.FullName
            }
            # Auto-expand drive nodes

        } catch {}
    }
}

# Add click event for This PC Button
$thisPCButton.Add_Click({
    if ($treeView.Visible) {
        $treeView.Visible = $false
        $thisPCButton.Text = "This PC ▶"
        $driveTimer.Stop()
    } else {
        $treeView.Visible = $true
        $thisPCButton.Text = "This PC ▼"
        if ($treeView.Nodes.Count -eq 0) {
            Populate-TreeView
        }
        $driveTimer.Start()
    }
})

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

# Form closing event to clean up timer
$form.Add_FormClosing({
    $driveTimer.Stop()
})

# Initial ListView population (using Desktop path from environment variable)
$desktopPath = [Environment]::GetFolderPath("Desktop")
Populate-ListView -path $desktopPath

# Show the form
[void]$form.ShowDialog()