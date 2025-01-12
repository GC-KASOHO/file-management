# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)

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
    if ($global:currentPath) {
        $folderName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter folder name:", "New Folder", "New Folder")
        if ($folderName) {
            $path = Join-Path $global:currentPath $folderName
            New-Item -Path $path -ItemType Directory
            Populate-ListView -path $global:currentPath
            $statusLabel.Text = "Created new folder: $folderName"
        }
    }
})

$refreshItem = New-Object System.Windows.Forms.ToolStripMenuItem
$refreshItem.Text = "Refresh"
$refreshItem.ShortcutKeys = [System.Windows.Forms.Keys]::F5
$refreshItem.Add_Click({
    if ($global:currentPath) {
        Populate-ListView -path $global:currentPath
        $statusLabel.Text = "View refreshed"
    }
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
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete $($selectedItem.Text)?",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq 'Yes') {
            Remove-Item $selectedItem.Tag -Force -Recurse
            Populate-ListView -path $global:currentPath
            $statusLabel.Text = "Deleted: $($selectedItem.Text)"
        }
    }
})

$renameItem = New-Object System.Windows.Forms.ToolStripMenuItem
$renameItem.Text = "Rename"
$renameItem.ShortcutKeys = [System.Windows.Forms.Keys]::F2
$renameItem.Add_Click({
    $selectedItem = $listView.SelectedItems[0]
    if ($selectedItem) {
        $oldName = $selectedItem.Text
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new name:", "Rename", $oldName)
        if ($newName -and ($newName -ne $oldName)) {
            $oldPath = $selectedItem.Tag
            $newPath = Join-Path (Split-Path $oldPath) $newName
            Rename-Item -Path $oldPath -NewName $newName
            Populate-ListView -path $global:currentPath
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
    $listView.View = [System.Windows.Forms.View]::Details
    $statusLabel.Text = "Switched to details view"
})

$iconsViewItem = New-Object System.Windows.Forms.ToolStripMenuItem
$iconsViewItem.Text = "Large Icons"
$iconsViewItem.Add_Click({
    $listView.View = [System.Windows.Forms.View]::LargeIcon
    $statusLabel.Text = "Switched to large icons view"
})

# Add menu items to their menus
$fileMenu.DropDownItems.AddRange(@($newFolderItem, $refreshItem, 
    (New-Object System.Windows.Forms.ToolStripSeparator), $exitItem))
$editMenu.DropDownItems.AddRange(@($deleteItem, $renameItem))
$viewMenu.DropDownItems.AddRange(@($detailsViewItem, $iconsViewItem))

# Add menus to menu strip
$menuStrip.Items.AddRange(@($fileMenu, $editMenu, $viewMenu))

# Create main TableLayoutPanel
$mainTableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainTableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$mainTableLayoutPanel.ColumnCount = 2
$mainTableLayoutPanel.RowCount = 2
$mainTableLayoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$mainTableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$mainTableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$mainTableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$mainTableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))

# Create left panel TableLayoutPanel
$leftTableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$leftTableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$leftTableLayoutPanel.ColumnCount = 1
$leftTableLayoutPanel.RowCount = 3
$leftTableLayoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$leftTableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$leftTableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 220))) # QuickAccess
$leftTableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))  # This PC Button
$leftTableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) # TreeView

# Create Quick Access panel
$quickAccessPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$quickAccessPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$quickAccessPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$quickAccessPanel.WrapContents = $false
$quickAccessPanel.AutoScroll = $true
$quickAccessPanel.BackColor = [System.Drawing.Color]::WhiteSmoke

# Create This PC Button
$thisPCButton = New-Object System.Windows.Forms.Button
$thisPCButton.Dock = [System.Windows.Forms.DockStyle]::Fill
$thisPCButton.Text = "This PC ▶"
$thisPCButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Popup
$thisPCButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$thisPCButton.BackColor = [System.Drawing.Color]::LightGray

# Create TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeView.Scrollable = $true
$treeView.Visible = $false
$treeView.BackColor = [System.Drawing.Color]::WhiteSmoke

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Dock = [System.Windows.Forms.DockStyle]::Fill
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.BackColor = [System.Drawing.Color]::WhiteSmoke

# Create status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusLabel.Text = "Ready"

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

# Add controls to left panel
$leftTableLayoutPanel.Controls.Add($quickAccessPanel, 0, 0)
$leftTableLayoutPanel.Controls.Add($thisPCButton, 0, 1)
$leftTableLayoutPanel.Controls.Add($treeView, 0, 2)

# Add controls to main panel
$mainTableLayoutPanel.Controls.Add($leftTableLayoutPanel, 0, 0)
$mainTableLayoutPanel.Controls.Add($listView, 1, 0)
$mainTableLayoutPanel.Controls.Add($statusLabel, 0, 1)
$mainTableLayoutPanel.SetColumnSpan($statusLabel, 2)

# [Previous functions remain the same: Get-CurrentDrives, Format-FileSize, etc.]
# [Include all the previous functions here exactly as they were]

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
        $statusLabel.Text = "Current path: $path"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error accessing path: $path`n$($_.Exception.Message)",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $statusLabel.Text = "Error accessing path: $path"
    }
}

# Function to create Quick Access buttons
function Add-QuickAccessButton($text, $path) {
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $text
    $button.Width = 230  # Slightly smaller to account for scrollbar
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
    
    # Get all drives
    Get-PSDrive -PSProvider FileSystem | ForEach-Object {
        $driveNode = $treeView.Nodes.Add($_.Root)
        $driveNode.Tag = $_.Root
        try {
            Get-ChildItem -Path $_.Root -Directory -ErrorAction Stop | ForEach-Object {
                $subNode = $driveNode.Nodes.Add($_.Name)
                $subNode.Tag = $_.FullName
            }
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

# Event handler for ListView column click (sorting)
$listView.Add_ColumnClick({
    param($sender, $e)
    
    $column = $e.Column
    $listView = $sender
    
    # If current sorting column is different from clicked column, sort ascending
    if ($script:sortColumn -ne $column) {
        $script:sortAscending = $true
    } else {
        # If same column, toggle sort direction
        $script:sortAscending = !$script:sortAscending
    }
    
    $script:sortColumn = $column
    
    # Sort the items
    $listView.ListViewItemSorter = New-Object System.Windows.Forms.ListViewItemComparer($column, $script:sortAscending)
    $listView.Sort()
})

# Custom comparer for ListView sorting
Add-Type -TypeDefinition @"
using System;
using System.Collections;
using System.Windows.Forms;

public class ListViewItemComparer : IComparer
{
    private int col;
    private bool ascending;
    
    public ListViewItemComparer(int column, bool asc)
    {
        col = column;
        ascending = asc;
    }
    
    public int Compare(object x, object y)
    {
        ListViewItem itemX = (ListViewItem)x;
        ListViewItem itemY = (ListViewItem)y;
        
        string textX = col == 0 ? itemX.Text : itemX.SubItems[col].Text;
        string textY = col == 0 ? itemY.Text : itemY.SubItems[col].Text;
        
        // Handle size column specially
        if (col == 2)
        {
            // If both are folders (empty size), sort by name
            if (string.IsNullOrEmpty(textX) && string.IsNullOrEmpty(textY))
                return ascending ? 
                    string.Compare(itemX.Text, itemY.Text) : 
                    string.Compare(itemY.Text, itemX.Text);
            
            // Folders always come before files
            if (string.IsNullOrEmpty(textX)) return ascending ? -1 : 1;
            if (string.IsNullOrEmpty(textY)) return ascending ? 1 : -1;
            
            // Try to parse the size values
            try {
                double sizeX = ParseSize(textX);
                double sizeY = ParseSize(textY);
                return ascending ? 
                    sizeX.CompareTo(sizeY) : 
                    sizeY.CompareTo(sizeX);
            }
            catch {
                return ascending ? 
                    string.Compare(textX, textY) : 
                    string.Compare(textY, textX);
            }
        }
        
        return ascending ? 
            string.Compare(textX, textY) : 
            string.Compare(textY, textX);
    }
    
    private double ParseSize(string size)
    {
        string[] parts = size.Split(' ');
        if (parts.Length != 2) return 0;
        
        double value = Convert.ToDouble(parts[0]);
        string unit = parts[1].ToUpper();
        
        switch (unit)
        {
            case "B": return value;
            case "KB": return value * 1024;
            case "MB": return value * 1024 * 1024;
            case "GB": return value * 1024 * 1024 * 1024;
            case "TB": return value * 1024 * 1024 * 1024 * 1024;
            default: return 0;
        }
    }
}
"@

# Form closing event to clean up timer
$form.Add_FormClosing({
    $driveTimer.Stop()
})

# Add MenuStrip and main TableLayoutPanel to form
$form.Controls.Add($menuStrip)
$form.Controls.Add($mainTableLayoutPanel)

# Initial ListView population (using Desktop path from environment variable)
$desktopPath = [Environment]::GetFolderPath("Desktop")
Populate-ListView -path $desktopPath

# Initialize sorting variables
$script:sortColumn = 0
$script:sortAscending = $true

# Show the form
[void]$form.ShowDialog()