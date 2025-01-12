﻿# Import required assemblies
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
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = [System.Windows.Forms.DockStyle]::Fill
$treeView.Scrollable = $true

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

# Show the form
[void]$form.ShowDialog()
#Piadozo Menu Bar