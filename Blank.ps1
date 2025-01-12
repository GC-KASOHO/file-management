# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.VisualStyles

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)

# Create MenuStrip (keeping existing menu code...)
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$form.MainMenuStrip = $menuStrip

# Create ToolStrip (Toolbar)
$toolStrip = New-Object System.Windows.Forms.ToolStrip
$toolStrip.ImageScalingSize = New-Object System.Drawing.Size(16, 16)
$toolStrip.BackColor = [System.Drawing.Color]::WhiteSmoke

# New Folder Button
$newFolderButton = New-Object System.Windows.Forms.ToolStripButton
$newFolderButton.Image = [System.Drawing.SystemIcons]::FolderClosed.ToBitmap()
$newFolderButton.Text = "New Folder"
$newFolderButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
$newFolderButton.Add_Click({
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode) {
        $folderName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter folder name:", "New Folder", "New Folder")
        if ($folderName) {
            $path = Join-Path $selectedNode.Tag $folderName
            New-Item -Path $path -ItemType Directory
            RefreshExplorer
            $statusLabel.Text = "Created new folder: $folderName"
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a location first.", "New Folder")
    }
})

# Copy Button
$copyButton = New-Object System.Windows.Forms.ToolStripButton
$copyButton.Image = [System.Drawing.SystemIcons]::WinLogo.ToBitmap()  # Using WinLogo as placeholder
$copyButton.Text = "Copy"
$copyButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
$copyButton.Add_Click({
    if ($listBox.SelectedItem) {
        $script:copySource = Join-Path $currentPath $listBox.SelectedItem
        $script:copyOperation = "Copy"
        $statusLabel.Text = "Ready to copy: $($listBox.SelectedItem)"
    }
})

# Cut Button
$cutButton = New-Object System.Windows.Forms.ToolStripButton
$cutButton.Image = [System.Drawing.SystemIcons]::Warning.ToBitmap()  # Using Warning as placeholder
$cutButton.Text = "Cut"
$cutButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
$cutButton.Add_Click({
    if ($listBox.SelectedItem) {
        $script:copySource = Join-Path $currentPath $listBox.SelectedItem
        $script:copyOperation = "Cut"
        $statusLabel.Text = "Ready to move: $($listBox.SelectedItem)"
    }
})

# Paste Button
$pasteButton = New-Object System.Windows.Forms.ToolStripButton
$pasteButton.Image = [System.Drawing.SystemIcons]::Information.ToBitmap()  # Using Information as placeholder
$pasteButton.Text = "Paste"
$pasteButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
$pasteButton.Add_Click({
    if ($script:copySource -and $treeView.SelectedNode) {
        $destination = $treeView.SelectedNode.Tag
        $fileName = Split-Path $script:copySource -Leaf
        $destinationPath = Join-Path $destination $fileName
        
        try {
            if ($script:copyOperation -eq "Copy") {
                Copy-Item -Path $script:copySource -Destination $destinationPath -Recurse
                $statusLabel.Text = "Copied: $fileName"
            } else {
                Move-Item -Path $script:copySource -Destination $destinationPath
                $statusLabel.Text = "Moved: $fileName"
            }
            RefreshExplorer
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error during $script:copyOperation operation: $_", "Error")
        }
        
        $script:copySource = $null
        $script:copyOperation = $null
    }
})

# Delete Button
$deleteButton = New-Object System.Windows.Forms.ToolStripButton
$deleteButton.Image = [System.Drawing.SystemIcons]::Error.ToBitmap()  # Using Error as placeholder
$deleteButton.Text = "Delete"
$deleteButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
$deleteButton.Add_Click({
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

# Refresh Button
$refreshButton = New-Object System.Windows.Forms.ToolStripButton
$refreshButton.Image = [System.Drawing.SystemIcons]::Question.ToBitmap()  # Using Question as placeholder
$refreshButton.Text = "Refresh"
$refreshButton.DisplayStyle = [System.Windows.Forms.ToolStripItemDisplayStyle]::ImageAndText
$refreshButton.Add_Click({
    RefreshExplorer
    $statusLabel.Text = "View refreshed"
})

# Add separators and buttons to toolbar
$toolStrip.Items.AddRange(@(
    $newFolderButton,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $copyButton,
    $cutButton,
    $pasteButton,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $deleteButton,
    (New-Object System.Windows.Forms.ToolStripSeparator),
    $refreshButton
))

# Create TableLayoutPanel for the main content
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.ColumnCount = 2
$tableLayoutPanel.RowCount = 2
$tableLayoutPanel.CellBorderStyle = [System.Windows.Forms.TableLayoutPanelCellBorderStyle]::None
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))

# Create TreeView and ListBox
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

# Script-level variables for copy/cut operations
$script:copySource = $null
$script:copyOperation = $null

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

# Add all controls to form
$form.Controls.Add($menuStrip)
$form.Controls.Add($toolStrip)
$form.Controls.Add($tableLayoutPanel)

# Position the TableLayoutPanel below the ToolStrip
$tableLayoutPanel.Location = New-Object System.Drawing.Point(0, $toolStrip.Height + $menuStrip.Height)

# Show the form
[void]$form.ShowDialog()
#Veslino