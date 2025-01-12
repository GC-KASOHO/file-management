# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create a TreeView to display directories (left side)
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 500)
$treeView.Location = New-Object System.Drawing.Point(10, 50)
$treeView.Scrollable = $true
$form.Controls.Add($treeView)

# Create a ListBox to display files (right side)
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Size = New-Object System.Drawing.Size(600, 500)
$listBox.Location = New-Object System.Drawing.Point(270, 50)
$form.Controls.Add($listBox)

function Show-ContextMenu {
    param (
        [string]$Path,
        [System.Windows.Forms.Control]$Control,
        [System.Windows.Forms.MouseEventArgs]$MouseEvent
    )
    
    # Create context menu
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    
    # Determine if path is file or directory
    $isDirectory = (Get-Item $Path) -is [System.IO.DirectoryInfo]
    
    # Open item
    $openMenuItem = $contextMenu.Items.Add("Open")
    $openMenuItem.Add_Click({
        if ($isDirectory) {
            Set-Location $Path
            Refresh-FileView
        } else {
            Start-Process $Path
        }
    })
    
    # Copy
    $copyMenuItem = $contextMenu.Items.Add("Copy")
    $copyMenuItem.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($Path)
        Write-Host "Path copied to clipboard: $Path"
    })
    
    # Cut
    $cutMenuItem = $contextMenu.Items.Add("Cut")
    $cutMenuItem.Add_Click({
        $script:cutPath = $Path
        $script:isCut = $true
        Write-Host "Item marked for moving: $Path"
    })
    
    # Paste
    if ($isDirectory) {
        $pasteMenuItem = $contextMenu.Items.Add("Paste")
        $pasteMenuItem.Add_Click({
            if ($script:cutPath) {
                if ($script:isCut) {
                    Move-Item $script:cutPath $Path
                    $script:cutPath = $null
                    $script:isCut = $false
                }
            } elseif ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
                $files = [System.Windows.Forms.Clipboard]::GetFileDropList()
                foreach ($file in $files) {
                    Copy-Item $file $Path
                }
            }
            Refresh-FileView
        })
    }
    
    # Add separator
    $contextMenu.Items.Add("-")
    
    # Delete
    $deleteMenuItem = $contextMenu.Items.Add("Delete")
    $deleteMenuItem.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete this item?",
            "Confirm Delete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-Item $Path -Force -Recurse
            Refresh-FileView
        }
    })
    
    # Rename
    $renameMenuItem = $contextMenu.Items.Add("Rename")
    $renameMenuItem.Add_Click({
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter new name:",
            "Rename",
            (Split-Path $Path -Leaf)
        )
        
        if ($newName -and $newName -ne "") {
            $newPath = Join-Path (Split-Path $Path -Parent) $newName
            Rename-Item $Path $newPath
            Refresh-FileView
        }
    })
    
    # Properties
    $propertiesMenuItem = $contextMenu.Items.Add("Properties")
    $propertiesMenuItem.Add_Click({
        $item = Get-Item $Path
        $info = @"
Properties for: $($item.Name)
Type: $($item.GetType().Name)
Created: $($item.CreationTime)
Modified: $($item.LastWriteTime)
"@
        if ($isDirectory) {
            $info += "`nContents: $((Get-ChildItem $Path | Measure-Object).Count) items"
        } else {
            $info += "`nSize: $([math]::Round($item.Length/1KB, 2)) KB"
        }
        
        [System.Windows.Forms.MessageBox]::Show($info, "Properties")
    })
    
    # Show the context menu at mouse position
    $contextMenu.Show($Control, $MouseEvent.Location)
}

# Right-click event for ListBox
$listBox.Add_MouseClick({
    param($sender, $e)
    
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $item = $listBox.SelectedItem
        if ($item) {
            Show-ContextMenu -Path $item -Control $listBox -MouseEvent $e
        }
    }
})

# Right-click event for TreeView
$treeView.Add_MouseClick({
    param($sender, $e)
    
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $node = $treeView.GetNodeAt($e.Location)
        if ($node) {
            $path = $node.Tag
            Show-ContextMenu -Path $path -Control $treeView -MouseEvent $e
        } else {
            # If right-clicked on empty space in TreeView, show context menu for creating new folder
            $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
            $newFolderMenuItem = $contextMenu.Items.Add("New Folder")
            $newFolderMenuItem.Add_Click({
                $newFolderName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new folder name:", "New Folder", "New Folder")
                if ($newFolderName -and $newFolderName -ne "") {
                    $selectedNode = $treeView.SelectedNode
                    if ($selectedNode) {
                        $newFolderPath = Join-Path $selectedNode.Tag $newFolderName
                        New-Item -ItemType Directory -Path $newFolderPath
                        Refresh-TreeView
                    }
                }
            })
            $contextMenu.Show($treeView, $e.Location)
        }
    }
})

# Function to refresh the ListBox based on the selected directory in the TreeView
function Refresh-FileView {
    $listBox.Items.Clear()
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode) {
        $path = $selectedNode.Tag
        Get-ChildItem -Path $path | ForEach-Object {
            $listBox.Items.Add($_.FullName) # You can customize what to display
        }
    }
}

# Function to refresh the TreeView
function Refresh-TreeView {
    $treeView.Nodes.Clear()
    $root = New-Object System.Windows.Forms.TreeNode("C:\") # Change this to your desired root path
    $root.Tag = "C:\" # Set the tag for the root node
    $treeView.Nodes.Add($root)
    Populate-TreeView $root
}

# Function to populate the TreeView with directories
function Populate-TreeView {
    param (
        [System.Windows.Forms.TreeNode]$node
    )
    
    $path = $node.Tag
    try {
        $directories = Get-ChildItem -Path $path -Directory
        foreach ($dir in $directories) {
            $childNode = New-Object System.Windows.Forms.TreeNode($dir.Name)
            $childNode.Tag = $dir.FullName  # Set the tag for the child node
            $node.Nodes.Add($childNode)
            Populate-TreeView $childNode  # Recursively populate child nodes
        }
    } catch {
        # Handle any errors (e.g., access denied)
        Write-Host "Error accessing path: $path - $_"
    }
}

# Initial population of the TreeView
Refresh-TreeView

# Show the form
[void]$form.ShowDialog()