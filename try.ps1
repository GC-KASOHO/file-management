# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::LightGray

# Create ImageList for TreeView icons
$imageList = New-Object System.Windows.Forms.ImageList
$imageList.ImageSize = New-Object System.Drawing.Size(16, 16)

# Add icons
$folderIcon = [System.Drawing.SystemIcons]::FolderLarge.ToBitmap()
$driveIcon = [System.Drawing.SystemIcons]::Application.ToBitmap()
$homeIcon = [System.Drawing.SystemIcons]::Information.ToBitmap()

$imageList.Images.Add("drive", $driveIcon)
$imageList.Images.Add("folder", $folderIcon)
$imageList.Images.Add("home", $homeIcon)

# Create TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 500)
$treeView.Location = New-Object System.Drawing.Point(10, 50)
$treeView.ImageList = $imageList
$treeView.ShowLines = $true
$form.Controls.Add($treeView)

# Create ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(600, 500)
$listView.Location = New-Object System.Drawing.Point(270, 50)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Columns.Add("Name", 250)
$listView.Columns.Add("Size", 100)
$listView.Columns.Add("Type", 100)
$listView.Columns.Add("Last Modified", 150)
$form.Controls.Add($listView)

# Variable to store recently viewed files
$recentFiles = @()

# Function to get special folders
function Get-SpecialFolders {
    return @(
        @{ Name = "Home"; Path = "recent"; Icon = 2 },
        @{ Name = "Desktop"; Path = [System.Environment]::GetFolderPath("Desktop"); Icon = 1 },
        @{ Name = "Documents"; Path = [System.Environment]::GetFolderPath("MyDocuments"); Icon = 1 },
        @{ Name = "Downloads"; Path = (Join-Path $env:USERPROFILE "Downloads"); Icon = 1 },
        @{ Name = "Pictures"; Path = [System.Environment]::GetFolderPath("MyPictures"); Icon = 1 },
        @{ Name = "Videos"; Path = [System.Environment]::GetFolderPath("MyVideos"); Icon = 1 }
    )
}

# Populate TreeView with special folders and drives
function Populate-TreeView {
    $treeView.Nodes.Clear()

    # Add special folders
    $specialFolders = Get-SpecialFolders
    foreach ($folder in $specialFolders) {
        $folderNode = New-Object System.Windows.Forms.TreeNode
        $folderNode.Text = $folder.Name
        $folderNode.Tag = $folder.Path
        $folderNode.ImageIndex = $folder.Icon
        $folderNode.SelectedImageIndex = $folder.Icon
        $treeView.Nodes.Add($folderNode)
    }

    # Add drives
    Get-PSDrive -PSProvider FileSystem | ForEach-Object {
        $driveNode = New-Object System.Windows.Forms.TreeNode
        $driveNode.Text = $_.Name + ":\"
        $driveNode.Tag = $_.Root
        $driveNode.ImageIndex = 0
        $driveNode.SelectedImageIndex = 0
        $treeView.Nodes.Add($driveNode)
    }
}

# Function to show file details in ListView
function Show-FileDetails {
    param($path)

    $listView.Items.Clear()
    if ($path -eq "recent") {
        # Show recent files
        foreach ($recentFile in $recentFiles) {
            $item = New-Object System.Windows.Forms.ListViewItem($recentFile.Name)
            $item.SubItems.Add($recentFile.Size)
            $item.SubItems.Add($recentFile.Type)
            $item.SubItems.Add($recentFile.LastModified)
            $listView.Items.Add($item)
        }
    } else {
        try {
            Get-ChildItem -Path $path | ForEach-Object {
                $item = New-Object System.Windows.Forms.ListViewItem($_.Name)
                if ($_.PSIsContainer) {
                    $item.SubItems.Add("<DIR>")
                    $item.SubItems.Add("Folder")
                } else {
                    $item.SubItems.Add("{0:N2} KB" -f ($_.Length / 1KB))
                    $item.SubItems.Add($_.Extension)
                }
                $item.SubItems.Add($_.LastWriteTime.ToString())
                $item.Tag = $_.FullName
                $listView.Items.Add($item)

                # Track recently viewed files
                if (-not $_.PSIsContainer) {
                    $recentFiles += @{
                        Name = $_.Name
                        Size = "{0:N2} KB" -f ($_.Length / 1KB)
                        Type = $_.Extension
                        LastModified = $_.LastWriteTime.ToString()
                    }
                    $recentFiles = $recentFiles | Select-Object -Unique | Select-Object -Last 10
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Unable to access path: $path", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Event handler for TreeView selection
$treeView.Add_AfterSelect({
    param($sender, $e)
    if ($e.Node.Tag) {
        Show-FileDetails $e.Node.Tag
    }
})

# Initialize TreeView
Populate-TreeView

# Show the form
[void]$form.ShowDialog()
