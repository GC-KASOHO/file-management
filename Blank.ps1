# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(1200, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)

# Create a TreeView to display directories (left side)
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(250, 600)
$treeView.Location = New-Object System.Drawing.Point(10, 50)
$treeView.Scrollable = $true
$treeView.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($treeView)

# Create a ListView to display files (right side)
$listView = New-Object System.Windows.Forms.ListView
$listView.Size = New-Object System.Drawing.Size(900, 600)
$listView.Location = New-Object System.Drawing.Point(270, 50)
$listView.View = [System.Windows.Forms.View]::Details
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.AllowColumnReorder = $true
$listView.BackColor = [System.Drawing.Color]::White

# Add columns to the ListView
$columns = @(
    @{Name="Name"; Width=300},
    @{Name="Type"; Width=100},
    @{Name="Size"; Width=100},
    @{Name="Modified Date"; Width=150},
    @{Name="Created Date"; Width=150},
    @{Name="Attributes"; Width=100}
)

foreach ($column in $columns) {
    $listView.Columns.Add($column.Name, $column.Width)
}

# Function to format file size
function Format-FileSize {
    param ([long]$size)
    $suffixes = "B", "KB", "MB", "GB", "TB"
    $i = 0
    while ($size -ge 1024 -and $i -lt ($suffixes.Count - 1)) {
        $size = $size / 1024
        $i++
    }
    return "{0:N2} {1}" -f $size, $suffixes[$i]
}

# Function to populate ListView with files and folders
function Update-FileList {
    param ([string]$path)
    
    $listView.Items.Clear()
    
    try {
        $items = Get-ChildItem -Path $path -Force -ErrorAction Stop
        
        foreach ($item in $items) {
            $listViewItem = New-Object System.Windows.Forms.ListViewItem($item.Name)
            
            # Set appropriate icon
            if ($item.PSIsContainer) {
                $type = "Folder"
                $size = ""
            } else {
                $type = $item.Extension.TrimStart(".").ToUpper()
                if ([string]::IsNullOrEmpty($type)) { $type = "File" }
                $size = Format-FileSize $item.Length
            }
            
            $listViewItem.SubItems.Add($type)
            $listViewItem.SubItems.Add($size)
            $listViewItem.SubItems.Add($item.LastWriteTime.ToString("g"))
            $listViewItem.SubItems.Add($item.CreationTime.ToString("g"))
            $listViewItem.SubItems.Add($item.Attributes.ToString())
            
            # Set different colors for folders and files
            if ($item.PSIsContainer) {
                $listViewItem.BackColor = [System.Drawing.Color]::FromArgb(240, 248, 255)
            }
            
            $listView.Items.Add($listViewItem)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error accessing path: $path`n`n$($_.Exception.Message)", 
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Add ListView to form
$form.Controls.Add($listView)

# Initialize with C: drive
Update-FileList -path "C:\"

# Add event handler for TreeView selection change
$treeView.Add_AfterSelect({
    if ($this.SelectedNode.FullPath) {
        Update-FileList -path $this.SelectedNode.FullPath
    }
})

# Show the form
[void]$form.ShowDialog()