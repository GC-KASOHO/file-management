# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"

# Set the background color (optional)
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

# Create a TextBox for the search functionality with black border
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Size = New-Object System.Drawing.Size(200, 20)
$searchBox.Location = New-Object System.Drawing.Point(600,12)
$searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$searchBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($searchBox)

# Create a Button for initiating the search
$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Size = New-Object System.Drawing.Size(70, 20)
$searchButton.Location = New-Object System.Drawing.Point(806, 11)
$form.Controls.Add($searchButton)

# Function to search for files and folders
function Search-FilesAndFolders {
    param (
        [string]$query,
        [string]$path
    )

    $listBox.Items.Clear()
    try {
        Get-ChildItem -Path $path -Recurse -ErrorAction Stop | Where-Object {
            $_.Name -like "*$query*"
        } | ForEach-Object {
            $listBox.Items.Add($_.FullName)
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error searching: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Handle the search button click
$searchButton.Add_Click({
    $query = $searchBox.Text
    $currentPath = if ($treeView.SelectedNode) { $treeView.SelectedNode.Tag } else { "C:\" }

    if (-not [string]::IsNullOrWhiteSpace($query)) {
        Search-FilesAndFolders -query $query -path $currentPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a search query.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Show the form
[void]$form.ShowDialog()