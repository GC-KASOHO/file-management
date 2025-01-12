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

# Show the form
[void]$form.ShowDialog()
