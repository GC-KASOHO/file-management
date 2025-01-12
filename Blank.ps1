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

# Calculate the center position for the smaller address bar
$addressBarWidth = 500  # Reduced width
$addressBarHeight = 20  # Smaller height
$addressBarLeft = ($form.Width - $addressBarWidth) / 2  # Center the address bar horizontally

# Create a smaller address bar TextBox (centered)
$addressBar = New-Object System.Windows.Forms.TextBox
$addressBar.Size = New-Object System.Drawing.Size($addressBarWidth, $addressBarHeight)
$addressBar.Location = New-Object System.Drawing.Point($addressBarLeft, 10)
$addressBar.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Regular)  # Smaller font
$addressBar.ForeColor = [System.Drawing.Color]::DarkBlue
$addressBar.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($addressBar)

# Event handler for pressing Enter in the address bar
$addressBar.Add_KeyDown({
    param ($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $path = $addressBar.Text
        if (Test-Path $path) {
            $listBox.Items.Clear()
            $files = Get-ChildItem -Path $path -File | ForEach-Object { $_.Name }
            $listBox.Items.AddRange($files)
        } else {
            [System.Windows.Forms.MessageBox]::Show("The path is invalid.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# Show the form
[void]$form.ShowDialog()
