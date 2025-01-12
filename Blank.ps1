# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form for the File Explorer
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell File Explorer with Full File Preview"
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
$listBox.Size = New-Object System.Drawing.Size(600, 250)
$listBox.Location = New-Object System.Drawing.Point(270, 50)
$form.Controls.Add($listBox)

# Create a Panel for File Preview (below the ListBox)
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Size = New-Object System.Drawing.Size(600, 200)
$previewPanel.Location = New-Object System.Drawing.Point(270, 310)
$form.Controls.Add($previewPanel)

# Create a PictureBox for Image Preview
$imagePreview = New-Object System.Windows.Forms.PictureBox
$imagePreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$imagePreview.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
$imagePreview.Visible = $false
$previewPanel.Controls.Add($imagePreview)

# Create a RichTextBox for Text File Preview
$textPreview = New-Object System.Windows.Forms.RichTextBox
$textPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$textPreview.Visible = $false
$previewPanel.Controls.Add($textPreview)

# Create a Windows Media Player for Video Preview
$wmpPreview = New-Object -ComObject WMPlayer.OCX
$wmpPreview.Dock = [System.Windows.Forms.DockStyle]::Fill
$wmpPreview.Visible = $false
$previewPanel.Controls.Add($wmpPreview)

# Function to Populate the TreeView with directories
function Populate-TreeView {
    $treeView.Nodes.Clear()
    $rootNode = $treeView.Nodes.Add("Computer", "This PC")
    $rootNode.Nodes.Add("C:\")
    $rootNode.Nodes.Add("D:\")
}

# Function to Populate ListBox with files from the selected folder
function Populate-ListBox {
    param($folderPath)
    $listBox.Items.Clear()
    if (Test-Path $folderPath) {
        $files = Get-ChildItem -Path $folderPath -File
        $listBox.Items.AddRange($files.Name)
    }
}

# Function to Preview File Content
function Preview-FileContent {
    param($filePath)

    # Hide all preview components first
    $imagePreview.Visible = $false
    $textPreview.Visible = $false
    $wmpPreview.Visible = $false

    if ($filePath -match '\.jpg|\.jpeg|\.png|\.gif|\.bmp|\.tiff$') {
        # Preview image files
        $imagePreview.Visible = $true
        $imagePreview.Image = [System.Drawing.Image]::FromFile($filePath)
    }
    elseif ($filePath -match '\.txt|\.log|\.csv|\.md|\.xml|\.json$') {
        # Preview text files
        $textPreview.Visible = $true
        $textPreview.Text = Get-Content $filePath -Raw
    }
    elseif ($filePath -match '\.mp4|\.avi|\.mov|\.mkv|\.wmv$') {
        # Preview video files using Windows Media Player
        $wmpPreview.Visible = $true
        $wmpPreview.URL = $filePath
        $wmpPreview.Ctlcontrols.play()
    }
    elseif ($filePath -match '\.pdf$') {
        # Preview PDF files (placeholder message)
        $textPreview.Visible = $true
        $textPreview.Text = "PDF preview not supported directly in PowerShell. Please use a PDF viewer."
    }
    elseif ($filePath -match '\.docx|\.doc|\.xlsx|\.xls|\.pptx|\.ppt$') {
        # Placeholder for Office files
        $textPreview.Visible = $true
        $textPreview.Text = "Office file preview not supported. Open in Office applications."
    }
    else {
        # For other file types, show a default message
        $textPreview.Visible = $true
        $textPreview.Text = "Preview not available for this file type."
    }
}

# Event to handle directory selection in TreeView
$treeView.Add_AfterSelect({
    $selectedNode = $treeView.SelectedNode
    if ($selectedNode) {
        $folderPath = $selectedNode.FullPath
        Populate-ListBox -folderPath $folderPath
    }
})

# Event to handle file selection in ListBox
$listBox.Add_SelectedIndexChanged({
    $selectedFile = $listBox.SelectedItem
    if ($selectedFile) {
        $selectedFilePath = Join-Path $folderPath $selectedFile
        Preview-FileContent -filePath $selectedFilePath
    }
})

# Populate the TreeView with initial directories
Populate-TreeView

# Show the form
[void]$form.ShowDialog()
