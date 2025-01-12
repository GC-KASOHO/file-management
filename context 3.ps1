# =====================================================================================
# =====================================================================================
# Create a ContextMenuStrip
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Add menu items
$menuItemOpen = $contextMenu.Items.Add("Open")
$menuItemCopy = $contextMenu.Items.Add("Copy")
$menuItemCut = $contextMenu.Items.Add("Cut")
$menuItemPaste = $contextMenu.Items.Add("Paste")
$menuItemDelete = $contextMenu.Items.Add("Delete")
$menuItemProperties = $contextMenu.Items.Add("Properties")
$menuItemRefresh = $contextMenu.Items.Add("Refresh")

# Create a sub-menu for New options
$newSubMenu = New-Object System.Windows.Forms.ToolStripMenuItem("New")
$contextMenu.Items.Add($newSubMenu)

# Add Folder option
$newSubMenu.DropDownItems.Add("Folder").Add_Click({
    Create-NewFolder
})

# Add separator
$newSubMenu.DropDownItems.Add("-")

# Add file format options
$newSubMenu.DropDownItems.Add("Text File").Add_Click({
    Create-NewFile -extension ".txt"
})
$newSubMenu.DropDownItems.Add("Rich Text File").Add_Click({
    Create-NewFile -extension ".rtf"
})
$newSubMenu.DropDownItems.Add("Word Document").Add_Click({
    Create-NewFile -extension ".docx"
})
$newSubMenu.DropDownItems.Add("Excel Workbook").Add_Click({
    Create-NewFile -extension ".xlsx"
})
$newSubMenu.DropDownItems.Add("PowerPoint Presentation").Add_Click({
    Create-NewFile -extension ".pptx"
})
$newSubMenu.DropDownItems.Add("PDF Document").Add_Click({
    Create-NewFile -extension ".pdf"
})
$newSubMenu.DropDownItems.Add("CSV File").Add_Click({
    Create-NewFile -extension ".csv"
})
$newSubMenu.DropDownItems.Add("Markdown File").Add_Click({
    Create-NewFile -extension ".md"
})
$newSubMenu.DropDownItems.Add("XML File").Add_Click({
    Create-NewFile -extension ".xml"
})
$newSubMenu.DropDownItems.Add("JSON File").Add_Click({
    Create-NewFile -extension ".json"
})

# Add these new functions after your existing functions:

# Function to show an input dialog using Windows Forms
function Show-InputDialog {
    param(
        [string]$prompt,
        [string]$title,
        [string]$defaultValue
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(350,150)
    $form.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = $prompt
    $form.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,50)
    $textBox.Size = New-Object System.Drawing.Size(310,20)
    $textBox.Text = $defaultValue
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(75,80)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(170,80)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    }
    return $null
}

# Function to create a new folder
function Create-NewFolder {
    $folderName = Show-InputDialog -prompt "Enter the name for the new folder:" -title "New Folder" -defaultValue "New Folder"
    
    if (![string]::IsNullOrWhiteSpace($folderName)) {
        $fullPath = Join-Path -Path $global:currentPath -ChildPath $folderName
        try {
            # Create the new folder
            New-Item -Path $fullPath -ItemType Directory -Force
            [System.Windows.Forms.MessageBox]::Show("Created: $fullPath", "Folder Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Populate-ListView -path $global:currentPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating folder: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Enhanced function to create a new file
function Create-NewFile {
    param (
        [string]$extension
    )

    $fileName = Show-InputDialog -prompt "Enter the name for the new file (without extension):" -title "New File" -defaultValue "New File"
    
    if (![string]::IsNullOrWhiteSpace($fileName)) {
        $fullPath = Join-Path -Path $global:currentPath -ChildPath ($fileName + $extension)
        try {
            # Create the new file
            $null = New-Item -Path $fullPath -ItemType File -Force
            
            # Add default content based on file type
            switch ($extension) {
                ".md" { 
                    Set-Content -Path $fullPath -Value "# New Markdown Document`n`nCreated on $(Get-Date -Format 'yyyy-MM-dd')"
                }
                ".xml" {
                    Set-Content -Path $fullPath -Value "<?xml version=`"1.0`" encoding=`"UTF-8`"?>`n<root>`n</root>"
                }
                ".json" {
                    Set-Content -Path $fullPath -Value "{`n    `"created`": `"$(Get-Date -Format 'yyyy-MM-dd')`"`n}"
                }
                ".html" {
                    Set-Content -Path $fullPath -Value "<!DOCTYPE html>`n<html>`n<head>`n    <title>New Document</title>`n</head>`n<body>`n    <h1>New Document</h1>`n</body>`n</html>"
                }
            }

            [System.Windows.Forms.MessageBox]::Show("Created: $fullPath", "File Created", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Populate-ListView -path $global:currentPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

# Event handler for Open
$menuItemOpen.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            Start-Process $itemPath
        }
    }
})

# Event handler for Copy
$menuItemCopy.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            $global:clipboardPath = $itemPath
            $global:isCut = $false
            [System.Windows.Forms.MessageBox]::Show("Copied to clipboard: $itemPath")
        }
    }
})

# Event handler for Cut
$menuItemCut.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            $global:clipboardPath = $itemPath
            $global:isCut = $true
            [System.Windows.Forms.MessageBox]::Show("Cut to clipboard: $itemPath")
        }
    }
})

# Event handler for Paste
$menuItemPaste.Add_Click({
    if ($global:clipboardPath -ne $null -and (Test-Path -Path $global:currentPath)) {
        $destinationPath = Join-Path -Path $global:currentPath -ChildPath (Get-Item $global:clipboardPath).Name
        try {
            if ($global:isCut) {
                Move-Item -Path $global:clipboardPath -Destination $destinationPath -Force
                [System.Windows.Forms.MessageBox]::Show("Moved to: $destinationPath")
            } else {
                Copy-Item -Path $global:clipboardPath -Destination $destinationPath -Force
                [System.Windows.Forms.MessageBox]::Show("Copied to: $destinationPath")
            }
            Populate-ListView -path $global:currentPath
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error during paste operation: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No item to paste or invalid destination.", "Paste Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Event handler for Delete
$menuItemDelete.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            $confirmation = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to delete this item: $itemPath?",
                "Confirm Delete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
                Remove-Item -Path $itemPath -Recurse -Force -ErrorAction SilentlyContinue
                Populate-ListView -path $global:currentPath
            }
        }
    }
})

# Event handler for Properties
$menuItemProperties.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            $properties = Get-Item -Path $itemPath
            
            # Determine the type
            $type = if ($properties.PSIsContainer) { 'Folder' } else { 'File' }
            
            # Determine the size
            $size = if ([System.IO.File]::Exists($itemPath)) {
                Format-FileSize $properties.Length
            } else {
                'N/A'
            }
            
            # Create the properties info string
            $propertiesInfo = "Name: $($properties.Name)`n" +
                              "Path: $($properties.FullName)`n" +
                              "Type: $type`n" +
                              "Size: $size`n" +
                              "Modified: $($properties.LastWriteTime.ToString('g'))"
            
            [System.Windows.Forms.MessageBox]::Show($propertiesInfo, "Properties", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# Event handler for Refresh
$menuItemRefresh.Add_Click({
    Populate-ListView -path $global:currentPath
})

# Associate the context menu with the ListView
$listView.ContextMenuStrip = $contextMenu

# Event handler for right-click on ListView
$listView.add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hitTest = $listView.HitTest($_.Location)
        if ($hitTest.Item -ne $null) {
            $listView.SelectedItems.Clear()
            $hitTest.Item.Selected = $true
        } else {
            # If right-clicked on empty space, clear selection
            $listView.SelectedItems.Clear()
        }
    }
})

# Associate the context menu with the TreeView
$treeView.ContextMenuStrip = $contextMenu

# Event handler for right-click on TreeView
$treeView.add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hitTest = $treeView.HitTest($_.Location)
        if ($hitTest.Node -ne $null) {
            $treeView.SelectedNode = $hitTest.Node
            # Optionally, you can add context menu actions for folders here
        } else {
            # If right-clicked on empty space, clear selection
            $treeView.SelectedNode = $null
        }
    }
})

# Initial TreeView population
Populate-TreeView

# Initial ListView population (using Desktop path from environment variable)
$desktopPath = [Environment]::GetFolderPath("Desktop")
Populate-ListView -path $desktopPath


# =====================================================================================
# =====================================================================================