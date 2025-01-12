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

# Variable to hold the path of the item to be cut or copied
$global:clipboardPath = $null
$global:isCut = $false

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
    # Check if clipboardPath is set and currentPath is valid
    if ($global:clipboardPath -ne $null -and (Test-Path -Path $global:currentPath)) {
        # Get the destination path by combining the current path with the name of the item to paste
        $destinationPath = Join-Path -Path $global:currentPath -ChildPath (Get-Item $global:clipboardPath).Name
        
        try {
            if ($global:isCut) {
                # Move the item if it was cut
                Move-Item -Path $global:clipboardPath -Destination $destinationPath -Force
                [System.Windows.Forms.MessageBox]::Show("Moved to: $destinationPath")
            } else {
                # Copy the item if it was copied
                Copy-Item -Path $global:clipboardPath -Destination $destinationPath -Force
                [System.Windows.Forms.MessageBox]::Show("Copied to: $destinationPath")
            }
            # Refresh the ListView to show the updated contents
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

# Show the form
[void]$form.ShowDialog()