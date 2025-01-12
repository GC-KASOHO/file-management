
# Create a ContextMenuStrip
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Add menu items
$menuItemOpen = $contextMenu.Items.Add("Open")
$menuItemCopy = $contextMenu.Items.Add("Copy")
$menuItemDelete = $contextMenu.Items.Add("Delete")
$menuItemProperties = $contextMenu.Items.Add("Properties")

# Event handler for context menu items
$menuItemOpen.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            Start-Process $itemPath
        }
    }
})

$menuItemCopy.Add_Click({
    if ($listView.SelectedItems.Count -gt 0) {
        $itemPath = $listView.SelectedItems[0].Tag
        if (Test-Path -Path $itemPath) {
            [System.Windows.Forms.Clipboard]::SetText($itemPath)
            [System.Windows.Forms.MessageBox]::Show("Copied to clipboard: $itemPath")
        }
    }
})

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