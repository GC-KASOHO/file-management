# Add Windows Forms assembly for context menu functionality
Add-Type -AssemblyName System.Windows.Forms

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
            # Navigate to directory - you'll need to implement this function
            Set-CurrentDirectory $Path
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
        # Store the path for later paste operation
        $script:cutPath = $Path
        $script:isCut = $true
        Write-Host "Item marked for moving: $Path"
    })
    
    # Paste (only show if in directory view)
    if ($isDirectory) {
        $pasteMenuItem = $contextMenu.Items.Add("Paste")
        $pasteMenuItem.Add_Click({
            if ($script:cutPath) {
                if ($script:isCut) {
                    # Move item
                    Move-Item $script:cutPath $Path
                    $script:cutPath = $null
                    $script:isCut = $false
                }
            } elseif ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
                # Handle files from clipboard
                $files = [System.Windows.Forms.Clipboard]::GetFileDropList()
                foreach ($file in $files) {
                    Copy-Item $file $Path
                }
            }
            # Refresh the view - you'll need to implement this function
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
            # Refresh the view - you'll need to implement this function
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
            # Refresh the view - you'll need to implement this function
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

# Example usage in your file explorer form:
$listView.Add_MouseClick({
    param($sender, $e)
    
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $item = $listView.GetItemAt($e.X, $e.Y)
        if ($item) {
            Show-ContextMenu -Path $item.Tag -Control $listView -MouseEvent $e
        }
    }
})