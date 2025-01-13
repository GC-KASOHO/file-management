# Function to get supported formats based on file extension
function Get-SupportedFormats {
    param (
        [string]$extension
    )
    
    $extension = $extension.ToLower()
    
    switch -Regex ($extension) {
        # Word documents
        '\.(doc|docx)$' {
            return @(
                [PSCustomObject]@{DisplayName="PDF (.pdf)"; Extension=".pdf"},
                [PSCustomObject]@{DisplayName="Text (.txt)"; Extension=".txt"},
                [PSCustomObject]@{DisplayName="Rich Text (.rtf)"; Extension=".rtf"},
                [PSCustomObject]@{DisplayName="HTML (.html)"; Extension=".html"}
            )
        }
        
        # PowerPoint presentations
        '\.(ppt|pptx)$' {
            return @(
                [PSCustomObject]@{DisplayName="PDF (.pdf)"; Extension=".pdf"},
                [PSCustomObject]@{DisplayName="PNG (.png)"; Extension=".png"},
                [PSCustomObject]@{DisplayName="JPEG (.jpg)"; Extension=".jpg"}
            )
        }
        
        # Excel workbooks
        '\.(xls|xlsx)$' {
            return @(
                [PSCustomObject]@{DisplayName="PDF (.pdf)"; Extension=".pdf"},
                [PSCustomObject]@{DisplayName="CSV (.csv)"; Extension=".csv"},
                [PSCustomObject]@{DisplayName="Text (.txt)"; Extension=".txt"}
            )
        }
        
        # Image files
        '\.(jpg|jpeg|png|gif|bmp)$' {
            return @(
                [PSCustomObject]@{DisplayName="JPEG (.jpg)"; Extension=".jpg"},
                [PSCustomObject]@{DisplayName="PNG (.png)"; Extension=".png"},
                [PSCustomObject]@{DisplayName="BMP (.bmp)"; Extension=".bmp"},
                [PSCustomObject]@{DisplayName="GIF (.gif)"; Extension=".gif"}
            )
        }
        
        # PDF documents
        '\.pdf$' {
            return @(
                [PSCustomObject]@{DisplayName="JPEG (.jpg)"; Extension=".jpg"},
                [PSCustomObject]@{DisplayName="PNG (.png)"; Extension=".png"},
                [PSCustomObject]@{DisplayName="Text (.txt)"; Extension=".txt"}
            )
        }
        
        default {
            return @()
        }
    }
}

function Show-ConversionDialog {
    param (
        [string]$sourceFile,
        [array]$formats
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Convert File"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(360, 20)
    $label.Text = "Select output format:"
    $form.Controls.Add($label)
    
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(10, 50)
    $comboBox.Size = New-Object System.Drawing.Size(360, 20)
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    
    # Create a list to store format objects
    $script:formatList = New-Object System.Collections.ArrayList
    
    foreach ($format in $formats) {
        # Add to combo box and store in our list
        [void]$comboBox.Items.Add($format.DisplayName)
        [void]$script:formatList.Add($format)
    }
    
    if ($comboBox.Items.Count -gt 0) {
        $comboBox.SelectedIndex = 0
    }
    
    $form.Controls.Add($comboBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(100, 100)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "Convert"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(200, 100)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)
    
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Return the selected format object from our list
        return $script:formatList[$comboBox.SelectedIndex]
    }
    return $null
}