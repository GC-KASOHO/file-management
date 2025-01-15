# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint

# Load iTextSharp for PDF operations - you'll need to install this first
$iTextSharpPath = Join-Path $PSScriptRoot "itextsharp.dll"
if (Test-Path $iTextSharpPath) {
    Add-Type -Path $iTextSharpPath
} else {
    Write-Warning "iTextSharp.dll not found. Some PDF conversions may not be available."
}

function Convert-File {
    param (
        [Parameter(Mandatory=$true)]
        [string]$sourcePath,
        [Parameter(Mandatory=$true)]
        [string]$targetFormat
    )
    
    try {
        $sourceExt = [System.IO.Path]::GetExtension($sourcePath).ToLower()
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat)
        
        switch -Regex ($sourceExt) {
            # Word document conversions
            '\.(doc|docx)$' {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $doc = $word.Documents.Open($sourcePath)
                
                switch ($targetFormat) {
                    '.pdf' { $doc.SaveAs([ref]$targetPath, [ref]17) } # wdFormatPDF = 17
                    '.txt' { $doc.SaveAs([ref]$targetPath, [ref]2) }  # wdFormatText = 2
                    '.rtf' { $doc.SaveAs([ref]$targetPath, [ref]6) }  # wdFormatRTF = 6
                    '.html' { $doc.SaveAs([ref]$targetPath, [ref]8) } # wdFormatHTML = 8
                }
                
                $doc.Close()
                $word.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
            }
            
            # Excel workbook conversions
            '\.(xls|xlsx)$' {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $workbook = $excel.Workbooks.Open($sourcePath)
                
                switch ($targetFormat) {
                    '.pdf' { $workbook.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $targetPath) }
                    '.csv' { $workbook.SaveAs($targetPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV) }
                    '.txt' { $workbook.SaveAs($targetPath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlText) }
                }
                
                $workbook.Close($false)
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
            }
            
            # PowerPoint presentation conversions
            '\.(ppt|pptx)$' {
                $ppt = New-Object -ComObject PowerPoint.Application
                $presentation = $ppt.Presentations.Open($sourcePath)
                
                switch ($targetFormat) {
                    '.pdf' { $presentation.SaveAs($targetPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF) }
                    '.jpg' { $presentation.SaveAs($targetPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsJPG) }
                    '.png' { $presentation.SaveAs($targetPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPNG) }
                }
                
                $presentation.Close()
                $ppt.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt)
            }
            
            # Image conversions
            '\.(jpg|jpeg|png|gif|bmp)$' {
                $image = [System.Drawing.Image]::FromFile($sourcePath)
                
                try {
                    switch ($targetFormat) {
                        '.jpg' { $image.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Jpeg) }
                        '.png' { $image.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png) }
                        '.bmp' { $image.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Bmp) }
                        '.gif' { $image.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Gif) }
                        '.pdf' {
                            try {
                                # Create PDF document with proper page sizing
                                $imageRatio = $image.Width / $image.Height
                                $pageSize = New-Object iTextSharp.text.Rectangle(
                                    [Math]::Min(595, $image.Width), # 595 = A4 width in points
                                    [Math]::Min(842, $image.Height)  # 842 = A4 height in points
                                )
                                
                                $document = New-Object iTextSharp.text.Document($pageSize, 0, 0, 0, 0)
                                $writer = [iTextSharp.text.pdf.PdfWriter]::GetInstance(
                                    $document, 
                                    [System.IO.File]::Create($targetPath)
                                )
                                $document.Open()
                                
                                # Convert image to bytes using memory stream
                                $ms = New-Object System.IO.MemoryStream
                                try {
                                    $image.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
                                    $imageBytes = $ms.ToArray()
                                    
                                    # Create PDF image and set positioning
                                    $pdfImage = [iTextSharp.text.Image]::GetInstance($imageBytes)
                                    
                                    # Scale image to fit page while maintaining aspect ratio
                                    $pdfImage.ScaleToFit($pageSize.Width, $pageSize.Height)
                                    
                                    # Center the image on the page
                                    $pdfImage.SetAbsolutePosition(
                                        ($pageSize.Width - $pdfImage.ScaledWidth) / 2,
                                        ($pageSize.Height - $pdfImage.ScaledHeight) / 2
                                    )
                                    
                                    $document.Add($pdfImage)
                                }
                                finally {
                                    $ms.Dispose()
                                }
                            }
                            finally {
                                if ($document) {
                                    $document.Close()
                                }
                                if ($writer) {
                                    $writer.Close()
                                }
                            }
                        }
                    }
                }
                finally {
                    $image.Dispose()
                }
            }
            
            # PDF conversions
            '\.pdf$' {
                switch ($targetFormat) {
                    '.docx' {
                        $word = New-Object -ComObject Word.Application
                        
                        try {
                            # Configure Word for conversion
                            $word.Visible = $true
                            $word.DisplayAlerts = 'wdAlertsAll'
                            
                            # Get full paths
                            $sourcePath = [System.IO.Path]::GetFullPath($sourcePath)
                            $targetPath = [System.IO.Path]::GetFullPath($targetPath)
                            
                            # Create and show a form to keep focus
                            $form = New-Object System.Windows.Forms.Form
                            $form.TopMost = $true
                            $form.ShowInTaskbar = $false
                            $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
                            $form.Size = New-Object System.Drawing.Size(1, 1)
                            $form.Location = New-Object System.Drawing.Point(-100, -100)
                            $form.Show()
                            
                            # Add delay to allow Word to initialize
                            Start-Sleep -Milliseconds 500
                            
                            # Show Word application
                            $word.Activate()
                            
                            # Open document and convert
                            $doc = $word.Documents.Open($sourcePath)
                            Start-Sleep -Milliseconds 500  # Give time for document to open
                            
                            # Ensure Word window and dialogs stay in front
                            $wordProcess = Get-Process -Name "WINWORD" | Select-Object -First 1
                            if ($wordProcess) {
                                # Import Windows API functions
                                Add-Type @"
                                    using System;
                                    using System.Runtime.InteropServices;
                                    public class Win32 {
                                        [DllImport("user32.dll")]
                                        public static extern bool SetForegroundWindow(IntPtr hWnd);
                                        
                                        [DllImport("user32.dll")]
                                        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
                                    }
"@
                                # Force Word window to front
                                [Win32]::SetForegroundWindow($wordProcess.MainWindowHandle)
                            }
                            
                            # SaveAs2 parameters: 16 = wdFormatDocx
                            $doc.SaveAs2($targetPath, 16)
                            $doc.Close()
                            
                            # Close the temporary form
                            $form.Close()
                            $form.Dispose()
                        }
                        catch {
                            Write-Error "PDF conversion failed: $_"
                            throw
                        }
                        finally {
                            if ($word) {
                                # Ensure Word is properly closed
                                $word.Quit()
                                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                                [System.GC]::Collect()
                                [System.GC]::WaitForPendingFinalizers()
                            }
                        }
                    }
                }
            }
            
            default {
                throw "Unsupported file format: $sourceExt"
            }
        }
        
        Write-Host "File converted successfully: $targetPath"
        return $true
    }
    catch {
        Write-Error "Conversion failed: $_"
        return $false
    }
    finally {
        # Clean up any remaining COM objects
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

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
                [PSCustomObject]@{DisplayName="Word Document (.docx)"; Extension=".docx"}
            )
        }
        
        default {
            return @()
        }
    }
}

function Show-FormatSelectionDialog {
    param (
        [Parameter(Mandatory=$true)]
        [string]$filePath
    )
    
    # Get the file extension and supported formats
    $extension = [System.IO.Path]::GetExtension($filePath)
    $formats = Get-SupportedFormats -extension $extension
    
    # Check if file type is supported
    if ($formats.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "This file type is not supported for conversion.",
            "Unsupported File Type",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $null
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Conversion Format"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterScreen"
    
    # Add file name label
    $fileLabel = New-Object System.Windows.Forms.Label
    $fileLabel.Location = New-Object System.Drawing.Point(10, 20)
    $fileLabel.Size = New-Object System.Drawing.Size(360, 20)
    $fileLabel.Text = "File: $([System.IO.Path]::GetFileName($filePath))"
    $form.Controls.Add($fileLabel)
    
    $formatLabel = New-Object System.Windows.Forms.Label
    $formatLabel.Location = New-Object System.Drawing.Point(10, 50)
    $formatLabel.Size = New-Object System.Drawing.Size(360, 20)
    $formatLabel.Text = "Select output format:"
    $form.Controls.Add($formatLabel)
    
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(10, 80)
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
    $okButton.Location = New-Object System.Drawing.Point(100, 120)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.Text = "Convert"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(200, 120)
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