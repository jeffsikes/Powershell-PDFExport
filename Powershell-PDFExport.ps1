$documents_path = 'E:\ExternalCertificatesWEBARCHIVEToPDF'

Add-type -AssemblyName office -ErrorAction SilentlyContinue
Add-Type -AssemblyName microsoft.office.interop.powerpoint -ErrorAction SilentlyContinue 

$word_app = New-Object -ComObject Word.Application
$excel_app = New-Object -ComObject Excel.Application
$powerpoint_app = New-Object -ComObject PowerPoint.Application

$powerpoint_pdf_format = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF 
$office_false = [Microsoft.Office.Core.MsoTriState]::msoFalse

# WARNING!!! Setting this to TRUE will delete the original file if the PDF Conversion is successful
$delete_original_file = $false
 

Try {
    Get-ChildItem -Path $documents_path -Include *.htm, *.html -recurse | ForEach-Object {

        $file_delete = $false

        Write-Host $_.FullName
        
        $filename = $_.FullName
        
        $filename_noextension = $_.BaseName
        
        $current_directory = $_.DirectoryName
        
        $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"
        
        $file_extension = $_.Extension.ToLower()
        
        switch($_.Extension.ToLower()) {
            {($_ -eq ".xls") -or ($_ -eq ".xlsx") -or $_ -eq ".csv" } 
                {
                    $workbook = $excel_app.Workbooks.Open($filename)        
                    $workbook.Saved = $true
                    $workbook.ExportAsFixedFormat($xlFixedFomrat::xlTypePDF,$pdf_filename)
                    $excel_app.Workbooks.Close()
                    $file_delete = $true
                }

            {($_ -eq ".one") } 
                {
                    $presentation = $powerpoint_app.Presentations.Open($filename,$office_false,$office_false,$office_false)
 
                    $presentation.SaveAs($pdf_filename, $powerpoint_pdf_format)
                    $presentation.Close()
                    $file_delete = $true
                }

            {($_ -eq ".ppt") -or ($_ -eq ".pptx") } 
                {
                    $presentation = $powerpoint_app.Presentations.Open($filename,$office_false,$office_false,$office_false)
 
                    $presentation.SaveAs($pdf_filename, $powerpoint_pdf_format)
                    $presentation.Close()
                    $file_delete = $true
                }

            {($_ -eq ".doc") -or ($_ -eq ".docx") -or ($_ -eq ".txt") -or ($_ -eq ".rtf") -or ($_ -eq ".htm") -or ($_ -eq ".html") -or ($_ -eq ".eml") -or ($_ -eq ".odt") -or ($_ -eq ".jpg")  -or ($_ -eq ".jpeg") -or ($_ -eq ".tif") -or ($_ -eq ".tiff") -or ($_ -eq ".gif") -or ($_ -eq ".png") -or ($_ -eq ".bmp")}
                {
                    if ($_ -eq ".jpg" $_ -eq ".jpeg" -or $_ -eq ".tif" -or $_ -eq ".tiff" -or $_ -eq ".png" -or $_ -eq ".bmp" ) {
                        $document = $word_app.Documents.Add()
                        $word_app.Selection.EndKey(6) | Out-Null
                        $word_app.Selection.InlineShapes.AddPicture($filename) | Out-Null
                    }
                    else {
                        $document = $word_app.Documents.Open($filename)
                    }

                    $document.SaveAs($pdf_filename, 17)

                    $document.Close(0)
                    $file_delete = $true
                }
                
            {($_ -eq ".xps")} 
                {
                    # I lucked out.  Most of my XPS were a single page.  This XpsRchVw.exe creates a PNG file for each page of the XPS, which would then have to be stitched back together...
                    $CMD = "XpsRchVw.exe"
                    $arg1 = $filename + " /o:" + $current_directory + "\" + $filename_noextension + ".png"
                    
                    #Write-Host $CMD $arg1
                    #& $CMD $arg1 $arg2 
                    
                    Start-Process -FilePath $CMD -ArgumentList $arg1 -Wait -WindowStyle Hidden

                    $image_to_delete = $current_directory + "\" + $filename_noextension + "-1.png"
                    
                    #Write-Host $image_to_delete

                    $document = $word_app.Documents.Add()
                    $word_app.Selection.EndKey(6) | Out-Null
                    $word_app.Selection.InlineShapes.AddPicture($image_to_delete) | Out-Null

                    $document.SaveAs($pdf_filename, 17)

                    $document.Close(0)
                    
                    Remove-item $image_to_delete 
                    
                    $file_delete = $true
                }
            
            default { 
                Write-Host "File Did Not Match Any Supported Types."
                
             }
        }
        
        if ($file_delete -eq $true -and $delete_original_file -eq $true) {
            Remove-item $filename 
        }
        
    }
}

Catch {
    Write-Host $_.Exception.Message
}
Finally {
    $word_app.Quit()
    $excel_app.Quit()
    $powerpoint_app.Quit()
}
    