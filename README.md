# Office to PDF Conversion Script

A PowerShell script for bulk converting Microsoft Office documents and other file types to PDF format.

## Overview

This script provides a robust solution for converting various file types to PDF format using Microsoft Office automation. It handles Word documents, Excel spreadsheets, PowerPoint presentations, images, and other text-based files.

## Features

- **Multi-format support**: Converts a wide range of file formats including:
  - Word documents (.doc, .docx)
  - Excel spreadsheets (.xls, .xlsx, .csv)
  - PowerPoint presentations (.ppt, .pptx)
  - Text files (.txt, .rtf)
  - Web pages (.htm, .html)
  - Images (.jpg, .jpeg, .png, .gif, .tif, .tiff, .bmp)

- **Batch processing**: Convert entire folders of documents at once
  
- **Subfolder support**: Option to recursively process files in subfolders

- **Custom output location**: Specify a separate folder for the converted PDF files

- **Selective conversion**: Option to delete original files after successful conversion

- **Skip existing**: Automatically skips files that have already been converted to avoid duplication

- **Memory management**: Restarts Office COM objects periodically to prevent crashes during large batch jobs

- **Detailed logging**: Color-coded console output shows conversion progress and results

## Requirements

- Windows operating system
- PowerShell 3.0 or higher
- Microsoft Office installed (Word, Excel, and PowerPoint)

## Installation

1. Save the script file as `Convert-OfficeToPDF.ps1` to your desired location
2. Open PowerShell with administrator privileges
3. Navigate to the folder containing the script
4. You may need to set the execution policy to run the script:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

## Usage

### Basic Usage

Convert all supported files in a folder:

```powershell
.\Convert-OfficeToPDF.ps1 -FolderPath "C:\Documents"
```

### Advanced Usage

Convert files with various options:

```powershell
.\Convert-OfficeToPDF.ps1 -FolderPath "C:\Documents" -OutputPath "C:\PDFs" -IncludeSubfolders -DeleteOriginal
```

### Function Import

You can also import the function into your PowerShell session:

```powershell
. .\Convert-OfficeToPDF.ps1
Convert-OfficeToPDF -FolderPath "C:\Documents" -OutputPath "C:\PDFs"
```

### Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| FolderPath | String | Yes | Path to the folder containing files to convert |
| OutputPath | String | No | Path to save the PDF files (defaults to input folder) |
| DeleteOriginal | Switch | No | If specified, original files will be deleted after successful conversion |
| IncludeSubfolders | Switch | No | If specified, files in subfolders will also be processed |
| FileTypes | String[] | No | Array of file extensions to process (defaults to all supported types) |

## Examples

### Example 1: Basic Conversion

Convert all supported files in a folder to PDF, saving them in the same location:

```powershell
Convert-OfficeToPDF -FolderPath "E:\Documents"
```

### Example 2: Process Subfolders and Save to Different Location

Convert all supported files, including those in subfolders, and save PDFs to a different location:

```powershell
Convert-OfficeToPDF -FolderPath "E:\Documents" -OutputPath "E:\PDFs" -IncludeSubfolders
```

### Example 3: Convert and Delete Originals

Convert only Word documents and delete the originals after successful conversion:

```powershell
Convert-OfficeToPDF -FolderPath "E:\Documents" -FileTypes @('.doc','.docx') -DeleteOriginal
```

## Troubleshooting

- **Files not converting**: Ensure Microsoft Office is installed and functioning correctly
- **'Access denied' errors**: Run PowerShell as Administrator
- **COM errors**: Make sure no Office applications are running when executing the script
- **Memory issues**: Try converting smaller batches of files at a time

## Limitations

- Requires Microsoft Office to be installed
- May not preserve all formatting perfectly in complex documents
- Performance depends on your system resources and the complexity of documents

## License

This script is released under the MIT License.

## Acknowledgements

This script was created to simplify the process of converting large collections of documents to PDF format automatically.
