# Metadata to CSV
A simple GUI for extracting all meta and XMP data from a batch of PDFs and organizing the data into a CSV.

![Capture](https://user-images.githubusercontent.com/108289013/184893025-ce1a4f9e-1c95-4b31-9d29-d3fc633a5022.PNG)

## Options ##
* **Search Directory:** Where the PDFs are stored
* **Output Location:** The folder and filename where the CSV will be stored
* **Search Subdirectories:** if true recurse through the subdirectories of the search directory otherwise stay on the first level
* **Hyperlink to PDF path:** if true output the file path as an excel formatted hyperlink rather than plaintext
* **Relative Hyperlink:** if true the hyperlink will be the relative location between the search directory and the output directory -- useful for sharing the data between file systems -- otherwise will be the full file path. 

## Usage ##
* Download the [latest version](https://github.com/henrystern/Metadata-Extractor/releases/latest "releases")
* Select your options and click run to generate the CSV.

## Run From Source ##
* Clone repository
* Open repository directory in VS Code
* Select Run > Run Without Debugging (CTRL+F5)

This is easier than installing java manually as VS Code gathers dependencies and prompts the correct installation.

## Tips ##
* To avoid encoding errors, import the CSV into an Excel sheet using Excel's "from text/CSV" tool
* After importing run this script in Excel to activate the hyperlinks:
  ```VBA
  Sub activateHyperlinks()
    
    Dim Table_Name As String
    Table_Name = "PDF_Metadata" 'replace PDF_Metadata with the name of your table
    
    Dim Column_Header As String
    Column_Header = "File Path" 'replace File Path with name of column
    
    Dim c As Range
    For Each c In Range(Table_Name & "[" & Column_Header & "]")
        c.Value = c.Value
    Next c

  End Sub
  ```
 * The hyperlinks won't work if the path is very long due to limitations in Excel. Paths can be shortened to DOS format using the FSO.ShortPath property in VBA.
