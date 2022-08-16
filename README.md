# Metadata_Extractor
A simple gui for extracting all meta and xml data from a batch of PDFs and organizing the data into a CSV.

![Capture](https://user-images.githubusercontent.com/108289013/184893025-ce1a4f9e-1c95-4b31-9d29-d3fc633a5022.PNG)

## Options ##
* **Search Directory:** Where the PDFs are stored
* **Output Location:** The folder and filename where the CSV will be stored
* **Search Subdirectories:** if true recurse through the subdirectories of the search directory otherwise stay on the first level
* **Hyperlink to PDF path:** if true output the file path as an excel formatted hyperlink rather than plaintext
* **Relative Hyperlink:** if true the hyperlink will be the relative location between the search directory and the output directory -- useful for sharing the data between file systems -- otherwise will be the full file path. 

## Usage ##
* Download [latest version](https://github.com/henrystern/Metadata-Extractor/releases/latest "releases")
* Select your options and click run to generate the CSV.

## Tips ##
* To avoid encoding errors, import the CSV into an Excel sheet using Excel's "from text/CSV" tool
* Clean up the appearance of the Excel sheet by find and replacing (CTRL+H) the string "null" to an empty string ""
