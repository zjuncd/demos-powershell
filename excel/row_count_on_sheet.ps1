# This is demo for operating excel via powershell
# The script will read sheets from specified workbook, and then try make following operations on cells:
# 1. Merge and Unmerge cells
# 2. Insert, Delete and Copy columns
# 3. Edit Cell text
# 4. Rename worksheet 

# Global Variables
$pathBaseFolder="C:\Demos\Powershell\ExcelFiles";
$pathSourceFile = "${pathBaseFolder}\Demo_Source.xlsx";

if (!(Test-Path ${pathSourceFile})) { 
    echo "Unable to find the file: ${pathSourceFile} !";
    exit
}

# Global Application for Excel
$appExcelObj=New-Object -ComObject Excel.Application;

# Read Files:
$objSrcWorkbook = $appExcelObj.workbooks.open(${pathSourceFile});
# Handle with first worksheet
$objWorksheet = $objSrcWorkbook.Worksheets.item(1);
$rows = $objWorksheet.UsedRange.Rows.Count
$objSrcWorkbook.Close();
#  End of Excel Operations, Kill Excel Processes
$appExcelObj.quit()
ps excel | kill  #for some reason Excel stays