# This is demo for operating excel via powershell
# The script will read sheets from specified workbook, and then try make following operations on cells:
# 1. Merge and Unmerge cells
# 2. Insert, Delete and Copy columns
# 3. Edit Cell text
# 4. Rename worksheet 

# Global Variables
$pathBaseFolder="C:\Demos\Powershell\ExcelFiles";
$pathOutputFolder="${pathBaseFolder}\Output";


if(!(Test-Path -Path ${pathOutputFolder})){ 
    echo "Creating Output Folder: ${pathOutputFolder} !";
    mkdir -p ${pathOutputFolder};
}

# Global Application for Excel
$appExcelObj=New-Object -ComObject Excel.Application;

# Enum for Cell Operations 
$xlShiftToRight = -4121;
$xlShiftToLeft = -4159;

# Read Files:
Get-ChildItem -Recurse -Name -Filter "*.xlsx" ${pathBaseFolder} | % {
    $pathSourceExcelFile = "${pathBaseFolder}\$_"
    $objSrcWorkbook = $appExcelObj.workbooks.open(${pathSourceExcelFile});
    # Handle with first worksheet
    $objWorksheet = $objSrcWorkbook.Worksheets.item(1);
    $strShowLog =  $_ + " ============= " + $objWorksheet.Name + " : " + $objWorksheet.cells.item(4,4).text
    echo $strShowLog
    # Rename the first Sheet 
    $objWorksheet.Name = "MainWork";
    # Unmerge the merged cell in D3
    $objWorksheet.Range("D3").UnMerge();

    # Insert New Column
    $range = $objWorksheet.Range("D:D").EntireColumn;
    $range.Insert($xlShiftToRight);
    $objWorksheet.cells.item(4,4) = 'DemoShow';
    $objWorksheet.cells.item(3,4) = $objWorksheet.cells.item(3,5).text;

    # Delete Column and shift right columns to left
    $range = $objWorksheet.Range("H:H").EntireColumn;
    $range.Delete($xlShiftToLeft);
    
    # Copy from one column to another
    $range = $objWorksheet.Range("E:E").EntireColumn;
    $range.copy($objWorksheet.Range("I:I").EntireColumn)
    
    # Merge Cells 
    $objWorksheet.Range("D3:G3").Merge();
    $objWorksheet.Range("H3:K3").Merge();

    # Save and close Workbook.
    $objSrcWorkbook.Save();
    $objSrcWorkbook.Close();
}

#  End of Excel Operations, Kill Excel Processes
$appExcelObj.quit()
ps excel | kill  #for some reason Excel stays