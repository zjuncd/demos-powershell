# This is demo for operating excel via powershell
# The script will read sheets from specified workbook, and then create a new workbook for each sheet and save them.

# Global Variables
$pathBaseFolder="C:\Demos\Powershell\ExcelFiles";
$pathOutputFolder="${pathBaseFolder}\Output";
$pathSourceFile = "${pathBaseFolder}\Demo_Source.xlsx";

if (!(Test-Path ${pathSourceFile})) { 
    echo "Unable to find the file: ${pathSourceFile} !";
    exit
}

if(!(Test-Path -Path ${pathOutputFolder})){ 
    echo "Creating Output Folder: ${pathOutputFolder} !";
    mkdir -p ${pathOutputFolder};
}

# Global Application for Excel
$appExcelObj=New-Object -ComObject Excel.Application;

$objSrcWorkbook = $appExcelObj.workbooks.open($pathSourceFile);

$arrSrcWorksheets = $objSrcWorkbook.Worksheets;

$arrSrcWorksheets | % {
    $objSrcWorksheet = $_ ;
    echo "===== Operating on " + $objSrcWorksheet.name ;
    
    $pathTargetWorkBookFile = $pathOutputFolder + "\" + $objSrcWorksheet.name + ".xlsx";
    if (Test-Path ${pathTargetWorkBookFile}) {  rm ${pathTargetWorkBookFile} }

    $objNewWorkbook = $appExcelObj.Workbooks.add();
    $objTargetSheet = $objNewWorkbook.worksheets.Item(1);
    $objSrcWorksheet.copy(${objTargetSheet});

    $objNewWorkbook.SaveAs(${pathTargetWorkBookFile});
    $objNewWorkbook.Close()
}
$objSrcWorkbook.Saved = $true
$objSrcWorkbook.close()

#  End of Excel Operations, Kill Excel Processes
$appExcelObj.quit()
ps excel | kill  #for some reason Excel stays
