# This is demo for operating excel via powershell

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


# Enum for Operations in Excel
$xlShiftToRight = -4121;
$xlShiftToLeft = -4159;

# Global Application for Excel
$appExcelObj=New-Object -ComObject Excel.Application;

# Start Each Demo here

# Demo 01: Convert Excel to CSV using UTF-8:

$objSrcWorkbook = $appExcelObj.workbooks.open($pathSourceFile);

$arrSrcWorksheets = $objSrcWorkbook.Worksheets;

$arrSrcWorksheets | % {
    $sheet = $_ ;
    echo "===== Operating on " + $sheet.name ;
    $pathTargetCSVFile = ${pathOutputFolder} + "\" + ${sheet}.name + ".csv" ;
    if (Test-Path ${pathTargetCSVFile}) {  rm ${pathTargetCSVFile} }

    #[Ref] Manual for XlFileFormat: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
    ${sheet}.SaveAs(${pathTargetCSVFile}, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlUnicodeText) ; 
}

$objSrcWorkbook.Saved = $true
$objSrcWorkbook.close()

ls "$(Split-Path ${pathTargetCSVFile})\*.csv" | % { (Get-Content $_) -replace '\t',',' | Set-Content $_ -Encoding utf8 }


#  End of Excel Operations, Kill Excel Processes
$appExcelObj.quit()
ps excel | kill  #for some reason Excel stays
