# This is demo for batch renaming files via powershell

$pathWorkFoler="C:\Demos\Powershell\DemoFiles";

cd $pathWorkFoler;

$strOriginalFileName = "MyOriginalFile.txt";
$strChangedFileName = "MyRenamedFile.txt";

# Rename Single file
Rename-Item ${strOriginalFileName} -NewName ${strChangedFileName}

# Batch Rename 

$i = 0 
Get-ChildItem -Path ${pathWorkFoler} -Filter *.pdf | % {
  $extension = $_.Extension 
  $newName = 'doc_{0:d2}{1}' -f $i, $extension
  $i++
  Rename-Item -Path $_.FullName -NewName $newName
}