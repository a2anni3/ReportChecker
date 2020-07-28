#Only needed when first ran
#Install-Module -Name ImportExcel -RequiredVersion 7.1.0

#imports ImportExcel module
    import-module importexcel

#path of the files, will need to copy and paste file path
$OMA = Get-ChildItem -Path C:\Users\annie\Downloads\O_export_all_$(get-date -f yyyy-MM-dd)*.xlsx -Recurse | Select-Object Name
$ENF = Get-ChildItem -Path C:\Users\annie\Downloads\E_export_all_$(get-date -f yyyy-MM-dd)*.xlsx -Recurse | Select-Object Name
$CHI = Get-ChildItem -Path C:\Users\annie\Downloads\C_export_all_$(get-date -f yyyy-MM-dd)*.xlsx -Recurse | Select-Object Name
$BEL = Get-ChildItem -Path C:\Users\annie\Downloads\B_export_all_$(get-date -f yyyy-MM-dd)*.xlsx -Recurse | Select-Object Name

$filePathOMA = "C:\Users\annie\Downloads\" + $OMA.Name
$filePathENF = "C:\Users\annie\Downloads\" + $ENF.Name
$filePathCHI = "C:\Users\annie\Downloads\" + $CHI.Name
$filePathBEL = "C:\Users\annie\Downloads\" + $BEL.Name

#name of worksheetsheet we're focusing on
$worksheetName = "Sheet5"
#row number of headers
$numHeaderRow = "1"

#gets excel data
$resultOMA = import-excel $filePathOMA -WorkSheetname $worksheetName -HeaderRow $numHeaderRow
$resultENF = import-excel $filePathENF -WorkSheetname $worksheetName -HeaderRow $numHeaderRow
$resultCHI = import-excel $filePathCHI -WorkSheetname $worksheetName -HeaderRow $numHeaderRow
$resultBEL = import-excel $filePathBEL -WorkSheetname $worksheetName -HeaderRow $numHeaderRow


#filters the data
$resultOMA | ? {$_."Free %" -le "15"} | ? {$_."Free MB" -le "10000"} | ? {$_."Disk" -cmatch '([A-Z]):\\'} 
$resultENF | ? {$_."Free %" -le "15"} | ? {$_."Free MB" -le "10000"} | ? {$_."Disk" -cmatch '([A-Z]):\\'} 
$resultCHI | ? {$_."Free %" -le "15"} | ? {$_."Free MB" -le "10000"} | ? {$_."Disk" -cmatch '([A-Z]):\\'} 
$resultBEL | ? {$_."Free %" -le "15"} | ? {$_."Free MB" -le "10000"} | ? {$_."Disk" -cmatch '([A-Z]):\\'} 