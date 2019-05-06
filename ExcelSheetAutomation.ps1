# Function to get file location
Function Get-OpenFile($initialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFile = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFile.initialDirectory = $initialDirectory
    $OpenFile.Title = "Choose report to open"
    $OpenFile.filter = "All files (*.xls*)| *.xls*"
    $OpenFile.DefaultExt = "xls"
    $OpenFile.AddExtension = $True
    if($openFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){    
        $OpenFile.filename
    }
    else{
        exit
    }
}

# Function get save location
Function Get-SaveFile($initialDirectory){
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $SaveFile = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFile.initialDirectory = $initialDirectory
    $SaveFile.Title = "Choose where to save file"
    $SaveFile.filter = "Excel Files (*.xls*)| *.xls*"
    $SaveFile.DefaultExt = "xls"
    $SaveFile.AddExtension = $True
    if($SaveFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){    
        $SaveFile.filename
    }
    else{
        exit
    }
}

#-------------Necessary Values for Pasting----------------------
$xlPasteValues             = -4163
$xlPasteAll                = -4104
$xlPasteFormats            = -4122
$xlShiftToRight = -4161

#-------------Necessary Values for Saving documents-------------
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook

$xclpath = Get-OpenFile $xclpath
$xcl = New-Object -ComObject Excel.Application
$xcl.Visible = $True
$xclWorkBook = $xcl.Workbooks.Open($xclPath)
$xclworksheet = $xcl.worksheets.item(1)

# Deletes the first couple of rows that are not needed
for($i = 0; $i -le 4; $i++){
    $xclworksheet.Cells.Item(1,1).EntireRow.Delete()
}
$maxR = $xclworksheet.UsedRange.Rows.Count

$xclworksheet.Cells.Item($maxR,1).EntireRow.Delete()

$maxR = $xclworksheet.UsedRange.Rows.Count

$xclworksheet.Cells.Item($maxR,1).EntireRow.Delete()

# Inserts  Patient Suffix column
$xclworksheet.Range("D:D").Insert($xlShiftToRight)
$xclworksheet.Range("D1:D1").Value() = "Patient Suffix"
$xclworksheet.Range("H1:H1").Value() = "Payer"
$xclworksheet.Range("M1:M1").Value() = "Provider"
$xclworksheet.Range("N1:N1").Value() = "Tax-id"
$xclworksheet.Range("O1:O1").Value() = "Npi"
$xclworksheet.Range("P1:P1").Value() = "Site-id"

$maxR = $xclworksheet.UsedRange.Rows.Count

# Sets values for columns M through P
$xclworksheet.Range("M2:M$maxR").Value() = "BOC"
$xclworksheet.Range("N2:N$maxR").Value() = #set taxid
$xclworksheet.Range("O2:O$maxR").Value() = #set npi
$xclworksheet.Range("P2:P$maxR").Value() = #set ID

# Sets the payer column to the appropriate value depending on the hospital
if($xclpath.ToLower().Contains("aetna")){
    $xclworksheet.Range("H2:H$maxR").value() = 20
}
elseif($xclpath.ToLower().Contains("uhc")){
    $xclworksheet.Range("H2:H$maxR").value() = 19
}
elseif($xclpath.ToLower().Contains("humana")){
    $xclworksheet.Range("H2:H$maxR").value() = 339
}
elseif($xclpath.ToLower().Contains("cross")){
    $xclworksheet.Range("H2:H$maxR").value() = 335
}

$xcl.displayAlerts = $False
$xclWorkBook.SaveAs($xclpath, $xlFixedFormat)
$xclWorkBook.Close()
$xcl.Quit()
[Runtime.Interopservices.Marshal]::ReleaseComObject($xcl)