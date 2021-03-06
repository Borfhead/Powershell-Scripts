$fileDialog = New-Object System.Windows.Forms.OpenFileDialog   #user friendly Dialog box to open the file path
$fileDialog.ShowDialog() | Out-Null
$file = $fileDialog.filename

$sheetName = Read-Host -Prompt "What is the name of the sheet?"   #prompt user for sheetname

$objExcel = New-Object -ComObject Excel.Application              #Creates a new Excel object, opening the desired file.
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible = $true


$objOutlook = New-Object -ComObject Outlook.Application                  #Creates a new Outlook object
$entries = $objOutlook.session.GetGlobalAddressList().AddressEntries


[int]$col = Read-Host -Prompt "What is the column number you'd like read? (Please enter an integer)"      #Set up variables for cells to read from
[int]$row = Read-Host -Prompt "What is the starting row? (Please enter an integer)"
[int]$rowmax = Read-Host -Prompt "What is the ending row? (Please enter an integer)"




for($i = $row; $i -le $rowmax; $i++)
{
    $cellText = $sheet.Cells.Item($i, $col).text     #Gets the name from the cell

    $splitName = $cellText.Split(",");               #Changes the name from "Last, First" to "First Last"
    $firstName = $splitName[1].Trim()
    $lastName = $splitName[0].Trim()
    $fullName = $firstName + " " + $lastName

        if($entries.Item($fullName).Name -eq $fullName)
        {
            #$entries.Item($fullName).Name + ":  " +$entries.Item($fullName).GetExchangeUser().OfficeLocation   #show up in debug window
            
            $sheet.Cells.Item($i, 3) = $entries.Item($fullName).GetExchangeUser().OfficeLocation     #update excel cells with location information
            
        }

        else
        {
            $fullName + ":  " 
            
            $sheet.Cells.Item($i, 3) = " "
        }
        
    

}