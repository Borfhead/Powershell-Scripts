$directory = "Removed"

$objExcel = New-Object -ComObject Excel.Application             
$readFile = "Removed"
$sheetName = "Removed"
$workbook = $objExcel.Workbooks.Open($readFile)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible = $true

$word = new-object -comobject Word.Application
$word.Visible = $false


for($i = 664; $i -le 718; $i++)                             #Loops through each desired cell, reading from the Master Textbook List
{
    $courseCode = $sheet.Cells.Item($i, 2).Text             #Gets the course code, formats it into "XXX123"
    $courseCode = $courseCode.replace(" ", "")
    
    $bookTitle = $sheet.Cells.Item($i, 4).Text              #Gets book title
    
    [string]$authors = $sheet.Cells.Item($i, 5).Text                #Gets the authors and splits it into an array
    [string]$authors = $authors.split("/")
    [string]$authors = $authors.split(";")
    
    $year = $sheet.Cells.Item($i, 7).Text                   #Gets year and publishers
    $publishers = $sheet.Cells.Item($i, 9).Text
    
[array]$wordFile = Get-childitem -path $directory -recurse -include $courseCode*.docx      #Look in the directory for any file that contains
                                                                                           #the course code.  Some of them have more than 1,
                                                                                           #in which case both will be looped through and opened.
        try
        {
            foreach($thing in $wordFile)
            {
                
                $doc = $word.documents.open($directory +"\"+ $thing.name)                  #Opens the word document with the correct name.
                $selection = $word.selection                                               #Initialize selection variable of the document
                
                
                if($selection.find.execute($bookTitle))                                    #Looks for book title first.  When found, the selection
                {                                                                          #will change to the paragraph.  It then loops through
                    write "found $bookTitle"                                               #the authors array searching for each one, resetting 
                    $selection.paragraphs.first.range.select()                             #the selection to select the entire paragraph again.
                    foreach($author in $authors)
                    {
                        $selection.find.execute($author) +" Author"
                        $selection.paragraphs.range.select()
                    }
                    $selection.paragraphs.first.range.select()                             #same thing is done for year and publisher.
                    $selection.find.execute($year) +" year"
                    $selection.paragraphs.first.range.select()
                    $selection.find.execute($publishers) +" publishers"
                         
                }
                else
                {
                    write "Could not find $bookTitle in $courseCode"
                }
            
            $doc.close()     
            }  
              
        }
        
        catch
        {
        write "Couldn't open $courseCode"                                                 #Some course codes don't have catalogs.
        }
          
}





