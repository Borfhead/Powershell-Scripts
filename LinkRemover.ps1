$word = new-object -comobject Word.Application
$word.Visible = $false
$directory = "Removed"

$ReplaceAll = 2
$findContinue = 1
$MatchCase = $false
$MatchPhrase = $false
$MatchWholeWord = $true
$MatchWildcards = $true
$MatchSoundsLike = $false
$MatchAllWordForms = $false
$Forward = $true
$Wrap = $FindContinue
$Format = $false
$deleteText = "Removed"





foreach($file in Get-ChildItem $directory)
{

$file.FullName

    try
    {
        $document = $word.Documents.open($file.FullName)
    }
    
    catch
    {
        write "Can't open it"
    }
    
    
    try
    {
        
        while($word.Selection.Find.Execute($deleteText))
            {
            $word.selection.paragraphs.first.range.select()
            $word.selection.range.delete()
            write "Deleted something"
            }    
    }   
    
    catch
    {
        write "Search Failed"
    }
    
    
    
   
    $document.Save()
    $document.Close()

    
}
