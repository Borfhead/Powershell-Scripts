$directory = "Removed"


foreach($file in Get-ChildItem $directory)
{
    $name = $file.FullName
    $name = [io.path]::GetFileNameWithoutExtension($name)
    $firstHalf, $secondHalf = $name.split("_", 2)
    
    
    if($firstHalf -match "SL$" -or $secondHalf -match "SL$")
    {
    $name = $firstHalf + "_20160622_SL.docx"
    #$name = $file.FullName
    #$name = [io.path]::GetFileNameWithoutExtension($name)
    
    }
    
    else
    {
    $name = $firstHalf + "_20160622.docx"
    }
    
    Rename-Item -path $file.FullName -newName $name
    $name
    
    
    
    
    
    
    
}