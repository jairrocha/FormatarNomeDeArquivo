'Este Script renomeia (adiciona o ZERO na frente de arquivos que começam com número de 1 a 9)
'os arquivos que estão sobre o mesmo diretório e subdiretórios.

'-------------------------------------------------------------------------------------------------------
Main "", null
Msgbox "Pronto"

'-------------------------------------------------------------------------------------------------------

Sub Main(path, ByRef oFS)

If Isnull(oFs) then
Set oFS = CreateObject("Scripting.FileSystemObject")
path = oFS.GetParentFolderName(WScript.ScriptFullName)
End if

RenomeaFiles path, oFS	

For Each Folder in oFS.GetFolder(path).SubFolders
  Main Folder.Path, oFS	
Next



end sub

'-------------------------------------------------------------------------------------------------------

Sub RenomeaFiles(path, ByRef oFS)

For Each File in oFS.GetFolder(path).Files
 If IsNumeric(Trim(Split(File.Name," ")(0))) then
   If (Trim(Split(File.Name," ")(0)) * 1) < 10 And Left(Trim(Split(File.Name," ")(0)), 1) <> "0" Then
     File.Name = "0"& File.Name
   end if
 end if
Next

end Sub

