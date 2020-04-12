Language="VBScript"


continueCicle = True 'In VBS and VBA variable type declaration is sometimes not necessary
Set FSO2 = CreateObject("Scripting.FileSystemObject")
cartellaLocale = FSO2.GetAbsolutePathName(".")

ELEM_ITEM = InputBox("Inserisci stringa da Trovare in questa Cartella e in tutte le sue Sotto-Cartelle","Cerca nei File Excel")
continueCicle = (ELEM_ITEM<>"") 'Just Testing is such assignment thing works

if continueCicle Then
    Set result0 = CreateObject("System.Collections.ArrayList")
    Set result0 = Ricors (cartellaLocale)

    'Set WinScriptHost = CreateObject("WScript.Shell")  :Doesn't work for me 

    for each elem in result0
            stringa = stringa & elem.name & VBcrlf
    next 
 
    msgbox "Ricerca Conclusa"
Else
    msgbox "Ricerca Annullata"
End if

Function Ricors(sPath) 
    'Recursive function that looks for all XLSX files inside all node and leaft Sub-Folders of the Root-Folder.
    'The Root-Folder of the entire searching will be the one where the whole script is issued.
    Set returncollection = CreateObject("System.Collections.ArrayList")     
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myFolder = FSO.GetFolder(sPath)

    For Each myFile In myFolder.Files
        nomefile = MyFile.name
        if Right(nomefile,5) = ".xlsx" Or Right(nomefile,4) = ".xls" Then 'Even checking filenames, I still cannot filter Windows backup files (those starting with '~$').
            returncollection.add MyFile
        End if
    Next

    For Each mySubFolder In myFolder.SubFolders
        For Each elem in Recurse(mySubFolder.Path)
            returncollection.add elem
        Next
    Next
    Set Recurse = returncollection

End Function
