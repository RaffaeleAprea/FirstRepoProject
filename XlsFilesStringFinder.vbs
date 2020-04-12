LANGUAGE="VBScript"

continueCicle = True
Set FSO2 = CreateObject("Scripting.FileSystemObject")
cartellaLocale = FSO2.GetAbsolutePathName(".")

ELEM_ITEM = InputBox("Inserisci stringa da Trovare in questa Cartella e in tutte le sue Sotto-Cartelle","Cerca nei File Excel")
continueCicle = (ELEM_ITEM<>"")
if continueCicle Then
    set result0 = CreateObject("System.Collections.ArrayList")
    set result0 = Recurse (cartellaLocale)

    Set WinScriptHost = CreateObject("WScript.Shell")

    for each elem in result0
            stringa = stringa & elem.name & VBcrlf
    next 
    msgbox stringa
    Set objExcel = CreateObject("Excel.Application")
    for each elem in result0
        'stringa = stringa & elem.name & VBcrlf
        if continuecicle and (FSO2.FileExists(elem.Path))then

            On Error Resume next
            
            Set objWorkbook = objExcel.Workbooks.Open(elem.Path)
            if err.Number = 0 then

                objExcel.Application.Visible = False
                'msgbox elem.Path
                For Each sh In objExcel.Worksheets
                    sh.Activate           
                    Set found = sh.usedrange.Find (ELEM_ITEM, , -4123) 'in VB-> "Find(what:="abc", LookIn:=xlFormulas)" bacause ':=' is not valid in VBS like it does in VB

                    If Not found Is Nothing Then
                    
                    found.Select
                    objExcel.Application.Visible = True
                    Msgbox "Trovato << " & ELEM_ITEM & " >> " & VBCRLF & VBCRLF &  "nel file" & elem.Path & VBCRLF & VBCRLF & "nella cella [" & found.address &"]"
      
                    Else
                        
                    End If
                        objWorkbook.Close
                        'Set objExcel = Nothing
                Next
                'objExcel.Application.Visible = True
            On Error Goto 0
            End if

        End if

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
