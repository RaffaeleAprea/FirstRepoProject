LANGUAGE              = "VBScript"

continueCicle         = True
Set FSO2              = CreateObject("Scripting.FileSystemObject")
CARTELLALOCALE        = FSO2.GetAbsolutePathName(".")
ELEM_ITEM             = InputBox("Inserisci stringa da Trovare in questa Cartella e in tutte le sue Sotto-Cartelle","Cerca nei File Excel")
continueCicle         = (ELEM_ITEM<>"")





if continueCicle Then
    set result0       = CreateObject("System.Collections.ArrayList")
    set result0       = Recurse (CARTELLALOCALE)

    Set WinScriptHost = CreateObject("WScript.Shell")

    'For Each elem in result0
    '        stringa = stringa & elem.name & VBCRLF
    'Next 
    'msgbox stringa

    Set objExcel      = CreateObject("Excel.Application")
    Set ReportExcel   = CreateObject("Excel.Application") 

    ReportExcel.Application.Visible = True
    Set ReportWorkbook = ReportExcel.Workbooks.Add()
    set RepSh = ReportExcel.Worksheets(1)
    cellNumb = 2

    
    RepSh.range("A1").ColumnWidth = 100
    RepSh.range("B1").ColumnWidth = 45
    RepSh.range("C1").ColumnWidth = 30
    RepSh.Range("$A$1").Value = "ELEMENTO PUNTATO"
    RepSh.Range("$A$1").Font.Bold = True
    RepSh.Range("$B$1").Value = "STATO"
    RepSh.Range("$B$1").Font.Bold = True
    RepSh.Range("$C$1").Value = "ESITO"
    RepSh.Range("$C$1").Font.Bold = True

    For Each elem in result0
        RepSh.Cells( cellNumb , 1 ).Value = elem.Path
        RepSh.Cells( cellNumb , 2 ).Value = "Non Valutato"
        RepSh.Cells( cellNumb , 3 ).Value = "Nessuno"
    cellNumb = cellNumb + 1
    Next 
   
    
    cellNumb = 2
    For Each elem in result0
        'stringa = stringa & elem.name & VBcrlf
        if continuecicle And (FSO2.FileExists(elem.Path)) Then

            RepSh.Cells( cellNumb , 2 ).Value = "Apertura in Corso .. "

            On Error Resume next
            Set objWorkbook = objExcel.Workbooks.Open(elem.Path)

            if err.Number = 0 Then
                RepSh.Cells( cellNumb , 2 ).Value = "Analisi in Corso"
                objExcel.Application.Visible = False
                RepSh.Cells( cellNumb , 3 ).Value = "Negativo"
                For Each sh In objExcel.Worksheets

                    sh.Activate           
                    Set found = sh.usedrange.Find (ELEM_ITEM, , -4123) '(what:="abc", LookIn:=xlFormulas)

                    If Not found Is Nothing Then
                    
                    found.Select
                    'Msgbox "Trovato << " & ELEM_ITEM & " >> " & VBCRLF & VBCRLF &  "nel file" & elem.Path & VBCRLF & VBCRLF & "nella cella [" & found.address &"]"
                    RepSh.Cells( cellNumb , 3 ).Value = "Positivo"
           
                    End If
                        
  
                Next
                objWorkbook.Close
                RepSh.Cells( cellNumb , 2 ).Value = "Analisi Conclusa"
            Else
                RepSh.Cells( cellNumb , 2 ).Value = "Apertura Fallita"
            End if
            
            On Error Goto 0

        End if

        cellNumb = cellNumb + 1

    Next
    Msgbox "Ricerca Conclusa"
Else
    Msgbox "Ricerca Annullata"
End if
    'msgbox stringa


Function Recurse(sPath) 
    Set returncollection = CreateObject("System.Collections.ArrayList")     
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Set myFolder = FSO.GetFolder(sPath)

    For Each myFile In myFolder.Files

        nomefile = MyFile.name
        if Right(nomefile,5) = ".xlsx" Or Right(nomefile,4) = ".xls" Then
            returncollection.add MyFile
        End if
           

        'msgbox returncollection.count
    Next

    For Each mySubFolder In myFolder.SubFolders
        
        for Each elem in Recurse(mySubFolder.Path)
            returncollection.add elem
           
        next
        
    Next
    set Recurse = returncollection

End Function