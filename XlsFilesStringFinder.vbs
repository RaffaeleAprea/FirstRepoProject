LANGUAGE              = "VBScript"

continueCicle         = True
Set FSO2              = CreateObject("Scripting.FileSystemObject")
CARTELLALOCALE        = FSO2.GetAbsolutePathName(".")
ELEM_ITEM             = InputBox("Inserisci stringa da Trovare in questa Cartella e in tutte le sue Sotto-Cartelle","Cerca nei File Excel")
continueCicle         = (ELEM_ITEM<>"")

If continueCicle Then

    Set objExcel      = CreateObject("Excel.Application")
    Set ReportExcel   = CreateObject("Excel.Application") 

    ReportExcel.Application.Visible = True
    Set ReportWorkbook = ReportExcel.Workbooks.Add()
    set RepSh = ReportExcel.Worksheets(1)
    cellNumb = 3

    RepSh.Range("$A$1").Value = "ELEMENTI TROVATI"
    RepSh.Range("$A$1").Font.Bold = True
    RepSh.Range("$B$1").Value = 0
    RepSh.Range("$B$1").Font.Bold = True
    RepSh.Range("$C$1").Value = "Stringa di Input"
    RepSh.Range("$C$1").Font.Bold = True
    RepSh.Range("$D$1").Value = ELEM_ITEM
    RepSh.Range("$D$1").Font.Bold = True
    RepSh.Range("$C$1").Value = "DataOra"
    RepSh.Range("$C$1").Font.Bold = True
    RepSh.Range("$D$1").Value = ELEM_ITEM
    RepSh.Range("$D$1").Font.Bold = True
    RepSh.range("A1").ColumnWidth = 20
    RepSh.range("B1").ColumnWidth = 45
    RepSh.range("C1").ColumnWidth = 30
    RepSh.Range("$A$2").Value = "ELEMENTO PUNTATO"
    RepSh.Range("$A$2").Font.Bold = True
    RepSh.Range("$B$2").Value = "STATO"
    RepSh.Range("$B$2").Font.Bold = True
    RepSh.Range("$C$2").Value = "ESITO"
    RepSh.Range("$C$2").Font.Bold = True

    set result0       = CreateObject("System.Collections.ArrayList")
    set result0       = Ricors (CARTELLALOCALE , RepSh )
    RepSh.range("A1").ColumnWidth = 100
    For Each elem in result0
        RepSh.Cells( cellNumb , 1 ).Value = elem.Path
        RepSh.Cells( cellNumb , 2 ).Value = "Non Valutato"
        RepSh.Cells( cellNumb , 3 ).Value = "Nessuno"
    cellNumb = cellNumb + 1
    Next 
   
    
    cellNumb = 3
    For Each elem in result0
        'stringa = stringa & elem.name & VBcrlf
        if continuecicle And (FSO2.FileExists(elem.Path)) Then

            On Error Resume next
            
            RepSh.Cells( cellNumb , 2 ).Value = "Apertura in Corso .. "
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
    msgbox "Ricerca Conclusa"
Else
    msgbox "Ricerca Annullata"
End if

Function Ricors(sPath , ByRef sheet) 
    Set returncollection = CreateObject("System.Collections.ArrayList")     
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myFolder = FSO.GetFolder(sPath)

    For Each myFile In myFolder.Files

        nomefile = MyFile.name
        if Right(nomefile,5) = ".xlsx" Or Right(nomefile,4) = ".xls" Then
            returncollection.add MyFile
            sheet.Range("$B$1").Value = sheet.Range("$B$1").Value + 1
        End if

    Next

    For Each mySubFolder In myFolder.SubFolders
        
        for Each elem in Recurse(mySubFolder.Path , sheet)
            returncollection.add elem
           
        next
        
    Next
    set Recurse = returncollection

End Function