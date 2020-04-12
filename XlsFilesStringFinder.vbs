'Alternative branch where I'm migrating into VBS a later version of the project that is instead written in VBdotNET

LANGUAGE = "VBScript"

'Main commands
continueCicle = True
FSO2 = CreateObject("Scripting.FileSystemObject")
CARTELLALOCALE = FSO2.GetAbsolutePathName(".")
ELEM_ITEM = InputBox("Inserisci stringa da Trovare in questa Cartella e in tutte le sue Sotto-Cartelle", "Cerca nei File Excel")
continueCicle = (ELEM_ITEM <> "") 'Same boolean variable is used also to break a loop in the script

If continueCicle Then

    Set objExcel = CreateObject("Excel.Application")        'OGGETTO D'ANALISI
    Set ReportExcel = CreateObject("Excel.Application")     'OGGETTO DI REPORT

    ReportExcel.Application.Visible = True

    Set ReportWorkbook = ReportExcel.Workbooks.Add()
    Set RepSh = ReportExcel.Worksheets(1)

    cellNumb = 3

    RepSh.Range("$A$1").Value = "ELEMENTI TROVATI:"
    RepSh.Range("$A$1").HorizontalAlignment = -4152 'destra
    RepSh.Range("$A$1").Font.Bold = True
    RepSh.Range("$B$1").Value = 0
    RepSh.Range("$B$1").HorizontalAlignment = -4131 'sinistra
    RepSh.Range("$B$1").Font.Bold = True
    RepSh.Range("$C$1").Value = "Stringa di Input:"
    RepSh.Range("$C$1").Font.Bold = True
    RepSh.Range("$C$1").HorizontalAlignment = -4152
    RepSh.Range("$D$1").Value = ELEM_ITEM
    RepSh.Range("$D$1").Font.Bold = True
    RepSh.Range("$E$1").Value = "DataOra:"
    RepSh.Range("$E$1").Font.Bold = True
    RepSh.Range("$E$1").HorizontalAlignment = -4152
    RepSh.Range("$F$1").Value = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
    RepSh.Range("$F$1").Font.Bold = True
    RepSh.Range("$F$1").NumberFormat = "d/m/yy h.mm;@"
    RepSh.Range("A1").ColumnWidth = 20
    RepSh.Range("B1").ColumnWidth = 45
    RepSh.Range("C1").ColumnWidth = 30
    RepSh.Range("E1").ColumnWidth = 20
    RepSh.Range("$A$2").Value = "ELEMENTO PUNTATO"
    RepSh.Range("$A$2").Font.Bold = True
    RepSh.Range("$B$2").Value = "STATO"
    RepSh.Range("$B$2").Font.Bold = True
    RepSh.Range("$C$2").Value = "ESITO"
    RepSh.Range("$C$2").Font.Bold = True
    RepSh.Range("$D$2").Value = "CELLA"
    RepSh.Range("$D$2").Font.Bold = True
    RepSh.Range("$E$2").Value = "FOGLIO"
    RepSh.Range("$E$2").Font.Bold = True

    Set result0 = CreateObject("System.Collections.ArrayList")
    Set result0 = Ricors(CARTELLALOCALE, RepSh)

    RepSh.range("A1").ColumnWidth = 90
    For Each elem In result0
        RepSh.Cells(cellNumb, 1).Value = elem.Path
        RepSh.Cells(cellNumb, 2).Value = "Non Valutato"
        RepSh.Cells(cellNumb, 3).Value = "Nessuno"
        RepSh.Cells(cellNumb, 4).Value = "Nessuna"
        RepSh.Cells(cellNumb, 5).Value = "Nessuno"
        cellNumb += 1
    Next

    cellNumb = 3
    For Each elem In result0

        If continueCicle And (FSO2.FileExists(elem.Path)) Then

            On Error Resume Next

            RepSh.Cells(cellNumb, 2).Value = "Apertura in Corso .. "
            objWorkbook = objExcel.Workbooks.Open(elem.Path)

            If Err.Number = 0 Then
                RepSh.Cells(cellNumb, 2).Value = "Analisi in Corso .."
                objExcel.Application.Visible = False
                RepSh.Cells(cellNumb, 3).Value = "Negativo"
                For Each sh In objExcel.Worksheets

                    sh.Activate
                    found = sh.usedrange.Find(ELEM_ITEM, , -4123) '(what:="abc", LookIn:=xlFormulas)

                    If Not found Is Nothing Then

                        found.Select
                        RepSh.Cells(cellNumb, 3).Value = "Positivo"
                        RepSh.Cells(cellNumb, 3).Interior.Color = RGB(180, 255, 180)
                        RepSh.Cells(cellNumb, 4).Value = "" & found.address
                        RepSh.Cells(cellNumb, 4).Interior.Color = RGB(180, 255, 180)
                        RepSh.Cells(cellNumb, 5).Value = "" & sh.Name
                        RepSh.Cells(cellNumb, 5).Interior.Color = RGB(180, 255, 180)

                    End If

                Next
                objWorkbook.Close(False)
                RepSh.Cells(cellNumb, 2).Value = "Analisi Conclusa"

            Else
                RepSh.Cells(cellNumb, 2).Value = "Apertura Fallita"
                RepSh.Cells(cellNumb, 2).Interior.Color = RGB(255, 180, 180)
            End If

            On Error GoTo 0

        End If

        cellNumb = cellNumb + 1

    Next
    MsgBox("Ricerca Conclusa")
Else
    MsgBox("Ricerca Annullata")
End If
'End Main commands

Function Ricors(sPath, ByRef sheet)

    Set returncollection = CreateObject("System.Collections.ArrayList")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myFolder = FSO.GetFolder(sPath)

    For Each myFile In myFolder.Files

        nomefile = myFile.name
        If Right(nomefile, 5) = ".xlsx" Or Right(nomefile, 4) = ".xls" Then
            returncollection.add(myFile)
            sheet.Range("$B$1").Value = sheet.Range("$B$1").Value + 1
        End If

    Next

    For Each mySubFolder In myFolder.SubFolders

        For Each elem In Ricors(mySubFolder.Path, sheet)
            returncollection.add(elem)

        Next

    Next
    Return returncollection

End Function
