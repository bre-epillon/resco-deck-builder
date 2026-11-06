Attribute VB_Name = "Module1"
Sub Button1_Click()
'Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.PrintCommunication = False

    Dim wsReport As Worksheet
    Dim wsReport3 As Worksheet
    Dim wsReport5 As Worksheet
    Dim wbSource As Workbook
    Dim sourceFilePath As String
    Dim folderPath As String
    
    Dim startTime As Single, timeElapsed As Single, min As Long
    startTime = Timer
    
    'The full path to your external .xlsm file
    fileName = Range("C1")
    sourceFilePath = Range("C2")
    maxyear = Range("C3").Value
    from = Range("c4").Value
    tooo = Range("c5").Value
    Sheets("Exp_judgement").Range("B4:L13").ClearContents
    Sheets("Exp_judgement").Range("B17:L26").ClearContents
    Sheets("Exp_judgement").Range("B30:L39").ClearContents
    
    
    Dim myPrintRange5 As Range
    Dim myPrintRange50 As Range
    Set wsReport5 = Sheets("Exp_judgement")
    Set myPrintRange5 = wsReport5.Range("a1:m40")
    Set myPrintRange50 = wsReport5.Range("a1:m26")
    
    'Define where to save (current folder of this Excel file)
    folderPath = Range("C6")
    
    Set wbSource = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True, UpdateLinks:=0)
    DoEvents
    Set wsReport = Sheets("ResCo tables_new")
    Dim myPrintRange1 As Range
    Dim myPrintRange10 As Range
    Dim myPrintRange2 As Range
    Dim myPrintRange3 As Range

    
    Set myPrintRange1 = wsReport.Range("C4:s30")
    Set myPrintRange10 = wsReport.Range("C4:s17")
    Set myPrintRange2 = wsReport.Range("C36:w90")
    Set myPrintRange3 = wsReport.Range("y24:aa27")

    
    'Loop through each value in your list (adjust range A2:A10 as needed)
    For i = from To tooo
        Workbooks("Print tables.xlsm").Sheets("PrintPNGs").Activate
        lob = Cells(i + 7, 3).Value
        resclass = Cells(i + 7, 2).Value
        tabletoprint = Cells(i + 7, 4)
        trianglestoprint = Cells(i + 7, 5)
        profitstoprint = Cells(i + 7, 6)
        trianglesCharts = Cells(i + 7, 7)
        patternsgraphs = Cells(i + 7, 8)
        expjudg = Cells(i + 7, 9)
        expjudgrange = Cells(i + 7, 10)
        '1. Change the trigger cell (e.g., B2) to the current list value
        Workbooks(fileName).Sheets("ResCo tables_new").Activate
        Range("B1").Value = lob

        '2. Force Excel to recalculate (optional but safer)
        Application.Calculate
        
        If tabletoprint = True Then
        Workbooks(fileName).Activate
        
        Dim imgPath1 As String
        imgPath1 = folderPath & lob & "_table.png"
        
        Application.Wait (Now + TimeValue("0:00:01")) ' Give it 1 second to breathe
        If resclass = "LIIND" Or resclass = "LISHT" Or resclass = "LIMED" Or resclass = "LIMLT" Or resclass = "LITOT" Then
        Call ExportRangeAsPNG(myPrintRange10, imgPath1)
        Else
        Call ExportRangeAsPNG(myPrintRange1, imgPath1)
        End If
        
        Application.Calculate
        DoEvents
        
        End If
        
        If trianglestoprint = True Then
        Dim imgPath2 As String
        imgPath2 = folderPath & lob & "_triangle.png"
        Call ExportRangeAsPNG(myPrintRange2, imgPath2)
            
        Application.Calculate
        DoEvents
       
        End If

        If profitstoprint = True Then
        Dim imgPath3 As String
        imgPath3 = folderPath & lob & "_profit.png"
        Call ExportRangeAsPNG(myPrintRange3, imgPath3)
            
        Application.Calculate
        DoEvents
        
        End If
        
        If trianglesCharts = True Then
        Workbooks(fileName).Sheets("ResCo tables_new").Activate
        ActiveSheet.ChartObjects("Chart 8").Activate
        Dim myChart_a As ChartObject
        Set myChart_a = wsReport.ChartObjects("Chart 8")
        
        myChart_a.Chart.Export fileName:=folderPath & lob & "_ChartPrem.png", FilterName:="PNG"
       
        Application.Calculate
        DoEvents
    
        ActiveSheet.ChartObjects("Chart 9").Activate
    
        Dim myChart_b As ChartObject
        Set myChart_b = wsReport.ChartObjects("Chart 9")
                
        myChart_b.Chart.Export fileName:=folderPath & lob & "_ChartPaid.png", FilterName:="PNG"
            
        Application.Calculate
        DoEvents
        
        ActiveSheet.ChartObjects("Chart 10").Activate
        Dim myChart_c As ChartObject
        Set myChart_c = wsReport.ChartObjects("Chart 10")
        
        myChart_c.Chart.Export fileName:=folderPath & lob & "_ChartIncurr.png", FilterName:="PNG"
            
        Application.Calculate
        DoEvents
        
        End If
        
        If patternsgraphs = True Then
        Workbooks(fileName).Activate
        Sheets(resclass).Activate
        
        Set wsReport3 = Sheets(resclass)
        Dim myChart1 As Chart
        Dim myChart2 As Chart
        
        'Sheets(resclass).Activate
        Set myChart1 = wsReport3.ChartObjects("Premiums").Chart
        Set myChart2 = wsReport3.ChartObjects("Incurred").Chart
        
        ActiveSheet.ChartObjects("Premiums").Activate
        myChart1.Export fileName:=folderPath & lob & "_PremiumsBookedGraph.png", FilterName:="PNG"
            
        Application.Calculate
        DoEvents
        
        ActiveSheet.ChartObjects("Incurred").Activate
        myChart2.Export fileName:=folderPath & lob & "_IncurredGraph.png", FilterName:="PNG"
            
        Application.Calculate
        DoEvents
        
        End If
        
        Application.Wait (Now + TimeValue("0:00:01")) ' Give it 1 second to breathe
        
        If expjudg = True Then
        Workbooks(fileName).Activate
        Sheets("ExpertJudgement").Activate
        Range(Cells(expjudgrange, 5), Cells(expjudgrange + maxyear - 2017, 15)).Select
        Selection.Copy
        Workbooks("Print tables.xlsm").Sheets("Exp_judgement").Activate
        Cells(4, 2).Select
        Selection.PasteSpecial Paste:=xlValues
        
        Workbooks(fileName).Activate
        Range(Cells(expjudgrange, 17), Cells(expjudgrange + maxyear - 2017, 24)).Select
        Selection.Copy
        Workbooks("Print tables.xlsm").Sheets("Exp_judgement").Activate
        Cells(17, 2).Select
        Selection.PasteSpecial Paste:=xlValues

        Workbooks(fileName).Activate
        Range(Cells(expjudgrange, 25), Cells(expjudgrange + maxyear - 2017, 26)).Select
        Selection.Copy
        Workbooks("Print tables.xlsm").Sheets("Exp_judgement").Activate
        Cells(17, 11).Select
        Selection.PasteSpecial Paste:=xlValues
        
        Workbooks(fileName).Activate
        Range(Cells(expjudgrange, 28), Cells(expjudgrange + maxyear - 2017, 38)).Select
        Selection.Copy
        Workbooks("Print tables.xlsm").Sheets("Exp_judgement").Activate
        Cells(30, 2).Select
        Selection.PasteSpecial Paste:=xlValues

        Dim imgPath5 As String
        imgPath5 = folderPath & lob & "_expjudg.png"
        If resclass = "LIIND" Or resclass = "LISHT" Or resclass = "LIMED" Or resclass = "LIMLT" Or resclass = "LITOT" Then
        Call ExportRangeAsPNG(myPrintRange50, imgPath5)
        Else
        Call ExportRangeAsPNG(myPrintRange5, imgPath5)
        End If
            
        Application.Calculate
        DoEvents

        End If
        
        
    Next i
    
    Workbooks(fileName).Activate
    ActiveWorkbook.Close SaveChanges:=False
    
    Workbooks("Print tables.xlsm").Sheets("PrintPNGs").Activate
    Range("A1").Select
    
    Application.PrintCommunication = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    timeElapsed = Timer - startTime
    mins = Round(timeElapsed / 60, 2)
    
    MsgBox "All files have been generated in " & mins & " minutes"

End Sub
Sub Button2_Click()
Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim pptSlide As PowerPoint.Slide
    Dim folderPath As String
    Dim fileName As String
    Dim fileCount As Integer
    
    from = Range("c4").Value
    tooo = Range("c5").Value
    
    
    '--- 1. SETTINGS ---
    folderPath = Range("C6")
    
    '--- 2. INITIALIZE POWERPOINT ---
    On Error Resume Next
    Set pptApp = GetObject(class:="PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = New PowerPoint.Application
    On Error GoTo 0
    
    Set pptPres = pptApp.Presentations.Add
    pptApp.Visible = True
    
   
    For i = from To tooo
        Workbooks("Print tables.xlsm").Sheets("PrintPNGs").Activate
        lob = Cells(i + 7, 3).Value
        resclass = Cells(i + 7, 2).Value
        tabletoprint = Cells(i + 7, 4)
        trianglestoprint = Cells(i + 7, 5)
        profitstoprint = Cells(i + 7, 6)
        trianglesCharts = Cells(i + 7, 7)
        patternsgraphs = Cells(i + 7, 8)
        expjudg = Cells(i + 7, 9)
        singleslide = Cells(i + 7, 10)

        
        'Insert the PDF as an image/object
        'Note: This inserts the first page of the PDF as a picture
        If tabletoprint = True Then
        'Add a blank slide
        Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
        
            If resclass = "LIIND" Or resclass = "LISHT" Or resclass = "LIMED" Or resclass = "LIMLT" Or resclass = "LITOT" Then
            pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_table.png", _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=95, Top:=150, Width:=Application.InchesToPoints(10), Height:=Application.InchesToPoints(2.4) 'Adjust dimensions as needed
            Else
            pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_table.png", _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=95, Top:=100, Width:=Application.InchesToPoints(10), Height:=Application.InchesToPoints(4) 'Adjust dimensions as needed
            End If
            
            
        End If
        
        
        If profitstoprint = True Then
        'Add a blank slide
        If singleslide = "N" Then
        Else
            If singleslide = "Y" Then
            Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
            End If
        End If
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_profit.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=693, Top:=400, Width:=Application.InchesToPoints(1.7), Height:=Application.InchesToPoints(0.7) 'Adjust dimensions as needed
        End If
        
        
        If trianglestoprint = True Then
        
        Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
            
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_triangle.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=20, Top:=100, Width:=Application.InchesToPoints(8.8), Height:=Application.InchesToPoints(5.3) 'Adjust dimensions as needed
        End If
  
 
        If trianglesCharts = True Then
        'Add a blank slide
        If singleslide = "N" Then
        Else
            If singleslide = "Y" Then
            Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
            End If
        End If
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_ChartPrem.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=670, Top:=100, Width:=Application.InchesToPoints(3.8), Height:=Application.InchesToPoints(1.5) 'Adjust dimensions as needed
        
        'Add a blank slide
        If singleslide = "N" Then
        Else
            If singleslide = "Y" Then
            Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
            End If
        End If
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_ChartPaid.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=670, Top:=220, Width:=Application.InchesToPoints(3.8), Height:=Application.InchesToPoints(1.5) 'Adjust dimensions as needed
            
        'Add a blank slide
        If singleslide = "N" Then
        Else
            If singleslide = "Y" Then
            Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
            End If
        End If
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_ChartIncurr.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=670, Top:=340, Width:=Application.InchesToPoints(3.8), Height:=Application.InchesToPoints(1.5) 'Adjust dimensions as needed
        End If
  
        If patternsgraphs = True Then
        'Add a blank slide
        Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_PremiumsBookedGraph.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=40, Top:=100, Width:=Application.InchesToPoints(6), Height:=Application.InchesToPoints(4.5) 'Adjust dimensions as needed
        
        'Add a blank slide
        If singleslide = "N" Then
        Else
            If singleslide = "Y" Then
            Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
            End If
        End If
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
                
        pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_IncurredGraph.png", _
            LinkToFile:=msoFalse, _
            SaveWithDocument:=msoTrue, _
            Left:=500, Top:=100, Width:=Application.InchesToPoints(6), Height:=Application.InchesToPoints(4.5) 'Adjust dimensions as needed
        End If
  
        If expjudg = True Then
        'Add a blank slide
        Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
        
        DoEvents
        'Sleep 500 ' Wait for 0.5 seconds (requires a declaration, see below)
            If resclass = "LIIND" Or resclass = "LISHT" Or resclass = "LIMED" Or resclass = "LIMLT" Or resclass = "LITOT" Then
            pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_expjudg.png", _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=150, Top:=115, Width:=Application.InchesToPoints(9), Height:=Application.InchesToPoints(4.4) 'Adjust dimensions as needed
            Else
            pptSlide.Shapes.AddPicture fileName:=folderPath & lob & "_expjudg.png", _
                LinkToFile:=msoFalse, _
                SaveWithDocument:=msoTrue, _
                Left:=150, Top:=80, Width:=Application.InchesToPoints(9), Height:=Application.InchesToPoints(5.9) 'Adjust dimensions as needed
            End If
        End If
  
    Next i
    
    Dim timestamp As String
    ' Use "HH-mm" for Hours and Minutes, or "HH-mm-ss" for Seconds
    timestamp = Format(Now, "yyyy-mm-dd_HH-mm")
    
    Dim newName As String
    ' Create a name like "Report_2024-05-20.pptx"
    newName = folderPath & "Graphs_" & timestamp & ".pptx"
    
    '--- 5. SAVE AS NEW FILE ---
    pptPres.SaveAs fileName:=newName
    
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01")) ' Give it 1 second to breathe
    
    pptPres.Close

    MsgBox "PDFs have been imported into PowerPoint!"
      
    ' Close the PowerPoint Application
    pptApp.Quit
  

End Sub


Sub ExportRangeAsPNG(rng As Range, filePath As String)
Dim tmpChart As ChartObject
    
    ' 1. Ensure the directory exists
    Dim folderPath As String
    folderPath = Left(filePath, InStrRev(filePath, "/"))


    ' 2. Copy the table
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    ' 3. Create the dummy chart
    Set tmpChart = rng.Parent.ChartObjects.Add(Left:=rng.Left, Top:=rng.Top, _
                                               Width:=rng.Width, Height:=rng.Height)
    
    With tmpChart
        .Border.LineStyle = xlLineStyleNone
        .Select
        
        ' 4. Paste and wait for the image to "render"
        .Chart.Paste
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01")) ' 1 second breather
        
        ' 5. THE FIX: Attempt export with error handling
        On Error Resume Next
        .Chart.Export fileName:=filePath, FilterName:="PNG"
        
        If Err.Number <> 0 Then
            ' If it fails, try one more time with a slightly different path
            ' sometimes file locks are the culprit
            .Chart.Export fileName:=filePath, FilterName:="PNG"
        End If
        On Error GoTo 0
        
        ' 6. Clean up
        .Delete
    End With
    
    Application.CutCopyMode = False
End Sub

