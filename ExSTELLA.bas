Attribute VB_Name = "ExSTELLA"
Option Explicit

' ==============================================================================
' ExSTELLA Processing Module
' Description:
'   Processes STELLA spectrometer data from an Exchel data set by collecting
'   white card and plant sample reflectance readings, calculating  reflectance
'   ratios, outputting results to a new sheet, and generating a chart.
' ==============================================================================

' --- Module-level worksheet settings ---
Private wavelengthMin As Long     ' Minimum wavelength in nm
Private wavelengthMax As Long     ' Maximum wavelength in nm
Private numSamples As Long        ' Number of rows per sample block
Private hasHeader As Boolean      ' Whether data has a header row

' ------------------------------------------------------------------------------
' Sub: SetVars
' Purpose: Initialize module-level worksheet constants.
' ------------------------------------------------------------------------------
Sub SetVars()
    wavelengthMin = 410
    wavelengthMax = 940
    numSamples = 18
    hasHeader = True
End Sub

' ------------------------------------------------------------------------------
' Sub: DeleteEmptyRows
' Purpose: Removes empty rows from the active worksheet. Use this to prep the
' data if the Excel or CSV file has blank rows.
' ------------------------------------------------------------------------------
Sub DeleteEmptyRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim isEmpty As Boolean

    Set ws = ActiveSheet
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' Loop from bottom to top to avoid skipping rows
    For r = lastRow To 1 Step -1
        isEmpty = WorksheetFunction.CountA(ws.Rows(r)) = 0
        If isEmpty Then
            ws.Rows(r).Delete
        End If
    Next r

    Application.ScreenUpdating = True
    MsgBox "Empty rows deleted.", vbInformation
End Sub

' ------------------------------------------------------------------------------
' Function: PromptForWorksheet
' Purpose: Prompts the user to choose a worksheet and returns the selection.
' Returns: Worksheet object or Nothing if canceled.
' ------------------------------------------------------------------------------
Function PromptForWorksheet() As Worksheet
    Dim wsName As String
    Dim wsNumber As Variant
    Dim ws As Worksheet
    Dim wsList As String
    Dim found As Boolean
    Dim wsResult As Worksheet
    Dim i As Long
    
    ' Check if there is only 1 sheet
    If ActiveWorkbook.Worksheets.count = 1 Then
        Set PromptForWorksheet = ActiveWorkbook.Worksheets(1)
        Exit Function
    End If

    ' Create a list of available sheets
    i = 1
    'MsgBox "#sheets " & ThisWorkbook.Worksheets.Count
    
    For Each ws In ActiveWorkbook.Worksheets
        wsList = wsList & i & ") " & ws.Name & vbNewLine
        i = i + 1
    Next ws

    ' Prompt user
    Do
        wsNumber = InputBox("Enter the number of the worksheet to use from the list below:" & vbNewLine & wsList, "Select Worksheet")

        If wsNumber = "" Then
            Set PromptForWorksheet = Nothing
            Exit Function
        End If

        On Error Resume Next
        wsNumber = CLng(wsNumber)
        ' MsgBox wsNumber & " " & ThisWorkbook.Worksheets(wsNumber).Name
        wsName = ActiveWorkbook.Worksheets(wsNumber).Name
        Set wsResult = ActiveWorkbook.Worksheets(wsName)
        On Error GoTo 0

        If Not wsResult Is Nothing Then
            wsResult.Activate ' Switch to selcted sheet
            Set PromptForWorksheet = wsResult
            Exit Function
        Else
            MsgBox "Worksheet '" & wsName & "' not found. Please try again.", vbExclamation
        End If
    Loop
End Function

' ------------------------------------------------------------------------------
' Function: AverageFromCollection
' Purpose: Computes the average of numeric values in a collection.
' Returns: Double (average value)
' ------------------------------------------------------------------------------
Function AverageFromCollection(col As Collection) As Double
    Dim total As Double
    Dim count As Long
    Dim val As Variant

    total = 0
    count = 0

    For Each val In col
        If IsNumeric(val) Then
            total = total + CDbl(val)
            count = count + 1
        End If
    Next val

    If count > 0 Then
        AverageFromCollection = total / count
    Else
        AverageFromCollection = 0
    End If
End Function

' ------------------------------------------------------------------------------
' Sub: HighlightSets
' Purpose: Highlights rows with the minimum wavelength value to mark the start
' of each data set.
' ------------------------------------------------------------------------------
Sub HighlightSets()
    ' Highlight start of set
    SetVars
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim i As Long
    Dim lastRow As Long
    Dim highlight As VbMsgBoxResult
    
    highlight = MsgBox("Do you want to highlight the start of each sample set (rows with " & wavelengthMin & "nm wavelength)?", vbYesNo)
    If highlight = vbYes Then
        lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).Row
        Dim dataStart As Long
        If hasHeader Then
            dataStart = 2
        Else
            dataStart = 1
        End If
        
        For i = dataStart To lastRow
            If ws.Cells(i, "H").value = wavelengthMin Then
                ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
            End If
        Next i
        
        MsgBox "Rows with wavelength_nm = " & wavelengthMin & " have been highlighted.", vbInformation
    End If
End Sub

' ------------------------------------------------------------------------------
' Function: SortDictionaryKeys
' Purpose: Sorts dictionary keys numerically.
' Returns: Variant array of sorted keys.
' ------------------------------------------------------------------------------
Function SortDictionaryKeys(dict As Object) As Variant
    Dim keys() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant

    keys = dict.keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If val(keys(i)) > val(keys(j)) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    SortDictionaryKeys = keys
End Function

' ------------------------------------------------------------------------------
' Function: CollectRanges
' Purpose: Prompts user to select blocks of sample data (white card or plant
' samples).
' Parameters:
'   - setName: Label for the range being collected.
' Returns: Collection of selected ranges.
' ------------------------------------------------------------------------------
Function CollectRanges(setName As String) As Collection
    SetVars
    Dim sampleRanges As New Collection
    Dim startRange As Range
    Dim response As VbMsgBoxResult

    Do
        On Error Resume Next
        Set startRange = Application.InputBox("Select the full " & numSamples & "-row range for a " & setName & " set", _
                                             "Select " & setName & " Data", Type:=8)
        On Error GoTo 0

        If startRange Is Nothing Then
            MsgBox "No range selected. Exiting.", vbExclamation
            Exit Function
        End If
        
        If startRange.Rows.count <> numSamples Then
            MsgBox "You must select exactly " & numSamples & " rows for each " & setName & " set.", vbExclamation
        Else
            sampleRanges.Add startRange
            response = MsgBox("Do you want to add another " & setName & " set?", vbYesNo)
            If response = vbNo Then Exit Do
        End If
    Loop
    Set CollectRanges = sampleRanges
End Function

' ------------------------------------------------------------------------------
' Function: CleanDictionary
' Purpose: Resets the wavelength dictionary to hold empty collections. This
' function mantains the existing keys (wavelengths).
' Returns: Updated dictionary object.
' ------------------------------------------------------------------------------
Function CleanDictionary(dict As Object)
    Dim key As Variant
    For Each key In dict.keys
        Set dict(key) = New Collection ' replace collection to hold raw counts
        dict(key).Add New Collection ' add collection for whitecards
        dict(key).Add New Collection ' add collection for plantsamples
    Next key
    Set CleanDictionary = dict
End Function

' ------------------------------------------------------------------------------
' Function: AddToWavelengthDict
' Purpose: Populates the dictionary with raw_counts for each wavelength.
' Parameters:
'   - dict: The dictionary to populate
'   - sampleRange: A Collection of selected ranges
'   - rangeName: Descriptive name for user messages
'   - index: 1 = white card, 2 = plant sample
'   - ws: Worksheet from which data is extracted
' Returns: Updated dictionary object
' ------------------------------------------------------------------------------
Function AddToWavelengthDict(dict As Object, sampleRange As Collection, rangeName As String, index As Long, ws As Worksheet)
    
    Dim r As Range
    Dim cellRow As Range
    Dim wl As Variant, rc As Variant
    Dim key As Variant
    Dim x As Long
    

    For Each r In sampleRange
        For Each cellRow In r.Rows
            wl = ws.Cells(cellRow.Row, 8).value ' current wavelength_nm
            rc = ws.Cells(cellRow.Row, 11).value ' current raw_counts
                        
            If IsNumeric(wl) And IsNumeric(rc) Then
                key = wl
                If Not dict.exists(key) Then 'when the current wavelength is not in the dictionary
                    dict.Add key, New Collection ' add wl & collection to hold raw counts
                    dict(key).Add New Collection ' add collection for whitecards
                    dict(key).Add New Collection ' add collection for plantsamples
                End If
                dict(key).Item(index).Add rc ' add the current raw count to item at index
            End If
        Next cellRow
    Next r

    If dict.count = 0 Then
        MsgBox "No valid data was found in selected " & rangeName & " blocks.", vbCritical
        Exit Function
    End If
    

    Set AddToWavelengthDict = dict
End Function


' ------------------------------------------------------------------------------
' Sub: ProcessSTELLAData
' Purpose: Main sub for processing data, collecting inputs, calculating
' reflectance ratios, outputting results, and creating charts.
' ------------------------------------------------------------------------------
Sub ProcessSTELLAData()
    SetVars
    Dim ws As Worksheet
    Set ws = PromptForWorksheet()
    If ws Is Nothing Then Exit Sub
    
    ' Dictionary to store wavelength => collection(collection whitecards, collection plantsamples)
    ' whitecard is index 1 & plantsample is index 2, 3, etc.
    Dim WavelengthDict As Object
    Set WavelengthDict = CreateObject("Scripting.Dictionary")
    
    Dim plantCount As Long
    plantCount = 0
    Dim plantList As New Collection
    Dim wsOut As Worksheet
    Set wsOut = Worksheets.Add(After:=Worksheets(Worksheets.count))
    
    Do
        ws.Activate ' switch back to main sheet
        Dim plantName As String ' put this in a loop if we want to plot mutiple plants
        plantName = InputBox("Enter the Plant Name for this sample set:", "Plant Name")
        If Trim(plantName) = "" Then
            MsgBox "Plant Name is required. Exiting.", vbExclamation
            Exit Sub
        Else
            plantCount = plantCount + 1
            plantList.Add (plantName)
        End If
    
        Dim whiteCardRanges As Collection
        Set whiteCardRanges = CollectRanges("White Cards")
        
        Dim plantSampleRanges As Collection
        Set plantSampleRanges = CollectRanges("Plant Samples")
    
        Set WavelengthDict = AddToWavelengthDict(WavelengthDict, whiteCardRanges, "White Cards", 1, ws)
        Set WavelengthDict = AddToWavelengthDict(WavelengthDict, plantSampleRanges, "Plant Samples", 2, ws)
        
        Dim outputRow As Long
        outputRow = 2 'start in row 2 (1 is for header)
        
        Dim key As Variant
        
        If plantCount = 1 Then ' fill first column with wavelengths
            wsOut.Cells(1, 1).value = "Wavelength"
                For Each key In WavelengthDict.keys
                    wsOut.Cells(outputRow, 1).value = key
                    outputRow = outputRow + 1
                Next key
        End If
        
        wsOut.Cells(1, plantCount + 1).value = plantName 'header for plant

        outputRow = 2 'start in row 2 (1 is for header)
        For Each key In WavelengthDict.keys
            wsOut.Cells(outputRow, plantCount + 1).value = AverageFromCollection(WavelengthDict(key).Item(2)) / AverageFromCollection(WavelengthDict(key).Item(1))
            outputRow = outputRow + 1
        Next key
                
        Set WavelengthDict = CleanDictionary(WavelengthDict)
                
        Dim response As VbMsgBoxResult
        response = MsgBox("Do you want to add another sample set?", vbYesNo)
        If response = vbNo Then Exit Do
        
    Loop
    
    Dim plantListOut As String
    Dim p As Variant
    For Each p In plantList
        plantListOut = plantListOut & " " & p
    Next p
    plantListOut = Left(plantListOut, 20) ' trim string length to avoid name maxiumum
    On Error Resume Next
    wsOut.Name = "Reflectance" & plantListOut
    On Error GoTo 0
    wsOut.Columns.AutoFit
    
    MsgBox "Done! Reflectance table created on new sheet: " & wsOut.Name, vbInformation
    
    Dim chartObj As ChartObject
    
    Dim lastRow As Long
    lastRow = wsOut.Cells(wsOut.Rows.count, "A").End(xlUp).Row
    Dim lastCol As Long

    Set chartObj = wsOut.ChartObjects.Add(Left:=300, Width:=500, Top:=50, Height:=300)

    With chartObj.Chart
        .ChartType = xlXYScatterSmooth
        
        
        Dim s As Variant
        For s = 1 To plantCount
            .SeriesCollection.NewSeries
            .SeriesCollection(s).XValues = wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(lastRow, 1))
            .SeriesCollection(s).Values = wsOut.Range(wsOut.Cells(2, s + 1), wsOut.Cells(lastRow, s + 1))
            .SeriesCollection(s).Name = wsOut.Cells(1, s + 1).value
        Next s
        
        .HasTitle = True
        .ChartTitle.Text = "Reflectance Ratio vs. Wavelength"

        With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Reflectance Ratio"
            .MinimumScale = 0 ' set y min to 0
        End With
        
        With .Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Wavelength (nm)"
            .MinimumScale = wavelengthMin ' Set X-axis bounds
            .MaximumScale = wavelengthMax
        End With
    End With
    
    wsOut.Activate ' switch to the new sheet with data & chart
    
End Sub
