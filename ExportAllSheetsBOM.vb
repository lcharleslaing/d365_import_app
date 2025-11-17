Sub Main()
    Dim doc = ThisDoc.Document
    Dim xl = CreateObject("Excel.Application")
    xl.Visible = False
    Dim wb = xl.Workbooks.Add()
    Dim wsCount = 0

    Logger.Info(">>---------------------------")

    ' === CREATE "Print Package" TAB FIRST ===
    ' NOTE: This tab collects BOM items, but not all BOM items have corresponding drawing files
    ' The OpenPrintPackageDrawings rule may fail if trying to open non-drawing part numbers
    Dim drawingsTab = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    
    ' Get the job number (part number) for prefixing
    Dim jobNumber As String = ""
    On Error Resume Next
    Dim fullPartNumber As String = ThisDoc.Document.PropertySets("Design Tracking Properties")("Part Number").Value
    If Err.Number <> 0 Then
        jobNumber = "UNKNOWN"
        Err.Clear
    Else
        ' Extract only the base job number (part before first space)
        If InStr(fullPartNumber, " ") > 0 Then
            jobNumber = Left(fullPartNumber, InStr(fullPartNumber, " ") - 1)
        Else
            jobNumber = fullPartNumber
        End If
    End If
    On Error GoTo 0
    
    drawingsTab.Name = jobNumber & "_Print Package"
    
    ' === COLLECT ALL BOM ITEMS FROM ALL SHEETS (BEFORE FILTERING) ===
    Dim allBOMItems As New List(Of Object)
    Dim currentRow As Integer = 1
    
    ' Write headers for Print Package tab
    drawingsTab.Cells(1, 1).Value = "SHEET NAME"
    drawingsTab.Cells(1, 2).Value = "KEMCO PART NUMBER"
    drawingsTab.Cells(1, 3).Value = "KEMCO DESCRIPTION"
    drawingsTab.Cells(1, 4).Value = "QTY"
    drawingsTab.Cells(1, 5).Value = "UM"
    drawingsTab.Cells(1, 6).Value = "AMOUNT"
    
    currentRow = 2 ' Start data at row 2
    
    ' Loop through all sheets and collect BOM data
    For Each sheet In doc.Sheets
        If Sheet.PartsLists.Count > 0 Then
            For i = 1 To Sheet.PartsLists.Count
                Dim pl = Sheet.PartsLists(i)
                
                ' Debug: Show available columns for this PartsList
                Dim availableColumns As String = ""
                For c = 1 To pl.PartsListColumns.Count
                    availableColumns += pl.PartsListColumns(c).Title & ", "
                Next
                Logger.Info("Available columns in " & Sheet.Name & ": " & availableColumns)
                
                For Each row In pl.PartsListRows
                    If Not Row.Visible Then Continue For
                    
                    ' Only add assemblies (UM = "BOM") to Print Package tab
                    On Error Resume Next
                    Dim umCol = GetColumnIndex(pl, "UM")
                    Dim umValue As String = ""
                    If umCol > 0 Then
                        umValue = UCase(Trim(CStr(Row.Item(umCol).Value)))
                    End If
                    If Err.Number <> 0 Then
                        umValue = ""
                        Err.Clear
                    End If
                    On Error GoTo 0
                    
                    ' Only proceed if this is an assembly (UM = "BOM")
                    If umValue = "BOM" Then
                        ' Get the part number first and validate it
                        On Error Resume Next
                        Dim partNumCol = GetColumnIndex(pl, "KEMCO PART NUMBER")
                        Dim bomPartNumber As String = ""
                        If partNumCol > 0 Then
                            bomPartNumber = Trim(CStr(Row.Item(partNumCol).Value))
                        End If
                        If Err.Number <> 0 Then
                            bomPartNumber = ""
                            Err.Clear
                        End If
                        On Error GoTo 0
                        
                        ' Only add if we have a valid part number (skip empty, N/A, Error values)
                        If bomPartNumber <> "" And bomPartNumber <> "N/A" And bomPartNumber <> "Error" Then
                            ' Add to Print Package tab with safe column access
                            drawingsTab.Cells(currentRow, 1).Value = Sheet.Name
                            drawingsTab.Cells(currentRow, 2).Value = bomPartNumber
                            
                            ' Safely get other column values with error handling
                            On Error Resume Next
                        
                        Dim descCol = GetColumnIndex(pl, "KEMCO DESCRIPTION")
                        If descCol > 0 Then
                            drawingsTab.Cells(currentRow, 3).Value = Row.Item(descCol).Value
                        Else
                            drawingsTab.Cells(currentRow, 3).Value = "N/A"
                        End If
                        If Err.Number <> 0 Then
                            drawingsTab.Cells(currentRow, 3).Value = "Error"
                            Err.Clear
                        End If
                        
                        Dim qtyCol = GetColumnIndex(pl, "QTY")
                        If qtyCol > 0 Then
                            drawingsTab.Cells(currentRow, 4).Value = Row.Item(qtyCol).Value
                        Else
                            drawingsTab.Cells(currentRow, 4).Value = "N/A"
                        End If
                        If Err.Number <> 0 Then
                            drawingsTab.Cells(currentRow, 4).Value = "Error"
                            Err.Clear
                        End If
                        
                        drawingsTab.Cells(currentRow, 5).Value = umValue ' Already got this above
                        
                        Dim amtCol = GetColumnIndex(pl, "AMOUNT")
                        If amtCol > 0 Then
                            drawingsTab.Cells(currentRow, 6).Value = Row.Item(amtCol).Value
                        Else
                            drawingsTab.Cells(currentRow, 6).Value = "N/A"
                        End If
                        If Err.Number <> 0 Then
                            drawingsTab.Cells(currentRow, 6).Value = "Error"
                            Err.Clear
                        End If
                                                    On Error GoTo 0
                            
                            currentRow += 1
                        End If
                    End If
                Next
            Next
        End If
    Next
    
    ' Format the Print Package tab
    With drawingsTab.UsedRange.Borders
        .LineStyle = 1
        .Weight = 2
    End With
    drawingsTab.Cells.HorizontalAlignment = -4131
    drawingsTab.Columns.AutoFit()
    
    ' === PAGE SETUP FOR PRINT PACKAGE TAB ===
    On Error Resume Next
    
    ' Basic page setup only - simplified to avoid hangs
    drawingsTab.PageSetup.CenterHorizontally = True
    drawingsTab.PageSetup.CenterVertically = True
    drawingsTab.PageSetup.FitToPagesWide = 1
    drawingsTab.PageSetup.FitToPagesTall = False
    drawingsTab.PageSetup.Orientation = 2 ' Landscape
    drawingsTab.PageSetup.PrintGridlines = True
    
    ' Simple header with filename only
    drawingsTab.PageSetup.CenterHeader = "&F"
    
    Logger.Info("Basic page setup configured for Print Package tab")
    
    On Error GoTo 0
    
    ' === NOW CONTINUE WITH YOUR ORIGINAL LOGIC ===
    For Each sheet In doc.Sheets
        If Sheet.PartsLists.Count = 0 Then Continue For

        wsCount += 1
        Dim ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))

        ' === Clean and assign safe Excel sheet name with job number prefix ===
        Dim instanceName As String
        instanceName = CleanName(Sheet.Name)
        
        ' Debug: Log the sheet name and cleaned name
        Logger.Info("Processing sheet: '" & Sheet.Name & "' -> Cleaned: '" & instanceName & "'")
        
        ' Create the prefixed sheet name based on sheet type
        Dim finalTabName As String = ""
        
        ' HEATER drawings
        If InStr(UCase(instanceName), "HEATER FINAL") > 0 Then
            finalTabName = jobNumber & "_HEATER FINAL"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "HEATER WELDMENT") > 0 Then
            finalTabName = jobNumber & ".1_HEATER WELDMENT"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "HEATER SHELL") > 0 Then
            finalTabName = jobNumber & ".2_HEATER SHELL"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "HEATER STACK") > 0 Then
            finalTabName = jobNumber & ".3_HEATER STACK"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "GAS TRAIN") > 0 Then
            finalTabName = jobNumber & ".4_GAS TRAIN"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "MODULAR PIPING") > 0 Then
            finalTabName = jobNumber & ".5_MODULAR PIPING"
            Logger.Info("  -> Assigned: " & finalTabName)
        ' TANK drawings
        ElseIf InStr(UCase(instanceName), "TANK") > 0 And InStr(UCase(instanceName), "FFA") > 0 Then
            finalTabName = jobNumber & "_TANK FFA"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "SHELL") > 0 And InStr(UCase(instanceName), "FFA") > 0 Then
            finalTabName = jobNumber & ".1_SHELL FFA"
            Logger.Info("  -> Assigned: " & finalTabName)
        ElseIf InStr(UCase(instanceName), "SUCTION FITTING") > 0 Then
            finalTabName = "SUCTION FITTING"
            Logger.Info("  -> Assigned: " & finalTabName)
        Else
            ' For any other sheet types, use the original naming logic
            Dim baseTabName As String = Left(instanceName, 28)
            finalTabName = baseTabName
            Dim tabSuffix As Integer: tabSuffix = 1
            Do While WorksheetNameExists(wb, finalTabName)
                finalTabName = Left(baseTabName, 25) & "_" & tabSuffix
                tabSuffix += 1
            Loop
            Logger.Info("  -> Assigned (default): " & finalTabName)
        End If

        On Error Resume Next
        ws.Name = finalTabName
        On Error GoTo 0

        ' === Write Headers ===
        Dim pl = Sheet.PartsLists(1)
        Dim colCount = pl.PartsListColumns.Count
        For c = 1 To colCount
            ws.Cells(1, c).Value = pl.PartsListColumns(c).Title
        Next

        ' === Write Combined Data for all PartsLists on the sheet ===
        Dim rowIdx = 2
        For i = 1 To Sheet.PartsLists.Count
            pl = Sheet.PartsLists(i)

            For Each row In pl.PartsListRows
                If Not Row.Visible Then Continue For
                For c = 1 To colCount
                    ws.Cells(rowIdx, c).Value = Row.Item(c).Value
                Next
                rowIdx += 1
            Next
        Next

        Dim used = ws.UsedRange
        Dim lastRow = used.Rows.Count

        ' Move QTY before KEMCO DESCRIPTION
        Dim descCol = GetColumnNumber(ws, "KEMCO DESCRIPTION")
        Dim qtyCol = GetColumnNumber(ws, "QTY")
        If descCol > 0 And qtyCol > 0 And qtyCol <> descCol - 1 Then
            ws.Columns(qtyCol).Cut()
            ws.Columns(descCol).Insert()
        End If
        qtyCol = GetColumnNumber(ws, "QTY")

        ' Insert WAREHOUSE before QTY
        If qtyCol > 0 Then
            ws.Columns(qtyCol).Insert()
            ws.Cells(1, qtyCol).Value = "WAREHOUSE"
            For r = 2 To lastRow
                ws.Cells(r, qtyCol).Value = "'000"
            Next
        End If

        ' Override QTY with AMOUNT
        Dim amtCol = GetColumnNumber(ws, "AMOUNT")
        qtyCol = GetColumnNumber(ws, "QTY")
        If amtCol > 0 And qtyCol > 0 Then
            For r = 2 To lastRow
                Dim val = Trim(CStr(ws.Cells(r, amtCol).Value))
                If val <> "" Then ws.Cells(r, qtyCol).Value = val
            Next
        End If

        ' Borders
        With ws.UsedRange.Borders
            .LineStyle = 1
            .Weight = 2
        End With

        ' Left align
        ws.Cells.HorizontalAlignment = -4131

        ' Grey fill
        Dim columnsToColor = {"KEMCO PART NUMBER", "WAREHOUSE", "QTY"}
        For Each colName In columnsToColor
            Dim cIdx = GetColumnNumber(ws, colName)
            If cIdx > 0 Then
                For r = 2 To lastRow
                    ws.Cells(r, cIdx).Interior.Color = RGB(220, 220, 220)
                Next
            End If
        Next

        ' Delete rows (QTY=0, QTY=REF, PART starts with 900-)
        Dim qtyColDel = GetColumnNumber(ws, "QTY")
        Dim partCol = GetColumnNumber(ws, "KEMCO PART NUMBER")
        If qtyColDel > 0 Or partCol > 0 Then
            For r = ws.UsedRange.Rows.Count To 2 Step -1
                Dim qtyVal = UCase(Trim(CStr(ws.Cells(r, qtyColDel).Value)))
                Dim partVal = Trim(CStr(ws.Cells(r, partCol).Value))
                If qtyVal = "0" Or qtyVal = "REF" Or Left(partVal, 4) = "900-" Then
                    ws.Rows(r).Delete()
                End If
            Next
        End If

        ' UM updates (PEGGED SUPPLY or PHANTOM)
        Dim umCol = GetColumnNumber(ws, "UM")
        Dim descColUM = GetColumnNumber(ws, "KEMCO DESCRIPTION")
        If umCol > 0 Then
            Dim isHeaterFinal = InStr(UCase(ws.Name), "HEATER FINAL") > 0
            For r = 2 To ws.UsedRange.Rows.Count
                Dim currentUM = UCase(Trim(CStr(ws.Cells(r, umCol).Value)))
                If currentUM = "BOM" Then
                    Dim description = ""
                    If descColUM > 0 Then description = UCase(CStr(ws.Cells(r, descColUM).Value))
                    If isHeaterFinal And (InStr(description, "GAS TRAIN") > 0 Or InStr(description, "HEATER, WELD") > 0) Then
                        ws.Cells(r, umCol).Value = "PEGGED SUPPLY"
                    Else
                        ws.Cells(r, umCol).Value = "PHANTOM"
                    End If
                End If
            Next
        End If

        ' Highlight mismatched 304/316
        descCol = GetColumnNumber(ws, "KEMCO DESCRIPTION")
        If descCol > 0 Then
            Dim count304 = 0
            Dim count316 = 0
            Dim matFlags(ws.UsedRange.Rows.Count)
            For r = 2 To ws.UsedRange.Rows.Count
                Dim desc = UCase(CStr(ws.Cells(r, descCol).Value))
                If InStr(desc, "304") > 0 Then count304 += 1 : matFlags(r) = "304"
                If InStr(desc, "316") > 0 Then count316 += 1 : matFlags(r) = "316"
            Next
            Dim majority = ""
            If count304 > 0 And count316 > 0 Then
                If count304 >= 2 * count316 Then majority = "304"
                If count316 >= 2 * count304 Then majority = "316"
            End If
            If majority <> "" Then
                For r = 2 To ws.UsedRange.Rows.Count
                    If matFlags(r) <> "" And matFlags(r) <> majority Then
                        ws.Cells(r, descCol).Interior.Color = RGB(255, 0, 0)
                        ws.Cells(r, descCol).Font.Color = RGB(255, 255, 255)
                    End If
                Next
            End If
        End If

        ' Highlight PIPE without 10S/40S
        If descCol > 0 Then
            For r = 2 To ws.UsedRange.Rows.Count
                Dim desc = UCase(Trim(CStr(ws.Cells(r, descCol).Value)))
                If (Left(desc, 12) = "PIPE, SS304" Or Left(desc, 12) = "PIPE, SS316") And InStr(desc, "10S") = 0 Then
                    ws.Cells(r, descCol).Interior.Color = RGB(255, 0, 0)
                    ws.Cells(r, descCol).Font.Color = RGB(255, 255, 255)
                End If
                If Left(desc, 8) = "PIPE, BI" And InStr(desc, "40S") = 0 Then
                    ws.Cells(r, descCol).Interior.Color = RGB(255, 0, 0)
                    ws.Cells(r, descCol).Font.Color = RGB(255, 255, 255)
                End If
            Next
        End If

        ' Final auto-fit
        ws.Columns.AutoFit()
        
        ' === PAGE SETUP CONFIGURATION ===
        ' Configure page setup for printing - simplified to avoid hangs
        On Error Resume Next
        
        ' Basic page setup only - skip complex properties that might cause hangs
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically = True
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = False
        ws.PageSetup.Orientation = 2 ' Landscape
        ws.PageSetup.PrintGridlines = True
        
        ' Simple header with filename only
        ws.PageSetup.CenterHeader = "&F"
        
        Logger.Info("Basic page setup configured for " & ws.Name)
        
        On Error GoTo 0
    Next

    ' Remove unused sheets
    While wb.Worksheets.Count > wsCount + 1 ' +1 for Print Package tab
        wb.Worksheets(1).Delete()
    End While

    ' Save as: {Part Number}-BOM.xlsx (ask user for destination folder)
    Dim partNumber = ThisDoc.Document.PropertySets("Design Tracking Properties")("Part Number").Value
    Dim safePart = CleanName(partNumber)
    Dim fileName = safePart & "-BOM.xlsx"

    ' Pick destination folder
    Dim selectedFolder As String = ""
    On Error Resume Next
    Dim fbd = New System.Windows.Forms.FolderBrowserDialog()
    fbd.Description = "Select a folder to save the BOM Excel file"
    fbd.ShowNewFolderButton = True
    fbd.SelectedPath = Environ("USERPROFILE") & "\Downloads"
    Dim result = fbd.ShowDialog()
    If Err.Number <> 0 Then
        Err.Clear
        selectedFolder = Environ("USERPROFILE") & "\Downloads"
    Else
        If result = System.Windows.Forms.DialogResult.OK Then
            selectedFolder = fbd.SelectedPath
        Else
            MsgBox("Export cancelled.", vbInformation, "Export BOM")
            wb.Close(False)
            xl.Quit()
            On Error GoTo 0
            Exit Sub
        End If
    End If
    On Error GoTo 0

    Dim savePath = IO.Path.Combine(selectedFolder, fileName)
    
    ' Delete existing file if it exists
    If Dir(savePath) <> "" Then 
        On Error Resume Next
        Kill(savePath)
        ' Wait a moment for file system to release the file
        Dim waitTime As Integer = 500 ' 0.5 seconds
        Dim startTime = Timer
        Do While Timer < startTime + (waitTime / 1000)
            ' Wait
        Loop
        On Error GoTo 0
    End If
    
    ' Save the workbook
    wb.SaveAs(savePath)
    wb.Close(False)
    
    ' Automatically open the Excel file after creation
    On Error Resume Next
    Logger.Info("Attempting to open Excel file: " & savePath)
    
    ' Wait a moment for file to be fully accessible
    Dim openWaitTime As Integer = 1000 ' 1 second
    Dim openStartTime = Timer
    Do While Timer < openStartTime + (openWaitTime / 1000)
        ' Wait
    Loop
    
    ' Try to open using WScript.Shell (most reliable method)
    Dim shell = CreateObject("WScript.Shell")
    shell.Run("""" & savePath & """")
    
    If Err.Number = 0 Then
        Logger.Info("Excel file opened successfully using WScript.Shell")
    Else
        Logger.Warn("Failed to open Excel file: " & Err.Description)
        Err.Clear
        ' Show user where the file is located
        MsgBox("BOM export completed successfully!" & vbCrLf & vbCrLf & "File saved to:" & vbCrLf & savePath & vbCrLf & vbCrLf & "The file will open automatically, or you can open it manually from the location above.", vbInformation, "Export Complete")
    End If
    
    On Error GoTo 0
    
    xl.Quit()

    ' No more dialog box - the Excel file opens automatically and that's enough feedback
End Sub

' === NEW HELPER FUNCTION ===
Function GetColumnIndex(pl, columnTitle)
    If pl Is Nothing OrElse pl.PartsListColumns Is Nothing Then
        GetColumnIndex = -1
        Exit Function
    End If
    
    On Error Resume Next
    For c = 1 To pl.PartsListColumns.Count
        If pl.PartsListColumns(c) IsNot Nothing AndAlso pl.PartsListColumns(c).Title IsNot Nothing Then
            If Trim(UCase(pl.PartsListColumns(c).Title)) = Trim(UCase(columnTitle)) Then
                GetColumnIndex = c
                On Error GoTo 0
                Exit Function
            End If
        End If
        If Err.Number <> 0 Then
            Err.Clear
        End If
    Next
    On Error GoTo 0
    GetColumnIndex = -1
End Function

Function CleanName(name)
    Dim i, ch, bad
    bad = ":/\?*[]()"
    For i = 1 To Len(bad)
        ch = Mid(bad, i, 1)
        name = Replace(name, ch, "_")
    Next
    CleanName = Trim(name)
End Function

Function GetColumnNumber(ws, title)
    For c = 1 To ws.UsedRange.Columns.Count
        If Trim(UCase(ws.Cells(1, c).Value)) = Trim(UCase(title)) Then
            GetColumnNumber = c : Exit Function
        End If
    Next
    GetColumnNumber = -1
End Function

Function WorksheetNameExists(wb, name)
    WorksheetNameExists = False
    For Each s In wb.Worksheets
        If s.Name = name Then WorksheetNameExists = True : Exit Function
    Next
End Function
