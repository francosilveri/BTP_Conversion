Imports System
Imports System.Configuration
Imports System.Collections.Specialized
Imports OfficeOpenXml
Imports System.IO

Module BTP_Conversion
    Dim macro As String = ""
    Dim inputfolder As String = ""
    Dim outputfolder As String = ""
    Dim filename As String = ""
    Dim ErrMsg As String = ""
    Dim btpOutFileName As String = ""
    Dim btpOutFileNameLog As String = ""

    Sub InitFromCommandLine(ByRef cmdArgs() As String)
        Dim flag As String
        Dim param As String
        Try
            If cmdArgs.Length = 0 Then
                ErrMsg = GetUsage("Invalid Syntax:")
                Exit Sub
            End If
            For Each arg As String In cmdArgs
                arg = RemoveQuote(arg)
                arg = arg.Trim
                If arg.StartsWith("-") Or arg.StartsWith("/") Then

                    If arg.IndexOf("=") > 0 Then
                        flag = arg.Split("="c)(0).Substring(1).ToLower
                        param = arg.Split("="c)(1).ToLower

                        Select Case flag
                            Case "btptype"
                                Select Case param
                                    Case "m", "e", "oralbm", "gen"
                                        macro = param
                                        Continue For
                                End Select
                                ErrMsg = GetUsage("Invalid Parameter: " + param)
                                Exit Sub
                            Case "inputfolder"
                                inputfolder = param
                                If Dir(inputfolder, FileAttribute.Directory).Length > 0 Then
                                    If inputfolder.EndsWith("\") Then inputfolder = inputfolder.Substring(0, inputfolder.Length - 1)
                                    Continue For
                                End If
                                ErrMsg = GetUsage("'" + param + "' folder does not exist or folder name is invalid.")
                                Exit Sub
                            Case "outputfolder"
                                outputfolder = param
                                If Dir(outputfolder, FileAttribute.Directory).Length > 0 Then
                                    If outputfolder.EndsWith("\") Then outputfolder = outputfolder.Substring(0, outputfolder.Length - 1)
                                    Continue For
                                End If
                                ErrMsg = GetUsage("Invalid Parameter:  " + param + ", folder does not exist or folder name is invalid.")
                                Exit Sub
                            Case "filename"
                                filename = param
                                If Dir(inputfolder & "\" & filename, FileAttribute.Directory).Length > 0 Then
                                    Continue For
                                End If
                                ErrMsg = GetUsage("Invalid Parameter:  " + param + ", filer does not exist or file name is invalid.")
                                Exit Sub
                        End Select
                        ErrMsg = GetUsage("Invalid Argument: " + arg)
                        Exit Sub
                    End If
                End If
            Next
        Catch ex As Exception
            ErrMsg = "Exception Error: " + ex.Message
        End Try

    End Sub

    Function Main(ByVal cmdArgs() As String) As Integer

        InitFromCommandLine(cmdArgs)

        If outputfolder <> "" Then
            btpOutFileName = outputfolder & "\" & ConfigurationManager.AppSettings.Get("btpOutFileName")
            btpOutFileNameLog = outputfolder & "\" & ConfigurationManager.AppSettings.Get("btpOutFileNameLog")
        Else
            btpOutFileName = Application.StartupPath & "\" & ConfigurationManager.AppSettings.Get("btpOutFileName")
            btpOutFileNameLog = Application.StartupPath & "\" & ConfigurationManager.AppSettings.Get("btpOutFileNameLog")
        End If

        If File.Exists(btpOutFileName) Then
            File.Delete(btpOutFileName)
        End If
        If File.Exists(btpOutFileNameLog) Then
            File.Delete(btpOutFileNameLog)
        End If

        Dim wr As New StreamWriter(btpOutFileNameLog)
        If ErrMsg = "" And (macro = "" Or inputfolder = "" Or outputfolder = "" Or filename = "") Then
            ErrMsg = GetUsage("Missig parameters.")
        End If
        If ErrMsg <> "" Then
            wr.WriteLine("ko")
            wr.WriteLine(ErrMsg)
            wr.Close()
            Return AppConstant.ExitCode.WRONG_ARGUMENTS
            Exit Function
        End If

        If macro = "m" Then
            ErrMsg = BTPMechReadExcelAndSaveCsv(inputfolder, filename)
        ElseIf macro = "e" Then
            ErrMsg = BTPElecReadExcelAndSaveCsv(inputfolder, filename)
        ElseIf macro = "oralbm" Then
            ErrMsg = BTPOralBMechReadExcelAndSaveCsv(inputfolder, filename)
        ElseIf macro = "gen" Then
            ErrMsg = BTPGenericReadExcelAndSaveCsv(inputfolder, filename)
        Else
            ErrMsg = ""
        End If

        If ErrMsg <> "ok" Then
            wr.WriteLine("ko")
            wr.WriteLine(ErrMsg)
            wr.Close()
            Return AppConstant.ExitCode.DRAFT_SOFTERROR
        ElseIf ErrMsg = "" Then
            wr.WriteLine("ko")
            Return AppConstant.ExitCode.CODE_EXCEPTION
        Else
            wr.WriteLine(ErrMsg)
            wr.Close()
            Return AppConstant.ExitCode.OK
        End If

    End Function

    Function BTPMechReadExcelAndSaveCsv(ByVal inputfolder As String, ByVal filename As String) As String

        Try
            Dim existingFile As FileInfo = New FileInfo(inputfolder & "\" & filename)
            Using package As ExcelPackage = New ExcelPackage(existingFile)
                'get the App Settings for BTP mechanical
                Dim csvSeparator = ConfigurationManager.AppSettings.Get("csvSeparator")
                Dim btpHeaderm() As String = ConfigurationManager.AppSettings.Get("btpHeaderm").Split(csvSeparator)
                Dim btpHeadermRow As Integer = ConfigurationManager.AppSettings.Get("btpHeadermRow")
                Dim btpHeadermSheet As String = ConfigurationManager.AppSettings.Get("btpHeadermSheet")

                'get the worksheet "Sheet1" in the workbook
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(btpHeadermSheet)
                If worksheet Is Nothing Then
                    Return "The input file is not a standard BTP mechanical"
                    Exit Function
                End If

                'check if the input file is a correct BTP mechanical file
                For col As Integer = 1 To btpHeaderm.Length
                    If worksheet.Cells(btpHeadermRow, col).Value.ToString().Trim() <> btpHeaderm(col - 1).Trim() Then
                        Return "The input file is not a standard BTP mechanical"
                        Exit Function
                    End If
                Next

                'output the input file to a standard CSV file, separated by pipeline
                Dim wr As New StreamWriter(btpOutFileName)
                Dim wsLine As String
                Dim row As Integer = btpHeadermRow
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                    wsLine = ""
                    For col As Integer = 1 To btpHeaderm.Length
                        If col = btpHeaderm.Length Then
                            wsLine += worksheet.Cells(row, col).Value.ToString().Trim()
                        Else
                            wsLine += worksheet.Cells(row, col).Value.ToString().Trim() & csvSeparator
                        End If
                    Next
                    wr.WriteLine(wsLine)
                    row += 1
                Loop
                wr.Close()
            End Using
            Return "ok"
        Catch ex As Exception
            ErrMsg = "Error: " + ex.Message
            Return ErrMsg
        End Try
    End Function

    Function BTPElecReadExcelAndSaveCsv(ByVal inputfolder As String, ByVal filename As String) As String

        Try
            Dim existingFile As FileInfo = New FileInfo(inputfolder & "\" & filename)
            Using package As ExcelPackage = New ExcelPackage(existingFile)
                'get the App Settings for BTP mechanical
                Dim csvSeparator = ConfigurationManager.AppSettings.Get("csvSeparator")
                Dim btpHeadere() As String = ConfigurationManager.AppSettings.Get("btpHeadere").Split(csvSeparator)
                Dim btpHeadereRow As Integer = ConfigurationManager.AppSettings.Get("btpHeadereRow")
                Dim btpHeadereSheet As String = ConfigurationManager.AppSettings.Get("btpHeadereSheet")

                'get the worksheet "Parts List" in the workbook
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(btpHeadereSheet)
                If worksheet Is Nothing Then
                    Return "The input file is not a standard BTP electrical"
                    Exit Function
                End If

                'check if the input file is a correct BTP electrical file
                For col As Integer = 1 To btpHeadere.Length
                    If worksheet.Cells(btpHeadereRow, col).Value.ToString().Trim() <> btpHeadere(col - 1).Trim() Then
                        Return "The input file is not a standard BTP electrical"
                        Exit Function
                    End If
                Next

                'Modify BTP electrical sheet to BTP mechanical sheet
                Dim btpHeaderm() As String = ConfigurationManager.AppSettings.Get("btpHeaderm").Split(csvSeparator)
                Dim btpHeadermRow As Integer = ConfigurationManager.AppSettings.Get("btpHeadermRow")
                Dim btpHeadermSheet As String = ConfigurationManager.AppSettings.Get("btpHeadermSheet")

                'check if the input file has the project name indication
                Dim projectRow As Integer
                Dim project As String = ""
                Dim row As Integer
                For row = 1 To btpHeadereRow
                    projectRow = row
                    If Not worksheet.Cells(row, 1).Value Is Nothing Then
                        If worksheet.Cells(row, 1).Value.ToString().ToLower().Contains("project name: ") Then
                            project = "PROJECT " & worksheet.Cells(row, 1).Value.ToString().ToLower().Replace("project name: ", "").ToUpper()
                            Exit For
                        End If
                    End If
                Next
                If projectRow >= btpHeadereRow + 1 Then
                    Return "The input file has not a valid Project Name defined"
                    Exit Function
                End If

                'check if the input file has the comma char instead of the point char as decimal separator
                'check the max string length for columns Installation (2), Location (3), Manufacturer(11), Part Number (5)
                Dim maxCol2 As Integer = 0, maxCol3 As Integer = 0, maxCol5 As Integer = 0, maxCol11 As Integer = 0
                Dim maxRow As Integer
                row = btpHeadereRow + 1
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                    If Not worksheet.Cells(row, 5).Value Is Nothing Then If Val(worksheet.Cells(row, 12).Value) = 0 Then worksheet.Cells(row, 12).Value = worksheet.Cells(row, 5).Value
                    If Not worksheet.Cells(row, 7).Value Is Nothing Then worksheet.Cells(row, 7).Value = worksheet.Cells(row, 7).Value.ToString().Replace(",", ".")
                    If Not worksheet.Cells(row, 8).Value Is Nothing Then worksheet.Cells(row, 8).Value = worksheet.Cells(row, 8).Value.ToString().Replace(",", ".")
                    If Not worksheet.Cells(row, 9).Value Is Nothing Then worksheet.Cells(row, 9).Value = worksheet.Cells(row, 9).Value.ToString().Replace(",", ".")
                    If Not worksheet.Cells(row, 2).Value Is Nothing Then If worksheet.Cells(row, 2).Value.ToString().Length > maxCol2 Then maxCol2 = worksheet.Cells(row, 2).Value.ToString().Length
                    If Not worksheet.Cells(row, 3).Value Is Nothing Then If worksheet.Cells(row, 3).Value.ToString().Length > maxCol3 Then maxCol3 = worksheet.Cells(row, 3).Value.ToString().Length
                    If Not worksheet.Cells(row, 5).Value Is Nothing Then If worksheet.Cells(row, 5).Value.ToString().Length > maxCol5 Then maxCol5 = worksheet.Cells(row, 5).Value.ToString().Length
                    If Not worksheet.Cells(row, 11).Value Is Nothing Then If worksheet.Cells(row, 11).Value.ToString().Length > maxCol11 Then maxCol11 = worksheet.Cells(row, 11).Value.ToString().Length
                    row += 1
                Loop
                maxRow = row - 1

                Dim sortCol As Integer = btpHeadere.Length + 1
                row = btpHeadereRow + 1
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                    If Not worksheet.Cells(row, 2).Value Is Nothing Then
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol + 1).Value & worksheet.Cells(row, 2).Value & Space(maxCol2 - worksheet.Cells(row, 2).Value.ToString().Length)
                    Else
                        worksheet.Cells(row, sortCol + 1).Value = worksheet.Cells(row, sortCol).Value & Space(maxCol2)
                    End If
                    If Not worksheet.Cells(row, 3).Value Is Nothing Then
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol).Value & worksheet.Cells(row, 3).Value & Space(maxCol3 - worksheet.Cells(row, 3).Value.ToString().Length)
                    Else
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol).Value & Space(maxCol3)
                    End If
                    If Not worksheet.Cells(row, 5).Value Is Nothing Then
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol).Value & worksheet.Cells(row, 5).Value & Space(maxCol5 - worksheet.Cells(row, 5).Value.ToString().Length)
                    Else
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol).Value & Space(maxCol5)
                    End If
                    If Not worksheet.Cells(row, 11).Value Is Nothing Then
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol).Value & worksheet.Cells(row, 11).Value & Space(maxCol11 - worksheet.Cells(row, 11).Value.ToString().Length)
                    Else
                        worksheet.Cells(row, sortCol).Value = worksheet.Cells(row, sortCol).Value & Space(maxCol11)
                    End If
                    row += 1
                Loop

                'sort the input file by Installation, Location, Manufacturer, Part Number
                'pack the input by qty for Installation, Location, Manufacturer, Part Number
                Dim tmp As Object
                For r1 As Integer = btpHeadereRow + 1 To maxRow - 1
                    For r2 As Integer = r1 + 1 To maxRow
                        If worksheet.Cells(r1, sortCol).Value.ToString() > worksheet.Cells(r2, sortCol).Value.ToString() Then
                            For c As Integer = 1 To sortCol
                                tmp = worksheet.Cells(r1, c).Value
                                worksheet.Cells(r1, c).Value = worksheet.Cells(r2, c).Value
                                worksheet.Cells(r2, c).Value = tmp
                            Next
                        End If
                    Next
                Next

                row = btpHeadereRow + 1
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                    If CDbl(worksheet.Cells(row, 7).Value) = 0 Or IsNothing(worksheet.Cells(row, 12).Value) Then
                        worksheet.DeleteRow(row)
                        row -= 1
                    ElseIf worksheet.Cells(row, sortCol).Value = worksheet.Cells(row + 1, sortCol).Value Then
                        worksheet.Cells(row, 7).Value = CDbl(worksheet.Cells(row, 7).Value) + CDbl(worksheet.Cells(row + 1, 7).Value)
                        worksheet.Cells(row, 8).Value = CDbl(worksheet.Cells(row, 8).Value) + CDbl(worksheet.Cells(row + 1, 8).Value)
                        worksheet.Cells(row, 9).Value = CDbl(worksheet.Cells(row, 9).Value) + CDbl(worksheet.Cells(row + 1, 9).Value)
                        worksheet.DeleteRow(row + 1)
                        row -= 1
                    End If
                    row += 1
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                Loop

                'Output the BTP electrical file to BTP mechanical file
                Dim ws As ExcelWorksheet = package.Workbook.Worksheets.Add("FeatureBOM")
                For col As Integer = 1 To btpHeaderm.Length
                    ws.Cells(1, col).Value = btpHeaderm(col - 1).Trim()
                Next

                Dim lastcode As Object
                row = btpHeadereRow + 1
                Dim destRow As Integer = 2
                Dim itemnumber As Integer
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                    lastcode = worksheet.Cells(row, 2).Value & worksheet.Cells(row, 3).Value
                    ws.Cells(destRow, 1).Value = 1
                    ws.Cells(destRow, 3).Value = worksheet.Cells(row, 3).Value
                    ws.Cells(destRow, 4).Value = worksheet.Cells(row, 3).Value
                    ws.Cells(destRow, 6).Value = project & worksheet.Cells(row, 3).Value
                    ws.Cells(destRow, 7).Value = 1
                    ws.Cells(destRow, 8).Value = "ea"
                    ws.Cells(destRow, 9).Value = project & worksheet.Cells(row, 3).Value
                    ws.Cells(destRow, 18).Value = "Assembly"
                    ws.Cells(destRow, 19).Value = "Commercial Electrical"
                    itemnumber = 10

                    destRow += 1
                    Do
                        If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                        If lastcode <> worksheet.Cells(row, 2).Value & worksheet.Cells(row, 3).Value Then Exit Do
                        ws.Cells(destRow, 1).Value = 2
                        ws.Cells(destRow, 2).Value = itemnumber
                        ws.Cells(destRow, 3).Value = worksheet.Cells(row, 3).Value
                        ws.Cells(destRow, 4).Value = worksheet.Cells(row, 12).Value
                        ws.Cells(destRow, 5).Value = worksheet.Cells(row, 5).Value
                        ws.Cells(destRow, 6).Value = worksheet.Cells(row, 12).Value
                        ws.Cells(destRow, 7).Value = worksheet.Cells(row, 7).Value
                        ws.Cells(destRow, 8).Value = "ea"
                        ws.Cells(destRow, 9).Value = worksheet.Cells(row, 10).Value
                        ws.Cells(destRow, 10).Value = worksheet.Cells(row, 4).Value
                        ws.Cells(destRow, 11).Value = worksheet.Cells(row, 11).Value
                        ws.Cells(destRow, 16).Value = worksheet.Cells(row, 6).Value
                        ws.Cells(destRow, 18).Value = "Cmponent"
                        ws.Cells(destRow, 19).Value = "Commercial Electrical"

                        itemnumber += 1
                        destRow += 1
                        row += 1
                    Loop
                Loop

                For col As Integer = 1 To btpHeaderm.Length
                    ws.Cells(btpHeadermRow, col).Value = btpHeaderm(col - 1).Trim()
                Next

                'output the input file to a standard CSV file, separated by pipeline
                Dim wr As New StreamWriter(btpOutFileName)
                Dim wsLine As String
                row = 1
                Do
                    If IsNothing(ws.Cells(row, 1).Value) Then Exit Do
                    wsLine = ""
                    For col As Integer = 1 To btpHeaderm.Length
                        If col = btpHeaderm.Length Then
                            wsLine += ws.Cells(row, col).Value
                        Else
                            wsLine += ws.Cells(row, col).Value & csvSeparator
                        End If
                    Next
                    wr.WriteLine(wsLine)
                    row += 1
                Loop
                wr.Close()
            End Using
            Return "ok"
        Catch ex As Exception
            ErrMsg = "Error: " + ex.Message
            Return ErrMsg
        End Try
    End Function

    Function BTPOralBMechReadExcelAndSaveCsv(ByVal inputfolder As String, ByVal filename As String) As String
        Try
            Dim existingFile As FileInfo = New FileInfo(inputfolder & "\" & filename)
            Using package As ExcelPackage = New ExcelPackage(existingFile)
                'get the App Settings for BTP Oral-B mechanical
                Dim csvSeparator = ConfigurationManager.AppSettings.Get("csvSeparator")
                Dim btpHeaderoralbm() As String = ConfigurationManager.AppSettings.Get("btpHeaderoralbm").Split(csvSeparator)
                Dim btpHeaderoralbmRow As Integer = ConfigurationManager.AppSettings.Get("btpHeaderoralbmRow")
                Dim btpHeaderoralbmSheet As String = ConfigurationManager.AppSettings.Get("btpHeaderoralbmSheet")

                'get the worksheet "Default" in the workbook
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(btpHeaderoralbmSheet)
                If worksheet Is Nothing Then
                    Return "The input file is not a standard BTP Oral-B mechanical"
                    Exit Function
                End If

                'check if the input file is a correct BTP Oral-B mechanical file
                For col As Integer = 1 To btpHeaderoralbm.Length
                    If worksheet.Cells(btpHeaderoralbmRow, col).Value.ToString().Trim() <> btpHeaderoralbm(col - 1).Trim() Then
                        Return "The input file is not a standard BTP Oral-B mechanical"
                        Exit Function
                    End If
                Next

                'check if the input file has all the columns in the correct format
                Dim maxRow As Integer, ErrMsg As String = ""
                Dim row As Integer = btpHeaderoralbmRow + 1
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) And IsNothing(worksheet.Cells(row + 1, 1).Value) And IsNothing(worksheet.Cells(row + 2, 1).Value) Then Exit Do
                    row += 1
                Loop
                maxRow = row - 1

                For row = btpHeaderoralbmRow + 1 To maxRow
                    If IsNothing(worksheet.Cells(row, 1).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Structure Level." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 2).Value) And row > 2 Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Line Number." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 3).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Number." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 4).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Tool/Equipment Name English." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 5).Value) And row > 2 Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Part ID." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 6).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Part ID Version." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 7).Value) And row > 2 Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Quantity." & vbCrLf
                    End If
                    If IsNothing(worksheet.Cells(row, 8).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Source." & vbCrLf
                    End If
                    If worksheet.Cells(row, 8).Value.ToString() = "Buy" And IsNothing(worksheet.Cells(row, 9).Value) And IsNothing(worksheet.Cells(row, 10).Value) And IsNothing(worksheet.Cells(row, 11).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Vendor." & vbCrLf
                    End If
                    If worksheet.Cells(row, 8).Value.ToString() = "Buy" And Not IsNothing(worksheet.Cells(row, 9).Value) And IsNothing(worksheet.Cells(row, 10).Value) And IsNothing(worksheet.Cells(row, 11).Value) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Order Number or Order Text." & vbCrLf
                    End If
                    If worksheet.Cells(row, 8).Value.ToString() = "Buy" And IsNothing(worksheet.Cells(row, 9).Value) And Not (IsNothing(worksheet.Cells(row, 10).Value) Or IsNothing(worksheet.Cells(row, 11).Value)) Then
                        ErrMsg = ErrMsg & "Error in row: " & row & ", the item has no Vendor." & vbCrLf
                    End If
                Next
                If ErrMsg <> "" Then Return ErrMsg

                'Modify BTP Oral-B sheet to BTP mechanical sheet
                Dim btpHeaderm() As String = ConfigurationManager.AppSettings.Get("btpHeaderm").Split(csvSeparator)
                Dim btpHeadermRow As Integer = ConfigurationManager.AppSettings.Get("btpHeadermRow")
                Dim btpHeadermSheet As String = ConfigurationManager.AppSettings.Get("btpHeadermSheet")

                'Output the BTP Oral-B mechanical file to BTP mechanical file
                Dim ws As ExcelWorksheet = package.Workbook.Worksheets.Add(btpHeadermSheet)
                For col As Integer = 1 To btpHeaderm.Length
                    ws.Cells(1, col).Value = btpHeaderm(col - 1).Trim()
                Next

                Dim StructureLevel As Integer = 0
                Dim StructureLevelAfter As Integer = 0
                Dim StructureLevelBefore As Integer = 0
                Dim LineNumber As String = ""
                Dim Number As String = ""
                Dim NumberBefore As String = ""
                Dim ToolEquipmentNameEnglish As String = ""
                Dim PartID As String = ""
                Dim Version As String = ""
                Dim VersionBefore As String = ""
                Dim Level As Integer = 0
                Dim ItemNumber As Integer = 0
                Dim MajorRevision As Integer = 0
                Dim MajorRevisionBefore As Integer = 0
                Dim ParentCode As String = ""
                Dim Quantity As String = ""
                Dim Source As String = ""
                Dim Vendor As String = ""
                Dim OrderNumber As String = ""
                Dim OrderText As String = ""
                Dim DrawingNumber As String = ""
                Dim EngineeringPartNumber As String = ""
                Dim MmrNumber As String = ""
                Dim Qt
                Dim UnitOfMeasure As String = ""
                Dim Qty As Integer = 0
                Dim PartNameTitle As String = ""
                Dim Manufacturer As String = ""
                Dim Specification As String = ""
                Dim PartType As String = ""
                Dim Presentation As String = ""

                For row = btpHeaderoralbmRow + 1 To maxRow
                    StructureLevel = worksheet.Cells(row, 1).Value + 1
                    StructureLevelAfter = worksheet.Cells(row + 1, 1).Value + 1
                    LineNumber = worksheet.Cells(row, 2).Value
                    Number = worksheet.Cells(row, 3).Value.ToString().Trim()
                    NumberBefore = worksheet.Cells(row - 1, 3).Value.Trim()
                    ToolEquipmentNameEnglish = worksheet.Cells(row, 4).Value.ToString().ToUpper()
                    PartID = worksheet.Cells(row, 5).Value
                    Version = worksheet.Cells(row, 6).Value
                    VersionBefore = worksheet.Cells(row - 1, 6).Value
                    Quantity = worksheet.Cells(row, 7).Value
                    If Quantity = "" And row = 2 Then
                        Quantity = "1 each"
                    End If
                    Source = worksheet.Cells(row, 8).Value
                    Vendor = worksheet.Cells(row, 9).Value
                    OrderNumber = worksheet.Cells(row, 10).Value
                    OrderText = worksheet.Cells(row, 11).Value

                    If worksheet.Cells(row - 1, 1).Value <> "Structure Level" Then
                        StructureLevelBefore = worksheet.Cells(row - 1, 1).Value + 1
                    Else
                        StructureLevelBefore = 0
                    End If

                    Level = StructureLevel
                    ItemNumber = LineNumber

                    If Version <> "" Then
                        MajorRevision = Val(Mid(Version, 1, InStr(1, Version, ".") - 1))
                    Else
                        MajorRevision = 0
                    End If
                    If VersionBefore <> "" And VersionBefore <> "Version" Then
                        MajorRevisionBefore = Val(Mid(VersionBefore, 1, InStr(1, VersionBefore, ".") - 1))
                    Else
                        MajorRevisionBefore = 0
                    End If

                    If row = 2 Then
                        ParentCode = Number & "-R" & Format(MajorRevision, "00")
                        worksheet.Cells(row, 2).Value = ParentCode
                    ElseIf Level > StructureLevelBefore And row > 2 Then
                        ParentCode = NumberBefore & "-R" & Format(MajorRevisionBefore, "00")
                        worksheet.Cells(row, 2).Value = ParentCode
                    ElseIf Level = StructureLevelBefore And row > 2 Then
                        worksheet.Cells(row, 2).Value = ParentCode
                    ElseIf Level < StructureLevelBefore And row > 2 Then
                        For rrow As Integer = row - 1 To 2 Step -1
                            If worksheet.Cells(rrow, 1).Value + 1 = Level Then
                                ParentCode = worksheet.Cells(rrow, 2).Value
                                worksheet.Cells(row, 2).Value = ParentCode
                                Exit For
                            End If
                        Next
                    End If

                    'If DestinationFile = "" And ParentCode <> "" Then
                    'DestinationFile = ParentCode
                    'End If

                    If Source = "Make" Then
                        DrawingNumber = Number & "-R" & Format(MajorRevision, "00")
                    ElseIf Source = "Buy" Then
                        DrawingNumber = Number
                    Else
                    End If

                    EngineeringPartNumber = Number
                    MmrNumber = PartID

                    'Erase Qt
                    Qt = Split(Quantity, " ")
                    If UBound(Qt) = -1 Then
                        Qty = 0
                        UnitOfMeasure = ""
                    Else
                        Qty = Qt(0)
                        If Qt(1) = "each" Then
                            UnitOfMeasure = "ea"
                        Else
                            Return "Error in row: " & row & ", the item has no valide Unit Of Measure." & vbCrLf
                            Exit Function
                        End If
                    End If

                    PartNameTitle = ToolEquipmentNameEnglish
                    Manufacturer = Vendor
                    If Source = "Make" Then
                        Specification = ""
                        PartType = "Fabricated"
                    ElseIf Source = "Buy" Then
                        If OrderNumber <> "" Then
                            Specification = OrderNumber
                        Else
                            Specification = ""
                        End If
                        If Specification <> "" And OrderText <> "" Then
                            Specification &= ";"
                        End If
                        If OrderText <> "" Then
                            Specification &= OrderText
                        End If
                        If Manufacturer <> "" Then
                            PartType = "Commercial"
                        Else
                            PartType = "Hardware"
                        End If
                    Else
                    End If

                    If StructureLevel < StructureLevelAfter Then
                        Presentation = "OBAssembly"
                    ElseIf StructureLevel >= StructureLevelAfter Then
                        Presentation = "OBCmponent"
                    Else
                    End If

                    ws.Cells(row, 1).Value = Level
                    ws.Cells(row, 2).Value = ItemNumber
                    ws.Cells(row, 3).Value = ParentCode
                    ws.Cells(row, 4).Value = DrawingNumber
                    ws.Cells(row, 5).Value = EngineeringPartNumber
                    ws.Cells(row, 6).Value = MmrNumber
                    ws.Cells(row, 7).Value = Qty
                    ws.Cells(row, 8).Value = UnitOfMeasure
                    ws.Cells(row, 9).Value = PartNameTitle
                    ws.Cells(row, 10).Value = PartNameTitle
                    ws.Cells(row, 11).Value = Manufacturer
                    ws.Cells(row, 16).Value = Specification
                    ws.Cells(row, 18).Value = Presentation
                    ws.Cells(row, 19).Value = PartType
                Next

                'output the input file to a standard CSV file, separated by pipeline
                Dim wr As New StreamWriter(btpOutFileName)
                Dim wsLine As String
                row = 1
                Do
                    If IsNothing(ws.Cells(row, 1).Value) Then Exit Do
                    wsLine = ""
                    For col As Integer = 1 To btpHeaderm.Length
                        If col = btpHeaderm.Length Then
                            wsLine += ws.Cells(row, col).Value
                        Else
                            wsLine += ws.Cells(row, col).Value & csvSeparator
                        End If
                    Next
                    wr.WriteLine(wsLine)
                    row += 1
                Loop
                wr.Close()
            End Using
            Return "ok"
        Catch ex As Exception
            ErrMsg = "Error: " + ex.Message
            Return ErrMsg
        End Try
    End Function

    Function BTPGenericReadExcelAndSaveCsv(ByVal inputfolder As String, ByVal filename As String) As String

        Try
            Dim existingFile As FileInfo = New FileInfo(inputfolder & "\" & filename)
            Using package As ExcelPackage = New ExcelPackage(existingFile)
                'get the App Settings for BTP mechanical
                Dim csvSeparator = ConfigurationManager.AppSettings.Get("csvSeparator")
                Dim btpHeaderm() As String = ConfigurationManager.AppSettings.Get("btpHeaderm").Split(csvSeparator)
                Dim btpHeadermRow As Integer = ConfigurationManager.AppSettings.Get("btpHeadermRow")
                Dim btpHeadermSheet As String = ConfigurationManager.AppSettings.Get("btpHeadermSheet")

                'get the worksheet "Sheet1" in the workbook
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(btpHeadermSheet)
                If worksheet Is Nothing Then
                    Return "The input file is not a standard BTP mechanical"
                    Exit Function
                End If

                'check if the input file is a correct BTP mechanical file
                For col As Integer = 1 To btpHeaderm.Length
                    If worksheet.Cells(btpHeadermRow, col).Value.ToString().Trim() <> btpHeaderm(col - 1).Trim() Then
                        Return "The input file is not a standard BTP mechanical"
                        Exit Function
                    End If
                Next

                'output the input file to a standard CSV file, separated by pipeline
                Dim wr As New StreamWriter(btpOutFileName)
                Dim wsLine As String
                Dim row As Integer = btpHeadermRow
                Do
                    If IsNothing(worksheet.Cells(row, 1).Value) Then Exit Do
                    wsLine = ""
                    For col As Integer = 1 To btpHeaderm.Length
                        If col = btpHeaderm.Length Then
                            wsLine += worksheet.Cells(row, col).Value.ToString().Trim()
                        Else
                            wsLine += worksheet.Cells(row, col).Value.ToString().Trim() & csvSeparator
                        End If
                    Next
                    wr.WriteLine(wsLine)
                    row += 1
                Loop
                wr.Close()
            End Using
            Return "ok"
        Catch ex As Exception
            ErrMsg = "Error: " + ex.Message
            Return ErrMsg
        End Try
    End Function


    Function GetUsage(ByVal ErrorLine As String) As String
        Dim str As String = vbCrLf

        str += "BTP_Conversion 1.0.0 - Fameccanica.Data" + vbCrLf + vbCrLf
        If ErrorLine <> "" Then
            str += "Error:" + vbCrLf
            str += ErrorLine + vbCrLf + vbCrLf
        End If
        str += vbCrLf
        str += "Usage:" + vbCrLf
        str += " -btptype|-inputfolder|-outputfolder|-filename" + vbCrLf
        str += "             Procedure to start" + vbCrLf
        str += vbCrLf
        str += " -btptype=" + vbCrLf
        str += "             Specify the type of BTP: m, e, oralbm" + vbCrLf
        str += vbCrLf
        str += " -inputfolder=" + vbCrLf
        str += "             Specify input folder" + vbCrLf
        str += vbCrLf
        str += " -outputfolder=" + vbCrLf
        str += "             Specify output folder" + vbCrLf
        str += vbCrLf
        str += " -filename=" + vbCrLf
        str += "             Specify filename" + vbCrLf
        str += vbCrLf
        Return str
    End Function

    Function RemoveQuote(ByVal line As String) As String

        Dim lineNoQuote As String
        lineNoQuote = line

        If line.Length > 1 Then
            While (lineNoQuote.StartsWith("""") And lineNoQuote.EndsWith("""")) Or (lineNoQuote.StartsWith("'") And lineNoQuote.EndsWith("'"))
                lineNoQuote = Mid(lineNoQuote, 2, lineNoQuote.Length - 2)
            End While
        End If
        Return lineNoQuote
    End Function

    Function AddQuote(ByVal line As String) As String
        Dim lineWithQuote As String

        If Len(line) = 0 Then Return """"""

        lineWithQuote = line

        If Not line.StartsWith("""") And Not line.EndsWith("""") Then
            lineWithQuote = """" + lineWithQuote + """"
        End If

        Return lineWithQuote
    End Function

End Module

Public Module AppConstant

    Public Enum ExitCode
        OK = 0
        WRONG_ARGUMENTS = 1
        DRAFT_SOFTERROR = 2
        DRAFT_HARDERROR = 3
        CODE_EXCEPTION = 4
    End Enum

End Module


