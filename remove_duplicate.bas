Attribute VB_Name = "remove_duplicate"
Sub CopyAndFillLatestModifiedRowsByDate_Fast_v3()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim data As Variant, result As Variant, headers As Variant
    Dim dateDict As Object, groupedRows As Object
    Dim i As Long, j As Long, destRow As Long
    Dim dateKey As String, modDate As Double
    Dim rowToKeep As Long
    Dim dateCol As Long, modCol As Long
    Dim colStart As Long, colEnd As Long
    Dim rowValues() As Variant
    Dim key As Variant, vRow As Variant
    Dim numCols As Long
    Dim specialStart As Long, specialEnd As Long
    Dim fillValue As Variant
    Dim latestMod As Double, bestRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '==== Settings ====
    Set wsSource = ThisWorkbook.Sheets(1)
    colStart = 2        ' Column B
    colEnd = 21         ' Column U
    dateCol = 2         ' DATE_ (B)
    modCol = 4          ' Modified_Date (D)
    specialStart = 12   ' Column L
    specialEnd = 21     ' Column U
    numCols = colEnd - colStart + 1

    '==== Load data ====
    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.Rows.Count, dateCol).End(xlUp).Row
    data = wsSource.Range(wsSource.Cells(1, colStart), wsSource.Cells(lastRow, colEnd)).Value
    headers = Application.Index(data, 1, 0)

    '==== Prepare destination ====
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("Filtered_Latest_Modified")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSource)
        wsDest.Name = "Filtered_Latest_Modified"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0

    '==== Initialize dictionaries ====
    Set dateDict = CreateObject("Scripting.Dictionary")
    Set groupedRows = CreateObject("Scripting.Dictionary")

    '==== Group rows and find latest modified ====
    For i = 2 To UBound(data, 1)
        If IsDate(data(i, dateCol - colStart + 1)) And IsDate(data(i, modCol - colStart + 1)) Then
            dateKey = Format$(CDate(data(i, dateCol - colStart + 1)), "yyyy-mm-dd")
            modDate = CDbl(CDate(data(i, modCol - colStart + 1)))

            If Not dateDict.Exists(dateKey) Then
                dateDict(dateKey) = Array(modDate, i)
            ElseIf modDate > dateDict(dateKey)(0) Then
                dateDict(dateKey) = Array(modDate, i)
            End If

            If Not groupedRows.Exists(dateKey) Then
                Set groupedRows(dateKey) = CreateObject("System.Collections.ArrayList")
            End If
            groupedRows(dateKey).Add i
        End If
    Next i

    '==== Prepare result ====
    ReDim result(1 To dateDict.Count, 1 To numCols)
    destRow = 1

    '==== Process each date group ====
    For Each key In dateDict.Keys
        rowToKeep = dateDict(key)(1)
        ReDim rowValues(1 To numCols)

        ' Load the "latest modified" row as base
        For j = 1 To numCols
            rowValues(j) = data(rowToKeep, j)
        Next j

        '=== Fill B–K (Cols 2–11) from earlier non-empty ===
        For j = 1 To (specialStart - colStart)
            If Trim(CStr(rowValues(j))) = "" Then
                For Each vRow In groupedRows(key)
                    If vRow = rowToKeep Then Exit For
                    If Trim(CStr(data(vRow, j))) <> "" Then
                        rowValues(j) = data(vRow, j)
                        Exit For
                    End If
                Next vRow
            End If
        Next j

        '=== Fill L–U (Cols 12–21) using most recent non-empty value ===
        For j = (specialStart - colStart + 1) To (specialEnd - colStart + 1)
            latestMod = 0
            fillValue = ""
            For Each vRow In groupedRows(key)
                If Trim(CStr(data(vRow, j))) <> "" Then
                    If IsDate(data(vRow, modCol - colStart + 1)) Then
                        modDate = CDbl(CDate(data(vRow, modCol - colStart + 1)))
                        If modDate > latestMod Then
                            latestMod = modDate
                            fillValue = data(vRow, j)
                        End If
                    End If
                End If
            Next vRow
            If fillValue <> "" Then rowValues(j) = fillValue
        Next j

        '=== Store filled row ===
        For j = 1 To numCols
            result(destRow, j) = rowValues(j)
        Next j

        destRow = destRow + 1
    Next key

    '==== Output ====
    wsDest.Range("A1").Resize(1, UBound(headers)).Value = headers
    wsDest.Range("A2").Resize(UBound(result, 1), UBound(result, 2)).Value = result

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Done: Latest modified rows copied and filled (L–U picks newest non-empty per column)."
End Sub

