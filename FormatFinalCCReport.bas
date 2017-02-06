Attribute VB_Name = "FormatFinalCCReport"
Sub formatGiftAdmin()

Application.ScreenUpdating = False

ActiveSheet.Range("A1").Activate

'gift amount = currency
Range("H:H").NumberFormat = "$#,##0.00"

'format header row and columns
Range("1:1").Select
    With Selection
        .EntireColumn.Interior.ColorIndex = xlNone
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .EntireColumn.RowHeight = 20
        .Font.Bold = True
        .EntireColumn.Font.Size = 12
        .Interior.ColorIndex = 36
        .EntireColumn.AutoFit
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With


'Move returns to bottom and delete from top list

Dim tran1 As Range
Dim tranLast As Range
Dim tranAll As Range
Dim lastRow As Integer
Dim cellRng As Range
Dim returnVoidCount As Integer

lastRow1 = Columns("H:H").Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 1
lastRow2 = lastRow1

Set tran1 = Range("X2")
Set tranLast = Range("A1").Offset(lastRow1 - 1, 23)
Set tranAll = Range(tran1, tranLast)
returnVoidCount = 0

'move voids and returns below
For Each cell In tranAll
    cell.Activate
    If cell.Value = "Void" Then
        cell.EntireRow.Copy
        Range("A1").Offset(lastRow2 + 2, 0).Activate
        ActiveCell.EntireRow.Insert
        ActiveCell.EntireRow.Font.Color = vbBlue
        lastRow2 = lastRow2 + 1
        returnVoidCount = returnVoidCount + 1
    ElseIf cell.Value = "Return" Then
        cell.EntireRow.Copy
        Range("A1").Offset(lastRow2 + 2, 0).Activate
        ActiveCell.EntireRow.Insert
        ActiveCell.EntireRow.Font.Color = vbRed
        lastRow2 = lastRow2 + 1
        returnVoidCount = returnVoidCount + 1
    End If
Next cell

'delete voids and returns above
    Dim c As Range
   
    Do
        Set c = tranAll.Find("Return", LookIn:=xlValues)
        If Not c Is Nothing Then
            c.EntireRow.Delete
        Else
            Set c = tranAll.Find("Void", LookIn:=xlValues)
            If Not c Is Nothing Then c.EntireRow.Delete
        End If
    Loop While Not c Is Nothing

'Find gifts range
    lastRow1 = lastRow1 - returnVoidCount
    Set firstGift = Range("H2")
    Range("A1").Offset(lastRow1 - 2, 7).Activate
    Set lastGift = Range(ActiveCell, ActiveCell)
    Set gifts = Range(firstGift, lastGift)
    
    giftTotal = 0
    
    'add up total gift amount
    For Each cell In gifts
        cell.Activate
        If ActiveCell.Offset(0, 16) = "Sale" Then
            giftTotal = giftTotal + cell.Value
        End If
    Next cell
    
    gifts.Activate
        
    
    'Format Total line
    Range("A1").Offset(lastRow1, 6).Select
    Selection = "Spreadsheet Total:"
    Selection.HorizontalAlignment = xlRight
    Range("A1").Offset(lastRow1, 7).Formula = "=sum(" + gifts.Address + ")"  'giftTotal
    Range("A1").Offset(lastRow1, 15) = Range("A1").Offset(lastRow1 + 1, 13).Value
    Range("A1").Offset(lastRow1 + 1, 15).Value = ""
    Range("A1").Offset(lastRow1, 16) = Range("A1").Offset(lastRow1 + 1, 14).Value
    Range("A1").Offset(lastRow1 + 1, 16).Value = ""

    Columns("G:G").Find("Spreadsheet Total:", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).EntireRow.Select
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        Selection.Font.Bold = True
    
    'Autofit columns after totaling amounts.
    Range("1:1").Select
        With Selection
            .EntireColumn.AutoFit
    
        End With
        
    Range("A1").Activate
    
    Application.ScreenUpdating = True

End Sub


