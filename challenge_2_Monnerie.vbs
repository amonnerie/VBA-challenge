Sub StockAnalysis()
    
    'Declaration of variables
    Dim openP As Double
    Dim closeP As Double
    Dim Ticker As String
    Dim iTicker As String
    Dim vol As LongLong
    Dim ws As Worksheet
    
    Dim j As Long
    Dim i As Long
    
    'Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
    'ws.Select
    'Assign variables
    openP = 0
    closeP = 0
    Ticker = ""
    iTicker = ""
    j = 1
    'MsgBox ("in For Each loop")
    'the loop
        For i = 2 To 760000
            vol = 0
            openP = ws.Cells(i, 3).Value 'Start on C2
            iTicker = ws.Cells(i, 1).Value 'Start on A2
            Ticker = iTicker
            If Ticker = "" Then
                Exit For
            End If
            j = j + 1
            Do While iTicker = Ticker
                vol = vol + ws.Cells(i, 7).Value
                nextT = ws.Cells(i + 1, 1).Value
                If Not iTicker = nextT Then
                    closeP = ws.Cells(i, 6).Value
                    Exit Do
                End If
               i = i + 1
               iTicker = nextT
               counter = counter + 1
            Loop
            
            'Display outputs
            ws.Cells(j, 9).Value = Ticker
            ws.Cells(j, 10).Value = closeP - openP
            ws.Cells(j, 11).Value = ((closeP - openP) / openP)
            ws.Cells(j, 12).Value = vol
        Next i
    Next ws
    'Application.ScreenUpdating = True
    
End Sub

Sub FindGreatest()
    
    'Declarations
    Dim i As Long
    Dim greatestIn As Double
    Dim greatestDe As Double
    Dim greatestVol As LongLong
    Dim current1 As Double
    Dim current2 As LongLong
    Dim Ticker As String
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    'Assignments
    greatestIn = 0
    greatestDe = 0
    Ticker = ""
'Find the greatest % increase
    For i = 2 To 3001
        current1 = ws.Cells(i, 11).Value
        If current1 > greatestIn Then
            greatestIn = current1
            Ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    ws.Cells(2, 16).Value = Ticker
    ws.Cells(2, 17).Value = greatestIn
    
    'Find the greatest % decrease
    Ticker = ""
    For i = 2 To 3001
        current1 = ws.Cells(i, 11).Value
        If current1 < greatestDe Then
            greatestDe = current1
            Ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    ws.Cells(3, 16).Value = Ticker
    ws.Cells(3, 17).Value = greatestDe
    
    'Find the greatest total volume
    Ticker = ""
    greatestVol = 0
    For i = 2 To 3001
        current2 = ws.Cells(i, 12).Value
        If current2 > greatestVol Then
            greatestVol = current2
            Ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    ws.Cells(4, 16).Value = Ticker
    ws.Cells(4, 17).Value = greatestVol
Next ws
End Sub
