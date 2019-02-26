VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocktest()
Dim WS As Worksheet
Dim ticker As String
Dim totalvolume As Double
Dim lopen As Double
Dim lclose As Double
Dim yearlychange As Double
Dim nextopen As Double
Dim percentagechange As Double
Dim i As Long
Dim lastrow As Long


For Each WS In Worksheets
WS.Activate
Dim stocktable As Integer
stocktable = 2


WS.Range("L1").Value = "Ticker"
WS.Range("M1").Value = "Yearly Change"
WS.Range("N1").Value = "Percentage Change"
WS.Range("O1").Value = "Total Stock Volume"

totalvolume = 0
lopen = WS.Cells(2, 3)
yearlychange = 0


For i = 2 To 78000

If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    ticker = WS.Cells(i, 1).Value
    lclose = WS.Cells(i, 6).Value
    yearlychange = (lclose - lopen)
    percentagechange = Round((yearlychange / lopen) * 100, 2)
    totalvolume = totalvolume + WS.Cells(i, 7).Value
    
    WS.Range("M" & stocktable).Value = yearlychange
    WS.Range("N" & stocktable).Value = percentagechange
    
    lopen = WS.Cells(i + 1, 3)
    WS.Range("L" & stocktable).Value = ticker
    WS.Range("O" & stocktable).Value = totalvolume
    

    stocktable = stocktable + 1
    totalvolume = 0
    
    Select Case yearlychange
               Case Is > 0
               WS.Range("M" & stocktable).Interior.ColorIndex = 4
               Case Is < 0
                   WS.Range("M" & stocktable).Interior.ColorIndex = 3
               Case Else
                   WS.Range("M" & stocktable).Interior.ColorIndex = 0
           End Select


Else
 totalvolume = totalvolume + WS.Cells(i, 7).Value
 
End If
 
Next i
