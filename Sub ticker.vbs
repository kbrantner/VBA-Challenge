Sub ticker()

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"

Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percentchange As Double

Dim ticker As String


Dim totalstock As Double
totalstock = 0

Dim stockrow As Integer
stockrow = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    ticker = Cells(i, 1).Value
    totalstock = totalstock + Cells(i, 7).Value
    Range("I" & stockrow).Value = ticker
    Range("L" & stockrow).Value = totalstock
    stockrow = stockrow + 1
    totalstock = 0
    
    Else
    totalstock = totalstock + Cells(i, 7).Value
    
    End If

Next i


End Sub
