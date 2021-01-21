Public Sub Border_Top(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeTop, RGB(red, green, blue))
End Sub

Public Sub Border_Bottom(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeBottom, RGB(red, green, blue))
End Sub

Public Sub Border_Left(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeLeft, RGB(red, green, blue))
End Sub

Public Sub Border_Right(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeRight, RGB(red, green, blue))
End Sub

Public Sub Border_DiagonalDown(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlDiagonalDown, RGB(red, green, blue))
End Sub

Public Sub Border_DiagonalUp(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlDiagonalUp, RGB(red, green, blue))
End Sub

Public Sub Border_Lattice(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeTop, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeBottom, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeLeft, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeRight, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlInsideHorizontal, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlInsideVertical, RGB(red, green, blue))
End Sub

Public Sub Border_OuterFrame(sheetName As String, cellAddress As String, borderType As Integer, red As Integer, green As Integer, blue As Integer)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeTop, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeBottom, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeLeft, RGB(red, green, blue))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), borderType,xlEdgeRight, RGB(red, green, blue))
End Sub

Public Sub Border_Delete(sheetName As String, cellAddress As String)
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlEdgeTop, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlEdgeBottom, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlEdgeLeft, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlEdgeRight, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlInsideHorizontal, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlInsideVertical, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlDiagonalDown, RGB(0, 0, 0))
    Call BorderWrite(Worksheets(sheetName).Range(cellAddress), 0, xlDiagonalUp, RGB(0, 0, 0))
End Sub

Private Sub BorderWrite(cellRange As Range, borderType As Integer, borderPosition As Integer, borderColor As Long)

    Select Case borderType
    Case 0
        With cellRange.Borders(borderPosition)
        .LineStyle = xlLineStyleNone
        End With
    Case 1
        With cellRange.Borders(borderPosition)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = borderColor
        End With
    Case 2
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDot
        .Weight = xlThin
        .Color = borderColor
        End With
    Case 3
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDashDotDot
        .Weight = xlThin
        .Color = borderColor
        End With
    Case 4
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDashDot
        .Weight = xlThin
        .Color = borderColor
        End With
    Case 5
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDash
        .Weight = xlThin
        .Color = borderColor
        End With
    Case 6
        With cellRange.Borders(borderPosition)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = borderColor
        End With
    Case 7
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDashDotDot
        .Weight = xlMedium
        .Color = borderColor
        End With
    Case 8
        With cellRange.Borders(borderPosition)
        .LineStyle = xlSlantDashDot
        .Weight = xlMedium
        .Color = borderColor
        End With
    Case 9
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDashDot
        .Weight = xlMedium
        .Color = borderColor
        End With
    Case 10
        With cellRange.Borders(borderPosition)
        .LineStyle = xlDash
        .Weight = xlMedium
        .Color = borderColor
        End With
    Case 11
        With cellRange.Borders(borderPosition)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = borderColor
        End With
    Case 12
        With cellRange.Borders(borderPosition)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = borderColor
        End With
    Case 13
         With cellRange.Borders(borderPosition)
        .LineStyle = xlDouble
        .Weight = xlThick
        .Color = borderColor
        End With
    End Select
End Sub

