Attribute VB_Name = "modGrid"
Option Explicit

Public Sub makeGrid(grid As MSFlexGrid, titulos As Variant, fixedCols As Integer, fixedRows As Integer, mode As Integer)
Dim columnas As Integer
Dim columna As Integer
    
    columnas = UBound(titulos) + 1
    grid.Clear
    grid.SelectionMode = mode
    grid.ScrollBars = flexScrollBarVertical
    grid.Rows = 2
    grid.Cols = columnas
    grid.fixedCols = fixedCols
    grid.fixedRows = fixedRows
    grid.Rows = 1
    For columna = 0 To columnas - 1
        grid.TextMatrix(0, columna) = titulos(columna)(0)
        grid.ColWidth(columna) = titulos(columna)(1)
    Next columna

End Sub

Public Sub letCheckCell(grid As MSFlexGrid, row As Integer, col As Integer, value As Boolean)

    grid.row = row
    grid.col = col
    grid.CellFontName = "Wingdings"
    grid.CellFontSize = 14
    grid.CellAlignment = 4
    grid.Text = IIf(value, Chr(254), Chr(113))

End Sub

Public Function getCheckCell(grid As MSFlexGrid, row As Integer, col As Integer) As Boolean
    
    getCheckCell = IIf(Asc(grid.TextMatrix(row, col)) = 254, True, False)

End Function

Public Function topGrid(grid As MSFlexGrid) As Long
    
    topGrid = grid.Top + grid.CellTop

End Function

Public Function leftGrid(grid As MSFlexGrid) As Long
    
    leftGrid = grid.Left + grid.CellLeft

End Function

Public Sub setColourGrid(grid As MSFlexGrid, row As Integer, col As Integer, pColor As Variant)
    
    grid.row = row
    grid.col = col
    grid.CellBackColor = modColor.variant2RGB(pColor)

End Sub

Public Function getColourGrid(grid As MSFlexGrid, row As Integer, col As Integer) As Variant
    
    grid.row = row
    grid.col = col
    getColourGrid = grid.CellBackColor

End Function

Public Sub swapColor(grid As MSFlexGrid, row As Integer, col As Integer, color1 As Variant, color2 As Variant)

    setColourGrid grid, row, col, IIf(getColourGrid(grid, row, col) = color1, color2, color1)
    
End Sub

Public Sub fieldAdd(line As String, field As Variant)

    line = line & field & Chr(9)
    
End Sub

Public Function array2itemGrid(pArray As Variant) As String
Dim linea As String

Dim ciclo As Integer

    linea = ""
    
    For ciclo = LBound(pArray) To UBound(pArray)
        fieldAdd linea, pArray(ciclo)
    Next ciclo
    
    array2itemGrid = linea
    
End Function

Public Sub setTextBox(grid As MSFlexGrid, pTextBox As TextBox)

    With pTextBox
        .Visible = False
        .Left = leftGrid(grid)
        .Top = topGrid(grid)
        .Width = grid.CellWidth
        .Height = grid.CellHeight
        .Visible = True
        .Text = grid.Text
        .SetFocus
    End With
    
End Sub

