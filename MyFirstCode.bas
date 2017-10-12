Attribute VB_Name = "MyFirstCode"

Option Explicit
'
'Sub createRectangle()
'ActivePage.shapes.DrawRectangle(0,0,100,100)
'
'
'End Sub

Sub colorShapes()
Dim shps As Shapes
Dim shp As Shape

shps = ActivePage.Shapes
For Each shp In shps
shp.Cells("FillForegnd").Formula = RGB(255, 0, 0)
Next shp

End Sub
