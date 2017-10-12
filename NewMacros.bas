Attribute VB_Name = "NewMacros"
Option Explicit

Sub Macro1()

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    Application.Windows.ItemEx("Visio2GitTest.vsdm").Activate
    Application.ActiveWindow.Page.Drop Application.Documents.Item("BASIC_M.VSSX").Masters.ItemU("Rectangle"), 4.640045, 8.605174

    Dim UndoScopeID1 As Long
    UndoScopeID1 = Application.BeginUndoScope("Size Object")
    Application.ActiveWindow.Page.Shapes.ItemFromID(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinX).FormulaU = "83.142857142857 mm"
    Application.ActiveWindow.Page.Shapes.ItemFromID(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormPinY).FormulaU = "192.53571428571 mm"
    Application.ActiveWindow.Page.Shapes.ItemFromID(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormWidth).FormulaU = "109.42857142857 mm"
    Application.ActiveWindow.Page.Shapes.ItemFromID(1).CellsSRC(visSectionObject, visRowXFormOut, visXFormHeight).FormulaU = "82.071428571429 mm"
    Application.EndUndoScope UndoScopeID1, True

    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices

End Sub
