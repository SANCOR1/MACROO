# MACROO

Sub Actualiza()
Application.ScreenUpdating = False
    Sheets("Consumo produccion").Select
    ActiveSheet.PivotTables("Tabla dinámica2").PivotCache.Refresh
    Columns("G:I").Select
    Selection.Copy
    Columns("K:K").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    MsgBox ("Se ha actualizado la data!"), vbInformation
    Range("K1").Select
End Sub
