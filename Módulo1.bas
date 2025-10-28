Attribute VB_Name = "M�dulo1"
' RESTABLECER TABLA DE PROYECTO
PUBLIC SUB RESTABLECER_TABLA_PROYECTO()
    ThisWorkbook.Worksheets("PROYECTO").Range("BQ94:BQ158").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("CE94:CE158").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("CS94:CS158").Value = ""
    ThisWorkbook.Worksheets("PROYECTO").Range("DL94:DL158").Value = ""
END SUB
' RESTABLECER TABLA INSPECCION
Public Sub RESTABLECER_TABLA_INSPECCION()
    Dim i As Long
    Dim filaActual As Long
    Dim columnas As Variant
    
    With ThisWorkbook.Worksheets("INSPECCION")
        ' Limpiar rango de datos
        .Range("BA84:DB140").Value = ""
        
        ' Valores iniciales en fila 84
        .Range("EF84").Value = "FR-XXXX"
        .Range("ER84").Value = "FR-XXXX"
        .Range("FD84").Value = "FR-XXXX"
        .Range("GB84").Value = "FR-XXXX"
        .Range("GZ84").Value = "FR-XXXX"
        .Range("HX84").Value = "FR-XXXX"
        
        ' Definir columnas a actualizar
        columnas = Array("EF", "ER", "FD", "GB", "GZ", "HX")
        
        ' Bucle para filas desde 88 hasta 140, de 4 en 4
        For i = 88 To 140 Step 4
            filaActual = i - 4  ' Fila anterior (referencia en la fórmula)
            
            ' Aplicar fórmula a cada columna
            Dim col As Variant
            For Each col In columnas
                .Range(col & i).FormulaLocal = "=SI(BA" & filaActual & "<>"""";""FR-XXXX"";"""")"
            Next col
        Next i
    End With
End Sub
