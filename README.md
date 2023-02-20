# excel_vba
Códigos de pequeños y distintos trabajos hechos en Excel VBA.
Sub QuitaRegistro()

'Elimina registro de la hoja cuando cumple condición y lo guarda en otra hoja

'Subrutina Código VBA para trasladar registro a otra hoja en función de un criterio, por ejemplo, que esté ‘pagado’, y eliminar ese registro de la tabla principal (máster).

Dim Ht1, Ht2 As Worksheet
Set Ht1 = Worksheets(«Hoja2»)
Set Ht2 = Worksheets(«Hoja3»)
Ht2.Select
nFilas2 = Cells.SpecialCells(xlLastCell).Row
Ht1.Select
nFilas1 = Cells.SpecialCells(xlLastCell).Row
'Trasladar Pagadas
For i = 2 To nFilas1
	If Cells(i, «B») = «Pagada» Then
		nFilas2 = nFilas2 + 1
		Ht1.Rows(i).Copy Destination:=Ht2.Rows(nFilas2)
	End If
Next i

'Depurar Pagadas
For i = 2 To nFilas1
	Volver:
	k = k + 1
	If k > nFilas1 Then Exit Sub
	
	If Cells(i, «B») = «Pagada» Then ‘Criterio
		Rows(i).Delete
		GoTo Volver:
	End If
Next i

MsgBox «Finalizado.»

End Sub
