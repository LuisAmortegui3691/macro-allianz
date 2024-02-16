Attribute VB_Name = "Módulo1"
Sub macroComisionesAllianz()

    ' Meidcion del flujo de trabajo
    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim duracionSegundos As Double
    Dim duracionMinutos As Double

    Dim documentosEntrada As String, documentosSalida As String
    Dim archivoCimosionesALFA As String, archivoCimosionesPlantilla As String
    Dim ultimaFilaPlanilla As Long, ultimaFilaALFASIS As Long
    Dim i As Long, j As Long
    Dim polizaPlantilla
    Dim polizaALFASIS
    Dim tieneSlash As Boolean
    Dim tieneGuion As Boolean
    Dim comisionALFASIS
    Dim comisionPlantilla
    Dim diferenciaComis
    Dim valiacionALFASISOK As String
    
    ' Registra el tiempo de inicio
    tiempoInicio = Timer
    
    documentosEntrada = ThisWorkbook.Sheets("main").Range("C2").Value
    documentosSalida = ThisWorkbook.Sheets("main").Range("C3").Value

    ' Validar que archivo hay en la carpeta Comisiones ALFASIS
    archivoCimosionesALFA = documentosEntrada & "Comisiones ALFASIS\"
    archivoCimosionesALFA = Dir(archivoCimosionesALFA)
    
    ' Validar que archivo hay en la carpeta Comisiones plantilla
    archivoCimosionesPlantilla = documentosEntrada & "Comisiones Planilla\"
    archivoCimosionesPlantilla = Dir(archivoCimosionesPlantilla)
    
    ' Abrir archivo comisiones ALFASIS
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "Comisiones ALFASIS\" & archivoCimosionesALFA
    Application.DisplayAlerts = True
    
    ' Eliminar fila 1 y 2 archivo alfasis
    Workbooks(archivoCimosionesALFA).Sheets("produccion").Rows("1:2").Delete

    ' Abrir archivo comisiones plantilla
    Application.DisplayAlerts = False
    Workbooks.OpenText Filename:=documentosEntrada & "Comisiones Planilla\" & archivoCimosionesPlantilla
    Application.DisplayAlerts = True

    ' Validar ultima fila plantilla
    ultimaFilaPlanilla = Workbooks(archivoCimosionesPlantilla).Sheets("Recibos Bancarios").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Validar ultima fila ALFASIS
    ultimaFilaALFASIS = Workbooks(archivoCimosionesALFA).Sheets("produccion").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Validar polizas
    For i = 2 To ultimaFilaPlanilla
        polizaPlantilla = Workbooks(archivoCimosionesPlantilla).Sheets("Recibos Bancarios").Range("D" & i).Value
        polizaPlantilla = Split(polizaPlantilla, "/")
        polizaPlantilla = polizaPlantilla(0)
        
        comisionPlantilla = Workbooks(archivoCimosionesPlantilla).Sheets("Recibos Bancarios").Range("F" & i).Value

        For j = 2 To ultimaFilaALFASIS
            valiacionALFASISOK = Workbooks(archivoCimosionesALFA).Sheets("produccion").Range("CD" & j).Value
            
            If valiacionALFASISOK <> "ok" Then
            
                polizaALFASIS = CStr(Workbooks(archivoCimosionesALFA).Sheets("produccion").Range("C" & j).Value)
                
                ' Verificar si el primer carácter es cero
                If Left(polizaALFASIS, 1) = "0" Then
                    ' Si es cero, eliminarlo
                    polizaALFASIS = Mid(polizaALFASIS, 2)
                    Debug.Print "Se eliminó el cero al comienzo. El valor resultante es: " & polizaALFASIS
                Else
                    ' Si no es cero, continuar con el flujo
                    Debug.Print "El valor no comienza con cero. Puedes continuar con el flujo del programa."
                End If
                
                
                
                ' Verificar si la cadena contiene "/"
                tieneSlash = InStr(polizaALFASIS, "/") > 0
                
                ' Si no contiene "/", continuar con el flujo
                If Not tieneSlash Then
                    ' Aquí puedes continuar con el flujo de tu programa
                    polizaALFASIS = polizaALFASIS
                Else
                    ' Si contiene "/", hacer algo más
                    polizaALFASIS = Split(polizaALFASIS, "/")
                    polizaALFASIS = polizaALFASIS(0)
                End If
                
                
                
                ' Verificar si la cadena contiene "-"
                tieneGuion = InStr(polizaALFASIS, "-") > 0
                
                ' Si no contiene "-", continuar con el flujo
                If Not tieneGuion Then
                    ' Aquí puedes continuar con el flujo de tu programa
                    Debug.Print "El valor no contiene un guion ('-'). Puedes continuar con el flujo del programa."
                Else
                    ' Si contiene "-", hacer algo más
                    ' Separar la cadena por el guion y obtener la primera parte
                    polizaALFASIS = Split(polizaALFASIS, "-")
                    polizaALFASIS = polizaALFASIS(0)
                End If
                
                comisionALFASIS = CDbl(Workbooks(archivoCimosionesALFA).Sheets("produccion").Range("Q" & j).Value)
                
                
                diferenciaComis = comisionPlantilla - comisionALFASIS
                
                If polizaPlantilla = polizaALFASIS Then
                    
                    
                    If diferenciaComis >= -200 Or diferenciaComis <= 200 Then
                        Workbooks(archivoCimosionesALFA).Sheets("produccion").Range("CD" & j).Value = "ok"
                        
                        Workbooks(archivoCimosionesALFA).Activate
                        Workbooks(archivoCimosionesALFA).Sheets("produccion").Rows(j).Select
                        ' Cambia el color de fondo de la selección al color #ffb3ff
                        With Selection.Interior
                            .Color = RGB(102, 255, 255) ' Código RGB para el color #ffb3ff
                        End With
                        ' Quita la selección
                        Application.CutCopyMode = False
                        
                        Exit For
                        
                    ElseIf diferenciaComis = 0 Then
                        
                        
                    End If
                End If
                
            End If
            
        Next j
        
        Workbooks(archivoCimosionesPlantilla).Sheets("Recibos Bancarios").Range("I" & i).Value = comisionALFASIS
        Workbooks(archivoCimosionesPlantilla).Sheets("Recibos Bancarios").Range("J" & i).Value = "ok"
        
    Next i
    
    ' Registra el tiempo de finalización
    tiempoFin = Timer
    ' Calcula la duración en segundos
    duracionSegundos = tiempoFin - tiempoInicio
    
    ' Muestra los resultados en la ventana inmediata (puedes ajustar esto según tus necesidades)
    Debug.Print "Duración en segundos: " & duracionSegundos

End Sub

