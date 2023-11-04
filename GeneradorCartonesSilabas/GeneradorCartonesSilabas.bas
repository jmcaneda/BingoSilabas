Attribute VB_Name = "Módulo1"
Sub GenerarBingoSilabas()
    Dim ppt As PowerPoint.Application
    Set ppt = New PowerPoint.Application
    
    Dim presentacion As PowerPoint.Presentation
    Set presentacion = ppt.Presentations.Add
    
    Dim silabas As Variant
    silabas = Array("pa", "pe", "pi", "po", "pu", "la", "le", "li", "lo", "lu", _
                    "ma", "me", "mi", "mo", "mu", "sa", "se", "si", "so", "su")
    
    Dim i As Integer, j As Integer
    Dim fila As Integer, columna As Integer
    Dim diapositiva As PowerPoint.Slide
    Dim tabla As PowerPoint.Table
    Dim texto As String
    Dim celdasEnBlanco(1 To 2, 1 To 2) As Integer  ' Para almacenar las posiciones de las celdas en blanco
    
    For i = 1 To 17
        Set diapositiva = presentacion.Slides.Add(i, 1)
        
        ' Agregar título a la diapositiva
        'Dim shape As PowerPoint.shape
        'Set shape = diapositiva.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
                          Left:=150, Top:=50, Width:=600, Height:=100)
        'shape.TextFrame.TextRange.Text = "Bingo de Sílabas"
        'shape.TextFrame.TextRange.Font.Size = 36
        'shape.TextFrame.TextRange.Font.Name = "Arial"
        'shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        
        ' Agregar tabla a la diapositiva
        Set tabla = diapositiva.Shapes.AddTable(NumRows:=3, NumColumns:=4, Left:=50, Top:=50, Width:=850, Height:=450).Table  ' Altura ajustada a 450
        
        ' Seleccionar aleatoriamente dos celdas para estar en blanco
        For j = 1 To 2
            Do
                celdasEnBlanco(j, 1) = Int(Rnd() * 3) + 1  ' Fila aleatoria
                celdasEnBlanco(j, 2) = Int(Rnd() * 4) + 1  ' Columna aleatoria
                ' Asegurar que las dos celdas en blanco sean únicas
            Loop While (j = 2 And (celdasEnBlanco(1, 1) = celdasEnBlanco(2, 1) And celdasEnBlanco(1, 2) = celdasEnBlanco(2, 2)))
        Next j
        
        ' Inicializar un diccionario para rastrear las sílabas usadas en esta diapositiva
        Dim silabasUsadas As Object
        Set silabasUsadas = CreateObject("Scripting.Dictionary")
        
        For fila = 1 To 3
            For columna = 1 To 4
                ' Verificar si la celda actual debe estar en blanco
                Do
                    texto = silabas(Int(Rnd() * UBound(silabas) + 1))
                Loop While silabasUsadas.Exists(texto)  ' Reintentar si la sílaba ya se ha usado en esta diapositiva
                
                For j = 1 To 2
                    If fila = celdasEnBlanco(j, 1) And columna = celdasEnBlanco(j, 2) Then
                        texto = ""
                        Exit For
                    End If
                Next j
                
                ' Agregar texto a la celda
                tabla.Cell(fila, columna).shape.TextFrame.TextRange.Text = texto
                tabla.Cell(fila, columna).shape.TextFrame.TextRange.Font.Size = 48
                tabla.Cell(fila, columna).shape.TextFrame.TextRange.Font.Name = "Arial"  ' Tipo de letra establecido a Arial
                tabla.Cell(fila, columna).shape.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter  ' Alineación horizontal al centro
                tabla.Cell(fila, columna).shape.TextFrame.VerticalAnchor = msoAnchorMiddle  ' Alineación vertical al medio
                
                ' Añadir la sílaba al diccionario de sílabas usadas si no está en blanco
                If texto <> "" Then silabasUsadas.Add texto, Nothing
            Next columna
        Next fila
        
        ' Limpiar el diccionario de sílabas usadas para la siguiente diapositiva
        Set silabasUsadas = Nothing
    Next i
    
    ppt.Visible = True
End Sub


