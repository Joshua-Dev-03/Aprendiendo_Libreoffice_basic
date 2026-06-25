REM  *****  BASIC  *****

Sub IniciarJuego()

    Dim oDoc As Object
    Dim hoja1 As Object
    Dim hoja2 As Object
    Dim hoja3 As Object

    oDoc = ThisComponent

    hoja1 = oDoc.Sheets.getByName("Hoja1")
    hoja2 = oDoc.Sheets.getByName("Hoja2")
    hoja3 = oDoc.Sheets.getByName("Hoja3")



   
    LimpiarTablero hoja1
    LimpiarTablero hoja3


    hoja2.getCellRangeByName("B1").Value = 1
    hoja2.getCellRangeByName("B2").Value = 0



    oDoc.CurrentController.setActiveSheet(hoja1)

    MsgBox "Jugador 1 colocar ejército"

End Sub





Sub FinalizarMovimientos()

    Dim oDoc As Object
    Dim hoja1 As Object
    Dim hoja2 As Object
    Dim hoja3 As Object

    oDoc = ThisComponent

    hoja1 = oDoc.Sheets.getByName("Hoja1")
    hoja2 = oDoc.Sheets.getByName("Hoja2")
    hoja3 = oDoc.Sheets.getByName("Hoja3")



    Dim fase As Integer

    fase = hoja2.getCellRangeByName("B2").Value



   
    If fase = 0 Then

        hoja2.getCellRangeByName("B2").Value = 1

        oDoc.CurrentController.setActiveSheet(hoja3)

        MsgBox "Jugador 2 colocar ejército"

        Exit Sub

    End If



   
    If fase = 1 Then

        hoja2.getCellRangeByName("B2").Value = 2

        hoja2.getCellRangeByName("B1").Value = 1

        oDoc.CurrentController.setActiveSheet(hoja1)

        MsgBox "Comienzar batalla - Turno Jugador 1"

        Exit Sub

    End If

End Sub


Sub LimpiarTablero(oHoja As Object)

    Dim fila As Integer
    Dim columna As Integer



    'Este es el tablero izquierdo

    For columna = 2 To 11

        For fila = 1 To 9

            With oHoja.getCellByPosition(columna, fila)

                .String = ""
                .CellBackColor = RGB(0,0,0)
                .CharColor = RGB(255,255,255)

            End With

        Next fila

    Next columna




    'Este es el tablero derecho

    For columna = 16 To 25

        For fila = 1 To 9

            With oHoja.getCellByPosition(columna, fila)

                .String = ""
                .CellBackColor = RGB(0,0,0)
                .CharColor = RGB(255,255,255)

            End With

        Next fila

    Next columna

End Sub


Sub LanzarAtaque()

    Dim oDoc As Object
    Dim hojaControl As Object
    Dim hojaJ1 As Object
    Dim hojaJ2 As Object
    Dim hojaActual As Object

    oDoc = ThisComponent

    hojaActual = oDoc.CurrentController.ActiveSheet

    hojaControl = oDoc.Sheets.getByName("Hoja2")
    hojaJ1 = oDoc.Sheets.getByName("Hoja1")
    hojaJ2 = oDoc.Sheets.getByName("Hoja3")



 

    Dim fase As Integer

    fase = hojaControl.getCellRangeByName("B2").Value

    If fase <> 2 Then

        MsgBox "La batalla aún no comienza"
        Exit Sub

    End If



   

    Dim turno As Integer

    turno = hojaControl.getCellRangeByName("B1").Value



    

    Dim seleccion As Object

    seleccion = oDoc.getCurrentSelection()



    Dim fila As Integer
    Dim columna As Integer

    fila = seleccion.RangeAddress.StartRow
    columna = seleccion.RangeAddress.StartColumn


If turno = 1 Then

    If columna < 16 Or columna > 25 Then

        MsgBox "Jugador 1 debe atacar en el tablero derecho"
        Exit Sub

    End If

ElseIf turno = 2 Then

    If columna < 2 Or columna > 11 Then

        MsgBox "Jugador 2 debe atacar en el tablero izquierdo"
        Exit Sub

    End If

End If



    If fila < 1 Or fila > 9 Then

        MsgBox "Debes atacar dentro del tablero"
        Exit Sub

    End If



    If seleccion.String = "X" Or seleccion.String = "O" Then

        MsgBox "Ya atacaste aquí"
        Exit Sub

    End If



    Dim colEnemiga As Integer

   If turno = 1 Then

    ' Q:Z -> C:L
    colEnemiga = columna - 14

Else

    ' C:L -> C:L
    colEnemiga = columna

End If



    If turno = 1 Then

        Dim objetivo As Object

        objetivo = hojaJ2.getCellByPosition(colEnemiga, fila)



        If objetivo.String <> "" Then

            seleccion.CellBackColor = RGB(255,0,0)
            seleccion.String = "X"

            MsgBox "¡Impacto!"

        Else

            seleccion.CellBackColor = RGB(120,120,120)
            seleccion.String = "O"

            MsgBox "Agua"

        End If



        hojaControl.getCellRangeByName("B1").Value = 2

        oDoc.CurrentController.setActiveSheet(hojaControl)

        MsgBox "Jugador 1 ha atacado" & CHR(10) & _
               "Turno del Jugador 2"

        Exit Sub

    End If


    If turno = 2 Then

        Dim objetivo2 As Object

        objetivo2 = hojaJ1.getCellByPosition(colEnemiga, fila)



        If objetivo2.String <> "" Then

            seleccion.CellBackColor = RGB(255,0,0)
            seleccion.String = "X"

            MsgBox "¡Impacto!"

        Else

            seleccion.CellBackColor = RGB(120,120,120)
            seleccion.String = "O"

            MsgBox "Agua"

        End If



        hojaControl.getCellRangeByName("B1").Value = 1

        oDoc.CurrentController.setActiveSheet(hojaControl)

        MsgBox "Jugador 2 ha atacado" & CHR(10) & _
               "Turno del Jugador 1"

        Exit Sub

    End If

End Sub

Sub ContinuarTurno()

    Dim oDoc As Object
    Dim hojaControl As Object
    Dim hoja1 As Object
    Dim hoja3 As Object

    oDoc = ThisComponent

    hojaControl = oDoc.Sheets.getByName("Hoja2")
    hoja1 = oDoc.Sheets.getByName("Hoja1")
    hoja3 = oDoc.Sheets.getByName("Hoja3")



    Dim turno As Integer

    turno = hojaControl.getCellRangeByName("B1").Value



    If turno = 1 Then

        oDoc.CurrentController.setActiveSheet(hoja1)

        MsgBox "Turno del Jugador 1"

    Else

        oDoc.CurrentController.setActiveSheet(hoja3)

        MsgBox "Turno del Jugador 2"

    End If

End Sub
