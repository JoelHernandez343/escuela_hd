Public Class FormConsultaCustom
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: FormConsultaCustom
    'Nombre del archivo: Form7
    '
    'Este form es el menú para ver LA TABLA requerida en alguna consulta PERSONALIZADA
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Dim i As Integer                         'Variable para subrutinas FOR
    Dim campo(8) As Windows.Forms.Label      'Variable que absorbe las propiedades de los label campoN
    Dim contenido(8) As Windows.Forms.Label  'Variable que absorbe las propieddaes de los label contenidoN

    '//////////////////////////////////////////////////////////////////////
    'FUNCIONES PARA MOVER EL FORM SIN BORDE ///////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Dim ex, ey As Integer
    Dim Arrastre As Boolean
    'Estas tres subrutinas permiten desplazar el formulario desde el FORM
    Private Sub Me_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        ex = e.X
        ey = e.Y
        Arrastre = True
    End Sub
    Private Sub Me_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        Arrastre = False
    End Sub
    Private Sub Me_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        'Si el formulario no tiene borde (FormBorderStyle = none) la siguiente linea funciona bien
        If Arrastre Then Me.Location = Me.PointToScreen(New Point(MousePosition.X - Me.Location.X - ex, MousePosition.Y - Me.Location.Y - ey))
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'FUNCIONES PARA VERIFICAR EL TEMA Y EL LENGUAJE ///////////////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Me_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
        If Form2.tema = True Then
            Me.BackColor = Color.FromArgb(32, 32, 32)
        Else
            Me.BackColor = Color.FromArgb(32, 103, 179)
        End If
        If Form2.ingles = True Then
            Button1.Text = " Back"
        Else
            Button1.Text = " Regresar"
        End If
    End Sub
    '//////////////////////////////////////////////////////////////////////
    'CONEXIÓN CON BASE DE DATOS ///////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////





















    '//////////////////////////////////////////////////////////////////////
    'FUNCIÓN DE ASIGNADOS INICIALES Y ASIGNADO DE FILAS AL FORM ///////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Form4_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        campo(0) = campo1
        campo(1) = campo2
        campo(2) = campo3
        campo(3) = campo4
        campo(4) = campo5
        campo(5) = campo6
        campo(6) = campo7
        campo(7) = campo8
        contenido(0) = contenido1
        contenido(1) = contenido2
        contenido(2) = contenido3
        contenido(3) = contenido4
        contenido(4) = contenido5
        contenido(5) = contenido6
        contenido(6) = contenido7
        contenido(7) = contenido8
        With FormCategorias
            For Me.i = 1 To 7
                campo(i).Visible = False
                contenido(i).Visible = False
            Next
            For Me.i = 0 To .conteo - 1
                campo(i).Visible = True
                contenido(i).Visible = True
            Next
            For Me.i = 0 To .conteo - 1
                Select Case .tableElegida(i)
                    Case 1
                        campo(i).Text = Form2.profesorCampos(.campoElegido(i) - 1)
                        Exit Select
                    Case 2
                        campo(i).Text = Form2.alumnoCampos(.campoElegido(i) - 1)
                        Exit Select
                    Case 3
                        campo(i).Text = Form2.materiaCampos(.campoElegido(i) - 1)
                        Exit Select
                    Case 4
                        campo(i).Text = Form2.grupoCampos(.campoElegido(i) - 1)
                        Exit Select
                    Case 5
                        campo(i).Text = Form2.semestreCampos(.campoElegido(i) - 1)
                        Exit Select
                End Select
            Next
        End With
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        FormCategorias.Show()
        Me.Close()
    End Sub
End Class