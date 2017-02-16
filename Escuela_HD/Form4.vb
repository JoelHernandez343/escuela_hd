Imports MetroFramework

Public Class Form4
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form4
    'Nombre del archivo: Form4
    '
    'Este form es el menú para ver LA TABLA requerida en alguna consulta predefinida.
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Dim i As Integer                         'Variable para subrutinas FOR
    Dim campo(8) As Windows.Forms.Label      'Variable que absorbe las propiedades de los label campoN
    Dim contenido(8) As Windows.Forms.Label  'Variable que absorbe las propieddaes de los label contenidoN
    Dim numCampos As Integer
    Public contador As Integer = 0
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

    Private Sub Form4_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
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
    'FUNCIÓN DE ASIGNADOS INICIALES Y ASIGNADO DE FILAS AL FORM ///////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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
        For Me.i = 1 To 7
            campo(i).Visible = False
            contenido(i).Visible = False
        Next
        Select Case Form2.tabla
            Case "Profesor"
                numCampos = 5
                For Me.i = 0 To numCampos
                    campo(i).Visible = True
                    contenido(i).Visible = True
                    campo(i).Text = Form2.profesorCampos(i)
                    contenido(i).Text = TablaLogica(0, i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Teachers"
                Else
                    Label7.Text = "Profesores"
                End If
                Exit Select
            Case "Alumno"
                numCampos = 7
                For Me.i = 0 To numCampos
                    campo(i).Visible = True
                    contenido(i).Visible = True
                    campo(i).Text = Form2.alumnoCampos(i)
                    contenido(i).Text = TablaLogica(0, i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Students"
                Else
                    Label7.Text = "Alumnos"
                End If
                Exit Select
            Case "Materia"
                numCampos = 3
                For Me.i = 0 To numCampos
                    campo(i).Visible = True
                    contenido(i).Visible = True
                    campo(i).Text = Form2.materiaCampos(i)
                    contenido(i).Text = TablaLogica(0, i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Signatures"
                Else
                    Label7.Text = "Materias"
                End If
                Exit Select
            Case "Grupo"
                numCampos = 2
                For Me.i = 0 To numCampos
                    campo(i).Visible = True
                    contenido(i).Visible = True
                    campo(i).Text = Form2.grupoCampos(i)
                    contenido(i).Text = TablaLogica(0, i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Groups"
                Else
                    Label7.Text = "Grupos"
                End If
                Exit Select
            Case "Semestre"
                numCampos = 0
                contenido(0).Text = TablaLogica(0, 0)
                If Form2.ingles = True Then
                    Label7.Text = "Semesters"
                    campo1.Text = Form2.semestreCampos(0)
                Else
                    Label7.Text = "Semestres"
                    campo1.Text = Form2.semestreCampos(0)
                End If
        End Select
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'CONEXIÓN CON BASE DE DATOS ///////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////





















    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form3.Show()
        Me.Close()
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        contador -= 1
        If contador < 0 Then
            MetroMessageBox.Show(Me, "No hay registros anteriores", "Lo sentimos :(", MessageBoxButtons.OK, MessageBoxIcon.Error)
            contador += 1
        Else
            For Me.i = 0 To numCampos
                contenido(i).Text = TablaLogica(contador, i)
            Next
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        contador += 1
        If contador > registros - 1 Then
            MetroMessageBox.Show(Me, "No hay registros siguientes", "Lo sentimos :(", MessageBoxButtons.OK, MessageBoxIcon.Error)
            contador -= 1
        Else
            For Me.i = 0 To numCampos
                contenido(i).Text = TablaLogica(contador, i)
            Next
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        contador = 0
        For Me.i = 0 To numCampos
            contenido(i).Text = TablaLogica(contador, i)
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        contador = registros - 1
        For Me.i = 0 To numCampos
            contenido(i).Text = TablaLogica(contador, i)
        Next
    End Sub
End Class