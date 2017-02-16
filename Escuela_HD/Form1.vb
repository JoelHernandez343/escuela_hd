Public Class Form1
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form1
    'Nombre del archivo: Form1
    '
    'Este form es el menpu para EDITAR la base de datos, y desde aquí se puede acceder a:
    'Form5
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////


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
    Private Sub Form1_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
        If Form2.tema = True Then
            Me.BackColor = Color.FromArgb(32, 32, 32)
        Else
            Me.BackColor = Color.FromArgb(32, 103, 179)
        End If
        If Form2.ingles = True Then
            Button1.Text = " Back"
            Label7.Text = "Edit"
            Label1.Text = "Choose the table you want to edit"
            Label2.Text = "Teachers"
            Label3.Text = "Students"
            Label4.Text = "Signatures"
            Label5.Text = "Groups"
            Label6.Text = "Semesters"
        Else
            Button1.Text = " Regresar"
            Label7.Text = "Editar"
            Label1.Text = "Elija la tabla que quiere editar"
            Label2.Text = "Profesores"
            Label3.Text = "Estudiantes"
            Label4.Text = "Materias"
            Label5.Text = "Grupos"
            Label6.Text = "Semestres"
        End If
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form2.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Form5.contador = 0
        Form2.tabla = "Materia"
        Form2.tablaElegida = Form2.materiaCamposT
        leer(Form2.tabla, Form2.tablaElegida, 4)
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Form5.contador = 0
        Form2.tabla = "Profesor"
        Form2.tablaElegida = Form2.profesorCamposT
        leer(Form2.tabla, Form2.tablaElegida, 6)
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Form5.contador = 0
        Form2.tabla = "Alumno"
        Form2.tablaElegida = Form2.alumnoCamposT
        leer(Form2.tabla, Form2.tablaElegida, 8)
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Form5.contador = 0
        Form2.tabla = "Grupo"
        Form2.tablaElegida = Form2.grupoCamposT
        leer(Form2.tabla, Form2.tablaElegida, 3)
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Form5.contador = 0
        Form2.tabla = "Semestre"
        Form2.tablaElegida = Form2.semestreCamposT
        leer(Form2.tabla, Form2.tablaElegida, 1)
        Form5.Show()
        Me.Hide()
    End Sub
End Class
