Public Class Form3
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form3
    'Nombre del archivo: Form3
    '
    'Este form es el menpu para CONSULTAR la base de datos, y desde aquí se puede acceder a:
    'Form4, FormCategorias
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
    Private Sub Form3_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
        If Form2.tema = True Then
            Me.BackColor = Color.FromArgb(32, 32, 32)
        Else
            Me.BackColor = Color.FromArgb(32, 103, 179)
        End If
        If Form2.ingles = True Then
            Button1.Text = " Back"
            Label7.Text = "Consult"
            Label1.Text = "Choose the way to consult the information"
            Label2.Text = "Teachers"
            Label3.Text = "Students"
            Label4.Text = "Signatures"
            Label5.Text = "Groups"
            Label6.Text = "Semesters"
            Label8.Text = "Or custom"
            Button8.Text = "Custom Search"
        Else
            Button1.Text = " Regresar"
            Label7.Text = "Consultar"
            Label1.Text = "Elija la forma de su consulta"
            Label2.Text = "Profesores"
            Label3.Text = "Estudiantes"
            Label4.Text = "Materias"
            Label5.Text = "Grupos"
            Label6.Text = "Semestres"
            Label8.Text = "O personalizada"
            Button8.Text = "Búsqueda por criterios"
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
        Form4.contador = 0
        Form2.tabla = "Materia"
        Form2.tablaElegida = Form2.materiaCamposT
        leer(Form2.tabla, Form2.tablaElegida, 4)
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Form4.contador = 0
        Form2.tabla = "Profesor"
        Form2.tablaElegida = Form2.profesorCamposT
        leer(Form2.tabla, Form2.tablaElegida, 6)
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Form4.contador = 0
        Form2.tabla = "Alumno"
        Form2.tablaElegida = Form2.alumnoCamposT
        leer(Form2.tabla, Form2.tablaElegida, 8)
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Form4.contador = 0
        Form2.tabla = "Grupo"
        Form2.tablaElegida = Form2.grupoCamposT
        leer(Form2.tabla, Form2.tablaElegida, 3)
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Form4.contador = 0
        Form2.tabla = "Semestre"
        Form2.tablaElegida = Form2.semestreCamposT
        leer(Form2.tabla, Form2.tablaElegida, 1)
        Form4.Show()
        Me.Hide()
    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        FormCategorias.Show()
        Me.Hide()
    End Sub

    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class