Imports MetroFramework

Public Class Form2
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form2
    'Nombre del archivo: Form2
    '
    'Este form es el FORMULARIO INICIAL, y sirve como menú inicial donde se pueden abrir los siguientes forms:
    'Form1, Form3
    'Desde aquí se pueden acceder a todas las funciones de todo el programa
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public abierto As Boolean = False        'Indica si está abierto el menú
    Public tema As Boolean = False           'Indica el tema global del programa
    Public ingles As Boolean = False         'Indica el idioma del programa
    Public tabla As String                   'Indica la tabla que será usada tanto en el Form1 como en el Form
    Public tablaElegida() As String

    Public profesorCampos(6) As String
    Public alumnoCampos(8) As String
    Public materiaCampos(4) As String
    Public grupoCampos(3) As String
    Public semestreCampos(1) As String

    Public profesorCamposES() As String = {"NSS", "Nombres", "Apellido Paterno", "Apellido Materno", "Dirección", "Teléfono"}
    Public alumnoCamposES() As String = {"Matrícula", "Nombres", "Apellido Paterno", "Apellido Materno", "Edad", "Sexo", "Grupo", "Tutor"}
    Public materiaCamposES() As String = {"Id", "Nombre", "Número de semestre", "Profesor"}
    Public grupoCamposES() As String = {"Grupo", "Número de semestre", "Asesor"}
    Public semestreCamposES() As String = {"Número de semestre"}

    Public profesorCamposEN() As String = {"NSS", "Names", "First Surname", "Last Surname", "Address", "Phone"}
    Public alumnoCamposEN() As String = {"Registration", "Names", "First Surname", "Last Surname", "Age", "Gender", "Group", "Parent"}
    Public materiaCamposEN() As String = {"Signature ID", "Name", "Area", "Semester number", "Teacher"}
    Public grupoCamposEN() As String = {"Group", "Semester number", "Adviser"}
    Public semestreCamposEN() As String = {"Semester number"}

    Public profesorCamposT() As String = {"Void", "NSS", "Nombres", "ApellidoPaterno", "ApellidoMaterno", "Direccion", "Telefono"}
    Public alumnoCamposT() As String = {"Void", "Matricula", "Nombres", "ApellidoPaterno", "ApellidoMaterno", "Edad", "Sexo", "Grupo", "Tutor"}
    Public materiaCamposT() As String = {"Void", "Id", "Nombre", "NumeroSemestre", "Profesor"}
    Public grupoCamposT() As String = {"Void", "Grupo", "NumeroSemestre", "Asesor"}
    Public semestreCamposT() As String = {"Void", "NumeroSemestre"}
    
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

    'Éstas subrutinas permiten desplazar el formulario desde el LABEL
    Private Sub Label7_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseDown
        ex = e.X
        ey = e.Y
        Arrastre = True
    End Sub
    Private Sub Label7_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseUp
        Arrastre = False
    End Sub
    Private Sub Label7_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseMove
        'Si el formulario no tiene borde (FormBorderStyle = none) la siguiente linea funciona bien
        If Arrastre Then Me.Location = Me.PointToScreen(New Point(MousePosition.X - Me.Location.X - ex - 36, MousePosition.Y - Me.Location.Y - ey - 20))
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'FUNCIÓN DE MENSAJE DE SALIDA /////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Private Sub Form2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim resultado As MsgBoxResult
        'Muestra el mensaje en Iglés o Español
        If ingles = False Then
            resultado = MetroMessageBox.Show(Me, "El sistema procederá a cerrar el programa", "¿Seguro que deseas salir?", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        Else
            resultado = MetroMessageBox.Show(Me, "The system will close down", "Sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        End If
        'Revisa si el usuario acepto o no salir
        If resultado = MsgBoxResult.Yes Then
            End
        Else
            e.Cancel = True
        End If
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'CONTROL DE ANIMACIONES DE BOTONES PRINCIPALES ////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub FondoNormal(ByVal boton As Windows.Forms.Button, ByVal texto1 As Windows.Forms.Label, ByVal texto2 As Windows.Forms.Label, ByVal boton2 As Windows.Forms.Button)
        If tema = False Then
            boton.BackColor = Color.FromArgb(78, 168, 221)
            texto1.BackColor = Color.FromArgb(32, 103, 179)
            texto2.BackColor = Color.FromArgb(32, 103, 179)
            boton2.BackColor = Color.FromArgb(32, 103, 179)
        Else
            boton.BackColor = Color.FromArgb(78, 168, 221)
            texto1.BackColor = Color.FromArgb(32, 32, 32)
            texto2.BackColor = Color.FromArgb(32, 32, 32)
            boton2.BackColor = Color.FromArgb(32, 32, 32)
        End If
    End Sub

    Public Sub FondoHover(ByVal boton As Windows.Forms.Button, ByVal texto1 As Windows.Forms.Label, ByVal texto2 As Windows.Forms.Label, ByVal boton2 As Windows.Forms.Button)
        If tema = False Then
            boton.BackColor = Color.FromArgb(113, 185, 227)
            texto1.BackColor = Color.FromArgb(38, 123, 214)
            texto2.BackColor = Color.FromArgb(38, 123, 214)
            boton2.BackColor = Color.FromArgb(38, 123, 214)
        Else
            boton.BackColor = Color.FromArgb(113, 185, 227)
            texto1.BackColor = Color.FromArgb(50, 50, 50)
            texto2.BackColor = Color.FromArgb(50, 50, 50)
            boton2.BackColor = Color.FromArgb(50, 50, 50)
        End If
    End Sub

    Public Sub FondoDown(ByVal boton As Windows.Forms.Button, ByVal texto1 As Windows.Forms.Label, ByVal texto2 As Windows.Forms.Label, ByVal boton2 As Windows.Forms.Button)
        If tema = False Then
            boton.BackColor = Color.FromArgb(58, 158, 218)
            texto1.BackColor = Color.FromArgb(116, 161, 209)
            texto2.BackColor = Color.FromArgb(116, 161, 209)
            boton2.BackColor = Color.FromArgb(116, 161, 209)
        Else
            boton.BackColor = Color.FromArgb(58, 158, 218)
            texto1.BackColor = Color.FromArgb(16, 16, 16)
            texto2.BackColor = Color.FromArgb(16, 16, 16)
            boton2.BackColor = Color.FromArgb(16, 16, 16)
        End If
    End Sub

    Public Sub BotCambioColor(ByVal boton As Windows.Forms.Button, ByVal texto1 As Windows.Forms.Label, ByVal texto2 As Windows.Forms.Label)
        If tema = False Then
            texto1.BackColor = Color.FromArgb(32, 103, 179)
            texto2.BackColor = Color.FromArgb(32, 103, 179)
            boton.BackColor = Color.FromArgb(32, 103, 179)
            boton.FlatAppearance.MouseDownBackColor = Color.FromArgb(116, 161, 209)
            boton.FlatAppearance.MouseOverBackColor = Color.FromArgb(38, 123, 214)
        Else
            texto1.BackColor = Color.FromArgb(32, 32, 32)
            texto2.BackColor = Color.FromArgb(32, 32, 32)
            boton.BackColor = Color.FromArgb(32, 32, 32)
            boton.FlatAppearance.MouseDownBackColor = Color.FromArgb(16, 16, 16)
            boton.FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 50, 50)
        End If
    End Sub

    'Botones de EDITAR -----------------------
    Private Sub Button1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseDown
        FondoDown(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Button1_MouseEnter(sender As Object, e As System.EventArgs) Handles Button1.MouseEnter
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Button1_MouseLeave(sender As Object, e As System.EventArgs) Handles Button1.MouseLeave
        FondoNormal(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Button1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button1.MouseUp
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub

    Private Sub Button4_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button4.MouseDown
        FondoDown(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Button4_MouseEnter(sender As Object, e As System.EventArgs) Handles Button4.MouseEnter
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Button4_MouseLeave(sender As Object, e As System.EventArgs) Handles Button4.MouseLeave
        FondoNormal(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Button4_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button4.MouseUp
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub

    Private Sub Label1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseDown
        FondoDown(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Label1_MouseEnter(sender As Object, e As System.EventArgs) Handles Label1.MouseEnter
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Label1_MouseLeave(sender As Object, e As System.EventArgs) Handles Label4.MouseLeave
        FondoNormal(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Label1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseUp
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub

    Private Sub Label2_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseDown
        FondoDown(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Label2_MouseEnter(sender As Object, e As System.EventArgs) Handles Label2.MouseEnter
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Label2_MouseLeave(sender As Object, e As System.EventArgs) Handles Label2.MouseLeave
        FondoNormal(Button1, Label1, Label2, Button4)
    End Sub
    Private Sub Label2_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseUp
        FondoHover(Button1, Label1, Label2, Button4)
    End Sub

    'Botones de CONSULTAR -----------------------
    Private Sub Button5_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button5.MouseDown
        FondoDown(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Button5_MouseEnter(sender As Object, e As System.EventArgs) Handles Button5.MouseEnter
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Button5_MouseLeave(sender As Object, e As System.EventArgs) Handles Button5.MouseLeave
        FondoNormal(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Button5_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button5.MouseUp
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub

    Private Sub Button3_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseDown
        FondoDown(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Button3_MouseEnter(sender As Object, e As System.EventArgs) Handles Button3.MouseEnter
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Button3_MouseLeave(sender As Object, e As System.EventArgs) Handles Button3.MouseLeave
        FondoNormal(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Button3_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button3.MouseUp
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub

    Private Sub Label4_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label4.MouseDown
        FondoDown(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Label4_MouseEnter(sender As Object, e As System.EventArgs) Handles Label4.MouseEnter
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Label4_MouseLeave(sender As Object, e As System.EventArgs) Handles Label4.MouseLeave
        FondoNormal(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Label4_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label4.MouseUp
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub

    Private Sub Label3_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label3.MouseDown
        FondoDown(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Label3_MouseEnter(sender As Object, e As System.EventArgs) Handles Label3.MouseEnter
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Label3_MouseLeave(sender As Object, e As System.EventArgs) Handles Label3.MouseLeave
        FondoNormal(Button3, Label3, Label4, Button5)
    End Sub
    Private Sub Label3_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label3.MouseUp
        FondoHover(Button3, Label3, Label4, Button5)
    End Sub

    'Botones de SALIR -----------------------
    Private Sub Button6_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button6.MouseDown
        FondoDown(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Button6_MouseEnter(sender As Object, e As System.EventArgs) Handles Button6.MouseEnter
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Button6_MouseLeave(sender As Object, e As System.EventArgs) Handles Button6.MouseLeave
        FondoNormal(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Button6_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button6.MouseUp
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub

    Private Sub Button2_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseDown
        FondoDown(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Button2_MouseEnter(sender As Object, e As System.EventArgs) Handles Button2.MouseEnter
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Button2_MouseLeave(sender As Object, e As System.EventArgs) Handles Button2.MouseLeave
        FondoNormal(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Button2_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Button2.MouseUp
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub

    Private Sub Label5_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label5.MouseDown
        FondoDown(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Label5_MouseEnter(sender As Object, e As System.EventArgs) Handles Label5.MouseEnter
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Label5_MouseLeave(sender As Object, e As System.EventArgs) Handles Label5.MouseLeave
        FondoNormal(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Label5_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label5.MouseUp
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub

    Private Sub Label6_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label6.MouseDown
        FondoDown(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Label56_MouseEnter(sender As Object, e As System.EventArgs) Handles Label6.MouseEnter
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Label6_MouseLeave(sender As Object, e As System.EventArgs) Handles Label6.MouseLeave
        FondoNormal(Button2, Label5, Label6, Button6)
    End Sub
    Private Sub Label6_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Label6.MouseUp
        FondoHover(Button2, Label5, Label6, Button6)
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'CERRAR PROGRAMA //////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Cerrar()
        conexion.Close()
        Me.Close()
        End
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Size = New System.Drawing.Size(733, 445)
        For i = 0 To 5
            profesorCampos(i) = profesorCamposES(i)
        Next
        For i = 0 To 7
            alumnoCampos(i) = alumnoCamposES(i)
        Next
        For i = 0 To 3
            materiaCampos(i) = materiaCamposES(i)
        Next
        For i = 0 To 2
            grupoCampos(i) = grupoCamposES(i)
        Next
        semestreCampos(0) = semestreCamposES(0)
        ConnectToAccess()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Cerrar()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Cerrar()
    End Sub
    Private Sub Label6_Click(sender As System.Object, e As System.EventArgs) Handles Label6.Click
        Cerrar()
    End Sub

    Private Sub Label5_Click(sender As System.Object, e As System.EventArgs) Handles Label5.Click
        Cerrar()
    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        Dim x, i As Integer
        If abierto = False Then
            For i = 0 To 1000 Step 3
                If (i Mod 2 = 0) Then
                    x += 1
                    Me.Size = New System.Drawing.Size(733 + x, 445)
                End If
            Next
            abierto = True
            Panel1.Visible = True
        Else

            For i = 0 To 1000 Step 3
                If (i Mod 2 = 0) Then
                    x += 1
                    Me.Size = New System.Drawing.Size(900 - x, 445)
                End If
            Next
            abierto = False
            Panel1.Visible = False
        End If
    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        If tema = False Then
            Button7.Text = "Blue"
            tema = True
            Me.BackColor = Color.FromArgb(32, 32, 32)
        Else
            tema = False
            Button7.Text = "Dark"
            Me.BackColor = Color.FromArgb(32, 103, 179)

        End If
        BotCambioColor(Button4, Label1, Label2)
        BotCambioColor(Button5, Label3, Label4)
        BotCambioColor(Button6, Label5, Label6)
    End Sub



    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Label2_Click(sender As System.Object, e As System.EventArgs) Handles Label2.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Label4_Click(sender As System.Object, e As System.EventArgs) Handles Label4.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Label3_Click(sender As System.Object, e As System.EventArgs) Handles Label3.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        If ingles = False Then
            ingles = True
            Button9.Text = "Español"
            Label1.Text = "Here you can modify the data base information"
            Label2.Text = "Edit"
            Label3.Text = "Here you can consult the school logs"
            Label4.Text = "Consult"
            Label5.Text = "Exit of the system"
            Label6.Text = "Exit"
            Label7.Text = "Menu"
            Label9.Text = "Theme"
            Label10.Text = "Language"
            For i = 0 To 5
                profesorCampos(i) = profesorCamposEN(i)
            Next
            For i = 0 To 7
                alumnoCampos(i) = alumnoCamposEN(i)
            Next
            For i = 0 To 3
                materiaCampos(i) = materiaCamposEN(i)
            Next
            For i = 0 To 2
                grupoCampos(i) = grupoCamposEN(i)
            Next
            semestreCampos(0) = semestreCamposES(0)
        Else
            ingles = False
            Button9.Text = "English"
            Label1.Text = "Aquí puedes modificar los datos de la base de datos"
            Label2.Text = "Editar"
            Label3.Text = "Aquí puedes visualizar los registros de la escuela"
            Label4.Text = "Consultar"
            Label5.Text = "Salir del sistema"
            Label6.Text = "Salir"
            Label7.Text = "Menú"
            Label9.Text = "Tema"
            Label10.Text = "Lenguaje"
            For i = 0 To 5
                profesorCampos(i) = profesorCamposES(i)
            Next
            For i = 0 To 7
                alumnoCampos(i) = alumnoCamposES(i)
            Next
            For i = 0 To 3
                materiaCampos(i) = materiaCamposES(i)
            Next
            For i = 0 To 2
                grupoCampos(i) = grupoCamposES(i)
            Next
            semestreCampos(0) = semestreCamposES(0)
        End If
    End Sub
End Class