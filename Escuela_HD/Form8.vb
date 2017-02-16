Public Class Form8
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form8
    'Nombre del archivo: Form8
    '
    'Este form permite elegir la tabla que se quiere consultar
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
    Private Sub Me_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
        If Form2.tema = True Then
            Me.BackColor = Color.FromArgb(64, 64, 64)
        Else
            Me.BackColor = Color.FromArgb(51, 131, 219)
        End If
        If Form2.ingles = True Then
            Button1.Text = " Back"
            Label1.Text = "Choose a table"
            Button2.Text = "Teacher"
            Button3.Text = "Student"
            Button4.Text = "Signature"
            Button5.Text = "Group"
            Button6.Text = "Semester"
        Else
            Button1.Text = " Regresar"
            Label1.Text = "Elige una tabla"
            Button2.Text = "Profesor"
            Button3.Text = "Alumno"
            Button4.Text = "Materia"
            Button5.Text = "Grupo"
            Button6.Text = "Semestre"
        End If
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'VOLVER AL FORM ANTERIOR //////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Volver_Form()
        Me.Close()
        FormCategorias.Enabled = True
        FormCategorias.Focus()
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'ASIGNAR VALORES A CADA VARIABLE //////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Asignación_Valores(ByVal NombreTabla As String, ByVal Valor As Integer)
        With FormCategorias
            .ntable(.botontabla).Text = NombreTabla
            .tableElegida(.botontabla) = Valor
            .campoElegido(.botontabla) = 0
            If Form2.ingles = True Then
                .ncampo(.botontabla).Text = "Choose field"
            Else
                .ncampo(.botontabla).Text = "Elegir campo"
            End If
        End With
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Volver_Form()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If Form2.ingles = True Then
            Asignación_Valores("Teacher", 1)
        Else
            Asignación_Valores("Profesor", 1)
        End If
        Volver_Form()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Form2.ingles = True Then
            Asignación_Valores("Student", 2)
        Else
            Asignación_Valores("Alumno", 2)
        End If
        Volver_Form()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Form2.ingles = True Then
            Asignación_Valores("Signature", 3)
        Else
            Asignación_Valores("Materia", 3)
        End If

        Volver_Form()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Form2.ingles = True Then
            Asignación_Valores("Group", 4)
        Else
            Asignación_Valores("Grupo", 4)
        End If

        Volver_Form()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Form2.ingles = True Then
            Asignación_Valores("Semester", 5)
        Else
            Asignación_Valores("Semestre", 5)
        End If
        Volver_Form()
    End Sub

    Private Sub Form8_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class