Public Class Form9
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form9
    'Nombre del archivo: Form9
    '
    'Este form permite elegir la tabla que se quiere consultar
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Dim i As Integer                         'Variable para subrutinas FOR
    Dim campo(8) As Windows.Forms.Button      'Variable que absorbe las propiedades de los botones campoN

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
            Label1.Text = "Choose a field of the table:"
        Else
            Button1.Text = " Regresar"
            Label1.Text = "Elige un campo de la tabla"
        End If
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'ASIGNAR VALORES A CADA VARIABLE //////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Asignación_Valores(ByVal NombreCampo As String, ByVal Valor As Integer)
        With FormCategorias
            .ncampo(.botoncampo).Text = NombreCampo
            .campoElegido(.botoncampo) = Valor
        End With
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
        For Me.i = 1 To 7
            campo(i).Visible = False
        Next
        Select Case FormCategorias.tableElegida(FormCategorias.botoncampo)
            Case 1
                For Me.i = 0 To 5
                    campo(i).Visible = True
                    campo(i).Text = Form2.profesorCampos(i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Teacher"
                Else
                    Label7.Text = "Profesor"
                End If
                Exit Select
            Case 2
                For Me.i = 0 To 7
                    campo(i).Text = Form2.alumnoCampos(i)
                    campo(i).Visible = True
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Student"
                Else
                    Label7.Text = "Alumno"
                End If
                Exit Select
            Case 3
                For Me.i = 0 To 3
                    campo(i).Visible = True
                    campo(i).Text = Form2.materiaCampos(i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Signature"
                Else
                    Label7.Text = "Materia"
                End If
                Exit Select
            Case 4
                For Me.i = 0 To 2
                    campo(i).Visible = True
                    campo(i).Text = Form2.grupoCampos(i)
                Next
                If Form2.ingles = True Then
                    Label7.Text = "Group"
                Else
                    Label7.Text = "Grupo"
                End If
                Exit Select
            Case 5
                campo1.Text = Form2.semestreCampos(0)
                If Form2.ingles = True Then
                    Label7.Text = "Semester"
                Else
                    Label7.Text = "Semestre"
                End If
                
        End Select
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Volver_Form()
    End Sub

    Private Sub campo1_Click(sender As Object, e As EventArgs) Handles campo1.Click
        Asignación_Valores(campo1.Text, 1)
        Volver_Form()
    End Sub
    Private Sub campo2_Click(sender As Object, e As EventArgs) Handles campo2.Click
        Asignación_Valores(campo2.Text, 2)
        Volver_Form()
    End Sub

    Private Sub campo3_Click(sender As Object, e As EventArgs) Handles campo3.Click
        Asignación_Valores(campo3.Text, 3)
        Volver_Form()
    End Sub

    Private Sub campo4_Click(sender As Object, e As EventArgs) Handles campo4.Click
        Asignación_Valores(campo4.Text, 4)
        Volver_Form()
    End Sub

    Private Sub campo5_Click(sender As Object, e As EventArgs) Handles campo5.Click
        Asignación_Valores(campo5.Text, 5)
        Volver_Form()
    End Sub

    Private Sub campo6_Click(sender As Object, e As EventArgs) Handles campo6.Click
        Asignación_Valores(campo6.Text, 6)
        Volver_Form()
    End Sub

    Private Sub campo7_Click(sender As Object, e As EventArgs) Handles campo7.Click
        Asignación_Valores(campo7.Text, 7)
        Volver_Form()
    End Sub

    Private Sub campo8_Click(sender As Object, e As EventArgs) Handles campo8.Click
        Asignación_Valores(campo8.Text, 8)
        Volver_Form()
    End Sub


End Class