Imports MetroFramework

Public Class Form5
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: Form5
    'Nombre del archivo: Form5
    '
    'Este form es el menú para ver LA TABLA requerida en alguna consulta predefinida.
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Dim i As Integer                         'Variable para subrutinas FOR
    Dim campo(8) As Windows.Forms.Label      'Variable que absorbe las propiedades de los label campoN
    Public contenido(8) As Windows.Forms.TextBox  'Variable que absorbe las propieddaes de los text contenidoN
    Dim mensajeError, tituloError As String
    Public contador As Integer = 0
    Dim numCampos As Integer = 0
    Dim guardar As Boolean = False

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

    Private Sub Form5_VisibleChanged(sender As Object, e As System.EventArgs) Handles Me.VisibleChanged
        If Form2.tema = True Then
            Me.BackColor = Color.FromArgb(32, 32, 32)
        Else
            Me.BackColor = Color.FromArgb(32, 103, 179)
        End If
        If Form2.ingles = True Then
            Button1.Text = " Back"
            Button6.Text = "New"
            LimpiarL.Text = "Clear"
            mensajeError = "Records eliminated can't be recovered"
            tituloError = "Are you sure you want to delete this record?"
        Else
            Button1.Text = " Regresar"
            Button6.Text = "Nuevo"
            LimpiarL.Text = "Limpiar"
            mensajeError = "Los registros eliminados no podrán recuperarse"
            tituloError = "¿Seguro que deseas borrar este registro?"
        End If
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'FUNCIÓN DE ASIGNADOS INICIALES Y ASIGNADO DE FILAS AL FORM ///////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Form5_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
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

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Form1.Show()
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        guardar = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False

        Button7.Visible = True
        Button8.Visible = True
        LimpiarB.Visible = True
        LimpiarL.Visible = True
        For Me.i = 0 To 7
            contenido(i).Enabled = True
        Next
        For Me.i = 0 To 7
            contenido(i).Text = ""
        Next
        'Código para agregar registro ----------------

    End Sub

    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button9.Enabled = True
        Button10.Enabled = True

        Button7.Visible = False
        Button8.Visible = False
        LimpiarB.Visible = False
        LimpiarL.Visible = False
        For Me.i = 0 To 7
            contenido(i).Enabled = False
        Next

        For Me.i = 0 To numCampos
            contenido(i).Text = TablaLogica(contador, i)
        Next

        'For i = 0 To 7
        'contenido(i).Text = ""
        ' Next
        'Código para cancelar el registro nuevo O PARA cancelar la edición de un registro

    End Sub

    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
        Button6.Enabled = True
        Button9.Enabled = True
        Button10.Enabled = True

        Button7.Visible = False
        Button8.Visible = False
        LimpiarB.Visible = False
        LimpiarL.Visible = False
        For Me.i = 0 To 7
            contenido(i).Enabled = False
        Next
        'Código para guardar el registro actual
        If guardar = True Then
            editar(Form2.tabla, Form2.tablaElegida, numCampos + 1)
        Else
            Nuevo(Form2.tabla, Form2.tablaElegida, numCampos + 1)
        End If
        leer(Form2.tabla, Form2.tablaElegida, numCampos + 1)
        If contador > registros - 1 Then
            contador = registros - 1
        End If
        For Me.i = 0 To numCampos
            contenido(i).Text = TablaLogica(contador, i)
        Next
    End Sub

    Private Sub LimpiarB_Click(sender As Object, e As EventArgs) Handles LimpiarB.Click
        For Me.i = 0 To 7
            contenido(i).Text = ""
        Next
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        guardar = True
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        Button6.Enabled = False
        Button9.Enabled = False
        Button10.Enabled = False

        Button7.Visible = True
        Button8.Visible = True
        LimpiarB.Visible = True
        LimpiarL.Visible = True
        For Me.i = 0 To 7
            contenido(i).Enabled = True
        Next
        'Código para editar el registro actual --------



    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Dim resultado As MsgBoxResult
        resultado = MetroMessageBox.Show(Me, mensajeError, tituloError, MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If resultado = MsgBoxResult.No Then
            Exit Sub
        Else
            'Código para eliminar registro actual
            Eliminar(Form2.tabla, Form2.tablaElegida, numCampos + 1)
            leer(Form2.tabla, Form2.tablaElegida, numCampos + 1)
            If contador > registros - 1 Then
                contador = registros - 1
            End If
            If contador > 0 Then
                contador -= 1
            End If
            For Me.i = 0 To numCampos
                contenido(i).Text = TablaLogica(contador, i)
            Next
           
        End If
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

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        contador = 0
        For Me.i = 0 To numCampos
            contenido(i).Text = TablaLogica(contador, i)
        Next
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        contador = registros - 1
        For Me.i = 0 To numCampos
            contenido(i).Text = TablaLogica(contador, i)
        Next
    End Sub
End Class