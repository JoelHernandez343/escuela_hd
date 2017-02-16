Imports MetroFramework

Public Class FormCategorias
    '//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Nombre del form: FormCategorias
    'Nombre del archivo: Form6
    '
    'Este form permite elegir las tablas y campos que se quieren consultar
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Dim mas(8) As Windows.Forms.Button       'Absorbe las propieddaes de los botones masN
    Dim table(8) As Windows.Forms.Button     'Absorbe las propiedades de los botones tableN
    Public ntable(8) As Windows.Forms.Label  'Absorbe las propiedades de los labels ntableN
    Dim campo(8) As Windows.Forms.Button     'Absorbe las propiedades de los botones campoN
    Public ncampo(8) As Windows.Forms.Label  'Absorbe las propiedades de los labels ncampoN
    Dim i, altura, locy, locx As Integer     'For, locación en x y y
    Public conteo As Integer = 0               'Variable que cuenta el número de ingresos puestos
    Dim mensaje1, mensaje2, mensaje3 As String

    Public botontabla, botoncampo As Integer 'Variables que comunican el botón elegido
    Public tableElegida(8) As Integer        'Absorbe la tabla que será utilizada por criterio
    Public campoElegido(8) As Integer        'Absorbe el campo que será utilizado por criterio

    'TABLA DE VALORES DE TABLA ELEGIDA
    '0 ---> No elegido
    '1 ---> Profesor
    '2 ---> Alumno
    '3 ---> Materia
    '4 ---> Grupo
    '5 ---> Semestre

    'TABLA DE VALORES DE CAMPO ELEGIDO
    '0 ---> No elegido
    '1..8 ---> Depende de la tabla

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
            Label7.Text = "Custom search"
            Label1.Text = "Choose the parameters of your query, you can up to 8 parameters and minimum 1"
            Label3.Text = "Reset"
            label2.Text = "Run"
            mensaje1 = "You didn't choose any table "
            mensaje2 = "You have to choose at least one criteria"
            mensaje3 = "You didn't choose any field or table according to the choosen criteria"
            For i = 0 To 5
                ntable(i).Text = "Choose table"
                ncampo(i).Text = "Choose field"
            Next
        Else
            Button1.Text = " Regresar"
            Label7.Text = "Consulta personalizada"
            Label1.Text = "Elige los parámetros de tu consulta, se pueden hasta máximo 8 parámetros y mínimo 1"
            Label3.Text = "Restablece"
            label2.Text = "Ejecutar"
            mensaje1 = "No has elegido ninguna tabla"
            mensaje2 = "Tienes que seleccionar al menos un criterio"
            mensaje3 = "No has seleccionado algún campo o tabla de acuerdo a los criterios que elegiste"
            For i = 0 To 5
                ntable(i).Text = "Elegir tabla"
                ncampo(i).Text = "Elegir campo"
            Next
        End If
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'FUNCION PARA AUMENTA DE TAMAÑO ///////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Public Sub AumentarMas(ByVal n As Integer)

        altura = Me.Size.Height
        locy = Me.Location.Y
        locx = Me.Location.X
        conteo = conteo + 1
        Dim y As Integer
        If n <> 6 Then
            For Me.i = 0 To 90 Step 2
                If (i Mod 2 = 0) Then
                    y += 1
                    Me.Size = New System.Drawing.Size(733, altura + y)
                    Me.Location = New Point(locx, locy - y / 2)
                End If
            Next

        End If
        mas(n - 1).Enabled = False
        table(n - 1).Visible = True
        ntable(n - 1).Visible = True
        campo(n - 1).Visible = True
        ncampo(n - 1).Visible = True
        If n <> 6 Then
            mas(n).Visible = True
        End If
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'FUNCION PARA RESTABLECER DE TAMAÑO ///////////////////////////////////
    '//////////////////////////////////////////////////////////////////////

    Public Sub RestablecerT()
        Valores_Iniciales()
        conteo = 0
        altura = Me.Size.Height
        locy = Me.Location.Y
        locx = Me.Location.X
        Dim y As Integer
        For Me.i = 0 To 5
            If i <> 0 Then
                mas(i).Visible = False

            End If
            mas(i).Enabled = True
            table(i).Visible = False
            ntable(i).Visible = False
            ntable(i).Text = "Elegir tabla"
            campo(i).Visible = False
            ncampo(i).Visible = False
            ncampo(i).Text = "Elegir campo"
        Next
        For Me.i = 0 To 800 Step 2
            If Me.Size.Height <= 271 Then
                Exit Sub
            End If
            If (i Mod 2 = 0) Then
                y += 1
                Me.Size = New System.Drawing.Size(733, altura - y)
                Me.Location = New Point(locx, locy + y / 2)
            End If
        Next
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'FUNCIÓN DE ASIGNADOS INICIALES Y ASIGNADO DE FILAS AL FORM ///////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub FormCategorias_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        mas(0) = mas1
        mas(1) = mas2
        mas(2) = mas3
        mas(3) = mas4
        mas(4) = mas5
        mas(5) = mas6
        'mas(6) = mas7
        'mas(7) = mas8

        table(0) = table1
        table(1) = table2
        table(2) = table3
        table(3) = table4
        table(4) = table5
        table(5) = table6
        'table(6) = table7
        'table(7) = table8


        ntable(0) = ntable1
        ntable(1) = ntable2
        ntable(2) = ntable3
        ntable(3) = ntable4
        ntable(4) = ntable5
        ntable(5) = ntable6
        'ntable(6) = ntable7
        'ntable(7) = ntable8

        campo(0) = campo1
        campo(1) = campo2
        campo(2) = campo3
        campo(3) = campo4
        campo(4) = campo5
        campo(5) = campo6
        'campo(6) = campo7
        'campo(7) = campo8

        ncampo(0) = ncampo1
        ncampo(1) = ncampo2
        ncampo(2) = ncampo3
        ncampo(3) = ncampo4
        ncampo(4) = ncampo5
        ncampo(5) = ncampo6
        'ncampo(6) = ncampo7
        'ncampo(7) = ncampo8

        Valores_Iniciales()
        Me.Size = New System.Drawing.Size(733, 271)
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'VALORES INICALES DE ASIGNADO DE TABLAS Y CAMPOS //////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Valores_Iniciales()
        'ASIGNACIÓN POR DEFAULT DE VALORES DE LAS TABLAS ELEGIDAS EN 0 ---> No elegida
        For Me.i = 0 To 6
            tableElegida(i) = 0
        Next
        'ASIGNACIÓN POR DEFAULT DE VALORES DE LOS CAMPOS ELEGIDOS EN 0 ---> No elegido
        For Me.i = 0 To 6
            campoElegido(i) = 0
        Next
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'IDENTIFICACIÓN DE LA TABLA O CONTENIDO DEL QUE SE HABLA //////////////
    '//////////////////////////////////////////////////////////////////////

    Private Sub Eleccion_Tabla(ByVal tabla As Integer)
        botontabla = tabla
        Form8.Show()
        Me.Enabled = False
    End Sub
    Private Sub Eleccion_Campo(ByVal tabla As Integer)
        If tableElegida(tabla) = 0 Then
            MetroMessageBox.Show(Me, mensaje1, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            botoncampo = tabla
            Form9.Show()
            Me.Enabled = False
        End If
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'VERFICACIÓN //////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Verificacion()
        If conteo = 0 Then
            MetroMessageBox.Show(Me, mensaje2, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
        For Me.i = 0 To conteo - 1
            If tableElegida(i) = 0 Or campoElegido(i) = 0 Then
                MetroMessageBox.Show(Me, mensaje3, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Next
        FormConsultaCustom.Show()
        Me.Hide()

    End Sub


    '//////////////////////////////////////////////////////////////////////
    'CÓDIGO DEL PROGRAMA //////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Form3.Show()
        Me.Close()
    End Sub

    Private Sub mas1_Click(sender As System.Object, e As System.EventArgs) Handles mas1.Click
        AumentarMas(1)
    End Sub

    Private Sub mas2_Click(sender As System.Object, e As System.EventArgs) Handles mas2.Click
        AumentarMas(2)
    End Sub

    Private Sub mas3_Click(sender As System.Object, e As System.EventArgs) Handles mas3.Click
        AumentarMas(3)
    End Sub

    Private Sub mas4_Click(sender As System.Object, e As System.EventArgs) Handles mas4.Click
        AumentarMas(4)
    End Sub

    Private Sub mas5_Click(sender As System.Object, e As System.EventArgs) Handles mas5.Click
        AumentarMas(5)
    End Sub

    Private Sub mas6_Click(sender As System.Object, e As System.EventArgs) Handles mas6.Click
        AumentarMas(6)
    End Sub





    Private Sub restablecer_Click(sender As System.Object, e As System.EventArgs) Handles restablecer.Click
        RestablecerT()
    End Sub

    Private Sub table1_Click(sender As System.Object, e As System.EventArgs) Handles table1.Click
        Eleccion_Tabla(0)
    End Sub

    Private Sub table2_Click(sender As System.Object, e As System.EventArgs) Handles table2.Click
        Eleccion_Tabla(1)
    End Sub

    Private Sub table3_Click(sender As System.Object, e As System.EventArgs) Handles table3.Click
        Eleccion_Tabla(2)
    End Sub

    Private Sub table4_Click(sender As System.Object, e As System.EventArgs) Handles table4.Click
        Eleccion_Tabla(3)
    End Sub

    Private Sub table5_Click(sender As System.Object, e As System.EventArgs) Handles table5.Click
        Eleccion_Tabla(4)
    End Sub

    Private Sub table6_Click(sender As System.Object, e As System.EventArgs) Handles table6.Click
        Eleccion_Tabla(5)
    End Sub

    Private Sub campo1_Click(sender As System.Object, e As System.EventArgs) Handles campo1.Click
        Eleccion_Campo(0)
    End Sub

    Private Sub campo2_Click(sender As System.Object, e As System.EventArgs) Handles campo2.Click
        Eleccion_Campo(1)
    End Sub

    Private Sub campo3_Click(sender As System.Object, e As System.EventArgs) Handles campo3.Click
        Eleccion_Campo(2)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox(tableElegida(0) & " " & tableElegida(1) & " " & tableElegida(2) & " " & tableElegida(3) & " " & tableElegida(4) & " " & tableElegida(5) & " " & tableElegida(6) & " " & tableElegida(7))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox(campoElegido(0) & " " & campoElegido(1) & " " & campoElegido(2) & " " & campoElegido(3) & " " & campoElegido(4) & " " & campoElegido(5) & " " & campoElegido(6) & " " & campoElegido(7))
    End Sub

    Private Sub ejecutar_Click(sender As Object, e As EventArgs) Handles ejecutar.Click
        Verificacion()
    End Sub

    Private Sub campo4_Click(sender As System.Object, e As System.EventArgs) Handles campo4.Click
        Eleccion_Campo(3)
    End Sub

    Private Sub campo5_Click(sender As System.Object, e As System.EventArgs) Handles campo5.Click
        Eleccion_Campo(4)
    End Sub

    Private Sub campo6_Click(sender As System.Object, e As System.EventArgs) Handles campo6.Click
        Eleccion_Campo(5)
    End Sub

End Class