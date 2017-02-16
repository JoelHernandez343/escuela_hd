Imports MetroFramework

Module Module1
    ''//////////////////////////////////////////////////////////////////////
    'DESCRPCIÓN INICIAL ///////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    'Modulo 1
    '
    'Aquí se guardan las funciones de conexión y consulta
    '
    '//////////////////////////////////////////////////////////////////////
    'DECLARACIÓN DE VARIABLES /////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public conexion As New OleDb.OleDbConnection
    Public comando As New OleDb.OleDbCommand
    Public lector As OleDb.OleDbDataReader
    Public adaptador As New OleDb.OleDbDataAdapter
    Public consulta As String = ""
    Dim i As Integer = 0
    Dim j As Integer = 0
    Dim cadenaP As String = ""
    Public TablaLogica(100, 100) As String
    Public registros As Integer = 0
    Dim numcampos As Integer = 0
    Dim editarCampo As String
    Public camposALeer As String = ""

    '//////////////////////////////////////////////////////////////////////
    'CONEXIÓN CON BASE DE DATOS - INTENTO /////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub ConnectToAccess()

        conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data source=" & CurDir() & "\Escuela_HD.accdb"
        Try
            conexion.Open()
        Catch ex As Exception
            If ex.ToString.Contains("El proveedor 'Microsoft.ACE.OLEDB.12.") Then
                MetroMessageBox.Show(Form2, "RECURSO NO ENCONTRADO" + vbCr + "Error AccesDatabaseEngineNotFound", "Debes instalar AccessDatabaseEngine", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                AbrirRecurso(ex)
            Else
                MetroMessageBox.Show(Form2, ex.Message, "Error inesperado", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                conexion.Close()
                End
            End If
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'ABRIR EL RECURSO EXTERNO PARA EL PROBLEMA CONOCIDO ///////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub AbrirRecurso(ex1 As Exception)
        Try
            System.Diagnostics.Process.Start(CurDir() & "\AccessDatabaseEngine.exe")
            conexion.Close()
            End
        Catch exs As Exception
            If exs.ToString.Contains("El usuario ha cancelado la op") Then
                MetroMessageBox.Show(Form2, "Intentalo de nuevo" & vbCr & "No podras editar ni consultar", "¡Debes instalarlo!", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                conexion.Close()
                End
            Else
                MsgBox(ex1.Message)
            End If
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'CREA STRING CONSULTA, REALIZA CONSULTA ///////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub leer(ByRef nombreTabla As String, ByRef tabla() As String, ByRef cantidad As Integer)

        camposALeer = ""
        For i = 1 To cantidad
            camposALeer = camposALeer & tabla(i)
            If i <> cantidad Then
                camposALeer = camposALeer & ", "
            End If
        Next
        consulta = "SELECT " & camposALeer & " FROM " & nombreTabla
        ConsultaTotal(lector, comando, consulta, cantidad)
        
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'REALIZA CONSULTA Y ASIGNA A ARREGLO BIDIMENSIONAL ////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub ConsultaTotal(ByRef lectorP As OleDb.OleDbDataReader, comandoP As OleDb.OleDbCommand, ByRef texto As String, ByRef limite As Integer)
        'Inicilización de la conexión
        comandoP.Connection = conexion
        comandoP.CommandType = CommandType.Text
        comandoP.CommandText = texto
        'Intenta la ejecución
        registros = 0
        Try
            lectorP = comandoP.ExecuteReader()
            If lectorP.HasRows Then
                While lectorP.Read()
                    For i = 0 To limite - 1
                        TablaLogica(j, i) = lector(i).ToString
                        cadenaP = cadenaP & TablaLogica(j,
                                                        i) & " "
                    Next
                    cadenaP = cadenaP & vbCr
                    j += 1
                    registros += 1
                End While
                j = 0
                'MsgBox(cadenaP)
            Else
                MsgBox("NO HAY DATOS")
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        cadenaP = ""
        lectorP.Close()
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'EDITAR CAMPOS ///////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub editar(ByRef tabla As String, ByRef campos() As String, ByRef contador As Integer)

        editarCampo = ""
        editarCampo = "UPDATE " & tabla & " SET "
        For i = 1 To contador
            editarCampo = editarCampo & campos(i) & " = '" & Form5.contenido(i - 1).Text & "'"
            If i <> contador Then
                editarCampo = editarCampo & " , "
            End If
        Next
        editarCampo = editarCampo & " WHERE " & campos(1) & " =" & Form5.contenido(0).Text
        MsgBox(editarCampo)
        comando.Connection = conexion
        comando.CommandType = CommandType.Text
        consulta = editarCampo
        comando.CommandText = consulta
        Try
            comando.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'conexion.Close()
    End Sub

    '//////////////////////////////////////////////////////////////////////
    'NUEVO ////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Nuevo(ByRef tabla As String, ByRef campos() As String, ByRef contador As Integer)

        editarCampo = ""
        editarCampo = "INSERT INTO " & tabla & "("
        For i = 1 To contador
            editarCampo = editarCampo & campos(i)
            If i <> contador Then
                editarCampo = editarCampo & " , "
            End If
        Next
        editarCampo = editarCampo & ") VALUES("
        For i = 1 To contador
            editarCampo = editarCampo & " '" & Form5.contenido(i - 1).Text & "'"
            If i <> contador Then
                editarCampo = editarCampo & " , "
            End If
        Next
        editarCampo = editarCampo & ")"
        comando.Connection = conexion
        comando.CommandType = CommandType.Text
        consulta = editarCampo
        comando.CommandText = consulta
        Try
            comando.ExecuteNonQuery()

        Catch ex As Exception
            If ex.ToString.Contains("valores duplicados") Then
                MsgBox("Ya esxiste esa ID")
            Else
                MsgBox(ex.Message)
            End If


        End Try
        'conexion.Close()
    End Sub


    '//////////////////////////////////////////////////////////////////////
    'ELIMINAR ////////////////////////////////////////////////////////////////
    '//////////////////////////////////////////////////////////////////////
    Public Sub Eliminar(ByRef tabla As String, ByRef campos() As String, ByRef contador As Integer)
        editarCampo = ""
        editarCampo = "DELETE FROM " & tabla & " WHERE " & campos(1) & " =" & Form5.contenido(0).Text
        comando.Connection = conexion
        comando.CommandType = CommandType.Text
        consulta = editarCampo
        comando.CommandText = consulta
        Try
            comando.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        'conexion.Close()
    End Sub
End Module
