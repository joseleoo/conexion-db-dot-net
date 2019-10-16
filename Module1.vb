Imports System.Data.SqlClient 'para conectar sql server
Imports System.Data.OleDb 'para conectar con access
Imports System.Data
Module Module1

    Private sqlconexion As SqlConnection 'para sql server
    Private sqlcomando As SqlCommand 'para sqlserver
    Private sqladapter As SqlDataAdapter
    Private oledbconexion As OleDbConnection 'para acces
    Private oledbcomando As OleDbCommand 'para acces
    Private oledbadapter As OleDbDataAdapter 'para acces
    Private objtabla As DataTable

    Private Sub iniciaSQL()
        Try
            sqlconexion = New SqlConnection
            sqlconexion.ConnectionString = "aqui va la cadena de conexion de sql Server"
            sqlconexion.Open()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub finalSQL()
        Try
            If sqlconexion.State = ConnectionState.Open Then
                sqlconexion.Close() 'cierra la base de datos
            End If
            sqlconexion.Dispose() 'elimina el objeto
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub iniciaOledb()
        Try
            oledbconexion = New OleDbConnection
            oledbconexion.ConnectionString = "aqui va la cadena de conexion de acces"
            oledbconexion.Open()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub finalOledb()
        Try
            If oledbconexion.State = ConnectionState.Open Then
                oledbconexion.Close()
            End If
            oledbconexion.Dispose()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Function operaSQL(ByVal str As String) As Boolean
        iniciaSQL()
        Try
            sqlcomando = New SqlCommand 'creo una instancia
            sqlcomando.CommandText = str 'le asigno el query
            sqlcomando.Connection = sqlconexion 'le asigno la conexion de la BD
            sqlcomando.ExecuteNonQuery() 'ejecuta un query
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
            finalSQL()
            Exit Function
        End Try
        finalSQL()
        Return True
        sqlcomando.Dispose()

    End Function

    Public Function operaOledb(ByVal str As String) As Boolean
        iniciaOledb()
        Try
            oledbcomando = New OleDbCommand
            oledbcomando.CommandText = str
            oledbcomando.Connection = oledbconexion
            oledbcomando.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
            finalOledb()
            Exit Function
        End Try
        finalOledb()
        Return True
        oledbcomando.Dispose()
    End Function

    Public Function GridSQL(ByVal str As String) As DataTable
        iniciaSQL() 'inicio la BD
        Try
            'cre la instancia asigno el query y le asigno el objeto de conexion
            sqladapter = New SqlDataAdapter(str, sqlconexion)
            'instancio el objeto de tipo datatable
            objtabla = New DataTable
            'la propiedad fill llena datos en la tabla
            sqladapter.Fill(objtabla)
        Catch ex As Exception
            MsgBox(ex.ToString)
            finalSQL()
        End Try
        Return objtabla
        finalSQL()
        sqladapter.Dispose()
    End Function
    Public Function GridOledb(ByVal str As String) As DataTable
        iniciaOledb() 'inicio la BD
        Try
            'cre la instancia asigno el query y le asigno el objeto de conexion
            oledbadapter = New OleDbDataAdapter(str, oledbconexion)
            'instancio el objeto de tipo datatable
            objtabla = New DataTable
            'la propiedad fill llena datos en la tabla
            oledbadapter.Fill(objtabla)
        Catch ex As Exception
            MsgBox(ex.ToString)
            finalOledb()
        End Try
        Return objtabla
        finalOledb()
        oledbadapter.Dispose()
    End Function


End Module
