Imports MySql.Data.MySqlClient

Module Module1

    Public tax_rate As Double = 0.25
    Public myadocon, conn As New MySqlConnection
    Public cmd As New MySqlCommand
    Public cmdread As MySqlDataReader
    Public db_server = "'localhost'" ' server name
    Public db_port = "'3306'" 'port
    Public db_uid = "'root'" ' user id
    Public db_pwd = "''" ' password
    Public db_name = "'ejdrepairshop_database'" ' database name
    Public strconn As String = "server=" & db_server & "; port=" & db_port & "; userid=" & db_uid & ";password=" & db_pwd & "; database=" & db_name & ";"

    Public Sub readquery(ByVal sql As String)
        Try
            With conn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = strconn
                .Open()
            End With
            With cmd
                .Connection = conn
                .CommandText = sql
                cmdread = .ExecuteReader
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function TestDatabaseConnection() As Boolean
        Try
            conn.ConnectionString = strconn
            conn.Open()
            If conn.State = ConnectionState.Open Then
                conn.Close()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function

End Module
