Imports System.Data.OleDb

Module DatabaseModule




    Public dbPath As String =
        Application.StartupPath &
        "\users.accdb"

        Public connString As String =
        "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=" & dbPath

        Public conn As New OleDbConnection(connString)

        Public Sub OpenConnection()

            If conn.State <> ConnectionState.Open Then

                conn.Open()

            End If

        End Sub

        Public Sub CloseConnection()

            If conn.State <> ConnectionState.Closed Then

                conn.Close()

            End If

        End Sub

    End Module


   