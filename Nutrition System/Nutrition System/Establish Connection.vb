Imports System.Data.OleDb

Module Establish_Connection

    Dim filepath As String

    'this defines conna as a public database connection variable
    Public conn As New OleDbConnection

    'this subroutine will define the connection string to the database
    Public Sub connect()

        'this try will run the connection code with the potential to fail if the database file is missing or corrupt
        Try

            'this defines the relative location of the database object
            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|datadirectory|\Nutrition_Database.accdb;"

            'this checks if the connection state is open or closed
            If conn.State = ConnectionState.Closed Then
                conn.Open()
            End If

            'this will catch any errors found in the above code and proceed with the alternative code
        Catch ex As Exception

            'this message box will alert the user that the error has occured
            MsgBox("Database Not Located, User Location required.")


        End Try
    End Sub

End Module
