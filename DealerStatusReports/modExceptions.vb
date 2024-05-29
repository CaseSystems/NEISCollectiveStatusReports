Imports System.Data.SqlClient
Module modExceptions
#Region "Connection Strings"

    Public cnError As New SqlClient.SqlConnection("Data Source='ERP-M1';Initial Catalog='M1_M1'; User Id='sa'; Password='n4819p+';")

#End Region
#Region "Exception Handling Functions"

    Public Sub ThrowError(ByVal ex As String)

        SaveError(ex)

    End Sub

    Function SaveError(ByVal strException As String) As Boolean
        Dim Complete As Boolean = False
        Try
            'Open the Connection, if Closed
            '------------------------------
            If Not cnError Is Nothing Then
                If cnError.State = ConnectionState.Closed Then

                    cnError.Open()

                End If
            End If

            'Get Next Error ID
            '-----------------
            Dim strNextErrorID As String = "SELECT TOP(1) uslErrorID FROM uShopFloorErrorLog ORDER BY uslErrorID DESC"
            Dim dtNextID As DataTable = csiGetDataTable(strNextErrorID, cnError)
            Dim iErrorID As Integer = 0

            'Replace Single and Double Quotes with Spaces
            '--------------------------------------------
            strException = strException.Replace("'", " ")


            'If Error have been Logged, increment by one
            '-------------------------------------------
            If dtNextID.Rows.Count > 0 Then
                iErrorID = CInt(Trim(dtNextID.Rows(0).Item("uslErrorID").ToString())) + 1
            Else
                iErrorID = 1
            End If

            Dim strSaveError As String = "INSERT INTO uShopFloorErrorLog (" &
                                         "uslErrorID, uslUserID, uslComputerName, uslErrorMessage, uslCreatedDate, uslAddressed, uslProgram) VALUES("

            'Error ID
            '--------
            strSaveError += iErrorID.ToString() & ", "

            'User ID
            '--------
            strSaveError += "'" & Trim(Environment.UserName) & "', "

            'Computer Name
            '-------------
            strSaveError += "'" & Trim(Environment.MachineName) & "', "

            'Error Message
            '-------------
            strSaveError += "'" & Trim(strException) & "', "

            'CreatedDate
            '-----------
            strSaveError += "GETDATE()" & ","

            'Addressed
            '-----------
            strSaveError += "'0', "

            'Program
            '-----------
            strSaveError += "'DealerStatusEmailer')"

            'Create SQL Command
            '------------------
            Dim cmdSaveError As New SqlClient.SqlCommand
            cmdSaveError.CommandText = strSaveError

            'Verify Connection
            '-----------------
            If Not cnError Is Nothing Then
                If cnError.State = ConnectionState.Closed Then

                    cnError.Open()

                End If
            End If

            'Set Connection
            '--------------
            cmdSaveError.Connection = cnError

            'Execute Connection
            '------------------
            If cmdSaveError.ExecuteNonQuery() Then
                Complete = True
            End If

            'Return Memory to System
            '-----------------------
            dtNextID.Dispose()

        Catch ex As Exception
            ThrowError(ex.ToString())
        Finally

            'Close the Connection, if Open
            '-----------------------------
            If Not cnError Is Nothing Then
                If cnError.State = ConnectionState.Open Then

                    cnError.Close()

                End If
            End If
        End Try
        Return Complete
    End Function

#End Region
End Module
