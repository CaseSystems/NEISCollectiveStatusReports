Imports System.IO

Module basData
#Region "View/Query Functions"
    Function csiGetDataAdapterForQuery(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sPrimaryKeyField As String, Optional ByVal sWhere As String = "", Optional ByVal sInnerJoin As String = "", Optional ByVal fUseIdentity As Boolean = True, Optional ByVal sIgnore As String = "") As SqlClient.SqlDataAdapter
        Dim da As New SqlClient.SqlDataAdapter
        Try
            'Difference is the sIgnore parameter.  This is an array list of those fields to not update/edit/insert
            'Get Primary Key
            da.SelectCommand = csiGetSelectCommand(sTable, cnSql, sInnerJoin, sWhere)
            da.InsertCommand = csiGetInsertCommandForQuery(sTable, cnSql, sPrimaryKeyField, fUseIdentity, sIgnore)
            da.UpdateCommand = csiGetUpdateCommandForQuery(sTable, cnSql, sPrimaryKeyField, fUseIdentity, sIgnore)
            da.DeleteCommand = csiGetDeleteCommand(sTable, cnSql, sPrimaryKeyField)
            Return da
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return da
    End Function
    Function csiGetInsertCommandForQuery(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sIdentity As String, ByVal fUseIdentity As Boolean, ByVal sIgnore As String) As SqlClient.SqlCommand
        Try
            Dim myCmd As New SqlClient.SqlCommand
            Dim sSelect As String = ""
            Dim sValues As String = ""
            Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM " & sTable & " WHERE 1=0", cnSql)
            Dim dt As New DataTable
            Dim dc As DataColumn
            Dim sFieldList As String = ""
            Dim fIdentityParamUsed As Boolean = False
            Dim sIgnores() As String = sIgnore.ToLower.Split("|")
            da.Fill(dt)
            For Each dc In dt.Columns
                If (dc.ColumnName.ToUpper <> sIdentity.ToUpper Or Not fUseIdentity) Then
                    If Array.IndexOf(sIgnores, dc.ColumnName.ToLower) < 0 Then
                        If dc.ColumnName.ToLower = sIdentity.ToLower Then
                            fIdentityParamUsed = True
                        End If
                        If sSelect <> "" Then
                            sSelect &= ","
                            sValues &= ","
                        End If

                        myCmd.Parameters.Add(csiDefineParameter(dc, ParameterDirection.Input, DataRowVersion.Default, cnSql))

                        sSelect &= dc.ColumnName
                        sValues &= "@" & dc.ColumnName

                    End If
                End If
            Next
            sSelect = "INSERT INTO " & sTable & " (" & sSelect & ")"
            sValues = "VALUES (" & sValues & ")"

            If fUseIdentity Then
                sValues &= " SELECT @" & sIdentity & "=@@Identity"
                If Not fIdentityParamUsed Then
                    myCmd.Parameters.Add(sIdentity, SqlDbType.Int, 4, sIdentity)
                End If
                myCmd.Parameters.Item(sIdentity).Direction = ParameterDirection.InputOutput
            End If

            myCmd.CommandText = sSelect & " " & sValues

            'AddHandler da.RowUpdated, AddressOf issGetNewID

            myCmd.Connection = cnSql

            Return myCmd

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function
    Function csiGetUpdateCommandForQuery(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sIdentity As String, ByVal fUseIdentity As Boolean, ByVal sIgnore As String) As SqlClient.SqlCommand
        Try
            Dim myCmd As New SqlClient.SqlCommand
            Dim sSets As String = ""

            'Build Parameters
            Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM " & sTable & " WHERE 1=0", cnSql)
            Dim dt As New DataTable
            Dim dc As DataColumn
            Dim sIgnores() As String = sIgnore.ToLower.Split("|")
            Dim fIdentityParamCreated As Boolean = False
            da.Fill(dt)
            For Each dc In dt.Columns
                If dc.ColumnName.ToUpper <> sIdentity.ToUpper Or Not fUseIdentity Then
                    If Array.IndexOf(sIgnores, dc.ColumnName.ToLower) < 0 Then
                        If dc.ColumnName.ToLower = sIdentity.ToLower Then
                            fIdentityParamCreated = True
                        End If
                        If sSets <> "" Then
                            sSets &= ","
                        End If
                        sSets &= dc.ColumnName & "=@" & dc.ColumnName
                        myCmd.Parameters.Add(csiDefineParameter(dc, ParameterDirection.Input, DataRowVersion.Current, cnSql))
                    End If
                End If
            Next

            'Add Parameter for sIdentity
            If Not fIdentityParamCreated Then
                myCmd.Parameters.Add(New SqlClient.SqlParameter("@" & sIdentity, SqlDbType.Int, 4, sIdentity))
            End If

            myCmd.CommandText = "UPDATE " & sTable & " SET " & sSets & " WHERE " & sIdentity & "=@" & sIdentity
            myCmd.Connection = cnSql

            Return myCmd

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function
#End Region
#Region "Table Functions"

    Function csiGetDataTable(ByVal ssql As String, ByVal mycn As SqlClient.SqlConnection) As DataTable
        Dim dt As DataTable = Nothing
        Try
            Dim da As New SqlClient.SqlDataAdapter(ssql, mycn)
            dt = New DataTable
            da.Fill(dt)
            da.Dispose()
            Return dt
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return dt
    End Function

    Function csiGetDataAdapter(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sPrimaryKeyField As String, Optional ByVal sWhere As String = "", Optional ByVal sInnerJoin As String = "", Optional ByVal fUseIdentity As Boolean = True) As SqlClient.SqlDataAdapter
        Dim da As New SqlClient.SqlDataAdapter
        Try
            'Get Primary Key
            da.SelectCommand = csiGetSelectCommand(sTable, cnSql, sInnerJoin, sWhere)
            da.InsertCommand = csiGetInsertCommand(sTable, cnSql, sPrimaryKeyField, fUseIdentity)
            da.UpdateCommand = csiGetUpdateCommand(sTable, cnSql, sPrimaryKeyField, fUseIdentity)
            da.DeleteCommand = csiGetDeleteCommand(sTable, cnSql, sPrimaryKeyField)
            Return da
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return da
    End Function

    Function csiGetSelectCommand(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sInnerJoin As String, ByVal sWhere As String) As SqlClient.SqlCommand
        Try
            Dim myCmd As New SqlClient.SqlCommand
            'Get the field list for this table
            Dim sFields As String = csiGetFieldList(sTable, cnSql)
            myCmd.CommandText = "SELECT " & sFields & " FROM " & sTable & " " & sInnerJoin
            If sWhere <> "" Then
                myCmd.CommandText &= " WHERE " & sWhere
            End If
            myCmd.Connection = cnSql
            myCmd.CommandType = CommandType.Text
            Return myCmd
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function

    Function csiGetInsertCommand(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sIdentity As String, ByVal fUseIdentity As Boolean) As SqlClient.SqlCommand
        Try
            Dim myCmd As New SqlClient.SqlCommand
            Dim sSelect As String = ""
            Dim sValues As String = ""
            Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM " & sTable & " WHERE 1=0", cnSql)
            Dim dt As New DataTable
            Dim dc As DataColumn
            Dim sFieldList As String = ""
            Dim fIdentityParamUsed As Boolean = False
            da.Fill(dt)
            For Each dc In dt.Columns
                If (dc.ColumnName.ToUpper <> sIdentity.ToUpper Or Not fUseIdentity) And dc.ColumnName.ToLower <> "rowguid" Then
                    If dc.ColumnName.ToLower = sIdentity.ToLower Then
                        fIdentityParamUsed = True
                    End If
                    If sSelect <> "" Then
                        sSelect &= ","
                        sValues &= ","
                    End If

                    myCmd.Parameters.Add(csiDefineParameter(dc, ParameterDirection.Input, DataRowVersion.Default, cnSql))

                    sSelect &= "[" & dc.ColumnName & "]"
                    sValues &= "@" & dc.ColumnName

                End If
            Next
            sSelect = "INSERT INTO " & sTable & " (" & sSelect & ")"
            sValues = "VALUES (" & sValues & ")"

            If fUseIdentity Then
                sValues &= " SELECT @" & sIdentity & "=@@Identity"
                If Not fIdentityParamUsed Then
                    myCmd.Parameters.Add(sIdentity, SqlDbType.Int, 4, sIdentity)
                End If
                myCmd.Parameters.Item(sIdentity).Direction = ParameterDirection.InputOutput
            End If

            myCmd.CommandText = sSelect & " " & sValues

            myCmd.Connection = cnSql

            Return myCmd

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function
    Function csiGetUpdateCommand(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sIdentity As String, ByVal fUseIdentity As Boolean) As SqlClient.SqlCommand
        Try
            Dim myCmd As New SqlClient.SqlCommand
            Dim sSets As String = ""

            'Build Parameters
            Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM " & sTable & " WHERE 1=0", cnSql)
            Dim dt As New DataTable
            Dim dc As DataColumn
            Dim fIdentityParamCreated As Boolean = False
            da.Fill(dt)
            For Each dc In dt.Columns
                If (dc.ColumnName.ToUpper <> sIdentity.ToUpper Or Not fUseIdentity) And dc.ColumnName.ToLower <> "rowguid" Then
                    If dc.ColumnName.ToLower = sIdentity.ToLower Then
                        fIdentityParamCreated = True
                    End If
                    If sSets <> "" Then
                        sSets &= ","
                    End If
                    sSets &= "[" & dc.ColumnName & "]=@" & dc.ColumnName
                    myCmd.Parameters.Add(csiDefineParameter(dc, ParameterDirection.Input, DataRowVersion.Current, cnSql))
                End If
            Next

            'Add Parameter for sIdentity
            If Not fIdentityParamCreated Then
                myCmd.Parameters.Add(New SqlClient.SqlParameter("@" & sIdentity, SqlDbType.Int, 4, sIdentity))
            End If

            myCmd.CommandText = "UPDATE " & sTable & " SET " & sSets & " WHERE " & sIdentity & "=@" & sIdentity
            myCmd.Connection = cnSql

            Return myCmd

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function

    Sub issGetNewID(ByVal sender As System.Object, ByVal e As SqlClient.SqlRowUpdatedEventArgs)
        Dim iNewID As Integer = 0
        If e.StatementType = StatementType.Insert Then
            Dim idCMD As New SqlClient.SqlCommand("SELECT @@IDENTITY", e.Command.Connection)
            iNewID = CInt(idCMD.ExecuteScalar)
            e.Row(0).item(0) = iNewID
        End If
    End Sub

    Function csiGetDeleteCommand(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection, ByVal sIdentity As String) As SqlClient.SqlCommand
        Try
            Dim myCmd As New SqlClient.SqlCommand

            myCmd.CommandText = "DELETE FROM " & sTable & " WHERE " & sIdentity & "=@" & sIdentity

            myCmd.Connection = cnSql


            'Build Parameters
            Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM " & sTable & " WHERE 1=0", cnSql)
            Dim dt As New DataTable
            Dim dc As DataColumn
            Dim sFieldList As String = ""
            da.Fill(dt)
            For Each dc In dt.Columns
                If dc.ColumnName.ToLower = sIdentity.ToLower Then
                    myCmd.Parameters.Add(csiDefineParameter(dc, ParameterDirection.Input, DataRowVersion.Original, cnSql))
                End If
            Next

            Return myCmd

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function

#End Region
#Region "Global Settings"
    Public Enum tySettingType
        xString
        xNumber
        xBoolean
    End Enum
    Function csiGetAppSetting(ByVal sSettingName As String, ByVal dbtType As tySettingType, ByVal cn As SqlClient.SqlConnection) As String
        Try
            Dim dt As DataTable = csiGetDataTable("SELECT * FROM tblGlobalSettings WHERE SettingName='" & sSettingName & "'", cn)

            Dim drr() As DataRow
            Dim dr As DataRow
            'Check for the username
            drr = dt.Select("UserName='" & Environment.UserName & "'")
            If drr.Length > 0 Then
                dr = drr(0)
            Else
                'Check machine name
                drr = dt.Select("MachineName='" & Environment.MachineName & "'")
                If drr.Length > 0 Then
                    dr = drr(0)
                Else
                    'Check Globals
                    drr = dt.Select("MachineName='Global' OR UserName='Global'")
                    If drr.Length > 0 Then
                        dr = drr(0)
                    Else
                        MessageBox.Show(sSettingName & " setting not found!")
                        Return ""
                    End If
                End If
            End If

            Select Case dbtType
                Case tySettingType.xString
                    Return dr.Item("ValueStr")
                Case tySettingType.xNumber
                    Return dr.Item("ValueNum")
                Case tySettingType.xBoolean
                    Return dr.Item("ValueBool")
                Case Else
                    Return ""
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return ""
    End Function
    Function csiGetSQLConnectionString(sDBName As String) As String
        Dim myCn As New SqlClient.SqlConnection("Data Source='SQL1';Initial Catalog='CaseGlobal';Integrated Security=sspi")
        Dim sRetVal As String = ""
        Try
            myCn.Open()

            Dim sSql As String = "SELECT Server FROM tblConnectionStrings WHERE DBName='" & sDBName & "'"

            Dim dc As New SqlClient.SqlCommand(sSql, myCn)
            Dim dr As SqlClient.SqlDataReader = dc.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                sRetVal = dr.Item(0).ToString
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        Finally
            If myCn.State = ConnectionState.Open Then
                myCn.Close()
            End If
        End Try

        sRetVal = "Data Source='" & sRetVal & "';Initial Catalog='" & sDBName & "'; Integrated Security=sspi"

        Return sRetVal
    End Function
#End Region
    Function csiDRHasChanges(ByVal dr As DataRow) As Boolean
        Try
            Dim dc As DataColumn
            For Each dc In dr.Table.Columns
                If dr.Item(dc.ColumnName) Is DBNull.Value And dr.Item(dc.ColumnName, DataRowVersion.Original) Is DBNull.Value Then
                Else
                    If (dr.Item(dc.ColumnName) Is DBNull.Value And Not dr.Item(dc.ColumnName, DataRowVersion.Original) Is DBNull.Value) Or (Not dr.Item(dc.ColumnName) Is DBNull.Value And dr.Item(dc.ColumnName, DataRowVersion.Original) Is DBNull.Value) Then
                        Return True
                    Else
                        If dr.Item(dc.ColumnName) <> dr.Item(dc.ColumnName, DataRowVersion.Original) Then
                            Return True
                        End If
                    End If
                End If
            Next
            Return False
        Catch ex As Exception
            Return True
        End Try
    End Function
    Function csiDefineParameter(ByVal dc As DataColumn, ByVal iDirection As ParameterDirection, ByVal iRowVersion As DataRowVersion, ByVal cnSql As SqlClient.SqlConnection) As SqlClient.SqlParameter
        Dim prmParam As SqlClient.SqlParameter

        Try
            prmParam = New SqlClient.SqlParameter
            With prmParam
                .ParameterName = "@" & dc.ColumnName
                .Direction = ParameterDirection.Input
                .SourceColumn = dc.ColumnName
                .SourceVersion = iRowVersion
                .Direction = iDirection
                If dc.ColumnName = "TopBLOB" Then
                    .SqlDbType = SqlDbType.Image
                Else
                    .SqlDbType = TypeToSqlDbType(dc.DataType)
                    .Size = SizeOfColumn(dc)
                End If
            End With

            Return prmParam

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function

    Function csiGetFieldList(ByVal sTable As String, ByVal cnSql As SqlClient.SqlConnection) As String
        Try
            'Return a string of fields for this table separated by "," for a SQL Server Table
            Dim da As New SqlClient.SqlDataAdapter("SELECT * FROM " & sTable & " WHERE 1=0", cnSql)
            Dim dt As New DataTable
            Dim dc As DataColumn
            Dim sFieldList As String = ""
            da.Fill(dt)
            For Each dc In dt.Columns
                If sFieldList <> "" Then
                    sFieldList &= ","
                End If
                sFieldList &= sTable & "." & dc.ColumnName
            Next
            Return (sFieldList)
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function

    Friend Function SizeOfColumn(ByVal pdcColumn As DataColumn) As Integer
        Dim myTypeCode As TypeCode
        Try
            myTypeCode = Type.GetTypeCode(pdcColumn.DataType)
            Select Case myTypeCode
                Case TypeCode.Boolean
                    Return 2
                Case TypeCode.Byte
                    Return 1
                Case TypeCode.Char
                    Return 2
                Case TypeCode.DateTime
                    Return 8
                Case TypeCode.DBNull
                    Return 4
                Case TypeCode.Decimal
                    Return 16
                Case TypeCode.Double
                    Return 8
                Case TypeCode.Empty
                    Return 4
                Case TypeCode.Int16
                    Return 2
                Case TypeCode.Int32
                    Return 4
                Case TypeCode.Int64
                    Return 8
                Case TypeCode.Object
                    Return 4
                Case TypeCode.SByte
                    Return 1
                Case TypeCode.Single
                    Return 4
                Case TypeCode.String
                    If pdcColumn.MaxLength > 0 Then
                        Return pdcColumn.MaxLength
                    Else
                        Return 0
                    End If
                Case TypeCode.UInt16
                    Return 2
                Case TypeCode.UInt32
                    Return 4
                Case TypeCode.UInt64
                    Return 8
            End Select

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function
    Friend Function TypeToSqlDbType(ByVal dbtType As System.Type) As SqlDbType
        Dim myTypeCode As TypeCode

        Try
            myTypeCode = Type.GetTypeCode(dbtType)
            Select Case myTypeCode
                Case TypeCode.Boolean
                    Return SqlDbType.Bit
                Case TypeCode.Byte
                    Return SqlDbType.Image
                Case TypeCode.Char
                    Return SqlDbType.VarChar
                Case TypeCode.DateTime
                    Return SqlDbType.DateTime
                Case TypeCode.DBNull
                    Return SqlDbType.Int
                Case TypeCode.Decimal
                    Return SqlDbType.Decimal
                Case TypeCode.Double
                    Return SqlDbType.Float
                Case TypeCode.Empty
                    Return SqlDbType.Text
                Case TypeCode.Int16
                    Return SqlDbType.SmallInt
                Case TypeCode.Int32
                    Return SqlDbType.Int
                Case TypeCode.Int64
                    Return SqlDbType.BigInt
                Case TypeCode.Object
                    'Return SqlDbType.Variant
                    Return SqlDbType.Binary
                Case TypeCode.SByte
                    Return SqlDbType.Image
                Case TypeCode.Single
                    Return SqlDbType.Real
                Case TypeCode.String
                    Return SqlDbType.VarChar
                Case TypeCode.UInt16
                    Return SqlDbType.UniqueIdentifier
                Case TypeCode.UInt32
                    Return SqlDbType.BigInt
                Case TypeCode.UInt64
                    Return SqlDbType.BigInt
                Case TypeCode.Byte

            End Select

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
        Return Nothing
    End Function


    Sub Ops_Log(strModule As String, strProcess As String, strMsg As String)
        Dim strFile As String = "log_" & Format(Now(), "yyyyMMdd") & ".txt"
        Dim fileExists As Boolean = File.Exists(strFile)
        Using sw As New StreamWriter(File.Open(strFile, FileMode.OpenOrCreate))
            sw.WriteLine(
                IIf(fileExists, DateTime.Now & "- Module: " & strModule & ", Process: " & strProcess & ", Message: " & strMsg, "******** Start Log ********"))
        End Using
    End Sub
End Module
