Module basUtilities

    Public Const dRad_to_Degree_Const As Double = 180 / Math.PI
    Public Const sPreAngle As String = "Ð"
    Public Const sPostAngle As String = "°"

    Public Sub csiRotateText(ByVal g As Graphics, ByVal f As Font, ByVal s As String, ByVal angle As Single, ByVal b As Brush, ByVal x As Single, ByVal y As Single)
        Try
            If angle > 360 Then
                While angle > 360
                    angle = angle - 360
                End While
            ElseIf angle < 0 Then
                While angle < 0
                    angle = angle + 360
                End While
            End If

            ' Create a matrix and rotate it n degrees.
            g.TranslateTransform(x, y)
            g.RotateTransform(angle)

            g.DrawString(s, f, b, 0, 0)

            'Rotate back
            g.RotateTransform(-angle)
            g.TranslateTransform(-x, -y)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub
    Function csiCleanDollarMask(ByVal sDollar As String) As Double
        Try
            Dim sRet As String = sDollar.Replace("$", "")
            sRet = sRet.Replace(",", "")
            Return ez(sRet)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function
    Function csiCleanPercentMask(ByVal sPerc As String) As Double
        Try
            Dim sRet As String = sPerc.Replace("%", "")
            sRet = sRet.Replace(",", "")
            Return ez(sRet)
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    Function csiCheckWebConnection() As Boolean
        Try
            If My.Computer.Network.IsAvailable Then
                ' You can pass a URL or an IP address to this function
                Try
                    If My.Computer.Network.Ping("www.CaseSystems.com", 1000) Then
                        Return True
                    Else
                        Return False
                    End If
                Catch ex As Exception
                    Return False
                End Try
            Else
                Return False
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.ToString)
            Return False
        End Try
    End Function
    Function csiReverseArraySplit(ByVal sArray() As String, ByVal sSeparator As String) As String
        Try
            Dim sC As String
            Dim sTmp As String = ""
            For Each sC In sArray
                If sTmp <> "" Then
                    sTmp &= sSeparator
                End If
                sTmp &= sC
            Next
            Return sTmp
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return ""
    End Function

    Enum eListViewSortType
        eText
        eNumber
        eDate
    End Enum

#Region "Zip Code"
    Function csiGZipFile(ByVal sFile As String, Optional ByVal sDestination As String = "") As String
        Try
            Dim sExe As String
            Dim sDestFile As String = ""

            sExe = Application.StartupPath & "\gzip.exe"


            'Parameters---- sFile = File you want to compress
            '           --- sDestination = Directory you want the compressed file to be placed in
            '                               If this is left blank, the original file will be compressed

            'GZIP.EXE must be in the same dir as your program exe
            'First Copy the file to the destination then run gzip on it
            If IO.File.Exists(sFile) Then

                'Check to copy file
                If sDestination <> "" Then
                    If Right(sDestination, 1) = "\" Then
                        sDestination = Left(sDestination, Len(sDestination) - 1)
                    End If
                    sDestFile = sDestination & "\" & csiGetFileFromPath(sFile)
                    IO.File.Copy(sFile, sDestFile, True)
                Else
                    sDestFile = sFile
                End If
                If IO.File.Exists(sExe) Then
                    'Create Process to run GZIP on the sDestFile
                    'gzip.exe filename
                    If IO.File.Exists(sDestFile & ".gz") Then
                        IO.File.Delete(sDestFile & ".gz")
                    End If


                    Shell(Application.StartupPath & "\gzip.exe """ & sDestFile & "", vbNormalFocus, True)

                    'Loop until file exists
                    sDestFile = sDestFile & ".gz"
                    Do While Not IO.File.Exists(sDestFile)
                        Application.DoEvents()
                    Loop
                    Return sDestFile
                Else
                    MessageBox.Show("GZip.exe needs to be located in " & Application.StartupPath)
                    Return sDestFile
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return Nothing
    End Function
    Function csiGUnZipFile(ByVal sFile As String, Optional ByVal sDestination As String = "") As String
        Try
            Dim sExe As String
            Dim sDestFile As String = ""

            sExe = Application.StartupPath & "\gzip.exe"


            'Parameters---- sFile = File you want to compress
            '           --- sDestination = Directory you want the compressed file to be placed in
            '                               If this is left blank, the original file will be compressed

            'GZIP.EXE must be in the same dir as your program exe
            'First Copy the file to the destination then run gzip on it
            If IO.File.Exists(sFile) Then

                'Check to copy file
                If sDestination <> "" Then
                    If Right(sDestination, 1) = "\" Then
                        sDestination = Left(sDestination, Len(sDestination) - 1)
                    End If
                    sDestFile = sDestination & "\" & csiGetFileFromPath(sFile)
                    IO.File.Copy(sFile, sDestFile, True)
                Else
                    sDestFile = sFile
                End If
                If IO.File.Exists(sExe) Then
                    'Create Process to run GZIP on the sDestFile
                    'gzip.exe filename
                    If IO.File.Exists(sDestFile.Replace(".gz", "")) Then
                        IO.File.Exists(sDestFile.Replace(".gz", ""))
                    End If


                    Shell(Application.StartupPath & "\gzip.exe -d """ & sDestFile & "", vbNormalFocus, True)

                    'Loop until file exists
                    sDestFile = sDestFile.Replace(".gz", "")
                    Do While Not IO.File.Exists(sDestFile)
                        Application.DoEvents()
                    Loop
                    Return sDestFile
                Else
                    MessageBox.Show("GZip.exe needs to be located in " & Application.StartupPath)
                    Return sDestFile
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return Nothing
    End Function
#End Region

    Sub csiSortListview(ByVal dList As ListView, ByVal iCol As Integer, Optional ByVal iType As eListViewSortType = eListViewSortType.eText)
        Try

            dList.BeginUpdate()

            'Tag of listview holds last sort, a/d & col#.  a means ascending, d means descending
            Dim fDesc As Boolean
            Dim iC As Integer
            If dList.Tag <> "" Then
                If dList.Tag.ToString.Substring(0, 1) = "a" Then
                    fDesc = False
                Else
                    fDesc = True
                End If
                iC = dList.Tag.ToString.Replace("a", "").Replace("d", "")
                If iC = iCol Then
                    'Same column, switch asc/desc
                    fDesc = Not fDesc
                Else
                    fDesc = False
                End If
            Else
                fDesc = False
            End If


            Dim arSort(dList.Items.Count) As String

            Dim iY As Integer = 0
            Dim iX As Integer = 0
            Dim itmX As ListViewItem
            Dim colX As ColumnHeader
            Dim sBuffer As String
            Dim sKey As String = ""

            'Load the listview into a ListSorter
            For Each itmX In dList.Items
                iY = 0
                sBuffer = ""
                For Each colX In dList.Columns
                    If iY = 0 Then
                        sBuffer = itmX.Text
                    Else
                        sBuffer &= "|" & itmX.SubItems(iY).Text
                    End If
                    If iY = iCol Then
                        sKey = itmX.SubItems(iY).Text
                    End If
                    iY += 1
                Next

                arSort(iX) = sKey & "|" & sBuffer
                iX += 1
            Next

            'Sort the array
            Array.Sort(arSort)


            Dim iFrom As Integer
            Dim iTo As Integer
            Dim iStep As Integer
            Dim arItem() As String

            'Load the ListSorter back into the list
            dList.Items.Clear()
            If fDesc Then
                iFrom = iX
                iTo = 1
                iStep = -1
            Else
                iFrom = 1
                iTo = iX
                iStep = 1
            End If

            For iC = iFrom To iTo Step iStep
                arItem = arSort(iC).ToString.Split("|")
                'First Item is the sort key, just ignore it
                itmX = dList.Items.Add(arItem(1))
                For iY = 2 To arItem.GetUpperBound(0)
                    itmX.SubItems.Add(arItem(iY))
                Next
            Next


            If fDesc Then
                dList.Tag = "d" & iCol
            Else
                dList.Tag = "a" & iCol
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        Finally
            dList.EndUpdate()
        End Try
    End Sub
    Function csiGetFileFromPath(ByVal strPath) As String
        Try


            Dim x As Integer
            Dim intStartPos As Integer

            intStartPos = 1

            Do
                x = InStr(intStartPos, strPath, "\")
                If x = 0 Then Exit Do
                intStartPos = x + 1
            Loop
            Return Mid(strPath, intStartPos)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return ""
    End Function
    Function csiRoundDecimalFields(ByVal dr As DataRow, ByVal iP As Integer) As DataRow
        Try
            Dim dc As DataColumn
            For Each dc In dr.Table.Columns
                If dc.DataType Is GetType(Double) Or dc.DataType Is GetType(Single) Then
                    dr.Item(dc.ColumnName) = Math.Round(ez(dr.Item(dc.ColumnName)), iP)
                End If
            Next
            Return dr
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return dr
    End Function
    Function csiMergeDigitStrings(ByVal sDigit1 As String, ByVal sDigit2 As String) As String
        Try
            Dim sTmp As String = ""
            Dim iC As Integer
            For iC = 0 To sDigit1.Length - 1
                If sDigit2.Length > iC Then
                    If sDigit1.Substring(iC, 1) = 1 Or sDigit2.Substring(iC, 1) = 1 Then
                        sTmp &= "1"
                    Else
                        sTmp &= "0"
                    End If
                Else
                    sTmp &= sDigit1.Substring(iC, 1)
                End If
            Next
            If iC < sDigit2.Length Then
                For iC = iC To sDigit2.Length - 1
                    sTmp &= sDigit2.Substring(iC, 1)
                Next
            End If
            Return sTmp
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
        Return ""
    End Function

    Function csiGetDirectoryFromPath(ByVal sPath As String) As String
        Return IO.Directory.GetParent(sPath).FullName
    End Function
    Function csiGetControlFromName(ByRef containerObj As Object,
                     ByVal name As String) As Control
        Try
            Dim tempCtrl As Control
            For Each tempCtrl In containerObj.Controls
                If tempCtrl.Name.ToUpper.Trim = name.ToUpper.Trim Then
                    Return tempCtrl
                End If
            Next tempCtrl
            'MessageBox.Show("(" & name & ") control not found!")
        Catch ex As Exception
        End Try
        Return Nothing
    End Function
    Function csiGetControlByTag(ByVal sTag As String, ByRef containerObj As Object) As Control
        Try
            Dim tempCtrl As Control
            For Each tempCtrl In containerObj.Controls
                If Not IsNothing(tempCtrl.Tag) Then
                    If tempCtrl.Tag.ToString.ToUpper.Trim = sTag.ToUpper.Trim Then
                        Return tempCtrl
                    End If
                End If
            Next tempCtrl
            'MessageBox.Show("(" & name & ") control not found!")
        Catch ex As Exception
        End Try
        Return Nothing
    End Function
    Public Function nz(ByVal vData As Object) As String
        Try
            If IsDBNull(vData) Or IsNothing(vData) Then
                Return "0"
            Else
                Return vData
            End If
        Catch ex As Exception
            Return "0"
        End Try
    End Function
    Public Function nz(ByVal vData As Object, ByVal sIfIsNull As String) As String
        Try
            If IsDBNull(vData) Or IsNothing(vData) Then
                Return sIfIsNull
            Else
                Return vData
            End If
        Catch ex As Exception
            Return sIfIsNull
        End Try
    End Function
    Public Function fnz(ByVal vData As Object, fIfIsNull As Boolean) As Boolean
        Try
            If IsDBNull(vData) Or IsNothing(vData) Then
                Return fIfIsNull
            Else
                Return vData
            End If
        Catch ex As Exception
            Return fIfIsNull
        End Try
    End Function
    Public Function ez(ByVal vData As Object, Optional ByVal sIfIsEmpty As Object = "0") As Object
        Try
            If vData.ToString = "" Then
                Return sIfIsEmpty
            Else
                Return vData
            End If
        Catch e As Exception
            Return sIfIsEmpty
        End Try
    End Function
    Public Function zn(ByVal vData As Object, Optional ByVal sIfZero As Object = Nothing) As Object
        Try
            If vData.ToString = "0" Then
                If sIfZero Is Nothing Then
                    Return DBNull.Value
                Else
                    Return sIfZero
                End If
            Else
                Return vData
            End If
        Catch ex As Exception
            Return sIfZero
        End Try
    End Function

    Public Function csiDlookup(ByVal sField As String, ByVal sWhere As String, ByVal sOrderBy As String, ByVal dt As DataTable) As String
        Dim drr() As DataRow
        Try
            If sOrderBy = "" Then
                drr = dt.Select(sWhere)
            Else
                drr = dt.Select(sWhere, sOrderBy)
            End If

            If drr.GetUpperBound(0) >= 0 Then
                If drr(0).Item(sField) Is Nothing Then
                    Return Nothing
                Else
                    Return nz(drr(0).Item(sField), Nothing)
                End If

            Else
                Return Nothing
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        Finally
        End Try
        Return Nothing
    End Function
    Public Function csiDlookup(ByVal sTable As String, ByVal sField As String, ByVal sWhere As String, ByVal myCn As SqlClient.SqlConnection) As String
        Dim sSql As String
        Dim ds As SqlClient.SqlDataReader = Nothing
        Dim dc As SqlClient.SqlCommand
        Try

            sSql = "SET CONCAT_NULL_YIELDS_NULL OFF; "
            sSql &= "SELECT " & sField & " FROM " & sTable & " WHERE " & sWhere & ""

            If myCn.State = ConnectionState.Closed Then
                myCn.Open()
            End If
            dc = New SqlClient.SqlCommand(sSql, myCn)
            ds = dc.ExecuteReader

            With ds
                .Read()
                If .HasRows Then
                    Return nz(.Item(0), Nothing)
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        Finally
            If Not ds Is Nothing Then
                ds.Close()
            End If
            If myCn.State = ConnectionState.Open Then
                myCn.Close()
            End If
        End Try
        Return Nothing
    End Function
    Function csiDSum(ByVal sTable As String, ByVal sField As String, ByVal sWhere As String, ByVal myCn As SqlClient.SqlConnection) As Double
        Try
            Dim dt As DataTable = csiGetDataTable("SELECT sum(" & sField & ") FROM " & sTable & " WHERE " & sWhere, myCn)
            If dt.Rows.Count > 0 Then
                Return dt.Rows(0).Item(0)
            Else
                Return 0
            End If
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function
    Function csiDSum(ByVal sField As String, ByVal sWhere As String, ByVal dt As DataTable) As Double
        Try
            Dim dRetVal As Double = 0
            For Each dr As DataRow In dt.Select(sWhere)
                dRetVal += dr.Item(sField)
            Next
            Return dRetVal
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    Public Sub csiLoadStates(ByVal dlList As ComboBox)

        dlList.Items.Clear()
        dlList.Items.Add("")
        dlList.Items.Add("AL")
        dlList.Items.Add("AK")
        dlList.Items.Add("AZ")
        dlList.Items.Add("AR")
        dlList.Items.Add("CA")
        dlList.Items.Add("CO")
        dlList.Items.Add("CT")
        dlList.Items.Add("DE")
        dlList.Items.Add("FL")
        dlList.Items.Add("GA")
        dlList.Items.Add("HI")
        dlList.Items.Add("ID")
        dlList.Items.Add("IL")
        dlList.Items.Add("IN")
        dlList.Items.Add("IA")
        dlList.Items.Add("KS")
        dlList.Items.Add("KY")
        dlList.Items.Add("LA")
        dlList.Items.Add("ME")
        dlList.Items.Add("MD")
        dlList.Items.Add("MA")
        dlList.Items.Add("MI")
        dlList.Items.Add("MN")
        dlList.Items.Add("MS")
        dlList.Items.Add("MO")
        dlList.Items.Add("MT")
        dlList.Items.Add("NE")
        dlList.Items.Add("NV")
        dlList.Items.Add("NH")
        dlList.Items.Add("NJ")
        dlList.Items.Add("NM")
        dlList.Items.Add("NY")
        dlList.Items.Add("NC")
        dlList.Items.Add("ND")
        dlList.Items.Add("OH")
        dlList.Items.Add("OK")
        dlList.Items.Add("OR")
        dlList.Items.Add("PA")
        dlList.Items.Add("RI")
        dlList.Items.Add("SC")
        dlList.Items.Add("SD")
        dlList.Items.Add("TN")
        dlList.Items.Add("TX")
        dlList.Items.Add("UT")
        dlList.Items.Add("VT")
        dlList.Items.Add("VA")
        dlList.Items.Add("WA")
        dlList.Items.Add("WV")
        dlList.Items.Add("WI")
        dlList.Items.Add("WY")

    End Sub
    Public Sub csiApplyPhoneMask(ByVal txt As TextBox)
        Dim sRetVal As String

        '(000) 000-0000
        sRetVal = csiCleanPhoneMask(txt.Text)

        Select Case Len(sRetVal)
            Case 1 To 3
                txt.Text = "(" & sRetVal & ")"
                txt.SelectionStart = 1 + Len(sRetVal)
            Case 4 To 6
                txt.Text = "(" & Left(sRetVal, 3) & ")" & " " & Mid(sRetVal, 4)
                txt.SelectionStart = 3 + Len(sRetVal)
            Case Is > 6
                txt.Text = "(" & Left(sRetVal, 3) & ")" & " " & Mid(sRetVal, 4, 3) & "-" & Mid(sRetVal, 7)
                txt.SelectionStart = 4 + Len(sRetVal)
        End Select
    End Sub
    Public Function csiCleanPhoneMask(ByVal sStr As String) As String

        sStr = Replace(sStr, "(", "")
        sStr = Replace(sStr, ")", "")
        sStr = Replace(sStr, " ", "")
        sStr = Replace(sStr, "-", "")

        csiCleanPhoneMask = sStr
    End Function
    Public Sub csiLoadCountrys(ByVal dlList As ComboBox)

        With dlList.Items
            .Clear()
            .Add("")
            .Add("UNITED STATES")
            .Add("CANADA")
            .Add("AFGHANISTAN")
            .Add("ALBANIA")
            .Add("ALGERIA")
            .Add("AMERICAN SAMOA")
            .Add("ANDORRA")
            .Add("ANGOLA")
            .Add("ANGUILLA")
            .Add("ANTARCTICA")
            .Add("ANTIGUA AND BARBUDA")
            .Add("ARGENTINA")
            .Add("ARMENIA")
            .Add("ARUBA")
            .Add("AUSTRALIA")
            .Add("AUSTRIA")
            .Add("AZERBAIJAN")
            .Add("BAHAMAS")
            .Add("BAHRAIN")
            .Add("BANGLADESH")
            .Add("BARBADOS")
            .Add("BELARUS")
            .Add("BELGIUM")
            .Add("BELIZE")
            .Add("BENIN")
            .Add("BERMUDA")
            .Add("BHUTAN")
            .Add("BOLIVIA")
            .Add("BOSNIA AND HERZEGOWINA")
            .Add("BOTSWANA")
            .Add("BOUVET ISLAND")
            .Add("BRAZIL")
            .Add("BRITISH INDIAN OCEAN TERRITORY")
            .Add("BRUNEI DARUSSALAM")
            .Add("BULGARIA")
            .Add("BURKINA FASO")
            .Add("BURUNDI")
            .Add("CAMBODIA")
            .Add("CAMEROON")
            .Add("CANADA")
            .Add("CAPE VERDE")
            .Add("CAYMAN ISLANDS")
            .Add("CENTRAL AFRICAN REPUBLIC")
            .Add("CHAD")
            .Add("CHILE")
            .Add("CHINA")
            .Add("CHRISTMAS ISLAND")
            .Add("COCOS (KEELING) ISLANDS")
            .Add("COLOMBIA")
            .Add("COMOROS")
            .Add("CONGO")
            .Add("CONGO, THE DEMOCRATIC REPUBLIC OF THE")
            .Add("COOK ISLANDS")
            .Add("COSTA RICA")
            .Add("COTE D'IVOIRE")
            .Add("CROATIA (local name: Hrvatska)")
            .Add("CUBA")
            .Add("CYPRUS")
            .Add("CZECH REPUBLIC")
            .Add("DENMARK")
            .Add("DJIBOUTI")
            .Add("DOMINICA")
            .Add("DOMINICAN REPUBLIC")
            .Add("EAST TIMOR")
            .Add("ECUADOR")
            .Add("EGYPT")
            .Add("EL SALVADOR")
            .Add("EQUATORIAL GUINEA")
            .Add("ERITREA")
            .Add("ESTONIA")
            .Add("ETHIOPIA")
            .Add("FALKLAND ISLANDS (MALVINAS)")
            .Add("FAROE ISLANDS")
            .Add("FIJI")
            .Add("FINLAND")
            .Add("FRANCE")
            .Add("FRANCE, METROPOLITAN")
            .Add("FRENCH GUIANA")
            .Add("FRENCH POLYNESIA")
            .Add("FRENCH SOUTHERN TERRITORIES")
            .Add("GABON")
            .Add("GAMBIA")
            .Add("GEORGIA")
            .Add("GERMANY")
            .Add("GHANA")
            .Add("GIBRALTAR")
            .Add("GREECE")
            .Add("GREENLAND")
            .Add("GRENADA")
            .Add("GUADELOUPE")
            .Add("GUAM")
            .Add("GUATEMALA")
            .Add("GUINEA")
            .Add("GUINEA-BISSAU")
            .Add("GUYANA")
            .Add("HAITI")
            .Add("HEARD AND MC DONALD ISLANDS")
            .Add("HOLY SEE (VATICAN CITY STATE)")
            .Add("HONDURAS")
            .Add("HONG KONG")
            .Add("HUNGARY")
            .Add("ICELAND")
            .Add("INDIA")
            .Add("INDONESIA")
            .Add("IRAN (ISLAMIC REPUBLIC OF)")
            .Add("IRAQ")
            .Add("IRELAND")
            .Add("ISRAEL")
            .Add("ITALY")
            .Add("JAMAICA")
            .Add("JAPAN")
            .Add("JORDAN")
            .Add("KAZAKHSTAN")
            .Add("KENYA")
            .Add("KIRIBATI")
            .Add("KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF")
            .Add("KOREA, REPUBLIC OF")
            .Add("KUWAIT")
            .Add("KYRGYZSTAN")
            .Add("LAO PEOPLE'S DEMOCRATIC REPUBLIC")
            .Add("LATVIA")
            .Add("LEBANON")
            .Add("LESOTHO")
            .Add("LIBERIA")
            .Add("LIBYAN ARAB JAMAHIRIYA")
            .Add("LIECHTENSTEIN")
            .Add("LITHUANIA")
            .Add("LUXEMBOURG")
            .Add("MACAU")
            .Add("MACEDONIA, THE FORMER YUGOSLAV REPUBLIC")
            .Add("MADAGASCAR")
            .Add("MALAWI")
            .Add("MALAYSIA")
            .Add("MALDIVES")
            .Add("MALI")
            .Add("MALTA")
            .Add("MARSHALL ISLANDS")
            .Add("MARTINIQUE")
            .Add("MAURITANIA")
            .Add("MAURITIUS")
            .Add("MAYOTTE")
            .Add("MEXICO")
            .Add("MICRONESIA, FEDERATED STATES OF")
            .Add("MOLDOVA, REPUBLIC OF")
            .Add("MONACO")
            .Add("MONGOLIA")
            .Add("MONTSERRAT")
            .Add("MOROCCO")
            .Add("MOZAMBIQUE")
            .Add("MYANMAR (Burma)")
            .Add("NAMIBIA")
            .Add("NAURU")
            .Add("NEPAL")
            .Add("NETHERLANDS")
            .Add("NETHERLANDS ANTILLES")
            .Add("NEW CALEDONIA")
            .Add("NEW ZEALAND")
            .Add("NICARAGUA")
            .Add("NIGER")
            .Add("NIGERIA")
            .Add("NIUE")
            .Add("NORFOLK ISLAND")
            .Add("NORTHERN MARIANA ISLANDS")
            .Add("NORWAY")
            .Add("OMAN")
            .Add("PAKISTAN")
            .Add("PALAU")
            .Add("PANAMA")
            .Add("PAPUA NEW GUINEA")
            .Add("PARAGUAY")
            .Add("PERU")
            .Add("PHILIPPINES")
            .Add("PITCAIRN")
            .Add("POLAND")
            .Add("PORTUGAL")
            .Add("PUERTO RICO")
            .Add("QATAR")
            .Add("REUNION")
            .Add("ROMANIA")
            .Add("RUSSIAN FEDERATION")
            .Add("RWANDA")
            .Add("SAINT KITTS AND NEVIS")
            .Add("SAINT LUCIA")
            .Add("SAINT VINCENT AND THE GRENADINES")
            .Add("SAMOA")
            .Add("SAN MARINO")
            .Add("SAO TOME AND PRINCIPE")
            .Add("SAUDI ARABIA")
            .Add("SENEGAL")
            .Add("SEYCHELLES")
            .Add("SIERRA LEONE")
            .Add("SINGAPORE")
            .Add("SLOVAKIA (Slovak Republic)")
            .Add("SLOVENIA")
            .Add("SOLOMON ISLANDS")
            .Add("SOMALIA")
            .Add("SOUTH AFRICA")
            .Add("SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS")
            .Add("SPAIN")
            .Add("SRI LANKA")
            .Add("ST. HELENA")
            .Add("ST. PIERRE AND MIQUELON")
            .Add("SUDAN")
            .Add("SURINAME")
            .Add("SVALBARD AND JAN MAYEN ISLANDS")
            .Add("SWAZILAND")
            .Add("SWEDEN")
            .Add("SWITZERLAND")
            .Add("SYRIAN ARAB REPUBLIC")
            .Add("TAIWAN, PROVINCE OF CHINA")
            .Add("TAJIKISTAN")
            .Add("TANZANIA, UNITED REPUBLIC OF")
            .Add("THAILAND")
            .Add("TOGO")
            .Add("TOKELAU")
            .Add("TONGA")
            .Add("TRINIDAD AND TOBAGO")
            .Add("TUNISIA")
            .Add("TURKEY")
            .Add("TURKMENISTAN")
            .Add("TURKS AND CAICOS ISLANDS")
            .Add("TUVALU")
            .Add("UGANDA")
            .Add("UKRAINE")
            .Add("UNITED ARAB EMIRATES")
            .Add("UNITED KINGDOM")
            .Add("UNITED STATES")
            .Add("UNITED STATES MINOR OUTLYING ISLANDS")
            .Add("URUGUAY")
            .Add("UZBEKISTAN")
            .Add("VANUATU")
            .Add("VENEZUELA")
            .Add("VIET NAM")
            .Add("VIRGIN ISLANDS (BRITISH)")
            .Add("VIRGIN ISLANDS (U.S.)")
            .Add("WALLIS AND FUTUNA ISLANDS")
            .Add("WESTERN SAHARA")
            .Add("YEMEN")
            .Add("YUGOSLAVIA")
            .Add("ZAMBIA")
            .Add("ZIMBABWE")
        End With
    End Sub
    Public Function csiGetListViewIndex(ByVal sText As String, ByVal vListView As ListView)
        Try
            Dim itmX As ListViewItem
            For Each itmX In vListView.Items
                If itmX.Text = sText Then
                    Return itmX.Index
                    Exit For
                End If
            Next
            Return -1
        Catch ex As Exception
            MessageBox.Show(ex.Message & "--" & ex.StackTrace.Substring(ex.StackTrace.Length - 10))
        End Try
        Return -1
    End Function
    Public Function csiGetListViewColIndex(ByVal sText As String, ByRef vListView As ListView, Optional ByVal fCreateIfNotFound As Boolean = False)
        Try
            Dim iC As ColumnHeader
            For Each iC In vListView.Columns
                If iC.Text = sText Then
                    Return iC.Index
                End If
            Next
            If fCreateIfNotFound Then
                iC = vListView.Columns.Add(sText, -2, HorizontalAlignment.Left)
                iC.Width = sText.Length * 8.8
                Return iC.Index
            Else
                Return -1
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message & "--" & ex.StackTrace.Substring(ex.StackTrace.Length - 10))
        End Try
        Return -1
    End Function
    Public Function csiFixDollarAmt(ByVal sAmt As String) As Double
        Try
            Return sAmt.Replace("$", "").Replace(",", "")
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return Nothing
        End Try
    End Function

    Public Function csiSafeDivide(ByVal sNumerator As Double, ByVal sDenominator As Double) As Double
        If sDenominator <> 0 Then
            csiSafeDivide = sNumerator / sDenominator
        Else
            csiSafeDivide = 0
        End If
    End Function
    Public Function csiGetJulianDate() As Integer
        Dim iRet As Integer = Now.DayOfYear
        iRet = iRet + ((Now.Year - 1900) * 365)
        Return iRet
    End Function
    Public Sub csiGetCheckedNodes(ByVal myParentNodes As TreeNodeCollection, ByRef arNodes() As TreeNode)
        Try
            Dim myNode As TreeNode

            For Each myNode In myParentNodes
                If myNode.Checked Then
                    'Add to the array
                    If IsNothing(arNodes) Then
                        Array.Resize(arNodes, 1)
                    Else
                        Array.Resize(arNodes, arNodes.GetUpperBound(0) + 2)
                    End If

                    arNodes(arNodes.GetUpperBound(0)) = myNode
                End If

                csiGetCheckedNodes(myNode.Nodes, arNodes)

            Next
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Function csiGetMyVersion() As String
        Try

            Dim myBuildInfo As System.Diagnostics.FileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
            Dim sRetVal As String = myBuildInfo.FileMajorPart & "." & myBuildInfo.FileMinorPart & "." & myBuildInfo.FileBuildPart

            Return sRetVal

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return "v."
        End Try
    End Function
    Function csiCleanFileName(ByVal sFileName As String) As String
        Try
            Dim sC() As Char = IO.Path.GetInvalidFileNameChars
            Dim sChar As Char
            Dim sBuffer As String = sFileName
            For Each sChar In sC
                sBuffer = sBuffer.Replace(sChar, "")
            Next
            Return sBuffer
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            Return "error"
        End Try
    End Function

End Module
