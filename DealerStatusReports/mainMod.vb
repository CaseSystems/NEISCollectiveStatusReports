Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.IO
Imports System.Net.Mail
Imports System.Runtime.InteropServices

Module mainMod
    Const Smtp_Server As String = "172.16.10.95"

    Const TempDocFolder As String = "\\fs2\Apps\joblog08\autorpts\neiscollective\"
    Const ReportExtension As String = ".pdf"
    Const DefaultReplyTo As String = "Char.Lloyd-King@casesystems.com"
    Const DefaultFromAddr As String = "shippingnotices@casesystems.com"

    Dim intStatCounter As Integer
    Dim intPriceCounter As Integer
    Dim intInfoDueCounter As Integer
    Dim strParam(3) As String


    '*****************************************************************
    '* Name:    Main                                                 *
    '* Purpose: Determine what reports to process based on           *
    '*          parameters. Main decision maker for the pogram.      *
    '*****************************************************************
    Sub NEISCollectiveReport()
        Form1.Show()
        SetStatus("Main")
        intStatCounter = 0
        intPriceCounter = 0
        intInfoDueCounter = 0
        SetAction("Initializing Temp Folder")
        Try
            Kill(TempDocFolder & "*.*")
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try


        'Send emails to contact persons with the CONTACT_GEN_RPTS set to True
        'debug - comment out for testing contact status only
        Try
            ProcessNEISCollectiveReport()
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try

        Form1.txtAction.Text = "Exiting.."
        End
    End Sub


    '*****************************************************************
    '* Name:    SplitName                                            *
    '* Purpose: Separate out first and last name from full name      *
    '*****************************************************************
    Sub SplitName(strFullName As String, ByRef strLastName As String, ByRef strFirstName As String)
        Dim Comma As Integer
        strLastName = Trim(strFullName)
        strFirstName = ""
        Comma = InStr(strFullName, ",")
        If Comma > 1 Then
            strLastName = Trim(Left(strFullName, Comma - 1))
            strFirstName = Trim(Mid(strFullName, Comma + 1))
        End If
    End Sub

    '*****************************************************************
    '* Name:    SetAction                                            *
    '* Purpose: Update action text field                             *
    '*****************************************************************
    Sub SetAction(strAction As String)
        Form1.txtAction.Text = strAction
        Form1.txtAction.Refresh()
        Form1.Refresh()
    End Sub

    '*****************************************************************
    '* Name:    SetStatus                                            *
    '* Purpose: Update status text field                             *
    '*****************************************************************
    Sub SetStatus(strStatus As String)
        Form1.txtStatus.Text = strStatus
        Form1.txtStatus.Refresh()
        Form1.Refresh()
    End Sub


    '*****************************************************************
    '* Name:    IsValidEmailFormat                                        
    '* Purpose: Check if format for email address is valid                        
    '*****************************************************************
    Function IsValidEmailFormat(ByVal s As String) As Boolean
        Try
            Dim a As New System.Net.Mail.MailAddress(s)
        Catch
            Return False
        End Try
        Return True
    End Function

    '*****************************************************************
    '* Name:    LaunchEmail                                         
    '* Purpose: Send reports via email                          
    '*****************************************************************
    Sub LaunchEmail(ByVal Address As String, ByVal Atch1 As String)

        SetStatus("LaunchEmail")

        Dim mail As New MailMessage

        If Atch1 = "" Then
            Exit Sub
        End If

        If Not IsValidEmailFormat(Address) Then
            Exit Sub
        End If

        mail.To.Clear()
        mail.To.Add(Address)
        mail.From = New MailAddress(DefaultFromAddr)
        mail.Subject = "NEIS Collective Status Report"
        mail.Body = "This message has been sent to you by Case Systems automated e-mail system. "
        mail.Body = mail.Body & "Attached you will find one or more documents containing the Status Reports for your projects. "
        mail.Body = mail.Body & "These documents have been created in xlsx format. You will need Microsoft Excel to view these documents." & vbCrLf & vbCrLf & "Do not reply to this e-mail. "
        mail.Body = mail.Body & "Contact your Case Systems representative or send e-mail responses to " & DefaultReplyTo & "." & vbCrLf

        If Atch1 <> "" And File.Exists(Atch1) Then
            Try
                mail.Attachments.Add(New Attachment(Atch1))
            Catch ex As Exception
                ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
            End Try
        End If


        Dim Smtp As New SmtpClient
        Smtp.Host = Smtp_Server

        If File.Exists(Atch1) Then
            Try
                Smtp.Send(mail)
            Catch ex As SmtpException
                ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
            End Try
        End If

        Smtp.Dispose()
        Threading.Thread.Sleep(2500)

    End Sub


    '*****************************************************************
    '* Name:    SendStatusReports                                    
    '* Purpose: Generate file names for reports and send them out    
    '*****************************************************************
    Sub SendStatusReports(ByVal Contact As String, ByVal EmailAddress As String, ByVal ReportPath As String)
        Try

            'If they exist...
            If Not (EmailAddress = "") Then
                SetAction("Generating NEIS Collective report for " & Contact & " - " & EmailAddress)
                Threading.Thread.Sleep(2500)

                CreateExcelReport(ReportPath, Contact)

                'debug - next line commented out for testing
                LaunchEmail(EmailAddress, ReportPath)
            Else
            End If

        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try

    End Sub


    '*****************************************************************
    '* Name:    ProcessDealerStatusSummary                           *
    '* Purpose: Process the dealer status summary reports for all    *
    '*          dealers who want a companywide status report.        *
    '*****************************************************************
    Sub ProcessNEISCollectiveReport()

        Dim strSQLSelect As String
        Dim strSQLFrom As String
        Dim strSQLWhere As String
        Dim strSQLOrder As String
        Dim strSQLQuery As String
        Dim strSQLQuery2 As String = ""
        Dim Contact As String
        '    Dim Dealer As String
        Dim EmailAddress As String
        Dim ReportPath As String = ""
        Dim ReportCounter As String = 1
        Dim dtContacts As New System.Data.DataTable

        Try
            SetAction("Processing dealer summary reports.")
            SetStatus("")

            strSQLSelect = "SELECT DISTINCT CONTACT_NAME_LAST, CONTACT_NAME_FIRST, CONTACT_EMAIL"
            strSQLFrom = " FROM Contact_List"
            strSQLWhere = " WHERE ((Contact_List.CONTACT_EXCEL_DLR_STATUS_RPTS = 1)"


            strSQLWhere = strSQLWhere & ")"
            strSQLOrder = " ORDER BY Contact_List.CONTACT_NAME_LAST, Contact_List.CONTACT_NAME_FIRST;"
            strSQLQuery = strSQLSelect & strSQLFrom & strSQLWhere & strSQLOrder
            dtContacts = csiGetDataTable(strSQLQuery, cnJobLog)
            strSQLQuery = ""


            For Each drContact In dtContacts.Rows
                Contact = nz(drContact.item("CONTACT_NAME_LAST"), "") & ", " & nz(drContact.item("CONTACT_NAME_FIRST"), "")
                '     Dealer = nz(drContact.item("DEALER_IDENT"), "")
                EmailAddress = nz(drContact.Item("CONTACT_EMAIL"), "")

                ReportPath = TempDocFolder + "NEISCollectiveStatusReport" & ReportCounter.ToString & ".xlsx"


                SendStatusReports(Contact, EmailAddress, ReportPath)

                ReportCounter += 1
                If strParam(1) = "3" Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        End Try
    End Sub




    '**************************************************************************
    ' EXCEL CODE
    '**************************************************************************

    Sub CreateExcelReport(ExcelReportPath As String, DealerContact As String)
        Try
            Dim dtExcelReport As New Data.DataTable
            InitializeDataTable(dtExcelReport)
            PopulateDataTable(dtExcelReport)
            ExportToExcel(dtExcelReport, DealerContact, ExcelReportPath)
        Catch ex As Exception
            ThrowError(System.Reflection.MethodBase.GetCurrentMethod().Name & " - " & ex.ToString())
        Finally

        End Try
    End Sub

    Sub InitializeDataTable(ByRef dt As Data.DataTable)
        dt.Columns.Clear()
        dt.Columns.Add("ProjectName")
        dt.Columns.Add("Dealer")
        dt.Columns.Add("JobNumber")
        dt.Columns.Add("DeliveryWeek")

        dt.Columns.Add("ShipToAddress")
        dt.Columns.Add("PONumber")
        dt.Columns.Add("JobContact")
        dt.Columns.Add("Status")
        dt.Columns.Add("Selling")
        dt.Columns.Add("SiteContact")

        dt.Columns.Add("Trailers")
        dt.Columns.Add("AWILabels")
        dt.Columns.Add("FORRequested")
        dt.Columns.Add("InfoDueDate")
        dt.Columns.Add("InfoDueStatus")
        dt.Columns.Add("Notes")
    End Sub


    Sub PopulateDataTable(ByRef dt As Data.DataTable)
        Dim JobLogReader As SqlDataReader = Nothing

        Dim strProjectName = ""
        Dim strDealer = ""
        Dim strJobNumber = ""
        Dim strDeliveryWeek = ""

        Dim strShipToAddress = ""
        Dim strPONumber = ""
        Dim strJobContact = ""
        Dim strStatus = ""
        Dim strSelling = ""
        Dim strSiteContact = ""

        Dim strTrailers = ""
        Dim strAWILabels = ""
        Dim strFORRequested = ""
        Dim strInfoDueDate = ""
        Dim strInfoDueStatus = ""
        Dim strNotes = ""

        Dim strJobLogCmd As String = "SELECT   J.JOB_NUMBER, J.BTecJobNumber, J.JOB_STATUS, P.PRODUCTION_DATE, J.DEALER_IDENT, D.DEALER_NAME, D.DEALER_ADDRESS_1, 
                                    D.DEALER_ADDRESS_2, D.DEALER_ADDRESS_3, D.DealerCity, D.DealerState, D.DealerZip, J.DEALER_PO_NBR, J.CO_NUM,
                                    J.PROJECT_COORDINATOR, J.JOB_NAME, J.CONTACT_PERSON, J.TOTAL_SELLING, J.TOTAL_LOADING, J.BUILDING_TYPE, E.ShipDate, E.FC_DUE,
                                    E.FC_IN, J.RevisePO, J.TrailerQty, P.SPECIALS_NOTES, E.REVNRESUBMIT, J.COSIGNEE, J.SHIP_CONTACT_PHONE, 
                                    J.COMBRELEASE, J.AWILABEL, E.DWGS_OUT, E.DWG_SUB_SENT, E.FINAL_DWGS_OUT, E.FINAL_DWG_SUB_SENT, E.FirstSubDue,
								    L.[LongLeadTimeItem1], L.[LongLeadTimeItem2], L.[LongLeadTimeItem3], J.Ship_to as ShipTo,
                                    coalesce(J.SHIPPING_ADDRESS_1, '') + ' ' + coalesce(J.SHIPPING_ADDRESS_2, '') + ' ' + coalesce(J.SHIPPING_ADDRESS_3,'') as ShipTo2, 
                                    coalesce(J.ShipToCity, '') + ', ' + coalesce(J.ShipToState, '') + ' ' + coalesce(J.ShipToZip, '') as ShipTo3,
                                    J.OutsourceCTops, J.StoredChargeable, J.ProjectedDelivery
                           FROM     [JobLog].[dbo].[Job_General_Info] J
                                    INNER Join [JobLog].[dbo].[Engineering_Info] E ON E.JOB_NUMBER = J.JOB_NUMBER
                                    INNER Join [JobLog].[dbo].[Dealer_List] D ON D.DEALER_IDENT = J.DEALER_IDENT
                                    INNER Join [JobLog].[dbo].[Production_Schedule] P ON P.JOB_NUMBER = J.JOB_NUMBER
                                    INNER JOIN [JobLog].[dbo].[vwLongLeadTimeItems] L ON L.JOB_NUMBER = J.JOB_NUMBER
                           WHERE    (not J.Job_Number LIKE '____L__') AND (D.DEALER_NAME like 'NEIS Collective%') 
                                    AND (J.JOB_STATUS='ACTIVE' Or J.JOB_STATUS='ON HOLD' Or J.JOB_STATUS='PENDING' Or J.JOB_STATUS='MFG-PENDING') 
                                    AND (P.PROD_RUN='' Or P.PROD_RUN='A' Or P.PROD_RUN Is Null)
                           ORDER BY CASE WHEN J.Job_Status='ON HOLD' THEN 1 ELSE 2 END, J.JOB_STATUS Desc, P.PRODUCTION_DATE"


        Try
            Dim cmdJobLog As New SqlCommand(strJobLogCmd, cnJobLog)
            cmdJobLog.CommandTimeout = 0
            cnJobLog.Open()
            JobLogReader = cmdJobLog.ExecuteReader()
            While JobLogReader.Read()

                strProjectName = JobLogReader("JOB_NAME").ToString
                strDealer = JobLogReader("DEALER_NAME").ToString
                strJobNumber = JobLogReader("JOB_NUMBER").ToString

                Dim NeedPO As Boolean = False 'No longer looking for this so it stays false!

                strDeliveryWeek = ShipDate(NeedPO, JobLogReader("BUILDING_TYPE").ToString, JobLogReader("JOB_STATUS").ToString, JobLogReader("PRODUCTION_DATE").ToString, JobLogReader("JOB_NUMBER").ToString, JobLogReader("StoredChargeable").ToString, ez(JobLogReader("ProjectedDelivery").ToString, "1/1/0001"))


                Dim strShipTo As String = JobLogReader("ShipTo").ToString
                Dim strShipTo2 As String = JobLogReader("ShipTo2").ToString
                Dim strShipTo3 As String = JobLogReader("ShipTo3").ToString

                strShipToAddress = strShipTo + vbCrLf + strShipTo2 + vbCrLf + strShipTo3


                strPONumber = JobLogReader("DEALER_PO_NBR").ToString
                strJobContact = JobLogReader("CONTACT_PERSON").ToString
                strStatus = JobLogReader("JOB_STATUS").ToString
                strSelling = JobLogReader("TOTAL_SELLING").ToString
                strSiteContact = JobLogReader("COSIGNEE").ToString

                strTrailers = JobLogReader("TrailerQty").ToString
                strAWILabels = JobLogReader("AWILABEL").ToString
                strFORRequested = JobLogReader("COMBRELEASE").ToString
                strInfoDueDate = JobLogReader("FC_DUE").ToString
                Dim strInfoReceivedDate As String = JobLogReader("FC_IN").ToString

                If strInfoReceivedDate <> "" Then
                    strInfoDueStatus = "Received"
                End If

                Dim strLongLeadTime1 As String = JobLogReader("LongLeadTimeItem1").ToString
                Dim strLongLeadTime2 As String = JobLogReader("LongLeadTimeItem2").ToString
                Dim strLongLeadTime3 As String = JobLogReader("LongLeadTimeItem3").ToString


                If strLongLeadTime1 <> "" And strLongLeadTime2 <> "" And strLongLeadTime3 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem1").ToString + ", " + JobLogReader("LongLeadTimeItem2").ToString + ", " + JobLogReader("LongLeadTimeItem3").ToString
                ElseIf strLongLeadTime1 <> "" And strLongLeadTime2 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem1").ToString + ", " + JobLogReader("LongLeadTimeItem2").ToString
                ElseIf strLongLeadTime1 <> "" And strLongLeadTime3 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem1").ToString + ", " + JobLogReader("LongLeadTimeItem3").ToString
                ElseIf strLongLeadTime2 <> "" And strLongLeadTime3 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem2").ToString + ", " + JobLogReader("LongLeadTimeItem3").ToString
                ElseIf strLongLeadTime1 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem1").ToString
                ElseIf strLongLeadTime2 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem2").ToString
                ElseIf strLongLeadTime3 <> "" Then
                    strNotes = "Long Lead Time Items: " + JobLogReader("LongLeadTimeItem3").ToString
                End If

                'Fill data tables
                dt.Rows.Add(strProjectName, strDealer, strJobNumber, strDeliveryWeek, strShipToAddress, strPONumber, strJobContact, strStatus, strSelling, strSiteContact, strTrailers, strAWILabels, strFORRequested, strInfoDueDate, strInfoDueStatus, strNotes)

            End While

            cnJobLog.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            cnJobLog.Close()
        End Try

    End Sub


    Sub ExportToExcel(ByVal dt As System.Data.DataTable, ByVal DealerContact As String, ByVal ExcelReportPath As String)
        Dim TITLE_COL As Integer = 1
        Dim REPORT_HEADER_ROW As Integer = 1
        Dim REPORT_DATE_RANGE_ROW As Integer = 2
        Dim REPORT_DATE_GENERATED_ROW As Integer = 3

        Dim DATA_HEADER_ROW As Integer = 5

        Dim ProjectName_COL_A As Integer = 1
        Dim Dealer_COL_B As Integer = 2
        Dim JobNumber_COL_C As Integer = 3
        Dim DeliveryWeek_COL_D As Integer = 4

        Dim ShipToAddress_COL_E As Integer = 5
        Dim PONumber_COL_F As Integer = 6
        Dim JobContact_COL_G As Integer = 7
        Dim Status_COL_H As Integer = 8
        Dim Selling_COL_I As Integer = 9
        Dim SiteContact_COL_J As Integer = 10

        Dim Trailers_COL_K As Integer = 11
        Dim AWILabels_COL_L As Integer = 12
        Dim FORRequested_COL_M As Integer = 13
        Dim InfoDueDate_COL_N As Integer = 14
        Dim InfoDueStatus_COL_O As Integer = 15
        Dim Notes_COL_P As Integer = 16

        Dim dc As DataColumn
        Dim dr As DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = DATA_HEADER_ROW

        Dim _excel As New Excel.Application
        Dim wBook As Excel.Workbook = Nothing
        Dim wSheet As Excel.Worksheet = Nothing

        Dim strPrevJob As String = ""
        Dim strCurrJob As String = ""

        Dim strPrevEntryID As String = ""
        Dim strCurrEntryID As String = ""

        Try
            wBook = _excel.Workbooks.Add()
            wSheet = wBook.ActiveSheet()

            SetHeaders(_excel, DealerContact)

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    _excel.Cells(rowIndex, colIndex) = dr(dc.ColumnName)
                Next
            Next

            FormatSpreadsheet(_excel, wSheet, dt)

            ' Save the workbook to a specified file path
            wBook.SaveAs(ExcelReportPath)

            'Threading.Thread.Sleep(2000)
            '' Close the workbook and Excel application
            'wBook.Close()
            '_excel.Quit()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            '_excel.Quit()
        Finally
            ' Release Excel objects
            ReleaseComObject(wSheet)
            ReleaseComObject(wBook)

            If _excel IsNot Nothing Then
                _excel.Quit()
                ReleaseComObject(_excel)
            End If

            ' Force garbage collection to release any remaining COM objects
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Sub SetHeaders(ByRef _excel As Excel.Application, ByVal DealerContact As String)
        Dim strOptionValue As String = ""
        Dim TITLE_COL As Integer = 1
        Dim REPORT_HEADER_ROW As Integer = 1
        Dim REPORT_DATE_RANGE_ROW As Integer = 2
        Dim REPORT_DATE_GENERATED_ROW As Integer = 3

        Dim DATA_HEADER_ROW As Integer = 5

        Dim ProjectName_COL_A As Integer = 1
        Dim Dealer_COL_B As Integer = 2
        Dim JobNumber_COL_C As Integer = 3
        Dim DeliveryWeek_COL_D As Integer = 4

        Dim ShipToAddress_COL_E As Integer = 5
        Dim PONumber_COL_F As Integer = 6
        Dim JobContact_COL_G As Integer = 7
        Dim Status_COL_H As Integer = 8
        Dim Selling_COL_I As Integer = 9
        Dim SiteContact_COL_J As Integer = 10

        Dim Trailers_COL_K As Integer = 11
        Dim AWILabels_COL_L As Integer = 12
        Dim FORRequested_COL_M As Integer = 13
        Dim InfoDueDate_COL_N As Integer = 14
        Dim InfoDueStatus_COL_O As Integer = 15
        Dim Notes_COL_P As Integer = 16

        Dim i As Integer = 0

        _excel.Cells(REPORT_HEADER_ROW, TITLE_COL) = "NEIS Collective Status Report"
        _excel.Cells(REPORT_DATE_RANGE_ROW, TITLE_COL) = "Att: " & DealerContact
        _excel.Cells(REPORT_DATE_GENERATED_ROW, TITLE_COL) = "Report Generated: " & Now

        _excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A) = "Project Name"
        _excel.Cells(DATA_HEADER_ROW, Dealer_COL_B) = "Project Division"
        _excel.Cells(DATA_HEADER_ROW, JobNumber_COL_C) = "Case Number"
        _excel.Cells(DATA_HEADER_ROW, DeliveryWeek_COL_D) = "Delivery Week"

        _excel.Cells(DATA_HEADER_ROW, ShipToAddress_COL_E) = "Ship To Address"
        _excel.Cells(DATA_HEADER_ROW, PONumber_COL_F) = "PO Number"
        _excel.Cells(DATA_HEADER_ROW, JobContact_COL_G) = "Job Contact"
        _excel.Cells(DATA_HEADER_ROW, Status_COL_H) = "Status"
        _excel.Cells(DATA_HEADER_ROW, Selling_COL_I) = "Selling"
        _excel.Cells(DATA_HEADER_ROW, SiteContact_COL_J) = "Site Contact"

        _excel.Cells(DATA_HEADER_ROW, Trailers_COL_K) = "Trailers"
        _excel.Cells(DATA_HEADER_ROW, AWILabels_COL_L) = "AWI Labels"
        _excel.Cells(DATA_HEADER_ROW, FORRequested_COL_M) = "F.O.R. Requested"
        _excel.Cells(DATA_HEADER_ROW, InfoDueDate_COL_N) = "Info Due Date"
        _excel.Cells(DATA_HEADER_ROW, InfoDueStatus_COL_O) = "Info Due Status"
        _excel.Cells(DATA_HEADER_ROW, Notes_COL_P) = "Notes"

    End Sub

    Sub FormatSpreadsheet(ByRef _excel As Excel.Application, ByRef wSheet As Excel.Worksheet, ByRef dt As Data.DataTable)
        Dim TITLE_COL As Integer = 1
        Dim REPORT_HEADER_ROW As Integer = 1

        Dim DATA_HEADER_ROW As Integer = 5

        Dim ProjectName_COL_A As Integer = 1
        Dim Dealer_COL_B As Integer = 2
        Dim JobNumber_COL_C As Integer = 3
        Dim DeliveryWeek_COL_D As Integer = 4

        Dim ShipToAddress_COL_E As Integer = 5
        Dim PONumber_COL_F As Integer = 6
        Dim JobContact_COL_G As Integer = 7
        Dim Status_COL_H As Integer = 8
        Dim Selling_COL_I As Integer = 9
        Dim SiteContact_COL_J As Integer = 10

        Dim Trailers_COL_K As Integer = 11
        Dim AWILabels_COL_L As Integer = 12
        Dim FORRequested_COL_M As Integer = 13
        Dim InfoDueDate_COL_N As Integer = 14
        Dim InfoDueStatus_COL_O As Integer = 15
        Dim Notes_COL_P As Integer = 16

        Dim eRange As Excel.Range


        'Format Report Font and Size
        eRange = wSheet.Range(_excel.Cells(REPORT_HEADER_ROW, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, Notes_COL_P))
        eRange.Font.Name = "Arial"
        eRange.Font.Size = 10


        'Format Report Title
        eRange = wSheet.Range(_excel.Cells(REPORT_HEADER_ROW, TITLE_COL), _excel.Cells(REPORT_HEADER_ROW, TITLE_COL))
        eRange.Font.Size = 16
        eRange.Font.Bold = True

        'Format currency columns
        eRange = wSheet.Range(_excel.Cells(REPORT_HEADER_ROW + 1, Selling_COL_I), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, Selling_COL_I))
        For Each cell As Range In eRange.Cells
            cell.Style = "Currency"
            cell.NumberFormat = "$#,##0"
        Next

        'Format date columns
        eRange = wSheet.Range(_excel.Cells(REPORT_HEADER_ROW + 1, InfoDueDate_COL_N), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, InfoDueDate_COL_N))
        For Each cell As Range In eRange.Cells
            cell.NumberFormat = "MM/dd/yy"
        Next


        'Bold Column headers
        eRange = wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW, Notes_COL_P))
        eRange.Font.Bold = True

        ''Underline
        'eRange.Font.Underline = True

        'Horizontal Alignment Center
        eRange = wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, Notes_COL_P))
        eRange.HorizontalAlignment = Excel.Constants.xlCenter

        'Horizontal Alignment Left
        eRange = wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ShipToAddress_COL_E), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, ShipToAddress_COL_E))
        eRange.HorizontalAlignment = Excel.Constants.xlLeft

        'Vertical Alignment Center
        eRange = wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, Notes_COL_P))
        eRange.VerticalAlignment = Excel.Constants.xlCenter

        'set custom height
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW + 1, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, Notes_COL_P)).RowHeight = 51 '38.25

        'set custom width
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A)).ColumnWidth = 28
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, Dealer_COL_B), _excel.Cells(DATA_HEADER_ROW, Dealer_COL_B)).ColumnWidth = 17
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, JobNumber_COL_C), _excel.Cells(DATA_HEADER_ROW, JobNumber_COL_C)).ColumnWidth = 10
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, DeliveryWeek_COL_D), _excel.Cells(DATA_HEADER_ROW, DeliveryWeek_COL_D)).ColumnWidth = 42
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ShipToAddress_COL_E), _excel.Cells(DATA_HEADER_ROW, ShipToAddress_COL_E)).ColumnWidth = 32
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, PONumber_COL_F), _excel.Cells(DATA_HEADER_ROW, PONumber_COL_F)).ColumnWidth = 12
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, JobContact_COL_G), _excel.Cells(DATA_HEADER_ROW, JobContact_COL_G)).ColumnWidth = 20
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, Status_COL_H), _excel.Cells(DATA_HEADER_ROW, Status_COL_H)).ColumnWidth = 10
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, Selling_COL_I), _excel.Cells(DATA_HEADER_ROW, Selling_COL_I)).ColumnWidth = 10
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, SiteContact_COL_J), _excel.Cells(DATA_HEADER_ROW, SiteContact_COL_J)).ColumnWidth = 14
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, Trailers_COL_K), _excel.Cells(DATA_HEADER_ROW, Trailers_COL_K)).ColumnWidth = 8
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, AWILabels_COL_L), _excel.Cells(DATA_HEADER_ROW, AWILabels_COL_L)).ColumnWidth = 8
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, FORRequested_COL_M), _excel.Cells(DATA_HEADER_ROW, FORRequested_COL_M)).ColumnWidth = 10
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, InfoDueDate_COL_N), _excel.Cells(DATA_HEADER_ROW, InfoDueDate_COL_N)).ColumnWidth = 10
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, InfoDueStatus_COL_O), _excel.Cells(DATA_HEADER_ROW, InfoDueStatus_COL_O)).ColumnWidth = 8
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, Notes_COL_P), _excel.Cells(DATA_HEADER_ROW, Notes_COL_P)).ColumnWidth = 32


        'wrap text
        wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ProjectName_COL_A), _excel.Cells(DATA_HEADER_ROW + dt.Rows.Count, Notes_COL_P)).WrapText = True
        'wSheet.Range(_excel.Cells(DATA_HEADER_ROW, JobNumber_COL_C), _excel.Cells(DATA_HEADER_ROW, JobNumber_COL_C)).WrapText = True
        'wSheet.Range(_excel.Cells(DATA_HEADER_ROW, ShipToAddress_COL_E), _excel.Cells(DATA_HEADER_ROW, ShipToAddress_COL_E)).WrapText = True
        'wSheet.Range(_excel.Cells(DATA_HEADER_ROW, Notes_COL_P), _excel.Cells(DATA_HEADER_ROW, Notes_COL_P)).WrapText = True

    End Sub


    '*****************************************************************
    '* Name:    ShipDate                                             *
    '* Purpose: Determine status of the ship date                    *
    '*****************************************************************
    '* 02/24/17:    We are using field checks for all info due so we
    '*              should only use that field for missing info. -MCB
    '*****************************************************************
    Function ShipDate(ByVal NeedPO As Boolean, ByVal Bldg As String, ByVal Stat As String, ByVal ProdDate As Date, ByVal JobNumber As String, Optional ByVal BuildAndStore As String = "0", Optional ByVal ProjectedDelivery As Date = Nothing) As String
        Dim Result As String = ""
        Dim NewDate As Date
        Dim NewDate2 As Date
        Dim MissingInfo As Boolean
        Dim strDateFormat As String = "MM/dd/yyyy"
        Dim strDeliveryCommit As String = ""
        Dim strRRRDeliveryCommit As String = ""
        Dim strDelivery As String = ""

        Dim blnFCReqd As Boolean = False
        Dim strFCDue As String = ""
        Dim strFCIn As String = ""
        Dim dtEngineering As New System.Data.DataTable

        dtEngineering = csiGetDataTable("SELECT coalesce(DELEVERY_COMMIT, DeliveryDate, '') as RRR_Delivery, 
                                                coalesce(DeliveryDate, '') as Delivery, coalesce(DELEVERY_COMMIT, '') as DeliveryCommit,                                                
                                                FC_REQD, FC_DUE, FC_IN  
                                        FROM    [JobLog].[dbo].[Engineering_Info]
                                        WHERE   JOB_NUMBER = '" & JobNumber & "'", cnJobLog)

        For Each dr As DataRow In dtEngineering.Rows
            strRRRDeliveryCommit = nz(dr.Item("RRR_Delivery"), "")
            strDeliveryCommit = nz(dr.Item("DeliveryCommit"), "")
            strDelivery = nz(dr.Item("Delivery"), "")
            blnFCReqd = nz(dr.Item("FC_REQD"))
            strFCDue = nz(dr.Item("FC_DUE"), "")
            strFCIn = nz(dr.Item("FC_IN"), "")
        Next

        If NeedPO Then
            Result = "Revised Purchase Order Required!"
        Else
            If BuildAndStore = "True" Then
                NewDate = DateValue(ProjectedDelivery)
            Else
                If IsDate(strRRRDeliveryCommit) Then
                    NewDate = DateValue(strRRRDeliveryCommit)
                Else
                    NewDate = DateAdd("d", 11, ProdDate)
                End If
                If IsDate(strDelivery) Then
                    NewDate2 = DateValue(strDelivery)
                Else
                    NewDate2 = DateAdd("d", 11, ProdDate)
                End If

                If IsDate(strDeliveryCommit) Then
                    strDeliveryCommit = DateValue(strDeliveryCommit).ToString(strDateFormat)
                Else
                    strDeliveryCommit = ""
                End If

                Dim strtest = ""
            End If
            If Stat = "ON HOLD" Then
                Result = "ON HOLD! Contact your Case Systems AM immediately."
            Else
                If Stat = "PENDING" Then
                    Result = "IN PROCESS. Pending material."
                Else
                    MissingInfo = InfoLate(blnFCReqd, nz(strFCDue, ""), nz(strFCIn, ""))

                    If MissingInfo Then
                        Result = "MISSING INFO! Contact your Case Systems AM immediately."
                    Else
                        If BuildAndStore = "True" Then
                            Result = "PROJECTED DELIVERY WEEK: " & NewDate.ToString(strDateFormat)
                        Else
                            If Bldg = "RRR" Or Bldg.StartsWith("RRR") Then
                                Result = "TENTATIVE SHIP DATE: " & NewDate.ToString(strDateFormat)
                            Else
                                'ORIGINAL DELIVERY WEEK: 
                                If IsDate(strDeliveryCommit) Then
                                    Result = "TENTATIVE DELIVERY WEEK: " & NewDate2.ToString(strDateFormat) & vbCrLf & "COMMITTED DELIVERY WEEK: " & strDeliveryCommit
                                Else
                                    Result = "TENTATIVE DELIVERY WEEK: " & NewDate2.ToString(strDateFormat)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return Result
    End Function


    '*****************************************************************
    '* Name:    InfoLate                                             *
    '* Purpose: Determine whether or not information is late         *
    '*****************************************************************
    Function InfoLate(ByVal blnRequired As Boolean, ByVal strDue As String, ByVal strRecorded As String) As Boolean
        If blnRequired Then
            If IsDate(strDue) Then
                If IsDate(strRecorded) Then
                    Return False
                Else
                    If DateValue(strDue) < DateValue(Now) Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            ' Handle exception, if necessary
        Finally
            GC.Collect()
        End Try
    End Sub
End Module
