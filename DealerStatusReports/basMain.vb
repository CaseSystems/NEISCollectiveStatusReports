Imports System.Data.SqlClient
Module basMain
    Public cnConfig As New SqlConnection("Data Source='SQL1';Initial Catalog='Config';Integrated Security=sspi")
    Public cnJobLog As New SqlConnection("Data Source='SQL1';Initial Catalog='JobLog';Integrated Security=sspi")
    ' Public cnBTec As New SqlConnection("Data Source='SQL2';Initial Catalog='BTecQuote';Integrated Security=sspi")
    Public cnM1_Live As New SqlConnection("Data Source='ERP-M1';Initial Catalog='M1_CS';User Id='sa'; Password='n4819p+'")
    '   Public cnM1_Pilot As New SqlConnection("Data Source='ERP';Initial Catalog='M1_CP';User Id='sa'; Password='n4819p+';")
End Module
