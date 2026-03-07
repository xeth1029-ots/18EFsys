Partial Class SD_09_011_R
    Inherits AuthBasePage

    'Dim conn As SqlConnection
    'Dim sql As String
    'Dim objreader As SqlDataReader
    Dim objconn As SqlConnection
    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then

            If sm.UserInfo.RID = "A" Then
                DistID = TIMS.Get_DistID(DistID)
            Else
                DistID = TIMS.Get_DistID(DistID)
                DistID.Enabled = False
            End If

            years = TIMS.GetSyear(years)
            months.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, 0))
            For i As Integer = 1 To 12
                months.Items.Add(i)
            Next
            DistID.SelectedValue = sm.UserInfo.DistID
            Common.SetListItem(years, Now.Year)
            Common.SetListItem(months, Now.Month)
        End If

        Button1.Attributes("onclick") = "javascript:return print();"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim SDate As Date = CDate(years.SelectedValue & "/" & months.SelectedValue & "/1")

        Dim EDate As Date = CDate(years.SelectedValue & "/" & (months.SelectedValue + 1) & "/1")

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_09_011_R", "DistID=" & DistID.SelectedValue & "&SDate=" & SDate & "&EDate=" & EDate & "")

    End Sub
End Class
