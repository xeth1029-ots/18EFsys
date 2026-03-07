Partial Class SD_04_013_R
    Inherits AuthBasePage

    'SD_04_013_R
    'SD_04_013_R_ds
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'objconn = DbAccess.GetConnection()
        Button1.Attributes("onclick") = "javascript:return ReportPrint();"
        If Not IsPostBack Then
            Years = TIMS.GetSyear(Years)
            Common.SetListItem(Years, Now.Year)
            For i As Integer = 1 To 12
                Months.Items.Add(New ListItem(i & "月份", i))
            Next
            Months.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Common.SetListItem(Months, Now.Month)

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Allday_TR.Style("display") = "none" '全期的隱藏
            Button3_Click(sender, e)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Printtype.Attributes("onclick") = "ChangeMode();"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim RID As String = ""
        Dim sYear As String = ""
        Dim STdate As String = ""
        Dim FTdate As String = ""
        Dim start_month As String = ""
        Dim end_month As String = ""

        If Years.SelectedValue <> "" And Months.SelectedValue <> "" Then
            start_month = Years.SelectedValue & "/" & Months.SelectedValue & "/1"
            end_month = Common.FormatDate(DateAdd(DateInterval.Month, 1, CDate(start_month)))
            sYear = (Years.SelectedValue - 1911)
        End If

        If s_date.Text <> "" And e_date.Text <> "" Then
            STdate = s_date.Text
            FTdate = e_date.Text
        End If

        If Me.RIDValue.Value = "" Then
            RID = sm.UserInfo.RID
        Else
            RID = Me.RIDValue.Value
        End If

        Dim sMyValue As String = ""
        sMyValue &= "RID=" & RID
        sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        sMyValue &= "&SDate=" & start_month
        sMyValue &= "&EDate=" & end_month
        sMyValue &= "&Years=" & sYear
        sMyValue &= "&Months=" & Me.Months.SelectedValue
        sMyValue &= "&OCID=" & Me.OCIDValue1.Value
        sMyValue &= "&STdate=" & STdate
        sMyValue &= "&FTdate=" & FTdate
        sMyValue &= "&CJOB_UNKEY=" & cjobValue.Value

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_04_013_R", sMyValue)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Printtype_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Printtype.SelectedIndexChanged
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        'Months_TR.Style("display") = "inline"
        Months_TR.Style("display") = ""
        Allday_TR.Style("display") = "none"
        s_date.Text = ""
        e_date.Text = ""
        Select Case Printtype.SelectedIndex
            Case 0
                'Months_TR.Style("display") = "inline"
                Months_TR.Style("display") = ""
                Allday_TR.Style("display") = "none"
                s_date.Text = ""
                e_date.Text = ""
            Case 1
                'Dim SQL As String
                If OCIDValue1.Value = "" Then
                    Common.MessageBox(Me.Page, "尚未選擇職類/班別")
                    Printtype.SelectedValue = "0"
                Else
                    Dim sql As String = "SELECT STDATE,FTDATE FROM CLASS_CLASSINFO WHERE OCID = " & OCIDValue1.Value & " "
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If Not dr Is Nothing Then
                        s_date.Text = TIMS.Cdate3(dr("STDate"))
                        e_date.Text = TIMS.Cdate3(dr("FTDate"))
                    End If
                    'Allday_TR.Style("display") = "inline"
                    Allday_TR.Style("display") = ""
                    Months_TR.Style("display") = "none"
                    Years.SelectedIndex = -1 '.SelectedValue = ""
                    Months.SelectedIndex = -1 '
                    'Years.SelectedValue = ""
                    'Months.SelectedValue = ""
                End If

        End Select
    End Sub

End Class
