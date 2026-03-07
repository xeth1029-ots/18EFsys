Partial Class SD_05_009_detail_add
    Inherits AuthBasePage

    Dim TechID As String
    Dim ocid As String
    Dim Years As String
    Dim Month As String

    'Dim objreader As SqlDataReader

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        TechID = Request("TechID")
        ocid = Request("OCID")
        Years = Request("Year")
        Month = Request("Month")
        Dim teacherName As String
        Dim teacherName_sql As String = "select b.TeachCName from Teach_PayHour a  join Teach_TeacherInfo b  on a.TechID=b.TechID where a.OCID='" & ocid & "' and a.TechID='" & TechID & "' "
        teacherName = Convert.ToString(DbAccess.ExecuteScalar(teacherName_sql, objconn))
        Me.Label1.Text = teacherName
        '檢查日期格式-Melody(2005/3/)
        start_date.Attributes("onchange") = "check_date();"
    End Sub

    Private Sub save_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles save_Button.Click
        Dim daPayHour As SqlDataAdapter = Nothing
        Dim sqldr As DataRow = Nothing
        Dim strMessage_date As String = ""
        Dim strMessage1 As String = "講師鐘點費新增成功!!"

        If Not Page.IsValid Then
            Exit Sub
        End If

        Dim dtPayHour As DataTable = DbAccess.GetDataTable("select * from Teach_PayHour where OCID='" & ocid & "' and TechID='" & TechID & "' and TeachDate='" & start_date.Text & "' and UnitPrice= '" & TB_Price.Text & "'", daPayHour, objconn)
        If dtPayHour.Rows.Count = 0 Then '沒有重覆,新增

            'newdatestr = Convert.ToDateTime(objdate.Text)
            If Month <> 10 And Month <> 11 And Month <> 12 Then
                Month = "0" & Month
            End If
            If DateTime.Parse(start_date.Text).ToString("yyyy/MM") <> Years + "/" + Month Then
                strMessage_date = "日期請選擇" & Years + "/" + Month & "區間"
                Common.MessageBox(Page, strMessage_date)
                Exit Sub
            End If

            sqldr = dtPayHour.NewRow
            dtPayHour.Rows.Add(sqldr)
            sqldr("ModifyAcct") = sm.UserInfo.UserID
            sqldr("ModifyDate") = Now()
            sqldr("OCID") = ocid
            sqldr("TechID") = TechID
            sqldr("TeachDate") = start_date.Text
            sqldr("UnitPrice") = TB_Price.Text
            sqldr("UnitHour") = TB_hour.Text
            DbAccess.UpdateDataTable(dtPayHour, daPayHour)

            Common.AddClientScript(Page, "alert('" & strMessage1 & "');")
            'Common.AddClientScript(Page, "opener.location.href='SD_05_009_List.aspx?OCID=" & ocid & "&TechID=" & TechID & "&Years=" & Years & "&Month=" & Month & "';")
            Common.AddClientScript(Page, "opener.document.form1.submit();")
            Common.AddClientScript(Page, "window.close();")
        Else '選擇合併

            If DateTime.Parse(start_date.Text).ToString("yyyy/MM") <> Years + "/" + Month Then
                strMessage_date = "日期請選擇" & Years + "/" + Month & "區間"
                Common.MessageBox(Page, strMessage_date)
                Exit Sub
            End If

            sqldr = dtPayHour.Rows(0)
            sqldr("UnitHour") = TB_hour.Text + sqldr("UnitHour")
            sqldr("ModifyAcct") = sm.UserInfo.UserID
            sqldr("ModifyDate") = Now()

            DbAccess.UpdateDataTable(dtPayHour, daPayHour)
            Common.AddClientScript(Page, "alert('" & strMessage1 & "');")
            'Common.AddClientScript(Page, "opener.document.form1.submit();='SD_05_009_List.aspx?OCID=" & ocid & "&TechID=" & TechID & "&Years=" & Years & "&Month=" & Month & "';")
            Common.AddClientScript(Page, "opener.document.form1.submit();")
            Common.AddClientScript(Page, "window.close();")
        End If

    End Sub

    Private Sub CustomValidator1_ServerValidate(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles CustomValidator1.ServerValidate
        If Request("BypassCheck") = "1" Then
            args.IsValid = True
            Exit Sub
        End If

        Dim sqlstr_count As String = "select * from Teach_PayHour where OCID='" & ocid & "' and TechID='" & TechID & "' and TeachDate='" & start_date.Text & "' and UnitPrice= '" & TB_Price.Text & "'"
        If DbAccess.GetCount(sqlstr_count) > 0 Then
            args.IsValid = False

            Page.RegisterHiddenField("BypassCheck", "1")

        Else
            args.IsValid = True
        End If
    End Sub
End Class
