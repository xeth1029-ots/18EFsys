Partial Class OB_01_ch
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9) '☆

        If Not IsPostBack Then
            ddl_years = TIMS.GetSyear(ddl_years, Year(Now) - 1, Year(Now) + 3, True)
            ddl_TPlanID = TIMS.Get_TPlan(ddl_TPlanID)
        End If
    End Sub

    Sub years()
        ddl_years.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        For i As Int16 = 0 To 4
            ddl_years.Items.Add(New ListItem((Year(Now) + i).ToString, (Year(Now) + i).ToString))
        Next
    End Sub

    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        Dim sql As String
        Dim dt As DataTable
        sql = ""
        If Request("sort") = "0" Then
            sql += " select ot.Tsn, TISn, ot.Years, op.PlanName, ot.TenderCName, convert(varchar,ot.TenderSDate,111) as TenderSDate " & vbCrLf
            sql += " From OB_Tender ot "
            sql += " JOIN OB_TenderItem oti on oti.tsn=ot.tsn "
            sql += " LEFT JOIN OB_Plan op on op.PlanSN=ot.PlanSN " & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
        ElseIf Request("sort") = "1" Then
            sql += " Select ot.Tsn, ot.Years, op.PlanName, ot.TenderCName, convert(varchar,ot.TenderSDate,111) as TenderSDate " & vbCrLf
            sql += " From OB_Tender ot "
            sql += " LEFT JOIN OB_Plan op on op.PlanSN=ot.PlanSN " & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
        End If

        If sm.UserInfo.DistID <> "000" Then
            sql += " and ot.DistID='" & sm.UserInfo.DistID & "'"
        End If

        If Len(ddl_years.SelectedValue) > 0 Then
            sql += " and ot.years=" & ddl_years.SelectedValue
        End If

        If Len(ddl_TPlanID.SelectedValue) > 0 Then
            sql += " and ot.TPlanID='" & ddl_TPlanID.SelectedValue & "'"
        End If

        PlanName.Text = PlanName.Text.Trim
        If Len(PlanName.Text) > 0 Then
            sql += " and op.PlanName LIKE '%" & PlanName.Text & "%'"
        End If

        If TIMS.Get_SQLRecordCount(sql) > 0 Then
            dt = DbAccess.GetDataTable(sql)
            msg.Visible = False
            Panel_View.Visible = True
            dg_Sch.DataSource = dt
            dg_Sch.DataBind()
        Else
            msg.Visible = True
            Panel_View.Visible = False

        End If
    End Sub

    Private Sub send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click
        If Request("tsn") <> "" Then
            Dim sql As String
            Dim dr As DataRow
            sql = "select tsn,TenderCName from OB_Tender where tsn=" & Request("tsn")
            dr = DbAccess.GetOneRow(sql)

            Common.RespWrite(Me, "<script language=javascript>")
            Common.RespWrite(Me, "function returnNum(){")
            Common.RespWrite(Me, "window.opener.document.form1.txt_Name.value='" & dr("TenderCName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.txt_tsn.value='" & dr("tsn") & "';")
            Common.RespWrite(Me, "window.close();")
            Common.RespWrite(Me, "}")
            Common.RespWrite(Me, "returnNum();")
            Common.RespWrite(Me, "</script>")
        Else
            Common.MessageBox(Me, "請先勾選標案!")
        End If
    End Sub
End Class
