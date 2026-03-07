Partial Class SD_06_003_R
    Inherits AuthBasePage

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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DataGridTable.Visible = False
            btnGetOneClass_Click(sender, e)
            For j As Integer = 1 To 12 '月
                PMonth.Items.Add(New ListItem(j, j))
            Next
            For i As Integer = 1 To 31 '日
                PDay.Items.Add(New ListItem(i, i))
            Next
            Years.Text = CInt(Year(Now)) - 1911 '預設年
            PMonth.SelectedValue = Month(Now) '預設月
            PDay.SelectedValue = Day(Now)   '預設日
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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
        Button1.Attributes("onclick") = "return CheckSearch();"
        Button5.Attributes("onclick") = "openClass('../02/SD_02_ch.aspx?RID='+document.getElementById('RIDValue').value);"

        'Dim SMpath As String = ""
        'SMpath = ReportQuery.GetReportQueryPath

        '列印加保申報表
        PrintA.Attributes("onclick") = "CheckPrint(0);"
        '列印退保申報表
        PrintB.Attributes("onclick") = "CheckPrint(1);"
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyCheck1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim MyCheck2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
                Dim OtherSubsidy As DropDownList = e.Item.FindControl("OtherSubsidy")

                MyCheck1.Value = drv("SOCID")
                MyCheck2.Value = drv("SOCID")
                If IsDBNull(drv("ApplyInsurance")) Then
                    MyCheck1.Disabled = True
                End If
                If IsDBNull(drv("DropoutInsurance")) Then
                    MyCheck2.Disabled = True
                End If

                If Convert.ToString(drv("OtherSubsidy")) = "Y" Then
                    OtherSubsidy.SelectedValue = "Y"
                Else
                    OtherSubsidy.SelectedValue = "N"
                    If Convert.ToString(drv("OtherSubsidy")) = "" Then
                        PrintA.Disabled = True '加保
                        PrintB.Disabled = True '退保
                    Else
                        PrintA.Disabled = False '加保
                        PrintB.Disabled = False '退保
                    End If
                End If

        End Select
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable

        sql = "" & vbCrLf
        sql += "SELECT a.OCID,a.SOCID,substr(a.StudentID,-2) StudentID,b.IDNO,b.Name,b.Birthday,c.InsureSalary,c.ApplyInsurance,c.DropoutInsurance,c.OtherSubsidy " & vbCrLf
        sql += "FROM Class_StudentsOfClass  a " & vbCrLf
        sql += "JOIN Stud_StudentInfo b ON a.SID=b.SID AND a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        sql += "JOIN Stud_Insurance c ON a.SOCID=c.SOCID " & vbCrLf
        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '排除在職者補助身分
            sql += " AND ( dbo.NVL(a.WorkSuppIdent,' ') !='Y') " & vbCrLf
        End If
        sql += "Order By a.OCID,substr(a.StudentID,-2)"

        dt = DbAccess.GetDataTable(sql, objconn)
        msg.Text = "查無申請資料!!"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    '存檔
    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別")
            Exit Sub
        End If

        Dim updataRow As Integer = 0
        Dim conn As SqlConnection
        conn = DbAccess.GetConnection()
        Dim da As SqlDataAdapter = Nothing '= TIMS.GetOneDA()
        Dim dt As DataTable
        Dim dr As DataRow
        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " select * " & vbCrLf
        Sql += " from Stud_Insurance p where 1=1" & vbCrLf
        Sql += " and exists (" & vbCrLf
        Sql += " 	select 'x' from Class_StudentsOfClass cs" & vbCrLf
        Sql += " 	where cs.ocid ='" & OCIDValue1.Value & "'  " & vbCrLf
        Sql += " 	and cs.socid =p.socid " & vbCrLf
        Sql += " )" & vbCrLf
        dt = DbAccess.GetDataTable(Sql, da, conn)
        If dt.Rows.Count > 0 Then
            For Each item As DataGridItem In DataGrid1.Items
                Dim OtherSubsidy As DropDownList = item.FindControl("OtherSubsidy")
                Dim MyCheck1 As HtmlInputCheckBox = item.FindControl("Checkbox1") 'MyCheck1.Value'SOCID  

                If dt.Select("SOCID='" & MyCheck1.Value & "'").Length > 0 Then
                    dr = dt.Select("SOCID='" & MyCheck1.Value & "'")(0)
                    dr("OtherSubsidy") = Convert.DBNull
                    If OtherSubsidy.SelectedValue <> "" Then
                        dr("OtherSubsidy") = OtherSubsidy.SelectedValue
                    End If
                    updataRow += 1
                End If

            Next
        End If
        DbAccess.UpdateDataTable(dt, da)
        TIMS.CloseDbConn(conn)

        If updataRow = 0 Then
            Common.MessageBox(Me, "查無申請資料，請確認申請資料!!")
        Else
            Common.MessageBox(Me, "儲存成功")
        End If
        PrintA.Disabled = False '加保
        PrintB.Disabled = False '退保
    End Sub

    '查班
    Private Sub btnGetOneClass_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGetOneClass.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

End Class
