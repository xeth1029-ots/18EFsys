Partial Class SYS_04_013
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
        PageControler1.PageDataGrid = Me.DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False

            If Convert.ToString(Request("s")) = "1" Then
                Call Search1()
            End If
        End If

        ''檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    'If sm.UserInfo.RoleID <> 0 Then
        '    'End If
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '       'Dim FunDr As DataRow = FunDrArray(0)
        '        btnSearch.Enabled = False
        '        btnInsert.Enabled = False
        '        If FunDr("Sech") = "1" Then btnSearch.Enabled = True
        '        If FunDr("Adds") = "1" Then btnInsert.Enabled = True
        '    End If
        'End If
        ''檢查帳號的功能權限-----------------------------------End

    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        Me.ViewState("subject") = ""
        Me.ViewState("schOSDate1") = ""
        Me.ViewState("schOSDate2") = ""

        If schSubject.Text <> "" Then
            Me.ViewState("subject") = schSubject.Text.Replace("'", "''")
        End If

        If schOSDate1.Text <> "" Then
            Me.ViewState("schOSDate1") = Common.FormatDate(schOSDate1.Text)
        End If

        If schOSDate2.Text <> "" Then
            Me.ViewState("schOSDate2") = Common.FormatDate(schOSDate2.Text)
        End If

        Call Search1()

    End Sub

    Sub Search1()
        Dim Sql As String = ""
        Sql += " SELECT * " & vbCrLf
        Sql += " FROM Auth_AccCal  " & vbCrLf
        Sql += " WHERE Account ='" & sm.UserInfo.UserID & "'" & vbCrLf

        If Me.ViewState("schOSDate1") <> "" Then
            Sql += " AND OSDate >=" & TIMS.To_date(Me.ViewState("schOSDate1")) & vbCrLf
        End If
        If Me.ViewState("schOSDate2") <> "" Then
            Sql += " AND OSDate <=" & TIMS.To_date(Me.ViewState("schOSDate2")) & vbCrLf
        End If
        If Me.ViewState("subject") <> "" Then
            Sql += " AND subject like '%'+'" & Me.ViewState("subject") & "'+'%'" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(Sql, objconn)

        msg.Text = "查無資料"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            'DataGrid1.DataSource = dt
            'DataGrid1.DataBind()
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "Edit"
                '修改
                TIMS.Utl_Redirect1(Me, "SYS_04_013_add.aspx?ID=" & Request("ID") & "&CalID=" & e.CommandArgument)

            Case "Del"
                Dim sCalID As String = ""
                sCalID = e.CommandArgument

                '刪除
                Dim sql As String = ""
                sql = ""
                sql += " select * from Auth_AccCal WHERE calID=" & sCalID & vbCrLf
                sql += " and Account='" & sm.UserInfo.UserID & "'" & vbCrLf
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count = 1 Then
                    '刪除成功
                    sql = ""
                    sql += " DELETE Auth_AccCal WHERE calID=" & sCalID & vbCrLf
                    sql += " and Account='" & sm.UserInfo.UserID & "'" & vbCrLf
                    Call DbAccess.ExecuteNonQuery(sql, objconn)

                    Common.MessageBox(Me, "刪除成功!!")
                    Call Search1()
                End If

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                '序號
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex

                Dim alsubject As Label = e.Item.FindControl("alsubject")
                Dim alOSDate As Label = e.Item.FindControl("alOSDate")
                Dim alOFDate As Label = e.Item.FindControl("alOFDate")
                Dim BtnEdit As LinkButton = e.Item.FindControl("BtnEdit")
                Dim BtnDel As LinkButton = e.Item.FindControl("BtnDel")

                alsubject.Text = Convert.ToString(drv("subject"))
                alOSDate.Text = Convert.ToString(CDate(drv("OSDate")).ToString("yyyy/MM/dd"))
                alOFDate.Text = Convert.ToString(CDate(drv("OFDate")).ToString("yyyy/MM/dd"))
                BtnDel.CommandArgument = drv("CalID")
                BtnDel.Attributes("onClick") = "return confirm('這樣會刪除此筆資料,\n確定要繼續刪除?');"

                BtnEdit.CommandArgument = drv("CalID")

        End Select
    End Sub

    '新增
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        TIMS.Utl_Redirect1(Me, "SYS_04_013_add.aspx?ID=" & Request("ID"))
    End Sub

End Class
