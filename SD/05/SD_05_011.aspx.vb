Partial Class SD_05_011
    Inherits AuthBasePage

    'Dim objreader As SqlDataReader
    'Dim FunDr As DataRow

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

        '分頁設定 Start
        PageControler1.PageDataGrid = Stud_DG
        '分頁設定 End

        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            If FunDr("Adds") = "1" Then
        '                Button1.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '            If FunDr("Sech") = "1" Then
        '                Button2.Enabled = True
        '            Else
        '                Button2.Enabled = False
        '            End If
        '            If FunDr("Del") = "1" Then
        '                check_del.Value = "1"
        '            Else
        '                check_del.Value = "0"
        '            End If
        '            If FunDr("Mod") = "1" Then
        '                check_mod.Value = "1"
        '            Else
        '                check_mod.Value = "0"
        '            End If
        '        End If
        '    End If
        'End If

        If Not Page.IsPostBack Then
            Plan_List = TIMS.Get_TPlan(Plan_List)
            Common.SetListItem(Plan_List, sm.UserInfo.TPlanID)

            DistrictList = TIMS.Get_DistID(DistrictList, Nothing, objconn)
        End If

        If sm.UserInfo.RID <> "A" Then
            Common.SetListItem(DistrictList, sm.UserInfo.DistID)
            DistrictList.Enabled = False
            Plan_List.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '請選擇年度
        If YearList.SelectedValue = "" OrElse YearList.SelectedValue = "0" Then
            Common.MessageBox(Me, "請選擇年度")
            Exit Sub
        End If

        Call GetDataGrid()
    End Sub

    Sub GetDataGrid()


        Dim sqlstr As String = ""
        Name.Text = TIMS.ClearSQM(Name.Text)
        SID.Text = TIMS.ClearSQM(SID.Text)
        If YearList.SelectedValue = "2004" Then
            'History_StudentInfo93
            sqlstr = "" & vbCrLf
            sqlstr &= " Select Serial As StdID,DistName,PlanName,TrinUnit,ClassName,Name,Sex,IDNO,Ident,TPlanID,DistID,'History_StudentInfo93' source" & vbCrLf
            sqlstr &= " FROM History_StudentInfo93" & vbCrLf
            sqlstr &= " WHERE 1=1" & vbCrLf
            'sqlstr &= SearchStr2
            If Plan_List.SelectedIndex <> 0 AndAlso Plan_List.SelectedValue <> "" Then
                sqlstr &= " and TPlanID='" & Plan_List.SelectedValue & "'"
            End If
            If Me.SID.Text <> "" Then
                sqlstr &= " and IDNO='" & SID.Text & "'"
            End If
            If Name.Text <> "" Then
                sqlstr &= " and Name like '%" & Name.Text & "%'"
            End If
            If DistrictList.SelectedIndex <> 0 Then
                sqlstr &= " and DistID='" & DistrictList.SelectedValue & "'"
            End If
        Else
            'StdAll
            sqlstr = "" & vbCrLf
            sqlstr &= " Select StdID,DistName,PlanName,TrinUnit,ClassName,Name,Sex,SID As IDNO,Ident,TPlanID,DistID,'StdAll' source" & vbCrLf
            sqlstr &= " FROM StdAll" & vbCrLf
            sqlstr &= " WHERE 1=1" & vbCrLf
            sqlstr &= " and Years='" & YearList.SelectedValue & "'" & vbCrLf
            If Plan_List.SelectedIndex <> 0 Then
                sqlstr &= " and TPlanID='" & Plan_List.SelectedValue & "'"
            End If
            If SID.Text <> "" Then
                sqlstr &= " and SID='" & SID.Text & "'"
            End If
            If Me.Name.Text <> "" Then
                sqlstr &= " and Name like '%" & Name.Text & "%'"
            End If
            If DistrictList.SelectedIndex <> 0 Then
                sqlstr &= " and DistID='" & DistrictList.SelectedValue & "'"
            End If
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        Stud_DG.Visible = False
        Panel2.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            Panel2.Visible = True
            Stud_DG.Visible = True
            msg.Text = ""

            PageControler1.PageDataTable = dt '.SqlString = sqlstr
            PageControler1.PrimaryKey = "StdID"
            PageControler1.Sort = "DistID,TPlanID"
            PageControler1.ControlerLoad()
        End If
        'If TIMS.Get_SQLRecordCount(sqlstr, objconn) > 0 Then
        '    PageControler1.SqlString = sqlstr
        '    PageControler1.PrimaryKey = "StdID"
        '    PageControler1.Sort = "DistID,TPlanID"
        '    PageControler1.ControlerLoad()
        'End If

    End Sub

#Region "NO USE"
    'dtOrgInfo = DbAccess.GetDataTable(sqlstr)

    'If dtOrgInfo.Rows.Count = 0 Then
    '    Stud_DG.Visible = False
    '    Panel2.Visible = False
    '    msg.Text = "查無資料!!"

    'Else
    '    Panel2.Visible = True
    '    Stud_DG.Visible = True
    '    msg.Text = ""

    '    dtOrgInfo.DefaultView.Sort = "DistID,TPlanID"
    '    Stud_DG.Visible = True
    '    Stud_DG.DataSource = dtOrgInfo
    '    Stud_DG.DataKeyField = "StdID"
    '    Stud_DG.DataBind()

    '    '分頁用-   Start
    '    DataGridPage1.MyRecord = TIMS.Get_SQLRecordCount(sqlstr)
    '    DataGridPage1.MySqlStr = sqlstr
    '    DataGridPage1.MyPrimaryKey = "StdID"
    '    DataGridPage1.MySort = "DistID,TPlanID"
    '    DataGridPage1.FirstTime()
    '    '分頁用-   End
    'End If

    'Dim sql As String
    'sql = TIMS.Get_SQLPAGE(sqlstr, 1, Stud_DG.PageSize, "Stdid", "DistID,TPlanID")
    'If dt.Rows.Count = 0 Then
    '    Stud_DG.Visible = False
    '    Panel2.Visible = False
    '    msg.Text = "查無資料!!"
    'Else
    '    Panel2.Visible = True
    '    Stud_DG.Visible = True
    '    msg.Text = ""
    '    dt.DefaultView.Sort = "DistID,TPlanID"
    '    Stud_DG.Visible = True
    '    Stud_DG.DataSource = dt
    '    Stud_DG.DataKeyField = "Stdid"
    '    Stud_DG.DataBind()

    '    '分頁用-   Start
    '    DataGridPage1.MyRecord = TIMS.Get_SQLRecordCount(sqlstr)
    '    DataGridPage1.MyDataTable = dt
    '    'DataGridPage1.MySqlStr = sqlstr
    '    'DataGridPage1.MyPrimaryKey = "Stdid"
    '    DataGridPage1.MySort = "DistID,TPlanID"
    '    DataGridPage1.FirstTime()
    '    '分頁用-   End
    'End If
#End Region

    Private Sub Stud_DG_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Stud_DG.ItemDataBound
        'dr = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim S_SEX As String = ""
                Select Case Convert.ToString(drv("Sex"))
                    Case "M"
                        S_SEX = "男"
                    Case "F"
                        S_SEX = "女"
                End Select
                e.Item.Cells(5).Text = S_SEX

                Dim edit_but As Button = e.Item.FindControl("edit_but") '修改 
                edit_but.Text = "修改"
                '登入帳號為職訓局
                If sm.UserInfo.DistID = "000" Then edit_but.Text = "查看"

                Dim del_but As Button = e.Item.FindControl("del_but") '刪除 
                Select Case Convert.ToString(drv("source"))
                    Case "StdAll"
                        del_but.Attributes("onclick") = "javascript:return confirm('此動作會刪除此筆歷史資料，是否確定刪除?');"
                        '不提供刪除
                        del_but.Visible = False
                    Case Else
                        '不提供刪除
                        del_but.Visible = False
                End Select

        End Select

    End Sub

    Private Sub Stud_DG_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Stud_DG.ItemCommand
        If e.CommandName = "edit" Then
            TIMS.Utl_Redirect1(Me, "SD_05_011_add.aspx?ProcessType=Update&stdid=" & e.Item.Cells(10).Text & "&source=" & e.Item.Cells(9).Text & "&ID=" & Request("ID"))
        ElseIf e.CommandName = "del" Then
            Dim sql As String
            If e.Item.Cells(9).Text = "StdAll" Then
                sql = "delete StdAll where Stdid='" & e.Item.Cells(10).Text & "'"
            Else
                sql = "delete History_StudentInfo93  where Serial='" & e.Item.Cells(10).Text & "'"
            End If

            Dim cmd As New SqlCommand(sql, objconn)
            TIMS.OpenDbConn(objconn)
            cmd.ExecuteNonQuery()

            msg.Text = "刪除成功!"
            Panel2.Visible = False
            Stud_DG.DataBind()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Response.Redirect("SD_05_011_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "")
        '20100208 按新增時代查詢之 身分證號嗎 & 姓名
        TIMS.Utl_Redirect1(Me, "SD_05_011_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&StuID=" & SID.Text & "&StuName=" & Name.Text & "")
    End Sub
End Class
