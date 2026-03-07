Partial Class TC_01_017
    Inherits AuthBasePage

    'SELECT * FROM KEY_ORGTYPE1 B --on a.OrgKind1=b.OrgTypeID1
    'SELECT * FROM V_ORGKIND1
    'SELECT ORGTYPEID1,ORGTYPE FROM VIEW_ORGTYPE1
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            bt_search.Enabled = False
            Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃與充電起飛計畫(在職)!!")
            Exit Sub
        End If
        'If Not TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    bt_search.Enabled = False
        '    Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃!!")
        '    Exit Sub
        'End If

        '0:署(局)
        '1:分署(中心)
        '2:委訓
        'Select Case sm.UserInfo.LID
        '    Case "0", "1"
        '    Case Else
        '        Common.MessageBox(Me, "開放給職訓局及中心階層以上者，權限不足，無法使用!!")
        '        Exit Sub
        'End Select
        '0:超級使用者
        '1:系統管理者
        '2:一級以上
        '3:一級
        '4:二級
        '5:承辦人
        '99:一般使用者
        'Select Case sm.UserInfo.RoleID
        '    Case "0", "1"
        '    Case Else
        '        Common.MessageBox(Me, "非系統管理者以上權限者，權限不足，無法使用!!")
        '        Exit Sub
        'End Select

        '分頁設定 Start
        PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = dg1
        '分頁設定 End

        If Not IsPostBack Then
            dg1.Visible = False
            PageControler1.Visible = False

            PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn)

            dl_typeid2.Items.Clear()
            dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
            dl_typeid2.SelectedIndex = 0
            dl_typeid2.Enabled = False
        End If
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            bt_search.Enabled = False
            Common.MessageBox(Me, "此功能目前僅開放給產業人才投資計劃與充電起飛計畫(在職)!!")
            Exit Sub
        End If

        tb_comidno.Text = Trim(tb_comidno.Text)

        Call loadData()
    End Sub

    Private Sub loadData()
        Dim sql As String = ""
        sql += " select distinct d.tplanid,a.orgid,a.orgname,a.comidno" & vbCrLf
        sql += " ,a.orgkind1" & vbCrLf
        sql += " ,b.typeid1" & vbCrLf
        'sql += " ,dbo.DECODE6(b.typeid1,'1','勞工在職進修計畫','2','勞工團體辦理勞工在職進修計畫','') plantype " & vbCrLf
        sql += " ,v1.Name plantype " & vbCrLf
        sql += " ,b.typeid2,b.typeid2name" & vbCrLf
        sql += " ,b.typeid2 + '-' + b.typeid2name as orgtype" & vbCrLf
        sql += " FROM ORG_ORGINFO a" & vbCrLf
        sql += " JOIN AUTH_RELSHIP c on a.orgid=c.orgid" & vbCrLf
        sql += " JOIN ID_PLAN d on c.planid=d.planid" & vbCrLf
        sql += " LEFT JOIN Key_OrgType1 b on a.OrgKind1=b.OrgTypeID1" & vbCrLf
        sql += " LEFT JOIN v_OrgKind1 v1 on v1.sort= b.typeid1" & vbCrLf
        sql += " where d.tplanid=@tplanid" & vbCrLf
        Select Case PlanPoint.SelectedValue
            Case "1", "2"
                sql += " and b.typeid1=@typeid1  " & vbCrLf
        End Select
        If dl_typeid2.SelectedIndex <> 0 Then
            sql += " and b.typeid2=@typeid2  " & vbCrLf
        End If
        If tb_orgname.Text <> "" Then
            sql += " and a.orgname like '%' + @orgname + '%'  " & vbCrLf
        End If
        If Trim(tb_comidno.Text) <> "" Then
            sql += " and a.comidno=@comidno "
        End If
        sql += " order by a.orgid "

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("tplanid", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
            Select Case PlanPoint.SelectedValue
                Case "1", "2"
                    .Parameters.Add("@typeid1", SqlDbType.Int).Value = CInt(PlanPoint.SelectedValue)
            End Select
            If dl_typeid2.SelectedIndex <> 0 Then
                .Parameters.Add("@typeid2", SqlDbType.VarChar).Value = dl_typeid2.SelectedValue.ToString
            End If
            If tb_orgname.Text <> "" Then
                .Parameters.Add("@orgname", SqlDbType.VarChar).Value = tb_orgname.Text
            End If
            If Trim(tb_comidno.Text) <> "" Then
                .Parameters.Add("@comidno", SqlDbType.VarChar).Value = Trim(tb_comidno.Text)
            End If
            dt.Load(.ExecuteReader())
        End With

        msg.Text = "查無資料!"
        PageControler1.Visible = False
        dg1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            PageControler1.Visible = True
            dg1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.Sort = "orgname"
            PageControler1.ControlerLoad()
        End If

    End Sub

    'Private Sub get_ddl_typeid2(ByVal TypeID1 As String)
    '    If TypeID1 = "" Or TypeID1 = "0" Then
    '        Exit Sub
    '    End If

    '    Dim sql As String = ""
    '    sql += " select TypeID1,TypeID2,  TypeID2 + '-' + TypeID2Name TypeID2Name  " & vbCrLf
    '    sql += " from Key_OrgType1" & vbCrLf
    '    sql += " where TypeID1=@TypeID1" & vbCrLf
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim dt As New DataTable
    '    Dim oCmd As New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("@TypeID1", SqlDbType.Int).Value = CInt(TypeID1)
    '        dt.Load(.ExecuteReader())
    '    End With
    '    If dt.Rows.Count > 0 Then
    '        dl_typeid2.Items.Clear()

    '        dl_typeid2.DataSource = dt
    '        dl_typeid2.DataValueField = "TypeID2"
    '        dl_typeid2.DataTextField = "TypeID2Name"
    '        dl_typeid2.DataBind()
    '        dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '        dl_typeid2.SelectedIndex = 0
    '        dl_typeid2.Enabled = True
    '    Else
    '        dl_typeid2.Items.Clear()
    '        dl_typeid2.Items.Insert(0, New ListItem("==請選擇==", ""))
    '        dl_typeid2.SelectedIndex = 0
    '        dl_typeid2.Enabled = False
    '    End If

    'End Sub

    Private Sub dg1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + dg1.PageSize * dg1.CurrentPageIndex
            Dim drv As DataRowView = e.Item.DataItem
            Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
            lbtEdit.CommandArgument = Convert.ToString(drv("orgid"))
        End If
    End Sub

    Private Sub dg1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Select Case e.CommandName
            Case "edit"
                Me.Page.RegisterStartupScript("edit", "<script language='javascript'>location.href='TC_01_017_add.aspx?orgid=" & e.CommandArgument & "';</script>")
        End Select
    End Sub

    Function IsInt(ByVal chkstr As String) As Boolean  '判斷是否為正整數
        Dim rst As Boolean = False
        Try
            If Int32.Parse(chkstr) = chkstr AndAlso Int32.Parse(chkstr) > 0 Then
                rst = True
            End If
        Catch ex As Exception
        End Try
        Return rst
    End Function

    Private Sub PlanPoint_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlanPoint.SelectedIndexChanged
        Dim v_PlanPoint As String = TIMS.GetListValue(PlanPoint)
        Call TIMS.GET_DDL_TYPEID12(dl_typeid2, objconn, v_PlanPoint)
    End Sub
End Class
