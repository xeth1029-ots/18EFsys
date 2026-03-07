Partial Class SD_05_027
    Inherits AuthBasePage


    'Dim conn As SqlConnection = DbAccess.GetConnection

    '2011 功能按鈕權限控管參數 ---------------------Start
    'Dim FunDr As DataRow = Nothing
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印
    '2011 功能按鈕權限控管參數 ---------------------End

#Region "Sub"
    '查詢
    Private Sub sSearch1()
        '取出計畫種類
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        '取出訓練計畫名稱
        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim pms1 As New Hashtable From {{"planid", sm.UserInfo.PlanID}}
        Dim sql As String = ""
        sql &= " select c.years,a.planid,b.orgid,a.ocid,c.distid,c.years,b.orgname"
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " from class_classinfo a "
        sql &= " join org_orginfo b on b.comidno=a.comidno "
        sql &= " join id_plan c on c.planid=a.planid "
        sql &= " join key_plan d on d.tplanid=c.tplanid "
        sql &= " where a.issuccess='Y' and a.planid=@planid "

        If sm.UserInfo.RID = "A" Then
            pms1.Add("Relship", Relship)
            sql &= " and a.rid in (select rid from auth_relship where relship like @Relship+'%') "
        Else
            If PlanKind = "2" Then
                pms1.Add("Relship", Relship) '2:委外
                sql &= " and a.rid in (select rid from auth_relship where relship like @Relship+'%') "
            End If

            If sm.UserInfo.LID = 0 Or sm.UserInfo.LID = 1 Then
                pms1.Add("DistID", sm.UserInfo.DistID)
                sql &= " and c.distid=@DistID "
            End If
        End If

        '通俗職類
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        jobvalue.Value = TIMS.ClearSQM(jobvalue.Value)
        trainvalue.Value = TIMS.ClearSQM(trainvalue.Value)
        If txtcjob_name.Text <> "" AndAlso cjobValue.Value <> "" Then
            pms1.Add("cjob_unkey", cjobValue.Value)
            sql &= " and a.cjob_unkey =@cjob_unkey"
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '訓練職類
            If jobvalue.Value <> "" Then
                pms1.Add("jobvalue", jobvalue.Value)
                sql &= " and (a.tmid=@jobvalue or a.tmid in ( select tmid from key_traintype where parent in ( "
                sql &= " select tmid from key_traintype where parent in ( "
                sql &= " select tmid from key_traintype where busid ='G') and tmid =@jobvalue ))) "
            End If
        Else
            '通俗職類
            If trainvalue.Value <> "" Then
                pms1.Add("trainvalue", trainvalue.Value)
                sql &= "and a.tmid=@trainvalue "
            End If
        End If

        '班級名稱
        txtSchClass.Text = TIMS.ClearSQM(txtSchClass.Text)
        If txtSchClass.Text <> "" Then
            pms1.Add("classcname", txtSchClass.Text)
            sql &= "and classcname like '%'+@classcname+'%' " & vbCrLf
        End If

        '期別
        txtCyclType.Text = TIMS.ClearSQM(txtCyclType.Text)
        If txtCyclType.Text <> "" Then
            pms1.Add("cycltype", (Int(txtCyclType.Text)).ToString.PadLeft(2, "0"))
            sql &= "and a.cycltype=@cycltype " & vbCrLf
        End If

        sql &= "order by c.years"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)

        If TIMS.dtNODATA(dt) Then
            DataGrid1.Visible = False
            PageControler1.Visible = False
            labMsg.Visible = True
            Exit Sub
        End If

        DataGrid1.Visible = True
        PageControler1.Visible = True
        labMsg.Visible = False

        DataGrid1.DataSource = dt 'ds.Tables(0)
        DataGrid1.DataBind()

        PageControler1.PageDataTable = dt 'ds.Tables(0)
        PageControler1.ControlerLoad()

    End Sub

    '載入資料
    Private Sub loadData()
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim dr As DataRow = Nothing

        Select Case Convert.ToString(ViewState("ProcKind"))
            Case "org" '單位設定
                Dim Sql As String = ""
                Sql &= " select a.leaveid,a.name,dbo.NVL(b.minuspoint,0) minuspoint "
                Sql &= " from key_leave a "
                Sql &= " left join org_leave b on b.leaveid=a.leaveid and b.planid=@planid and b.orgid=@orgid "
                Sql &= " order by a.leaveid"
                With sda
                    .SelectCommand = New SqlCommand(Sql, objconn)
                    .SelectCommand.Parameters.Clear()
                    .SelectCommand.Parameters.Add("planid", SqlDbType.VarChar).Value = sm.UserInfo.PlanID 'Convert.ToString(sm.UserInfo.PlanID).PadLeft(2, "0")
                    .SelectCommand.Parameters.Add("orgid", SqlDbType.VarChar).Value = hidOrgID.Value
                    .Fill(ds)
                End With

            Case "class" '班級個別設定
                Dim Sql As String = ""
                Sql &= " select a.leaveid,a.name,dbo.NVL(b.minuspoint,0) minuspoint "
                Sql &= " from key_leave a"
                Sql &= " left join class_leave b on b.leaveid=a.leaveid and b.ocid=@ocid "
                Sql &= " order by a.leaveid"

                With sda
                    .SelectCommand = New SqlCommand(Sql, objconn)
                    .SelectCommand.Parameters.Clear()
                    .SelectCommand.Parameters.Add("ocid", SqlDbType.VarChar).Value = hidOCID.Value
                    .Fill(ds)
                End With
        End Select

        If ds.Tables(0).Rows.Count > 0 Then
            Datagrid2.DataSource = ds.Tables(0)
            Datagrid2.DataBind()
        End If
    End Sub

    '清除維護頁資料
    Private Sub clsValue()
        hidOCID.Value = ""

        labYear.Text = ""
        labOrg.Text = ""
        labClass.Text = ""

        hidOrgID.Value = ""
    End Sub
#End Region

#Region "Function"
    '判斷是否有設定內容
    Private Function chkUsed(ByVal strOCID As String) As Boolean
        'Dim sda As New SqlDataAdapter 'Dim ds As New DataSet
        Dim bolRtn As Boolean = False
        Dim dt As New DataTable
        Dim Sql As String = "select ocid from class_leave where ocid=@ocid"
        Using SCMD As New SqlCommand(Sql, objconn)
            With SCMD
                .Parameters.Add("ocid", SqlDbType.VarChar).Value = strOCID
                dt.Load(.ExecuteReader())
            End With
        End Using
        If TIMS.dtHaveDATA(dt) Then bolRtn = True
        Return bolRtn
    End Function
#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值

        If Not Me.IsPostBack Then
            'planid.Value = sm.UserInfo.PlanID

            RIDValue.Value = sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
            org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), sm.UserInfo.Years)

            btnSave.Attributes.Add("onclick", "return chkSave();")
            tbEdit.Visible = False
            PageControler1.Visible = False
        End If

        TIMS.ShowHistoryRID(Me, historyrid, "HistoryList2", "RIDValue", "center")
        If historyrid.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        PageControler1.PageDataGrid = DataGrid1

        '2011 功能按鈕權限控管--Start
        Dim strSechObjID As String = "" '查詢按鈕物件ID
        Dim strAddsObjID As String = "" '維護按鈕物件ID
        Dim strPrntObjID As String = "" '列印按鈕物件ID

        strSechObjID = btnSch.ClientID
        strAddsObjID = btnOrgSet.ClientID & "," & btnSave.ClientID

        TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
        '2011 功能按鈕權限控管--End
    End Sub

    Private Sub btnSch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSch.Click
        txtCyclType.Text = TIMS.ClearSQM(txtCyclType.Text)
        If txtCyclType.Text <> "" Then
            If Not IsNumeric(txtCyclType.Text) Then
                Common.MessageBox(Me, "期別需輸入數字型態!!")
                Exit Sub
            End If
        End If

        Call sSearch1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labDYear As Label = e.Item.FindControl("labDYear")
                Dim labDOrgName As Label = e.Item.FindControl("labDOrgName")
                Dim labDClassName As Label = e.Item.FindControl("labDClassName")
                Dim labStatus As Label = e.Item.FindControl("labStatus")

                Dim btnEdit As Button = e.Item.FindControl("btnEdit")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                labDYear.Text = Convert.ToString(drv("years"))
                labDOrgName.Text = Convert.ToString(drv("orgname"))
                labDClassName.Text = Convert.ToString(drv("classname"))

                If chkUsed(Convert.ToString(drv("ocid"))) Then
                    labStatus.Text = "有"
                Else
                    labStatus.Text = "無"
                End If

                btnEdit.CommandArgument = Convert.ToString(drv("ocid"))


                '2011 功能按鈕權限控管--Start
                'If blnCanAdds = True Then '維護
                'Else
                '    btnEdit.Visible = False
                'End If
                '2011 功能按鈕權限控管--End
        End Select
    End Sub

    Public Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edt"
                Dim labDYear As Label = Nothing
                Dim labDOrgName As Label = Nothing
                Dim labDClassName As Label = Nothing

                ViewState("ProcKind") = "class"
                tbSch.Visible = False
                tbEdit.Visible = True
                trClass.Visible = True
                clsValue()

                hidOCID.Value = e.CommandArgument

                labDYear = DataGrid1.Items(e.Item.ItemIndex).FindControl("labDYear")
                labDOrgName = DataGrid1.Items(e.Item.ItemIndex).FindControl("labDOrgName")
                labDClassName = DataGrid1.Items(e.Item.ItemIndex).FindControl("labDClassName")

                labYear.Text = labDYear.Text
                labOrg.Text = labDOrgName.Text
                labClass.Text = labDClassName.Text

                loadData()
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidLeaveID As HtmlInputHidden = e.Item.FindControl("hidLeaveID")
                Dim labName As Label = e.Item.FindControl("labName")
                Dim txtItem As TextBox = e.Item.FindControl("txtItem")
                Dim dMinusPoint As Double = 0

                dMinusPoint = Convert.ToDouble(drv("minuspoint"))

                hidLeaveID.Value = Convert.ToString(drv("leaveid"))
                labName.Text = Convert.ToString(drv("name"))
                txtItem.Text = String.Format("{0:##0.00}", dMinusPoint)
        End Select
    End Sub

    '儲存
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim sda As New SqlDataAdapter
        Dim trans As SqlTransaction = Nothing
        Dim bolCnt As Boolean = False
        Dim hidLeaveID As HtmlInputHidden = Nothing
        Dim txtItem As TextBox = Nothing

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(conn)
        trans = DbAccess.BeginTrans(conn)

        Try
            'conn.Open()
            Select Case Convert.ToString(ViewState("ProcKind"))
                Case "org"
                    '刪除操行分數設定
                    Dim Sql_d2 As String = "delete org_leave where planid=@planid and orgid=@orgid "
                    With sda
                        .DeleteCommand = New SqlCommand(Sql_d2, conn, trans)
                        .DeleteCommand.Parameters.Clear()
                        .DeleteCommand.Parameters.Add("planid", SqlDbType.VarChar).Value = sm.UserInfo.PlanID 'Convert.ToString(sm.UserInfo.TPlanID).PadLeft(2, "0")
                        .DeleteCommand.Parameters.Add("orgid", SqlDbType.VarChar).Value = hidOrgID.Value
                        .DeleteCommand.ExecuteNonQuery()
                    End With

                    '新增操行分數設定
                    Dim Sql_i2 As String = ""
                    Sql_i2 = " insert into org_leave(planid,orgid,leaveid,minuspoint,modifyacct,modifydate) "
                    Sql_i2 &= " values(@planid,@orgid,@leaveid,@minuspoint,@modifyacct,getdate())"
                    sda.InsertCommand = New SqlCommand(Sql_i2, conn, trans)

                    For Each dgi As DataGridItem In Datagrid2.Items
                        hidLeaveID = dgi.FindControl("hidLeaveID")
                        txtItem = dgi.FindControl("txtItem")

                        With sda
                            .InsertCommand.Parameters.Clear()
                            .InsertCommand.Parameters.Add("planid", SqlDbType.VarChar).Value = sm.UserInfo.PlanID 'Convert.ToString(sm.UserInfo.TPlanID).PadLeft(2, "0")
                            .InsertCommand.Parameters.Add("orgid", SqlDbType.VarChar).Value = hidOrgID.Value
                            .InsertCommand.Parameters.Add("leaveid", SqlDbType.VarChar).Value = hidLeaveID.Value
                            .InsertCommand.Parameters.Add("minuspoint", SqlDbType.Float).Value = txtItem.Text
                            .InsertCommand.Parameters.Add("modifyacct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                            .InsertCommand.ExecuteNonQuery()
                        End With
                    Next
                Case "class"
                    '刪除操行分數設定
                    Dim Sql_d As String = " delete class_leave where ocid=@ocid"
                    With sda
                        .DeleteCommand = New SqlCommand(Sql_d, conn, trans)
                        .DeleteCommand.Parameters.Clear()
                        .DeleteCommand.Parameters.Add("ocid", SqlDbType.VarChar).Value = hidOCID.Value
                        .DeleteCommand.ExecuteNonQuery()
                    End With

                    '新增操行分數設定
                    Dim Sql_i As String = ""
                    Sql_i = " insert into class_leave(ocid,leaveid,minuspoint,modifyacct,modifydate) "
                    Sql_i &= " values(@ocid,@leaveid,@minuspoint,@modifyacct,getdate())"
                    sda.InsertCommand = New SqlCommand(Sql_i, conn, trans)


                    For Each dgi As DataGridItem In Datagrid2.Items
                        hidLeaveID = dgi.FindControl("hidLeaveID")
                        txtItem = dgi.FindControl("txtItem")

                        With sda
                            .InsertCommand.Parameters.Clear()
                            .InsertCommand.Parameters.Add("ocid", SqlDbType.VarChar).Value = hidOCID.Value
                            .InsertCommand.Parameters.Add("leaveid", SqlDbType.VarChar).Value = hidLeaveID.Value
                            .InsertCommand.Parameters.Add("minuspoint", SqlDbType.Float).Value = txtItem.Text
                            .InsertCommand.Parameters.Add("modifyacct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                            .InsertCommand.ExecuteNonQuery()
                        End With
                    Next

            End Select

            DbAccess.CommitTrans(trans)
            bolCnt = True
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Finally
            '    If Not sda Is Nothing Then sda.Dispose()
            '    If Not trans Is Nothing Then trans.Dispose()
            '    If Not conn Is Nothing Then conn.Close()
        End Try
        Call TIMS.CloseDbConn(conn)

        If bolCnt = True Then
            Common.MessageBox(Me, "儲存成功!")
            btnBack_Click(sender, e)
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        tbSch.Visible = True
        tbEdit.Visible = False
        Call sSearch1()
        '2011 功能按鈕權限控管--Start
        'If blnCanSech = True Then '查詢
        '    search()
        'End If
        '2011 功能按鈕權限控管--End
    End Sub

    '單位操行分數設定
    Private Sub btnOrgSet_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOrgSet.Click
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim dr As DataRow = Nothing

        tbSch.Visible = False
        tbEdit.Visible = True
        trClass.Visible = False
        clsValue()

        Call TIMS.OpenDbConn(objconn)

        ViewState("ProcKind") = "org"

        '查詢OrgID,Name
        Dim sql As String = ""
        sql &= " select a.orgid,b.orgname "
        sql &= " from auth_relship a "
        sql &= " join org_orginfo b on b.orgid=a.orgid "
        sql &= " where a.rid=@rid "
        With sda
            .SelectCommand = New SqlCommand(sql, objconn)
            .SelectCommand.Parameters.Clear()
            .SelectCommand.Parameters.Add("rid", SqlDbType.VarChar).Value = RIDValue.Value
            .Fill(ds, "org")
        End With

        If ds.Tables("org").Rows.Count > 0 Then
            dr = ds.Tables("org").Rows(0)

            hidOrgID.Value = Convert.ToString(dr("orgid"))
            labYear.Text = Convert.ToString(sm.UserInfo.Years)
            labOrg.Text = Convert.ToString(dr("orgname"))
        End If

        'Try
        'conn.Open()
        'Call TIMS.OpenDbConn(conn)

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        'If Not ds Is Nothing Then ds.Dispose()
        'If Not sda Is Nothing Then sda.Dispose()
        'If Not conn Is Nothing Then conn.Close()
        'End Try

        loadData()
    End Sub
End Class
