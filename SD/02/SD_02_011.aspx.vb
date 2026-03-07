Partial Class SD_02_011
    Inherits AuthBasePage

#Region "Sub"

    '查詢
    Private Sub search()
        '取出計畫種類
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        '取出訓練計畫名稱
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
        If RelShip = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        'Dim flag_use_relship As Boolean = False
        'If sm.UserInfo.RID = "A" Then flag_use_relship = True

        Dim sql As String = ""
        sql &= " SELECT e.woid " & vbCrLf
        sql &= " ,a.planid " & vbCrLf
        sql &= " ,b.orgid " & vbCrLf
        sql &= " ,a.ocid " & vbCrLf
        sql &= " ,c.distid " & vbCrLf
        sql &= " ,c.years " & vbCrLf
        sql &= " ,b.orgname " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME ,a.CYCLTYPE) classname" & vbCrLf
        sql &= " ,e.writeresult " & vbCrLf
        sql &= " ,e.oralresult " & vbCrLf
        sql &= " FROM class_classinfo a " & vbCrLf
        sql &= " JOIN org_orginfo b ON b.comidno = a.comidno " & vbCrLf
        sql &= " JOIN id_plan c ON c.planid = a.planid " & vbCrLf
        sql &= " JOIN key_plan d ON d.tplanid = c.tplanid " & vbCrLf
        sql &= " JOIN auth_relship r ON r.rid = a.rid" & vbCrLf
        sql &= " LEFT JOIN org_writeoral e ON e.ocid = a.ocid " & vbCrLf
        sql &= " WHERE a.issuccess = 'Y' " & vbCrLf
        sql &= " AND a.planid = @planid " & vbCrLf
        If ddlSchYear.SelectedValue <> "" Then sql &= " AND c.years = @years " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND r.relship LIKE '" & RelShip & "%' " & vbCrLf
        Else
            If PlanKind = "2" Then sql &= " AND r.relship LIKE '" & RelShip & "%' " & vbCrLf

            Select Case sm.UserInfo.LID
                Case "1", "2"
                    sql &= " AND c.distid = '" & sm.UserInfo.DistID & "' " & vbCrLf
            End Select
        End If

        '通俗職類
        If txtCJOB_NAME.Text <> "" Then sql &= " AND a.cjob_unkey = " & cjobValue.Value & " " & vbCrLf

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '訓練職類
            If jobvalue.Value <> "" Then
                sql &= " AND (a.tmid = " & jobvalue.Value & " " & vbCrLf
                sql &= "      OR a.tmid IN ( " & vbCrLf
                sql &= "         SELECT tmid FROM key_traintype WHERE parent IN ( " & vbCrLf
                sql &= "         SELECT tmid FROM key_traintype WHERE parent in ( " & vbCrLf
                sql &= "         SELECT tmid FROM key_traintype WHERE busid = 'G') " & vbCrLf
                sql &= " AND tmid = " & jobvalue.Value & "))) " & vbCrLf
            End If
        Else
            '通俗職類
            If trainValue.Value <> "" Then sql &= " AND a.tmid = '" & trainValue.Value & "' " & vbCrLf
        End If

        '班級名稱
        txtschclass.Text = TIMS.ClearSQM(txtschclass.Text)
        If txtschclass.Text <> "" Then sql &= " AND a.classcname LIKE '%" & txtschclass.Text & "%' " & vbCrLf

        '期別
        txtcycltype.Text = TIMS.ClearSQM(txtcycltype.Text)
        If txtcycltype.Text <> "" Then
            txtcycltype.Text = Val(txtcycltype.Text)
            txtcycltype.Text = Int(txtcycltype.Text).ToString.PadLeft(2, "0")
        End If
        If txtcycltype.Text <> "" Then sql &= " AND a.cycltype = '" & txtcycltype.Text & "' " & vbCrLf
        sql &= " ORDER BY c.years " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("@planid", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
            If ddlSchYear.SelectedValue <> "" Then .Parameters.Add("@years", SqlDbType.VarChar).Value = ddlSchYear.SelectedValue
            dt.Load(.ExecuteReader())
        End With

        DataGrid1.Visible = False
        pagecontroler1.Visible = False
        labMsg.Visible = True
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            pagecontroler1.Visible = True
            labMsg.Visible = False
            'DataGrid1.DataSource = dt
            'DataGrid1.DataBind()
            pagecontroler1.PageDataTable = dt
            pagecontroler1.ControlerLoad()
        End If
    End Sub

    '代入資料
    Private Sub loadData(ByVal strFlag As String)
        Dim sql As String = ""

        Select Case strFlag
            Case "add", "edit"
                sql = ""
                sql &= " SELECT c.years ,e.name distname ,b.orgname"
                sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                sql &= " ,f.writeresult ,f.oralresult "
                sql &= " FROM class_classinfo a "
                sql &= " JOIN org_orginfo b ON b.comidno = a.comidno "
                sql &= " JOIN id_plan c ON c.planid = a.planid "
                sql &= " JOIN key_plan d ON d.tplanid = c.tplanid "
                sql &= " JOIN id_district e ON e.distid = c.distid "
                sql &= " LEFT JOIN org_writeoral f ON f.ocid = a.ocid "
                sql &= " WHERE 1=1 "
                If strFlag = "add" Then
                    sql &= " AND a.planid = @planid "
                    sql &= " AND b.orgid = @orgid "
                    sql &= " AND a.ocid = @ocid "
                End If
                If strFlag = "edit" Then sql &= " AND f.woid = @woid "
            Case "org"
                sql = ""
                sql &= " SELECT a.years ,b.name distname"
                sql &= " ,(SELECT ORGNAME FROM ORG_ORGINFO WHERE orgid=@orgid) orgname"
                sql &= " ,c.writeresult ,c.oralresult "
                sql &= " FROM id_plan a "
                sql &= " JOIN id_district b ON b.distid = a.distid "
                sql &= " LEFT JOIN (SELECT * FROM org_writeoral WHERE ocid IS NULL AND orgid = @orgid) c ON c.planid = a.planid "
                sql &= " WHERE 1=1 "
                sql &= " AND a.years = @years "
                sql &= " AND a.planid = @planid "
                sql &= " AND a.distid = @distid "
        End Select
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            Select Case strFlag
                Case "add", "edit"
                    If strFlag = "add" Then
                        .Parameters.Add("@planid", SqlDbType.VarChar).Value = If(hidPlanID.Value <> "", Val(hidPlanID.Value), 0)
                        .Parameters.Add("@orgid", SqlDbType.VarChar).Value = If(hidOrgID.Value <> "", Val(hidOrgID.Value), 0) 'hidOrgID.Value
                        .Parameters.Add("@ocid", SqlDbType.VarChar).Value = hidOCID.Value
                    End If
                    If strFlag = "edit" Then .Parameters.Add("@woid", SqlDbType.VarChar).Value = hidWOID.Value
                Case "org"
                    .Parameters.Add("@orgid", SqlDbType.VarChar).Value = If(hidOrgID.Value <> "", Val(hidOrgID.Value), 0)
                    .Parameters.Add("@years", SqlDbType.VarChar).Value = hidYears.Value
                    .Parameters.Add("@planid", SqlDbType.VarChar).Value = If(hidPlanID.Value <> "", Val(hidPlanID.Value), 0)
                    .Parameters.Add("@distid", SqlDbType.VarChar).Value = hidDistID.Value
            End Select
            dt.Load(.ExecuteReader())
        End With

        Dim dr As DataRow = Nothing
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            labYear.Text = Convert.ToString(dr("years"))
            labDist.Text = Convert.ToString(dr("distname"))
            labOrg.Text = Convert.ToString(dr("orgname"))
            If strFlag <> "org" Then
                trclass.Visible = True
                labClass.Text = Convert.ToString(dr("classname"))
            Else
                trclass.Visible = False
                labClass.Text = ""
            End If
            txtWrite.Text = Convert.ToString(dr("writeresult"))
            txtOral.Text = Convert.ToString(dr("oralresult"))
        End If
    End Sub

    '清除維護頁資料
    Private Sub clsValue()
        hidWOID.Value = ""
        hidYears.Value = ""
        hidDistID.Value = ""
        hidPlanID.Value = ""
        hidOrgID.Value = ""
        hidOCID.Value = ""
        txtWrite.Text = ""
        txtOral.Text = ""
    End Sub

#End Region

#Region "Function"

    '判斷是否可使用
    'Private Function chkUsed() As Boolean
    '    Dim sda As New SqlDataAdapter
    '    Dim ds As New DataSet
    '    Dim bolRtn As Boolean = False

    '    Try
    '        conn.Open()
    '        sql = " SELECT gvid FROM Sys_GlobalVar WHERE tplanid = @tplanid AND gvid = '23' AND itemvar1 = 'Y' "
    '        With sda
    '            .SelectCommand = New SqlCommand(sql, conn)
    '            .SelectCommand.Parameters.Clear()
    '            .SelectCommand.Parameters.Add("@tplanid", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
    '            .Fill(ds)
    '        End With
    '        If ds.Tables(0).Rows.Count > 0 Then bolRtn = True
    '        conn.Close()
    '        If Not sda Is Nothing Then sda.Dispose()
    '        If Not ds Is Nothing Then ds.Dispose()
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '    End Try

    '    Return bolRtn
    'End Function

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        pagecontroler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            CCreate1()
        End If

        Dim v_ddlSchYear As String = TIMS.GetListValue(ddlSchYear)
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), v_ddlSchYear)

        TIMS.ShowHistoryRID(Me, historyrid, "HistoryList2", "RIDValue", "center")
        If historyrid.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Sub CCreate1()
        ddlSchYear = TIMS.GetSyear(ddlSchYear)
        ddlSchYear.Items.RemoveAt(0)
        Common.SetListItem(ddlSchYear, sm.UserInfo.Years)
        'tplanid.Value = sm.UserInfo.TPlanID
        RIDValue.Value = sm.UserInfo.RID
        center.Text = sm.UserInfo.OrgName

        btnSave.Attributes.Add("onclick", "return chkSave();")
        tbEdit.Visible = False
        pagecontroler1.Visible = False

        '年度計畫未開放設定,則不能使用功能
        Dim sChkUsed As String = TIMS.GetGlobalVar(Me, "23", "1", objconn)
        If sChkUsed = "" Then
            btnOrgSet.Enabled = False
            btnsch.Enabled = False
            Common.MessageBox(Me, "該分署計畫尚未設定-系統參數!")
            Return
        End If
        If sChkUsed <> "Y" Then
            btnOrgSet.Enabled = False
            btnsch.Enabled = False
            Common.MessageBox(Me, "該分署計畫未開放單位成績計算比例設定!")
            Return
        End If

    End Sub

    Private Sub btnOrgSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOrgSet.Click

        tbsch.Visible = False
        tbEdit.Visible = True
        Call clsValue()

        '查詢OrgID,Name
        'Dim v_orgid As String = sm.UserInfo.OrgID
        'Dim v_planid As String = sm.UserInfo.PlanID
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        hidYears.Value = sm.UserInfo.Years
        hidDistID.Value = sm.UserInfo.DistID
        hidPlanID.Value = sm.UserInfo.PlanID
        hidOrgID.Value = sm.UserInfo.OrgID ' dtOrg.Rows(0)("orgid")
        Dim parms As New Hashtable From {{"RID", RIDValue.Value}}
        Dim sql As String = ""
        sql &= " SELECT A.rid ,a.orgid ,b.orgname" & vbCrLf
        sql &= " ,ISNULL(ip.DISTNAME,d.NAME) DISTNAME" & vbCrLf
        sql &= " ,ip.planid,ip.years,ip.distid" & vbCrLf
        sql &= " FROM AUTH_RELSHIP a" & vbCrLf
        sql &= " LEFT JOIN VIEW_PLAN ip on ip.PlanID=a.PlanID" & vbCrLf
        sql &= " join ID_DISTRICT d on d.DISTID=a.DISTID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO b ON b.orgid = a.orgid" & vbCrLf
        sql &= " WHERE A.RID=@RID" & vbCrLf

        Dim dtOrg As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dtOrg.Rows.Count > 0 Then
            Dim drOrg As DataRow = Nothing
            drOrg = dtOrg.Rows(0)
            If (Convert.ToString(drOrg("years")) <> "") Then hidYears.Value = Convert.ToString(drOrg("years"))
            If (Convert.ToString(drOrg("distid")) <> "") Then hidDistID.Value = Convert.ToString(drOrg("distid"))
            If (Convert.ToString(drOrg("planid")) <> "") Then hidPlanID.Value = Convert.ToString(drOrg("planid"))
            If (Convert.ToString(drOrg("orgid")) <> "") Then hidOrgID.Value = Convert.ToString(drOrg("orgid"))
        End If

        '查詢機構設定
        Dim sql2 As String = ""
        sql2 = " SELECT woid FROM org_writeoral WHERE planid = @planid AND orgid = @orgid AND ocid IS NULL "
        Dim parms2 As New Hashtable
        parms2.Clear()
        parms2.Add("planid", If(hidPlanID.Value <> "", Val(hidPlanID.Value), 0))
        parms2.Add("orgid", If(hidOrgID.Value <> "", Val(hidOrgID.Value), 0))
        Dim dtChk As DataTable = Nothing
        dtChk = DbAccess.GetDataTable(sql2, objconn, parms2)
        If dtChk.Rows.Count > 0 Then
            hidWOID.Value = Convert.ToString(dtChk.Rows(0)("woid"))
        End If

        Call loadData("org")
    End Sub

    Private Sub btnSch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnsch.Click
        If txtcycltype.Text <> "" Then
            If Not IsNumeric(txtcycltype.Text) Then
                Common.MessageBox(Me, "期別需輸入數字型態!!")
                Exit Sub
            End If
        End If
        Call search()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labDYear As Label = e.Item.FindControl("labDYear")
                Dim labDOrgName As Label = e.Item.FindControl("labDOrgName")
                Dim labDClassName As Label = e.Item.FindControl("labDClassName")
                Dim labDWrite As Label = e.Item.FindControl("labDWrite")
                Dim labDOral As Label = e.Item.FindControl("labDOral")
                Dim btnEdit As LinkButton = e.Item.FindControl("btnEdit")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                labDYear.Text = Convert.ToString(drv("years"))
                labDOrgName.Text = Convert.ToString(drv("orgname"))
                labDClassName.Text = Convert.ToString(drv("classname"))
                labDWrite.Text = Convert.ToString(drv("writeresult"))
                labDOral.Text = Convert.ToString(drv("oralresult"))
                If Convert.ToString(drv("woid")) <> "" Then
                    btnEdit.CommandArgument = "0" & "," & Convert.ToString(drv("woid"))
                Else
                    btnEdit.CommandArgument = "1" & "," & Convert.ToString(drv("years")) & "," & Convert.ToString(drv("distid")) & "," & Convert.ToString(drv("planid")) & "," & Convert.ToString(drv("orgid")) & "," & Convert.ToString(drv("ocid"))
                End If
        End Select
    End Sub

    Public Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edt"
                Dim strArr() As String = Split(Convert.ToString(e.CommandArgument), ",")
                tbsch.Visible = False
                tbEdit.Visible = True
                clsValue()
                Select Case strArr(0)
                    Case "0"
                        hidWOID.Value = strArr(1)
                        loadData("edit")
                    Case "1"
                        hidYears.Value = strArr(1)
                        hidDistID.Value = strArr(2)
                        hidPlanID.Value = strArr(3)
                        hidOrgID.Value = strArr(4)
                        hidOCID.Value = strArr(5)
                        loadData("add")
                End Select

        End Select
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Dim sda As New SqlDataAdapter
        Dim intCnt As Integer = 0
        Dim sql As String = ""
        'Dim conn As SqlConnection = Nothing
        'conn = DbAccess.GetConnection()
        'Call TIMS.OpenDbConn(conn)
        'TIMS.TestDbConn(Me, conn)

        Dim sda As New SqlDataAdapter
        Call TIMS.OpenDbConn(objconn)

        If txtWrite.Text <> "" And txtOral.Text <> "" Then
            If hidWOID.Value = "" Then
                '新增資料
                Dim iWOID As Integer = DbAccess.GetNewId(objconn, "ORG_WRITEORAL_WOID_SEQ,ORG_WRITEORAL,WOID")
                sql = " INSERT INTO ORG_WRITEORAL (WOID,planid,orgid,ocid,writeresult,oralresult,modifyacct,modifydate) "
                sql &= " VALUES (@woid,@planid,@orgid,@ocid,@writeresult,@oralresult,@modifyacct,GETDATE()) "
                With sda
                    .InsertCommand = New SqlCommand(sql, objconn)
                    .InsertCommand.Parameters.Clear()
                    .InsertCommand.Parameters.Add("@woid", SqlDbType.Int).Value = iWOID
                    .InsertCommand.Parameters.Add("@planid", SqlDbType.VarChar).Value = hidPlanID.Value
                    .InsertCommand.Parameters.Add("@orgid", SqlDbType.VarChar).Value = hidOrgID.Value
                    .InsertCommand.Parameters.Add("@ocid", SqlDbType.VarChar).Value = If(hidOCID.Value <> "", hidOCID.Value, Convert.DBNull)
                    .InsertCommand.Parameters.Add("@writeresult", SqlDbType.VarChar).Value = Convert.ToDecimal(txtWrite.Text)
                    .InsertCommand.Parameters.Add("@oralresult", SqlDbType.VarChar).Value = Convert.ToDecimal(txtOral.Text)
                    .InsertCommand.Parameters.Add("@modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.InsertCommand.Parameters.Add("@modifydate", SqlDbType.VarChar).Value = Now.ToString("yyyy/MM/dd")
                    '.InsertCommand.ExecuteNonQuery()  'edit，by:20181016
                    DbAccess.ExecuteNonQuery(sql, objconn, .InsertCommand.Parameters)  'edit，by:20181016
                End With
            Else
                '修改資料
                sql = " UPDATE org_writeoral SET writeresult = @writeresult ,oralresult = @oralresult, modifyacct = @modifyacct ,modifydate = GETDATE() WHERE woid = @woid "
                With sda
                    .UpdateCommand = New SqlCommand(sql, objconn)
                    .UpdateCommand.Parameters.Clear()
                    .UpdateCommand.Parameters.Add("@writeresult", SqlDbType.Float).Value = Convert.ToDecimal(txtWrite.Text)
                    .UpdateCommand.Parameters.Add("@oralresult", SqlDbType.Float).Value = Convert.ToDecimal(txtOral.Text)
                    .UpdateCommand.Parameters.Add("@modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .UpdateCommand.Parameters.Add("@woid", SqlDbType.VarChar).Value = hidWOID.Value
                    '.UpdateCommand.ExecuteNonQuery()  'edit，by:20181016
                    DbAccess.ExecuteNonQuery(sql, objconn, .UpdateCommand.Parameters)  'edit，by:20181016
                End With
            End If
        Else
            If hidWOID.Value <> "" Then
                '刪除資料
                sql = " DELETE org_writeoral WHERE woid = @woid "
                With sda
                    .DeleteCommand = New SqlCommand(sql, objconn)
                    .DeleteCommand.Parameters.Clear()
                    .DeleteCommand.Parameters.Add("@woid", SqlDbType.VarChar).Value = hidWOID.Value
                    '.DeleteCommand.ExecuteNonQuery()  'edit，by:20181016
                    DbAccess.ExecuteNonQuery(sql, objconn, .DeleteCommand.Parameters)  'edit，by:20181016
                End With
            End If
        End If

        intCnt = 1
        'Call TIMS.CloseDbConn(conn)

        If intCnt = 1 Then
            Common.MessageBox(Me, "儲存成功!")
            'btnBack_Click(sender, e)
            tbsch.Visible = True
            tbEdit.Visible = False
            Call search()
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        tbsch.Visible = True
        tbEdit.Visible = False
        Call search()
    End Sub
End Class