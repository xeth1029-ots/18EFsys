Partial Class OB_01_004
    Inherits AuthBasePage

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
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            Call cCreate1()
        End If

    End Sub

    Sub cCreate1()
        ddlDistID = TIMS.Get_DistID(ddlDistID, Nothing, objconn)
        ddlEDistID = TIMS.Get_DistID(ddlEDistID, Nothing, objconn)
        ddlOrg = TIMS.Get_OrgType(ddlOrg, objconn)
        ddlEDistID.SelectedValue = sm.UserInfo.DistID
        ddlEDistID.Enabled = False

        DataGridTable.Visible = False
        panelSch.Visible = True
        panelEdit.Visible = False

        '上層為委外訓練
        trPlanName.Visible = False
        If Request("tsn") <> "" And Request("Action") = "con" Then
            ViewState("tsn") = Request("tsn")
            Me.LabPlanName.Text = TIMS.Get_OB_Tendere(ViewState("tsn"), "PlanName")
            Me.LabTenderCName.Text = TIMS.Get_OB_Tendere(ViewState("tsn"), "TenderCName")
            trPlanName.Visible = True
            Query()
        Else
            ViewState("tsn") = ""
            btnAdd.Visible = False
            btnBack.Visible = False
        End If

        LitZip.Text = TIMS.Get_WorkZIPB3Link2()

        Dim btnCityZip_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(txtZip, txtZIPB3, hidZIP6W, txtCity, txtAddr)
        btnCityZip.Attributes.Add("onclick", btnCityZip_Attr_VAL)
    End Sub
    Sub Query()
        Dim sql As String = ""
        Dim dt As DataTable

        sql = " select b.TCsn, b.csn, a.OrgName, a.ComIDNO, a.ComCIDNO "
        sql += " from OB_Contractor a "
        sql += " join OB_TContractor b on b.csn=a.csn "
        sql += " WHERE 1=1 "

        If ViewState("tsn") <> "" Then
            sql += " AND b.tsn='" & ViewState("tsn") & "'"
        End If

        If ddlDistID.SelectedIndex > 0 Then
            sql += " AND b.DistID='" & ddlDistID.SelectedValue & "'"
        End If

        If txtOrgName.Text.Trim <> "" Then
            txtOrgName.Text = txtOrgName.Text.Trim
            sql += " AND a.OrgName like '%" & txtOrgName.Text & "%'"
        End If

        If txtComIDNO.Text.Trim <> "" Then
            txtComIDNO.Text = txtComIDNO.Text.Trim
            sql += " AND a.ComIDNO like '%" & txtComIDNO.Text & "%'"
        End If

        If txtComCIDNO.Text.Trim <> "" Then
            txtComCIDNO.Text = txtComCIDNO.Text.Trim
            sql += " AND a.ComCIDNO like '%" & txtComCIDNO.Text & "%'"
        End If

        dt = DbAccess.GetDataTable(sql, objconn)

        Me.LabActionType.Text = ""

        DataGridTable.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        If Page.IsValid Then
            Query()
        End If
    End Sub

    Sub ItemSet(ByVal strChk As String)
        If strChk = "edit" Then
            ddlEDistID.Visible = True
            lblDist.Visible = False
            txtTitle.BorderStyle = BorderStyle.NotSet
            txtTitle.ReadOnly = False
            ddlOrg.Visible = True
            lblOrg.Visible = False
            txtEComIDNO.BorderStyle = BorderStyle.NotSet
            txtEComIDNO.ReadOnly = False
            txtEComCIDNO.BorderStyle = BorderStyle.NotSet
            txtEComCIDNO.ReadOnly = False
            txtTel.BorderStyle = BorderStyle.NotSet
            txtTel.ReadOnly = False
            txtFax.BorderStyle = BorderStyle.NotSet
            txtFax.ReadOnly = False
            'txtZip.BorderStyle = BorderStyle.NotSet
            'txtZIPB3.BorderStyle = BorderStyle.NotSet

            txtCity.BorderStyle = BorderStyle.NotSet
            txtCity.ReadOnly = False
            btnCityZip.Visible = True
            txtAddr.BorderStyle = BorderStyle.NotSet
            txtAddr.ReadOnly = False
            txtMName.BorderStyle = BorderStyle.NotSet
            txtMName.ReadOnly = False
            txtMIDNO.BorderStyle = BorderStyle.NotSet
            txtMIDNO.ReadOnly = False
            txtPlanMaster.BorderStyle = BorderStyle.NotSet
            txtPlanMaster.ReadOnly = False
            txtPMPhone.BorderStyle = BorderStyle.NotSet
            txtPMPhone.ReadOnly = False
            txtPMFax.BorderStyle = BorderStyle.NotSet
            txtPMFax.ReadOnly = False
            txtCName.BorderStyle = BorderStyle.NotSet
            txtCName.ReadOnly = False
            rblCSex.Visible = True
            lblCSex.Visible = False
            txtCPhone.BorderStyle = BorderStyle.NotSet
            txtCPhone.ReadOnly = False
            txtCCell.BorderStyle = BorderStyle.NotSet
            txtCCell.ReadOnly = False
            txtCEMail.BorderStyle = BorderStyle.NotSet
            txtCEMail.ReadOnly = False
            txtCFax.BorderStyle = BorderStyle.NotSet
            txtCFax.ReadOnly = False
            btnSave.Visible = True
        Else 'view
            lblDist.Text = ddlEDistID.SelectedItem.Text
            ddlEDistID.Visible = False
            lblDist.Visible = True
            txtTitle.BorderStyle = BorderStyle.None
            txtTitle.ReadOnly = True
            lblOrg.Text = ddlOrg.SelectedItem.Text
            ddlOrg.Visible = False
            lblOrg.Visible = True
            txtEComIDNO.BorderStyle = BorderStyle.None
            txtEComIDNO.ReadOnly = True
            txtEComCIDNO.BorderStyle = BorderStyle.None
            txtEComCIDNO.ReadOnly = True
            txtTel.BorderStyle = BorderStyle.None
            txtTel.ReadOnly = True
            txtFax.BorderStyle = BorderStyle.None
            txtFax.ReadOnly = True
            'txtZip.BorderStyle = BorderStyle.None
            'txtZIPB3.BorderStyle = BorderStyle.None

            txtCity.BorderStyle = BorderStyle.None
            txtCity.ReadOnly = True
            btnCityZip.Visible = False
            txtAddr.BorderStyle = BorderStyle.None
            txtAddr.ReadOnly = True
            txtMName.BorderStyle = BorderStyle.None
            txtMName.ReadOnly = True
            txtMIDNO.BorderStyle = BorderStyle.None
            txtMIDNO.ReadOnly = True
            txtPlanMaster.BorderStyle = BorderStyle.None
            txtPlanMaster.ReadOnly = True
            txtPMPhone.BorderStyle = BorderStyle.None
            txtPMPhone.ReadOnly = True
            txtPMFax.BorderStyle = BorderStyle.None
            txtPMFax.ReadOnly = True
            txtCName.BorderStyle = BorderStyle.None
            txtCName.ReadOnly = True
            lblCSex.Text = rblCSex.SelectedItem.Text
            rblCSex.Visible = False
            lblCSex.Visible = True
            txtCPhone.BorderStyle = BorderStyle.None
            txtCPhone.ReadOnly = True
            txtCCell.BorderStyle = BorderStyle.None
            txtCCell.ReadOnly = True
            txtCEMail.BorderStyle = BorderStyle.None
            txtCEMail.ReadOnly = True
            txtCFax.BorderStyle = BorderStyle.None
            txtCFax.ReadOnly = True
            btnSave.Visible = False
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        panelSch.Visible = False
        panelEdit.Visible = True

        'ddlEDistID.SelectedIndex = 0
        ddlEDistID.SelectedValue = sm.UserInfo.DistID
        hidTCsn.Value = ""
        txtTitle.Text = ""
        ddlOrg.SelectedIndex = 0
        txtEComIDNO.Text = ""
        txtEComCIDNO.Text = ""
        txtTel.Text = ""
        txtFax.Text = ""
        txtZip.Value = ""
        txtZIPB3.Value = ""
        hidZIP6W.Value = ""
        txtCity.Text = ""
        txtAddr.Text = ""
        txtMName.Text = ""
        txtMIDNO.Text = ""
        txtPlanMaster.Text = ""
        txtPMPhone.Text = ""
        txtPMFax.Text = ""
        txtCName.Text = ""
        rblCSex.SelectedIndex = 0
        txtCPhone.Text = ""
        txtCCell.Text = ""
        txtCEMail.Text = ""
        txtCFax.Text = ""

        Me.LabActionType.Text = "新增"
        ItemSet("edit") '可編輯模式

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnView As Button = e.Item.FindControl("btnView")
                Dim btnEdit As Button = e.Item.FindControl("btnEdit")
                Dim btnDel As Button = e.Item.FindControl("btnDel")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                btnView.CommandArgument = drv("TCsn")
                If ViewState("tsn") <> "" Then
                    btnEdit.CommandArgument = drv("TCsn")
                    btnDel.CommandArgument = drv("csn")
                    btnDel.Attributes("onclick") = "return confirm('您確定要刪除 " & e.Item.Cells(1).Text & " 的資料?');"
                Else
                    btnEdit.Visible = False
                    btnDel.Visible = False
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Return
        Dim i_CSN As Integer = Val(TIMS.ClearSQM(e.CommandArgument))
        Dim sql As String = ""
        Select Case e.CommandName
            Case "del"

                Try
                    Dim parms As New Hashtable
                    parms.Add("CSN", i_CSN)
                    sql = "SELECT TCSN FROM OB_TCONTRACTOR WHERE CSN=@CSN"
                    Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

                    If dt.Rows.Count = 1 Then
                        Dim d_parms As New Hashtable
                        d_parms.Add("CSN", i_CSN)
                        Dim d_sql As String = "DELETE OB_CONTRACTOR WHERE CSN=@CSN"
                        DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)
                    End If

                    Dim d_parms2 As New Hashtable
                    d_parms2.Add("CSN", i_CSN)
                    d_parms2.Add("TSN", Val(ViewState("tsn")))
                    Dim d_sql2 As String = "DELETE OB_TCONTRACTOR WHERE CSN=@CSN AND TSN=@TSN "
                    DbAccess.ExecuteNonQuery(d_sql2, objconn, d_parms2)

                    Common.MessageBox(Me, "刪除成功！")
                    Query()

                Catch ex As Exception
                    'trans.Rollback()
                    Common.MessageBox(Me, "刪除錯誤：" + ex.Message.ToString)
                End Try


            Case "edit"
                panelSch.Visible = False
                panelEdit.Visible = True
                lblDist.Visible = False
                lblOrg.Visible = False
                btnSave.Visible = True

                hidTCsn.Value = e.CommandArgument
                Me.LabActionType.Text = "修改"
                LoadData("edit")

            Case "view"
                panelSch.Visible = False
                panelEdit.Visible = True
                lblDist.Visible = True
                lblOrg.Visible = True
                btnSave.Visible = False

                hidTCsn.Value = e.CommandArgument
                Me.LabActionType.Text = "檢視"
                LoadData("view")
        End Select
    End Sub

    Sub LoadData(ByVal strChk As String)
        Dim sql As String
        Dim dr As DataRow

        hidTCsn.Value = TIMS.ClearSQM(hidTCsn.Value)
        sql = "select * from OB_Contractor a "
        sql += " join OB_TContractor b on b.csn=a.csn "
        sql += " WHERE b.TCsn='" & hidTCsn.Value & "'"
        If ViewState("tsn") <> "" Then
            ViewState("tsn") = TIMS.ClearSQM(ViewState("tsn"))
            sql += " AND b.tsn='" & ViewState("tsn") & "'"
        End If
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then
            ddlEDistID.SelectedValue = Convert.ToString(dr("DistID"))
            txtTitle.Text = Convert.ToString(dr("OrgName"))
            ddlOrg.SelectedValue = Convert.ToString(dr("OrgKind"))
            txtEComIDNO.Text = Convert.ToString(dr("ComIDNO"))
            txtEComCIDNO.Text = Convert.ToString(dr("ComCIDNO"))
            txtTel.Text = Convert.ToString(dr("Phone"))
            txtFax.Text = Convert.ToString(dr("Fax"))
            txtZip.Value = Convert.ToString(dr("Zip"))
            hidZIP6W.Value = Convert.ToString(dr("ZIP6W"))
            txtZIPB3.Value = TIMS.GetZIPCODEB3(dr("ZIP6W"))
            txtCity.Text = "(" & txtZip.Value & ")" & TIMS.Get_ZipName(txtZip.Value) & "[" & TIMS.Get_ZipLName(txtZip.Value, objconn) & "]"
            txtAddr.Text = Convert.ToString(dr("Address"))
            txtMName.Text = Convert.ToString(dr("MasterName"))
            txtMIDNO.Text = Convert.ToString(dr("MasterIDNO"))
            txtPlanMaster.Text = Convert.ToString(dr("PlanMaster"))
            txtPMPhone.Text = Convert.ToString(dr("PlanMasterPhone"))
            txtPMFax.Text = Convert.ToString(dr("PlanMasterFax"))
            txtCName.Text = Convert.ToString(dr("ContactName"))
            rblCSex.SelectedValue = Convert.ToString(dr("ContactSex"))
            txtCPhone.Text = Convert.ToString(dr("ContactPhone"))
            txtCCell.Text = Convert.ToString(dr("ContactCell"))
            txtCEMail.Text = Convert.ToString(dr("ContactEMail"))
            txtCFax.Text = Convert.ToString(dr("ContactFax"))

            ItemSet(strChk)
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim sql As String = ""
        Dim intCsn As Integer = 0
        Dim dr As DataRow = Nothing 'OB_Contractor
        Dim drChk As DataRow = Nothing 'OB_TContractor

        Dim blnCmdUC As Boolean = False
        Dim blnCmdIC As Boolean = False
        Dim blnCmdUTC As Boolean = False
        Dim blnCmdITC As Boolean = False

        If hidTCsn.Value = "" Then '新增
            sql = "select csn from OB_Contractor "
            sql += "where ComIDNO='" & txtEComIDNO.Text & "'"
            dr = DbAccess.GetOneRow(sql, objconn)
            If Not dr Is Nothing Then
                intCsn = CInt(dr("csn"))
                sql = "select TCsn from OB_TContractor "
                sql += "where csn='" & intCsn & "' and tsn='" & ViewState("tsn") & "'"
                drChk = DbAccess.GetOneRow(sql, objconn)
                If Not drChk Is Nothing Then
                    Common.MessageBox(Me, "該廠商已存在，無法新增！")
                    Exit Sub
                End If
                blnCmdUC = True
            Else
                blnCmdIC = True
            End If
            blnCmdITC = True
        Else '修改
            sql = "select csn from OB_TContractor where TCsn='" & hidTCsn.Value & "' "
            drChk = DbAccess.GetOneRow(sql, objconn)
            If Not drChk Is Nothing Then
                sql = "select csn from OB_Contractor "
                sql += "where ComIDNO='" & txtEComIDNO.Text & "' and csn<>'" & drChk("csn") & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    Common.MessageBox(Me, "該廠商已存在，無法修改！")
                    Exit Sub
                End If
                intCsn = CInt(drChk("csn"))
                sql = "select TCsn from OB_TContractor "
                sql += "where csn='" & intCsn & "' and tsn='" & ViewState("tsn") & "' "
                sql += "and TCsn<>'" & hidTCsn.Value & "'"
                drChk = DbAccess.GetOneRow(sql, objconn)
                If Not drChk Is Nothing Then
                    Common.MessageBox(Me, "該廠商已存在，無法修改！")
                    Exit Sub
                End If
                blnCmdUC = True
            End If
            blnCmdUTC = True
        End If

        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            sql = "insert into OB_Contractor (OrgName, OrgKind, ComIDNO, ComCIDNO, ModifyAcct, ModifyTime) "
            sql += "values(@OrgName, @OrgKind, @ComIDNO, @ComCIDNO, '" & sm.UserInfo.UserID & "', getdate())"
            Dim cmdIC As New SqlCommand(sql, objconn, trans)

            sql = "update OB_Contractor set "
            sql += "OrgName=@OrgName, OrgKind=@OrgKind, ComIDNO=@ComIDNO, ComCIDNO=@ComCIDNO, "
            sql += "ModifyAcct='" & sm.UserInfo.UserID & "', ModifyTime=getdate() "
            sql += "where csn='" & dr("csn") & "'"
            Dim cmdUC As New SqlCommand(sql, objconn, trans)

            sql = "insert into OB_TContractor "
            sql += "(tsn, csn, DistID, Phone, Fax, Zip, ZIP6W, Address, MasterName, MasterIDNO, "
            sql += " PlanMaster, PlanMasterPhone, PlanMasterFax, ContactName, ContactSex, ContactPhone, "
            sql += " ContactCell, ContactEMail, ContactFax, CreateAcct, CreateTime, ModifyAcct, ModifyTime) "
            sql += "values"
            sql += "(@tsn, @csn, @DistID, @Phone, @Fax, @Zip, @ZIP6W, @Address, @MasterName, @MasterIDNO, "
            sql += " @PlanMaster, @PlanMasterPhone, @PlanMasterFax, @ContactName, @ContactSex, @ContactPhone, "
            sql += " @ContactCell, @ContactEMail, @ContactFax, '" & sm.UserInfo.UserID & "', getdate(), '" & sm.UserInfo.UserID & "', getdate())"
            Dim cmdITC As New SqlCommand(sql, objconn, trans)

            sql = "update OB_TContractor set "
            sql += "tsn=@tsn, csn=@csn, DistID=@DistID, Phone=@Phone, Fax=@Fax, Zip=@Zip, ZIP6W=@ZIP6W, "
            sql += "Address=@Address, MasterName=@MasterName, MasterIDNO=@MasterIDNO, PlanMaster=@PlanMaster, "
            sql += "PlanMasterPhone=@PlanMasterPhone, PlanMasterFax=@PlanMasterFax, ContactName=@ContactName, "
            sql += "ContactSex=@ContactSex, ContactPhone=@ContactPhone, ContactCell=@ContactCell, "
            sql += "ContactEMail=@ContactEMail, ContactFax=@ContactFax, ModifyAcct='" & sm.UserInfo.UserID & "', "
            sql += "ModifyTime=getdate() "
            sql += "where TCsn='" & hidTCsn.Value & "'"
            Dim cmdUTC As New SqlCommand(sql, objconn, trans)

            If blnCmdIC Then
                With cmdIC
                    .Parameters.Clear()
                    .Parameters.Add("OrgName", SqlDbType.NVarChar).Value = Convert.ToString(txtTitle.Text)
                    .Parameters.Add("OrgKind", SqlDbType.NVarChar).Value = Convert.ToString(ddlOrg.SelectedValue)
                    .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = Convert.ToString(txtEComIDNO.Text)
                    .Parameters.Add("ComCIDNO", SqlDbType.NVarChar).Value = Convert.ToString(txtEComCIDNO.Text)
                    .ExecuteNonQuery()
                End With
            End If
            If blnCmdUC Then
                With cmdUC
                    .Parameters.Clear()
                    .Parameters.Add("OrgName", SqlDbType.NVarChar).Value = Convert.ToString(txtTitle.Text)
                    .Parameters.Add("OrgKind", SqlDbType.NVarChar).Value = Convert.ToString(ddlOrg.SelectedValue)
                    .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = Convert.ToString(txtEComIDNO.Text)
                    .Parameters.Add("ComCIDNO", SqlDbType.NVarChar).Value = Convert.ToString(txtEComCIDNO.Text)
                    .ExecuteNonQuery()
                End With
            End If

            hidZIP6W.Value = TIMS.GetZIPCODE6W(txtZip.Value, txtZIPB3.Value)
            If blnCmdITC Then
                'sql = "SELECT identity" 'OB_CONTRACTOR_CSN_SEQ
                If intCsn = 0 Then intCsn = DbAccess.GetNewId(trans, "OB_CONTRACTOR_CSN_SEQ")

                With cmdITC
                    .Parameters.Clear()
                    .Parameters.Add("tsn", SqlDbType.Int).Value = CInt(ViewState("tsn"))
                    .Parameters.Add("csn", SqlDbType.Int).Value = intCsn
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = Convert.ToString(ddlEDistID.SelectedValue)
                    .Parameters.Add("Phone", SqlDbType.VarChar).Value = Convert.ToString(txtTel.Text)
                    .Parameters.Add("Fax", SqlDbType.VarChar).Value = Convert.ToString(txtFax.Text)
                    .Parameters.Add("Zip", SqlDbType.Int).Value = CInt(txtZip.Value)
                    .Parameters.Add("ZIP6W", SqlDbType.VarChar).Value = hidZIP6W.Value
                    .Parameters.Add("Address", SqlDbType.NVarChar).Value = Convert.ToString(txtAddr.Text)
                    .Parameters.Add("MasterName", SqlDbType.NVarChar).Value = Convert.ToString(txtMName.Text)
                    .Parameters.Add("MasterIDNO", SqlDbType.Char).Value = Convert.ToString(txtMIDNO.Text)
                    .Parameters.Add("PlanMaster", SqlDbType.NVarChar).Value = Convert.ToString(txtPlanMaster.Text)
                    .Parameters.Add("PlanMasterPhone", SqlDbType.VarChar).Value = Convert.ToString(txtPMPhone.Text)
                    .Parameters.Add("PlanMasterFax", SqlDbType.VarChar).Value = Convert.ToString(txtPMFax.Text)
                    .Parameters.Add("ContactName", SqlDbType.NVarChar).Value = Convert.ToString(txtCName.Text)
                    .Parameters.Add("ContactSex", SqlDbType.Char).Value = Convert.ToString(rblCSex.SelectedValue)
                    .Parameters.Add("ContactPhone", SqlDbType.VarChar).Value = Convert.ToString(txtCPhone.Text)
                    .Parameters.Add("ContactCell", SqlDbType.VarChar).Value = Convert.ToString(txtCCell.Text)
                    .Parameters.Add("ContactEMail", SqlDbType.VarChar).Value = Convert.ToString(txtCEMail.Text)
                    .Parameters.Add("ContactFax", SqlDbType.VarChar).Value = Convert.ToString(txtCFax.Text)
                    .ExecuteNonQuery()
                End With
            End If

            If blnCmdUTC Then
                With cmdUTC
                    .Parameters.Clear()
                    .Parameters.Add("tsn", SqlDbType.Int).Value = CInt(ViewState("tsn"))
                    .Parameters.Add("csn", SqlDbType.Int).Value = intCsn
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = Convert.ToString(ddlEDistID.SelectedValue)
                    .Parameters.Add("Phone", SqlDbType.VarChar).Value = Convert.ToString(txtTel.Text)
                    .Parameters.Add("Fax", SqlDbType.VarChar).Value = Convert.ToString(txtFax.Text)
                    .Parameters.Add("Zip", SqlDbType.Int).Value = CInt(txtZip.Value)
                    .Parameters.Add("ZIP6W", SqlDbType.VarChar).Value = hidZIP6W.Value
                    .Parameters.Add("Address", SqlDbType.NVarChar).Value = Convert.ToString(txtAddr.Text)
                    .Parameters.Add("MasterName", SqlDbType.NVarChar).Value = Convert.ToString(txtMName.Text)
                    .Parameters.Add("MasterIDNO", SqlDbType.Char).Value = Convert.ToString(txtMIDNO.Text)
                    .Parameters.Add("PlanMaster", SqlDbType.NVarChar).Value = Convert.ToString(txtPlanMaster.Text)
                    .Parameters.Add("PlanMasterPhone", SqlDbType.VarChar).Value = Convert.ToString(txtPMPhone.Text)
                    .Parameters.Add("PlanMasterFax", SqlDbType.VarChar).Value = Convert.ToString(txtPMFax.Text)
                    .Parameters.Add("ContactName", SqlDbType.NVarChar).Value = Convert.ToString(txtCName.Text)
                    .Parameters.Add("ContactSex", SqlDbType.Char).Value = Convert.ToString(rblCSex.SelectedValue)
                    .Parameters.Add("ContactPhone", SqlDbType.VarChar).Value = Convert.ToString(txtCPhone.Text)
                    .Parameters.Add("ContactCell", SqlDbType.VarChar).Value = Convert.ToString(txtCCell.Text)
                    .Parameters.Add("ContactEMail", SqlDbType.VarChar).Value = Convert.ToString(txtCEMail.Text)
                    .Parameters.Add("ContactFax", SqlDbType.VarChar).Value = Convert.ToString(txtCFax.Text)
                    .ExecuteNonQuery()
                End With
            End If
            trans.Commit()

            Common.RespWrite(Me, "<script>alert('儲存成功');</script>")
            panelSch.Visible = True
            panelEdit.Visible = False
            Query()
        Catch ex As Exception
            trans.Rollback()
            Common.MessageBox(Me, ex.ToString)
        Finally
            objconn.Close()
            objconn.Dispose()
            objconn = Nothing
        End Try
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.LabActionType.Text = ""
        panelSch.Visible = True
        panelEdit.Visible = False
    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        Dim strMessage As String = ""
        For Each obj As WebControls.BaseValidator In Page.Validators
            If obj.IsValid = False Then
                strMessage &= obj.ErrorMessage & vbCrLf
            End If
        Next

        If strMessage <> "" Then
            Common.MessageBox(Page, strMessage)
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        TIMS.Utl_Redirect1(Me, "OB_01_001.aspx?ID=" & Request("ID"))
    End Sub


End Class
