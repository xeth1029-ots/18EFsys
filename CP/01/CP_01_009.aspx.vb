Partial Class CP_01_009
    Inherits AuthBasePage

    'CP_01_009empty2
    'CP_01_009
    'Dim FunDr As DataRow
    'Const Cst_separator As String = "+"

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            'center.Text = sm.UserInfo.OrgName
            'RIDValue.Value = sm.UserInfo.RID
            btnSearch.Attributes("onclick") = "javascript:return search();"
            btnSave.Attributes("onclick") = "javascript:return saveChk();"
            txtQ1A.Attributes("onkeyup") = "javascript:return Q1Chk();"
            txtQ1B.Attributes("onkeyup") = "javascript:return Q1Chk();"
            DataGridTable.Visible = False
            panelSch.Visible = True
            panelEdit.Visible = False
        End If

#Region "(No Use)"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & CInt(Request("ID")) & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = 1 Then
        '            btnAdd.Enabled = True
        '        Else
        '            btnAdd.Enabled = False
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            btnSearch.Enabled = True
        '        Else
        '            btnSearch.Enabled = False
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

#End Region
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim V_VisitorDate1 As String = If(flag_ROC, TIMS.Cdate18(txtSDate.Text), TIMS.Cdate3(txtSDate.Text))  'edit，by:20181018
        Dim V_VisitorDate2 As String = If(flag_ROC, TIMS.Cdate18(txtEDate.Text), TIMS.Cdate3(txtEDate.Text))  'edit，by:20181018

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT o.OrgID, o.PlanID, o.VisitorDate, oi.OrgName" & vbCrLf
        sql &= " ,concat(ip.Years , id.Name , k.PlanName , ip.seq) AS PlanName" & vbCrLf
        sql &= " FROM Org_Report o" & vbCrLf
        sql &= " JOIN org_orginfo oi ON oi.orgid = o.OrgID" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID = o.PlanID" & vbCrLf
        sql &= " JOIN ID_District id ON id.DistID = ip.DistID" & vbCrLf
        sql &= " JOIN Key_Plan k ON k.TPlanID = ip.TPlanID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        hidOrgID.Value = TIMS.ClearSQM(hidOrgID.Value)
        hidPlanID.Value = TIMS.ClearSQM(hidPlanID.Value)
        If hidOrgID.Value <> "" And hidPlanID.Value <> "" And txtTBplan.Text <> "" Then sql += " AND o.OrgID = '" & hidOrgID.Value & "' AND o.PlanID = '" & hidPlanID.Value & "' "
        If V_VisitorDate1 <> "" Then sql += " AND o.VisitorDate >= " & TIMS.To_date(V_VisitorDate1) & vbCrLf
        If V_VisitorDate2 <> "" Then sql += " AND o.VisitorDate < DATEADD(DAY, 1, " & TIMS.To_date(V_VisitorDate2) & ") " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            'PageControler1.SqlDataCreate(sql, "VisitorDate")
            PageControler1.PageDataTable = dt
            PageControler1.Sort = "VisitorDate"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnEdit As Button = e.Item.Cells(3).FindControl("btnEdit")
                Dim btnDel As Button = e.Item.Cells(3).FindControl("btnDel")
                Dim btnPrint As Button = e.Item.Cells(3).FindControl("btnPrint")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OrgID", Convert.ToString(drv("OrgID")))
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "VisitorDate", Convert.ToString(drv("VisitorDate")))

                btnEdit.CommandArgument = sCmdArg 'drv("OrgID") & Cst_separator & drv("PlanID") & Cst_separator & drv("VisitorDate")
                btnDel.CommandArgument = sCmdArg '"OrgID='" & drv("OrgID") & "' and PlanID='" & drv("PlanID") & "' and VisitorDate='" & drv("VisitorDate") & "'"
                btnDel.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                Dim s_myValue As String = "OrgID=" & Convert.ToString(drv("OrgID")) & "&vdate=" & Convert.ToString(drv("VisitorDate")) & "&planid=" & Convert.ToString(drv("PlanID"))
                btnPrint.Attributes("onclick") = ReportQuery.ReportScript(Me, "Report", "CP_01_009", s_myValue)
        End Select
        'If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
        'End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing Then Return
        If Convert.ToString(e.CommandArgument) = "" Then Return
        If Convert.ToString(e.CommandName) = "" Then Return
        Dim sCmdArg As String = e.CommandArgument

        hidEditOrgID.Value = TIMS.GetMyValue(sCmdArg, "OrgID") 'Split(e.CommandArgument, Cst_separator)(0)
        hidEditPlanID.Value = TIMS.GetMyValue(sCmdArg, "PlanID") 'Split(e.CommandArgument, Cst_separator)(1)
        Dim v_VisitorDate As String = TIMS.GetMyValue(sCmdArg, "VisitorDate") 'Split(e.CommandArgument, Cst_separator)(2)
        txtVDate.Text = If(flag_ROC, TIMS.Cdate17(v_VisitorDate), TIMS.Cdate3(v_VisitorDate))  'edit，by:20181018

        Select Case e.CommandName
            Case "edit"
                panelSch.Visible = False
                panelEdit.Visible = True
                btnChoice.Visible = False
                btnClear.Visible = False
                panelImg.Visible = False
                hidChkEdit.Value = "Y"

                LoadData()
            Case "del"

                Dim s_parms As New Hashtable
                s_parms.Add("OrgID", hidEditOrgID.Value)
                s_parms.Add("PlanID", hidEditPlanID.Value)
                s_parms.Add("VisitorDate", v_VisitorDate)

                Dim sql As String = ""
                sql = ""
                sql &= " SELECT 1 FROM ORG_REPORT"
                sql &= " WHERE 1=1"
                sql &= " AND OrgID=@OrgID"
                sql &= " AND PlanID=@PlanID"
                sql &= " AND VisitorDate=@VisitorDate"
                Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, s_parms)
                If dt1.Rows.Count = 1 Then
                    '查詢到一筆資料
                    sql = ""
                    sql &= " DELETE Org_Report" '& e.CommandArgument
                    sql &= " WHERE 1=1"
                    sql &= " AND OrgID=@OrgID"
                    sql &= " AND PlanID=@PlanID"
                    sql &= " AND VisitorDate=@VisitorDate"
                    Dim d_sql As String = sql
                    DbAccess.ExecuteNonQuery(d_sql, objconn, s_parms)
                    Common.MessageBox(Me, "刪除成功!!")
                End If

                btnSearch_Click(btnSearch, e)
        End Select
        'If e.CommandName = "edit" Then
        'ElseIf e.CommandName = "del" Then
        'End If
    End Sub

    Private Sub LoadData()
        Dim V_tVDate1 As String = If(flag_ROC, TIMS.Cdate18(txtVDate.Text), TIMS.Cdate3(txtVDate.Text))  'edit，by:20181018
        hidEditOrgID.Value = TIMS.ClearSQM(hidEditOrgID.Value)
        hidEditPlanID.Value = TIMS.ClearSQM(hidEditPlanID.Value)

        Dim dr As DataRow
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT o.*" & vbCrLf
        sql &= " ,concat(ip.Years , id.Name , k.PlanName , ip.seq) AS PlanName" & vbCrLf
        sql &= " FROM Org_Report o" & vbCrLf
        sql &= " JOIN org_orginfo oi ON oi.orgid = o.OrgID" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID = o.PlanID" & vbCrLf
        sql &= " JOIN ID_District id ON id.DistID = ip.DistID" & vbCrLf
        sql &= " JOIN Key_Plan k ON k.TPlanID = ip.TPlanID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql += " AND o.OrgID = '" & hidEditOrgID.Value & "' " & vbCrLf
        sql += " AND o.PlanID = '" & hidEditPlanID.Value & "' " & vbCrLf
        sql += " AND o.VisitorDate = " & TIMS.To_date(V_tVDate1) & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr IsNot Nothing Then
            txtTBplan2.Text = Convert.ToString(dr("PlanName"))

            rdoQ11.Checked = False
            rdoQ10.Checked = False
            Select Case Convert.ToString(dr("Q1"))
                Case "0"
                    rdoQ10.Checked = True
                Case "1"
                    rdoQ11.Checked = True
            End Select

            txtQ1Per.Text = Convert.ToString(dr("Q1Per"))
            txtQ1A.Text = Convert.ToString(dr("Q1A"))
            txtQ1B.Text = Convert.ToString(dr("Q1B"))
            txtQ1Note.Text = Convert.ToString(dr("Q1Note"))

            rdoQ21.Checked = False
            rdoQ22.Checked = False
            rdoQ20.Checked = False
            Select Case Convert.ToString(dr("Q2"))
                Case "0"
                    rdoQ20.Checked = True
                Case "1"
                    rdoQ21.Checked = True
                Case "2"
                    rdoQ22.Checked = True
            End Select
            txtQ2Other.Text = Convert.ToString(dr("Q2Other"))
            txtQ2Note.Text = Convert.ToString(dr("Q2Note"))

            rdoQ31.Checked = False
            rdoQ32.Checked = False
            rdoQ30.Checked = False
            Select Case Convert.ToString(dr("Q3"))
                Case "0"
                    rdoQ30.Checked = True
                Case "1"
                    rdoQ31.Checked = True
                Case "2"
                    rdoQ32.Checked = True
            End Select
            txtQ3Other.Text = Convert.ToString(dr("Q3Other"))
            txtQ3Note.Text = Convert.ToString(dr("Q3Note"))

            rdoQ41.Checked = False
            rdoQ42.Checked = False
            rdoQ40.Checked = False
            Select Case Convert.ToString(dr("Q4"))
                Case "0"
                    rdoQ40.Checked = True
                Case "1"
                    rdoQ41.Checked = True
                Case "2"
                    rdoQ42.Checked = True
            End Select
            txtQ4Other.Text = Convert.ToString(dr("Q4Other"))
            txtQ4Note.Text = Convert.ToString(dr("Q4Note"))

            rdoQ51.Checked = False
            rdoQ52.Checked = False
            rdoQ53.Checked = False
            rdoQ50.Checked = False
            Select Case Convert.ToString(dr("Q5"))'dr("Q5")
                Case "0"
                    rdoQ50.Checked = True
                Case "1"
                    rdoQ51.Checked = True
                Case "2"
                    rdoQ52.Checked = True
                Case "3"
                    rdoQ53.Checked = True
            End Select
            txtQ5Other.Text = Convert.ToString(dr("Q5Other"))
            txtQ5Note.Text = Convert.ToString(dr("Q5Note"))

            txtQ6.Text = Convert.ToString(dr("Q6"))
            txtQ6Other.Text = Convert.ToString(dr("Q6Other"))
            txtQ7Note.Text = Convert.ToString(dr("Q7Note"))
            txtQ8Note.Text = Convert.ToString(dr("Q8Note"))

            txtUnitName.Text = Convert.ToString(dr("UnitName"))
            txtFillerName.Text = Convert.ToString(dr("FillerName"))
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim V_tVDate1 As String = If(flag_ROC, TIMS.Cdate18(txtVDate.Text), TIMS.Cdate3(txtVDate.Text))  'edit，by:20181018
        hidEditOrgID.Value = TIMS.ClearSQM(hidEditOrgID.Value)
        hidEditPlanID.Value = TIMS.ClearSQM(hidEditPlanID.Value)

        Dim sql As String
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable
        Dim dr As DataRow

        TIMS.OpenDbConn(objconn)

        If hidChkEdit.Value = "" Then '新增
            sql = ""
            sql &= " Select PlanID FROM Org_Report"
            sql &= " WHERE 1=1"
            sql &= " And OrgID = '" & hidEditOrgID.Value & "'"
            sql &= " AND PlanID = '" & hidEditPlanID.Value & "'"
            sql &= " AND VisitorDate = " & TIMS.To_date(V_tVDate1) & vbCrLf
            dr = DbAccess.GetOneRow(sql, objconn)
            If Not dr Is Nothing Then
                Common.MessageBox(Me, "該訪查已存在，無法新增！")
                Exit Sub
            End If
        End If

        If hidChkEdit.Value = "" Then '新增
            sql = ""
            sql = " SELECT * FROM Org_Report WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, da, objconn)
            dr = dt.NewRow
            dt.Rows.Add(dr)

            dr("OrgID") = Convert.ToString(hidEditOrgID.Value)
            dr("PlanID") = Convert.ToString(hidEditPlanID.Value)
            dr("VisitorDate") = Convert.ToDateTime(V_tVDate1)
        Else '修改
            sql = ""
            sql &= " Select * FROM Org_Report"
            sql &= " WHERE 1=1"
            sql &= " AND OrgID = '" & hidEditOrgID.Value & "'"
            sql &= " AND PlanID = '" & hidEditPlanID.Value & "'"
            sql &= " AND VisitorDate = " & TIMS.To_date(V_tVDate1) & vbCrLf
            dt = DbAccess.GetDataTable(sql, da, objconn)
            dr = dt.Rows(0)

        End If

        Dim vQ1 As String = ""
        If rdoQ11.Checked = True Then
            vQ1 = "1"
        ElseIf rdoQ10.Checked = True Then
            vQ1 = "0"
        End If
        dr("Q1") = vQ1
        dr("Q1Per") = txtQ1Per.Text
        dr("Q1A") = CInt(txtQ1A.Text)
        dr("Q1B") = CInt(txtQ1B.Text)
        dr("Q1Note") = Convert.ToString(txtQ1Note.Text)

        Dim vQ2 As String = ""
        'dr("Q2Other") = "" '選其他 才會帶入值
        If rdoQ21.Checked = True Then
            vQ2 = "1"
        ElseIf rdoQ22.Checked = True Then
            vQ2 = "2"
            'dr("Q2Other") = Convert.ToString(txtQ2Other.Text) '選其他 才會帶入值
        ElseIf rdoQ20.Checked = True Then
            vQ2 = "0"
        End If
        dr("Q2") = "2"
        dr("Q2Other") = If(vQ2 = "2", txtQ2Other.Text, Convert.DBNull) '選其他 才會帶入值
        dr("Q2Note") = Convert.ToString(txtQ2Note.Text)

        Dim vQ3 As String = ""
        'dr("Q3Other") = "" '選其他 才會帶入值
        If rdoQ31.Checked = True Then
            vQ3 = "1"
        ElseIf rdoQ32.Checked = True Then
            vQ3 = "2"
            'dr("Q3Other") = Convert.ToString(txtQ3Other.Text) '選其他 才會帶入值
        ElseIf rdoQ30.Checked = True Then
            vQ3 = "0"
        End If
        dr("Q3") = vQ3
        dr("Q3Other") = If(vQ3 = "2", txtQ3Other.Text, Convert.DBNull) '選其他 才會帶入值
        dr("Q3Note") = Convert.ToString(txtQ3Note.Text)

        Dim vQ4 As String = ""
        'dr("Q4Other") = "" '選其他 才會帶入值
        If rdoQ41.Checked = True Then
            vQ4 = "1"
        ElseIf rdoQ42.Checked = True Then
            vQ4 = "2"
            'dr("Q4Other") = Convert.ToString(txtQ4Other.Text) '選其他 才會帶入值
        ElseIf rdoQ40.Checked = True Then
            vQ4 = "0"
        End If
        dr("Q4") = vQ4
        dr("Q4Other") = If(vQ4 = "2", txtQ4Other.Text, Convert.DBNull) '選其他 才會帶入值
        dr("Q4Note") = Convert.ToString(txtQ4Note.Text)

        Dim vQ5 As String = ""
        'dr("Q5Other") = "" '選其他 才會帶入值
        If rdoQ51.Checked = True Then
            vQ5 = "1"
        ElseIf rdoQ52.Checked = True Then
            vQ5 = "2"
            'dr("Q5Other") = Convert.ToString(txtQ5Other.Text) '選其他 才會帶入值
        ElseIf rdoQ53.Checked = True Then
            vQ5 = "3"
        ElseIf rdoQ50.Checked = True Then
            vQ5 = "0"
        End If
        dr("Q5") = vQ5
        dr("Q5Other") = If(vQ5 = "2", txtQ5Other.Text, Convert.DBNull) '選其他 才會帶入值
        dr("Q5Note") = Convert.ToString(txtQ5Note.Text)

        dr("Q6") = Convert.ToString(txtQ6.Text)
        dr("Q6Other") = Convert.ToString(txtQ6Other.Text)
        dr("Q7Note") = Convert.ToString(txtQ7Note.Text)
        dr("Q8Note") = Convert.ToString(txtQ8Note.Text)
        dr("UnitName") = Convert.ToString(txtUnitName.Text)
        dr("FillerName") = Convert.ToString(txtFillerName.Text)

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Try
            DbAccess.UpdateDataTable(dt, da)
            Common.RespWrite(Me, "<script>alert('儲存成功');</script>")
            panelSch.Visible = True
            panelEdit.Visible = False
            btnSearch_Click(btnSearch, e)
        Catch ex As Exception
            'Common.RespWrite(Me, ex)
            Throw ex
        End Try
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        panelSch.Visible = False
        panelEdit.Visible = True
        btnChoice.Visible = True
        btnClear.Visible = True
        panelImg.Visible = True
        hidChkEdit.Value = ""

        txtTBplan2.Text = ""
        hidEditRID.Value = ""
        hidEditOrgID.Value = ""
        hidEditPlanID.Value = ""
        txtVDate.Text = ""
        rdoQ11.Checked = False
        rdoQ10.Checked = False
        txtQ1Per.Text = ""
        txtQ1A.Text = ""
        txtQ1B.Text = ""
        txtQ1Note.Text = ""
        rdoQ21.Checked = False
        rdoQ22.Checked = False
        rdoQ20.Checked = False
        txtQ2Other.Text = ""
        txtQ2Note.Text = ""
        rdoQ31.Checked = False
        rdoQ32.Checked = False
        rdoQ30.Checked = False
        txtQ3Other.Text = ""
        txtQ3Note.Text = ""
        rdoQ41.Checked = False
        rdoQ42.Checked = False
        rdoQ40.Checked = False
        txtQ4Other.Text = ""
        txtQ4Note.Text = ""
        rdoQ51.Checked = False
        rdoQ52.Checked = False
        rdoQ53.Checked = False
        rdoQ50.Checked = False
        txtQ5Other.Text = ""
        txtQ5Note.Text = ""
        txtQ6.Text = ""
        txtQ6Other.Text = ""
        txtQ7Note.Text = ""
        txtQ8Note.Text = ""
        txtUnitName.Text = ""
        txtFillerName.Text = ""
    End Sub

    Private Sub btnERpt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnERpt.Click
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "CP_01_009empty", "")
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "CP_01_009empty2", "")
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        panelSch.Visible = True
        panelEdit.Visible = False
        btnChoice.Visible = True
        btnClear.Visible = True
        panelImg.Visible = True
    End Sub
End Class