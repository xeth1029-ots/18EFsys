Partial Class LevPlan2
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then Call ShowFunction()
    End Sub

    Sub ShowFunction()
        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow

        Years.Items.Clear()
        sql = " SELECT DISTINCT years FROM ID_Plan where years <> ' ' ORDER BY 1 "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            'location.href='Secret_Login.aspx';
            Common.RespWrite(Me, "<script>alert('尚無計畫可登入');</script>")
        Else
            With Years
                .DataSource = dt
                .DataTextField = "Years"
                .DataValueField = "Years"
                .DataBind()
            End With
            Dim vYears As String = CStr(Now.Year)
            If Not Request.Cookies("LoginYears") Is Nothing Then vYears = RSA20031.AesDecrypt2(Request.Cookies("LoginYears").Value)
            Common.SetListItem(Years, vYears)
        End If

        drpDist.Items.Clear()
        sql = " SELECT * FROM ID_District WHERE DistID <> '000' ORDER BY 1 "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.RespWrite(Me, "<script>alert('尚無計畫可登入');</script>")
        Else
            With drpDist
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "DistID"
                .DataBind()
            End With
            Me.drpDist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            'If Not Request.Cookies("LoginDistID") Is Nothing Then Common.SetListItem(drpDist, Request.Cookies("LoginDistID").Value)
        End If
    End Sub

    Sub ShowPlanID()
        txtPlan.Items.Clear()
        Dim v_drpDist As String = TIMS.GetListValue(drpDist)
        Dim v_Years As String = TIMS.GetListValue(Years)
        If v_drpDist = "" Then Exit Sub
        If v_Years = "" Then Exit Sub
        Dim sql As String
        Dim dt As DataTable
        sql = ""
        sql &= " SELECT * FROM dbo.VIEW_LOGINPLAN "
        sql += " WHERE PlanID IN (SELECT PlanID FROM ID_Plan WHERE Years = '" & v_Years & "') AND DistID = '" & v_drpDist & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        With txtPlan
            .DataSource = dt
            .DataTextField = "PlanName"
            .DataValueField = "PlanID"
            .DataBind()
        End With
        Me.txtPlan.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'Clear Gov Org
        drpGov.Items.Clear()
        drpOrg.Items.Clear()
    End Sub

    Function GetRIDValue(ByVal PlanID As String, ByVal OrgID As String) As String
        Dim RID As String = ""
        PlanID = TIMS.ClearSQM(PlanID)
        OrgID = TIMS.ClearSQM(OrgID)
        If OrgID = "" Then Return RID
        If PlanID = "" Then Return RID

        Dim sql As String = ""
        'Dim dt As DataTable
        sql = ""
        sql += " SELECT RID "
        sql += " FROM AUTH_RELSHIP "
        sql += " WHERE PlanID = " & PlanID & " AND OrgID= " & OrgID
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then RID = Convert.ToString(dr("RID")) '.ToString
        Return RID
    End Function

    Function GetTPlanID(ByVal PlanID As String) As String
        Dim TPlanID As String = ""
        PlanID = TIMS.ClearSQM(PlanID)
        If PlanID = "" Then Return TPlanID
        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        sql = " SELECT * FROM VIEW_LOGINPLAN WHERE PLANID =@PLANID" '" & PlanID & "' "
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PLANID", SqlDbType.VarChar).Value = PlanID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Return TPlanID
        dr = dt.Rows(0) ' DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then TPlanID = dr.Item("TPlanID").ToString()
        Return TPlanID
    End Function

    Sub txtPlanSelectedChange()
        Dim sql As String = ""
        Dim dt As DataTable
        'Dim TPlanID, PlanID As String
        'PlanID = txtPlan.SelectedValue
        Me.PlanIDValue.Value = TIMS.GetListValue(txtPlan) '.SelectedValue
        Me.PlanIDValue.Value = TIMS.ClearSQM(Me.PlanIDValue.Value)
        Dim v_PlanID As String = Me.PlanIDValue.Value
        Me.RIDValue.Value = ""
        Me.drpGov.Items.Clear()
        Me.drpOrg.Items.Clear()
        'drpGov.Visible = False
        Dim TPlanID As String = GetTPlanID(PlanIDValue.Value)
        If TPlanID = "17" Then '政府補助計劃須列示政府補助單位
            Dim v_drpDist As String = TIMS.GetListValue(drpDist)
            sql = "" & vbCrLf
            sql += " SELECT DISTINCT d.orgid, d.orgname " & vbCrLf
            sql += " FROM AUTH_RELSHIP a " & vbCrLf
            sql += " JOIN (SELECT * FROM id_plan WHERE DistID = '" & v_drpDist & "' AND  years = '" & TIMS.GetListValue(Years) & "') b ON a.planid = b.planid " & vbCrLf
            sql += " JOIN (SELECT * FROM key_plan WHERE tplanid = '17') c ON b.tplanid = c.tplanid " & vbCrLf
            sql += " JOIN (SELECT * FROM org_orginfo WHERE ISCONUNIT=1) d ON d.orgid = a.orgid " & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)
            With drpGov
                .DataSource = dt
                .DataTextField = "OrgName"
                .DataValueField = "OrgID"
                .DataBind()
            End With
            'drpGov.Visible = True
            Me.drpGov.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Else
            'drpOrg.Items.Clear()
            Dim v_drpDist As String = TIMS.GetListValue(drpDist)
            If v_drpDist <> "" And v_PlanID <> "" Then Call ShowOrgID("", v_drpDist, v_PlanID)
        End If
    End Sub

    Sub ShowOrgID(ByVal RIDValue As String, ByVal DistValue As String, Optional ByVal PlanID As String = "")
        Dim sql As String = ""
        Dim dt As DataTable
        RIDValue = TIMS.ClearSQM(RIDValue)
        PlanID = TIMS.ClearSQM(PlanID)
        DistValue = TIMS.ClearSQM(DistValue)
        If RIDValue <> "" And PlanID = "" Then
            sql = "" & vbCrLf
            sql &= " SELECT a.orgid ,a.orgname ,b.DistID " & vbCrLf
            sql += " FROM Org_OrgInfo a " & vbCrLf
            sql += " JOIN (SELECT RSID ,DistID ,OrgID FROM AUTH_RELSHIP WHERE (PlanID=0 AND DistID='" & DistValue & "') OR RID LIKE '" & RIDValue & "%') b ON a.OrgID = b.OrgID " & vbCrLf
            sql += " JOIN Org_OrgPlanInfo c ON b.RSID = c.RSID " & vbCrLf
            sql += " ORDER BY b.DistID " & vbCrLf
        ElseIf DistValue <> "000" Then 'PlanID <> ""
            sql = "" & vbCrLf
            sql &= " SELECT a.orgid ,a.orgname ,b.DistID " & vbCrLf
            sql += " FROM Org_OrgInfo a " & vbCrLf
            sql += " JOIN (SELECT RSID ,DistID ,OrgID FROM AUTH_RELSHIP WHERE DistID = '" & DistValue & "' AND PlanID = " & PlanID & ") b ON a.OrgID = b.OrgID " & vbCrLf
            sql += " JOIN Org_OrgPlanInfo c ON b.RSID = c.RSID " & vbCrLf
            sql += " ORDER BY b.DistID " & vbCrLf
        ElseIf DistValue = "000" Then 'PlanID <> ""
            sql = "" & vbCrLf
            sql &= " SELECT a.orgid ,a.orgname ,b.DistID " & vbCrLf
            sql += " FROM Org_OrgInfo a " & vbCrLf
            sql += " JOIN (SELECT RSID,DistID,OrgID FROM AUTH_RELSHIP WHERE PlanID = " & PlanID & ") b ON a.OrgID = b.OrgID " & vbCrLf
            sql += " JOIN Org_OrgPlanInfo c ON b.RSID = c.RSID " & vbCrLf
            sql += " ORDER BY b.DistID " & vbCrLf
        End If
        If sql <> "" Then
            dt = DbAccess.GetDataTable(sql, objconn)
            With drpOrg
                .DataSource = dt
                .DataTextField = "OrgName"
                .DataValueField = "OrgID"
                .DataBind()
            End With
            Me.drpOrg.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            'Me.drpOrg.Items.Insert(0, New ListItem("=(機構不存在)另新增機構=", "-1"))
            'If PlanID = "" Then Common.SetListItem(drpOrg, drpGov.SelectedValue)
        End If
    End Sub

    Sub clearForm()
        'drpDist.SelectedValue = ""
        Common.SetListItem(drpDist, "")
        PlanIDValue.Value = ""
        RIDValue.Value = ""
        OrgIDValue.Value = ""
        txtPlan.Items.Clear()
        drpGov.Items.Clear()
        drpOrg.Items.Clear()
    End Sub

    Private Sub Years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Years.SelectedIndexChanged
        Me.YearsValue.Value = TIMS.GetListValue(Years) '.SelectedValue
        ShowPlanID()
        'ShowPlanIDdetail(Years.SelectedValue, drpDist.SelectedValue)
    End Sub

    Private Sub drpDist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles drpDist.SelectedIndexChanged
        Me.DistValue.Value = TIMS.GetListValue(drpDist) '.SelectedValue
        ShowPlanID()
        'ShowPlanIDdetail(Years.SelectedValue, drpDist.SelectedValue)
    End Sub

    Private Sub txtPlan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPlan.SelectedIndexChanged
        txtPlanSelectedChange()
    End Sub

    Private Sub drpGov_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles drpGov.SelectedIndexChanged
        RIDValue.Value = GetRIDValue(TIMS.GetListValue(txtPlan), TIMS.GetListValue(drpGov))
        drpOrg.Items.Clear()
        Dim v_drpDist As String = TIMS.GetListValue(drpDist)
        If RIDValue.Value <> "" AndAlso v_drpDist <> "" Then Call ShowOrgID(RIDValue.Value, v_drpDist)
    End Sub

    Private Sub clear_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear.ServerClick
        clearForm()
    End Sub
End Class