Partial Class SD_01_004_email_add
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            btnSave.Attributes("onclick") = "return checkData();"
            RIDValue.Value = Me.Request("RID").ToString  '單位Rid
            OrgID.Value = Me.Request("OrgID")
            OCID.Value = Me.Request("OCID")
            'SetingYear = Me.Request("Year").ToString  '設定年份
            PlanYear.Text = Me.Request("Year").ToString
            If Convert.ToString(OCID.Value) <> "" Then '班級
                TR_Class.Visible = True
            Else
                TR_Class.Visible = False
            End If
            LoadOrgClassName()
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing
            Me.ViewState("RID") = RIDValue.Value
            LoadData()
        End If
    End Sub

    Sub LoadOrgClassName()
        Dim da As New SqlDataAdapter
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim dr As DataRow = Nothing

        Dim sqlStr As String = ""
        If OCID.Value.ToString <> "" Then  '班級
            sqlStr = ""
            sqlStr &= " SELECT a.rid, a.ocid, f.DistID, e.OrgName, e.OrgID"
            sqlStr &= " ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) ClassCName" & vbCrLf
            sqlStr &= " FROM Class_ClassInfo a "
            sqlStr &= " JOIN Auth_Relship f ON a.RID = f.RID "
            sqlStr &= " JOIN Org_OrgInfo e ON f.OrgID = e.Orgid "
            sqlStr &= " WHERE a.ocid = @OCID AND f.rid = @RID "
        Else  '訓練機構
            sqlStr = ""
            sqlStr &= " SELECT a.DistID, b.OrgID, b.Orgname "
            sqlStr &= " FROM Auth_Relship a "
            sqlStr &= " JOIN Org_orginfo b ON a.orgid = b.orgid "
            sqlStr &= " WHERE a.RID = @RID "
        End If

        Try
            With da
                da.SelectCommand = New SqlCommand(sqlStr, objconn)
                If OCID.Value.ToString <> "" Then   '班級
                    da.SelectCommand.Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID.Value.ToString
                    da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value.ToString
                Else
                    da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value.ToString
                End If
                .Fill(ds, "QueryTB")
                dt = ds.Tables("QueryTB")
            End With
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString())
        End Try

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            If Convert.ToString(OCID.Value) <> "" Then ClassName.Text = Convert.ToString(dr("ClassCName")) '班級
            Me.OrgID.Value = Convert.ToString(dr("OrgID"))
            OrgName.Text = Convert.ToString(dr("OrgName"))
            Me.DistID.Text = TIMS.Get_DistName1(Convert.ToString(dr("DistID")))
        End If

        TR_CtrlOrg.Visible = False
        Dim ParentName As String = TIMS.Get_ParentRID(RIDValue.Value, objconn)
        If ParentName <> "" Then
            TR_CtrlOrg.Visible = True
            CtrlOrg.Text = ParentName
        End If
    End Sub

    Private Sub LoadData()
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing

        Try
            If OCID.Value <> "" Then 'by 班級
                dt = TIMS.Get_FinalEComment("Class", Convert.ToString(OrgID.Value), Convert.ToString(OCID.Value), Convert.ToString(RIDValue.Value), Convert.ToString(Me.sm.UserInfo.DistID), Convert.ToString(Me.sm.UserInfo.PlanID), Nothing, objconn)
            Else  'by 機構(分區) -- RID
                dt = TIMS.Get_FinalEComment("Org", Convert.ToString(OrgID.Value), "", Convert.ToString(RIDValue.Value), Convert.ToString(Me.sm.UserInfo.DistID), Convert.ToString(Me.sm.UserInfo.PlanID), Nothing, objconn)
            End If

            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Me.OrgID.Value = Convert.ToString(dr("OrgID"))
                Me.eComment.Value = Convert.ToString(dr("eComment"))
            Else
                Me.eComment.Value = ""
                dt = TIMS.Get_OrgEComment(OrgID.Value, RIDValue.Value, Convert.ToString(Me.sm.UserInfo.PlanID), objconn)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then dr = dt.Rows(0)
            End If
        Catch ex As Exception
            Page.RegisterStartupScript("errmsg", "<script>alert('錯誤訊息：" & ex.Message.ToString & "');</script>")
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim Errmsg As String = ""
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        If SaveData(Errmsg) Then
            Page.RegisterStartupScript("", "<script>alert('e網審核郵件設定-成功!'); window.location.href='SD_01_004_email.aspx?ID=" & Request("ID") & "';</script>")
        Else
            Page.RegisterStartupScript("errmsg", "<script>alert('" & Errmsg & "');</script>")
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        'Page.RegisterStartupScript("", "<script> window.location.href='SD_01_004_email.aspx?ID=" & Request("ID") & "';</script>")
        TIMS.Utl_Redirect(Me, objconn, "SD_01_004_email.aspx?ID=" & Request("ID"))
    End Sub

#Region "(No Use)"

    'Private Function Get_DistName(ByVal DistID As String) As String
    '    Dim retVal As String = ""

    '    Select Case Convert.ToString(DistID)
    '        Case "000"
    '            retVal = "職訓局"
    '        Case "001"
    '            retVal = "北區職訓中心"
    '        Case "002"
    '            retVal = "泰山職訓中心"
    '        Case "003"
    '            retVal = "桃園職訓中心"
    '        Case "004"
    '            retVal = "中區職訓中心"
    '        Case "005"
    '            retVal = "台南職訓中心"
    '        Case "006"
    '            retVal = "南區職訓中心"
    '    End Select

    '    Return retVal
    'End Function

#End Region

    Private Function SaveData(ByRef Errmsg As String) As Boolean
        Dim retVal As Boolean = False
        Errmsg = ""
        OCID.Value = TIMS.ClearSQM(OCID.Value)
        eComment.Value = TIMS.ClearSQM(eComment.Value)
        If eComment.Value.Length > 500 Then eComment.Value = Mid(eComment.Value, 1, 500)

        Dim sql As String = ""
        sql = " UPDATE Class_ClassInfo SET eComment = @eComment WHERE OCID = @OCID "
        'Dim uCmd1 As New SqlCommand(sql, objconn)
        Dim uSql1 As String = sql

        sql = ""
        sql &= " UPDATE Org_eComment "
        sql &= " SET eComment = @eComment ,ModifyAcct = @ModifyAcct ,ModifyDate = getdate() "
        sql &= " WHERE 1=1 "
        sql &= " AND PlanID = @PlanID AND OrgID = @OrgID "
        'Dim uCmd2 As New SqlCommand(sql, objconn)
        Dim uSql2 As String = sql

        sql = ""
        sql &= " INSERT INTO Org_eComment (PlanID, OrgID, eComment, ModifyAcct, ModifyDate) "
        sql &= " VALUES (@PlanID, @OrgID, @eComment, @ModifyAcct, getdate()) "
        'Dim iCmd2 As New SqlCommand(sql, objconn)
        Dim iSql2 As String = sql

        Call TIMS.OpenDbConn(objconn)

        If Convert.ToString(OCID.Value) <> "" Then '班級
            '.Parameters.Clear()
            Dim myParam As Hashtable = New Hashtable
            myParam.Add("eComment", If(eComment.Value <> "", eComment.Value, Convert.DBNull))
            myParam.Add("OCID", OCID.Value)
            DbAccess.ExecuteNonQuery(uSql1, objconn, myParam)
            retVal = True 'Return True
        Else
            Dim dt As DataTable = Nothing
            dt = TIMS.Get_OrgEComment(OrgID.Value, (RIDValue.Value), Convert.ToString(Me.sm.UserInfo.PlanID), objconn)
            If dt Is Nothing Then Return False

            If dt.Rows.Count = 0 Then
                '新增
                Dim myParam As Hashtable = New Hashtable
                myParam.Add("PlanID", Convert.ToString(Me.sm.UserInfo.PlanID))
                myParam.Add("OrgID", Convert.ToString(OrgID.Value))
                myParam.Add("eComment", If(eComment.Value <> "", eComment.Value, Convert.DBNull))
                myParam.Add("ModifyAcct", Me.sm.UserInfo.UserID)
                DbAccess.ExecuteNonQuery(iSql2, objconn, myParam)
            Else
                '異動
                '.Parameters.Clear()
                Dim myParam As Hashtable = New Hashtable
                myParam.Add("eComment", If(eComment.Value <> "", eComment.Value, Convert.DBNull))
                myParam.Add("ModifyAcct", Me.sm.UserInfo.UserID)
                myParam.Add("PlanID", Convert.ToString(Me.sm.UserInfo.PlanID))
                myParam.Add("OrgID", Convert.ToString(OrgID.Value))
                DbAccess.ExecuteNonQuery(uSql2, objconn, myParam)
            End If
            retVal = True 'Return True
        End If
        Return retVal
    End Function
End Class