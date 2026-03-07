Partial Class SD_02_004_add
    Inherits AuthBasePage

    Const cst_chkTypeClass As String = "Class"
    Const cst_chkTypeOrg As String = "Org"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not Session("_SearchStr") Is Nothing Then
            Session("_SearchStr") = Session("_SearchStr")
        End If

        If Not IsPostBack Then
            btnSave.Attributes("onclick") = "return checkData();"

            RIDValue.Value = TIMS.ClearSQM(Request("RID")) '單位Rid
            OrgID.Value = TIMS.ClearSQM(Request("OrgID"))
            OCID.Value = TIMS.ClearSQM(Request("OCID"))
            'SetingYear = Request("Year").ToString  '設定年份
            PlanYear.Text = TIMS.ClearSQM(Request("Year"))
            hidDistID.Value = TIMS.ClearSQM(Request("DistID")) '機構設定用
            hidPlanID.Value = TIMS.ClearSQM(Request("PlanID")) '機構設定用

            hidDistID.Value = IIf(hidDistID.Value = "", sm.UserInfo.DistID, hidDistID.Value)
            hidPlanID.Value = IIf(hidPlanID.Value = "", sm.UserInfo.PlanID, hidPlanID.Value)

            TR_Class.Visible = False
            If OCID.Value <> "" Then '班級
                TR_Class.Visible = True
            End If

            Call LoadOrgClassName()

            Call LoadData()
        End If
    End Sub

    Sub LoadOrgClassName()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim sqlS As String = ""
        ClassName.Text = ""
        OrgID.Value = ""
        OrgName.Text = ""
        DistID.Text = ""
        If OCID.Value <> "" Then
            '班級
            Dim drCC As DataRow = TIMS.GetOCIDDate(OCID.Value, objconn)
            If drCC Is Nothing Then Exit Sub
            ClassName.Text = Convert.ToString(drCC("ClassCName2"))
            OrgID.Value = Convert.ToString(drCC("OrgID"))
            OrgName.Text = Convert.ToString(drCC("OrgName"))
            DistID.Text = TIMS.Get_DistName1(drCC("DistID"))
        Else
            '訓練機構
            sqlS = ""
            sqlS &= " select a.DistID,b.OrgID,b.Orgname "
            sqlS &= " from Auth_Relship a"
            sqlS &= " join Org_orginfo b on a.orgid=b.orgid"
            sqlS &= " where a.RID=@RID "
            Dim sCmd2 As New SqlCommand(sqlS, objconn)
            Dim dt2 As New DataTable
            With sCmd2
                .Parameters.Clear()
                .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                'dt2.Load(.ExecuteReader())
                dt2 = DbAccess.GetDataTable(sCmd2.CommandText, objconn, sCmd2.Parameters)
            End With
            If dt2.Rows.Count > 0 Then
                Dim dr2 As DataRow = dt2.Rows(0)
                OrgID.Value = Convert.ToString(dr2("OrgID"))
                OrgName.Text = Convert.ToString(dr2("OrgName"))
                DistID.Text = TIMS.Get_DistName1(Convert.ToString(dr2("DistID")))
            End If
        End If

        'Dim drR2 As DataRow = TIMS.GET_RELSHIP23(RIDValue.Value, sm.UserInfo.PlanID, objconn)
        Dim drR2 As DataRow = TIMS.GET_RELSHIP23(RIDValue.Value, hidPlanID.Value, objconn)
        If Not drR2 Is Nothing Then
            CtrlOrg.Text = drR2("ORGNAME2")
            ParentOrgID.Value = drR2("ORGID2")
            ParentRID.Value = drR2("RID2")
        End If
    End Sub

    '載入資料
    Sub LoadData()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing

        If OCID.Value <> "" Then 'by 班級
            'dt = Get_FinalComment(cst_chkTypeClass,
            '                      OrgID.Value, OCID.Value,
            '                      CStr(sm.UserInfo.PlanID), Nothing)
            dt = Get_FinalComment(cst_chkTypeClass,
                                  OrgID.Value, OCID.Value,
                                  hidPlanID.Value, Nothing)
        Else
            'by 機構(分區) -- RID
            'dt = Get_FinalComment(cst_chkTypeOrg,
            '                      OrgID.Value, "",
            '                      CStr(sm.UserInfo.PlanID), Nothing)
            dt = Get_FinalComment(cst_chkTypeOrg,
                                  OrgID.Value, "",
                                  hidPlanID.Value, Nothing)
        End If
        Dim flagSchSetRP1 As Boolean = False '查無有效資料為true
        If dt Is Nothing Then flagSchSetRP1 = True
        If Not flagSchSetRP1 AndAlso Not dt Is Nothing Then
            If dt.Rows.Count = 0 Then flagSchSetRP1 = True
        End If

        'mark 下面這行，如果是先設定單位的再設定班，這行就永遠撈不到班的設定結果
        'If Not dt Is Nothing Then flagSchSetRP1 = True

        If flagSchSetRP1 Then
            'dt = Get_OrgComment(OrgID.Value, CStr(sm.UserInfo.PlanID))
            dt = Get_OrgComment(OrgID.Value, hidPlanID.Value)
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                Notice.Value = Convert.ToString(dr("Notice"))
            End If
        Else
            dr = dt.Rows(0)
            Notice.Value = Convert.ToString(dr("Notice"))
        End If
    End Sub

    '依需求不同計畫機構須有不同甄試通知單說明事項設定
    Function Get_OrgComment(ByVal OrgID As String,
                            ByVal PlanID As String) As DataTable
        Dim dt As New DataTable
        If PlanID = "" Then Return dt
        If OrgID = "" Then Return dt

        Dim sqls As String = ""
        sqls = ""
        sqls &= " select PlanID,OrgID,Notice "
        sqls &= " FROM ORG_NOTICE "
        sqls &= " where 1=1 "
        sqls &= " and PlanID=@PlanID "
        sqls &= " and OrgID=@OrgID "
        sqls &= " and OCID is null "  '針對單位設定OCID不能有值

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PlanID", PlanID)
        parms.Add("OrgID", OrgID)

        dt = DbAccess.GetDataTable(sqls, objconn, parms)
        Return dt
    End Function

    '20090601 by Jimmy 取得「班級」甄試通知單說明事項設定
    Function Get_ClassComment(ByVal OrgID As String,
                              ByVal OCID As String,
                              ByVal PlanID As String) As DataTable
        Dim dt As New DataTable
        If PlanID = "" Then Return dt
        If OrgID = "" Then Return dt
        If OCID = "" Then Return dt

        Dim sqls As String = ""
        sqls = ""
        sqls &= " select PlanID,OrgID,Notice"
        sqls &= " from Org_Notice   "
        sqls &= " where PlanID=@PlanID "
        sqls &= " and OrgID=@OrgID "
        sqls &= " and OCID=@OCID "
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PlanID", PlanID)
        parms.Add("OrgID", OrgID)
        parms.Add("OCID", TIMS.GetValue1(OCID))
        dt = DbAccess.GetDataTable(sqls, objconn, parms)
        Return dt
    End Function

    Function Get_FinalComment(ByVal chkType As String,
                              ByVal OrgID As String, ByVal OCID As String,
                              ByVal PlanID As String, ByVal retDt As DataTable) As DataTable
        Dim dr As DataRow = Nothing
        Dim chkflag As Boolean = False
        Select Case chkType
            Case cst_chkTypeClass 'by 班級
                retDt = Get_ClassComment(OrgID, OCID, PlanID)
            Case cst_chkTypeOrg   'by 機構(分區)
                retDt = Get_OrgComment(OrgID, PlanID) 'for 不同年度、計畫、轄區
        End Select
        If Not retDt Is Nothing Then
            If retDt.Rows.Count > 0 Then
                dr = retDt.Rows(0)
                If Convert.ToString(dr("Notice")) <> "" Then
                    chkflag = True
                End If
            End If
        End If
        If Not chkflag Then
            Select Case chkType
                Case cst_chkTypeClass '班級沒有說明事項時，繼續向上找單位
                    retDt = Get_OrgComment(ParentOrgID.Value, PlanID)
                Case cst_chkTypeOrg
                    retDt = Nothing
            End Select
        End If
        Return retDt
    End Function

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        'Session("_SearchStr") = Me.ViewState("_SearchStr")
        If SaveData() Then
            Page.RegisterStartupScript("", "<script>alert('甄試通知說明設定-成功!'); window.location.href='SD_02_004.aspx?ID=" & Request("ID") & "';</script>")
        End If
    End Sub

    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "SD_02_004.aspx?ID=" & Request("ID"))
    End Sub

    Private Function SaveData() As Boolean
        Dim retVal As Boolean = True
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Return False '異常

        Dim sqls As String = ""
        Dim da As New SqlDataAdapter
        'Dim tmpComment As String = ""
        Dim dt As DataTable = Nothing
        Call TIMS.OpenDbConn(objconn)
        Dim iSql As String = ""
        Dim uSql As String = ""
        Dim dSql As String = ""
        Dim parms As Hashtable = New Hashtable()
        Dim newNoticeID As Integer = 0

        If OCID.Value <> "" Then '班級   '090907 andy test
            '新增
            iSql = " insert into Org_Notice(NOTICEID,PlanID,OrgID,OCID,Notice,MODIFYACCT,MODIFYDATE) "
            iSql &= " values (@NOTICEID,@PlanID,@OrgID,@OCID,@Notice,@MODIFYACCT,getdate())"
            'da.InsertCommand = New SqlCommand(sqls, objconn)
            '刪除
            dSql = " delete Org_Notice"
            dSql &= " where 1=1 "
            dSql &= " and PlanID=@PlanID "
            dSql &= " and OrgID=@OrgID "
            dSql &= " and OCID=@OCID "
            dSql &= " and OCID  is not null"
            'da.DeleteCommand = New SqlCommand(sqls, objconn)
            '修改
            uSql = "  update Org_Notice "
            uSql &= " set Notice=@Notice "
            uSql &= " ,MODIFYACCT=@MODIFYACCT "
            uSql &= " ,MODIFYDATE=getdate() "
            uSql &= " where 1=1 "
            uSql &= " and PlanID=@PlanID "
            uSql &= " and OrgID=@OrgID "
            uSql &= " and OCID=@OCID "
            uSql &= " and OCID  is not null"
            'da.UpdateCommand = New SqlCommand(sqls, objconn)

            'dt = Get_ClassComment(OrgID.Value, OCID.Value, Convert.ToString(sm.UserInfo.PlanID))
            dt = Get_ClassComment(OrgID.Value, OCID.Value, hidPlanID.Value)
            If dt.Rows.Count = 0 Then
                '新增
                'With da.InsertCommand
                '    .Parameters.Clear()
                '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.PlanID)
                '    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.OrgID)
                '    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID.Value
                '    .Parameters.Add("Notice", SqlDbType.NVarChar).Value = TIMS.GetValue1(Notice.Value)
                '    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = CStr(sm.UserInfo.UserID)
                '    .ExecuteNonQuery()
                'End With

                newNoticeID = DbAccess.GetNewId(objconn, "ORG_NOTICE_NOTICEID_SEQ, ORG_NOTICE, NOTICEID")

                parms.Clear()
                parms.Add("NOTICEID", newNoticeID)
                'parms.Add("PlanID", Convert.ToString(sm.UserInfo.PlanID))
                parms.Add("PlanID", hidPlanID.Value)
                'parms.Add("OrgID", Convert.ToString(sm.UserInfo.OrgID))
                parms.Add("OrgID", OrgID.Value) '署可維護分署的資料
                parms.Add("OCID", OCID.Value)
                parms.Add("Notice", TIMS.GetValue1(Notice.Value))
                parms.Add("MODIFYACCT", Convert.ToString(sm.UserInfo.UserID))
                ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                DbAccess.ExecuteNonQuery(iSql, objconn, parms)
            Else
                If Notice.Value = "" Then
                    '異動'刪除
                    'With da.DeleteCommand
                    '    .Parameters.Clear()
                    '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.PlanID)
                    '    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.OrgID)
                    '    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID.Value
                    '    .ExecuteNonQuery()
                    'End With

                    parms.Clear()
                    'parms.Add("PlanID", Convert.ToString(sm.UserInfo.PlanID))
                    parms.Add("PlanID", hidPlanID.Value)
                    'parms.Add("OrgID", Convert.ToString(sm.UserInfo.OrgID))
                    parms.Add("OrgID", OrgID.Value) '署可維護分署的資料
                    parms.Add("OCID", OCID.Value)
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(dSql, objconn, parms)
                Else
                    '異動'修改
                    'With da.UpdateCommand
                    '    .Parameters.Clear()
                    '    .Parameters.Add("Notice", SqlDbType.NVarChar).Value = Notice.Value
                    '    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = CStr(sm.UserInfo.UserID)
                    '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.PlanID)
                    '    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.OrgID)
                    '    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID.Value
                    '    .ExecuteNonQuery()
                    'End With

                    parms.Clear()
                    parms.Add("Notice", TIMS.GetValue1(Notice.Value))
                    parms.Add("MODIFYACCT", Convert.ToString(sm.UserInfo.UserID))
                    'parms.Add("PlanID", Convert.ToString(sm.UserInfo.PlanID))
                    parms.Add("PlanID", hidPlanID.Value)
                    'parms.Add("OrgID", Convert.ToString(sm.UserInfo.OrgID))
                    parms.Add("OrgID", OrgID.Value) '署可維護分署的資料
                    parms.Add("OCID", OCID.Value)
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(uSql, objconn, parms)
                End If
            End If
        Else
            '新增
            iSql = " insert into Org_Notice(NOTICEID,PlanID,OrgID,Notice,MODIFYACCT,MODIFYDATE) "
            iSql &= " values (@NOTICEID,@PlanID,@OrgID,@Notice,@MODIFYACCT,getdate())"
            'da.InsertCommand = New SqlCommand(sqls, objconn)

            '刪除
            dSql = " delete Org_Notice "
            dSql &= " where 1=1 "
            dSql &= " and PlanID=@PlanID "
            dSql &= " and OrgID=@OrgID "
            dSql &= " and OCID is null"
            'da.DeleteCommand = New SqlCommand(sqls, objconn)

            '修改
            uSql = "  update Org_Notice "
            uSql &= " set Notice=@Notice "
            uSql &= " ,MODIFYACCT=@MODIFYACCT "
            uSql &= " ,MODIFYDATE=getdate() "
            uSql &= " where 1=1 "
            uSql &= " and PlanID=@PlanID "
            uSql &= " and OrgID=@OrgID "
            uSql &= " and OCID is null"
            'da.UpdateCommand = New SqlCommand(sqls, objconn)

            'dt = Get_OrgComment(OrgID.Value, CStr(sm.UserInfo.PlanID))
            dt = Get_OrgComment(OrgID.Value, hidPlanID.Value)
            If dt.Rows.Count = 0 Then
                '新增
                'With da.InsertCommand
                '    .Parameters.Clear()
                '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.PlanID)
                '    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.OrgID)
                '    .Parameters.Add("Notice", SqlDbType.NVarChar).Value = TIMS.GetValue1(Notice.Value)
                '    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = CStr(sm.UserInfo.UserID)
                '    .ExecuteNonQuery()
                'End With

                newNoticeID = DbAccess.GetNewId(objconn, "ORG_NOTICE_NOTICEID_SEQ, ORG_NOTICE, NOTICEID")

                parms.Clear()
                parms.Add("NOTICEID", newNoticeID)
                'parms.Add("PlanID", Convert.ToString(sm.UserInfo.PlanID))
                parms.Add("PlanID", hidPlanID.Value)
                'parms.Add("OrgID", Convert.ToString(sm.UserInfo.OrgID))
                parms.Add("OrgID", OrgID.Value) '署可維護分署的資料
                parms.Add("Notice", TIMS.GetValue1(Notice.Value))
                parms.Add("MODIFYACCT", Convert.ToString(sm.UserInfo.UserID))
                ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                DbAccess.ExecuteNonQuery(iSql, objconn, parms)
            Else
                If Notice.Value = "" Then
                    '異動'刪除
                    'With da.DeleteCommand
                    '    .Parameters.Clear()
                    '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.PlanID)
                    '    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.OrgID)
                    '    .ExecuteNonQuery()
                    'End With

                    parms.Clear()
                    'parms.Add("PlanID", Convert.ToString(sm.UserInfo.PlanID))
                    parms.Add("PlanID", hidPlanID.Value)
                    'parms.Add("OrgID", Convert.ToString(sm.UserInfo.OrgID))
                    parms.Add("OrgID", OrgID.Value) '署可維護分署的資料
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(dSql, objconn, parms)
                Else
                    '異動'修改
                    'With da.UpdateCommand
                    '    .Parameters.Clear()
                    '    .Parameters.Add("Notice", SqlDbType.NVarChar).Value = Notice.Value
                    '    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = CStr(sm.UserInfo.UserID)
                    '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.PlanID)
                    '    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = CStr(sm.UserInfo.OrgID)
                    '    .ExecuteNonQuery()
                    'End With

                    parms.Clear()
                    parms.Add("Notice", TIMS.GetValue1(Notice.Value))
                    parms.Add("MODIFYACCT", Convert.ToString(sm.UserInfo.UserID))
                    'parms.Add("PlanID", Convert.ToString(sm.UserInfo.PlanID))
                    parms.Add("PlanID", hidPlanID.Value)
                    'parms.Add("OrgID", Convert.ToString(sm.UserInfo.OrgID))
                    parms.Add("OrgID", OrgID.Value) '署可維護分署的資料
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(uSql, objconn, parms)
                End If
            End If
        End If

        da = Nothing
        'da2 = Nothing
        Return retVal '正常結束
    End Function
End Class
