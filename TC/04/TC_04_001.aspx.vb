Partial Class TC_04_001
    Inherits AuthBasePage

    '若是產學訓計畫，則跳到 TC_04_002.aspx 執行審核動作

    Dim vMsg1 As String = ""

    Const cst_DG1_col編號 As Integer = 0
    Const cst_DG1_col計畫年度 As Integer = 1
    Const cst_DG1_col班別名稱 As Integer = 2
    Const cst_DG1_col統一編號 As Integer = 3
    Const cst_DG1_col訓練單位 As Integer = 4
    Const cst_DG1_col公司地址 As Integer = 5
    Const cst_DG1_col聯絡人 As Integer = 6
    Const cst_DG1_col電話 As Integer = 7
    Const cst_DG1_col審核 As Integer = 8
    Const cst_DG1_col原因 As Integer = 9

    Const cst_DG2_col序號 As Integer = 0
    Const cst_DG2_col申請日期 As Integer = 1
    Const cst_DG2_col訓練日期 As Integer = 2
    Const cst_DG2_col訓練迄日 As Integer = 3
    Const cst_DG2_col訓練機構 As Integer = 4
    Const cst_DG2_col班別名稱 As Integer = 5
    Const cst_DG2_col轉班 As Integer = 6
    Const cst_DG2_col取消審核 As Integer = 7

    'Dim blnCanAdds As Boolean=False '新增
    'Dim blnCanMod As Boolean=False '修改
    'Dim blnCanDel As Boolean=False '刪除
    'Dim blnCanSech As Boolean=False '查詢
    'Dim blnCanPrnt As Boolean=False '列印
    'Dim au As New cAUTH
    Dim Auth_Relship As DataTable
    Dim dtOrgBlack As DataTable  '取出系統現有黑名單
    Const cst_isBlackMsg As String = TIMS.cst_gBlackMsg1
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '一般計畫審核計畫專用
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '取出系統現有黑名單
        dtOrgBlack = TIMS.Get_OrgBlackList(Me, objconn)

        '若是產學訓計畫，則跳到 TC_04_002.aspx 執行審核動作
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            LabTMID.Text = "訓練業別"
            Server.Transfer("TC_04_002.aspx")
            Return ' Exit Sub
        End If

        PageControler2.PageDataGrid = DataGrid2

        '依登入者RID、TPlanID、Years、PlanID  取得 Auth_Relship.RID,OrgName
        Auth_Relship = TIMS.sUtl_GetAuthRelship(Me, objconn)

        bntAdd.Enabled = True
        btnQuery.Enabled = True

#Region "(No Use)"

        'bntAdd.Enabled=False
        'If au.blnCanAdds Then bntAdd.Enabled=True
        'btnQuery.Enabled=False
        'If au.blnCanSech Then btnQuery.Enabled=True

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable=sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow=FunDt.Select("FunID='" & Request("ID") & "'")
        '    FunDr=FunDrArray(0)
        '    If FunDr("Sech")=1 Then
        '        btnQuery.Enabled=True
        '    Else
        '        btnQuery.Enabled=False
        '    End If
        '    If FunDr("Adds")=1 Then
        '        bntAdd.Enabled=True
        '    Else
        '        bntAdd.Enabled=False
        '    End If
        'End If

#End Region

        If Not IsPostBack Then
            '取得訓練計畫
            TPlanid.Value = sm.UserInfo.TPlanID 'DbAccess.ExecuteScalar(Sqlstr, objconn)

            '(加強操作便利性)2005/4/1-melody
            RIDValue.Value = sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName  'orgname

            DataGridTable1.Visible = False
            DataGridTable2.Visible = False
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "return SavaData();"

        '帶入查詢參數
        If Not IsPostBack Then
            PlanMode = TIMS.Get_PlanMode(PlanMode)

            If Session("search") IsNot Nothing Then
                Dim s_sess_search As String = Convert.ToString(Session("search"))
                Session("search") = Nothing
                Dim MyValue As String = ""
                MyValue = TIMS.GetMyValue(s_sess_search, "prg")
                If MyValue = "TC_04_001" Then
                    TB_career_id.Text = TIMS.GetMyValue(s_sess_search, "TB_career_id")
                    trainValue.Value = TIMS.GetMyValue(s_sess_search, "trainValue")
                    jobValue.Value = TIMS.GetMyValue(s_sess_search, "jobValue")
                    txtCJOB_NAME.Text = TIMS.GetMyValue(s_sess_search, "txtCJOB_NAME")
                    cjobValue.Value = TIMS.GetMyValue(s_sess_search, "cjobValue")
                    center.Text = TIMS.GetMyValue(s_sess_search, "center")
                    RIDValue.Value = TIMS.GetMyValue(s_sess_search, "RIDValue")
                    UNIT_SDATE.Text = TIMS.GetMyValue(s_sess_search, "UNIT_SDATE")
                    UNIT_EDATE.Text = TIMS.GetMyValue(s_sess_search, "UNIT_EDATE")

                    MyValue = TIMS.GetMyValue(s_sess_search, "PlanMode")
                    'Common.SetListItem(PlanMode, MyValue)
                    PlanMode.SelectedIndex = MyValue
                    If PlanMode.SelectedIndex = 0 Then
                        'dgPlan.PageIndex=TIMS.GetMyValue(s_sess_search, "PageIndex")
                    Else
                        PageControler2.PageIndex = 0
                        'PageControler2.PageIndex=TIMS.GetMyValue(s_sess_search, "PageIndex")
                        MyValue = TIMS.GetMyValue(s_sess_search, "PageIndex")
                        If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                            MyValue = CInt(MyValue)
                            PageControler2.PageIndex = MyValue
                        End If
                    End If
                    'CurrentPageIndex
                    'btnQuery_Click(sender, e)
                    Call Search1()
                End If

            End If
        End If

        '確認機構是否為黑名單
        Dim vsMsg2 As String = ""
        If Chk_OrgBlackList(vsMsg2) Then
            bntAdd.Enabled = False
            TIMS.Tooltip(bntAdd, vsMsg2)
            Button1.Enabled = False
            TIMS.Tooltip(Button1, vsMsg2)

            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            isBlack.Value = "Y"
            orgname.Value = sm.UserInfo.OrgName
            'btnAdd.Visible=False 'Button8.Visible=False
        End If
        Return rst
    End Function

    '查詢 SQL '若有改正審核搜尋條件，請順更檢查 主頁訊息功能。
    Sub Search1()
        Dim flgROLEIDx0xLIDx0 As Boolean = TIMS.IsSuperUser(sm, 1)

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        UNIT_SDATE.Text = TIMS.Cdate3(UNIT_SDATE.Text)
        UNIT_EDATE.Text = TIMS.Cdate3(UNIT_EDATE.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(UNIT_SDATE.Text, UNIT_EDATE.Text) Then
            Dim T_DATE1 As String = UNIT_SDATE.Text
            UNIT_SDATE.Text = UNIT_EDATE.Text
            UNIT_EDATE.Text = T_DATE1
        End If

        Dim flag_BlackOpen As Boolean = False '預設 不啟用黑名單
        flag_BlackOpen = TIMS.Check_OrgBlackList(Me, "", objconn) '檢測計畫是否啟用黑名單
        'TIMS.Check_OrgBlackList(Me, "", objconn) 該計畫是否啟用黑名單功能
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_BlackOpen = True '產投計畫啟用黑名單

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim v_RIDValue As String = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim RelShip As String = TIMS.GET_RelshipforRID(v_RIDValue, objconn)

        'PlanMode 審核類型  SelectedIndex : '0:審核中 1:已通過
        If PlanMode.SelectedIndex = 0 Then
            DataGridTable1.Visible = True
            DataGridTable2.Visible = False
            bntAdd.Enabled = True
            Label1.Visible = True

            Dim sql_p As String = ""
            sql_p &= " Select P1.PlanID , P1.ComIDNO, P1.SeqNo,P1.RID, P1.PlanYear, P1.ClassName, P1.CyclType" & vbCrLf
            sql_p &= " , O1.OrgName, O2.Address, O2.ContactName, O2.Phone, P1.AppliedResult, P1.TransFlag" & vbCrLf
            sql_p &= " , p2.DistID ,'N' isBlack" & vbCrLf
            sql_p &= " FROM Plan_PlanInfo P1" & vbCrLf
            sql_p &= " JOIN Auth_Relship A1 ON P1.RID=A1.RID" & vbCrLf
            If (flgROLEIDx0xLIDx0) Then
                sql_p &= "  AND ISNULL(P1.AppliedResult,'') != 'Y'" & vbCrLf
            Else
                sql_p &= "  AND (P1.AppliedResult IS NULL OR P1.AppliedResult='O')" & vbCrLf
            End If
            'sql_p &= " AND P1.IsApprPaper='Y'" & vbCrLf
            sql_p &= String.Concat(" AND A1.relship LIKE '", String.Concat(RelShip, "%"), "'") & vbCrLf
            sql_p &= " JOIN Org_OrgInfo O1 ON A1.OrgID=O1.OrgID" & vbCrLf
            sql_p &= " JOIN Org_OrgPlanInfo O2 ON A1.RSID=O2.RSID" & vbCrLf
            sql_p &= " JOIN ID_Plan P2 ON P1.PlanID=P2.PlanID" & vbCrLf
            sql_p &= " WHERE P1.IsApprPaper='Y'" & vbCrLf
            If RIDValue.Value <> "" Then sql_p &= " AND P1.RID='" & RIDValue.Value & "'" & vbCrLf

            Select Case sm.UserInfo.RID
                Case "A" '署(局)
                    sql_p &= " AND P2.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                    sql_p &= " AND P2.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                Case Else '非署(局)
                    sql_p &= " AND P1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            End Select

            jobValue.Value = TIMS.ClearSQM(jobValue.Value)
            trainValue.Value = TIMS.ClearSQM(trainValue.Value)
            cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                'Me.LabTMID.Text="訓練業別"
                'If jobValue.Value <> "" Then jobValue.Value=Trim(Me.jobValue.Value)
                If jobValue.Value <> "" Then
                    sql_p &= " AND ( P1.TMID=" & jobValue.Value & vbCrLf
                    sql_p &= " OR P1.TMID IN (" & vbCrLf
                    sql_p &= "    SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '職類別
                    sql_p &= "    SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '業別
                    sql_p &= "    SELECT TMID FROM Key_TrainType WHERE busid='G')" & vbCrLf '產業人才投資方案類
                    sql_p &= " AND tmid =" & jobValue.Value & " )))" & vbCrLf

                End If
            Else
                'If trainValue.Value <> "" Then trainValue.Value=Trim(Me.trainValue.Value)
                If trainValue.Value <> "" Then sql_p &= " AND P1.TMID=" & trainValue.Value & vbCrLf
            End If

            '通俗職類
            'If txtCJOB_NAME.Text <> "" Then txtCJOB_NAME.Text=Trim(Me.txtCJOB_NAME.Text)
            If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then sql_p &= " AND P1.CJOB_UNKEY=" & cjobValue.Value & "" & vbCrLf

            'If trainValue.Value <> "" Then sql=sql & " AND P1.TMID='" & trainValue.Value & "' "
            If UNIT_SDATE.Text <> "" Then sql_p &= " AND P1.AppliedDate >= " & TIMS.To_date(Me.UNIT_SDATE.Text) & vbCrLf
            If UNIT_EDATE.Text <> "" Then sql_p &= " AND P1.AppliedDate <= " & TIMS.To_date(Me.UNIT_EDATE.Text) & vbCrLf

            ClassName.Text = TIMS.ClearSQM(ClassName.Text)
            'If ClassName.Text <> "" Then ClassName.Text=Trim(ClassName.Text)
            If ClassName.Text <> "" Then sql_p &= " AND P1.ClassName LIKE N'%" & ClassName.Text & "%'" & vbCrLf

            CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
            If CyclType.Text <> "" Then sql_p &= " AND P1.CyclType='" & CyclType.Text & "'" & vbCrLf

            sql_p &= " ORDER BY P1.STDate ,P1.ClassName "

            Dim dt1 As DataTable = DbAccess.GetDataTable(sql_p, objconn)
            'sqlAdapter=New SqlDataAdapter(sql, objconn)
            'sqlTable=New DataTable
            'sqlAdapter.Fill(sqlTable)

            If flag_BlackOpen AndAlso dt1.Rows.Count > 0 Then
                '檢測黑名單機構
                For Each odr As DataRow In dt1.Rows
                    If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS<>'38'").Length > 0 Then
                        odr("isBlack") = "Y"
                    Else
                        If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS='38' AND DistID='" & odr("DistID") & "'").Length > 0 Then odr("isBlack") = "Y"
                    End If
                Next
                dt1.AcceptChanges()
            End If

            dgPlan.DataSource = dt1
            dgPlan.DataBind()
        Else
            '1:已通過
            DataGridTable1.Visible = False
            DataGridTable2.Visible = True '顯示

            Dim sql_p1 As String = ""
            'sql_p1=""
            sql_p1 &= " SELECT P1.PlanID ,P1.ComIDNO ,P1.SeqNo ,P1.RID" & vbCrLf
            sql_p1 &= " ,P1.AppliedDate ,P1.STDate ,P1.FDDate" & vbCrLf
            sql_p1 &= " ,P1.ClassName ,P1.CyclType ,O1.OrgName" & vbCrLf
            sql_p1 &= " ,P1.TransFlag TFlag" & vbCrLf
            sql_p1 &= " ,CASE WHEN P1.TransFlag='Y' THEN '是' ELSE '否' END AS TransFlag" & vbCrLf
            sql_p1 &= " ,A1.relship" & vbCrLf
            sql_p1 &= " ,p2.DistID" & vbCrLf
            sql_p1 &= " ,'N' isBlack" & vbCrLf
            sql_p1 &= " FROM Plan_PlanInfo P1" & vbCrLf
            sql_p1 &= " JOIN Org_OrgInfo O1 ON P1.ComIDNO=O1.ComIDNO" & vbCrLf
            sql_p1 &= " JOIN Auth_Relship A1 ON A1.RID=P1.RID" & vbCrLf
            sql_p1 &= " JOIN ID_Plan P2 ON P1.PlanID=P2.PlanID" & vbCrLf
            sql_p1 &= "  AND A1.relship LIKE '" & RelShip & "%'" & vbCrLf
            'sql += " AND P1.RID IN (SELECT RID FROM Auth_Relship WHERE relship LIKE '" & sm.UserInfo.Relship & "%')" & vbCrLf
            sql_p1 &= " WHERE P1.IsApprPaper='Y'" & vbCrLf
            sql_p1 &= " AND P1.AppliedResult='Y'" & vbCrLf
            'sql_p1 &= " AND P1.IsApprPaper='Y'" & vbCrLf

            If RIDValue.Value <> "" Then sql_p1 &= " AND P1.RID='" & RIDValue.Value & "'" & vbCrLf
            'sql += " WHERE P2.PlanKind=2" & vbCrLf '計畫種類'1.自辦(內訓) 2.委外
            If sm.UserInfo.TPlanID = "02" Then
                sql_p1 &= " AND P2.PlanKind=1" & vbCrLf '計畫種類'1.自辦(內訓) 2.委外
            Else
                sql_p1 &= " AND P2.PlanKind=2" & vbCrLf '計畫種類'1.自辦(內訓) 2.委外
            End If
            Select Case sm.UserInfo.RID
                Case "A"
                    sql_p1 &= " AND P2.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                    sql_p1 &= " AND P2.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                Case Else '非署(局)
                    sql_p1 &= " AND P1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            End Select

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                'Me.LabTMID.Text="訓練業別"
                If jobValue.Value <> "" Then jobValue.Value = Trim(Me.jobValue.Value)
                If jobValue.Value <> "" Then
                    sql_p1 &= " AND (P1.TMID=" & jobValue.Value & vbCrLf
                    sql_p1 &= " OR P1.TMID IN (" & vbCrLf
                    sql_p1 &= "    SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '職類別
                    sql_p1 &= "    SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '業別
                    sql_p1 &= "    SELECT TMID FROM Key_TrainType WHERE busid='G')" & vbCrLf '產業人才投資方案類
                    sql_p1 &= " AND tmid=" & jobValue.Value & " )))" & vbCrLf
                End If
            Else
                If trainValue.Value <> "" Then trainValue.Value = Trim(Me.trainValue.Value)
                If trainValue.Value <> "" Then sql_p1 &= " AND P1.TMID=" & trainValue.Value & vbCrLf
            End If

            '通俗職類
            If txtCJOB_NAME.Text <> "" Then txtCJOB_NAME.Text = Trim(Me.txtCJOB_NAME.Text)
            If txtCJOB_NAME.Text <> "" Then sql_p1 &= " AND P1.CJOB_UNKEY=" & cjobValue.Value & "" & vbCrLf

            'If trainValue.Value <> "" Then sql=sql & " AND P1.TMID='" & trainValue.Value & "' "
            If UNIT_SDATE.Text <> "" Then sql_p1 &= " AND P1.AppliedDate >= " & TIMS.To_date(Me.UNIT_SDATE.Text) & vbCrLf
            If UNIT_EDATE.Text <> "" Then sql_p1 &= " AND P1.AppliedDate <= " & TIMS.To_date(Me.UNIT_EDATE.Text) & vbCrLf
            ClassName.Text = TIMS.ClearSQM(ClassName.Text)
            If ClassName.Text <> "" Then sql_p1 &= " AND P1.ClassName LIKE N'%" & ClassName.Text & "%'" & vbCrLf

            CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
            If CyclType.Text <> "" Then sql_p1 &= " AND P1.CyclType='" & CyclType.Text & "'" & vbCrLf

            Dim dt As DataTable = DbAccess.GetDataTable(sql_p1, objconn)
            If flag_BlackOpen AndAlso dt.Rows.Count > 0 Then
                '檢測黑名單機構
                For Each odr As DataRow In dt.Rows
                    If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS <> '38' ").Length > 0 Then
                        odr("isBlack") = "Y"
                    Else
                        If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS='38' AND DistID='" & odr("DistID") & "' ").Length > 0 Then odr("isBlack") = "Y"
                    End If
                Next
                dt.AcceptChanges()
            End If

            '查無資料
            DataGrid2.Visible = False
            Button1.Enabled = False '儲存鈕
            PageControler2.Visible = False
            If TIMS.dtNODATA(dt) Then Return

            '查無資料
            DataGrid2.Visible = True
            If sm.UserInfo.TPlanID = "02" Then
                Button1.Enabled = False '儲存鈕 '自辦不用儲存
                TIMS.Tooltip(Button1, "登入為自辦計畫，停用儲存鈕", True)
            Else
                Button1.Enabled = True '儲存鈕 '委外提供儲存
            End If
            PageControler2.Visible = True
            'PageControler2.SqlDataCreate(sql, "STDate,ClassName")
            PageControler2.PageDataTable = dt
            PageControler2.Sort = "STDate,ClassName"
            PageControler2.ControlerLoad()
        End If
    End Sub

    '查詢 
    Private Sub BtnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call Search1()
    End Sub

    '審核通過/不通過
    Private Sub BntAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        For Each eItem As DataGridItem In dgPlan.Items
            'Dim myitem As DataGridItem=dgPlan.Items(i)
            Dim Hid_PLANID As HiddenField = eItem.FindControl("Hid_PLANID")
            Dim Hid_COMIDNO As HiddenField = eItem.FindControl("Hid_COMIDNO")
            Dim Hid_SEQNO As HiddenField = eItem.FindControl("Hid_SEQNO")
            Dim KeyValue As HtmlInputHidden = eItem.FindControl("KeyValue")
            Dim objReason As HtmlControls.HtmlTextArea = eItem.FindControl("Reason")
            Dim objApply As DropDownList = eItem.FindControl("AppliedResult")
            Dim objApply1 As DropDownList = eItem.FindControl("AppliedResult1")

            If objApply.SelectedIndex <> 0 OrElse objApply1.SelectedIndex <> 0 Then
                Dim vsPlanID As String = TIMS.ClearSQM(Hid_PLANID.Value)
                Dim vsComIDNO As String = TIMS.ClearSQM(Hid_COMIDNO.Value)
                Dim vsSeqNo As String = TIMS.ClearSQM(Hid_SEQNO.Value)

                '照顧服務員自訓自用訓練計畫 '因違反第20點規定：自處分日期起1年內，該單位不得申請照顧服務員自訓自用訓練計畫之訓練單位。處分日期:2017/05/23
                Dim Errmsg As String = ""
                If TIMS.Check_OrgBlackList2(Me, vsComIDNO, iBlackType, OBTERMS.cst_c07, objconn) Then
                    Errmsg += OBTERMS.cst_c07_altMsg1 & vbCrLf
                    'Return False
                End If
                If TIMS.Check_OrgBlackList2(Me, vsComIDNO, iBlackType, OBTERMS.cst_c20, objconn) Then
                    Errmsg += OBTERMS.cst_c20_altMsg1 & vbCrLf
                    'Return False
                End If
                If TIMS.Check_OrgBlackList2(Me, vsComIDNO, iBlackType, OBTERMS.cst_c21, objconn) Then
                    Errmsg += OBTERMS.cst_c21_altMsg1 & vbCrLf
                    'Return False
                End If
                If Errmsg <> "" Then
                    Common.MessageBox(Me, Errmsg)
                    Exit Sub
                End If
            End If
        Next

        Dim s_SAVEMSG70 As String = ""
        Dim iRow_update As Integer = 0
        For Each eItem As DataGridItem In dgPlan.Items
            'Dim myitem As DataGridItem=dgPlan.Items(i)
            Dim Hid_PLANID As HiddenField = eItem.FindControl("Hid_PLANID")
            Dim Hid_COMIDNO As HiddenField = eItem.FindControl("Hid_COMIDNO")
            Dim Hid_SEQNO As HiddenField = eItem.FindControl("Hid_SEQNO")
            Dim Hid_RIDV As HiddenField = eItem.FindControl("Hid_RIDV")
            Dim KeyValue As HtmlInputHidden = eItem.FindControl("KeyValue")
            Dim objReason As HtmlControls.HtmlTextArea = eItem.FindControl("Reason")
            Dim objApply As DropDownList = eItem.FindControl("AppliedResult")
            Dim objApply1 As DropDownList = eItem.FindControl("AppliedResult1")

            Dim v_objApply1 As String = TIMS.GetListValue(objApply1)
            Dim v_objApply As String = TIMS.GetListValue(objApply)
            Dim V_AppliedResult As String = If(Not objApply.Visible AndAlso v_objApply1 <> "", v_objApply1, If(Not objApply1.Visible AndAlso v_objApply <> "", v_objApply, ""))

            Dim vsPlanID As String = TIMS.ClearSQM(Hid_PLANID.Value)
            Dim vsComIDNO As String = TIMS.ClearSQM(Hid_COMIDNO.Value)
            Dim vsSeqNo As String = TIMS.ClearSQM(Hid_SEQNO.Value)
            Dim vsRID As String = TIMS.ClearSQM(Hid_RIDV.Value)

            s_SAVEMSG70 = ""
            If (objApply.SelectedIndex <> 0 OrElse objApply1.SelectedIndex <> 0) AndAlso V_AppliedResult <> "" Then
                If V_AppliedResult = "Y" AndAlso TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    'vMsg1=""
                    'vMsg1 &= String.Concat("##SAVE_CHANGEDATA_T70: vsPlanID: ", vsPlanID, ",vsComIDNO: ", vsComIDNO, ",vsSeqNo: ", vsSeqNo, ",vsRID: ", vsRID)
                    'vMsg1 &= String.Concat(",v_objApply1: ", v_objApply1, ",v_objApply: ", v_objApply, ",V_AppliedResult: ", V_AppliedResult)
                    'TIMS.LOG.Info(vMsg1)

                    Dim ePMS_70 As New Hashtable From {{"APPLIEDRESULT", V_AppliedResult}, {"MODIFYACCT", sm.UserInfo.UserID},
                        {"PlanID", vsPlanID}, {"ComIDNO", vsComIDNO}, {"SeqNo", vsSeqNo}, {"RID", vsRID}}
                    s_SAVEMSG70 = SAVE_CHANGEDATA_T70(ePMS_70)
                    If s_SAVEMSG70 <> "" Then
                        s_SAVEMSG70 = String.Concat(s_SAVEMSG70, ".", vsPlanID, ".", vsComIDNO, ".", vsSeqNo)
                        Exit For
                    End If
                Else
                    'SqlCmd="update Plan_PlanInfo set AppliedResult='" & AppliedResult & "',ModifyAcct='" & sm.UserInfo.UserID & "',ModifyDate='" & Common.FormatNow(DateTime.Today) & "' where PlanID='" & myitem.Cells(1).Text & "' and ComIDNO='" & myitem.Cells(5).Text & "' and SeqNo='" & myitem.Cells(2).Text & "'"
                    '& TIMS.to_date(DateTime.Now)
                    Dim uPMS_u2 As New Hashtable From {{"APPLIEDRESULT", V_AppliedResult}, {"MODIFYACCT", sm.UserInfo.UserID},
                        {"PlanID", vsPlanID}, {"ComIDNO", vsComIDNO}, {"SeqNo", vsSeqNo}, {"RID", vsRID}}
                    Dim sql_u2 As String = ""
                    sql_u2 &= " UPDATE PLAN_PLANINFO "
                    sql_u2 &= " SET APPLIEDRESULT=@APPLIEDRESULT ,MODIFYACCT =@MODIFYACCT,MODIFYDATE=GETDATE()"
                    sql_u2 &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo AND RID=@RID"
                    DbAccess.ExecuteNonQuery(sql_u2, objconn, uPMS_u2)
                End If
                Call TIMS.Plan_VerRecord_Update(vsPlanID, vsComIDNO, vsSeqNo, sm.UserInfo.UserID, "", "", objReason.Value, objconn)
                iRow_update += 1
            End If
        Next

        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If s_SAVEMSG70 <> "" Then
                Common.MessageBox(Me, s_SAVEMSG70)
                Return
            End If
            If iRow_update = 0 Then
                Common.MessageBox(Me, "目前沒有儲存任何資料!!請重新審核!")
                Return
            End If
            Common.RespWrite(Me, "<script>alert('班級審核轉入作業完成!');</script>")
            Common.RespWrite(Me, $"<script>location.href='../04/TC_04_001.aspx?ID={TIMS.ClearSQM(Request("ID"))}'</script>")
        Else
            If iRow_update = 0 Then
                Common.MessageBox(Me, "目前沒有任何資料!!請重新審核!")
                Return
            End If
            Common.RespWrite(Me, "<script>alert('班級審核作業完成!!');</script>")
            Common.RespWrite(Me, $"<script>location.href='../04/TC_04_001.aspx?ID={TIMS.ClearSQM(Request("ID"))}'</script>")
        End If
    End Sub

    Private Sub dgPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgPlan.ItemCommand
        KeepSearchStr()
        'Dim url1 As String=""
        Dim sCmdArg As String = e.CommandArgument
        'If e.CommandArgument="" Then Exit Sub
        If sCmdArg = "" Then Exit Sub

        Dim vsPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vsComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
        Dim vsSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")
        If vsPlanID = "" OrElse vsComIDNO = "" OrElse vsSeqNo = "" Then Exit Sub

        Dim rqMID As String = TIMS.Get_MRqID(Me)

        Dim url1 As String = "../03/TC_03_001.aspx?ID=" & rqMID & "&todo=1&" & e.CommandArgument
        '企訓專用
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then url1 = "../03/TC_03_006.aspx?ID=" & rqMID & "&todo=1&" & e.CommandArgument

        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub dgPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgPlan.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass="SD_TD1"
                Dim SelectAll As DropDownList = e.Item.FindControl("SelectAll")
                SelectAll.Attributes("onchange") = "ChangeAll(this.selectedIndex);"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labSEQNO As Label = e.Item.FindControl("labSEQNO")
                Dim Hid_PLANID As HiddenField = e.Item.FindControl("Hid_PLANID")
                Dim Hid_COMIDNO As HiddenField = e.Item.FindControl("Hid_COMIDNO")
                Dim Hid_SEQNO As HiddenField = e.Item.FindControl("Hid_SEQNO")
                Dim Hid_RIDV As HiddenField = e.Item.FindControl("Hid_RIDV")
                Dim KeyValue As HtmlInputHidden = e.Item.FindControl("KeyValue")
                Dim objReason As HtmlControls.HtmlTextArea = e.Item.FindControl("Reason")
                Dim AppliedResult As DropDownList = e.Item.FindControl("AppliedResult")
                Dim AppliedResult1 As DropDownList = e.Item.FindControl("AppliedResult1")

                'e.Item.Cells(0).Text=e.Item.ItemIndex + 1 + dgPlan.CurrentPageIndex * dgPlan.PageSize
                labSEQNO.Text = e.Item.ItemIndex + 1 + dgPlan.CurrentPageIndex * dgPlan.PageSize
                Hid_PLANID.Value = Convert.ToString(drv("PlanID"))
                Hid_COMIDNO.Value = Convert.ToString(drv("ComIDNO"))
                Hid_SEQNO.Value = Convert.ToString(drv("SeqNo"))
                Hid_RIDV.Value = Convert.ToString(drv("RID"))
                KeyValue.Value = String.Concat(drv("PlanID"), "x", drv("ComIDNO"), "x", drv("SeqNo"))
                'objControl=e.Item.FindControl("AppliedResult")

                AppliedResult.Visible = True
                AppliedResult1.Visible = False
                'If drv("AppliedResult").ToString="" Then
                '    AppliedResult.Visible=True
                '    AppliedResult1.Visible=False
                'Else
                If drv("AppliedResult").ToString = "O" Then
                    AppliedResult.Visible = False
                    AppliedResult1.Visible = True
                    If drv("TransFlag").ToString = "N" Then AppliedResult1.Items.Add(New ListItem("審核不通過", "N"))
                End If

                Dim s_ClassName As String = ""
                Dim LinkButton1 As LinkButton = e.Item.FindControl("LinkButton1")
                LinkButton1.Text = TIMS.GET_CLASSNAME(drv("ClassName"), drv("CyclType"))
                LinkButton1.CommandArgument = String.Concat("PlanID=", drv("PlanID"), "&ComIDNO=", drv("ComIDNO"), "&SeqNo=", drv("SeqNo"))

                If Convert.ToString(drv("isBlack")) = "Y" Then
                    '該機構，已列入處分名單!!
                    AppliedResult.Enabled = False
                    TIMS.Tooltip(AppliedResult, cst_isBlackMsg)

                    AppliedResult1.Enabled = False
                    TIMS.Tooltip(AppliedResult1, cst_isBlackMsg)

                    LinkButton1.Enabled = False
                    TIMS.Tooltip(LinkButton1, cst_isBlackMsg)
                End If

            Case ListItemType.Footer
                If dgPlan.Items.Count = 0 Then
                    dgPlan.ShowFooter = True
                    Dim mycell As New TableCell
                    mycell.ColumnSpan = e.Item.Cells.Count
                    mycell.Text = "目前沒有任何資料!"
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(mycell)
                    e.Item.HorizontalAlign = HorizontalAlign.Center
                Else
                    dgPlan.ShowFooter = False
                End If
        End Select
    End Sub

    Sub KeepSearchStr()
        Dim s_sess_search As String = ""
        TIMS.SetMyValue(s_sess_search, "prg", "TC_04_001")
        TIMS.SetMyValue(s_sess_search, "center", center.Text)
        TIMS.SetMyValue(s_sess_search, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(s_sess_search, "TB_career_id", TB_career_id.Text)

        TIMS.SetMyValue(s_sess_search, "trainValue", trainValue.Value)
        TIMS.SetMyValue(s_sess_search, "txtCJOB_NAME", txtCJOB_NAME.Text)
        TIMS.SetMyValue(s_sess_search, "cjobValue", cjobValue.Value)
        TIMS.SetMyValue(s_sess_search, "jobValue", jobValue.Value)
        TIMS.SetMyValue(s_sess_search, "UNIT_SDATE", UNIT_SDATE.Text)
        TIMS.SetMyValue(s_sess_search, "UNIT_EDATE", UNIT_EDATE.Text)
        TIMS.SetMyValue(s_sess_search, "PlanMode", PlanMode.SelectedIndex)
        If PlanMode.SelectedIndex <> 0 Then
            TIMS.SetMyValue(s_sess_search, "PageIndex", (DataGrid2.CurrentPageIndex + 1))
        End If

        Session("search") = s_sess_search
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        'Case ListItemType.Header
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedResult2 As DropDownList = e.Item.FindControl("AppliedResult2")
                Dim KeyValue As HtmlInputHidden = e.Item.FindControl("KeyValue")

                e.Item.Cells(cst_DG2_col序號).Text = e.Item.ItemIndex + 1 + DataGrid2.CurrentPageIndex * DataGrid2.PageSize

                'orgname
                If Split(drv("relship"), "/").Length >= 3 Then
                    Dim ParenrRID As String = Split(drv("relship"), "/")(Split(drv("relship"), "/").Length - 3)
                    Dim ParentName As String = ""
                    If Auth_Relship.Select("RID='" & ParenrRID & "'").Length <> 0 Then ParentName = Auth_Relship.Select("RID='" & ParenrRID & "'")(0)("OrgName")
                    If ParentName <> "" Then e.Item.Cells(cst_DG2_col訓練機構).Text = String.Format("<font color='Blue'>{0}</font>-{1}", ParentName, drv("OrgName"))
                End If
                e.Item.Cells(cst_DG2_col班別名稱).Text = TIMS.GET_CLASSNAME(drv("ClassName"), drv("CyclType"))

                KeyValue.Value = String.Concat(drv("PlanID"), "x", drv("ComIDNO"), "x", drv("SeqNo"))

                AppliedResult2.Enabled = False '轉班後不可取消審核
                If Convert.ToString(drv("TFlag")) <> "Y" Then AppliedResult2.Enabled = True '轉班前可取消審核

                '轉班後不可取消審核
                '已開班轉入，不可再做審核還原之功能
                If Not AppliedResult2.Enabled Then TIMS.Tooltip(AppliedResult2, "已開班轉入，不可再做取消審核之功能")

                Select Case sm.UserInfo.RID
                    Case "A" '署(局)
                        If Not AppliedResult2.Enabled Then
                            AppliedResult2.Enabled = True
                            'TIMS.Tooltip(AppliedResult2, "局登入，開放取消審核之功能@@")
                            TIMS.Tooltip(AppliedResult2, "署登入，開放取消審核之功能 ")
                        End If
                End Select

                If Convert.ToString(drv("isBlack")) = "Y" Then
                    AppliedResult2.Enabled = False
                    TIMS.Tooltip(AppliedResult2, TIMS.cst_gBlackMsg1)
                End If
            Case ListItemType.Footer
                If DataGrid2.Items.Count = 0 Then
                    DataGrid2.ShowFooter = True
                    Dim mycell As New TableCell
                    mycell.ColumnSpan = e.Item.Cells.Count
                    mycell.Text = "目前沒有任何資料!"
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(mycell)
                    e.Item.HorizontalAlign = HorizontalAlign.Center
                Else
                    DataGrid2.ShowFooter = False
                End If
        End Select
    End Sub

    ''' <summary>取消審核 按鈕</summary>
    ''' <param name="PCS"></param>
    Sub UPDATE_PLANINFO_AppliedResult(ByRef PCS As String)
        'Dim PCS As String=TIMS.GetMyValue2(htPP, "PCS")
        Dim uParms As New Hashtable
        uParms.Add("ModifyAcct", sm.UserInfo.UserID)
        uParms.Add("PCS", PCS)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim uSql As String = " UPDATE PLAN_PLANINFO SET AppliedResult=NULL,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() WHERE concat(PLANID,'x',COMIDNO,'x',SEQNO)=@PCS"
            DbAccess.ExecuteNonQuery(uSql, objconn, uParms)
        Else
            Dim uSql As String = " UPDATE PLAN_PLANINFO SET AppliedResult ='O',ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() WHERE concat(PLANID,'x',COMIDNO,'x',SEQNO)=@PCS"
            DbAccess.ExecuteNonQuery(uSql, objconn, uParms)
        End If

        Dim uParms2 As New Hashtable
        uParms2.Add("PCS", PCS)
        Dim uSql2 As String = " DELETE PLAN_VERRECORD WHERE concat(PLANID,'x',COMIDNO,'x',SEQNO)=@PCS"
        DbAccess.ExecuteNonQuery(uSql2, objconn, uParms2)
    End Sub

    '取消審核 按鈕。
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        For Each item As DataGridItem In DataGrid2.Items
            Dim KeyValue As HtmlInputHidden = item.FindControl("KeyValue")
            Dim AppliedResult2 As DropDownList = item.FindControl("AppliedResult2")
            '取消審核:AppliedResult2.SelectedIndex <> 0 
            If AppliedResult2.SelectedIndex <> 0 Then
                UPDATE_PLANINFO_AppliedResult(KeyValue.Value)
            End If
        Next
        Common.MessageBox(Me, "儲存成功")
        'btnQuery_Click(sender, e)
        Call Search1()
    End Sub

    ''' <summary>轉入資料(SAVE) PLAN_PLANINFO CLASS_CLASSINFO</summary>
    Function SAVE_CHANGEDATA_T70(rPMS As Hashtable) As String
        Dim RETURN_MSG As String = ""
        Dim v_PlanID As String = TIMS.GetMyValue2(rPMS, "PlanID")
        Dim v_ComIDNO As String = TIMS.GetMyValue2(rPMS, "ComIDNO")
        Dim v_SeqNo As String = TIMS.GetMyValue2(rPMS, "SeqNo")
        Dim rqRID As String = TIMS.GetMyValue2(rPMS, "RID")

        Dim iCnt As Integer = 0
        Dim drPlaninfo As DataRow = TIMS.Get_PlanInfoDataRow(objconn, v_PlanID, v_ComIDNO, v_SeqNo, rqRID, "", iCnt)
        'Common.MessageBox(Me, "計畫資料有誤，請重新選擇!!")
        If drPlaninfo Is Nothing Then RETURN_MSG = "計畫資料有誤，請重新選擇!!"
        If RETURN_MSG <> "" Then Return RETURN_MSG

        '登入者檢查
        Hid_UserComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_UserComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    'Errmsg &= "於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                    'Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業!")
                    RETURN_MSG = "於處分日期起的期間，已審核通過的班級不可進行轉班作業!" 'Exit Sub '有錯誤訊息 'Return False '不可儲存
                    If RETURN_MSG <> "" Then Return RETURN_MSG
            End Select
        End If

        '轉入班級者檢查
        If TIMS.Check_OrgBlackList2(Me, Convert.ToString(drPlaninfo("COMIDNO")), iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    'Errmsg &= "於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                    'Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業!!")
                    RETURN_MSG = "於處分日期起的期間，已審核通過的班級不可進行轉班作業!!" 'Exit Sub '有錯誤訊息 'Return False '不可儲存
                    If RETURN_MSG <> "" Then Return RETURN_MSG
            End Select
        End If

        Dim pms_r1 As New Hashtable() From {{"RID", drPlaninfo("RID")}}
        Dim sql_r1 As String = " SELECT RELSHIP FROM AUTH_RELSHIP WHERE RID=@RID"
        Dim relship As String = Convert.ToString(DbAccess.ExecuteScalar(sql_r1, objconn, pms_r1))

        'Dim pms_cd As New Hashtable() From {{"CLSID", vCLSID}}
        'Dim sql_cd As String=" SELECT CLASSENAME FROM ID_CLASS WHERE CLSID=@CLSID"
        'Dim ClassEngName As String=Convert.ToString(DbAccess.ExecuteScalar(sql_cd, objconn, pms_cd))

        Dim pms_tc As New Hashtable() From {{"PlanID", drPlaninfo("PlanID")}, {"ComIDNO", drPlaninfo("ComIDNO")}, {"SeqNo", drPlaninfo("SeqNo")}}
        Dim sql_tc As String = " SELECT PCONT,PNAME FROM PLAN_TRAINDESC WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO ORDER BY PTDID"
        Dim dtTRAINDESC As DataTable = DbAccess.GetDataTable(sql_tc, objconn, pms_tc)

        Dim class_PName As String = ""
        For Each drTra As DataRow In dtTRAINDESC.Rows
            Dim sPName As String = TIMS.ClearSQM(drTra("PName"))
            class_PName &= String.Concat(If(class_PName <> "", ",", ""), sPName)
        Next

        Dim htPV As New Hashtable 'htPV.Clear()
        htPV.Add("RID", Convert.ToString(drPlaninfo("RID")))
        htPV.Add("TMID", Convert.ToString(drPlaninfo("TMID")))
        If sm.UserInfo.LID = 0 Then
            Dim drPN As DataRow = TIMS.GetPlanID1(v_PlanID, objconn)
            If drPN Is Nothing Then RETURN_MSG = "傳入計畫代碼有誤，請檢查系統參數!"
            If RETURN_MSG <> "" Then Return RETURN_MSG
            htPV.Add("TPLANID", drPN("TPLANID"))
            htPV.Add("DISTID", drPN("DISTID"))
            htPV.Add("YEARS", drPN("YEARS"))
        Else
            htPV.Add("TPLANID", sm.UserInfo.TPlanID)
            htPV.Add("DISTID", sm.UserInfo.DistID)
            htPV.Add("YEARS", sm.UserInfo.Years)
        End If
        htPV.Add("CJOB_UNKEY", Convert.ToString(drPlaninfo("CJOB_UNKEY")))
        htPV.Add("CLASSNAME", Convert.ToString(drPlaninfo("CLASSNAME")))
        htPV.Add("ClassEngName", Convert.ToString(drPlaninfo("ClassEngName")))
        htPV.Add("CONTENT", class_PName)

        Dim dr70c As DataRow = TIMS.Get_ClassID_70(sm, objconn, htPV)
        If dr70c Is Nothing Then
            'Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            RETURN_MSG = "設定班別代碼有誤，請檢查輸入參數!" 'Exit Sub
            If RETURN_MSG <> "" Then Return RETURN_MSG
        End If

        Dim vCLASSENAME As String = Convert.ToString(dr70c("CLASSENAME"))
        Dim vCLSID As String = Convert.ToString(dr70c("CLSID"))
        If IsDBNull(dr70c("CJOB_UNKEY")) Then '如果有CJOB_UNKEY是NULL的
            'Dim strScript2 As String=""
            'strScript2 += "<script language=""javascript"">" + vbCrLf
            'strScript2 += "alert(' 轉入失敗,請聯絡承辦人設定此班別代碼的通俗職類資料,才可進行開班轉入動作!!');" + vbCrLf
            'strScript2 += "</script>"
            'Page.RegisterStartupScript("", strScript2)
            RETURN_MSG = "轉入失敗,請聯絡承辦人設定此班別代碼的通俗職類資料,才可進行開班轉入動作!!"
            If RETURN_MSG <> "" Then Return RETURN_MSG
        Else '如果有CJOB_UNKEY的值
            Dim uPMS9 As New Hashtable From {{"CJOB_UNKEY", dr70c("CJOB_UNKEY")}, {"PLANID", v_PlanID}, {"COMIDNO", v_ComIDNO}, {"SEQNO", v_SeqNo}}
            Dim sql9u As String = ""
            sql9u &= " UPDATE PLAN_PLANINFO "
            sql9u &= " SET CJOB_UNKEY =@CJOB_UNKEY"
            sql9u &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
            DbAccess.ExecuteNonQuery(sql9u, objconn, uPMS9)
        End If

        Dim PMSck As New Hashtable() From {{"CLSID", vCLSID}, {"CLASSCNAME", drPlaninfo("CLASSNAME")}, {"PlanID", drPlaninfo("PlanID")}, {"RID", drPlaninfo("RID")}}
        Dim check_sql As String = ""
        check_sql &= " SELECT concat('(',dbo.FN_CLASSID2(cc.CLSID),')',dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE)) CLASSCNAME" & vbCrLf
        check_sql &= " FROM dbo.CLASS_CLASSINFO cc" & vbCrLf
        check_sql &= " WHERE cc.CLSID=@CLSID AND cc.CLASSCNAME=@CLASSCNAME" & vbCrLf '依班別代碼/班名(重複)
        check_sql &= " AND cc.PlanID=@PlanID AND cc.RID=@RID" & vbCrLf 'PlanID,機構
        If $"{drPlaninfo("CyclType")}" <> "" Then
            check_sql &= " AND cc.CyclType=@CyclType" & vbCrLf '期別(重複)
            PMSck.Add("CyclType", drPlaninfo("CyclType"))
        Else
            check_sql &= " AND cc.CyclType IS NULL" & vbCrLf '期別(重複)
        End If
        Dim dr9ck As DataRow = DbAccess.GetOneRow(check_sql, objconn, PMSck)

        Dim blnChkIsDouble As Boolean = If(dr9ck IsNot Nothing, True, False) '重複

        If blnChkIsDouble Then '重複
            'Dim strScript2 As String=String.Concat("轉入班級資料 班別代碼與期別重複!!", vbCrLf, dr9ck("classcname"))
            'Common.MessageBox(Me, strScript2)
            RETURN_MSG = String.Concat("轉入班級資料 班別代碼/班名與期別重複!", vbCrLf, dr9ck("CLASSCNAME"))
            If RETURN_MSG <> "" Then Return RETURN_MSG
        End If

        Dim sPMS_cc As New Hashtable From {{"PLANID", v_PlanID}, {"COMIDNO", v_ComIDNO}, {"SEQNO", v_SeqNo}}
        Dim sql_cc As String = ""
        sql_cc &= " SELECT 'X' FROM dbo.CLASS_CLASSINFO"
        sql_cc &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim dtCC As DataTable = DbAccess.GetDataTable(sql_cc, objconn, sPMS_cc)
        'Common.MessageBox(Me, "新增開班資料重複(已有轉班資料!!!)")
        If dtCC.Rows.Count > 0 Then RETURN_MSG = "新增開班資料重複(已有轉班資料!!!)" 'Exit Sub '有錯誤訊息 'Return False '不可儲存 
        If RETURN_MSG <> "" Then Return RETURN_MSG

        'TIMS專用
        Hid_RID1.Value = Convert.ToString(drPlaninfo("RID")).Substring(0, 1)

        Dim pp_years As String = CInt(drPlaninfo("PlanYear"))
        Try
            Dim sqldr As DataRow = Nothing
            'Dim objTrans As SqlTransaction
            Dim sqlAdapter As SqlDataAdapter = Nothing
            Dim sql_c As String = " SELECT * FROM CLASS_CLASSINFO WHERE 1<>1 "
            Dim sqlTable As DataTable = DbAccess.GetDataTable(sql_c, sqlAdapter, objconn)

            sqldr = sqlTable.NewRow 'CLASS_CLASSINFO
            sqlTable.Rows.Add(sqldr)

            '新增一組OCID
            Dim iOCID_New As Integer = DbAccess.GetNewId(objconn, "CLASS_CLASSINFO_OCID_SEQ,CLASS_CLASSINFO,OCID")
            sqldr("OCID") = iOCID_New
            sqldr("ONSHELLDATE") = TIMS.Cdate2(CDate(Now).ToString("yyyy/MM/dd"))
            sqldr("LastState") = "A" 'A: 新增(最後異動狀態)

            sqldr("Relship") = relship
            'sqldr("Content")=class_cont
            sqldr("Content") = class_PName   '2007/11/19 修改成將訓練內容簡介的課程單元帶入--Charles
            sqldr("Years") = pp_years.Substring(2) '012
            sqldr("PlanID") = drPlaninfo("PlanID")
            sqldr("ComIDNO") = drPlaninfo("ComIDNO")
            sqldr("SeqNO") = drPlaninfo("SeqNO")
            sqldr("RID") = drPlaninfo("RID")
            sqldr("TMID") = drPlaninfo("TMID")
            sqldr("TPropertyID") = 1 '1:在職／2:接受企業委託
            'CJOB_UNKEY
            sqldr("CJOB_UNKEY") = If(Convert.ToString(drPlaninfo("CJOB_UNKEY")) <> "", drPlaninfo("CJOB_UNKEY"), dr70c("CJOB_UNKEY")) '通俗職類
            sqldr("CLSID") = vCLSID
            sqldr("ClassCName") = drPlaninfo("ClassName")
            sqldr("ClassEngName") = If(Convert.ToString(drPlaninfo("ClassEngName")) <> "", Convert.ToString(drPlaninfo("ClassEngName")), vCLASSENAME)

            Dim vCyclType As String = Convert.ToString(drPlaninfo("CyclType"))
            If vCyclType = "" Then vCyclType = TIMS.cst_Default_CyclType
            vCyclType = TIMS.FmtCyclType(vCyclType)
            sqldr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)

            sqldr("ClassNum") = 1 'vCyclType '班數
            sqldr("ADVANCE") = drPlaninfo("ADVANCE") '訓練課程類型 ADVANCE
            sqldr("TNum") = drPlaninfo("TNum")
            sqldr("THours") = drPlaninfo("THours")
            sqldr("STDate") = drPlaninfo("STDate")
            sqldr("FTDate") = drPlaninfo("FDDate")
            'SELECT SENTERDATE,FENTERDATE,EXAMDATE,ExamPeriod FROM PLAN_PLANINFO WHERE ROWNUM  <=10
            'SELECT SENTERDATE,FENTERDATE,EXAMDATE,FENTERDATE2,ExamPeriod FROM CLASS_CLASSINFO  WHERE ROWNUM  <=10
            sqldr("SENTERDATE") = drPlaninfo("SENTERDATE")
            sqldr("FENTERDATE") = drPlaninfo("FENTERDATE")
            sqldr("EXAMDATE") = drPlaninfo("EXAMDATE")
            sqldr("ExamPeriod") = drPlaninfo("ExamPeriod")
            Dim sFENTERDATE As String = Convert.ToString(drPlaninfo("FENTERDATE"))
            Dim sEXAMDATE As String = Convert.ToString(drPlaninfo("EXAMDATE"))
            Dim SS1 As String = ""
            TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value)
            TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
            Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
            If sFENTERDATE2 <> "" Then sqldr("FENTERDATE2") = CDate(sFENTERDATE2) 'TIMS.GET_FENTERDATE2()
            sqldr("CheckInDate") = drPlaninfo("CheckInDate")
            '2005/8/11新增轉入訓練地點--Melody
            sqldr("TaddressZip") = drPlaninfo("TaddressZip")
            sqldr("TaddressZIP6W") = drPlaninfo("TaddressZIP6W")
            sqldr("TAddress") = drPlaninfo("TAddress")
            '2005/8/12新增轉入課程目標--Melody，2007/9/26 修改成將訓練目標帶入即可--Charles
            'sqldr("Purpose")="一、學科：" & drPlaninfo("PurScience") & "二、術科：" & drPlaninfo("PurTech")
            sqldr("Purpose") = drPlaninfo("PurScience")
            sqldr("NotOpen") = "N" '不開班
            'sqldr("NORID")=Convert.DBNull '不開班原因代碼
            'sqldr("OtherReason")=Convert.DBNull '不開班其他原因說明
            'sqldr("LastState")="M" 'M: 修改(最後異動狀態)
            'sqldr("Companyname")=If(Companyname.Text <> "", Companyname.Text, Convert.DBNull) '企業名稱
            sqldr("IsCalculate") = "N" '是否試算
            sqldr("IsClosed") = "N" '是否結訓
            sqldr("IsSuccess") = "Y" '是否轉入成功
            sqldr("BGTime") = 0 '勾稽次數
            sqldr("IsApplic") = "N" '納入志願
            '班級英文名稱
            sqldr("CLASSENGNAME") = drPlaninfo("CLASSENGNAME")
            '訓練時段'取得鍵值-訓練時段
            sqldr("TPERIOD") = drPlaninfo("TPERIOD")
            sqldr("NOTE3") = drPlaninfo("NOTE3")
            '「訓練期限」
            sqldr("TDEADLINE") = drPlaninfo("TDEADLINE")
            '導師名稱
            sqldr("CTName") = drPlaninfo("CTName")
            'EADDRESSZIP,EADDRESS,EADDRESSZIP6W
            sqldr("EADDRESSZIP") = drPlaninfo("EADDRESSZIP")
            sqldr("EADDRESSZIP6W") = drPlaninfo("EADDRESSZIP6W")
            sqldr("EADDRESS") = drPlaninfo("EADDRESS")
            sqldr("Class_Unit") = Convert.DBNull ' drPlaninfo("Class_Unit")
            sqldr("ModifyAcct") = sm.UserInfo.UserID
            sqldr("ModifyDate") = Now()
            'INSERT/UPDATE CLASS_CLASSINFO
            DbAccess.UpdateDataTable(sqlTable, sqlAdapter)

            'UPDATE PLAN_PLANINFO
            Dim uPMS_up1 As New Hashtable From {{"APPLIEDRESULT", "Y"}, {"MODIFYACCT", sm.UserInfo.UserID}, {"PlanID", v_PlanID}, {"ComIDNO", v_ComIDNO}, {"SeqNo", v_SeqNo}}
            Dim sql_up1 As String = ""
            sql_up1 &= " UPDATE PLAN_PLANINFO "
            sql_up1 &= " SET APPLIEDRESULT=@APPLIEDRESULT ,TransFlag='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()"
            sql_up1 &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
            DbAccess.ExecuteNonQuery(sql_up1, objconn, uPMS_up1)
            'Common.MessageBox(Me, "班級已轉入!")
        Catch ex As Exception
            'objTrans.Rollback()
            'Common.MessageBox(Page, String.Concat("班級轉入失敗!!", ex.Message))
            TIMS.LOG.Error(ex.Message, ex)
            RETURN_MSG = String.Concat("班級轉入失敗!!", ex.Message)
            If RETURN_MSG <> "" Then Return RETURN_MSG
            'Throw ex 'Finally ' objconn.Close()
        End Try

        Return RETURN_MSG
    End Function

End Class
