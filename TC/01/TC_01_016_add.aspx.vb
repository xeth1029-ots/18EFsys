Partial Class TC_01_016_add
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    'Plan_PlanInfo
    'Plan_CostItem
    'Plan_Revise
    'Plan_Teacher
    'Plan_TrainDesc
    'Plan_TrainDesc_Revise
    'Plan_VerReport
    'Plan_VerRecord

    'Class_ClassInfo
    'Stud_SelResult 
    'Stud_EnterType2 
    'Stud_EnterType
    'Stud_DataLid
    'Org_StudRecord
    'Class_Visitor
    'Class_UnexpectVisitor
    'Class_UnexpectTel
    'Class_RestTime
    'Sys_DelLog

    'PLAN_DEPOT
    'Plan_BusPackage '計畫包班事業單位(產學訓)
    'Plan_PersonCost '一人份材料明細(產學訓)
    'Plan_CommonCost '共同材料明細(產學訓)
    'Plan_SheetCost    '教材費用 (產學訓)
    'Plan_OtherCost '其他費用 (產學訓)
    Const cst_table_rows As String = "PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost"

    Dim objconn As SqlConnection
    'Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        'PageControler1.PageDataGrid = DataGrid2
        '2005/1/6新增輸入郵遞區號回傳地區名稱的Javascript----------------------------------Start
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        TBCity.Attributes("onblur") = "getzipname(this.value,'TBCity','city_code');"
        '2005/1/6新增輸入郵遞區號回傳地區名稱的Javascript----------------------------------End

        Check_TPlanID28()
        '判斷登入角色若為
        '1.產業人才投資計畫角色登入者則為唯讀
        '2.提升勞工自主學習計畫角色登入者則為可以使用

        Me.ViewState("RWPlanRID") = ""
        Me.ViewState("RSID") = ""
        Me.ViewState("ProcessType") = ""
        Me.ViewState("Re_orgid") = ""
        Me.ViewState("Re_planid") = ""
        Me.ViewState("Re_rid") = ""
        Me.ViewState("Re_distid") = ""
        Me.ViewState("Re_ComIDNO") = ""
        Me.ViewState("orglevel") = "" '機構階層
        'Me.ViewState("ReOrgName") = "" '機構名稱
        Me.ViewState("orgName") = "" '轄區_機構名稱
        Me.ViewState("plan_name") = "" '年度+轄區+計畫+計畫代碼檔_序號
        Me.ViewState("AppliedResult") = ""

        If Not Request("orgid") Is Nothing Then
            Me.ViewState("Re_orgid") = Request("orgid")
            OrgIDValue.Value = Me.ViewState("Re_orgid")
        End If
        If Not Request("RWPlanRID") Is Nothing Then Me.ViewState("RWPlanRID") = Request("RWPlanRID") 'PlanID,RID (是登入計畫&權限，非使用者建立之計畫&權限)
        If Not Request("RSID") Is Nothing Then Me.ViewState("RSID") = Request("RSID")
        If Not Request("ProcessType") Is Nothing Then Me.ViewState("ProcessType") = Request("ProcessType")
        If Not Request("planid") Is Nothing Then Me.ViewState("Re_planid") = Request("planid") 'RWPlanID (是登入計畫，非使用者建立之權限計畫)
        If Not Request("rid") Is Nothing Then Me.ViewState("Re_rid") = Request("rid") 'RID (是登入權限，非使用者建立之計畫權限)
        If Not Request("distid") Is Nothing Then Me.ViewState("Re_distid") = Request("distid")
        If Not Request("comidno") Is Nothing Then Me.ViewState("Re_ComIDNO") = Request("comidno")
        If Not Request("AppliedResult") Is Nothing Then Me.ViewState("AppliedResult") = Request("AppliedResult")

        MenuTable.Style.Item("display") = "none"

        If Not Session("_Search") Is Nothing Then
            Me.ViewState("OrgSearchStr") = Session("_Search")
            Session("_Search") = Nothing
        End If

        If Not Page.IsPostBack Then
            Page.RegisterStartupScript("Load", "<script>SetLastYearExeRate();</script>")
            Select Case Me.ViewState("ProcessType")
                Case "modify"
                    Me.lblProecessType.Text = "計畫調動"
                Case Else
                    Me.lblProecessType.Text = "(開發檢視中) 情況不明-請將操作步驟告知系統管理者"
            End Select
            'DistrictList = TIMS.Get_DistID(DistrictList)
            DistrictList = TIMS.Get_DistID(DistrictList, TIMS.dtNothing(), objconn)
            DistrictList.Items.Remove(DistrictList.Items.FindByValue(""))
            'If Me.ViewState("Re_distid") = "0" Then Me.level_list.Items.Insert(0, New ListItem("局", "0"))  '轄區是署(局)
            If sm.UserInfo.RoleID = "0" Then '表示角色為超級管理者 【階層為署(局)】
                If sm.UserInfo.DistID <> "000" Then
                    Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                    Me.DistrictList.Enabled = False
                    TIMS.Tooltip(DistrictList, "管理者權限")
                End If
            ElseIf sm.UserInfo.RoleID = "1" Then '角色為系統管理者【階層為分署(中心)】
                Me.level_list.Items.Remove(DistrictList.Items.FindByValue("1"))
                'Me.level_list.Items.Insert(1, New ListItem("中心", "1"))
                Me.level_list.Items.Insert(1, New ListItem("分署", "1"))
                Common.SetListItem(level_list, sm.UserInfo.LID)
                'Me.DistrictList.SelectedValue = sm.UserInfo.DistID
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                Me.DistrictList.Enabled = False
                TIMS.Tooltip(DistrictList, "分署的權限")

            ElseIf sm.UserInfo.RoleID > "1" Then '角色為非系統管理者、非超級管理者【階層為分署(中心)以下】
                Me.level_list.Enabled = False
                TIMS.Tooltip(level_list, "委訓單位的權限")
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                Me.DistrictList.Enabled = False
                TIMS.Tooltip(DistrictList, "委訓單位的權限")

            End If
            OrgKindList = TIMS.Get_OrgType(OrgKindList, objconn)
            'Dim City As String
            'Dim planid As String
            Select Case Me.ViewState("ProcessType")
                Case "modify"
                    TBplan.Enabled = False
                    TIMS.Tooltip(TBplan, "不能修改~")
                    Dim sub_sql As String = ""
                    If Me.ViewState("Re_rid").ToString.Trim.Length <> 1 Then
                        '委訓卡計畫
                        sub_sql = "" & vbCrLf
                        sub_sql &= " SELECT 'x' x " & vbCrLf
                        sub_sql &= " FROM dbo.Org_orginfo a " & vbCrLf
                        sub_sql &= " JOIN dbo.Auth_Relship b ON a.orgid = b.orgid " & vbCrLf
                        sub_sql &= " LEFT JOIN dbo.Class_classinfo c ON b.PlanID = c.PlanID " & vbCrLf
                        sub_sql &= " WHERE a.orgid = '" & Me.ViewState("Re_orgid") & "' " & vbCrLf
                        sub_sql &= " AND b.PlanID = '" & Me.ViewState("Re_planid") & "' " & vbCrLf
                        sub_sql &= " AND c.RID = '" & Me.ViewState("Re_rid") & "' " & vbCrLf
                    Else
                        '分署(中心)不卡計畫
                        sub_sql = "" & vbCrLf
                        sub_sql &= " SELECT 'x' x " & vbCrLf
                        sub_sql &= " FROM dbo.Org_orginfo a " & vbCrLf
                        sub_sql &= " JOIN dbo.Auth_Relship b ON a.orgid = b.orgid " & vbCrLf
                        sub_sql &= " LEFT JOIN dbo.Class_classinfo c ON b.PlanID = c.PlanID " & vbCrLf
                        sub_sql &= " WHERE a.orgid = '" & Me.ViewState("Re_orgid") & "' "
                        sub_sql &= " AND c.RID = '" & Me.ViewState("Re_rid") & "' " & vbCrLf
                    End If

                    If DbAccess.GetCount(sub_sql, objconn) > 0 Then
                        Me.level_list.Enabled = False
                        Me.TBplan.Enabled = False
                        Me.DistrictList.Enabled = False
                        TIMS.Tooltip(level_list, "有開課,不能修改~")
                        TIMS.Tooltip(TBplan, "有開課,不能修改~")
                        TIMS.Tooltip(DistrictList, "有開課,不能修改~")
                        TBID.Enabled = False
                        TIMS.Tooltip(TBID, "有開課,不能修改~")
                    End If
                    '機構階層
                    If Me.ViewState("Re_rid").ToString.Trim.Length <> 1 Then
                        '委訓卡計畫
                        sub_sql = "" & vbCrLf
                        sub_sql &= " SELECT b.orglevel,a.comidno" & vbCrLf
                        sub_sql &= " FROM dbo.Org_orginfo a" & vbCrLf
                        sub_sql &= " JOIN dbo.Auth_Relship b ON a.orgid = b.orgid " & vbCrLf
                        sub_sql &= " WHERE a.orgid = '" & Me.ViewState("Re_orgid") & "'" & vbCrLf
                        sub_sql &= " AND b.PlanID = '" & Me.ViewState("Re_planid") & "'" & vbCrLf
                        sub_sql &= " AND b.RID = '" & Me.ViewState("Re_rid") & "' " & vbCrLf
                    Else
                        '分署(中心)不卡計畫
                        sub_sql = "" & vbCrLf
                        sub_sql &= " SELECT b.orglevel,a.comidno" & vbCrLf
                        sub_sql &= " FROM dbo.Org_orginfo a" & vbCrLf
                        sub_sql &= " JOIN dbo.Auth_Relship b ON a.orgid = b.orgid " & vbCrLf
                        sub_sql &= " WHERE a.orgid = '" & Me.ViewState("Re_orgid") & "'" & vbCrLf
                        sub_sql &= " AND b.RID = '" & Me.ViewState("Re_rid") & "' " & vbCrLf
                    End If
                    Dim dr_o As DataRow = DbAccess.GetOneRow(sub_sql, objconn)

                    Me.ViewState("orglevel") = Convert.ToString(dr_o("orglevel"))
                    TBID.Text = Convert.ToString(dr_o("comidno"))

                    If Me.ViewState("orglevel") > "1" Then Common.SetListItem(Me.level_list, "2") '設定為委訓
                    '計算共用數量
                    sub_sql = " SELECT * FROM Auth_Relship WHERE orgid = " & Me.ViewState("Re_orgid")
                    If DbAccess.GetCount(sub_sql, objconn) > 1 Then '有共用過,訓練機構共同資料不能修改
                        DistrictList.Enabled = False
                        TBID.Enabled = False
                        TIMS.Tooltip(DistrictList, "有共用過,訓練機構共同資料不能修改")
                        TIMS.Tooltip(TBID, "有共用過,訓練機構共同資料不能修改")
                    End If

                    Dim flag_x As Boolean = create1()
                    'If Not create1() Then End If 'Exit Sub '未共用

                    ''Show_OrgInfoData
                    'If Not Show_OrgInfoData(Me.ViewState("Re_orgid"), Me.ViewState("Re_planid"), Me.ViewState("Re_rid"), msg.Text) Then
                    '    Exit Sub
                    'End If
                    ''Get_OrgPlanNameList1
                    'If Get_OrgPlanNameList1(Me.TBID.Text, Me.ViewState("Re_planid")) = False Then
                    '    Exit Sub
                    'End If
            End Select
        End If

        LastYearExeRate.Enabled = False

        '郵遞區號查詢
        Litcity_code.Text = TIMS.Get_WorkZIPB3Link2()

        MenuTable.Style.Item("display") = ""
        HistoryTable.Style.Item("display") = "none"

        ''返回呼叫頁目方式
        'If Session("Redirect") Is Nothing Then
        '    Me.Button1.Visible = False
        'Else
        '    Me.ViewState("Redirect") = Session("Redirect")
        '    Session("Redirect") = Nothing
        '    Me.Button1.Visible = True
        'End If

        '設定為失效頁面
        Table1.Disabled = True
        DistrictList.Enabled = False
        level_list.Enabled = False
        OrgKindList.Enabled = False

        'Me.lblProecessType.Text = "檢視"
        If Not Page.IsPostBack Then
            center.Text = TIMS.GET_OrgName(OrgIDValue.Value, objconn) 'sm.UserInfo.OrgName
            RIDValue.Value = Me.ViewState("Re_rid") 'sm.UserInfo.RID
            'HistoryTable.Disabled = True
            HistoryTable.Disabled = False
            SearchHistory(Me.TBID.Text, "", "", Me.ViewState("Re_planid"), Me.ViewState("Re_rid"), Me.ViewState("AppliedResult"))
        End If

        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        'If HistoryRID.Rows.Count <> 0 Then
        '    center.Attributes("onclick") = "showObj('HistoryList2');"
        '    center.Style("CURSOR") = "hand"
        'End If
        'HistoryRID.Visible = False

        'BtnOrg.Visible = False
        '?btnName=Button1

        Dim s_javascript_openOrg_FMT1 As String = String.Concat("javascript:openOrg('../../Common/LevOrg{0}.aspx?TC_01_016_add=", sm.UserInfo.Years, "&btnName=btnPlanSearch');")
        BtnOrg.Attributes("onclick") = String.Format(s_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

    End Sub

    Function create1() As Boolean
        Dim rst As Boolean = True  '設定顯示資料
        'create1 = True  '設定顯示資料
        'Show_OrgInfoData
        If Not Show_OrgInfoData(Me.ViewState("Re_orgid"), Me.ViewState("Re_planid"), Me.ViewState("Re_rid"), msg.Text) Then
            rst = False
            Return rst 'Exit Function
        End If

        ''Get_OrgPlanNameList1
        'If Get_OrgPlanNameList1(Me.TBID.Text, Me.ViewState("RWPlanRID")) = False Then
        '    create1 = False
        '    Exit Function
        'End If
        'Get_OrgPlanNameList1
        'Me.ViewState("Re_rid")

        If Get_OrgPlanNameList1(Me.ViewState("Re_rid"), Me.ViewState("RWPlanRID")) = False Then
            rst = False
            Return rst 'Exit Function
        End If
        Return rst
    End Function

    Sub Check_TPlanID28()
        '-若為產業人才投資方案計劃則顯示轉用輸入表格

        '**by Milor 20080522--所有的計畫都要顯示計畫主持人，不限制只有產學訓----start
        Dim TrShow As Boolean = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then TrShow = True
        'TPlanID28A.Visible = TPlanID28.Visible
        LastYear.Value = (CInt(sm.UserInfo.Years) - 1).ToString

        '**by Milor 20080502--User依據新的使用手冊要求把動態顯示改為直接顯示上年度字樣----start
        LabLastYear.Text = "上年度是否辦理本計劃"
        LabLastYear2.Text = "上年度核定人數執行率"
        '**by Milor 20080502----end
    End Sub

    Function Show_OrgInfoData(ByVal OrgID As String, ByVal PlanID As String, ByVal RID As String, ByVal Errmsg As String) As Boolean
        Dim rst As Boolean = True
        'Show_OrgInfoData = True
        Dim Parent_list As String = ""
        Dim list As DataRow
        Dim sub_sql As String
        If RID.ToString.Trim.Length = 1 Then PlanID = "0" '分署(中心)自動將計畫清除

        sub_sql = "" & vbCrLf
        sub_sql &= " SELECT a.OrgID, a.OrgName, a.IsConUnit ,a.ComIDNO, a.ComCIDNO ,c.ActNo, c.ZipCode, c.ZipCODE6W, c.Address " & vbCrLf
        sub_sql &= " ,c.phone, c.mastername, c.ContactName ,c.ContactEmail, c.ContactCellPhone ,c.TrainCap, c.FireControlState " & vbCrLf
        sub_sql &= " ,c.ProTrainKind, c.ComSumm ,b.Relship, b.PlanID, b.DistID ,a.OrgKind, c.OrgPName " & vbCrLf
        sub_sql &= " ,c.PlanMaster, c.PlanMasterPhone ,c.ContactFax, a.LastYearExeRate " & vbCrLf
        sub_sql &= " FROM dbo.Org_orginfo a " & vbCrLf
        sub_sql &= " JOIN dbo.Auth_Relship b ON a.orgid = b.orgid " & vbCrLf
        sub_sql &= " JOIN dbo.Org_OrgPlanInfo c ON c.RSID = b.RSID " & vbCrLf
        sub_sql &= " WHERE 1=1 " & vbCrLf
        sub_sql &= " AND a.OrgID = '" & OrgID & "' " & vbCrLf
        sub_sql &= " AND b.PlanID = '" & PlanID & "' " & vbCrLf
        sub_sql &= " AND b.RID = '" & RID & "' " & vbCrLf
        list = DbAccess.GetOneRow(sub_sql, objconn)
        'Me.DistrictList.SelectedValue = list("DistID")

        If list Is Nothing Then
            bt_save.Enabled = False '儲存鈕失效
            bt_save.Visible = False '儲存鈕隱藏
            Errmsg += vbCrLf & "計畫權限取得有誤，請重新輸入查詢值!!"
            rst = False
            Return rst ' Show_OrgInfoData = False Exit Function
        End If
        Dim myarray As Array
        myarray = list("Relship").Split("/")
        Dim range As Integer = myarray.Length - 3
        If range >= 0 Then
            Parent_list = myarray(range)
        Else
            Errmsg += vbCrLf & "計畫權限取得有誤，請重新輸入查詢值!!"
            rst = False
            Return rst ' Show_OrgInfoData = False Exit Function
        End If
        sub_sql = ""
        sub_sql &= " SELECT b.OrgName PlanName"
        sub_sql &= " FROM dbo.Auth_Relship a "
        sub_sql &= " JOIN dbo.org_orginfo b ON a.orgid = b.orgid "
        sub_sql &= " WHERE a.RID = '" & Parent_list & "' "
        Me.ViewState("orgName") = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))

        'c.Years+d.Name+e.PlanName+c.seq+'_'
        sub_sql = ""
        sub_sql &= " SELECT c.Years + d.Name + e.PlanName + c.seq + '_' AS PlanName"
        sub_sql &= " FROM dbo.Auth_Relship a"
        sub_sql &= " JOIN dbo.org_orginfo b ON a.orgid = b.orgid"
        sub_sql &= " JOIN dbo.ID_Plan c ON c.PlanID = a.PlanID "
        sub_sql &= " JOIN dbo.ID_District d ON d.DistID = c.DistID"
        sub_sql &= " JOIN dbo.Key_Plan e ON c.TPlanID = e.TPlanID"
        sub_sql &= " WHERE a.planid = '" & PlanID & "' AND a.RID = '" & RID & "' "
        Me.ViewState("plan_name") = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))
        Me.TBplan.Text = Me.ViewState("plan_name") & Me.ViewState("orgName")

        '是否為管控單位
        Common.SetListItem(IsConUnit, Convert.ToString(list("IsConUnit")))

        If list("OrgID").ToString <> "" Then OrgIDValue.Value = list("OrgID").ToString
        'PlanID = list("PlanID")
        PlanIDValue.Value = list("PlanID")
        TBtitle.Text = list("OrgName")
        TBID.Text = list("ComIDNO")
        ViewState("comidno") = TBID.Text
        TBseqno.Text = list("ComCIDNO")

        TB_ActNo.Text = Convert.ToString(list("ActNo")) 'ActNo

        city_code.Value = Convert.ToString(list("ZipCode"))
        hidZipCODE6W.Value = Convert.ToString(list("ZipCODE6W"))
        ZipCODEB3.Value = TIMS.GetZIPCODEB3(hidZipCODE6W.Value)
        TBCity.Text = TIMS.GET_FullCCTName(objconn, city_code.Value, hidZipCODE6W.Value)
        TBaddress.Text = Convert.ToString(list("Address")) 'list("Address")

        TBseqno.Text = Convert.ToString(list("ComCIDNO")) 'list("ComCIDNO")
        Me.TBtel.Text = Convert.ToString(list("Phone")) 'phone
        Me.TBm_name.Text = Convert.ToString(list("MasterName")) 'master
        Me.TBContactName.Text = Convert.ToString(list("ContactName")) 'Cname

        Me.TBmail.Text = Convert.ToString(list("ContactEmail")) ' Cemail
        Me.TBcontact_cellphone.Text = Convert.ToString(list("ContactCellPhone")) ' C_cell
        Me.TB_TrainCap.Text = Convert.ToString(list("TrainCap")) ' T_cap
        Me.TB_FireControlState.Text = Convert.ToString(list("FireControlState")) ' Fire_con
        Me.TB_ProTrainKind.Text = Convert.ToString(list("ProTrainKind")) ' Pro_train
        Me.ComSumm.Text = Convert.ToString(list("ComSumm"))

        If list("DistID").ToString <> "" Then Common.SetListItem(Me.DistrictList, list("DistID"))
        'Me.OrgKindList.SelectedValue = list("OrgKind")
        Common.SetListItem(OrgKindList, Convert.ToString(list("OrgKind")))

        TB_OrgPName.Text = Convert.ToString(list("OrgPName"))
        PlanMaster.Text = Convert.ToString(list("PlanMaster")) 'list("PlanMaster").ToString
        PlanMasterPhone.Text = Convert.ToString(list("PlanMasterPhone")) 'list("PlanMasterPhone").ToString
        ContactFax.Text = Convert.ToString(list("ContactFax")) 'list("ContactFax").ToString
        Dim flag_LastYearExeRate As Boolean = If(Convert.ToString(list("LastYearExeRate")) <> "", If(Val(list("LastYearExeRate")) < 0, False, True), False)
        '上年度是否辦理本計劃
        Common.SetListItem(LastYearExeRate, If(flag_LastYearExeRate, "1", "-1")) '是"1" '否"-1"
        If flag_LastYearExeRate Then txtLastYearExeRate.Text = Convert.ToString(list("LastYearExeRate"))

        Return rst ' Show_OrgInfoData = False Exit Function
    End Function

    Sub SearchHistory(ByVal ComIDNO_Val As String, ByVal TPlanID_Val As String, ByVal Years_Val As String, ByVal PlanID_Val As String, ByVal RID_Val As String, ByVal AppliedResult As String)
        'ByVal ComIDNO_Val As String, Optional ByVal TPlanID_Val As String = "", Optional ByVal Years_Val As String = "", Optional ByVal PlanID_Val As String = "", Optional ByVal RID_Val As String = "", Optional ByVal AppliedResult As String = ""
        'Dim Sql As String = ""
        'Dim dt As DataTable
        '==測==
        Dim s_pagename As String = "TC_01_016_add"
        'Dim fms_sqlstr As String = String.Format("##{0}, strSql: {1}", s_pagename, strSql)
        Dim ss_parms As String = ""
        TIMS.SetMyValue(ss_parms, "ComIDNO_Val", ComIDNO_Val)
        TIMS.SetMyValue(ss_parms, "TPlanID_Val", TPlanID_Val)
        TIMS.SetMyValue(ss_parms, "Years_Val", Years_Val)
        TIMS.SetMyValue(ss_parms, "PlanID_Val", PlanID_Val)
        TIMS.SetMyValue(ss_parms, "RID_Val", RID_Val)
        TIMS.SetMyValue(ss_parms, "AppliedResult", AppliedResult)
        Dim fms_Param As String = String.Format("##{0}, parms: {1}", s_pagename, ss_parms)
        Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        'Dim slogMsg1 As String = fms_sqlstr & vbCrLf & fms_Param & vbCrLf
        Dim slogMsg1 As String = fms_Param & vbCrLf
        If flag_chktest Then TIMS.WriteLog(Me, slogMsg1)
        '==測==

        If ComIDNO_Val = "" Then Exit Sub
        DataGrid2.CurrentPageIndex = 0

        Dim parms As Hashtable = New Hashtable()
        Dim dt As DataTable
        Dim strSql As String = ""
        strSql = "" & vbCrLf
        strSql &= " SELECT cc.OCID" & vbCrLf
        strSql &= " ,ip.DistID" & vbCrLf
        strSql &= " ,pp.TPlanID" & vbCrLf
        strSql &= " ,kp.PlanName " & vbCrLf
        strSql &= " ,CASE WHEN cc.ocid IS NOT NULL THEN '(' + CONVERT(VARCHAR, cc.ocid) + ')'" & vbCrLf
        strSql &= "  + dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) " & vbCrLf
        strSql &= " ELSE dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) END ClassName " & vbCrLf
        strSql &= " ,pp.ComIDNO,pp.PlanID,pp.SeqNo,oo.ORGNAME" & vbCrLf
        strSql &= " ,pp.RID, pp.PlanYear" & vbCrLf
        strSql &= " ,pp.STDate, pp.FDDate" & vbCrLf
        strSql &= " ,pp.TMID, idt.Name DistName " & vbCrLf
        strSql &= " ,CASE WHEN K2.JobID IS NULL THEN K2.TrainName ELSE K2.JobName END TrainName" & vbCrLf
        strSql &= " ,CONVERT(varchar, pp.STDate, 111) + '<BR>|<BR>' + CONVERT(varchar, pp.FDDate, 111) TRound " & vbCrLf

        'strSql &= " ,RIGHT('000' + CONVERT(VARCHAR, YEAR(pp.STDate)-1911), 3) + FORMAT(pp.STDate, '/MM/dd') " & vbCrLf
        'strSql &= " + '<BR>|<BR>' " & vbCrLf
        'strSql &= " + RIGHT('000' + CONVERT(VARCHAR, YEAR(pp.FDDate)-1911), 3) + FORMAT(pp.FDDate, '/MM/dd') TRound_ROC " & vbCrLf  'edit，by:20181002

        strSql &= " FROM PLAN_PLANINFO pp " & vbCrLf
        strSql &= " LEFT JOIN CLASS_CLASSINFO cc ON cc.planid = pp.planid AND cc.comidno = pp.comidno AND cc.seqno = pp.seqno AND cc.rid = pp.rid " & vbCrLf
        strSql &= " JOIN ID_PLAN ip ON pp.PlanID = ip.PlanID " & vbCrLf
        strSql &= " JOIN ORG_ORGINFO oo on oo.COMIDNO=pp.COMIDNO" & vbCrLf
        strSql &= " JOIN Key_Plan kp ON pp.TPlanID = kp.TPlanID " & vbCrLf
        strSql &= " JOIN Key_TrainType K2 ON pp.TMID = K2.TMID " & vbCrLf
        strSql &= " JOIN ID_District idt ON ip.DistID = idt.DistID " & vbCrLf
        strSql &= " WHERE 1=1 " & vbCrLf
        If Trim(AppliedResult) <> "" Then
            AppliedResult = Trim(AppliedResult)
            Select Case AppliedResult
                Case "X"
                    strSql &= " AND pp.APPLIEDRESULT IS NULL" & vbCrLf
                Case Else
                    strSql &= " AND pp.APPLIEDRESULT = @APPLIEDRESULT " & vbCrLf
                    parms.Add("APPLIEDRESULT", AppliedResult)
            End Select
        Else
            strSql &= " AND pp.AppliedResult = 'Y' " & vbCrLf
        End If
        strSql &= " AND pp.IsApprPaper = 'Y' " & vbCrLf

        If ComIDNO_Val <> "" Then
            strSql &= " AND pp.COMIDNO = @COMIDNO " & vbCrLf
            parms.Add("COMIDNO", ComIDNO_Val)
        End If
        If TPlanID_Val <> "" Then
            strSql &= " AND ip.TPLANID = @TPLANID " & vbCrLf
            parms.Add("TPLANID", TPlanID_Val)
        End If
        If Years_Val <> "" Then
            strSql &= " AND pp.PLANYEAR = @PLANYEAR " & vbCrLf
            parms.Add("PLANYEAR", Years_Val)
        End If
        If PlanID_Val <> "" Then
            strSql &= " AND pp.PLANID = @PLANID " & vbCrLf
            parms.Add("PLANID", PlanID_Val)
        End If
        If RID_Val <> "" Then
            strSql &= " AND pp.RID = @RID " & vbCrLf
            parms.Add("RID", RID_Val)
        End If

        '==測==
        writeLog_1(strSql, parms)
        '==測==

        dt = DbAccess.GetDataTable(strSql, objconn, parms)

        DataGrid2.Visible = False
        'HistoryTable.Style.Item("display") = "none"
        msg.Text = "查無資料!"

        If dt.Rows.Count > 0 Then
            DataGrid2.Visible = True
            'HistoryTable.Style.Item("display") = "inline"
            msg.Text = ""
            'RecordCount.Text = dt.Rows.Count
            If Me.ViewState("sort") Is Nothing Then Me.ViewState("sort") = " PlanYear,TRound,ClassName,DistID"
            dt.DefaultView.Sort = Me.ViewState("sort")
            'PageControler1.PageDataTable = dt
            'PageControler1.Sort = Me.ViewState("sort") ' "IDNO,Birthday,TRound"
            'PageControler1.ControlerLoad()
            DataGrid2.DataSource = dt
            DataGrid2.DataBind()

            dt.Dispose()
            dt = Nothing
        End If
    End Sub

    '==測==
    Sub writeLog_1(ByRef strSql As String, ByRef parms As Hashtable)
        Dim s_pagename As String = "TC_01_016_add"
        Dim fms_sqlstr As String = String.Format("##{0}, strSql: {1}", s_pagename, strSql)
        Dim fms_Param As String = ""
        If parms IsNot Nothing Then
            fms_Param = String.Format("##{0}, parms: {1}", s_pagename, TIMS.GetMyValue3(parms))
        End If
        Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()

        Dim slogMsg1 As String = ""
        slogMsg1 = fms_sqlstr & vbCrLf
        If fms_Param <> "" Then slogMsg1 &= fms_Param & vbCrLf

        If flag_chktest Then TIMS.WriteLog(Me, slogMsg1)
    End Sub

    '共用機構取得 (若無共機計畫則回傳 False)
    Function Get_OrgPlanNameList1(ByVal RID As String, ByVal RWPlanRID_Val As String) As Boolean
        Dim rst As Boolean = True
        'Get_OrgPlanNameList1 = True
        Dim sqlstr As String = ""
        Dim dt As DataTable

#Region "(No Use)"

        'sqlstr = "" & vbCrLf
        'sqlstr += " select a.orgid,b.RSID,b.relship,a.OrgName,c.name,a.ComIDNO" & vbCrLf
        'sqlstr += " ,f.Address,b.distid,b.PlanID,b.RID,d.PlanName" & vbCrLf
        'sqlstr += " ,f.ActNo,f.ContactName,f.ContactEmail, f.modifyAcct" & vbCrLf
        'sqlstr += " from " & vbCrLf
        'sqlstr += "Org_orginfo a " & vbCrLf
        'sqlstr += "join Auth_Relship b on a.orgid = b.orgid " & vbCrLf
        'sqlstr += "join ID_District c on b.distid=c.distid " & vbCrLf
        'sqlstr += "join view_LoginPlan d on d.PlanID=b.PlanID " & vbCrLf
        ''sqlstr += "join ID_Plan d on d.PlanID=b.PlanID " & vbCrLf
        ''sqlstr += "join Key_Plan e on e.TPlanID=d.TPlanID " & vbCrLf
        'sqlstr += "join Org_OrgPlanInfo f on f.RSID=b.RSID " & vbCrLf
        'sqlstr += " where d.TPlanID NOT IN ('17','28','36') " & vbCrLf

        '移除下列計畫
        'drpPlan.Items.Remove(drpPlan.Items.FindByValue("17")) '補助地方政府訓練
        'drpPlan.Items.Remove(drpPlan.Items.FindByValue("28")) '產業人才投資方案
        'drpPlan.Items.Remove(drpPlan.Items.FindByValue("36")) '青年職涯啟動計畫

#End Region

        'Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT vrp.PlanName" & vbCrLf
        sqlstr &= " ,vrp.RWPlanRID" & vbCrLf
        sqlstr &= " ,vrp.RWPlanID" & vbCrLf
        sqlstr &= " ,vrp.PlanID" & vbCrLf
        sqlstr &= " ,vrp.Relship" & vbCrLf
        sqlstr &= " ,vrp.RID" & vbCrLf
        sqlstr &= " ,vrp.OrgID" & vbCrLf
        sqlstr &= " ,vrp.Years" & vbCrLf
        sqlstr &= " ,vrp.ORGLEVEL" & vbCrLf
        sqlstr &= " ,vrp.COMIDNO" & vbCrLf
        sqlstr &= " FROM dbo.VIEW_RWPLANRID vrp" & vbCrLf
        sqlstr &= " WHERE 1=1" & vbCrLf
        'sqlstr += " AND vrp.OrgLevel>='" & sm.UserInfo.OrgLevel & "'" & vbCrLf '登入層級
        'sqlstr += " AND vrp.Years='" & sm.UserInfo.Years & "'" & vbCrLf '登入年度
        If RID.Length > 1 Then
            sqlstr &= " AND vrp.RID = '" & RID & "'" & vbCrLf ' 業務權限
        Else
            sqlstr &= " AND vrp.OrgID = '" & OrgIDValue.Value & "' " & vbCrLf '同一機構
        End If
        sqlstr &= " AND vrp.RWPlanRID != '" & RWPlanRID_Val & "' " & vbCrLf '不同權限
        '限定年度大於等於登入年度
        sqlstr &= " AND vrp.Years >= '" & sm.UserInfo.Years & "' " & vbCrLf
        sqlstr &= " ORDER BY vrp.Years, vrp.RWPlanRID " & vbCrLf '排序依年度

        '==測==
        writeLog_1(sqlstr, Nothing)
        '==測==

        'PlanName
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "此機構計畫尚未共用")
            rst = False
            Return rst 'Exit Function
        End If

        '重整 PlanName
        Dim sub_plan_name As String = ""
        Dim sub_orgName As String = ""
        Dim sub_sql As String = ""
        Dim Parent_list_RID As String = ""

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim list As DataRow = dt.Rows(i)

            If list("PlanID") <> 0 Then
                If Convert.ToString(list("ORGLEVEL")) <> "2" Then
                    Common.MessageBox(Me, "計畫權限取得有誤，請重新輸入查詢值!!")
                    rst = False
                    Return rst 'Exit Function
                End If

                Parent_list_RID = Left(list("RID"), 1)
                sub_sql = "SELECT b.OrgName FROM AUTH_RELSHIP a JOIN ORG_ORGINFO b ON a.orgid = b.orgid WHERE a.RID = '" & Parent_list_RID & "' "
                sub_orgName = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))

                sub_sql = "" & vbCrLf
                sub_sql &= " SELECT c.Years + d.Name + e.PlanName + c.seq + '_' PlanName " & vbCrLf
                sub_sql &= " FROM Auth_Relship a " & vbCrLf
                sub_sql &= " JOIN org_orginfo b ON a.orgid = b.orgid " & vbCrLf
                sub_sql &= " JOIN ID_Plan c ON c.PlanID = a.PlanID " & vbCrLf
                sub_sql &= " JOIN ID_District d ON d.DistID = c.DistID " & vbCrLf
                sub_sql &= " JOIN Key_Plan e ON c.TPlanID = e.TPlanID " & vbCrLf
                sub_sql &= " WHERE a.planid = '" & list("RWPlanID") & "' " & vbCrLf
                sub_sql &= " AND a.RID = '" & list("RID") & "' " & vbCrLf
                sub_plan_name = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))
                list("PlanName") = sub_plan_name + sub_orgName
            End If

            'If list("PlanID") <> 0 Then
            '    Dim myarray As Array
            '    myarray = list("Relship").Split("/")
            '    Dim range As Integer = myarray.Length - 3
            '    If range >= 0 Then
            '        Parent_list = myarray(range)
            '    Else
            '        Common.MessageBox(Me, "計畫權限取得有誤，請重新輸入查詢值!!")
            '        rst = False
            '        Return rst 'Exit Function
            '    End If

            '    sub_sql = ""
            '    sub_sql &= " Select b.OrgName As PlanName"
            '    sub_sql &= " FROM dbo.Auth_Relship a"
            '    sub_sql &= " JOIN dbo.org_orginfo b On a.orgid = b.orgid"
            '    sub_sql &= " WHERE a.RID = '" & Parent_list & "' "
            '    sub_orgName = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))

            '    sub_sql = " SELECT c.Years + d.Name + e.PlanName + c.seq + '_' AS PlanName"
            '    sub_sql &= " FROM dbo.Auth_Relship a "
            '    sub_sql &= " JOIN dbo.org_orginfo b ON a.orgid = b.orgid "
            '    sub_sql &= " JOIN dbo.ID_Plan c ON c.PlanID = a.PlanID "
            '    sub_sql &= " JOIN dbo.ID_District d ON d.DistID = c.DistID "
            '    sub_sql &= " JOIN dbo.Key_Plan e ON c.TPlanID = e.TPlanID "
            '    sub_sql &= " WHERE a.planid = '" & list("RWPlanID") & "' AND a.RID = '" & list("RID") & "' "
            '    sub_plan_name = Convert.ToString(DbAccess.ExecuteScalar(sub_sql, objconn))
            '    list("PlanName") = sub_plan_name & sub_orgName
            'End If
        Next

        With OrgPlanNameList
            .Items.Clear()
            .DataSource = dt
            .DataTextField = "PlanName"
            .DataValueField = "RWPlanRID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        Return rst 'Exit Function
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Session("_Search") = Me.ViewState("OrgSearchStr")
        Me.ViewState("OrgSearchStr") = Nothing
        TIMS.Utl_Redirect1(Me, "TC_01_016.aspx?ID=" & Request("ID") & "")
#Region "(No Use)"

        'If Me.ViewState("Redirect") Is Nothing Then
        '    Session("_Search") = Me.ViewState("OrgSearchStr")
        '    Me.ViewState("OrgSearchStr") = Nothing
        '   TIMS.Utl_Redirect1(Me, "TC_01_016.aspx?ID=" & Request("ID") & "")
        'Else
        '   TIMS.Utl_Redirect1(Me, Me.ViewState("Redirect") & "?ID=" & Request("ID") & "")
        'End If

#End Region
    End Sub

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        Dim s_Errmsg As String = ""
        Hid_Errmsg.Value = "" '= Nothing

        Dim v_OrgPlanNameList As String = TIMS.GetListValue(OrgPlanNameList) 'OrgPlanNameList.SelectedValue
        If v_OrgPlanNameList = "" Then Hid_Errmsg.Value &= "請選擇，計畫調動位置" & vbCrLf

        Me.ViewState("Class_SeqNo") = ""
        Me.ViewState("Class_SeqNo") = Get_PlanClassSeqNo(s_Errmsg)
        If s_Errmsg <> "" Then Hid_Errmsg.Value &= s_Errmsg

        If Hid_Errmsg.Value = "" Then
            'Common.MessageBox(Me, "計畫調動開始")
            Chang_OrgPlanClass(Me.ViewState("RWPlanRID"), v_OrgPlanNameList, Me.ViewState("Class_SeqNo"))
            SearchHistory(Me.TBID.Text, "", "", Me.ViewState("Re_planid"), Me.ViewState("Re_rid"), Me.ViewState("AppliedResult"))
        Else
            Common.MessageBox(Me, Hid_Errmsg.Value)
            'If Not create1() Then Exit Sub
            Exit Sub
        End If
    End Sub

    Function Get_PlanClassSeqNo(ByRef Errmsg As String) As String
        Errmsg = ""
        Dim tmpSeqNo As String = ""
        Dim intRows As Integer = 0
        For Each Item As DataGridItem In DataGrid2.Items
            Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
            Dim SeqNO As HtmlInputHidden = Item.FindControl("SeqNO")
            If Checkbox1.Checked Then
                intRows += 1
                If tmpSeqNo <> "" Then tmpSeqNo &= ","
                tmpSeqNo &= SeqNO.Value
            End If
        Next
        If intRows = 0 OrElse tmpSeqNo = "" Then Errmsg += "必須要有選擇班級" & vbCrLf
        Return tmpSeqNo
    End Function

    '儲存
    Sub Chang_OrgPlanClass(ByVal RWPlanRID1 As String, ByVal RWPlanRID2 As String, ByVal strSeqNO As String)
        Dim sql As String = ""
        Dim dtC As DataTable
        Dim dtP As DataTable
        Dim dr As DataRow
        Dim iMinSeqNo1 As Integer = 0 '來源最小序號
        Dim iMaxSeqNo1 As Integer = 1 '來源最大序號
        Dim iMaxSeqNo2 As Integer = 0 '目標最大序號
        Dim intDiffSeqNo As Integer = 0 '差距數
        Dim dr1 As DataRow
        Dim dr2 As DataRow
        Dim sub_sql As String = ""
        'Dim strSeqNO As String = ""
        Dim strOCID As String = ""
        Dim strWhere As String = ""
        'Dim str_now As String = ""
        Dim str_user As String = "-" & sm.UserInfo.UserID '有調動者使用「-」為依據，來做為日後判斷

        'sub_sql = ""
        'sub_sql += "SELECT getdate() now  " & vbCrLf
        'sub_sql += "FROM view_RIDName vr " & vbCrLf
        'Common.FormatNow(TIMS.GetSysDateNow(objconn))
        'str_now = Common.FormatNow(DbAccess.ExecuteScalar(sub_sql, objconn))
        Dim str_now As String = ""
        str_now = Common.FormatNow(TIMS.GetSysDateNow(objconn))

        '確認機構是否共用1
        sub_sql = ""
        sub_sql &= " SELECT vr.*, oo.ComIDNO, ip.TPlanID " & vbCrLf
        sub_sql &= " FROM dbo.VIEW_RWPLANRID vr " & vbCrLf
        sub_sql &= " JOIN Org_OrgInfo oo ON oo.OrgID = vr.OrgID " & vbCrLf
        sub_sql &= " JOIN id_Plan ip ON ip.PlanID = vr.RWPlanID " & vbCrLf
        sub_sql &= " WHERE 1=1 "
        sub_sql &= " AND vr.RWPlanRID = '" & RWPlanRID1 & "' "
        dr1 = DbAccess.GetOneRow(sub_sql, objconn)

        '確認機構是否共用2
        sub_sql = ""
        sub_sql &= " SELECT vr.*, oo.ComIDNO, ip.TPlanID " & vbCrLf
        sub_sql &= " FROM dbo.VIEW_RWPLANRID vr " & vbCrLf
        sub_sql &= " JOIN Org_OrgInfo oo ON oo.OrgID = vr.OrgID " & vbCrLf
        sub_sql &= " JOIN id_Plan ip ON ip.PlanID = vr.RWPlanID " & vbCrLf
        sub_sql &= " WHERE 1=1 "
        sub_sql &= " AND vr.RWPlanRID = '" & RWPlanRID2 & "' "
        dr2 = DbAccess.GetOneRow(sub_sql, objconn)

        If dr1 Is Nothing Or dr2 Is Nothing Then
            Common.MessageBox(Me, "轉換計畫資訊有誤，請洽系統管理人員！")
            Exit Sub
        End If
        If strSeqNO = "" Then
            Common.MessageBox(Me, "未選擇要調動的班級，請確認")
            Exit Sub
        End If

        'Exit Sub
        'sub_sql1
        Me.ViewState("PlanYear1") = dr1("Years")
        Me.ViewState("Years1") = Right(dr1("Years").ToString, 2)
        Me.ViewState("TPlanID1") = dr1("TPlanID")
        Me.ViewState("PlanID1") = dr1("RWPlanID")
        Me.ViewState("ComIDNO1") = dr1("ComIDNO")
        Me.ViewState("RID1") = dr1("RID")

        'sub_sql2
        Me.ViewState("PlanYear2") = dr2("Years")
        Me.ViewState("Years2") = Right(dr2("Years").ToString, 2)
        Me.ViewState("TPlanID2") = dr2("TPlanID")
        Me.ViewState("PlanID2") = dr2("RWPlanID")
        Me.ViewState("ComIDNO2") = dr2("ComIDNO") '應與 Me.ViewState("ComIDNO1") 相同
        Me.ViewState("RID2") = dr2("RID")

        strWhere = ""
        strWhere += " PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
        strWhere += " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
        strWhere += " AND TPlanID = '" & Me.ViewState("TPlanID1") & "' " & vbCrLf
        strWhere += " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
        strWhere += " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf '使用傳入的 SeqNO
        dtP = TIMS.Get_KeyTable("Plan_PlanInfo", strWhere, objconn)
        strSeqNO = ""
        If dtP.Rows.Count > 0 Then
            For i As Integer = 0 To dtP.Rows.Count - 1
                dr = dtP.Rows(i)
                If strSeqNO <> "" Then strSeqNO &= ","
                strSeqNO &= dr("SeqNO")
            Next
        End If

        If strSeqNO = "" Then Exit Sub '重新確認符合條件的 SeqNO

        strWhere = ""
        strWhere += " PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
        strWhere += " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
        strWhere += " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
        strWhere += " AND SeqNO IN (" & strSeqNO & ")  " & vbCrLf
        dtC = TIMS.Get_KeyTable("Class_ClassInfo", strWhere, objconn)
        strOCID = "" '取得符合條件的 OCID
        If dtC.Rows.Count > 0 Then
            For i As Integer = 0 To dtC.Rows.Count - 1
                dr = dtC.Rows(i)
                If Convert.ToString(dr("OCID")) <> "" Then
                    If strOCID <> "" Then strOCID &= ","
                    strOCID &= dr("OCID")
                End If
            Next
        End If

        sql = ""
        sql = " SELECT ISNULL(Min(SeqNO),0) SeqNO FROM Plan_PlanInfo " & vbCrLf
        sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
        sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
        sql &= " AND TPlanID = '" & Me.ViewState("TPlanID1") & "' " & vbCrLf
        sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
        sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
        iMinSeqNo1 = DbAccess.ExecuteScalar(sql, objconn) '來源最小序號

        sql = ""
        sql = " SELECT ISNULL(MAX(SeqNO),0) SeqNO FROM Plan_PlanInfo " & vbCrLf
        sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
        sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
        sql &= " AND TPlanID = '" & Me.ViewState("TPlanID1") & "' " & vbCrLf
        sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
        sql &= " AND SeqNO IN (" & strSeqNO & ")  " & vbCrLf
        iMaxSeqNo1 = DbAccess.ExecuteScalar(sql, objconn) '來源最大序號

        sql = ""
        sql = " SELECT ISNULL(MAX(SeqNO),0) SeqNO FROM Plan_PlanInfo " & vbCrLf
        sql &= " WHERE PlanID = '" & Me.ViewState("PlanID2") & "' " & vbCrLf
        sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO2") & "' " & vbCrLf
        'sql += " AND TPlanID='" & Me.ViewState("TPlanID2") & "' " & vbCrLf
        'sql += " AND RID='" & Me.ViewState("RID2") & "' " & vbCrLf
        iMaxSeqNo2 = DbAccess.ExecuteScalar(sql, objconn) '目標最大序號

        If Not iMaxSeqNo2 < iMinSeqNo1 Then intDiffSeqNo = iMaxSeqNo2 - iMinSeqNo1 + 1 '差距數相加1

        '下列功能為避免重複序號(1.2衝突)
        Dim dt1 As DataTable
        Dim flag1 As Boolean = True
        sql = "" & vbCrLf
        sql &= " SELECT * FROM Plan_PlanInfo " & vbCrLf
        sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
        sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
        'sql += " AND TPlanID='" & Me.ViewState("TPlanID1") & "' " & vbCrLf
        'sql += " AND RID='" & Me.ViewState("RID1") & "' " & vbCrLf
        sql &= " AND SeqNO IN (" & strSeqNO & ")  " & vbCrLf
        dt1 = DbAccess.GetDataTable(sql, objconn)
        Do While True '試著判斷有無衝突，直到無衝突為止
            flag1 = True '無衝突
            Dim Tempdr As DataRow '暫用 dr (Tempdr)
            Tempdr = Nothing
            For i As Integer = 0 To dt1.Rows.Count - 1
                dr = dt1.Rows(i)
                sql = "" & vbCrLf
                sql &= " SELECT * FROM Plan_PlanInfo " & vbCrLf
                sql &= " WHERE SeqNO = " & dr("SeqNO") + intDiffSeqNo & " " & vbCrLf
                sql &= " AND PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                Tempdr = DbAccess.GetOneRow(sql, objconn)
                If Not Tempdr Is Nothing Then flag1 = False '衝突

                '無衝突
                Tempdr = Nothing
                If flag1 Then
                    sql = "" & vbCrLf
                    sql &= " SELECT * FROM Plan_PlanInfo " & vbCrLf
                    sql &= " WHERE SeqNO = " & dr("SeqNO") + intDiffSeqNo & " " & vbCrLf
                    sql &= " AND PlanID = '" & Me.ViewState("PlanID2") & "' " & vbCrLf
                    sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO2") & "' " & vbCrLf
                    Tempdr = DbAccess.GetOneRow(sql, objconn)
                    If Not Tempdr Is Nothing Then flag1 = False '衝突
                End If
            Next

            If Not flag1 Then intDiffSeqNo += 1 '衝突+1
            If flag1 Then Exit Do  '完全無衝突離開
        Loop

        'PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost
        'Dim s_table_rows As String = cst_table_rows '"PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost"

        'Dim da As SqlDataAdapter = nothing
        Dim tConn As SqlConnection = DbAccess.GetConnection
        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            TIMS.DeleteCookieTable(Me.ViewState("PlanID1"), trans)
            TIMS.DeleteCookieTable(Me.ViewState("PlanID2"), trans)

#Region "(No Use)"

            'trans = DbAccess.BeginTrans(tConn)
            ''序號問題有衝突，重新定義序號
            'If Not MaxSeqNo2 < MinSeqNo1 Then
            '    Do Until MaxSeqNo2 < MinSeqNo1 '直到 來源最小序號 大於 目標最大序號
            '    Loop
            'End If

#End Region

            If intDiffSeqNo > 0 Then
                '序號問題有衝突，重新定義序號
                'Plan_PlanInfo
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_PlanInfo " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                'sql += " AND TPlanID = '" & Me.ViewState("TPlanID1") & "' " & vbCrLf
                'sql += " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_Revise
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_Revise " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_CostItem
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_CostItem " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_Teacher
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_Teacher " & vbCrLf
                sql &= " SET SeqNO= SeqNO + " & intDiffSeqNo & " ,ModifyAcct='" & str_user & "' ,ModifyDate=" & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID='" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO='" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_TrainDesc
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_TrainDesc " & vbCrLf
                sql &= " SET SeqNO= SeqNO + " & intDiffSeqNo & " ,ModifyAcct='" & str_user & "' ,ModifyDate=" & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID='" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO='" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_TrainDesc_Revise
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_TrainDesc_Revise " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " " & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_VerReport
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_VerReport " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Plan_VerRecord
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Plan_VerRecord " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost
                'Dim s_table_rows As String = cst_table_rows '"PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost"
                For Each s_tbn As String In cst_table_rows.Split(",")
                    '修改 來源最小序號
                    sql = "" & vbCrLf
                    sql &= " UPDATE " & s_tbn & vbCrLf
                    sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                    sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                    sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                    sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                    DbAccess.ExecuteNonQuery(sql, trans)
                Next

                If strOCID <> "" Then
                    'Class_ClassInfo
                    '修改 來源最小序號
                    sql = "" & vbCrLf
                    sql &= " UPDATE Class_ClassInfo " & vbCrLf
                    sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                    sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                    sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                    sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                    sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                    DbAccess.ExecuteNonQuery(sql, trans)
                End If

                'Sys_DelLog
                '修改 來源最小序號
                sql = "" & vbCrLf
                sql &= " UPDATE Sys_DelLog " & vbCrLf
                sql &= " SET SeqNO = SeqNO + " & intDiffSeqNo & " ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                '取得最小值
                'sql = ""
                'sql = "SELECT ISNULL(Min(SeqNO),0) AS SeqNO FROM Plan_PlanInfo " & vbCrLf
                'sql += " WHERE PlanID='" & Me.ViewState("PlanID1") & "' " & vbCrLf
                'sql += " AND ComIDNO='" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                'MinSeqNo1 = DbAccess.ExecuteScalar(sql, trans)

                ''取得最大值()
                'sql = ""
                'sql = "SELECT ISNULL(MAX(SeqNO),0) AS SeqNO FROM Plan_PlanInfo " & vbCrLf
                'sql += " WHERE PlanID='" & Me.ViewState("PlanID2") & "' " & vbCrLf
                'sql += " AND ComIDNO='" & Me.ViewState("ComIDNO2") & "' " & vbCrLf
                'MaxSeqNo2 = DbAccess.ExecuteScalar(sql, trans)

                'strWhere = ""
                'strWhere += " PlanID='" & Me.ViewState("PlanID1") & "' " & vbCrLf
                'strWhere += " AND ComIDNO='" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                'strWhere += " AND TPlanID='" & Me.ViewState("TPlanID1") & "' " & vbCrLf
                'strWhere += " AND RID='" & Me.ViewState("RID1") & "' " & vbCrLf
                'dtP = TIMS.Get_KeyTable("Plan_PlanInfo", strWhere)
                'strOCID = ""

                '因為有差距值，所以重新設定 SeqNO
                strSeqNO = ""
                If dtP.Rows.Count > 0 Then
                    For i As Integer = 0 To dtP.Rows.Count - 1
                        dr = dtP.Rows(i)
                        If strSeqNO <> "" Then strSeqNO &= ","
                        strSeqNO &= Val(dr("SeqNO") + intDiffSeqNo).ToString
                    Next
                End If
            End If

            '序號問題沒有衝突 ( MaxSeqNo2 < MinSeqNo1 )
            'Plan_PlanInfo
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_PlanInfo " & vbCrLf
            sql &= " SET TPlanID = '" & Me.ViewState("TPlanID2") & "' " & vbCrLf
            sql &= "  ,PlanID = '" & Me.ViewState("PlanID2") & "' " & vbCrLf
            sql &= "  ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' " & vbCrLf
            sql &= "  ,RID = '" & Me.ViewState("RID2") & "' " & vbCrLf
            sql &= "  ,PlanYear = '" & Me.ViewState("PlanYear2") & "' " & vbCrLf
            sql &= "  ,ModifyAcct = '" & str_user & "' " & vbCrLf
            sql &= "  ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_CostItem
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_CostItem " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_Revise
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_Revise " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_Teacher
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_Teacher " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct ='" & str_user & "' ,ModifyDate =" & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID ='" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO ='" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_TrainDesc
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_TrainDesc " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_TrainDesc_Revise
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_TrainDesc_Revise " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' " & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ")  " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_VerReport
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_VerReport " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'Plan_VerRecord
            sql = "" & vbCrLf
            sql &= " UPDATE Plan_VerRecord " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            'PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost
            'Dim s_table_rows As String = cst_table_rows '"PLAN_DEPOT,Plan_BusPackage,Plan_PersonCost,Plan_CommonCost,Plan_SheetCost,Plan_OtherCost"
            For Each s_tbn As String In cst_table_rows.Split(",")
                sql = "" & vbCrLf
                sql &= " UPDATE " & s_tbn & vbCrLf
                sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)
            Next

            If strOCID <> "" Then
                'Class_ClassInfo
                sql = "" & vbCrLf
                sql &= " UPDATE Class_ClassInfo " & vbCrLf
                sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,RID = '" & Me.ViewState("RID2") & "' ,Years = '" & Me.ViewState("Years2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
                sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                sql &= " AND SeqNO IN (" & strSeqNO & ")  " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Stud_SelResult 
                sql = "" & vbCrLf
                sql &= " UPDATE Stud_SelResult " & vbCrLf
                sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND RID='" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Stud_EnterType2 
                sql = "" & vbCrLf
                sql &= " UPDATE Stud_EnterType2 " & vbCrLf
                sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID1 IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Stud_EnterType
                sql = "" & vbCrLf
                sql &= " UPDATE Stud_EnterType " & vbCrLf
                sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
                sql &= " AND RID='" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID1 IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Stud_DataLid
                sql = "" & vbCrLf
                sql &= " UPDATE Stud_DataLid " & vbCrLf
                sql &= " SET TPlanID= '" & Me.ViewState("TPlanID2") & "' ,RID= '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE TPlanID = '" & Me.ViewState("TPlanID1") & "' " & vbCrLf
                sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Org_StudRecord
                sql = "" & vbCrLf
                sql &= " UPDATE Org_StudRecord " & vbCrLf
                sql &= " SET RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Class_Visitor
                sql = "" & vbCrLf
                sql &= " UPDATE Class_Visitor " & vbCrLf
                sql &= " SET RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Class_UnexpectVisitor
                sql = "" & vbCrLf
                sql &= " UPDATE Class_UnexpectVisitor " & vbCrLf
                sql &= " SET RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE RID='" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Class_UnexpectTel
                sql = "" & vbCrLf
                sql &= " UPDATE Class_UnexpectTel " & vbCrLf
                sql &= " SET RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)

                'Class_RestTime
                sql = "" & vbCrLf
                sql &= " UPDATE Class_RestTime " & vbCrLf
                sql &= " SET TPlanID = '" & Me.ViewState("TPlanID2") & "' ,RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
                sql &= " WHERE TPlanID = '" & Me.ViewState("TPlanID1") & "' " & vbCrLf
                sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
                sql &= " AND OCID IN (" & strOCID & ") " & vbCrLf
                DbAccess.ExecuteNonQuery(sql, trans)
            End If

            'Sys_DelLog
            sql = "" & vbCrLf
            sql &= " UPDATE Sys_DelLog " & vbCrLf
            sql &= " SET PlanID = '" & Me.ViewState("PlanID2") & "' ,ComIDNO = '" & Me.ViewState("ComIDNO2") & "' ,RID = '" & Me.ViewState("RID2") & "' ,ModifyAcct = '" & str_user & "' ,ModifyDate = " & TIMS.To_date(str_now) & vbCrLf
            sql &= " WHERE PlanID = '" & Me.ViewState("PlanID1") & "' " & vbCrLf
            sql &= " AND ComIDNO = '" & Me.ViewState("ComIDNO1") & "' " & vbCrLf
            sql &= " AND RID = '" & Me.ViewState("RID1") & "' " & vbCrLf
            sql &= " AND SeqNO IN (" & strSeqNO & ") " & vbCrLf
            DbAccess.ExecuteNonQuery(sql, trans)

            DbAccess.CommitTrans(trans)
            Call TIMS.CloseDbConn(tConn)
            'DbAccess.RollbackTrans(trans) '測試回復
            Common.MessageBox(Me, "計畫調動完成!!")
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(tConn)
            Common.MessageBox(Me, "計畫調動失敗，請重新再試或連絡系統人員!!")
            'Throw ex
        End Try
        Call TIMS.CloseDbConn(tConn)
    End Sub

    Private Sub btnPlanSearch_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlanSearch.ServerClick
        Show_OrgInfoData(Me.ViewState("Re_orgid"), Me.ViewState("Re_planid"), Me.ViewState("Re_rid"), msg.Text)
        Get_OrgPlanNameList1(Me.RIDValue.Value, Me.ViewState("RWPlanRID"))
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                '參考SD_01_002_R
                Dim SeqNO As HtmlInputHidden = e.Item.FindControl("SeqNO")
                SeqNO.Value = "'" & drv("SeqNO") & "'"
                If Convert.ToString(drv("OCID")) = "" Then TIMS.Tooltip(e.Item.Cells(0), "班級尚未建立", True)
                'If flag_ROC Then
                '    e.Item.Cells(3).Text = String.Format("{0:000}", Convert.ToInt32(drv("PlanYear").ToString.Trim) - 1911)  'edit，by:20181001
                '    e.Item.Cells(7).Text = drv("TRound_ROC").ToString.Trim  'edit，by:20181001
                'End If
        End Select
    End Sub

    Protected Sub OrgPlanNameList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OrgPlanNameList.SelectedIndexChanged

    End Sub
End Class