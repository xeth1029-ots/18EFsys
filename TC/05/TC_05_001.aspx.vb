Partial Class TC_05_001
    Inherits AuthBasePage

#Region "(No Use)"

    'Const Cst_訓練期間 As Integer = 1
    'Const Cst_訓練時段 As Integer = 2
    'Const Cst_訓練地點 As Integer = 3
    'Const Cst_課程編配 As Integer = 4
    'Const Cst_訓練師資 As Integer = 5
    'Const Cst_班別名稱 As Integer = 6
    'Const Cst_期別 As Integer = 7
    'Const Cst_上課地址 As Integer = 8
    'Const Cst_停辦 As Integer = 9
    'Const Cst_上課時段 As Integer = 10
    'Const Cst_師資 As Integer = 11
    'Const Cst_助教 As Integer = 20  '20120213 BY AMU (產投用助教)
    'Const Cst_核定人數 As Integer = 12  'Cst_招生人數  as Integer = 12
    'Const Cst_增班 As Integer = 13
    'Const Cst_科場地 As Integer = 14
    'Const Cst_上課時間 As Integer = 15
    'Const Cst_其他 As Integer = 16
    'Const Cst_報名日期 As Integer = 17  '20080825 andy  add 報名日期
    'Const Cst_課程表 As Integer = 18  '20080626 andy add 課程表
    'Const Cst_包班種類 As Integer = 19  '20111208 BY AMU 
    'Const Cst_訓練費用 As Integer = 21  '20170908 (職前)

#End Region


    '技檢訓練時數 '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_t1 As String = "技檢訓練時數,目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時可儲存，若不符合上述條件，該資料不會存入資料庫。"
    '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_Use_TMID As String = "672"

    'v_SearchMode
    Const cst_sess_tc05001_search1 As String = ""
    Const Cst_cPlanInfo As String = "Plan_PlanInfo"  '申請
    Const Cst_cRevise As String = "Plan_Revise"  '變更結果
    'Dim ChgItemName As Array  '儲存變更項目名稱的陣列
    Dim ChgItemName As String() '將變更項目名稱定義到陣列之中
    Dim rqMID As String = "" 'TIMS.Get_MRqID(Me)

    Const cst_But_Dir_txt_查詢 As String = "查詢"
    Const cst_But_Dir_txt_修改 As String = "修改"
    Const cst_BTN_OL_SUBVIEW_txt_查看 As String = "查看"

    'Cells / Columns
    'CACCTNAME

    'Const Cst_pl28_年度計畫 As Integer = 0 'PlanYear
    'Const Cst_pl28_班別名稱 As Integer = 1 'ClassName2
    Const Cst_pl28_訓練職類 As Integer = 2 'TrainName
    Const Cst_pl28_變更項目 As Integer = 3 '(labAltDataID)
    Const Cst_pl28_申請變更日 As Integer = 4 'CDate
    Const Cst_pl28_申請人姓名 As Integer = 5 'sql &= " ,za.NAME REVISEACCT_Name" & vbCrLf 'REVISEACCT_Name
    Const Cst_pl28_審核時間 As Integer = 6 'modifydate
    Const Cst_pl28_計畫狀態 As Integer = 7 '(labPrjstatus),計畫狀態/申請狀態
    'Const Cst_pl28_功能 As Integer = 8 '(But_Dir)
    'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
    Const Cst_pl28_線上送件 As Integer = 9 'BTN_OL_EDIT1,BTN_OL_SEND1／'(線上送件)編輯／'(線上送件)送出
    Const Cst_pl28_列印 As Integer = 10 '(Button3,bt_print)

    Dim gsCmd As SqlCommand
    'Dim iPYNum14 As Integer = 1 'TIMS.sUtl_GetPYNum14(Me)
    Dim prtFilename As String = "" '列印表件名稱

    'Const Cst_pl_年度計畫 As Integer = 0 'PlanYear
    'Const Cst_pl_班別名稱 As Integer = 1 'ClassName2
    'Const Cst_pl_訓練職類 As Integer = 2 'TrainName
    Const Cst_pl_變更項目 As Integer = 3 '(labAltDataID)
    Const Cst_pl_申請變更日 As Integer = 4 'CDate
    Const Cst_pl_申請人姓名 As Integer = 5 'sql &= " ,za.NAME REVISEACCT_Name" & vbCrLf 'REVISEACCT_Name
    Const Cst_pl_審核時間 As Integer = 6 'modifydate
    Const Cst_pl_計畫狀態 As Integer = 7 '(labPrjstatus),計畫狀態/申請狀態
    'Const Cst_pl_功能 As Integer = 8 '(But_Dir)
    Const Cst_計畫狀態_Header_txt_計畫狀態 As String = "計畫狀態"
    Const Cst_計畫狀態_Header_txt_申請狀態 As String = "申請狀態"

    Const cst_CommandName_But_Dir_appChg As String = "appChg"
    'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
    Const cst_CommandName_BTN_OL_EDIT1 As String = "OL_EDIT1" '(線上送件)編輯
    Const cst_CommandName_BTN_OL_SEND1 As String = "OL_SEND1" '(線上送件)送出
    Const cst_CommandName_BTN_OL_DEL1 As String = "OL_DEL1" '(線上送件)刪除

    '審核通過Y / 審核後修正O / 審核不通過N
    Const cst_AppliedResult_txt_Y As String = "審核通過" '審核通過 AppliedResult Y
    Const cst_AppliedResult_txt_O As String = "審核後修正" '審核後修正 AppliedResult O
    Const cst_AppliedResult_txt_N As String = "審核不通過" '審核不通過 AppliedResult N
    Const cst_AppliedResult_txt_oth As String = "審核中" '審核中 AppliedResult oth
    Const cst_AppliedResult_txt_PARTREDUC_Y As String = "待修正" '待修正 AppliedResult oth
    Dim BlnTest1 As Boolean = False '正式環境為false
    Dim vsMsg2 As String = "" '確認機構是否為黑名單

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        BlnTest1 = TIMS.sUtl_ChkTest()

        Call Utl_EveryCreate1()

        If Not Page.IsPostBack Then
            CCreate1()
        End If

#Region "(No Use)"

        '檢查帳號的功能權限-----------------------------------Start
        'But_Search.Enabled = False
        'If au.blnCanSech Then But_Search.Enabled = True
        ''檢查帳號的功能權限-----------------------------------End

#End Region

    End Sub

    Private Sub Utl_EveryCreate1()
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End
        '將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改
        '產學訓套用的顯示字串  / '非產學訓套用的顯示字串
        ChgItemName = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.TPlanID28ChgItemName, TIMS.TPlanIDChgItemName)

        PageControler1.PageDataGrid = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, PlanList28, PlanList)

        'Dim rqMID As String = ""'TIMS.Get_MRqID(Me)
        rqMID = TIMS.Get_MRqID(Me)

        msg.Text = ""

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        'search_act1
        SearchMode.Attributes("onclick") = "SearchMode_CHGACT1();"

        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then gsCmd = SD_14_010.Get_SEL_REVISE_SQLCMD1(objconn)

        'True '分署(中心)審核 / 'False 委訓單位不顯示 審核完成時間
        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            PlanList28.Columns(Cst_pl28_審核時間).Visible = If(sm.UserInfo.LID = 2, False, True) '分署(中心)審核
        Else
            PlanList.Columns(Cst_pl_審核時間).Visible = If(sm.UserInfo.LID = 2, False, True) '分署(中心)審核
        End If

        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            tr_TMIDVALUE_TP06.Visible = False
            tr_CJOBVALUE_TP06.Visible = False
            'LabTMID 'TB_career_id 'Button1'LabCJOB_UNKEY'btu_sel2
            trainValue.Value = ""
            jobValue.Value = ""
            txtCJOB_NAME.Text = ""
            cjobValue.Value = ""
        Else
            tr_TMIDVALUE_TP06.Visible = True
            tr_CJOBVALUE_TP06.Visible = True
        End If
    End Sub

    Sub CCreate1()
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        hid_USE_PLAN_REVISESUB.Value = TIMS.Utl_GetConfigVAL(objconn, "USE_PLAN_REVISESUB")

        ROC_Years.Value = (sm.UserInfo.Years - 1911)
        divTip.Visible = False
        DataGridTable28.Visible = False
        DataGridTable.Visible = False
        PageControler1.Visible = False

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = TIMS.Get_APPSTAGE2(AppStage) 'If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
        End If
        Call UseKeepSession1()

        '確認機構是否為黑名單 'Dim vsMsg2 As String = ""
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            'Button2.Enabled = False 'TIMS.Tooltip(Button2, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    ''' <summary>機構黑名單內容(訓練單位處分功能)</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        If isBlack.Value = "Y" Then Return True '機構黑名單(訓練單位處分)
        Errmsg = ""
        If sm.UserInfo.OrgID Is Nothing Then Return True
        If Convert.ToString(sm.UserInfo.OrgID) = "" Then Return True
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = String.Concat(sm.UserInfo.OrgName, "，已列入處分名單!!")
            isBlack.Value = "Y"
            orgname.Value = sm.UserInfo.OrgName
            'btnAdd.Visible = False 'Button8.Visible = False
        End If
        Return rst
    End Function

    ''' <summary>
    ''' 計畫狀態/申請狀態 : AppliedResult 申請 plan_planinfo / ReviseStatus 變更結果 PLAN_REVISE
    ''' </summary>
    ''' <param name="drv"></param>
    ''' <param name="v_SearchMode"></param>
    ''' <returns></returns>
    Function Get_txtPrjstatus(ByRef drv As DataRowView, ByVal v_SearchMode As String) As String
        Dim rst As String = ""
        Select Case v_SearchMode
            Case Cst_cPlanInfo 'UCase() '申請
                ''計畫狀態
                Select Case Convert.ToString(drv("AppliedResult")) 'AppliedResult 申請 plan_planinfo / ReviseStatus 變更結果 PLAN_REVISE
                    Case "Y"
                        rst = cst_AppliedResult_txt_Y'"審核通過"
                    Case "O"
                        rst = cst_AppliedResult_txt_O'"審核後修正"
                    Case "N"
                        rst = cst_AppliedResult_txt_N '"審核不通過"
                    Case Else
                        rst = cst_AppliedResult_txt_oth '"審核中"
                End Select

            Case Cst_cRevise 'UCase() '變更結果
                '申請狀態
                Select Case Convert.ToString(drv("AppliedResult")) 'AppliedResult 申請 plan_planinfo / ReviseStatus 變更結果 PLAN_REVISE
                    Case "Y"
                        rst = cst_AppliedResult_txt_Y'"審核通過"
                    Case "O"
                        rst = cst_AppliedResult_txt_O'"審核後修正"
                    Case "N"
                        rst = cst_AppliedResult_txt_N '"審核不通過"
                    Case Else
                        'PLAN_REVISE '"待修正"／'"審核中"
                        rst = If($"{drv("PARTREDUC")}" = TIMS.cst_YES, cst_AppliedResult_txt_PARTREDUC_Y, cst_AppliedResult_txt_oth)
                End Select

        End Select
        Return rst
    End Function

    ''' <summary> 查詢[] </summary>
    Sub sUtl_Search1()
        'Call TIMS.sUtl_TxtPageSize(Me, TxtPageSize, PlanList)
        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, PlanList28)
        Else
            Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, PlanList)
        End If

        msg.Text = "查無資料!!"
        divTip.Visible = False
        DataGridTable28.Visible = False
        DataGridTable.Visible = False
        PageControler1.Visible = False

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        Dim strDataSort As String = ""
        Dim sql As String = ""

        '查詢模式
        Dim v_SearchMode As String = TIMS.GetListValue(SearchMode)
        '審核狀態
        Dim v_CheckMode As String = TIMS.GetListValue(CheckMode)

        Dim rParms As New Hashtable

        Select Case v_SearchMode
            Case Cst_cPlanInfo 'UCase() '申請
                strDataSort = "STDate,ClassName,CDate"

                rParms.Add("RID", RIDValue.Value)

                sql = ""
                sql &= " SELECT a.PlanID ,a.ComIDNO ,a.SeqNo ,a.PlanYear,a.STDate,a.FDDATE FTDATE,a.ClassName,a.CyclType" & vbCrLf
                'ClassName2
                sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSNAME,a.CYCLTYPE) CLASSNAME2" & vbCrLf
                ' AppliedResult 申請 PLAN_PLANINFO / REVISESTATUS 變更結果 PLAN_REVISE
                sql &= " ,a.AppliedResult,0 AltDataID" & vbCrLf
                sql &= " ,format(a.ModifyDate ,'yyyy/MM/dd') CDate" & vbCrLf '申請變更日
                sql &= " ,b.PlanName" & vbCrLf
                sql &= " ,a.TMID,CASE WHEN c.JobID IS NULL THEN concat('[',c.TrainID,']',c.TrainName) ELSE concat('[',c.JobID,']',c.JobName) END TrainName ,f.IsClosed" & vbCrLf
                '20090320 andy 加入變更時間(空欄位，因尚未審核)
                sql &= " ,NULL ModifyDate ,NULL REVISEACCT_Name" & vbCrLf 'REVISEACCT_Name
                sql &= " ,CASE d.ORGKIND WHEN '10' THEN 'W' ELSE 'G' END ORGKINDGW" & vbCrLf
                'objstr += " ,e.DistID,'N' isBlack" & vbCrLf
                sql &= " FROM PLAN_PLANINFO a" & vbCrLf
                sql &= " JOIN KEY_PLAN b ON a.TPlanID=b.TPlanID" & vbCrLf
                sql &= " JOIN KEY_TRAINTYPE c ON a.TMID=c.TMID" & vbCrLf
                sql &= " JOIN ORG_ORGINFO d ON a.ComIDNO=d.ComIDNO" & vbCrLf
                '業務RID
                sql &= " JOIN AUTH_RELSHIP e ON d.OrgID=e.OrgID AND e.RID=@RID" & vbCrLf
                sql &= " JOIN CLASS_CLASSINFO f ON a.PlanID=f.PlanID AND a.ComIDNO=f.ComIDNO AND a.SeqNo=f.SeqNo" & vbCrLf
                sql &= " WHERE a.RID=@RID" & vbCrLf

                '20080716 andy 修正班級變更申請為「不顯示」未轉班資料
                sql &= " AND f.IsSuccess='Y'" & vbCrLf '轉入成功的資料
                sql &= " AND (f.NotOpen IS NULL OR f.NotOpen='N')" & vbCrLf '只顯示要開班的計畫
                sql &= " AND a.IsApprPaper='Y'" & vbCrLf '正式送出的計畫
                '審核狀態
                Select Case v_CheckMode 'CheckMode.SelectedValue
                    Case "1" '審核不通過
                        sql &= " AND a.AppliedResult='N'" & vbCrLf
                    Case "0" '審核完成
                        sql &= " AND a.AppliedResult='Y'" & vbCrLf
                    Case "2" '審核中
                        sql &= " AND a.AppliedResult IS NULL" & vbCrLf
                End Select

            Case Cst_cRevise 'UCase() '變更結果
                strDataSort = "STDate,ClassName,CDate,ModifyDate"

                rParms.Add("RID", RIDValue.Value)
                sql = ""
                sql &= " SELECT a.PlanID ,a.ComIDNO ,a.SeqNo ,a.PlanYear,a.STDate,a.FDDATE FTDATE,a.ClassName,a.CyclType" & vbCrLf
                'ClassName2
                sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSNAME,a.CYCLTYPE) CLASSNAME2" & vbCrLf
                ' AppliedResult 申請 PLAN_PLANINFO / REVISESTATUS 變更結果 PLAN_REVISE
                sql &= " ,z.ReviseStatus AppliedResult" & vbCrLf
                sql &= " ,z.ReviseStatus" & vbCrLf
                sql &= " ,b.PlanName" & vbCrLf
                sql &= " ,a.TMID,CASE WHEN c.JobID IS NULL THEN concat('[',c.TrainID,']',c.TrainName) ELSE concat('[',c.JobID,']',c.JobName) END TrainName" & vbCrLf
                'PLAN_REVISE z
                sql &= " ,format(z.CDate,'yyyy/MM/dd') CDate" & vbCrLf '申請變更日
                sql &= " ,z.SubSeqNo" & vbCrLf
                sql &= " ,z.AltDataID" & vbCrLf
                sql &= " ,f.IsClosed" & vbCrLf
                sql &= " ,z.ModifyDate" & vbCrLf '20090320 andy 加入變更時間
                sql &= " ,za.NAME REVISEACCT_Name" & vbCrLf 'REVISEACCT_Name
                'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
                sql &= " ,z.PARTREDUC,z.REDUCACCT,z.REDUCDATE" & vbCrLf
                'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                sql &= " ,z.ONLINESENDSTATUS,z.ONLINESENDACCT,z.ONLINESENDDATE" & vbCrLf
                sql &= " ,CASE d.ORGKIND WHEN '10' THEN 'W' ELSE 'G' END ORGKINDGW" & vbCrLf
                'objstr += " ,e.DistID ,'N' isBlack" & vbCrLf
                sql &= " FROM PLAN_REVISE z" & vbCrLf
                sql &= " JOIN PLAN_PLANINFO a ON z.PlanID=a.PlanID AND z.ComIDNO=a.ComIDNO AND z.SeqNO=a.SeqNo" & vbCrLf
                sql &= " JOIN KEY_PLAN b ON a.TPlanID=b.TPlanID" & vbCrLf
                sql &= " JOIN KEY_TRAINTYPE c ON a.TMID=c.TMID" & vbCrLf
                sql &= " JOIN ORG_ORGINFO d ON a.ComIDNO=d.ComIDNO" & vbCrLf
                '業務RID
                sql &= " JOIN AUTH_RELSHIP e ON d.OrgID=e.OrgID AND e.RID=@RID" & vbCrLf
                sql &= " JOIN CLASS_CLASSINFO f ON a.PlanID=f.PlanID AND a.ComIDNO=f.ComIDNO AND a.SeqNo=f.SeqNo" & vbCrLf
                sql &= " LEFT JOIN AUTH_ACCOUNT za on za.account=z.REVISEACCT" & vbCrLf
                sql &= " WHERE a.RID=@RID" & vbCrLf
                sql &= " AND f.IsSuccess='Y'" & vbCrLf '轉入成功的資料
                sql &= " AND a.IsApprPaper='Y'" & vbCrLf '正式送出的計畫
                sql &= " AND a.AppliedResult='Y'" & vbCrLf
                '審核狀態
                Select Case v_CheckMode 'CheckMode.SelectedValue
                    Case "1" '審核不通過
                        sql &= " AND z.ReviseStatus='N'" & vbCrLf
                    Case "0" '審核完成
                        sql &= " AND z.ReviseStatus='Y'" & vbCrLf
                    Case "2" '審核中
                        sql &= " AND z.ReviseStatus IS NULL" & vbCrLf
                End Select
        End Select

        'ClassName.Text = Trim(ClassName.Text)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)

        If ClassName.Text <> "" Then
            rParms.Add("lkClassName", ClassName.Text)
            sql &= " AND a.ClassName LIKE '%'+@lkClassName+'%'" & vbCrLf
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'LabTMID.Text = "訓練業別"
            If jobValue.Value <> "" Then
                rParms.Add("TMID", jobValue.Value)
                sql &= " AND (c.TMID=@TMID OR c.TMID IN (" & vbCrLf
                sql &= "  SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '職類別
                sql &= "  SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '業別
                sql &= "  SELECT TMID FROM Key_TrainType WHERE busid = 'G')" & vbCrLf '產業人才投資方案類
                sql &= "  AND TMID=@TMID)))" & vbCrLf
            End If
        Else
            '訓練職類
            If trainValue.Value <> "" Then
                rParms.Add("TMID", trainValue.Value)
                sql &= " AND c.TMID=@TMID" & vbCrLf
            End If
        End If

        '通俗職類
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            rParms.Add("CJOB_UNKEY", cjobValue.Value)
            sql &= " AND a.CJOB_UNKEY =@CJOB_UNKEY" & vbCrLf
        End If

        'PLAN_PLANINFO a
        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If tr_AppStage_TP28.Visible AndAlso v_AppStage <> "" Then
            rParms.Add("APPSTAGE", v_AppStage)
            sql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        End If

        '期別
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then
            rParms.Add("CyclType", CyclType.Text)
            sql &= " AND a.CyclType =@CyclType" & vbCrLf
        End If

        '署(局)與其他
        If sm.UserInfo.RID = "A" Then
            rParms.Add("TPlanID", sm.UserInfo.TPlanID)
            rParms.Add("PlanYear", sm.UserInfo.Years)
            sql &= " AND a.TPlanID =@TPlanID" & vbCrLf
            sql &= " AND a.PlanYear =@PlanYear" & vbCrLf
        Else
            rParms.Add("PlanID", TIMS.VAL1(sm.UserInfo.PlanID))
            sql &= " AND a.PlanID =@PlanID" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, rParms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        msg.Text = ""
        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            divTip.Visible = (v_SearchMode = Cst_cRevise)
            DataGridTable28.Visible = True
        Else
            DataGridTable.Visible = True
        End If

        'PageControler1.SqlDataCreate(objstr, strDataSort)
        PageControler1.PageDataTable = dt
        PageControler1.Sort = strDataSort
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub But_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Search.Click
        '查詢模式
        Dim v_SearchMode As String = TIMS.GetListValue(SearchMode)
        Select Case v_SearchMode
            Case Cst_cPlanInfo, Cst_cRevise
            Case Else
                '請選擇查詢模式
                Dim s_msg1 As String = "請選擇查詢模式!"
                Common.MessageBox(Me, s_msg1)
                Return
        End Select

        Call sUtl_Search1()
    End Sub

    Sub UseKeepSession1()
        If Session(cst_sess_tc05001_search1) Is Nothing Then Return

        Dim MyValue As String = ""
        Dim KeepSessionStr1 As String = Convert.ToString(Session(cst_sess_tc05001_search1))
        Session(cst_sess_tc05001_search1) = Nothing

        MyValue = TIMS.GetMyValue(KeepSessionStr1, "prg")
        If MyValue <> "TC_05_001" Then Return

        TB_career_id.Text = TIMS.GetMyValue(KeepSessionStr1, "TB_career_id")
        trainValue.Value = TIMS.GetMyValue(KeepSessionStr1, "trainValue")
        jobValue.Value = TIMS.GetMyValue(KeepSessionStr1, "jobValue")
        txtCJOB_NAME.Text = TIMS.GetMyValue(KeepSessionStr1, "txtCJOB_NAME")
        cjobValue.Value = TIMS.GetMyValue(KeepSessionStr1, "cjobValue")
        center.Text = TIMS.GetMyValue(KeepSessionStr1, "center")
        RIDValue.Value = TIMS.GetMyValue(KeepSessionStr1, "RIDValue")
        ClassName.Text = TIMS.GetMyValue(KeepSessionStr1, "ClassName")
        MyValue = TIMS.GetMyValue(KeepSessionStr1, "SearchMode")
        If MyValue <> "" Then Common.SetListItem(SearchMode, MyValue)

        MyValue = TIMS.GetMyValue(KeepSessionStr1, "CheckMode")
        If MyValue <> "" Then Common.SetListItem(CheckMode, MyValue)
        CyclType.Text = TIMS.GetMyValue(KeepSessionStr1, "CyclType")
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        MyValue = TIMS.GetMyValue(KeepSessionStr1, "PageIndex")
        If IsNumeric(MyValue) Then PageControler1.PageIndex = TIMS.CINT1(MyValue)
        'But_Search_Click(sender, e)
        Call sUtl_Search1()

    End Sub

    Sub KeepSession1()
        Dim v_SearchMode As String = TIMS.GetListValue(SearchMode)
        Dim v_CheckMode As String = TIMS.GetListValue(CheckMode)
        'Session(cst_sess_tc05001_search1) = Nothing
        Dim s_search As String = String.Empty
        s_search = "prg=TC_05_001"
        s_search += "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        s_search += "&trainValue=" & TIMS.ClearSQM(trainValue.Value)
        s_search += "&jobValue=" & TIMS.ClearSQM(jobValue.Value)
        s_search += "&txtCJOB_NAME=" & TIMS.ClearSQM(txtCJOB_NAME.Text)
        s_search += "&cjobValue=" & TIMS.ClearSQM(cjobValue.Value)
        s_search += "&center=" & TIMS.ClearSQM(center.Text)
        s_search += "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        s_search += "&ClassName=" & TIMS.ClearSQM(ClassName.Text)
        s_search += "&SearchMode=" & v_SearchMode
        s_search += "&CheckMode=" & v_CheckMode
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        s_search += "&CyclType=" & TIMS.ClearSQM(CyclType.Text)
        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            s_search += "&PageIndex=" & PlanList28.CurrentPageIndex + 1
        Else
            s_search += "&PageIndex=" & PlanList.CurrentPageIndex + 1
        End If
        Session(cst_sess_tc05001_search1) = s_search
    End Sub

    Sub Utl_PlanList28_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        If e.CommandName = "" OrElse e.CommandArgument = "" Then Return
        Dim sCmdArg As String = e.CommandArgument
        'If isBlack.Value = "Y" Then '機構黑名單(訓練單位處分)
        '    Dim b_OBL As Boolean = Chk_OrgBlackList(vsMsg2)
        '    Common.MessageBox(Me, vsMsg2)
        '    Exit Sub
        'End If

        Call KeepSession1()
        Select Case e.CommandName '功能鈕
            Case cst_CommandName_BTN_OL_EDIT1 '(線上送件)編輯／查看
                'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                Dim url1 As String = String.Concat("TC_05_001_FL?ID=", TIMS.Get_MRqID(Me), sCmdArg)
                Call TIMS.Utl_Redirect(Me, objconn, url1)

            Case cst_CommandName_BTN_OL_SEND1 '(線上送件)送出
                'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                Dim rPMS As New Hashtable
                TIMS.SetMyValue2(rPMS, "ORGKINDGW", TIMS.GetMyValue(sCmdArg, "ORGKINDGW"))
                TIMS.SetMyValue2(rPMS, "ALTDATAID", TIMS.GetMyValue(sCmdArg, "AltDataID"))
                TIMS.SetMyValue2(rPMS, "PLANID", TIMS.GetMyValue(sCmdArg, "PlanID"))
                TIMS.SetMyValue2(rPMS, "COMIDNO", TIMS.GetMyValue(sCmdArg, "cid"))
                TIMS.SetMyValue2(rPMS, "SEQNO", TIMS.GetMyValue(sCmdArg, "no"))
                TIMS.SetMyValue2(rPMS, "CDATE", TIMS.GetMyValue(sCmdArg, "CDate"))
                TIMS.SetMyValue2(rPMS, "SUBSEQNO", TIMS.GetMyValue(sCmdArg, "subno"))

                Dim tmpMSG As String = ""
                Dim iProgress As Integer = TIMS.GET_iPROGRESS_PR(objconn, tmpMSG, rPMS)
                Dim EMSG As String = ""
                If iProgress < 100 Then
                    EMSG = $"線上申辦進度 未達100%，不可送出!,{iProgress}{vbCrLf}{If(tmpMSG <> "", $"請檢查：({tmpMSG})", "")}"
                    Common.MessageBox(Me, EMSG)
                    Return
                End If

                'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                Dim uParms As New Hashtable From {{"ONLINESENDACCT", sm.UserInfo.UserID}}
                'uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                'TIMS.SetMyValue2(uParms, "ORGKINDGW", TIMS.GetMyValue(sCmdArg, "ORGKINDGW"))
                TIMS.SetMyValue2(uParms, "PLANID", TIMS.GetMyValue(sCmdArg, "PlanID"))
                TIMS.SetMyValue2(uParms, "COMIDNO", TIMS.GetMyValue(sCmdArg, "cid"))
                TIMS.SetMyValue2(uParms, "SEQNO", TIMS.GetMyValue(sCmdArg, "no"))
                TIMS.SetMyValue2(uParms, "CDATE", TIMS.GetMyValue(sCmdArg, "CDate"))
                TIMS.SetMyValue2(uParms, "SUBSEQNO", TIMS.GetMyValue(sCmdArg, "subno"))
                TIMS.SetMyValue2(uParms, "ALTDATAID", TIMS.GetMyValue(sCmdArg, "AltDataID"))
                Dim usSql As String = ""
                usSql &= " UPDATE PLAN_REVISE" & vbCrLf
                usSql &= " SET ONLINESENDSTATUS='Y' ,ONLINESENDACCT=@ONLINESENDACCT ,ONLINESENDDATE=GETDATE()" & vbCrLf
                'usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                usSql &= " AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO AND ALTDATAID=@ALTDATAID" & vbCrLf
                Dim iRst As Integer = DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
                If iRst > 0 Then
                    Common.MessageBox(Me, "班級變更申請-線上送件-已送出!")
                    Call sUtl_Search1()
                End If

            Case cst_CommandName_BTN_OL_DEL1  '(線上送件)刪除

                Dim pms_sel As New Hashtable 'From {{"ONLINESENDACCT", sm.UserInfo.UserID}}
                TIMS.SetMyValue2(pms_sel, "PLANID", TIMS.GetMyValue(sCmdArg, "PlanID"))
                TIMS.SetMyValue2(pms_sel, "COMIDNO", TIMS.GetMyValue(sCmdArg, "cid"))
                TIMS.SetMyValue2(pms_sel, "SEQNO", TIMS.GetMyValue(sCmdArg, "no"))
                TIMS.SetMyValue2(pms_sel, "CDATE", TIMS.GetMyValue(sCmdArg, "CDate"))
                TIMS.SetMyValue2(pms_sel, "SUBSEQNO", TIMS.GetMyValue(sCmdArg, "subno"))
                TIMS.SetMyValue2(pms_sel, "ALTDATAID", TIMS.GetMyValue(sCmdArg, "AltDataID"))
                Dim Sql_sel As String = ""
                Sql_sel &= " SELECT FILENAME1,FILEPATH1 FROM PLAN_REVISESUBFL"
                Sql_sel &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                Sql_sel &= " AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO AND ALTDATAID=@ALTDATAID" & vbCrLf
                Sql_sel &= " AND FILENAME1 IS NOT NULL AND FILEPATH1 IS NOT NULL" & vbCrLf
                Dim dtFL As DataTable = DbAccess.GetDataTable(Sql_sel, objconn, pms_sel)
                If TIMS.dtHaveDATA(dtFL) Then
                    Dim oFILENAME1 As String = ""
                    Dim oUploadPath As String = ""
                    For Each drFL As DataRow In dtFL.Rows
                        Try
                            oFILENAME1 = Convert.ToString(drFL("FILENAME1"))
                            oUploadPath = Convert.ToString(drFL("FILEPATH1"))
                            If oFILENAME1 <> "" Then TIMS.MyFileDelete(Server.MapPath(oUploadPath & oFILENAME1))
                        Catch ex As Exception
                            TIMS.LOG.Warn(ex.Message, ex)
                            'Common.MessageBox(Me, ex.Message)
                            Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                            strErrmsg &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
                            strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
                            strErrmsg &= String.Concat("Server.MapPath(oUploadPath & oFILENAME1): ", Server.MapPath(oUploadPath & oFILENAME1), vbCrLf)
                            TIMS.WriteTraceLog(Me, ex, strErrmsg)
                        End Try
                    Next
                End If
                'DbAccess.ExecuteNonQuery(Sql_sel, objconn, pms_sel)

                Dim pms_del As New Hashtable 'From {{"ONLINESENDACCT", sm.UserInfo.UserID}}
                TIMS.SetMyValue2(pms_del, "PLANID", TIMS.GetMyValue(sCmdArg, "PlanID"))
                TIMS.SetMyValue2(pms_del, "COMIDNO", TIMS.GetMyValue(sCmdArg, "cid"))
                TIMS.SetMyValue2(pms_del, "SEQNO", TIMS.GetMyValue(sCmdArg, "no"))
                TIMS.SetMyValue2(pms_del, "CDATE", TIMS.GetMyValue(sCmdArg, "CDate"))
                TIMS.SetMyValue2(pms_del, "SUBSEQNO", TIMS.GetMyValue(sCmdArg, "subno"))
                TIMS.SetMyValue2(pms_del, "ALTDATAID", TIMS.GetMyValue(sCmdArg, "AltDataID"))
                Dim Sql_del As String = ""
                Sql_del &= " UPDATE PLAN_REVISE"
                Sql_del &= " SET ONLINESENDSTATUS=NULL ,ONLINESENDACCT=NULL ,ONLINESENDDATE=NULL"
                Sql_del &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                Sql_del &= " AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO AND ALTDATAID=@ALTDATAID" & vbCrLf
                DbAccess.ExecuteNonQuery(Sql_del, objconn, pms_del)

                Dim pms_del2 As New Hashtable 'From {{"ONLINESENDACCT", sm.UserInfo.UserID}}
                TIMS.SetMyValue2(pms_del2, "PLANID", TIMS.GetMyValue(sCmdArg, "PlanID"))
                TIMS.SetMyValue2(pms_del2, "COMIDNO", TIMS.GetMyValue(sCmdArg, "cid"))
                TIMS.SetMyValue2(pms_del2, "SEQNO", TIMS.GetMyValue(sCmdArg, "no"))
                TIMS.SetMyValue2(pms_del2, "CDATE", TIMS.GetMyValue(sCmdArg, "CDate"))
                TIMS.SetMyValue2(pms_del2, "SUBSEQNO", TIMS.GetMyValue(sCmdArg, "subno"))
                TIMS.SetMyValue2(pms_del2, "ALTDATAID", TIMS.GetMyValue(sCmdArg, "AltDataID"))
                Dim Sql_del2 As String = ""
                Sql_del2 &= " DELETE PLAN_REVISESUBFL FROM PLAN_REVISESUBFL"
                Sql_del2 &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                Sql_del2 &= " AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO AND ALTDATAID=@ALTDATAID" & vbCrLf
                DbAccess.ExecuteNonQuery(Sql_del2, objconn, pms_del2)

                Common.MessageBox(Me, "班級變更申請-刪除線上送件-已完成!")
                Call sUtl_Search1()
            Case cst_CommandName_But_Dir_appChg 'appChg"
                'Response.Redirect(e.CommandArgument)
                Dim url1 As String = e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Sub Utl_PlanList_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        If e.CommandName = "" Then Return
        'If Me.isBlack.Value = "Y" Then'機構黑名單(訓練單位處分)
        '    Dim b_OBL As Boolean = Chk_OrgBlackList(vsMsg2)
        '    Common.MessageBox(Me, vsMsg2)
        '    Exit Sub
        'End If

        Call KeepSession1()
        Select Case e.CommandName '功能鈕
            Case cst_CommandName_But_Dir_appChg 'appChg"
                'Response.Redirect(e.CommandArgument)
                Dim url1 As String = e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub PlanList_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles PlanList28.ItemCommand, PlanList.ItemCommand
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Utl_PlanList28_ItemCommand(source, e)
        Else
            Utl_PlanList_ItemCommand(source, e)
        End If
    End Sub

    Sub Utl_PlanList28_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim v_SearchMode As String = TIMS.GetListValue(SearchMode)
        Select Case e.Item.ItemType
            Case ListItemType.Header
                '查詢模式
                Select Case v_SearchMode
                    Case Cst_cPlanInfo 'UCase() '申請
                        PlanList28.Columns(Cst_pl28_訓練職類).Visible = True
                        PlanList28.Columns(Cst_pl28_變更項目).Visible = False
                        PlanList28.Columns(Cst_pl28_申請變更日).Visible = False
                        PlanList28.Columns(Cst_pl28_申請人姓名).Visible = False
                        PlanList28.Columns(Cst_pl28_審核時間).Visible = False
                        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                        PlanList28.Columns(Cst_pl28_線上送件).Visible = False
                        PlanList28.Columns(Cst_pl28_列印).Visible = False
                        e.Item.Cells(Cst_pl28_計畫狀態).Text = Cst_計畫狀態_Header_txt_計畫狀態 ' "計畫狀態"

                    Case Cst_cRevise 'UCase() '變更結果
                        PlanList28.Columns(Cst_pl28_訓練職類).Visible = False
                        PlanList28.Columns(Cst_pl28_變更項目).Visible = True
                        PlanList28.Columns(Cst_pl28_申請變更日).Visible = True
                        PlanList28.Columns(Cst_pl28_申請人姓名).Visible = True
                        PlanList28.Columns(Cst_pl28_審核時間).Visible = True
                        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                        PlanList28.Columns(Cst_pl28_線上送件).Visible = (hid_USE_PLAN_REVISESUB.Value = "Y") 'True
                        PlanList28.Columns(Cst_pl28_列印).Visible = True
                        e.Item.Cells(Cst_pl28_計畫狀態).Text = Cst_計畫狀態_Header_txt_申請狀態 ' "申請狀態"

                End Select

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labPrjstatus As Label = e.Item.FindControl("labPrjstatus") '計畫狀態
                Dim But_Dir As LinkButton = e.Item.FindControl("But_Dir") '功能鈕
                Dim labAltDataID As Label = e.Item.FindControl("labAltDataID") '變更項目
                'Dim labCDate As Label = e.Item.FindControl("labCDate") '申請變更日
                'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                Dim BTN_OL_EDIT1 As LinkButton = e.Item.FindControl("BTN_OL_EDIT1") '(線上送件)編輯／查看-"線上送件-編輯鈕"
                Dim BTN_OL_SEND1 As LinkButton = e.Item.FindControl("BTN_OL_SEND1") '(線上送件)送出-"線上送件-送出鈕"
                Dim BTN_OL_DEL1 As LinkButton = e.Item.FindControl("BTN_OL_DEL1") '(線上送件)刪除
                Dim Button3 As HtmlInputButton = e.Item.FindControl("Button3")   '(列印)計畫變更表
                Dim bt_print As HtmlInputButton = e.Item.FindControl("bt_print") '(列印)變更後課程表

                Select Case v_SearchMode
                    Case Cst_cPlanInfo 'UCase() '申請
                        'PlanList28.Columns(Cst_pl28_列印).Visible = False 'Button3.Visible = False 'bt_print.Visible = False

                        '計畫狀態
                        labPrjstatus.Text = Get_txtPrjstatus(drv, v_SearchMode)
                        '審核通過Y / 審核後修正O / 審核不通過N
                        Select Case Convert.ToString(drv("AppliedResult"))
                            Case "Y", "O"
                                Dim sUrl As String = ""
                                sUrl = "TC_05_001_chg.aspx?ID=" & rqMID
                                'sCmdArg = "TC_05_001_chg.aspx?ID=" & rqMID
                                Dim sCmdArg As String = ""
                                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                                TIMS.SetMyValue(sCmdArg, "cid", Convert.ToString(drv("ComIDNO")))
                                TIMS.SetMyValue(sCmdArg, "no", Convert.ToString(drv("SeqNO")))
                                TIMS.SetMyValue(sCmdArg, "check", v_SearchMode)

                                But_Dir.CommandArgument = sUrl & sCmdArg
                                '"TC_05_001_chg.aspx?ID=" & rqMID & "&PlanID=" & drv("PlanID") & "&cid=" & drv("ComIDNO") & "&no=" & drv("SeqNO") & "&check=" & Me.SearchMode.SelectedValue
                            Case "N" '審核不通過
                                But_Dir.Visible = False
                            Case Else
                                But_Dir.Visible = False
                        End Select
                        If Not BlnTest1 Then '正式環境為false
                            '結訓的話禁止申請
                            But_Dir.Enabled = True
                            If drv("IsClosed").ToString = "Y" Then
                                But_Dir.Enabled = False
                                But_Dir.ToolTip = "結訓班級不可以申請變更資料"
                                e.Item.ForeColor = Color.Red
                            End If
                        End If
                        'PlanList.Columns.Item(Cst_變更項目).Visible = False
                        'PlanList.Columns.Item(Cst_申請變更日).Visible = False
                        'e.Item.Cells(Cst_審核時間).Visible = False

                    Case Cst_cRevise 'UCase() '變更結果
                        Dim SCDate As String = If(Not IsDBNull(drv("CDate")), TIMS.Cdate3(drv("CDate")), "") 'yyyy/MM/dd

                        '-20081016 andy add 列印變更課程表(產學訓) -課程表-start
                        Dim sAltDataID As String = Convert.ToString(drv("AltDataID"))
                        bt_print.Disabled = (sAltDataID = "9" OrElse sAltDataID = "16" OrElse SCDate < CDate("2008-9-20"))
                        If (bt_print.Disabled) Then TIMS.Tooltip(bt_print, "停辦、其它")

                        If Not bt_print.Disabled Then
                            'SD_14_010_R1_c
                            Dim iPTDRID As Integer = SD_14_010.Get_PTDRID(drv("PlanID"), Convert.ToString(drv("ComIDNO")), drv("SeqNO"), SCDate, drv("SubSeqNO"), gsCmd)
                            prtFilename = If(Convert.ToString(drv("TMID")) = TIMS.cst_EHour_Use_TMID, SD_14_010.cst_printFN1d, SD_14_010.cst_printFN1c)
                            bt_print.Attributes("onclick") = String.Concat("openPrint('../../SQControl.aspx?filename=", prtFilename, "&PTDRID=", iPTDRID, "&AltDataID=", sAltDataID, "');")
                        End If
                        '-20081016 andy add 列印變更課程表(產學訓) -課程表-end
                        '列印變更申請表
                        prtFilename = SD_14_010.cst_printFN2 '"SD_14_010_b"
                        'Button3.Attributes("onclick") = "openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=" & prtFilename & "&path=" & SMpath & "&Years=" & (sm.UserInfo.Years - 1911) & "&PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNo") & "&CDate=" & SCDate & "&SubSeqNO=" & drv("SubSeqNO") & "&Title='+escape('" & KindValue.Value & "')+' &');"
                        'http://163.29.199.222:8080/ReportServer3/report.do?GUID=38506&RptID=SD_14_010_b&Years=107&PlanID=4519&ComIDNO=40760667&SeqNo=15&CDate=2019/01/08&SubSeqNO=1&Title=x&UserID=L7100071
                        'select dbo.FN_REV_PLAN_ONCLASS(4519,'40760667',15,1,convert(date,'2019/01/08'),'VWANDT')
                        'Dim REVIEW_JS As String = "openPrint('../../SQControl.aspx?filename=" & prtFilename & "&Years=" & (sm.UserInfo.Years - 1911) & "&PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNo") & "&CDate=" & SCDate & "&SubSeqNO=" & drv("SubSeqNO") & "&Title='+escape('" & KindValue.Value & "')+'&');"
                        Dim REVIEW_JS As String = "openPrint('../../SQControl.aspx?filename=" & prtFilename & "&Years=" & ROC_Years.Value & "&PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNo") & "&CDate=" & SCDate & "&SubSeqNO=" & drv("SubSeqNO") & "&');"
                        Button3.Attributes("onclick") = REVIEW_JS

                        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
                        Dim fg_BTN_OL_EDIT1_VIEW As Boolean = (Convert.ToString(drv("ReviseStatus")) <> "" OrElse Convert.ToString(drv("ONLINESENDSTATUS")) <> "")
                        BTN_OL_SEND1.Visible = Not fg_BTN_OL_EDIT1_VIEW  '(線上送件)送出'(已有審核結果)／線上送件 已送出
                        Dim OL_EDIT_TXT As String = If(fg_BTN_OL_EDIT1_VIEW, cst_BTN_OL_SUBVIEW_txt_查看, "")
                        Dim sEDIT_TIP As String = If(fg_BTN_OL_EDIT1_VIEW, "線上送件-查看鈕", "線上送件-編輯鈕")
                        Dim fg_BTN_OL_SEND1_NOSHOW As Boolean = If(fg_BTN_OL_EDIT1_VIEW, True, False)
                        '列印變更申請表
                        Dim s_ReviseStatus_N As String = ""
                        '**by Milor 20080502--只有審核中的資料才能按列印----start
                        Select Case Convert.ToString(drv("ReviseStatus"))
                            Case ""
                                s_ReviseStatus_N = cst_AppliedResult_txt_oth '"審核中"
                                Button3.Disabled = False
                            Case "Y"
                                s_ReviseStatus_N = cst_AppliedResult_txt_Y '"審核通過"
                                Button3.Disabled = True
                                Button3.Style.Add("background-color", "lightgray")  'edit，by:20181120
                                'If (Not bt_print.Disabled) Then bt_print.Disabled = True
                            Case "N"
                                s_ReviseStatus_N = cst_AppliedResult_txt_N '"審核失敗"
                                Button3.Disabled = True
                                Button3.Style.Add("background-color", "lightgray")  'edit，by:20181120
                                'If (Not bt_print.Disabled) Then bt_print.Disabled = True
                        End Select
                        '(若有[查看]文字)
                        If OL_EDIT_TXT <> "" Then BTN_OL_EDIT1.Text = OL_EDIT_TXT
                        Dim s_SUBVIEW As String = BTN_OL_EDIT1.Text '查看／編輯
                        If s_ReviseStatus_N <> "" Then TIMS.Tooltip(BTN_OL_EDIT1, $"{s_ReviseStatus_N}-{sEDIT_TIP}", True)

                        '(查看)且(已有審核狀態：通過／不通過)且 ONLINESENDSTATUS:NULL
                        If (OL_EDIT_TXT = cst_BTN_OL_SUBVIEW_txt_查看 AndAlso s_ReviseStatus_N <> cst_AppliedResult_txt_oth) AndAlso Convert.ToString(drv("ONLINESENDSTATUS")) = "" Then
                            BTN_OL_EDIT1.Visible = False '(不顯示線上送件查看按鈕)
                        End If

                        '(已有審核結果-不送)
                        BTN_OL_SEND1.Visible = If(fg_BTN_OL_SEND1_NOSHOW, False, True)
                        If s_ReviseStatus_N <> "" Then TIMS.Tooltip(BTN_OL_SEND1, $"{s_ReviseStatus_N}-線上送件-送出鈕", True)
                        If s_ReviseStatus_N <> "" Then TIMS.Tooltip(Button3, s_ReviseStatus_N, True)
                        If s_ReviseStatus_N <> "" Then TIMS.Tooltip(bt_print, s_ReviseStatus_N, True)
                        '**by Milor 20080502----end

                        '申請狀態
                        labPrjstatus.Text = Get_txtPrjstatus(drv, v_SearchMode)
                        If labPrjstatus.Text = cst_AppliedResult_txt_PARTREDUC_Y Then
                            'OJT-25071401：<系統> 產投_班級變更申請：調整還原時之列印按鈕卡控邏輯
                            '【申請狀態】為「待修正」，則「訓練計畫變更表」及「變更後課程表」按鈕反灰，不提供列印
                            Const cst_tit1 As String = "【申請狀態】待修正，不可列印"
                            If Not Button3.Disabled Then
                                Button3.Disabled = True '(列印)計畫變更表
                                TIMS.Tooltip(Button3, cst_tit1, True)
                            End If
                            If Not bt_print.Disabled Then
                                bt_print.Disabled = True '(列印)變更後課程表
                                TIMS.Tooltip(bt_print, cst_tit1, True)
                            End If
                        End If

                        '審核通過Y / 審核後修正O / 審核不通過N
                        But_Dir.Text = cst_But_Dir_txt_查詢 ' "查詢"

                        Dim s_PARTREDUC1 As String = ""
                        '計畫狀態/申請狀態 : AppliedResult 申請 plan_planinfo / ReviseStatus 變更結果 PLAN_REVISE
                        If Convert.ToString(drv("AppliedResult")) = "" AndAlso Convert.ToString(drv("PARTREDUC")) = TIMS.cst_YES Then
                            But_Dir.Text = cst_But_Dir_txt_修改 ' "修改"
                            TIMS.Tooltip(But_Dir, "狀態為未審還原，可修改部份內容")
                            s_PARTREDUC1 = "Y" '可修改部份內容
                        End If

                        Dim sUrl As String = "TC_05_001_chg.aspx?ID=" & rqMID
                        'sCmdArg = "TC_05_001_chg.aspx?ID=" & rqMID
                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                        TIMS.SetMyValue(sCmdArg, "CDate", TIMS.Cdate3(drv("CDate"))) '申請變更日
                        TIMS.SetMyValue(sCmdArg, "subno", Convert.ToString(drv("SubSeqNo")))
                        TIMS.SetMyValue(sCmdArg, "AltDataID", Convert.ToString(drv("AltDataID")))
                        TIMS.SetMyValue(sCmdArg, "PARTREDUC1", s_PARTREDUC1)

                        TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                        TIMS.SetMyValue(sCmdArg, "cid", Convert.ToString(drv("ComIDNO")))
                        TIMS.SetMyValue(sCmdArg, "no", Convert.ToString(drv("SeqNO")))
                        TIMS.SetMyValue(sCmdArg, "check", v_SearchMode)
                        TIMS.SetMyValue(sCmdArg, "SUBVIEW", s_SUBVIEW)

                        Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                        Dim fgLIDx0xROLEIDx1 As Boolean = TIMS.ChkUserLIDRole(sm, 0, 1)
                        BTN_OL_DEL1.Visible = If(flagS1 OrElse fgLIDx0xROLEIDx1, True, False)
                        'BTN_OL_DEL1.Style.Item("display") = "none"
                        BTN_OL_DEL1.CommandArgument = sCmdArg
                        BTN_OL_DEL1.Attributes("onclick") = "javascript:return confirm('此動作會刪除上傳資料，是否確定?');"
                        TIMS.Tooltip(BTN_OL_DEL1, "(刪除)線上送件-(刪除)上傳資料", True)

                        If $"{drv("ONLINESENDSTATUS")}" = "" AndAlso $"{drv("ONLINESENDACCT")}" = "" AndAlso $"{drv("ONLINESENDDATE")}" = "" Then
                            BTN_OL_DEL1.Style.Item("display") = "none" '(查無資料為空)
                        End If

                        But_Dir.CommandArgument = sUrl & sCmdArg
                        BTN_OL_EDIT1.CommandArgument = sCmdArg
                        BTN_OL_SEND1.CommandArgument = sCmdArg
                        BTN_OL_DEL1.CommandArgument = sCmdArg

                        'PlanList.Columns.Item(Cst_變更項目).Visible = True
                        'PlanList.Columns.Item(Cst_申請變更日).Visible = True '申請變更日
                        'e.Item.Cells(Cst_審核時間).Visible = True
                        '**by Milor 20080506--將變更項目改由年度、是否產學訓判斷，會顯示不同的變更項目名稱。 '變數值改由陣列獲取
                        Dim s_AltDataID_N As String = ""
                        'Dim s_CDate_N As String = TIMS.cdate3(drv("Cdate"))'申請變更日
                        Try
                            'e.Item.Cells(Cst_變更項目).Text = chgItem(CInt(drv("AltDataID")) - 1) & " (" & CStr(drv("SubSeqNo")) & ")"
                            s_AltDataID_N = ChgItemName(CInt(drv("AltDataID")) - 1) & " (" & CStr(drv("SubSeqNo")) & ")"
                        Catch ex As Exception
                            'e.Item.Cells(Cst_變更項目).Text = "<FONT color='red'>陣列資料有誤</FONT>"
                            s_AltDataID_N = "<FONT color='red'>陣列資料有誤</FONT>"
                            Dim s_err1 As String = String.Format("TC_05_001,陣列資料有誤,{0}", ex.Message)
                            TIMS.LOG.Error(s_err1, ex)
                        End Try
                        labAltDataID.Text = s_AltDataID_N
                        'labCDate.Text = s_CDate_N'申請變更日
                        '**by Milor 20080506----end
                End Select
                If isBlack.Value = "Y" Then
                    'But_Dir.Enabled = False'機構黑名單(訓練單位處分)
                    TIMS.Tooltip(But_Dir, TIMS.cst_gBlackMsg1)
                End If
        End Select
    End Sub

    Sub Utl_PlanList_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        'Dim dr As DataRowView = e.Item.DataItem
        'Dim myitem As DataGridItem
        Dim v_SearchMode As String = TIMS.GetListValue(SearchMode)
        Select Case e.Item.ItemType
            Case ListItemType.Header
                '查詢模式
                Select Case v_SearchMode
                    Case Cst_cPlanInfo 'UCase() '申請
                        PlanList.Columns(Cst_pl_變更項目).Visible = False
                        PlanList.Columns(Cst_pl_申請變更日).Visible = False
                        PlanList.Columns(Cst_pl_申請人姓名).Visible = False
                        PlanList.Columns(Cst_pl_審核時間).Visible = False
                        e.Item.Cells(Cst_pl_計畫狀態).Text = Cst_計畫狀態_Header_txt_計畫狀態 ' "計畫狀態"
                    Case Cst_cRevise 'UCase() '變更結果
                        PlanList.Columns(Cst_pl_變更項目).Visible = True
                        PlanList.Columns(Cst_pl_申請變更日).Visible = True
                        PlanList.Columns(Cst_pl_申請人姓名).Visible = True
                        PlanList.Columns(Cst_pl_審核時間).Visible = True
                        e.Item.Cells(Cst_pl_計畫狀態).Text = Cst_計畫狀態_Header_txt_申請狀態 ' "申請狀態"
                End Select

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labPrjstatus As Label = e.Item.FindControl("labPrjstatus") '計畫狀態
                Dim But_Dir As LinkButton = e.Item.FindControl("But_Dir") '功能鈕
                Dim labAltDataID As Label = e.Item.FindControl("labAltDataID") '變更項目
                'Dim labCDate As Label = e.Item.FindControl("labCDate") '申請變更日

                Select Case v_SearchMode
                    Case Cst_cPlanInfo 'UCase() '申請
                        '計畫狀態
                        labPrjstatus.Text = Get_txtPrjstatus(drv, v_SearchMode)
                        '審核通過Y / 審核後修正O / 審核不通過N
                        Select Case Convert.ToString(drv("AppliedResult"))
                            Case "Y", "O"
                                Dim sUrl As String = ""
                                sUrl = "TC_05_001_chg.aspx?ID=" & rqMID
                                'sCmdArg = "TC_05_001_chg.aspx?ID=" & rqMID
                                Dim sCmdArg As String = ""
                                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                                TIMS.SetMyValue(sCmdArg, "cid", Convert.ToString(drv("ComIDNO")))
                                TIMS.SetMyValue(sCmdArg, "no", Convert.ToString(drv("SeqNO")))
                                TIMS.SetMyValue(sCmdArg, "check", v_SearchMode)

                                But_Dir.CommandArgument = sUrl & sCmdArg
                                '"TC_05_001_chg.aspx?ID=" & rqMID & "&PlanID=" & drv("PlanID") & "&cid=" & drv("ComIDNO") & "&no=" & drv("SeqNO") & "&check=" & Me.SearchMode.SelectedValue
                            Case "N" '審核不通過
                                But_Dir.Visible = False
                            Case Else
                                But_Dir.Visible = False
                        End Select
                        If Not BlnTest1 Then '正式環境為false
                            '結訓的話禁止申請
                            But_Dir.Enabled = True
                            If drv("IsClosed").ToString = "Y" Then
                                But_Dir.Enabled = False
                                But_Dir.ToolTip = "結訓班級不可以申請變更資料"
                                e.Item.ForeColor = Color.Red
                            End If
                        End If
                        'PlanList.Columns.Item(Cst_變更項目).Visible = False
                        'PlanList.Columns.Item(Cst_申請變更日).Visible = False
                        'e.Item.Cells(Cst_審核時間).Visible = False

                    Case Cst_cRevise 'UCase() '變更結果
                        '申請狀態
                        labPrjstatus.Text = Get_txtPrjstatus(drv, v_SearchMode)

                        '審核通過Y / 審核後修正O / 審核不通過N
                        But_Dir.Text = cst_But_Dir_txt_查詢 ' "查詢"

                        Dim s_PARTREDUC1 As String = "" '可修改部份內容-狀態為未審還原，可修改部份內容
                        '計畫狀態/申請狀態 : AppliedResult 申請 plan_planinfo / ReviseStatus 變更結果 PLAN_REVISE
                        If Convert.ToString(drv("AppliedResult")) = "" AndAlso Convert.ToString(drv("PARTREDUC")) = TIMS.cst_YES Then
                            But_Dir.Text = cst_But_Dir_txt_修改 ' "修改"
                            TIMS.Tooltip(But_Dir, "狀態為未審還原，可修改部份內容")
                            s_PARTREDUC1 = "Y" '可修改部份內容
                        End If

                        Dim sUrl As String = "TC_05_001_chg.aspx?ID=" & rqMID 'sCmdArg = "TC_05_001_chg.aspx?ID=" & rqMID
                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "CDate", TIMS.Cdate3(drv("CDate"))) '申請變更日
                        TIMS.SetMyValue(sCmdArg, "subno", Convert.ToString(drv("SubSeqNo")))
                        TIMS.SetMyValue(sCmdArg, "AltDataID", Convert.ToString(drv("AltDataID")))
                        TIMS.SetMyValue(sCmdArg, "PARTREDUC1", s_PARTREDUC1)

                        TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                        TIMS.SetMyValue(sCmdArg, "cid", Convert.ToString(drv("ComIDNO")))
                        TIMS.SetMyValue(sCmdArg, "no", Convert.ToString(drv("SeqNO")))
                        TIMS.SetMyValue(sCmdArg, "check", v_SearchMode)

                        But_Dir.CommandArgument = sUrl & sCmdArg

                        'PlanList.Columns.Item(Cst_變更項目).Visible = True
                        'PlanList.Columns.Item(Cst_申請變更日).Visible = True '申請變更日
                        'e.Item.Cells(Cst_審核時間).Visible = True
                        '**by Milor 20080506--將變更項目改由年度、是否產學訓判斷，會顯示不同的變更項目名稱。 '變數值改由陣列獲取
                        Dim s_AltDataID_N As String = ""
                        'Dim s_CDate_N As String = TIMS.cdate3(drv("Cdate"))'申請變更日
                        Try
                            'e.Item.Cells(Cst_變更項目).Text = chgItem(CInt(drv("AltDataID")) - 1) & " (" & CStr(drv("SubSeqNo")) & ")"
                            s_AltDataID_N = ChgItemName(CInt(drv("AltDataID")) - 1) & " (" & CStr(drv("SubSeqNo")) & ")"
                        Catch ex As Exception
                            'e.Item.Cells(Cst_變更項目).Text = "<FONT color='red'>陣列資料有誤</FONT>"
                            s_AltDataID_N = "<FONT color='red'>陣列資料有誤</FONT>"
                            Dim s_err1 As String = String.Format("TC_05_001,陣列資料有誤,{0}", ex.Message)
                            TIMS.LOG.Error(s_err1, ex)
                        End Try
                        labAltDataID.Text = s_AltDataID_N
                        'labCDate.Text = s_CDate_N'申請變更日
                        '**by Milor 20080506----end
                End Select
                If isBlack.Value = "Y" Then
                    'But_Dir.Enabled = False'機構黑名單(訓練單位處分)
                    TIMS.Tooltip(But_Dir, TIMS.cst_gBlackMsg1)
                End If
        End Select
    End Sub

    Private Sub list_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles PlanList28.ItemDataBound, PlanList.ItemDataBound
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Utl_PlanList28_ItemDataBound(sender, e)
        Else
            Utl_PlanList_ItemDataBound(sender, e)
        End If
    End Sub

End Class
