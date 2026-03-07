Partial Class TC_02_002
    Inherits AuthBasePage

    Const cst_printFN1 As String = "TC_02_002A"
    'dtPlan 'Dim CPdt As DataTable 'SearchData1
    Const Cst_index As Integer = 0 '序號
    'Const Cst_PlanYear As Integer = 1 'PlanYear '計畫年度
    'Const Cst_AppliedDate As Integer = 2 'AppliedDate '申請日期
    'Const Cst_STDate As Integer = 3 'STDate '訓練起日
    'Const Cst_FDDate As Integer = 4 'FDDate '訓練迄日
    'Const Cst_OrgName2 As Integer = 5 'OrgName2 '管控單位
    'Const Cst_OrgName As Integer = 6 'OrgName '機構名稱
    'Const Cst_ClassName As Integer = 7 'ClassName '班名
    'Const Cst_AppliedResult As Integer = 8 '審核狀態
    'Const Cst_FUNCTION1 As Integer = 9 '功能 'lbtSFEDIT1 'lbtSFDEL1

    'Const cst_Sort As String = "Sort"
    '' e.CommandName
    Const cst_lbtSFEDIT1_Txt_申復 As String = "申復"
    Const cst_lbtSFEDIT1_Txt_修改 As String = "修改"
    Const cst_lbtSFEDIT1_Txt_查看 As String = "查看"

    Const cst_stopmsg_12 As String = "申請申複階段受理期間未開放，請確認後再操作!"
    Const cst_search_tc02002 As String = "_search_tc02002"
    'Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        'center.Attributes.Add("onfocus", "this.blur();")
        TIMS.INPUT_ReadOnly2(center)

        TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = dtPlan '分頁設定

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '帶入查詢參數
        If Not IsPostBack Then
            Call CCreate1()
        End If

        '因有傳入值 yearlist.SelectedValue.ToString 故放此位置，才可讀到值
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), sm.UserInfo.Years)

        '確認機構是否為黑名單
        Dim StrMsg2 As String = ""
        If Chk_OrgBlackList(StrMsg2) Then
            Dim strScript As String = String.Concat("<script>alert('", StrMsg2, "');</script>")
            Page.RegisterStartupScript("", strScript)
        End If
    End Sub
    Sub CCreate1()
        Call SHOW_PANEL(0)
        '(加強操作便利性)
        RIDValue.Value = sm.UserInfo.RID
        center.Text = sm.UserInfo.OrgName

        DataGridTable.Visible = False
        msg1.Text = ""

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Dim v_APPSTAGE As String = TIMS.GET_CANUSE_APPSTAGE(objconn, CStr(sm.UserInfo.Years), TIMS.cst_APPSTAGE_PTYPE1_02)
        Dim v_APPSTAGE_SCH_DEF As String = "1"
        Common.SetListItem(ddlAPPSTAGE_SCH, If(v_APPSTAGE <> "", v_APPSTAGE, v_APPSTAGE_SCH_DEF))

        Call UseKeepSearchStr()
    End Sub

    ''' <summary>顯示目前執行區塊 0:list 1:one data</summary>
    ''' <param name="iType"></param>
    Sub SHOW_PANEL(iType As Integer)
        '0: show panelSch / '1: show panelEdit1 
        panelSch.Visible = If(iType = 0, True, False)
        panelEdit1.Visible = If(iType = 1, True, False)
    End Sub


    ''' <summary>'已列入處分名單 提醒功能</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = False
        Errmsg = ""
        Dim StrComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, StrComIDNO, objconn) Then
            Rst = True
            isBlack.Value = "Y"
            orgname.Value = sm.UserInfo.OrgName
            Errmsg = String.Concat(sm.UserInfo.OrgName, "，已列入處分名單!!")
        End If
        Return Rst
    End Function

    ''' <summary>審核文字說明</summary>
    ''' <param name="s_TPLANID"></param>
    ''' <param name="s_AppliedResult"></param>
    ''' <param name="s_RESULTBUTTON"></param>
    ''' <returns></returns>
    Function Get_AppliedResultTxt(ByRef s_TPLANID As String, ByRef s_AppliedResult As String, ByRef s_RESULTBUTTON As String) As String
        Dim rst As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(s_TPLANID) > -1 Then
            '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
            'Dim strMsg As String = "班級審核中" '= ""

            rst = "班級審核中" '= ""
            Dim flag_AppliedResult_Red_1 As Boolean = False '(紅字加強)
            Select Case s_AppliedResult
                Case "Y"
                    rst = "班級審核通過" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"班級審核通過"
                Case "N"
                    flag_AppliedResult_Red_1 = True
                    rst = "班級審核不通過" 'e.Item.Cells(Cst_AppliedResult).Text += "<font color=red>" & strMsg & "</font>"
                Case "R"
                    rst = "班級退件修正" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg ' "班級退件修正"
                Case "M"
                    rst = "請修正資料" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"請修正資料"
                Case "O"
                    '產投為審核中狀態。
                    rst = "班級審核中(審核後修正)" 'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"班級審核中(審核後修正)" '"審核後修正"
                Case Else
                    's_RESULTBUTTON Y/R 
                    Select Case s_RESULTBUTTON'Convert.ToString(drv("RESULTBUTTON"))
                        Case TIMS.cst_ResultButton_尚未送出_待送審 'Y
                            rst = "待送審"
                        Case TIMS.cst_ResultButton_尚未送出_未送出 'R
                            flag_AppliedResult_Red_1 = True
                            rst = "(未正式儲存)"
                            'Case Else 'strMsg = "班級審核中"
                    End Select
                    'e.Item.Cells(Cst_AppliedResult).Text += strMsg '"班級審核中"
            End Select
            'e.Item.Cells(Cst_AppliedResult).Text = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(drv("AppliedResult")))
            If flag_AppliedResult_Red_1 Then rst = String.Concat("<font color=red>", rst, "</font>")
            'e.Item.Cells(Cst_AppliedResult).Text = strMsg
        Else
            '非產投
            If s_AppliedResult = "" Then
                rst = "審核中"
                Return rst
            End If
            Select Case s_AppliedResult 'drv("AppliedResult")
                Case "Y"
                    rst = "審核通過"
                Case "N"
                    rst = "審核不通過"
                Case "R"
                    rst = "退件修正"
                Case "M"
                    rst = "請修正資料"
                Case "O"
                    rst = "審核後修正"
            End Select
        End If
        Return rst
    End Function

    ''' <summary>清理單筆資料</summary>
    Sub CLEAR_DATA2()
        Hid_PSOID.Value = "" 'Convert.ToString(dr("PSOID"))
        Hid_PSNO28.Value = "" 'Convert.ToString(dr("PSNO28"))

        lbYEARS_ROC.Text = "" 'TIMS.GET_YEARS_ROC(dr("YEARS"))
        lbAPPSTAGE_N.Text = ""
        lbDistName.Text = "" 'Convert.ToString(dr("DISTNAME"))
        lbOrgName.Text = "" ' Convert.ToString(dr("ORGNAME"))
        lbPSNO28.Text = "" ' Convert.ToString(dr("PSNO28"))
        lbClassName.Text = "" 'Convert.ToString(dr("CLASSCNAME"))
        lbSFTDate.Text = "" ' String.Format("{0}~{1}", dr("STDATE"), dr("FTDATE"))

        '分署確認課程分類 / 職類課程 / 訓練業別
        lbGCODEPNAME.Text = "" ' Convert.ToString(dr("GCODEPNAME"))
        'lbGCNAME.Text = ""'Convert.ToString(dr("GCNAME")) '訓練業別
        lbCCNAME.Text = "" 'Convert.ToString(dr("CCNAME")) '訓練職能
        lbTNum.Text = "" 'Convert.ToString(dr("TNUM"))
        lbTHours.Text = "" ' Convert.ToString(dr("THOURS"))

        '核班結果,未核班原因/審查意見
        'labNGREASON.Enabled = False
        'labNGREASON.Text = "" 'Convert.ToString(dr("NGREASON"))
        NGREASON.Text = "" 'Convert.ToString(dr("NGREASON"))
        'If labNGREASON.Text = "" Then labNGREASON.Text = TIMS.cst_EmptyDataValue

        SFCONTNAME.Text = "" 'Convert.ToString(dr("SFCONTNAME"))
        SFCONTTEL.Text = "" 'Convert.ToString(dr("SFCONTTEL"))
        SFCONTTITLE.Text = "" 'Convert.ToString(dr("SFCONTTITLE"))
        SFCONTEMAIL.Text = "" 'Convert.ToString(dr("SFCONTEMAIL"))
        '申復理由及說明
        SFCONTREASONS.Text = "" 'Convert.ToString(dr("SFCONTREASONS"))
    End Sub

    ''' <summary>顯示單筆資料</summary>
    ''' <param name="dr"></param>
    Sub SHOW_DATA2(ByRef dr As DataRow)
        If dr Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Call SHOW_PANEL(1)
        Call CLEAR_DATA2()
        Dim fg_SFCID_Y As Boolean = (Convert.ToString(dr("SFCID_Y")) <> "")
        btnSAVE1.Enabled = (Not fg_SFCID_Y)
        TIMS.Tooltip(btnSAVE1, If(Not btnSAVE1.Enabled AndAlso fg_SFCID_Y, "申復線上送件-使用中", ""), True)

        Hid_PSOID.Value = Convert.ToString(dr("PSOID"))
        Hid_PSNO28.Value = Convert.ToString(dr("PSNO28"))

        lbYEARS_ROC.Text = TIMS.GET_YEARS_ROC(dr("YEARS"))
        lbAPPSTAGE_N.Text = Convert.ToString(dr("APPSTAGE_N"))
        lbDistName.Text = Convert.ToString(dr("DISTNAME"))
        lbOrgName.Text = Convert.ToString(dr("ORGNAME"))
        lbPSNO28.Text = Convert.ToString(dr("PSNO28"))
        lbClassName.Text = Convert.ToString(dr("CLASSCNAME"))
        lbSFTDate.Text = String.Format("{0}~{1}", dr("STDATE"), dr("FTDATE"))

        '分署確認課程分類 / 職類課程 / 訓練業別
        lbGCODEPNAME.Text = Convert.ToString(dr("GCODEPNAME"))
        'lbGCNAME.Text = Convert.ToString(dr("GCNAME")) '訓練業別
        lbCCNAME.Text = Convert.ToString(dr("CCNAME")) '訓練職能
        lbTNum.Text = Convert.ToString(dr("TNUM"))
        lbTHours.Text = Convert.ToString(dr("THOURS"))

        '核班結果,未核班原因/審查意見
        'labNGREASON.Enabled = False
        'labNGREASON.Text = Convert.ToString(dr("NGREASON"))
        'If labNGREASON.Text = "" Then labNGREASON.Text = TIMS.cst_EmptyDataValue
        NGREASON.Text = Convert.ToString(dr("NGREASON"))
        If NGREASON.Text = "" Then NGREASON.Text = TIMS.cst_EmptyDataValue
        NGREASON.ReadOnly = True
        NGREASON.ApplyStyle(TIMS.GET_RO_STYLE())

        SFCONTNAME.Text = Convert.ToString(dr("SFCONTNAME"))
        SFCONTTEL.Text = Convert.ToString(dr("SFCONTTEL"))
        SFCONTEMAIL.Text = Convert.ToString(dr("SFCONTEMAIL"))
        SFCONTTITLE.Text = Convert.ToString(dr("SFCONTTITLE"))
        'TIMS.PL_settextbox1(SFCONTNAME, dr("SFCONTNAME"))
        'TIMS.PL_settextbox1(SFCONTTEL, dr("SFCONTTEL"))
        'TIMS.PL_settextbox1(SFCONTEMAIL, dr("SFCONTEMAIL"))

        'Dim drO1 As DataRow = Nothing
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        'If drRR IsNot Nothing Then
        '    Dim htPMS As New Hashtable
        '    htPMS.Add("RID", Convert.ToString(drRR("RID")))
        '    htPMS.Add("PLANID", Convert.ToString(drRR("PLANID")))
        '    htPMS.Add("COMIDNO", Convert.ToString(drRR("COMIDNO")))
        '    drO1 = GET_ORG_ORGPLANINFO_row(htPMS)
        'End If
        'If drRR IsNot Nothing AndAlso drO1 IsNot Nothing Then
        '    'p.RSID,op.CONTACTNAME,op.CONTACTCELLPHONE,op.CONTACTEMAIL,op.PHONE
        '    If SFCONTNAME.Text = "" AndAlso Convert.ToString(drO1("CONTACTNAME")) <> "" Then TIMS.PL_placeholder(SFCONTNAME, Convert.ToString(drO1("CONTACTNAME")))
        '    If SFCONTTEL.Text = "" AndAlso Convert.ToString(drO1("PHONE")) <> "" Then TIMS.PL_placeholder(SFCONTTEL, Convert.ToString(drO1("PHONE")))
        '    If SFCONTEMAIL.Text = "" AndAlso Convert.ToString(drO1("CONTACTEMAIL")) <> "" Then TIMS.PL_placeholder(SFCONTEMAIL, Convert.ToString(drO1("CONTACTEMAIL")))
        'End If

        '申復理由及說明
        SFCONTREASONS.Text = Convert.ToString(dr("SFCONTREASONS"))

        SFCONTNAME.ReadOnly = (Not btnSAVE1.Enabled)
        SFCONTTEL.ReadOnly = (Not btnSAVE1.Enabled)
        SFCONTEMAIL.ReadOnly = (Not btnSAVE1.Enabled)
        SFCONTTITLE.ReadOnly = (Not btnSAVE1.Enabled)
        SFCONTREASONS.ReadOnly = (Not btnSAVE1.Enabled)
        If Not btnSAVE1.Enabled Then
            SFCONTNAME.ApplyStyle(TIMS.GET_RO_STYLE())
            SFCONTTEL.ApplyStyle(TIMS.GET_RO_STYLE())
            SFCONTEMAIL.ApplyStyle(TIMS.GET_RO_STYLE())
            SFCONTTITLE.ApplyStyle(TIMS.GET_RO_STYLE())
            SFCONTREASONS.ApplyStyle(TIMS.GET_RO_STYLE())
        Else
            SFCONTNAME.BackColor = Color.White
            SFCONTTEL.BackColor = Color.White
            SFCONTEMAIL.BackColor = Color.White
            SFCONTTITLE.BackColor = Color.White
            SFCONTREASONS.BackColor = Color.White
        End If
        TIMS.Tooltip(SFCONTNAME, If(Not btnSAVE1.Enabled AndAlso fg_SFCID_Y, "申復線上送件-使用中", ""), True)
        TIMS.Tooltip(SFCONTTEL, If(Not btnSAVE1.Enabled AndAlso fg_SFCID_Y, "申復線上送件-使用中", ""), True)
        TIMS.Tooltip(SFCONTEMAIL, If(Not btnSAVE1.Enabled AndAlso fg_SFCID_Y, "申復線上送件-使用中", ""), True)
        TIMS.Tooltip(SFCONTTITLE, If(Not btnSAVE1.Enabled AndAlso fg_SFCID_Y, "申復線上送件-使用中", ""), True)
        TIMS.Tooltip(SFCONTREASONS, If(Not btnSAVE1.Enabled AndAlso fg_SFCID_Y, "申復線上送件-使用中", ""), True)
    End Sub

    ''' <summary>查詢資料list TB</summary>
    Sub SSearchDATA1()
        SHOW_PANEL(0)
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dtPlan)
        TPlanName.Text = TIMS.GetTPlanName(Convert.ToString(sm.UserInfo.TPlanID), objconn)

        DataGridTable.Visible = False
        msg1.Text = ""

        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return
        End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If drR Is Nothing Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return
        End If
        '申請階段管理-受理期間設定 APPLISTAGE
        Dim aParms As New Hashtable From {{"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}}
        '開放受理之申請階段／PLAN_APPSTAGE
        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE02(objconn, aParms)
        '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
        If Not fg_can_applistage Then
            Common.MessageBox(Me, cst_stopmsg_12)
            Return
        End If

        Dim sParms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}}
        Dim sSql As String = ""
        'sSql &= " SELECT TOP 999 pf.PSOID,pf.PSNO28" & vbCrLf
        sSql &= " SELECT pf.PSOID,pf.PSNO28" & vbCrLf
        sSql &= " ,pp.YEARS,pp.APPSTAGE,pp.CLASSCNAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,CONCAT(dbo.FN_GET_ROC_YEAR(pp.YEARS),dbo.FN_GET_APPSTAGE2(pp.APPSTAGE)) YEARSROCAG" & vbCrLf
        sSql &= " ,format(pp.APPLIEDDATE,'yyyy/MM/dd') APPLIEDDATE" & vbCrLf
        sSql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sSql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sSql &= " ,pp.DISTNAME" & vbCrLf
        sSql &= " ,pp.ORGNAME,pp.TNUM,pp.THOURS" & vbCrLf
        sSql &= " ,pp.ISAPPRPAPER,pp.APPLIEDRESULT,pp.RESULTBUTTON,pp.PVR_ISAPPRPAPER,pp.DATANOTSENT" & vbCrLf
        'sSql &= " /* '分署確認課程分類 / 職類課程 / 訓練業別 */" & vbCrLf
        sSql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sSql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sSql &= " ,kc.CCNAME" & vbCrLf
        sSql &= " ,pf.NGREASON,pf.SFCONTREASONS" & vbCrLf
        sSql &= " ,CASE WHEN pf.SFCONTREASONS IS NOT NULL THEN 'Y' END SFCONT_CANDEL" & vbCrLf
        sSql &= " ,CASE WHEN pf.SFCONTREASONS IS NOT NULL THEN 'Y' END SFCONT_CANEDIT" & vbCrLf
        sSql &= " ,CASE WHEN pf.SFCONTREASONS IS NOT NULL THEN 'Y' END SFCONT_CANPRINT" & vbCrLf
        sSql &= " ,pf.SFCONTNAME,pf.SFCONTTITLE,pf.SFCONTTEL,pf.SFCONTEMAIL" & vbCrLf
        sSql &= " ,pf.SFCONTACCT,pf.SFCONTDATE" & vbCrLf
        '申復線上送件使用中，請先刪除申復線上送件!
        sSql &= " ,(SELECT MIN(a.SFCID) MIN_SFCID FROM ORG_SFCASEPI a WHERE a.PLANID=pp.PLANID AND a.COMIDNO=pp.COMIDNO AND a.SEQNO=pp.SEQNO) SFCID_Y" & vbCrLf
        '申復線上送件使用中，請先刪除申復線上送件!
        sSql &= " ,(SELECT MIN(a.SFCID) MIN_SFCID FROM ORG_SFCASE a WHERE a.PLANID=pp.PLANID AND a.RID=pp.RID AND a.APPSTAGE=pp.APPSTAGE) SFCID_RID" & vbCrLf
        sSql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sSql &= " JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        sSql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sSql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf

        sSql &= " WHERE pp.ISAPPRPAPER='Y' AND pp.PVR_ISAPPRPAPER='Y' AND pp.RESULTBUTTON IS NULL" & vbCrLf
        sSql &= " AND pp.DATANOTSENT IS NULL AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND pp.APPSTAGE=@APPSTAGE" & vbCrLf

        Select Case sm.UserInfo.LID
            Case 0
                If RIDValue.Value.Length > 1 Then
                    sParms.Add("RID", RIDValue.Value)
                    sSql &= " AND pp.RID=@RID" & vbCrLf
                ElseIf RIDValue.Value.Length = 1 Then
                    sParms.Add("DistID", drR("DistID"))
                    sSql &= " AND pp.DistID=@DistID" & vbCrLf
                Else
                    '(無資訊)
                    sSql &= " AND 1!=1" & vbCrLf
                End If
            Case 1
                sParms.Add("PlanID", sm.UserInfo.PlanID)
                sParms.Add("DistID", sm.UserInfo.DistID)
                sSql &= " AND pp.PlanID=@PlanID" & vbCrLf
                sSql &= " AND pp.DistID=@DistID" & vbCrLf
                If RIDValue.Value.Length > 1 Then
                    sParms.Add("RID", RIDValue.Value)
                    sSql &= " AND pp.RID=@RID" & vbCrLf
                End If
            Case Else
                sParms.Add("PlanID", sm.UserInfo.PlanID)
                sParms.Add("DistID", sm.UserInfo.DistID)
                sParms.Add("RID", sm.UserInfo.RID)
                sSql &= " AND pp.PlanID=@PlanID" & vbCrLf
                sSql &= " AND pp.DistID=@DistID" & vbCrLf
                sSql &= " AND pp.RID=@RID" & vbCrLf
        End Select

        '班名查詢
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then
            sParms.Add("ClassName", ClassName.Text)
            sSql &= " and pp.ClassName LIKE '%'+@ClassName +'%'" & vbCrLf
        End If
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then
            sParms.Add("CyclType", CyclType.Text)
            sSql &= " AND pp.CyclType=@CyclType" & vbCrLf
        End If
        'rbl_TransFlagS '增加【轉班上架】欄位，選項：不區分、未轉班、已轉班
        Dim v_rbl_TransFlagS As String = TIMS.GetListValue(rbl_TransFlagS)
        v_rbl_TransFlagS = If(v_rbl_TransFlagS = "A", "", v_rbl_TransFlagS)
        If v_rbl_TransFlagS <> "" Then
            sParms.Add("TransFlag", v_rbl_TransFlagS)
            sSql &= " AND pp.TransFlag=@TransFlag" & vbCrLf '轉班上架
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)

        DataGridTable.Visible = False
        msg1.Text = "查無資料"
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg1.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>查詢-單筆資料</summary>
    ''' <param name="htPMS"></param>
    ''' <returns></returns>
    Function SSearchDATA2(htPMS As Hashtable) As DataRow
        Hid_PSOID.Value = TIMS.GetMyValue2(htPMS, "PSOID") 'TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.GetMyValue2(htPMS, "PSNO28") ' TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSOID.Value = "" OrElse Hid_PSNO28.Value = "" Then Return Nothing

        Dim sParms As New Hashtable From {{"PSOID", Val(Hid_PSOID.Value)}, {"PSNO28", Hid_PSNO28.Value}}
        Dim sSql As String = ""
        'sSql &= " SELECT TOP 999 pf.PSOID,pf.PSNO28" & vbCrLf
        sSql &= " SELECT pf.PSOID,pf.PSNO28" & vbCrLf
        sSql &= " ,pp.YEARS,pp.APPSTAGE,pp.CLASSCNAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,CONCAT(dbo.FN_GET_ROC_YEAR(pp.YEARS),dbo.FN_GET_APPSTAGE2(pp.APPSTAGE)) YEARSROCAG" & vbCrLf
        sSql &= " ,format(pp.APPLIEDDATE,'yyyy/MM/dd') APPLIEDDATE" & vbCrLf
        sSql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sSql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sSql &= " ,pp.DISTID,pp.DISTNAME" & vbCrLf
        sSql &= " ,pp.ORGNAME,pp.TNUM,pp.THOURS" & vbCrLf
        sSql &= " ,pp.APPLIEDRESULT,pp.RESULTBUTTON,pp.PVR_ISAPPRPAPER,pp.DATANOTSENT" & vbCrLf
        sSql &= " ,pp.RID,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        'sSql &= " /* '分署確認課程分類 / 職類課程 / 訓練業別 */" & vbCrLf
        sSql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sSql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sSql &= " ,kc.CCNAME" & vbCrLf
        sSql &= " ,pf.NGREASON,pf.SFCONTREASONS" & vbCrLf
        sSql &= " ,pf.SFCONTNAME,pf.SFCONTTITLE,pf.SFCONTTEL,pf.SFCONTEMAIL" & vbCrLf
        sSql &= " ,pf.SFCONTACCT,pf.SFCONTDATE" & vbCrLf
        '申復線上送件使用中，請先刪除申復線上送件!
        sSql &= " ,(SELECT MIN(a.SFCID) MIN_SFCID FROM ORG_SFCASEPI a WHERE a.PLANID=pp.PLANID AND a.COMIDNO=pp.COMIDNO AND a.SEQNO=pp.SEQNO) SFCID_Y" & vbCrLf
        '申復線上送件使用中，請先刪除申復線上送件!
        sSql &= " ,(SELECT MIN(a.SFCID) MIN_SFCID FROM ORG_SFCASE a WHERE a.PLANID=pp.PLANID AND a.RID=pp.RID AND a.APPSTAGE=pp.APPSTAGE) SFCID_RID" & vbCrLf
        sSql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sSql &= " JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        sSql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sSql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf

        sSql &= " WHERE pf.PSOID=@PSOID AND pf.PSNO28=@PSNO28" & vbCrLf
        Dim drA2 As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms)
        Return drA2
    End Function

    ''' <summary>查詢值保留</summary>
    Sub KeepSearchStr()
        Dim str_search As String = ""
        TIMS.SetMyValue(str_search, "prg", "TC_02_002")
        TIMS.SetMyValue(str_search, "center", TIMS.ClearSQM(center.Text))
        TIMS.SetMyValue(str_search, "RIDValue", TIMS.ClearSQM(RIDValue.Value))
        TIMS.SetMyValue(str_search, "ClassName", TIMS.ClearSQM(ClassName.Text))
        TIMS.SetMyValue(str_search, "CyclType", TIMS.ClearSQM(CyclType.Text))
        TIMS.SetMyValue(str_search, "APPSTAGE_SCH", TIMS.GetListValue(ddlAPPSTAGE_SCH))
        TIMS.SetMyValue(str_search, "TransFlagS", TIMS.GetListValue(rbl_TransFlagS))
        TIMS.SetMyValue(str_search, "PageIndex", (dtPlan.CurrentPageIndex + 1))
        Session(cst_search_tc02002) = str_search
    End Sub

    ''' <summary>使用查詢值</summary>
    Sub UseKeepSearchStr()
        If Session(cst_search_tc02002) Is Nothing Then Return

        Dim str_search As String = Convert.ToString(Session(cst_search_tc02002))
        Session(cst_search_tc02002) = Nothing

        Dim MyValue As String = TIMS.GetMyValue(str_search, "prg")
        If MyValue <> "TC_02_002" Then Return

        center.Text = TIMS.GetMyValue(str_search, "center")
        RIDValue.Value = TIMS.GetMyValue(str_search, "RIDValue")
        ClassName.Text = TIMS.GetMyValue(str_search, "ClassName")
        Common.SetListItem(ddlAPPSTAGE_SCH, TIMS.GetMyValue(str_search, "APPSTAGE_SCH"))
        Common.SetListItem(rbl_TransFlagS, TIMS.GetMyValue(str_search, "TransFlagS"))

        Call SSearchDATA1()
        MyValue = TIMS.GetMyValue(str_search, "PageIndex")
        If MyValue <> "" AndAlso IsNumeric(MyValue) Then PageControler1.PageIndex = MyValue
    End Sub

    ''' <summary>檢核正常為true 異常 false</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" AndAlso Not TIMS.IsNumberStr(CyclType.Text) Then Errmsg &= "期別需輸入數字型態!!" & vbCrLf

        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then Errmsg &= "請選擇，申請階段" & vbCrLf

        rst = (Errmsg = "")
        Return rst
    End Function

    ''' <summary>查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Call SSearchDATA1()
    End Sub

    ''' <summary>檢核 ORG_SFCASEPI</summary>
    ''' <param name="oConn"></param>
    ''' <param name="s_PSNO28"></param>
    ''' <returns></returns>
    Public Shared Function CHK_ORG_SFCASEPI_EXISTS(ByRef oConn As SqlConnection, ByVal s_PSNO28 As String) As Boolean
        Dim rst As Boolean = False
        If s_PSNO28 = "" Then Return rst
        Dim dt1 As New DataTable
        Dim sSql As String = ""
        sSql &= " SELECT a.SFCID,a.SFCPID,a.PLANID,a.COMIDNO,a.SEQNO ,pp.PSNO28" & vbCrLf
        sSql &= " FROM PLAN_PLANINFO pp WITH(NOLOCK)" & vbCrLf
        sSql &= " JOIN ORG_SFCASEPI a WITH(NOLOCK) on a.PLANID=pp.PLANID AND a.COMIDNO=pp.COMIDNO AND a.SEQNO=pp.SEQNO" & vbCrLf
        sSql &= " WHERE pp.PSNO28=@PSNO28" & vbCrLf
        Dim sCmd As New SqlCommand(sSql, oConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            dt1.Load(.ExecuteReader)
        End With
        rst = (dt1.Rows.Count > 0)
        Return rst
    End Function

    ''' <summary>修改，清空所有欄位</summary>
    ''' <param name="htPMS"></param>
    Private Sub UPDATE_DELETE_DATA2(htPMS As Hashtable)
        Dim iRst As Integer = 0

        Hid_PSOID.Value = TIMS.GetMyValue2(htPMS, "PSOID") 'TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.GetMyValue2(htPMS, "PSNO28") ' TIMS.ClearSQM(Hid_PSNO28.Value)
        Hid_PSNO28.Value = TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSOID.Value = "" OrElse Hid_PSNO28.Value = "" Then Return

        Dim fg_SF_EXISTS As Boolean = CHK_ORG_SFCASEPI_EXISTS(objconn, Hid_PSNO28.Value)
        If fg_SF_EXISTS Then
            Common.MessageBox(Me, "申復線上送件使用中，請先刪除申復線上送件!")
            Return
        End If

        SFCONTNAME.Text = TIMS.ClearSQM(SFCONTNAME.Text)
        SFCONTTITLE.Text = TIMS.ClearSQM(SFCONTTITLE.Text)
        SFCONTTEL.Text = TIMS.ClearSQM(SFCONTTEL.Text)
        SFCONTEMAIL.Text = TIMS.ClearSQM(SFCONTEMAIL.Text)
        SFCONTREASONS.Text = Trim(SFCONTREASONS.Text)

        'uParms.Add("MODIFYDATE", MODIFYDATE)
        Dim uParms As New Hashtable From {
            {"SFCONTNAME", DBNull.Value},
            {"SFCONTTITLE", DBNull.Value},
            {"SFCONTTEL", DBNull.Value},
            {"SFCONTEMAIL", DBNull.Value},
            {"SFCONTREASONS", DBNull.Value},
            {"SFCONTACCT", DBNull.Value},
            {"SFCONTDATE", DBNull.Value},
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"PSOID", Val(Hid_PSOID.Value)},
            {"PSNO28", Hid_PSNO28.Value}
        }

        Dim usSql As String = ""
        usSql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        usSql &= " SET SFCONTNAME=@SFCONTNAME,SFCONTTITLE=@SFCONTTITLE,SFCONTTEL=@SFCONTTEL,SFCONTEMAIL=@SFCONTEMAIL" & vbCrLf
        usSql &= " ,SFCONTREASONS=@SFCONTREASONS,SFCONTACCT=@SFCONTACCT,SFCONTDATE=@SFCONTDATE" & vbCrLf
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        iRst = DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If

        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
        Call SSearchDATA1()
    End Sub
    Private Sub DtPlan_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles dtPlan.ItemCommand
        'Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim sCmdArg As String = e.CommandArgument
        Dim rPSOID As String = TIMS.GetMyValue(sCmdArg, "PSOID")
        Dim rPSNO28 As String = TIMS.GetMyValue(sCmdArg, "PSNO28")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue(sCmdArg, "APPSTAGE")
        If rPSOID = "" Then Return
        Dim iPSOID As Integer = TIMS.VAL1(rPSOID)
        Hid_PSOID.Value = rPSOID
        Hid_PSNO28.Value = rPSNO28

        '申請階段管理-受理期間設定 APPLISTAGE
        Dim aParms As New Hashtable From {{"YEARS", vYEARS}, {"APPSTAGE", vAPPSTAGE}}
        '開放受理之申請階段／PLAN_APPSTAGE
        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE02(objconn, aParms)
        '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
        If Not fg_can_applistage Then
            Common.MessageBox(Me, cst_stopmsg_12)
            Return
        End If

        Select Case e.CommandName
            Case "lbtSFEDIT1" '申復
                Dim hParms As New Hashtable From {
                    {"PSOID", Hid_PSOID.Value},
                    {"PSNO28", Hid_PSNO28.Value}
                }
                Dim dr2 As DataRow = SSearchDATA2(hParms)
                Call SHOW_DATA2(dr2)

            Case "lbtSFDEL1" '刪除
                Dim hParms As New Hashtable From {
                    {"PSOID", Hid_PSOID.Value},
                    {"PSNO28", Hid_PSNO28.Value}
                }
                Call UPDATE_DELETE_DATA2(hParms)

            Case "lbtSFPRINT1" '列印 申復意見表
                Dim hParms As New Hashtable From {
                    {"PSOID", Hid_PSOID.Value},
                    {"PSNO28", Hid_PSNO28.Value}
                }
                Dim dr2 As DataRow = SSearchDATA2(hParms)
                If dr2 Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Return
                End If
                Dim vDISTID As String = Convert.ToString(dr2("DISTID"))
                Dim MyValue1 As String = ""
                TIMS.SetMyValue(MyValue1, "TPlanID", sm.UserInfo.TPlanID)
                TIMS.SetMyValue(MyValue1, "YEARS", sm.UserInfo.Years)
                TIMS.SetMyValue(MyValue1, "PSOID", Hid_PSOID.Value)
                TIMS.SetMyValue(MyValue1, "PSNO28", Hid_PSNO28.Value)
                TIMS.SetMyValue(MyValue1, "DISTID", vDISTID)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)

        End Select
    End Sub
    Private Sub DtPlan_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labAppliedResult As Label = e.Item.FindControl("labAppliedResult")
                Dim lbtSFEDIT1 As LinkButton = e.Item.FindControl("lbtSFEDIT1")
                Dim lbtSFDEL1 As LinkButton = e.Item.FindControl("lbtSFDEL1")
                Dim lbtSFPRINT1 As LinkButton = e.Item.FindControl("lbtSFPRINT1")
                '
                e.Item.Cells(Cst_index).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim strMsg As String = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(drv("AppliedResult")), Convert.ToString(drv("RESULTBUTTON")))
                labAppliedResult.Text = strMsg
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PSOID", Convert.ToString(drv("PSOID")))
                TIMS.SetMyValue(sCmdArg, "PSNO28", Convert.ToString(drv("PSNO28")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                TIMS.SetMyValue(sCmdArg, "APPSTAGE", Convert.ToString(drv("APPSTAGE")))
                TIMS.SetMyValue(sCmdArg, "SFCID_Y", Convert.ToString(drv("SFCID_Y")))
                lbtSFEDIT1.CommandArgument = sCmdArg
                lbtSFDEL1.CommandArgument = sCmdArg
                lbtSFPRINT1.CommandArgument = sCmdArg

                '申復線上送件使用中，請先刪除申復線上送件!
                Dim fg_SFCID_RID As Boolean = (Convert.ToString(drv("SFCID_RID")) <> "")
                Dim fg_SFCID_Y As Boolean = (Convert.ToString(drv("SFCID_Y")) <> "")
                If fg_SFCID_Y Then
                    lbtSFEDIT1.Text = cst_lbtSFEDIT1_Txt_查看
                    lbtSFDEL1.Enabled = False
                    TIMS.Tooltip(lbtSFDEL1, "申復線上送件-使用中", True)
                Else
                    lbtSFEDIT1.Text = If(Convert.ToString(drv("SFCONT_CANEDIT")) = "Y", cst_lbtSFEDIT1_Txt_修改, cst_lbtSFEDIT1_Txt_申復)
                    'If fg_SFCID_RID Then lbtSFEDIT1.Attributes("onclick") = "return confirm('(單位業務)申復線上送件-使用中!');"
                    TIMS.Tooltip(lbtSFEDIT1, If(fg_SFCID_RID, "單位申復線上送件-此班未加入", ""), True)

                    lbtSFDEL1.Enabled = (Convert.ToString(drv("SFCONT_CANDEL")) = "Y")
                    TIMS.Tooltip(lbtSFDEL1, If(lbtSFDEL1.Enabled, "", "無申復資料"), True)
                    If lbtSFDEL1.Enabled Then lbtSFDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                End If
                lbtSFPRINT1.Enabled = (Convert.ToString(drv("SFCONT_CANPRINT")) = "Y")
                TIMS.Tooltip(lbtSFPRINT1, If(lbtSFPRINT1.Enabled, "", "無申復資料"), True)
        End Select
    End Sub

    ''' <summary>回上一頁</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        Call SSearchDATA1()
    End Sub

    ''' <summary>儲存資料 (單筆)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSAVE1_Click(sender As Object, e As EventArgs) Handles btnSAVE1.Click
        Call SAVEDATA1()
    End Sub

    ''' <summary>SAVEDATA1 儲存資料 (單筆) UPDATE PLAN_STAFFOPIN </summary>
    Private Sub SAVEDATA1()
        Dim iRst As Integer = 0

        Hid_PSOID.Value = TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSOID.Value = "" OrElse Hid_PSNO28.Value = "" Then Return

        SFCONTNAME.Text = TIMS.ClearSQM(SFCONTNAME.Text)
        SFCONTTITLE.Text = TIMS.ClearSQM(SFCONTTITLE.Text)
        SFCONTTEL.Text = TIMS.ClearSQM(SFCONTTEL.Text)
        SFCONTEMAIL.Text = TIMS.ClearSQM(SFCONTEMAIL.Text)
        SFCONTREASONS.Text = Trim(SFCONTREASONS.Text)

        Dim sERRMSG As String = ""
        If SFCONTNAME.Text = "" Then sERRMSG &= "請填寫，聯絡人!" & vbCrLf
        If SFCONTTEL.Text = "" Then sERRMSG &= "請填寫，聯絡電話!" & vbCrLf
        If SFCONTTITLE.Text = "" Then sERRMSG &= "請填寫，職稱!" & vbCrLf
        If SFCONTREASONS.Text = "" Then sERRMSG &= "請填寫，EMAIL!" & vbCrLf
        '申復理由及說明
        If SFCONTREASONS.Text = "" Then sERRMSG &= "請填寫，申復理由及說明!" & vbCrLf
        If SFCONTREASONS.Text <> "" AndAlso SFCONTREASONS.Text.Length < 2 Then sERRMSG &= "請填寫(完整)，申復理由及說明!" & vbCrLf

        SFCONTNAME.Text = TIMS.Get_Substr1(SFCONTNAME.Text, 60)
        SFCONTTITLE.Text = TIMS.Get_Substr1(SFCONTTITLE.Text, 100)
        SFCONTTEL.Text = TIMS.Get_Substr1(SFCONTTEL.Text, 33)
        SFCONTEMAIL.Text = TIMS.Get_Substr1(SFCONTEMAIL.Text, 66)
        If SFCONTEMAIL.Text <> "" AndAlso Not TIMS.CheckEmail(SFCONTEMAIL.Text) Then sERRMSG &= "EMAIL格式有誤" & vbCrLf
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Return
        End If

        'uParms.Add("SFCONTDATE", SFCONTDATE)
        Dim uParms As New Hashtable From {
            {"PSOID", Val(Hid_PSOID.Value)},
            {"PSNO28", Hid_PSNO28.Value},
            {"SFCONTNAME", SFCONTNAME.Text},
            {"SFCONTTITLE", SFCONTTITLE.Text},
            {"SFCONTTEL", SFCONTTEL.Text},
            {"SFCONTEMAIL", SFCONTEMAIL.Text},
            {"SFCONTREASONS", SFCONTREASONS.Text},
            {"SFCONTACCT", sm.UserInfo.UserID},
            {"MODIFYACCT", sm.UserInfo.UserID}
        }
        'uParms.Add("MODIFYDATE", MODIFYDATE)
        Dim usSql As String = ""
        usSql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        usSql &= " SET SFCONTNAME=@SFCONTNAME,SFCONTTITLE=@SFCONTTITLE,SFCONTTEL=@SFCONTTEL,SFCONTEMAIL=@SFCONTEMAIL" & vbCrLf
        usSql &= " ,SFCONTREASONS=@SFCONTREASONS,SFCONTACCT=@SFCONTACCT,SFCONTDATE=GETDATE()" & vbCrLf '申復理由及說明
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        iRst = DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

    ''' <summary>取得部份機構資訊</summary>
    ''' <param name="htParms"></param>
    ''' <returns></returns>
    Function GET_ORG_ORGPLANINFO_row(htParms As Hashtable) As DataRow
        Dim vRID As String = TIMS.GetMyValue2(htParms, "RID")
        Dim vPLANID As String = TIMS.GetMyValue2(htParms, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(htParms, "COMIDNO")

        Dim sPMS As New Hashtable From {{"RID", vRID}, {"PLANID", vPLANID}, {"COMIDNO", vCOMIDNO}}
        Dim sSql As String = ""
        sSql &= " SELECT op.RSID,op.CONTACTNAME,op.CONTACTCELLPHONE,op.CONTACTEMAIL,op.PHONE" & vbCrLf
        sSql &= " ,ar.RID,oo.COMCIDNO,ip.PLANID" & vbCrLf
        sSql &= " FROM ORG_ORGPLANINFO op" & vbCrLf
        sSql &= " JOIN AUTH_RELSHIP ar on ar.RSID=op.RSID" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.ORGID=ar.ORGID" & vbCrLf
        sSql &= " JOIN ID_PLAN ip on ip.PLANID =ar.PLANID" & vbCrLf
        sSql &= " WHERE ar.RID=@RID AND ip.PLANID=@PLANID AND oo.COMIDNO=@COMIDNO" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sSql, objconn, sPMS)
        Return dr
    End Function

#Region "NO USE"
    'Protected Sub btnONLINE1_Click(sender As Object, e As EventArgs) Handles btnONLINE1.Click
    '    Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
    '    If v_APPSTAGE_SCH = "" Then
    '        msg1.Text = TIMS.cst_NODATAMsg2
    '        Return
    '    End If
    '    '申請階段管理-受理期間設定 APPLISTAGE
    '    Dim aParms As New Hashtable
    '    aParms.Add("YEARS", sm.UserInfo.Years)
    '    aParms.Add("APPSTAGE", v_APPSTAGE_SCH)
    '    '開放受理之申請階段／PLAN_APPSTAGE
    '    Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE02(objconn, aParms)
    '    '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
    '    If Not fg_can_applistage Then
    '        Common.MessageBox(Me, cst_stopmsg_12)
    '        Return
    '    End If

    '    SHOW_ONLINE_DATA1()

    'End Sub

#End Region

End Class
