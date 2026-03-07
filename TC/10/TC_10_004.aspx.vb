Partial Class TC_10_004
    Inherits AuthBasePage

    Const cst_printFN1 As String = "TC_10_004R"
    '因為「政策性產業」課程並非像「上、下半年」整合五分署依職類進行審查，而是由各分署各自辦理 '會議雖是同一天，但其實是各自舉辦
    Const cst_printFN2 As String = "TC_10_004R2"

    Const cst_UPD1 As String = "UPD1" '修改
    Const cst_DEL1 As String = "DEL1" '刪除
    Const cst_PRT1 As String = "PRT1" '列印
    'Const cst_EXP1 As String = "EXP1" '匯出

    Dim dic_AGE As Dictionary(Of String, String) = TIMS.Get_ACCEPTSTAGE_DIC()
    Dim a_vAGE1() As String = {"A1", "B1", "C1", "D1"} '1 (初次申請)
    Dim a_vAGE2() As String = {"A2", "B2", "C2", "D2"} '2 (申復)

    'Dim dtGCODE3 As DataTable = Nothing
    Dim dtDist As DataTable = Nothing

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

        '分署資料查詢
        dtDist = TIMS.Get_DISTIDT2(objconn)
        '審查職類代碼
        'dtGCODE3 = TIMS.Get_GOVCODE3dt(objconn)
        'dtGCODE3_5 = TIMS.GET_dtGCODE3(objconn, hid_GOVCODE3.Value)
        'Dim dt As DataTable = Get_GOVCODE3dt(oConn)
        msg1.Text = ""
        ShowButton1()

        If Not IsPostBack Then cCreate1()
    End Sub

    Sub ShowButton1()
        '分署登入鎖定登入者的分署，署登入可選擇下拉選單。
        ' 其他角色： 暫不開放使用(目前功能暫不開放給其他群組用)

        'btnSAVE1.Enabled = If(sm.UserInfo.LID = 0, True, False)
        'BtnSEARCH.Enabled = If(sm.UserInfo.LID = 0, True, False)
        'BtnADDNEW.Enabled = If(sm.UserInfo.LID = 0, True, False)

        'If (sm.UserInfo.LID = 1) AndAlso (sm.UserInfo.RoleID <= 1) Then
        '    btnSAVE1.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
        '    BtnSEARCH.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
        '    BtnADDNEW.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
        'End If

        'Const cst_tipmsg1 As String = "分署-系統管理者：可使用"
        'If Not btnSAVE1.Enabled Then TIMS.Tooltip(btnSAVE1, cst_tipmsg1)
        'If Not BtnSEARCH.Enabled Then TIMS.Tooltip(BtnSEARCH, cst_tipmsg1)
        'If Not BtnADDNEW.Enabled Then TIMS.Tooltip(BtnADDNEW, cst_tipmsg1)
    End Sub

    Sub cCreate1()
        'hid_EXAMINER_TABLE_GUID1.Value = TIMS.GetGUID()
        'Session(hid_EXAMINER_TABLE_GUID1.Value) = Nothing
        PageControler1.Visible = False
        btnSAVE1.Attributes("onclick") = "return chkSaveData1();"
        '單筆
        ddlDISTID = TIMS.Get_DistID(ddlDISTID, dtDist) '主責分署/轄區分署
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '轄區分署
        ddlDISTID.Enabled = Not (sm.UserInfo.LID <> 0)
        If Not ddlDISTID.Enabled Then TIMS.Tooltip(ddlDISTID, "分署：分署各自維護")

        '下拉式選單。最大值以當年度+1
        Dim iSYears As Integer = 2023 'OJT-22090601：<系統> 產投-會議與出席場次管理：邏輯調整
        Dim iEYears As Integer = If((Now.Year + 1) < 2023, 2023, (Now.Year + 1))

        ddlMYEARS = TIMS.GetSyear(ddlMYEARS, iSYears, iEYears, True) '年度
        Common.SetListItem(ddlMYEARS, sm.UserInfo.Years) '年度

        TIMS.SUB_SET_HR_MI(HR1, MM1) '時分
        Common.SetListItem(HR1, "09")
        Common.SetListItem(MM1, "00")

        TIMS.SUB_SET_HR_MI(HR2, MM2) '時分
        Common.SetListItem(HR2, "18")
        Common.SetListItem(MM2, "00")

        '受理階段 ACCEPTSTAGE
        ddlACCEPTSTAGE = TIMS.Get_ACCEPTSTAGE(ddlACCEPTSTAGE)
        '受理階段 ACCEPTSTAGE
        ddlACCEPTSTAGE_sch = TIMS.Get_ACCEPTSTAGE(ddlACCEPTSTAGE_sch)

        '查詢 SCH
        ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, dtDist) '轄區分署
        Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID) '轄區分署
        '轄區：每個轄區分署各自舉辦，各分署各自維護出席審查委員
        '跨區：多個分署共同參加(輪流主辦)，由主責轄區分署維護出席審查委員
        ddlDISTID_SCH.Enabled = True
        If (sm.UserInfo.LID <> 0) Then ddlDISTID_SCH.Enabled = False
        If Not ddlDISTID_SCH.Enabled Then TIMS.Tooltip(ddlDISTID_SCH, "分署：分署各自維護")


        ddlMYEARS_SCH = TIMS.GetSyear(ddlMYEARS_SCH, iSYears, iEYears, True)  '年度
        'Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID) '轄區分署(查詢)
        Common.SetListItem(ddlMYEARS_SCH, sm.UserInfo.Years) '年度(查詢)

        SHOW_PANEL(0)
    End Sub

    ''' <summary> 查詢list </summary>
    Sub sSearch1()
        Dim s_ERRMSG As String = ""

        Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH) '轄區分署
        Dim v_ddlMYEARS_SCH As String = TIMS.GetListValue(ddlMYEARS_SCH) '年度
        'Dim v_rblCATEGORY_SCH As String = TIMS.GetListValue(rblCATEGORY_SCH) '審查會議類別
        Dim v_cblORGPLANKIND_sch As String = TIMS.GetCblValue(cblORGPLANKIND_sch) '計畫別 G,W
        Dim v_ddlACCEPTSTAGE_sch As String = TIMS.GetListValue(ddlACCEPTSTAGE_sch) '受理階段 ACCEPTSTAGE

        SMEETDATE_sch1.Text = TIMS.Cdate3(TIMS.ClearSQM(SMEETDATE_sch1.Text)) '會議日期/時間-開始1
        SMEETDATE_sch2.Text = TIMS.Cdate3(TIMS.ClearSQM(SMEETDATE_sch2.Text)) '會議日期/時間-開始2
        If TIMS.ChkDateErr3(SMEETDATE_sch1.Text, SMEETDATE_sch2.Text) Then
            Dim T_DATE1 As String = SMEETDATE_sch1.Text
            SMEETDATE_sch1.Text = SMEETDATE_sch2.Text
            SMEETDATE_sch2.Text = T_DATE1
        End If

        MEETPLACE_sch.Text = TIMS.ClearSQM(MEETPLACE_sch.Text) '會議地點
        MODERATOR_sch.Text = TIMS.ClearSQM(MODERATOR_sch.Text) '主席 文字框，30個字元

        Dim lk_MEETPLACE_sch As String = If(MEETPLACE_sch.Text <> "", String.Format("%{0}%", MEETPLACE_sch.Text), "")
        Dim lk_MODERATOR_sch As String = If(MODERATOR_sch.Text <> "", String.Format("%{0}%", MODERATOR_sch.Text), "")

        If v_ddlMYEARS_SCH = "" Then s_ERRMSG &= "請選擇 查詢年度" & vbCrLf

        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, s_ERRMSG)
            Return
        End If

        Dim sSql As String = ""
        sSql &= " SELECT a.MRSEQ ,a.DISTID,d1.NAME DISTNAME" & vbCrLf
        sSql &= " ,a.MYEARS" & vbCrLf
        sSql &= " ,a.ORGPLANKIND" & vbCrLf '計畫別 G,W
        sSql &= " ,CASE a.ORGPLANKIND WHEN 'G' THEN '產業人才投資計畫' WHEN 'W' THEN '提升勞工自主學習計畫' WHEN 'G,W' THEN '產業人才投資、提升勞工自主' END ORGPLANKIND_N" & vbCrLf
        sSql &= " ,a.ACCEPTSTAGE" & vbCrLf '受理階段 ACCEPTSTAGE
        sSql &= " ,CASE a.ACCEPTSTAGE" & vbCrLf
        For i As Integer = 0 To a_vAGE1.Length - 1
            Dim vID1 As String = a_vAGE1(i) : Dim vNM1 As String = dic_AGE(vID1)
            Dim vID2 As String = a_vAGE2(i) : Dim vNM2 As String = dic_AGE(vID2)
            sSql &= String.Format("  WHEN '{0}' THEN '{1}' WHEN '{2}' THEN '{3}'", vID1, vNM1, vID2, vNM2)
            sSql &= If(i = (a_vAGE1.Length - 1), " END ACCEPTSTAGE_N", "") & vbCrLf
        Next
        'sql &= "  WHEN 'A1' THEN '上半年' WHEN 'A2' THEN '上半年申復'" & vbCrLf
        'sql &= "  WHEN 'B1' THEN '政策性' WHEN 'B2' THEN '政策性申復'" & vbCrLf
        'sql &= "  WHEN 'C1' THEN '下半年' WHEN 'C2' THEN '下半年申復'" & vbCrLf
        'sql &= "  WHEN 'D1' THEN '進階政策性' WHEN 'D2' THEN '進階政策性申復' END ACCEPTSTAGE_N" & vbCrLf

        sSql &= " ,a.SMEETDATE" & vbCrLf
        sSql &= " ,a.FMEETDATE" & vbCrLf
        '會議時間
        sSql &= " ,CONCAT(format(a.SMEETDATE,'yyyy/MM/dd HH:mm'),' ~ ',format(a.FMEETDATE,'HH:mm')) SFMEETDATE_N" & vbCrLf
        sSql &= " ,a.MEETPLACE" & vbCrLf
        sSql &= " ,a.MEETADDRESS" & vbCrLf
        sSql &= " ,a.SPEECHMAN" & vbCrLf
        sSql &= " ,a.MHOSTER" & vbCrLf
        sSql &= " ,a.AGENDA" & vbCrLf
        'sSql &= " ,a.CREATEACCT" & vbCrLf
        'sSql &= " ,a.CREATEDATE" & vbCrLf
        'sSql &= " ,a.MODIFYACCT" & vbCrLf
        'sSql &= " ,a.MODIFYDATE" & vbCrLf
        sSql &= " ,a.RID" & vbCrLf
        sSql &= " FROM dbo.OA_MEETINGRPT a" & vbCrLf
        sSql &= " JOIN dbo.ID_DISTRICT d1 on d1.DISTID=a.DISTID" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf

        '主責分署
        If sm.UserInfo.LID > 0 Then
            sSql &= " AND a.DISTID='" & v_ddlDISTID_SCH & "'" & vbCrLf
        ElseIf v_ddlDISTID_SCH <> "" Then
            sSql &= " AND a.DISTID='" & v_ddlDISTID_SCH & "'" & vbCrLf
        End If
        '年度
        If v_ddlMYEARS_SCH <> "" Then sSql &= " AND a.MYEARS='" & v_ddlMYEARS_SCH & "'" & vbCrLf

        '受理階段 A1:上半年 A2:上半年申復
        '受理階段 B1:政策性 B2:政策性申復
        '受理階段 C1:下半年 B2:下半年申復
        If v_ddlACCEPTSTAGE_sch <> "" Then sSql &= " AND a.ACCEPTSTAGE='" & v_ddlACCEPTSTAGE_sch & "'" & vbCrLf

        '會議日期
        If SMEETDATE_sch1.Text <> "" Then sSql &= " AND CONVERT(date,a.SMEETDATE)>=CONVERT(date,'" & SMEETDATE_sch1.Text & "') " & vbCrLf
        If SMEETDATE_sch2.Text <> "" Then sSql &= " AND CONVERT(date,a.SMEETDATE)<=CONVERT(date,'" & SMEETDATE_sch2.Text & "') " & vbCrLf
        'If SMEETDATE_sch1.Text <> "" Then sSql &= " AND a.SMEETDATE >='" & SMEETDATE_sch1.Text & "'" & vbCrLf
        'If SMEETDATE_sch2.Text <> "" Then sSql &= " AND a.SMEETDATE <='" & SMEETDATE_sch2.Text & "'" & vbCrLf

        '地點/地址
        If lk_MEETPLACE_sch <> "" Then
            sSql &= " AND (a.MEETPLACE like '" & lk_MEETPLACE_sch & "' OR a.MEETADDRESS like '" & lk_MEETPLACE_sch & "')" & vbCrLf
        End If

        '致詞主席/主責分署主持人
        If lk_MODERATOR_sch <> "" Then
            sSql &= " AND (a.SPEECHMAN like '" & lk_MODERATOR_sch & "' OR a.MHOSTER like '" & lk_MODERATOR_sch & "')" & vbCrLf
        End If

        msg1.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn)
        If dt.Rows.Count = 0 Then Return 'Common.MessageBox(Me, "查無資料") 'Exit Sub

        msg1.Text = ""
        DataGrid1.Visible = True
        PageControler1.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary> SHOW_PANEL </summary>
    ''' <param name="iType"></param>
    Sub SHOW_PANEL(ByVal iType As Integer)
        'iType 0:查詢/1:修改 
        panelEdit.Visible = False
        panelSch.Visible = False

        Select Case iType
            Case 0
                panelSch.Visible = True
            Case 1
                panelEdit.Visible = True
        End Select
    End Sub

    Sub ClearData1()
        Hid_MRSEQ.Value = ""
        'hid_GOVCODE3.Value = ""
        'Hid_MTSEQ.Value = ""
        'Session(hid_EXAMINER_TABLE_GUID1.Value) = Nothing
        'CreateTableDG2("", 1, If(Hid_MTSEQ.Value = "", 0, Val(Hid_MTSEQ.Value)))
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '轄區分署
        Common.SetListItem(ddlMYEARS, sm.UserInfo.Years) '年度

        'rblCATEGORY.SelectedIndex = -1 '審查會議類別
        TIMS.SetCblValue(cblORGPLANKIND, "") '計畫別 G,W

        ddlACCEPTSTAGE.SelectedIndex = -1 '受理階段

        SMEETDATE.Text = "" '會議日期/時間-開始
        Common.SetListItem(HR1, "09") '會議日期/時間-開始
        Common.SetListItem(MM1, "00") '會議日期/時間-開始
        'FMEETDATE.Text = "" '會議日期/時間-結束
        Common.SetListItem(HR2, "18") '會議日期/時間-結束
        Common.SetListItem(MM2, "00") '會議日期/時間-結束

        'FMEETDATE.TEXT = ""
        MEETPLACE.Text = "" '會議/地點
        MEETADDRESS.Text = "" '地址

        SPEECHMAN.Text = "" '致詞主席/主席 文字框，30個字元
        MHOSTER.Text = "" '主責分署主持人 文字框，30個字元
        AGENDA.Text = "" '議程
    End Sub

    Sub LoadData1(ByRef iMRSEQ As Integer)
        Hid_MRSEQ.Value = TIMS.ClearSQM(Hid_MRSEQ.Value)
        If Hid_MRSEQ.Value = "" Then Return
        If iMRSEQ <> Val(Hid_MRSEQ.Value) Then Return

        Dim sql As String = ""
        sql = "SELECT * FROM dbo.OA_MEETINGRPT WHERE MRSEQ=@MRSEQ "
        Dim parms As New Hashtable From {{"MRSEQ", iMRSEQ}}
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr1 Is Nothing Then Return

        Common.SetListItem(ddlDISTID, Convert.ToString(dr1("DISTID"))) '主責分署
        Common.SetListItem(ddlMYEARS, Convert.ToString(dr1("MYEARS"))) '年度
        TIMS.SetCblValue(cblORGPLANKIND, Convert.ToString(dr1("ORGPLANKIND"))) '計畫別 G,W
        Common.SetListItem(ddlACCEPTSTAGE, Convert.ToString(dr1("ACCEPTSTAGE"))) '受理階段

        SMEETDATE.Text = TIMS.Cdate3(dr1("SMEETDATE")) '會議日期/時間-開始
        If SMEETDATE.Text <> "" Then TIMS.SET_DateHM(CDate(dr1("SMEETDATE")), HR1, MM1)
        'FMEETDATE.Text = TIMS.cdate3(dr1("FMEETDATE")) '會議日期/時間-結束

        Dim v_FMEETDATE As String = TIMS.Cdate3(dr1("FMEETDATE")) '會議日期/時間-結束
        If v_FMEETDATE <> "" Then TIMS.SET_DateHM(CDate(dr1("FMEETDATE")), HR2, MM2)

        'FMEETDATE.TEXT = ""
        MEETPLACE.Text = Convert.ToString(dr1("MEETPLACE")) '地點
        MEETADDRESS.Text = Convert.ToString(dr1("MEETADDRESS")) '地址

        SPEECHMAN.Text = Convert.ToString(dr1("SPEECHMAN")) '致詞主席 文字框，30個字元
        MHOSTER.Text = Convert.ToString(dr1("MHOSTER")) '主責分署主持人
        AGENDA.Text = Convert.ToString(dr1("AGENDA")) '議程 AGENDA

        SHOW_PANEL(1)
    End Sub

    ''' <summary> '檢查 </summary>
    ''' <param name="s_ERRMSG">有值為異常:False</param>
    ''' <returns></returns>
    Function CheckData1(ByRef s_ERRMSG As String) As Boolean
        Dim rst As Boolean = True
        s_ERRMSG = ""

        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID) '主責分署/轄區分署
        Dim v_ddlMYEARS As String = TIMS.GetListValue(ddlMYEARS) '年度
        'Dim v_rblCATEGORY As String = TIMS.GetListValue(rblCATEGORY) '審查會議類別
        Dim v_cblORGPLANKIND As String = TIMS.GetCblValue(cblORGPLANKIND) '計畫別 G,W
        Dim v_ddlACCEPTSTAGE As String = TIMS.GetListValue(ddlACCEPTSTAGE) '受理階段 ACCEPTSTAGE
        SMEETDATE.Text = TIMS.Cdate3(TIMS.ClearSQM(SMEETDATE.Text)) '會議日期/時間-開始
        Dim s_SMEETDATE As String = TIMS.GET_DateHM(SMEETDATE, HR1, MM1) '會議日期/時間-開始
        Dim s_FMEETDATE As String = TIMS.GET_DateHM(SMEETDATE, HR2, MM2) '會議日期/時間-結束 (Single)

        MEETPLACE.Text = TIMS.ClearSQM(MEETPLACE.Text) '會議地點/地點
        MEETADDRESS.Text = TIMS.ClearSQM(MEETADDRESS.Text) '會議地址/地址
        SPEECHMAN.Text = TIMS.ClearSQM(SPEECHMAN.Text) '致詞主席/主席 SPEECHMAN 文字框，30個字元
        MHOSTER.Text = TIMS.ClearSQM(MHOSTER.Text) '主責分署主持人 MHOSTER 文字框，30個字元
        AGENDA.Text = TIMS.ClearSQM2(AGENDA.Text) '議程 AGENDA 

        If v_ddlDISTID = "" Then s_ERRMSG &= "請選擇 主責分署" & vbCrLf
        If v_ddlMYEARS = "" Then s_ERRMSG &= "請選擇 年度" & vbCrLf
        'If v_rblCATEGORY = "" Then s_ERRMSG &= "請選擇 審查會議類別" & vbCrLf
        If v_cblORGPLANKIND = "" Then s_ERRMSG &= "請選擇 計畫別" & vbCrLf '計畫別 G,W

        If v_ddlACCEPTSTAGE = "" Then s_ERRMSG &= "請選擇 受理階段" & vbCrLf '受理階段 ACCEPTSTAGE
        If s_SMEETDATE = "" Then s_ERRMSG &= "請選擇輸入 日期/時間 -日期" & vbCrLf
        'If s_FMEETDATE = "" Then s_ERRMSG &= "請選擇輸入 日期/時間 -結束" & vbCrLf
        If MEETPLACE.Text = "" Then s_ERRMSG &= "請輸入 地點" & vbCrLf
        If MEETPLACE.Text = "" Then s_ERRMSG &= "請輸入 地址" & vbCrLf
        'If SPEECHMAN.Text = "" Then s_ERRMSG &= "請輸入 致詞主席" & vbCrLf
        'If MHOSTER.Text = "" Then s_ERRMSG &= "請輸入 主責分署主持人" & vbCrLf
        'If AGENDA.Text = "" Then s_ERRMSG &= "請輸入 議程" & vbCrLf
        If s_ERRMSG <> "" Then Return False

        If s_SMEETDATE <> "" AndAlso s_FMEETDATE <> "" Then
            If IsDate(s_SMEETDATE) AndAlso IsDate(s_FMEETDATE) Then
                If DateDiff(DateInterval.Minute, CDate(s_SMEETDATE), CDate(s_FMEETDATE)) < 0 Then s_ERRMSG &= "日期/時間起迄，順序有誤" & vbCrLf
            Else
                s_ERRMSG &= "日期/時間起迄，格式有誤" & vbCrLf
            End If
        End If

        '分署，只能選擇自已
        If sm.UserInfo.LID <> 0 AndAlso v_ddlDISTID <> sm.UserInfo.DistID Then
            s_ERRMSG &= " 主責分署 與登入分署不同(不可儲存)" & vbCrLf
        End If
        If s_ERRMSG <> "" Then Return False

        '取得 審查課程職類 TABLE 
        '申請階段:1:上半年/2:下半年/3:政策性產業/4:進階政策性產業
        Dim v_APPSTAGE As String = TIMS.GET_APPSTAGE_12(v_ddlACCEPTSTAGE)
        '3:政策性產業/4:進階政策性產業 不檢核 
        Dim fg_CanCheck1 As Boolean = If(v_APPSTAGE = "1", True, If(v_APPSTAGE = "2", True, False))
        If s_ERRMSG <> "" Then Return False

        Dim sParms As New Hashtable
        TIMS.SetMyValue2(sParms, "MYEARS", v_ddlMYEARS)
        TIMS.SetMyValue2(sParms, "ACCEPTSTAGE", v_ddlACCEPTSTAGE)
        TIMS.SetMyValue2(sParms, "SMEETDATE", SMEETDATE.Text)
        Dim fgCHK1 As Boolean = CHK_SAVEROLE1(sParms, Hid_MRSEQ.Value)
        If fgCHK1 Then s_ERRMSG &= "該 年度／受理階段／會議日期 已有資料 (不可儲存)!" & vbCrLf

        If s_ERRMSG <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>檢核日期只能有一筆資料:查無資料:false/有資料:true</summary>
    ''' <param name="sParms"></param>
    ''' <param name="s_MRSEQ"></param>
    ''' <returns></returns>
    Private Function CHK_SAVEROLE1(ByVal sParms As Hashtable, ByVal s_MRSEQ As String) As Boolean
        Dim iMRSEQ As Integer = If(s_MRSEQ <> "", Val(s_MRSEQ), 0)

        Dim s_MYEARS As String = TIMS.GetMyValue2(sParms, "MYEARS")
        Dim s_ACCEPTSTAGE As String = TIMS.GetMyValue2(sParms, "ACCEPTSTAGE")
        Dim s_SMEETDATE As String = TIMS.GetMyValue2(sParms, "SMEETDATE")
        Dim rst As Boolean = False

        Dim pParms As New Hashtable From {
            {"MYEARS", s_MYEARS},
            {"ACCEPTSTAGE", s_ACCEPTSTAGE},
            {"SMEETDATE", s_SMEETDATE}
        }
        If iMRSEQ > 0 Then pParms.Add("MRSEQ", iMRSEQ)

        Dim sSql As String = "SELECT 1 FROM OA_MEETINGRPT a" & vbCrLf
        sSql &= " WHERE a.MYEARS=@MYEARS" & vbCrLf
        sSql &= " and a.ACCEPTSTAGE=@ACCEPTSTAGE" & vbCrLf
        sSql &= " and CONVERT(date,a.SMEETDATE)=CONVERT(date,@SMEETDATE)" & vbCrLf
        If iMRSEQ > 0 Then sSql &= " AND MRSEQ!=@MRSEQ"

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    ''' <summary> '儲存 SAVE - 審查會議-預計參加審查委員名單</summary>
    Sub SaveData1()
        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID) '主責分署/轄區分署
        Dim v_ddlMYEARS As String = TIMS.GetListValue(ddlMYEARS) '年度
        'Dim v_rblCATEGORY As String = TIMS.GetListValue(rblCATEGORY) '審查會議類別
        Dim v_cblORGPLANKIND As String = TIMS.GetCblValue(cblORGPLANKIND) '計畫別 G,W
        Dim v_ddlACCEPTSTAGE As String = TIMS.GetListValue(ddlACCEPTSTAGE) '受理階段 ACCEPTSTAGE
        SMEETDATE.Text = TIMS.Cdate3(TIMS.ClearSQM(SMEETDATE.Text)) '會議日期/時間-開始
        'FMEETDATE.Text = TIMS.cdate3(TIMS.ClearSQM(FMEETDATE.Text)) '會議日期/時間-結束
        'Dim vFMEETDATE As String = SMEETDATE.Text ' 會議日期/時間-結束 (Single)
        Dim s_SMEETDATE As String = TIMS.GET_DateHM(SMEETDATE, HR1, MM1) '會議日期/時間-開始
        Dim s_FMEETDATE As String = TIMS.GET_DateHM(SMEETDATE, HR2, MM2) '會議日期/時間-結束 (Single)

        MEETPLACE.Text = TIMS.ClearSQM(MEETPLACE.Text) '會議地點/地點
        MEETADDRESS.Text = TIMS.ClearSQM(MEETADDRESS.Text) '會議地址/地址
        SPEECHMAN.Text = TIMS.ClearSQM(SPEECHMAN.Text) '致詞主席/主席 SPEECHMAN 文字框，30個字元
        MHOSTER.Text = TIMS.ClearSQM(MHOSTER.Text) '主責分署主持人 MHOSTER 文字框，30個字元
        AGENDA.Text = TIMS.ClearSQM2(AGENDA.Text) '議程 AGENDA 

        Dim rst As Integer = 0
        Dim flagSaveOK1 As Boolean = False

        Dim iSql As String = ""
        iSql = " INSERT INTO OA_MEETINGRPT(MRSEQ,DISTID,MYEARS,ORGPLANKIND,ACCEPTSTAGE,SMEETDATE,FMEETDATE,MEETPLACE,MEETADDRESS" & vbCrLf
        iSql &= " ,SPEECHMAN,MHOSTER,AGENDA,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,RID)" & vbCrLf
        iSql &= " VALUES (@MRSEQ,@DISTID,@MYEARS,@ORGPLANKIND,@ACCEPTSTAGE,@SMEETDATE,@FMEETDATE,@MEETPLACE,@MEETADDRESS" & vbCrLf
        iSql &= " ,@SPEECHMAN,@MHOSTER,@AGENDA,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE(),@RID)" & vbCrLf

        Dim uSql As String = ""
        uSql = " UPDATE OA_MEETINGRPT" & vbCrLf
        uSql &= " SET MRSEQ=@MRSEQ ,DISTID=@DISTID ,MYEARS=@MYEARS ,ORGPLANKIND=@ORGPLANKIND ,ACCEPTSTAGE=@ACCEPTSTAGE" & vbCrLf
        uSql &= " ,SMEETDATE=@SMEETDATE ,FMEETDATE=@FMEETDATE ,MEETPLACE=@MEETPLACE ,MEETADDRESS=@MEETADDRESS" & vbCrLf
        uSql &= " ,SPEECHMAN=@SPEECHMAN ,MHOSTER=@MHOSTER ,AGENDA=@AGENDA,RID=@RID" & vbCrLf
        uSql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        uSql &= " WHERE MRSEQ=@MRSEQ" & vbCrLf

        Dim iMRSEQ As Integer = If(Hid_MRSEQ.Value <> "", Val(Hid_MRSEQ.Value), 0)
        Dim parms As Hashtable = New Hashtable
        Hid_MRSEQ.Value = TIMS.ClearSQM(Hid_MRSEQ.Value)
        If Hid_MRSEQ.Value = "" Then
            '新增
            iMRSEQ = DbAccess.GetNewId(objconn, "OA_MEETINGRPT_MTSEQ_SEQ,OA_MEETINGRPT,MRSEQ")
            parms.Clear()
            parms.Add("MRSEQ ", iMRSEQ)
            parms.Add("DISTID ", If(v_ddlDISTID <> "", v_ddlDISTID, Convert.DBNull))
            parms.Add("MYEARS ", If(v_ddlMYEARS <> "", v_ddlMYEARS, Convert.DBNull))
            parms.Add("ORGPLANKIND ", If(v_cblORGPLANKIND <> "", v_cblORGPLANKIND, Convert.DBNull))
            parms.Add("ACCEPTSTAGE ", If(v_ddlACCEPTSTAGE <> "", v_ddlACCEPTSTAGE, Convert.DBNull))
            parms.Add("SMEETDATE ", If(s_SMEETDATE <> "", s_SMEETDATE, Convert.DBNull))
            parms.Add("FMEETDATE", If(s_FMEETDATE <> "", s_FMEETDATE, Convert.DBNull))
            parms.Add("MEETPLACE ", If(MEETPLACE.Text <> "", MEETPLACE.Text, Convert.DBNull))
            parms.Add("MEETADDRESS ", If(MEETADDRESS.Text <> "", MEETADDRESS.Text, Convert.DBNull))
            parms.Add("SPEECHMAN ", If(SPEECHMAN.Text <> "", SPEECHMAN.Text, Convert.DBNull))
            parms.Add("MHOSTER ", If(MHOSTER.Text <> "", MHOSTER.Text, Convert.DBNull))
            parms.Add("AGENDA ", If(AGENDA.Text <> "", AGENDA.Text, Convert.DBNull))
            parms.Add("RID", sm.UserInfo.RID)
            parms.Add("CREATEACCT ", sm.UserInfo.UserID)
            parms.Add("MODIFYACCT ", sm.UserInfo.UserID)

            rst = DbAccess.ExecuteNonQuery(iSql, objconn, parms)
            flagSaveOK1 = True
        Else
            '修改
            'iMRSEQ = Val(Hid_MRSEQ.Value)
            parms.Clear()
            parms.Add("DISTID ", If(v_ddlDISTID <> "", v_ddlDISTID, Convert.DBNull))
            parms.Add("MYEARS ", If(v_ddlMYEARS <> "", v_ddlMYEARS, Convert.DBNull))
            parms.Add("ORGPLANKIND ", If(v_cblORGPLANKIND <> "", v_cblORGPLANKIND, Convert.DBNull))
            parms.Add("ACCEPTSTAGE ", If(v_ddlACCEPTSTAGE <> "", v_ddlACCEPTSTAGE, Convert.DBNull))
            parms.Add("SMEETDATE ", If(s_SMEETDATE <> "", s_SMEETDATE, Convert.DBNull))
            parms.Add("FMEETDATE", If(s_FMEETDATE <> "", s_FMEETDATE, Convert.DBNull))

            parms.Add("MEETPLACE ", If(MEETPLACE.Text <> "", MEETPLACE.Text, Convert.DBNull))
            parms.Add("MEETADDRESS ", If(MEETADDRESS.Text <> "", MEETADDRESS.Text, Convert.DBNull))
            parms.Add("SPEECHMAN ", If(SPEECHMAN.Text <> "", SPEECHMAN.Text, Convert.DBNull))
            parms.Add("MHOSTER ", If(MHOSTER.Text <> "", MHOSTER.Text, Convert.DBNull))
            parms.Add("AGENDA ", If(AGENDA.Text <> "", AGENDA.Text, Convert.DBNull))
            parms.Add("RID", sm.UserInfo.RID)
            parms.Add("MODIFYACCT ", sm.UserInfo.UserID)

            parms.Add("MRSEQ ", iMRSEQ)
            rst = DbAccess.ExecuteNonQuery(uSql, objconn, parms)
            flagSaveOK1 = True
        End If

        '申請階段:1:上半年/2:下半年/3:政策性產業/4:進階政策性產業
        Dim v_APPSTAGE As String = TIMS.GET_APPSTAGE_12(v_ddlACCEPTSTAGE)
        '3:政策性產業/4:進階政策性產業 不檢核 
        Dim fg_CanCheck1 As Boolean = If(v_APPSTAGE = "1", True, If(v_APPSTAGE = "2", True, False))

        If Not flagSaveOK1 Then '儲存-失敗
            Common.MessageBox(Me, "儲存失敗!")
            Exit Sub
        End If

        SHOW_PANEL(0)

        '儲存成功 'Hid_EMSEQ.Value = ""
        Call ClearData1()
        Common.MessageBox(Me, "儲存成功!")
        Call sSearch1()
    End Sub


    ''' <summary>'查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>

    Protected Sub BtnSEARCH_Click(sender As Object, e As EventArgs) Handles BtnSEARCH.Click
        sSearch1()
    End Sub

    ''' <summary>'新增</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnADDNEW_Click(sender As Object, e As EventArgs) Handles BtnADDNEW.Click
        Call ClearData1()

        '北分署新增了一筆110年、跨區、上半年之審查會議，則以北分署為主責分署，其他分署不可再新增，除非北分署刪掉該筆會議資料。(理論上非主責分署不會去新增)
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '轄區分署
        ddlDISTID.Enabled = If(sm.UserInfo.LID = 0, True, False)

        SHOW_PANEL(1)
    End Sub

    ''' <summary>'儲存 SAVE - 審查會議-預計參加審查委員名單</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSAVE1_Click(sender As Object, e As EventArgs) Handles btnSAVE1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return
        End If

        Call SaveData1()
    End Sub

    Protected Sub btnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        SHOW_PANEL(0)
        Call ClearData1()
    End Sub

    ''' <summary> 署：全部使用者 +  分署：系統管理員才能使用【匯出名單】功能。 </summary>
    ''' <param name="iMTSEQ"></param>
    Sub Export1(ByRef iMTSEQ As Integer)
        '遴聘類別、審查委員姓名、現職服務機構、職稱、學歷、專業背景、連絡電話、手機、電子郵件、地址。 其中聯絡電話、手機、電子郵件、地址以抓取該名委員第1筆資料。
        Dim dt As DataTable = Nothing 'GET_TABLEDG3(iMTSEQ)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim sFileName1 As String = String.Concat("MEETRPT", TIMS.GetDateNo2())
        'Response.Clear()
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim sPattern As String = "遴聘類別,審查委員姓名,現職服務機構,職稱,學歷,專業背景,連絡電話,手機,電子郵件,地址"
        Dim sColumn As String = "RECRUIT_N,MBRNAME,UNITNAME,JOBTITLE,DEGREE,SPECIALTY,PHONE,CELLPHONE,EMAIL,MADDRESS"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = "<tr>"
        ExportStr &= "<td>序號</td>"
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i))
        Next
        ExportStr &= "</tr>" & vbCrLf
        sbHTML.Append(TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dt.Rows
            iNum += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", iNum)
            For i As Integer = 0 To sColumnA.Length - 1
                ExportStr &= String.Format("<td>{0}</td>", dr(sColumnA(i))) '& vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(TIMS.sUtl_AntiXss(ExportStr))
        Next
        sbHTML.Append("</div>")

        'parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        Dim parmsExp As New Hashtable From {
            {"ExpType", "EXCEL"},
            {"FileName", sFileName1},
            {"strSTYLE", strSTYLE},
            {"strHTML", sbHTML.ToString()},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

    ''' <summary>刪除</summary>
    ''' <param name="iMRSEQ"></param>
    Private Sub DELETE_MEETINGRPT(iMRSEQ As Integer)
        Dim rst As Integer = 0
        If iMRSEQ = 0 Then Return

        '查詢1筆
        Dim parms As New Hashtable From {{"MRSEQ", iMRSEQ}}
        Dim sSql As String = " SELECT 'X' FROM OA_MEETINGRPT WHERE MRSEQ=@MRSEQ" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, parms)
        If dt1.Rows.Count <> 1 Then Return

        '備份存檔
        Dim parms_up As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"MRSEQ", iMRSEQ}}
        Dim sql_up As String = "" & vbCrLf
        sql_up &= " UPDATE OA_MEETINGRPT" & vbCrLf
        sql_up &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        sql_up &= " WHERE MRSEQ=@MRSEQ" & vbCrLf
        rst = DbAccess.ExecuteNonQuery(sql_up, objconn, parms_up)

        Dim s_COL As String = "MRSEQ,DISTID,MYEARS,ORGPLANKIND,ACCEPTSTAGE,SMEETDATE,FMEETDATE,MEETPLACE,MEETADDRESS,SPEECHMAN,MHOSTER,AGENDA,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,RID"

        '備份存檔 'Dim parms As Hashtable = New Hashtable 'Dim sql As String = ""
        Dim parms_bk As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"MRSEQ", iMRSEQ}}
        Dim sql_bk As String = "" & vbCrLf
        sql_bk &= String.Concat(" INSERT INTO OA_MEETINGRPTDEL(", s_COL, ")") & vbCrLf
        sql_bk &= String.Concat(" SELECT ", s_COL, " FROM OA_MEETINGRPT") & vbCrLf
        sql_bk &= " WHERE MODIFYACCT=@MODIFYACCT AND MRSEQ=@MRSEQ" & vbCrLf
        rst = DbAccess.ExecuteNonQuery(sql_bk, objconn, parms_bk)

        '刪除 'Dim parms As Hashtable = New Hashtable 'Dim sql As String = ""
        Dim parms_d As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"MRSEQ", iMRSEQ}}
        Dim sql_d As String = "" & vbCrLf
        sql_d &= " DELETE OA_MEETINGRPT" & vbCrLf
        sql_d &= " WHERE MODIFYACCT=@MODIFYACCT AND MRSEQ=@MRSEQ" & vbCrLf
        rst = DbAccess.ExecuteNonQuery(sql_d, objconn, parms_d)
    End Sub


    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        If e Is Nothing Then Return
        If e.CommandName Is Nothing OrElse String.IsNullOrEmpty(e.CommandName) Then Return
        If e.CommandArgument Is Nothing OrElse String.IsNullOrEmpty(e.CommandArgument) Then Return
        Dim s_CmdArg As String = e.CommandArgument

        If e.CommandName = cst_UPD1 Then Call ClearData1()

        Hid_MRSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MRSEQ")
        Dim iMRSEQ As Integer = If(Hid_MRSEQ.Value <> "", Val(Hid_MRSEQ.Value), 0)
        Select Case e.CommandName
            Case cst_UPD1 '修改
                If iMRSEQ = 0 Then Return
                Call LoadData1(iMRSEQ)

            Case cst_DEL1 '刪除
                If String.IsNullOrEmpty(Hid_MRSEQ.Value) Then
                    Common.MessageBox(Me, "查無資料，不可刪除!!")
                    Return
                ElseIf iMRSEQ = 0 Then
                    Common.MessageBox(Me, "查無資料，不可刪除!!")
                    Return
                End If

                Call DELETE_MEETINGRPT(iMRSEQ)
                Dim s_msg2 As String = "資料已刪除！"
                Common.MessageBox(Me, s_msg2)
                sSearch1()
                Return

            Case cst_PRT1 '列印
                If iMRSEQ = 0 Then Return
                Dim s_ACCEPTSTAGE As String = TIMS.GetMyValue(s_CmdArg, "ACCEPTSTAGE")
                Dim s_DISTID As String = TIMS.GetMyValue(s_CmdArg, "DISTID")
                Dim s_MYEARS As String = TIMS.GetMyValue(s_CmdArg, "MYEARS")
                Dim myValue As String = ""
                myValue &= String.Concat("&MRSEQ=", iMRSEQ)
                myValue &= String.Concat("&DISTID=", s_DISTID) '"&DISTID=" & FTDate1.Text
                myValue &= String.Concat("&MYEARS=", s_MYEARS) '"&MYEARS=" & FTDate2.Text

                Select Case s_ACCEPTSTAGE
                    Case "B1", "B2"
                        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, myValue)
                    Case Else
                        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
                End Select

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim BTNUPD1 As Button = e.Item.FindControl("BTNUPD1") '修改
                Dim BTNDEL1 As Button = e.Item.FindControl("BTNDEL1") '刪除
                Dim BTNPRT1 As Button = e.Item.FindControl("BTNPRT1") '列印
                'Dim BTNEXP1 As Button = e.Item.FindControl("BTNEXP1") '匯出
                'BTNEXP1.Visible = False ' 署：全部使用者 +  分署：系統管理員才能使用【匯出名單】功能。
                BTNPRT1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID = 1 AndAlso sm.UserInfo.RoleID < 2, True, False))
                'BTNEXP1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID = 1 AndAlso sm.UserInfo.RoleID < 2, True, False))

                BTNDEL1.Attributes("onclick") = "javascript:return confirm('此動作會刪除 職類審查會，是否確定刪除?');"
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim s_CmdArg As String = ""
                TIMS.SetMyValue(s_CmdArg, "ACCEPTSTAGE", drv("ACCEPTSTAGE"))
                TIMS.SetMyValue(s_CmdArg, "MRSEQ", drv("MRSEQ"))
                TIMS.SetMyValue(s_CmdArg, "DISTID", drv("DISTID"))
                TIMS.SetMyValue(s_CmdArg, "MYEARS", drv("MYEARS"))

                BTNUPD1.CommandArgument = s_CmdArg
                BTNDEL1.CommandArgument = s_CmdArg
                TIMS.Tooltip(BTNDEL1, Convert.ToString(drv("MRSEQ")))
                BTNPRT1.CommandArgument = s_CmdArg
                'BTNEXP1.CommandArgument = s_CmdArg

                '2.	分署：僅可查詢、查看明細
                'BTNUPD1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID <> 0 AndAlso sm.UserInfo.DistID = Convert.ToString(drv("DISTID")), True, False)) '修改
                'BTNDEL1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID <> 0 AndAlso sm.UserInfo.DistID = Convert.ToString(drv("DISTID")), True, False)) '刪除
        End Select
    End Sub

End Class
