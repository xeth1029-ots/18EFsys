Partial Class TC_10_002_b
    Inherits AuthBasePage

    'TC_10_002 --OLD 2022年以前規則
    'TC_10_002_b --OLD 2023年新規則

    'OA_MEETING 會議 / OA_MEETINGDEL  會議 (刪除歷史) / OA_MEETGRADE 審查課程職類
    'OA_MEETEXAM 會議與出席場次管理
    'OA_EXAMINER 審查委員
    'OA_EXAMINERJOB 審查委員職類 (審查職類代碼)
    'ALTER TABLE OA_MEETING ALTER COLUMN [CATEGORY] [varchar](3) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL
    'ALTER TABLE OA_MEETINGDEL ALTER COLUMN [CATEGORY] [varchar](3) COLLATE Chinese_Taiwan_Stroke_CI_AS NULL

    Const cst_ADD1 As String = "ADD1" '新增
    Const cst_UPD1 As String = "UPD1" '修改
    Const cst_DEL1 As String = "DEL1" '刪除
    Const cst_EDIT3 As String = "EDIT3" '管理出席狀況/名單 'BTNEDIT3
    Const cst_VIEW4 As String = "VIEW4" '查看出席狀況/名單
    Const cst_EDIT4 As String = "EDIT4" '管理-委員計畫職類 BTNEDIT4
    Const cst_EXP1 As String = "EXP1" '匯出名單
    '<asp:Button ID="BTNEXP1" runat="server" Text="匯出名單" CommandName="EXP1" CssClass="asp_button_M" />
    '<asp:Button ID="BTNVIEW4" runat="server" Text="查看出席狀況/名單" CommandName="VIEW4" CssClass="asp_button_M" />
    'CATEGORY 1:轄區 2:跨區 審查會議類別
    'Const cst_CATEGORY_轄區 As String = "1"
    'Const cst_CATEGORY_跨區 As String = "2"
    Const cst_ORGPLANKIND_GW As String = "G,W"
    Const cst_ORGPLANKIND_G As String = "G"
    Const cst_ORGPLANKIND_W As String = "W"

    Dim dic_AGE As Dictionary(Of String, String) = TIMS.Get_ACCEPTSTAGE_DIC()
    'Dim a_vAGE1() As String = {"A1", "B1", "C1", "D1"}
    'Dim a_vAGE2() As String = {"A2", "B2", "C2", "D2"}
    Dim a_vAGE1() As String = {"A1", "B1", "C1", "D1"} '1 (初次申請)
    Dim a_vAGE2() As String = {"A2", "B2", "C2", "D2"} '2 (申復)
    'sql &= "  WHEN 'A1' THEN '上半年' WHEN 'A2' THEN '上半年申復'" & vbCrLf
    'sql &= "  WHEN 'B1' THEN '政策性' WHEN 'B2' THEN '政策性申復'" & vbCrLf
    'sql &= "  WHEN 'C1' THEN '下半年' WHEN 'C2' THEN '下半年申復'" & vbCrLf
    'sql &= "  WHEN 'D1' THEN '進階政策性' WHEN 'D2' THEN '進階政策性申復' END ACCEPTSTAGE_N" & vbCrLf

    Dim gFlag_TEST As Boolean = False '測試環境啟用

    Dim ff3 As String = ""
    Const Cst_EXAMINERpkName As String = "EMSEQ"

    'Dim BTNUPD1 As Button = e.Item.FindControl("BTNUPD1") '修改
    'Dim BTNDEL1 As Button = e.Item.FindControl("BTNDEL1") '刪除
    'Dim BTNEDIT3 As Button = e.Item.FindControl("BTNEDIT3") '管理出席狀況/名單
    'Dim BTNVIEW4 As Button = e.Item.FindControl("BTNVIEW4") '查看出席狀況/名單-分署

    Dim dtMEETEXAM As DataTable = Nothing 'OA_MEETEXAM 會議與出席場次管理
    Dim dtGCODE3 As DataTable = Nothing '審查課程職類-依dtGOVCODE3 審查職類代碼 SELECT GCODE,CNAME FROM ID_GOVCLASSCAST3 WHERE PARENTS IS NULL ORDER BY GCODE
    Dim dtGCODE3_5 As DataTable = Nothing
    Dim dtDist As DataTable = Nothing 'TIMS.Get_DistIDdt(objconn)
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

        gFlag_TEST = TIMS.sUtl_ChkTest() '測試環境啟用

        'OJT-22090601：系統-產投-會議與出席場次管理：邏輯調整
        Dim me_class_name As String = Me.GetType().BaseType.Name 'Me.GetType().Name
        If me_class_name = "TC_10_002" AndAlso sm.UserInfo.Years >= 2023 Then '登入年度2023
            Call TIMS.Utl_Redirect(Me, objconn, String.Concat("TC_10_002_b.aspx?ID=", TIMS.Get_MRqID(Me))) ' Request("ID"))
        ElseIf me_class_name = "TC_10_002_b" AndAlso sm.UserInfo.Years < 2023 Then
            Call TIMS.Utl_Redirect(Me, objconn, String.Concat("TC_10_002.aspx?ID=", TIMS.Get_MRqID(Me))) ' Request("ID"))
        End If

        '分頁設定
        PageControler1.PageDataGrid = DataGrid1
        '分署資料查詢
        dtDist = TIMS.Get_DISTIDdt(objconn) 'Dim dtDist As DataTable = TIMS.Get_DistIDdt(objconn)
        '審查職類代碼
        dtGCODE3 = TIMS.Get_GOVCODE3dt(objconn)
        dtGCODE3_5 = TIMS.GET_dtGCODE3(objconn, hid_GOVCODE3.Value)
        'Dim dt As DataTable = Get_GOVCODE3dt(oConn)
        msg1.Text = ""
        ShowButton1()

        If Not IsPostBack Then cCreate1()
    End Sub

    Sub ShowButton1()
        '署： 可使用全部功能【新增】、【修改】、【刪除】、【管理出席狀況/名單】
        '分署：'系統管理者：  可使用全部功能【新增】、【修改】、【刪除】、【管理出席狀況/名單】 (目前分署系統管理者看不到，如下圖)
        ' 其他角色： 暫不開放使用(目前功能暫不開放給其他群組用)

        '1.	署：可使用所有功能 /'2.	分署：僅可查詢、查看明細
        btnSave1.Enabled = If(sm.UserInfo.LID = 0, True, False)
        btnSave3.Enabled = If(sm.UserInfo.LID = 0, True, False)
        Button29.Enabled = If(sm.UserInfo.LID = 0, True, False)
        BtnSearch.Enabled = If(sm.UserInfo.LID = 0, True, False)
        BtnAddnew.Enabled = If(sm.UserInfo.LID = 0, True, False)

        If (sm.UserInfo.LID = 1) AndAlso (sm.UserInfo.RoleID <= 1) Then
            btnSave1.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
            btnSave3.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
            Button29.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
            BtnSearch.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
            BtnAddnew.Enabled = True 'If(sm.UserInfo.LID = 0, True, False)
        End If

        Const cst_tipmsg1 As String = "分署-系統管理者：可使用"
        If Not btnSave1.Enabled Then TIMS.Tooltip(btnSave1, cst_tipmsg1)
        If Not btnSave3.Enabled Then TIMS.Tooltip(btnSave3, cst_tipmsg1)
        If Not Button29.Enabled Then TIMS.Tooltip(Button29, cst_tipmsg1)
        If Not BtnSearch.Enabled Then TIMS.Tooltip(BtnSearch, cst_tipmsg1)
        If Not BtnAddnew.Enabled Then TIMS.Tooltip(BtnAddnew, cst_tipmsg1)
    End Sub

    Sub cCreate1()
        hid_EXAMINER_TABLE_GUID1.Value = TIMS.GetGUID()
        Session(hid_EXAMINER_TABLE_GUID1.Value) = Nothing

        PageControler1.Visible = False
        btnSave1.Attributes("onclick") = "return chkSaveData1();"

        '單筆
        ddlDISTID = TIMS.Get_DistID(ddlDISTID, dtDist) '轄區分署
        '轄區：每個轄區分署各自舉辦，各分署各自維護出席審查委員
        '跨區：多個分署共同參加(輪流主辦)，由主責轄區分署維護出席審查委員
        ddlDISTID.Enabled = True
        If (sm.UserInfo.LID <> 0) Then ddlDISTID.Enabled = False
        If Not ddlDISTID.Enabled Then TIMS.Tooltip(ddlDISTID, "分署：分署各自維護")
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '轄區分署

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

        '審查課程職類
        cblGOVCODE3 = TIMS.Get_GOVCODE3(objconn, cblGOVCODE3)
        cblGOVCODE3.Attributes("onclick") = "SelectAll('cblGOVCODE3','cblGOVCODE3Hidden');"
        '受理階段 ACCEPTSTAGE
        ddlACCEPTSTAGE = TIMS.Get_ACCEPTSTAGE(ddlACCEPTSTAGE)
        '受理階段 ACCEPTSTAGE
        ddlACCEPTSTAGE_sch = TIMS.Get_ACCEPTSTAGE(ddlACCEPTSTAGE_sch)

        '查詢 SCH
        ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, dtDist) '轄區分署
        '轄區：每個轄區分署各自舉辦，各分署各自維護出席審查委員
        '跨區：多個分署共同參加(輪流主辦)，由主責轄區分署維護出席審查委員
        ddlDISTID_SCH.Enabled = True
        If (sm.UserInfo.LID <> 0) Then ddlDISTID_SCH.Enabled = False
        If Not ddlDISTID_SCH.Enabled Then TIMS.Tooltip(ddlDISTID_SCH, "分署：分署各自維護")
        '審查課程職類
        cblGOVCODE3_sch = TIMS.Get_GOVCODE3(objconn, cblGOVCODE3_sch)
        cblGOVCODE3_sch.Attributes("onclick") = "SelectAll('cblGOVCODE3_sch','cblGOVCODE3_schHidden');"

        ddlMYEARS_SCH = TIMS.GetSyear(ddlMYEARS_SCH, iSYears, iEYears, True)  '年度
        'Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID) '轄區分署(查詢)
        Common.SetListItem(ddlMYEARS_SCH, sm.UserInfo.Years) '年度(查詢)

        SHOW_PANEL(0)
    End Sub

    ''' <summary> 查詢list </summary>
    Sub SSearch1()
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

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.MTSEQ ,a.DISTID" & vbCrLf
        sql &= " ,d1.NAME DISTNAME" & vbCrLf
        sql &= " ,a.MYEARS" & vbCrLf
        sql &= " ,a.CATEGORY" & vbCrLf '審查會議類別
        sql &= " ,CASE a.CATEGORY WHEN '1' THEN '轄區' WHEN '2' THEN '跨區' END CATEGORY_N" & vbCrLf

        sql &= " ,a.ORGPLANKIND" & vbCrLf '計畫別 G,W
        sql &= " ,CASE a.ORGPLANKIND WHEN 'G' THEN '產業人才投資計畫' WHEN 'W' THEN '提升勞工自主學習計畫' WHEN 'G,W' THEN '產業人才投資、提升勞工自主' END ORGPLANKIND_N" & vbCrLf
        '審查課程職類
        sql &= " ,dbo.FN_GET_MEETING(a.MTSEQ,'GOVCODE3') GOVCODE3" & vbCrLf '審查課程職類
        sql &= " ,dbo.FN_GET_MEETING(a.MTSEQ,'GOVCODE3_N') GOVCODE3_N" & vbCrLf '審查課程職類

        sql &= " ,a.ACCEPTSTAGE" & vbCrLf '受理階段 ACCEPTSTAGE
        sql &= " ,CASE a.ACCEPTSTAGE" & vbCrLf
        For i As Integer = 0 To a_vAGE1.Length - 1
            Dim vID1 As String = a_vAGE1(i) : Dim vNM1 As String = dic_AGE(vID1)
            Dim vID2 As String = a_vAGE2(i) : Dim vNM2 As String = dic_AGE(vID2)
            sql &= String.Format("  WHEN '{0}' THEN '{1}' WHEN '{2}' THEN '{3}'", vID1, vNM1, vID2, vNM2)
            sql &= If(i = (a_vAGE1.Length - 1), " END ACCEPTSTAGE_N", "") & vbCrLf
        Next
        'sql &= "  WHEN 'A1' THEN '上半年' WHEN 'A2' THEN '上半年申復'" & vbCrLf
        'sql &= "  WHEN 'B1' THEN '政策性' WHEN 'B2' THEN '政策性申復'" & vbCrLf
        'sql &= "  WHEN 'C1' THEN '下半年' WHEN 'C2' THEN '下半年申復'" & vbCrLf
        'sql &= "  WHEN 'D1' THEN '進階政策性' WHEN 'D2' THEN '進階政策性申復' END ACCEPTSTAGE_N" & vbCrLf

        sql &= " ,a.SMEETDATE" & vbCrLf
        sql &= " ,a.FMEETDATE" & vbCrLf
        '會議時間
        sql &= " ,CONCAT(format(a.SMEETDATE,'yyyy/MM/dd HH:mm'),'<br>~',format(a.FMEETDATE,'yyyy/MM/dd HH:mm')) SFMEETDATE_N" & vbCrLf
        sql &= " ,a.MEETPLACE" & vbCrLf
        sql &= " ,a.MODERATOR" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.OA_MEETING a" & vbCrLf
        sql &= " JOIN dbo.ID_DISTRICT d1 on d1.DISTID=a.DISTID" & vbCrLf

        sql &= " WHERE 1=1" & vbCrLf

        'Select Case sm.UserInfo.LID
        '    Case 0
        '    Case Else
        '        '分署-資料範圍只能搜到自己轄區+跨區
        '        sql &= " AND (1!=1" & vbCrLf
        '        sql &= " OR (a.CATEGORY='1' AND a.DISTID='" & sm.UserInfo.DistID & "')" & vbCrLf
        '        sql &= " OR a.CATEGORY='2'" & vbCrLf
        '        sql &= " )" & vbCrLf
        'End Select

        '轄區分署
        If v_ddlDISTID_SCH <> "" Then sql &= " AND a.DISTID='" & v_ddlDISTID_SCH & "'" & vbCrLf
        '年度
        If v_ddlMYEARS_SCH <> "" Then sql &= " AND a.MYEARS='" & v_ddlMYEARS_SCH & "'" & vbCrLf
        '會議類別 1:轄區 2:跨區 審查會議類別
        'If v_rblCATEGORY_SCH <> "" Then sql &= " AND a.CATEGORY='" & v_rblCATEGORY_SCH & "'" & vbCrLf

        '受理階段 A1:上半年 A2:上半年申復
        '受理階段 B1:政策性 B2:政策性申復
        '受理階段 C1:下半年 B2:下半年申復
        If v_ddlACCEPTSTAGE_sch <> "" Then sql &= " AND a.ACCEPTSTAGE='" & v_ddlACCEPTSTAGE_sch & "'" & vbCrLf

        '會議日期
        If SMEETDATE_sch1.Text <> "" Then sql &= " AND CONVERT(date,a.SMEETDATE)>=CONVERT(date,'" & SMEETDATE_sch1.Text & "') " & vbCrLf
        If SMEETDATE_sch2.Text <> "" Then sql &= " AND CONVERT(date,a.SMEETDATE)<=CONVERT(date,'" & SMEETDATE_sch2.Text & "') " & vbCrLf

        '會議地點
        If lk_MEETPLACE_sch <> "" Then sql &= " AND a.MEETPLACE like '" & lk_MEETPLACE_sch & "'" & vbCrLf
        '主席
        If lk_MODERATOR_sch <> "" Then sql &= " AND a.MODERATOR like '" & lk_MODERATOR_sch & "'" & vbCrLf

        msg1.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
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
        'iType 0:查詢/1:修改/3:管理/4:查詢/5:委員計畫職類
        panelView4.Visible = False
        panelEdit3.Visible = False
        panelEdit4.Visible = False '委員計畫職類 
        panelEdit.Visible = False
        panelSch.Visible = False

        Select Case iType
            Case 0
                panelSch.Visible = True
            Case 1
                panelEdit.Visible = True
            Case 3
                panelEdit3.Visible = True
            Case 4
                panelView4.Visible = True
            Case 5
                panelEdit4.Visible = True
        End Select
    End Sub

    Sub ClearData1()
        hid_GOVCODE3.Value = ""
        Hid_MTSEQ.Value = ""
        Session(hid_EXAMINER_TABLE_GUID1.Value) = Nothing
        CreateTableDG2("", 1, If(Hid_MTSEQ.Value = "", 0, Val(Hid_MTSEQ.Value)))

        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '轄區分署
        Common.SetListItem(ddlMYEARS, sm.UserInfo.Years) '年度

        'rblCATEGORY.SelectedIndex = -1 '審查會議類別
        TIMS.SetCblValue(cblORGPLANKIND, "") '計畫別 G,W

        TIMS.SetCblValue(cblGOVCODE3, "") '審查課程職類

        ddlACCEPTSTAGE.SelectedIndex = -1 '受理階段

        SMEETDATE.Text = "" '會議日期/時間-開始
        Common.SetListItem(HR1, "09") '會議日期/時間-開始
        Common.SetListItem(MM1, "00") '會議日期/時間-開始

        FMEETDATE.Text = "" '會議日期/時間-結束
        Common.SetListItem(HR2, "18") '會議日期/時間-結束
        Common.SetListItem(MM2, "00") '會議日期/時間-結束

        'FMEETDATE.TEXT = ""
        MEETPLACE.Text = "" '會議地點
        MODERATOR.Text = "" '主席 文字框，30個字元
    End Sub

    Sub LoadData1()
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then Return

        Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)
        Dim sql As String = ""
        sql = "SELECT * FROM dbo.OA_MEETING WHERE MTSEQ=@MTSEQ "
        Dim parms As New Hashtable
        parms.Add("MTSEQ", iMTSEQ)
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr1 Is Nothing Then Return

        '審查課程職類 3
        Dim v_GOVCODE3 As String = TIMS.GET_GOVCODE3(objconn, iMTSEQ)
        'Dim v_GOVCODE3 As String = GET_GOVCODE3(iMTSEQ) '審查課程職類
        TIMS.SetCblValue(cblGOVCODE3, v_GOVCODE3) '審查課程職類

        Common.SetListItem(ddlDISTID, Convert.ToString(dr1("DISTID"))) '轄區分署
        Common.SetListItem(ddlMYEARS, Convert.ToString(dr1("MYEARS"))) '年度
        'Common.SetListItem(rblCATEGORY, Convert.ToString(dr1("CATEGORY"))) '審查會議類別
        TIMS.SetCblValue(cblORGPLANKIND, Convert.ToString(dr1("ORGPLANKIND"))) '計畫別 G,W

        Common.SetListItem(ddlACCEPTSTAGE, Convert.ToString(dr1("ACCEPTSTAGE"))) '受理階段

        SMEETDATE.Text = TIMS.Cdate3(dr1("SMEETDATE")) '會議日期/時間-開始
        If SMEETDATE.Text <> "" Then TIMS.SET_DateHM(CDate(dr1("SMEETDATE")), HR1, MM1)

        FMEETDATE.Text = TIMS.Cdate3(dr1("FMEETDATE")) '會議日期/時間-結束
        'Dim v_FMEETDATE As String = TIMS.cdate3(dr1("FMEETDATE")) '會議日期/時間-結束
        If FMEETDATE.Text <> "" Then TIMS.SET_DateHM(CDate(dr1("FMEETDATE")), HR2, MM2)

        'FMEETDATE.TEXT = ""
        MEETPLACE.Text = Convert.ToString(dr1("MEETPLACE")) '會議地點
        MODERATOR.Text = Convert.ToString(dr1("MODERATOR")) '主席 文字框，30個字元

        'OA_MEETEXAM 會議與出席場次管理
        Dim sEMSEQVAL As String = TIMS.GET_OA_MEETEXAM_EMSEQVAL(objconn, iMTSEQ)
        'OA_MEETEXAM 會議與出席場次管理
        CreateTableDG2(sEMSEQVAL, 1, If(Hid_MTSEQ.Value = "", 0, Val(Hid_MTSEQ.Value)))

        SHOW_PANEL(1)
    End Sub

    ''' <summary> MTSEQ </summary>
    Sub LoadData3()
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then Return

        Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)

        '取得會議資料
        Dim dr1 As DataRow = GET_MEETING_Data(iMTSEQ)
        If dr1 Is Nothing Then Return

        'labDISTID.Text = TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(dr1("DISTID"))) '轄區分署
        labDISTID.Text = Convert.ToString(dr1("DISTNAME")) '轄區分署
        labMYEARS.Text = Convert.ToString(dr1("MYEARS")) '年度
        'labCATEGORY.Text = Convert.ToString(dr1("CATEGORY_N")) '審查會議類別
        labORGPLANKIND.Text = Convert.ToString(dr1("ORGPLANKIND_N")) '計畫別 G,W
        '審查課程職類
        labGOVCODE3.Text = Convert.ToString(dr1("GOVCODE3_N")) '審查課程職類

        labACCEPTSTAGE.Text = Convert.ToString(dr1("ACCEPTSTAGE_N")) '受理階段 ACCEPTSTAGE
        '會議日期/時間-開始 '會議日期/時間-結束
        labSFMEETDATE.Text = Convert.ToString(dr1("SFMEETDATE_N")) '會議日期/時間-開始 會議時間
        labMEETPLACE.Text = Convert.ToString(dr1("MEETPLACE")) '會議地點
        labMODERATOR.Text = Convert.ToString(dr1("MODERATOR")) '主席 文字框，30個字元

        '審查委員出席狀況
        CreateTableDG3(iMTSEQ)

        SHOW_PANEL(3)
    End Sub

    ''' <summary>
    ''' MTSEQ
    ''' </summary>
    Sub LoadData4()
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then Return

        Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)

        '取得會議資料
        Dim dr1 As DataRow = GET_MEETING_Data(iMTSEQ)
        If dr1 Is Nothing Then Return

        'labDISTID.Text = TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(dr1("DISTID"))) '轄區分署
        labDISTID4.Text = Convert.ToString(dr1("DISTNAME")) '轄區分署
        labMYEARS4.Text = Convert.ToString(dr1("MYEARS")) '年度
        'labCATEGORY4.Text = Convert.ToString(dr1("CATEGORY_N")) '審查會議類別
        labORGPLANKIND4.Text = Convert.ToString(dr1("ORGPLANKIND_N")) '計畫別 G,W
        labGOVCODE3_4.Text = Convert.ToString(dr1("GOVCODE3_N")) '審查課程職類

        labACCEPTSTAGE4.Text = Convert.ToString(dr1("ACCEPTSTAGE_N")) '受理階段 ACCEPTSTAGE
        '會議日期/時間-開始 '會議日期/時間-結束
        labSFMEETDATE4.Text = Convert.ToString(dr1("SFMEETDATE_N")) '會議日期/時間-開始
        labMEETPLACE4.Text = Convert.ToString(dr1("MEETPLACE")) '會議地點
        labMODERATOR4.Text = Convert.ToString(dr1("MODERATOR")) '主席 文字框，30個字元

        '審查委員名單 
        CreateTableDG4(iMTSEQ)

        SHOW_PANEL(4)
    End Sub

    '查詢 委員計畫職類
    Sub LoadData5()
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then Return

        Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)

        '取得會議資料
        Dim dr1 As DataRow = GET_MEETING_Data(iMTSEQ)
        If dr1 Is Nothing Then Return

        labDISTIDp4.Text = Convert.ToString(dr1("DISTNAME")) '轄區分署
        labMYEARSp4.Text = Convert.ToString(dr1("MYEARS")) '年度
        'labCATEGORYp4.Text = Convert.ToString(dr1("CATEGORY_N")) '審查會議類別
        labORGPLANKINDp4.Text = Convert.ToString(dr1("ORGPLANKIND_N")) '計畫別 G,W
        labGOVCODE3_p4.Text = Convert.ToString(dr1("GOVCODE3_N")) '審查課程職類
        hid_GOVCODE3.Value = Convert.ToString(dr1("GOVCODE3"))
        dtGCODE3_5 = TIMS.GET_dtGCODE3(objconn, hid_GOVCODE3.Value)

        labACCEPTSTAGEp4.Text = Convert.ToString(dr1("ACCEPTSTAGE_N")) '受理階段 ACCEPTSTAGE
        '會議日期/時間-開始 '會議日期/時間-結束
        labSFMEETDATEp4.Text = Convert.ToString(dr1("SFMEETDATE_N")) '會議日期/時間-開始
        labMEETPLACEp4.Text = Convert.ToString(dr1("MEETPLACE")) '會議地點
        labMODERATORp4.Text = Convert.ToString(dr1("MODERATOR")) '主席 文字框，30個字元

        '審查委員名單 
        CreateTableDG5(iMTSEQ)

        SHOW_PANEL(5)
    End Sub

    '審查委員名單 
    Private Sub CreateTableDG5(ByRef iMTSEQ As Integer)
        Dim dt2 As DataTable = GET_TABLEDG3(iMTSEQ)
        'labmsg_DG3.Text = String.Format("審查委員名單, 共有{0}筆", dt2.Rows.Count)
        If dt2.Rows.Count = 0 Then labmsg_DG3.Text = "審查委員名單, 尚未建立!"
        With DataGrid5
            .DataSource = dt2
            '.DataKeyField = Cst_EXAMINERpkName
            .DataBind()
        End With
    End Sub

    ''' <summary>取得會議資料</summary>
    ''' <param name="iMTSEQ"></param>
    ''' <returns></returns>
    Function GET_MEETING_Data(ByRef iMTSEQ As Integer) As DataRow
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " Select a.MTSEQ" & vbCrLf
        sql &= " , a.DISTID" & vbCrLf
        sql &= " , d1.NAME DISTNAME" & vbCrLf
        sql &= " , a.MYEARS" & vbCrLf
        sql &= " , a.CATEGORY" & vbCrLf '審查會議類別
        sql &= " ,CASE a.CATEGORY WHEN '1' THEN '轄區' WHEN '2' THEN '跨區' END CATEGORY_N" & vbCrLf

        sql &= " ,a.ORGPLANKIND" & vbCrLf '計畫別 G,W
        sql &= " ,CASE a.ORGPLANKIND WHEN 'G' THEN '產業人才投資計畫' WHEN 'W' THEN '提升勞工自主學習計畫' WHEN 'G,W' THEN '產業人才投資、提升勞工自主' END ORGPLANKIND_N" & vbCrLf

        sql &= " ,a.ACCEPTSTAGE" & vbCrLf '受理階段 ACCEPTSTAGE
        sql &= " ,CASE a.ACCEPTSTAGE" & vbCrLf
        For i As Integer = 0 To a_vAGE1.Length - 1
            Dim vID1 As String = a_vAGE1(i) : Dim vNM1 As String = dic_AGE(vID1)
            Dim vID2 As String = a_vAGE2(i) : Dim vNM2 As String = dic_AGE(vID2)
            sql &= String.Format("  WHEN '{0}' THEN '{1}' WHEN '{2}' THEN '{3}'", vID1, vNM1, vID2, vNM2)
            sql &= If(i = (a_vAGE1.Length - 1), " END ACCEPTSTAGE_N", "") & vbCrLf
        Next
        'sql &= " ,CASE a.ACCEPTSTAGE" & vbCrLf
        'sql &= "  WHEN 'A1' THEN '上半年' WHEN 'A2' THEN '上半年申復' " & vbCrLf
        'sql &= "  WHEN 'B1' THEN '政策性' WHEN 'B2' THEN '政策性申復' " & vbCrLf
        'sql &= "  WHEN 'C1' THEN '下半年' WHEN 'C2' THEN '下半年申復' END ACCEPTSTAGE_N" & vbCrLf
        '審查課程職類
        sql &= " ,dbo.FN_GET_MEETING(a.MTSEQ,'GOVCODE3') GOVCODE3" & vbCrLf '審查課程職類
        sql &= " ,dbo.FN_GET_MEETING(a.MTSEQ,'GOVCODE3_N') GOVCODE3_N" & vbCrLf '審查課程職類

        sql &= " ,a.SMEETDATE" & vbCrLf
        sql &= " ,a.FMEETDATE" & vbCrLf
        '會議時間
        sql &= " ,CONCAT(format(a.SMEETDATE,'yyyy/MM/dd HH:mm'),'~',format(a.FMEETDATE,'yyyy/MM/dd HH:mm')) SFMEETDATE_N" & vbCrLf
        sql &= " ,a.MEETPLACE" & vbCrLf
        sql &= " ,a.MODERATOR" & vbCrLf
        'sql &= " ,a.CREATEACCT" & vbCrLf
        'sql &= " ,a.CREATEDATE" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.OA_MEETING a" & vbCrLf
        sql &= " JOIN dbo.ID_DISTRICT d1 on d1.DISTID=a.DISTID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.MTSEQ=@MTSEQ "
        Dim parms As New Hashtable
        parms.Add("MTSEQ", iMTSEQ)
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        Return dr1
        'If dr1 Is Nothing Then Return
    End Function

    '不要再限定會議唯一值，每個年度、受理階段、19大類讓它可以開多場
    '2.	轄區會議-(5)每個年度、轄區、類別、受理階段、計畫別只能有一筆會議。 (有重複true 沒有false) 
    Function Check_Rule1(ByRef s_parms As Hashtable) As Boolean
        Dim rst As Boolean = False

        Dim parms As New Hashtable
        parms.Add("MYEARS", TIMS.GetMyValue2(s_parms, "MYEARS")) '年度
        parms.Add("DISTID", TIMS.GetMyValue2(s_parms, "DISTID")) '轄區分署
        'parms.Add("CATEGORY", TIMS.GetMyValue2(s_parms, "CATEGORY")) '審查會議類別
        parms.Add("ACCEPTSTAGE", TIMS.GetMyValue2(s_parms, "ACCEPTSTAGE")) '受理階段 ACCEPTSTAGE
        parms.Add("ORGPLANKIND", TIMS.GetMyValue2(s_parms, "ORGPLANKIND")) '計畫別 G,W
        Dim v_MTSEQ As String = TIMS.GetMyValue2(s_parms, "MTSEQ") '序號
        If v_MTSEQ <> "" Then parms.Add("MTSEQ", Val(v_MTSEQ)) '序號

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM dbo.OA_MEETING a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND MYEARS=@MYEARS" & vbCrLf
        sql &= " AND DISTID=@DISTID" & vbCrLf
        'sql &= " AND CATEGORY=@CATEGORY" & vbCrLf
        sql &= " AND ACCEPTSTAGE=@ACCEPTSTAGE" & vbCrLf
        sql &= " AND ORGPLANKIND=@ORGPLANKIND" & vbCrLf
        If v_MTSEQ <> "" Then sql &= " AND MTSEQ!=@MTSEQ" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then Return rst '沒有false
        rst = True
        Return rst '有重複true
    End Function

    '不要再限定會議唯一值，每個年度、受理階段、19大類讓它可以開多場
    '3.	跨區會議-(3)每個年度、類別、受理階段、計畫別只能有一筆會議 (有重複true 沒有false)
    Function Check_Rule2(ByRef s_parms As Hashtable) As Boolean
        Dim rst As Boolean = False

        Dim parms As New Hashtable
        parms.Add("MYEARS", TIMS.GetMyValue2(s_parms, "MYEARS"))
        'parms.Add("DISTID", TIMS.GetMyValue2(s_parms, "DISTID"))
        'parms.Add("CATEGORY", TIMS.GetMyValue2(s_parms, "CATEGORY")) '審查會議類別
        parms.Add("ACCEPTSTAGE", TIMS.GetMyValue2(s_parms, "ACCEPTSTAGE")) '受理階段
        parms.Add("ORGPLANKIND", TIMS.GetMyValue2(s_parms, "ORGPLANKIND")) '計畫別 G,W
        Dim v_MTSEQ As String = TIMS.GetMyValue2(s_parms, "MTSEQ")
        If v_MTSEQ <> "" Then parms.Add("MTSEQ", Val(v_MTSEQ))

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM dbo.OA_MEETING a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND MYEARS=@MYEARS" & vbCrLf
        'sql &= " AND DISTID=@DISTID" & vbCrLf
        'sql &= " AND CATEGORY=@CATEGORY" & vbCrLf  '審查會議類別
        sql &= " AND ACCEPTSTAGE=@ACCEPTSTAGE" & vbCrLf '受理階段 ACCEPTSTAGE
        sql &= " AND ORGPLANKIND=@ORGPLANKIND" & vbCrLf '計畫別 G,W
        If v_MTSEQ <> "" Then sql &= " AND MTSEQ!=@MTSEQ" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then Return rst '沒有false
        rst = True
        Return rst '有重複true
    End Function

    '(5)申復：就看有沒有需要，也可能不辦。上半年由哪個分署主辦，申復就一樣由該分署辦理。 (有true 沒有false)
    Function Check_Rule3(ByRef s_parms As Hashtable) As Boolean
        Dim rst As Boolean = True

        'Dim vORGPLANKIND As String = TIMS.GetMyValue2(s_parms, "ORGPLANKIND")
        Dim v_MTSEQ As String = TIMS.GetMyValue2(s_parms, "MTSEQ")

        Dim parms As New Hashtable
        parms.Add("MYEARS", TIMS.GetMyValue2(s_parms, "MYEARS")) '年度
        parms.Add("DISTID", TIMS.GetMyValue2(s_parms, "DISTID")) '轄區
        'parms.Add("CATEGORY", TIMS.GetMyValue2(s_parms, "CATEGORY")) '類別 審查會議類別
        parms.Add("ACCEPTSTAGE", TIMS.GetMyValue2(s_parms, "ACCEPTSTAGE")) '受理階段 ACCEPTSTAGE
        'parms.Add("ORGPLANKIND", vORGPLANKIND) '計畫別 G,W
        If v_MTSEQ <> "" Then parms.Add("MTSEQ", Val(v_MTSEQ))

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM dbo.OA_MEETING a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND MYEARS=@MYEARS" & vbCrLf
        sql &= " AND DISTID=@DISTID" & vbCrLf
        'sql &= " AND CATEGORY=@CATEGORY" & vbCrLf '審查會議類別
        sql &= " AND ACCEPTSTAGE=@ACCEPTSTAGE" & vbCrLf  '受理階段 ACCEPTSTAGE
        'sql &= " AND ORGPLANKIND=@ORGPLANKIND" & vbCrLf '計畫別 G,W
        If v_MTSEQ <> "" Then sql &= " AND MTSEQ!=@MTSEQ" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then Return rst '有true --沒有false

        ''申復檢核改為，判斷上半年計畫別，是否有產投或提升勞工自主任一、或是有兩計畫合併，有的話都可以提申覆
        'Dim flag_G As Boolean = False
        'parms.Clear()
        'parms.Add("MYEARS", TIMS.GetMyValue2(s_parms, "MYEARS"))
        'parms.Add("DISTID", TIMS.GetMyValue2(s_parms, "DISTID"))
        'parms.Add("CATEGORY", TIMS.GetMyValue2(s_parms, "CATEGORY"))
        'parms.Add("ACCEPTSTAGE", TIMS.GetMyValue2(s_parms, "ACCEPTSTAGE"))
        'parms.Add("ORGPLANKIND", cst_ORGPLANKIND_G)
        'If v_MTSEQ <> "" Then parms.Add("MTSEQ", Val(v_MTSEQ))
        'Dim dtG As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        'If dtG.Rows.Count > 0 Then flag_G = True '有true --沒有false
        'If flag_G Then Return rst '任1即可

        ''申復檢核改為，判斷上半年計畫別，是否有產投或提升勞工自主任一、或是有兩計畫合併，有的話都可以提申覆
        'Dim flag_W As Boolean = False
        'parms.Clear()
        'parms.Add("MYEARS", TIMS.GetMyValue2(s_parms, "MYEARS"))
        'parms.Add("DISTID", TIMS.GetMyValue2(s_parms, "DISTID"))
        'parms.Add("CATEGORY", TIMS.GetMyValue2(s_parms, "CATEGORY"))
        'parms.Add("ACCEPTSTAGE", TIMS.GetMyValue2(s_parms, "ACCEPTSTAGE"))
        'parms.Add("ORGPLANKIND", cst_ORGPLANKIND_W)
        'If v_MTSEQ <> "" Then parms.Add("MTSEQ", Val(v_MTSEQ))
        'Dim dtW As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        'If dtW.Rows.Count > 0 Then flag_W = False '有true --沒有false        
        'If flag_W Then Return rst '任1即可

        rst = False
        Return rst '沒有false
    End Function

    ''' <summary> '檢查 </summary>
    ''' <param name="s_ERRMSG">有值為異常</param>
    ''' <returns></returns>
    Function CheckData1(ByRef s_ERRMSG As String) As Boolean
        Dim rst As Boolean = True
        s_ERRMSG = ""

        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID) '轄區分署
        Dim v_ddlMYEARS As String = TIMS.GetListValue(ddlMYEARS) '年度
        'Dim v_rblCATEGORY As String = TIMS.GetListValue(rblCATEGORY) '審查會議類別
        Dim v_cblORGPLANKIND As String = TIMS.GetCblValue(cblORGPLANKIND) '計畫別 G,W
        Dim v_cblGOVCODE3 As String = TIMS.GetCblValue(cblGOVCODE3) '審查課程職類
        Dim v_ddlACCEPTSTAGE As String = TIMS.GetListValue(ddlACCEPTSTAGE) '受理階段 ACCEPTSTAGE

        Dim s_SMEETDATE As String = TIMS.GET_DateHM(SMEETDATE, HR1, MM1) '會議日期/時間-開始
        Dim s_FMEETDATE As String = TIMS.GET_DateHM(FMEETDATE, HR2, MM2) '會議日期/時間-結束 (Single)
        If TIMS.ChkDateErr3(s_SMEETDATE, s_FMEETDATE) Then
            Dim T_DATE1 As String = s_SMEETDATE
            s_SMEETDATE = s_FMEETDATE
            s_FMEETDATE = T_DATE1
        End If
        MEETPLACE.Text = TIMS.ClearSQM(MEETPLACE.Text) '會議地點
        MODERATOR.Text = TIMS.ClearSQM(MODERATOR.Text) '主席 文字框，30個字元

        If v_ddlDISTID = "" Then s_ERRMSG &= "請選擇 轄區分署" & vbCrLf
        If v_ddlMYEARS = "" Then s_ERRMSG &= "請選擇 年度" & vbCrLf
        'If v_rblCATEGORY = "" Then s_ERRMSG &= "請選擇 審查會議類別" & vbCrLf
        If v_cblORGPLANKIND = "" Then s_ERRMSG &= "請選擇 計畫別" & vbCrLf '計畫別 G,W
        If v_cblGOVCODE3 = "" Then s_ERRMSG &= "請選擇 審查課程職類" & vbCrLf '審查課程職類
        If v_ddlACCEPTSTAGE = "" Then s_ERRMSG &= "請選擇 受理階段" & vbCrLf '受理階段 ACCEPTSTAGE
        If s_SMEETDATE = "" Then s_ERRMSG &= "請選擇輸入 會議日期/時間-開始" & vbCrLf
        If s_FMEETDATE = "" Then s_ERRMSG &= "請選擇輸入 會議日期/時間-結束" & vbCrLf
        If MEETPLACE.Text = "" Then s_ERRMSG &= "請輸入 會議地點" & vbCrLf
        If MODERATOR.Text = "" Then s_ERRMSG &= "請輸入 主席" & vbCrLf
        If s_ERRMSG <> "" Then Return False

        '分署，只能選擇自已
        If sm.UserInfo.LID <> 0 AndAlso v_ddlDISTID <> sm.UserInfo.DistID Then
            s_ERRMSG &= " 轄區分署 與登入分署不同(不可儲存)" & vbCrLf
        End If
        If s_ERRMSG <> "" Then Return False

        '取得 審查課程職類 TABLE 
        '申請階段:1:上半年/2:下半年/3:政策性產業/4:進階政策性產業
        Dim v_APPSTAGE As String = TIMS.GET_APPSTAGE_12(v_ddlACCEPTSTAGE)
        '3:政策性產業/4:進階政策性產業 不檢核 
        Dim fg_CanCheck1 As Boolean = If(v_APPSTAGE = "1", True, If(v_APPSTAGE = "2", True, False))

        If fg_CanCheck1 Then
            Dim HtPP As New Hashtable
            Dim dtGCODE As DataTable = Nothing
            If v_cblGOVCODE3 <> "" Then
                '有選擇 審查課程職類
                HtPP.Clear() 'Dim HtPP As New Hashtable
                HtPP.Add("YEARS", v_ddlMYEARS)
                HtPP.Add("APPSTAGE", v_APPSTAGE)
                HtPP.Add("DISTID", v_ddlDISTID)
                HtPP.Add("cblGOVCODE3", v_cblGOVCODE3)
                dtGCODE = TIMS.GET_GCODEREVIEdt(objconn, HtPP) 'GET SYS_GCODEREVIE Dim dtGCODE As DataTable 
                If dtGCODE Is Nothing OrElse dtGCODE.Rows.Count = 0 Then
                    s_ERRMSG &= " 請選擇 年度／申請階段(受理階段)／轄區分署 有效的 審查課程職類" & vbCrLf
                    Return False
                End If
                'If s_ERRMSG <> "" Then Return False
            Else
                '沒有選擇 審查課程職類
                HtPP.Add("YEARS", v_ddlMYEARS)
                HtPP.Add("APPSTAGE", v_APPSTAGE)
                HtPP.Add("DISTID", v_ddlDISTID)
                'HtPP.Add("cblGOVCODE3", v_cblGOVCODE3)
                dtGCODE = TIMS.GET_GCODEREVIEdt(objconn, HtPP) 'GET SYS_GCODEREVIE
                If dtGCODE Is Nothing OrElse dtGCODE.Rows.Count = 0 Then
                    s_ERRMSG &= " 年度／申請階段(受理階段)／轄區分署 查無 審查課程職類" & vbCrLf
                    Return False
                End If
                'If s_ERRMSG <> "" Then Return False
            End If
        End If

        'CATEGORY 1:轄區 2:跨區 審查會議類別
        Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("MYEARS", v_ddlMYEARS) '年度
        s_parms.Add("DISTID", v_ddlDISTID) '轄區
        's_parms.Add("CATEGORY", v_rblCATEGORY) '類別 審查會議類別
        s_parms.Add("ACCEPTSTAGE", v_ddlACCEPTSTAGE) '受理階段
        s_parms.Add("ORGPLANKIND", v_cblORGPLANKIND) '計畫別 G,W
        s_parms.Add("MTSEQ", Hid_MTSEQ.Value) '有值為修改 / 無值為新增
        '不要再限定會議唯一值，每個年度、受理階段、19大類讓它可以開多場
        '轄區會議-(5)每個年度、轄區、類別、受理階段、計畫別只能有一筆會議。
        'If v_rblCATEGORY.Equals(cst_CATEGORY_轄區) Then
        '    Dim flag_CHK1 As Boolean = Check_Rule1(s_parms)
        '    If flag_CHK1 Then s_ERRMSG &= "轄區會議-每個年度、轄區、類別、受理階段、計畫別只能有一筆會議。" & vbCrLf
        'End If

        s_parms.Clear()
        s_parms.Add("MYEARS", v_ddlMYEARS) '年度
        's_parms.Add("DISTID", v_ddlDISTID)
        's_parms.Add("CATEGORY", v_rblCATEGORY) '類別 審查會議類別
        s_parms.Add("ACCEPTSTAGE", v_ddlACCEPTSTAGE) '受理階段
        s_parms.Add("ORGPLANKIND", v_cblORGPLANKIND) '計畫別 G,W
        s_parms.Add("MTSEQ", Hid_MTSEQ.Value) '有值為修改 / 無值為新增
        '不要再限定會議唯一值，每個年度、受理階段、19大類讓它可以開多場
        '跨區會議-每個年度、類別、受理階段、計畫別只能有一筆會議。
        'If v_rblCATEGORY.Equals(cst_CATEGORY_跨區) Then
        '    Dim flag_CHK2 As Boolean = Check_Rule2(s_parms)
        '    If flag_CHK2 Then s_ERRMSG &= "跨區會議-每個年度、類別、受理階段、計畫別只能有一筆會議。" & vbCrLf
        'End If

        '(5)申復：就看有沒有需要，也可能不辦。上半年由哪個分署主辦，申覆就一樣由該分署辦理。
        Dim v_ACCEPTSTAGE_LASTTIME As String = TIMS.GET_ACCEPTSTAGE_LASTTIME(v_ddlACCEPTSTAGE)
        '(申復)受理階段 取得 申請階段 中文
        Dim s_YTXT As String = TIMS.GET_YTXT(v_ACCEPTSTAGE_LASTTIME)
        '(申復)檢核
        If v_ACCEPTSTAGE_LASTTIME <> "" AndAlso s_YTXT <> "" Then
            s_parms.Clear()
            s_parms.Add("MYEARS", v_ddlMYEARS) '年度
            s_parms.Add("DISTID", v_ddlDISTID) '轄區
            's_parms.Add("CATEGORY", v_rblCATEGORY) '類別 （轄區/跨區）
            s_parms.Add("ACCEPTSTAGE", v_ACCEPTSTAGE_LASTTIME) '受理階段 ACCEPTSTAGE
            's_parms.Add("ORGPLANKIND", v_cblORGPLANKIND) '計畫別 G,W
            s_parms.Add("MTSEQ", Hid_MTSEQ.Value) '有值為修改 / 無值為新增
            Dim flag_CHK3 As Boolean = Check_Rule3(s_parms)
            If Not flag_CHK3 Then s_ERRMSG &= String.Format("{0}申復-{0}由哪個分署主辦，{0}申復就一樣由該分署辦理(查無該分署主辦資料)。", s_YTXT) & vbCrLf
        End If
        If s_ERRMSG <> "" Then Return False

        Return rst
    End Function

    ''' <summary> '儲存 SAVE - 審查會議-預計參加審查委員名單</summary>
    Sub SaveData1()
        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID) '轄區分署
        Dim v_ddlMYEARS As String = TIMS.GetListValue(ddlMYEARS) '年度
        'Dim v_rblCATEGORY As String = TIMS.GetListValue(rblCATEGORY) '審查會議類別
        Dim v_cblORGPLANKIND As String = TIMS.GetCblValue(cblORGPLANKIND) '計畫別
        Dim v_ddlACCEPTSTAGE As String = TIMS.GetListValue(ddlACCEPTSTAGE) '受理階段 ACCEPTSTAGE
        SMEETDATE.Text = TIMS.Cdate3(TIMS.ClearSQM(SMEETDATE.Text)) '會議日期/時間-開始
        FMEETDATE.Text = TIMS.Cdate3(TIMS.ClearSQM(FMEETDATE.Text)) '會議日期/時間-結束
        'Dim vFMEETDATE As String = SMEETDATE.Text ' 會議日期/時間-結束 (Single)
        MEETPLACE.Text = TIMS.ClearSQM(MEETPLACE.Text) '會議地點
        MODERATOR.Text = TIMS.ClearSQM(MODERATOR.Text) '主席 文字框，30個字元
        Dim s_SMEETDATE As String = TIMS.GET_DateHM(SMEETDATE, HR1, MM1) '會議日期/時間-開始
        Dim s_FMEETDATE As String = TIMS.GET_DateHM(FMEETDATE, HR2, MM2) '會議日期/時間-結束 (Single)

        Dim rst As Integer = 0
        Dim flagSaveOK1 As Boolean = False

        Dim i_sql As String = ""
        i_sql = "" & vbCrLf
        i_sql &= " INSERT INTO OA_MEETING( MTSEQ ,DISTID ,MYEARS ,ORGPLANKIND ,ACCEPTSTAGE ,SMEETDATE ,FMEETDATE" & vbCrLf
        i_sql &= " ,MEETPLACE ,MODERATOR,RID ,CREATEACCT ,CREATEDATE ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES ( @MTSEQ ,@DISTID ,@MYEARS ,@ORGPLANKIND ,@ACCEPTSTAGE ,@SMEETDATE ,@FMEETDATE" & vbCrLf
        i_sql &= " ,@MEETPLACE ,@MODERATOR,@RID ,@CREATEACCT ,GETDATE() ,@MODIFYACCT ,GETDATE())" & vbCrLf

        Dim u_sql As String = ""
        u_sql = "" & vbCrLf
        u_sql &= " UPDATE OA_MEETING" & vbCrLf
        u_sql &= " Set DISTID=@DISTID ,MYEARS=@MYEARS" & vbCrLf
        'u_sql &= " ,CATEGORY=@CATEGORY" & vbCrLf
        u_sql &= " ,ORGPLANKIND=@ORGPLANKIND" & vbCrLf
        u_sql &= " ,ACCEPTSTAGE=@ACCEPTSTAGE" & vbCrLf
        u_sql &= " ,SMEETDATE=@SMEETDATE" & vbCrLf
        u_sql &= " ,FMEETDATE=@FMEETDATE" & vbCrLf
        u_sql &= " ,MEETPLACE=@MEETPLACE" & vbCrLf
        u_sql &= " ,MODERATOR=@MODERATOR" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE MTSEQ=@MTSEQ" & vbCrLf

        'Dim v_ddlRECRUIT As String = TIMS.GetListValue(ddlRECRUIT)
        'Dim v_cbPUSHDISTID As String = TIMS.GetCblValue(cbPUSHDISTID)
        'Dim v_cbTRAINDISTID As String = TIMS.GetCblValue(cbTRAINDISTID)
        'Dim v_rblRUNTRAIN As String = TIMS.GetListValue(rblRUNTRAIN)
        'Dim v_rblSTOPUSE As String = TIMS.GetListValue(rblSTOPUSE)

        Dim iMTSEQ As Integer = 0
        Dim parms As Hashtable = New Hashtable
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then
            '新增
            iMTSEQ = DbAccess.GetNewId(objconn, "OA_MEETING_MTSEQ_SEQ,OA_MEETING,MTSEQ")
            parms.Clear()
            parms.Add("MTSEQ ", iMTSEQ)
            parms.Add("DISTID ", If(v_ddlDISTID <> "", v_ddlDISTID, Convert.DBNull))
            parms.Add("MYEARS ", If(v_ddlMYEARS <> "", v_ddlMYEARS, Convert.DBNull))
            'parms.Add("CATEGORY", If(v_rblCATEGORY <> "", v_rblCATEGORY, Convert.DBNull))
            parms.Add("ORGPLANKIND ", If(v_cblORGPLANKIND <> "", v_cblORGPLANKIND, Convert.DBNull))
            parms.Add("ACCEPTSTAGE ", If(v_ddlACCEPTSTAGE <> "", v_ddlACCEPTSTAGE, Convert.DBNull))
            parms.Add("SMEETDATE ", If(s_SMEETDATE <> "", s_SMEETDATE, Convert.DBNull))
            parms.Add("FMEETDATE", If(s_FMEETDATE <> "", s_FMEETDATE, Convert.DBNull))
            parms.Add("MEETPLACE ", If(MEETPLACE.Text <> "", MEETPLACE.Text, Convert.DBNull))
            parms.Add("MODERATOR", If(MODERATOR.Text <> "", MODERATOR.Text, Convert.DBNull))
            parms.Add("RID", sm.UserInfo.RID)
            parms.Add("CREATEACCT ", sm.UserInfo.UserID)
            parms.Add("MODIFYACCT ", sm.UserInfo.UserID)

            rst = DbAccess.ExecuteNonQuery(i_sql, objconn, parms)
            flagSaveOK1 = True
        Else
            '修改
            iMTSEQ = Val(Hid_MTSEQ.Value)
            parms.Clear()
            parms.Add("DISTID ", If(v_ddlDISTID <> "", v_ddlDISTID, Convert.DBNull))
            parms.Add("MYEARS ", If(v_ddlMYEARS <> "", v_ddlMYEARS, Convert.DBNull))
            'parms.Add("CATEGORY", If(v_rblCATEGORY <> "", v_rblCATEGORY, Convert.DBNull))
            parms.Add("ORGPLANKIND ", If(v_cblORGPLANKIND <> "", v_cblORGPLANKIND, Convert.DBNull))
            parms.Add("ACCEPTSTAGE ", If(v_ddlACCEPTSTAGE <> "", v_ddlACCEPTSTAGE, Convert.DBNull))
            parms.Add("SMEETDATE ", If(s_SMEETDATE <> "", s_SMEETDATE, Convert.DBNull))
            parms.Add("FMEETDATE", If(s_FMEETDATE <> "", s_FMEETDATE, Convert.DBNull))
            parms.Add("MEETPLACE ", If(MEETPLACE.Text <> "", MEETPLACE.Text, Convert.DBNull))
            parms.Add("MODERATOR", If(MODERATOR.Text <> "", MODERATOR.Text, Convert.DBNull))
            'parms.Add("RID", sm.UserInfo.RID)
            parms.Add("MODIFYACCT ", sm.UserInfo.UserID)

            parms.Add("MTSEQ ", iMTSEQ)
            rst = DbAccess.ExecuteNonQuery(u_sql, objconn, parms)
            flagSaveOK1 = True
        End If

        '申請階段:1:上半年/2:下半年/3:政策性產業/4:進階政策性產業
        Dim v_APPSTAGE As String = TIMS.GET_APPSTAGE_12(v_ddlACCEPTSTAGE)
        '3:政策性產業/4:進階政策性產業 不檢核 
        Dim fg_CanCheck1 As Boolean = If(v_APPSTAGE = "1", True, If(v_APPSTAGE = "2", True, False))

        Dim v_cblGOVCODE3 As String = TIMS.GetCblValue(cblGOVCODE3)

        '審查課程職類
        If fg_CanCheck1 Then
            '申請階段:1:上半年/2:下半年
            '取得 審查課程職類 TABLE 
            Dim HtPP As New Hashtable
            HtPP.Add("YEARS", v_ddlMYEARS)
            HtPP.Add("APPSTAGE", v_APPSTAGE)
            HtPP.Add("DISTID", v_ddlDISTID)
            HtPP.Add("cblGOVCODE3", v_cblGOVCODE3)
            Dim dtGCODE As DataTable = TIMS.GET_GCODEREVIEdt(objconn, HtPP) 'GET SYS_GCODEREVIE
            Call SaveData1C(iMTSEQ, dtGCODE)
        Else
            '3:政策性產業/4:進階政策性產業 不檢核 
            Call SaveData1D(iMTSEQ, v_cblGOVCODE3)
        End If

        '預計參加審查委員名單
        Call SaveData1B(iMTSEQ)

        If Not flagSaveOK1 Then '儲存-失敗
            Common.MessageBox(Me, "儲存失敗!")
            Exit Sub
        End If

        SHOW_PANEL(0)
        '儲存成功 'Hid_EMSEQ.Value = ""
        Call ClearData1()
        Common.MessageBox(Me, "儲存成功!")
        Call SSearch1()
    End Sub

    ''' <summary>SAVE - 會議與出席場次管理-預計參加審查委員名單</summary>
    ''' <param name="iMTSEQ"></param>
    Sub SaveData1B(ByRef iMTSEQ As Integer)
        Dim sEMSEQVAL As String = ""
        For Each eItem As DataGridItem In DataGrid2.Items
            Dim drv As DataRowView = eItem.DataItem
            Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ")
            Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
            Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))
            'Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
            If sEMSEQVAL <> "" Then sEMSEQVAL &= ","
            sEMSEQVAL &= Hid_EMSEQ.Value
        Next
        'OA_MEETEXAM 會議與出席場次管理
        Dim d_parms As Hashtable = New Hashtable 'd_parms.Clear()
        d_parms.Add("MTSEQ", iMTSEQ)
        Dim d_sql As String = ""
        d_sql = "" & vbCrLf
        d_sql &= " DELETE OA_MEETEXAM" & vbCrLf
        d_sql &= " WHERE 1=1" & vbCrLf
        d_sql &= " And MTSEQ=@MTSEQ" & vbCrLf
        If sEMSEQVAL <> "" Then d_sql &= String.Concat(" AND EMSEQ NOT IN (", sEMSEQVAL, ")")
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)

        Dim s_parms As Hashtable = New Hashtable
        Dim s_sql As String = ""
        s_sql = " Select 'X' FROM OA_MEETEXAM WHERE 1=1 AND MTSEQ=@MTSEQ AND EMSEQ=@EMSEQ" & vbCrLf
        'Dim vMTSEQ
        Dim i_parms As Hashtable = New Hashtable
        Dim i_sql As String = ""
        i_sql = "" & vbCrLf
        i_sql &= " INSERT INTO OA_MEETEXAM(MTSEQ ,EMSEQ ,RID ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@MTSEQ ,@EMSEQ ,@RID ,@MODIFYACCT ,GETDATE())" & vbCrLf

        For Each eItem As DataGridItem In DataGrid2.Items
            Dim drv As DataRowView = eItem.DataItem
            Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ")
            Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))
            'Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))

            s_parms.Clear()
            s_parms.Add("MTSEQ", iMTSEQ)
            s_parms.Add("EMSEQ", iEMSEQ)
            Dim s_dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, s_parms)
            If s_dt.Rows.Count = 0 Then
                i_parms.Clear()
                i_parms.Add("MTSEQ", iMTSEQ)
                i_parms.Add("EMSEQ", iEMSEQ)
                i_parms.Add("RID", sm.UserInfo.RID)
                i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
            End If
        Next
    End Sub

    ''' <summary> SAVE - 會議與出席場次管理-審查課程職類</summary>
    ''' <param name="iMTSEQ"></param>
    ''' <param name="dtGCODE"></param>
    Sub SaveData1C(ByRef iMTSEQ As Integer, ByRef dtGCODE As DataTable)
        'OA_MEETGRADE 審查課程職類
        Dim d_parms As Hashtable = New Hashtable
        d_parms.Clear()
        d_parms.Add("MTSEQ", iMTSEQ)
        Dim d_sql As String = " DELETE OA_MEETGRADE WHERE 1=1 AND MTSEQ=@MTSEQ" & vbCrLf 'sql &= " AND SGRID=@SGRID" & vbCrLf
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)

        Dim s_parms As Hashtable = New Hashtable
        Dim s_sql As String = ""
        's_sql = " SELECT 1 FROM OA_MEETGRADE WHERE 1=1 AND MTSEQ=@MTSEQ AND SGRID=@SGRID" & vbCrLf
        s_sql = " SELECT 1 FROM OA_MEETGRADE WHERE 1=1 AND MTSEQ=@MTSEQ AND GCODE=@GCODE" & vbCrLf

        Dim i_parms As Hashtable = New Hashtable
        Dim i_sql As String = ""
        i_sql = "" & vbCrLf
        i_sql &= " INSERT INTO OA_MEETGRADE(MTSEQ ,GCODE,SGRID ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@MTSEQ ,@GCODE,@SGRID ,@MODIFYACCT ,GETDATE())" & vbCrLf

        For Each dr1 As DataRow In dtGCODE.Rows
            'Dim iMTSEQ As Integer xx
            Dim v_GCODE As String = Convert.ToString(dr1("GCODE"))
            Dim iSGRID As Integer = Val(dr1("SGRID"))
            'Dim s_parms As New Hashtable
            s_parms.Clear()
            s_parms.Add("MTSEQ", iMTSEQ)
            s_parms.Add("GCODE", v_GCODE)
            Dim s_dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, s_parms)
            If s_dt.Rows.Count = 0 Then
                'Dim i_parms As New Hashtable
                i_parms.Clear()
                i_parms.Add("MTSEQ", iMTSEQ)
                i_parms.Add("GCODE", v_GCODE)
                i_parms.Add("SGRID", iSGRID)
                i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
            End If
        Next
    End Sub

    ''' <summary>SAVE - 會議與出席場次管理-審查課程職類</summary>
    ''' <param name="iMTSEQ"></param>
    ''' <param name="v_cblGOVCODE3"></param>
    Private Sub SaveData1D(iMTSEQ As Integer, v_cblGOVCODE3 As String)
        'OA_MEETGRADE 審查課程職類
        Dim d_parms As Hashtable = New Hashtable
        d_parms.Clear()
        d_parms.Add("MTSEQ", iMTSEQ)
        Dim d_sql As String = " DELETE OA_MEETGRADE WHERE 1=1 AND MTSEQ=@MTSEQ" & vbCrLf 'sql &= " AND SGRID=@SGRID" & vbCrLf
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)

        Dim s_parms As Hashtable = New Hashtable
        Dim s_sql As String = ""
        s_sql = " SELECT 1 FROM OA_MEETGRADE WHERE 1=1 AND MTSEQ=@MTSEQ AND GCODE=@GCODE" & vbCrLf

        Dim i_parms As Hashtable = New Hashtable
        Dim i_sql As String = ""
        i_sql = "" & vbCrLf
        i_sql &= " INSERT INTO OA_MEETGRADE(MTSEQ,GCODE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@MTSEQ,@GCODE,@MODIFYACCT,GETDATE())" & vbCrLf

        For Each s_GCODE As String In v_cblGOVCODE3.Split(",")
            s_GCODE = TIMS.ClearSQM(s_GCODE)
            s_parms.Clear()
            s_parms.Add("MTSEQ", iMTSEQ)
            s_parms.Add("GCODE", s_GCODE)
            Dim s_dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, s_parms)
            If s_dt.Rows.Count = 0 Then
                'Dim i_parms As New Hashtable
                i_parms.Clear()
                i_parms.Add("MTSEQ", iMTSEQ)
                i_parms.Add("GCODE", s_GCODE)
                i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
            End If
        Next
    End Sub

    ''' <summary>SAVE-審查會議-審查委員出席狀況</summary>
    ''' <param name="iMTSEQ"></param>
    Sub SaveData3(ByRef iMTSEQ As Integer)
        '	儲存時間需>會議日期之起始時間， 如2021/3/8 09:00
        Dim Errmsg As String = ""
        Dim flag_OK As Boolean = CHK_SAVEDATA1(iMTSEQ)
        If Not flag_OK Then
            Errmsg &= "儲存(現在)時間需 大於 會議日期之起始時間！"
            Common.MessageBox(Me, Errmsg)
            Return
        End If

        For Each eItem As DataGridItem In DataGrid3.Items
            Dim drv As DataRowView = eItem.DataItem

            Dim cbATTEND As HtmlInputCheckBox = eItem.FindControl("cbATTEND") '出席(勾選框)
            Dim cbNOTINABS As HtmlInputCheckBox = eItem.FindControl("cbNOTINABS") '不列入缺席
            Dim REASON As TextBox = eItem.FindControl("REASON") '不列入缺席原因
            Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ")
            Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))
            'Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
            'REASON.Text = TIMS.ClearSQM2(REASON.Text)

            If (cbATTEND.Checked AndAlso cbNOTINABS.Checked) Then
                Errmsg &= "出席與不列入缺席，不可同時勾選！" & vbCrLf
                Exit For
            End If
        Next
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return
        End If

        Dim s_parms As Hashtable = New Hashtable
        Dim s_sql As String = ""
        s_sql = " SELECT 'X' FROM OA_MEETEXAM WHERE 1=1 AND MTSEQ=@MTSEQ AND EMSEQ=@EMSEQ" & vbCrLf
        'Dim vMTSEQ
        Dim u_parms As Hashtable = New Hashtable
        Dim u_sql As String = ""
        u_sql = "" & vbCrLf
        u_sql &= " UPDATE OA_MEETEXAM" & vbCrLf
        u_sql &= " SET ATTEND=@ATTEND" & vbCrLf
        u_sql &= " ,NOTINABS=@NOTINABS" & vbCrLf
        u_sql &= " ,REASON=@REASON" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE 1=1" & vbCrLf
        u_sql &= " AND MTSEQ=@MTSEQ" & vbCrLf
        u_sql &= " AND EMSEQ=@EMSEQ" & vbCrLf
        For Each eItem As DataGridItem In DataGrid3.Items
            Dim drv As DataRowView = eItem.DataItem

            Dim cbATTEND As HtmlInputCheckBox = eItem.FindControl("cbATTEND") '出席(勾選框)
            Dim cbNOTINABS As HtmlInputCheckBox = eItem.FindControl("cbNOTINABS") '不列入缺席
            Dim REASON As TextBox = eItem.FindControl("REASON") '不列入缺席原因
            Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ")
            Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))
            'Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
            REASON.Text = TIMS.ClearSQM2(REASON.Text)

            s_parms.Clear()
            s_parms.Add("MTSEQ", iMTSEQ)
            s_parms.Add("EMSEQ", iEMSEQ)
            Dim s_dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, s_parms)
            If s_dt.Rows.Count > 0 Then
                u_parms.Clear()
                u_parms.Add("ATTEND", If(cbATTEND.Checked, "Y", Convert.DBNull))
                u_parms.Add("NOTINABS", If(cbNOTINABS.Checked, "Y", Convert.DBNull))
                u_parms.Add("REASON", If(REASON.Text <> "", REASON.Text, Convert.DBNull))
                u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)

                u_parms.Add("MTSEQ", iMTSEQ)
                u_parms.Add("EMSEQ", iEMSEQ)
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
            End If
        Next

        SHOW_PANEL(0)
        '儲存成功 'Hid_EMSEQ.Value = ""
        Call ClearData1()
        Common.MessageBox(Me, "儲存成功!")
        Call SSearch1()
    End Sub

    ''' <summary>'查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        SSearch1()
    End Sub

    ''' <summary>'新增</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnAddnew_Click(sender As Object, e As EventArgs) Handles BtnAddnew.Click
        Call ClearData1()

        '北分署新增了一筆110年、跨區、上半年之審查會議，則以北分署為主責分署，其他分署不可再新增，除非北分署刪掉該筆會議資料。(理論上非主責分署不會去新增)
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '轄區分署
        ddlDISTID.Enabled = If(sm.UserInfo.LID = 0, True, False)

        SHOW_PANEL(1)
    End Sub

    ''' <summary>'儲存 SAVE - 審查會議-預計參加審查委員名單</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return
        End If

        Call SaveData1()
    End Sub

    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        SHOW_PANEL(0)
        Call ClearData1()
    End Sub

    ''' <summary>檢核使用狀況 true:使用中 false:無使用</summary>
    ''' <param name="iMTSEQ"></param>
    ''' <returns></returns>
    Function CHECK_MEETEXAM(ByRef iMTSEQ As Integer) As Boolean
        Dim rst As Boolean = True
        If iMTSEQ = 0 Then Return rst
        If dtMEETEXAM IsNot Nothing Then dtMEETEXAM.Clear()
        'OA_MEETEXAM 會議與出席場次管理
        'ff3 = String.Format("EMSEQ={0} AND ATTEND='Y'", Convert.ToString(drv("EMSEQ")))
        Dim sql As String = ""
        sql = " SELECT EMSEQ,ATTEND FROM OA_MEETEXAM WHERE MTSEQ=@MTSEQ" & vbCrLf
        dtMEETEXAM = New DataTable

        TIMS.OpenDbConn(objconn)
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("MTSEQ", SqlDbType.Int).Value = iMTSEQ
            dtMEETEXAM.Load(.ExecuteReader())
        End With
        If dtMEETEXAM.Rows.Count > 0 Then Return rst
        Return False
    End Function

    '刪除
    Sub DELETE_MEETING(ByRef iMTSEQ As Integer)
        Dim rst As Integer = 0
        If iMTSEQ = 0 Then Return

        '查詢1筆
        Dim parms As Hashtable = New Hashtable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM OA_MEETING" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND MTSEQ=@MTSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MTSEQ", iMTSEQ)
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt1.Rows.Count <> 1 Then Return

        'Dim parms As Hashtable = New Hashtable
        'Dim sql As String = ""
        '備份存檔
        Dim sql_up As String = "" & vbCrLf
        sql_up &= " UPDATE OA_MEETING" & vbCrLf
        sql_up &= " SET MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql_up &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql_up &= " WHERE 1=1" & vbCrLf
        sql_up &= " AND MTSEQ=@MTSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("MTSEQ", iMTSEQ)
        rst = DbAccess.ExecuteNonQuery(sql_up, objconn, parms)

        Dim s_COL As String = "MTSEQ,DISTID,MYEARS,ORGPLANKIND,ACCEPTSTAGE,SMEETDATE,FMEETDATE,MEETPLACE,MODERATOR,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,RID"

        '備份存檔 'Dim parms As Hashtable = New Hashtable 'Dim sql As String = ""
        Dim sql_bk As String = "" & vbCrLf
        sql_bk &= String.Concat(" INSERT INTO OA_MEETINGDEL(", s_COL, ")") & vbCrLf
        sql_bk &= String.Concat(" SELECT ", s_COL, " FROM OA_MEETING") & vbCrLf
        sql_bk &= " WHERE 1=1" & vbCrLf
        sql_bk &= " AND MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql_bk &= " AND MTSEQ=@MTSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("MTSEQ", iMTSEQ)
        rst = DbAccess.ExecuteNonQuery(sql_bk, objconn, parms)

        '刪除 'Dim parms As Hashtable = New Hashtable 'Dim sql As String = ""
        Dim sql_d As String = "" & vbCrLf
        sql_d &= " DELETE OA_MEETING" & vbCrLf
        sql_d &= " WHERE 1=1" & vbCrLf
        sql_d &= " AND MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql_d &= " AND MTSEQ=@MTSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("MTSEQ", iMTSEQ)
        rst = DbAccess.ExecuteNonQuery(sql_d, objconn, parms)
    End Sub


    ''' <summary> 署：全部使用者 +  分署：系統管理員才能使用【匯出名單】功能。 </summary>
    ''' <param name="iMTSEQ"></param>
    Sub Export1(ByRef iMTSEQ As Integer)
        '遴聘類別、審查委員姓名、現職服務機構、職稱、學歷、專業背景、連絡電話、手機、電子郵件、地址。 其中聯絡電話、手機、電子郵件、地址以抓取該名委員第1筆資料。
        Dim dt As DataTable = GET_TABLEDG3(iMTSEQ)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim sFileName1 As String = "EXAMINER" & TIMS.GetDateNo2()
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

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Call TIMS.OpenDbConn(objconn)
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        If e Is Nothing Then Return
        If e.CommandName Is Nothing Then Return
        If e.CommandName = "" Then Return
        If e.CommandArgument Is Nothing Then Return
        If e.CommandArgument = "" Then Return
        Dim s_CmdArg As String = e.CommandArgument

        'Const cst_UPD1 As String = "UPD1" '修改
        'Const cst_DEL1 As String = "DEL1" '刪除
        'Const cst_EDIT3 As String = "EDIT3" '管理出席狀況/名單
        'Const cst_VIEW1 As String = "VIEW1" '查看出席狀況/名單 
        'MTSEQ
        Select Case e.CommandName
            Case cst_UPD1 '修改
                Call ClearData1()
                Hid_MTSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MTSEQ")
                Call LoadData1()

            Case cst_DEL1 '刪除
                Hid_MTSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MTSEQ")
                Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)
                Dim flag_MEETEXAM As Boolean = False
                If iMTSEQ > 0 Then flag_MEETEXAM = CHECK_MEETEXAM(iMTSEQ)
                If flag_MEETEXAM Then
                    Common.MessageBox(Me, "使用中，不可刪除!!")
                    Return
                End If
                If iMTSEQ = 0 Then
                    Common.MessageBox(Me, "查無資料，不可刪除!!")
                    Return
                End If

                Call DELETE_MEETING(iMTSEQ)
                Dim s_msg2 As String = "資料已刪除！"
                Common.MessageBox(Me, s_msg2)
                SSearch1()
                Return

            Case cst_EDIT4 '審查委員計畫職類
                Call ClearData1()
                Hid_MTSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MTSEQ")
                Call LoadData5()

            Case cst_EDIT3 '管理出席狀況/名單
                Call ClearData1()
                Hid_MTSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MTSEQ")
                Call LoadData3()

            Case cst_VIEW4 '查看出席狀況/名單
                Call ClearData1()
                Hid_MTSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MTSEQ")
                Call LoadData4()

            Case cst_EXP1 '匯出名單
                Hid_MTSEQ.Value = TIMS.GetMyValue(s_CmdArg, "MTSEQ")
                If Hid_MTSEQ.Value = "" Then Return
                Call Export1(Val(Hid_MTSEQ.Value))
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim BTNUPD1 As Button = e.Item.FindControl("BTNUPD1") '修改
                Dim BTNDEL1 As Button = e.Item.FindControl("BTNDEL1") '刪除
                Dim BTNEDIT4 As Button = e.Item.FindControl("BTNEDIT4") '委員計畫職類
                Dim BTNEDIT3 As Button = e.Item.FindControl("BTNEDIT3") '管理出席狀況/名單
                Dim BTNVIEW4 As Button = e.Item.FindControl("BTNVIEW4") '查看出席狀況/名單-分署
                Dim BTNEXP1 As Button = e.Item.FindControl("BTNEXP1") '匯出名單
                'BTNEXP1.Visible = False ' 署：全部使用者 +  分署：系統管理員才能使用【匯出名單】功能。
                BTNEXP1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID = 1 AndAlso sm.UserInfo.RoleID < 2, True, False))

                BTNDEL1.Attributes("onclick") = "javascript:return confirm('此動作會刪除會議與出席場次管理，是否確定刪除?');"
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim s_CmdArg As String = ""
                TIMS.SetMyValue(s_CmdArg, "MTSEQ", drv("MTSEQ"))

                BTNUPD1.CommandArgument = s_CmdArg
                BTNDEL1.CommandArgument = s_CmdArg
                'TIMS.Tooltip(BTNDEL1, Convert.ToString(drv("MTSEQ")))
                BTNEDIT4.CommandArgument = s_CmdArg
                BTNEDIT3.CommandArgument = s_CmdArg
                BTNVIEW4.CommandArgument = s_CmdArg
                BTNEXP1.CommandArgument = s_CmdArg

                '2.	分署：僅可查詢、查看明細
                BTNUPD1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID <> 0 AndAlso sm.UserInfo.DistID = Convert.ToString(drv("DISTID")), True, False)) '修改
                BTNDEL1.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID <> 0 AndAlso sm.UserInfo.DistID = Convert.ToString(drv("DISTID")), True, False)) '刪除
                '管理出席狀況/名單
                '(4)各轄區會議的委員名單，由各分署各自維護，每個分署只能看到自己分署的會議名單。
                '管理出席狀況/名單
                '(6)跨區會議的委員名單，由主責分署維護，其他分署只可查看，不可修改。
                BTNEDIT3.Visible = If(sm.UserInfo.LID = 0, True, If(sm.UserInfo.LID <> 0 AndAlso sm.UserInfo.DistID = Convert.ToString(drv("DISTID")), True, False)) '管理出席狀況/名單
                BTNEDIT4.Visible = BTNEDIT3.Visible

                'A.轄區會議 ：各分署僅可看到自己轄區分署的會議資料。可使用修改、刪除、管理出席狀況/名單按鈕。
                'B.跨區會議 ：只有主責分署(即當初新增的分署)可使用修改、刪除、管理出席狀況/名單按鈕，其餘分署可使用查看委員名單按鈕。
                If Not gFlag_TEST Then
                    '(4)各轄區會議的委員名單，由各分署各自維護，每個分署只能看到自己分署的會議名單。
                    '管理出席狀況/名單
                    '(6)跨區會議的委員名單，由主責分署維護，其他分署只可查看，不可修改。
                    'BTNVIEW4.Visible = If(sm.UserInfo.LID = 0, False, True) '查看出席狀況/名單-分署
                    BTNVIEW4.Visible = False
                    If sm.UserInfo.LID <> 0 AndAlso sm.UserInfo.DistID <> Convert.ToString(drv("DISTID")) Then BTNVIEW4.Visible = True
                End If

        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        If Session(hid_EXAMINER_TABLE_GUID1.Value) Is Nothing Then Exit Sub
        Dim dt As DataTable = Session(hid_EXAMINER_TABLE_GUID1.Value) '取得SESSION到 dt

        Select Case e.CommandName
            Case cst_DEL1 '"DEL1"
                Dim DGobj As DataGrid = DataGrid2
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = "查無可刪除資料！"
                    Exit Sub
                End If
                '出席狀況

                ff3 = String.Format("{0}={1}", Cst_EXAMINERpkName, DGobj.DataKeys(e.Item.ItemIndex))
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(ff3).Length <> 0 Then
                    For Each dr As DataRow In dt.Select(ff3)
                        If dr.RowState <> DataRowState.Deleted Then
                            dr.Delete() '刪除
                            Exit For
                        End If
                    Next
                End If

                Session(hid_EXAMINER_TABLE_GUID1.Value) = dt
                DGobj.DataSource = dt
                DGobj.DataBind()
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labPUSHDISTID_N As Label = e.Item.FindControl("labPUSHDISTID_N")
                labPUSHDISTID_N.Text = TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(drv("PUSHDISTID")))
                Dim Hid_EMSEQ As HiddenField = e.Item.FindControl("Hid_EMSEQ")
                Dim BTNDEL1 As Button = e.Item.FindControl("BTNDEL1") '刪除
                BTNDEL1.Attributes("onclick") = "javascript:return confirm('此動作會刪除資料，是否確定刪除?');"
                Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim s_CmdArg As String = ""
                TIMS.SetMyValue(s_CmdArg, "EMSEQ", drv("EMSEQ"))

                BTNDEL1.CommandArgument = s_CmdArg
                BTNDEL1.Visible = If(sm.UserInfo.LID <> 0, False, True)

                If Not dtMEETEXAM Is Nothing Then
                    ff3 = String.Format("EMSEQ={0} AND ATTEND='Y'", Convert.ToString(drv("EMSEQ")))
                    If dtMEETEXAM.Select(ff3).Length > 0 Then
                        BTNDEL1.Enabled = False
                        TIMS.Tooltip(BTNDEL1, "有出席資料！", True)
                    End If
                End If
        End Select
    End Sub

    ''' <summary> CreateTableDG2 </summary>
    ''' <param name="sEMSEQVAL"></param>
    ''' <param name="iType">iType 1:新增 2:修改</param>
    Sub CreateTableDG2(ByVal sEMSEQVAL As String, ByVal iType As Integer, ByRef iMTSEQ As Integer)
        Dim sql As String = ""
        Dim flag_MEETEXAM As Boolean = False
        If iMTSEQ > 0 Then flag_MEETEXAM = CHECK_MEETEXAM(iMTSEQ)

        'iType 1:新增 2:修改
        If iType = 2 Then
            For Each eItem As DataGridItem In DataGrid2.Items
                Dim drv As DataRowView = eItem.DataItem
                Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ")
                Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
                'Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))
                'Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
                sEMSEQVAL &= String.Concat(If(sEMSEQVAL <> "", ",", ""), Hid_EMSEQ.Value)
            Next
        End If

        Dim dt As DataTable = Nothing
        'Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT r.EMSEQ" & vbCrLf '審查委員
        sql &= " ,r.RECRUIT" & vbCrLf
        sql &= " ,CASE r.RECRUIT WHEN 'A' THEN 'A-產業界' WHEN 'B' THEN 'B-學術界' WHEN 'C' THEN 'C-勞工團體代表' END RECRUIT_N" & vbCrLf
        sql &= " ,r.UNITNAME" & vbCrLf
        sql &= " ,r.MBRNAME" & vbCrLf
        sql &= " ,r.JOBTITLE" & vbCrLf
        sql &= " ,r.PUSHDISTID" & vbCrLf
        sql &= " FROM dbo.OA_EXAMINER r" & vbCrLf '審查委員 
        sql &= " WHERE 1=1" & vbCrLf
        sql &= If(sEMSEQVAL <> "", String.Concat(" AND r.EMSEQ in (", sEMSEQVAL, ")"), " AND 1<>1")
        '順序排序： 1.先依遴聘類別 2.依姓名筆劃 chinese_taiwan_stroke_cs_as_ks_ws CHINESE_TAIWAN_STROKE_CI_AS
        sql &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE CHINESE_TAIWAN_STROKE_CI_AS" & vbCrLf
        'sql2 &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE Chinese_PRC_Stroke_ci_as " & vbCrLf

        dt = DbAccess.GetDataTable(sql, objconn)
        dt.Columns(Cst_EXAMINERpkName).AutoIncrement = True
        dt.Columns(Cst_EXAMINERpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_EXAMINERpkName).AutoIncrementStep = -1
        Session(hid_EXAMINER_TABLE_GUID1.Value) = dt
        With DataGrid2
            .DataSource = dt
            .DataKeyField = Cst_EXAMINERpkName
            .DataBind()
        End With
    End Sub

    ''' <summary> 審查委員出席狀況 ／ 會議與出席場次管理 OA_MEETEXAM</summary>
    ''' <param name="iMTSEQ"></param>
    ''' <returns></returns>
    Function GET_TABLEDG3(ByRef iMTSEQ As Integer) As DataTable
        Dim dt2 As DataTable = Nothing
        'Dim sPattern As String = "遴聘類別,審查委員姓名,現職服務機構,職稱,學歷,專業背景,連絡電話,手機,電子郵件,地址"
        'Dim sColumn As String = "RECRUIT_N,MBRNAME,UNITNAME,JOBTITLE,DEGREE,SPECIALTY,PHONE,CELLPHONE,EMAIL,MADDRESS"
        Dim sql2 As String = ""
        sql2 = "" & vbCrLf
        sql2 &= " SELECT m.MTSEQ" & vbCrLf '(會議序號)
        sql2 &= " ,m.EMSEQ" & vbCrLf '(審查委員序號)
        sql2 &= " ,m.ATTEND" & vbCrLf
        sql2 &= " ,m.NOTINABS" & vbCrLf
        sql2 &= " ,m.REASON" & vbCrLf
        sql2 &= " ,r.RECRUIT" & vbCrLf
        sql2 &= " ,CASE r.RECRUIT WHEN 'A' THEN 'A-產業界' WHEN 'B' THEN 'B-學術界' WHEN 'C' THEN 'C-勞工團體代表' END RECRUIT_N" & vbCrLf '遴聘類別
        sql2 &= " ,r.UNITNAME" & vbCrLf '現職服務機構
        sql2 &= " ,r.MBRNAME" & vbCrLf '審查委員姓名
        sql2 &= " ,r.JOBTITLE" & vbCrLf '職稱
        sql2 &= " ,r.PUSHDISTID" & vbCrLf
        sql2 &= " ,r.DEGREE" & vbCrLf '學歷
        sql2 &= " ,r.SPECIALTY" & vbCrLf '專業背景
        sql2 &= " ,r.PHONE" & vbCrLf '連絡電話
        sql2 &= " ,r.CELLPHONE" & vbCrLf '手機
        sql2 &= " ,r.CONFAX" & vbCrLf
        sql2 &= " ,r.EMAIL" & vbCrLf '電子郵件
        'sql2 &= " ,r.MZIPCODE,r.MZIPCODE2W
        sql2 &= " ,r.MADDRESS" & vbCrLf '地址
        sql2 &= " ,m.ORGKINDG" & vbCrLf '產業人才投資計畫
        sql2 &= " ,m.ORGKINDW" & vbCrLf '提升勞工自主學習計畫
        sql2 &= " ,m.GCODE300" & vbCrLf '審查課程職類-依dtGOVCODE3

        sql2 &= " FROM dbo.OA_MEETEXAM m" & vbCrLf 'OA_MEETEXAM 會議與出席場次管理
        sql2 &= " JOIN dbo.OA_EXAMINER r ON r.EMSEQ=m.EMSEQ" & vbCrLf
        sql2 &= " WHERE 1=1" & vbCrLf
        sql2 &= " AND m.MTSEQ=@MTSEQ" & vbCrLf
        '順序排序： 1.先依遴聘類別 2.依姓名筆劃 chinese_taiwan_stroke_cs_as_ks_ws CHINESE_TAIWAN_STROKE_CI_AS
        sql2 &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE CHINESE_TAIWAN_STROKE_CI_AS" & vbCrLf
        'sql2 &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE Chinese_PRC_Stroke_ci_as " & vbCrLf
        Dim parms2 As Hashtable = New Hashtable
        parms2.Add("MTSEQ", iMTSEQ)
        dt2 = DbAccess.GetDataTable(sql2, objconn, parms2)
        Return dt2
    End Function

    ''' <summary> 審查委員出席狀況 </summary>
    ''' <param name="iMTSEQ"></param>
    Sub CreateTableDG3(ByRef iMTSEQ As Integer)
        'iType 1:新增 2:修改
        'If iType = 2 Then
        '    For Each eItem As DataGridItem In DataGrid2.Items
        '        Dim drv As DataRowView = eItem.DataItem
        '        Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ")
        '        Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
        '        'Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))
        '        'Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
        '        If sEMSEQVAL <> "" Then sEMSEQVAL &= ","
        '        sEMSEQVAL &= Hid_EMSEQ.Value
        '    Next
        'End If

        Dim dt2 As DataTable = GET_TABLEDG3(iMTSEQ)
        'labmsg_DG3.Text = String.Format("審查委員名單, 共有{0}筆", dt2.Rows.Count)
        If dt2.Rows.Count = 0 Then labmsg_DG3.Text = "審查委員名單, 尚未建立!"
        With DataGrid3
            .DataSource = dt2
            '.DataKeyField = Cst_EXAMINERpkName
            .DataBind()
        End With
    End Sub

    ''' <summary>審查委員名單</summary>
    ''' <param name="iMTSEQ"></param>
    Sub CreateTableDG4(ByRef iMTSEQ As Integer)
        Dim dt2 As DataTable = Nothing
        Dim sql2 As String = ""
        sql2 = "" & vbCrLf
        sql2 &= " SELECT m.MTSEQ" & vbCrLf
        sql2 &= " ,m.EMSEQ" & vbCrLf
        sql2 &= " ,m.ATTEND" & vbCrLf
        sql2 &= " ,m.NOTINABS" & vbCrLf
        sql2 &= " ,m.REASON" & vbCrLf
        'sql2 &= " ,m.MODIFYACCT" & vbCrLf
        'sql2 &= " ,m.MODIFYDATE" & vbCrLf
        'sql &= " SELECT r.EMSEQ" & vbCrLf
        sql2 &= " ,r.RECRUIT" & vbCrLf
        sql2 &= " ,CASE r.RECRUIT WHEN 'A' THEN 'A-產業界' WHEN 'B' THEN 'B-學術界' WHEN 'C' THEN 'C-勞工團體代表' END RECRUIT_N" & vbCrLf
        sql2 &= " ,r.UNITNAME" & vbCrLf
        sql2 &= " ,r.MBRNAME" & vbCrLf
        sql2 &= " ,r.JOBTITLE" & vbCrLf
        sql2 &= " ,r.PUSHDISTID" & vbCrLf
        sql2 &= " FROM dbo.OA_MEETEXAM m" & vbCrLf
        sql2 &= " JOIN dbo.OA_EXAMINER r ON r.EMSEQ=m.EMSEQ" & vbCrLf
        sql2 &= " WHERE 1=1" & vbCrLf
        sql2 &= " AND m.MTSEQ=@MTSEQ" & vbCrLf
        '順序排序： 1.先依遴聘類別 2.依姓名筆劃 chinese_taiwan_stroke_cs_as_ks_ws CHINESE_TAIWAN_STROKE_CI_AS
        sql2 &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE CHINESE_TAIWAN_STROKE_CI_AS" & vbCrLf
        'sql2 &= " ORDER BY r.RECRUIT ,r.MBRNAME COLLATE Chinese_PRC_Stroke_ci_as " & vbCrLf

        Dim parms2 As Hashtable = New Hashtable
        parms2.Add("MTSEQ", iMTSEQ)
        dt2 = DbAccess.GetDataTable(sql2, objconn, parms2)

        labmsg_DG4.Text = String.Format("審查委員名單, 共有{0}筆", dt2.Rows.Count)
        If dt2.Rows.Count = 0 Then labmsg_DG4.Text = "審查委員名單, 尚未建立!"

        With DataGrid4
            .DataSource = dt2
            '.DataKeyField = Cst_EXAMINERpkName
            .DataBind()
        End With
    End Sub

    '新增
    Protected Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        hid_EMSEQVAL.Value = TIMS.ClearSQM(hid_EMSEQVAL.Value)
        Dim sEMSEQVAL As String = hid_EMSEQVAL.Value
        txtEXAMINER.Text = ""
        hid_EMSEQVAL.Value = ""
        'If sEMSEQVAL = "" Then Return
        CreateTableDG2(sEMSEQVAL, 2, If(Hid_MTSEQ.Value = "", 0, Val(Hid_MTSEQ.Value)))
    End Sub

    Protected Sub btnBack3_Click(sender As Object, e As EventArgs) Handles btnBack3.Click
        SHOW_PANEL(0)
        'Call ClearData1()
    End Sub

    ''' <summary>儲存-審查委員出席狀況</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSave3_Click(sender As Object, e As EventArgs) Handles btnSave3.Click
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then Return
        Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)
        Call SaveData3(iMTSEQ)
    End Sub

    Private Sub DataGrid3_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"

            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem

                Dim cbATTEND As HtmlInputCheckBox = e.Item.FindControl("cbATTEND") '出席(勾選框)
                Dim cbNOTINABS As HtmlInputCheckBox = e.Item.FindControl("cbNOTINABS") '不列入缺席
                Dim REASON As TextBox = e.Item.FindControl("REASON") '不列入缺席原因
                Dim Hid_EMSEQ As HiddenField = e.Item.FindControl("Hid_EMSEQ")

                cbATTEND.Checked = If(Convert.ToString(drv("ATTEND")).Equals("Y"), True, False)
                cbNOTINABS.Checked = If(Convert.ToString(drv("NOTINABS")).Equals("Y"), True, False)
                REASON.Text = Convert.ToString(drv("REASON"))
                Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))

                cbATTEND.Attributes("onclick") = "Click_cbATTEND('" & cbATTEND.ClientID & "','" & cbNOTINABS.ClientID & "');"
                cbNOTINABS.Attributes("onclick") = "Click_cbNOTINABS('" & cbATTEND.ClientID & "','" & cbNOTINABS.ClientID & "');"

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

        End Select
    End Sub

    Protected Sub btnBack4_Click(sender As Object, e As EventArgs) Handles btnBack4.Click
        SHOW_PANEL(0)
        Call ClearData1()
    End Sub

    Private Sub DataGrid4_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labPUSHDISTID_N As Label = e.Item.FindControl("labPUSHDISTID_N")
                labPUSHDISTID_N.Text = TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(drv("PUSHDISTID")))
                Dim Hid_EMSEQ As HiddenField = e.Item.FindControl("Hid_EMSEQ")
                Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
                '序號
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

        End Select
    End Sub

    '回上頁
    Protected Sub BtnBACKp4_Click(sender As Object, e As EventArgs) Handles BtnBACKp4.Click
        SHOW_PANEL(0)
        'Call ClearData1()
    End Sub

    ''' <summary>儲存-審查委員計畫職類</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSAVEDATAp4_Click(sender As Object, e As EventArgs) Handles BtnSAVEDATAp4.Click
        'Call SAVEDATA_5()
        Hid_MTSEQ.Value = TIMS.ClearSQM(Hid_MTSEQ.Value)
        If Hid_MTSEQ.Value = "" Then Return
        Dim iMTSEQ As Integer = Val(Hid_MTSEQ.Value)
        Call SaveData4(iMTSEQ)
    End Sub

    ''' <summary>儲存前-檢核1</summary>
    ''' <param name="iMTSEQ"></param>
    ''' <returns></returns>
    Function CHK_SAVEDATA1(ByRef iMTSEQ As Integer) As Boolean
        Dim Errmsg As String = ""

        Dim s_parms2 As Hashtable = New Hashtable
        s_parms2.Clear()
        s_parms2.Add("MTSEQ", iMTSEQ)
        Dim s_sql2 As String = ""
        s_sql2 = ""
        s_sql2 &= " SELECT 'X' FROM OA_MEETING" & vbCrLf
        s_sql2 &= " WHERE DATEDIFF(MI,SMEETDATE,GETDATE())>0 AND MTSEQ=@MTSEQ" & vbCrLf
        Dim s_dt2 As DataTable = DbAccess.GetDataTable(s_sql2, objconn, s_parms2)
        If s_dt2.Rows.Count = 0 Then
            'Errmsg &= "儲存(現在)時間需 大於 會議日期之起始時間！"
            'Common.MessageBox(Me, Errmsg)
            Return False
        End If
        Return True
    End Function

    ''' <summary>儲存-審查委員計畫職類 SAVEDATA</summary>
    ''' <param name="iMTSEQ"></param>
    Sub SaveData4(ByRef iMTSEQ As Integer)
        '	儲存時間需>會議日期之起始時間， 如2021/3/8 09:00
        Dim Errmsg As String = ""
        'Dim flag_OK As Boolean = CHK_SAVEDATA1(iMTSEQ)
        'If Not flag_OK Then
        '    Errmsg &= "儲存(現在)時間需 大於 會議日期之起始時間！"
        '    Common.MessageBox(Me, Errmsg)
        '    Return
        'End If

        Dim s_parms As Hashtable = New Hashtable
        Dim s_sql As String = "SELECT 'X' FROM OA_MEETEXAM WHERE MTSEQ=@MTSEQ AND EMSEQ=@EMSEQ"
        'Dim vMTSEQ
        Dim u_parms As Hashtable = New Hashtable
        Dim u_sql As String = ""
        u_sql = " UPDATE OA_MEETEXAM SET ORGKINDG=@ORGKINDG ,ORGKINDW=@ORGKINDW ,GCODE300=@GCODE300" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE MTSEQ=@MTSEQ AND EMSEQ=@EMSEQ" & vbCrLf
        For Each eItem As DataGridItem In DataGrid5.Items
            Dim drv As DataRowView = eItem.DataItem

            Dim Hid_EMSEQ As HiddenField = eItem.FindControl("Hid_EMSEQ") 'OA_EXAMINER 審查委員/審查委員序號
            Dim cbORGKIND_G As HtmlInputCheckBox = eItem.FindControl("cbORGKIND_G")
            Dim cbORGKIND_W As HtmlInputCheckBox = eItem.FindControl("cbORGKIND_W")
            Dim cblGOVCODE3_dg5 As CheckBoxList = eItem.FindControl("cblGOVCODE3_dg5")
            Dim v_cblGOVCODE3_dg5 As String = TIMS.GetCblValue(cblGOVCODE3_dg5)
            Dim iEMSEQ As Integer = TIMS.GetValue1(Val(Hid_EMSEQ.Value))

            s_parms.Clear()
            s_parms.Add("MTSEQ", iMTSEQ)
            s_parms.Add("EMSEQ", iEMSEQ)
            Dim s_dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, s_parms)
            If s_dt.Rows.Count > 0 Then
                u_parms.Clear()
                u_parms.Add("ORGKINDG", If(cbORGKIND_G.Checked, "Y", Convert.DBNull))
                u_parms.Add("ORGKINDW", If(cbORGKIND_W.Checked, "Y", Convert.DBNull))
                u_parms.Add("GCODE300", If(v_cblGOVCODE3_dg5 <> "", v_cblGOVCODE3_dg5, Convert.DBNull))
                u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                u_parms.Add("MTSEQ", iMTSEQ)
                u_parms.Add("EMSEQ", iEMSEQ)
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
            End If
        Next

        SHOW_PANEL(0)
        '儲存成功 'Hid_EMSEQ.Value = ""
        Call ClearData1()
        Common.MessageBox(Me, "儲存成功!")
        Call SSearch1()
    End Sub

    Private Sub DataGrid5_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid5.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Dim Hid_EMSEQ As HiddenField = e.Item.FindControl("Hid_EMSEQ") 'OA_EXAMINER 審查委員/審查委員序號
                Dim cbORGKIND_G As HtmlInputCheckBox = e.Item.FindControl("cbORGKIND_G")
                Dim cbORGKIND_W As HtmlInputCheckBox = e.Item.FindControl("cbORGKIND_W")
                Dim cblGOVCODE3_dg5 As CheckBoxList = e.Item.FindControl("cblGOVCODE3_dg5")
                '產業人才投資計畫
                cbORGKIND_G.Checked = If(Convert.ToString(drv("ORGKINDG")) = "Y", True, False)
                '提升勞工自主學習計畫
                cbORGKIND_W.Checked = If(Convert.ToString(drv("ORGKINDW")) = "Y", True, False)
                '審查課程職類-依dtGOVCODE3
                cblGOVCODE3_dg5 = TIMS.Get_GOVCODE3(dtGCODE3_5, cblGOVCODE3_dg5, False)
                TIMS.SetCblValue(cblGOVCODE3_dg5, Convert.ToString(drv("GCODE300")))
                'Dim s_CmdArg As String = "" 'TIMS.SetMyValue(s_CmdArg, "EMSEQ", drv("EMSEQ"))
                Hid_EMSEQ.Value = Convert.ToString(drv("EMSEQ"))
        End Select
    End Sub

End Class
