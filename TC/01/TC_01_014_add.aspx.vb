Partial Class TC_01_014_add
    Inherits AuthBasePage

    '#Region "(No Use)"

    'Dim objconn As SqlConnection
    'Dim objreader As SqlDataReader
    'Plan_OnClass,'Plan_TrainDesc,'Teach_TeacherInfo,'Plan_Teacher,
    'select * from Plan_OnClass where PlanID='1783'
    'select * from Plan_TrainDesc where PlanID='1783'
    'select * from Plan_Teacher where PlanID='1783'
    'select * from Teach_TeacherInfo p
    'where 1=1
    'and exists (
    '	select 'x' from Plan_Teacher x where PlanID='1783' and x.TechID=p.TechID
    ')

    ''取得教師學歷
    'Public Shared Function Get_TeacherDegree(ByVal TechID As Object) As String
    '    Dim objstr As String
    '    objstr="select DegreeID from Teach_TeacherInfo where TechID='" & TechID & "'"
    '    Return DbAccess.ExecuteScalar(objstr, objconn)
    'End Function


    '#Region "參數/變數 設定"
    Const cst_ptInsert As String = "Insert"
    Const cst_ptUpdate As String = "Update"
    Const cst_ptView As String = "View"
    Const cst_msg_memo8a As String = "本課程非屬於「職業安全衛生教育訓練規則」所訂定之訓練課程，無法作為時數認列。"
    Const cst_SaveDef As String = "草稿儲存"
    Const cst_SaveRcc As String = "正式送出"
    Const cst_PlanID As String = "PlanID"
    Const cst_ComIDNO As String = "ComIDNO"
    Const cst_SeqNO As String = "SeqNO"
    Const cst_ProcessType As String = "ProcessType" 'ProcessType @Insert/Update/View

    Dim rqPlanID As String = "" 'Request(cst_PlanID)
    Dim rqComIDNO As String = "" 'Request(cst_ComIDNO)
    Dim rqSeqNO As String = "" 'Request(cst_SeqNO)
    Dim rqProcessType As String = "" 'Request(cst_ProcessType)
    Dim rqPlanYear As String = "" 'Request("PlanYear")
    Dim rqTPlanID As String = "" 'Request("TPlanID")
    Dim rqTMID As String = "" 'Request("TMID")
    Dim rqRID As String = "" 'Request("RID")
    Dim rqIsApprPaper As String = ""
    Dim sWOScript1 As String = ""
    Dim iSeqno As Integer = 0

    'Dim iPYNum As Integer=1 'iPYNum=TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ''#Region "在這裡放置使用者程式碼以初始化網頁"

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'iPYNum=TIMS.sUtl_GetPYNum(Me)
        rqPlanID = TIMS.ClearSQM(Request(cst_PlanID))
        rqComIDNO = TIMS.ClearSQM(Request(cst_ComIDNO))
        rqSeqNO = TIMS.ClearSQM(Request(cst_SeqNO))
        rqProcessType = TIMS.ClearSQM(Request(cst_ProcessType)) 'ProcessType @Insert/Update/View
        rqPlanYear = TIMS.ClearSQM(Request("PlanYear"))
        rqTPlanID = TIMS.ClearSQM(Request("TPlanID"))
        rqTMID = TIMS.ClearSQM(Request("TMID"))
        rqRID = TIMS.ClearSQM(Request("RID"))
        rqIsApprPaper = TIMS.ClearSQM(Request("IsApprPaper"))

        Dim drP As DataRow = TIMS.GetPCSDate(rqPlanID, rqComIDNO, rqSeqNO, objconn)
        If drP Is Nothing Then
            Dim url1 As String = "TC_01_014.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
            Call TIMS.Utl_Redirect(Me, objconn, url1)
        End If
        If rqRID = "" Then
            rqRID = sm.UserInfo.RID '若為空白，取得目前登入者RID 
            If Not drP Is Nothing Then rqRID = Convert.ToString(drP("RID"))
        End If
        RIDValue.Value = rqRID '只能是委訓單位

        If Not IsPostBack Then
            TIMS.PL_placeholder(tPOWERNEED1)
            TIMS.PL_placeholder(tPOWERNEED2)
            TIMS.PL_placeholder(tPOWERNEED4)
            TIMS.ChkTextLength(TMethodOth, 100)
            TIMS.ChkTextLength(tPOWERNEED1, 1000)
            TIMS.ChkTextLength(tPOWERNEED2, 1000)
            TIMS.ChkTextLength(tPOWERNEED3, 200)
            TIMS.ChkTextLength(tPOWERNEED4, 200)
            Button6.Text = cst_SaveDef
            bt_addrow.Text = cst_SaveRcc

            '建立物件--Start
            Call CreateItem()
            Call GET_PLAN_TRAINPLACE() '(rqComIDNO)
            Dim dt2 As DataTable = Get_PLAN_PLANINFO()
            Call CreateClassTime()
            Call GET_PLAN_TEACHER12()
            Call CreateTrainDesc()
            '建立物件--End

            If dt2.Rows.Count = 0 Then Exit Sub
            Call SHOW_DATA1(dt2)
            Call CHK_OBJENABLED() '確認各物件的屬性
        End If

        If Not Session("_search") Is Nothing Then
            ViewState("_search") = Session("_search")
            Session("_search") = Nothing
        End If

        Dim strScript As String
        strScript = "<script>showPanel();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    '確認各物件的屬性
    Sub CHK_OBJENABLED()
        ClassName.ReadOnly = True
        ClassCate.Enabled = False
        ClassID.Enabled = False '此課程班別，與課程轉入有部份相同功能，暫不輸入，可顯示 by 豪哥／AMU 2008-04-28
        TIMS.Tooltip(ClassID, "課程班別，暫不輸入")
        Tnum.ReadOnly = True
        THours.ReadOnly = True
        start_date.ReadOnly = True
        end_date.ReadOnly = True
        IMG1.Visible = False
        IMG2.Visible = False
        CapDegree.Enabled = False
        tPlanCause.ReadOnly = True
        tPurScience.ReadOnly = True
        tPurTech.ReadOnly = True
        tPurMoral.ReadOnly = True
        CostDesc.ReadOnly = True
        'TRA1.Disabled=True
        'TRA2.Disabled=True
        'TRB1.Disabled=True
        'TRB2.Disabled=True
        'TRC1.Disabled=True

        'Tnum2.Enabled=False
        'HwDesc2.Enabled=False
        'Tnum3.Enabled=False
        'HwDesc3.Enabled=False
        'OtherDesc23.Enabled=False

        Select Case rqProcessType 'ProcessType @Insert/Update/View
            Case cst_ptView '查詢功能不提供儲存
                Button6.Visible = False '草稿儲存
                bt_addrow.Visible = False '正式送出
                'ChoiceTechBtn.Visible=False '師資選單
                Dim strScript As String = ""
                strScript = "<script>lock1();</script>"
                TIMS.RegisterStartupScript(Me, TIMS.xBlockName, strScript)
            Case cst_ptUpdate '修改功能檢查是否為草稿
                '已經為正式，不可使用草稿送出
                If rqIsApprPaper = "Y" Then Button6.Visible = False Else Button6.Visible = True
                bt_addrow.Visible = True '正式送出
                'ChoiceTechBtn.Visible=True '師資選單
        End Select
    End Sub

    Sub SHOW_DATA1(ByVal dt2 As DataTable)
        '#Region "SHOW_DATA1"

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If dt2.Rows.Count = 0 Then Exit Sub
        Dim dr2 As DataRow = dt2.Rows(0)

        Common.SetListItem(PlanYear, rqPlanYear)
        ClassName.Text = Convert.ToString(dr2("classname"))
        If Convert.ToString(dr2("ClassCate")) <> "" Then Common.SetListItem(ClassCate, dr2("ClassCate"))
        Tnum.Text = Convert.ToString(dr2("Tnum")) 'Request("Tnum")
        THours.Text = Convert.ToString(dr2("THours")) 'Request("THours")
        start_date.Text = TIMS.Cdate3(dr2("STDate")) 'Request("STDate")
        end_date.Text = TIMS.Cdate3(dr2("FDDate")) 'Request("FDDate")
        If Convert.ToString(dr2("CapDegree")) <> "" Then Common.SetListItem(CapDegree, dr2("CapDegree"))
        DefGovCost.Text = Convert.ToString(dr2("DefGovCost"))
        If DefGovCost.Text = "" Then DefGovCost.Text = 0
        DefStdCost.Text = Convert.ToString(dr2("DefStdCost"))
        If DefStdCost.Text = "" Then DefStdCost.Text = 0
        TotalCost.Text = CInt(DefGovCost.Text) + CInt(DefStdCost.Text)
        DefGovCost_Tnum.Text = ""
        DefStdCost_Tnum.Text = ""
        If Tnum.Text <> "" AndAlso Tnum.Text <> "0" Then
            DefGovCost_Tnum.Text = CInt(DefGovCost.Text) / CInt(Tnum.Text)
            DefStdCost_Tnum.Text = CInt(DefStdCost.Text) / CInt(Tnum.Text)
            TotalCost_Tnum.Text = CInt(TotalCost.Text) / CInt(Tnum.Text)
        End If

        '接收 Request的內容
        Select Case Convert.ToString(rqProcessType) 'ProcessType @Insert/Update/View
            Case cst_ptUpdate, cst_ptView
                Call SHOW_DATA2()
            Case Else 'Insert
                Exit Sub
        End Select
    End Sub

    'PLAN_VERREPORT
    Sub SHOW_DATA2()
        '#Region "SHOW_DATA2"

        '接收 Request的內容
        Select Case Convert.ToString(rqProcessType) 'ProcessType @Insert/Update/View
            Case cst_ptUpdate, cst_ptView
            Case Else 'Insert
                Exit Sub
        End Select

        Dim PMS1 As New Hashtable From {{"PlanID", TIMS.CINT1(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", TIMS.CINT1(rqSeqNO)}}
        Dim sql As String = " SELECT * FROM PLAN_VERREPORT WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "尚未建立此班開班計劃資料維護")
            Dim url1 As String = "../01/TC_01_014.aspx?ID=" & Request("ID")
            Call TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        If TIMS.dtNODATA(dt) Then Exit Sub

        Dim dr As DataRow = dt.Rows(0)
        Common.SetListItem(rblFuncLevel, Convert.ToString(dr("FuncLevel")))
        If Convert.ToString(dr("TMethod")) <> "" Then TIMS.SetCblValue(cblTMethod, Convert.ToString(dr("TMethod")))
        TMethodOth.Text = Convert.ToString(dr("TMethodOth"))
        Common.SetListItem(ClassID, dr("ClassID"))
        TIMS.PL_settextbox1(tPOWERNEED1, dr("POWERNEED1"))
        TIMS.PL_settextbox1(tPOWERNEED2, dr("POWERNEED2"))
        TIMS.PL_settextbox1(tPOWERNEED3, dr("POWERNEED3"))
        cbPOWERNEED4.Checked = False
        If Convert.ToString(dr("POWERNEED4CHK")) = TIMS.cst_YES Then cbPOWERNEED4.Checked = True
        If cbPOWERNEED4.Checked AndAlso Convert.ToString(dr("POWERNEED4")) <> "" Then TIMS.PL_settextbox1(tPOWERNEED4, dr("POWERNEED4"))
        If tPlanCause.Text = "" AndAlso Convert.ToString(dr("PlanCause")) <> "" Then tPlanCause.Text = Convert.ToString(dr("PlanCause"))
        If tPurScience.Text = "" AndAlso Convert.ToString(dr("PurScience")) <> "" Then tPurScience.Text = Convert.ToString(dr("PurScience"))
        If tPurTech.Text = "" AndAlso Convert.ToString(dr("PurTech")) <> "" Then tPurTech.Text = Convert.ToString(dr("PurTech"))
        If tPurMoral.Text = "" AndAlso Convert.ToString(dr("PurMoral")) <> "" Then tPurMoral.Text = Convert.ToString(dr("PurMoral"))
        CapAll.Text = dr("CapAll").ToString
        If CostDesc.Text = "" Then
            If dr("CostDesc").ToString <> "" Then CostDesc.Text = dr("CostDesc").ToString
        End If
        TrainMode.Enabled = False
        TrainMode.Text = "(請勾選教學方法)"
        RecDesc.Text = Convert.ToString(dr("RecDesc")) '.ToString
        LearnDesc.Text = Convert.ToString(dr("LearnDesc")) '.ToString
        ActDesc.Text = Convert.ToString(dr("ActDesc")) '.ToString
        ResultDesc.Text = Convert.ToString(dr("ResultDesc")) '.ToString
        OtherDesc.Text = Convert.ToString(dr("OtherDesc")) '.ToString

        chk_RecDesc.Checked = False
        If RecDesc.Text <> "" Then chk_RecDesc.Checked = True
        chk_LearnDesc.Checked = False
        If LearnDesc.Text <> "" Then chk_LearnDesc.Checked = True
        chk_ActDesc.Checked = False
        If ActDesc.Text <> "" Then chk_ActDesc.Checked = True
        chk_ResultDesc.Checked = False
        If ResultDesc.Text <> "" Then chk_ResultDesc.Checked = True
        chk_OtherDesc.Checked = False
        If OtherDesc.Text <> "" Then chk_OtherDesc.Checked = True

        '是否為iCAP課程 / 是, 請填寫/否/ 課程相關說明
        Dim sISiCAPCOUR As String = Convert.ToString(dr("ISiCAPCOUR"))
        RB_ISiCAPCOUR_Y.Checked = If(sISiCAPCOUR = "Y", True, False)
        RB_ISiCAPCOUR_N.Checked = If(sISiCAPCOUR = "N", True, False)
        iCAPCOURDESC.Text = Convert.ToString(dr("iCAPCOURDESC")) '課程相關說明
        Recruit.Text = Convert.ToString(dr("Recruit")) '招訓方式
        Selmethod.Text = Convert.ToString(dr("Selmethod")) '遴選方式
        Inspire.Text = Convert.ToString(dr("Inspire")) '學員激勵辦法

        TGovExamCY.Checked = False
        TGovExamCN.Checked = False
        Select Case Convert.ToString(dr("TGovExam"))
            Case "Y"
                TGovExamCY.Checked = True
            Case "N"
                TGovExamCN.Checked = True
        End Select
        TGovExamName.Text = dr("TGovExamName").ToString
        chkMEMO8C1.Checked = False
        chkMEMO8C2.Checked = False
        txtMemo8.Text = ""
        If Convert.ToString(dr("memo8")) <> "" Then chkMEMO8C1.Checked = True
        If Convert.ToString(dr("memo82")) <> "" Then
            chkMEMO8C2.Checked = True
            txtMemo8.Text = Convert.ToString(dr("memo82"))
        End If
    End Sub

    ''' <summary>
    ''' 建立上課時間
    ''' </summary>
    Sub CreateClassTime()

        Dim PMS1 As New Hashtable From {{"PlanID", TIMS.CINT1(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", TIMS.CINT1(rqSeqNO)}}
        Dim sql As String = " SELECT * FROM PLAN_ONCLASS WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)
        DataGrid1Table.Visible = False

        If dt.Rows.Count > 0 Then
            DataGrid1Table.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    ''' <summary>
    ''' 計畫訓練內容簡介
    ''' </summary>
    Sub CreateTrainDesc()
        Dim PMS1 As New Hashtable From {{"PlanID", TIMS.CINT1(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", TIMS.CINT1(rqSeqNO)}}
        Dim sql As String = ""
        sql &= " SELECT * FROM PLAN_TRAINDESC WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        sql &= " ORDER BY STrainDate,PName"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)

        Datagrid3Table.Style.Item("display") = "none"
        If dt.Rows.Count > 0 Then
            Datagrid3Table.Style.Item("display") = ""  '"inline"
            With Datagrid3
                .DataSource = dt
                .DataKeyField = "PTDID"
                .DataBind()
            End With
        End If
    End Sub

    '建立可選教師列表
    Sub GET_PLAN_TEACHER12()
        '#Region "建立可選教師列表"

        If rqRID = "" Then
            Common.RespWrite(Me, "<script>alert('查詢 教師列表 資料有誤!!');</script>")
            Common.RespWrite(Me, "<script>location.href='TC_01_014.aspx?ID=" & Request("ID") & "'</script>")
            Exit Sub
        End If

        Dim PMS1 As New Hashtable From {{"PlanID", TIMS.CINT1(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", TIMS.CINT1(rqSeqNO)}, {"RID", rqRID}}
        Dim sql As String = ""
        sql &= " SELECT a.TechID " & vbCrLf '教師ID
        sql &= " ,a.TeachCName " & vbCrLf '教師姓名 
        sql &= " ,a.DegreeID " & vbCrLf '學歷
        sql &= " ,c.Name DegreeName " & vbCrLf '學歷
        '專業領域
        sql &= " ,REPLACE(ISNULL(a.Specialty1, ' '),',',' ') " & vbCrLf
        sql &= " + REPLACE(ISNULL(a.Specialty2, ' '),',',' ') " & vbCrLf
        sql &= " + REPLACE(ISNULL(a.Specialty3, ' '),',',' ') " & vbCrLf
        sql &= " + REPLACE(ISNULL(a.Specialty4, ' '),',',' ') " & vbCrLf
        sql &= " + REPLACE(ISNULL(a.Specialty5, ' '),',',' ') major " & vbCrLf '專業領域
        '專業證照-相關證照
        sql &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2 " & vbCrLf
        sql &= " ELSE a.ProLicense END ProLicense " & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'A', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
        sql &= " FROM TEACH_TEACHERINFO a " & vbCrLf
        sql &= " JOIN ( SELECT DISTINCT TechID, PLANID, COMIDNO, SEQNO FROM PLAN_TRAINDESC " & vbCrLf
        sql &= " WHERE TECHID IS NOT NULL AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo " & vbCrLf
        sql &= " ) b ON a.TechID=b.TechID " & vbCrLf
        sql &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID=c.DegreeID " & vbCrLf
        sql &= " WHERE a.WorkStatus='1' AND a.RID=@RID" & vbCrLf
        sql &= " ORDER BY a.TechID " & vbCrLf
        Dim dtT As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)

        iSeqno = 0
        Datagrid2Table.Visible = False
        If dtT.Rows.Count > 0 Then
            Datagrid2Table.Visible = True
            DataGrid2.DataSource = dtT
            DataGrid2.DataBind()
        End If

        Dim PMS2 As New Hashtable From {{"PlanID", TIMS.CINT1(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", TIMS.CINT1(rqSeqNO)}, {"RID", rqRID}}
        Dim sql2 As String = ""
        sql2 &= " SELECT a.TechID " & vbCrLf '教師ID
        sql2 &= " ,a.TeachCName " & vbCrLf '教師姓名 
        sql2 &= " ,a.DegreeID " & vbCrLf '學歷
        sql2 &= " ,c.Name DegreeName " & vbCrLf '學歷
        sql2 &= " ,REPLACE(ISNULL(a.Specialty1, ' '),',',' ') " & vbCrLf
        sql2 &= " + REPLACE(ISNULL(a.Specialty2, ' '),',',' ') " & vbCrLf
        sql2 &= " + REPLACE(ISNULL(a.Specialty3, ' '),',',' ') " & vbCrLf
        sql2 &= " + REPLACE(ISNULL(a.Specialty4, ' '),',',' ') " & vbCrLf
        sql2 &= " + REPLACE(ISNULL(a.Specialty5, ' '),',',' ') major " & vbCrLf '專業領域
        '專業證照-相關證照
        sql2 &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2 " & vbCrLf
        sql2 &= " ELSE a.ProLicense END ProLicense " & vbCrLf
        sql2 &= " ,dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'B', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
        sql2 &= " FROM TEACH_TEACHERINFO a " & vbCrLf
        sql2 &= " JOIN ( SELECT DISTINCT TECHID2 TechID, planid, comidno, seqno " & vbCrLf
        sql2 &= " FROM PLAN_TRAINDESC " & vbCrLf
        sql2 &= " WHERE TECHID2 IS NOT NULL AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo " & vbCrLf
        sql2 &= " ) b ON a.TechID=b.TechID " & vbCrLf
        sql2 &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID=c.DegreeID " & vbCrLf
        sql2 &= " WHERE a.WorkStatus='1' AND a.RID=@RID " & vbCrLf
        sql2 &= " ORDER BY a.TechID " & vbCrLf
        Dim dtT2 As DataTable = DbAccess.GetDataTable(sql2, objconn, PMS2)
        iSeqno = 0
        Datagrid2Table2.Visible = False

        If dtT2.Rows.Count > 0 Then
            Datagrid2Table2.Visible = True
            DataGrid22.DataSource = dtT2
            DataGrid22.DataBind()
        End If
    End Sub

    'Not Page.IsPostBack 建立物件
    Sub CreateItem()
        '#Region "CreateItem"

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        lbMEMO8.Text = cst_msg_memo8a
        PlanYear = TIMS.GetSyear(PlanYear)
        Call TIMS.Get_ClassCatelog(ClassCate, objconn)
        Dim sql_Degree As String = ""
        Dim dt As DataTable
        sql_Degree = " SELECT DEGREEID, NAME FROM KEY_DEGREE ORDER BY DEGREEID "
        dt = DbAccess.GetDataTable(sql_Degree, objconn)

        With CapDegree
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "DegreeID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            '.SelectedValue=Request("CapDegree")
        End With

        ClassID = TIMS.Get_ClassID(ClassID, rqTPlanID, rqTMID, sm.UserInfo.DistID, objconn)
        'Dim CHARID As String=TIMS.Get_CHARID(sm.UserInfo.OrgID, sm.UserInfo.Years)
    End Sub

    '取得PlanInfo
    Function Get_PLAN_PLANINFO() As DataTable
        '#Region "取得PlanInfo"

        Dim PMS2 As New Hashtable From {{"PlanID", TIMS.CINT1(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", TIMS.CINT1(rqSeqNO)}}
        Dim sql As String = " SELECT * FROM PLAN_PLANINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo " & vbCrLf
        If rqPlanYear <> "" Then sql &= " AND PlanYear='" & rqPlanYear & "' "
        If rqTMID <> "" Then sql &= " AND TMID='" & rqTMID & "' "
        If rqRID <> "" Then sql &= " AND RID='" & rqRID & "' "
        If rqTPlanID <> "" Then sql &= " AND TPlanID='" & rqTPlanID & "' "

        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn, PMS2)
        If dt2.Rows.Count = 0 Then Return dt2

        Dim dr As DataRow = dt2.Rows(0)
        tPlanCause.Text = Convert.ToString(dr("PlanCause"))
        tPurScience.Text = Convert.ToString(dr("PurScience"))
        tPurTech.Text = Convert.ToString(dr("PurTech"))
        tPurMoral.Text = Convert.ToString(dr("PurMoral"))
        CostDesc.Text = dr("Note").ToString

        Return dt2
    End Function

    '取得TrainPlace
    Sub GET_PLAN_TRAINPLACE() '(ByVal ComIDNO As String)
        '#Region "取得TrainPlace"

        Dim OtherMsg As String = ""
        Dim objstr As String = ""
        Dim dr As DataRow
        Dim dt As DataTable

        OtherMsg = ""
        objstr = "" & vbCrLf
        objstr &= " SELECT b.connum " & vbCrLf
        objstr &= " ,b.hwdesc " & vbCrLf
        objstr &= " ,b.OtherDesc " & vbCrLf
        objstr &= " ,b.PTID " & vbCrLf
        objstr &= " ,b.PLACEID " & vbCrLf
        objstr &= " ,b.ClassIFICation " & vbCrLf
        objstr &= " FROM PLAN_PLANINFO a " & vbCrLf
        objstr &= " LEFT JOIN PLAN_TRAINPLACE b ON a.ComIDNO=b.ComIDNO AND a.SCIPLACEID=b.PLACEID " & vbCrLf
        objstr &= " AND b.ComIDNO='" & rqComIDNO & "' "
        objstr &= " WHERE 1=1 " & vbCrLf
        objstr &= " AND a.PlanID='" & rqPlanID & "' " & vbCrLf
        objstr &= " AND a.ComIDNO='" & rqComIDNO & "' " & vbCrLf
        objstr &= " AND a.SeqNo='" & rqSeqNO & "' " & vbCrLf
        objstr &= " AND b.ClassIFICation IN (1,3) " & vbCrLf '學科共用。
        dt = DbAccess.GetDataTable(objstr, objconn)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Tnum2.Text = Convert.ToString(dr("connum")) 'dr("connum").ToString
            HwDesc2.Text = dr("hwdesc").ToString
            If dr("OtherDesc").ToString <> "" Then OtherMsg &= dr("OtherDesc").ToString & vbCrLf
        End If
        If OtherMsg <> "" Then OtherMsg &= vbCrLf '換行

        objstr = "" & vbCrLf
        objstr &= " SELECT b.connum " & vbCrLf
        objstr &= " ,b.hwdesc " & vbCrLf
        objstr &= " ,b.OtherDesc " & vbCrLf
        objstr &= " ,b.PTID " & vbCrLf
        objstr &= " ,b.PLACEID " & vbCrLf
        objstr &= " ,b.ClassIFICation " & vbCrLf
        objstr &= " FROM PLAN_PLANINFO a " & vbCrLf
        objstr &= " LEFT JOIN PLAN_TRAINPLACE b ON a.ComIDNO=b.ComIDNO AND a.TECHPLACEID=b.PLACEID " & vbCrLf
        objstr &= " AND b.ComIDNO='" & rqComIDNO & "' "
        objstr &= " WHERE 1=1 " & vbCrLf
        objstr &= " AND a.PlanID='" & rqPlanID & "' " & vbCrLf
        objstr &= " AND a.ComIDNO='" & rqComIDNO & "' " & vbCrLf
        objstr &= " AND a.SeqNo='" & rqSeqNO & "' " & vbCrLf
        objstr &= " AND b.ClassIFICation IN (2,3) " & vbCrLf '術科共用。
        dt = DbAccess.GetDataTable(objstr, objconn)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Tnum3.Text = Convert.ToString(dr("connum")) '.ToString
            HwDesc3.Text = dr("hwdesc").ToString
            If dr("OtherDesc").ToString <> "" Then OtherMsg &= dr("OtherDesc").ToString & vbCrLf
        End If
        '術科
        If OtherMsg <> "" Then OtherDesc23.Text = OtherMsg
    End Sub


    '檢查 班級申請老師
    Function CHK_PLAN_TEACHER12(ByRef errmsg As String) As Boolean
        '#Region "檢查 班級申請老師"

        Dim rst As Boolean = True
        Const Cst_授課教師限制數 As Integer = 0 '10 '0:無限制

        Select Case rqProcessType 'ProcessType @Insert/Update/View
            Case cst_ptInsert, cst_ptUpdate
                Dim i As Integer = 0
                Dim errT As String = ""
                Dim errI2 As Integer = 0
                For Each eItem As DataGridItem In DataGrid2.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    Dim seqno As Label = eItem.FindControl("seqno")
                    Dim TeachCName As Label = eItem.FindControl("TeachCName")
                    Dim DegreeName As Label = eItem.FindControl("DegreeName")
                    Dim major As Label = eItem.FindControl("major")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA") 'TechTYPE: A:師資/B:助教
                    i += 1
                    If TeacherDesc.Text = "" Then
                        errT = seqno.Text & ":" & TeachCName.Text
                        errI2 += 1
                        Exit For
                    End If
                Next

                If i = 0 Then
                    errmsg &= "至少選擇1筆授課教師" & vbCrLf
                    Return False
                End If
                If errI2 > 0 Then
                    errmsg &= "授課教師-" & errT & "-遴選辦法說明辦法為必填" & vbCrLf
                    Return False
                End If
                If Cst_授課教師限制數 <> 0 Then '0:無限制
                    If Not (i <= Cst_授課教師限制數) Then
                        errmsg &= "僅可選擇" & Cst_授課教師限制數 & "筆授課教師" & vbCrLf
                        Return False
                    End If
                End If

                Dim errTB As String = ""
                Dim errI2B As Integer = 0
                For Each eItem As DataGridItem In DataGrid22.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    Dim seqno As Label = eItem.FindControl("seqno")
                    Dim TeachCName As Label = eItem.FindControl("TeachCName")
                    Dim DegreeName As Label = eItem.FindControl("DegreeName")
                    Dim major As Label = eItem.FindControl("major")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    Dim btn_TCTYPEB As HtmlInputButton = eItem.FindControl("btn_TCTYPEB") 'TechTYPE: A:師資/B:助教
                    If TeacherDesc.Text = "" Then
                        errTB = seqno.Text & ":" & TeachCName.Text
                        errI2B += 1
                        Exit For
                    End If
                Next

                If errI2B > 0 Then
                    errmsg &= "授課助教-" & errTB & "-遴選辦法說明辦法為必填" & vbCrLf
                    Return False
                End If
        End Select

        Return rst
    End Function

    '儲存 班級申請老師
    Sub SAVE_PLAN_TEACHER(ByVal tConn As SqlConnection)
        '#Region "儲存 班級申請老師"

        Select Case rqProcessType 'ProcessType @Insert/Update/View
            Case cst_ptInsert, cst_ptUpdate
                Dim dt As DataTable = Nothing
                Dim dr As DataRow = Nothing
                Dim da As SqlDataAdapter = Nothing
                Dim sql As String = ""

                Call TIMS.OpenDbConn(tConn)
                Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
                Try
                    sql = ""
                    sql &= " DELETE PLAN_TEACHER "
                    sql &= " WHERE PlanID='" & rqPlanID & "' AND ComIDNO='" & rqComIDNO & "' AND SeqNo='" & rqSeqNO & "'"
                    DbAccess.ExecuteNonQuery(sql, trans)

                    sql = " SELECT * FROM PLAN_TEACHER WHERE 1<>1 "
                    dt = DbAccess.GetDataTable(sql, da, trans)

                    For Each eItem As DataGridItem In DataGrid2.Items
                        Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                        Dim seqno As Label = eItem.FindControl("seqno")
                        Dim TeachCName As Label = eItem.FindControl("TeachCName")
                        Dim DegreeName As Label = eItem.FindControl("DegreeName")
                        Dim major As Label = eItem.FindControl("major")
                        Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                        Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA")

                        If HidTechID.Value <> "" Then
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("PlanID") = rqPlanID
                            dr("ComIDNO") = rqComIDNO
                            dr("SeqNo") = rqSeqNO
                            dr("TechID") = HidTechID.Value
                            dr("TechTYPE") = "A" 'TechTYPE: A:師資/B:助教
                            dr("TeacherDesc") = TIMS.ClearSQM(TeacherDesc.Text)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now
                        End If
                    Next

                    For Each eItem As DataGridItem In DataGrid22.Items
                        Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                        Dim seqno As Label = eItem.FindControl("seqno")
                        Dim TeachCName As Label = eItem.FindControl("TeachCName")
                        Dim DegreeName As Label = eItem.FindControl("DegreeName")
                        Dim major As Label = eItem.FindControl("major")
                        Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                        Dim btn_TCTYPEB As HtmlInputButton = eItem.FindControl("btn_TCTYPEB")

                        If HidTechID.Value <> "" Then
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("PlanID") = rqPlanID
                            dr("ComIDNO") = rqComIDNO
                            dr("SeqNo") = rqSeqNO
                            dr("TechID") = HidTechID.Value
                            dr("TechTYPE") = "B" 'TechTYPE: A:師資/B:助教
                            dr("TeacherDesc") = TIMS.ClearSQM(TeacherDesc.Text)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now
                        End If
                    Next
                    DbAccess.UpdateDataTable(dt, da, trans)
                    DbAccess.CommitTrans(trans)
                Catch ex As Exception
                    DbAccess.RollbackTrans(trans)
                    Call TIMS.CloseDbConn(tConn)
                    Common.MessageBox(Me, "儲存失敗!!")
                    Exit Sub
                End Try
        End Select


    End Sub

    '儲存 班級申請老師(CLASS_TEACHER)
    Sub SAVE_CLASS_TEACHER(ByVal iOCID As Integer, ByVal tConn As SqlConnection)
        '#Region "儲存 班級申請老師(CLASS_TEACHER)"

        '更新師資表--Start
        'Dim tConn As SqlConnection=DbAccess.GetConnection()
        'Call TIMS.OpenDbConn(tConn)
        'Dim trans As SqlTransaction=DbAccess.BeginTrans(tConn)
        Const cst_iMaxLen_TeacherDesc As Integer = 500
        '更新師資表 'TechTYPE: A:師資/B:助教
        Const cst_tTECHTYPE_A As String = "A"
        Const cst_tTECHTYPE_B As String = "B"

        Dim dParms As New Hashtable From {{"OCID", iOCID}}
        Dim dSql As String = " DELETE CLASS_TEACHER WHERE OCID=@OCID"
        DbAccess.ExecuteNonQuery(dSql, tConn, dParms)

        Dim iSqlc As String = ""
        iSqlc &= " INSERT INTO CLASS_TEACHER ( CTRID ,OCID,TECHID,MODIFYACCT,MODIFYDATE,TECHTYPE,TEACHERDESC)" & vbCrLf
        iSqlc &= " VALUES ( @CTRID ,@OCID,@TECHID,@MODIFYACCT,GETDATE(),@TECHTYPE,@TEACHERDESC)" & vbCrLf

        Dim sSql1 As String = ""
        sSql1 = " SELECT 1 FROM CLASS_TEACHER WHERE OCID=@OCID AND TECHID=@TECHID AND TECHTYPE=@TECHTYPE" & vbCrLf

        For Each eItem As DataGridItem In DataGrid2.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label=eItem.FindControl("seqno")
            'Dim TeachCName As Label=eItem.FindControl("TeachCName")
            'Dim DegreeName As Label=eItem.FindControl("DegreeName")
            'Dim major As Label=eItem.FindControl("major")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEA As HtmlInputButton=eItem.FindControl("btn_TCTYPEA")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable
                sParms.Add("OCID", iOCID)
                sParms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                sParms.Add("TECHTYPE", cst_tTECHTYPE_A) 'TechTYPE: A:師資/B:助教
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(tConn, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    Dim parms As New Hashtable 'parms.Clear()
                    parms.Add("CTRID", iCTRID)
                    parms.Add("OCID", iOCID)
                    parms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                    parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    parms.Add("TECHTYPE", cst_tTECHTYPE_A)
                    parms.Add("TEACHERDESC", tTEACHERDESC)
                    DbAccess.ExecuteNonQuery(iSqlc, tConn, parms)
                End If
            End If
        Next

        For Each eItem As DataGridItem In DataGrid22.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label=eItem.FindControl("seqno")
            'Dim TeachCName As Label=eItem.FindControl("TeachCName")
            'Dim DegreeName As Label=eItem.FindControl("DegreeName")
            'Dim major As Label=eItem.FindControl("major")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEB As HtmlInputButton=eItem.FindControl("btn_TCTYPEB")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable
                sParms.Add("OCID", iOCID)
                sParms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                sParms.Add("TECHTYPE", cst_tTECHTYPE_B) 'TechTYPE: A:師資/B:助教
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(tConn, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    Dim parms As New Hashtable 'parms.Clear()
                    parms.Add("CTRID", iCTRID)
                    parms.Add("OCID", iOCID)
                    parms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                    parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    parms.Add("TECHTYPE", cst_tTECHTYPE_B)
                    parms.Add("TEACHERDESC", tTEACHERDESC)
                    DbAccess.ExecuteNonQuery(iSqlc, tConn, parms)
                End If
            End If
        Next


        'Try


        '    'DbAccess.UpdateDataTable(dt, da, trans)
        '    DbAccess.CommitTrans(trans)
        'Catch ex As Exception
        '    DbAccess.RollbackTrans(trans)
        '    Call TIMS.CloseDbConn(tConn)
        '    Common.MessageBox(Me, "儲存失敗!!")
        '    Exit Sub
        'End Try
        '更新師資表--End
    End Sub

    '儲存 開班計畫表資料維護
    Sub SAVE_PLAN_VERREPORT(ByVal SaveType1 As String)
        '#Region "儲存 開班計畫表資料維護"

        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing

        Dim vTMethod As String = TIMS.GetCblValue(cblTMethod)

        Dim iPVID As Integer = 0
        Select Case rqProcessType 'ProcessType @Insert/Update/View
            Case cst_ptInsert
                sql = ""
                sql &= " SELECT * FROM PLAN_VERREPORT "
                sql &= " WHERE PlanID='" & rqPlanID & "' "
                sql &= " AND ComIDNO='" & rqComIDNO & "' "
                sql &= " AND SeqNo='" & rqSeqNO & "' "
                dt = DbAccess.GetDataTable(sql, da, objconn)

                If dt.Rows.Count = 0 Then
                    iPVID = DbAccess.GetNewId(objconn, "PLAN_VERREPORT_PVID_SEQ,PLAN_VERREPORT,PVID")
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("PVID") = iPVID
                End If
                If dt.Rows.Count = 1 Then
                    '新增 卻有資料
                    dr = dt.Rows(0)
                    iPVID = dt.Rows(0)("PVID")
                End If
                If dt.Rows.Count > 1 Then
                    Common.MessageBox(Me, "儲存資料有誤!(請洽系統管理者)!!")
                    Exit Sub
                End If
            Case cst_ptUpdate
                sql = ""
                sql &= " SELECT * FROM PLAN_VERREPORT "
                sql &= " WHERE PlanID='" & rqPlanID & "' "
                sql &= " AND ComIDNO='" & rqComIDNO & "' "
                sql &= " AND SeqNo='" & rqSeqNO & "' "
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count <> 1 Then
                    Common.MessageBox(Me, "儲存資料有誤!(請洽系統管理者)!!")
                    Exit Sub
                End If
                dr = dt.Rows(0)
                iPVID = dt.Rows(0)("PVID")
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        dr("PlanID") = rqPlanID
        dr("ComIDNO") = rqComIDNO
        dr("SeqNo") = rqSeqNO
        Dim v_rblFuncLevel As String = TIMS.GetListValue(rblFuncLevel)
        dr("FuncLevel") = If(v_rblFuncLevel <> "", v_rblFuncLevel, Convert.DBNull)
        dr("TMethod") = vTMethod
        dr("TMethodOth") = TIMS.ClearSQM(TMethodOth.Text)
        dr("ClassID") = ClassID.SelectedValue
        TIMS.Chk_placeholder(tPOWERNEED1)
        TIMS.Chk_placeholder(tPOWERNEED2)
        TIMS.Chk_placeholder(tPOWERNEED4)
        dr("POWERNEED1") = If(tPOWERNEED1.Text <> "", tPOWERNEED1.Text, Convert.DBNull)
        dr("POWERNEED2") = If(tPOWERNEED2.Text <> "", tPOWERNEED2.Text, Convert.DBNull)
        dr("POWERNEED3") = If(tPOWERNEED3.Text <> "", tPOWERNEED3.Text, Convert.DBNull)
        Dim objD4CHK As Object = Convert.DBNull
        If cbPOWERNEED4.Checked Then objD4CHK = TIMS.cst_YES
        dr("POWERNEED4CHK") = objD4CHK
        If Not cbPOWERNEED4.Checked Then tPOWERNEED4.Text = ""
        dr("POWERNEED4") = If(tPOWERNEED4.Text <> "", tPOWERNEED4.Text, Convert.DBNull)
        dr("PlanCause") = If(tPlanCause.Text <> "", tPlanCause.Text, Convert.DBNull)
        dr("PurScience") = If(tPurScience.Text <> "", tPurScience.Text, Convert.DBNull)
        dr("PurTech") = If(tPurTech.Text <> "", tPurTech.Text, Convert.DBNull)
        dr("PurMoral") = If(tPurMoral.Text <> "", tPurMoral.Text, Convert.DBNull)
        dr("Domain") = Convert.DBNull 'Me.Domain.Text
        dr("CapAll") = If(Me.CapAll.Text <> "", CapAll.Text, Convert.DBNull) 'Me.CapAll.Text

        dr("CostDesc") = If(Me.CostDesc.Text <> "", CostDesc.Text, Convert.DBNull) 'Me.CostDesc.Text
        dr("RecDesc") = If(Me.RecDesc.Text <> "", RecDesc.Text, Convert.DBNull) 'Me.RecDesc.Text
        dr("LearnDesc") = If(Me.LearnDesc.Text <> "", LearnDesc.Text, Convert.DBNull) 'Me.LearnDesc.Text
        dr("ActDesc") = If(Me.ActDesc.Text <> "", ActDesc.Text, Convert.DBNull) ' ActDesc.Text
        dr("ResultDesc") = If(Me.ResultDesc.Text <> "", ResultDesc.Text, Convert.DBNull) 'Me.ResultDesc.Text
        dr("OtherDesc") = If(Me.OtherDesc.Text <> "", OtherDesc.Text, Convert.DBNull) 'Me.OtherDesc.Text

        '是否為iCAP課程 / 是, 請填寫/否/ 課程相關說明
        Dim sISiCAPCOUR As String = ""
        sISiCAPCOUR = If(RB_ISiCAPCOUR_Y.Checked, "Y", If(RB_ISiCAPCOUR_N.Checked, "N", ""))
        dr("ISiCAPCOUR") = If(sISiCAPCOUR <> "", sISiCAPCOUR, Convert.DBNull) '是否為iCAP課程
        iCAPCOURDESC.Text = TIMS.Get_Substr1(TIMS.ClearSQM(iCAPCOURDESC.Text), 500) '課程相關說明
        dr("iCAPCOURDESC") = If(iCAPCOURDESC.Text <> "", iCAPCOURDESC.Text, Convert.DBNull) '(500)
        dr("Recruit") = If(Recruit.Text <> "", Recruit.Text, Convert.DBNull) '招訓方式 Recruit.Text
        dr("Selmethod") = If(Selmethod.Text <> "", Selmethod.Text, Convert.DBNull) '遴選方式
        dr("Inspire") = If(Inspire.Text <> "", Inspire.Text, Convert.DBNull) '學員激勵辦法

        Dim sTGovExamC As String = ""
        If TGovExamCY.Checked Then sTGovExamC = "Y"
        If TGovExamCN.Checked Then sTGovExamC = "N"
        dr("TGovExam") = If(sTGovExamC <> "", sTGovExamC, Convert.DBNull)
        If TGovExamName.Text <> "" Then TGovExamName.Text = Trim(TGovExamName.Text)
        dr("TGovExamName") = If(TGovExamName.Text <> "", TGovExamName.Text, Convert.DBNull)

        dr("memo8") = If(chkMEMO8C1.Checked, cst_msg_memo8a, Convert.DBNull)
        txtMemo8.Text = TIMS.ClearSQM(Me.txtMemo8.Text)
        If chkMEMO8C2.Checked AndAlso txtMemo8.Text = "" Then txtMemo8.Text = " " '若沒有值，輸入一個空白
        dr("memo82") = If(chkMEMO8C2.Checked, txtMemo8.Text, Convert.DBNull)

        'If sender.text.ToString=cst_DefSave Then dr("IsApprPaper")="N" Else dr("IsApprPaper")="Y"
        Select Case SaveType1
            Case cst_SaveDef
                dr("IsApprPaper") = "N"
            Case Else
                dr("IsApprPaper") = "Y"
        End Select
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)

    End Sub

    Function CheckData1(ByRef ErrMsg As String) As Boolean
        '#Region "CheckData1"

        Dim rst As Boolean = False
        RecDesc.Text = TIMS.ClearSQM(RecDesc.Text)
        LearnDesc.Text = TIMS.ClearSQM(LearnDesc.Text)
        ActDesc.Text = TIMS.ClearSQM(ActDesc.Text)
        ResultDesc.Text = TIMS.ClearSQM(ResultDesc.Text)
        OtherDesc.Text = TIMS.ClearSQM(OtherDesc.Text)

        Select Case Hid_sender1.Value
            Case cst_SaveDef
                Return True
        End Select

        '正式檢測。
        If RecDesc.Text = "" AndAlso chk_RecDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 反應評估，請輸入內容" & vbCrLf
        If LearnDesc.Text = "" AndAlso chk_LearnDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 學習評估，請輸入內容" & vbCrLf
        If ActDesc.Text = "" AndAlso chk_ActDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 行為評估，請輸入內容" & vbCrLf
        If ResultDesc.Text = "" AndAlso chk_ResultDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 成果評估，請輸入內容" & vbCrLf
        If OtherDesc.Text = "" AndAlso chk_OtherDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 其他機制，請輸入內容" & vbCrLf
        If ErrMsg <> "" Then Return False

        If RecDesc.Text <> "" AndAlso Not chk_RecDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 反應評估，請勿輸入內容" & vbCrLf
        If LearnDesc.Text <> "" AndAlso Not chk_LearnDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 學習評估，請勿輸入內容" & vbCrLf
        If ActDesc.Text <> "" AndAlso Not chk_ActDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 行為評估，請勿輸入內容" & vbCrLf
        If ResultDesc.Text <> "" AndAlso Not chk_ResultDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 成果評估，請勿輸入內容" & vbCrLf
        If OtherDesc.Text <> "" AndAlso Not chk_OtherDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 其他機制，請勿輸入內容" & vbCrLf
        If ErrMsg <> "" Then Return False

        Dim i_chk2 As Integer = 0
        If chk_RecDesc.Checked Then i_chk2 += 1
        If chk_LearnDesc.Checked Then i_chk2 += 1
        If chk_ActDesc.Checked Then i_chk2 += 1
        If chk_ResultDesc.Checked Then i_chk2 += 1
        If chk_OtherDesc.Checked Then i_chk2 += 1
        If i_chk2 = 0 Then ErrMsg &= "四、訓練績效評估-未勾選 (至少要勾選一項)" & vbCrLf
        If ErrMsg <> "" Then Return False

        Dim flagCPT As Boolean = CHK_PLAN_TEACHER12(ErrMsg)
        If Not flagCPT Then Return False

        Dim vTMethod As String = TIMS.GetCblValue(cblTMethod)
        If vTMethod.IndexOf("99") > -1 AndAlso TMethodOth.Text = "" Then ErrMsg &= "教學方法-若選「其他教學方法」，需填寫輸入其它說明，上限100個字" & vbCrLf
        If cbPOWERNEED4.Checked AndAlso tPOWERNEED4.Text = "" Then ErrMsg &= "訓練需求調查 -若勾選「課程須符合目的事業主管機關相關規定」，需填寫它說明，上限200個字" & vbCrLf
        If CapAll.Text = "" Then ErrMsg &= "未輸入學員資格，請確認" & vbCrLf
        If Inspire.Text = "" Then ErrMsg &= "未輸入學員激勵辦法，請確認" & vbCrLf

        If ErrMsg = "" Then rst = True
        Return rst


    End Function

    '草稿儲存/正式儲存
    Private Sub bt_addrow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_addrow.Click, Button6.Click
        '#Region "草稿儲存/正式儲存"

        Hid_sender1.Value = sender.text
        Dim sErrMsg As String = ""
        Dim rst As Boolean = CheckData1(sErrMsg)

        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        Select Case sender.text.ToString
            Case cst_SaveDef '草稿
                '儲存 開班計畫表資料維護
                Call SAVE_PLAN_VERREPORT(cst_SaveDef)
            Case cst_SaveRcc '正式
                '儲存 開班計畫表資料維護
                Call SAVE_PLAN_VERREPORT(cst_SaveRcc)
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        '儲存 班級申請老師
        Call SAVE_PLAN_TEACHER(objconn)

        Dim drPP As DataRow = TIMS.GetPCSDate(rqPlanID, rqComIDNO, rqSeqNO, objconn)
        If Convert.ToString(drPP("OCID")) <> "" Then Call SAVE_CLASS_TEACHER(drPP("OCID"), objconn) '修改班級師資資料
        Session("_search") = ViewState("_search")
        Common.RespWrite(Me, "<script>alert('儲存成功!!');</script>")
        Common.RespWrite(Me, "<script>location.href='TC_01_014.aspx?ID=" & Request("ID") & "'</script>")


    End Sub

    '回上一頁
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '#Region "回上一頁"

        Session("_search") = ViewState("_search")
        Dim url1 As String = "TC_01_014.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)


    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        '#Region "DataGrid1_ItemDataBound"

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Weeks As Label = e.Item.FindControl("Weeks1")
                Dim Times As Label = e.Item.FindControl("Times1")
                Dim drv As DataRowView = e.Item.DataItem
                Weeks.Text = drv("Weeks").ToString
                Times.Text = drv("Times").ToString
            Case ListItemType.EditItem
                Dim Weeks As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times As TextBox = e.Item.FindControl("Times2")
                Dim drv As DataRowView = e.Item.DataItem
                With Weeks
                    .Items.Add(New ListItem("==請選擇==", ""))
                    .Items.Add(New ListItem("星期一", "星期一"))
                    .Items.Add(New ListItem("星期二", "星期二"))
                    .Items.Add(New ListItem("星期三", "星期三"))
                    .Items.Add(New ListItem("星期四", "星期四"))
                    .Items.Add(New ListItem("星期五", "星期五"))
                    .Items.Add(New ListItem("星期六", "星期六"))
                    .Items.Add(New ListItem("星期日", "星期日"))
                End With
                Common.SetListItem(Weeks, drv("Weeks").ToString)
                Times.Text = drv("Times").ToString
        End Select


    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        '#Region "Datagrid3_ItemDataBound"

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim STrainDateLabel As Label = e.Item.FindControl("STrainDateLabel")
                Dim PNameLabel As Label = e.Item.FindControl("PNameLabel")
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel")
                Dim PContText As TextBox = e.Item.FindControl("PContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1")
                Dim drpPTID As DropDownList = e.Item.FindControl("drpPTID")
                Dim Tech1Value As HtmlInputHidden = e.Item.FindControl("Tech1Value")
                Dim Tech1Text As TextBox = e.Item.FindControl("Tech1Text")
                Dim Tech2Value As HtmlInputHidden = e.Item.FindControl("Tech2Value")
                Dim Tech2Text As TextBox = e.Item.FindControl("Tech2Text")
                If drv("STrainDate").ToString <> "" Then STrainDateLabel.Text = Common.FormatDate(drv("STrainDate").ToString)
                PNameLabel.Text = drv("PName").ToString
                PHourLabel.Text = drv("PHour").ToString
                PContText.Text = TIMS.HtmlDecode1(Convert.ToString(drv("PCont")))

                If drv("Classification1").ToString <> "" Then
                    Common.SetListItem(drpClassification1, drv("Classification1").ToString)
                    Select Case drpClassification1.SelectedValue
                        Case "1" '學科
                            If rqComIDNO = "" Then
                                If Hid_COMIDNO.Value = "" Then
                                    Hid_COMIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                                End If
                                drpPTID = TIMS.Get_SciPTID(drpPTID, Hid_COMIDNO.Value, 1, objconn)
                            Else
                                drpPTID = TIMS.Get_SciPTID(drpPTID, rqComIDNO, 1, objconn)
                            End If
                        Case "2" '術科
                            If rqComIDNO = "" Then
                                If Hid_COMIDNO.Value = "" Then
                                    Hid_COMIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                                End If
                                drpPTID = TIMS.Get_TechPTID(drpPTID, Hid_COMIDNO.Value, 1, objconn)
                            Else
                                drpPTID = TIMS.Get_TechPTID(drpPTID, rqComIDNO, 1, objconn)
                            End If
                    End Select
                    If drv("PTID").ToString <> "" Then Common.SetListItem(drpPTID, drv("PTID").ToString)
                End If

                If drv("TechID").ToString <> "" Then
                    Tech1Value.Value = drv("TechID").ToString
                    Tech1Text.Text = TIMS.Get_TeachCName(Tech1Value.Value, objconn) 'TIMS.Get_TeacherName(drv("TechID").ToString)
                End If

                If Convert.ToString(drv("TechID2")) <> "" Then
                    Tech2Value.Value = drv("TechID2").ToString
                    Tech2Text.Text = TIMS.Get_TeachCName(Tech2Value.Value, objconn) 'TIMS.Get_TeacherName(drv("TechID2").ToString)
                End If
        End Select


    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        '#Region "DataGrid2_ItemDataBound"

        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem

                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim ProLicense As Label = e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEA As HtmlInputButton = e.Item.FindControl("btn_TCTYPEA")
                sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                btn_TCTYPEA.Attributes("onclick") = sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                iSeqno += 1
                seqno.Text = iSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                ProLicense.Text = Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                TeacherDesc.ReadOnly = False
                btn_TCTYPEA.Visible = True

                Select Case rqProcessType 'ProcessType @Insert/Update/View
                    Case cst_ptView '查詢功能不提供儲存
                        TeacherDesc.ReadOnly = True
                        btn_TCTYPEA.Visible = False
                End Select
        End Select


    End Sub

    Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged
    End Sub

    Private Sub DataGrid22_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid22.ItemDataBound
        '#Region "DataGrid22_ItemDataBound"

        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem

                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim ProLicense As Label = e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEB As HtmlInputButton = e.Item.FindControl("btn_TCTYPEB")
                sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                btn_TCTYPEB.Attributes("onclick") = sWOScript1
                HidTechID.Value = Convert.ToString(drv("TechID"))
                iSeqno += 1
                seqno.Text = iSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                ProLicense.Text = Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                TeacherDesc.ReadOnly = False
                btn_TCTYPEB.Visible = True

                Select Case rqProcessType 'ProcessType @Insert/Update/View
                    Case cst_ptView '查詢功能不提供儲存
                        TeacherDesc.ReadOnly = True
                        btn_TCTYPEB.Visible = False
                End Select
        End Select


    End Sub

    Protected Sub DataGrid22_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid22.SelectedIndexChanged
    End Sub
End Class