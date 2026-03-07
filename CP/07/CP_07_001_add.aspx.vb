Partial Class CP_07_001_add
    Inherits AuthBasePage

    'ALTER TABLE STUD_QUESTRAINING ADD [Q2_4] [numeric] (10,0)  NULL 
    'ALTER TABLE STUD_QUESTRAINING alter column  [SIGNER] [nvarchar](30) COLLATE Chinese_Taiwan_Stroke_CS_AS NULL
    Dim rRID As String = ""
    Dim rOCID As String = ""
    Dim rSOCID As String = "" '= Request("socid")
    Dim rPlanID As String = "" '= Request("PlanID")
    Dim rStatus As String = "" '= Request("status") 'edit,add
    'Dim rType As String = "" '= Request("Type") '1,2
    'Dim rInquireType As String = "" '= Request("InquireType") '1,2

    Const cst_SearchStr As String = "CP_07_001_SearchStr"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        rOCID = TIMS.ClearSQM(Request("ocid"))
        rSOCID = TIMS.ClearSQM(Request("socid"))
        rRID = TIMS.ClearSQM(Request("rid"))

        rPlanID = TIMS.ClearSQM(Request("PlanID"))
        rStatus = TIMS.ClearSQM(Request("status")) 'edit,add
        'rType = Request("Type") '1,2
        'rInquireType = Request("InquireType") '1,2
        'InquireType = Convert.ToInt16(Request("InquireType"))

        If Session(cst_SearchStr) IsNot Nothing Then
            'Session(cst_SearchStr) = Session(cst_SearchStr)
            ViewState(cst_SearchStr) = Session(cst_SearchStr)
        Else
            If Session(cst_SearchStr) Is Nothing AndAlso ViewState(cst_SearchStr) IsNot Nothing Then
                Session(cst_SearchStr) = Me.ViewState(cst_SearchStr)
            End If
        End If

        If Not IsPostBack Then
            cCreate1()
        End If

    End Sub

    Sub cCreate1()
        OCIDValue1.Value = rOCID
        Hid_socid.Value = rSOCID
        RIDValue.Value = rRID

        txt_Suggestion.Text = ""
        bt_save.Attributes.Add("onclick", "return confirm('確定儲存?');")
        'bt_save.Attributes.Add("onclick", "return chkinput();")
        'dl_InquireType.SelectedValue = Convert.ToInt16(Request("InquireType"))  '系統登打：1 、電話訪查：2 
        'If Convert.ToInt16(Request("InquireType")) <> 2 Then
        '    tr_1.Visible = False
        'Else
        '    tr_1.Visible = True
        'End If
        Call sub_ClearData1()
        Call sub_LoadData1()

        'rStatus status : add/edit/check/clear （新增/修改/查看/清除重填）
        If rStatus = "check" Then Call sub_check1()

        'chkInquireType()
        'txFillDate.Text = TIMS.cdate3(TIMS.GetSysDate(objconn))
    End Sub

    Sub sub_check1()
        'rStatus status : add/edit/check/clear （新增/修改/查看/清除重填）
        If rStatus <> "check" Then Return
        Dim s_t1 As String = "僅供查詢"
        bt_save.Visible = False
        txFillDate.Enabled = False
        TIMS.Tooltip(txFillDate, s_t1)

        span_FillDate.Visible = False
        signer.Enabled = False
        TIMS.Tooltip(signer, s_t1)

        txt_Suggestion.Enabled = False
        TIMS.Tooltip(txt_Suggestion, s_t1)

        'tb_Ques1.Disabled = True
        'TIMS.Tooltip(tb_Ques1, s_t1)

        Q1_1.Enabled = False
        Q1_2.Enabled = False
        Q1_3.Enabled = False
        Q2_1.Enabled = False
        Q2_2.Enabled = False
        Q2_3.Enabled = False
        Q2_4.Enabled = False
        Q3_1.Enabled = False
        Q3_2.Enabled = False
        Q3_3.Enabled = False
        Q4_1.Enabled = False
        Q4_2.Enabled = False
        Q4_3.Enabled = False
        Q4_4.Enabled = False

        TIMS.Tooltip(Q1_1, s_t1)
        TIMS.Tooltip(Q1_2, s_t1)
        TIMS.Tooltip(Q1_3, s_t1)
        TIMS.Tooltip(Q2_1, s_t1)
        TIMS.Tooltip(Q2_2, s_t1)
        TIMS.Tooltip(Q2_3, s_t1)
        TIMS.Tooltip(Q2_4, s_t1)
        TIMS.Tooltip(Q3_1, s_t1)
        TIMS.Tooltip(Q3_2, s_t1)
        TIMS.Tooltip(Q3_3, s_t1)
        TIMS.Tooltip(Q4_1, s_t1)
        TIMS.Tooltip(Q4_2, s_t1)
        TIMS.Tooltip(Q4_3, s_t1)
        TIMS.Tooltip(Q4_4, s_t1)
    End Sub

    Sub sub_ClearData1()
        txFillDate.Text = TIMS.Cdate3(TIMS.GetSysDate(objconn))
        Q1_1.SelectedIndex = -1
        Q1_2.SelectedIndex = -1
        Q1_3.SelectedIndex = -1
        Q2_1.SelectedIndex = -1
        Q2_2.SelectedIndex = -1
        Q2_3.SelectedIndex = -1
        Q2_4.SelectedIndex = -1

        Q3_1.SelectedIndex = -1
        Q3_2.SelectedIndex = -1
        Q3_3.SelectedIndex = -1
        Q4_1.SelectedIndex = -1
        Q4_2.SelectedIndex = -1
        Q4_3.SelectedIndex = -1
        Q4_4.SelectedIndex = -1
        txt_Suggestion.Text = ""
    End Sub

    '呼叫單筆資料。。
    Sub sub_LoadData1()
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then Return '異常

        'lb_STDate.Text = Convert.ToString(Convert.ToDateTime().Year()) & "年" & Convert.ToString(Convert.ToDateTime(drCC("STDate")).Month()) & "月" & Convert.ToString(Convert.ToDateTime(drCC("STDate")).Day()) & "日"
        'lb_FTDate.Text = Convert.ToString(Convert.ToDateTime(drCC("FTDate")).Year()) & "年" & Convert.ToString(Convert.ToDateTime(drCC("FTDate")).Month()) & "月" & Convert.ToString(Convert.ToDateTime(drCC("FTDate")).Day()) & "日"
        lb_STDate.Text = String.Format("{0}年{1}月{2}日", CDate(drCC("STDate")).ToString("yyyy"), CDate(drCC("STDate")).ToString("MM"), CDate(drCC("STDate")).ToString("dd"))
        lb_FTDate.Text = String.Format("{0}年{1}月{2}日", CDate(drCC("FTDate")).ToString("yyyy"), CDate(drCC("FTDate")).ToString("MM"), CDate(drCC("FTDate")).ToString("dd"))
        lb_OrgName.Text = Convert.ToString(drCC("OrgName"))
        lb_OCID.Text = Convert.ToString(drCC("CLASSCNAME2"))
        'studID.Text = Convert.ToString(dr("StudentID")).ToUpper()
        lb_PlanName.Text = TIMS.GetPlanName(drCC("PlanID").ToString(), objconn)
        'lb_DistID.Text = Convert.ToString(dr("DistIDName"))

        Dim drSS As DataRow = TIMS.Get_StudData(rSOCID, objconn)
        If drSS Is Nothing Then Return '異常
        lb_STDNAME.Text = String.Format("{0}({1})", drSS("STDNAME"), drSS("STUDID2"))

        If Convert.ToString(drSS("ocid")) <> OCIDValue1.Value Then Return '異常

        'rStatus status : add/edit/check/clear （新增/修改/查看/清除重填）
        If rStatus = "clear" Then
            'STUD_QUESTRAININGDEL
            Dim dt2 As DataTable
            Dim parms2 As New Hashtable
            parms2.Add("SOCID", rSOCID)
            Dim sql2 As String = ""
            sql2 = "SELECT * FROM STUD_QUESTRAINING WHERE SOCID=@SOCID"
            dt2 = DbAccess.GetDataTable(sql2, objconn, parms2)
            If dt2.Rows.Count <> 1 Then Return
            Call TIMS.InsertDelTableLog("STUD_QUESTRAININGDEL", dt2, objconn)
            Dim parms3 As New Hashtable
            parms3.Add("SOCID", rSOCID)
            Dim sql3 As String = ""
            sql3 = "DELETE STUD_QUESTRAINING WHERE SOCID=@SOCID"
            DbAccess.ExecuteNonQuery(sql3, objconn, parms3)
            Return
        End If

        Dim parms As New Hashtable
        parms.Add("PlanID", rPlanID)
        parms.Add("OCID", OCIDValue1.Value)
        'parms.Add("RID", RIDValue.Value)
        parms.Add("SOCID", Hid_socid.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select ip.YEARS" & vbCrLf
        sql &= " ,ip.DISTNAME" & vbCrLf
        sql &= " ,cc.PlanName" & vbCrLf
        sql &= " ,cc.ORGNAME" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.CLASSCNAME2" & vbCrLf
        'sql &= " ,dbo.FN_CDATE(cc.STDATE)  CSTDATE" & vbCrLf
        'sql &= " ,dbo.FN_CDATE(cc.FTDATE)  CFTDATE" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,cs.SOCID" & vbCrLf
        'sql &= " ,dbo.FN_CDATE(g.FILLFORMDATE)  CFILLFORMDATE" & vbCrLf
        sql &= " ,format(g.FILLFORMDATE,'yyyy/MM/dd') FILLFORMDATE" & vbCrLf
        'sql &= " ,g.OCID" & vbCrLf
        sql &= " ,g.SOCID  queSOCID" & vbCrLf '/*PK*/"
        'sql &= " ,g.RID" & vbCrLf
        'sql &= " ,g.FILLFORMDATE" & vbCrLf
        'sql &= " ,g.TYPE" & vbCrLf
        sql &= " ,g.Q1_1" & vbCrLf
        sql &= " ,g.Q1_2" & vbCrLf
        sql &= " ,g.Q1_3" & vbCrLf
        sql &= " ,g.Q2_1" & vbCrLf
        sql &= " ,g.Q2_2" & vbCrLf
        sql &= " ,g.Q2_3" & vbCrLf
        sql &= " ,g.Q2_4" & vbCrLf
        sql &= " ,g.Q3_1" & vbCrLf
        sql &= " ,g.Q3_2" & vbCrLf
        sql &= " ,g.Q3_3" & vbCrLf
        sql &= " ,g.Q4_1" & vbCrLf
        sql &= " ,g.Q4_2" & vbCrLf
        sql &= " ,g.Q4_3" & vbCrLf
        sql &= " ,g.Q4_4" & vbCrLf
        sql &= " ,g.SUGGESTION" & vbCrLf
        sql &= " ,g.MODIFYACCT" & vbCrLf
        sql &= " ,g.MODIFYDATE" & vbCrLf
        sql &= " ,g.SIGNER" & vbCrLf
        '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
        sql &= " ,g.DASOURCE" & vbCrLf
        sql &= " from VIEW2 cc" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.PLANID=cc.PLANID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.ocid =cc.ocid" & vbCrLf
        sql &= " LEFT JOIN STUD_QUESTRAINING g on g.socid=cs.socid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql &= " and ip.tplanid ='06'" & vbCrLf
        'sql &= " and cc.ocid =111424" & vbCrLf
        sql &= " and ip.PlanID=@PlanID "
        sql &= " and cc.OCID=@OCID "
        'sql &= " and cc.RID=@RID "
        sql &= " and cs.SOCID=@SOCID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then Return
        Dim dr As DataRow = dt.Rows(0)

        txFillDate.Text = TIMS.Cdate3(dr("FillFormDate"))
        Common.SetListItem(Q1_1, dr("Q1_1"))
        Common.SetListItem(Q1_2, dr("Q1_2"))
        Common.SetListItem(Q1_3, dr("Q1_3"))
        Common.SetListItem(Q2_1, dr("Q2_1"))
        Common.SetListItem(Q2_2, dr("Q2_2"))
        Common.SetListItem(Q2_3, dr("Q2_3"))

        Common.SetListItem(Q2_4, dr("Q2_4"))

        Common.SetListItem(Q3_1, dr("Q3_1"))
        Common.SetListItem(Q3_2, dr("Q3_2"))
        Common.SetListItem(Q3_3, dr("Q3_3"))

        Common.SetListItem(Q4_1, dr("Q4_1"))
        Common.SetListItem(Q4_2, dr("Q4_2"))
        Common.SetListItem(Q4_3, dr("Q4_3"))
        Common.SetListItem(Q4_4, dr("Q4_4"))

        'Q1_1.SelectedValue = Convert.ToInt16(dr("Q1_1"))
        'Q1_2.SelectedValue = Convert.ToInt16(dr("Q1_2"))
        'Q1_3.SelectedValue = Convert.ToInt16(dr("Q1_3"))
        'Q2_1.SelectedValue = Convert.ToInt16(dr("Q2_1"))
        'Q2_2.SelectedValue = Convert.ToInt16(dr("Q2_2"))
        'Q2_3.SelectedValue = Convert.ToInt16(dr("Q2_3"))
        'Q3_1.SelectedValue = Convert.ToInt16(dr("Q3_1"))
        'Q3_2.SelectedValue = Convert.ToInt16(dr("Q3_2"))
        'Q3_3.SelectedValue = Convert.ToInt16(dr("Q3_3"))
        'Q4_1.SelectedValue = Convert.ToInt16(dr("Q4_1"))
        'Q4_2.SelectedValue = Convert.ToInt16(dr("Q4_2"))
        'Q4_3.SelectedValue = Convert.ToInt16(dr("Q4_3"))
        'Q4_4.SelectedValue = Convert.ToInt16(dr("Q4_4"))
        'dl_InquireType.SelectedValue = IIf(IsDBNull(dr("TYPE")), "0", Convert.ToString(dr("TYPE")))
        signer.Text = Convert.ToString(dr("signer"))
        txt_Suggestion.Text = Convert.ToString(dr("Suggestion"))

        'lb_FillDate.Text = Convert.ToString(Convert.ToDateTime(dr("FillFormDate")).Year()) & "年" & Convert.ToString(Convert.ToDateTime(dr("FillFormDate")).Month()) & "月" & Convert.ToString(Convert.ToDateTime(dr("FillFormDate")).Day()) & "日"
        ''WeekDay.Text = ConvertNum(Convert.ToString(Convert.ToDateTime(dr("FillFormDate")).DayOfWeek()))
        'WeekDay.Text = TIMS.GetWeekDay(CDate(dr("FillFormDate")).DayOfWeek)
        'msg.Text = "" '使用 UPDATE

    End Sub

    '儲存。
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        Dim errMsg As String = ""
        If Not chkInputAnswer(errMsg) Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If
        'status : add/edit/check/clear （新增/修改/查看/清除重填）
        Call savedata()
    End Sub

    '儲存[SQL]
    Private Sub savedata()
        'status : add/edit/check/clear （新增/修改/查看/清除重填）
        Dim flag_can_save As Boolean = False
        If rStatus = "add" OrElse rStatus = "edit" Then flag_can_save = True
        If Not flag_can_save Then Return

        'Dim iType As String = "1" '1.系統登打2.電話訪查
        'iType = CInt(dl_InquireType.SelectedValue)
        txt_Suggestion.Text = TIMS.ClearSQM2(txt_Suggestion.Text)
        signer.Text = TIMS.ClearSQM(signer.Text)

        Dim Q1_1_selval As String = TIMS.GetListValue(Q1_1)
        Dim Q1_2_selval As String = TIMS.GetListValue(Q1_2)
        Dim Q1_3_selval As String = TIMS.GetListValue(Q1_3)
        Dim Q2_1_selval As String = TIMS.GetListValue(Q2_1)
        Dim Q2_2_selval As String = TIMS.GetListValue(Q2_2)
        Dim Q2_3_selval As String = TIMS.GetListValue(Q2_3)
        Dim Q2_4_selval As String = TIMS.GetListValue(Q2_4)
        Dim Q3_1_selval As String = TIMS.GetListValue(Q3_1)
        Dim Q3_2_selval As String = TIMS.GetListValue(Q3_2)
        Dim Q3_3_selval As String = TIMS.GetListValue(Q3_3)
        Dim Q4_1_selval As String = TIMS.GetListValue(Q4_1)
        Dim Q4_2_selval As String = TIMS.GetListValue(Q4_2)
        Dim Q4_3_selval As String = TIMS.GetListValue(Q4_3)
        Dim Q4_4_selval As String = TIMS.GetListValue(Q4_4)

        Dim dr1 As DataRow = Nothing
        '取得一筆資料。
        dr1 = TIMS.GetData1_SQ1(Hid_socid.Value, objconn)

        'Dim sql As String = ""
        'Dim dt As DataTable = Nothing
        'Dim oCmd As SqlCommand = Nothing
        'sql = "SELECT 'X' FROM STUD_QUESTRAINING WHERE SOCID =@SOCID "
        'Call TIMS.OpenDbConn(objconn)
        'dt = New DataTable
        'oCmd = New SqlCommand(sql, objconn)
        'With oCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("SOCID", SqlDbType.Int).Value = rSOCID
        '    dt.Load(.ExecuteReader())
        'End With
        'msg1.Text = "尚未填寫！" '使用新增功能
        'If dt.Rows.Count > 0 Then msg1.Text = ""

        If dr1 Is Nothing Then
            Dim i_sql As String = ""
            i_sql = ""
            i_sql &= " INSERT INTO STUD_QUESTRAINING (" & vbCrLf
            i_sql &= " OCID" & vbCrLf
            i_sql &= " ,SOCID" & vbCrLf
            i_sql &= " ,RID" & vbCrLf
            i_sql &= " ,FILLFORMDATE" & vbCrLf
            i_sql &= " ,TYPE" & vbCrLf
            i_sql &= " ,Q1_1" & vbCrLf
            i_sql &= " ,Q1_2" & vbCrLf
            i_sql &= " ,Q1_3" & vbCrLf
            i_sql &= " ,Q2_1" & vbCrLf
            i_sql &= " ,Q2_2" & vbCrLf
            i_sql &= " ,Q2_3" & vbCrLf
            i_sql &= " ,Q2_4" & vbCrLf
            i_sql &= " ,Q3_1" & vbCrLf
            i_sql &= " ,Q3_2" & vbCrLf
            i_sql &= " ,Q3_3" & vbCrLf
            i_sql &= " ,Q4_1" & vbCrLf
            i_sql &= " ,Q4_2" & vbCrLf
            i_sql &= " ,Q4_3" & vbCrLf
            i_sql &= " ,Q4_4" & vbCrLf
            i_sql &= " ,SUGGESTION" & vbCrLf
            i_sql &= " ,SIGNER" & vbCrLf
            '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
            i_sql &= " ,DASOURCE" & vbCrLf
            i_sql &= " ,MODIFYACCT" & vbCrLf
            i_sql &= " ,MODIFYDATE" & vbCrLf
            i_sql &= " ) VALUES (" & vbCrLf
            i_sql &= " @OCID" & vbCrLf
            i_sql &= " ,@SOCID" & vbCrLf
            i_sql &= " ,@RID" & vbCrLf
            i_sql &= " ,GETDATE() " & vbCrLf 'FILLFORMDATE
            i_sql &= " ,@TYPE" & vbCrLf
            i_sql &= " ,@Q1_1" & vbCrLf
            i_sql &= " ,@Q1_2" & vbCrLf
            i_sql &= " ,@Q1_3" & vbCrLf
            i_sql &= " ,@Q2_1" & vbCrLf
            i_sql &= " ,@Q2_2" & vbCrLf
            i_sql &= " ,@Q2_3" & vbCrLf
            i_sql &= " ,@Q2_4" & vbCrLf
            i_sql &= " ,@Q3_1" & vbCrLf
            i_sql &= " ,@Q3_2" & vbCrLf
            i_sql &= " ,@Q3_3" & vbCrLf
            i_sql &= " ,@Q4_1" & vbCrLf
            i_sql &= " ,@Q4_2" & vbCrLf
            i_sql &= " ,@Q4_3" & vbCrLf
            i_sql &= " ,@Q4_4" & vbCrLf
            i_sql &= " ,@SUGGESTION" & vbCrLf
            i_sql &= " ,@SIGNER" & vbCrLf
            '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
            i_sql &= " ,@DASOURCE" & vbCrLf
            i_sql &= " ,@MODIFYACCT" & vbCrLf
            i_sql &= " ,GETDATE()" & vbCrLf
            i_sql &= " )" & vbCrLf

            Call TIMS.OpenDbConn(objconn)
            Dim iCmd As New SqlCommand(i_sql, objconn)
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.Int).Value = CInt(OCIDValue1.Value)
                .Parameters.Add("SOCID", SqlDbType.Int).Value = CInt(Hid_socid.Value)
                .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                .Parameters.Add("FILLFORMDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(txFillDate.Text)
                .Parameters.Add("TYPE", SqlDbType.Int).Value = 1 '1.系統登打 2.電話訪查 3.報名網寫入

                .Parameters.Add("Q1_1", SqlDbType.Int).Value = If(Q1_1_selval <> "", Q1_1_selval, Convert.DBNull) '.SelectedValue
                .Parameters.Add("Q1_2", SqlDbType.Int).Value = If(Q1_2_selval <> "", Q1_2_selval, Convert.DBNull) '
                .Parameters.Add("Q1_3", SqlDbType.Int).Value = If(Q1_3_selval <> "", Q1_3_selval, Convert.DBNull) '

                .Parameters.Add("Q2_1", SqlDbType.Int).Value = If(Q2_1_selval <> "", Q2_1_selval, Convert.DBNull) '
                .Parameters.Add("Q2_2", SqlDbType.Int).Value = If(Q2_2_selval <> "", Q2_2_selval, Convert.DBNull) '
                .Parameters.Add("Q2_3", SqlDbType.Int).Value = If(Q2_3_selval <> "", Q2_3_selval, Convert.DBNull) '
                .Parameters.Add("Q2_4", SqlDbType.Int).Value = If(Q2_4_selval <> "", Q2_4_selval, Convert.DBNull) '

                .Parameters.Add("Q3_1", SqlDbType.Int).Value = If(Q3_1_selval <> "", Q3_1_selval, Convert.DBNull) '
                .Parameters.Add("Q3_2", SqlDbType.Int).Value = If(Q3_2_selval <> "", Q3_2_selval, Convert.DBNull) '
                .Parameters.Add("Q3_3", SqlDbType.Int).Value = If(Q3_3_selval <> "", Q3_3_selval, Convert.DBNull) '

                .Parameters.Add("Q4_1", SqlDbType.Int).Value = If(Q4_1_selval <> "", Q4_1_selval, Convert.DBNull) '
                .Parameters.Add("Q4_2", SqlDbType.Int).Value = If(Q4_2_selval <> "", Q4_2_selval, Convert.DBNull) '
                .Parameters.Add("Q4_3", SqlDbType.Int).Value = If(Q4_3_selval <> "", Q4_3_selval, Convert.DBNull) '
                .Parameters.Add("Q4_4", SqlDbType.Int).Value = If(Q4_4_selval <> "", Q4_4_selval, Convert.DBNull) '
                '貳、其它意見
                .Parameters.Add("SUGGESTION", SqlDbType.NVarChar).Value = If(txt_Suggestion.Text <> "", txt_Suggestion.Text, Convert.DBNull)  'SUGGESTIO 
                '抽訪人員姓名 簽章
                .Parameters.Add("SIGNER", SqlDbType.VarChar).Value = If(signer.Text <> "", signer.Text, Convert.DBNull) 'SIGNER
                '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統後台()
                .Parameters.Add("DASOURCE", SqlDbType.Int).Value = 2 'SIGNERFROM
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .ExecuteNonQuery()
            End With
        Else
            Dim u_sql As String = ""
            u_sql = ""
            u_sql &= " UPDATE STUD_QUESTRAINING" & vbCrLf
            u_sql &= " SET RID= @RID" & vbCrLf
            u_sql &= " ,FILLFORMDATE= @FILLFORMDATE" & vbCrLf
            u_sql &= " ,TYPE=@TYPE" & vbCrLf
            u_sql &= " ,Q1_1=@Q1_1" & vbCrLf
            u_sql &= " ,Q1_2=@Q1_2" & vbCrLf
            u_sql &= " ,Q1_3=@Q1_3" & vbCrLf
            u_sql &= " ,Q2_1=@Q2_1" & vbCrLf
            u_sql &= " ,Q2_2=@Q2_2" & vbCrLf
            u_sql &= " ,Q2_3=@Q2_3" & vbCrLf
            u_sql &= " ,Q2_4=@Q2_4" & vbCrLf
            u_sql &= " ,Q3_1=@Q3_1" & vbCrLf
            u_sql &= " ,Q3_2=@Q3_2" & vbCrLf
            u_sql &= " ,Q3_3=@Q3_3" & vbCrLf
            u_sql &= " ,Q4_1=@Q4_1" & vbCrLf
            u_sql &= " ,Q4_2=@Q4_2" & vbCrLf
            u_sql &= " ,Q4_3=@Q4_3" & vbCrLf
            u_sql &= " ,Q4_4=@Q4_4" & vbCrLf
            u_sql &= " ,SUGGESTION=@SUGGESTION" & vbCrLf
            u_sql &= " ,SIGNER=@SIGNER" & vbCrLf
            u_sql &= " ,DASOURCE=@DASOURCE" & vbCrLf
            u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            u_sql &= " ,MODIFYDATE= GETDATE()" & vbCrLf
            u_sql &= " WHERE OCID=@OCID AND SOCID=@SOCID" & vbCrLf
            Call TIMS.OpenDbConn(objconn)
            Dim uCmd As New SqlCommand(u_sql, objconn)

            With uCmd
                .Parameters.Clear()
                '.Parameters.Add("FILLFORMDATE", SqlDbType.DateTime).Value = FILLFORMDATE
                '.Parameters.Add("TYPE", SqlDbType.Int).Value = iType
                .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                .Parameters.Add("FILLFORMDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(txFillDate.Text)
                .Parameters.Add("TYPE", SqlDbType.Int).Value = 1 '1.系統登打2.電話訪查

                .Parameters.Add("Q1_1", SqlDbType.Int).Value = If(Q1_1_selval <> "", Q1_1_selval, Convert.DBNull) '.SelectedValue
                .Parameters.Add("Q1_2", SqlDbType.Int).Value = If(Q1_2_selval <> "", Q1_2_selval, Convert.DBNull) '
                .Parameters.Add("Q1_3", SqlDbType.Int).Value = If(Q1_3_selval <> "", Q1_3_selval, Convert.DBNull) '

                .Parameters.Add("Q2_1", SqlDbType.Int).Value = If(Q2_1_selval <> "", Q2_1_selval, Convert.DBNull) '
                .Parameters.Add("Q2_2", SqlDbType.Int).Value = If(Q2_2_selval <> "", Q2_2_selval, Convert.DBNull) '
                .Parameters.Add("Q2_3", SqlDbType.Int).Value = If(Q2_3_selval <> "", Q2_3_selval, Convert.DBNull) '
                .Parameters.Add("Q2_4", SqlDbType.Int).Value = If(Q2_4_selval <> "", Q2_4_selval, Convert.DBNull) '

                .Parameters.Add("Q3_1", SqlDbType.Int).Value = If(Q3_1_selval <> "", Q3_1_selval, Convert.DBNull) '
                .Parameters.Add("Q3_2", SqlDbType.Int).Value = If(Q3_2_selval <> "", Q3_2_selval, Convert.DBNull) '
                .Parameters.Add("Q3_3", SqlDbType.Int).Value = If(Q3_3_selval <> "", Q3_3_selval, Convert.DBNull) '

                .Parameters.Add("Q4_1", SqlDbType.Int).Value = If(Q4_1_selval <> "", Q4_1_selval, Convert.DBNull) '
                .Parameters.Add("Q4_2", SqlDbType.Int).Value = If(Q4_2_selval <> "", Q4_2_selval, Convert.DBNull) '
                .Parameters.Add("Q4_3", SqlDbType.Int).Value = If(Q4_3_selval <> "", Q4_3_selval, Convert.DBNull) '
                .Parameters.Add("Q4_4", SqlDbType.Int).Value = If(Q4_4_selval <> "", Q4_4_selval, Convert.DBNull) '

                '貳、其它意見
                .Parameters.Add("SUGGESTION", SqlDbType.NVarChar).Value = If(txt_Suggestion.Text <> "", txt_Suggestion.Text, Convert.DBNull)  'SUGGESTIO 
                '抽訪人員姓名 簽章
                .Parameters.Add("SIGNER", SqlDbType.VarChar).Value = If(signer.Text <> "", signer.Text, Convert.DBNull) 'SIGNERFROM
                '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                .Parameters.Add("DASOURCE", SqlDbType.Int).Value = 2 'SIGNERFROM
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                .Parameters.Add("OCID", SqlDbType.Int).Value = CInt(OCIDValue1.Value)
                .Parameters.Add("SOCID", SqlDbType.Int).Value = rSOCID
                '.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                .ExecuteNonQuery()
            End With
        End If

        Dim s_FUNID As String = TIMS.Get_MRqID(Me)
        Common.MessageBox(Me, "儲存成功！")
        TIMS.Utl_Redirect(Me, objconn, "CP_07_001.aspx?ID=" & s_FUNID)
    End Sub

    '檢查
    Private Function chkInputAnswer(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True
        errMsg = ""
        Dim RtnMsg As String = ""

        'Select Case dl_InquireType.SelectedValue
        '    Case "1", "2"
        '    Case Else
        '        errMsg &= "請選擇「問 卷 調 查 類 型」\n"
        'End Select

        'If dl_InquireType.SelectedValue = "2" Then
        '    If Trim(signer.Text) = "" Then
        '        errMsg &= "請填寫「抽訪人員姓名 簽章欄位」\n"
        '    Else
        '        RtnMsg = TIMS.CheckInputRtn(signer.Text)
        '        If RtnMsg <> "" Then
        '            errMsg &= "「抽訪人員姓名 簽章」 欄位輸入字串中含有不合法字元" & RtnMsg
        '        End If
        '    End If
        'End If

        Dim Q1_1_selval As String = TIMS.GetListValue(Q1_1)
        Dim Q1_2_selval As String = TIMS.GetListValue(Q1_2)
        Dim Q1_3_selval As String = TIMS.GetListValue(Q1_3)
        Dim Q2_1_selval As String = TIMS.GetListValue(Q2_1)
        Dim Q2_2_selval As String = TIMS.GetListValue(Q2_2)
        Dim Q2_3_selval As String = TIMS.GetListValue(Q2_3)
        Dim Q2_4_selval As String = TIMS.GetListValue(Q2_4)
        Dim Q3_1_selval As String = TIMS.GetListValue(Q3_1)
        Dim Q3_2_selval As String = TIMS.GetListValue(Q3_2)
        Dim Q3_3_selval As String = TIMS.GetListValue(Q3_3)
        Dim Q4_1_selval As String = TIMS.GetListValue(Q4_1)
        Dim Q4_2_selval As String = TIMS.GetListValue(Q4_2)
        Dim Q4_3_selval As String = TIMS.GetListValue(Q4_3)
        Dim Q4_4_selval As String = TIMS.GetListValue(Q4_4)

        txFillDate.Text = TIMS.Cdate3(txFillDate.Text)
        If txFillDate.Text = "" Then
            errMsg &= "填表日期不得為空！" & vbCrLf
        ElseIf txFillDate.Text <> "" Then
            If TIMS.Cdate3(txFillDate.Text) > TIMS.Cdate3(Date.Now) Then errMsg &= "填表日期不得大於今日！" & vbCrLf
        End If

        If Q1_1_selval = "" Then errMsg &= "請填寫 [一、課程安排] - 1.您對整體課程安排是否滿意？" & vbCrLf
        If Q1_2_selval = "" Then errMsg &= "請填寫 [一、課程安排] - 2.您對課程安排內容銜接情形是否滿意？" & vbCrLf
        If Q1_3_selval = "" Then errMsg &= "請填寫 [一、課程安排] - 3.您對課程內容時數安排是否滿意？" & vbCrLf
        If Q2_1_selval = "" Then errMsg &= "請填寫 [二、師資、助教及教學] - 1.您對老師的教學方式是否滿意？" & vbCrLf
        If Q2_2_selval = "" Then errMsg &= "請填寫 [二、師資、助教及教學] - 2.您對老師的教學態度是否滿意？" & vbCrLf
        If Q2_3_selval = "" Then errMsg &= "請填寫 [二、師資、助教及教學] - 3.您對老師的專業知識是否滿意？" & vbCrLf
        'If Q2_4_selval = "" Then errMsg &= "請填寫 [二、師資、助教及教學] - 4.您對助教的協助教學是否滿意? (無助教不須勾選)" & vbCrLf
        If Q3_1_selval = "" Then errMsg &= "請填寫 [三、設備和教材] - 1.您對上課期間，教材設備的充分利用情形是否滿意？" & vbCrLf
        If Q3_2_selval = "" Then errMsg &= "請填寫 [三、設備和教材] - 2.您對教材的適用度是否滿意？" & vbCrLf
        If Q3_3_selval = "" Then errMsg &= "請填寫 [三、設備和教材] - 3.您對教材內容的難易程度是否滿意？" & vbCrLf
        If Q4_1_selval = "" Then errMsg &= "請填寫 [四、行政措施] - 1.您對導師關心學員學習狀況及解決問題的能力是否滿意？" & vbCrLf
        If Q4_2_selval = "" Then errMsg &= "請填寫 [四、行政措施] - 2.您對求助及申訴管道是否滿意？" & vbCrLf
        If Q4_3_selval = "" Then errMsg &= "請填寫 [四、行政措施] - 3.您對學習環境是否滿意？" & vbCrLf
        If Q4_4_selval = "" Then errMsg &= "請填寫 [四、行政措施] - 4.您對上課地點場地及週遭環境清潔是否滿意？" & vbCrLf

        If txt_Suggestion.Text <> "" Then
            RtnMsg = TIMS.CheckInputRtn(txt_Suggestion.Text)
            If RtnMsg <> "" Then
                errMsg &= "貳、其它意見 欄位輸入字串中含有不合法字元" & RtnMsg
            End If
        End If

        If errMsg <> "" Then rst = False
        Return rst
    End Function

    Protected Sub bt_back_Click(sender As Object, e As EventArgs) Handles bt_back.Click
        Dim s_FUNID As String = TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, "CP_07_001.aspx?ID=" & s_FUNID)
    End Sub

#Region "NO USE"
    'Public Function ConvertNum(ByVal Num As String) As String    '
    '    Dim NumChar As String = ""
    '    Select Case Num
    '        Case 1
    '            NumChar = "一"
    '        Case 2
    '            NumChar = "二"
    '        Case 3
    '            NumChar = "三"
    '        Case 4
    '            NumChar = "四"
    '        Case 5
    '            NumChar = "五"
    '        Case 6
    '            NumChar = "六"
    '        Case 0
    '            NumChar = "日"
    '    End Select
    '    Return NumChar
    'End Function

    'Private Function ReplaceStr(ByVal InputStr)
    '    InputStr = Replace(InputStr, "'", "''")
    '    Return InputStr
    'End Function



    'Private Sub PrePage()
    '    Dim parmstr As String = ""

    '    Session(SearchStr) = Me.ViewState("SearchStr")
    '    Session("_" & SearchStr) = Me.ViewState("_SearchStr")
    '    'If Request("DOCID") <> "" Then
    '    '    TIMS.Utl_Redirect1(Me, "CP_07_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID"))
    '    'Else
    '    '    TIMS.Utl_Redirect1(Me, "CP_07_001.aspx?ID=" & Request("ID"))
    '    'End If


    '    parmstr = "'CP_07_001.aspx"
    '    If ocid <> "" Then
    '        parmstr += "?ocid=" & ocid
    '    End If
    '    If socid <> "" Then
    '        parmstr += "&socid=" & socid
    '    End If
    '    If rid <> "" Then
    '        parmstr += "&rid=" & rid
    '    End If
    '    parmstr += "';"

    '    If Me.ViewState("parmstr") <> "" Then
    '        Common.RespWrite(Me, "<script>location.href=" & Me.ViewState("parmstr") & "</script>")
    '    Else
    '        Common.RespWrite(Me, "<script>location.href=" & parmstr & "</script>")
    '    End If
    'End Sub

    'Private Sub bt_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_back.Click
    '    PrePage()
    'End Sub

    'Private Sub load_SOCID()

    '    Dim dt As New DataTable
    '    Dim sql As String = ""
    '    Dim conn As New SqlConnection
    '    conn = DbAccess.GetConnection()
    '    conn.Open()
    '    Try
    '        sql = "SELECT StudentID, case "
    '        sql += "when len(a.StudentID)=12 then b.Name+'('+RIGHT(a.StudentID,3)+')' "
    '        sql += "else b.Name+'('+RIGHT(a.StudentID,2)+')' "
    '        sql += "end as Name, a.SOCID "
    '        sql += "FROM (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "') a "
    '        sql += "JOIN (SELECT * FROM Stud_StudentInfo) b ON a.SID=b.SID"
    '        dt = DbAccess.GetDataTable(sql)
    '        dt.DefaultView.Sort = "StudentID"
    '        With ddl_SOCID
    '            .DataSource = dt
    '            .DataTextField = "Name"
    '            .DataValueField = "SOCID"
    '            .DataBind()
    '        End With
    '        Common.SetListItem(ddl_SOCID, socid)
    '    Catch ex As Exception
    '        Dim strScript As String
    '        strScript = "<script language=""javascript"">" + vbCrLf
    '        strScript += "alert('發生錯誤!! " & ex.Message.ToString() & "');" + vbCrLf
    '        strScript += "</script>"
    '        Page.RegisterStartupScript("", strScript)
    '    Finally
    '        conn.Close()
    '        dt.Dispose()
    '    End Try

    'End Sub

    'Private Sub ddl_SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    socid = ddl_SOCID.SelectedValue
    '    load_Data(ddl_SOCID.SelectedValue)
    'End Sub
#End Region

    '問卷調查類型(1:系統登打 2:電話訪查)
    'Private Sub dl_InquireType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dl_InquireType.SelectedIndexChanged
    '    chkInquireType()
    'End Sub

    'Private Sub chkInquireType()
    '    ' 抽訪人員姓名 簽章
    '    tr_1.Visible = False
    '    If dl_InquireType.SelectedValue = "2" Then
    '        tr_1.Visible = True
    '    End If
    'End Sub

    '    Private Sub LinkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
    '        If txtFillDate.Value <> "" Then
    '            lb_FillDate.Text = Convert.ToString(CDate(txtFillDate.Value).Year()) & "年" & Convert.ToString(CDate(txtFillDate.Value).Month()) & "月" & Convert.ToString(CDate(txtFillDate.Value).Day()) & "日"
    '            WeekDay.Text = TIMS.ConvertNum(Convert.ToString(CDate(txtFillDate.Value).DayOfWeek()))
    '        Else
    '            lb_FillDate.Text = ""
    '            WeekDay.Text = ""
    '            Common.MessageBox(Me, "未輸入日期。")
    '            Exit Sub
    '        End If

    '    End Sub
End Class
