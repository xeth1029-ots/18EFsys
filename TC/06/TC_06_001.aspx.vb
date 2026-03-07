Partial Class TC_06_001
    Inherits AuthBasePage

    Dim vMsg As String = ""
    Dim dtOrgBlack As DataTable  '取出系統現有黑名單
    Const cst_isBlackMsg As String = TIMS.cst_gBlackMsg1 '檢測黑名單機構
    Dim ChgItemName As String() '將變更項目名稱定義到陣列之中
    'Dim BlnTest1 As Boolean = False '正式環境為false

    '線上送件時間
    Dim fg_SHOW_ONLINESENDDATE As Boolean = False
    Const cst_DG_PRL_COL_線上送件時間 As Integer = 9
    Const cst_DG_PRL_COL_備註 As Integer = 10
    Const cst_ss_search1 As String = "_search1"
    'DataGrid : PlanReviseList ：PageControler1
    Const cst_dgAct_View1 As String = "View1" '檢視 btnView1
    Const cst_dgAct_Updat1 As String = "Updat1" '修改 btnUpdat1
    Const cst_dgAct_Edit1 As String = "Edit1" '審核 btnEdit
    Const cst_dgAct_Delete1 As String = "Delete1" '刪除 btnDelete
    Const cst_dgAct_PartReduc1 As String = "PartReduc1" '還原 btnPartReduc

    'UPDATE TABLE : Plan_Revise
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = PlanReviseList

        Call Utl_EveryCreate1()

        If Not Page.IsPostBack Then
            Call SCreate2()
        End If

        'But_Search.Enabled = False
        'If au.blnCanSech Then But_Search.Enabled = True
    End Sub

    Sub Utl_EveryCreate1()
        'BlnTest1 = TIMS.sUtl_ChkTest()
        'https://jira.turbotech.com.tw/browse/TIMSC-148
        trReviseStatus2854.Visible = (TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) > -1)
        'If TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) > -1 Then trReviseStatus28.Visible = True
        '線上送件時間(產投、充飛)
        fg_SHOW_ONLINESENDDATE = (TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) > -1)
        Dim flag_can_ExpType1 As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID = 0, True, False)
        trRBListExpType.Visible = flag_can_ExpType1 'False
        BtnExp1.Visible = flag_can_ExpType1 'False

        '取出系統現有黑名單
        dtOrgBlack = TIMS.Get_OrgBlackList(Me, objconn)

        '將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改
        '產學訓套用的顯示字串  / '非產學訓套用的顯示字串
        ChgItemName = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.TPlanID28ChgItemName, TIMS.TPlanIDChgItemName)

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then LabTMID.Text = "訓練業別"

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button1.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
    End Sub

    Sub SCreate2()
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        'dg@PlanReviseList
        PlanReviseList.Columns(cst_DG_PRL_COL_線上送件時間).Visible = fg_SHOW_ONLINESENDDATE ' False
        PlanReviseList.Columns(cst_DG_PRL_COL_備註).Visible = fg_SHOW_ONLINESENDDATE ' False

        If Not Session(cst_ss_search1) Is Nothing Then
            Dim sSearch11 As String = Convert.ToString(Session(cst_ss_search1))
            Dim MyValue As String = TIMS.GetMyValue(sSearch11, "prg")
            If MyValue = "TC_06_001" Then
                center.Text = TIMS.GetMyValue(sSearch11, "center")
                RIDValue.Value = TIMS.GetMyValue(sSearch11, "RIDValue")
                TB_career_id.Text = TIMS.GetMyValue(sSearch11, "TB_career_id")
                trainValue.Value = TIMS.GetMyValue(sSearch11, "trainValue")
                jobValue.Value = TIMS.GetMyValue(sSearch11, "jobValue")
                txtCJOB_NAME.Text = TIMS.GetMyValue(sSearch11, "txtCJOB_NAME")
                cjobValue.Value = TIMS.GetMyValue(sSearch11, "cjobValue")
                ClassCName.Text = TIMS.GetMyValue(sSearch11, "ClassCName")
                CyclType.Text = TIMS.GetMyValue(sSearch11, "CyclType")
                ApplySDate.Text = TIMS.GetMyValue(sSearch11, "ApplySDate")
                ApplyEDate.Text = TIMS.GetMyValue(sSearch11, "ApplyEDate")
                MyValue = TIMS.GetMyValue(sSearch11, "rblReviseStatus")
                If MyValue <> "" Then Common.SetListItem(rblReviseStatus, MyValue)
                'Me.ViewState("PageIndex") = TIMS.GetMyValue(sSearch11, "PageIndex")
                'If IsNumeric(ViewState("PageIndex")) Then PageControler1.PageIndex = ViewState("PageIndex")
                MyValue = TIMS.GetMyValue(sSearch11, "KeepSearch")
                If MyValue = "1" Then Call Search1()
            End If
            Session(cst_ss_search1) = Nothing
        End If
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        ClassCName.Text = TIMS.ClearSQM(ClassCName.Text)
        ApplySDate.Text = TIMS.ClearSQM(ApplySDate.Text)
        ApplyEDate.Text = TIMS.ClearSQM(ApplyEDate.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(ApplySDate.Text, ApplyEDate.Text) Then
            Dim T_DATE1 As String = ApplySDate.Text
            ApplySDate.Text = ApplyEDate.Text
            ApplyEDate.Text = T_DATE1
        End If

        ViewState("ClassCName") = ""
        If Trim(ClassCName.Text) <> "" Then
            'Me.ViewState("ClassCName") = TIMS.ChangeSQM(UCase(ClassCName.Text))
            ViewState("ClassCName") = TIMS.ChangeSQM(ClassCName.Text)
            'Me.ViewState("Request_sig") = Request("sig")
            If TIMS.CheckInput(ViewState("ClassCName")) Then Errmsg += "班級名稱 輸入格式異常，請重新輸入" & vbCrLf
        End If
        If ApplySDate.Text <> "" Then
            If Not TIMS.IsDate1(ApplySDate.Text) Then Errmsg += "申請日期 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then ApplySDate.Text = CDate(ApplySDate.Text).ToString("yyyy/MM/dd")
        Else
            'Errmsg += "申請日期 起始日期 為必填" & vbCrLf
        End If
        If ApplyEDate.Text <> "" Then
            If Not TIMS.IsDate1(ApplyEDate.Text) Then Errmsg += "申請日期 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then ApplyEDate.Text = CDate(ApplyEDate.Text).ToString("yyyy/MM/dd")
        Else
            'Errmsg += "結訓日期 迄止日期 為必填" & vbCrLf
        End If
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    ''' <summary> 查詢SQL 語法 </summary>
    ''' <returns></returns>
    Function Search_SQL() As String
        'Me.ViewState("Search") = KeepSearch() '會產 vsSearch1 'Dim objtable As DataTable
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sql As String = ""
        sql &= " SELECT a.PLANID ,a.COMIDNO ,a.SEQNO" & vbCrLf
        '申請變更日期
        sql &= " ,format(a.CDATE ,'yyyy/MM/dd') CDATE" & vbCrLf
        sql &= " ,a.SUBSEQNO,a.ALTDATAID" & vbCrLf
        '審核狀態
        sql &= " ,a.REVISECONT,a.REVISESTATUS" & vbCrLf
        '審核時間
        sql &= " ,a.REVISEDATE" & vbCrLf
        sql &= " ,ip.PLANNAME,ip.DISTNAME,ip.DISTID" & vbCrLf
        sql &= " ,pp.PLANYEAR,pp.RID,oo.ORGNAME,pp.CLASSNAME,pp.CYCLTYPE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) CLASSNAME2" & vbCrLf
        sql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(pp.FDDATE,'yyyy/MM/dd') FDDATE" & vbCrLf
        sql &= " ,format(a.ONLINESENDDATE,'yyyy/MM/dd tt hh:mm:ss') ONLINESENDDATE" & vbCrLf
        '申請變更函送日期
        sql &= " ,format(a.SENDDATE4,'yyyy/MM/dd') SENDDATE4" & vbCrLf
        sql &= " ,a.STATUS4,a.ISPASS4,a.OVERWEEK4,a.NOINC4,a.NODEDUC4" & vbCrLf
        sql &= " ,op.ADDRESS ,op.CONTACTNAME,op.PHONE" & vbCrLf
        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
        sql &= " ,a.PARTREDUC,a.REDUCACCT,a.REDUCDATE" & vbCrLf
        sql &= " ,'N' isBlack" & vbCrLf '檢測黑名單機構
        sql &= " ,dbo.FN_GET_ALTDATAID_N28(a.ALTDATAID) ALTDATAID_N" & vbCrLf
        sql &= " ,case a.STATUS4 when 1 then '依規定辦理' when 2 then '逾期(扣分)' when 3 then '逾期(不扣分)' end STATUS4_N" & vbCrLf
        sql &= " ,case a.ISPASS4 when 'R' then '駁回' when 'Y' then '符合規定' when 'N' then '不符合規定' end ISPASS4_N" & vbCrLf
        sql &= " ,case a.OVERWEEK4 when 1 then '1週以內' when 2 then '1週以上' when 3 then '停辦' when 9 then '無逾期' end OVERWEEK4_N" & vbCrLf
        sql &= " ,case a.NOINC4 when 'Y' then '是' when 'N' then '否' end NOINC4_N" & vbCrLf
        sql &= " ,case a.NODEDUC4 when 'Y' then '是' when 'N' then '否' end NODEDUC4_N" & vbCrLf
        'sql &= " ,case WHEN a.REVISESTATUS IS NOT NULL THEN '已審核' ELSE '未審核' end REVISESTATUS_N" & vbCrLf
        sql &= " ,case a.REVISESTATUS WHEN 'Y' THEN '審核通過' WHEN 'N' THEN '審核不通過' ELSE '審核中' end REVISESTATUS_N" & vbCrLf
        sql &= " ,r.ORGPLANNAME" & vbCrLf
        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)
        sql &= " ,pp.APPSTAGE" & vbCrLf
        sql &= " ,CASE pp.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' END APPSTAGE_N" & vbCrLf
        sql &= " ,pp.PSNO28" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,CASE WHEN cc.OCID IS NOT NULL THEN convert(VARCHAR,cc.OCID) ELSE pp.PSNO28 END OCPSNO28" & vbCrLf

        sql &= " FROM PLAN_REVISE a" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.PLANID=a.PLANID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.COMIDNO=a.COMIDNO" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID AND pp.COMIDNO=a.COMIDNO AND pp.SEQNO=a.SEQNO" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME r on r.RID=pp.RID" & vbCrLf
        sql &= " JOIN ORG_ORGPLANINFO op on op.RSID=r.RSID" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO cc on cc.PLANID=pp.PLANID and cc.COMIDNO=pp.COMIDNO and cc.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " WHERE ip.TPLANID = '" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " AND ip.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf
        'sql &= " WHERE 1=1 AND a.STATUS4 is not null and a.MODIFYDATE>=GETDATE()-10" & vbCrLf
        '沒有值查無資料
        sql &= If(RIDValue.Value = "", " AND 1<>1", String.Concat(" AND pp.RID LIKE '", RIDValue.Value, "%'")) & vbCrLf

        If Not (sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1) Then
            sql &= " AND ip.PLANID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND ip.DISTID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        If sm.UserInfo.LID = 2 Then sql &= " AND 1<>1" & vbCrLf

        'https://jira.turbotech.com.tw/browse/TIMSC-148
        'trReviseStatus28 'X:未審核 O:已審核 Y:通過 N:不通過
        Dim v_rblReviseStatus As String = TIMS.GetListValue(rblReviseStatus)
        If trReviseStatus2854.Visible Then
            '限定 產業人才投資方案
            Select Case v_rblReviseStatus'rblReviseStatus.SelectedValue
                Case "O"
                    sql &= " AND a.ReviseStatus IS NOT NULL" & vbCrLf '正式
                Case "Y"
                    sql &= " AND a.ReviseStatus='Y'" & vbCrLf '正式
                Case "N"
                    sql &= " AND a.ReviseStatus='N'" & vbCrLf '正式
                Case Else
                    sql &= " AND a.ReviseStatus IS NULL" & vbCrLf '正式'未審核
            End Select
        Else
            '非-產業人才投資方案
            sql &= " AND a.ReviseStatus IS NULL" & vbCrLf '正式
            'If Not BlnTest1 Then '正式環境為false
            '     sql &=  " AND b.ReviseStatus IS NULL" & vbCrLf '正式
            'End If
        End If

        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Me.LabTMID.Text = "訓練業別"
            If jobValue.Value <> "" Then
                sql &= " AND ( pp.TMID = " & jobValue.Value & " OR pp.TMID IN (" & vbCrLf
                sql &= "   SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '職類別
                sql &= "   SELECT TMID FROM Key_TrainType WHERE parent IN (" & vbCrLf '業別
                sql &= "   SELECT TMID FROM Key_TrainType WHERE busid = 'G')" & vbCrLf '產業人才投資方案類
                sql &= " AND TMID=" & jobValue.Value & " )))" & vbCrLf
            End If
        Else
            If trainValue.Value <> "" Then sql &= " AND pp.TMID=" & trainValue.Value & vbCrLf
        End If
        txtCJOB_NAME.Text = TIMS.ClearSQM(txtCJOB_NAME.Text)
        If txtCJOB_NAME.Text <> "" Then sql &= " AND pp.CJOB_UNKEY=" & cjobValue.Value & "" & vbCrLf  '通俗職類
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then sql &= " AND pp.CyclType='" & CyclType.Text & "'" & vbCrLf

        ViewState("ClassCName") = TIMS.ClearSQM(ViewState("ClassCName"))
        ApplySDate.Text = TIMS.ClearSQM(ApplySDate.Text)
        ApplyEDate.Text = TIMS.ClearSQM(ApplyEDate.Text)
        If ViewState("ClassCName") <> "" Then sql &= " AND pp.ClassName LIKE '%" & ViewState("ClassCName") & "%'" & vbCrLf 'fix ORA-01722: invalid number
        If ApplySDate.Text <> "" Then sql &= " AND a.CDate>=" & TIMS.To_date(ApplySDate.Text) & vbCrLf
        If ApplyEDate.Text <> "" Then sql &= " AND a.CDate<=" & TIMS.To_date(ApplyEDate.Text) & vbCrLf
        ' sql &=  " AND b.ReviseStatus IS NULL "
        Return sql
    End Function

    '查詢
    Sub Search1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Me.ViewState("Search") = KeepSearch() '會產 vsSearch1 'Dim objtable As DataTable
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        '查詢SQL 語法 
        Dim sql As String = Search_SQL()

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            '檢測黑名單機構
            For Each odr As DataRow In dt.Rows
                If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS<>'38'").Length > 0 Then
                    odr("isBlack") = "Y"
                Else
                    If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS='38' AND DistID='" & odr("DistID") & "'").Length > 0 Then odr("isBlack") = "Y"
                End If
            Next
            'dt.AcceptChanges()
        End If

        DataGridTable.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.Sort = "STDate,ClassName"
        PageControler1.ControlerLoad()
    End Sub

    '查詢 '搜尋時寫入
    Private Sub But_Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Search.Click
        Call Search1()
    End Sub

    Function Get_PTDRID(ByVal tmpID As Integer, ByVal tmpNO As String, ByVal tmpSNO As Integer, ByVal tmpDate As String, ByVal tmpSubSNO As Integer) As Integer
        Dim iRst As Integer = 0
        TIMS.OpenDbConn(objconn)

        Dim sql As String = ""
        sql &= " SELECT PTDRID" & vbCrLf
        sql &= " FROM PLAN_TRAINDESC_REVISE" & vbCrLf
        sql &= " WHERE PlanID=@planid AND ComIDNO=@comidno AND SeqNO=@seqno" & vbCrLf
        sql &= " AND CDate=@cdate AND SubSeqNO=@subseqno" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("planid", SqlDbType.Int).Value = tmpID
            .Parameters.Add("comidno", SqlDbType.VarChar).Value = tmpNO
            .Parameters.Add("seqno", SqlDbType.Int).Value = tmpSNO
            .Parameters.Add("cdate", SqlDbType.DateTime).Value = Convert.ToDateTime(tmpDate)
            .Parameters.Add("subseqno", SqlDbType.Int).Value = tmpSubSNO
            iRst = .ExecuteScalar()
        End With

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
        Return iRst
    End Function

    ''' <summary> 已經有新申請變更資料，因順序問題不可執行-還原 </summary>
    ''' <param name="vsPlanID"></param>
    ''' <param name="vsComIDNO"></param>
    ''' <param name="vsSeqNO"></param>
    ''' <param name="vsCDate"></param>
    ''' <param name="vsSubSeqNO"></param>
    ''' <param name="vsAltDataID"></param>
    ''' <returns></returns>
    Function Check_PartReduc_PlanRevise(ByVal vsPlanID As String, ByVal vsComIDNO As String, ByVal vsSeqNO As String, ByVal vsCDate As String, ByVal vsSubSeqNO As String, ByVal vsAltDataID As String) As String
        Dim msg1 As String = ""
        Const cst_msg1 As String = "已經有新申請變更資料，因順序問題不可執行還原!"

        'parms.Add("CDate", TIMS.cdate3(vsCDate)) 'parms.Add("SubSeqNO", vsSubSeqNO)'parms.Clear()
        'u_Sql &= " And CDate = CONVERT(DATETIME, @CDate, 111) And SubSeqNO = @SubSeqNO" & vbCrLf
        Dim parms As New Hashtable From {{"PlanID", vsPlanID}, {"ComIDNO", vsComIDNO}, {"SeqNO", vsSeqNO}, {"AltDataID", vsAltDataID}}
        Dim sql As String = ""
        sql &= " SELECT MAX(CDate) CDate FROM PLAN_REVISE" & vbCrLf
        sql &= " WHERE PlanID=@PlanID And ComIDNO=@ComIDNO And SeqNO=@SeqNO" & vbCrLf
        sql &= " AND AltDataID=@AltDataID" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt1.Rows.Count = 0 Then Return msg1
        Dim dr1 As DataRow = dt1.Rows(0) '應該是同一筆
        If TIMS.Cdate3(dr1("CDate")) <> TIMS.Cdate3(vsCDate) Then msg1 = cst_msg1

        'Dim parms As New Hashtable 'parms.Add("SubSeqNO", vsSubSeqNO)
        Dim parms2 As New Hashtable From {{"PlanID", vsPlanID}, {"ComIDNO", vsComIDNO}, {"SeqNO", vsSeqNO}, {"CDate", TIMS.Cdate3(vsCDate)}, {"AltDataID", vsAltDataID}}
        Dim sql2 As String = ""
        sql2 &= " Select MAX(SubSeqNO) SubSeqNO FROM PLAN_REVISE" & vbCrLf
        sql2 &= " WHERE PlanID=@PlanID And ComIDNO=@ComIDNO And SeqNO=@SeqNO" & vbCrLf
        sql2 &= " And CDate=CONVERT(DATETIME, @CDate, 111) And AltDataID=@AltDataID" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, parms2)
        If dt2.Rows.Count = 0 Then Return msg1
        Dim dr2 As DataRow = dt2.Rows(0) '應該是同一筆
        If Convert.ToString(dr2("SubSeqNO")) <> vsSubSeqNO Then msg1 = cst_msg1

        Return msg1
    End Function

    ''' <summary> 儲存-執行還原 </summary>
    ''' <param name="vsPlanID"></param>
    ''' <param name="vsComIDNO"></param>
    ''' <param name="vsSeqNO"></param>
    ''' <param name="vsCDate"></param>
    ''' <param name="vsSubSeqNO"></param>
    ''' <param name="vsAltDataID"></param>
    Sub UPDATE_PartReduc_PlanRevise(ByVal vsPlanID As String, ByVal vsComIDNO As String, ByVal vsSeqNO As String, ByVal vsCDate As String, ByVal vsSubSeqNO As String, ByVal vsAltDataID As String)
        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
        Dim u_Sql As String = ""
        u_Sql &= " UPDATE PLAN_REVISE"
        u_Sql &= " SET REDUCACCT=@REDUCACCT ,REDUCDATE=GETDATE()" & vbCrLf
        u_Sql &= " ,PARTREDUC=@PARTREDUC" & vbCrLf '新增「還原」按鈕
        u_Sql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        u_Sql &= " ,ONLINESENDSTATUS=NULL ,ONLINESENDACCT=NULL ,ONLINESENDDATE=NULL" & vbCrLf
        u_Sql &= " WHERE PlanID=@PlanID And ComIDNO=@ComIDNO And SeqNO=@SeqNO" & vbCrLf
        u_Sql &= " And CDate=CONVERT(DATETIME, @CDate, 111)" & vbCrLf
        u_Sql &= " And SubSeqNO=@SubSeqNO And AltDataID=@AltDataID" & vbCrLf

        'u_Parms.Clear()
        Dim u_Parms As New Hashtable
        u_Parms.Add("REDUCACCT", sm.UserInfo.UserID) '新增「還原」按鈕
        u_Parms.Add("PARTREDUC", TIMS.cst_YES) '新增「還原」按鈕
        u_Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_Parms.Add("PlanID", vsPlanID)
        u_Parms.Add("ComIDNO", vsComIDNO)
        u_Parms.Add("SeqNO", vsSeqNO)
        u_Parms.Add("CDate", TIMS.Cdate3(vsCDate))
        u_Parms.Add("SubSeqNO", vsSubSeqNO)
        u_Parms.Add("AltDataID", vsAltDataID)
        DbAccess.ExecuteNonQuery(u_Sql, objconn, u_Parms)  'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
    End Sub

    Private Sub PlanReviseList_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles PlanReviseList.ItemCommand
        '保留查詢值
        Call KeepSession()
        If e.CommandArgument = "" Then Exit Sub

        Dim vsPlanID As String = TIMS.GetMyValue(e.CommandArgument, "PlanID")
        Dim vsComIDNO As String = TIMS.GetMyValue(e.CommandArgument, "ComIDNO")
        Dim vsSeqNO As String = TIMS.GetMyValue(e.CommandArgument, "SeqNO")
        Dim vsCDate As String = TIMS.GetMyValue(e.CommandArgument, "CDate")
        Dim vsSubSeqNO As String = TIMS.GetMyValue(e.CommandArgument, "SubSeqNO")
        Dim vsAltDataID As String = TIMS.GetMyValue(e.CommandArgument, "AltDataID")
        If vsPlanID = "" Then Exit Sub
        If vsComIDNO = "" Then Exit Sub
        If vsSeqNO = "" Then Exit Sub

        Select Case e.CommandName
            Case cst_dgAct_View1, cst_dgAct_Updat1, cst_dgAct_Edit1
                '檢視/修改/審核
                'Dim RqID As String = TIMS.Get_MRqID(Me)
                Dim url1 As String = ""
                url1 &= "TC_06_001_chk.aspx?ID=" & TIMS.Get_MRqID(Me) 'Request("ID")
                url1 &= "&act=" & e.CommandName 'cst_dgAct_View1/cst_dgAct_Updat1/cst_dgAct_Edit1
                url1 &= "&PlanID=" & vsPlanID
                url1 &= "&cid=" & vsComIDNO
                url1 &= "&no=" & vsSeqNO
                url1 &= "&CDate=" & vsCDate
                url1 &= "&subno=" & vsSubSeqNO
                url1 &= "&AltDataID=" & vsAltDataID
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case cst_dgAct_PartReduc1 '還原 btnPartReduc
                'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
                Dim s_NGMSG1 As String = Check_PartReduc_PlanRevise(vsPlanID, vsComIDNO, vsSeqNO, vsCDate, vsSubSeqNO, vsAltDataID)
                If s_NGMSG1 <> "" Then
                    Common.MessageBox(Me, s_NGMSG1)
                    Return ' Exit Sub
                End If
                'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
                Call UPDATE_PartReduc_PlanRevise(vsPlanID, vsComIDNO, vsSeqNO, vsCDate, vsSubSeqNO, vsAltDataID)

                Call Search1()
                Common.MessageBox(Me, "還原成功")
                Exit Sub

            Case cst_dgAct_Delete1 '刪除
                Try
                    Dim PTDRID As Integer = Get_PTDRID(vsPlanID, vsComIDNO, vsSeqNO, vsCDate, vsSubSeqNO)
                    '刪除 DELETE PLAN_TRAINDESC_REVISEITEM. DELETE PLAN_TRAINDESC_REVISE.
                    If PTDRID <> 0 Then Call TIMS.DEL_PLAN_TRAINDESC_REVISEITEM(sm, PTDRID, objconn)
                    '刪除 DELETE PLAN_REVISE
                    Call TIMS.DELETE_PLANREVISE(vsPlanID, vsComIDNO, vsSeqNO, vsCDate, vsSubSeqNO, vsAltDataID, objconn)
                Catch ex As Exception
                    Dim exMessage As String = ex.Message
                    vMsg = vbCrLf
                    vMsg &= "PlanReviseList_ItemCommand" & vbCrLf
                    vMsg &= "ex.ToString:" & ex.ToString & vbCrLf
                    vMsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    'vMsg = Replace(vMsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(vMsg)

                    Common.MessageBox(Me, $"刪除失敗!!,{exMessage}")
                    Return
                    'Throw ex
                End Try

                Call Search1()
                Common.MessageBox(Me, "刪除成功")
                Exit Sub

        End Select
    End Sub

    Private Sub PlanReviseList_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles PlanReviseList.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnView1 As Button = e.Item.FindControl("btnView1") '檢視 btnView1
                Dim btnUpdat1 As Button = e.Item.FindControl("btnUpdat1") '修改 btnUpdat1
                Dim btnEdit As Button = e.Item.FindControl("btnEdit") '審核 btnEdit
                Dim btnDelete As Button = e.Item.FindControl("btnDelete") '刪除 btnDelete
                Dim btnPartReduc As Button = e.Item.FindControl("btnPartReduc") '還原 btnPartReduc

                Dim labApplyDate As Label = e.Item.FindControl("labApplyDate")
                Dim labMemoD1 As Label = e.Item.FindControl("labMemoD1")
                Dim labChgTypeN1 As Label = e.Item.FindControl("labChgTypeN1")
                btnView1.CommandName = cst_dgAct_View1
                btnUpdat1.Visible = False '修改不顯示
                btnUpdat1.CommandName = cst_dgAct_Updat1 '修改不顯示
                btnEdit.CommandName = cst_dgAct_Edit1
                btnDelete.CommandName = cst_dgAct_Delete1
                btnPartReduc.CommandName = cst_dgAct_PartReduc1

                btnDelete.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                TIMS.Tooltip(btnDelete, "刪除錯誤申請的變更")
                'PARTREDUC
                btnPartReduc.Attributes("onclick") = "return confirm('您確定要還原這一筆資料?');"
                TIMS.Tooltip(btnPartReduc, "還原申請變更")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "ComIDNO", Convert.ToString(drv("ComIDNO")))
                TIMS.SetMyValue(sCmdArg, "SeqNO", Convert.ToString(drv("SeqNO")))
                TIMS.SetMyValue(sCmdArg, "CDate", Common.FormatDate(drv("CDate")))
                TIMS.SetMyValue(sCmdArg, "SubSeqNO", Convert.ToString(drv("SubSeqNO")))
                TIMS.SetMyValue(sCmdArg, "AltDataID", Convert.ToString(drv("AltDataID")))
                TIMS.SetMyValue(sCmdArg, "STDate", Convert.ToString(drv("STDate")))
                btnView1.CommandArgument = sCmdArg
                btnUpdat1.CommandArgument = sCmdArg
                btnEdit.CommandArgument = sCmdArg
                btnDelete.CommandArgument = sCmdArg
                btnPartReduc.CommandArgument = sCmdArg

                btnView1.Visible = True
                btnEdit.Visible = False
                btnDelete.Visible = False '其他計畫暫不啟用刪除鈕
                btnPartReduc.Visible = False '其他計畫暫不啟用還原鈕
                labMemoD1.Text = ""
                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產投專用
                    Select Case Convert.ToString(drv("AltDataID"))
                        Case "1" '訓練期間,顯示該開訓日前三天與後三天的開訓班數
                            'https://jira.turbotech.com.tw/browse/TIMSC-148
                            labMemoD1.Text = TIMS.Get_MemoD1(sCmdArg, objconn)
                    End Select
                End If

                If $"{drv("ReviseStatus")}" = "" Then '尚未審核
                    btnView1.Visible = False '停用檢視
                    btnEdit.Visible = True '未審可審資料
                    Dim flag_can_use_delete_btn As Boolean = False
                    Dim flag_can_use_PARTREDUC_btn As Boolean = False '「還原」按鈕
                    '產投
                    If TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        ' OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
                        '產業人才投資方案 ' 審核狀態為「未審核」' 查詢頁面新增「還原」按鈕，訓練單位向分署提出修改需求時，分署按「還原」後，單位端可進行修改。
                        '「還原」並非退件，僅是針對同一筆班級變更開放修改，故不會影響變更次數。
                        flag_can_use_PARTREDUC_btn = If($"{drv("PARTREDUC")}" = TIMS.cst_YES, False, True)
                        If flag_can_use_PARTREDUC_btn Then btnPartReduc.Visible = True '「還原」按鈕
                    End If

                    '產投充飛
                    If TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '產投充飛-署-開放刪除鈕
                        If sm.UserInfo.LID = 0 Then flag_can_use_delete_btn = True
                        '產投充飛-分署-開放刪除鈕
                        If sm.UserInfo.LID = 1 Then flag_can_use_delete_btn = True
                    End If
                    'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then btnDelete.Visible = True  '產投開放刪除鈕
                    'If sm.UserInfo.RID = "A" AndAlso sm.UserInfo.RoleID <= 1 Then btnDelete.Visible = True '署(局)開放刪除鈕
                    If flag_can_use_delete_btn Then btnDelete.Visible = True '刪除鈕

                    If $"{drv("PARTREDUC")}" = TIMS.cst_YES Then
                        '資料還原中，不可刪除，不可審核
                        Dim s_tip1 As String = "資料還原中，不可審核"
                        Dim s_tip2 As String = "資料還原中，不可刪除"
                        btnEdit.Enabled = False
                        btnDelete.Enabled = False
                        TIMS.Tooltip(btnEdit, s_tip1)
                        TIMS.Tooltip(btnDelete, s_tip2)
                    End If
                End If

                '(產投)可檢視／即可修改--'非委訓單位／可修改 cst_dgAct_Updat1
                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso btnView1.Visible AndAlso sm.UserInfo.LID <> 2 Then
                    btnView1.Visible = False '不檢視
                    btnUpdat1.Visible = True '可修改
                End If

                Select Case Convert.ToString(drv("ReviseStatus"))
                    Case "Y" '通過
                        TIMS.Tooltip(btnView1, "通過")
                    Case "N" '不通過
                        TIMS.Tooltip(btnView1, "不通過")
                End Select
                labApplyDate.Text = Common.FormatDate(drv("CDate"))

                '**by Milor 20080507--將變更項目改由年度、是否產學訓判斷，會顯示不同的變更項目名稱。 (變數值改由陣列獲取)
                Dim sChgTypeN1 As String = "未定義" & " (" & CStr(drv("SubSeqNo")) & ")"
                Try
                    sChgTypeN1 = ChgItemName(CInt(drv("AltDataID")) - 1) & " (" & CStr(drv("SubSeqNo")) & ")"
                Catch ex As Exception
                    Dim strErrmsg As String = ""
                    strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg)
                End Try
                labChgTypeN1.Text = sChgTypeN1
                '**by Milor 20080507----end

                If Convert.ToString(drv("isBlack")) = "Y" Then
                    '該機構，已列入處分名單!!
                    'btnDelete.Enabled = False
                    TIMS.Tooltip(btnDelete, cst_isBlackMsg)
                    'btnEdit.Enabled = False
                    TIMS.Tooltip(btnEdit, cst_isBlackMsg)
                End If

            Case ListItemType.Footer
                PlanReviseList.ShowFooter = False
                If PlanReviseList.Items.Count = 0 Then
                    PlanReviseList.ShowFooter = True
                    Dim mycell As New TableCell
                    mycell.ColumnSpan = e.Item.Cells.Count
                    mycell.Text = "目前沒有任何資料!"
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(mycell)
                    e.Item.HorizontalAlign = HorizontalAlign.Center
                End If
        End Select
    End Sub

    '保留查詢值
    Sub KeepSession()
        Dim v_rblReviseStatus As String = TIMS.GetListValue(rblReviseStatus)
        Session(cst_ss_search1) = Nothing
        Dim sSearch As String = ""
        TIMS.SetMyValue(sSearch, "prg", "TC_06_001")
        TIMS.SetMyValue(sSearch, "center", center.Text)
        TIMS.SetMyValue(sSearch, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(sSearch, "TB_career_id", TB_career_id.Text)
        TIMS.SetMyValue(sSearch, "trainValue", trainValue.Value)
        TIMS.SetMyValue(sSearch, "jobValue", jobValue.Value)
        TIMS.SetMyValue(sSearch, "txtCJOB_NAME", txtCJOB_NAME.Text)
        TIMS.SetMyValue(sSearch, "cjobValue", cjobValue.Value)
        TIMS.SetMyValue(sSearch, "ClassCName", ClassCName.Text)
        TIMS.SetMyValue(sSearch, "CyclType", CyclType.Text)
        TIMS.SetMyValue(sSearch, "ApplySDate", ApplySDate.Text)
        TIMS.SetMyValue(sSearch, "ApplyEDate", ApplyEDate.Text)
        TIMS.SetMyValue(sSearch, "rblReviseStatus", v_rblReviseStatus) 'rblReviseStatus.SelectedValue)
        TIMS.SetMyValue(sSearch, "KeepSearch", "1")
        'TIMS.SetMyValue(sSearch, "PageIndex", CStr(PlanReviseList.CurrentPageIndex + 1))
        Session(cst_ss_search1) = sSearch
    End Sub

#Region "(No Use)"

    '刪除資料
    'Sub UPDATE_Delete_PlanRevise(ByVal vsPlanID As String, ByVal vsComIDNO As String, ByVal vsSeqNO As String, ByVal vsCDate As String, ByVal vsSubSeqNO As String, ByVal vsAltDataID As String)

    '    Dim u_Sql As String = ""
    '    u_Sql = ""
    '    u_Sql &= " UPDATE PLAN_REVISE "
    '    u_Sql &= " Set MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE()" & vbCrLf
    '    u_Sql &= " WHERE 1=1 "
    '    u_Sql &= " And PlanID = @PlanID" & vbCrLf
    '    u_Sql &= " And ComIDNO = @ComIDNO" & vbCrLf
    '    u_Sql &= " And SeqNO = @SeqNO" & vbCrLf
    '    u_Sql &= " And CDate = CONVERT(DATETIME, @CDate, 111)" & vbCrLf
    '    u_Sql &= " And SubSeqNO = @SubSeqNO" & vbCrLf
    '    u_Sql &= " And AltDataID = @AltDataID" & vbCrLf

    '    Dim dt1 As DataTable = TIMS.Get_DataTable(objconn, "PLAN_REVISE")
    '    Dim str_COLUMN_1 As String = TIMS.Get_DataTableCOLUMN(dt1) '""

    '    Dim i_Sql As String = ""
    '    i_Sql = "" & vbCrLf
    '    i_Sql &= " INSERT INTO PLAN_REVISEDEL (" & str_COLUMN_1 & ")" & vbCrLf
    '    i_Sql &= " Select " & str_COLUMN_1 & vbCrLf
    '    i_Sql &= " FROM PLAN_REVISE" & vbCrLf
    '    i_Sql &= " WHERE 0=0" & vbCrLf
    '    i_Sql &= " And PlanID = @PlanID" & vbCrLf
    '    i_Sql &= " And ComIDNO = @ComIDNO" & vbCrLf
    '    i_Sql &= " And SeqNO = @SeqNO" & vbCrLf
    '    i_Sql &= " And CDate = CONVERT(DATETIME, @CDate, 111)" & vbCrLf
    '    i_Sql &= " And SubSeqNO = @SubSeqNO" & vbCrLf
    '    i_Sql &= " And AltDataID = @AltDataID" & vbCrLf

    '    Dim d_Sql As String = ""
    '    d_Sql = ""
    '    d_Sql &= " DELETE PLAN_REVISE "
    '    d_Sql &= " WHERE 1=1 "
    '    d_Sql &= " And PlanID = @PlanID" & vbCrLf
    '    d_Sql &= " And ComIDNO = @ComIDNO" & vbCrLf
    '    d_Sql &= " And SeqNO = @SeqNO" & vbCrLf
    '    d_Sql &= " And CDate = CONVERT(DATETIME, @CDate, 111)" & vbCrLf
    '    d_Sql &= " And SubSeqNO = @SubSeqNO" & vbCrLf
    '    d_Sql &= " And AltDataID = @AltDataID" & vbCrLf

    '    Dim u_Parms As New Hashtable
    '    u_Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
    '    u_Parms.Add("PlanID", vsPlanID)
    '    u_Parms.Add("ComIDNO", vsComIDNO)
    '    u_Parms.Add("SeqNO", vsSeqNO)
    '    u_Parms.Add("CDate", TIMS.cdate3(vsCDate))
    '    u_Parms.Add("SubSeqNO", vsSubSeqNO)
    '    u_Parms.Add("AltDataID", vsAltDataID)
    '    DbAccess.ExecuteNonQuery(u_Sql, objconn, u_Parms)  'edit，by:20181012

    '    Dim i_Parms As New Hashtable
    '    i_Parms.Add("PlanID", vsPlanID)
    '    i_Parms.Add("ComIDNO", vsComIDNO)
    '    i_Parms.Add("SeqNO", vsSeqNO)
    '    i_Parms.Add("CDate", TIMS.cdate3(vsCDate))
    '    i_Parms.Add("SubSeqNO", vsSubSeqNO)
    '    i_Parms.Add("AltDataID", vsAltDataID)
    '    DbAccess.ExecuteNonQuery(i_Sql, objconn, i_Parms)  'edit，by:20181012

    '    Dim d_Parms As New Hashtable
    '    d_Parms.Add("PlanID", vsPlanID)
    '    d_Parms.Add("ComIDNO", vsComIDNO)
    '    d_Parms.Add("SeqNO", vsSeqNO)
    '    d_Parms.Add("CDate", TIMS.cdate3(vsCDate))
    '    d_Parms.Add("SubSeqNO", vsSubSeqNO)
    '    d_Parms.Add("AltDataID", vsAltDataID)
    '    DbAccess.ExecuteNonQuery(d_Sql, objconn, d_Parms)  'edit，by:20181012

    'End Sub

    '刪除 DELETE PLAN_TRAINDESC_REVISE DELETE PLAN_TRAINDESC_REVISEITEM
    'Sub Delete_PlanTrainDescRevise_Item(ByVal tmpPTDRID As Integer)
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim sql As String = ""
    '    sql = " DELETE PLAN_TRAINDESC_REVISEITEM WHERE PTDRID = @PTDRID" & vbCrLf
    '    Dim dCmd As New SqlCommand(sql, objconn)
    '    With dCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("PTDRID", SqlDbType.Int).Value = tmpPTDRID
    '        .ExecuteNonQuery()  'edit，by:20181012
    '        'DbAccess.ExecuteNonQuery(dCmd.CommandText, objconn, dCmd.Parameters)  'edit，by:20181012
    '    End With

    '    sql = " DELETE PLAN_TRAINDESC_REVISE WHERE PTDRID = @PTDRID" & vbCrLf
    '    Dim dCmd2 As New SqlCommand(sql, objconn)
    '    With dCmd2
    '        .Parameters.Clear()
    '        .Parameters.Add("PTDRID", SqlDbType.Int).Value = tmpPTDRID
    '        .ExecuteNonQuery()  'edit，by:20181012
    '        'DbAccess.ExecuteNonQuery(dCmd2.CommandText, objconn, dCmd2.Parameters)  'edit，by:20181012
    '    End With
    '    'Try
    '    'Catch ex As Exception
    '    '    Common.MessageBox(Me, ex.ToString)
    '    'End Try
    'End Sub

    'Function KeepSearch() As String
    '    Dim vsSearch1 As String = ""
    '    '090402 andy edit
    '    '---------------
    '    vsSearch1 = ""
    '    If center.Text <> "" Then vsSearch1 += "&center=" & HttpUtility.UrlEncode(center.Text)
    '    If RIDValue.Value <> "" Then vsSearch1 += "&RIDValue=" & RIDValue.Value
    '    If ClassCName.Text <> "" Then vsSearch1 += "&ClassCName=" & HttpUtility.UrlEncode(ClassCName.Text)
    '    If TB_career_id.Text <> "" Then vsSearch1 += "&TB_career_id=" & HttpUtility.UrlEncode(TB_career_id.Text)
    '    If trainValue.Value <> "" Then vsSearch1 += "&trainValue=" & trainValue.Value
    '    If jobValue.Value <> "" Then vsSearch1 += "&jobValue=" & jobValue.Value
    '    If txtCJOB_NAME.Text <> "" Then vsSearch1 += "&txtCJOB_NAME=" & txtCJOB_NAME.Text
    '    If cjobValue.Value <> "" Then vsSearch1 += "&cjobValue=" & cjobValue.Value
    '    If ApplySDate.Text <> "" Then vsSearch1 += "&ApplySDate=" & ApplySDate.Text
    '    If ApplyEDate.Text <> "" Then vsSearch1 += "&ApplyEDate=" & ApplyEDate.Text
    '    'vsSearch1 += IIf(PlanReviseList.CurrentPageIndex <> 0, "&PlanReviseList=" & PlanReviseList.CurrentPageIndex, "&PlanReviseList=0")
    '    vsSearch1 += "&KeepSearch=1"
    '    '---------------
    '    Return vsSearch1
    'End Function

#End Region

    ''' <summary> 匯出鈕 </summary>
    Sub Export1()
        Dim flag_can_ExpType1 As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID = 0, True, False)
        If Not flag_can_ExpType1 Then
            trRBListExpType.Visible = flag_can_ExpType1 'False
            BtnExp1.Visible = flag_can_ExpType1 'False
            Common.MessageBox(Me, "權限不足!!")
            Return
        End If

        'OJT-22022301：<系統> 產投 - 班級變更審核：增加【報名開始日期】、【報名結束日期】兩欄位+匯出功能
        '三、 增加匯出excel或ods功能，匯出欄位
        '計畫別、分署、訓練單位、班別名稱、開訓日期、結訓日期、變更項目、變更說明、申請變更日期、申請變更函送日期、申請變更函送狀態、函送資料是否符合規定、逾期週數、是否納入審查計分變更次數(是/否)、政策性課程不扣分(是/否)、審核狀態(未審核、已審核)、審核時間(含日期及時間， 若該筆未審核則為空值)。
        '計畫別,分署,訓練單位,班別名稱,開訓日期,結訓日期,變更項目,變更說明,申請變更日期,申請變更函送日期,申請變更函送狀態,函送資料是否符合規定,逾期週數,是否納入審查計分變更次數,政策性課程不扣分,審核狀態,審核時間
        '權限僅限署可使用此匯出按鈕。
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Common.MessageBox(Me, "該計畫不提供此功能!!")
            Exit Sub
        End If

        Dim sErrMsg As String = ""
        Call CheckData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        '查詢SQL 語法 
        Dim sql As String = Search_SQL()
        If sql = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim dr1 As DataRow = dt.Rows(0)

        Const cst_TitleS1 As String = "REVISE"
        Dim strFilename1 As String = String.Concat(cst_TitleS1, TIMS.GetDateNo2())
        Dim sPattern As String = ""
        Dim sColumn As String = ""
        Dim sTitle1 As String = "班級變更審核匯出"
        sPattern = "計畫別,分署,訓練單位,班別名稱,申請階段,課程代碼,開訓日期,結訓日期,變更項目,變更說明,申請變更日期,申請變更函送日期,申請變更函送狀態,函送資料是否符合規定"
        sPattern &= ",逾期週數,不納入審查計分變更次數,政策性課程不扣分,審核狀態,審核時間"
        sColumn = "ORGPLANNAME,DISTNAME,ORGNAME,CLASSNAME,APPSTAGE_N,OCPSNO28,STDATE,FDDATE,ALTDATAID_N,REVISECONT,CDATE,SENDDATE4,STATUS4_N,ISPASS4_N"
        sColumn &= ",OVERWEEK4_N,NOINC4_N,NODEDUC4_N,REVISESTATUS_N,REVISEDATE"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dt, parms)
    End Sub

    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles BtnExp1.Click
        Call Export1()
    End Sub

    Protected Sub PlanReviseList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlanReviseList.SelectedIndexChanged

    End Sub
End Class
