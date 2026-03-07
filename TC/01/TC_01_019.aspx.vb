Public Class TC_01_019
    Inherits AuthBasePage 'System.Web.UI.Page

    'Const cst_mmo1 As String = "※確認送出後即鎖定不可再修改!"
    'ORG_TTQSLOCK / ORG_TTQS2

    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(MRqID, TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'labmmo1.Text = cst_mmo1

        If Not IsPostBack Then
            Call cCreate1()
        End If

    End Sub

    ''' <summary>
    ''' false 無開放單位確認
    ''' </summary>
    ''' <returns></returns>
    Function Check_TTQSLOCK() As Boolean
        Dim s_TTQSLOCKMsg1 As String = TIMS.Get_TTQSLOCKMsg1(objconn)
        If s_TTQSLOCKMsg1 = "" Then
            Dim v_msg_1 As String = "目前無開放單位確認!"
            Common.MessageBox(Me, v_msg_1)
            Return False
        End If
        Return True
    End Function

    ''' <summary>
    ''' 年度／月份 要 指定
    ''' </summary>
    Sub SET_YMVALUE()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT YEAR(GETDATE()) YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(YEAR(GETDATE())) YEARS_ROC" & vbCrLf
        sql &= " ,CASE WHEN MONTH(GETDATE())<=6 THEN '01' ELSE '07' END MONTHS" & vbCrLf
        'sql &= " ,CASE WHEN MONTH(GETDATE())<=6 THEN '1月' ELSE '7月' END MONTHS_N" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " select ISNULL(k.YEARS,c.YEARS) YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(ISNULL(k.YEARS,c.YEARS)) YEARS_ROC" & vbCrLf
        sql &= " ,ISNULL(k.MONTHS,c.MONTHS) MONTHS" & vbCrLf
        sql &= " ,CASE ISNULL(k.MONTHS,c.MONTHS) WHEN '01' THEN '1月' WHEN '07' THEN '7月' END MONTHS_N" & vbCrLf
        sql &= " FROM ORG_TTQSLOCK k WITH(NOLOCK)" & vbCrLf
        sql &= " CROSS JOIN WC1 c" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND (k.CSDATE <= GETDATE() AND k.CFDATE >= GETDATE())" & vbCrLf
        sql &= " AND k.ISDELETE IS NULL" & vbCrLf
        sql &= " ORDER BY k.CCDATE DESC,k.MODIFYDATE DESC" & vbCrLf
        Dim s_sql As String = sql
        Dim dr1 As DataRow = DbAccess.GetOneRow(s_sql, objconn)
        If dr1 Is Nothing Then
            'Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " SELECT YEAR(GETDATE()) YEARS" & vbCrLf
            sql &= " ,dbo.FN_CYEAR2(YEAR(GETDATE())) YEARS_ROC" & vbCrLf
            sql &= " ,CASE WHEN MONTH(GETDATE())<=6 THEN '01' ELSE '07' END MONTHS" & vbCrLf
            sql &= " ,CASE WHEN MONTH(GETDATE())<=6 THEN '1月' ELSE '7月' END MONTHS_N" & vbCrLf
            s_sql = sql
            dr1 = DbAccess.GetOneRow(s_sql, objconn)
        End If
        YEARS_ROC.Text = Convert.ToString(dr1("YEARS_ROC")) '年度N
        MONTHS.Text = Convert.ToString(dr1("MONTHS")) '月份N
        Hid_YEARS.Value = Convert.ToString(dr1("YEARS")) '年度
        Hid_MONTHS.Value = Convert.ToString(dr1("MONTHS")) '月份
        'PLANNAME.Text = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn) '(大)計畫
        PLANNAME.Text = TIMS.GetPlanName(sm.UserInfo.PlanID, objconn) '計畫
        'DISTNAME.Text = TIMS.GET_DistName(sm.UserInfo.DistID, objconn) '分署
    End Sub

    Sub cCreate1()
        lab_msg1.Text = ""

        '年度／月份 要 指定
        SET_YMVALUE()

        If Not Check_TTQSLOCK() Then Return

        With rblCONFIRM
            .Items.Clear()
            .Items.Add(New ListItem("正確", "Y"))
            .Items.Add(New ListItem("資料有誤(原因說明)", "N"))
        End With
        rblCONFIRM.AppendDataBoundItems = True

        Hid_ORGID.Value = sm.UserInfo.OrgID

        Loaddata1(Hid_ORGID.Value)
    End Sub

    Sub Loaddata1(ByVal vORGID As String)
        BTN_EDIT1.Enabled = True
        BTN_SAVE1.Enabled = True
        If vORGID = "" Then
            Hid_OTTID.Value = "" 'Convert.ToString(dr("OTTID"))
            Hid_ORGID.Value = "" 'Convert.ToString(dr("ORGID"))
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.OTTID" & vbCrLf
        sql &= " ,a.ORGID" & vbCrLf
        sql &= " ,o.ORGNAME" & vbCrLf '機構名稱
        sql &= " ,o.ORGKIND" & vbCrLf
        sql &= " ,k1.NAME ORGKIND_N" & vbCrLf '機構別
        sql &= " ,a.COMIDNO" & vbCrLf '統一編號
        sql &= " ,a.YEARS" & vbCrLf '年度
        sql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,a.MONTHS" & vbCrLf '月份
        sql &= " ,a.SENDVER" & vbCrLf
        sql &= " ,a.RESULT" & vbCrLf
        sql &= " ,v1.VNAME SENDVER_N" & vbCrLf '評核版別
        sql &= " ,v2.VNAME RESULT_N" & vbCrLf '評核結果 
        sql &= " ,a.EXTLICENS" & vbCrLf
        sql &= " ,a.GOAL" & vbCrLf '申請目的
        sql &= " ,a.EXTEND" & vbCrLf '展延
        sql &= " ,FORMAT(a.IMPORTDATE,'yyyy/MM/dd HH:mm:ss') IMPORTDATE" & vbCrLf '轉入／資料更新時間
        sql &= " ,FORMAT(a.SENDDATE,'yyyy/MM/dd') SENDDATE" & vbCrLf '評核日期
        sql &= " ,FORMAT(a.ISSUEDATE,'yyyy/MM/dd') ISSUEDATE" & vbCrLf '發文日期 
        sql &= " ,FORMAT(a.VALIDDATE,'yyyy/MM/dd') VALIDDATE" & vbCrLf '有效期限 
        sql &= " ,a.MEMO" & vbCrLf
        sql &= " ,a.EVALSCOPE" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(a.EVALSCOPE,'-')='-' then a.MEMO ELSE concat(a.MEMO,'(',a.EVALSCOPE,')') END MEMO2" & vbCrLf
        sql &= " ,a.IMPORTACCT" & vbCrLf
        sql &= " ,a.CONFIRM" & vbCrLf
        sql &= " ,a.CONFIRMACCT" & vbCrLf
        sql &= " ,a.CONFIRMDATE" & vbCrLf
        sql &= " ,a.REASON1" & vbCrLf
        sql &= " ,a.APPLIEDRESULT" & vbCrLf
        sql &= " ,a.APPLIEDACCT" & vbCrLf
        sql &= " ,a.APPLIEDDATE" & vbCrLf
        sql &= " ,a.LOCKACCT" & vbCrLf
        sql &= " ,a.LOCKDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.ORG_TTQS2 a" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO o on o.OrgID=a.OrgID" & vbCrLf
        sql &= " LEFT JOIN dbo.V_SENDVER v1 on v1.VID=a.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.V_RESULT v2 on v2.VID=a.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_ORGTYPE k1 on k1.ORGTYPEID=o.ORGKIND" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and a.TPLANID=@TPLANID" & vbCrLf
        sql &= " and a.DISTID=@DISTID" & vbCrLf
        sql &= " and a.ORGID=@ORGID" & vbCrLf
        sql &= " and a.YEARS=@YEARS" & vbCrLf
        sql &= " and a.MONTHS=@MONTHS" & vbCrLf
        sql &= " ORDER BY a.YEARS DESC,a.MONTHS DESC" & vbCrLf

        Hid_YEARS.Value = TIMS.ClearSQM(Hid_YEARS.Value)
        Hid_MONTHS.Value = TIMS.ClearSQM(Hid_MONTHS.Value)
        Dim dt As DataTable
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("DISTID", sm.UserInfo.DistID)
        parms.Add("ORGID", Val(Hid_ORGID.Value))
        'parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("YEARS", Hid_YEARS.Value) '(指定)
        parms.Add("MONTHS", Hid_MONTHS.Value) '(指定)

        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then
            Hid_OTTID.Value = "" 'Convert.ToString(dr("OTTID"))
            Hid_ORGID.Value = "" 'Convert.ToString(dr("ORGID"))
            lab_msg1.Text = TIMS.cst_NODATAMsg1
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            BTN_EDIT1.Enabled = False
            BTN_SAVE1.Enabled = False
            TIMS.Tooltip(BTN_EDIT1, TIMS.cst_NODATAMsg1)
            TIMS.Tooltip(BTN_SAVE1, TIMS.cst_NODATAMsg1)
            Return
        End If
        Dim dr As DataRow = dt.Rows(0)

        lab_msg1.Text = ""
        Hid_OTTID.Value = Convert.ToString(dr("OTTID"))
        Hid_ORGID.Value = Convert.ToString(dr("ORGID"))

        YEARS_ROC.Text = Convert.ToString(dr("YEARS_ROC")) '年度
        MONTHS.Text = Convert.ToString(dr("MONTHS")) '月份
        IMPORTDATE.Text = Convert.ToString(dr("IMPORTDATE")) '資料更新時間
        ORGNAME.Text = Convert.ToString(dr("ORGNAME")) '機構名稱
        COMIDNO.Text = Convert.ToString(dr("COMIDNO")) '統一編號

        'ORGKIND.Text = Convert.ToString(dr("ORGKIND")) '機構別
        ORGKIND_N.Text = Convert.ToString(dr("ORGKIND_N")) '機構別

        'SENDVER.Text = Convert.ToString(dr("SENDVER")) '評核版別
        SENDVER_N.Text = Convert.ToString(dr("SENDVER_N")) '評核版別
        GOAL.Text = Convert.ToString(dr("GOAL")) '申請目的

        'RESULT.Text = Convert.ToString(dr("RESULT")) '評核結果
        RESULT_N.Text = Convert.ToString(dr("RESULT_N")) '評核結果

        EXTLICENS.Text = Convert.ToString(dr("EXTLICENS")) '展延
        SENDDATE.Text = Convert.ToString(dr("SENDDATE")) '評核日期

        SENDDATE.Text = Convert.ToString(dr("SENDDATE")) '機構別
        VALIDDATE.Text = Convert.ToString(dr("VALIDDATE"))
        EXTLICENS.Text = Convert.ToString(dr("EXTLICENS"))

        ISSUEDATE.Text = Convert.ToString(dr("ISSUEDATE")) '發文日期
        VALIDDATE.Text = Convert.ToString(dr("VALIDDATE")) '有效期限
        'MEMO.Text = Convert.ToString(dr("MEMO")) 'TTQS訓練機構名稱
        MEMO2.Text = Convert.ToString(dr("MEMO2")) 'TTQS訓練機構名稱

        Common.SetListItem(rblCONFIRM, Convert.ToString(dr("CONFIRM"))) '單位確認
        txREASON1.Text = Convert.ToString(dr("REASON1")) '原因說明

        Dim flag_is_lock As Boolean = False
        If Convert.ToString(dr("LOCKACCT")) <> "" AndAlso Convert.ToString(dr("LOCKACCT")) <> "" Then
            flag_is_lock = True
        End If

        rblCONFIRM.Enabled = True
        txREASON1.Enabled = True
        BTN_EDIT1.Enabled = True
        BTN_SAVE1.Enabled = True
        If flag_is_lock Then
            Dim v_lock_msg1 As String = "已按確認送出!!"
            TIMS.Tooltip(rblCONFIRM, v_lock_msg1)
            TIMS.Tooltip(txREASON1, v_lock_msg1)
            TIMS.Tooltip(BTN_EDIT1, v_lock_msg1)
            TIMS.Tooltip(BTN_SAVE1, v_lock_msg1)
            rblCONFIRM.Enabled = False
            txREASON1.Enabled = False
            BTN_EDIT1.Enabled = False
            BTN_SAVE1.Enabled = False
        End If

    End Sub

    Function CheckData1(ByRef s_ERRMSG As String) As Boolean
        Dim rst As Boolean = True
        s_ERRMSG = ""
        If Hid_ORGID.Value = "" Then s_ERRMSG &= "請確認 未讀取到有效資料!" & vbCrLf
        txREASON1.Text = TIMS.ClearSQM(txREASON1.Text)
        Dim v_rblCONFIRM As String = TIMS.GetListValue(rblCONFIRM)
        If v_rblCONFIRM = "" Then s_ERRMSG &= "請選擇 單位確認結果!" & vbCrLf
        Select Case v_rblCONFIRM
            Case "N"
                If txREASON1.Text = "" Then
                    s_ERRMSG &= "單位確認有誤 原因說明為必填!" & vbCrLf
                End If
        End Select
        If (s_ERRMSG <> "") Then rst = False
        Return rst
    End Function

    Sub SaveData1(ByVal s_LOCK As String)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE ORG_TTQS2" & vbCrLf
        sql &= " SET CONFIRM=@CONFIRM" & vbCrLf
        sql &= " ,CONFIRMACCT=@CONFIRMACCT" & vbCrLf
        sql &= " ,CONFIRMDATE=GETDATE()" & vbCrLf
        sql &= " ,REASON1=@REASON1" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        If s_LOCK = "Y" Then
            sql &= " ,LOCKACCT=@LOCKACCT" & vbCrLf
            sql &= " ,LOCKDATE=GETDATE()" & vbCrLf
        End If
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND OTTID=@OTTID" & vbCrLf
        sql &= " AND ORGID=@ORGID" & vbCrLf
        sql &= " AND TPLANID=@TPLANID" & vbCrLf
        sql &= " AND DISTID=@DISTID" & vbCrLf
        sql &= " and YEARS=@YEARS" & vbCrLf
        sql &= " and MONTHS=@MONTHS" & vbCrLf
        Dim u_sql As String = sql

        Hid_YEARS.Value = TIMS.ClearSQM(Hid_YEARS.Value)
        Hid_MONTHS.Value = TIMS.ClearSQM(Hid_MONTHS.Value)
        Dim v_rblCONFIRM As String = TIMS.GetListValue(rblCONFIRM)
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("CONFIRM", v_rblCONFIRM)
        parms.Add("CONFIRMACCT", sm.UserInfo.UserID)
        parms.Add("REASON1", txREASON1.Text)
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        If s_LOCK = "Y" Then parms.Add("LOCKACCT", sm.UserInfo.UserID)
        parms.Add("OTTID", Val(Hid_OTTID.Value))
        parms.Add("ORGID", Val(Hid_ORGID.Value))
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("DISTID", sm.UserInfo.DistID)
        'parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("YEARS", Hid_YEARS.Value) '(指定)
        parms.Add("MONTHS", Hid_MONTHS.Value) '(指定)
        DbAccess.ExecuteNonQuery(u_sql, objconn, parms)

        Dim s_msg_1 As String = "儲存成功!"
        If s_LOCK = "Y" AndAlso v_rblCONFIRM = "N" Then s_msg_1 = "請單位檢附相關資料送至分署確認!"

        '儲存成功
        Dim url1 As String = "TC_01_019.aspx?id=" & TIMS.Get_MRqID(Me)
        Common.MessageBox(Me, s_msg_1)
        TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub

    Protected Sub BTN_EDIT1_Click(sender As Object, e As EventArgs) Handles BTN_EDIT1.Click
        If Not Check_TTQSLOCK() Then Return

        Dim s_ERRMSG As String = ""
        Dim rst As Boolean = CheckData1(s_ERRMSG)
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, s_ERRMSG)
            Return
        End If

        SaveData1("")

    End Sub

    Protected Sub BTN_SAVE1_Click(sender As Object, e As EventArgs) Handles BTN_SAVE1.Click
        If Not Check_TTQSLOCK() Then Return

        Dim s_ERRMSG As String = ""
        Dim rst As Boolean = CheckData1(s_ERRMSG)
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, s_ERRMSG)
            Return
        End If

        SaveData1("Y")

    End Sub
End Class