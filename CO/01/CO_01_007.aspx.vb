Public Class CO_01_007
    Inherits AuthBasePage 'System.Web.UI.Page

    '[ORG_TTQSQUERY]
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call CCreate1() '頁面初始化
            '20190103 若有先前查詢條件記錄，則將資料重新讀取到頁面中
            Call USE_KeepSearch1()
        End If
    End Sub

    '頁面初始化
    Sub CCreate1()
        lab_msg1.Text = ""
        tb_Sch.Visible = False
        lab_msg2.Text = ""

        'LOAD1
        ddlYEARS_S1 = TIMS.Get_ROCYEARS2(ddlYEARS_S1, objconn)
        ddlYEARS_S1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose2, "")) '"==請選擇==" 
        Common.SetListItem(ddlAPPSTAGE_S1, sm.UserInfo.Years)
        ddlAPPSTAGE_S1 = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_S1)
        Common.SetListItem(ddlAPPSTAGE_S1, "")
        ddlTTQSLOCK_S1 = TIMS.GET_ddlTTQSLOCK(ddlTTQSLOCK_S1, objconn)
        ddlTTQSLOCK_S1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose2, "")) '"==請選擇==" 

        'DATA1
        ddlYEARS = TIMS.Get_ROCYEARS2(ddlYEARS, objconn)
        ddlAPPSTAGE = TIMS.Get_APPSTAGE2(ddlAPPSTAGE)
        Common.SetListItem(ddlAPPSTAGE, "1")
        ddlTTQSLOCK = TIMS.GET_ddlTTQSLOCK(ddlTTQSLOCK, objconn)

        Call TIMS.SUB_SET_HR_MI(ddlSPREDATE_HH, ddlSPREDATE_MM)
        Call TIMS.SUB_SET_HR_MI(ddlFPREDATE_HH, ddlFPREDATE_MM)

        Call TIMS.SUB_SET_HR_MI(ddlQSDATE_HH, ddlQSDATE_MM)
        Call TIMS.SUB_SET_HR_MI(ddlQFDATE_HH, ddlQFDATE_MM)

        Utl_SHOW(0)
    End Sub
    '資料查詢
    Sub SSearch1()
        Dim v_ddlYEARS_S1 As String = TIMS.GetListValue(ddlYEARS_S1)
        Dim v_ddlAPPSTAGE_S1 As String = TIMS.GetListValue(ddlAPPSTAGE_S1)
        Dim v_ddlTTQSLOCK_S1 As String = TIMS.GetListValue(ddlTTQSLOCK_S1)
        schQCDATE1.Text = TIMS.ClearSQM(schQCDATE1.Text)
        schQCDATE2.Text = TIMS.ClearSQM(schQCDATE2.Text)
        Dim v_schQCDATE1 As String = TIMS.Cdate3(schQCDATE1.Text)
        Dim v_schQCDATE2 As String = TIMS.Cdate3(schQCDATE2.Text)

        Dim pParms As New Hashtable()
        '審查計分區間
        Dim sSql As String = "
SELECT a.OTQID
,format(a.QCDATE,'yyyy/MM/dd') QCDATE
,format(a.QSDATE,'yyyy/MM/dd HH:mm') QSDATE,format(a.QFDATE,'yyyy/MM/dd HH:mm') QFDATE,a.QEXPLAIN
,format(a.SPREDATE,'yyyy/MM/dd HH:mm') SPREDATE,format(a.FPREDATE,'yyyy/MM/dd HH:mm') FPREDATE,a.REMIND1
,a.ISDELETE,a.YEARS
,a.APPSTAGE ,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N
,a.OTLID,dbo.FN_GET_TTQSLOCK_N(a.OTLID) TTQSLOCK_N
FROM ORG_TTQSQUERY a
WHERE a.ISDELETE IS NULL
"
        '包含已刪除
        'sSql &= If(CHK_ISDELETE.Checked, " WHERE 1=1", " WHERE a.ISDELETE IS NULL")
        If v_ddlYEARS_S1 <> "" Then
            pParms.Add("YEARS", v_ddlYEARS_S1)
            sSql &= " AND a.YEARS=@YEARS" & vbCrLf
        End If
        If v_ddlAPPSTAGE_S1 <> "" Then
            pParms.Add("APPSTAGE", v_ddlAPPSTAGE_S1)
            sSql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        End If
        If v_ddlTTQSLOCK_S1 <> "" Then
            pParms.Add("OTLID", v_ddlTTQSLOCK_S1)
            sSql &= " AND a.OTLID=@OTLID" & vbCrLf
        End If
        If v_schQCDATE1 <> "" Then
            pParms.Add("schQCDATE1", TIMS.Cdate2(v_schQCDATE1))
            sSql &= " AND a.QCDATE >= @schQCDATE1" & vbCrLf
        End If
        If v_schQCDATE2 <> "" Then
            pParms.Add("schQCDATE2", TIMS.Cdate2(v_schQCDATE2))
            sSql &= " AND a.QCDATE <= @schQCDATE2" & vbCrLf
        End If
        sSql &= " ORDER BY a.QCDATE DESC " & vbCrLf

        lab_msg1.Text = "查無資料"
        tb_Sch.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        If TIMS.dtNODATA(dt) Then Return

        lab_msg1.Text = ""
        tb_Sch.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>'資料讀取</summary>
    ''' <param name="iOTQID"></param>
    Private Sub LoadData1(ByVal iOTQID As Integer)
        Dim pParms As New Hashtable From {{"OTQID", iOTQID}}
        Dim sSql As String = "
SELECT a.OTQID
,format(a.QCDATE,'yyyy/MM/dd') QCDATE
,format(a.QSDATE,'yyyy/MM/dd') QSDATE,format(a.QSDATE,'HH') QSDATE_HH ,format(a.QSDATE,'mm') QSDATE_MM
,format(a.QFDATE,'yyyy/MM/dd') QFDATE,format(a.QFDATE,'HH') QFDATE_HH,format(a.QFDATE,'mm') QFDATE_MM
,a.QEXPLAIN
,format(a.SPREDATE,'yyyy/MM/dd') SPREDATE,format(a.SPREDATE,'HH') SPREDATE_HH ,format(a.SPREDATE,'mm') SPREDATE_MM
,format(a.FPREDATE,'yyyy/MM/dd') FPREDATE,format(a.FPREDATE,'HH') FPREDATE_HH,format(a.FPREDATE,'mm') FPREDATE_MM
,a.REMIND1
,a.ISDELETE,a.YEARS,a.APPSTAGE
,a.OTLID,dbo.FN_GET_TTQSLOCK_N(a.OTLID) TTQSLOCK_N
FROM ORG_TTQSQUERY a
WHERE a.OTQID=@OTQID
"
        'sql &= " AND a.ISDELETE IS NULL" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim dr As DataRow = dt.Rows(0)
        txQCDATE.Text = TIMS.Cdate3(dr("QCDATE"))
        Common.SetListItem(ddlYEARS, Convert.ToString(dr("YEARS")))
        Common.SetListItem(ddlAPPSTAGE, Convert.ToString(dr("APPSTAGE")))
        Common.SetListItem(ddlTTQSLOCK, Convert.ToString(dr("OTLID")))

        txSPREDATE.Text = TIMS.Cdate3(dr("SPREDATE"))
        Common.SetListItem(ddlSPREDATE_HH, Convert.ToString(dr("SPREDATE_HH")))
        Common.SetListItem(ddlSPREDATE_MM, Convert.ToString(dr("SPREDATE_MM")))
        txFPREDATE.Text = TIMS.Cdate3(dr("FPREDATE"))
        Common.SetListItem(ddlFPREDATE_HH, Convert.ToString(dr("FPREDATE_HH")))
        Common.SetListItem(ddlFPREDATE_MM, Convert.ToString(dr("FPREDATE_MM")))
        txREMIND1.Text = $"{dr("REMIND1")}"

        txQSDATE.Text = TIMS.Cdate3(dr("QSDATE"))
        Common.SetListItem(ddlQSDATE_HH, Convert.ToString(dr("QSDATE_HH")))
        Common.SetListItem(ddlQSDATE_MM, Convert.ToString(dr("QSDATE_MM")))
        txQFDATE.Text = TIMS.Cdate3(dr("QFDATE"))
        Common.SetListItem(ddlQFDATE_HH, Convert.ToString(dr("QFDATE_HH")))
        Common.SetListItem(ddlQFDATE_MM, Convert.ToString(dr("QFDATE_MM")))
        txQEXPLAIN.Text = $"{dr("QEXPLAIN")}"

        '包含已刪除
        LabISDELETE.Visible = (Convert.ToString(dr("ISDELETE")) = "Y")
        '已刪除資料
        bt_save.Enabled = If(Convert.ToString(dr("ISDELETE")) = "Y", False, True)
        TIMS.Tooltip(bt_save, If(Not bt_save.Enabled, "已刪除資料", ""), True)
    End Sub


    Sub Utl_SHOW(ByVal iType As Integer)
        div_sch1.Visible = False
        div_edit.Visible = False
        Select Case iType
            Case 0
                div_sch1.Visible = True
                Exit Sub
            Case 1
                div_edit.Visible = True
                Exit Sub
        End Select

    End Sub

    Sub Utl_Clear()
        txQCDATE.Text = TIMS.Cdate3(Now.Date)
        Dim dr1 As DataRow = TIMS.Get_drYMVALUE1(objconn)
        Common.SetListItem(ddlYEARS, dr1("YEARS"))
        Common.SetListItem(ddlAPPSTAGE, "1")
        Common.SetListItem(ddlTTQSLOCK, "")

        txSPREDATE.Text = ""
        Call TIMS.SUB_SET_HR_MI(ddlSPREDATE_HH, ddlSPREDATE_MM)
        txFPREDATE.Text = ""
        Call TIMS.SUB_SET_HR_MI(ddlFPREDATE_HH, ddlFPREDATE_MM)
        txREMIND1.Text = ""

        txQSDATE.Text = ""
        Call TIMS.SUB_SET_HR_MI(ddlQSDATE_HH, ddlQSDATE_MM)
        txQFDATE.Text = ""
        Call TIMS.SUB_SET_HR_MI(ddlQFDATE_HH, ddlQFDATE_MM)
        txQEXPLAIN.Text = ""

        Hid_OTQID.Value = ""
        '包含已刪除
        LabISDELETE.Visible = False
    End Sub

    ''' <summary>查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Utl_SHOW(0)
        SSearch1()
    End Sub

    ''' <summary>新增</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_add_Click(sender As Object, e As EventArgs) Handles bt_add.Click
        Dim Errmsg As String = ""
        If sm.UserInfo.LID <> 0 Then
            Errmsg &= "權限不足 不可新增，請連絡系統管理者!" & vbCrLf
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Utl_SHOW(1)
        Utl_Clear()
    End Sub

    '送出前檢核 ---> SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        'Dim rst As Boolean=True
        Errmsg = ""

        'txQCDATE'txQSDATE'ddlQSDATE_HH'ddlQSDATE_MM'txQFDATE'ddlQFDATE_HH'ddlQFDATE_MM
        'ddlYEARS '計畫年度
        'ddlAPPSTAGE '申請階段
        'ddlSCORING '審查計分區間
        'txQEXPLAIN '控制說明／作業提醒
        txQCDATE.Text = TIMS.ClearSQM(txQCDATE.Text)
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_ddlAPPSTAGE As String = TIMS.GetListValue(ddlAPPSTAGE)
        Dim v_ddlTTQSLOCK As String = TIMS.GetListValue(ddlTTQSLOCK)

        txSPREDATE.Text = TIMS.ClearSQM(txSPREDATE.Text)
        txFPREDATE.Text = TIMS.ClearSQM(txFPREDATE.Text)
        txREMIND1.Text = TIMS.ClearSQM(txREMIND1.Text)

        txQSDATE.Text = TIMS.ClearSQM(txQSDATE.Text)
        txQFDATE.Text = TIMS.ClearSQM(txQFDATE.Text)
        txQEXPLAIN.Text = TIMS.ClearSQM(txQEXPLAIN.Text)

        If txQCDATE.Text = "" Then Errmsg &= "設定日期 不可為空" & vbCrLf
        If v_ddlYEARS = "" Then Errmsg &= "請選擇 計畫年度" & vbCrLf
        If v_ddlAPPSTAGE = "" Then Errmsg &= "請選擇 申請階段" & vbCrLf
        If v_ddlTTQSLOCK = "" Then Errmsg &= "請選擇 審查計分區間" & vbCrLf

        If txSPREDATE.Text = "" Then Errmsg &= "(初審)開放時間 不可為空" & vbCrLf
        If txFPREDATE.Text = "" Then Errmsg &= "(初審)結束時間 不可為空" & vbCrLf
        If txREMIND1.Text = "" Then Errmsg &= "作業提醒 不可為空" & vbCrLf

        If txQSDATE.Text = "" Then Errmsg &= "控制起始日期 不可為空" & vbCrLf
        If txQFDATE.Text = "" Then Errmsg &= "控制結束日期 不可為空" & vbCrLf
        If txQEXPLAIN.Text = "" Then Errmsg &= "控制說明／作業提醒 不可為空" & vbCrLf
        If Errmsg <> "" Then Return False

        Dim s_SPREDATE As String = TIMS.GET_DateHM(txSPREDATE, ddlSPREDATE_HH, ddlSPREDATE_MM)
        Dim s_FPREDATE As String = TIMS.GET_DateHM(txFPREDATE, ddlFPREDATE_HH, ddlFPREDATE_MM)
        If DateDiff(DateInterval.Minute, CDate(s_SPREDATE), CDate(s_FPREDATE)) = 0 Then Errmsg &= "(初審)開放時間與(初審)結束時間 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(s_SPREDATE), CDate(s_FPREDATE)) < 0 Then Errmsg &= "(初審)開放時間與(初審)結束時間 順序異常!!" & vbCrLf

        If DateDiff(DateInterval.Minute, CDate(txQCDATE.Text), CDate(s_SPREDATE)) < 0 Then Errmsg &= "設定日期 與(初審)開放時間 順序異常!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(txQCDATE.Text), CDate(s_FPREDATE)) = 0 Then Errmsg &= "設定日期 與(初審)結束時間 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(txQCDATE.Text), CDate(s_FPREDATE)) < 0 Then Errmsg &= "設定日期 與(初審)結束時間 順序異常!!" & vbCrLf

        Dim s_txQSDATE As String = TIMS.GET_DateHM(txQSDATE, ddlQSDATE_HH, ddlQSDATE_MM)
        Dim s_txQFDATE As String = TIMS.GET_DateHM(txQFDATE, ddlQFDATE_HH, ddlQFDATE_MM)
        If DateDiff(DateInterval.Minute, CDate(s_txQSDATE), CDate(s_txQFDATE)) = 0 Then Errmsg &= "控制起始日期與控制結束日期 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(s_txQSDATE), CDate(s_txQFDATE)) < 0 Then Errmsg &= "控制起始日期與控制結束日期 順序異常!!" & vbCrLf

        If DateDiff(DateInterval.Minute, CDate(txQCDATE.Text), CDate(s_txQSDATE)) < 0 Then Errmsg &= "設定日期 與控制起始日期 順序異常!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(txQCDATE.Text), CDate(s_txQFDATE)) = 0 Then Errmsg &= "設定日期 與控制結束日期 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(txQCDATE.Text), CDate(s_txQFDATE)) < 0 Then Errmsg &= "設定日期 與控制結束日期 順序異常!!" & vbCrLf

        If sm.UserInfo.LID <> 0 Then
            Errmsg &= "權限不足 不可儲存，請連絡系統管理者!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        Hid_OTQID.Value = TIMS.ClearSQM(Hid_OTQID.Value)
        Dim parms As New Hashtable
        Dim sql As String = ""
        If Hid_OTQID.Value <> "" Then
            'UPDATE 'ORG_TTQSLOCK
            parms.Clear()
            parms.Add("OTQID", Val(Hid_OTQID.Value))
            parms.Add("YEARS", v_ddlYEARS)
            parms.Add("APPSTAGE", v_ddlAPPSTAGE)
            sql = "SELECT YEARS,APPSTAGE FROM ORG_TTQSQUERY WHERE OTQID!=@OTQID AND YEARS=@YEARS AND APPSTAGE=@APPSTAGE" & vbCrLf
        Else
            'INSERT
            parms.Clear()
            parms.Add("YEARS", v_ddlYEARS)
            parms.Add("APPSTAGE", v_ddlAPPSTAGE)
            sql = " SELECT YEARS,APPSTAGE FROM ORG_TTQSQUERY WHERE YEARS=@YEARS AND APPSTAGE=@APPSTAGE" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then
            Errmsg &= "計畫年度／申請階段 已存在，請使用修改功能!!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        Dim parms2 As New Hashtable
        Dim sql2 As String = ""
        If Hid_OTQID.Value <> "" Then
            parms2.Clear()
            parms2.Add("OTQID", Val(Hid_OTQID.Value))
            parms2.Add("OTLID", Val(v_ddlTTQSLOCK))
            sql2 = "SELECT OTQID,OTLID FROM ORG_TTQSQUERY WHERE OTQID!=@OTQID AND OTLID=@OTLID" & vbCrLf
        Else
            parms2.Clear()
            parms2.Add("OTLID", Val(v_ddlTTQSLOCK))
            sql2 = "SELECT OTQID,OTLID FROM ORG_TTQSQUERY WHERE OTLID=@OTLID" & vbCrLf
        End If
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, parms2)
        If dt2.Rows.Count > 0 Then
            Errmsg &= "審查計分區間 已存在，請使用修改功能!!" & vbCrLf
        End If
        Return If(Errmsg <> "", False, True)
    End Function

    '儲存(part-1)
    Sub SaveData1()
        Dim flagSaveOK1 As Boolean = True

        txQCDATE.Text = TIMS.Cdate3(txQCDATE.Text)
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_ddlAPPSTAGE As String = TIMS.GetListValue(ddlAPPSTAGE)
        Dim v_ddlTTQSLOCK As String = TIMS.GetListValue(ddlTTQSLOCK)

        Dim s_SPREDATE As String = TIMS.GET_DateHM(txSPREDATE, ddlSPREDATE_HH, ddlSPREDATE_MM)
        Dim s_FPREDATE As String = TIMS.GET_DateHM(txFPREDATE, ddlFPREDATE_HH, ddlFPREDATE_MM)
        txREMIND1.Text = TIMS.ClearSQM(txREMIND1.Text)

        Dim s_txQSDATE As String = TIMS.GET_DateHM(txQSDATE, ddlQSDATE_HH, ddlQSDATE_MM)
        Dim s_txQFDATE As String = TIMS.GET_DateHM(txQFDATE, ddlQFDATE_HH, ddlQFDATE_MM)
        txQEXPLAIN.Text = TIMS.ClearSQM(txQEXPLAIN.Text)

        Hid_OTQID.Value = TIMS.ClearSQM(Hid_OTQID.Value)
        Dim iOTQID As Integer = 0
        Dim iRst As Integer = 0
        If Hid_OTQID.Value = "" Then
            '新增
            iOTQID = DbAccess.GetNewId(objconn, "ORG_TTQSQUERY_OTQID_SEQ,ORG_TTQSQUERY,OTQID")

            Dim iParms As New Hashtable From {
                {"OTQID", iOTQID},
                {"QCDATE", txQCDATE.Text},
                {"QSDATE", s_txQSDATE},
                {"QFDATE", s_txQFDATE},
                {"QEXPLAIN", txQEXPLAIN.Text},
                {"SPREDATE", s_SPREDATE},
                {"FPREDATE", s_FPREDATE},
                {"REMIND1", txREMIND1.Text},
                {"YEARS", v_ddlYEARS},
                {"APPSTAGE", Val(v_ddlAPPSTAGE)},
                {"OTLID", Val(v_ddlTTQSLOCK)},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim isSql As String = "
INSERT INTO ORG_TTQSQUERY(OTQID,QCDATE,QSDATE,QFDATE,QEXPLAIN,SPREDATE,FPREDATE,REMIND1
,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE, YEARS,APPSTAGE,OTLID)
VALUES (@OTQID,@QCDATE,@QSDATE,@QFDATE,@QEXPLAIN,@SPREDATE,@FPREDATE,@REMIND1
,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE(), @YEARS,@APPSTAGE,@OTLID)
"
            iRst += DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
            Hid_OTQID.Value = iOTQID
        Else
            iOTQID = TIMS.CINT1(Hid_OTQID.Value)
            '修改
            Dim uParms As New Hashtable From {
                {"QCDATE", txQCDATE.Text},
                {"QSDATE", s_txQSDATE},
                {"QFDATE", s_txQFDATE},
                {"QEXPLAIN", txQEXPLAIN.Text},
                {"SPREDATE", s_SPREDATE},
                {"FPREDATE", s_FPREDATE},
                {"REMIND1", txREMIND1.Text},
                {"YEARS", v_ddlYEARS},
                {"APPSTAGE", TIMS.CINT1(v_ddlAPPSTAGE)},
                {"OTLID", TIMS.CINT1(v_ddlTTQSLOCK)},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"OTQID", iOTQID}
            }
            Dim usSql As String = "
UPDATE ORG_TTQSQUERY
SET QCDATE=@QCDATE,QSDATE=@QSDATE,QFDATE=@QFDATE,QEXPLAIN=@QEXPLAIN ,SPREDATE=@SPREDATE,FPREDATE=@FPREDATE,REMIND1=@REMIND1
,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE(),YEARS=@YEARS,APPSTAGE=@APPSTAGE,OTLID=@OTLID
WHERE OTQID=@OTQID
"
            iRst += DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        End If

        If flagSaveOK1 Then
            '儲存成功
            Dim url1 As String = $"CO_01_007.aspx?id={TIMS.Get_MRqID(Me)}"
            Common.MessageBox(Me, "儲存成功!")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

    End Sub
    Sub KeepSearch1()
        Dim V_ddlYEARS_S1 As String = TIMS.GetListValue(ddlYEARS_S1)
        Dim V_ddlAPPSTAGE_S1 As String = TIMS.GetListValue(ddlAPPSTAGE_S1)
        Dim V_ddlTTQSLOCK_S1 As String = TIMS.GetListValue(ddlTTQSLOCK_S1)
        schQCDATE1.Text = TIMS.ClearSQM(schQCDATE1.Text)
        schQCDATE2.Text = TIMS.ClearSQM(schQCDATE2.Text)
        Dim v_schQCDATE1 As String = TIMS.Cdate3(schQCDATE1.Text)
        Dim v_schQCDATE2 As String = TIMS.Cdate3(schQCDATE2.Text)

        Dim s_SearchStr1 As String = "" 'Session("_SearchStr")
        s_SearchStr1 = "k=1"
        s_SearchStr1 &= $"&YEARS_S1={V_ddlYEARS_S1}&APPSTAGE_S1={V_ddlAPPSTAGE_S1}&TTQSLOCK_S1={V_ddlTTQSLOCK_S1}"
        s_SearchStr1 &= $"&QCDATE1={schQCDATE1.Text}&QCDATE2={schQCDATE2.Text}"
        Session("_SearchStr") = s_SearchStr1
    End Sub

    '20190103 若有先前查詢條件記錄，則將資料重新讀取到頁面中
    Sub USE_KeepSearch1()
        If Session("_SearchStr") IsNot Nothing Then
            Dim MYVALUE1 As String = ""
            Dim s_SearchStr1 As String = Session("_SearchStr")
            Session("_SearchStr") = Nothing

            MYVALUE1 = TIMS.GetMyValue(s_SearchStr1, "YEARS_S1")
            If MYVALUE1 <> "" Then Common.SetListItem(ddlYEARS_S1, MYVALUE1)
            MYVALUE1 = TIMS.GetMyValue(s_SearchStr1, "APPSTAGE_S1")
            If MYVALUE1 <> "" Then Common.SetListItem(ddlAPPSTAGE_S1, MYVALUE1)
            MYVALUE1 = TIMS.GetMyValue(s_SearchStr1, "TTQSLOCK_S1")
            If MYVALUE1 <> "" Then Common.SetListItem(ddlTTQSLOCK_S1, MYVALUE1)
            schQCDATE1.Text = TIMS.GetMyValue(s_SearchStr1, "QCDATE1")
            schQCDATE2.Text = TIMS.GetMyValue(s_SearchStr1, "QCDATE2")
            SSearch1()
        End If
    End Sub

    Sub Utl_Delete(ByVal s_OTQID As String)
        Dim uParms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"OTQID", s_OTQID}}
        Dim usql As String = "
UPDATE ORG_TTQSQUERY
SET ISDELETE='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()
WHERE OTQID=@OTQID
"
        DbAccess.ExecuteNonQuery(usql, objconn, uParms)

        Common.MessageBox(Me, "刪除成功")
    End Sub

    ''' <summary>儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_save_Click(sender As Object, e As EventArgs) Handles bt_save.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    Protected Sub Bt_cancle_Click(sender As Object, e As EventArgs) Handles bt_cancle.Click
        Utl_Clear()
        Utl_SHOW(0)
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                '序號
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim btnEDIT1 As LinkButton = e.Item.FindControl("btnEDIT1")
                'Dim btnDEL1 As LinkButton=e.Item.FindControl("btnDEL1")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OTQID", TIMS.CStr1(drv("OTQID")))
                btnEDIT1.CommandArgument = sCmdArg
                'btnDEL1.CommandArgument=sCmdArg
                'btnDEL1.Attributes("onclick")=TIMS.cst_confirm_delmsg2

                'Errmsg &= "權限不足 不可儲存，請連絡系統管理者!" & vbCrLf
                If sm.UserInfo.LID <> 0 Then
                    btnEDIT1.Visible = False
                    'btnDEL1.Visible=False
                End If
                btnEDIT1.Enabled = If(Convert.ToString(drv("ISDELETE")) = "Y", False, True)
                TIMS.Tooltip(btnEDIT1, If(Not btnEDIT1.Enabled, "已刪除資料", ""), True)
                'btnDEL1.Enabled=If(Convert.ToString(drv("ISDELETE"))="Y", False, True)
                'TIMS.Tooltip(btnDEL1, If(Not btnDEL1.Enabled,"已刪除資料",""), True)
        End Select

    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Utl_Clear()
        KeepSearch1()

        Dim sCmdArg As String = e.CommandArgument
        Select Case e.CommandName
            Case "edit"
                Utl_SHOW(1)
                Hid_OTQID.Value = TIMS.GetMyValue(sCmdArg, "OTQID")
                LoadData1(Hid_OTQID.Value)
            Case "del"
                Hid_OTQID.Value = TIMS.GetMyValue(sCmdArg, "OTQID")
                Utl_Delete(Hid_OTQID.Value)
                Call SSearch1()
        End Select
    End Sub

End Class
