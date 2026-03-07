Public Class CO_01_006
    Inherits AuthBasePage 'System.Web.UI.Page

    'Const cst_months_01 As String = "01"
    'Const cst_months_07 As String = "07"
    '[ORG_TTQSLOCK]
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            Call sCreate1() '頁面初始化

            '20190103 若有先前查詢條件記錄，則將資料重新讀取到頁面中
            If Session("_SearchStr") IsNot Nothing Then
                Dim s_SearchStr1 As String = Session("_SearchStr")
                schCCDATE1.Text = TIMS.GetMyValue(s_SearchStr1, "schCCDATE1")
                schCCDATE2.Text = TIMS.GetMyValue(s_SearchStr1, "schCCDATE2")
                schCSDATE1.Text = TIMS.GetMyValue(s_SearchStr1, "schCSDATE1")
                schCSDATE2.Text = TIMS.GetMyValue(s_SearchStr1, "schCSDATE2")
                schCFDATE1.Text = TIMS.GetMyValue(s_SearchStr1, "schCFDATE1")
                schCFDATE2.Text = TIMS.GetMyValue(s_SearchStr1, "schCFDATE2")
                Session("_SearchStr") = Nothing

                sSearch1()

            End If
        End If
    End Sub

    '資料查詢
    Sub sSearch1()
        schCCDATE1.Text = TIMS.ClearSQM(schCCDATE1.Text)
        schCCDATE2.Text = TIMS.ClearSQM(schCCDATE2.Text)
        schCSDATE1.Text = TIMS.ClearSQM(schCSDATE1.Text)
        schCSDATE2.Text = TIMS.ClearSQM(schCSDATE2.Text)
        schCFDATE1.Text = TIMS.ClearSQM(schCFDATE1.Text)
        schCFDATE2.Text = TIMS.ClearSQM(schCFDATE2.Text)

        Dim v_schCCDATE1 As String = TIMS.Cdate3(schCCDATE1.Text)
        Dim v_schCCDATE2 As String = TIMS.Cdate3(schCCDATE2.Text)
        Dim v_schCSDATE1 As String = TIMS.Cdate3(schCSDATE1.Text)
        Dim v_schCSDATE2 As String = TIMS.Cdate3(schCSDATE2.Text)
        Dim v_schCFDATE1 As String = TIMS.Cdate3(schCFDATE1.Text)
        Dim v_schCFDATE2 As String = TIMS.Cdate3(schCFDATE2.Text)

        Dim sql As String = ""
        sql &= " SELECT a.OTLID" & vbCrLf '  /*PK*/
        sql &= " ,format(a.CCDATE,'yyyy/MM/dd') CCDATE" & vbCrLf
        sql &= " ,format(a.CSDATE,'yyyy/MM/dd HH:mm') CSDATE" & vbCrLf
        sql &= " ,format(a.CFDATE,'yyyy/MM/dd HH:mm') CFDATE" & vbCrLf
        sql &= " ,a.EXPLAIN" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,a.MONTHS" & vbCrLf
        sql &= " ,CASE a.MONTHS WHEN '01' THEN '1月' WHEN '07' THEN '7月' END MONTHS_N" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.ISDELETE" & vbCrLf
        sql &= " FROM ORG_TTQSLOCK a" & vbCrLf
        'sql &= " WHERE a.ISDELETE IS NULL" & vbCrLf
        '包含已刪除
        sql &= If(CHK_ISDELETE.Checked, " WHERE 1=1", " WHERE a.ISDELETE IS NULL")
        If v_schCCDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.CCDATE, 111) >= @schCCDATE1 " & vbCrLf
        If v_schCCDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.CCDATE, 111) <= @schCCDATE2 " & vbCrLf
        If v_schCSDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.CSDATE, 111) >= @schCSDATE1 " & vbCrLf
        If v_schCSDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.CSDATE, 111) <= @schCSDATE2 " & vbCrLf
        If v_schCFDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.CFDATE, 111) >= @schCFDATE1 " & vbCrLf
        If v_schCFDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.CFDATE, 111) <= @schCFDATE2 " & vbCrLf
        sql &= " ORDER BY a.CSDATE DESC " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If v_schCCDATE1 <> "" Then parms.Add("schCCDATE1", v_schCCDATE1)
        If v_schCCDATE2 <> "" Then parms.Add("schCCDATE2", v_schCCDATE2)
        If v_schCSDATE1 <> "" Then parms.Add("schCSDATE1", v_schCSDATE1)
        If v_schCSDATE2 <> "" Then parms.Add("schCSDATE2", v_schCSDATE2)
        If v_schCFDATE1 <> "" Then parms.Add("schCFDATE1", v_schCFDATE1)
        If v_schCFDATE2 <> "" Then parms.Add("schCFDATE2", v_schCFDATE2)

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        lab_msg1.Text = "查無資料"
        tb_Sch.Visible = False

        If TIMS.dtNODATA(dt) Then Return

        lab_msg1.Text = ""
        tb_Sch.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '資料讀取
    Private Sub LoadData1(ByVal iOTLID As Integer)
        Dim sql As String = ""
        sql &= " SELECT a.OTLID" & vbCrLf '  /*PK*/
        sql &= " ,format(a.CCDATE,'yyyy/MM/dd') CCDATE" & vbCrLf

        sql &= " ,format(a.CSDATE,'yyyy/MM/dd') CSDATE" & vbCrLf
        sql &= " ,format(a.CSDATE,'HH') CSDATEHH" & vbCrLf
        sql &= " ,format(a.CSDATE,'mm') CSDATEMM" & vbCrLf

        sql &= " ,format(a.CFDATE,'yyyy/MM/dd') CFDATE" & vbCrLf
        sql &= " ,format(a.CFDATE,'HH') CFDATEHH" & vbCrLf
        sql &= " ,format(a.CFDATE,'mm') CFDATEMM" & vbCrLf

        sql &= " ,a.EXPLAIN" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,a.MONTHS" & vbCrLf
        sql &= " ,CASE a.MONTHS WHEN '01' THEN '1月' WHEN '07' THEN '7月' END MONTHS_N" & vbCrLf

        'sql &= " ,dbo.FN_CYEAR2(a.YEARS1) YEARS1_ROC" & vbCrLf
        sql &= " ,a.YEARS1" & vbCrLf
        sql &= " ,a.HALFYEAR1" & vbCrLf
        sql &= " ,a.YEARS2" & vbCrLf
        sql &= " ,a.HALFYEAR2" & vbCrLf
        'sql &= " ,a.CREATEACCT" & vbCrLf
        'sql &= " ,a.CREATEDATE" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.ISDELETE" & vbCrLf
        sql &= " FROM ORG_TTQSLOCK a" & vbCrLf
        sql &= " WHERE a.OTLID=@OTLID" & vbCrLf
        'sql &= " AND a.ISDELETE IS NULL" & vbCrLf

        Dim parms As New Hashtable From {{"OTLID", iOTLID}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim dr As DataRow = dt.Rows(0)
        txCCDATE.Text = TIMS.Cdate3(dr("CCDATE"))
        txCSDATE.Text = TIMS.Cdate3(dr("CSDATE"))
        Common.SetListItem(ddlCSDATE_HH, Convert.ToString(dr("CSDATEHH")))
        Common.SetListItem(ddlCSDATE_MM, Convert.ToString(dr("CSDATEMM")))

        txCFDATE.Text = TIMS.Cdate3(dr("CFDATE"))
        Common.SetListItem(ddlCFDATE_HH, Convert.ToString(dr("CFDATEHH")))
        Common.SetListItem(ddlCFDATE_MM, Convert.ToString(dr("CFDATEMM")))
        txEXPLAIN.Text = Convert.ToString(dr("EXPLAIN"))

        Common.SetListItem(ddlYEARS, Convert.ToString(dr("YEARS")))
        Common.SetListItem(rblMONTHS, Convert.ToString(dr("MONTHS")))

        Common.SetListItem(ddlYEARS1, Convert.ToString(dr("YEARS1")))
        Common.SetListItem(ddlHALFYEAR1, Convert.ToString(dr("HALFYEAR1")))
        Common.SetListItem(ddlYEARS2, Convert.ToString(dr("YEARS2")))
        Common.SetListItem(ddlHALFYEAR2, Convert.ToString(dr("HALFYEAR2")))
        '包含已刪除
        LabISDELETE.Visible = (Convert.ToString(dr("ISDELETE")) = "Y")
        '已刪除資料
        bt_save.Enabled = If(Convert.ToString(dr("ISDELETE")) = "Y", False, True)
        TIMS.Tooltip(bt_save, If(Not bt_save.Enabled, "已刪除資料", ""), True)
    End Sub

    ''' <summary> 月份設計 01/07 </summary>
    ''' <param name="obj"></param>
    ''' <param name="tConn"></param>
    ''' <returns></returns>
    Function Get_MONTHS2(ByVal obj As ListControl, ByVal tConn As SqlConnection) As ListControl
        Dim sql As String = ""
        sql &= " SELECT '01' MONTHS, N'1月' MONTHS_N UNION SELECT '07' MONTHS, N'7月' MONTHS_N" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, tConn)
        With obj
            .DataSource = dt
            .DataTextField = "MONTHS_N"
            .DataValueField = "MONTHS"
            .DataBind()
            .AppendDataBoundItems = True
        End With
        Return obj
    End Function

    '頁面初始化
    Sub sCreate1()
        lab_msg1.Text = ""
        tb_Sch.Visible = False
        lab_msg2.Text = ""

        ddlYEARS = TIMS.Get_ROCYEARS2(ddlYEARS, objconn)
        rblMONTHS = Get_MONTHS2(rblMONTHS, objconn)

        ddlYEARS1 = TIMS.Get_ROCYEARS2(ddlYEARS1, objconn)
        ddlHALFYEAR1 = TIMS.Get_ddlHALFYEAR(ddlHALFYEAR1)

        ddlYEARS2 = TIMS.Get_ROCYEARS2(ddlYEARS2, objconn)
        ddlHALFYEAR2 = TIMS.Get_ddlHALFYEAR(ddlHALFYEAR2)

        Dim dr1 As DataRow = TIMS.Get_drYMVALUE1(objconn)
        Common.SetListItem(ddlYEARS, dr1("YEARS"))
        Common.SetListItem(rblMONTHS, dr1("MONTHS"))

        Call TIMS.SUB_SET_HR_MI(ddlCSDATE_HH, ddlCSDATE_MM)
        Call TIMS.SUB_SET_HR_MI(ddlCFDATE_HH, ddlCFDATE_MM)

        Utl_SHOW(0)
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
        txCCDATE.Text = TIMS.Cdate3(Now.Date)
        txCSDATE.Text = ""
        txCFDATE.Text = ""
        'ddlCFDATE_hh1.SelectedIndex = -1
        'ddlCFDATE_mm1.SelectedIndex = -1
        Call TIMS.SUB_SET_HR_MI(ddlCSDATE_HH, ddlCSDATE_MM)
        Call TIMS.SUB_SET_HR_MI(ddlCFDATE_HH, ddlCFDATE_MM)
        txEXPLAIN.Text = ""
        Hid_OTLID.Value = ""
        Dim dr1 As DataRow = TIMS.Get_drYMVALUE1(objconn)
        Common.SetListItem(ddlYEARS, dr1("YEARS"))
        Common.SetListItem(rblMONTHS, dr1("MONTHS"))
        '包含已刪除
        LabISDELETE.Visible = False
    End Sub

    ''' <summary>
    ''' 查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Utl_SHOW(0)
        sSearch1()
    End Sub

    ''' <summary>
    ''' 新增
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_add_Click(sender As Object, e As EventArgs) Handles bt_add.Click
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
        Dim rst As Boolean = True
        Errmsg = ""

        txCCDATE.Text = TIMS.ClearSQM(txCCDATE.Text)
        txCSDATE.Text = TIMS.ClearSQM(txCSDATE.Text)
        txCFDATE.Text = TIMS.ClearSQM(txCFDATE.Text)
        txEXPLAIN.Text = TIMS.ClearSQM(txEXPLAIN.Text)

        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_rblMONTHS As String = TIMS.GetListValue(rblMONTHS)
        Dim v_ddlYEARS1 As String = TIMS.GetListValue(ddlYEARS1)
        Dim v_ddlHALFYEAR1 As String = TIMS.GetListValue(ddlHALFYEAR1)
        Dim v_ddlYEARS2 As String = TIMS.GetListValue(ddlYEARS2)
        Dim v_ddlHALFYEAR2 As String = TIMS.GetListValue(ddlHALFYEAR2)

        If txCCDATE.Text = "" Then Errmsg &= "設定日期 不可為空" & vbCrLf
        If txCSDATE.Text = "" Then Errmsg &= "控制起始日期 不可為空" & vbCrLf
        If txCFDATE.Text = "" Then Errmsg &= "控制結束日期 不可為空" & vbCrLf
        If txEXPLAIN.Text = "" Then Errmsg &= "控制說明 不可為空" & vbCrLf
        If v_ddlYEARS = "" Then Errmsg &= "請選擇 截止年度" & vbCrLf
        If v_rblMONTHS = "" Then Errmsg &= "請選擇 截止月份" & vbCrLf
        If v_ddlYEARS1 = "" Then Errmsg &= "請選擇 審查計分區間起始-年" & vbCrLf
        If v_ddlHALFYEAR1 = "" Then Errmsg &= "請選擇 審查計分區間起始-半年" & vbCrLf
        If v_ddlYEARS2 = "" Then Errmsg &= "請選擇 審查計分區間迄止-年" & vbCrLf
        If v_ddlHALFYEAR2 = "" Then Errmsg &= "請選擇 審查計分區間迄止-半年" & vbCrLf

        If Errmsg <> "" Then Return False


        Dim s_CSDATE As String = TIMS.GET_DateHM(txCSDATE, ddlCSDATE_HH, ddlCSDATE_MM)
        Dim s_CFDATE As String = TIMS.GET_DateHM(txCFDATE, ddlCFDATE_HH, ddlCFDATE_MM)

        If DateDiff(DateInterval.Minute, CDate(s_CSDATE), CDate(s_CFDATE)) = 0 Then Errmsg &= "控制起始日期與控制結束日期 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(s_CSDATE), CDate(s_CFDATE)) < 0 Then Errmsg &= "控制起始日期與控制結束日期 順序異常!!" & vbCrLf
        If sm.UserInfo.LID <> 0 Then
            Errmsg &= "權限不足 不可儲存，請連絡系統管理者!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        Dim parms As New Hashtable
        Dim sql As String = ""
        If Hid_OTLID.Value <> "" Then
            'UPDATE 'ORG_TTQSLOCK
            parms.Clear()
            parms.Add("OTLID", Val(Hid_OTLID.Value))
            parms.Add("YEARS", v_ddlYEARS)
            parms.Add("MONTHS", v_rblMONTHS)
            sql = " SELECT YEARS,MONTHS FROM ORG_TTQSLOCK WHERE OTLID!=@OTLID AND YEARS=@YEARS AND MONTHS=@MONTHS" & vbCrLf
        Else
            'INSERT
            parms.Clear()
            parms.Add("YEARS", v_ddlYEARS)
            parms.Add("MONTHS", v_rblMONTHS)
            sql = " SELECT YEARS,MONTHS FROM ORG_TTQSLOCK WHERE YEARS=@YEARS AND MONTHS=@MONTHS" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then
            Errmsg &= "截止年度／月份 已存在，請使用修改功能!!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        If Val(v_ddlYEARS1) > Val(v_ddlYEARS2) Then Errmsg &= "審查計分區間 年度 順序異常!!" & vbCrLf

        If Val(v_ddlYEARS2) > Val(v_ddlYEARS) Then
            Errmsg &= "審查計分區間 與 截止年度 年度 順序異常!!" & vbCrLf
        End If

        If Val(v_ddlYEARS2) = Val(v_ddlYEARS) Then
            If Val(v_ddlHALFYEAR1) < Val(v_ddlHALFYEAR2) Then
                Errmsg &= "審查計分區間 與 截止年度 半年順序異常!! (同年)" & vbCrLf
            End If
            If Val(v_ddlHALFYEAR1) = Val(v_ddlHALFYEAR2) Then
                Errmsg &= "審查計分區間 與 截止年度 半年 順序異常!! (同半年)" & vbCrLf
            End If
        End If
        If Val(v_ddlYEARS1) = Val(v_ddlYEARS2) Then
            If Val(v_ddlHALFYEAR1) > Val(v_ddlHALFYEAR2) Then
                Errmsg &= "審查計分區間 半年 順序異常!!" & vbCrLf
            End If
            If Val(v_ddlHALFYEAR1) = Val(v_ddlHALFYEAR2) Then
                Errmsg &= "審查計分區間 半年 順序異常!! (同半年)" & vbCrLf
            End If
        End If

        If Val(v_ddlYEARS1) < Val(v_ddlYEARS2) AndAlso Val(v_ddlHALFYEAR1) < Val(v_ddlHALFYEAR2) Then
            Errmsg &= "審查計分區間 半年 順序異常!!" & vbCrLf
        End If
        If Val(v_ddlYEARS1) < Val(v_ddlYEARS2) AndAlso Val(v_ddlYEARS1) + 1 <> Val(v_ddlYEARS2) Then
            Errmsg &= "審查計分區間 年度 範圍異常!!" & vbCrLf
        End If
        If Val(v_ddlYEARS2) < Val(v_ddlYEARS) AndAlso Val(v_ddlYEARS2) + 1 <> Val(v_ddlYEARS) Then
            Errmsg &= "審查計分區間 與 截止年度 年度範圍異常!!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        If Val(v_ddlHALFYEAR1) = Val(v_ddlHALFYEAR2) Then
            Errmsg &= $"審查計分區間 半年 異常!! (相同),{v_ddlHALFYEAR2}" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存(part-1)
    Sub SaveData1()
        Dim flagSaveOK1 As Boolean = True

        Dim i_sql As String = ""
        i_sql &= " INSERT INTO ORG_TTQSLOCK( OTLID,CCDATE,CSDATE,CFDATE,EXPLAIN,YEARS,MONTHS" & vbCrLf
        i_sql &= " ,YEARS1,HALFYEAR1,YEARS2,HALFYEAR2" & vbCrLf
        i_sql &= " ,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES ( @OTLID,@CCDATE,@CSDATE,@CFDATE,@EXPLAIN,@YEARS,@MONTHS" & vbCrLf
        i_sql &= " ,@YEARS1,@HALFYEAR1,@YEARS2,@HALFYEAR2" & vbCrLf
        i_sql &= " ,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf

        Dim u_sql As String = ""
        u_sql &= " UPDATE ORG_TTQSLOCK" & vbCrLf
        u_sql &= " SET OTLID=@OTLID" & vbCrLf
        u_sql &= " ,CCDATE=@CCDATE" & vbCrLf
        u_sql &= " ,CSDATE=@CSDATE" & vbCrLf
        u_sql &= " ,CFDATE=@CFDATE" & vbCrLf
        u_sql &= " ,EXPLAIN=@EXPLAIN" & vbCrLf
        u_sql &= " ,YEARS=@YEARS" & vbCrLf
        u_sql &= " ,MONTHS=@MONTHS" & vbCrLf
        u_sql &= " ,YEARS1=@YEARS1,HALFYEAR1=@HALFYEAR1" & vbCrLf
        u_sql &= " ,YEARS2=@YEARS2,HALFYEAR2=@HALFYEAR2" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE OTLID=@OTLID" & vbCrLf

        txCCDATE.Text = TIMS.Cdate3(txCCDATE.Text)
        Dim s_CSDATE As String = TIMS.GET_DateHM(txCSDATE, ddlCSDATE_HH, ddlCSDATE_MM)
        Dim s_CFDATE As String = TIMS.GET_DateHM(txCFDATE, ddlCFDATE_HH, ddlCFDATE_MM)
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_rblMONTHS As String = TIMS.GetListValue(rblMONTHS)
        Dim v_ddlYEARS1 As String = TIMS.GetListValue(ddlYEARS1)
        Dim v_ddlHALFYEAR1 As String = TIMS.GetListValue(ddlHALFYEAR1)
        Dim v_ddlYEARS2 As String = TIMS.GetListValue(ddlYEARS2)
        Dim v_ddlHALFYEAR2 As String = TIMS.GetListValue(ddlHALFYEAR2)

        Dim iOTLID As Integer = 0
        Dim iRst As Integer = 0
        If Hid_OTLID.Value = "" Then
            '新增
            iOTLID = DbAccess.GetNewId(objconn, "ORG_TTQSLOCK_OTLID_SEQ,ORG_TTQSLOCK,OTLID")
            Dim parms As New Hashtable
            parms.Clear()
            parms.Add("OTLID", iOTLID)
            parms.Add("CCDATE", txCCDATE.Text)
            parms.Add("CSDATE", s_CSDATE)
            parms.Add("CFDATE", s_CFDATE)
            parms.Add("EXPLAIN", txEXPLAIN.Text)
            parms.Add("YEARS", If(v_ddlYEARS <> "", v_ddlYEARS, Convert.DBNull))
            parms.Add("MONTHS", If(v_rblMONTHS <> "", v_rblMONTHS, Convert.DBNull))
            parms.Add("YEARS1", If(v_ddlYEARS1 <> "", v_ddlYEARS1, Convert.DBNull))
            parms.Add("HALFYEAR1", If(v_ddlHALFYEAR1 <> "", v_ddlHALFYEAR1, Convert.DBNull))
            parms.Add("YEARS2", If(v_ddlYEARS2 <> "", v_ddlYEARS2, Convert.DBNull))
            parms.Add("HALFYEAR2", If(v_ddlHALFYEAR2 <> "", v_ddlHALFYEAR2, Convert.DBNull))

            parms.Add("CREATEACCT", sm.UserInfo.UserID)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            iRst += DbAccess.ExecuteNonQuery(i_sql, objconn, parms)
            Hid_OTLID.Value = iOTLID
        Else
            iOTLID = Val(Hid_OTLID.Value)
            '修改
            Dim parms As New Hashtable
            parms.Clear()
            parms.Add("CCDATE", txCCDATE.Text)
            parms.Add("CSDATE", s_CSDATE)
            parms.Add("CFDATE", s_CFDATE)
            parms.Add("EXPLAIN", txEXPLAIN.Text)
            parms.Add("YEARS", If(v_ddlYEARS <> "", v_ddlYEARS, Convert.DBNull))
            parms.Add("MONTHS", If(v_rblMONTHS <> "", v_rblMONTHS, Convert.DBNull))
            parms.Add("YEARS1", If(v_ddlYEARS1 <> "", v_ddlYEARS1, Convert.DBNull))
            parms.Add("HALFYEAR1", If(v_ddlHALFYEAR1 <> "", v_ddlHALFYEAR1, Convert.DBNull))
            parms.Add("YEARS2", If(v_ddlYEARS2 <> "", v_ddlYEARS2, Convert.DBNull))
            parms.Add("HALFYEAR2", If(v_ddlHALFYEAR2 <> "", v_ddlHALFYEAR2, Convert.DBNull))

            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            parms.Add("OTLID", iOTLID)
            iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, parms)
        End If

        If flagSaveOK1 Then
            '儲存成功
            Dim url1 As String = "CO_01_006.aspx?id=" & TIMS.Get_MRqID(Me)
            Common.MessageBox(Me, "儲存成功!")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

    End Sub


    Sub KeepSearch1()
        Dim s_SearchStr1 As String = "" 'Session("_SearchStr")
        s_SearchStr1 = "k=1"
        s_SearchStr1 &= "&schCCDATE1=" & schCCDATE1.Text
        s_SearchStr1 &= "&schCCDATE2=" & schCCDATE2.Text
        s_SearchStr1 &= "&schCSDATE1=" & schCSDATE1.Text
        s_SearchStr1 &= "&schCSDATE2=" & schCSDATE2.Text
        s_SearchStr1 &= "&schCFDATE1=" & schCFDATE1.Text
        s_SearchStr1 &= "&schCFDATE2=" & schCFDATE2.Text
        Session("_SearchStr") = s_SearchStr1
    End Sub

    Sub Utl_Delete(ByVal s_OTLID As String)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE ORG_TTQSLOCK " & vbCrLf
        sql &= " SET ISDELETE = 'Y' " & vbCrLf
        sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " ,MODIFYACCT = @MODIFYACCT " & vbCrLf
        sql &= " WHERE OTLID = @OTLID " & vbCrLf
        Dim parms As Hashtable = New Hashtable()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("OTLID", s_OTLID)
        DbAccess.ExecuteNonQuery(sql, objconn, parms)

        Common.MessageBox(Me, "刪除成功")
    End Sub


    ''' <summary>儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_save_Click(sender As Object, e As EventArgs) Handles bt_save.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    Protected Sub bt_cancle_Click(sender As Object, e As EventArgs) Handles bt_cancle.Click
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
                Dim btnDEL1 As LinkButton = e.Item.FindControl("btnDEL1")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OTLID", TIMS.CStr1(drv("OTLID")))
                btnEDIT1.CommandArgument = sCmdArg
                btnDEL1.CommandArgument = sCmdArg
                btnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg2

                If sm.UserInfo.LID <> 0 Then
                    'Errmsg &= "權限不足 不可儲存，請連絡系統管理者!" & vbCrLf
                    btnEDIT1.Visible = False
                    btnDEL1.Visible = False
                End If

                btnDEL1.Enabled = If(Convert.ToString(drv("ISDELETE")) = "Y", False, True)
                TIMS.Tooltip(btnDEL1, If(Not btnDEL1.Enabled, "已刪除資料", ""), True)
                'btnEDIT1.Enabled = If(Convert.ToString(drv("ISDELETE")) = "Y", False, True)
                'TIMS.Tooltip(btnEDIT1, If(Not btnEDIT1.Enabled, "已刪除資料", ""), True)
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
                Hid_OTLID.Value = TIMS.GetMyValue(sCmdArg, "OTLID")
                LoadData1(Hid_OTLID.Value)
            Case "del"
                Hid_OTLID.Value = TIMS.GetMyValue(sCmdArg, "OTLID")
                Utl_Delete(Hid_OTLID.Value)
                Call sSearch1()

        End Select
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
