Partial Class CR_01_007
    Inherits AuthBasePage

    '課程審查/一階審查/審查幕僚意見開關機制,產業人才投資方案／PLAN_STAFFOPINSWITCH
    Dim flag_ROC As Boolean = True
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    Private Sub CCreate1()
        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = ""

        panel_EDIT1.Visible = False

        '計畫年度
        Dim iSYears As Integer = 2023 '(起始年度)
        Dim iYearsLast As Integer = If((sm.UserInfo.Years + 1) > (Year(Now) + 1), sm.UserInfo.Years + 1, Year(Now) + 1) '(sm.UserInfo.Years + 1)
        Dim iSYears2 As Integer = (iSYears + 2)
        Dim iEYears As Integer = If(iYearsLast > iSYears2, iYearsLast, iSYears2)
        ddlYearlist_sch = TIMS.GetSyear(ddlYearlist_sch, iSYears, iEYears, True)
        Common.SetListItem(ddlYearlist_sch, sm.UserInfo.Years)

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        ddlAppStage_sch = TIMS.Get_APPSTAGE2(ddlAppStage_sch)

        ddlYEARS = TIMS.GetSyear(ddlYEARS, iSYears, iEYears, True)
        'Common.SetListItem(ddlYEARS, sm.UserInfo.Years)
        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)
        ddlAppStage = TIMS.Get_APPSTAGE2(ddlAppStage)

        Call TIMS.SUB_SET_HR_MI(TB_SOPENDATE_HR, TB_SOPENDATE_MM)
        Call TIMS.SUB_SET_HR_MI(TB_FOPENDATE_HR, TB_FOPENDATE_MM)
    End Sub

    Protected Sub Bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Call sSearch1()
    End Sub

    Private Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = "查無資料"

        Dim v_ddlYearlist_sch As String = TIMS.GetListValue(ddlYearlist_sch)
        Dim v_ddlAppStage_sch As String = TIMS.GetListValue(ddlAppStage_sch)

        Dim sParms As New Hashtable
        Dim sSql As String = ""
        sSql &= " SELECT a.PSSID ,a.YEARS" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,a.APPSTAGE,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,format(a.SOPENDATE,'yyyy/MM/dd') SOPENDATE" & vbCrLf
        sSql &= " ,format(a.FOPENDATE,'yyyy/MM/dd') FOPENDATE" & vbCrLf
        sSql &= " ,concat(format(a.SOPENDATE,'yyyy/MM/dd HH:mm'),'~',format(a.FOPENDATE,'yyyy/MM/dd HH:mm')) SFOPENDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.SOPENDATE) SOPENDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.FOPENDATE) FOPENDATE_ROC" & vbCrLf
        sSql &= " ,concat(dbo.FN_CDATE1C(a.SOPENDATE),'~',dbo.FN_CDATE1C(a.FOPENDATE)) SFOPENDATE_ROC" & vbCrLf
        sSql &= " FROM PLAN_STAFFOPINSWITCH a" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf
        If v_ddlYearlist_sch <> "" Then
            sParms.Add("YEARS", v_ddlYearlist_sch)
            sSql &= " AND a.YEARS=@YEARS" & vbCrLf
        End If
        If v_ddlAppStage_sch <> "" Then
            sParms.Add("APPSTAGE", v_ddlAppStage_sch)
            sSql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Protected Sub Bt_addnew_Click(sender As Object, e As EventArgs) Handles bt_addnew.Click
        Call CLEARDATA1()
        panel_sch.Visible = False
        panel_EDIT1.Visible = True
    End Sub

    Private Sub CLEARDATA1()
        ddlYEARS.Enabled = True
        ddlAppStage.Enabled = True
        TIMS.Tooltip(ddlYEARS, "")
        TIMS.Tooltip(ddlAppStage, "")

        Hid_PSSID.Value = ""
        Common.SetListItem(ddlYEARS, "")
        Common.SetListItem(ddlAppStage, "")
        TB_SOPENDATE.Text = ""
        TB_FOPENDATE.Text = ""

        '日期修正為今天 00:00
        Dim s_DATE1 As String = String.Concat(Now.ToString("yyyy/MM/dd"), " 00:00")
        TIMS.SET_DateHM(CDate(s_DATE1), TB_SOPENDATE_HR, TB_SOPENDATE_MM)
        '日期修正為今天 23:59
        Dim s_DATE2 As String = String.Concat(Now.ToString("yyyy/MM/dd"), " 23:59")
        TIMS.SET_DateHM(CDate(s_DATE2), TB_FOPENDATE_HR, TB_FOPENDATE_MM)
    End Sub

    Protected Sub BtnSAVEDATA1_Click(sender As Object, e As EventArgs) Handles BtnSAVEDATA1.Click
        Dim sErrMsg1 As String = CheckSaveData1()
        If sErrMsg1 <> "" Then
            Common.MessageBox(Me, sErrMsg1)
            Exit Sub
        End If

        Call SAVEDATA1()
    End Sub

    Private Function CheckSaveData1() As String
        Dim sErrMsg1 As String = ""
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_ddlAppStage As String = TIMS.GetListValue(ddlAppStage)

        TB_SOPENDATE.Text = TIMS.ClearSQM(TB_SOPENDATE.Text)
        TB_FOPENDATE.Text = TIMS.ClearSQM(TB_FOPENDATE.Text)
        Dim v_TB_SOPENDATE As String = If(flag_ROC, TIMS.Cdate7(TB_SOPENDATE.Text), TIMS.Cdate3(TB_SOPENDATE.Text))
        Dim v_TB_FOPENDATE As String = If(flag_ROC, TIMS.Cdate7(TB_FOPENDATE.Text), TIMS.Cdate3(TB_FOPENDATE.Text))
        Dim vSOPENDATE As String = TIMS.GET_DateHM(v_TB_SOPENDATE, TB_SOPENDATE_HR, TB_SOPENDATE_MM)
        Dim vFOPENDATE As String = TIMS.GET_DateHM(v_TB_FOPENDATE, TB_FOPENDATE_HR, TB_FOPENDATE_MM)

        If v_ddlYEARS = "" Then sErrMsg1 &= "年度 不可為空" & vbCrLf
        If v_ddlAppStage = "" Then sErrMsg1 &= "申請階段 不可為空" & vbCrLf
        If TB_SOPENDATE.Text = "" Then sErrMsg1 &= "「審查幕僚意見」開放增修 起始日期 不可為空" & vbCrLf
        If TB_FOPENDATE.Text = "" Then sErrMsg1 &= "「審查幕僚意見」開放增修 結束日期 不可為空" & vbCrLf
        'If sErrMsg1 <> "" Then Return sErrMsg1

        '(_i 正確值比對)
        If vSOPENDATE <> "" AndAlso vFOPENDATE <> "" Then
            '有勾選且有填上架日期的資料，再進一步檢核設定結果不得超過報名日期
            Dim iMinute3 As Long = DateDiff(DateInterval.Minute, CDate(vSOPENDATE), CDate(vFOPENDATE))
            If iMinute3 <= 0 Then sErrMsg1 &= String.Concat("「審查幕僚意見」開放增修 [起始日期][結束日期]起始時間大於結束順序有誤!", iMinute3) & vbCrLf
            If sErrMsg1 <> "" Then Return sErrMsg1

            Dim iDay3 As Long = DateDiff(DateInterval.Day, CDate(vSOPENDATE), CDate(vFOPENDATE))
            If iDay3 <= 3 Then sErrMsg1 &= $"「審查幕僚意見」開放增修 [起始日期][結束日期]時間長度有誤!{iDay3}{vbCrLf}(不足3天)"
            If iDay3 >= 180 Then sErrMsg1 &= $"「審查幕僚意見」開放增修 [起始日期][結束日期]時間長度有誤!{iDay3}{vbCrLf}(超過180天)"
            If sErrMsg1 <> "" Then Return sErrMsg1
        End If
        If sErrMsg1 <> "" Then Return sErrMsg1

        Hid_PSSID.Value = TIMS.ClearSQM(Hid_PSSID.Value)
        Dim sParms1 As New Hashtable From {{"YEARS", v_ddlYEARS}, {"APPSTAGE", v_ddlAppStage}}
        Dim sSql1 As String = "SELECT PSSID FROM PLAN_STAFFOPINSWITCH WHERE YEARS=@YEARS AND APPSTAGE=@APPSTAGE "
        If Hid_PSSID.Value <> "" Then
            sParms1.Add("PSSID", TIMS.VAL1(Hid_PSSID.Value))
            sSql1 &= " AND PSSID!=@PSSID"
        End If
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms1)
        If dr1 IsNot Nothing Then sErrMsg1 &= String.Concat("年度／申請階段 資料已經存在，請使用修改功能!", dr1("PSSID")) & vbCrLf
        If sErrMsg1 <> "" Then Return sErrMsg1

        Return sErrMsg1
    End Function

    Private Sub SAVEDATA1()
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_ddlAppStage As String = TIMS.GetListValue(ddlAppStage)

        TB_SOPENDATE.Text = TIMS.ClearSQM(TB_SOPENDATE.Text)
        TB_FOPENDATE.Text = TIMS.ClearSQM(TB_FOPENDATE.Text)
        Dim v_TB_SOPENDATE As String = If(flag_ROC, TIMS.Cdate7(TB_SOPENDATE.Text), TIMS.Cdate3(TB_SOPENDATE.Text))
        Dim v_TB_FOPENDATE As String = If(flag_ROC, TIMS.Cdate7(TB_FOPENDATE.Text), TIMS.Cdate3(TB_FOPENDATE.Text))
        Dim vSOPENDATE As String = TIMS.GET_DateHM(v_TB_SOPENDATE, TB_SOPENDATE_HR, TB_SOPENDATE_MM)
        Dim vFOPENDATE As String = TIMS.GET_DateHM(v_TB_FOPENDATE, TB_FOPENDATE_HR, TB_FOPENDATE_MM)

        Hid_PSSID.Value = TIMS.ClearSQM(Hid_PSSID.Value)
        Dim iRst As Integer = 0
        If Hid_PSSID.Value = "" Then
            Dim iPSSID As Integer = DbAccess.GetNewId(objconn, "PLAN_STAFFOPINSWITCH_PSSID_SEQ,PLAN_STAFFOPINSWITCH,PSSID")
            Dim iParms As New Hashtable
            iParms.Add("PSSID", iPSSID)
            iParms.Add("YEARS", v_ddlYEARS)
            iParms.Add("APPSTAGE", v_ddlAppStage)
            iParms.Add("SOPENDATE", vSOPENDATE)
            iParms.Add("FOPENDATE", vFOPENDATE)
            iParms.Add("CREATEACCT", sm.UserInfo.UserID)
            iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
            Dim i_Sql As String = ""
            i_Sql &= " INSERT INTO PLAN_STAFFOPINSWITCH(PSSID ,YEARS ,APPSTAGE ,SOPENDATE ,FOPENDATE,CREATEACCT ,CREATEDATE ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
            i_Sql &= " VALUES (@PSSID ,@YEARS ,@APPSTAGE ,@SOPENDATE ,@FOPENDATE,@CREATEACCT ,GETDATE(),@MODIFYACCT ,GETDATE())" & vbCrLf
            iRst = DbAccess.ExecuteNonQuery(i_Sql, objconn, iParms)
            If iRst = 0 Then
                Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
                Return
            End If
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
            'Return
        Else
            Dim iPSSID As Integer = TIMS.VAL1(Hid_PSSID.Value)
            Dim uParms As New Hashtable
            uParms.Add("SOPENDATE", vSOPENDATE)
            uParms.Add("FOPENDATE", vFOPENDATE)
            uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
            uParms.Add("PSSID", iPSSID)

            Dim UsSql As String = ""
            UsSql = " UPDATE PLAN_STAFFOPINSWITCH" & vbCrLf
            UsSql &= " SET SOPENDATE=@SOPENDATE ,FOPENDATE=@FOPENDATE" & vbCrLf
            UsSql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
            UsSql &= " WHERE PSSID=@PSSID" & vbCrLf
            iRst = DbAccess.ExecuteNonQuery(UsSql, objconn, uParms)
            If iRst = 0 Then
                Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
                Return
            End If
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
            'Return
        End If
        Call sSearch1()
        panel_sch.Visible = True 'False
        panel_EDIT1.Visible = False 'True
    End Sub

    Protected Sub BtnBack1_Click2(sender As Object, e As EventArgs) Handles BtnBack1.Click
        panel_sch.Visible = True 'False
        panel_EDIT1.Visible = False 'True
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing OrElse e.CommandName = "" OrElse e.CommandArgument = "" Then Return ' Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim iPSSID As Integer = TIMS.VAL1(TIMS.GetMyValue(sCmdArg, "PSSID"))

        Select Case e.CommandName
            Case "BTNUPDATE" '修改
                Call CLEARDATA1()
                Call SHOW_DETAIL1(iPSSID)

            Case "BTNDELETE" '刪除
                'Call CLEARDATA1()
                Call DELETEDATA1(iPSSID)
        End Select
    End Sub

    Private Sub DELETEDATA1(iPSSID As Integer)
        Dim iRst As Integer = 0
        Dim dParms As New Hashtable
        dParms.Add("PSSID", iPSSID)
        'dParms.Add("MODIFYACCT", sm.UserInfo.UserID)
        Dim dsSql As String = "DELETE PLAN_STAFFOPINSWITCH WHERE PSSID=@PSSID" & vbCrLf
        iRst = DbAccess.ExecuteNonQuery(dsSql, objconn, dParms)
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Return
        End If
        Common.MessageBox(Me, TIMS.cst_DELETEOKMsg1)
        'Return
        Call sSearch1()
        panel_sch.Visible = True 'False
        panel_EDIT1.Visible = False 'True
    End Sub

    Private Sub SHOW_DETAIL1(iPSSID As Integer)
        If iPSSID = 0 Then Return
        Hid_PSSID.Value = iPSSID

        Dim sParms1 As New Hashtable
        sParms1.Add("PSSID", iPSSID)
        Dim sSql As String = ""
        sSql &= " SELECT a.PSSID ,a.YEARS" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,a.APPSTAGE,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,format(a.SOPENDATE,'yyyy/MM/dd') SOPENDATE" & vbCrLf
        sSql &= " ,format(a.FOPENDATE,'yyyy/MM/dd') FOPENDATE" & vbCrLf
        sSql &= " ,format(a.SOPENDATE,'yyyy/MM/dd HH:mm') SOPENDATE_HM" & vbCrLf
        sSql &= " ,format(a.FOPENDATE,'yyyy/MM/dd HH:mm') FOPENDATE_HM" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.SOPENDATE) SOPENDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.FOPENDATE) FOPENDATE_ROC" & vbCrLf
        sSql &= " FROM PLAN_STAFFOPINSWITCH a" & vbCrLf
        sSql &= " WHERE a.PSSID=@PSSID " & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms1)
        If dr1 Is Nothing Then Return

        Common.SetListItem(ddlYEARS, dr1("YEARS"))
        Common.SetListItem(ddlAppStage, dr1("APPSTAGE"))

        TB_SOPENDATE.Text = If(flag_ROC, TIMS.ClearSQM(dr1("SOPENDATE_ROC")), TIMS.Cdate3(dr1("SOPENDATE")))
        TIMS.SET_DateHM(CDate(dr1("SOPENDATE_HM")), TB_SOPENDATE_HR, TB_SOPENDATE_MM)
        TB_FOPENDATE.Text = If(flag_ROC, TIMS.ClearSQM(dr1("FOPENDATE_ROC")), TIMS.Cdate3(dr1("FOPENDATE")))
        TIMS.SET_DateHM(CDate(dr1("FOPENDATE_HM")), TB_FOPENDATE_HR, TB_FOPENDATE_MM)

        ddlYEARS.Enabled = False
        ddlAppStage.Enabled = False
        TIMS.Tooltip(ddlYEARS, "年度不可修改", True)
        TIMS.Tooltip(ddlAppStage, "申請階段不可修改", True)

        panel_sch.Visible = False
        panel_EDIT1.Visible = True
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim BTNUPDATE As LinkButton = e.Item.FindControl("BTNUPDATE") '修改
                Dim BTNDELETE As LinkButton = e.Item.FindControl("BTNDELETE") '刪除
                Dim iDGSeqNo As Integer = TIMS.Get_DGSeqNo(sender, e) ' e.Item.ItemIndex + 1 + DG_Org.PageSize * DG_Org.CurrentPageIndex
                e.Item.Cells(0).Text = iDGSeqNo ' e.Item.ItemIndex + 1 + DG_Org.PageSize * DG_Org.CurrentPageIndex
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PSSID", drv("PSSID"))
                BTNUPDATE.CommandArgument = sCmdArg '修改
                BTNDELETE.CommandArgument = sCmdArg '刪除
                Dim sDELMSG1 As String = String.Concat("return confirm('", "您確定要刪除第", iDGSeqNo, "筆資料嗎?", "');")
                BTNDELETE.Attributes.Add("onclick", sDELMSG1)
        End Select
    End Sub

End Class