Partial Class TC_01_021
    Inherits AuthBasePage

    Const cst_title_申請階段種類_1 As String = "申請階段種類"
    '開放受理之申請階段／PLAN_APPSTAGE
    Dim flag_ROC As Boolean = True
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        'Call sUtl_PageInit1()
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

        Call TIMS.SUB_SET_HR_MI(TB_SACCEPTDATE_HR, TB_SACCEPTDATE_MM)
        Call TIMS.SUB_SET_HR_MI(TB_FACCEPTDATE_HR, TB_FACCEPTDATE_MM)
    End Sub

    Protected Sub Bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Call SSearch1()
    End Sub

    Private Sub SSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = "查無資料"

        Dim v_ddlYearlist_sch As String = TIMS.GetListValue(ddlYearlist_sch)
        Dim v_ddlAppStage_sch As String = TIMS.GetListValue(ddlAppStage_sch)
        Dim v_rblPTYPE1_sch As String = TIMS.GetListValue(rblPTYPE1_sch)
        If v_rblPTYPE1_sch = "" Then
            Common.MessageBox(Me, String.Concat("請先選擇，", cst_title_申請階段種類_1, "!"))
            Return
        End If
        If v_rblPTYPE1_sch = "" Then Return

        Dim sParms As New Hashtable From {{"PTYPE1", v_rblPTYPE1_sch}}
        Dim sSql As String = ""
        sSql &= " SELECT a.PAID" & vbCrLf
        sSql &= " ,a.YEARS ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,a.APPSTAGE,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,a.PTYPE1,CASE WHEN a.PTYPE1='01' THEN '申請' WHEN a.PTYPE1='02' THEN '申複' END PTYPE1_N" & vbCrLf
        sSql &= " ,format(a.SACCEPTDATE,'yyyy/MM/dd') SACCEPTDATE" & vbCrLf
        sSql &= " ,format(a.FACCEPTDATE,'yyyy/MM/dd') FACCEPTDATE" & vbCrLf
        sSql &= " ,concat(format(a.SACCEPTDATE,'yyyy/MM/dd HH:mm'),'~',format(a.FACCEPTDATE,'yyyy/MM/dd HH:mm')) SFACCEPTDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.SACCEPTDATE) SACCEPTDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.FACCEPTDATE) FACCEPTDATE_ROC" & vbCrLf
        sSql &= " ,concat(dbo.FN_CDATE1C(a.SACCEPTDATE),'~',dbo.FN_CDATE1C(a.FACCEPTDATE)) SFACCEPTDATE_ROC" & vbCrLf
        'sSql &= " ,format(a.FACCEPTDATE,'yyyy/MM/dd') FACCEPTDATE" & vbCrLf
        'sSql &= " ,a.CREATEACCT ,a.CREATEDATE ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sSql &= " FROM PLAN_APPSTAGE a" & vbCrLf
        sSql &= " WHERE a.PTYPE1=@PTYPE1" & vbCrLf
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

        Dim v_ddlYearlist_sch As String = TIMS.GetListValue(ddlYearlist_sch)
        Dim v_ddlAppStage_sch As String = TIMS.GetListValue(ddlAppStage_sch)
        Dim v_rblPTYPE1_sch As String = TIMS.GetListValue(rblPTYPE1_sch)
        If v_rblPTYPE1_sch = "" Then
            Common.MessageBox(Me, String.Concat("請先選擇，", cst_title_申請階段種類_1, "!"))
            Return
        End If
        If v_rblPTYPE1_sch = "" Then Return

        Common.SetListItem(ddlYEARS, v_ddlYearlist_sch)
        Common.SetListItem(ddlAppStage, v_ddlAppStage_sch)
        Common.SetListItem(rblPTYPE1, v_rblPTYPE1_sch)
        'rblPTYPE1.Enabled = False
        'TIMS.Tooltip(rblPTYPE1, "申請階段種類不可修改!", True)

        panel_sch.Visible = False
        panel_EDIT1.Visible = True
    End Sub

    Private Sub CLEARDATA1()
        ddlYEARS.Enabled = True
        ddlAppStage.Enabled = True
        TIMS.Tooltip(ddlYEARS, "")
        TIMS.Tooltip(ddlAppStage, "")

        Hid_PAID.Value = ""
        Common.SetListItem(ddlYEARS, "")
        Common.SetListItem(ddlAppStage, "")

        rblPTYPE1.Enabled = True
        TIMS.Tooltip(rblPTYPE1, "")
        Common.SetListItem(rblPTYPE1, "")

        TB_SACCEPTDATE.Text = ""
        TB_FACCEPTDATE.Text = ""
        '日期修正為今天 00:00
        Dim s_DATE1 As String = String.Concat(Now.ToString("yyyy/MM/dd"), " 00:00")
        TIMS.SET_DateHM(CDate(s_DATE1), TB_SACCEPTDATE_HR, TB_SACCEPTDATE_MM)
        '日期修正為今天 23:59
        Dim s_DATE2 As String = String.Concat(Now.ToString("yyyy/MM/dd"), " 23:59")
        TIMS.SET_DateHM(CDate(s_DATE2), TB_FACCEPTDATE_HR, TB_FACCEPTDATE_MM)
        '勾選後,關閉(更新班級清單)按鈕
        CBK1_CLOSE_RETRIEVE.Checked = False
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
        Dim v_rblPTYPE1 As String = TIMS.GetListValue(rblPTYPE1)

        'TB_SACCEPTDATE.Text = TIMS.cdate3(TIMS.ClearSQM(TB_SACCEPTDATE.Text))
        'TB_FACCEPTDATE.Text = TIMS.cdate3(TIMS.ClearSQM(TB_FACCEPTDATE.Text))
        TB_SACCEPTDATE.Text = TIMS.ClearSQM(TB_SACCEPTDATE.Text)
        TB_FACCEPTDATE.Text = TIMS.ClearSQM(TB_FACCEPTDATE.Text)
        Dim v_TB_SACCEPTDATE As String = If(flag_ROC, TIMS.Cdate7(TB_SACCEPTDATE.Text), TIMS.Cdate3(TB_SACCEPTDATE.Text))
        Dim v_TB_FACCEPTDATE As String = If(flag_ROC, TIMS.Cdate7(TB_FACCEPTDATE.Text), TIMS.Cdate3(TB_FACCEPTDATE.Text))
        Dim vSACCEPTDATE As String = TIMS.GET_DateHM(v_TB_SACCEPTDATE, TB_SACCEPTDATE_HR, TB_SACCEPTDATE_MM)
        Dim vFACCEPTDATE As String = TIMS.GET_DateHM(v_TB_FACCEPTDATE, TB_FACCEPTDATE_HR, TB_FACCEPTDATE_MM)

        If v_ddlYEARS = "" Then sErrMsg1 &= "年度 不可為空" & vbCrLf
        If v_ddlAppStage = "" Then sErrMsg1 &= "申請階段 不可為空" & vbCrLf
        If v_rblPTYPE1 = "" Then sErrMsg1 &= "階段種類 不可為空" & vbCrLf
        If TB_SACCEPTDATE.Text = "" Then sErrMsg1 &= "受理期間設定起始日期 不可為空" & vbCrLf
        If TB_FACCEPTDATE.Text = "" Then sErrMsg1 &= "受理期間設定結束日期 不可為空" & vbCrLf
        'If sErrMsg1 <> "" Then Return sErrMsg1

        '(_i 正確值比對)
        If vSACCEPTDATE <> "" AndAlso vFACCEPTDATE <> "" Then
            '有勾選且有填上架日期的資料，再進一步檢核設定結果不得超過報名日期
            Dim iMinute3 As Long = DateDiff(DateInterval.Minute, CDate(vSACCEPTDATE), CDate(vFACCEPTDATE))
            If iMinute3 <= 0 Then sErrMsg1 &= String.Concat("[受理期間設定起始日期][結束日期]起始時間大於結束順序有誤!", iMinute3) & vbCrLf
            If sErrMsg1 <> "" Then Return sErrMsg1

            Dim iDay3 As Long = DateDiff(DateInterval.Day, CDate(vSACCEPTDATE), CDate(vFACCEPTDATE))
            If iDay3 > 180 Then
                sErrMsg1 &= String.Concat("[受理期間設定起始日期][結束日期]時間長度有誤(大於180天)!", iDay3) & vbCrLf
            End If
            If sErrMsg1 <> "" Then Return sErrMsg1
        End If
        If sErrMsg1 <> "" Then Return sErrMsg1

        Hid_PAID.Value = TIMS.ClearSQM(Hid_PAID.Value)
        Select Case v_rblPTYPE1
            Case TIMS.cst_APPSTAGE_PTYPE1_01
                '申請階段
                Dim sParms1 As New Hashtable From {{"YEARS", v_ddlYEARS}, {"APPSTAGE", v_ddlAppStage}}
                Dim sSql1 As String = " SELECT PAID,SACCEPTDATE FROM PLAN_APPSTAGE WHERE PTYPE1='01' AND YEARS=@YEARS AND APPSTAGE=@APPSTAGE"
                If Hid_PAID.Value <> "" Then
                    sParms1.Add("PAID", TIMS.VAL1(Hid_PAID.Value))
                    sSql1 &= " AND PAID!=@PAID"
                End If
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms1)
                If Hid_PAID.Value <> "" AndAlso dr1 IsNot Nothing AndAlso TIMS.Cdate3(dr1("SACCEPTDATE")) = TIMS.Cdate3(TB_SACCEPTDATE.Text) Then
                    '(修改時重複)(再判斷開始日期，如果相同才顯示此訊息)
                    sErrMsg1 &= String.Concat("年度／申請階段資料已經存在，請使用修改功能或檢查重複資料!", dr1("PAID")) & vbCrLf
                ElseIf Hid_PAID.Value = "" AndAlso dr1 IsNot Nothing Then
                    '(新增時重複)
                    sErrMsg1 &= String.Concat("年度／申請階段資料已經存在，請使用修改功能或檢查重複資料!!", dr1("PAID")) & vbCrLf
                End If

            Case TIMS.cst_APPSTAGE_PTYPE1_02
                '申複階段
                Dim sParms1 As New Hashtable From {{"YEARS", v_ddlYEARS}, {"APPSTAGE", v_ddlAppStage}}
                Dim sSql1 As String = " SELECT PAID,SACCEPTDATE FROM PLAN_APPSTAGE WHERE PTYPE1='02' AND YEARS=@YEARS AND APPSTAGE=@APPSTAGE"
                If Hid_PAID.Value <> "" Then
                    sParms1.Add("PAID", TIMS.VAL1(Hid_PAID.Value))
                    sSql1 &= " AND PAID!=@PAID"
                End If
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms1)
                If Hid_PAID.Value <> "" AndAlso dr1 IsNot Nothing AndAlso TIMS.Cdate3(dr1("SACCEPTDATE")) = TIMS.Cdate3(TB_SACCEPTDATE.Text) Then
                    '(修改時重複)(再判斷開始日期，如果相同才顯示此訊息)
                    sErrMsg1 &= String.Concat("年度／申複階段資料已經存在，請使用修改功能或檢查重複資料!", dr1("PAID")) & vbCrLf
                ElseIf Hid_PAID.Value = "" AndAlso dr1 IsNot Nothing Then
                    '(新增時重複)
                    sErrMsg1 &= String.Concat("年度／申複階段資料已經存在，請使用修改功能或檢查重複資料!!", dr1("PAID")) & vbCrLf
                End If
        End Select

        If sErrMsg1 <> "" Then Return sErrMsg1
        Return sErrMsg1
    End Function

    Private Sub SAVEDATA1()
        Dim v_ddlYEARS As String = TIMS.GetListValue(ddlYEARS)
        Dim v_ddlAppStage As String = TIMS.GetListValue(ddlAppStage)
        Dim v_rblPTYPE1 As String = TIMS.GetListValue(rblPTYPE1)

        'TB_SACCEPTDATE.Text = TIMS.cdate3(TIMS.ClearSQM(TB_SACCEPTDATE.Text))
        'TB_FACCEPTDATE.Text = TIMS.cdate3(TIMS.ClearSQM(TB_FACCEPTDATE.Text))
        TB_SACCEPTDATE.Text = TIMS.ClearSQM(TB_SACCEPTDATE.Text)
        TB_FACCEPTDATE.Text = TIMS.ClearSQM(TB_FACCEPTDATE.Text)
        Dim v_TB_SACCEPTDATE As String = If(flag_ROC, TIMS.Cdate7(TB_SACCEPTDATE.Text), TIMS.Cdate3(TB_SACCEPTDATE.Text))
        Dim v_TB_FACCEPTDATE As String = If(flag_ROC, TIMS.Cdate7(TB_FACCEPTDATE.Text), TIMS.Cdate3(TB_FACCEPTDATE.Text))
        Dim vSACCEPTDATE As String = TIMS.GET_DateHM(v_TB_SACCEPTDATE, TB_SACCEPTDATE_HR, TB_SACCEPTDATE_MM)
        Dim vFACCEPTDATE As String = TIMS.GET_DateHM(v_TB_FACCEPTDATE, TB_FACCEPTDATE_HR, TB_FACCEPTDATE_MM)
        '勾選後,關閉(更新班級清單)按鈕
        Dim V_CLOSE_RETRIEVE As String = If(CBK1_CLOSE_RETRIEVE.Checked, "Y", "")

        Hid_PAID.Value = TIMS.ClearSQM(Hid_PAID.Value)
        Dim iRst As Integer = 0
        If Hid_PAID.Value = "" Then
            Dim iPAID As Integer = DbAccess.GetNewId(objconn, "PLAN_APPSTAGE_PAID_SEQ,PLAN_APPSTAGE,PAID")
            Dim iParms As New Hashtable From {
                {"PAID", iPAID},
                {"YEARS", v_ddlYEARS},
                {"APPSTAGE", v_ddlAppStage},
                {"PTYPE1", v_rblPTYPE1},
                {"SACCEPTDATE", vSACCEPTDATE},
                {"FACCEPTDATE", vFACCEPTDATE},
                {"CLOSE_RETRIEVE", If(V_CLOSE_RETRIEVE <> "", V_CLOSE_RETRIEVE, Convert.DBNull)},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim i_Sql As String = ""
            i_Sql = " INSERT INTO PLAN_APPSTAGE(PAID ,YEARS ,APPSTAGE,PTYPE1, SACCEPTDATE,FACCEPTDATE,CLOSE_RETRIEVE ,CREATEACCT,CREATEDATE ,MODIFYACCT,MODIFYDATE)" & vbCrLf
            i_Sql &= " VALUES (@PAID ,@YEARS ,@APPSTAGE,@PTYPE1 ,@SACCEPTDATE,@FACCEPTDATE,@CLOSE_RETRIEVE ,@CREATEACCT ,GETDATE(),@MODIFYACCT ,GETDATE())" & vbCrLf
            iRst = DbAccess.ExecuteNonQuery(i_Sql, objconn, iParms)
            If iRst = 0 Then
                Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
                Return
            End If
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
            'Return
        Else
            Dim iPAID As Integer = TIMS.VAL1(Hid_PAID.Value)
            Dim uParms As New Hashtable From {
                {"SACCEPTDATE", vSACCEPTDATE},
                {"FACCEPTDATE", vFACCEPTDATE},
                {"CLOSE_RETRIEVE", If(V_CLOSE_RETRIEVE <> "", V_CLOSE_RETRIEVE, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"PAID", iPAID},
                {"PTYPE1", v_rblPTYPE1}
            }

            Dim UsSql As String = ""
            UsSql &= " UPDATE PLAN_APPSTAGE" & vbCrLf
            UsSql &= " SET SACCEPTDATE=@SACCEPTDATE ,FACCEPTDATE=@FACCEPTDATE,CLOSE_RETRIEVE=@CLOSE_RETRIEVE" & vbCrLf
            UsSql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
            UsSql &= " WHERE PAID=@PAID AND PTYPE1=@PTYPE1" & vbCrLf
            iRst = DbAccess.ExecuteNonQuery(UsSql, objconn, uParms)
            If iRst = 0 Then
                Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
                Return
            End If
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
            'Return
        End If
        Call SSearch1()
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
        Dim iPAID As Integer = TIMS.VAL1(TIMS.GetMyValue(sCmdArg, "PAID"))

        Select Case e.CommandName
            Case "BTNUPDATE" '修改
                Call CLEARDATA1()
                Call SHOW_DETAIL1(iPAID)

            Case "BTNDELETE" '刪除
                'Call CLEARDATA1()
                Call DELETEDATA1(iPAID)
        End Select
    End Sub

    Private Sub DELETEDATA1(iPAID As Integer)
        Dim iRst As Integer = 0
        Dim dParms As New Hashtable From {{"PAID", iPAID}}
        Dim dsSql As String = "DELETE PLAN_APPSTAGE WHERE PAID=@PAID"
        iRst = DbAccess.ExecuteNonQuery(dsSql, objconn, dParms)
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Return
        End If
        Common.MessageBox(Me, TIMS.cst_DELETEOKMsg1)
        'Return
        Call SSearch1()
        panel_sch.Visible = True 'False
        panel_EDIT1.Visible = False 'True
    End Sub

    Private Sub SHOW_DETAIL1(iPAID As Integer)
        If iPAID <= 0 Then Return
        Hid_PAID.Value = iPAID

        Dim sParms1 As New Hashtable From {{"PAID", iPAID}}
        Dim sSql As String = ""
        sSql &= " SELECT a.PAID" & vbCrLf
        sSql &= " ,a.YEARS ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,a.APPSTAGE,dbo.FN_GET_APPSTAGE(a.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,a.PTYPE1,CASE WHEN a.PTYPE1='01' THEN '申請' WHEN a.PTYPE1='02' THEN '申複' END PTYPE1_N" & vbCrLf
        sSql &= " ,format(a.SACCEPTDATE,'yyyy/MM/dd') SACCEPTDATE" & vbCrLf
        sSql &= " ,format(a.FACCEPTDATE,'yyyy/MM/dd') FACCEPTDATE" & vbCrLf
        sSql &= " ,format(a.SACCEPTDATE,'yyyy/MM/dd HH:mm') SACCEPTDATE_HM" & vbCrLf
        sSql &= " ,format(a.FACCEPTDATE,'yyyy/MM/dd HH:mm') FACCEPTDATE_HM" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.SACCEPTDATE) SACCEPTDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.FACCEPTDATE) FACCEPTDATE_ROC" & vbCrLf
        sSql &= " ,a.CLOSE_RETRIEVE" & vbCrLf
        'sSql &= " ,format(a.SACCEPTDATE,'HH') TB_SACCEPTDATE_HR" & vbCrLf
        'sSql &= " ,format(a.SACCEPTDATE,'mm') TB_SACCEPTDATE_MM" & vbCrLf
        'sSql &= " ,format(a.FACCEPTDATE,'HH') TB_FACCEPTDATE_HR" & vbCrLf
        'sSql &= " ,format(a.FACCEPTDATE,'mm') TB_FACCEPTDATE_MM" & vbCrLf
        'sSql &= " ,a.CREATEACCT ,a.CREATEDATE ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sSql &= " FROM PLAN_APPSTAGE a" & vbCrLf
        sSql &= " WHERE a.PAID=@PAID " & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms1)
        If dr1 Is Nothing Then Return

        Common.SetListItem(ddlYEARS, dr1("YEARS"))
        Common.SetListItem(ddlAppStage, dr1("APPSTAGE"))
        Common.SetListItem(rblPTYPE1, dr1("PTYPE1"))

        TB_SACCEPTDATE.Text = If(flag_ROC, TIMS.ClearSQM(dr1("SACCEPTDATE_ROC")), TIMS.Cdate3(dr1("SACCEPTDATE")))
        TIMS.SET_DateHM(CDate(dr1("SACCEPTDATE_HM")), TB_SACCEPTDATE_HR, TB_SACCEPTDATE_MM)
        TB_FACCEPTDATE.Text = If(flag_ROC, TIMS.ClearSQM(dr1("FACCEPTDATE_ROC")), TIMS.Cdate3(dr1("FACCEPTDATE")))
        TIMS.SET_DateHM(CDate(dr1("FACCEPTDATE_HM")), TB_FACCEPTDATE_HR, TB_FACCEPTDATE_MM)
        CBK1_CLOSE_RETRIEVE.Checked = If($"{dr1("CLOSE_RETRIEVE")}" = "Y", True, False)
        'Common.SetListItem(TB_SACCEPTDATE_HR, dr1("TB_SACCEPTDATE_HR"))
        'Common.SetListItem(TB_SACCEPTDATE_MM, dr1("TB_SACCEPTDATE_MM"))
        'Common.SetListItem(TB_FACCEPTDATE_HR, dr1("TB_FACCEPTDATE_HR"))
        'Common.SetListItem(TB_FACCEPTDATE_MM, dr1("TB_FACCEPTDATE_MM"))
        ddlYEARS.Enabled = False
        ddlAppStage.Enabled = False
        TIMS.Tooltip(ddlYEARS, "年度不可修改", True)
        TIMS.Tooltip(ddlAppStage, "申請階段不可修改", True)
        rblPTYPE1.Enabled = False
        TIMS.Tooltip(rblPTYPE1, String.Concat(cst_title_申請階段種類_1, "不可修改!"), True)

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
                TIMS.SetMyValue(sCmdArg, "PAID", drv("PAID"))
                BTNUPDATE.CommandArgument = sCmdArg '修改
                BTNDELETE.CommandArgument = sCmdArg '刪除
                Dim sDELMSG1 As String = String.Concat("return confirm('", "您確定要刪除第", iDGSeqNo, "筆資料嗎?", "');")
                BTNDELETE.Attributes.Add("onclick", sDELMSG1)
        End Select
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    '    Sub ClearDATA1()
    '        TB_OrgName.Text = "" 'Convert.ToString(dr2("ORGNAME"))

    '        Hid_OCID.Value = "" 'Convert.ToString(dr("OCID"))
    '        Hid_PlanID.Value = "" 'Convert.ToString(dr("PlanID"))
    '        Hid_ComIDNO.Value = "" 'Convert.ToString(dr("ComIDNO"))
    '        Hid_SeqNo.Value = "" 'Convert.ToString(dr("SeqNo"))

    '        TB_OCID.Text = "" 'Convert.ToString(dr("OCID"))
    '        TB_ClassCName.Text = "" 'Convert.ToString(dr("CLASSCNAME"))
    '        TB_CYCLTYPE.Text = "" 'TIMS.FmtCyclType(dr("CyclType"))

    '        TB_ContactName.Text = "" 'dr("ContactName").ToString
    '        TB_ContactPhone.Text = "" 'dr("ContactPhone").ToString
    '        TB_ContactEmail.Text = "" 'dr("ContactEmail").ToString
    '        TB_ContactFax.Text = "" 'dr("ContactFax").ToString
    '    End Sub

    '    Sub Save_REVISE(ByVal OCIDVal As String)
    '        OCIDVal = TIMS.ClearSQM(OCIDVal)

    '        Dim sParms As New Hashtable
    '        sParms.Add("OCID", OCIDVal)
    '        Dim sSql As String = "SELECT 1 FROM CLASS_SUBINFO WHERE OCID=@OCID"
    '        Dim dr As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms)
    '        If dr Is Nothing Then Return

    '        Dim uParms As New Hashtable
    '        uParms.Add("REVISEACCT", sm.UserInfo.UserID)
    '        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
    '        uParms.Add("OCID", OCIDVal)
    '        Dim u_Sql As String = ""
    '        u_Sql &= " UPDATE CLASS_SUBINFO SET REVISEACCT=@REVISEACCT,REVISEDATE=GETDATE(),UNLOCKSTATE=NULL" & vbCrLf
    '        u_Sql &= ",MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() WHERE OCID=@OCID" & vbCrLf
    '        DbAccess.ExecuteNonQuery(u_Sql, objconn, uParms)
    '    End Sub

    '    Sub Save_PLANINFO()
    '        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
    '        Hid_PlanID.Value = TIMS.ClearSQM(Hid_PlanID.Value)
    '        Hid_ComIDNO.Value = TIMS.ClearSQM(Hid_ComIDNO.Value)
    '        Hid_SeqNo.Value = TIMS.ClearSQM(Hid_SeqNo.Value)

    '        If Hid_OCID.Value = "" OrElse Hid_PlanID.Value = "" OrElse Hid_ComIDNO.Value = "" OrElse Hid_SeqNo.Value = "" Then Return

    '        Dim sParms As New Hashtable
    '        sParms.Add("OCID", Hid_OCID.Value)
    '        sParms.Add("PlanID", Hid_PlanID.Value)
    '        sParms.Add("ComIDNO", Hid_ComIDNO.Value)
    '        sParms.Add("SeqNO", Hid_SeqNo.Value)
    '        Dim sql As String = ""
    '        sql = "SELECT * FROM dbo.CLASS_CLASSINFO WHERE OCID=@OCID AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
    '        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, sParms)
    '        If dr Is Nothing Then Return

    '        Dim sParms2 As New Hashtable
    '        sParms2.Add("PlanID", Hid_PlanID.Value)
    '        sParms2.Add("ComIDNO", Hid_ComIDNO.Value)
    '        sParms2.Add("SeqNO", Hid_SeqNo.Value)
    '        Dim sql2 As String = ""
    '        sql2 = " SELECT * FROM dbo.PLAN_PLANINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
    '        Dim dr2 As DataRow = DbAccess.GetOneRow(sql2, objconn, sParms2)
    '        If dr2 Is Nothing Then Return

    '        TB_ContactName.Text = TIMS.ClearSQM(TB_ContactName.Text)
    '        TB_ContactPhone.Text = TIMS.ClearSQM(TB_ContactPhone.Text)
    '        TB_ContactEmail.Text = TIMS.ClearSQM(TB_ContactEmail.Text)
    '        TB_ContactFax.Text = TIMS.ClearSQM(TB_ContactFax.Text)
    '        TB_ContactEmail.Text = TIMS.ChangeEmail(TB_ContactEmail.Text)

    '        Dim uParms As New Hashtable
    '        uParms.Add("ContactName", TB_ContactName.Text)
    '        uParms.Add("ContactPhone", TB_ContactPhone.Text)
    '        uParms.Add("ContactEmail", TB_ContactEmail.Text)
    '        uParms.Add("ContactFax", TB_ContactFax.Text)
    '        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
    '        uParms.Add("PlanID", Hid_PlanID.Value)
    '        uParms.Add("ComIDNO", Hid_ComIDNO.Value)
    '        uParms.Add("SeqNO", Hid_SeqNo.Value)
    '        Dim u_Sql As String = ""
    '        u_Sql = "" & vbCrLf
    '        u_Sql &= " UPDATE PLAN_PLANINFO" & vbCrLf
    '        u_Sql &= " SET CONTACTNAME=@ContactName" & vbCrLf
    '        u_Sql &= " ,CONTACTPHONE=@ContactPhone" & vbCrLf
    '        u_Sql &= " ,CONTACTEMAIL=@ContactEmail" & vbCrLf
    '        u_Sql &= " ,CONTACTFAX=@ContactFax" & vbCrLf
    '        u_Sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
    '        u_Sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
    '        u_Sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO" & vbCrLf
    '        DbAccess.ExecuteNonQuery(u_Sql, objconn, uParms)
    '    End Sub

    '    Sub SHOW_DETAIL1()
    '        EdtPanel1.Visible = True
    '        SchPanel.Visible = False
    '        SchPanel2.Visible = SchPanel.Visible

    '        TB_OrgName.Enabled = False
    '        TB_OCID.Enabled = False
    '        TB_ClassCName.Enabled = False
    '        TB_CYCLTYPE.Enabled = False
    '    End Sub

    '    Sub SHOW_SEARCH1()
    '        EdtPanel1.Visible = False
    '        SchPanel.Visible = True
    '        SchPanel2.Visible = SchPanel.Visible
    '        'If Not Me.ViewState("ClassSearchStr") Is Nothing Then Session("ClassSearchStr") = Me.ViewState("ClassSearchStr")
    '        ''Response.Redirect("TC_01_004.aspx?ID=" & Request("ID") & "")
    '        'Dim url1 As String = "TC_01_004.aspx?ID=" & Request("ID") & ""
    '        'Call TIMS.Utl_Redirect(Me, objconn, url1)
    '    End Sub

    '    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
    '        Call ClearDATA1()
    '        Call SHOW_SEARCH1()
    '    End Sub

    '    Protected Sub BtnSAVEDATA1_Click(sender As Object, e As EventArgs) Handles BtnSAVEDATA1.Click
    '        Call Save_PLANINFO()  '儲存(PLAN_PLANINFO)
    '        Call Save_REVISE(Hid_OCID.Value)
    '        Common.MessageBox(Me, "儲存完畢")
    '        Call ClearDATA1()
    '        Call SHOW_SEARCH1()
    '        Call sSearch1()
    '    End Sub
End Class