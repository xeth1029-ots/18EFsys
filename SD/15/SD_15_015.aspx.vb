Public Class SD_15_015
    Inherits AuthBasePage

    Dim vsMyValue As String = ""
    Const cst_printFN1 As String = "SD_15_015_R"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1
        '檢查Session是否存在 End

        '訓練機構
        If Not IsPostBack Then
            Call cCreate1()
        End If
    End Sub

    Sub cCreate1()
        msg.Text = ""
        DataGridTable.Visible = False

        '年度
        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        '轄區
        DistID = TIMS.Get_DistID(DistID)
        Common.SetListItem(DistID, sm.UserInfo.DistID)

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_AppStage2(AppStage), TIMS.Get_AppStage(AppStage))
        End If

        btnSearch1.Attributes("onclick") = "javascript:return CheckSearch1();"
        'btnPrint1
        btnPrint1.Attributes("onclick") = "javascript:return CheckSearch1();"
    End Sub

    ''' <summary>
    ''' 設定 ViewState Value
    ''' </summary>
    Sub sUtl_ViewStateValue()
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)
        'ViewState("Years") = If(Me.yearlist.SelectedValue <> "", Me.yearlist.SelectedValue, "")
        'ViewState("DistID") = If(Me.DistID.SelectedValue <> "", Me.DistID.SelectedValue, "")
        ViewState("Years") = TIMS.GetListValue(yearlist)
        ViewState("DistID") = TIMS.GetListValue(DistID)
        ViewState("TPlanID") = sm.UserInfo.TPlanID
        ViewState("TeachCName") = ""
        ViewState("IDNO") = ""

        If txtName.Text <> "" Then ViewState("TeachCName") = TIMS.Get_SplitValeu1(txtName.Text)
        If txtIDNO.Text <> "" Then ViewState("IDNO") = TIMS.Get_SplitValeu1(txtIDNO.Text)

        ViewState("STDate1") = If(Me.STDate1.Text <> "", Me.STDate1.Text, "")
        ViewState("STDate2") = If(Me.STDate2.Text <> "", Me.STDate2.Text, "")
        ViewState("FTDate1") = If(Me.FTDate1.Text <> "", Me.FTDate1.Text, "")
        ViewState("FTDate2") = If(Me.FTDate2.Text <> "", Me.FTDate2.Text, "")
        'Dim v_audit As String = TIMS.GetListValue(audit)  'audit.SelectedValue '審核狀態有選值
        'ViewState("audit") = v_audit
        ViewState("audit") = TIMS.GetListValue(audit)  'audit.SelectedValue '審核狀態有選值
        Select Case ViewState("audit")
            Case "Y", "N"
            Case Else
                ViewState("audit") = ""
        End Select
        'Dim v_AppStage As String = ""
        'If tr_AppStage_TP28.Visible Then v_AppStage = TIMS.GetListValue(AppStage)
        ViewState("AppStage") = TIMS.GetListValue(AppStage)  'audit.SelectedValue '審核狀態有選值

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&Years=" & ViewState("Years")
        MyValue &= "&DistID=" & ViewState("DistID")
        MyValue &= "&TPlanID=" & ViewState("TPlanID")
        MyValue &= "&TeachCName=" & ViewState("TeachCName").ToString.Replace("'", "\'")
        MyValue &= "&IDNO=" & ViewState("IDNO").ToString.Replace("'", "\'")
        MyValue &= "&STDate1=" & ViewState("STDate1")
        MyValue &= "&STDate2=" & ViewState("STDate2")
        MyValue &= "&FTDate1=" & ViewState("FTDate1")
        MyValue &= "&FTDate2=" & ViewState("FTDate2")
        Select Case ViewState("audit")
            Case "Y"
                MyValue &= "&auditY=" & ViewState("audit")
            Case "N"
                MyValue &= "&auditN=" & ViewState("audit")
            Case Else
                ViewState("audit") = ""
        End Select
        MyValue &= "&audit=" & ViewState("audit")
        MyValue &= "&AppStage=" & ViewState("AppStage")
        vsMyValue = MyValue
    End Sub

    ''' <summary>取得資料</summary>
    ''' <returns></returns>
    Function GetSchDt(iTYPE As Integer) As DataTable
        'iTYPE: 1:明細資料 2:統計資料
        ViewState("audit") = TIMS.ClearSQM(ViewState("audit"))

        Dim sql As String = ""
        'WP1 PLAN_PLANINFO/CLASS_CLASSINFO
        sql &= " WITH WP1 AS  ( SELECT oo.ORGNAME,PP.CLASSNAME ,PP.CYCLTYPE" & vbCrLf
        sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        sql &= " ,pp.STDATE,pp.FDDATE" & vbCrLf
        sql &= " ,cc.OCID,pp.PSNO28" & vbCrLf
        sql &= " ,pp.AppStage" & vbCrLf
        sql &= " FROM PLAN_PLANINFO pp" & vbCrLf
        sql &= " JOIN PLAN_VERREPORT PVR WITH(NOLOCK) ON PVR.PLANID=PP.PLANID AND PVR.COMIDNO=PP.COMIDNO AND PVR.SEQNO=PP.SEQNO" & vbCrLf        '
        sql &= " JOIN AUTH_RELSHIP rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.planid=rr.planid" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.COMIDNO=pp.COMIDNO" & vbCrLf
        '如果審核狀態有選值
        If ViewState("audit") <> "" Then
            Select Case ViewState("audit")
                Case "N" '如果是選審核中
                    sql &= " LEFT JOIN CLASS_CLASSINFO cc on cc.PLANID=pp.PLANID and cc.COMIDNO=pp.COMIDNO and cc.SEQNO=pp.SEQNO" & vbCrLf
                    sql &= " AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                Case Else ' "Y" '如果是選(已通過)已審核
                    sql &= " JOIN CLASS_CLASSINFO cc on cc.PLANID=pp.PLANID and cc.COMIDNO=pp.COMIDNO and cc.SEQNO=pp.SEQNO" & vbCrLf
                    sql &= " AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
            End Select
        Else
            '(沒有值)
            sql &= " LEFT JOIN CLASS_CLASSINFO cc on cc.PLANID=pp.PLANID and cc.COMIDNO=pp.COMIDNO and cc.SEQNO=pp.SEQNO" & vbCrLf
            sql &= " AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
        End If
        sql &= " WHERE PP.ISAPPRPAPER='Y' AND PVR.ISAPPRPAPER='Y'" & vbCrLf
        sql &= " AND pp.RESULTBUTTON IS NULL " & vbCrLf
        'sql &= " and ip.Years='2019' and ip.DistID ='001' and ip.TPlanID ='28'" & vbCrLf
        sql &= " and ip.Years='" & ViewState("Years") & "'" & vbCrLf
        sql &= " and ip.DistID ='" & ViewState("DistID") & "'" & vbCrLf
        sql &= " and ip.TPlanID ='" & ViewState("TPlanID") & "'" & vbCrLf
        If ViewState("STDate1") <> "" Then sql &= " and pp.STDate >= " & TIMS.To_date(ViewState("STDate1")) & vbCrLf
        If ViewState("STDate2") <> "" Then sql &= " and pp.STDate <= " & TIMS.To_date(ViewState("STDate2")) & vbCrLf
        If ViewState("FTDate1") <> "" Then sql &= " and pp.FDDATE >= " & TIMS.To_date(ViewState("FTDate1")) & vbCrLf
        If ViewState("FTDate2") <> "" Then sql &= " and pp.FDDATE <= " & TIMS.To_date(ViewState("FTDate2")) & vbCrLf
        If ViewState("AppStage") <> "" Then sql &= " AND pp.AppStage='" & ViewState("AppStage") & "'" & vbCrLf '依申請階段
        '如果審核狀態有選值
        If ViewState("audit") <> "" Then
            Select Case ViewState("audit")
                Case "N" '如果是選審核中 RESULTBUTTON IS NULL
                    'sql &= " AND (pp.AppliedResult NOT IN ('Y','N') OR pp.AppliedResult IS NULL) " & vbCrLf
                    sql &= " AND (pp.AppliedResult NOT IN ('Y') OR pp.AppliedResult IS NULL)" & vbCrLf
                Case Else ' "Y" '如果是選(已通過)已審核
                    'sql &= " AND pp.AppliedResult IN ('Y','N') " & vbCrLf
                    sql &= " AND pp.AppliedResult IN ('Y')" & vbCrLf
            End Select
        End If
        sql &= " )" & vbCrLf

        'WP2 PLAN_TRAINDESC
        sql &= " ,WP2 AS (" & " SELECT pp.PLANID,pp.COMIDNO,pp.SEQNO,pt.TECHID,pt.PNAME ,pt.STrainDate,pt.PHOUR" & vbCrLf
        sql &= " FROM WP1 pp" & vbCrLf
        sql &= " JOIN PLAN_TRAINDESC pt ON pt.PLANID=pp.PLANID and pt.COMIDNO=pp.COMIDNO and pt.SEQNO=pp.SEQNO" & " )" & vbCrLf
        'SEL
        Dim dt1 As DataTable = Nothing
        If iTYPE = 1 Then
            'WP3 GROUP 
            sql &= " ,WP3 AS ( SELECT pp.PLANID,pp.COMIDNO,pp.SEQNO,pt.PNAME" & vbCrLf
            sql &= " ,dbo.FN_GETWEEKDAY(pt.STrainDate) WEEKDAY1" & vbCrLf
            sql &= " ,SUM(pt.PHOUR) PHOUR" & vbCrLf
            sql &= " ,MAX(pt.TECHID) TECHID" & vbCrLf
            sql &= " FROM WP1 pp" & vbCrLf
            sql &= " JOIN WP2 pt ON pt.PLANID=pp.PLANID and pt.COMIDNO=pp.COMIDNO and pt.SEQNO=pp.SEQNO" & vbCrLf
            sql &= " GROUP BY pp.PLANID,pp.COMIDNO,pp.SEQNO,pt.TECHID,pt.PNAME" & vbCrLf
            sql &= " ,dbo.FN_GETWEEKDAY(pt.STrainDate)" & " )" & vbCrLf

            sql &= " SELECT TT.TEACHCNAME,TT.IDNO" & vbCrLf
            sql &= " ,dbo.FN_GET_MASK1(TT.IDNO) IDNO_MK" & vbCrLf
            'sql &= " ,dbo.SUBSTR3(TT.IDNO,1,2) + '*****' + dbo.SUBSTR3(TT.IDNO,-2,2) IDNO2" & vbCrLf
            sql &= " ,p1.ORGNAME" & vbCrLf
            sql &= " ,p1.CLASSNAME" & vbCrLf
            sql &= " ,p1.OCID" & vbCrLf
            sql &= " ,p1.PSNO28" & vbCrLf
            sql &= " ,dbo.FN_GET_APPSTAGE(p1.AppStage) AppStage" & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(p1.CLASSNAME ,p1.CYCLTYPE) CLASSNAME2" & vbCrLf
            sql &= " ,dbo.FN_CDATE1B(p1.STDATE) STDATE" & vbCrLf
            sql &= " ,dbo.FN_CDATE1B(p1.FDDATE) FTDATE" & vbCrLf
            sql &= " ,dbo.FN_CDATE1B(p1.STDATE)+'~'+dbo.FN_CDATE1B(p1.FDDATE) SFTDATE" & vbCrLf
            sql &= " ,p3.PNAME ,p3.PHOUR ,p3.WEEKDAY1" & vbCrLf
            sql &= " FROM TEACH_TEACHERINFO tt" & vbCrLf
            sql &= " JOIN WP3 p3 on p3.TECHID=tt.TECHID" & vbCrLf
            sql &= " JOIN WP1 p1 on p1.PLANID=p3.PLANID and p1.COMIDNO=p3.COMIDNO and p1.SEQNO=p3.SEQNO" & vbCrLf
            sql &= " WHERE 1=1 " & vbCrLf
            If ViewState("TeachCName") <> "" Then sql &= " and tt.TeachCName in (" & ViewState("TeachCName") & ") " & vbCrLf
            If ViewState("IDNO") <> "" Then sql &= " and tt.IDNO in (" & ViewState("IDNO") & ") " & vbCrLf

            sql &= " ORDER BY tt.TeachCName ,tt.IDNO" & vbCrLf
            dt1 = DbAccess.GetDataTable(sql, objconn)
            Return dt1

        ElseIf iTYPE = 2 Then
            'WP3 GROUP 
            sql &= " ,WP3 AS ( SELECT pp.PLANID,pp.COMIDNO,pp.SEQNO,pt.TECHID" & vbCrLf
            sql &= " ,SUM(pt.PHOUR) PHOUR" & vbCrLf
            sql &= " FROM WP1 pp" & vbCrLf
            sql &= " JOIN WP2 pt ON pt.PLANID=pp.PLANID and pt.COMIDNO=pp.COMIDNO and pt.SEQNO=pp.SEQNO" & vbCrLf
            sql &= " GROUP BY pp.PLANID,pp.COMIDNO,pp.SEQNO,pt.TECHID" & " )" & vbCrLf

            sql &= " SELECT TT.TEACHCNAME,TT.IDNO" & vbCrLf
            sql &= " ,dbo.FN_GET_MASK1(TT.IDNO) IDNO_MK" & vbCrLf
            'sql &= " ,dbo.SUBSTR3(TT.IDNO,1,2) + '*****' + dbo.SUBSTR3(TT.IDNO,-2,2) IDNO2" & vbCrLf
            sql &= " ,p1.ORGNAME" & vbCrLf
            sql &= " ,p1.CLASSNAME" & vbCrLf
            sql &= " ,p1.OCID" & vbCrLf
            sql &= " ,p1.PSNO28" & vbCrLf
            sql &= " ,dbo.FN_GET_APPSTAGE(p1.AppStage) AppStage" & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(p1.CLASSNAME ,p1.CYCLTYPE) CLASSNAME2" & vbCrLf
            sql &= " ,dbo.FN_CDATE1B(p1.STDATE) STDATE" & vbCrLf
            sql &= " ,dbo.FN_CDATE1B(p1.FDDATE) FTDATE" & vbCrLf
            sql &= " ,dbo.FN_CDATE1B(p1.STDATE)+'~'+dbo.FN_CDATE1B(p1.FDDATE) SFTDATE" & vbCrLf
            sql &= " ,p3.PHOUR" & vbCrLf
            sql &= " FROM TEACH_TEACHERINFO tt" & vbCrLf
            sql &= " JOIN WP3 p3 on p3.TECHID=tt.TECHID" & vbCrLf
            sql &= " JOIN WP1 p1 on p1.PLANID=p3.PLANID and p1.COMIDNO=p3.COMIDNO and p1.SEQNO=p3.SEQNO" & vbCrLf
            sql &= " WHERE 1=1 " & vbCrLf
            If ViewState("TeachCName") <> "" Then sql &= " and tt.TeachCName in (" & ViewState("TeachCName") & ") " & vbCrLf
            If ViewState("IDNO") <> "" Then sql &= " and tt.IDNO in (" & ViewState("IDNO") & ") " & vbCrLf

            sql &= " ORDER BY tt.TeachCName ,tt.IDNO" & vbCrLf
            dt1 = DbAccess.GetDataTable(sql, objconn)
            Return dt1

        End If
        Return dt1
    End Function

    'SQL
    Sub search1()
        Dim dt As DataTable = GetSchDt(1)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
        End Select
    End Sub

#Region "ExportData1"
    Sub ExportData1()
        Dim dt As DataTable = GetSchDt(1)
        'Dim rtnPath As String = Request.FilePath
        'If dt Is Nothing Then
        '    Common.MessageBox(Me, "資料庫查詢失敗，請重新查詢", rtnPath)
        '    Exit Sub
        'End If
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料，請重新查詢")
            Exit Sub
        End If

        ExpReport1(dt)
    End Sub

    '匯出 Response 
    Sub ExpReport1(ByRef dt As DataTable)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dt.Rows(0)

        '匯出表頭名稱
        Dim sFileName1 As String = String.Concat("師資授課時數統計明細表", TIMS.GetDateNo2())

        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= String.Concat("&ACT=", sFileName1) & vbCrLf
        'sMemo &= "&TPLANID=" & sm.UserInfo.TPlanID & vbCrLf
        'sMemo &= "&OCIDValue1=" & OCIDValue1.Value & vbCrLf
        sMemo &= String.Concat("&NAME=", dr1("TEACHCNAME")) & vbCrLf
        sMemo &= String.Concat("&IDNO=", dr1("IDNO_MK")) & vbCrLf
        sMemo &= String.Concat("&COUNT=", dt.Rows.Count) & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip1, "", sMemo, objconn)

        Const cst_tit1 As String = "師資姓名,身分證字號,訓練機構,班別名稱,班級申請流水編號,訓練期間,星期,上課時間,授課時數"
        Const cst_tit2 As String = "TEACHCNAME,IDNO_MK,ORGNAME,CLASSNAME2,PSNO28,SFTDATE,WEEKDAY1,PNAME,PHOUR"
        Dim sta_tit1 As String() = Split(cst_tit1, ",")
        Dim sta_tit2 As String() = Split(cst_tit2, ",")

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        'Dim s_TDVALUE As String = ""
        'Dim flag_isDate As Boolean = False
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = "<tr>" & vbCrLf
        For Each str1 As String In sta_tit1
            ExportStr &= "<td>" & str1 & "</td>" & vbTab
        Next
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Rows
            'For Each dr As DataRow In dt.Rows
            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            For Each str_cl2 As String In sta_tit2
                'flag_isDate = False
                'If str_cl2.Equals("AppliedDate") Then flag_isDate = True
                'If str_cl2.Equals("STDate") Then flag_isDate = True
                'If str_cl2.Equals("FDDate") Then flag_isDate = True
                's_TDVALUE = If(flag_isDate, TIMS.cdate3(dr(str_cl2)), Convert.ToString(dr(str_cl2)))
                'If str_cl2.Equals("AppliedResult") Then
                '    s_TDVALUE = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(dr("AppliedResult")), Convert.ToString(dr("RESULTBUTTON")))
                'End If
                'ExportStr &= "<td>" & s_TDVALUE & "</td>" & vbTab
                ExportStr &= String.Concat("<td>", Convert.ToString(dr(str_cl2)), "</td>") & vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub
#End Region

    ''' <summary>
    ''' 查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnSearch1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch1.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Call sUtl_ViewStateValue()  '設定 ViewState Value
        Call search1()
    End Sub

    ''' <summary>
    ''' 列印
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnPrint1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint1.Click
        Call sUtl_ViewStateValue()  '設定 ViewState Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, vsMyValue)
    End Sub


    ''' <summary> 匯出明細表 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        'Dim Errmsg As String = ""
        'Call CheckData1(Errmsg)
        'If Errmsg <> "" Then
        '    Common.MessageBox(Me, Errmsg)
        '    Exit Sub
        'End If
        Call sUtl_ViewStateValue()  '設定 ViewState Value
        Call ExportData1()
    End Sub

#Region "ExportData2"
    Sub ExportData2()
        Dim dt As DataTable = GetSchDt(2)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料，請重新查詢")
            Exit Sub
        End If
        ExpReport2(dt)
    End Sub

    '匯出 Response 
    Sub ExpReport2(ByRef dt As DataTable)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dt.Rows(0)

        '匯出表頭名稱
        Dim sFileName1 As String = String.Concat("師資授課時數統計表", TIMS.GetDateNo2())

        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= String.Concat("&ACT=", sFileName1) & vbCrLf
        'sMemo &= "&TPLANID=" & sm.UserInfo.TPlanID & vbCrLf
        'sMemo &= "&OCIDValue1=" & OCIDValue1.Value & vbCrLf
        sMemo &= String.Concat("&NAME=", dr1("TEACHCNAME")) & vbCrLf
        sMemo &= String.Concat("&IDNO=", dr1("IDNO_MK")) & vbCrLf
        sMemo &= String.Concat("&COUNT=", dt.Rows.Count) & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip1, "", sMemo, objconn)

        Const cst_tit1 As String = "師資姓名,身分證字號,訓練機構,班別名稱,班級申請流水編號,訓練期間,授課時數"
        Const cst_tit2 As String = "TEACHCNAME,IDNO_MK,ORGNAME,CLASSNAME2,PSNO28,SFTDATE,PHOUR"
        Dim sta_tit1 As String() = Split(cst_tit1, ",")
        Dim sta_tit2 As String() = Split(cst_tit2, ",")

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        'Dim s_TDVALUE As String = ""
        'Dim flag_isDate As Boolean = False
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = "<tr>" & vbCrLf
        For Each str1 As String In sta_tit1
            ExportStr &= "<td>" & str1 & "</td>" & vbTab
        Next
        ExportStr += "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Rows
            'For Each dr As DataRow In dt.Rows
            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            For Each str_cl2 As String In sta_tit2
                'flag_isDate = False
                'If str_cl2.Equals("AppliedDate") Then flag_isDate = True
                'If str_cl2.Equals("STDate") Then flag_isDate = True
                'If str_cl2.Equals("FDDate") Then flag_isDate = True
                's_TDVALUE = If(flag_isDate, TIMS.cdate3(dr(str_cl2)), Convert.ToString(dr(str_cl2)))
                'If str_cl2.Equals("AppliedResult") Then
                '    s_TDVALUE = Get_AppliedResultTxt(sm.UserInfo.TPlanID, Convert.ToString(dr("AppliedResult")), Convert.ToString(dr("RESULTBUTTON")))
                'End If
                'ExportStr &= "<td>" & s_TDVALUE & "</td>" & vbTab
                ExportStr &= String.Concat("<td>", Convert.ToString(dr(str_cl2)), "</td>") & vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub
#End Region

    Protected Sub btnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click
        Call sUtl_ViewStateValue()  '設定 ViewState Value
        Call ExportData2()
    End Sub
End Class
