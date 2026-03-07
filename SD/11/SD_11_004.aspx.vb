Partial Class SD_11_004
    Inherits AuthBasePage

    Dim WriteStatus As Boolean '可寫入資料庫狀態
    Dim FillFormDate As String '填寫日期 
    Dim SOCID As String        '學員編號

    '(產投)
    'STUD_QUESTIONFAC (2016 OLD)/STUD_QUESTIONFAC2 (2017) /dbo.fn_GET_GOVCNT
    'SD_11_004_*.jrxml
    'SD_14_012 (參訓學員補助經費申請表)/ SD_15_007 (訓後意見調查統計表)

    'ReportQuery
    'SQControl.aspx
    'SD_11_004_emp & (year)
    'SD_11_004_emp08
    'SD_11_004_n08
    'SD_11_004_08
    '(SD_11_004_)(SD_11_004_emp)
    'SD_11_004_emp12 @BussinessTrain 
    'SD_11_004_n12 @BussinessTrain
    'SD_11_004_12 @BussinessTrain
    'SD_05_008_D.aspx 注意修改。

    ''' <summary>
    ''' EDIT NAME ?
    ''' </summary>
    Dim sPrtFileName1 As String = "" 'SD_11_004
    Const cst_rptEmp As String = "SD_11_004_emp" '空白表 SD_11_004_emp*.jrxml SD_11_004_emp17*.jrxml -空白
    Const cst_rptN As String = "SD_11_004_n" '班級問卷表 SD_11_004_n17 *.jrxml -空白
    Const cst_rptData As String = "SD_11_004_" '學員答案表 SD_11_004_17.jrxml 
    Const cst_Addaspx As String = "SD_11_004_add@xxx.aspx?" '@xxx 程式用 

    Const cst_ptInsert As String = "Insert" 'e.CommandName'.aspx
    Const cst_ptDelete As String = "Delete" '.aspx
    Const cst_ptCheck As String = "Check" 'e.CommandName'.aspx
    Const cst_ptEdit As String = "Edit" 'e.CommandName'.aspx
    Const cst_ptClear As String = "Clear" 'e.CommandName
    Const cst_ptView As String = "View"
    Const cst_ptPrint As String = "Print"
    Const cst_ptBack As String = "Back" '.aspx(ProcessType)
    Const ss_QuestionFacSearchStr As String = "QuestionFacSearchStr"

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在-------------------------- Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End

        msg.Text = ""
        eMeng.Style("display") = VeMeng.Text 'none/""

        'check_add.Value = "0"
        'check_search.Value = "0"
        'check_del.Value = "0"
        'check_mod.Value = "0"
        'If blnCanAdds Then check_add.Value = "1"
        'If blnCanSech Then check_search.Value = "1"
        'If blnCanDel Then check_del.Value = "1"
        'If blnCanMod Then check_mod.Value = "1"
        'search.Enabled = True
        'If Not blnCanSech Then search.Enabled = False
        'If Not blnCanSech Then TIMS.Tooltip(search, "無此權限")

        'search.Attributes("onclick") = "javascript:return search1()"
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

#Region "(No Use)"

        ''分頁設定 Start
        'DataGridPage1.MyDataGrid = DG_stud
        ''分頁設定 End

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        check_add.Value = "0"
        '        check_search.Value = "0"
        '        check_del.Value = "0"
        '        check_mod.Value = "0"

        '        If FunDr("Adds") = "1" Then check_add.Value = "1"
        '        If FunDr("Sech") = "1" Then check_search.Value = "1"
        '        If FunDr("Del") = "1" Then check_del.Value = "1"
        '        If FunDr("Mod") = "1" Then check_mod.Value = "1"

        '        search.Enabled = True
        '        If check_search.Value = "0" Then
        '            search.Enabled = False
        '            TIMS.Tooltip(search, "無此權限")
        '        End If
        '    End If
        'End If

        '給予1年緩衝期。20140224黃芊雅。
        'If sm.UserInfo.LID = 2 Then '如果是委訓單位
        '    check_add.Value = "0" '不可新增
        '    check_mod.Value = "0" '不可修改
        'End If

#End Region

        If Not Page.IsPostBack Then
            Years.Value = sm.UserInfo.Years - 1911
            eMeng.Style("display") = VeMeng.Text
            VeMeng.Text = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            StudentTable.Style.Item("display") = "none"
            Dim re_ocid As String = TIMS.ClearSQM(Request("ocid"))
            Dim ProcessType As String = TIMS.ClearSQM(Request("ProcessType"))

            Select Case ProcessType
                Case cst_ptBack
                    If Not Session(ss_QuestionFacSearchStr) Is Nothing Then
                        Dim MyValue As String = ""
                        center.Text = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "center")
                        RIDValue.Value = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "RIDValue")
                        TMID1.Text = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "TMID1")
                        OCID1.Text = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "OCID1")
                        TMIDValue1.Value = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "TMIDValue1")
                        OCIDValue1.Value = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "OCIDValue1")
                        MyValue = TIMS.GetMyValue(Session(ss_QuestionFacSearchStr), "Button1")
                        If MyValue = "True" Then
                            Call Search1()
                            If OCIDValue1.Value <> "" Then Call Search2(OCIDValue1.Value)
                            'Call search_Click(sender, e)
                        End If
                        Session(ss_QuestionFacSearchStr) = Nothing
                    End If
            End Select

            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        '**by Milor 20070418----end
        '依不同的年度呼叫不同的空白問卷
        Dim sql As String = ""
        Dim okind As String = ""
        sql = " SELECT b.ORGKIND FROM AUTH_ACCOUNT a JOIN Org_OrgInfo b ON a.OrgID = b.OrgID WHERE a.Account = '" & sm.UserInfo.UserID & "' "
        okind = DbAccess.ExecuteScalar(sql, objconn)
        PrintBlank.Text = "列印空白表單"
        PrintBlank2.Text = "列印空白表單(" & TIMS.Get_PName28(Me, "W", objconn) & ")"
        PrintBlank.Visible = True '列印空白表單(產業人才)
        PrintBlank2.Visible = False '列印空白表單(在職勞工)

        Select Case sm.UserInfo.TPlanID
            Case "28"
                PrintBlank.Text = "列印空白表單(" & TIMS.Get_PName28(Me, "G", objconn) & ")" 'cst_OrgKind2GTxt
                PrintBlank2.Text = "列印空白表單(" & TIMS.Get_PName28(Me, "W", objconn) & ")" 'cst_OrgKind2WTxt
                If sm.UserInfo.LID = "2" Then    '如果是委訓單位，必須依據機構別帶出列印空白問卷按鈕
                    If okind = "10" Then '提升勞工自主學習計畫
                        PrintBlank.Visible = False
                        PrintBlank2.Visible = True
                    Else    '產業人才投資計畫
                        PrintBlank.Visible = True
                        PrintBlank2.Visible = False
                    End If
                Else    '非委訓單位則將兩個列印空白問卷按鈕皆帶出
                    PrintBlank.Visible = True
                    PrintBlank2.Visible = True
                End If
        End Select

#Region "(No Use)"

        'Dim myyear As String = ""
        'Select Case sm.UserInfo.Years
        '    Case Is <= 2007
        '        myyear = "07"
        '        HyperLink1.Visible = True
        '        Hyperlink2.Visible = False
        '    Case Is <= 2011
        '        myyear = "08"
        '        HyperLink1.Visible = False
        '        Hyperlink2.Visible = True
        '    Case Else
        '        myyear = "12"
        '        HyperLink1.Visible = False
        '        Hyperlink2.Visible = True
        'End Select
        'Dim myyear As String = ""
        'Select Case sm.UserInfo.Years
        '    Case Is <= 2007
        '        myyear = "07"
        '    Case Is <= 2011
        '        myyear = "08"
        '    Case Else
        '        myyear = "12"
        'End Select

#End Region

        sPrtFileName1 = Get_SD11004rpt(Me, cst_rptEmp)
        If sPrtFileName1 <> "" Then
            '列印空白表單(產業人才)
            PrintBlank.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, "&OrgKind=01&Years=" & CStr(sm.UserInfo.Years) & "&TPlanID=" & CStr(sm.UserInfo.TPlanID))
            '列印空白表單(在職勞工)
            PrintBlank2.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, "&OrgKind=10&Years=" & CStr(sm.UserInfo.Years) & "&TPlanID=" & CStr(sm.UserInfo.TPlanID))
        End If
        '**by Milor 20040418----end
    End Sub

    '查詢[SQL]
    Sub Search1()
        If objconn Is Nothing Then objconn = DbAccess.GetConnection()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Me.PanelDataGrid1.Visible = True
        Me.PanelDG_stud.Visible = False
        StudentTable.Style.Item("display") = "none"
        msg2.Visible = False
        Label1.Visible = False

        'Dim sqlstr_class As String = ""
        'sqlstr_class = "" & vbCrLf
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( " & vbCrLf
        sql &= "   SELECT a.OCID ,a.ClassCName ,a.CyclType ,a.LevelType ,a.STDate ,a.FTDate " & vbCrLf
        sql &= "   FROM CLASS_CLASSINFO a" & vbCrLf
        sql &= "   JOIN ID_PLAN ip on ip.planid = a.planid " & vbCrLf
        sql &= "   WHERE 1=1 " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        If RIDValue.Value <> "" Then sql &= " AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND a.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        sql &= " ) " & vbCrLf

        sql &= " ,WC2 AS ( " & vbCrLf
        sql &= "   SELECT cs.OCID " & vbCrLf
        sql &= "   ,COUNT(CASE WHEN cs.STUDSTATUS NOT IN (2,3) THEN 1 END) total " & vbCrLf
        sql &= "   ,COUNT(CASE WHEN f.socid IS NOT NULL THEN 1 END) num1 " & vbCrLf
        sql &= "   ,COUNT(CASE WHEN f2.socid IS NOT NULL THEN 1 END) num2 " & vbCrLf
        sql &= "   FROM WC1 cc " & vbCrLf
        sql &= "   JOIN CLASS_STUDENTSOFCLASS cs ON cs.ocid = cc.ocid " & vbCrLf
        sql &= "   LEFT JOIN STUD_QUESTIONFAC f ON f.socid = cs.socid " & vbCrLf
        sql &= "   LEFT JOIN STUD_QUESTIONFAC2 f2 ON f2.socid = cs.socid " & vbCrLf
        sql &= "   WHERE 1=1 " & vbCrLf
        sql &= "   GROUP BY cs.OCID " & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT cc.OCID ,cc.ClassCName ,cc.CyclType " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,cc.LevelType ,cc.STDate ,cc.FTDate ,ss.total " & vbCrLf
        sql &= " ,ISNULL(ss.num2,0) + ISNULL(ss.num1,0) num1 " & vbCrLf
        sql &= " FROM WC1 cc " & vbCrLf
        sql &= " JOIN WC2 ss ON cc.ocid = ss.ocid " & vbCrLf
        sql &= " ORDER BY cc.OCID "

        Dim dtC As DataTable
        dtC = DbAccess.GetDataTable(sql, objconn)
        If dtC.Rows.Count = 0 Then
            msg.Text = "查無資料!!"
            DataGrid1.Style.Item("display") = "none"
            DataGrid1.Visible = False
            Exit Sub
        End If
        msg.Text = ""
        DataGrid1.Style.Item("display") = ""
        DataGrid1.Visible = True
        DataGrid1.DataKeyField = "OCID"
        DataGrid1.DataSource = dtC
        DataGrid1.DataBind()
    End Sub

    '查詢
    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        Call Search1()
    End Sub

    '保留Session
    Sub KeepSearch()
        Dim strSch1 As String = ""
        strSch1 = "center=" & center.Text
        strSch1 &= "&RIDValue=" & RIDValue.Value
        strSch1 &= "&TMID1=" & TMID1.Text
        strSch1 &= "&OCID1=" & OCID1.Text
        strSch1 &= "&TMIDValue1=" & TMIDValue1.Value
        strSch1 &= "&OCIDValue1=" & OCIDValue1.Value
        strSch1 &= "&Button1=" & DG_stud.Visible
        strSch1 &= "&StudentTable=" & StudentTable.Style.Item("display")
        Session(ss_QuestionFacSearchStr) = strSch1
    End Sub

    '查詢學員LIST[SQL]
    Sub Search2(ByVal sOCIDValue1 As String)
        'Dim sqlstr_stud As String = ""
        'Dim sOCIDValue1 As String = e.CommandArgument
        sOCIDValue1 = TIMS.ClearSQM(sOCIDValue1)
        Dim drCC As DataRow = TIMS.GetOCIDDate(sOCIDValue1, objconn)
        If drCC Is Nothing Then Exit Sub

        Label1.Text = "班別：" & Convert.ToString(drCC("ClassCName2"))
        If Not IsDBNull(drCC("LevelType")) Then
            If CInt(drCC("LevelType")) <> 0 Then Label1.Text &= "第" & TIMS.GetChtNum(CInt(drCC("LevelType"))) & "階段"
        End If

        Dim sql As String = ""
        'SELECT * FROM V_STUDENTCOUNT WHERE rownum <=10
        sql = " SELECT * FROM V_STUDENTCOUNT WHERE ocid =" & sOCIDValue1
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then Exit Sub
        Label1.Text &= "(開訓人數:" & dr("opencount").ToString & "&nbsp;&nbsp;在結訓人數:" & dr("TrainCount").ToString & "&nbsp;&nbsp;離退訓人數:" & dr("LeaveCount").ToString & ")"

        'Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.AppliedResultM ,b.studentid ,b.StudStatus " & vbCrLf
        sql &= " ,dbo.fn_CSTUDID2(b.studentid) STUDID2" & vbCrLf
        sql &= " ,c.name ,b.OCID ,b.SOCID " & vbCrLf
        sql &= " ,CONVERT(varchar, b.RejectTDate1, 111) RejectTDate1 " & vbCrLf
        sql &= " ,CONVERT(varchar, b.RejectTDate2, 111) RejectTDate2 " & vbCrLf
        sql &= " ,ISNULL(d2.SOCID, d.SOCID) facSOCID " & vbCrLf
        sql &= " ,ISNULL(d2.DaSource, d.DaSource) DaSource " & vbCrLf
        sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.ocid = b.ocid " & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO c ON b.sid = c.sid " & vbCrLf
        sql &= " LEFT JOIN STUD_QUESTIONFAC d ON d.SOCID = b.SOCID " & vbCrLf
        sql &= " LEFT JOIN STUD_QUESTIONFAC2 d2 ON d2.SOCID = b.SOCID " & vbCrLf
        sql &= " WHERE a.OCID = '" & sOCIDValue1 & "' " & vbCrLf
        sql &= " ORDER BY b.studentid " & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Me.PanelDataGrid1.Visible = True
            Me.PanelDG_stud.Visible = False
            msg2.Text = "查無此班學生資料!"
            StudentTable.Style.Item("display") = "none"
            Label1.Visible = False
            Exit Sub
        End If

        Me.PanelDataGrid1.Visible = False
        Me.PanelDG_stud.Visible = True
        msg2.Text = ""
        StudentTable.Style.Item("display") = "" '"inline"
        Label1.Visible = True
        'Session("DTable_Stuednt") = dt
        'DG_stud.DataKeyField = "SOCID"
        DG_stud.DataSource = dt
        DG_stud.DataBind()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Select Case e.CommandName
            Case cst_ptView
                Dim sOCIDValue1 As String = e.CommandArgument
                Call Search2(sOCIDValue1)

            Case cst_ptPrint '列印空白調查表
                Dim sOCIDValue1 As String = e.CommandArgument
                sOCIDValue1 = TIMS.ClearSQM(sOCIDValue1)
                Dim drCC As DataRow = TIMS.GetOCIDDate(sOCIDValue1, objconn)
                If drCC Is Nothing Then Exit Sub

                '受訓學員意見調查表 (空白調查表)
                sPrtFileName1 = Get_SD11004rpt(Me, cst_rptN)
                If sPrtFileName1 = "" Then
                    Common.MessageBox(Me, "查無列印報表!!")
                    Exit Sub
                End If
                Dim iYears As Integer = sm.UserInfo.Years - 1911
                Dim myValue As String = "&Years=" & CStr(iYears) & "&OCID=" & sOCIDValue1
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrtFileName1, myValue)

                '列印調查表
                'mybut2.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_n" & myyear, "&Years=" & Years & "&OCID=" & OCID)
                '**by Milor 20070421----end
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim MYBUT1 As Button = e.Item.FindControl("Button1") '查詢
                Dim Printblank3 As Button = e.Item.FindControl("Printblank3")  '列印空白
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                MYBUT1.CommandName = cst_ptView
                MYBUT1.Enabled = True
                'If check_search.Value <> "1" Then
                '    MYBUT1.Enabled = False
                '    TIMS.Tooltip(MYBUT1, "無此權限")
                'End If
                MYBUT1.CommandArgument = Convert.ToString(drv("OCID")) 'DataGrid1.DataKeys(e.Item.ItemIndex)
                Printblank3.CommandArgument = Convert.ToString(drv("OCID"))
        End Select
    End Sub

    Private Sub DG_stud_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_stud.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim rSOCID As String = TIMS.GetMyValue(sCmdArg, "socid")
        Dim rOCID As String = TIMS.GetMyValue(sCmdArg, "ocid")
        Dim rDaSource As String = TIMS.GetMyValue(sCmdArg, "DaSource")

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        If rSOCID = "" Then Exit Sub
        If rOCID = "" Then Exit Sub
        '**by Milor 20080414----start
        'Dim sqlstr1 As String = ""
        'Dim okind As String = ""
        'Dim cyear As String = ""
        '將計畫方案分類----start
        Dim sqlstr1 As String = ""
        Dim okind As String = ""
        sqlstr1 = " SELECT ORGKIND FROM VIEW_RIDNAME WHERE RID = '" & RIDValue.Value & "' "
        Dim row_list2 As DataRow = DbAccess.GetOneRow(sqlstr1, objconn)
        If Not row_list2 Is Nothing Then
            If row_list2("OrgKind") <> "10" Then
                okind = "G" '產業人才投資方案
            Else
                okind = "W" '提升勞工自主學習計畫
            End If
        End If
        '將計畫方案分類----end

        sPrtFileName1 = Get_SD11004rpt(Me, cst_Addaspx)
        If sPrtFileName1 = "" Then
            Common.MessageBox(Me, "問卷參數有誤!!查無問卷!!(請連絡系統管理者)")
            Exit Sub
        End If

        'Request("ID")
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim myUrl1 As String = ""
        Select Case e.CommandName
            Case cst_ptInsert
                Call KeepSearch()
                TIMS.CloseDbConn(objconn)
                myUrl1 = "ProcessType=" & cst_ptInsert & "&SOCID=" & rSOCID & "&ocid=" & rOCID & "&ID=" & rqMID & "&orgkind=" & okind
                'Dim s_MYURL2 As String = myUrl1 & "&TM=" & TIMS.GetDateNo3()
                'Dim rMYURL1_H As String = TIMS.EncryptAes(s_MYURL2)
                'myUrl1 &= "&TK1=" & rMYURL1_H
                TIMS.Utl_Redirect1(Me, sPrtFileName1 & myUrl1)

            Case cst_ptClear '清除重填。
                Call KeepSearch()
                'myUrl1 = sPrtFileName1
                myUrl1 = "ProcessType=" & cst_ptDelete & "&SOCID=" & rSOCID & "&ocid=" & rOCID & "&ID=" & rqMID & "&orgkind=" & okind
                'Dim s_MYURL2 As String = myUrl1 & "&TM=" & TIMS.GetDateNo3()
                'Dim rMYURL1_H As String = TIMS.EncryptAes(s_MYURL2)
                'myUrl1 &= "&TK1=" & rMYURL1_H
                'Response.Redirect(sPrtFileName1 & myUrl1)
                Dim sConfirm1 As String = "此動作會刪除受訓學員意見調查表資料，是否確定刪除?"
                Dim strScript As String = ""
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "if (window.confirm('" & sConfirm1 & "')){" + vbCrLf
                strScript += "location.href ='" & sPrtFileName1 & myUrl1 & "';}" + vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript(TIMS.xBlockName, strScript)

            Case cst_ptCheck
                Call KeepSearch()
                TIMS.CloseDbConn(objconn)
                myUrl1 = "ProcessType=" & cst_ptCheck & "&SOCID=" & rSOCID & "&ocid=" & rOCID & "&ID=" & rqMID & "&orgkind=" & okind
                'Dim s_MYURL2 As String = myUrl1 & "&TM=" & TIMS.GetDateNo3()
                'Dim rMYURL1_H As String = TIMS.EncryptAes(s_MYURL2)
                'myUrl1 &= "&TK1=" & rMYURL1_H
                TIMS.Utl_Redirect1(Me, sPrtFileName1 & myUrl1)

            Case cst_ptEdit
                Call KeepSearch()
                TIMS.CloseDbConn(objconn)
                myUrl1 = "ProcessType=" & cst_ptEdit & "&SOCID=" & rSOCID & "&ocid=" & rOCID & "&ID=" & rqMID & "&orgkind=" & okind
                'Dim s_MYURL2 As String = myUrl1 & "&TM=" & TIMS.GetDateNo3()
                'Dim rMYURL1_H As String = TIMS.EncryptAes(s_MYURL2)
                'myUrl1 &= "&TK1=" & rMYURL1_H
                TIMS.Utl_Redirect1(Me, sPrtFileName1 & myUrl1)

            Case cst_ptPrint
                sPrtFileName1 = Get_SD11004rpt(Me, cst_rptData)
                If sPrtFileName1 = "" Then
                    Common.MessageBox(Me, "參數有誤!!查無列印報表!!(請連絡系統管理者)")
                    Exit Sub
                End If
                If rDaSource <> "1" Then
                    Dim iYears As Integer = sm.UserInfo.Years - 1911
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrtFileName1, "&Years=" & CStr(iYears) & "&SOCID=" & rSOCID)
                    'Select Case Val(sm.UserInfo.Years)
                    '    Case Is <= 2007
                    '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004", "&Years=" & Years & "&SOCID=" & drv("SOCID")) ' e.Item.Cells(6).Text)
                    '    Case Is <= 2011
                    '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_08", "&Years=" & Years & "&SOCID=" & drv("SOCID")) 'e.Item.Cells(6).Text)
                    '    Case Else
                    '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_12", "&Years=" & Years & "&SOCID=" & drv("SOCID")) 'e.Item.Cells(6).Text)
                    'End Select
                    'TIMS.Tooltip(Print, "非學員自行填寫，可列印")
                Else
                    'If Convert.ToString(drv("DaSource")) = "1" Then
                    'If sm.UserInfo.LID = 2 Then
                    Select Case Convert.ToString(sm.UserInfo.LID)
                        Case "2"
                        Case Else
                            Dim iYears As Integer = sm.UserInfo.Years - 1911
                            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrtFileName1, "&Years=" & CStr(iYears) & "&SOCID=" & rSOCID)
                    End Select
                End If
        End Select
    End Sub

    '含有一些限制規則。
    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim but4 As Button = e.Item.FindControl("Button4") '新增
                Dim Edit As Button = e.Item.FindControl("Edit")    '修改
                Dim but5 As Button = e.Item.FindControl("Button5") '查看
                Dim Print As Button = e.Item.FindControl("Print")  '列印
                Dim but6 As Button = e.Item.FindControl("Button6") '清除重填
                but4.CommandName = cst_ptInsert '新增
                Edit.CommandName = cst_ptEdit '修改
                but5.CommandName = cst_ptCheck '查看
                Print.CommandName = cst_ptPrint '列印
                but6.CommandName = cst_ptClear '清除重填

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "socid", Convert.ToString(drv("socid")))
                TIMS.SetMyValue(sCmdArg, "ocid", Convert.ToString(drv("ocid")))
                TIMS.SetMyValue(sCmdArg, "DaSource", Convert.ToString(drv("DaSource")))

                but4.CommandArgument = sCmdArg
                Edit.CommandArgument = sCmdArg
                but5.CommandArgument = sCmdArg
                Print.CommandArgument = sCmdArg
                but6.CommandArgument = sCmdArg

                Dim sName As String = Convert.ToString(drv("name"))
                If Convert.ToString(drv("RejectTDate1")) <> "" Then sName &= "(" & Convert.ToString(drv("RejectTDate1")) & ")"
                If Convert.ToString(drv("RejectTDate2")) <> "" Then sName &= "(" & Convert.ToString(drv("RejectTDate2")) & ")"
                e.Item.Cells(1).Text = sName

                If Convert.ToString(drv("facSOCID")) <> "" Then
                    '已有資料
                    e.Item.Cells(2).Text = "是"
                    but4.Enabled = False '不可新增
                    TIMS.Tooltip(but4, "已有資料不可新增")
                    Print.Enabled = True '可以列印
                    Edit.Enabled = True '可以修改
                    but5.Enabled = True '可以查看
                    but6.Enabled = True '可清除重填
                    'If check_mod.Value = "0" Then
                    '    Edit.Enabled = False
                    '    TIMS.Tooltip(Edit, "無此權限，不可修改")
                    'End If
                    'If check_search.Value = "0" Then
                    '    but5.Enabled = False
                    '    TIMS.Tooltip(but5, "無此權限，不可查看")
                    'End If
                    'If check_mod.Value = "0" And check_del.Value = "0" Then '兩者功能皆沒有時,不能使用 清除重填
                    '    but6.Enabled = False
                    '    TIMS.Tooltip(but6, "無修改刪除權限，不可重填")
                    'End If
                Else
                    '無資料
                    e.Item.Cells(2).Text = "否"
                    but4.Enabled = True '可新增
                    Edit.Enabled = False '不可修改
                    but5.Enabled = False '不可查看
                    but6.Enabled = False '不可清除重填
                    Print.Enabled = False '不可以列印
                    TIMS.Tooltip(Edit, "沒有資料，不可修改")
                    TIMS.Tooltip(but5, "沒有資料，不可查看")
                    TIMS.Tooltip(but6, "沒有資料，不可清除重填")
                    TIMS.Tooltip(Print, "沒有資料，不可以列印")
                    'If check_add.Value = "0" Then
                    '    but4.Enabled = False
                    '    TIMS.Tooltip(but4, "無此權限，不可新增")
                    'End If
                End If

                If Convert.ToString(drv("AppliedResultM")) = "Y" Then
                    but4.Enabled = False '不可新增
                    Edit.Enabled = False '不可修改
                    but6.Enabled = False '不可清除重填
                    TIMS.Tooltip(but4, "審核結果已通過，不可新增")
                    TIMS.Tooltip(Edit, "審核結果已通過，不可修改")
                    TIMS.Tooltip(but6, "審核結果已通過，不可清除重填")
                End If

                'Dim Years As Integer = sm.UserInfo.Years - 1911
                '依據年度不同，呼叫不同的已填妥問卷
                '學員問卷列印
                'Dim myyear As String = ""

                '單位填寫'DaSource 資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                '學員填寫不可列印。
                If Convert.ToString(drv("DaSource")) <> "1" Then
                    'Select Case Val(sm.UserInfo.Years)
                    '    Case Is <= 2007
                    '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004", "&Years=" & Years & "&SOCID=" & drv("SOCID")) ' e.Item.Cells(6).Text)
                    '    Case Is <= 2011
                    '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_08", "&Years=" & Years & "&SOCID=" & drv("SOCID")) 'e.Item.Cells(6).Text)
                    '    Case Else
                    '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_12", "&Years=" & Years & "&SOCID=" & drv("SOCID")) 'e.Item.Cells(6).Text)
                    'End Select
                    TIMS.Tooltip(Print, "非學員自行填寫，可列印")
                End If

                '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
                '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
                '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
                '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。
                'PS: 訓練單位登入,若是學員於外網填寫,是不可修改,不可查詢,不可列印,不可清除重填   但若是分署登入,若是學員於外網填寫,是不可修改,可查詢,可列印,不可清除重填
                If Convert.ToString(drv("DaSource")) = "1" Then
                    If sm.UserInfo.LID = 2 Then
                        '委訓單位
                        but4.Enabled = False '不可新增
                        Edit.Enabled = False '不可修改
                        but5.Enabled = False '不可查看
                        but6.Enabled = False '不可清除重填
                        Print.Enabled = False '不可以列印
                        TIMS.Tooltip(but4, "學員外網填寫，不可新增")
                        TIMS.Tooltip(Edit, "學員外網填寫，不可修改")
                        TIMS.Tooltip(but5, "學員外網填寫，不可查看")
                        TIMS.Tooltip(but6, "學員外網填寫，不可清除重填")
                        TIMS.Tooltip(Print, "學員外網填寫，不可以列印")
                        but4.CommandArgument = ""
                        Edit.CommandArgument = ""
                        but5.Enabled = True '不可查看
                        but5.CommandArgument = ""
                        but5.Attributes("onclick") = "alert('學員外網填寫，不可查看');return false;"
                        but6.CommandArgument = ""
                        Print.CommandArgument = ""
                    Else
                        '非委訓單位(署、分署)
                        but4.Enabled = False '不可新增
                        Edit.Enabled = False '不可修改
                        but5.Enabled = True '可查看
                        but6.Enabled = False '不可清除重填
                        Print.Enabled = True '可列印
                        TIMS.Tooltip(but4, "學員外網填寫，不可新增")
                        TIMS.Tooltip(Edit, "學員外網填寫，不可修改")
                        TIMS.Tooltip(but5, "非委訓單位登入，可查看")
                        TIMS.Tooltip(but6, "學員外網填寫，不可清除重填")
                        TIMS.Tooltip(Print, "非委訓單位登入，可列印")
                        but4.CommandArgument = ""
                        Edit.CommandArgument = ""
                        'but5.Enabled = True '不可查看
                        'but5.CommandArgument = ""
                        'but5.Attributes("onclick") = "alert('學員外網填寫，不可查看');return false;"
                        but6.CommandArgument = ""
                        'Print.CommandArgument = ""
                        'Select Case Val(sm.UserInfo.Years)
                        '    Case Is <= 2007
                        '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004", "&Years=" & Years & "&SOCID=" & drv("SOCID")) ' e.Item.Cells(6).Text)
                        '    Case Is <= 2011
                        '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_08", "&Years=" & Years & "&SOCID=" & drv("SOCID")) 'e.Item.Cells(6).Text)
                        '    Case Else
                        '        Print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "BussinessTrain", "SD_11_004_12", "&Years=" & Years & "&SOCID=" & drv("SOCID")) 'e.Item.Cells(6).Text)
                        'End Select
                    End If
                End If
        End Select
    End Sub

#Region "(No Use)"

    ''民國年改為西元年
    'Function ChangeTWDate(ByVal TWDate As String) As String
    '    Return CStr(CInt(Left(TWDate, 3)) + 1911) & "/" & Mid(TWDate, 4, 2) & "/" & Right(TWDate, 2)
    'End Function

    ''匯入調查表
    'Private Sub Button13_Click(ByVal sender As Object, ByVal e As System.EventArgs) 'Handles Button13.Click
    '    Dim Upload_Path As String = "~/SD/11/Temp/"
    '    'Dim I As Integer
    '    Dim MyFile As System.IO.File
    '    Dim FileOCIDValue, MyFileName As String
    '    Dim MyFileType As String
    '    Dim flag As String

    '    If File1.Value <> "" Then
    '        '檢查檔案格式與大小----------   Start
    '        If File1.PostedFile.ContentLength = 0 Then
    '            Common.MessageBox(Me, "檔案位置錯誤!")
    '            Exit Sub
    '        Else
    '            '取出檔案名稱
    '            MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
    '            '取出檔案類型
    '            If MyFileName.IndexOf(".") = -1 Then
    '                Common.MessageBox(Me, "檔案類型錯誤!")
    '                Exit Sub
    '            Else
    '                MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
    '                If LCase(MyFileType) = "csv" Then
    '                    flag = ","
    '                Else
    '                    Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
    '                    Exit Sub
    '                End If
    '            End If
    '        End If
    '        '檢查檔案格式與大小----------   End

    '        '上傳檔案
    '        File1.PostedFile.SaveAs(Server.MapPath(Upload_Path & MyFileName))
    '        'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)

    '        '將檔案讀出放入記憶體
    '        Dim sr As System.IO.Stream
    '        Dim srr As System.IO.StreamReader
    '        sr = MyFile.OpenRead(Server.MapPath(Upload_Path & MyFileName))
    '        srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

    '        Dim RowIndex As Integer = 0 '讀取行累計數
    '        Dim OneRow As String        'srr.ReadLine 一行一行的資料
    '        Dim col As String           '欄位
    '        Dim colArray As Array

    '        '取出資料庫的所有欄位--------   Start
    '        Dim sql As String
    '        Dim dtStuOfClass As DataTable
    '        Dim dr As DataRow
    '        Dim da As SqlDataAdapter

    '        Dim dt As DataTable
    '        Dim STDate As Date
    '        Dim Correct_Dlid As String
    '        Dim Next_Dlid As String
    '        Dim writeflag As Boolean

    '        Dim Reason As String                '儲存錯誤的原因
    '        Dim dtWrong As New DataTable        '儲存錯誤資料的DataTable
    '        Dim drWrong As DataRow

    '        '建立錯誤資料格式Table----------------Start
    '        dtWrong.Columns.Add(New DataColumn("Index"))
    '        dtWrong.Columns.Add(New DataColumn("FillFormDate"))
    '        dtWrong.Columns.Add(New DataColumn("StudID"))
    '        dtWrong.Columns.Add(New DataColumn("Status"))
    '        dtWrong.Columns.Add(New DataColumn("Reason"))
    '        '建立錯誤資料格式Table----------------End

    '        Do While srr.Peek >= 0
    '            OneRow = srr.ReadLine
    '            If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈

    '            If RowIndex <> 0 Then
    '                WriteStatus = True
    '                writeflag = True
    '                Reason = ""
    '                colArray = Split(OneRow, flag)


    '                '建立SOCID欄位值
    '                If OCIDValue1.Value = "" Then
    '                    Reason += "未選擇 開班編號(OCID) 無法匯入<BR>"
    '                    Exit Do
    '                End If

    '                If colArray.Length > 2 Then
    '                    '==補強EXCEL可能去除零值之可能性==Start
    '                    'AMU 2006/12/14
    '                    If Len(colArray(0).ToString) = 6 Then
    '                        colArray(0) = "0" & colArray(0).ToString
    '                    End If
    '                    '==補強EXCEL可能去除零值之可能性==End

    '                    Try
    '                        If Len(colArray(1).ToString) = 0 Then
    '                            Reason += "未輸入 學號(StudID) 無法匯入<BR>"
    '                            WriteStatus = False
    '                            writeflag = False

    '                        ElseIf getSOCID(OCIDValue1.Value, colArray(1).ToString, objconn) = "0" Then
    '                            Reason += "學號 " & colArray(1).ToString & " 無此學員資料無法匯入<BR>"
    '                            WriteStatus = False
    '                            writeflag = False
    '                        End If
    '                    Catch ex As Exception
    '                        Reason += "匯入 學號格式有誤，無法匯入<BR>"
    '                        WriteStatus = False
    '                        writeflag = False
    '                    End Try
    '                Else
    '                    Reason += "匯入格式有誤，無法匯入<BR>"
    '                    WriteStatus = False
    '                    writeflag = False
    '                End If

    '                If Reason = "" Then Reason += CheckImportData(colArray, writeflag) '檢查資料正確性

    '                '通過檢查，開始輸入資料---------------------Start
    '                If Reason = "" Then
    '                    WriteDB(colArray)
    '                Else
    '                    If writeflag Then WriteDB(colArray) '若 writeflag 為true 則可繼續新增到資料庫

    '                    '錯誤資料，填入錯誤資料表
    '                    drWrong = dtWrong.NewRow
    '                    dtWrong.Rows.Add(drWrong)

    '                    drWrong("Index") = RowIndex
    '                    If colArray.Length > 2 Then
    '                        drWrong("FillFormDate") = colArray(0) '填寫日期
    '                        drWrong("StudID") = colArray(1)       '學號
    '                        If WriteStatus = False Then drWrong("Status") = "未匯入" Else drWrong("Status") = "匯入"
    '                        drWrong("Reason") = Reason & IIf(WriteStatus = True, " 錯誤答案部份不寫入資料中", "")
    '                    Else
    '                        drWrong("Reason") = Reason
    '                    End If
    '                End If
    '            End If
    '            RowIndex += 1 '讀取行累計數
    '        Loop

    '        sr.Close()
    '        srr.Close()
    '        MyFile.Delete(Server.MapPath(Upload_Path & MyFileName))

    '        '開始判別欄位存入------------   End
    '        If Me.ViewState("WriteDBmsg") <> "" Then
    '            Common.MessageBox(Me, Me.ViewState("WriteDBmsg"))
    '        End If

    '        If dtWrong.Rows.Count = 0 Then
    '            If Reason = "" Then
    '                Common.MessageBox(Me, "資料匯入成功")
    '            Else
    '                Common.MessageBox(Me, Reason)
    '            End If
    '        Else

    '            'Session("MyWrongTable") = dtWrong
    '            Datagrid2.Style.Item("display") = "inline"
    '            Datagrid2.Visible = True
    '            Datagrid2.DataSource = dtWrong
    '            Datagrid2.DataBind()
    '            Common.MessageBox(Me, "資料匯入成功,但有錯誤資料請檢示原因!!!")

    '            For I As Integer = 1 To 100
    '                If I = 100 Then eMeng.Style.Item("display") = "inline"
    '                'Page.RegisterStartupScript("", "<script>{window.document.getElementById('eMeng').style.visibility='visible';}</script>")
    '            Next
    '        End If

    '    End If
    '    Call search_Click(sender, e)
    'End Sub

    ''匯入檢查資料正確性
    'Function CheckImportData(ByVal colArray As Array, ByRef writeflag As Boolean) '檢查資料正確性
    '    'amu 20061221 因為同意可寫入某些錯誤的資料，但還是要show訊息

    '    Dim Reason As String = ""
    '    Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
    '    Dim sql As String = ""
    '    Dim Num3 As String = "123"
    '    Dim Num4 As String = "1234"
    '    Dim Num5 As String = "12345"
    '    Dim Num6 As String = "123456"
    '    Dim Q1_1, Q1_2, Q1_3, Q1_4, Q1_5, Q1_6 As String
    '    Dim Q1_6_CCourName, Q1_6_CHour, Q1_6_MCourName, Q1_6_MHour As String
    '    Dim Q2_1, Q2_2, Q2_3, Q2_4, Q2_5 As String
    '    Dim Q2_1_CCourName, Q2_1_CHour, Q2_1_MCourName, Q2_1_MHour As String
    '    Dim Q3_1, Q3_2, Q3_3 As String
    '    Dim Q4, Q5, Q6, Q7, Q8, Q9 As String
    '    Dim Q7_8, Q7_9 As String

    '    Dim Q5_Note_News, Q5_Note_Other As String
    '    Dim Q6_Note1, Q6_Note2 As String
    '    Dim Q9_1, Q9_2, Q9_3, Q10, Q11 As String
    '    Dim dr As DataRow
    '    Dim cst_len As Integer

    '    Select Case sm.UserInfo.Years
    '        Case Is <= 2007 '96年度項目數
    '            cst_len = 35
    '        Case Is <= 2011 '97年度項目數
    '            cst_len = 39
    '        Case Else '101年度項目數
    '            cst_len = 41
    '    End Select

    '    If colArray.Length < cst_len Then
    '        'Reason += "欄位數量不正確(應該為" & cst_len & "個欄位)<BR>"
    '        Reason += "欄位對應有誤<BR>"
    '        Reason += "請注意欄位中是否有半形逗點<BR>"
    '    End If

    '    Try
    '        FillFormDate = colArray(0).ToString     '填寫日期
    '        SOCID = getSOCID(OCIDValue1.Value, colArray(1).ToString, objconn) '學號

    '        Q1_1 = colArray(2).ToString             '課程內容與工作性質是否相關
    '        Q1_2 = colArray(3).ToString             '課程名稱是否適當
    '        Q1_3 = colArray(4).ToString             '教材內容是否適當
    '        Q1_4 = colArray(5).ToString             '本項訓練發給教材情形
    '        Q1_5 = colArray(6).ToString             '發給方式
    '        Q1_6 = colArray(7).ToString             '訓練時數是否適當 
    '        Q1_6_CCourName = colArray(8).ToString   '應增加課程名稱
    '        Q1_6_CHour = colArray(9).ToString       '應增加 小時數
    '        Q1_6_MCourName = colArray(10).ToString  '應減少課程名稱
    '        Q1_6_MHour = colArray(11).ToString      '應減少 小時數
    '        Q2_1 = colArray(12).ToString            '術科時數是否適當
    '        Q2_1_CCourName = colArray(13).ToString  '應增加課程名稱
    '        Q2_1_CHour = colArray(14).ToString      '應增加 小時數
    '        Q2_1_MCourName = colArray(15).ToString  '應減少課程名稱
    '        Q2_1_MHour = colArray(16).ToString      '應減少 小時數
    '        Q2_2 = colArray(17).ToString            '術科內容是否適當
    '        Q2_3 = colArray(18).ToString            '術科操作解說是否充分
    '        Q2_4 = colArray(19).ToString            '訓練設備是否充足
    '        Q2_5 = colArray(20).ToString            '訓練設備現狀
    '        Q3_1 = colArray(21).ToString            '教師的教學態度
    '        Q3_2 = colArray(22).ToString            '教師師資的教學方法或技巧
    '        Q3_3 = colArray(23).ToString            '講授課程時間控制是否適當
    '        Q4 = colArray(24).ToString              '你對整體課程瞭解的程度
    '        Q5 = colArray(25).ToString              '你獲得招訓消息的來源為
    '        Q5_Note_News = colArray(26).ToString    '報紙名稱
    '        Q5_Note_Other = colArray(27).ToString   '其他來源名稱
    '        Q6 = colArray(28).ToString              '除政府補助外，自行繳納費用負擔方式
    '        Q6_Note1 = colArray(29).ToString        '自行負擔金額
    '        Q6_Note2 = colArray(30).ToString        '服務單位負擔
    '        Q7 = colArray(31).ToString              '96年->參加本項訓練後，對就業安定工作是否有幫助；97年->參加本訓練後，你對於訓練課程所教授知識或技能的掌握程度。
    '        '**by Milor 20070421----start
    '        Select Case sm.UserInfo.Years
    '            Case Is <= 2007 '96年度項目數
    '                Q8 = colArray(32).ToString              '你對訓練單位的行政服務滿意度
    '                Q9 = colArray(33).ToString      '整體而言，你對於參加本計畫訓練的滿意度為：
    '            Case Is <= 2011 '97年度項目數
    '                Q8 = colArray(32).ToString              '你對訓練單位的行政服務滿意度
    '                Q9_1 = colArray(33).ToString    '你對產業人才投資計畫規劃設計的瞭解程度->補助對象
    '                Q9_2 = colArray(34).ToString    '補助經費標準
    '                Q9_3 = colArray(35).ToString    '補助流程
    '                Q10 = colArray(36).ToString     '整體而言，你對於產業人才投資計畫是否滿意
    '                Q11 = colArray(37).ToString     '若無補助訓練經費，你每年願意以自費方式參加相關訓練課程之金額?元
    '            Case Else '101年度項目數
    '                Q7_8 = colArray(32).ToString              '(八)參加本項課程訓練後，你有把握自己能所學的知識應用到工作上
    '                Q7_9 = colArray(33).ToString              '(九)完成訓練後，你願意找機會將所學的知識／技能應用在工作中
    '                Q8 = colArray(34).ToString              '你對訓練單位的行政服務滿意度
    '                Q9_1 = colArray(35).ToString    '你對產業人才投資計畫規劃設計的瞭解程度->補助對象
    '                Q9_2 = colArray(36).ToString    '補助經費標準
    '                Q9_3 = colArray(37).ToString    '補助流程
    '                Q10 = colArray(38).ToString     '整體而言，你對於產業人才投資計畫是否滿意
    '                Q11 = colArray(39).ToString     '若無補助訓練經費，你每年願意以自費方式參加相關訓練課程之金額?元
    '        End Select
    '        '**by Milor 20070421----end
    '    Catch ex As Exception
    '        Reason += "欄位對應有誤<BR>"
    '        Reason += "請注意欄位中是否有半形逗點<BR>"
    '    End Try

    '    '欄位若完全塞入，再進行下列檢核功能，反之宣告有誤
    '    If Reason = "" Then
    '        '讀卡日期
    '        If FillFormDate = "" Or Len(FillFormDate) <> 7 Or IsNumeric(FillFormDate) <> True Then
    '            Reason += "填寫日期有誤，必須是民國年格式(yyymmdd)<BR>"
    '            writeflag = False
    '            WriteStatus = False
    '        Else
    '            'FillFormDate = CStr(CInt(Left(FillFormDate, 3)) + 1911) & "/" & Mid(FillFormDate, 4, 2) & "/" & Right(FillFormDate, 2)
    '            FillFormDate = ChangeTWDate(FillFormDate)
    '            If IsDate(FillFormDate) = False Then
    '                Reason += "填寫日期有誤，必須是民國年格式(yyymmdd)<BR>"
    '                writeflag = False
    '                WriteStatus = False
    '            Else
    '                If CDate(FillFormDate) < "1900/1/1" Or CDate(FillFormDate) > "2100/1/1" Then
    '                    Reason += "填寫日期範圍有誤<BR>"
    '                    writeflag = False
    '                    WriteStatus = False
    '                End If
    '            End If
    '        End If

    '        If Q1_1 <> "" Then
    '            If Len(Q1_1) <> 1 Then Reason += "1.1題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q1_1) = 0 Then Reason += "1.1題答案應在1-5之間<BR>"
    '        End If
    '        If Q1_2 <> "" Then
    '            If Len(Q1_2) <> 1 Then Reason += "1.2題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q1_2) = 0 Then Reason += "1.2題答案應在1-5之間<BR>"
    '        End If
    '        If Q1_3 <> "" Then
    '            If Len(Q1_3) <> 1 Then Reason += "1.3題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q1_3) = 0 Then Reason += "1.3題答案應在1-5之間<BR>"
    '        End If
    '        If Q1_4 <> "" Then
    '            If Len(Q1_4) <> 1 Then Reason += "1.4題答案應在1-4之間<BR>"
    '            If InStr(Num4, Q1_4) = 0 Then Reason += "1.4題答案應在1-4之間<BR>"
    '        End If
    '        If Q1_5 <> "" Then
    '            If Len(Q1_5) <> 1 Then Reason += "1.5題答案應在1-3之間<BR>"
    '            If InStr(Num3, Q1_5) = 0 Then Reason += "1.5題答案應在1-3之間<BR>"
    '        End If
    '        If Q1_6 <> "" Then
    '            If Len(Q1_6) <> 1 Then Reason += "1.6題答案應在1-3之間<BR>"
    '            If InStr(Num3, Q1_6) = 0 Then Reason += "1.6題答案應在1-3之間<BR>"
    '        End If
    '        If Q1_6 = "2" Then
    '            If Q1_6_CCourName = "" Then Reason += "必須填寫應增加課程名稱<BR>"
    '            If Q1_6_CHour = "" Then
    '                Reason += "必須填寫應增加小時數<BR>"
    '            Else
    '                If Not IsNumeric(Q1_6_CHour) Then
    '                    Reason += "應增加小時數，必須為數字格式<BR>"
    '                    writeflag = False
    '                End If
    '            End If
    '        End If
    '        If Q1_6 = "3" Then
    '            If Q1_6_MCourName = "" Then Reason += "必須填寫應減少課程名稱<BR>"
    '            If Q1_6_MHour = "" Then
    '                Reason += "必須填寫應減少小時數<BR>"
    '            Else
    '                If Not IsNumeric(Q1_6_MHour) Then
    '                    Reason += "應減少小時數，必須為數字格式<BR>"
    '                    writeflag = False
    '                End If
    '            End If
    '        End If
    '        If Q2_1 <> "" Then
    '            If Len(Q2_1) <> 1 Then Reason += "2.1題答案應在1-3之間<BR>"
    '            If InStr(Num3, Q2_1) = 0 Then Reason += "2.1題答案應在1-3之間<BR>"
    '        End If
    '        If Q2_1 = "2" Then
    '            If Q2_1_CCourName = "" Then Reason += "必須填寫應增加術科課程名稱<BR>"
    '            If Q2_1_CHour = "" Then
    '                Reason += "必須填寫應增加術科小時數<BR>"
    '            Else
    '                If Not IsNumeric(Q2_1_CHour) Then
    '                    Reason += " 應增加術科小時數，必須為數字格式<BR>"
    '                    writeflag = False
    '                End If
    '            End If
    '        End If
    '        If Q2_1 = "3" Then
    '            If Q2_1_MCourName = "" Then Reason += "必須填寫應減少術科課程名稱<BR>"
    '            If Q2_1_MHour = "" Then
    '                Reason += "必須填寫應減少術科小時數<BR>"
    '            Else
    '                If Not IsNumeric(Q2_1_MHour) Then
    '                    Reason += " 應減少術科小時數，必須為數字格式<BR>"
    '                    writeflag = False
    '                End If
    '            End If
    '        End If
    '        If Q2_2 <> "" Then
    '            If Len(Q2_2) <> 1 Then Reason += "2.2題答案應在1-4之間<BR>"
    '            If InStr(Num4, Q2_2) = 0 Then Reason += "2.2題答案應在1-4之間<BR>"
    '        End If
    '        If Q2_3 <> "" Then
    '            If Len(Q2_3) <> 1 Then Reason += "2.3題答案應在1-4之間<BR>"
    '            If InStr(Num4, Q2_3) = 0 Then Reason += "2.3題答案應在1-4之間<BR>"
    '        End If
    '        If Q2_4 <> "" Then
    '            If Len(Q2_4) <> 1 Then Reason += "2.4題答案應在1-4之間<BR>"
    '            If InStr(Num4, Q2_4) = 0 Then Reason += "2.4題答案應在1-4之間<BR>"
    '        End If
    '        If Q2_5 <> "" Then
    '            If Len(Q2_5) <> 1 Then Reason += "2.5題答案應在1-4之間<BR>"
    '            If InStr(Num4, Q2_5) = 0 Then Reason += "2.5題答案應在1-4之間<BR>"
    '        End If
    '        If Q3_1 <> "" Then
    '            If Len(Q3_1) <> 1 Then Reason += "3.1題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q3_1) = 0 Then Reason += "3.1題答案應在1-5之間<BR>"
    '        End If
    '        If Q3_2 <> "" Then
    '            If Len(Q3_2) <> 1 Then Reason += "3.2題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q3_2) = 0 Then Reason += "3.2題答案應在1-5之間<BR>"
    '        End If
    '        If Q3_3 <> "" Then
    '            If Len(Q3_3) <> 1 Then Reason += "3.3題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q3_3) = 0 Then Reason += "3.3題答案應在1-5之間<BR>"
    '        End If
    '        If Q4 <> "" Then
    '            If Len(Q4) <> 1 Then Reason += "4 題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q4) = 0 Then Reason += "4 題答案應在1-5之間<BR>"
    '        End If
    '        If Q5 <> "" Then
    '            If Len(Q5) <> 1 Then Reason += "5 題答案應在1-6之間<BR>"
    '            If InStr(Num6, Q5) = 0 Then Reason += "5 題答案應在1-6之間<BR>"
    '        End If
    '        If Q5 = "4" Then
    '            If Q5_Note_News = "" Then Reason += "必須填寫報紙名稱<BR>"
    '        End If
    '        If Q5 = "6" Then
    '            If Q5_Note_Other = "" Then Reason += "必須填寫其他來源名稱<BR>"
    '        End If
    '        If Q6 <> "" Then
    '            If Len(Q6) <> 1 Then Reason += "6 題答案應在1-3之間<BR>"
    '            If InStr(Num3, Q6) = 0 Then Reason += "6 題答案應在1-3之間<BR>"
    '        End If
    '        If Q6 = "1" Then
    '            If Q6_Note1 = "" Then Reason += "必須填寫自行負擔金額<BR>"
    '        End If
    '        If Q6 = "2" Then
    '            If Q6_Note2 = "" Then Reason += "必須填寫服務單位負擔<BR>"
    '        End If
    '        If Q6 = "3" Then
    '            If Q6_Note1 = "" Then
    '                Reason += "必須填寫自行負擔金額<BR>"
    '            Else
    '                If Not IsNumeric(Q6_Note1) Then
    '                    Reason += "自行負擔金額，必須為數字格式<BR>"
    '                    writeflag = False
    '                End If
    '            End If
    '            If Q6_Note2 = "" Then
    '                Reason += "必須填寫服務單位負擔金額<BR>"
    '            Else
    '                If Not IsNumeric(Q6_Note2) Then
    '                    Reason += "服務單位負擔金額，必須為數字格式<BR>"
    '                    writeflag = False
    '                End If
    '            End If
    '        End If
    '        If Q7 <> "" Then
    '            If Len(Q7) <> 1 Then Reason += "7 題答案應在1-5之間<BR>"
    '            If InStr(Num5, Q7) = 0 Then Reason += "7 題答案應在1-5之間<BR>"
    '        End If

    '        '**by Milor 20070421----start
    '        Select Case sm.UserInfo.Years
    '            Case Is <= 2007 '96年版
    '                If Q8 <> "" Then
    '                    If InStr(Num5, Q8) = 0 Then Reason += "8 題答案應在1-5之間<BR>"
    '                End If
    '                If Q9 <> "" Then
    '                    If InStr(Num5, Q9) = 0 Then Reason += "9 題答案應在1-5之間<BR>"
    '                End If
    '            Case Is <= 2011 '97年版
    '                If Q8 <> "" Then
    '                    If InStr(Num5, Q8) = 0 Then Reason += "8 題答案應在1-5之間<BR>"
    '                End If
    '                If Q9_1 <> "" Then
    '                    If InStr(Num5, Q9_1) = 0 Then Reason += "9.1題答案應在1-5間<BR>"
    '                End If
    '                If Q9_2 <> "" Then
    '                    If InStr(Num5, Q9_2) = 0 Then Reason += "9.2題答案應在1-5間<BR>"
    '                End If
    '                If Q9_3 <> "" Then
    '                    If InStr(Num5, Q9_3) = 0 Then Reason += "9.3題答案應在1-5間<BR>"
    '                End If
    '                If Q10 <> "" Then
    '                    If InStr(Num5, Q10) = 0 Then Reason += "10題答案應在1-5間<BR>"
    '                End If
    '                If Q11 = "" Then
    '                    Reason += "11題必須填寫每年願意自費之金額<BR>"
    '                Else
    '                    If Not IsNumeric(Q11) Then
    '                        Reason += "11題 每年願意自費之金額，必須為數字格式<BR>"
    '                        writeflag = False
    '                    End If
    '                End If
    '            Case Else '101年度項目數
    '                If Q7_8 <> "" Then
    '                    If Len(Q7_8) <> 1 Then Reason += "8.答案應在1-5間<BR>"
    '                    If InStr(Num5, Q7_8) = 0 Then Reason += "8.答案應在1-5間<BR>"
    '                End If
    '                If Q7_9 <> "" Then
    '                    If Len(Q7_9) <> 1 Then Reason += "9.答案應在1-5間<BR>"
    '                    If InStr(Num5, Q7_9) = 0 Then Reason += "9.答案應在1-5間<BR>"
    '                End If
    '                If Q8 <> "" Then
    '                    If Len(Q8) <> 1 Then Reason += "10.答案應在1-5之間<BR>"
    '                    If InStr(Num5, Q8) = 0 Then Reason += "10.答案應在1-5之間<BR>"
    '                End If
    '                If Q9_1 <> "" Then
    '                    If Len(Q9_1) <> 1 Then Reason += "11.1題答案應在1-5間<BR>"
    '                    If InStr(Num5, Q9_1) = 0 Then Reason += "11.1題答案應在1-5間<BR>"
    '                End If
    '                If Q9_2 <> "" Then
    '                    If Len(Q9_2) <> 1 Then Reason += "11.2題答案應在1-5間<BR>"
    '                    If InStr(Num5, Q9_2) = 0 Then Reason += "11.2題答案應在1-5間<BR>"
    '                End If
    '                If Q9_3 <> "" Then
    '                    If Len(Q9_3) <> 1 Then Reason += "11.3題答案應在1-5間<BR>"
    '                    If InStr(Num5, Q9_3) = 0 Then Reason += "11.3題答案應在1-5間<BR>"
    '                End If
    '                If Q10 <> "" Then
    '                    If Len(Q10) <> 1 Then Reason += "12題答案應在1-5間<BR>"
    '                    If InStr(Num5, Q10) = 0 Then Reason += "12題答案應在1-5間<BR>"
    '                End If
    '                If Q11 = "" Then
    '                    Reason += "13題必須填寫每年願意自費之金額<BR>"
    '                Else
    '                    If Not IsNumeric(Q11) Then
    '                        Reason += "13題 每年願意自費之金額，必須為數字格式<BR>"
    '                        writeflag = False
    '                    End If
    '                End If
    '        End Select
    '        '**by Milor 20070421----end
    '    End If
    '    Return Reason
    'End Function

    '匯入寫DB
    'Private Sub WriteDB(ByVal colArray As System.Array)
    '    '將資料寫入資料表中
    '    'Dim trans As SqlTransaction
    '    Dim dt As DataTable
    '    Dim sql As String
    '    'Dim dr As DataRow
    '    'Dim i As Integer

    '    FillFormDate = colArray(0).ToString       '填寫日期 
    '    FillFormDate = ChangeTWDate(FillFormDate) '填寫日期 
    '    SOCID = getSOCID(OCIDValue1.Value, colArray(1).ToString, objconn)  '學號 

    '    Me.ViewState("WriteDBmsg") = ""

    '    Dim tConn As SqlConnection = DbAccess.GetConnection()
    '    Call TIMS.OpenDbConn(tConn)

    '    Dim dr As DataRow = Nothing
    '    sql = " SELECT * FROM Stud_QuestionFac WHERE SOCID ='" & SOCID & "' "
    '    dt = DbAccess.GetDataTable(sql, tConn)
    '    If dt.Rows.Count > 0 Then
    '        dr = dt.Rows(0)
    '        '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
    '        '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
    '        '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
    '        '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
    '        '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。
    '        If Convert.ToString(dr("DaSource")) = "1" Then Exit Sub '無法逕行修改，僅保留填寫狀態供訓練單位查詢。
    '    End If

    '    Dim da As SqlDataAdapter = Nothing
    '    Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
    '    Try
    '        sql = " SELECT * FROM Stud_QuestionFac WHERE SOCID ='" & SOCID & "' "
    '        dt = DbAccess.GetDataTable(sql, da, trans)
    '        If dt.Rows.Count = 0 Then
    '            dr = dt.NewRow
    '            dt.Rows.Add(dr)
    '            dr("SOCID") = SOCID
    '        Else
    '            dr = dt.Rows(0)
    '            ''單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
    '            'If Convert.ToString(dr("DaSource")) = "1" Then Exit Sub '無法逕行修改，僅保留填寫狀態供訓練單位查詢。
    '        End If
    '        If colArray(2).ToString = "" Then dr("Q1_1") = Convert.DBNull Else dr("Q1_1") = colArray(2).ToString
    '        If colArray(3).ToString = "" Then dr("Q1_2") = Convert.DBNull Else dr("Q1_2") = colArray(3).ToString
    '        If colArray(4).ToString = "" Then dr("Q1_3") = Convert.DBNull Else dr("Q1_3") = colArray(4).ToString
    '        If colArray(5).ToString = "" Then dr("Q1_4") = Convert.DBNull Else dr("Q1_4") = colArray(5).ToString
    '        If colArray(6).ToString = "" Then dr("Q1_5") = Convert.DBNull Else dr("Q1_5") = colArray(6).ToString
    '        If colArray(7).ToString = "" Then dr("Q1_6") = Convert.DBNull Else dr("Q1_6") = colArray(7).ToString
    '        If colArray(8).ToString = "" Then dr("Q1_6_CCourName") = Convert.DBNull Else dr("Q1_6_CCourName") = Left(colArray(8).ToString, 30)
    '        If colArray(9).ToString = "" Then dr("Q1_6_CHour") = Convert.DBNull Else dr("Q1_6_CHour") = colArray(9).ToString
    '        If colArray(10).ToString = "" Then dr("Q1_6_MCourName") = Convert.DBNull Else dr("Q1_6_MCourName") = Left(colArray(10).ToString, 30)
    '        If colArray(11).ToString = "" Then dr("Q1_6_MHour") = Convert.DBNull Else dr("Q1_6_MHour") = colArray(11).ToString
    '        If colArray(12).ToString = "" Then dr("Q2_1") = Convert.DBNull Else dr("Q2_1") = colArray(12).ToString
    '        If colArray(13).ToString = "" Then dr("Q2_1_CCourName") = Convert.DBNull Else dr("Q2_1_CCourName") = Left(colArray(13).ToString, 30)
    '        If colArray(14).ToString = "" Then dr("Q2_1_CHour") = Convert.DBNull Else dr("Q2_1_CHour") = colArray(14).ToString
    '        If colArray(15).ToString = "" Then dr("Q2_1_MCourName") = Convert.DBNull Else dr("Q2_1_MCourName") = Left(colArray(15).ToString, 30)
    '        If colArray(16).ToString = "" Then dr("Q2_1_MHour") = Convert.DBNull Else dr("Q2_1_MHour") = colArray(16).ToString
    '        If colArray(17).ToString = "" Then dr("Q2_2") = Convert.DBNull Else dr("Q2_2") = colArray(17).ToString
    '        If colArray(18).ToString = "" Then dr("Q2_3") = Convert.DBNull Else dr("Q2_3") = colArray(18).ToString
    '        If colArray(19).ToString = "" Then dr("Q2_4") = Convert.DBNull Else dr("Q2_4") = colArray(19).ToString
    '        If colArray(20).ToString = "" Then dr("Q2_5") = Convert.DBNull Else dr("Q2_5") = colArray(20).ToString
    '        If colArray(21).ToString = "" Then dr("Q3_1") = Convert.DBNull Else dr("Q3_1") = colArray(21).ToString
    '        If colArray(22).ToString = "" Then dr("Q3_2") = Convert.DBNull Else dr("Q3_2") = colArray(22).ToString
    '        If colArray(23).ToString = "" Then dr("Q3_3") = Convert.DBNull Else dr("Q3_3") = colArray(23).ToString
    '        If colArray(24).ToString = "" Then dr("Q4") = Convert.DBNull Else dr("Q4") = colArray(24).ToString
    '        If colArray(25).ToString = "" Then dr("Q5") = Convert.DBNull Else dr("Q5") = colArray(25).ToString
    '        If colArray(26).ToString = "" Then dr("Q5_Note_News") = Convert.DBNull Else dr("Q5_Note_News") = Left(colArray(26).ToString, 100)
    '        If colArray(27).ToString = "" Then dr("Q5_Note_Other") = Convert.DBNull Else dr("Q5_Note_Other") = Left(colArray(27).ToString, 100)
    '        If colArray(28).ToString = "" Then dr("Q6") = Convert.DBNull Else dr("Q6") = colArray(28).ToString
    '        If colArray(29).ToString = "" Then dr("Q6_Note1") = Convert.DBNull Else dr("Q6_Note1") = Left(colArray(29).ToString, 100)
    '        If colArray(30).ToString = "" Then dr("Q6_Note2") = Convert.DBNull Else dr("Q6_Note2") = Left(colArray(30).ToString, 100)
    '        If colArray(31).ToString = "" Then dr("Q7") = Convert.DBNull Else dr("Q7") = colArray(31).ToString

    '        Select Case sm.UserInfo.Years
    '            Case Is <= 2007 '96年度項目數
    '                If colArray(32).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(32).ToString
    '                If colArray(33).ToString = "" Then dr("Q9") = Convert.DBNull Else dr("Q9") = colArray(33).ToString
    '                If colArray(34).ToString = "" Then dr("Q9_Note") = Convert.DBNull Else dr("Q9_Note") = Left(colArray(34).ToString, 100)
    '            Case Is <= 2011 '97年度項目數
    '                If colArray(32).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(32).ToString
    '                If colArray(33).ToString = "" Then dr("Q9_1") = Convert.DBNull Else dr("Q9_1") = colArray(33).ToString
    '                If colArray(34).ToString = "" Then dr("Q9_2") = Convert.DBNull Else dr("Q9_2") = colArray(34).ToString
    '                If colArray(35).ToString = "" Then dr("Q9_3") = Convert.DBNull Else dr("Q9_3") = colArray(35).ToString
    '                If colArray(36).ToString = "" Then dr("Q10") = Convert.DBNull Else dr("Q10") = colArray(36).ToString
    '                If colArray(37).ToString = "" Then dr("Q11") = Convert.DBNull Else dr("Q11") = colArray(37).ToString
    '                If colArray(38).ToString = "" Then dr("Q12") = Convert.DBNull Else dr("Q12") = colArray(38).ToString
    '            Case Else '101年度項目數
    '                If colArray(32).ToString = "" Then dr("Q7_8") = Convert.DBNull Else dr("Q7_8") = colArray(32).ToString
    '                If colArray(33).ToString = "" Then dr("Q7_9") = Convert.DBNull Else dr("Q7_9") = colArray(33).ToString
    '                If colArray(34).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(34).ToString
    '                If colArray(35).ToString = "" Then dr("Q9_1") = Convert.DBNull Else dr("Q9_1") = colArray(35).ToString
    '                If colArray(36).ToString = "" Then dr("Q9_2") = Convert.DBNull Else dr("Q9_2") = colArray(36).ToString
    '                If colArray(37).ToString = "" Then dr("Q9_3") = Convert.DBNull Else dr("Q9_3") = colArray(37).ToString
    '                If colArray(38).ToString = "" Then dr("Q10") = Convert.DBNull Else dr("Q10") = colArray(38).ToString
    '                If colArray(39).ToString = "" Then dr("Q11") = Convert.DBNull Else dr("Q11") = colArray(39).ToString
    '                If colArray(40).ToString = "" Then dr("Q12") = Convert.DBNull Else dr("Q12") = colArray(40).ToString
    '        End Select

    '        ''If colArray(32).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(32).ToString
    '        ''**by Milor 20070412----start
    '        'If sm.UserInfo.Years <= 2007 Then    '96年版
    '        '    If colArray(33).ToString = "" Then dr("Q9") = Convert.DBNull Else dr("Q9") = colArray(33).ToString
    '        '    If colArray(34).ToString = "" Then dr("Q9_Note") = Convert.DBNull Else dr("Q9_Note") = Left(colArray(34).ToString, 100)
    '        'Else    '97年版
    '        '    If colArray(33).ToString = "" Then dr("Q9_1") = Convert.DBNull Else dr("Q9_1") = colArray(33).ToString
    '        '    If colArray(34).ToString = "" Then dr("Q9_2") = Convert.DBNull Else dr("Q9_2") = colArray(34).ToString
    '        '    If colArray(35).ToString = "" Then dr("Q9_3") = Convert.DBNull Else dr("Q9_3") = colArray(35).ToString
    '        '    If colArray(36).ToString = "" Then dr("Q10") = Convert.DBNull Else dr("Q10") = colArray(36).ToString
    '        '    If colArray(37).ToString = "" Then dr("Q11") = Convert.DBNull Else dr("Q11") = colArray(37).ToString
    '        '    If colArray(38).ToString = "" Then dr("Q12") = Convert.DBNull Else dr("Q12") = colArray(38).ToString
    '        'End If
    '        ''**by Milor 20070412----end
    '        dr("DaSource") = "2" '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
    '        dr("ModifyAcct") = sm.UserInfo.UserID
    '        dr("ModifyDate") = CDate(FillFormDate) 'Now

    '        DbAccess.UpdateDataTable(dt, da, trans)
    '        DbAccess.CommitTrans(trans)
    '    Catch ex As Exception
    '        DbAccess.RollbackTrans(trans)
    '        Me.ViewState("WriteDBmsg") += ex.ToString
    '        'Throw ex
    '        'Common.MessageBox(Me, ex.ToString)
    '    End Try
    '    Call TIMS.CloseDbConn(tConn)
    'End Sub
    '確認學員填寫狀況。取得學員學號。
    'Public Shared Function getSOCID(ByVal OCID As String, ByVal StudID As String, ByRef tConn As SqlConnection) As String
    '    Dim Rst As String = ""
    '    Dim objstr As String
    '    Dim dt As DataTable
    '    objstr = " SELECT dbo.fn_GET_AllID(CONVERT(numeric, '" & OCID & "'),'STUDID','" & StudID & "') AS SOCID  "
    '    dt = DbAccess.GetDataTable(objstr, tConn)
    '    If dt.Rows.Count > 0 Then Rst = dt.Rows(0)("SOCID") 'DbAccess.ExecuteScalar(objstr, objconn)
    '    Return Rst
    'End Function
#End Region

    '判斷機構是否只有一個班級
    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGrid1.Style("display") = "none"
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGrid1.Style("display") = "none"
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    ''' <summary>
    ''' 取得報表名稱使用變數方式-
    ''' </summary>
    ''' <param name="Mpg"></param>
    ''' <param name="sType"></param>
    ''' <returns></returns>
    Function Get_SD11004rpt(ByRef Mpg As Page, ByVal sType As String) As String
        Dim rst As String = ""
        If sType = "" Then Return rst
        '(直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Mpg) Then Return rst

        'Const cst_y07 As String = "07"
        'Const cst_y08 As String = "08"
        Const cst_y12 As String = "12"
        Const cst_y17 As String = "17"

        Dim myyear As String = cst_y12 '"12" '2012年的版本
        'If Val(sm.UserInfo.Years) <= 2007 Then myyear = cst_y07 '"07" '顯示96年版的檔案連結
        'If Val(sm.UserInfo.Years) >= 2008 Then myyear = cst_y08 '"08" '顯示97年版的檔案連結(, 兩個檔案連結的描述是一樣的, 但檔案不同)
        If Val(sm.UserInfo.Years) >= 2012 Then myyear = cst_y12 '"12" '2012年的版本
        If Val(sm.UserInfo.Years) >= 2017 Then myyear = cst_y17 '"17" '2017年的版本

        Dim iPYNum17 As Integer = TIMS.sUtl_GetPYNum17(Me)
        If iPYNum17 = 2 Then myyear = cst_y17 '"17"

        Select Case sType
            Case cst_rptEmp '空白表
                rst = cst_rptEmp & myyear
            Case cst_rptN '班級問卷表
                rst = cst_rptN & myyear
            Case cst_rptData '學員答案表
                rst = cst_rptData & myyear
            Case cst_Addaspx '"SD_11_005_add@xxx.aspx" '@xxx 程式用
                rst = Replace(cst_Addaspx, "@xxx", myyear)
        End Select
        Return rst
    End Function

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Protected Sub DG_stud_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DG_stud.SelectedIndexChanged

    End Sub
End Class