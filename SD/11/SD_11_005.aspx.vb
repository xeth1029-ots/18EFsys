Partial Class SD_11_005
    Inherits AuthBasePage

#Region "(No Use)"

    'STUD_QUESTIONFIN 產投
    'ReportQuery
    'SD_11_005_emp  +myyear @BussinessTrain 
    'SD_11_005_n +NN @BussinessTrain
    'SD_11_005_ +rptyear @BussinessTrain

    '2012 old:
    'SD_11_005_emp09 @BussinessTrain 
    'SD_11_005_n09 @BussinessTrain
    'SD_11_005_09 @BussinessTrain

    '2012 new:產投使用 12
    'SD_11_005_emp12 @BussinessTrain 
    'SD_11_005_n12 @BussinessTrain
    'SD_11_005_12 @BussinessTrain

    '未設計。12b
    '2013 new:充電起飛使用 12b SD_11_005*.jrxml
    'SD_11_005_emp12b @BussinessTrain 
    'SD_11_005_n12b @BussinessTrain
    'SD_11_005_12b @BussinessTrain

    'Dim sqlAdapter As SqlDataAdapter
    'Dim stud_table As DataTable
    'Dim class_table As DataTable

    'Dim FunDr As DataRow
    'Dim Q1, Q2, Q3, Q4, Q5, Q6, Q7, Q10, Q11, Q12, Q13, Q14, Q15, Q16 As String
    'Dim Q8, Q8_1_Note, Q8_2_Note, Q8_3_Note, Q9_1_Note, Q9_2_Note, Q9_3_Note, Q10_1_Note, Q10_2_Note, Q10_3_Note, BusName As String
    'Dim Q6_7 As String
    'Dim Q6_8 As String
    'Dim WriteStatus As Boolean
    'Dim FillFormDate As String '填寫日期 
    'Dim SOCID As String        '學員編號

#End Region

    '特殊開放
    Dim sOpenOCID As String = "" '"59796、65430、59798、66584"

    '填寫訓後動態調查表限定3個月~4個月

    Dim sPrtFileName1 As String = "" 'SD_11_005
    Const cst_rptEmp As String = "SD_11_005_emp" '空白表
    Const cst_rptN As String = "SD_11_005_n" '班級問卷表
    Const cst_rptData As String = "SD_11_005_" '學員答案表

    Const cst_Addaspx As String = "SD_11_005_add@xxx.aspx?" '@xxx 程式用

    '依結訓日 判斷使用規則
    Const cst_wwFTDate1 As String = "2012/02/01" '(特殊規則區間1)(Old)
    Const cst_wwFTDate2 As String = "2012/09/01" '(特殊規則區間2)(Old)
    Const cst_wwEndDate1 As String = "2013/01/01" '(特殊規則區間3結束時間)(Old)

    Const ss_QuestionFinSearchStr As String = "QuestionFinSearchStr" 'STUD_QUESTIONFIN
    'Dim rqOCID As String = "" 'Request("ocid") ';
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

        'objconn = DbAccess.GetConnection()
        'rqOCID = Request("ocid")
        'rqOCID = TIMS.ClearSQM(rqOCID)
        ProcessType.Value = TIMS.ClearSQM(Request("ProcessType"))
        Re_OCID.Value = TIMS.ClearSQM(Request("OCID"))
        Re_ID.Value = TIMS.ClearSQM(Request("ID"))
        'ProcessType.Value = TIMS.ClearSQM(ProcessType.Value)
        'Re_OCID.Value = TIMS.ClearSQM(Re_OCID.Value)
        'Re_ID.Value = TIMS.ClearSQM(Re_ID.Value)

        msg.Text = ""
        eMeng.Style("display") = VeMeng.Text
        'search.Attributes("onclick") = "javascript:return search1()"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
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

        'check_add.Value = "0"
        'If blnCanAdds Then check_add.Value = "1"
        'search.Enabled = False
        'check_search.Value = "0"
        'If blnCanSech Then search.Enabled = True
        'If blnCanSech Then check_search.Value = "1"
        'If Not search.Enabled Then TIMS.Tooltip(search, "(Sech)無權限使用該功能", True)

        'check_del.Value = "0"
        'check_mod.Value = "0"
        'If blnCanDel Then check_del.Value = "1"
        'If blnCanMod Then check_mod.Value = "1"

#End Region

        '特殊開放
        Dim SD_11_005_OPEN_OCID As String = TIMS.Utl_GetConfigSet("SD11005OPENOCID")
        SD_11_005_OPEN_OCID = TIMS.ClearSQM(SD_11_005_OPEN_OCID)
        If SD_11_005_OPEN_OCID <> "" Then sOpenOCID = SD_11_005_OPEN_OCID

        If Not Page.IsPostBack Then
            eMeng.Style("display") = VeMeng.Text
            VeMeng.Text = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            StudentTable.Style.Item("display") = "none"
            If Convert.ToString(Request("ProcessType")) = "Back" Then
                If Not Session(ss_QuestionFinSearchStr) Is Nothing Then
                    'Dim StudentTableState As String = ""
                    center.Text = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "center")
                    RIDValue.Value = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "RIDValue")
                    TMID1.Text = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "TMID1")
                    OCID1.Text = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "OCID1")
                    TMIDValue1.Value = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "TMIDValue1")
                    OCIDValue1.Value = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "OCIDValue1")
                    'StudentTableState = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "StudentTable")
                    Dim xValue As String = TIMS.GetMyValue(Session(ss_QuestionFinSearchStr), "Button1")
                    If xValue = "True" Then search_Click(sender, e)
                    Session(ss_QuestionFinSearchStr) = Nothing
                End If
            End If
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button14_Click(sender, e)
            End If
        End If
        Years.Value = sm.UserInfo.Years - 1911
        '**by Milor 20080424開啟空白問卷----start

        Dim okind As String = ""
        If Re_OCID.Value <> "" Then
            Dim sql As String = ""
            sql &= " SELECT b.OrgKind "
            sql &= " FROM Class_ClassInfo a "
            sql &= " JOIN Org_OrgInfo b ON a.ComIDNO = b.ComIDNO "
            sql &= " WHERE a.OCID = " & Re_OCID.Value & ""
            'fix 有接到ocid時才query
            okind = DbAccess.ExecuteScalar(sql, objconn)
        End If

        PrintBlank.Text = "列印空白表單"
        PrintBlank2.Text = "列印空白表單(" & TIMS.Get_PName28(Me, "W", objconn) & ")"
        PrintBlank.Visible = True '列印空白表單(產業人才)
        PrintBlank2.Visible = False '列印空白表單(在職勞工)

        Select Case sm.UserInfo.TPlanID
            Case "28"
                'PrintBlank.Text = "列印空白表單(" & TIMS.cst_OrgKind2txtG & ")"
                'PrintBlank2.Text = "列印空白表單(" & TIMS.cst_OrgKind2txtW & ")"
                PrintBlank.Text = "列印空白表單(" & TIMS.Get_PName28(Me, "G", objconn) & ")"
                PrintBlank2.Text = "列印空白表單(" & TIMS.Get_PName28(Me, "W", objconn) & ")"
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

        sPrtFileName1 = Get_SD11005rpt(Me, cst_rptEmp)
        Dim MyValue As String = ""
        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue &= "&Years=" & sm.UserInfo.Years
        PrintBlank.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, "&OrgKind=01" & MyValue) '列印
        PrintBlank2.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, "&OrgKind=10" & MyValue) '列印
        '**by Milor 20080424----end
    End Sub

    '取得報表名稱使用變數方式
    Function Get_SD11005rpt(ByRef Mpg As Page, ByVal sType As String) As String
        Dim rst As String = ""
        If sType = "" Then Return rst
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Mpg) Then Return rst

        'Const cst_y07 As String = "07"
        'Const cst_y08 As String = "08"
        'Const cst_y09 As String = "09"
        Const cst_y12 As String = "12"
        Const cst_y17 As String = "17"

        Dim myyear As String = cst_y12 '"12" '2012年的版本
        'If Val(sm.UserInfo.Years) <= 2007 Then myyear = cst_y07 '"07"  '顯示96年版的檔案連結
        'If Val(sm.UserInfo.Years) = 2008 Then myyear = cst_y08 '"08" '顯示97年版的檔案連結(, 兩個檔案連結的描述是一樣的, 但檔案不同)
        'If Val(sm.UserInfo.Years) >= 2009 Then myyear = cst_y09 '"09" '2009年的版本
        If Val(sm.UserInfo.Years) >= 2012 Then myyear = cst_y12 '"12" '2012年的版本
        If Val(sm.UserInfo.Years) >= 2017 Then myyear = cst_y17 '"17" '2017年的版本
        Dim iPYNum17 As Integer = TIMS.sUtl_GetPYNum17(Me)
        If iPYNum17 = 2 Then myyear = cst_y17 '"17"

#Region "(No Use)"

        'If Val(sm.UserInfo.Years) <= 2007 Then    '顯示96年版的檔案連結
        'ElseIf sm.UserInfo.Years = 2008 Then
        'ElseIf sm.UserInfo.Years > 2008 Then
        'End If
        ''2012年的版本
        'If sm.UserInfo.Years >= 2012 Then myyear = "12"
        ''54	充電起飛計畫(補助在職勞工參訓)
        'If sm.UserInfo.TPlanID = "54" Then myyear = "12b"

#End Region

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

    Sub Search1()
        PanelDataGrid1.Visible = True
        PanelDG_stud.Visible = False

        StudentTable.Style.Item("display") = "none"
        msg2.Visible = False
        Label1.Visible = False

        Dim sql As String = ""
        sql &= " SELECT a.OCID ,a.CyclType ,a.LevelType ,a.ClassCName" & vbCrLf
        sql &= " ,a.FTDate" & vbCrLf
        sql &= " ,ISNULL(c.total,0) total" & vbCrLf
        sql &= " ,ISNULL(c.num1,0) num1" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = a.planid" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= "  SELECT cs.OCID" & vbCrLf
        sql &= "  ,COUNT(CASE WHEN cs.StudStatus NOT IN (2,3) THEN 1 END) total" & vbCrLf
        sql &= "  ,COUNT(CASE WHEN sq.socid IS NOT NULL THEN 1 END) num1" & vbCrLf
        sql &= "  FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= "  LEFT JOIN STUD_SUBSIDYCOST sc ON sc.SOCID = cs.SOCID" & vbCrLf
        sql &= "  LEFT JOIN STUD_QUESTIONFIN sq ON sq.SOCID = cs.SOCID" & vbCrLf
        'STUD_QUESTIONFIN '201608 加入排除條件 AMU 'A.排除離退訓(離退訓作業功能)  B.排除有結訓未申請(補助申請功能)  C.排除審核不通過(補助審核功能)的學員
        sql &= "  WHERE cs.STUDSTATUS NOT IN (2,3)" & vbCrLf '非離退
        sql &= "  AND sc.SOCID IS NOT NULL" & vbCrLf '有申請資料
        sql &= "  AND ISNULL(sc.AppliedStatusM,'Y') = 'Y'" & vbCrLf '審核通過 或申請中的
        sql &= "  GROUP BY cs.OCID" & vbCrLf
        sql &= " ) c on c.OCID = a.OCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years=  '" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then sql &= " AND a.RID = '" & RIDValue.Value & "'" & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND a.OCID = '" & OCIDValue1.Value & "'" & vbCrLf
        sql &= " ORDER BY a.OCID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        'class_table = DbAccess.GetDataTable(sqlstr, sqlAdapter, objconn)

        DataGrid1.Visible = False
        DataGrid1.Style.Item("display") = "none"
        msg.Text = "查無資料!!"
        msg.Visible = True

        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            DataGrid1.Style.Item("display") = "inline"
            msg.Text = ""
            msg.Visible = False
            DataGrid1.DataKeyField = "OCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
            ''分頁用-   Start
            'DataGridPage1.MyDataTable = stud_table
            'DataGridPage1.FirstTime()
            ''分頁用-   End
        End If

    End Sub

    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        Search1()
    End Sub

    '保留Session
    Sub GetSearchStr()
        Session(ss_QuestionFinSearchStr) = "center=" & center.Text
        Session(ss_QuestionFinSearchStr) += "&RIDValue=" & RIDValue.Value
        Session(ss_QuestionFinSearchStr) += "&TMID1=" & TMID1.Text
        Session(ss_QuestionFinSearchStr) += "&OCID1=" & OCID1.Text
        Session(ss_QuestionFinSearchStr) += "&TMIDValue1=" & TMIDValue1.Value
        Session(ss_QuestionFinSearchStr) += "&OCIDValue1=" & OCIDValue1.Value
        Session(ss_QuestionFinSearchStr) += "&Button1=" & DG_stud.Visible
        Session(ss_QuestionFinSearchStr) += "&StudentTable=" & StudentTable.Style.Item("display")
    End Sub

    Function GET_LABEL1TXT(ByRef s_OCID As String) As String
        Dim rst As String = ""
        If s_OCID = "" Then Return rst

        Dim drCC As DataRow = TIMS.GetOCIDDate(s_OCID, objconn)
        If drCC Is Nothing Then Return rst

        rst = "班別：" & Convert.ToString(drCC("ClassCName2"))
        If Not IsDBNull(drCC("LevelType")) Then
            If CInt(drCC("LevelType")) <> 0 Then rst &= "第" & TIMS.GetChtNum(CInt(drCC("LevelType"))) & "階段"
        End If

        Dim sql As String = ""
        sql = " SELECT * FROM V_STUDENTCOUNT WHERE ocid =" & s_OCID
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then Return rst
        rst &= "(開訓人數:" & dr("opencount").ToString & "&nbsp;&nbsp;在結訓人數:" & dr("TrainCount").ToString & "&nbsp;&nbsp;離退訓人數:" & dr("LeaveCount").ToString & ")"
        Return rst
    End Function

    Function GET_STUDENT1(ByRef s_OCID As String) As DataTable
        Dim dt2 As DataTable = Nothing

        Dim sql As String = ""
        sql &= " SELECT b.studentid ,b.StudStatus ,c.name" & vbCrLf
        sql &= " ,b.OCID ,b.SOCID" & vbCrLf
        sql &= " ,b.RejectTDate1 ,b.RejectTDate2" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.FTDate, 111) FTDate" & vbCrLf  '結訓日期
        sql &= " ,CONVERT(VARCHAR, DATEADD(DAY,1,DATEADD(MONTH,3,A.FTDATE)), 111) WRDATE1" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, DATEADD(MONTH,4,A.FTDATE), 111) WRDATE2" & vbCrLf
        '填寫訓後動態調查表限定3個月~4個月
        '可填寫未填寫時 canWrite="1"
        sql &= "  ,CASE WHEN DATEDIFF(DAY, DATEADD(DAY,1, DATEADD(MONTH,3,A.FTDATE)), GETDATE()) >= 0" & vbCrLf
        '未達結訓後3個月 (4個月?)
        sql &= "  AND DATEDIFF(DAY, GETDATE(), DateAdd(MONTH, 4, A.FTDATE)) >= 0 THEN 1 ELSE 0 END CANWRITE" & vbCrLf
        '系統目前日期
        sql &= "  ,CONVERT(varchar, GETDATE(), 111) aToday ,d.SOCID finSOCID " & vbCrLf
        sql &= " ,d.DASOURCE" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO A" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS B ON A.OCID = B.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO C ON B.SID = C.SID" & vbCrLf
        sql &= " LEFT JOIN STUD_SUBSIDYCOST sc ON sc.SOCID = b.SOCID" & vbCrLf
        sql &= " LEFT JOIN STUD_QUESTIONFIN d ON d.SOCID = b.SOCID" & vbCrLf
        '201608 加入排除條件 AMU	'A.排除離退訓(離退訓作業功能)	'B.排除有結訓未申請(補助申請功能)	'C.排除審核不通過(補助審核功能)的學員	      
        sql &= " WHERE b.STUDSTATUS NOT IN (2,3)" & vbCrLf '非離退
        sql &= " AND sc.SOCID IS NOT NULL" & vbCrLf '有申請資料
        sql &= " AND ISNULL(sc.AppliedStatusM, 'Y') = 'Y'" & vbCrLf '審核通過 或申請中的
        sql &= " AND a.OCID = @OCID" & vbCrLf
        sql &= " ORDER BY B.STUDENTID" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = s_OCID 'e.CommandArgument
            'dt2.Load(.ExecuteReader())
        End With
        dt2 = DbAccess.GetDataTable(sCmd2.CommandText, objconn, sCmd2.Parameters)
        'Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
        Return dt2
    End Function


    '班級
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing Then Return

        Select Case e.CommandName
            Case "view"
                Dim s_OCID As String = TIMS.ClearSQM(e.CommandArgument)
                Dim s_Label1 As String = GET_LABEL1TXT(s_OCID)
                Label1.Text = s_Label1

                Dim dt2 As DataTable = GET_STUDENT1(s_OCID)

                'Session("DTable_Stuednt") = Nothing
                msg2.Visible = True
                msg2.Text = "查無此班學生資料!"
                StudentTable.Style.Item("display") = "none"
                Label1.Visible = False
                PanelDataGrid1.Visible = True
                PanelDG_stud.Visible = False

                If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
                    PanelDataGrid1.Visible = False
                    PanelDG_stud.Visible = True
                    'msg.Visible = False

                    msg2.Text = ""
                    StudentTable.Style.Item("display") = "" '"inline"
                    Label1.Visible = True

                    DG_stud.DataSource = dt2
                    DG_stud.DataBind()
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            Dim mybut1 As Button = e.Item.FindControl("Button1")
            Dim mybut2 As Button = e.Item.FindControl("Button3")
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1

            Dim OCID As String = drv("OCID").ToString

            e.Item.Cells(1).Text = TIMS.GET_CLASSNAME(Convert.ToString(drv("ClassCName")), Convert.ToString(drv("CyclType")))

            mybut1.CommandArgument = DataGrid1.DataKeys(e.Item.ItemIndex)

            'Dim iYears As Integer = sm.UserInfo.Years - 1911
            sPrtFileName1 = Get_SD11005rpt(Me, cst_rptN)
            Dim myValue As String = ""
            myValue &= "&TPlanID=" & sm.UserInfo.TPlanID
            myValue &= "&Years=" & sm.UserInfo.Years
            myValue &= "&OCID=" & OCID
            mybut2.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, myValue)
        End If
    End Sub

    '學員
    Private Sub DG_stud_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_stud.ItemCommand
        Dim cmdArg As String = e.CommandArgument
        If cmdArg = "" Then Exit Sub
        Dim rSOCID As String = TIMS.GetMyValue(cmdArg, "socid")
        Dim rOCID As String = TIMS.GetMyValue(cmdArg, "ocid")
        Dim canWrite As String = TIMS.GetMyValue(cmdArg, "canWrite")
        If rSOCID = "" Then Exit Sub
        If rOCID = "" Then Exit Sub
        sPrtFileName1 = Get_SD11005rpt(Me, cst_Addaspx)
        Select Case e.CommandName
            Case "insert" '新增
                If canWrite = "1" Then '可填寫
                    Call GetSearchStr()
                    'Response.Redirect("SD_11_005_add" & rptyear & ".aspx?ProcessType=Insert&SOCID=" & rSOCID & "&ocid=" & rOCID & "&ID=" & Request("ID"))
                    Dim url1 As String = sPrtFileName1 & "ID=" & Request("ID")
                    url1 &= "&ProcessType=Insert"
                    url1 &= "&SOCID=" & rSOCID
                    url1 &= "&ocid=" & rOCID
                    Call TIMS.Utl_Redirect(Me, objconn, url1)
                End If
            Case "clear" '清除重填
                Call GetSearchStr()
                Dim url1 As String = sPrtFileName1 & "ID=" & Request("ID")
                url1 &= "&ProcessType=del"
                url1 &= "&SOCID=" & rSOCID
                url1 &= "&ocid=" & rOCID
                Dim strScript As String
                strScript = "<script language=""javascript"">" + vbCrLf
                strScript += "if (window.confirm('此動作會刪除受訓學員訓後動態調查表資料，是否確定刪除?')){" + vbCrLf
                strScript += "location.href ='" & url1 & "';}" + vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript("", strScript)
            Case "check" '查詢
                Call GetSearchStr()
                Dim url1 As String = sPrtFileName1 & "ID=" & Request("ID")
                url1 &= "&ProcessType=check"
                url1 &= "&SOCID=" & rSOCID
                url1 &= "&ocid=" & rOCID
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "Edit" '修改
                Call GetSearchStr()
                'Response.Redirect("SD_11_005_add" & rptyear & ".aspx?ProcessType=Edit&SOCID=" & rSOCID & "&ocid=" & rOCID & "&ID=" & Request("ID"))
                Dim url1 As String = sPrtFileName1 & "ID=" & Request("ID")
                url1 &= "&ProcessType=Edit"
                url1 &= "&SOCID=" & rSOCID
                url1 &= "&ocid=" & rOCID
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    'DG_stud
    Private Sub DG_stud_ItemDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                '特殊開放
                'Dim flag_can_write_1 As Boolean = False '不可填寫
                Dim flag_can_write_1 As Boolean = True 'False :不可填寫
                If Convert.ToString(drv("canWrite")) <> "1" Then
                    flag_can_write_1 = False '必不可填
                End If
                If sOpenOCID <> "" AndAlso sOpenOCID.IndexOf(Convert.ToString(drv("OCID"))) > -1 Then
                    flag_can_write_1 = True '特殊開放'必可填
                End If

                Dim but4 As Button = e.Item.FindControl("Button4")  '新增
                Dim Edit As Button = e.Item.FindControl("Edit")     '修改
                Dim but5 As Button = e.Item.FindControl("Button5")  '查看
                Dim but6 As Button = e.Item.FindControl("Button6")  '清除重填
                Dim Print As Button = e.Item.FindControl("Print")   '列印
                Dim cmdArg As String = ""
                TIMS.SetMyValue(cmdArg, "socid", drv("socid"))
                TIMS.SetMyValue(cmdArg, "ocid", drv("ocid"))
                If flag_can_write_1 Then
                    TIMS.SetMyValue(cmdArg, "canWrite", "1") '必可填
                Else
                    TIMS.SetMyValue(cmdArg, "canWrite", Convert.ToString(drv("canWrite")))
                End If

                but4.CommandArgument = cmdArg
                Edit.CommandArgument = cmdArg
                but5.CommandArgument = cmdArg
                Print.CommandArgument = cmdArg
                but6.CommandArgument = cmdArg
                If Len(drv("StudentID")) = 12 Then e.Item.Cells(0).Text = Right(drv("StudentID"), 3) Else e.Item.Cells(0).Text = Right(drv("StudentID"), 2)
                If drv("RejectTDate1").ToString <> "" Then e.Item.Cells(1).Text += "(" & FormatDateTime(drv("RejectTDate1"), 2) & ")"
                If drv("RejectTDate2").ToString <> "" Then e.Item.Cells(1).Text += "(" & FormatDateTime(drv("RejectTDate2"), 2) & ")"
                If Convert.ToString(drv("finSOCID")) <> "" Then
                    '已有資料
                    e.Item.Cells(2).Text = "是"
                    but4.Enabled = False '不可新增
                    TIMS.Tooltip(but4, "已有資料,不可新增", True)
                    Edit.Enabled = True '可以修改
                    but5.Enabled = True '可以查看
                    but6.Enabled = True '可清除重填
                    Print.Enabled = True '可以列印
                    'but6.Enabled = True
                    'If check_mod.Value = "0" AndAlso check_del.Value = "0" Then '兩者功能皆沒有時,不能使用
                    '    but6.Enabled = False
                    '    TIMS.Tooltip(but6, "(Mod.Del)無權限使用該功能", True)
                    'End If
                    'but5.Enabled = True
                    'If check_search.Value = "0" Then
                    '    but5.Enabled = False
                    '    TIMS.Tooltip(but5, "(Sech)無權限使用該功能", True)
                    'End If
                Else
                    e.Item.Cells(2).Text = "否"
                    but4.Enabled = True '可新增
                    Edit.Enabled = False '不可修改
                    but5.Enabled = False '不可查看
                    but6.Enabled = False '不可清除重填
                    Print.Enabled = False '不可以列印
                    TIMS.Tooltip(Edit, "未填寫資料", True)
                    TIMS.Tooltip(but5, "未填寫資料", True)
                    TIMS.Tooltip(but6, "未填寫資料", True)
                    TIMS.Tooltip(Print, "未填寫資料", True)
                    'but4.Enabled = True
                    'If check_add.Value = "0" Then
                    '    but4.Enabled = False
                    '    TIMS.Tooltip(but4, "(Adds)無權限使用該功能", True)
                    'End If
                End If

                '新 '20120830 by AMU '使用下列新規則 'but4.CommandArgument = ""
                '特殊開放
                'If sOpenOCID.IndexOf(Convert.ToString(drv("OCID"))) > -1 Then but4.CommandArgument += "&canWrite=1"
                'but4.CommandArgument += "&canWrite=" & Convert.ToString(drv("canWrite"))
                If Not flag_can_write_1 Then
                    '非填寫時間，註解說明
                    Dim vMsg As String = ""
                    vMsg = ""
                    vMsg += "該班結訓日期為" & Convert.ToString(drv("FTDate")) & "\n"
                    vMsg += "填寫日期為結訓後3個月才可開放填寫\n"
                    vMsg += "即" & Convert.ToString(drv("WrDate1")) & "~" & Convert.ToString(drv("WrDate2")) & "日期區間" & "\n"
                    but4.Attributes("onclick") = "alert('" & vMsg & "');return false;"
                    Edit.Attributes("onclick") = "alert('" & vMsg & "');return false;"
                    but6.Attributes("onclick") = "alert('" & vMsg & "');return false;"
                End If

                'DASOURCE  '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。) 2:TIMS系統() 學員填寫不可列印。
                If Convert.ToString(drv("DaSource")) <> "1" Then
                    sPrtFileName1 = Get_SD11005rpt(Me, cst_rptData)
                    Dim myValue As String = ""
                    myValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                    myValue &= "&Years=" & sm.UserInfo.Years
                    myValue &= "&OCID=" & Convert.ToString(drv("OCID"))
                    myValue &= "&SOCID=" & Convert.ToString(drv("SOCID"))
                    Print.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, myValue)
                    TIMS.Tooltip(Print, "非學員自行填寫，可列印")
                End If

                '單位填寫'資料來源	Null:未知 1:報名網()(學員外網填寫。)2@TIMS系統()
                '新增「填寫來源」檢核機制，由系統判斷填寫來源為產投報名網或TIMS系統：
                '1.若為報名網，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰，無法逕行修改，僅保留填寫狀態供訓練單位查詢。
                '2.若為TIMS系統，保留訓練單位各項功能(新增、修改、查詢、列印、清除重填)。
                '3.若有訓練單位先於TIMS系統協助填寫，學員後於報名網修正之情形者，訓練單位端之各項功能(新增、修改、查詢、列印、清除重填等)將反灰。
                'PS:訓練單位登入,若是學員於外網填寫,是不可修改,不可查詢,不可列印,不可清除重填   但若是分署登入,若是學員於外網填寫,是不可修改,可查詢,可列印,不可清除重填
                If Convert.ToString(drv("DaSource")) = "1" Then
                    '委訓單位
                    If sm.UserInfo.LID = 2 Then
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
                        Print.CommandArgument = ""

                        sPrtFileName1 = Get_SD11005rpt(Me, cst_rptData)
                        Dim myValue As String = ""
                        myValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                        myValue &= "&Years=" & sm.UserInfo.Years
                        myValue &= "&OCID=" & Convert.ToString(drv("OCID"))
                        myValue &= "&SOCID=" & Convert.ToString(drv("SOCID"))
                        Print.Attributes("onclick") = ReportQuery.ReportScript(Me, sPrtFileName1, myValue)
                    End If
                End If
        End Select
    End Sub

#Region "(No Use)"

    '將民國日期改為西元日期 0651010-> 1976/10/10
    'Function ChangeTWDate(ByVal TWDate As String) As String
    '    Return CStr(CInt(Left(TWDate, 3)) + 1911) & "/" & Mid(TWDate, 4, 2) & "/" & Right(TWDate, 2)
    'End Function

    '學號確認。
    'Function getSOCID(ByVal OCID As String, ByVal StudID As String) As String
    '    Dim Rst As String = ""
    '    Dim objstr As String
    '    Dim dt As DataTable
    '    objstr = "select dbo.fn_GET_AllID(" & OCID & ",'STUDID','" & StudID & "') SOCID "
    '    dt = DbAccess.GetDataTable(objstr, objconn)
    '    If dt.Rows.Count > 0 Then
    '        Rst = Convert.ToString(dt.Rows(0)("SOCID"))
    '    End If
    '    Return Rst
    'End Function

#End Region

    '判斷機構是否只有一個班級
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGrid1.Style.Item("display") = "none"
        StudentTable.Style.Item("display") = "none"
        eMeng.Style.Item("display") = "none"
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGrid1.Style.Item("display") = "none"
        StudentTable.Style.Item("display") = "none"
        eMeng.Style.Item("display") = "none"
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

#Region "(No Use)"

    ''匯入
    'Private Sub Button13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button13.Click
    '    Dim Upload_Path As String = "~/SD/11/Temp/"
    '    Dim I As Integer
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
    '        Dim da As SqlDataAdapter = Nothing
    '        Dim trans As SqlTransaction
    '        'Dim conn As SqlConnection = DbAccess.GetConnection
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

    '                '==補強EXCEL可能去除零值之可能性==Start
    '                'AMU 2006/12/14
    '                If Len(colArray(0).ToString) = 6 Then
    '                    colArray(0) = "0" & colArray(0).ToString
    '                End If
    '                '==補強EXCEL可能去除零值之可能性==End

    '                '建立SOCID欄位值
    '                If OCIDValue1.Value = "" Then
    '                    Reason += "未選擇 開班編號(OCID) 無法匯入<BR>"
    '                    Exit Do
    '                End If

    '                If colArray.Length >= 2 Then
    '                    If Len(colArray(1).ToString) = 0 Then
    '                        Reason += "未輸入 學號(StudID) 無法匯入<BR>"
    '                        WriteStatus = False
    '                        writeflag = False

    '                    Else
    '                        SOCID = getSOCID(OCIDValue1.Value, colArray(1).ToString)
    '                        If SOCID <> "" Then SOCID = Trim(SOCID)
    '                        If SOCID = "0" OrElse SOCID = "" Then
    '                            Reason += "學號 " & colArray(1).ToString & " 無此學員資料無法匯入<BR>"
    '                            WriteStatus = False
    '                            writeflag = False
    '                        End If
    '                    End If
    '                Else
    '                    Reason += "資料欄位不足!!<BR>"
    '                    WriteStatus = False
    '                    writeflag = False
    '                End If

    '                If Reason = "" Then Reason += CheckImportData(colArray, writeflag) '檢查資料正確性

    '                '通過檢查，開始輸入資料---------------------Start
    '                If Reason = "" Then
    '                    Call WriteDB(colArray, sr, srr, objconn)
    '                Else
    '                    Dim yearLen As Integer = 0 '各年度欄位長度不同
    '                    If sm.UserInfo.Years <= 2007 Then
    '                        yearLen = 21 - 1
    '                    ElseIf sm.UserInfo.Years = 2008 Then
    '                        yearLen = 22 - 1
    '                    ElseIf sm.UserInfo.Years > 2008 Then
    '                        '2012年的版本
    '                        If sm.UserInfo.Years >= 2012 Then
    '                            yearLen = 17 - 1
    '                        Else
    '                            yearLen = 15 - 1
    '                        End If
    '                    End If
    '                    '欄位過短 不寫入
    '                    If colArray.Length < yearLen Then
    '                        Reason += "資料欄位不足!!<BR>"
    '                        WriteStatus = False
    '                        writeflag = False
    '                    End If

    '                    'If writeflag Then WriteDB(colArray, sr, srr) '若 writeflag 為true 則可繼續新增到資料庫

    '                    '錯誤資料，填入錯誤資料表
    '                    drWrong = dtWrong.NewRow
    '                    dtWrong.Rows.Add(drWrong)

    '                    drWrong("Index") = RowIndex
    '                    If colArray.Length > 2 Then
    '                        drWrong("FillFormDate") = colArray(0) '填寫日期
    '                        drWrong("StudID") = colArray(1)       '學號
    '                        If WriteStatus = False Then drWrong("Status") = "未匯入" Else drWrong("Status") = "匯入"
    '                        drWrong("Reason") = Reason & IIf(WriteStatus = True, " 錯誤答案部份不寫入資料中", "")
    '                    End If
    '                End If
    '            End If
    '            RowIndex += 1 '讀取行累計數
    '        Loop

    '        '開始判別欄位存入------------   End
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
    '            For I = 1 To 100
    '                If I = 100 Then eMeng.Style.Item("display") = "inline"
    '                'Page.RegisterStartupScript("", "<script>{window.document.getElementById('eMeng').style.visibility='visible';}</script>")
    '            Next
    '        End If
    '        sr.Close()
    '        srr.Close()
    '        MyFile.Delete(Server.MapPath(Upload_Path & MyFileName))
    '    End If
    '    Call search_Click(sender, e)
    'End Sub

    ''匯入檢查
    'Function CheckImportData(ByVal colArray As Array, ByRef writeflag As Boolean) '檢查資料正確性
    '    'amu 20061221 因為同意可寫入某些錯誤的資料，但還是要show訊息

    '    Dim Reason As String
    '    Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
    '    Dim sql As String
    '    Dim Num3 As String = "123"
    '    Dim Num4 As String = "1234"
    '    Dim Num5 As String = "12345"
    '    Dim Num6 As String = "123456"
    '    Dim dr As DataRow
    '    Const cst_Len As Integer = 23

    '    Dim yearLen As Integer = 0 '各年度欄位長度不同
    '    If sm.UserInfo.Years <= 2007 Then
    '        yearLen = 21 - 1
    '    ElseIf sm.UserInfo.Years = 2008 Then
    '        yearLen = 22 - 1
    '    ElseIf sm.UserInfo.Years > 2008 Then
    '        yearLen = 15 - 1
    '    End If

    '    If colArray.Length > cst_Len Then
    '        'Reason += "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
    '        Reason += "欄位對應有誤<BR>"
    '        Reason += "請注意欄位中是否有半形逗點<BR>"

    '    ElseIf colArray.Length < yearLen Then
    '        Reason += "欄位對應有誤<BR>"
    '        Reason += "請注意欄位中是否有半形逗點<BR>"

    '    Else
    '        '將民國日期改為西元日期 0651010-> 1976/10/10
    '        FillFormDate = colArray(0).ToString   '填寫日期
    '        SOCID = getSOCID(OCIDValue1.Value, colArray(1).ToString) '學號
    '        Q1 = colArray(2).ToString             '學員目前的近況為何
    '        Q2 = colArray(3).ToString             '學員於結訓後薪資有提升嗎
    '        Q3 = colArray(4).ToString             '學員的職位有變化嗎
    '        Q4 = colArray(5).ToString             '學員的工作滿意度是否提升
    '        Q5 = colArray(6).ToString             '學員目前的工作內容是否與參訓課程內容相關
    '        Q6 = colArray(7).ToString             '學員認為參加訓練對目前的工作是否有幫助

    '        Try
    '            If sm.UserInfo.Years <= 2007 Then
    '                Q7 = colArray(8).ToString             '96年-學員是否有繼續參與進修訓練的意願；97年-承上題，參加本項訓練對學員的幫助是在哪方面
    '                Q8_1_Note = colArray(9).ToString      '學員認為還需要加強哪方面的專業知識使工作進行得更順利 1
    '                Q8_2_Note = colArray(10).ToString     '學員認為還需要加強哪方面的專業知識使工作進行得更順利 2
    '                Q8_3_Note = colArray(11).ToString     '學員認為還需要加強哪方面的專業知識使工作進行得更順利 3
    '                Q9_1_Note = colArray(12).ToString     '學員常和本課程的哪些學員、教師或職員聯絡 1
    '                Q9_2_Note = colArray(13).ToString     '學員常和本課程的哪些學員、教師或職員聯絡 2
    '                Q9_3_Note = colArray(14).ToString     '學員常和本課程的哪些學員、教師或職員聯絡 3
    '                BusName = colArray(15).ToString       '企業名稱
    '                Q10 = colArray(16).ToString           '96年-學員受訓後工作態度是否有改善
    '                Q11 = colArray(17).ToString           '96年-學員受訓後知識技術是否有提升；97年-學員受訓後工作態度是否有改善
    '                Q12 = colArray(18).ToString           '96年-參訓學員工作能力是否有提升；97年-學員受訓後知識技術是否有提升
    '                Q13 = colArray(19).ToString           '96年-企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變；97年-參訓學員工作能力是否有提升
    '                Q14 = colArray(20).ToString           '96年-企業之離職率是否有改變；97年-企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變
    '                Q15 = colArray(21).ToString           '96年-企業繼續辦理專班的計畫；97年-企業之離職率是否有改變

    '            ElseIf sm.UserInfo.Years = 2008 Then
    '                Q7 = colArray(8).ToString             '96年-學員是否有繼續參與進修訓練的意願；97年-承上題，參加本項訓練對學員的幫助是在哪方面
    '                Q8 = colArray(9).ToString           '學員是否有繼續參與進修訓練的意願
    '                Q9_1_Note = colArray(10).ToString   '學員認為還需要加強哪方面的專業知識使工作進行得更順利 1
    '                Q9_2_Note = colArray(11).ToString   '學員認為還需要加強哪方面的專業知識使工作進行得更順利 2
    '                Q9_3_Note = colArray(12).ToString   '學員認為還需要加強哪方面的專業知識使工作進行得更順利 3
    '                Q10_1_Note = colArray(13).ToString  '學員常和本課程的哪些學員、教師或職員連絡 1
    '                Q10_2_Note = colArray(14).ToString  '學員常和本課程的哪些學員、教師或職員連絡 2
    '                Q10_3_Note = colArray(15).ToString  '學員常和本課程的哪些學員、教師或職員連絡 3
    '                BusName = colArray(16).ToString     '企業名稱
    '                Q11 = colArray(17).ToString           '96年-學員受訓後知識技術是否有提升；97年-學員受訓後工作態度是否有改善
    '                Q12 = colArray(18).ToString           '96年-參訓學員工作能力是否有提升；97年-學員受訓後知識技術是否有提升
    '                Q13 = colArray(19).ToString           '96年-企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變；97年-參訓學員工作能力是否有提升
    '                Q14 = colArray(20).ToString           '96年-企業之離職率是否有改變；97年-企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變
    '                Q15 = colArray(21).ToString           '96年-企業繼續辦理專班的計畫；97年-企業之離職率是否有改變
    '                Q16 = colArray(22).ToString '企業繼續辦理專班的計畫

    '            ElseIf sm.UserInfo.Years > 2008 Then
    '                '2012年的版本
    '                If sm.UserInfo.Years >= 2012 Then
    '                    Dim xCol As Integer = 0
    '                    Q6_7 = Convert.ToString(colArray(8))
    '                    Q6_8 = Convert.ToString(colArray(9))

    '                    Q7 = colArray(10).ToString             '96年-學員是否有繼續參與進修訓練的意願；97年-承上題，參加本項訓練對學員的幫助是在哪方面
    '                    Q8 = Convert.ToString(colArray(11)) '學員是否有繼續參與進修訓練的意願
    '                    Q9_1_Note = ""
    '                    Q9_2_Note = ""
    '                    Q9_3_Note = ""
    '                    Q10_1_Note = ""
    '                    Q10_2_Note = ""
    '                    Q10_3_Note = ""
    '                    xCol = 12
    '                    If colArray.Length > xCol Then Q9_1_Note = Convert.ToString(colArray(xCol)) '學員認為還需要加強哪方面的專業知識使工作進行得更順利 1
    '                    xCol = 13
    '                    If colArray.Length > xCol Then Q9_2_Note = Convert.ToString(colArray(xCol)) '學員認為還需要加強哪方面的專業知識使工作進行得更順利 2
    '                    xCol = 14
    '                    If colArray.Length > xCol Then Q9_3_Note = Convert.ToString(colArray(xCol)) '學員認為還需要加強哪方面的專業知識使工作進行得更順利 3
    '                    xCol = 15
    '                    If colArray.Length > xCol Then Q10_1_Note = Convert.ToString(colArray(xCol)) '學員常和本課程的哪些學員、教師或職員連絡 1
    '                    xCol = 16
    '                    If colArray.Length > xCol Then Q10_2_Note = Convert.ToString(colArray(xCol)) '學員常和本課程的哪些學員、教師或職員連絡 2
    '                    xCol = 17
    '                    If colArray.Length > xCol Then Q10_3_Note = Convert.ToString(colArray(xCol)) '學員常和本課程的哪些學員、教師或職員連絡 3

    '                Else
    '                    Q7 = colArray(8).ToString             '96年-學員是否有繼續參與進修訓練的意願；97年-承上題，參加本項訓練對學員的幫助是在哪方面
    '                    Q8 = Convert.ToString(colArray(9)) '學員是否有繼續參與進修訓練的意願
    '                    Q9_1_Note = ""
    '                    Q9_2_Note = ""
    '                    Q9_3_Note = ""
    '                    Q10_1_Note = ""
    '                    Q10_2_Note = ""
    '                    Q10_3_Note = ""

    '                    If colArray.Length > 10 Then Q9_1_Note = Convert.ToString(colArray(10)) '學員認為還需要加強哪方面的專業知識使工作進行得更順利 1
    '                    If colArray.Length > 11 Then Q9_2_Note = Convert.ToString(colArray(11)) '學員認為還需要加強哪方面的專業知識使工作進行得更順利 2
    '                    If colArray.Length > 12 Then Q9_3_Note = Convert.ToString(colArray(12)) '學員認為還需要加強哪方面的專業知識使工作進行得更順利 3
    '                    If colArray.Length > 13 Then Q10_1_Note = Convert.ToString(colArray(13)) '學員常和本課程的哪些學員、教師或職員連絡 1
    '                    If colArray.Length > 14 Then Q10_2_Note = Convert.ToString(colArray(14)) '學員常和本課程的哪些學員、教師或職員連絡 2
    '                    If colArray.Length > 15 Then Q10_3_Note = Convert.ToString(colArray(15)) '學員常和本課程的哪些學員、教師或職員連絡 3
    '                End If
    '            End If
    '        Catch ex As Exception
    '            Reason += "欄位對應可能有誤<BR>"
    '            Reason += "請注意欄位中是否有半形逗點<BR>"
    '        End Try
    '        'Q11 = colArray(17).ToString           '96年-學員受訓後知識技術是否有提升；97年-學員受訓後工作態度是否有改善
    '        'Q12 = colArray(18).ToString           '96年-參訓學員工作能力是否有提升；97年-學員受訓後知識技術是否有提升
    '        'Q13 = colArray(19).ToString           '96年-企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變；97年-參訓學員工作能力是否有提升
    '        'Q14 = colArray(20).ToString           '96年-企業之離職率是否有改變；97年-企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變
    '        'Q15 = colArray(21).ToString           '96年-企業繼續辦理專班的計畫；97年-企業之離職率是否有改變

    '        '填寫日期
    '        If FillFormDate = "" Or Len(FillFormDate) <> 7 Or IsNumeric(FillFormDate) <> True Then
    '            Reason += "填寫日期有誤，必須是民國年格式(yyymmdd)<BR>"
    '            writeflag = False
    '            WriteStatus = False
    '        Else
    '            'FillFormDate = CStr(CInt(Left(FillFormDate, 3)) + 1911) & "/" & Mid(FillFormDate, 4, 2) & "/" & Right(FillFormDate, 2)
    '            '將民國日期改為西元日期 0651010-> 1976/10/10
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

    '        If Q1 <> "" Then
    '            If InStr(Num4, Q1) = 0 Then Reason += "1.1題答案應在1-4之間<BR>"
    '        End If
    '        If Q2 <> "" Then
    '            If InStr(Num5, Q2) = 0 Then Reason += "1.2題答案應在1-5之間<BR>"
    '        End If
    '        If Q3 <> "" Then
    '            If InStr(Num4, Q3) = 0 Then Reason += "1.3題答案應在1-4之間<BR>"
    '        End If
    '        If Q4 <> "" Then
    '            If InStr(Num5, Q4) = 0 Then Reason += "1.4題答案應在1-5之間<BR>"
    '        End If
    '        If Q5 <> "" Then
    '            If InStr(Num5, Q5) = 0 Then Reason += "1.5題答案應在1-5之間<BR>"
    '        End If
    '        If Q6 <> "" Then
    '            If InStr(Num5, Q6) = 0 Then Reason += "1.6題答案應在1-5之間<BR>"
    '        End If
    '        If sm.UserInfo.Years <= 2007 Then
    '            If Q7 <> "" Then
    '                If InStr(Num5, Q7) = 0 Then Reason += "1.7題答案應在1-5之間<BR>"
    '            End If
    '            If Q10 <> "" Then
    '                If InStr(Num5, Q10) = 0 Then Reason += "2.1 題答案應在1-5之間<BR>"
    '            End If
    '            If Q11 <> "" Then
    '                If InStr(Num5, Q11) = 0 Then Reason += "2.2 題答案應在1-5之間<BR>"
    '            End If
    '            If Q12 <> "" Then
    '                If InStr(Num5, Q12) = 0 Then Reason += "2.3 題答案應在1-5之間<BR>"
    '            End If
    '            If Q13 <> "" Then
    '                If InStr(Num5, Q13) = 0 Then Reason += "2.4 題答案應在1-5之間<BR>"
    '            End If
    '            If Q14 <> "" Then
    '                If InStr(Num5, Q14) = 0 Then Reason += "2.5 題答案應在1-5之間<BR>"
    '            End If
    '            If Q15 <> "" Then
    '                If InStr(Num4, Q15) = 0 Then Reason += "2.6 題答案應在1-4之間<BR>"
    '            End If

    '        ElseIf sm.UserInfo.Years = 2008 Then

    '            If Q7 <> "" Then
    '                If InStr(Num3, Q7) = 0 Then Reason += "1.7題答案應在1-3之間<BR>"
    '            End If
    '            If Q8 <> "" Then
    '                If InStr(Num5, Q8) = 0 Then Reason += "1.8題答案應在1-5之間<BR>"
    '            End If
    '            If Q11 <> "" Then
    '                If InStr(Num5, Q11) = 0 Then Reason += "2.2 題答案應在1-5之間<BR>"
    '            End If
    '            If Q12 <> "" Then
    '                If InStr(Num5, Q12) = 0 Then Reason += "2.3 題答案應在1-5之間<BR>"
    '            End If
    '            If Q13 <> "" Then
    '                If InStr(Num5, Q13) = 0 Then Reason += "2.4 題答案應在1-5之間<BR>"
    '            End If
    '            If Q14 <> "" Then
    '                If InStr(Num5, Q14) = 0 Then Reason += "2.5 題答案應在1-5之間<BR>"
    '            End If
    '            If Q15 <> "" Then
    '                If InStr(Num5, Q15) = 0 Then Reason += "2.5 題答案應在1-5之間<BR>"
    '            End If
    '            If Q16 <> "" Then
    '                If InStr(Num4, Q16) = 0 Then Reason += "2.6題答案應在1-4之間<BR>"
    '            End If
    '        ElseIf sm.UserInfo.Years > 2008 Then
    '            '2012年的版本
    '            If sm.UserInfo.Years >= 2012 Then
    '                If Q6_7 <> "" Then
    '                    If InStr(Num5, Q6_7) = 0 Then Reason += "1.7題答案應在1-5之間<BR>"
    '                End If
    '                If Q6_8 <> "" Then
    '                    If InStr(Num5, Q6_8) = 0 Then Reason += "1.8題答案應在1-5之間<BR>"
    '                End If
    '                If Q7 <> "" Then
    '                    If InStr(Num3, Q7) = 0 Then Reason += "1.9題答案應在1-3之間<BR>"
    '                End If
    '                If Q8 <> "" Then
    '                    If InStr(Num5, Q8) = 0 Then Reason += "1.10題答案應在1-5之間<BR>"
    '                End If
    '            Else
    '                If Q7 <> "" Then
    '                    If InStr(Num3, Q7) = 0 Then Reason += "1.7題答案應在1-3之間<BR>"
    '                End If
    '                If Q8 <> "" Then
    '                    If InStr(Num5, Q8) = 0 Then Reason += "1.8題答案應在1-5之間<BR>"
    '                End If
    '            End If
    '        End If
    '        'If Q11 <> "" Then
    '        '    If InStr(Num5, Q11) = 0 Then Reason += "2.2 題答案應在1-5之間<BR>"
    '        'End If
    '        'If Q12 <> "" Then
    '        '    If InStr(Num5, Q12) = 0 Then Reason += "2.3 題答案應在1-5之間<BR>"
    '        'End If
    '        'If Q13 <> "" Then
    '        '    If InStr(Num5, Q13) = 0 Then Reason += "2.4 題答案應在1-5之間<BR>"
    '        'End If
    '        'If Q14 <> "" Then
    '        '    If InStr(Num5, Q14) = 0 Then Reason += "2.5 題答案應在1-5之間<BR>"
    '        'End If
    '    End If

    '    Return Reason
    'End Function

    ''匯入寫DB
    'Private Sub WriteDB(ByVal colArray As System.Array, ByRef sr As System.IO.Stream, ByRef srr As System.IO.StreamReader, ByVal tConn As SqlConnection)
    '    '將資料寫入資料表中
    '    Dim trans As SqlTransaction
    '    'Dim conn As SqlConnection = DbAccess.GetConnection
    '    Dim dt As DataTable
    '    Dim sql As String
    '    Dim dr As DataRow
    '    Dim da As SqlDataAdapter = Nothing
    '    Dim i As Integer

    '    FillFormDate = colArray(0).ToString       '填寫日期 
    '    FillFormDate = ChangeTWDate(FillFormDate) '填寫日期 
    '    SOCID = getSOCID(OCIDValue1.Value, colArray(1).ToString) '學號 

    '    Try
    '        trans = DbAccess.BeginTrans(tConn)
    '        sql = "SELECT * FROM STUD_QUESTIONFIN WHERE SOCID ='" & SOCID & "'"
    '        dt = DbAccess.GetDataTable(sql, da, trans)
    '        If dt.Rows.Count = 0 Then
    '            dr = dt.NewRow
    '            dt.Rows.Add(dr)
    '            dr("SOCID") = SOCID
    '        Else
    '            dr = dt.Rows(0)
    '        End If
    '        If colArray(2).ToString = "" Then dr("Q1") = Convert.DBNull Else dr("Q1") = colArray(2).ToString

    '        If sm.UserInfo.Years > 2008 Then             '如果是2009年第一題答4,第2,3,4,5題就不存值
    '            If colArray(2).ToString = "4" Then
    '                dr("Q2") = Convert.DBNull
    '                dr("Q3") = Convert.DBNull
    '                dr("Q4") = Convert.DBNull
    '                dr("Q5") = Convert.DBNull
    '            Else
    '                If colArray(3).ToString = "" Then dr("Q2") = Convert.DBNull Else dr("Q2") = colArray(3).ToString
    '                If colArray(4).ToString = "" Then dr("Q3") = Convert.DBNull Else dr("Q3") = colArray(4).ToString
    '                If colArray(5).ToString = "" Then dr("Q4") = Convert.DBNull Else dr("Q4") = colArray(5).ToString
    '                If colArray(6).ToString = "" Then dr("Q5") = Convert.DBNull Else dr("Q5") = colArray(6).ToString

    '            End If
    '        Else
    '            If colArray(3).ToString = "" Then dr("Q2") = Convert.DBNull Else dr("Q2") = colArray(3).ToString
    '            If colArray(4).ToString = "" Then dr("Q3") = Convert.DBNull Else dr("Q3") = colArray(4).ToString
    '            If colArray(5).ToString = "" Then dr("Q4") = Convert.DBNull Else dr("Q4") = colArray(5).ToString
    '            If colArray(6).ToString = "" Then dr("Q5") = Convert.DBNull Else dr("Q5") = colArray(6).ToString
    '        End If

    '        'If colArray(3).ToString = "" Then dr("Q2") = Convert.DBNull Else dr("Q2") = colArray(3).ToString
    '        'If colArray(4).ToString = "" Then dr("Q3") = Convert.DBNull Else dr("Q3") = colArray(4).ToString
    '        'If colArray(5).ToString = "" Then dr("Q4") = Convert.DBNull Else dr("Q4") = colArray(5).ToString
    '        'If colArray(6).ToString = "" Then dr("Q5") = Convert.DBNull Else dr("Q5") = colArray(6).ToString
    '        If colArray(7).ToString = "" Then dr("Q6") = Convert.DBNull Else dr("Q6") = colArray(7).ToString
    '        If sm.UserInfo.Years <= 2007 Then
    '            If colArray(8).ToString = "" Then dr("Q7") = Convert.DBNull Else dr("Q7") = colArray(8).ToString
    '            If colArray(9).ToString = "" Then dr("Q8_1_Note") = Convert.DBNull Else dr("Q8_1_Note") = Left(colArray(9).ToString, 100)
    '            If colArray(10).ToString = "" Then dr("Q8_2_Note") = Convert.DBNull Else dr("Q8_2_Note") = Left(colArray(10).ToString, 100)
    '            If colArray(11).ToString = "" Then dr("Q8_3_Note") = Convert.DBNull Else dr("Q8_3_Note") = Left(colArray(11).ToString, 100)
    '            If colArray(12).ToString = "" Then dr("Q9_1_Note") = Convert.DBNull Else dr("Q9_1_Note") = Left(colArray(12).ToString, 100)
    '            If colArray(13).ToString = "" Then dr("Q9_2_Note") = Convert.DBNull Else dr("Q9_2_Note") = Left(colArray(13).ToString, 100)
    '            If colArray(14).ToString = "" Then dr("Q9_3_Note") = Convert.DBNull Else dr("Q9_3_Note") = Left(colArray(14).ToString, 100)
    '            If colArray(15).ToString = "" Then dr("BusName") = Convert.DBNull Else dr("BusName") = Left(colArray(15).ToString, 100)
    '            If colArray(16).ToString = "" Then dr("Q10") = Convert.DBNull Else dr("Q10") = colArray(16).ToString
    '            If colArray(17).ToString = "" Then dr("Q11") = Convert.DBNull Else dr("Q11") = colArray(17).ToString
    '            If colArray(18).ToString = "" Then dr("Q12") = Convert.DBNull Else dr("Q12") = colArray(18).ToString
    '            If colArray(19).ToString = "" Then dr("Q13") = Convert.DBNull Else dr("Q13") = colArray(19).ToString
    '            If colArray(20).ToString = "" Then dr("Q14") = Convert.DBNull Else dr("Q14") = colArray(20).ToString
    '            If colArray(21).ToString = "" Then dr("Q15") = Convert.DBNull Else dr("Q15") = colArray(21).ToString

    '        ElseIf sm.UserInfo.Years = 2008 Then

    '            If colArray(8).ToString = "" Then dr("Q7") = Convert.DBNull Else dr("Q7") = colArray(8).ToString
    '            If colArray(9).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(9).ToString
    '            If colArray(10).ToString = "" Then dr("Q9_1_Note") = Convert.DBNull Else dr("Q9_1_Note") = Left(colArray(10).ToString, 100)
    '            If colArray(11).ToString = "" Then dr("Q9_2_Note") = Convert.DBNull Else dr("Q9_2_Note") = Left(colArray(11).ToString, 100)
    '            If colArray(12).ToString = "" Then dr("Q9_3_Note") = Convert.DBNull Else dr("Q9_3_Note") = Left(colArray(12).ToString, 100)
    '            If colArray(13).ToString = "" Then dr("Q10_1_Note") = Convert.DBNull Else dr("Q10_1_Note") = Left(colArray(13).ToString, 100)
    '            If colArray(14).ToString = "" Then dr("Q10_2_Note") = Convert.DBNull Else dr("Q10_2_Note") = Left(colArray(14).ToString, 100)
    '            If colArray(15).ToString = "" Then dr("Q10_3_Note") = Convert.DBNull Else dr("Q10_3_Note") = Left(colArray(15).ToString, 100)
    '            If colArray(16).ToString = "" Then dr("BusName") = Convert.DBNull Else dr("BusName") = Left(colArray(16).ToString, 100)
    '            If colArray(17).ToString = "" Then dr("Q11") = Convert.DBNull Else dr("Q11") = colArray(17).ToString
    '            If colArray(18).ToString = "" Then dr("Q12") = Convert.DBNull Else dr("Q12") = colArray(18).ToString
    '            If colArray(19).ToString = "" Then dr("Q13") = Convert.DBNull Else dr("Q13") = colArray(19).ToString
    '            If colArray(20).ToString = "" Then dr("Q14") = Convert.DBNull Else dr("Q14") = colArray(20).ToString
    '            If colArray(21).ToString = "" Then dr("Q15") = Convert.DBNull Else dr("Q15") = colArray(21).ToString
    '            If colArray(22).ToString = "" Then dr("Q16") = Convert.DBNull Else dr("Q16") = colArray(22).ToString

    '        ElseIf sm.UserInfo.Years > 2008 Then

    '            '2012年的版本
    '            If sm.UserInfo.Years >= 2012 Then
    '                If colArray(8).ToString = "" Then dr("Q6_7") = Convert.DBNull Else dr("Q6_7") = colArray(8).ToString
    '                If colArray(9).ToString = "" Then dr("Q6_8") = Convert.DBNull Else dr("Q6_8") = colArray(9).ToString
    '                If colArray(10).ToString = "" Then dr("Q7") = Convert.DBNull Else dr("Q7") = colArray(10).ToString
    '                If colArray(11).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(11).ToString

    '                If Q9_1_Note = "" Then dr("Q9_1_Note") = Convert.DBNull Else dr("Q9_1_Note") = Left(Q9_1_Note, 100)
    '                If Q9_2_Note = "" Then dr("Q9_2_Note") = Convert.DBNull Else dr("Q9_2_Note") = Left(Q9_2_Note, 100)
    '                If Q9_3_Note = "" Then dr("Q9_3_Note") = Convert.DBNull Else dr("Q9_3_Note") = Left(Q9_3_Note, 100)
    '                If Q10_1_Note = "" Then dr("Q10_1_Note") = Convert.DBNull Else dr("Q10_1_Note") = Left(Q10_1_Note, 100)
    '                If Q10_2_Note = "" Then dr("Q10_2_Note") = Convert.DBNull Else dr("Q10_2_Note") = Left(Q10_2_Note, 100)
    '                If Q10_3_Note = "" Then dr("Q10_3_Note") = Convert.DBNull Else dr("Q10_3_Note") = Left(Q10_3_Note, 100)
    '            Else
    '                If colArray(8).ToString = "" Then dr("Q7") = Convert.DBNull Else dr("Q7") = colArray(8).ToString
    '                If colArray(9).ToString = "" Then dr("Q8") = Convert.DBNull Else dr("Q8") = colArray(9).ToString

    '                If Q9_1_Note = "" Then dr("Q9_1_Note") = Convert.DBNull Else dr("Q9_1_Note") = Left(Q9_1_Note, 100)
    '                If Q9_2_Note = "" Then dr("Q9_2_Note") = Convert.DBNull Else dr("Q9_2_Note") = Left(Q9_2_Note, 100)
    '                If Q9_3_Note = "" Then dr("Q9_3_Note") = Convert.DBNull Else dr("Q9_3_Note") = Left(Q9_3_Note, 100)
    '                If Q10_1_Note = "" Then dr("Q10_1_Note") = Convert.DBNull Else dr("Q10_1_Note") = Left(Q10_1_Note, 100)
    '                If Q10_2_Note = "" Then dr("Q10_2_Note") = Convert.DBNull Else dr("Q10_2_Note") = Left(Q10_2_Note, 100)
    '                If Q10_3_Note = "" Then dr("Q10_3_Note") = Convert.DBNull Else dr("Q10_3_Note") = Left(Q10_3_Note, 100)
    '            End If

    '        End If

    '        'If colArray(17).ToString = "" Then dr("Q11") = Convert.DBNull Else dr("Q11") = colArray(17).ToString
    '        'If colArray(18).ToString = "" Then dr("Q12") = Convert.DBNull Else dr("Q12") = colArray(18).ToString
    '        'If colArray(19).ToString = "" Then dr("Q13") = Convert.DBNull Else dr("Q13") = colArray(19).ToString
    '        'If colArray(20).ToString = "" Then dr("Q14") = Convert.DBNull Else dr("Q14") = colArray(20).ToString
    '        'If colArray(21).ToString = "" Then dr("Q15") = Convert.DBNull Else dr("Q15") = colArray(21).ToString
    '        dr("ModifyAcct") = sm.UserInfo.UserID
    '        dr("ModifyDate") = CDate(FillFormDate) 'Now

    '        DbAccess.UpdateDataTable(dt, da, trans)
    '        DbAccess.CommitTrans(trans)
    '    Catch ex As Exception
    '        DbAccess.RollbackTrans(trans)
    '        sr.Close()
    '        srr.Close()
    '        Throw ex
    '    End Try
    'End Sub

#End Region
End Class