Partial Class SD_15_007
    Inherits AuthBasePage

    'SD_15_007_R_Prt (2017)(不含) 前
    'SD_15_007_R_Prt2 (2017)(含)後

    'SD_15_007_R11~SD_15_007_R24.aspx :(SD_15_007_R_Prt.aspx)

    'Const cst_rpt性別 As String = "SD_15_007_R11"
    'Const cst_rpt年齡 As String = "SD_15_007_R12"
    'Const cst_rpt教育程度 As String = "SD_15_007_R13"
    'Const cst_rpt身分別 As String = "SD_15_007_R14"
    'Const cst_rpt工作年資 As String = "SD_15_007_R15"
    'Const cst_rpt地理分佈 As String = "SD_15_007_R16"
    'Const cst_rpt公司行業別 As String = "SD_15_007_R17"
    'Const cst_rpt公司規模 As String = "SD_15_007_R18"
    'Const cst_rpt參訓動機 As String = "SD_15_007_R19"
    'Const cst_rpt訓後動向 As String = "SD_15_007_R20"
    'Const cst_rpt參訓單位類別 As String = "SD_15_007_R21"
    'Const cst_rpt參加課程職能別 As String = "SD_15_007_R22"
    'Const cst_rpt參加課程型態別 As String = "SD_15_007_R23"

    Dim s_gTITLE3 As String() = {"產投", "自主", "小計"}

    '新2009年後
    Public Shared Sub Get_FuncList1(ByRef obj As ListControl)
        With obj
            If TypeOf obj Is DropDownList Then
                '.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                .Items.Insert(0, New ListItem("性別", "SD_15_007_R11"))
                .Items.Insert(1, New ListItem("年齡", "SD_15_007_R12"))
                .Items.Insert(2, New ListItem("教育程度", "SD_15_007_R13"))
                .Items.Insert(3, New ListItem("身分別", "SD_15_007_R14"))
                .Items.Insert(4, New ListItem("工作年資", "SD_15_007_R15"))
                .Items.Insert(5, New ListItem("地理分佈", "SD_15_007_R16"))
                .Items.Insert(6, New ListItem("公司行業別", "SD_15_007_R17"))
                .Items.Insert(7, New ListItem("公司規模", "SD_15_007_R18"))
                .Items.Insert(8, New ListItem("參訓動機", "SD_15_007_R19"))

                .Items.Insert(9, New ListItem("訓後動向", "SD_15_007_R20"))
                .Items.Insert(10, New ListItem("參訓單位類別", "SD_15_007_R21"))
                .Items.Insert(11, New ListItem("參加課程職能別", "SD_15_007_R22"))
                .Items.Insert(12, New ListItem("參加課程型態別", "SD_15_007_R23"))
                '.Items.Insert(13, New ListItem("訓練業別", "SD_15_007_R24"))
            End If
        End With
    End Sub

    Dim objconn As SqlConnection = Nothing

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
        '檢查Session是否存在 End

        '訓練機構
        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
            Common.SetListItem(SearchPlan, "A")

            yearlist = TIMS.GetSyear(yearlist)
            'Common.SetListItem(yearlist, Year(Now))
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            trPlanKind.Style("display") = "none"
            trPackageType.Style("display") = "none"
            '54:充電起飛計畫（在職）判斷方式
            If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                trPackageType.Style("display") = TIMS.cst_inline1 '"inline"
            Else
                '28:產業人才投資方案
                '計畫範圍 產投
                If sm.UserInfo.Years >= 2008 Then
                    trPlanKind.Style("display") = TIMS.cst_inline1 '"inline"
                End If
            End If

            'If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '    Get_FuncList1(FuncID)
            'Else
            '    If sm.UserInfo.Years < 2008 Then
            '        Get_FuncList(FuncID)
            '    Else
            '        Get_FuncList1(FuncID)
            '    End If
            'End If
            Call Get_FuncList1(FuncID)

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button5_Click(sender, e)
            End If
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        Years.Value = sm.UserInfo.Years

        Hid_LID.Value = sm.UserInfo.LID

        '跨年度查詢功能:False '分署(中心)、署(局) 才可使用
        Dim flag_can_over_years As Boolean = If(sm.UserInfo.LID <> 2, True, False)

        Dim Gv_yearlist As String = TIMS.GetListValue(yearlist) '取年度選擇值
        Hid_YEARS.Value = If(Gv_yearlist <> "", Gv_yearlist, sm.UserInfo.Years.ToString()) '必要年度傳入client
        Dim flag_selected_year As Boolean = (Gv_yearlist <> "" AndAlso Gv_yearlist <> sm.UserInfo.Years)
        Dim s_selected_year_Js As String = If(flag_selected_year, String.Concat("?selected_year=", Gv_yearlist), "")

        'flag_can_over_years (非委訓單位)'分署(中心)、署(局) 才可使用
        'flag_can_over_years = True '跨年度查詢功能:True
        tr_yearlist.Visible = If(flag_can_over_years, True, False) '委訓單位隱藏 
        trSTDate12.Visible = If(flag_can_over_years, True, False)
        trFTDate12.Visible = If(flag_can_over_years, True, False)
        BtnPrint3.Visible = If(flag_can_over_years, True, False)
        btnPrint4.Visible = If(flag_can_over_years, True, False)
        btnExport5.Visible = If(sm.UserInfo.LID = 0, True, False)
        lab_yearlist.Text = If(sm.UserInfo.LID = 0, "(匯出調查統計分析表)年度不選時，開訓期間為必填!", "")

        If flag_can_over_years Then
            Dim s_javascript_openOrg_FMT1 As String = String.Concat("javascript:openOrg('../../Common/LevOrg{0}.aspx", If(flag_selected_year, String.Concat("?selected_year=", Gv_yearlist), ""), "');")
            Button2.Attributes("onclick") = String.Format(s_javascript_openOrg_FMT1, If(flag_can_over_years OrElse sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        Else
            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
            Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End If

        Button1.Attributes("onclick") = "javascript:return print();"
    End Sub

#Region "NO USE"
    '舊2008年之前
    'Public Shared Sub Get_FuncList(ByRef obj As ListControl)
    '    With obj
    '        If TypeOf obj Is DropDownList Then
    '            .Items.Insert(0, New ListItem("性別", "SD_15_007_R1"))
    '            .Items.Insert(1, New ListItem("年齡", "SD_15_007_R2"))
    '            .Items.Insert(2, New ListItem("教育程度", "SD_15_007_R3"))
    '            .Items.Insert(3, New ListItem("身分別", "SD_15_007_R"))
    '            .Items.Insert(4, New ListItem("工作年資", "SD_15_007_R5"))
    '            .Items.Insert(5, New ListItem("地理分佈", "SD_15_007_R6"))
    '            .Items.Insert(6, New ListItem("公司行業別", "SD_15_007_R7"))
    '            .Items.Insert(7, New ListItem("公司規模", "SD_15_007_R8"))
    '            .Items.Insert(8, New ListItem("參訓動機", "SD_15_007_R9"))
    '            .Items.Insert(9, New ListItem("訓後動向", "SD_15_007_R10"))
    '        End If
    '    End With
    'End Sub

    '.Items.Insert(11, New ListItem("參加課程業別", "SD_15_007_R12"))
    '.Items.Insert(12, New ListItem("參加課程職類別", "SD_15_007_R13"))
#End Region

    '列印
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '28:產業人才投資方案
        Dim SearchPlan1 As String = TIMS.GetListValue(SearchPlan)
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then SearchPlan1 = ""
        If (SearchPlan1 = "A") Then SearchPlan1 = ""

        Dim strScript As String
        If RIDValue.Value.ToString = "A" Then RIDValue.Value = ""

        Dim sPackType As String = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If v_PackageType <> "A" Then sPackType = v_PackageType ' PackageType.SelectedValue
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sPrintAspx As String = "SD_15_007_R_Prt.aspx?" '2016 OLD
        If v_yearlist >= "2017" Then sPrintAspx = "SD_15_007_R_Prt2.aspx?" '2017

        Dim MyValue As String = ""
        MyValue = "filename=" & TIMS.ClearSQM(FuncID.SelectedValue)
        MyValue += "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue += "&Years=" & v_yearlist 'yearlist.SelectedValue
        MyValue += "&OCID=" & OCIDValue1.Value
        MyValue += "&RID=" & RIDValue.Value
        MyValue += "&SearchPlan=" & SearchPlan1 '"",G,W
        MyValue += "&PackageType=" & sPackType '"",2,3
        'MyValue += "&STDate1=" & TIMS.cdate3(STDate1.Text)
        'MyValue += "&STDate2=" & TIMS.cdate3(STDate2.Text)
        'MyValue += "&FTDate1=" & TIMS.cdate3(FTDate1.Text)
        'MyValue += "&FTDate2=" & TIMS.cdate3(FTDate2.Text)

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "window.open('" & sPrintAspx & MyValue & "');"
        strScript += "</script>"
        Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = Convert.ToString(dr("trainname"))
        OCID1.Text = Convert.ToString(dr("classname"))
        TMIDValue1.Value = Convert.ToString(dr("trainid"))
        OCIDValue1.Value = Convert.ToString(dr("ocid"))
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist = "" Then
            Errmsg += "請選擇年度" & vbCrLf
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Errmsg += "請選擇 職類/班別" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '搜㝷('2016 old)
    Function Search_Query3() As String
        Dim Rst As String = ""

        '28:產業人才投資方案
        Dim SearchPlan1 As String = TIMS.GetListValue(SearchPlan)
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then SearchPlan1 = ""
        If (SearchPlan1 = "A") Then SearchPlan1 = ""

        Dim sPackType As String = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If v_PackageType <> "A" Then sPackType = v_PackageType ' PackageType.SelectedValue
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sql As String = ""
        sql &= " select oo.orgname " & vbCrLf
        'sql &= " ,cc.classcname + ' 第' + cc.cycltype + '期' classcname" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.ocid " & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        'Sql += " ,ss.idno" & vbCrLf
        sql &= " ,ss.name as cname" & vbCrLf
        sql &= " ,cs.SOCID " & vbCrLf

        sql &= " ,fc1.Q1_1" & vbCrLf
        sql &= " ,fc1.Q1_2" & vbCrLf
        sql &= " ,fc1.Q1_3" & vbCrLf
        sql &= " ,fc1.Q1_4" & vbCrLf
        sql &= " ,fc1.Q1_5" & vbCrLf
        sql &= " ,fc1.Q1_6" & vbCrLf

        sql &= " ,fc1.Q2_1" & vbCrLf
        sql &= " ,fc1.Q2_2" & vbCrLf
        sql &= " ,fc1.Q2_3" & vbCrLf
        sql &= " ,fc1.Q2_4" & vbCrLf
        sql &= " ,fc1.Q2_5" & vbCrLf

        sql &= " ,fc1.Q3_1" & vbCrLf
        sql &= " ,fc1.Q3_2" & vbCrLf
        sql &= " ,fc1.Q3_3" & vbCrLf

        sql &= " ,fc1.Q4" & vbCrLf
        sql &= " ,fc1.Q5" & vbCrLf
        'Sql += " ,fc1.Q5_Note_News" & vbCrLf
        'Sql += " ,fc1.Q5_Note_Other" & vbCrLf
        sql &= " ,fc1.Q6" & vbCrLf
        'Sql += " ,fc1.Q6_Note1" & vbCrLf
        'Sql += " ,fc1.Q6_Note2" & vbCrLf
        sql &= " ,fc1.Q7" & vbCrLf
        sql &= " ,fc1.Q7_8" & vbCrLf
        sql &= " ,fc1.Q7_9" & vbCrLf

        sql &= " ,fc1.Q8" & vbCrLf
        'Sql += " ,fc1.Q9_Note" & vbCrLf
        sql &= " ,fc1.Q9" & vbCrLf
        sql &= " ,fc1.Q9_1" & vbCrLf
        sql &= " ,fc1.Q9_2" & vbCrLf
        sql &= " ,fc1.Q9_3" & vbCrLf
        sql &= " ,fc1.Q10" & vbCrLf

        sql &= " FROM dbo.CLASS_CLASSINFO CC" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO PP ON PP.PLANID =CC.PLANID  AND PP.COMIDNO= CC.COMIDNO AND PP.SEQNO = CC.SEQNO" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN IP ON IP.PLANID =CC.PLANID " & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO OO ON OO.COMIDNO =CC.COMIDNO " & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS CS ON CS.OCID =CC.OCID " & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO SS ON SS.SID =CS.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_QUESTIONFAC fc1 ON fc1.SOCID=CS.SOCID " & vbCrLf
        sql &= " WHERE cs.StudStatus NOT IN (2,3) " & vbCrLf
        sql &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        'Sql += " and ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        sql &= " and ip.Years='" & v_yearlist & "'" & vbCrLf
        sql &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        If SearchPlan1 <> "" Then
            sql &= " and oo.OrgKind2='" & SearchPlan1 & "'" & vbCrLf
        End If
        If sPackType <> "" Then
            sql &= " and pp.PackageType='" & sPackType & "'" & vbCrLf
        End If
        sql &= " order by cs.SOCID" & vbCrLf

        Rst = sql
        Return Rst
    End Function

    '列印明細('2016 old)
    Sub sUtl_Export3()
        'Dim tConn As SqlConnection
        'tConn = DbAccess.GetConnection()

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(Search_Query3, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        dt.DefaultView.Sort = "SOCID"
        dt = TIMS.dv2dt(dt.DefaultView)

        Dim sFileName1 As String = "受訓學員意見調查表" & OCIDValue1.Value
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table>")

        Dim num As Integer = 0
        Dim ExportStr As String = ""           '建立輸出文字

        ExportStr = "訓練單位名稱"
        ExportStr += vbTab & "課程名稱"
        ExportStr += vbTab & "課程代碼"
        ExportStr += vbTab & "開訓日期"
        ExportStr += vbTab & "結訓日期"
        ExportStr += vbTab & "學員姓名"

        ExportStr += vbTab & "1-1.課程內容與工作性質是否相關"
        ExportStr += vbTab & "1-2.課程名稱是否適當"
        ExportStr += vbTab & "1-3.教材內容是否適當"
        ExportStr += vbTab & "1-4.本項訓練發給教材情形"
        ExportStr += vbTab & "1-5.發給方式"
        ExportStr += vbTab & "1-6.訓練時數是否適當"

        ExportStr += vbTab & "2-1.術科時數是否適當"
        ExportStr += vbTab & "2-2.術科內容是否適當"
        ExportStr += vbTab & "2-3.術科操作解說是否充分"
        ExportStr += vbTab & "2-4.訓練設備是否充足"
        ExportStr += vbTab & "2-5.訓練設備現狀"

        ExportStr += vbTab & "3-1.教師的教學態度"
        ExportStr += vbTab & "3-2.教師師資的教學方法或技巧"
        ExportStr += vbTab & "3-3.講授課程時間控制是否適當"

        ExportStr += vbTab & "4.對整體課程瞭解的程度"
        ExportStr += vbTab & "5.獲得招訓消息的來源"
        ExportStr += vbTab & "6.自行繳納費用負擔方式"
        ExportStr += vbTab & "7.你能掌握訓練教授知識或技能 "
        ExportStr += vbTab & "8.參加本項課程訓練後，你有把握自己能所學的知識應用到工作上"
        ExportStr += vbTab & "9.完成訓練後，你願意找機會將所學的知識／技能應用在工作中"

        ExportStr += vbTab & "10.對訓練單位的行政服務滿意度"
        ExportStr += vbTab & "11-1.瞭解補助對象"
        ExportStr += vbTab & "11-2.瞭解補助經費標準"
        ExportStr += vbTab & "11-3.瞭解補助流程"

        ExportStr += vbTab & "12.你對於產業人才投資方案是否滿意"

        'ExportStr += vbTab & vbCrLf
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        strHTML &= TIMS.Get_TABLETR(Replace(ExportStr, vbTab, ","))

        '建立資料面
        For Each dr As DataRow In dt.DefaultView.Table.Rows
            num += 1
            ExportStr = ""
            ExportStr += dr("orgname").ToString
            ExportStr += vbTab & dr("classcname").ToString '"課程名稱"
            ExportStr += vbTab & dr("ocid").ToString '"課程代碼"
            ExportStr += vbTab & dr("stdate").ToString '"開訓日期"
            ExportStr += vbTab & dr("ftdate").ToString '"結訓日期"
            ExportStr += vbTab & dr("cname").ToString '"學員姓名"

            ExportStr += vbTab & dr("Q1_1").ToString '"1-1.課程內容與工作性質是否相關"
            ExportStr += vbTab & dr("Q1_2").ToString '"1-2.課程名稱是否適當"
            ExportStr += vbTab & dr("Q1_3").ToString '"1-3.教材內容是否適當"
            ExportStr += vbTab & dr("Q1_4").ToString '"1-4.本項訓練發給教材情形"
            ExportStr += vbTab & dr("Q1_5").ToString '"1-5.發給方式"
            ExportStr += vbTab & dr("Q1_6").ToString '"1-6.訓練時數是否適當"

            ExportStr += vbTab & dr("Q2_1").ToString '"2-1.術科時數是否適當"
            ExportStr += vbTab & dr("Q2_2").ToString '"2-2.術科內容是否適當"
            ExportStr += vbTab & dr("Q2_3").ToString '"2-3.術科操作解說是否充分"
            ExportStr += vbTab & dr("Q2_4").ToString '"2-4.訓練設備是否充足"
            ExportStr += vbTab & dr("Q2_5").ToString '"2-5.訓練設備現狀"

            ExportStr += vbTab & dr("Q3_1").ToString '"3-1.教師的教學態度"
            ExportStr += vbTab & dr("Q3_2").ToString '"3-2.教師師資的教學方法或技巧"
            ExportStr += vbTab & dr("Q3_3").ToString '"3-3.講授課程時間控制是否適當"

            ExportStr += vbTab & dr("Q4").ToString '"4.對整體課程瞭解的程度"
            ExportStr += vbTab & dr("Q5").ToString '"5.獲得招訓消息的來源"
            ExportStr += vbTab & dr("Q6").ToString '"6.自行繳納費用負擔方式"
            ExportStr += vbTab & dr("Q7").ToString '"7.對於訓練課程所教授知識或技能的掌握程度"
            ExportStr += vbTab & dr("Q7_8").ToString '"8.參加本項課程訓練後，你有把握自己能所學的知識應用到工作上"
            ExportStr += vbTab & dr("Q7_9").ToString '"9.完成訓練後，你願意找機會將所學的知識／技能應用在工作中"

            ExportStr += vbTab & dr("Q8").ToString '"10.對訓練單位的行政服務滿意度"
            ExportStr += vbTab & dr("Q9_1").ToString '"11-1.瞭解補助對象"
            ExportStr += vbTab & dr("Q9_2").ToString '"11-2.瞭解補助經費標準"
            ExportStr += vbTab & dr("Q9_3").ToString '"11-3.瞭解補助流程"

            ExportStr += vbTab & dr("Q10").ToString '"12.你對於產業人才投資方案是否滿意"

            'ExportStr += vbCrLf
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
            strHTML &= TIMS.Get_TABLETR(Replace(ExportStr, vbTab, ","))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '列印其他意見('2016 old) STUD_QUESTIONFAC
    Sub sUtl_Export3b()
        'Dim num As Integer = 0
        'Dim ExportStr As String = ""  '建立輸出文字
        Dim Errmsg As String = ""

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist = "" Then
            Errmsg += "請選擇年度" & vbCrLf
        End If

        If Trim(STDate1.Text) <> "" Then
            STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            STDate1.Text = ""
        End If

        If Trim(STDate2.Text) <> "" Then
            STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            STDate2.Text = ""
        End If

        If Trim(FTDate1.Text) <> "" Then
            FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate1.Text = ""
        End If

        If Trim(FTDate2.Text) <> "" Then
            FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate2.Text = ""
        End If

        Dim xFlag1 As Boolean = False
        If Trim(STDate1.Text) <> "" AndAlso Trim(STDate2.Text) <> "" Then
            xFlag1 = True
        End If
        If Trim(FTDate1.Text) <> "" AndAlso Trim(FTDate2.Text) <> "" Then
            xFlag1 = True
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If OCIDValue1.Value = "" Then
            If Not xFlag1 Then
                Errmsg += "開訓區間 或 結訓區間 為必填資訊" & vbCrLf
            End If
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '28:產業人才投資方案
        Dim SearchPlan1 As String = TIMS.GetListValue(SearchPlan)
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then SearchPlan1 = ""
        If (SearchPlan1 = "A") Then SearchPlan1 = ""

        Dim sPackType As String = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If v_PackageType <> "A" Then sPackType = v_PackageType ' PackageType.SelectedValue

        Dim sql As String = ""
        sql &= " select oo.orgname " & vbCrLf
        'sql &= " ,cc.classcname + ' 第' + cc.cycltype + '期' classcname" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.ocid " & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        'Sql += " ,ss.idno" & vbCrLf
        sql &= " ,ss.name as cname" & vbCrLf
        sql &= " ,cs.SOCID " & vbCrLf
        sql &= " ,fc1.Q12" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO CC" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO PP ON PP.PLANID =CC.PLANID  AND PP.COMIDNO= CC.COMIDNO AND PP.SEQNO = CC.SEQNO" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN IP ON IP.PLANID =CC.PLANID " & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO OO ON OO.COMIDNO =CC.COMIDNO " & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS CS ON CS.OCID =CC.OCID " & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO SS ON SS.SID =CS.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_QUESTIONFAC fc1 ON fc1.SOCID=CS.SOCID " & vbCrLf
        sql &= " WHERE fc1.Q12 IS NOT NULL" & vbCrLf
        sql &= " and cs.StudStatus NOT IN (2,3) " & vbCrLf
        sql &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " and ip.Years='" & v_yearlist & "'" & vbCrLf
        If OCIDValue1.Value <> "" Then
            sql &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        Else
            '未選擇班級
            '階層代碼 0.署(局) 1.分署(中心) 2.委訓 【SELECT LID ,COUNT(1) CNT FROM AUTH_ACCOUNT GROUP BY LID ORDER BY 1】
            Select Case sm.UserInfo.LID
                Case "0"
                    If Len(RIDValue.Value) > 1 Then
                        '指定單位
                        sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                    End If
                    If Len(RIDValue.Value) = 1 AndAlso RIDValue.Value <> "A" Then
                        '指定分署
                        sql &= " and ip.DistID ='" & TIMS.Get_DistID_RID(RIDValue.Value, objconn) & "'" & vbCrLf
                    End If
                Case "1" '1.分署(中心)
                    '限定登入分署
                    sql &= " and ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                    If Len(RIDValue.Value) > 1 Then
                        '指定單位
                        sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                    End If
                Case Else '2.委訓
                    '限定登入分署
                    sql &= " and ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                    If RIDValue.Value <> "" Then
                        sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                    Else
                        sql &= " and cc.RID ='" & sm.UserInfo.RID & "'" & vbCrLf
                    End If
            End Select
        End If

        If SearchPlan1 <> "" Then
            sql &= " and oo.OrgKind2='" & SearchPlan1 & "'" & vbCrLf
        End If
        If sPackType <> "" Then
            sql &= " and pp.PackageType='" & sPackType & "'" & vbCrLf
        End If
        '開訓區間
        If STDate1.Text <> "" Then
            sql &= " and cc.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf '" & STDate1.Text & "'" & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql &= " and cc.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
        End If
        '結訓區間
        If FTDate1.Text <> "" Then
            sql &= " and cc.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql &= " and cc.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
        End If
        sql &= " order by cs.socid" & vbCrLf

        Dim CPdt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
        If CPdt1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無有效資料，請重新查詢!!")
            Exit Sub
        End If

        Dim sFileName1 As String = "受訓學員意見調查" & OCIDValue1.Value
        Dim strSTYLE As String = ""
        strSTYLE &= "<style> .text { mso-number-format:\@; text-align@center;} </style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table border='1' cellspacing='0' style='font-family:標楷體;border-collapse@collapse;border@solid thin #000000;'>"
        '欄位名稱列
        strHTML &= "<tr>"
        strHTML &= "<td>訓練單位名稱</td>"
        strHTML &= "<td>課程名稱</td>"
        strHTML &= "<td>課程代碼</td>"
        strHTML &= "<td>開訓日期</td>"
        strHTML &= "<td>結訓日期</td>"
        strHTML &= "<td>學員姓名</td>"
        strHTML &= "<td>其他建議</td>"
        strHTML &= "</tr>"
        '資料列
        'Dim iRow As Integer = 0
        For Each drv As DataRow In CPdt1.Rows
            'iRow += 1
            strHTML &= "<tr>"
            'strHTML &= "<td>" & iRow & "</td>"
            strHTML &= "<td>" & Convert.ToString(drv("orgname")) & "</td>" '單位名稱
            strHTML &= "<td>" & Convert.ToString(drv("classcname")) & "</td>" '課程名稱
            strHTML &= "<td>" & Convert.ToString(drv("ocid")) & "</td>" '課程代碼
            strHTML &= "<td>" & Convert.ToString(drv("STDATE")) & "</td>" '開訓日期
            strHTML &= "<td>" & Convert.ToString(drv("FTDATE")) & "</td>" '結訓日期
            strHTML &= "<td>" & Convert.ToString(drv("CNAME")) & "</td>" '學員姓名
            strHTML &= "<td>" & Convert.ToString(drv("Q12")) & "</td>" '其他建議
            strHTML &= "</tr>" & vbCrLf
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(strHTML))
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '搜㝷('2017) STUD_QUESTIONFAC2
    Function Search_Query4() As String
        Dim Rst As String = ""

        '28:產業人才投資方案
        Dim SearchPlan1 As String = TIMS.GetListValue(SearchPlan)
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then SearchPlan1 = ""
        If (SearchPlan1 = "A") Then SearchPlan1 = ""

        Dim sPackType As String = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If v_PackageType <> "A" Then sPackType = v_PackageType ' PackageType.SelectedValue

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sql As String = ""
        sql &= " select oo.orgname " & vbCrLf
        'sql &= " ,cc.classcname + '第' + cc.cycltype + '期' classcname" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.ocid " & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        'Sql += " ,ss.idno" & vbCrLf
        sql &= " ,ss.name cname" & vbCrLf
        sql &= " ,cs.SOCID " & vbCrLf

        sql &= " ,fc2.A1_1" & vbCrLf
        sql &= " ,fc2.A1_2" & vbCrLf
        sql &= " ,fc2.A1_3" & vbCrLf
        sql &= " ,fc2.A1_4" & vbCrLf
        sql &= " ,fc2.A1_5" & vbCrLf
        sql &= " ,fc2.A1_6" & vbCrLf
        sql &= " ,fc2.A1_7" & vbCrLf
        sql &= " ,fc2.A1_8" & vbCrLf
        sql &= " ,fc2.A1_9" & vbCrLf
        sql &= " ,fc2.A1_10" & vbCrLf
        sql &= " ,fc2.A2" & vbCrLf
        sql &= " ,fc2.A3" & vbCrLf
        sql &= " ,fc2.A4" & vbCrLf
        sql &= " ,fc2.A5" & vbCrLf
        sql &= " ,fc2.A6" & vbCrLf
        sql &= " ,fc2.A7" & vbCrLf
        sql &= " ,fc2.B11" & vbCrLf
        sql &= " ,fc2.B12" & vbCrLf
        sql &= " ,fc2.B13" & vbCrLf
        sql &= " ,fc2.B14" & vbCrLf
        sql &= " ,fc2.B15" & vbCrLf
        sql &= " ,fc2.B21" & vbCrLf
        sql &= " ,fc2.B22" & vbCrLf
        sql &= " ,fc2.B23" & vbCrLf
        sql &= " ,fc2.B31" & vbCrLf
        sql &= " ,fc2.B32" & vbCrLf
        sql &= " ,fc2.B41" & vbCrLf
        sql &= " ,fc2.B42" & vbCrLf
        sql &= " ,fc2.B43" & vbCrLf
        sql &= " ,fc2.B44" & vbCrLf
        sql &= " ,fc2.B51" & vbCrLf
        sql &= " ,fc2.B61" & vbCrLf
        sql &= " ,fc2.B62" & vbCrLf
        sql &= " ,fc2.B63" & vbCrLf
        sql &= " ,fc2.B71" & vbCrLf
        sql &= " ,fc2.B72" & vbCrLf
        sql &= " ,fc2.B73" & vbCrLf
        sql &= " ,fc2.B74" & vbCrLf
        sql &= " ,fc2.C11" & vbCrLf

        sql &= " FROM dbo.CLASS_CLASSINFO CC" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO PP ON PP.PLANID =CC.PLANID  AND PP.COMIDNO= CC.COMIDNO AND PP.SEQNO = CC.SEQNO" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN IP ON IP.PLANID =CC.PLANID " & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO OO ON OO.COMIDNO =CC.COMIDNO " & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS CS ON CS.OCID =CC.OCID " & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO SS ON SS.SID =CS.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_QUESTIONFAC2 fc2 ON fc2.SOCID=CS.SOCID " & vbCrLf
        sql &= " WHERE cs.StudStatus NOT IN (2,3) " & vbCrLf
        sql &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        'Sql += " and ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        sql &= " and ip.Years='" & v_yearlist & "'" & vbCrLf
        sql &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        If SearchPlan1 <> "" Then
            sql &= " and oo.OrgKind2='" & SearchPlan1 & "'" & vbCrLf
        End If
        If sPackType <> "" Then
            sql &= " and pp.PackageType='" & sPackType & "'" & vbCrLf
        End If
        sql &= " order by cc.OCID,cs.SOCID" & vbCrLf

        'Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        'Dim slogMsg1 As String = ""
        'slogMsg1 &= "##SD_15_007, sql: " & sql & vbCrLf
        'If flag_chktest Then TIMS.writeLog(Me, slogMsg1)
        Rst = sql
        Return Rst
    End Function

    '列印明細('2017)
    Sub sUtl_Export4()
        'Dim tConn As SqlConnection
        'tConn = DbAccess.GetConnection()
        Dim dt As DataTable = DbAccess.GetDataTable(Search_Query4(), objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        dt.DefaultView.Sort = "SOCID"
        dt = TIMS.dv2dt(dt.DefaultView)

        Dim sFileName1 As String = "受訓學員意見調查表" & OCIDValue1.Value
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= (".DecFormat2{mso-number-format:""0.00"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'strHTML &= ("<tr>")

        Dim sPattern As String = ""
        Dim sColumn As String = ""
        sPattern = "訓練單位名稱,課程名稱,課程代碼,開訓日期,結訓日期,學員姓名,1-1-1.您獲得本次課程的訊息來源-本署或分署網站,1-1-2.您獲得本次課程的訊息來源-就業服務中心,1-1-3.您獲得本次課程的訊息來源-訓練單位,1-1-4.您獲得本次課程的訊息來源-搜尋網站,1-1-5.您獲得本次課程的訊息來源-報紙,1-1-6.您獲得本次課程的訊息來源-廣播,1-1-7.您獲得本次課程的訊息來源-電視,1-1-8.您獲得本次課程的訊息來源-親友介紹,1-1-9.您獲得本次課程的訊息來源-社群媒體(ex：臉書、LINE),1-1-10.您獲得本次課程的訊息來源-其他,1-2.參加本次課程的主要原因,1-3.選擇本訓練單位的主要原因,1-4.沒有參加本方案訓練之前，每年參加訓練支出的費用？,1-5.如果沒有補助訓練費用，你每年願意自費參加訓練課程的金額？,1-6.您認為本次課程的訓練費用是否合理？,1-7.結訓後對於工作的規劃？,2-1-1.課程內容符合期望,2-1-2.課程難易安排適當,2-1-3.課程總時數適當,2-1-4.課程符合實務需求,2-1-5.課程符合產業發展趨勢,2-2-1.滿意講師的教學態度,2-2-2.滿意講師的教學方法,2-2-3.滿意講師的課程專業度,2-3-1.對於訓練教材感到滿意,2-3-2.訓練教材能夠輔助課程學習,2-4-1.您對於訓練場地感到滿意,2-4-2.您對於訓練設備感到滿意,2-4-3.您認為實作設備的數量適當,2-4-4.您認為實作設備新穎,2-5.能促進學習效果,2-6-1.您認為在訓練課程中，課程內容能讓您專注,2-6-2.您在完成訓練後，已充份學習訓練課程所教授知識或技能,2-6-3.您在完成訓練後，有學習到新的知識或技能,2-7-1.您對於訓練單位的課程安排與授課情形感到滿意,2-7-2.您對於訓練單位的行政服務感到滿意,2-7-3.您對於產業人才投資方案感到滿意,2-7-4.您認為完成本訓練課程對於目前或未來工作有幫助,3-1.若本訓練課程沒有補助，是否會全額自費參訓？"
        sColumn = "ORGNAME,CLASSCNAME,OCID,STDATE,FTDATE,CNAME,A1_1,A1_2,A1_3,A1_4,A1_5,A1_6,A1_7,A1_8,A1_9,A1_10,A2,A3,A4,A5,A6,A7,B11,B12,B13,B14,B15,B21,B22,B23,B31,B32,B41,B42,B43,B44,B51,B61,B62,B63,B71,B72,B73,B74,C11"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = "<tr>"
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= "<td>" & sPatternA(i) & "</td>" '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dt.DefaultView.Table.Rows
            iNum += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                'Select Case CStr(sColumnA(i))
                '    Case "Phone1", "Phone2", "CellPhone"
                '        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
                '    Case Else
                '        ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
                'End Select
                ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
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

    '列印其他意見('2017)
    Sub sUtl_Export4b()
        Dim Errmsg As String = ""

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist = "" Then
            Errmsg += "請選擇年度" & vbCrLf
        End If

        If Trim(STDate1.Text) <> "" Then
            STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            STDate1.Text = ""
        End If

        If Trim(STDate2.Text) <> "" Then
            STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            STDate2.Text = ""
        End If

        If Trim(FTDate1.Text) <> "" Then
            FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate1.Text = ""
        End If

        If Trim(FTDate2.Text) <> "" Then
            FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate2.Text = ""
        End If

        Dim xFlag1 As Boolean = False
        If Trim(STDate1.Text) <> "" AndAlso Trim(STDate2.Text) <> "" Then
            xFlag1 = True
        End If
        If Trim(FTDate1.Text) <> "" AndAlso Trim(FTDate2.Text) <> "" Then
            xFlag1 = True
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If OCIDValue1.Value = "" Then
            If Not xFlag1 Then
                Errmsg += "開訓區間 或 結訓區間 為必填資訊" & vbCrLf
            End If
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '28:產業人才投資方案
        Dim SearchPlan1 As String = TIMS.GetListValue(SearchPlan)
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then SearchPlan1 = ""
        If (SearchPlan1 = "A") Then SearchPlan1 = ""

        Dim sPackType As String = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If v_PackageType <> "A" Then sPackType = v_PackageType ' PackageType.SelectedValue
        'Dim v_yearlist As String = TIMS.GetListValue(yearlist)

        Dim sql As String = ""
        sql &= " select oo.orgname " & vbCrLf
        'sql &= " ,cc.classcname + ' 第' + cc.cycltype + '期' classcname" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.ocid " & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        'Sql += " ,ss.idno" & vbCrLf
        sql &= " ,ss.name as cname" & vbCrLf
        sql &= " ,cs.SOCID " & vbCrLf
        sql &= " ,fc2.C21_NOTE C21NOTE" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO CC" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO PP ON PP.PLANID =CC.PLANID  AND PP.COMIDNO= CC.COMIDNO AND PP.SEQNO = CC.SEQNO" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN IP ON IP.PLANID =CC.PLANID " & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO OO ON OO.COMIDNO =CC.COMIDNO " & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS CS ON CS.OCID =CC.OCID " & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO SS ON SS.SID =CS.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_QUESTIONFAC2 fc2 ON fc2.SOCID=CS.SOCID " & vbCrLf
        sql &= " WHERE fc2.C21_NOTE IS NOT NULL" & vbCrLf
        sql &= " and cs.StudStatus NOT IN (2,3) " & vbCrLf
        sql &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " and ip.Years='" & v_yearlist & "'" & vbCrLf
        If OCIDValue1.Value <> "" Then
            sql &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        Else
            '未選擇班級
            '階層代碼 0.署(局) 1.分署(中心) 2.委訓 【SELECT LID ,COUNT(1) CNT FROM AUTH_ACCOUNT GROUP BY LID ORDER BY 1】
            Select Case sm.UserInfo.LID
                Case "0"
                    If Len(RIDValue.Value) > 1 Then
                        '指定單位
                        sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                    End If
                    If Len(RIDValue.Value) = 1 AndAlso RIDValue.Value <> "A" Then
                        '指定分署
                        sql &= " and ip.DistID ='" & TIMS.Get_DistID_RID(RIDValue.Value, objconn) & "'" & vbCrLf
                    End If
                Case "1" '1.分署(中心)
                    '限定登入分署
                    sql &= " and ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                    If Len(RIDValue.Value) > 1 Then
                        '指定單位
                        sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                    End If
                Case Else '2.委訓
                    '限定登入分署
                    sql &= " and ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                    If RIDValue.Value <> "" Then
                        sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                    Else
                        sql &= " and cc.RID ='" & sm.UserInfo.RID & "'" & vbCrLf
                    End If
            End Select
        End If
        If SearchPlan1 <> "" Then sql &= " and oo.OrgKind2='" & SearchPlan1 & "'" & vbCrLf
        If sPackType <> "" Then sql &= " and pp.PackageType='" & sPackType & "'" & vbCrLf
        '開訓區間
        If STDate1.Text <> "" Then sql &= " and cc.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf '" & STDate1.Text & "'" & vbCrLf 
        If STDate2.Text <> "" Then sql &= " and cc.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf 
        '結訓區間
        If FTDate1.Text <> "" Then sql &= " and cc.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf 
        If FTDate2.Text <> "" Then sql &= " and cc.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf 
        sql &= " ORDER BY cc.ocid,cs.socid" & vbCrLf

        Dim CPdt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
        If CPdt1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無有效資料，請重新查詢!!")
            Exit Sub
        End If

        Dim sFileName1 As String = "受訓學員意見調查" & OCIDValue1.Value
        Dim strSTYLE As String = ""
        strSTYLE &= "<style> .text { mso-number-format:\@; text-align@center;} </style>"

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= "<table border='1' cellspacing='0' style='font-family:標楷體;border-collapse@collapse;border@solid thin #000000;'>"
        '欄位名稱列
        strHTML &= "<tr>"
        strHTML &= "<td>訓練單位名稱</td>"
        strHTML &= "<td>課程名稱</td>"
        strHTML &= "<td>課程代碼</td>"
        strHTML &= "<td>開訓日期</td>"
        strHTML &= "<td>結訓日期</td>"
        strHTML &= "<td>學員姓名</td>"
        strHTML &= "<td>其他建議</td>"
        strHTML &= "</tr>"

        '資料列
        'Dim iRow As Integer = 0
        For Each drv As DataRow In CPdt1.Rows
            'iRow += 1
            strHTML &= "<tr>"
            'strHTML &= "<td>" & iRow & "</td>"
            strHTML &= "<td>" & Convert.ToString(drv("orgname")) & "</td>" '單位名稱
            strHTML &= "<td>" & Convert.ToString(drv("classcname")) & "</td>" '課程名稱
            strHTML &= "<td>" & Convert.ToString(drv("ocid")) & "</td>" '課程代碼
            strHTML &= "<td>" & Convert.ToString(drv("STDATE")) & "</td>" '開訓日期
            strHTML &= "<td>" & Convert.ToString(drv("FTDATE")) & "</td>" '結訓日期
            strHTML &= "<td>" & Convert.ToString(drv("CNAME")) & "</td>" '學員姓名
            strHTML &= "<td>" & Convert.ToString(drv("C21NOTE")) & "</td>" '其他建議
            strHTML &= "</tr>" & vbCrLf
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '列印明細
    Private Sub BtnPrint3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint3.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Const cst_msg1 As String = "不提供訓練單位 列印明細功能!!"
        '不提供訓練單位 列印明細功能!!
        If Not TIMS.Chk_LID2CanPrint(Me) Then
            Common.MessageBox(Me, cst_msg1)
            Exit Sub
        End If

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim sYearType As String = "2016" '2016 OLD
        If v_yearlist >= "2017" Then sYearType = "2017" '2017
        Select Case sYearType
            Case "2017"
                Call sUtl_Export4()
            Case "2016"
                Call sUtl_Export3()
            Case Else
                Call sUtl_Export3()
        End Select
    End Sub

    '列印其他意見
    Protected Sub btnPrint4_Click(sender As Object, e As EventArgs) Handles btnPrint4.Click
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim sYearType As String = "2016" '2016 OLD
        If v_yearlist >= "2017" Then sYearType = "2017" '2017
        Select Case sYearType
            Case "2017"
                Call sUtl_Export4b()
            Case "2016"
                Call sUtl_Export3b()
            Case Else
                Call sUtl_Export3b()
        End Select
    End Sub

    Protected Sub btnExport5_Click(sender As Object, e As EventArgs) Handles btnExport5.Click
        Dim Errmsg As String = ""

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
        Common.SetListItem(SearchPlan, "A")

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        '開訓區間／'結訓區間
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        FTDate1.Text = TIMS.Cdate3(FTDate1.Text)
        FTDate2.Text = TIMS.Cdate3(FTDate2.Text)

        Dim fg_YEAR1 As Boolean = If(v_yearlist <> "", True, False)
        Dim fg_STDT1 As Boolean = If(STDate1.Text <> "" AndAlso STDate2.Text <> "", True, False)
        Dim fg_STDT2 As Boolean = If(fg_STDT1 AndAlso DateDiff(DateInterval.Day, CDate(STDate1.Text), CDate(STDate2.Text)) < 500, True, False)
        If Not fg_YEAR1 Then
            If Not fg_STDT1 Then
                Errmsg &= "年度不選時，開訓期間為必填" & vbCrLf
            ElseIf Not fg_STDT2 Then
                Errmsg &= "(超過範圍限制)開訓期間計算天數，不可超過500天" & vbCrLf
            End If
        End If
        'If v_yearlist = "" Then Errmsg &= "請選擇年度" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Call sUtl_Export51()
    End Sub

    Function Get_TITLE3(s_TITLE3 As String()) As String
        Dim rst As String = ""
        For i As Integer = 1 To 5
            For Each t3 As String In s_TITLE3
                rst &= String.Concat("<td>", t3, "</td>")
            Next
        Next
        Return rst
    End Function

    ''' <summary>取得問卷題目</summary>
    ''' <returns></returns>
    Function GET_QUESTION_T1() As DataTable
        Dim sSql As String = ""
        sSql &= " SELECT 'B11' BNO,'(一) 訓練課程,1.課程內容符合期望' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B12' BNO,'(一) 訓練課程,2.課程難易安排適當' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B13' BNO,'(一) 訓練課程,3.課程總時數適當' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B14' BNO,'(一) 訓練課程,4.課程符合實務需求' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B15' BNO,'(一) 訓練課程,5.課程符合產業發展趨勢' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        'sSql &= " " & vbCrLf
        sSql &= " UNION SELECT 'B21' BNO,'(二)講師,1.滿意講師的教學態度' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B22' BNO,'(二)講師,2.滿意講師的教學方法' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B23' BNO,'(二)講師,3.滿意講師的課程專業度' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        'sSql &= " " & vbCrLf
        sSql &= " UNION SELECT 'B31' BNO,'(三)教材,1.對於訓練教材感到滿意' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B32' BNO,'(三)教材,2.訓練教材能夠輔助課程學習' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        'sSql &= " " & vbCrLf
        sSql &= " UNION SELECT 'B41' BNO,'(四)訓練環境,1.您對於訓練場地感到滿意' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B42' BNO,'(四)訓練環境,2.您對於訓練設備感到滿意' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B43' BNO,'(四)訓練環境,3.您認為實作設備的數量適當' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B44' BNO,'(四)訓練環境,4.您認為實作設備新穎' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        'sSql &= " " & vbCrLf
        sSql &= " UNION SELECT 'B51' BNO,'(五)訓練評量,1.能促進學習效果' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B61' BNO,'(六)立即學習效果,1.您認為在訓練課程中，課程內容能讓您專注' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B62' BNO,'(六)立即學習效果,2.您在完成訓練後，已充份學習訓練課程所教授知識或技能' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B63' BNO,'(六)立即學習效果,3.您在完成訓練後，有學習到新的知識或技能' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        'sSql &= " " & vbCrLf
        sSql &= " UNION SELECT 'B71' BNO,'(七)整體意見,1.您對於訓練單位的課程安排與授課情形感到滿意' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B72' BNO,'(七)整體意見,2.您對於訓練單位的行政服務感到滿意' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B73' BNO,'(七)整體意見,3.您對於產業人才投資方案感到滿意' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        sSql &= " UNION SELECT 'B74' BNO,'(七)整體意見,4.您認為完成本訓練課程對於目前或未來工作有幫助' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf

        sSql &= " UNION SELECT 'T31' BNO,'全部問題加總' BTITLE,'非常同意' T1,'同意' T2,'普通' T3,'不同意' T4,'非常不同意' T5" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn)
        Return dt
    End Function

    ''' <summary>取得 調查統計分析表 分析資料</summary>
    ''' <returns></returns>
    Function GET_STUD_QUESTIONFAC2_dt() As DataTable
        Dim rPMS As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}}
        Dim sSql As String = ""
        'sSql &= " SELECT rr.ORGKINDGW,ip.TPLANID,ip.DISTID,ip.YEARS" & vbCrLf
        sSql &= " SELECT rr.ORGKINDGW,ip.DISTID" & vbCrLf
        sSql &= " ,CASE f2.B11 WHEN 1 THEN 1 END B11_1,CASE f2.B11 WHEN 2 THEN 1 END B11_2,CASE f2.B11 WHEN 3 THEN 1 END B11_3,CASE f2.B11 WHEN 4 THEN 1 END B11_4,CASE f2.B11 WHEN 5 THEN 1 END B11_5" & vbCrLf
        sSql &= " ,CASE f2.B12 WHEN 1 THEN 1 END B12_1,CASE f2.B12 WHEN 2 THEN 1 END B12_2,CASE f2.B12 WHEN 3 THEN 1 END B12_3,CASE f2.B12 WHEN 4 THEN 1 END B12_4,CASE f2.B12 WHEN 5 THEN 1 END B12_5" & vbCrLf
        sSql &= " ,CASE f2.B13 WHEN 1 THEN 1 END B13_1,CASE f2.B13 WHEN 2 THEN 1 END B13_2,CASE f2.B13 WHEN 3 THEN 1 END B13_3,CASE f2.B13 WHEN 4 THEN 1 END B13_4,CASE f2.B13 WHEN 5 THEN 1 END B13_5" & vbCrLf
        sSql &= " ,CASE f2.B14 WHEN 1 THEN 1 END B14_1,CASE f2.B14 WHEN 2 THEN 1 END B14_2,CASE f2.B14 WHEN 3 THEN 1 END B14_3,CASE f2.B14 WHEN 4 THEN 1 END B14_4,CASE f2.B14 WHEN 5 THEN 1 END B14_5" & vbCrLf
        sSql &= " ,CASE f2.B15 WHEN 1 THEN 1 END B15_1,CASE f2.B15 WHEN 2 THEN 1 END B15_2,CASE f2.B15 WHEN 3 THEN 1 END B15_3,CASE f2.B15 WHEN 4 THEN 1 END B15_4,CASE f2.B15 WHEN 5 THEN 1 END B15_5" & vbCrLf
        sSql &= " ,CASE f2.B21 WHEN 1 THEN 1 END B21_1,CASE f2.B21 WHEN 2 THEN 1 END B21_2,CASE f2.B21 WHEN 3 THEN 1 END B21_3,CASE f2.B21 WHEN 4 THEN 1 END B21_4,CASE f2.B21 WHEN 5 THEN 1 END B21_5" & vbCrLf
        sSql &= " ,CASE f2.B22 WHEN 1 THEN 1 END B22_1,CASE f2.B22 WHEN 2 THEN 1 END B22_2,CASE f2.B22 WHEN 3 THEN 1 END B22_3,CASE f2.B22 WHEN 4 THEN 1 END B22_4,CASE f2.B22 WHEN 5 THEN 1 END B22_5" & vbCrLf
        sSql &= " ,CASE f2.B23 WHEN 1 THEN 1 END B23_1,CASE f2.B23 WHEN 2 THEN 1 END B23_2,CASE f2.B23 WHEN 3 THEN 1 END B23_3,CASE f2.B23 WHEN 4 THEN 1 END B23_4,CASE f2.B23 WHEN 5 THEN 1 END B23_5" & vbCrLf
        sSql &= " ,CASE f2.B31 WHEN 1 THEN 1 END B31_1,CASE f2.B31 WHEN 2 THEN 1 END B31_2,CASE f2.B31 WHEN 3 THEN 1 END B31_3,CASE f2.B31 WHEN 4 THEN 1 END B31_4,CASE f2.B31 WHEN 5 THEN 1 END B31_5" & vbCrLf
        sSql &= " ,CASE f2.B32 WHEN 1 THEN 1 END B32_1,CASE f2.B32 WHEN 2 THEN 1 END B32_2,CASE f2.B32 WHEN 3 THEN 1 END B32_3,CASE f2.B32 WHEN 4 THEN 1 END B32_4,CASE f2.B32 WHEN 5 THEN 1 END B32_5" & vbCrLf
        sSql &= " ,CASE f2.B41 WHEN 1 THEN 1 END B41_1,CASE f2.B41 WHEN 2 THEN 1 END B41_2,CASE f2.B41 WHEN 3 THEN 1 END B41_3,CASE f2.B41 WHEN 4 THEN 1 END B41_4,CASE f2.B41 WHEN 5 THEN 1 END B41_5" & vbCrLf
        sSql &= " ,CASE f2.B42 WHEN 1 THEN 1 END B42_1,CASE f2.B42 WHEN 2 THEN 1 END B42_2,CASE f2.B42 WHEN 3 THEN 1 END B42_3,CASE f2.B42 WHEN 4 THEN 1 END B42_4,CASE f2.B42 WHEN 5 THEN 1 END B42_5" & vbCrLf
        sSql &= " ,CASE f2.B43 WHEN 1 THEN 1 END B43_1,CASE f2.B43 WHEN 2 THEN 1 END B43_2,CASE f2.B43 WHEN 3 THEN 1 END B43_3,CASE f2.B43 WHEN 4 THEN 1 END B43_4,CASE f2.B43 WHEN 5 THEN 1 END B43_5" & vbCrLf
        sSql &= " ,CASE f2.B44 WHEN 1 THEN 1 END B44_1,CASE f2.B44 WHEN 2 THEN 1 END B44_2,CASE f2.B44 WHEN 3 THEN 1 END B44_3,CASE f2.B44 WHEN 4 THEN 1 END B44_4,CASE f2.B44 WHEN 5 THEN 1 END B44_5" & vbCrLf
        sSql &= " ,CASE f2.B51 WHEN 1 THEN 1 END B51_1,CASE f2.B51 WHEN 2 THEN 1 END B51_2,CASE f2.B51 WHEN 3 THEN 1 END B51_3,CASE f2.B51 WHEN 4 THEN 1 END B51_4,CASE f2.B51 WHEN 5 THEN 1 END B51_5" & vbCrLf
        sSql &= " ,CASE f2.B61 WHEN 1 THEN 1 END B61_1,CASE f2.B61 WHEN 2 THEN 1 END B61_2,CASE f2.B61 WHEN 3 THEN 1 END B61_3,CASE f2.B61 WHEN 4 THEN 1 END B61_4,CASE f2.B61 WHEN 5 THEN 1 END B61_5" & vbCrLf
        sSql &= " ,CASE f2.B62 WHEN 1 THEN 1 END B62_1,CASE f2.B62 WHEN 2 THEN 1 END B62_2,CASE f2.B62 WHEN 3 THEN 1 END B62_3,CASE f2.B62 WHEN 4 THEN 1 END B62_4,CASE f2.B62 WHEN 5 THEN 1 END B62_5" & vbCrLf
        sSql &= " ,CASE f2.B63 WHEN 1 THEN 1 END B63_1,CASE f2.B63 WHEN 2 THEN 1 END B63_2,CASE f2.B63 WHEN 3 THEN 1 END B63_3,CASE f2.B63 WHEN 4 THEN 1 END B63_4,CASE f2.B63 WHEN 5 THEN 1 END B63_5" & vbCrLf
        sSql &= " ,CASE f2.B71 WHEN 1 THEN 1 END B71_1,CASE f2.B71 WHEN 2 THEN 1 END B71_2,CASE f2.B71 WHEN 3 THEN 1 END B71_3,CASE f2.B71 WHEN 4 THEN 1 END B71_4,CASE f2.B71 WHEN 5 THEN 1 END B71_5" & vbCrLf
        sSql &= " ,CASE f2.B72 WHEN 1 THEN 1 END B72_1,CASE f2.B72 WHEN 2 THEN 1 END B72_2,CASE f2.B72 WHEN 3 THEN 1 END B72_3,CASE f2.B72 WHEN 4 THEN 1 END B72_4,CASE f2.B72 WHEN 5 THEN 1 END B72_5" & vbCrLf
        sSql &= " ,CASE f2.B73 WHEN 1 THEN 1 END B73_1,CASE f2.B73 WHEN 2 THEN 1 END B73_2,CASE f2.B73 WHEN 3 THEN 1 END B73_3,CASE f2.B73 WHEN 4 THEN 1 END B73_4,CASE f2.B73 WHEN 5 THEN 1 END B73_5" & vbCrLf
        sSql &= " ,CASE f2.B74 WHEN 1 THEN 1 END B74_1,CASE f2.B74 WHEN 2 THEN 1 END B74_2,CASE f2.B74 WHEN 3 THEN 1 END B74_3,CASE f2.B74 WHEN 4 THEN 1 END B74_4,CASE f2.B74 WHEN 5 THEN 1 END B74_5" & vbCrLf

        sSql &= " ,CASE f3.B11 WHEN 1 THEN 1 END B11_G1,CASE f3.B11 WHEN 2 THEN 1 END B11_G2,CASE f3.B11 WHEN 3 THEN 1 END B11_G3,CASE f3.B11 WHEN 4 THEN 1 END B11_G4,CASE f3.B11 WHEN 5 THEN 1 END B11_G5" & vbCrLf
        sSql &= " ,CASE f3.B12 WHEN 1 THEN 1 END B12_G1,CASE f3.B12 WHEN 2 THEN 1 END B12_G2,CASE f3.B12 WHEN 3 THEN 1 END B12_G3,CASE f3.B12 WHEN 4 THEN 1 END B12_G4,CASE f3.B12 WHEN 5 THEN 1 END B12_G5" & vbCrLf
        sSql &= " ,CASE f3.B13 WHEN 1 THEN 1 END B13_G1,CASE f3.B13 WHEN 2 THEN 1 END B13_G2,CASE f3.B13 WHEN 3 THEN 1 END B13_G3,CASE f3.B13 WHEN 4 THEN 1 END B13_G4,CASE f3.B13 WHEN 5 THEN 1 END B13_G5" & vbCrLf
        sSql &= " ,CASE f3.B14 WHEN 1 THEN 1 END B14_G1,CASE f3.B14 WHEN 2 THEN 1 END B14_G2,CASE f3.B14 WHEN 3 THEN 1 END B14_G3,CASE f3.B14 WHEN 4 THEN 1 END B14_G4,CASE f3.B14 WHEN 5 THEN 1 END B14_G5" & vbCrLf
        sSql &= " ,CASE f3.B15 WHEN 1 THEN 1 END B15_G1,CASE f3.B15 WHEN 2 THEN 1 END B15_G2,CASE f3.B15 WHEN 3 THEN 1 END B15_G3,CASE f3.B15 WHEN 4 THEN 1 END B15_G4,CASE f3.B15 WHEN 5 THEN 1 END B15_G5" & vbCrLf
        sSql &= " ,CASE f3.B21 WHEN 1 THEN 1 END B21_G1,CASE f3.B21 WHEN 2 THEN 1 END B21_G2,CASE f3.B21 WHEN 3 THEN 1 END B21_G3,CASE f3.B21 WHEN 4 THEN 1 END B21_G4,CASE f3.B21 WHEN 5 THEN 1 END B21_G5" & vbCrLf
        sSql &= " ,CASE f3.B22 WHEN 1 THEN 1 END B22_G1,CASE f3.B22 WHEN 2 THEN 1 END B22_G2,CASE f3.B22 WHEN 3 THEN 1 END B22_G3,CASE f3.B22 WHEN 4 THEN 1 END B22_G4,CASE f3.B22 WHEN 5 THEN 1 END B22_G5" & vbCrLf
        sSql &= " ,CASE f3.B23 WHEN 1 THEN 1 END B23_G1,CASE f3.B23 WHEN 2 THEN 1 END B23_G2,CASE f3.B23 WHEN 3 THEN 1 END B23_G3,CASE f3.B23 WHEN 4 THEN 1 END B23_G4,CASE f3.B23 WHEN 5 THEN 1 END B23_G5" & vbCrLf
        sSql &= " ,CASE f3.B31 WHEN 1 THEN 1 END B31_G1,CASE f3.B31 WHEN 2 THEN 1 END B31_G2,CASE f3.B31 WHEN 3 THEN 1 END B31_G3,CASE f3.B31 WHEN 4 THEN 1 END B31_G4,CASE f3.B31 WHEN 5 THEN 1 END B31_G5" & vbCrLf
        sSql &= " ,CASE f3.B32 WHEN 1 THEN 1 END B32_G1,CASE f3.B32 WHEN 2 THEN 1 END B32_G2,CASE f3.B32 WHEN 3 THEN 1 END B32_G3,CASE f3.B32 WHEN 4 THEN 1 END B32_G4,CASE f3.B32 WHEN 5 THEN 1 END B32_G5" & vbCrLf
        sSql &= " ,CASE f3.B41 WHEN 1 THEN 1 END B41_G1,CASE f3.B41 WHEN 2 THEN 1 END B41_G2,CASE f3.B41 WHEN 3 THEN 1 END B41_G3,CASE f3.B41 WHEN 4 THEN 1 END B41_G4,CASE f3.B41 WHEN 5 THEN 1 END B41_G5" & vbCrLf
        sSql &= " ,CASE f3.B42 WHEN 1 THEN 1 END B42_G1,CASE f3.B42 WHEN 2 THEN 1 END B42_G2,CASE f3.B42 WHEN 3 THEN 1 END B42_G3,CASE f3.B42 WHEN 4 THEN 1 END B42_G4,CASE f3.B42 WHEN 5 THEN 1 END B42_G5" & vbCrLf
        sSql &= " ,CASE f3.B43 WHEN 1 THEN 1 END B43_G1,CASE f3.B43 WHEN 2 THEN 1 END B43_G2,CASE f3.B43 WHEN 3 THEN 1 END B43_G3,CASE f3.B43 WHEN 4 THEN 1 END B43_G4,CASE f3.B43 WHEN 5 THEN 1 END B43_G5" & vbCrLf
        sSql &= " ,CASE f3.B44 WHEN 1 THEN 1 END B44_G1,CASE f3.B44 WHEN 2 THEN 1 END B44_G2,CASE f3.B44 WHEN 3 THEN 1 END B44_G3,CASE f3.B44 WHEN 4 THEN 1 END B44_G4,CASE f3.B44 WHEN 5 THEN 1 END B44_G5" & vbCrLf
        sSql &= " ,CASE f3.B51 WHEN 1 THEN 1 END B51_G1,CASE f3.B51 WHEN 2 THEN 1 END B51_G2,CASE f3.B51 WHEN 3 THEN 1 END B51_G3,CASE f3.B51 WHEN 4 THEN 1 END B51_G4,CASE f3.B51 WHEN 5 THEN 1 END B51_G5" & vbCrLf
        sSql &= " ,CASE f3.B61 WHEN 1 THEN 1 END B61_G1,CASE f3.B61 WHEN 2 THEN 1 END B61_G2,CASE f3.B61 WHEN 3 THEN 1 END B61_G3,CASE f3.B61 WHEN 4 THEN 1 END B61_G4,CASE f3.B61 WHEN 5 THEN 1 END B61_G5" & vbCrLf
        sSql &= " ,CASE f3.B62 WHEN 1 THEN 1 END B62_G1,CASE f3.B62 WHEN 2 THEN 1 END B62_G2,CASE f3.B62 WHEN 3 THEN 1 END B62_G3,CASE f3.B62 WHEN 4 THEN 1 END B62_G4,CASE f3.B62 WHEN 5 THEN 1 END B62_G5" & vbCrLf
        sSql &= " ,CASE f3.B63 WHEN 1 THEN 1 END B63_G1,CASE f3.B63 WHEN 2 THEN 1 END B63_G2,CASE f3.B63 WHEN 3 THEN 1 END B63_G3,CASE f3.B63 WHEN 4 THEN 1 END B63_G4,CASE f3.B63 WHEN 5 THEN 1 END B63_G5" & vbCrLf
        sSql &= " ,CASE f3.B71 WHEN 1 THEN 1 END B71_G1,CASE f3.B71 WHEN 2 THEN 1 END B71_G2,CASE f3.B71 WHEN 3 THEN 1 END B71_G3,CASE f3.B71 WHEN 4 THEN 1 END B71_G4,CASE f3.B71 WHEN 5 THEN 1 END B71_G5" & vbCrLf
        sSql &= " ,CASE f3.B72 WHEN 1 THEN 1 END B72_G1,CASE f3.B72 WHEN 2 THEN 1 END B72_G2,CASE f3.B72 WHEN 3 THEN 1 END B72_G3,CASE f3.B72 WHEN 4 THEN 1 END B72_G4,CASE f3.B72 WHEN 5 THEN 1 END B72_G5" & vbCrLf
        sSql &= " ,CASE f3.B73 WHEN 1 THEN 1 END B73_G1,CASE f3.B73 WHEN 2 THEN 1 END B73_G2,CASE f3.B73 WHEN 3 THEN 1 END B73_G3,CASE f3.B73 WHEN 4 THEN 1 END B73_G4,CASE f3.B73 WHEN 5 THEN 1 END B73_G5" & vbCrLf
        sSql &= " ,CASE f3.B74 WHEN 1 THEN 1 END B74_G1,CASE f3.B74 WHEN 2 THEN 1 END B74_G2,CASE f3.B74 WHEN 3 THEN 1 END B74_G3,CASE f3.B74 WHEN 4 THEN 1 END B74_G4,CASE f3.B74 WHEN 5 THEN 1 END B74_G5" & vbCrLf

        sSql &= " ,CASE f4.B11 WHEN 1 THEN 1 END B11_W1,CASE f4.B11 WHEN 2 THEN 1 END B11_W2,CASE f4.B11 WHEN 3 THEN 1 END B11_W3,CASE f4.B11 WHEN 4 THEN 1 END B11_W4,CASE f4.B11 WHEN 5 THEN 1 END B11_W5" & vbCrLf
        sSql &= " ,CASE f4.B12 WHEN 1 THEN 1 END B12_W1,CASE f4.B12 WHEN 2 THEN 1 END B12_W2,CASE f4.B12 WHEN 3 THEN 1 END B12_W3,CASE f4.B12 WHEN 4 THEN 1 END B12_W4,CASE f4.B12 WHEN 5 THEN 1 END B12_W5" & vbCrLf
        sSql &= " ,CASE f4.B13 WHEN 1 THEN 1 END B13_W1,CASE f4.B13 WHEN 2 THEN 1 END B13_W2,CASE f4.B13 WHEN 3 THEN 1 END B13_W3,CASE f4.B13 WHEN 4 THEN 1 END B13_W4,CASE f4.B13 WHEN 5 THEN 1 END B13_W5" & vbCrLf
        sSql &= " ,CASE f4.B14 WHEN 1 THEN 1 END B14_W1,CASE f4.B14 WHEN 2 THEN 1 END B14_W2,CASE f4.B14 WHEN 3 THEN 1 END B14_W3,CASE f4.B14 WHEN 4 THEN 1 END B14_W4,CASE f4.B14 WHEN 5 THEN 1 END B14_W5" & vbCrLf
        sSql &= " ,CASE f4.B15 WHEN 1 THEN 1 END B15_W1,CASE f4.B15 WHEN 2 THEN 1 END B15_W2,CASE f4.B15 WHEN 3 THEN 1 END B15_W3,CASE f4.B15 WHEN 4 THEN 1 END B15_W4,CASE f4.B15 WHEN 5 THEN 1 END B15_W5" & vbCrLf
        sSql &= " ,CASE f4.B21 WHEN 1 THEN 1 END B21_W1,CASE f4.B21 WHEN 2 THEN 1 END B21_W2,CASE f4.B21 WHEN 3 THEN 1 END B21_W3,CASE f4.B21 WHEN 4 THEN 1 END B21_W4,CASE f4.B21 WHEN 5 THEN 1 END B21_W5" & vbCrLf
        sSql &= " ,CASE f4.B22 WHEN 1 THEN 1 END B22_W1,CASE f4.B22 WHEN 2 THEN 1 END B22_W2,CASE f4.B22 WHEN 3 THEN 1 END B22_W3,CASE f4.B22 WHEN 4 THEN 1 END B22_W4,CASE f4.B22 WHEN 5 THEN 1 END B22_W5" & vbCrLf
        sSql &= " ,CASE f4.B23 WHEN 1 THEN 1 END B23_W1,CASE f4.B23 WHEN 2 THEN 1 END B23_W2,CASE f4.B23 WHEN 3 THEN 1 END B23_W3,CASE f4.B23 WHEN 4 THEN 1 END B23_W4,CASE f4.B23 WHEN 5 THEN 1 END B23_W5" & vbCrLf
        sSql &= " ,CASE f4.B31 WHEN 1 THEN 1 END B31_W1,CASE f4.B31 WHEN 2 THEN 1 END B31_W2,CASE f4.B31 WHEN 3 THEN 1 END B31_W3,CASE f4.B31 WHEN 4 THEN 1 END B31_W4,CASE f4.B31 WHEN 5 THEN 1 END B31_W5" & vbCrLf
        sSql &= " ,CASE f4.B32 WHEN 1 THEN 1 END B32_W1,CASE f4.B32 WHEN 2 THEN 1 END B32_W2,CASE f4.B32 WHEN 3 THEN 1 END B32_W3,CASE f4.B32 WHEN 4 THEN 1 END B32_W4,CASE f4.B32 WHEN 5 THEN 1 END B32_W5" & vbCrLf
        sSql &= " ,CASE f4.B41 WHEN 1 THEN 1 END B41_W1,CASE f4.B41 WHEN 2 THEN 1 END B41_W2,CASE f4.B41 WHEN 3 THEN 1 END B41_W3,CASE f4.B41 WHEN 4 THEN 1 END B41_W4,CASE f4.B41 WHEN 5 THEN 1 END B41_W5" & vbCrLf
        sSql &= " ,CASE f4.B42 WHEN 1 THEN 1 END B42_W1,CASE f4.B42 WHEN 2 THEN 1 END B42_W2,CASE f4.B42 WHEN 3 THEN 1 END B42_W3,CASE f4.B42 WHEN 4 THEN 1 END B42_W4,CASE f4.B42 WHEN 5 THEN 1 END B42_W5" & vbCrLf
        sSql &= " ,CASE f4.B43 WHEN 1 THEN 1 END B43_W1,CASE f4.B43 WHEN 2 THEN 1 END B43_W2,CASE f4.B43 WHEN 3 THEN 1 END B43_W3,CASE f4.B43 WHEN 4 THEN 1 END B43_W4,CASE f4.B43 WHEN 5 THEN 1 END B43_W5" & vbCrLf
        sSql &= " ,CASE f4.B44 WHEN 1 THEN 1 END B44_W1,CASE f4.B44 WHEN 2 THEN 1 END B44_W2,CASE f4.B44 WHEN 3 THEN 1 END B44_W3,CASE f4.B44 WHEN 4 THEN 1 END B44_W4,CASE f4.B44 WHEN 5 THEN 1 END B44_W5" & vbCrLf
        sSql &= " ,CASE f4.B51 WHEN 1 THEN 1 END B51_W1,CASE f4.B51 WHEN 2 THEN 1 END B51_W2,CASE f4.B51 WHEN 3 THEN 1 END B51_W3,CASE f4.B51 WHEN 4 THEN 1 END B51_W4,CASE f4.B51 WHEN 5 THEN 1 END B51_W5" & vbCrLf
        sSql &= " ,CASE f4.B61 WHEN 1 THEN 1 END B61_W1,CASE f4.B61 WHEN 2 THEN 1 END B61_W2,CASE f4.B61 WHEN 3 THEN 1 END B61_W3,CASE f4.B61 WHEN 4 THEN 1 END B61_W4,CASE f4.B61 WHEN 5 THEN 1 END B61_W5" & vbCrLf
        sSql &= " ,CASE f4.B62 WHEN 1 THEN 1 END B62_W1,CASE f4.B62 WHEN 2 THEN 1 END B62_W2,CASE f4.B62 WHEN 3 THEN 1 END B62_W3,CASE f4.B62 WHEN 4 THEN 1 END B62_W4,CASE f4.B62 WHEN 5 THEN 1 END B62_W5" & vbCrLf
        sSql &= " ,CASE f4.B63 WHEN 1 THEN 1 END B63_W1,CASE f4.B63 WHEN 2 THEN 1 END B63_W2,CASE f4.B63 WHEN 3 THEN 1 END B63_W3,CASE f4.B63 WHEN 4 THEN 1 END B63_W4,CASE f4.B63 WHEN 5 THEN 1 END B63_W5" & vbCrLf
        sSql &= " ,CASE f4.B71 WHEN 1 THEN 1 END B71_W1,CASE f4.B71 WHEN 2 THEN 1 END B71_W2,CASE f4.B71 WHEN 3 THEN 1 END B71_W3,CASE f4.B71 WHEN 4 THEN 1 END B71_W4,CASE f4.B71 WHEN 5 THEN 1 END B71_W5" & vbCrLf
        sSql &= " ,CASE f4.B72 WHEN 1 THEN 1 END B72_W1,CASE f4.B72 WHEN 2 THEN 1 END B72_W2,CASE f4.B72 WHEN 3 THEN 1 END B72_W3,CASE f4.B72 WHEN 4 THEN 1 END B72_W4,CASE f4.B72 WHEN 5 THEN 1 END B72_W5" & vbCrLf
        sSql &= " ,CASE f4.B73 WHEN 1 THEN 1 END B73_W1,CASE f4.B73 WHEN 2 THEN 1 END B73_W2,CASE f4.B73 WHEN 3 THEN 1 END B73_W3,CASE f4.B73 WHEN 4 THEN 1 END B73_W4,CASE f4.B73 WHEN 5 THEN 1 END B73_W5" & vbCrLf
        sSql &= " ,CASE f4.B74 WHEN 1 THEN 1 END B74_W1,CASE f4.B74 WHEN 2 THEN 1 END B74_W2,CASE f4.B74 WHEN 3 THEN 1 END B74_W3,CASE f4.B74 WHEN 4 THEN 1 END B74_W4,CASE f4.B74 WHEN 5 THEN 1 END B74_W5" & vbCrLf

        sSql &= " FROM STUD_QUESTIONFAC2 f2 WITH(NOLOCK)" & vbCrLf
        sSql &= " JOIN CLASS_STUDENTSOFCLASS cs WITH(NOLOCK) on cs.SOCID=f2.SOCID" & vbCrLf
        sSql &= " JOIN CLASS_CLASSINFO cc WITH(NOLOCK) on cc.OCID=cs.OCID" & vbCrLf
        sSql &= " JOIN ID_PLAN ip WITH(NOLOCK) on ip.PLANID=cc.PLANID" & vbCrLf
        sSql &= " JOIN VIEW_RIDNAME rr on rr.RID=cc.RID" & vbCrLf
        sSql &= " LEFT JOIN STUD_QUESTIONFAC2 f3 WITH(NOLOCK) ON f3.SOCID=cs.SOCID AND rr.ORGKINDGW='G'" & vbCrLf
        sSql &= " LEFT JOIN STUD_QUESTIONFAC2 f4 WITH(NOLOCK) ON f4.SOCID=cs.SOCID AND rr.ORGKINDGW='W'" & vbCrLf
        sSql &= " WHERE ip.TPLANID=@TPLANID"
        sSql &= " AND rr.ORGKINDGW IN ('G','W')"
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist <> "" Then
            rPMS.Add("YEARS", v_yearlist)
            sSql &= " AND ip.YEARS=@YEARS"
        End If

        '開訓區間／'結訓區間
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        FTDate1.Text = TIMS.Cdate3(FTDate1.Text)
        FTDate2.Text = TIMS.Cdate3(FTDate2.Text)

        If STDate1.Text <> "" Then sSql &= " AND cc.STDate >=@STDate1" ' & TIMS.to_date(STDate1.Text) & vbCrLf '" & STDate1.Text & "'" & vbCrLf        End If
        If STDate2.Text <> "" Then sSql &= " AND cc.STDate <=@STDate2"
        If FTDate1.Text <> "" Then sSql &= " AND cc.FTDate >=@FTDate1"
        If FTDate2.Text <> "" Then sSql &= " AND cc.FTDate <=@FTDate2"

        If STDate1.Text <> "" Then rPMS.Add("STDate1", STDate1.Text)
        If STDate2.Text <> "" Then rPMS.Add("STDate2", STDate2.Text)
        If FTDate1.Text <> "" Then rPMS.Add("FTDate1", FTDate1.Text)
        If FTDate2.Text <> "" Then rPMS.Add("FTDate2", FTDate2.Text)

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, rPMS)
        Return dt
    End Function

    ''' <summary>創建一個空的 全部問題加總 統計欄位</summary>
    ''' <param name="dtDISTs"></param>
    ''' <returns></returns>
    Function GET_CREATE_FQALL_dt(dtDISTs As DataTable) As DataTable
        Dim sql_c1 As String = "SELECT '000' DISTID,'G' ORGKINDGW,0 BITEM ,0 BCOUNT"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql_c1, objconn)
        If dt1.Rows.Count > 0 Then dt1.Rows.Clear()

        Dim ary_ORGKINDGW1 As String() = {"G", "W", "GW"}
        For Each drD1 As DataRow In dtDISTs.Rows
            For Each s_ORGKINDGW1 As String In ary_ORGKINDGW1
                For i_B As Integer = 1 To 5
                    Dim dr1 As DataRow = Nothing
                    dr1 = dt1.NewRow
                    dt1.Rows.Add(dr1)
                    dr1("DISTID") = drD1("DISTID")
                    dr1("ORGKINDGW") = s_ORGKINDGW1
                    dr1("BITEM") = i_B
                    dr1("BCOUNT") = 0
                Next
            Next
        Next
        Return dt1
    End Function

    ''' <summary>將資料加入 全部問題加總 統計欄位</summary>
    ''' <param name="dtC1"></param>
    ''' <param name="s_DISTID"></param>
    ''' <param name="s_ORGKINDGW1"></param>
    ''' <param name="iBITEM"></param>
    ''' <param name="iBCOUNT"></param>
    Sub CHANGE_CT1(ByRef dtC1 As DataTable, s_DISTID As String, s_ORGKINDGW1 As String, iBITEM As Integer, iBCOUNT As Integer)
        Dim find_c1 As String = String.Format("DISTID='{0}' AND ORGKINDGW='{1}' AND BITEM={2}", s_DISTID, s_ORGKINDGW1, iBITEM)
        If dtC1.Select(find_c1).Length > 0 Then
            Dim iCNT1 As Integer = dtC1.Select(find_c1)(0)("BCOUNT")
            dtC1.Select(find_c1)(0)("BCOUNT") = (iCNT1 + iBCOUNT)
        End If
    End Sub

    ''' <summary>取得 全部問題加總 統計欄位</summary>
    ''' <param name="dtC1"></param>
    ''' <param name="vDISTID"></param>
    ''' <param name="vORGKINDGW1"></param>
    ''' <param name="iBITEM"></param>
    ''' <returns></returns>
    Private Function GET_CT1(ByRef dtC1 As DataTable, vDISTID As String, vORGKINDGW1 As String, iBITEM As Integer) As Integer
        Dim iCNT1 As Integer = 0
        Dim find_c1 As String = String.Format("DISTID='{0}' AND ORGKINDGW='{1}' AND BITEM={2}", vDISTID, vORGKINDGW1, iBITEM)
        If dtC1.Select(find_c1).Length > 0 Then Return dtC1.Select(find_c1)(0)("BCOUNT")
        Return iCNT1
    End Function

    Function GET_QUESTION_T2(ByRef drT1 As DataRow, ByRef dtDISTs As DataTable, ByRef dtB1 As DataTable, ByRef dtC1 As DataTable) As StringBuilder
        Dim sbExp As New StringBuilder
        Dim sbExp_2 As New StringBuilder
        Dim sbExp_3 As New StringBuilder
        sbExp.Append("<tr>")
        sbExp.Append(String.Concat("<td>", drT1("BTITLE"), "</td>"))
        sbExp.Append(Get_TITLE3(s_gTITLE3))
        sbExp.Append("</tr>")

        Const cst_DISTID_001 As String = "001"
        Const cst_DISTID_003 As String = "003"
        Const cst_DISTID_004 As String = "004"
        Const cst_DISTID_005 As String = "005"
        Const cst_DISTID_006 As String = "006"
        'Dim s_DIST_001 As String = ""
        'Dim s_DIST_003 As String = ""
        'Dim s_DIST_004 As String = ""
        'Dim s_DIST_005 As String = ""
        'Dim s_DIST_006 As String = ""
        Dim i_SUM_DIST_001 As Double = 0
        Dim i_SUM_DIST_003 As Double = 0
        Dim i_SUM_DIST_004 As Double = 0
        Dim i_SUM_DIST_005 As Double = 0
        Dim i_SUM_DIST_006 As Double = 0

        Dim i_TOTAL_DIST_001 As Double = 0
        Dim i_TOTAL_DIST_003 As Double = 0
        Dim i_TOTAL_DIST_004 As Double = 0
        Dim i_TOTAL_DIST_005 As Double = 0
        Dim i_TOTAL_DIST_006 As Double = 0

        Dim i_SUM_DIST_001_G As Double = 0
        Dim i_SUM_DIST_003_G As Double = 0
        Dim i_SUM_DIST_004_G As Double = 0
        Dim i_SUM_DIST_005_G As Double = 0
        Dim i_SUM_DIST_006_G As Double = 0

        Dim i_SUM_DIST_001_W As Double = 0
        Dim i_SUM_DIST_003_W As Double = 0
        Dim i_SUM_DIST_004_W As Double = 0
        Dim i_SUM_DIST_005_W As Double = 0
        Dim i_SUM_DIST_006_W As Double = 0

        Dim s_fff_G As String = ""
        Dim s_fff_W As String = ""
        Dim s_fff_GW As String = ""
        For i_B As Integer = 1 To 5
            sbExp.Append("<tr>")
            Dim s_TT As String = String.Concat(drT1(String.Concat("T", i_B)), "(", 6 - i_B, ")")
            sbExp.Append(String.Concat("<td>", s_TT, "</td>"))
            For Each drD1 As DataRow In dtDISTs.Rows
                s_fff_G = String.Concat("DISTID='", drD1("DISTID"), "' AND ", drT1("BNO"), "_G", i_B, "=1")
                s_fff_W = String.Concat("DISTID='", drD1("DISTID"), "' AND ", drT1("BNO"), "_W", i_B, "=1")
                s_fff_GW = String.Concat("DISTID='", drD1("DISTID"), "' AND ", drT1("BNO"), "_", i_B, "=1")
                sbExp.Append(String.Concat("<td class='noDecFormat'>", dtB1.Select(s_fff_G).Length, "</td>"))
                sbExp.Append(String.Concat("<td class='noDecFormat'>", dtB1.Select(s_fff_W).Length, "</td>"))
                sbExp.Append(String.Concat("<td class='noDecFormat'>", dtB1.Select(s_fff_GW).Length, "</td>"))
                Dim vDISTID As String = Convert.ToString(drD1("DISTID"))
                CHANGE_CT1(dtC1, vDISTID, "G", i_B, dtB1.Select(s_fff_G).Length)
                CHANGE_CT1(dtC1, vDISTID, "W", i_B, dtB1.Select(s_fff_W).Length)
                CHANGE_CT1(dtC1, vDISTID, "GW", i_B, dtB1.Select(s_fff_GW).Length)
                Select Case vDISTID 'Convert.ToString(drD1("DISTID"))
                    Case cst_DISTID_001
                        i_SUM_DIST_001_G += dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_001_W += dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_001 += dtB1.Select(s_fff_GW).Length * (6 - i_B)
                        i_TOTAL_DIST_001 += dtB1.Select(s_fff_GW).Length
                    Case cst_DISTID_003
                        i_SUM_DIST_003_G += dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_003_W += dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_003 += dtB1.Select(s_fff_GW).Length * (6 - i_B)
                        i_TOTAL_DIST_003 += dtB1.Select(s_fff_GW).Length
                    Case cst_DISTID_004
                        i_SUM_DIST_004_G += dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_004_W += dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_004 += dtB1.Select(s_fff_GW).Length * (6 - i_B)
                        i_TOTAL_DIST_004 += dtB1.Select(s_fff_GW).Length
                    Case cst_DISTID_005
                        i_SUM_DIST_005_G += dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_005_W += dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_005 += dtB1.Select(s_fff_GW).Length * (6 - i_B)
                        i_TOTAL_DIST_005 += dtB1.Select(s_fff_GW).Length
                    Case cst_DISTID_006
                        i_SUM_DIST_006_G += dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_006_W += dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_006 += dtB1.Select(s_fff_GW).Length * (6 - i_B)
                        i_TOTAL_DIST_006 += dtB1.Select(s_fff_GW).Length
                End Select
            Next
            sbExp.Append("</tr>")
        Next

        Dim s_TMP1 As String = ""
        sbExp_3.Append("<tr>")
        sbExp_3.Append(String.Concat("<td>分數(以100分計)</td>"))
        sbExp_2.Append("<tr>")
        sbExp_2.Append(String.Concat("<td>分數</td>"))
        sbExp.Append("<tr>")
        sbExp.Append(String.Concat("<td>小計</td>"))
        For Each drD1 As DataRow In dtDISTs.Rows
            sbExp_3.Append(String.Concat("<td class='noDecFormat'> </td>"))
            sbExp_3.Append(String.Concat("<td class='noDecFormat'> </td>"))

            sbExp_2.Append(String.Concat("<td class='noDecFormat'> </td>"))
            sbExp_2.Append(String.Concat("<td class='noDecFormat'> </td>"))
            Select Case Convert.ToString(drD1("DISTID"))
                Case cst_DISTID_001
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_001_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_001_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_001_G + i_SUM_DIST_001_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_001 > 0, TIMS.ROUND(i_SUM_DIST_001 / i_TOTAL_DIST_001, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_001 > 0, TIMS.ROUND(i_SUM_DIST_001 / i_TOTAL_DIST_001 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_003
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_003_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_003_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_003_G + i_SUM_DIST_003_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_003 > 0, TIMS.ROUND(i_SUM_DIST_003 / i_TOTAL_DIST_003, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_003 > 0, TIMS.ROUND(i_SUM_DIST_003 / i_TOTAL_DIST_003 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_004
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_004_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_004_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_004_G + i_SUM_DIST_004_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_004 > 0, TIMS.ROUND(i_SUM_DIST_004 / i_TOTAL_DIST_004, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_004 > 0, TIMS.ROUND(i_SUM_DIST_004 / i_TOTAL_DIST_004 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_005
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_005_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_005_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_005_G + i_SUM_DIST_005_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_005 > 0, TIMS.ROUND(i_SUM_DIST_005 / i_TOTAL_DIST_005, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_005 > 0, TIMS.ROUND(i_SUM_DIST_005 / i_TOTAL_DIST_005 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_006
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_006_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_006_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_006_G + i_SUM_DIST_006_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_006 > 0, TIMS.ROUND(i_SUM_DIST_006 / i_TOTAL_DIST_006, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_006 > 0, TIMS.ROUND(i_SUM_DIST_006 / i_TOTAL_DIST_006 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
            End Select

        Next
        sbExp.Append("</tr>")
        sbExp_2.Append("</tr>")
        sbExp_3.Append("</tr>")

        sbExp.Append(sbExp_2)
        sbExp.Append(sbExp_3)

        Dim s_colspan15 As String = "colspan='15'"
        sbExp.Append(String.Concat("<tr><td></td>", "<td ", s_colspan15, "></td></tr>"))

        Return sbExp
    End Function

    Function GET_QUESTION_T31(ByRef drT1 As DataRow, ByRef dtDISTs As DataTable, ByRef dtB1 As DataTable, ByRef dtC1 As DataTable) As StringBuilder
        Dim sbExp As New StringBuilder
        Dim sbExp_2 As New StringBuilder
        Dim sbExp_3 As New StringBuilder
        sbExp.Append("<tr>")
        sbExp.Append(String.Concat("<td>", drT1("BTITLE"), "</td>"))
        sbExp.Append(Get_TITLE3(s_gTITLE3))
        sbExp.Append("</tr>")

        Const cst_DISTID_001 As String = "001"
        Const cst_DISTID_003 As String = "003"
        Const cst_DISTID_004 As String = "004"
        Const cst_DISTID_005 As String = "005"
        Const cst_DISTID_006 As String = "006"
        'Dim s_DIST_001 As String = ""
        'Dim s_DIST_003 As String = ""
        'Dim s_DIST_004 As String = ""
        'Dim s_DIST_005 As String = ""
        'Dim s_DIST_006 As String = ""
        Dim i_SUM_DIST_001 As Double = 0
        Dim i_SUM_DIST_003 As Double = 0
        Dim i_SUM_DIST_004 As Double = 0
        Dim i_SUM_DIST_005 As Double = 0
        Dim i_SUM_DIST_006 As Double = 0

        Dim i_TOTAL_DIST_001 As Double = 0
        Dim i_TOTAL_DIST_003 As Double = 0
        Dim i_TOTAL_DIST_004 As Double = 0
        Dim i_TOTAL_DIST_005 As Double = 0
        Dim i_TOTAL_DIST_006 As Double = 0

        Dim i_SUM_DIST_001_G As Double = 0
        Dim i_SUM_DIST_003_G As Double = 0
        Dim i_SUM_DIST_004_G As Double = 0
        Dim i_SUM_DIST_005_G As Double = 0
        Dim i_SUM_DIST_006_G As Double = 0

        Dim i_SUM_DIST_001_W As Double = 0
        Dim i_SUM_DIST_003_W As Double = 0
        Dim i_SUM_DIST_004_W As Double = 0
        Dim i_SUM_DIST_005_W As Double = 0
        Dim i_SUM_DIST_006_W As Double = 0

        'Dim s_fff_G As String = ""
        'Dim s_fff_W As String = ""
        'Dim s_fff_GW As String = ""
        For i_B As Integer = 1 To 5
            sbExp.Append("<tr>")
            Dim s_TT As String = String.Concat(drT1(String.Concat("T", i_B)), "(", 6 - i_B, ")")
            sbExp.Append(String.Concat("<td>", s_TT, "</td>"))
            For Each drD1 As DataRow In dtDISTs.Rows
                Dim vDISTID As String = Convert.ToString(drD1("DISTID"))
                Dim i_CNT_G As Integer = GET_CT1(dtC1, vDISTID, "G", i_B)
                Dim i_CNT_W As Integer = GET_CT1(dtC1, vDISTID, "W", i_B)
                Dim i_CNT_GW As Integer = GET_CT1(dtC1, vDISTID, "GW", i_B)
                sbExp.Append(String.Concat("<td class='noDecFormat'>", i_CNT_G, "</td>"))
                sbExp.Append(String.Concat("<td class='noDecFormat'>", i_CNT_W, "</td>"))
                sbExp.Append(String.Concat("<td class='noDecFormat'>", i_CNT_GW, "</td>"))
                Select Case vDISTID 'Convert.ToString(drD1("DISTID"))
                    Case cst_DISTID_001
                        i_SUM_DIST_001_G += i_CNT_G
                        i_SUM_DIST_001_W += i_CNT_W
                        i_SUM_DIST_001 += i_CNT_GW * (6 - i_B)
                        i_TOTAL_DIST_001 += i_CNT_GW
                    Case cst_DISTID_003
                        i_SUM_DIST_003_G += i_CNT_G
                        i_SUM_DIST_003_W += i_CNT_W 'dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_003 += i_CNT_GW * (6 - i_B)
                        i_TOTAL_DIST_003 += i_CNT_GW
                    Case cst_DISTID_004
                        i_SUM_DIST_004_G += i_CNT_G 'dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_004_W += i_CNT_W 'dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_004 += i_CNT_GW * (6 - i_B)
                        i_TOTAL_DIST_004 += i_CNT_GW
                    Case cst_DISTID_005
                        i_SUM_DIST_005_G += i_CNT_G 'dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_005_W += i_CNT_W 'dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_005 += i_CNT_GW * (6 - i_B)
                        i_TOTAL_DIST_005 += i_CNT_GW
                    Case cst_DISTID_006
                        i_SUM_DIST_006_G += i_CNT_G 'dtB1.Select(s_fff_G).Length
                        i_SUM_DIST_006_W += i_CNT_W 'dtB1.Select(s_fff_W).Length
                        i_SUM_DIST_006 += i_CNT_GW * (6 - i_B)
                        i_TOTAL_DIST_006 += i_CNT_GW
                End Select
            Next
            sbExp.Append("</tr>")
        Next

        Dim s_TMP1 As String = ""
        sbExp_3.Append("<tr>")
        sbExp_3.Append(String.Concat("<td>分數(以100分計)</td>"))
        sbExp_2.Append("<tr>")
        sbExp_2.Append(String.Concat("<td>分數</td>"))
        sbExp.Append("<tr>")
        sbExp.Append(String.Concat("<td>小計</td>"))
        For Each drD1 As DataRow In dtDISTs.Rows
            sbExp_3.Append(String.Concat("<td class='noDecFormat'> </td>"))
            sbExp_3.Append(String.Concat("<td class='noDecFormat'> </td>"))

            sbExp_2.Append(String.Concat("<td class='noDecFormat'> </td>"))
            sbExp_2.Append(String.Concat("<td class='noDecFormat'> </td>"))
            Select Case Convert.ToString(drD1("DISTID"))
                Case cst_DISTID_001
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_001_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_001_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_001_G + i_SUM_DIST_001_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_001 > 0, TIMS.ROUND(i_SUM_DIST_001 / i_TOTAL_DIST_001, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_001 > 0, TIMS.ROUND(i_SUM_DIST_001 / i_TOTAL_DIST_001 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_003
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_003_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_003_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_003_G + i_SUM_DIST_003_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_003 > 0, TIMS.ROUND(i_SUM_DIST_003 / i_TOTAL_DIST_003, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_003 > 0, TIMS.ROUND(i_SUM_DIST_003 / i_TOTAL_DIST_003 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_004
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_004_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_004_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_004_G + i_SUM_DIST_004_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_004 > 0, TIMS.ROUND(i_SUM_DIST_004 / i_TOTAL_DIST_004, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_004 > 0, TIMS.ROUND(i_SUM_DIST_004 / i_TOTAL_DIST_004 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_005
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_005_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_005_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_005_G + i_SUM_DIST_005_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_005 > 0, TIMS.ROUND(i_SUM_DIST_005 / i_TOTAL_DIST_005, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_005 > 0, TIMS.ROUND(i_SUM_DIST_005 / i_TOTAL_DIST_005 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                Case cst_DISTID_006
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_006_G, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", i_SUM_DIST_006_W, "</td>"))
                    sbExp.Append(String.Concat("<td class='noDecFormat'>", (i_SUM_DIST_006_G + i_SUM_DIST_006_W), "</td>"))

                    s_TMP1 = If(i_TOTAL_DIST_006 > 0, TIMS.ROUND(i_SUM_DIST_006 / i_TOTAL_DIST_006, 2), "0")
                    sbExp_2.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
                    s_TMP1 = If(i_TOTAL_DIST_006 > 0, TIMS.ROUND(i_SUM_DIST_006 / i_TOTAL_DIST_006 * 20, 2), "0")
                    sbExp_3.Append(String.Concat("<td class='DecFormat2'>", s_TMP1, "</td>"))
            End Select

        Next
        sbExp.Append("</tr>")
        sbExp_2.Append("</tr>")
        sbExp_3.Append("</tr>")

        sbExp.Append(sbExp_2)
        sbExp.Append(sbExp_3)

        Dim s_colspan15 As String = "colspan='15'"
        sbExp.Append(String.Concat("<tr><td></td>", "<td ", s_colspan15, "></td></tr>"))

        Return sbExp
    End Function

    Private Sub sUtl_Export51()
        Dim dtDIST As DataTable = TIMS.Get_DISTNAME3dt(objconn)
        Dim dtT1 As DataTable = GET_QUESTION_T1()
        Dim dtB1 As DataTable = GET_STUD_QUESTIONFAC2_dt()
        Dim dtC1 As DataTable = GET_CREATE_FQALL_dt(dtDIST)

        Dim sFileName1 As String = String.Concat("調查統計分析表", "_", TIMS.GetDateNo2(1))

        'Dim s_colspan16 As String = " colspan='16' "
        Dim s_colspan15 As String = "colspan='15'"
        Dim s_colspan3 As String = "colspan='3'"
        Dim s_TPlanName As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim sROC_YEAR As String = If(v_yearlist <> "", (CInt(v_yearlist) - 1911).ToString(), "")
        Dim s_TITLE1 As String = String.Concat(s_TPlanName, " ", If(sROC_YEAR <> "", String.Concat(sROC_YEAR, "年 "), ""), "參訓學員見調查統計分析表")
        Dim s_TITLE2 As String = String.Concat("資料抓取:(", Now.Year - 1911, Now.ToString(".MM.dd"), "),訓後意見調查表")

        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= ("td{mso-number-format:""\@"";text-align@center;}") '將所有td欄位格式改 為"文字"置中
        strSTYLE &= (".text {mso-number-format:\@;text-align@center;} ") '.text欄位格式改 為"文字"置中
        strSTYLE &= (".noDecFormat {mso-number-format:""0"";}")
        strSTYLE &= (".DecFormat2{mso-number-format:""0.00"";}")
        strSTYLE &= "</style>"

        Dim ExpStr As String = ""
        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table border='1' cellspacing='0' style='font-family:標楷體;border-collapse@collapse;border@solid thin #000000;'>")

        ExpStr = String.Concat("<tr><td></td>", "<td ", s_colspan15, ">", s_TITLE1, "</td></tr>")
        sbHTML.Append(ExpStr)

        ExpStr = "<tr>"
        ExpStr &= "<td></td>"
        For Each drD1 As DataRow In dtDIST.Rows
            ExpStr &= String.Concat("<td ", s_colspan3, ">", drD1("DISTNAME3"), "</td>")
        Next
        ExpStr &= "</tr>"
        sbHTML.Append(ExpStr)

        ExpStr = ""
        For Each drT1 As DataRow In dtT1.Rows
            Select Case drT1("BNO")
                Case "T31"
                    sbHTML.Append(GET_QUESTION_T31(drT1, dtDIST, dtB1, dtC1))
                Case Else
                    sbHTML.Append(GET_QUESTION_T2(drT1, dtDIST, dtB1, dtC1))
            End Select
        Next

        ExpStr = String.Concat("<tr><td></td>", "<td ", s_colspan15, ">", s_TITLE2, "</td></tr>")
        sbHTML.Append(ExpStr)

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

End Class
