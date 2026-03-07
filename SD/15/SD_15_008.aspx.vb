Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Partial Class SD_15_008
    Inherits AuthBasePage

    'ReportQuery
    'SD_15_RPT FuncID.SelectedValue("SD_15_008_Rnn_2009")
    'SD_15_008_R11_2012~SD_15_008_R24_2012 (SD_15_008_R_Prt.aspx)
    'SD_15_008_R_Prt.aspx

    ' 共用設定 'Dim fontName As String="標楷體"
    Dim fontSize12s As Single = 12.0F
    Dim print_lock As New Object '(); //lock

    'FuncID
    Public Shared Sub Get_FuncList2012(ByRef obj As ListControl)
        Dim str1 As String = ""
        str1 &= "不設定"
        str1 &= ",性別,年齡,教育程度"
        str1 &= ",身分別,工作年資,地理分佈"
        str1 &= ",公司行業別,公司規模,參訓動機"
        str1 &= ",訓後動向,參訓單位類別,參加課程職能別"
        str1 &= ",參加課程型態別,訓練業別"

        Dim str2 As String = ""
        str2 &= "SD_15_008_R25_2012"
        str2 &= ",SD_15_008_R11_2012,SD_15_008_R12_2012,SD_15_008_R13_2012"
        str2 &= ",SD_15_008_R14_2012,SD_15_008_R15_2012,SD_15_008_R16_2012"
        str2 &= ",SD_15_008_R17_2012,SD_15_008_R18_2012,SD_15_008_R19_2012"
        str2 &= ",SD_15_008_R20_2012,SD_15_008_R21_2012,SD_15_008_R22_2012"
        str2 &= ",SD_15_008_R23_2012,SD_15_008_R24_2012"
        Dim str1A() As String = str1.Split(",")
        Dim str2A() As String = str2.Split(",")
        With obj
            For idx1 As Integer = 0 To str1A.Length - 1
                .Items.Insert(idx1, New ListItem(str1A(idx1), str2A(idx1)))
            Next
        End With

    End Sub

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁 '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        'This JavaScript will blur (remove focus from) the textbox immediately if it gets focus.
        'TMID1.Attributes.Add("onfocus", "this.blur();")
        'OCID1.Attributes.Add("onfocus", "this.blur();")
        TIMS.INPUT_ReadOnly2(TMID1)
        TIMS.INPUT_ReadOnly2(OCID1)

        '訓練機構
        If Not IsPostBack Then
            CCREATE1()
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
        'Years.Value = sm.UserInfo.Years

        Hid_LID.Value = sm.UserInfo.LID

        '跨年度查詢功能:False '(非委訓單位)
        Dim flag_can_over_years As Boolean = If(sm.UserInfo.LID <> 2, True, False)

        Dim Gv_yearlist As String = TIMS.GetListValue(yearlist) '取年度選擇值
        Hid_YEARS.Value = If(Gv_yearlist <> "", Gv_yearlist, sm.UserInfo.Years.ToString()) '必要年度傳入client

        Dim flag_selected_year As Boolean = (Gv_yearlist <> "" AndAlso Gv_yearlist <> sm.UserInfo.Years)
        Dim s_selected_year_Js As String = If(flag_selected_year, String.Concat("?selected_year=", Gv_yearlist), "")
        'Dim flag_can_over_years As Boolean = False '跨年度查詢功能:False

        If flag_can_over_years Then '(非委訓單位)
            'flag_can_over_years = True '跨年度查詢功能:True
            tr_yearlist.Visible = True
            BtnPrint3.Visible = True '分署(中心)、署(局) 才可列印

            Dim s_javascript_openOrg_FMT1 As String = String.Concat("javascript:openOrg('../../Common/LevOrg{0}.aspx", If(flag_selected_year, String.Concat("?selected_year=", Gv_yearlist), ""), "');")
            Button2.Attributes("onclick") = String.Format(s_javascript_openOrg_FMT1, If(flag_can_over_years OrElse sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        Else
            tr_yearlist.Visible = False '委訓單位隱藏 
            BtnPrint3.Visible = False

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
            Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End If

        Button1.Attributes("onclick") = "javascript:return print();"
    End Sub

    Sub CCREATE1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
        Common.SetListItem(SearchPlan, "A")

        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)

        '計畫範圍 產投
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

        'If sm.UserInfo.Years < 2008 Then,Get_FuncList(FuncID),Else,If sm.UserInfo.Years = 2008 Then,Get_FuncList1(FuncID),Else,Get_FuncList2009(FuncID),End If,End If,
        'Call Get_FuncList2009(FuncID)
        Call Get_FuncList2012(FuncID)

        If sm.UserInfo.LID <> "2" Then
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        Else
            UTL_OnlyOne_OCID()
        End If
    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RIDValue.Value.ToString = "A" Then RIDValue.Value = ""

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

        Dim sPrintAspx As String = "SD_15_008_R_Prt.aspx?" '2016 OLD
        If v_yearlist >= "2017" Then sPrintAspx = "SD_15_008_R_Prt2.aspx?" '2017
        'If yearlist.SelectedValue >= "2017" Then,Common.MessageBox(Me, TIMS.cst_Error2),Exit Sub,End If,

        Dim MyValue As String = ""
        MyValue = "filename=" & TIMS.ClearSQM(FuncID.SelectedValue)
        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue &= "&Years=" & v_yearlist 'yearlist.SelectedValue
        MyValue &= "&OCID=" & OCIDValue1.Value
        MyValue &= "&RID=" & RIDValue.Value
        MyValue &= "&SearchPlan=" & SearchPlan1 '"",G,W
        MyValue &= "&PackageType=" & sPackType '"",2,3
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_15_RPT", FuncID.SelectedValue, MyValue)
        Dim strScript As String = String.Concat("<script language=""javascript"">", "window.open('", sPrintAspx, MyValue & "');", "</script>")
        Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    '列印明細(匯出Excel)
    Private Sub BtnPrint3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnPrint3.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '匯出EXCEL
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim sYearType As String = "2016" '2016 OLD
        If v_yearlist >= "2017" Then sYearType = "2017" '2017
        Select Case sYearType
            Case "2017"
                Call SUtl_Export4()
            Case "2016"
                Call SUtl_Export3()
            Case Else
                Call SUtl_Export3()
        End Select
    End Sub

    Sub UTL_OnlyOne_OCID()
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        UTL_OnlyOne_OCID()
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
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" AndAlso OCIDValue1.Value = "" Then
            Errmsg += "請選擇 職類/班別" & vbCrLf
        ElseIf RIDValue.Value = "" Then
            Errmsg += "請選擇 訓練機構" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    'sch 2016 old
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
        sql &= " select oo.orgname" & vbCrLf
        'sql &= " ,cc.classcname + ' 第' + cc.cycltype + '期' classcname" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        'Sql += " ,ss.idno" & vbCrLf
        sql &= " ,ss.name cname" & vbCrLf
        sql &= " ,cs.SOCID" & vbCrLf

        sql &= " ,ss2.Q1" & vbCrLf
        sql &= " ,ss2.Q2" & vbCrLf
        sql &= " ,ss2.Q3" & vbCrLf
        sql &= " ,ss2.Q4" & vbCrLf
        sql &= " ,ss2.Q5" & vbCrLf
        sql &= " ,ss2.Q6" & vbCrLf
        sql &= " ,ss2.Q6_7" & vbCrLf
        sql &= " ,ss2.Q6_8" & vbCrLf
        sql &= " ,ss2.Q7" & vbCrLf
        sql &= " ,ss2.Q8_1_Note" & vbCrLf
        sql &= " ,ss2.Q8_2_Note" & vbCrLf
        sql &= " ,ss2.Q8_3_Note" & vbCrLf
        sql &= " ,ss2.Q9_1_Note" & vbCrLf
        sql &= " ,ss2.Q9_2_Note" & vbCrLf
        sql &= " ,ss2.Q9_3_Note" & vbCrLf
        sql &= " ,ss2.BusName" & vbCrLf
        sql &= " ,ss2.Q10" & vbCrLf
        sql &= " ,ss2.Q11" & vbCrLf
        sql &= " ,ss2.Q12" & vbCrLf
        sql &= " ,ss2.Q13" & vbCrLf
        sql &= " ,ss2.Q14" & vbCrLf
        sql &= " ,ss2.Q15" & vbCrLf
        sql &= " ,ss2.Q8" & vbCrLf
        sql &= " ,ss2.Q10_1_Note" & vbCrLf
        sql &= " ,ss2.Q10_2_Note" & vbCrLf
        sql &= " ,ss2.Q10_3_Note" & vbCrLf
        sql &= " ,ss2.Q16" & vbCrLf

        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.Planid =cc.Planid  AND pp.comidno= cc.comidno AND pp.seqno = cc.seqno" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.planid =cc.planid" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.ocid =cc.ocid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        sql &= " JOIN STUD_QUESTIONFIN ss2 on ss2.socid=cs.socid" & vbCrLf
        sql &= " LEFT JOIN STUD_SUBSIDYCOST sc ON sc.SOCID =cs.SOCID" & vbCrLf
        'STUD_QUESTIONFIN '201608 加入排除條件 AMU 'A.排除離退訓(離退訓作業功能) 'B.排除有結訓未申請(補助申請功能) 'C.排除審核不通過(補助審核功能)的學員
        sql &= " WHERE cs.STUDSTATUS NOT IN (2,3)" & vbCrLf '非離退
        sql &= " AND sc.SOCID IS NOT NULL" & vbCrLf '有申請資料
        sql &= " AND ISNULL(sc.AppliedStatusM,'Y')='Y'" & vbCrLf '審核通過 或申請中的
        sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " AND ip.Years='" & v_yearlist & "'" & vbCrLf
        sql &= " AND cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        If SearchPlan1 <> "" Then
            sql &= " AND oo.OrgKind2='" & SearchPlan1 & "'" & vbCrLf
        End If
        If sPackType <> "" Then
            sql &= " AND pp.PackageType='" & sPackType & "'" & vbCrLf
        End If

        sql &= " order by cs.socid" & vbCrLf

        Rst = sql
        Return Rst
    End Function

    '匯出EXCEL (2016 old)
    Sub SUtl_Export3()
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(Search_Query3, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        dt.DefaultView.Sort = "SOCID"
        dt = TIMS.dv2dt(dt.DefaultView)

        Dim strFileName As String = "參訓學員訓後動態調查表" & OCIDValue1.Value '"SD_15_008_list"

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strFileName, System.Text.Encoding.UTF8) & ".xls")
        Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim num As Integer = 0
        Dim ExportStr As String = ""           '建立輸出文字

        ExportStr = "訓練單位名稱"
        ExportStr &= vbTab & "課程名稱"
        ExportStr &= vbTab & "課程代碼"
        ExportStr &= vbTab & "開訓日期"
        ExportStr &= vbTab & "結訓日期"
        ExportStr &= vbTab & "學員姓名"

        ExportStr &= vbTab & "1.學員目前的近況為何?"
        ExportStr &= vbTab & "2.學員於結訓後薪資有提升嗎?"
        ExportStr &= vbTab & "3.學員的職位有變化嗎?"
        ExportStr &= vbTab & "4.學員對目前工作的滿意度是否有變化?"
        ExportStr &= vbTab & "5.學員目前的工作內容是否與參訓課程內容相關?"
        'ExportStr &= vbTab & "6." '"6.學員是否同意參加訓練對目前工作表現有幫助?"
        ExportStr &= vbTab & "6-1. 學員是否同意參加訓練對目前工作表現有幫助？" '"7.學員是否同意參加訓練對未來工作表現有幫助?"
        ExportStr &= vbTab & "6-2. 學員是否同意參加訓練對第二專長培育有幫助？" '"8.學員是否同意參加訓練對第二專長培育有幫助?"

        'ExportStr &= vbTab & "9.承上題，參加本項訓練對學員的幫助是在哪方面?"
        ExportStr &= vbTab & "7. 學員是否有繼續參與進修訓練的意願？" '"10.學員是否有繼續參與進修訓練的意願?"
        ExportStr &= vbTab & "8. 學員認為還需要加強哪方面的專業知識使工作進行得更順利？" '"11.學員認為還需要加強哪方面的專業知識使工作進行得更順利?"
        ExportStr &= vbTab & "8-1." '"11-1."
        ExportStr &= vbTab & "8-2." '"11-2."
        ExportStr &= vbTab & "8-3." '"11-3."

        ExportStr &= vbTab & "9. 學員常和本課程的哪些學員、教師或職員聯絡？" '"12.學員常和本課程的哪些學員、教師或職員聯絡?"
        ExportStr &= vbTab & "9-1." '"12-1."
        ExportStr &= vbTab & "9-2." '"12-2."
        ExportStr &= vbTab & "9-3." '"12-3."

        ExportStr &= vbTab & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面 Stud_QuestionFin
        For Each dr As DataRow In dt.DefaultView.Table.Rows
            num += 1
            ExportStr = ""
            ExportStr &= dr("orgname").ToString
            ExportStr &= vbTab & dr("classcname").ToString '"課程名稱"
            ExportStr &= vbTab & dr("ocid").ToString '"課程代碼"
            ExportStr &= vbTab & dr("stdate").ToString '"開訓日期"
            ExportStr &= vbTab & dr("ftdate").ToString '"結訓日期"
            ExportStr &= vbTab & dr("cname").ToString '"學員姓名"

            ExportStr &= vbTab & dr("Q1").ToString  '"1.學員目前的近況為何?"
            ExportStr &= vbTab & dr("Q2").ToString  '"2.學員於結訓後薪資有提升嗎?"
            ExportStr &= vbTab & dr("Q3").ToString  '"3.學員的職位有變化嗎?"
            ExportStr &= vbTab & dr("Q4").ToString  '"4.學員對目前工作的滿意度是否有變化?"
            ExportStr &= vbTab & dr("Q5").ToString  '"5.學員目前的工作內容是否與參訓課程內容相關?"
            'ExportStr &= vbTab & "" ' dr("Q6").ToString  '"6.學員是否同意參加訓練對目前工作表現有幫助?"
            ExportStr &= vbTab & dr("Q6_7").ToString  '6-1"7.學員是否同意參加訓練對未來工作表現有幫助?"
            ExportStr &= vbTab & dr("Q6_8").ToString  '6-2"8.學員是否同意參加訓練對第二專長培育有幫助?"

            ExportStr &= vbTab & dr("Q7").ToString  '"7.承上題，參加本項訓練對學員的幫助是在哪方面?"
            ExportStr &= vbTab & dr("Q8").ToString  '"8.學員是否有繼續參與進修訓練的意願?"
            'ExportStr &= vbTab & "" 'dr("Q9").ToString  '"9.學員認為還需要加強哪方面的專業知識使工作進行得更順利?"
            ExportStr &= vbTab & dr("Q9_1_Note").ToString
            ExportStr &= vbTab & dr("Q9_2_Note").ToString
            ExportStr &= vbTab & dr("Q9_3_Note").ToString
            ExportStr &= vbTab & "" 'dr("Q10").ToString  '"10.學員常和本課程的哪些學員、教師或職員聯絡?"
            ExportStr &= vbTab & dr("Q10_1_Note").ToString
            ExportStr &= vbTab & dr("Q10_2_Note").ToString
            ExportStr &= vbTab & dr("Q10_3_Note").ToString

            ExportStr &= vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    'sch 2017
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

        Dim vPLANID As String = ""
        Dim vDISTID As String = ""
        Dim vRIDValue As String = ""
        If OCIDValue1.Value = "" Then
            Select Case sm.UserInfo.LID
                Case 0
                    vDISTID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    vRIDValue = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
                    If vRIDValue.Length = 1 Then vRIDValue = ""
                Case 1
                    vPLANID = sm.UserInfo.PlanID
                    vDISTID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    If vDISTID = "" Then vDISTID = sm.UserInfo.DistID
                    vRIDValue = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
                    If vRIDValue.Length = 1 Then vRIDValue = ""
                Case Else
                    vPLANID = sm.UserInfo.PlanID
                    vDISTID = sm.UserInfo.DistID
                    vRIDValue = sm.UserInfo.RID 'If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
            End Select
        End If

        Dim sql As String = ""
        sql &= " SELECT oo.ORGNAME" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        'Sql += " ,ss.idno" & vbCrLf
        sql &= " ,ss.NAME CNAME ,cs.SOCID" & vbCrLf
        sql &= " ,ss2.Q1,ss2.Q2,ss2.Q3,ss2.Q4,ss2.Q5,ss2.Q8" & vbCrLf
        'sql &= " ,ss2.Q1A ,ss2.Q1B,ss2.Q1C" & vbCrLf
        sql &= " ,ss2.Q7MR1,ss2.Q7MR2,ss2.Q7MR3,ss2.Q7MR4" & vbCrLf
        sql &= " ,ss2.Q211,ss2.Q212,ss2.Q213,ss2.Q214" & vbCrLf
        sql &= " ,ss2.Q215,ss2.Q216,ss2.Q217,ss2.Q218" & vbCrLf
        sql &= " ,ss2.Q221,ss2.Q222,ss2.Q223,ss2.Q224" & vbCrLf
        sql &= " ,ss2.Q225,ss2.Q226,ss2.Q3_NOTE" & vbCrLf
        'sql &= " ,DASOURCE,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.Planid =cc.Planid  AND pp.comidno= cc.comidno AND pp.seqno = cc.seqno" & vbCrLf
        sql &= " JOIN ID_Plan ip on ip.planid =cc.planid" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.ocid =cc.ocid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        sql &= " JOIN STUD_QUESTIONFIN ss2 on ss2.socid=cs.socid" & vbCrLf
        sql &= " LEFT JOIN STUD_SUBSIDYCOST sc ON sc.SOCID =cs.SOCID" & vbCrLf
        'STUD_QUESTIONFIN '201608 加入排除條件 AMU 'A.排除離退訓(離退訓作業功能) 'B.排除有結訓未申請(補助申請功能)'C.排除審核不通過(補助審核功能)的學員
        'If TIMS.sUtl_ChkTest() Then
        '    sql &= " AND ip.TPlanID='28'  and cc.ocid =51891  and ss2.modifydate >=DATEADD(DAY, getdate(), -3)"
        'End If
        sql &= " WHERE cs.STUDSTATUS NOT IN (2,3)" & vbCrLf '非離退
        sql &= " AND sc.SOCID IS NOT NULL" & vbCrLf '有申請資料
        sql &= " AND ISNULL(sc.AppliedStatusM,'Y')='Y'" & vbCrLf '審核通過 或申請中的
        sql &= " AND ip.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql &= " AND ip.YEARS='" & v_yearlist & "'" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
                If vDISTID <> "" Then sql &= " AND ip.DISTID='" & vDISTID & "'" & vbCrLf
                If vRIDValue <> "" Then sql &= " AND cc.RID='" & vRIDValue & "'" & vbCrLf
            Case Else
                If vPLANID <> "" Then sql &= " AND ip.PLANID='" & vPLANID & "'" & vbCrLf
                If vDISTID <> "" Then sql &= " AND ip.DISTID='" & vDISTID & "'" & vbCrLf
                If vRIDValue <> "" Then sql &= " AND cc.RID='" & vRIDValue & "'" & vbCrLf
        End Select

        Dim s_OCIDVX As String = OCIDValue1.Value
        If s_OCIDVX <> "" Then sql &= " AND cc.OCID='" & s_OCIDVX & "'" & vbCrLf

        If vPLANID = "" AndAlso vDISTID = "" AndAlso vRIDValue = "" AndAlso s_OCIDVX = "" Then
            sql &= " AND 1!=1" & vbCrLf '(查無資料)
        End If

        If SearchPlan1 <> "" Then
            sql &= " AND oo.OrgKind2='" & SearchPlan1 & "'" & vbCrLf
        End If
        If sPackType <> "" Then
            sql &= " AND pp.PackageType='" & sPackType & "'" & vbCrLf
        End If
        sql &= " ORDER BY cs.socid" & vbCrLf

        Rst = sql
        Return Rst
    End Function

    '匯出EXCEL (2017)
    Sub SUtl_Export4()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim dt As DataTable = DbAccess.GetDataTable(Search_Query4, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        dt.DefaultView.Sort = "SOCID"
        dt = TIMS.dv2dt(dt.DefaultView)

        '"SD_15_008_list"
        Dim str_OCIDVX As String = If(OCIDValue1.Value <> "", String.Concat(OCIDValue1.Value, "x"), "")
        Dim strFileName As String = String.Concat("參訓學員訓後動態調查表x", str_OCIDVX, TIMS.GetDateNo2(3))

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strFileName, System.Text.Encoding.UTF8) & ".xls")
        Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        'Dim num As Integer = 0
        'Dim ExportStr As String = ""           '建立輸出文字

        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        '套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'Common.RespWrite(Me, "<tr>")

        'Dim sPattern As String = "訓練單位名稱,課程名稱,課程代碼,開訓日期,結訓日期,學員姓名,1-1.請問您目前的就業狀況為何?,1-2.請問您的薪資於結訓後有提升嗎?,1-3.請問您擔任的職位有變化嗎?,1-4.請問您對目前工作的滿意度是否有變化?,1-5.請問您目前工作內容是否與本次參訓課程有相關?,1-6.請問您是否有繼續參與本方案的意願?,1-7-1.結訓後與教師保持聯絡,1-7-2.結訓後與學員持聯絡,1-7-3.結訓後與工作人員保持聯絡,1-7-4.結訓後無保持聯絡,2-1-1.參加訓練後，對工作能力更有信心,2-1-2.參加訓練後，有助於提升工作技能,2-1-3.參加訓練後，有助於提升工作效率,2-1-4.參加訓練後，能增進我的問題解決能力,2-1-5.參加訓練後，能將所學應用到工作上,2-1-6.參加訓練後，能將所學應用於日常生活中,2-1-7.是否同意參加訓練對第二專長有幫助,2-1-8.是否同意參加訓練對目前工作表現有幫助,2-2-1.參加訓練後，有助於提升我的績效考核,2-2-2.參加訓練後，有助於薪資的調升,2-2-3.參加訓練後，有助於職位的升遷,2-2-4.參加訓練後，有助於獲得證照,2-2-5.參加訓練後，有助於發展職涯,2-2-6.參加訓練後，有助於強化個人職場競爭力"
        Dim sPattern As String = ""
        sPattern &= "訓練單位名稱,課程名稱,課程代碼,開訓日期,結訓日期,學員姓名,1-1.請問您目前的就業狀況為何? 1.留任原公司/2.轉換至同產業的公司/3.轉換至不同產業的公司/4.創業/5.已離職待業中/6.其他,1-2.請問您的薪資於結訓後有提升嗎?1.大幅提升/2.小幅提升/3.沒有變化/4.小幅減少/5.大幅減少"
        sPattern &= ",1-3.請問您擔任的職位有變化嗎?1.升職/2.調職/3.沒有變化/4.降職,1-4.請問您對目前工作的滿意度是否有變化?1.大幅提升/2.小幅提升/3.沒有變化/4.小幅減少/5.大幅減少,1-5.請問您目前工作內容是否與本次參訓課程有相關?1.非常相關/2.相關/3.尚可/4.不相關/5.非常不相關"
        sPattern &= ",1-6.請問您是否有繼續參與本方案的意願?1.非常想參與/2.想參與/3.尚可/4.不想參與/5.非常不想參與,1-7-1.結訓後與教師保持聯絡,1-7-2.結訓後與學員持聯絡,1-7-3.結訓後與工作人員保持聯絡,1-7-4.結訓後無保持聯絡"
        sPattern &= ",2-1-1.參加訓練後，對工作能力更有信心/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-1-2.參加訓練後，有助於提升工作技能/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-1-3.參加訓練後，有助於提升工作效率/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意"
        sPattern &= ",2-1-4.參加訓練後，能增進我的問題解決能力/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-1-5.參加訓練後，能將所學應用到工作上/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-1-6.參加訓練後，能將所學應用於日常生活中/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意"
        sPattern &= ",2-1-7.是否同意參加訓練對第二專長有幫助/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-1-8.是否同意參加訓練對目前工作表現有幫助/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-2-1.參加訓練後，有助於提升我的績效考核/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意"
        sPattern &= ",2-2-2.參加訓練後，有助於薪資的調升/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-2-3.參加訓練後，有助於職位的升遷/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-2-4.參加訓練後，有助於獲得證照/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意"
        sPattern &= ",2-2-5.參加訓練後，有助於發展職涯/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,2-2-6.參加訓練後，有助於強化個人職場競爭力/1.非常同意/2.同意/3.普通/4.不同意/5.非常不同意,其他建議"
        Dim sColumn As String = "ORGNAME,CLASSCNAME,OCID,STDATE,FTDATE,CNAME,Q1,Q2,Q3,Q4,Q5,Q8,Q7MR1,Q7MR2,Q7MR3,Q7MR4,Q211,Q212,Q213,Q214,Q215,Q216,Q217,Q218,Q221,Q222,Q223,Q224,Q225,Q226,Q3_NOTE"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = "<tr>"
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= "<td>" & sPatternA(i) & "</td>" '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

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
                ExportStr &= String.Concat("<td>", dr(sColumnA(i)), "</td>") '& vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Dim ERRMSG1 As String = ""
        Dim V_YEARS As String = TIMS.GetListValue(yearlist)
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        If V_YEARS = "" AndAlso STDate1.Text = "" AndAlso STDate2.Text = "" Then
            ERRMSG1 &= "(匯出助益率分析表)年度不選時，開訓期間為必填!" & vbCrLf
            'ElseIf STDate1.Text <> "" AndAlso STDate2.Text = "" Then
            '    ERRMSG1 &= "(匯出助益率分析表)開訓期間結束為必填!" & vbCrLf
            'ElseIf STDate1.Text = "" AndAlso STDate2.Text <> "" Then
            '    ERRMSG1 &= "(匯出助益率分析表)開訓期間開始為必填!" & vbCrLf
        End If
        If ERRMSG1 <> "" Then
            Common.MessageBox(Me, ERRMSG1)
            Return
        End If

        'Dim htPP As New Hashtable
        ExportXlsStd28()
    End Sub

    Function SEARCH_DATA1_dt() As DataTable
        Dim ndt As DateTime = Now
        'Dim smYear As String = sm.UserInfo.Years
        Dim V_YEARS As String = TIMS.GetListValue(yearlist)
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)

        Dim PMS1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}}
        Dim SSQL As String = ""
        SSQL &= " WITH WT1 AS (SELECT * FROM (VALUES ('1'),('2'),('3'),('4'),('5')) AS T1(SEQ))" & vbCrLf
        SSQL &= " ,WC1 AS ( SELECT a.Q217,a.Q218,cs.ORGKIND2" & vbCrLf
        SSQL &= " FROM dbo.V_STUDENTINFO cs" & vbCrLf
        SSQL &= " JOIN dbo.STUD_QUESTIONFIN a WITH(NOLOCK) ON a.SOCID =cs.SOCID" & vbCrLf
        'SSQL &= " WHERE CS.TPLANID=@TPLANID AND CS.YEARS IN ('2022','2023') AND CS.STDATE>=CONVERT(DATE,'2023/01/01') AND CS.STDATE<=CONVERT(DATE,'2023/09/30') )" & vbCrLf
        SSQL &= " WHERE CS.TPLANID=@TPLANID" & vbCrLf
        If V_YEARS <> "" Then
            SSQL &= " AND CS.YEARS=@YEARS" & vbCrLf
            PMS1.Add("YEARS", V_YEARS)
        End If
        If STDate1.Text <> "" Then
            SSQL &= " AND CS.STDATE>=CONVERT(DATE,@STDate1)" & vbCrLf
            PMS1.Add("STDate1", STDate1.Text)
        End If
        If STDate2.Text <> "" Then
            SSQL &= " AND CS.STDATE<=CONVERT(DATE,@STDate2)" & vbCrLf
            PMS1.Add("STDate2", STDate2.Text)
        End If
        SSQL &= " )" & vbCrLf
        SSQL &= " ,WQ7 AS (SELECT Q217 SEQ,COUNT(CASE ORGKIND2 WHEN 'G' THEN 1 END) Q217CTG,COUNT(CASE ORGKIND2 WHEN 'W' THEN 1 END) Q217CTW FROM WC1 GROUP BY Q217)" & vbCrLf
        SSQL &= " ,WQ8 AS (SELECT Q218 SEQ,COUNT(CASE ORGKIND2 WHEN 'G' THEN 1 END) Q218CTG,COUNT(CASE ORGKIND2 WHEN 'W' THEN 1 END) Q218CTW FROM WC1 GROUP BY Q218)" & vbCrLf
        SSQL &= " ,WQ7S AS (SELECT SUM(Q217CTG) SUM_Q217CTG,SUM(Q217CTW) SUM_Q217CTW" & vbCrLf
        SSQL &= " ,SUM(CASE WHEN SEQ='1' OR SEQ='2' THEN Q217CTG END) SUM2_Q217CTG,SUM(CASE WHEN SEQ='1' OR SEQ='2' THEN Q217CTW END) SUM2_Q217CTW FROM WQ7)" & vbCrLf
        SSQL &= " ,WQ8S AS (SELECT SUM(Q218CTG) SUM_Q218CTG,SUM(Q218CTW) SUM_Q218CTW" & vbCrLf
        SSQL &= " ,SUM(CASE WHEN SEQ='1' OR SEQ='2' THEN Q218CTG END) SUM2_Q218CTG,SUM(CASE WHEN SEQ='1' OR SEQ='2' THEN Q218CTW END) SUM2_Q218CTW FROM WQ8)" & vbCrLf
        SSQL &= " SELECT A.SEQ,B1.Q217CTG,B1.Q217CTW,B2.Q218CTG,B2.Q218CTW" & vbCrLf
        SSQL &= " ,C1.SUM_Q217CTG,C1.SUM_Q217CTW,C2.SUM_Q218CTG,C2.SUM_Q218CTW" & vbCrLf
        SSQL &= " ,C1.SUM2_Q217CTG,C1.SUM2_Q217CTW,C2.SUM2_Q218CTG,C2.SUM2_Q218CTW" & vbCrLf
        SSQL &= " FROM WT1 A" & vbCrLf
        SSQL &= " LEFT JOIN WQ7 B1 ON B1.SEQ=A.SEQ" & vbCrLf
        SSQL &= " LEFT JOIN WQ8 B2 ON B2.SEQ=A.SEQ" & vbCrLf
        SSQL &= " CROSS JOIN WQ7S C1" & vbCrLf
        SSQL &= " CROSS JOIN WQ8S C2" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt
    End Function

    ''' <summary>
    ''' 匯出名冊 產投
    ''' </summary>
    Sub ExportXlsStd28()
        Const Cst_FileSavePath As String = "~/SD/15/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const cst_SampleXLS As String = "~\SD\15\SAMPLE_SD15008_1.xlsx" '& cst_files_ext
        'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If

        Dim strErrmsg As String = ""
        Dim sFileName As String = String.Concat("~\SD\15\Temp\", TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim sMyFile1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), sMyFile1, True)
        Catch ex As Exception
            strErrmsg = String.Concat("目錄名稱或磁碟區標籤語法錯誤!!!", vbCrLf, " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)", vbCrLf, ex.Message, vbCrLf)
            Common.MessageBox(Me, strErrmsg)
            TIMS.LOG.Error(ex.Message, ex)
            Return 'Exit Sub
        End Try

        Dim dtXls1 As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim ndt As DateTime = Now
        Dim ROC_Y2 As String = ndt.Year - 1911
        Dim ROC_M2 As String = ndt.Month
        'Dim smYear As String = sm.UserInfo.Years
        Dim V_yearlist As String = TIMS.GetListValue(yearlist)
        Dim ROC_YT As String = If(V_yearlist <> "", String.Concat(CInt(V_yearlist) - 1911, "年度"), "")
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        Dim ROC_Date1 As String = ""
        If STDate1.Text <> "" AndAlso STDate2.Text <> "" Then
            ROC_Date1 = String.Concat(TIMS.Cdate17(STDate1.Text), "~", TIMS.Cdate17(STDate2.Text))
        ElseIf STDate1.Text <> "" AndAlso STDate2.Text = "" Then
            ROC_Date1 = String.Concat("起始", TIMS.Cdate17(STDate1.Text), "~截至為止")
        ElseIf STDate1.Text <> "" AndAlso STDate2.Text <> "" Then
            ROC_Date1 = String.Concat("~截至", TIMS.Cdate17(STDate2.Text))
        End If
        'Dim SF_TITLE1 As String = String.Concat(ROC_Y, "年度訓後動態數據-截至", ROC_Y2, "年", ROC_M2, "月底") '"Excel標題：OOO年度 114/01/01~114/06/30 訓後動態數據 "
        Dim SF_TITLE1 As String = String.Concat(ROC_YT, ROC_Date1, "訓後動態數據")
        Dim s_FILENAME1 As String = String.Concat(ROC_Y2, "年度截至", ROC_M2, "月底產投自主參訓學員訓後動態調查(助益率只取前2題)-", TIMS.GetDateNo2(3))
        Dim fg_RespWriteEnd As Boolean = False
        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial 'ExcelPackage.LicenseContext=LicenseContext.NonCommercial
            'Dim file1 As New FileInfo(filePath1) 'Dim ndt As DateTime = Now

            '開檔
            Using fs1 As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim ep As New ExcelPackage(fs1)
                Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)
                'Dim ep As New ExcelPackage()
                'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)

                ws.Cells("A1:M1").Value = SF_TITLE1
                ws.Cells("A1:M1").Style.Font.Bold = True

                Dim idxStr1 As Integer = 4
                Dim idxStr2 As Integer = 9
                Dim idxStr3 As Integer = 14

                For Each dr As DataRow In dtXls1.Rows
                    Dim SUM2_CTG As Double = 0
                    Dim SUM2_CTW As Double = 0
                    Dim SUM_CTG As Double = 0
                    Dim SUM_CTW As Double = 0

                    SUM_CTG = TIMS.VAL1(dr("SUM_Q217CTG"))
                    SUM_CTW = TIMS.VAL1(dr("SUM_Q217CTW"))
                    Dim I1 As Double = TIMS.VAL1(dr("Q217CTG")) + TIMS.VAL1(dr("Q217CTW"))
                    Dim J1 As Double = I1 / (SUM_CTG + SUM_CTW)
                    ws.Cells("C" & idxStr1).Value = TIMS.VAL1(dr("Q217CTG"))
                    ws.Cells("D" & idxStr1).Value = TIMS.VAL1(dr("Q217CTG")) / SUM_CTG
                    ws.Cells("F" & idxStr1).Value = TIMS.VAL1(dr("Q217CTW"))
                    ws.Cells("G" & idxStr1).Value = TIMS.VAL1(dr("Q217CTW")) / SUM_CTW
                    ws.Cells("I" & idxStr1).Value = I1
                    ws.Cells("J" & idxStr1).Value = J1
                    ws.Cells("D" & idxStr1).Style.Numberformat.Format = "0.00%"
                    ws.Cells("G" & idxStr1).Style.Numberformat.Format = "0.00%"
                    ws.Cells("J" & idxStr1).Style.Numberformat.Format = "0.00%"
                    If idxStr1 = 4 Then
                        SUM2_CTG = TIMS.VAL1(dr("SUM2_Q217CTG"))
                        SUM2_CTW = TIMS.VAL1(dr("SUM2_Q217CTW"))
                        ws.Cells("E" & idxStr1).Value = SUM2_CTG / SUM_CTG
                        ws.Cells("H" & idxStr1).Value = SUM2_CTW / SUM_CTW
                        ws.Cells("K" & idxStr1).Value = (SUM2_CTG + SUM2_CTW) / (SUM_CTG + SUM_CTW)
                        ws.Cells("E" & idxStr1).Style.Numberformat.Format = "0.00%"
                        ws.Cells("H" & idxStr1).Style.Numberformat.Format = "0.00%"
                        ws.Cells("K" & idxStr1).Style.Numberformat.Format = "0.00%"
                    End If

                    SUM_CTG = TIMS.VAL1(dr("SUM_Q218CTG"))
                    SUM_CTW = TIMS.VAL1(dr("SUM_Q218CTW"))
                    Dim I2 As Double = TIMS.VAL1(dr("Q218CTG")) + TIMS.VAL1(dr("Q218CTW"))
                    Dim J2 As Double = I2 / (SUM_CTG + SUM_CTW)
                    ws.Cells("C" & idxStr2).Value = TIMS.VAL1(dr("Q218CTG"))
                    ws.Cells("D" & idxStr2).Value = TIMS.VAL1(dr("Q218CTG")) / SUM_CTG
                    ws.Cells("F" & idxStr2).Value = TIMS.VAL1(dr("Q218CTW"))
                    ws.Cells("G" & idxStr2).Value = TIMS.VAL1(dr("Q218CTW")) / SUM_CTW
                    ws.Cells("I" & idxStr2).Value = I2
                    ws.Cells("J" & idxStr2).Value = J2
                    ws.Cells("D" & idxStr2).Style.Numberformat.Format = "0.00%"
                    ws.Cells("G" & idxStr2).Style.Numberformat.Format = "0.00%"
                    ws.Cells("J" & idxStr2).Style.Numberformat.Format = "0.00%"
                    If idxStr2 = 9 Then
                        SUM2_CTG = TIMS.VAL1(dr("SUM2_Q218CTG"))
                        SUM2_CTW = TIMS.VAL1(dr("SUM2_Q218CTW"))
                        ws.Cells("E" & idxStr2).Value = SUM2_CTG / SUM_CTG
                        ws.Cells("H" & idxStr2).Value = SUM2_CTW / SUM_CTW
                        ws.Cells("K" & idxStr2).Value = (SUM2_CTG + SUM2_CTW) / (SUM_CTG + SUM_CTW)
                        ws.Cells("E" & idxStr2).Style.Numberformat.Format = "0.00%"
                        ws.Cells("H" & idxStr2).Style.Numberformat.Format = "0.00%"
                        ws.Cells("K" & idxStr2).Style.Numberformat.Format = "0.00%"
                    End If

                    SUM_CTG = TIMS.VAL1(dr("SUM_Q217CTG")) + TIMS.VAL1(dr("SUM_Q218CTG"))
                    SUM_CTW = TIMS.VAL1(dr("SUM_Q217CTW")) + TIMS.VAL1(dr("SUM_Q218CTW"))
                    Dim C3 As Double = TIMS.VAL1(dr("Q217CTG")) + TIMS.VAL1(dr("Q218CTG"))
                    Dim D3 As Double = C3 / SUM_CTG
                    Dim F3 As Double = TIMS.VAL1(dr("Q217CTW")) + TIMS.VAL1(dr("Q218CTW"))
                    Dim G3 As Double = F3 / SUM_CTW
                    Dim I3 As Double = C3 + F3
                    Dim J3 As Double = I3 / (SUM_CTG + SUM_CTW)

                    ws.Cells("C" & idxStr3).Value = C3
                    ws.Cells("D" & idxStr3).Value = D3
                    ws.Cells("F" & idxStr3).Value = F3
                    ws.Cells("G" & idxStr3).Value = G3
                    ws.Cells("I" & idxStr3).Value = I3
                    ws.Cells("J" & idxStr3).Value = J3
                    ws.Cells("D" & idxStr3).Style.Numberformat.Format = "0.00%"
                    ws.Cells("G" & idxStr3).Style.Numberformat.Format = "0.00%"
                    ws.Cells("J" & idxStr3).Style.Numberformat.Format = "0.00%"
                    If idxStr3 = 14 Then
                        SUM2_CTG = TIMS.VAL1(dr("SUM2_Q217CTG")) + TIMS.VAL1(dr("SUM2_Q218CTG"))
                        SUM2_CTW = TIMS.VAL1(dr("SUM2_Q217CTW")) + TIMS.VAL1(dr("SUM2_Q218CTW"))
                        ws.Cells("E" & idxStr3).Value = SUM2_CTG / SUM_CTG
                        ws.Cells("H" & idxStr3).Value = SUM2_CTW / SUM_CTW
                        ws.Cells("K" & idxStr3).Value = (SUM2_CTG + SUM2_CTW) / (SUM_CTG + SUM_CTW)
                        ws.Cells("E" & idxStr3).Style.Numberformat.Format = "0.00%"
                        ws.Cells("H" & idxStr3).Style.Numberformat.Format = "0.00%"
                        ws.Cells("K" & idxStr3).Style.Numberformat.Format = "0.00%"
                    End If

                    idxStr1 += 1
                    idxStr2 += 1
                    idxStr3 += 1
                Next

                Dim iFivePoint As Double = 0
                Dim iFiveCnt As Double = 0
                Dim iFivePoint2 As Double = 0
                Dim iFiveCnt2 As Double = 0
                Dim i5 As Double = 5
                For idx As Integer = 4 To 8
                    iFivePoint += TIMS.VAL1(ws.Cells("I" & idx).Value) * i5
                    iFiveCnt += TIMS.VAL1(ws.Cells("I" & idx).Value)
                    iFivePoint2 += TIMS.VAL1(ws.Cells("I" & (idx + 5)).Value) * i5
                    iFiveCnt2 += TIMS.VAL1(ws.Cells("I" & (idx + 5)).Value)
                    i5 -= 1
                Next
                ws.Cells("L4").Value = iFivePoint / iFiveCnt
                ws.Cells("L9").Value = iFivePoint2 / iFiveCnt2
                ws.Cells("M4").Value = iFivePoint / iFiveCnt * 20
                ws.Cells("M9").Value = iFivePoint2 / iFiveCnt2 * 20
                ws.Cells("L4").Style.Numberformat.Format = "0.00"
                ws.Cells("L9").Style.Numberformat.Format = "0.00"
                ws.Cells("M4").Style.Numberformat.Format = "0.00"
                ws.Cells("M9").Style.Numberformat.Format = "0.00"

                ' 設定貨幣格式，小數位數為 0
                'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format="$#,##0" ' 美元符號，您可以根據需要更改
                'ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width=33

                ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom=70 無法運行 修正為 ws.View.ZoomScale=70 才可運行
                'ws.View.ZoomScale = 90

                TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                fg_RespWriteEnd = True
                'TIMS.Utl_RespWriteEnd(Me, objconn, "")
                'Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
                'Select Case V_ExpType
                '    Case "EXCEL"
                '        TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                '        TIMS.Utl_RespWriteEnd(Me, objconn, "")
                '    Case "ODS"
                '        TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                '        TIMS.Utl_RespWriteEnd(Me, objconn, "")
                '    Case Else
                '        Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                '        Common.MessageBox(Me, s_log1)
                '        Return ' Exit Sub
                'End Select
            End Using
            Call TIMS.MyFileDelete(sMyFile1)
            If fg_RespWriteEnd Then TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub

End Class
