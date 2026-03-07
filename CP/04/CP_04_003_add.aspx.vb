Partial Class CP_04_003_add
    Inherits AuthBasePage

    '列印： CP_04_003_Rpt.jrxml SUB_EXPORT
    '"window.open('CP_04_003_01.aspx?ID=" & Request("ID") & "&Student_Data=" & drv("OCID") & "','OCID','width=500,height=500'); return false;"
    'Const cst_printFN1 As String="CP_04_003_Rpt"
    Const cst_printFN1 As String = "CP_04_003_R2"
    Const cst_search As String = "_search" ' Sssion(cst_search)=strSession 'CP_04_003_add.aspx 使用

    Const cst_fmt_date_1 As String = "yyyy.MM.dd"
    Const cst_fmt_date_2 As String = "yyyy.MM.dd HH:mm:ss"

    Const cst_vssort As String = "sort"
    Dim newDistrictName As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Sub UTL_CLS_VIEW_1()
        ViewState("prog") = "" 'TIMS.GetMyValue(Session_search1, "prog")
        ViewState("itemstr") = "" 'TIMS.GetMyValue(Session_search1, "itemstr")
        ViewState("itemplan") = "" ' TIMS.GetMyValue(Session_search1, "itemplan")
        ViewState("SSTDate") = "" 'TIMS.GetMyValue(Session_search1, "SSTDate")
        ViewState("ESTDate") = "" 'TIMS.GetMyValue(Session_search1, "ESTDate")
        ViewState("SFTDate") = "" 'TIMS.GetMyValue(Session_search1, "SFTDate")
        ViewState("EFTDate") = "" 'TIMS.GetMyValue(Session_search1, "EFTDate")
        ViewState("NotOpenStaus") = "" 'TIMS.GetMyValue(Session_search1, "NotOpenStaus")
        ViewState("newDistID") = "" 'TIMS.GetMyValue(Session_search1, "newDistID")
        ViewState("newTPlanID") = "" 'TIMS.GetMyValue(Session_search1, "newTPlanID")
        ViewState("OrgName") = "" 'TIMS.GetMyValue(Session_search1, "OrgName")
        ViewState("OrgID") = "" 'TIMS.GetMyValue(Session_search1, "OrgID")
        ViewState("ClassCName") = "" 'TIMS.GetMyValue(Session_search1, "ClassCName")
        ViewState("TMID") = "" 'TIMS.GetMyValue(Session_search1, "TMID")
        ' ViewState("itembudget")="" 'TIMS.GetMyValue(Session_search1, "itembudget")
        ' ViewState("newBudgetID")="" 'TIMS.GetMyValue(Session_search1, "newBudgetID")
        ViewState("newICityName") = "" 'TIMS.GetMyValue(Session_search1, "newICityName")
        ViewState("newTPlanIDName") = "" 'TIMS.GetMyValue(Session_search1, "newTPlanIDName")
        ViewState("NotOpenStausStr") = "" 'TIMS.GetMyValue(Session_search1, "NotOpenStausStr")
        ViewState("TMIDName") = "" 'TIMS.GetMyValue(Session_search1, "TMIDName")
        ' ViewState("newBudgetName")="" 'TIMS.GetMyValue(Session_search1, "newBudgetName")
        ViewState("RBListExpType") = "" 'TIMS.GetMyValue(s_Sess_search1, "RBListExpType")

        ViewState("itemcity") = ""
    End Sub

    Function CHK_SESSION_1() As String
        Dim rst As String = ""

        Dim s_itemcity_SEL As String = If(Session("itemcity") IsNot Nothing, Convert.ToString(Session("itemcity")), "")
        ViewState("itemcity") = s_itemcity_SEL

        Call UTL_CLS_VIEW_1()
        Dim s_Sess_search1 As String = If(Session(cst_search) IsNot Nothing, Convert.ToString(Session(cst_search)), "")
        If s_Sess_search1 = "" Then Return rst
        ViewState("prog") = TIMS.GetMyValue(s_Sess_search1, "prog")
        ViewState("itemstr") = TIMS.GetMyValue(s_Sess_search1, "itemstr")
        ViewState("itemplan") = TIMS.GetMyValue(s_Sess_search1, "itemplan")
        ViewState("SSTDate") = TIMS.GetMyValue(s_Sess_search1, "SSTDate")
        ViewState("ESTDate") = TIMS.GetMyValue(s_Sess_search1, "ESTDate")
        ViewState("SFTDate") = TIMS.GetMyValue(s_Sess_search1, "SFTDate")
        ViewState("EFTDate") = TIMS.GetMyValue(s_Sess_search1, "EFTDate")
        ViewState("NotOpenStaus") = TIMS.GetMyValue(s_Sess_search1, "NotOpenStaus")
        ViewState("newDistID") = TIMS.GetMyValue(s_Sess_search1, "newDistID")
        ViewState("newTPlanID") = TIMS.GetMyValue(s_Sess_search1, "newTPlanID")
        ViewState("OrgName") = TIMS.GetMyValue(s_Sess_search1, "OrgName")
        ViewState("OrgID") = TIMS.GetMyValue(s_Sess_search1, "OrgID")
        ViewState("ClassCName") = TIMS.GetMyValue(s_Sess_search1, "ClassCName")
        ViewState("TMID") = TIMS.GetMyValue(s_Sess_search1, "TMID")
        ' ViewState("itembudget")=TIMS.GetMyValue(s_Sess_search1, "itembudget")
        ' ViewState("newBudgetID")=TIMS.GetMyValue(s_Sess_search1, "newBudgetID")
        ViewState("newICityName") = TIMS.GetMyValue(s_Sess_search1, "newICityName")
        ViewState("newTPlanIDName") = TIMS.GetMyValue(s_Sess_search1, "newTPlanIDName")
        ViewState("NotOpenStausStr") = TIMS.GetMyValue(s_Sess_search1, "NotOpenStausStr")
        ViewState("TMIDName") = TIMS.GetMyValue(s_Sess_search1, "TMIDName")
        ' ViewState("newBudgetName")=TIMS.GetMyValue(s_Sess_search1, "newBudgetName")
        ViewState("RBListExpType") = TIMS.GetMyValue(s_Sess_search1, "RBListExpType")

        Return rst
    End Function

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        Call CHK_SESSION_1()

        If Not IsPostBack Then
            Call cCreate1()
            ' Button2.Attributes.Add("onclick", "location.href='CP_04_003.aspx?ID=" & Request("ID") & "';return false;")  '回上一頁
        End If
    End Sub

    Sub cCreate1()
        Dim yearlist As String = TIMS.ClearSQM(Request("yearlist"))
        Dim Export As String = TIMS.ClearSQM(Request("export"))
        Dim prog As String = ViewState("prog")

        Dim sql As String = Get_Sqlstr1()
        '開班資料明細**by Milor 20070411** '匯出
        If Export = "Y" Then
            'Call TIMS.WriteTraceLog(sql)
            Call SUB_EXPORT(sql)
            Return
        End If
        '建立最上排基本資料

        Dim dt As DataTable = Nothing
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Const cst_errMsg1 As String = "訓練資料查詢查詢有誤,請重新輸入查詢資料!!"
            Dim strErrmsg As String = ""
            strErrmsg &= cst_errMsg1 & vbCrLf
            strErrmsg += "/* sqlstr: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/* ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, cst_errMsg1)
            Exit Sub
        End Try

        '顯示所選年度
        YearLabel.Text = yearlist
        '顯示所選轄區
        Call area()

        Dim iCOUNT As Integer = 0 '+= 1
        Dim iCNum As Integer = 0 ' += Val(dr("CNum"))'招生人數(CNum/TNum)
        Dim iSNum As Integer = 0 '+= Val(dr("SNum"))
        Dim iESNum As Integer = 0 '+= Val(dr("ESNum"))
        'Dim iESNum2 As Integer=0 '+= Val(dr("ESNum2")) 'Dim iGSNum As Integer=0 '+= Val(dr("GSNum"))
        If TIMS.dtHaveDATA(dt) Then
            For Each dr As DataRow In dt.Rows
                iCOUNT += 1
                iCNum += Val(dr("CNum")) 'CLASS_CLASSINFO.TNUM'招生人數(CNum/TNum)
                iSNum += Val(dr("SNum"))
                iESNum += Val(dr("ESNum"))
                'iESNum2 += Val(dr("ESNum2")) 'iGSNum += Val(dr("GSNum")) 'iGSNum += Val(dr("SUM_INUM"))
            Next
        End If
        CountLabel.Text = iCOUNT
        STNum.Text = iCNum '招生人數(CNum/TNum)
        SSNum.Text = iSNum
        SESNum.Text = iESNum
        ' SESNum2.Text=iESNum2 ' SGSNum.Text=iGSNum

        Const cst_開訓人數 As Integer = 13
        Const cst_結訓人數 As Integer = 14
        'Const cst_離退訓人數 As Integer=15,'Const cst_提前就業人數 As Integer=15,'Const cst_就業人數 As Integer=16,'Const cst_就業關聯人數 As Integer=17,'Const cst_就業累計達1個月以上 As Integer=18,'Const cst_請領津貼人數 As Integer=19,'Const cst_請領津貼就業人數 As Integer=20,
        Label3.Visible = True
        Label5.Visible = True
        'Label2.Visible=True
        SSNum.Visible = True
        SESNum.Visible = True
        ' SGSNum.Visible=True
        DataGrid1.Columns(cst_開訓人數).Visible = True
        DataGrid1.Columns(cst_結訓人數).Visible = True
        ' DataGrid1.Columns(cst_提前就業人數).Visible=True' DataGrid1.Columns(cst_就業人數).Visible=True' DataGrid1.Columns(cst_就業關聯人數).Visible=True' DataGrid1.Columns(cst_就業累計達1個月以上).Visible=True' DataGrid1.Columns(cst_請領津貼人數).Visible=True' DataGrid1.Columns(cst_請領津貼就業人數).Visible=True
        If TIMS.dtNODATA(dt) Then
            NoData.Text = "<font color=red>查無資料</font>"
            DataGrid1.Visible = False
            PageControler1.Visible = False
            btnPrint.Enabled = False
            Exit Sub
        End If

        btnPrint.Enabled = True
        DataGrid1.Visible = True
        '程式來源不同
        If prog = "CP_04_008" Then
            Label3.Visible = False
            Label5.Visible = False
            'Label2.Visible=False
            SSNum.Visible = False
            ' SGSNum.Visible=False
            SESNum.Visible = False
            DataGrid1.Columns(cst_開訓人數).Visible = False
            DataGrid1.Columns(cst_結訓人數).Visible = False
            ' DataGrid1.Columns(cst_提前就業人數).Visible=False
            ' DataGrid1.Columns(cst_就業人數).Visible=False
            ' DataGrid1.Columns(cst_就業關聯人數).Visible=False
            ' DataGrid1.Columns(cst_就業累計達1個月以上).Visible=False
            ' DataGrid1.Columns(cst_請領津貼人數).Visible=False
            ' DataGrid1.Columns(cst_請領津貼就業人數).Visible=False
        End If

        PageControler1.Visible = True
        If ViewState(cst_vssort) = "" Then ViewState(cst_vssort) = "DistID,PlanID,STDate,OCID DESC" '"DistID,PlanID,STDate desc"
        If Convert.ToString(ViewState("SSSDTRID")) = "" Then ViewState("SSSDTRID") = TIMS.GetRnd6Eng()
        PageControler1.SSSDTRID = Convert.ToString(ViewState("SSSDTRID")) 'TIMS.GetRnd6Eng()

        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "OCID"
        PageControler1.Sort = ViewState(cst_vssort) '"DistID,PlanID,STDate,OCID desc"
        PageControler1.ControlerLoad()
        'PageControler1.SqlPrimaryKeyDataCreate(sqlstr, "OCID", "DistID,PlanID,STDate,OCID desc")
    End Sub


    ''' <summary> 匯出 [使用 xls]</summary>
    ''' <param name="Search_Sql"></param>
    Sub SUB_EXPORT(ByVal Search_Sql As String)
        '轄區/'訓練計畫/'訓練機構名稱/'班別名稱/'訓練職類/'訓練性質/'訓練時段/'開訓日期/'結訓日期/'招生人數/'時數/'開訓人數/'結訓人數
        '功能：     首頁>> 學員動態管理 >> 教務報表管理 >> 開班資料(僅此計畫! 自辦沒有)
        '1.增加匯出欄位：「委訓單位名稱」、「委訓單位類型」、「訓練對象」，位置請見圖一
        '2.增加「訓練費用」欄位(位置請見圖一)，其訓練費用欄位之資料由班級申請-【訓練費用】頁籤裡的「總計」欄位資料帶入。(圖二)
        Dim flag_TPlanID07_show As Boolean = False
        If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_TPlanID07_show = True

        '2025年啟用 work2025x01 :2025 政府政策性產業 (產投,自辦)
        Dim fg_SHOW_2025_1 As Boolean = TIMS.SHOW_2025_1(sm)

        Dim dtAll As DataTable = DbAccess.GetDataTable(Search_Sql, objconn) '就業人數 

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("student", System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType="Application/octet-stream"
        'Response.ContentEncoding=System.Text.Encoding.GetEncoding("Big5")

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8) & ".xls")
        ''Response.ContentType="Application/octet-stream"
        'Response.ContentEncoding=System.Text.Encoding.GetEncoding("Big5")
        ''Response.ContentEncoding=System.Text.Encoding.GetEncoding("UTF-8")
        ''Response.ContentType="application/vnd.ms-excel " '內容型態設為Excel
        ''文件內容指定為Excel
        ''Response.ContentType="application/ms-excel;charset=utf-8"
        ''http://stackoverflow.com/questions/974079/setting-mime-type-for-excel-document
        'Response.ContentType="application/x-ms-excel" '內容型態設為Excel
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        'Common.RespWrite(Me, "<html>")
        'Common.RespWrite(Me, "<head>")
        'Common.RespWrite(Me, "<meta http-equiv=""Content-Type"" content=""text/html;charset=BIG5"">")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>

        Dim sFileName1 As String = $"openclasslist{TIMS.GetDateNo3()}"

        '套CSS值
        'http://cosicimiento.blogspot.tw/2008/11/styling-excel-cells-with-mso-number.html
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        'mso-number-format:"0" 
        strSTYLE &= (".NDFormat{mso-number-format:""0"";}") 'NO Decimals
        strSTYLE &= (".NDFormat2{mso-number-format:""\@"";}") 'Text
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '政策性產業類別、報名開始日期、報名結束日期、甄試日期、報名人數、甄試人數、離退訓人數'
        '建立輸出文字
        Dim ExportStr As String = ""
        ExportStr = "<tr>"
        ExportStr &= "<td>轄區</td>"
        ExportStr &= "<td>訓練計畫</td>"
        ExportStr &= "<td>訓練機構名稱</td>"
        ExportStr &= "<td>縣市別</td>"
        ExportStr &= "<td>班別名稱</td>"
        ExportStr &= String.Format("<td>{0}</td>", "期別") '期別 CyclType
        ExportStr &= String.Format("<td>{0}</td>", "班級代碼") '班級代碼 OCID

        ExportStr &= "<td>訓練職類(大類)</td>" 'BUSNAME
        ExportStr &= "<td>訓練職類(中類)</td>" 'JOBNAME
        ExportStr &= "<td>訓練職類(小類)</td>" 'TRAINNAME
        ExportStr &= "<td>通俗職類(大類)</td>" 'CJOBNAME1
        ExportStr &= "<td>通俗職類(中類)</td>" 'CJOBNAME3
        ExportStr &= "<td>通俗職類(小類)</td>" 'CJOBNAME2

        ExportStr &= "<td>訓練性質</td>"
        If flag_TPlanID07_show Then
            ExportStr &= "<td>委訓單位名稱</td>" 'TRNUNITNAME
            ExportStr &= "<td>委訓單位類型</td>" 'TRNUNITCHO_N-委訓單位類型
            ExportStr &= "<td>委訓單位類型(其他說明)</td>" 'TRNUNITTYPE-委訓單位類型(其他說明)
            ExportStr &= "<td>訓練對象</td>" 'TRNUNITEE
        End If
        ExportStr &= "<td>訓練時段</td>"

        ExportStr &= "<td>報名開始日期</td>"
        ExportStr &= "<td>報名結束日期</td>"
        ExportStr &= "<td>甄試日期</td>"
        ExportStr &= "<td>開訓日期</td>"
        ExportStr &= "<td>結訓日期</td>"

        ExportStr &= "<td>招生人數</td>"
        ExportStr &= "<td>時數</td>"
        ExportStr &= "<td>報名人數</td>"
        ExportStr &= "<td>甄試人數</td>"
        ExportStr &= String.Format("<td>{0}</td>", "到考率(%)") '到考率(%) ATTENRATE
        ExportStr &= String.Format("<td>{0}</td>", "錄取人數") '錄取人數 STUDETNUM3
        ExportStr &= String.Format("<td>{0}</td>", "錄取率(%)") '錄取率(%) ACCEPRATE

        ExportStr &= "<td>就安開訓人數</td>" 'SNum02
        ExportStr &= "<td>就保開訓人數</td>" 'SNum03
        ExportStr &= "<td>合計開訓人數</td>" 'SNum
        ExportStr &= String.Format("<td>{0}</td>", "開訓人數比率(%)") '開訓人數比率(%) TRAINRATE

        ExportStr &= "<td>就安結訓人數</td>" 'ESNum02
        ExportStr &= "<td>就保結訓人數</td>" 'ESNum03
        ExportStr &= "<td>合計結訓人數</td>" 'ESNum
        ExportStr &= "<td>就安-離退訓人數</td>" 'JSNum02
        ExportStr &= "<td>就保-離退訓人數</td>" 'JSNum03
        ExportStr &= "<td>合計離退訓人數</td>" 'JSNum

        ExportStr &= String.Format("<td>{0}</td>", "離退訓率(%)") '離退訓率(%) RTIRERATE
        ExportStr &= String.Format("<td>{0}</td>", "第1部分滿意度") '第1部分滿意度 Q1_AVERAGE
        ExportStr &= String.Format("<td>{0}</td>", "第2部分滿意度") '第2部分滿意度 Q2_AVERAGE
        ExportStr &= String.Format("<td>{0}</td>", "第3部分滿意度") '第3部分滿意度 Q3_AVERAGE
        ExportStr &= String.Format("<td>{0}</td>", "第4部分滿意度") '第4部分滿意度 Q4_AVERAGE
        ExportStr &= String.Format("<td>{0}</td>", "第5部分滿意度") '第5部分滿意度 Q5_AVERAGE
        ExportStr &= String.Format("<td>{0}</td>", "平均滿意度") '平均滿意度 AVERAGE

        ExportStr &= "<td>就安-開訓男性人數</td>" 'CNT1M02
        ExportStr &= "<td>就保-開訓男性人數</td>" 'CNT1M03
        ExportStr &= "<td>合計開訓男性人數</td>" 'CNT1M
        ExportStr &= "<td>就安-開訓女性人數</td>" 'CNT1F02
        ExportStr &= "<td>就保-開訓女性人數</td>" 'CNT1F03
        ExportStr &= "<td>合計開訓女性人數</td>" 'CNT1F
        '開訓15~19歲人數、開訓20~24歲人數、開訓25~29歲人數、開訓30~34歲人數、開訓35~39歲人數、
        '開訓40~44歲人數、開訓45~49歲人數、開訓50~54歲人數、開訓55~59歲人數、開訓60~64歲人數、開訓65歲以上人數。
        ExportStr &= "<td>開訓15~19歲人數</td>"
        ExportStr &= "<td>開訓20~24歲人數</td>"
        ExportStr &= "<td>開訓25~29歲人數</td>"
        ExportStr &= "<td>開訓30~34歲人數</td>"
        ExportStr &= "<td>開訓35~39歲人數</td>"
        ExportStr &= "<td>開訓40~44歲人數</td>"
        ExportStr &= "<td>開訓45~49歲人數</td>"
        ExportStr &= "<td>開訓50~54歲人數</td>"
        ExportStr &= "<td>開訓55~59歲人數</td>"
        ExportStr &= "<td>開訓60~64歲人數</td>"
        ExportStr &= "<td>開訓65歲以上人數</td>"
        If flag_TPlanID07_show Then
            ExportStr &= "<td>訓練費用</td>" 'TOTALCOST
        Else
            ExportStr &= "<td>每人政府負擔費用</td>"
            ExportStr &= "<td>每人個人負擔費用</td>"
        End If
        'ExportStr &= "<td>在職者</td>" 'SUM_ISWORK,'ExportStr &= "<td>不就業人數</td>" 'NOJOBX,'ExportStr &= "<td>訓練職類(大類)</td>" 'BUSName,
        ExportStr &= "<td>5+2產業創新計畫</td>"
        ExportStr &= "<td>台灣AI行動計畫</td>"
        ExportStr &= "<td>數位國家創新經濟發展方案</td>"
        ExportStr &= "<td>國家資通安全發展方案</td>"
        ExportStr &= "<td>前瞻基礎建設計畫</td>"
        ExportStr &= "<td>新南向政策</td>"
        If fg_SHOW_2025_1 Then
            Dim D25NMs As String() = "亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策".Split(",")
            For Each SD25V1 As String In D25NMs
                ExportStr &= $"<td>{SD25V1}</td>"
            Next
        End If
        ExportStr &= "<td>訓練課程類型</td>" '、訓練課程類型 ADVANCE_N
        ExportStr &= "<td>訓練時段2</td>" '、訓練時段2 NOTE3
        ExportStr &= "<td>是否輔導考照</td>" '、是否輔導考照 COACHING_N
        For i As Integer = 1 To 3
            ExportStr &= String.Concat("<td>可參加檢定職類群(", i, ")</td>") '、可參加檢定職類群(1) JGNAME1
            ExportStr &= String.Concat("<td>可參加檢定職類(", i, ")</td>") '、可參加檢定職類(1) EXAMNAME1
            ExportStr &= String.Concat("<td>級別(", i, ")</td>") '、級別(1) EXAMLVN1
        Next
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dtAll.Rows
            iNum += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", dr("DistName")) '轄區
            ExportStr &= String.Format("<td>{0}</td>", dr("PlanName")) '訓練計畫
            ExportStr &= String.Format("<td>{0}</td>", dr("OrgName")) '訓練機構名稱
            ExportStr &= String.Format("<td>{0}</td>", dr("CityName")) '縣市別
            ExportStr &= String.Format("<td>{0}</td>", dr("ClassCName")) '班別名稱
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CyclType"))) ' 期別 CyclType
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("OCID"))) '班級代碼 OCID

            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("BUSNAME"))) '訓練職類(大類)
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("JOBNAME"))) '訓練職類(中類)
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("TRAINNAME"))) '訓練職類(小類)
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CJOBNAME1"))) '通俗職類(大類)
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CJOBNAME3"))) '通俗職類(中類)
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CJOBNAME2"))) '通俗職類(小類)

            ExportStr &= "<td> " & dr("TPropertyIDN").ToString & "</td>"  '訓練性質
            If flag_TPlanID07_show Then
                ExportStr &= "<td> " & dr("TRNUNITNAME").ToString & "</td>" '委訓單位名稱
                ExportStr &= "<td> " & Convert.ToString(dr("TRNUNITCHO_N")) & "</td>" 'TRNUNITCHO_N-委訓單位類型
                ExportStr &= "<td> " & Convert.ToString(dr("TRNUNITTYPE")) & "</td>" 'TRNUNITTYPE-委訓單位類型(其他說明)
                ExportStr &= "<td> " & dr("TRNUNITEE").ToString & "</td>" '訓練對象
            End If
            ExportStr &= "<td> " & dr("HourRanName").ToString & "</td>"  '訓練時段

            ExportStr &= "<td> " & TIMS.Cdate3(dr("SENTERDATE"), cst_fmt_date_2) & "</td>" '報名開始日期
            ExportStr &= "<td> " & TIMS.Cdate3(dr("FENTERDATE"), cst_fmt_date_2) & "</td>" '報名結束日期
            ExportStr &= "<td> " & TIMS.Cdate3(dr("EXAMDATE"), cst_fmt_date_1) & "</td>" '甄試日期
            ExportStr &= "<td> " & TIMS.Cdate3(dr("STDATE"), cst_fmt_date_1) & "</td>" '開訓日期
            ExportStr &= "<td> " & TIMS.Cdate3(dr("FTDATE"), cst_fmt_date_1) & "</td>" '結訓日期

            ExportStr &= "<td> " & Convert.ToString(dr("CNum")) & "</td>" '招生人數(CNum/TNum)
            ExportStr &= "<td> " & Convert.ToString(dr("THours")) & "</td>"  '時數
            ExportStr &= "<td> " & Convert.ToString(dr("STUDETNUM")) & "</td>" '報名人數
            ExportStr &= "<td> " & Convert.ToString(dr("STUDETNUM2")) & "</td>" '甄試人數
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("ATTENRATE"))) '到考率(%) ATTENRATE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("STUDETNUM3"))) '錄取人數 STUDETNUM3
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("ACCEPRATE"))) '錄取率(%) ACCEPRATE

            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("SNum02"))) '"<td>就安開訓人數</td>" 'SNum02
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("SNum03"))) '"<td>就保開訓人數</td>" 'SNum03
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("SNum"))) '"<td>合計開訓人數</td>" 'SNum
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("TRAINRATE"))) '開訓人數比率(%) TRAINRATE

            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("ESNum02"))) '"<td>就安結訓人數</td>" 'ESNum02
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("ESNum03"))) '"<td>就保結訓人數</td>" 'ESNum03
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("ESNum"))) '"<td>合計結訓人數</td>" 'ESNum
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("JSNum02"))) '"<td>就安-離退訓人數</td>" 'JSNum02
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("JSNum03"))) '"<td>就保-離退訓人數</td>" 'JSNum03
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("JSNum"))) '"<td>合計離退訓人數</td>" 'JSNum

            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("RTIRERATE"))) '離退訓率(%) RTIRERATE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("Q1_AVERAGE"))) '第1部分滿意度 Q1_AVERAGE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("Q2_AVERAGE"))) '第2部分滿意度 Q2_AVERAGE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("Q3_AVERAGE"))) '第3部分滿意度 Q3_AVERAGE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("Q4_AVERAGE"))) '第4部分滿意度 Q4_AVERAGE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("Q5_AVERAGE"))) '第5部分滿意度 Q5_AVERAGE
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("AVERAGE"))) '平均滿意度 AVERAGE

            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CNT1M02"))) '"<td>就安-開訓男性人數</td>" 'CNT1M02
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CNT1M03"))) '"<td>就保-開訓男性人數</td>" 'CNT1M03
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CNT1M"))) '"<td>合計開訓男性人數</td>" 'CNT1M
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CNT1F02"))) '"<td>就安-開訓女性人數</td>" 'CNT1F02
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CNT1F03"))) '"<td>就保-開訓女性人數</td>" 'CNT1F03
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CNT1F"))) '"<td>合計開訓女性人數</td>" 'CNT1F

            '開訓15~19歲人數、開訓20~24歲人數、開訓25~29歲人數、開訓30~34歲人數、開訓35~39歲人數、
            '開訓40~44歲人數、開訓45~49歲人數、開訓50~54歲人數、開訓55~59歲人數、開訓60~64歲人數、開訓65歲以上人數。
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD15"))) ' "<td>開訓15~19歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD20"))) ' "<td>開訓20~24歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD25"))) ' "<td>開訓25~29歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD30"))) ' "<td>開訓30~34歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD35"))) ' "<td>開訓35~39歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD40"))) ' "<td>開訓40~44歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD45"))) ' "<td>開訓45~49歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD50"))) ' "<td>開訓50~54歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD55"))) ' "<td>開訓55~59歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD60"))) ' "<td>開訓60~64歲人數</td>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("YOD65"))) ' "<td>開訓65歲以上人數</td>"
            If flag_TPlanID07_show Then
                ExportStr &= "<td> " & dr("TOTALCOST").ToString & "</td>"
            Else
                ExportStr &= "<td> " & dr("DEFGOVCOST").ToString & "</td>"
                ExportStr &= "<td> " & dr("DEFSTDCOST").ToString & "</td>"
            End If
            'ExportStr &= "<td>" & dr("SUM_ISWORK").ToString & "</td>" 'SUM_ISWORK'在職者
            'ExportStr &= "<td>" & dr("NOJOBX").ToString & "</td>" 'NOJOBX'不就業人數
            'ExportStr &= "<td>" & dr("BUSName").ToString & "</td>" 'BUSName'訓練職類(大類)
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME1")) & "</td>" '5+2產業創新計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME2")) & "</td>" '台灣AI行動計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME3")) & "</td>" '數位國家創新經濟發展方案
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME4")) & "</td>" '國家資通安全發展方案
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME5")) & "</td>" '前瞻基礎建設計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME6")) & "</td>" '新南向政策
            If fg_SHOW_2025_1 Then
                Dim D25VALs As String() = "D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6".Split(",")
                For Each SD25V2 As String In D25VALs
                    ExportStr &= $"<td>{dr(SD25V2)}</td>"
                Next
            End If
            ExportStr &= String.Format("<td>{0}</td>", dr("ADVANCE_N")) '訓練課程類型 ADVANCE_N
            ExportStr &= String.Format("<td>{0}</td>", dr("NOTE3")) '訓練時段2 NOTE3
            ExportStr &= String.Format("<td>{0}</td>", dr("COACHING_N")) '是否輔導考照 COACHING_N

            For i As Integer = 1 To 3
                Dim c_JGNAME_N As String = String.Concat("JGNAME", i)
                Dim c_EXAMNAME_N As String = String.Concat("EXAMNAME", i)
                Dim c_EXAMLVN_N As String = String.Concat("EXAMLVN", i)
                ExportStr &= String.Format("<td>{0}</td>", dr(c_JGNAME_N)) '可參加檢定職類群(1) JGNAME1
                ExportStr &= String.Format("<td>{0}</td>", dr(c_EXAMNAME_N)) '可參加檢定職類(1) EXAMNAME1
                ExportStr &= String.Format("<td>{0}</td>", dr(c_EXAMLVN_N)) '級別(1) EXAMLVN1
            Next

            ExportStr &= "</tr>"
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim v_ExpType As String = ViewState("RBListExpType") 'String'TIMS.GetListValue(RBListExpType)
        Dim parmsExp As New Hashtable From {
            {"ExpType", v_ExpType}, 'EXCEL/PDF/ODS
            {"FileName", sFileName1},
            {"strSTYLE", strSTYLE},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Call TIMS.CloseDbConn(objconn) 'Response.End()
    End Sub

    ''' <summary>班級範圍</summary>
    ''' <returns></returns>
    Function W_CLS_SQL_1() As String
        Dim yearlist As String = TIMS.ClearSQM(Request("yearlist"))
        Dim Export As String = TIMS.ClearSQM(Request("export"))
        Dim prog As String = ViewState("prog")
        Dim itemstr As String = ViewState("itemstr")
        Dim itemplan As String = ViewState("itemplan")
        Dim itemcity As String = ViewState("itemcity")
        Dim SSTDate As String = ViewState("SSTDate")
        Dim ESTDate As String = ViewState("ESTDate")
        Dim SFTDate As String = ViewState("SFTDate")
        Dim EFTDate As String = ViewState("EFTDate")
        Dim NotOpenStaus As String = ViewState("NotOpenStaus")
        'Dim itembudget As String= ViewState("itembudget")

        Dim vs_TMID As String = ViewState("TMID")
        Dim vs_OrgID As String = ViewState("OrgID")
        Dim vs_OrgName As String = ViewState("OrgName")
        Dim vs_ClassCName As String = ViewState("ClassCName")

        '訓練職類(中類)、通俗職類(大類)、通俗職類(中類)、通俗職類(小類)、
        '就安開訓人數、就保開訓人數、就安結訓人數、就保結訓人數、就安離退訓人數、就保離退訓人數、就安開訓男性人數、就保開訓男性人數、就安開訓女性人數、就保開訓女性人數
        '招生人數(CNum/TNum)
        Dim sql As String = ""
        sql &= " SELECT a.OCID,a.ClassCName,a.IsClosed,a.TNum CNum,a.STDATE,a.FTDATE,a.SENTERDATE,a.FENTERDATE,a.EXAMDATE,a.THours,a.TPropertyID,a.PlanID" & vbCrLf
        'sql &= " ,Case When a.TPropertyID=0 Then '職前'WHEN a.TPropertyID=1 THEN '在職'WHEN a.TPropertyID=2 THEN '委託訓練'ELSE CONVERT(VARCHAR, a.TPropertyID) END TPropertyIDN" & vbCrLf
        sql &= " ,dbo.FN_GET_TPROPERTY(a.TPropertyID) TPropertyIDN" & vbCrLf
        sql &= " ,ip.DistID,ip.DISTNAME,ip.PlanName" & vbCrLf
        '需求編號：   OJT-20080601
        sql &= " ,d.BUSNAME" & vbCrLf '訓練職類(大類)
        sql &= " ,d.JOBNAME" & vbCrLf '訓練職類(中類)
        sql &= " ,d.TRAINNAME" & vbCrLf '訓練職類(小類)
        sql &= " ,e.HourRanName" & vbCrLf
        sql &= " ,i.OrgName" & vbCrLf
        sql &= " ,a.CyclType" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.CyclType) CLASSCNAME2" & vbCrLf
        sql &= " ,COALESCE(iz.CTID,iz1.CTID,iz2.CTID) CTID" & vbCrLf
        sql &= " ,COALESCE(iz.CTName,iz1.CTName,iz2.CTName) CITYNAME" & vbCrLf
        sql &= " ,j.DEFGOVCOST" & vbCrLf
        sql &= " ,j.DEFSTDCOST" & vbCrLf
        '/*取到小數點第三位*/
        sql &= " ,CASE WHEN ISNULL(a.TNUM,0)=0 THEN NULL ELSE ROUND(CAST(j.DEFGOVCOST/a.TNUM AS FLOAT),3) END DEFGOVCOST1" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(a.TNUM,0)=0 THEN NULL ELSE ROUND(CAST(j.DEFSTDCOST/a.TNUM AS FLOAT),3) END DEFSTDCOST1" & vbCrLf

        sql &= " ,j.TRNUNITNAME" & vbCrLf '委訓單位名稱
        '委訓單位類型分成兩欄：
        '1.委訓單位類型：政府機關 / 公民營事業機構 / 學校 / 團體 / 其他
        '2.委訓單位類型(其他說明)
        sql &= " ,cd1.CODE_CNAME TRNUNITCHO_N" & vbCrLf
        sql &= " ,j.TRNUNITTYPE" & vbCrLf '委訓單位類型-'2.委訓單位類型(其他說明)
        sql &= " ,j.TRNUNITEE" & vbCrLf '訓練對象
        sql &= " ,j.TOTALCOST" & vbCrLf '訓練費用
        sql &= " ,D2.KID20,D2.D20KNAME1,D2.D20KNAME2,D2.D20KNAME3,D2.D20KNAME4,D2.D20KNAME5,D2.D20KNAME6" & vbCrLf
        ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
        sql &= " ,D2.D25KNAME1,D2.D25KNAME2,D2.D25KNAME3,D2.D25KNAME4,D2.D25KNAME5,D2.D25KNAME6,D2.D25KNAME7,D2.D25KNAME8" & vbCrLf
        '需求編號：   OJT-20080601
        '通俗職類(大類)、通俗職類(中類)、通俗職類(小類)
        sql &= " ,b3.CJOBNAME1,b3.CJOBNAME3,b3.CJOBNAME2" & vbCrLf
        '訓練課程類型 ADVANCE /訓練時段2 NOTE3 /是否輔導考照 COACHING
        sql &= " ,j.ADVANCE,j.NOTE3,j.COACHING" & vbCrLf
        '可參加檢定職類群 JGNAME1/可參加檢定職類 EXAMNAME1/級別 EXAMLVN1
        sql &= " ,j.EXAMIDS1,j.EXAMLVID1" & vbCrLf
        sql &= " ,j.EXAMIDS2,j.EXAMLVID2" & vbCrLf
        sql &= " ,j.EXAMIDS3,j.EXAMLVID3" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAM3('JG',J.EXAMIDS1) JGNAME1" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAM3('NM',J.EXAMIDS1) EXAMNAME1" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAMLV(J.EXAMLVID1) EXAMLVN1" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAM3('JG',J.EXAMIDS2) JGNAME2" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAM3('NM',J.EXAMIDS2) EXAMNAME2" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAMLV(J.EXAMLVID2) EXAMLVN2" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAM3('JG',J.EXAMIDS3) JGNAME3" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAM3('NM',J.EXAMIDS3) EXAMNAME3" & vbCrLf
        'sql &= " ,dbo.FN_GET_EXAMLV(J.EXAMLVID3) EXAMLVN3" & vbCrLf

        sql &= " FROM dbo.CLASS_CLASSINFO a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP h WITH(NOLOCK) ON h.RID=a.RID" & vbCrLf
        sql &= " JOIN dbo.ID_CLASS c WITH(NOLOCK) ON c.CLSID=a.CLSID" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE d ON d.TMID=a.TMID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip ON ip.PlanID=a.PlanID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO i WITH(NOLOCK) ON i.ComIDNO=a.ComIDNO AND i.OrgID=h.OrgID" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO j WITH(NOLOCK) ON j.PlanID=a.PlanID AND j.ComIDNO=a.ComIDNO AND j.SeqNO=a.SeqNO" & vbCrLf
        sql &= " JOIN dbo.V_SHARECJOB b3 ON j.CJOB_UNKEY=b3.CJOB_UNKEY" & vbCrLf

        sql &= " LEFT JOIN dbo.KEY_HOURRAN e ON e.HRID=a.TPeriod" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz ON iz.ZipCode=a.TaddressZip" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_TRAINPLACE sp ON sp.PTID=j.AddressSciPTID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz1 ON iz1.zipCode=sp.ZipCode" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_TRAINPLACE tp ON tp.PTID=j.AddressTechPTID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz2 ON iz2.zipCode=tp.ZipCode" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT D2 ON D2.PLANID=J.PLANID AND D2.COMIDNO=J.COMIDNO AND D2.SEQNO=J.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.SYS_SHAREDCODE CD1 WITH(NOLOCK) ON CD1.CODE_KIND='TRNUNITCHO' AND CD1.CODE_ID=j.TRNUNITCHO" & vbCrLf

        '是否轉入成功'PLAN_PLANINFO
        sql &= " WHERE a.IsSuccess='Y'AND j.IsApprPaper='Y' AND j.AppliedResult='Y' AND j.TransFlag='Y'" & vbCrLf
        '程式來源不同
        If prog = "CP_04_008" Then sql &= " AND a.STDate<=GETDATE()" & vbCrLf '已開訓
        '開訓日期起
        If SSTDate <> "" Then sql &= " AND a.STDate>=" & TIMS.To_date(SSTDate) & vbCrLf
        '開訓日期迄
        If ESTDate <> "" Then sql &= " AND a.STDate<=" & TIMS.To_date(ESTDate) & vbCrLf '" & ESTDate & "'" & vbCrLf
        '結訓日期起
        If SFTDate <> "" Then sql &= " AND a.FTDate>=" & TIMS.To_date(SFTDate) & vbCrLf '" & SFTDate & "'" & vbCrLf
        '結訓日期迄
        If EFTDate <> "" Then sql &= " AND a.FTDate<=" & TIMS.To_date(EFTDate) & vbCrLf '" & EFTDate & "'" & vbCrLf
        '選擇年度
        If yearlist <> "" Then sql &= " AND ip.Years='" & yearlist & "'" & vbCrLf
        '選擇轄區
        If itemstr <> "" Then sql &= " AND ip.DistID IN (" & itemstr & ")" & vbCrLf
        '選擇縣市
        If itemcity <> "" Then
            sql &= $" AND (iz.CTID IN ({itemcity}) OR iz1.CTID IN ({itemcity}) OR iz2.CTID IN ({itemcity}))" & vbCrLf
        End If
        '選擇訓練計畫
        If itemplan <> "" Then itemplan = TIMS.CombiSQM2IN(itemplan) '重新組合加入單引號
        If itemplan <> "" Then sql &= $" AND ip.TPlanID IN ({itemplan})" & vbCrLf
        '開班狀態
        If NotOpenStaus <> "" Then sql &= " AND a.NotOpen='" & NotOpenStaus & "'" & vbCrLf
        '職類
        If vs_TMID <> "" Then sql &= " AND a.TMID='" & vs_TMID & "'" & vbCrLf

        '機構OrgID若有輸入，則不比對OrgName
        If vs_OrgID <> "" Then
            sql &= " AND i.OrgID LIKE '%" & vs_OrgID & "%'" & vbCrLf
        Else
            '機構名稱
            If vs_OrgName <> "" Then
                sql &= " AND i.OrgName LIKE '%" & vs_OrgName & "%'" & vbCrLf
            End If
        End If
        '班級名稱
        If vs_ClassCName <> "" Then
            sql &= " AND a.ClassCName LIKE '%" & vs_ClassCName & "%'" & vbCrLf
        End If
        Return sql
    End Function

    ''' <summary>
    ''' 學員統計範圍
    ''' </summary>
    ''' <returns></returns>
    Function W_STD_SQL_2() As String
        '需求編號：   OJT-20080601 '就安開訓人數、就保開訓人數、就安結訓人數、就保結訓人數、就安離退訓人數、就保離退訓人數、就安開訓男性人數、就保開訓男性人數、就安開訓女性人數、就保開訓女性人數
        Dim sql As String = ""
        sql &= " SELECT cs.OCID" & vbCrLf
        '01公務02就安03就保04再出發
        '就安/'就保開訓人數
        sql &= " ,COUNT(CASE when cs.BUDGETID='02' then 1 end) SNum02" & vbCrLf
        sql &= " ,COUNT(CASE when cs.BUDGETID='03' then 1 end) SNum03" & vbCrLf
        sql &= " ,COUNT(1) SNum" & vbCrLf '開訓人數
        '就安/'就保結訓人數
        sql &= " ,COUNT(CASE when cs.StudStatus=5 and cs.BUDGETID='02' then 1 end) ESNum02" & vbCrLf
        sql &= " ,COUNT(CASE when cs.StudStatus=5 and cs.BUDGETID='03' then 1 end) ESNum03" & vbCrLf
        sql &= " ,COUNT(case when cs.StudStatus=5 then 1 end) ESNum" & vbCrLf '結訓人數
        '就安/'就保離退訓人數
        sql &= " ,COUNT(CASE when cs.StudStatus IN (2,3) and cs.BUDGETID='02' then 1 end) JSNum02" & vbCrLf
        sql &= " ,COUNT(CASE when cs.StudStatus IN (2,3) and cs.BUDGETID='03' then 1 end) JSNum03" & vbCrLf
        sql &= " ,COUNT(case when cs.StudStatus IN (2,3) then 1 end) JSNum" & vbCrLf '離退訓人數
        '就安/'就保開訓男性人數
        sql &= " ,COUNT(CASE WHEN SS.SEX ='M' and cs.BUDGETID='02' THEN 1 END) CNT1M02" & vbCrLf
        sql &= " ,COUNT(CASE WHEN SS.SEX ='M' and cs.BUDGETID='03' THEN 1 END) CNT1M03" & vbCrLf
        sql &= " ,COUNT(CASE WHEN SS.SEX ='M' THEN 1 END) CNT1M" & vbCrLf
        '就安/'就保開訓女性人數
        sql &= " ,COUNT(CASE WHEN SS.SEX ='F' and cs.BUDGETID='02' THEN 1 END) CNT1F02" & vbCrLf
        sql &= " ,COUNT(CASE WHEN SS.SEX ='F' and cs.BUDGETID='03' THEN 1 END) CNT1F03" & vbCrLf
        sql &= " ,COUNT(CASE WHEN SS.SEX ='F' THEN 1 END) CNT1F" & vbCrLf
        '開訓15~19歲人數、開訓20~24歲人數、開訓25~29歲人數、開訓30~34歲人數、開訓35~39歲人數、
        '開訓40~44歲人數、開訓45~49歲人數、開訓50~54歲人數、開訓55~59歲人數、開訓60~64歲人數、開訓65歲以上人數。
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=15 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=19 THEN 1 END) YOD15" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=20 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=24 THEN 1 END) YOD20" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=25 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=29 THEN 1 END) YOD25" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=30 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=34 THEN 1 END) YOD30" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=35 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=39 THEN 1 END) YOD35" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=40 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=44 THEN 1 END) YOD40" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=45 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=49 THEN 1 END) YOD45" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=50 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=54 THEN 1 END) YOD50" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=55 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=59 THEN 1 END) YOD55" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=60 AND dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)<=64 THEN 1 END) YOD60" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY)>=65 THEN 1 END) YOD65" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS cs WITH(NOLOCK) ON cs.OCID=cc.OCID" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO ss WITH(NOLOCK) ON ss.sid=cs.sid" & vbCrLf
        sql &= " WHERE cs.MAKESOCID IS NULL" & vbCrLf
        sql &= " GROUP BY cs.OCID" & vbCrLf
        Return sql
    End Function

    ''' <summary>
    ''' 報名學員統計
    ''' </summary>
    ''' <returns></returns>
    Function W_ETR_SQL_3() As String
        Dim sql As String = ""
        sql &= " SELECT cc.OCID" & vbCrLf
        '報名人數
        sql &= " ,COUNT(1) STUDETNUM" & vbCrLf
        '甄試人數
        sql &= " ,COUNT(CASE WHEN b.TOTALRESULT>=0 THEN 1 END) STUDETNUM2" & vbCrLf
        '錄取人數
        '2、錄取人數：為【錄訓作業】功能中的「正取」人數。
        sql &= " ,COUNT(CASE WHEN c.SELRESULTID ='01' THEN 1 END) STUDETNUM3" & vbCrLf

        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b WITH(NOLOCK) on b.OCID1=cc.OCID" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a WITH(NOLOCK) on a.setid=b.setid" & vbCrLf
        sql &= " LEFT JOIN STUD_SELRESULT c WITH(NOLOCK) on c.setid=b.setid and c.enterdate=b.enterdate and c.sernum=b.sernum and c.ocid=b.ocid1" & vbCrLf
        sql &= " GROUP BY cc.OCID" & vbCrLf
        Return sql
    End Function

    Function Get_Sqlstr1() As String
        Dim str_WC1 As String = W_CLS_SQL_1() '班級範圍 -WC1 cc
        Dim str_WC2 As String = W_STD_SQL_2() '學員統計範圍 -WC2 gs
        Dim str_WC3 As String = W_ETR_SQL_3() '報名學員統計 -WC3 gs3

        Dim sql As String = ""
        sql &= " WITH WC1 AS (" & str_WC1 & ")"
        sql &= " ,WC2 AS (" & str_WC2 & ")"
        sql &= " ,WC3 AS (" & str_WC3 & ")"

        sql &= " SELECT cc.ClassCName" & vbCrLf
        sql &= " ,cc.IsClosed" & vbCrLf
        sql &= " ,cc.CNum" & vbCrLf '招生人數(CNum/TNum)
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,FORMAT(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        'sql &= " ,FORMAT(cc.SENTERDATE,'yyyy/MM/dd') SENTERDATE" & vbCrLf
        'sql &= " ,FORMAT(cc.FENTERDATE,'yyyy/MM/dd') FENTERDATE" & vbCrLf
        sql &= " ,cc.SENTERDATE" & vbCrLf
        sql &= " ,cc.FENTERDATE" & vbCrLf
        sql &= " ,FORMAT(cc.EXAMDATE,'yyyy/MM/dd') EXAMDATE" & vbCrLf
        sql &= " ,cc.DistID" & vbCrLf
        sql &= " ,cc.PlanID" & vbCrLf
        sql &= " ,cc.DistName" & vbCrLf
        sql &= " ,cc.THours" & vbCrLf
        sql &= " ,cc.TPropertyID" & vbCrLf
        sql &= " ,cc.TPropertyIDN" & vbCrLf
        '需求編號： OJT-20080601
        sql &= " ,cc.BUSNAME" & vbCrLf '訓練職類(大類)
        sql &= " ,cc.JOBNAME" & vbCrLf '訓練職類(中類)
        sql &= " ,cc.TRAINNAME" & vbCrLf '訓練職類(小類)
        '通俗職類(大類)、通俗職類(中類)、通俗職類(小類)
        sql &= " ,cc.CJOBNAME1,cc.CJOBNAME3,cc.CJOBNAME2" & vbCrLf

        sql &= " ,cc.HourRanName" & vbCrLf
        sql &= " ,cc.PlanName" & vbCrLf
        sql &= " ,cc.OrgName" & vbCrLf
        sql &= " ,cc.CyclType" & vbCrLf
        sql &= " ,cc.ClassCName2" & vbCrLf
        sql &= " ,cc.CTID" & vbCrLf
        sql &= " ,cc.CityName" & vbCrLf
        sql &= " ,cc.DEFGOVCOST" & vbCrLf
        sql &= " ,cc.DEFSTDCOST" & vbCrLf
        sql &= " ,cc.TRNUNITNAME" & vbCrLf '委訓單位名稱
        sql &= " ,cc.TRNUNITCHO_N" & vbCrLf '委訓單位類型-N
        sql &= " ,cc.TRNUNITTYPE" & vbCrLf '委訓單位類型(其他說明)
        sql &= " ,cc.TRNUNITEE" & vbCrLf '訓練對象
        sql &= " ,cc.TOTALCOST" & vbCrLf '訓練費用

        sql &= " ,cc.KID20,cc.D20KNAME1,cc.D20KNAME2,cc.D20KNAME3,cc.D20KNAME4,cc.D20KNAME5,cc.D20KNAME6" & vbCrLf
        sql &= " ,cc.D25KNAME1,cc.D25KNAME2,cc.D25KNAME3,cc.D25KNAME4,cc.D25KNAME5,cc.D25KNAME6,cc.D25KNAME7,cc.D25KNAME8" & vbCrLf

        '訓練課程類型 ADVANCE /訓練時段2 NOTE3 /是否輔導考照 COACHING
        sql &= " ,cc.ADVANCE,CASE cc.ADVANCE WHEN '01' THEN '基礎' WHEN '02' THEN '進階' END ADVANCE_N" & vbCrLf
        sql &= " ,cc.NOTE3" & vbCrLf
        sql &= " ,cc.COACHING,CASE cc.COACHING WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END COACHING_N" & vbCrLf
        '可參加檢定職類群 JGNAME1/可參加檢定職類 EXAMNAME1/級別 EXAMLVN1
        sql &= " ,dbo.FN_GET_EXAM3('JG',cc.EXAMIDS1) JGNAME1" & vbCrLf
        sql &= " ,dbo.FN_GET_EXAM3('NM',cc.EXAMIDS1) EXAMNAME1" & vbCrLf
        sql &= " ,dbo.FN_GET_EXAMLV(cc.EXAMLVID1) EXAMLVN1" & vbCrLf

        sql &= " ,dbo.FN_GET_EXAM3('JG',cc.EXAMIDS2) JGNAME2" & vbCrLf
        sql &= " ,dbo.FN_GET_EXAM3('NM',cc.EXAMIDS2) EXAMNAME2" & vbCrLf
        sql &= " ,dbo.FN_GET_EXAMLV(cc.EXAMLVID2) EXAMLVN2" & vbCrLf

        sql &= " ,dbo.FN_GET_EXAM3('JG',cc.EXAMIDS3) JGNAME3" & vbCrLf
        sql &= " ,dbo.FN_GET_EXAM3('NM',cc.EXAMIDS3) EXAMNAME3" & vbCrLf
        sql &= " ,dbo.FN_GET_EXAMLV(cc.EXAMLVID3) EXAMLVN3" & vbCrLf

        '避免有班級為空學員資料
        sql &= " ,ISNULL(gs.SNum,0) SNum" & vbCrLf '開訓人數
        sql &= " ,ISNULL(gs.ESNum,0) ESNum" & vbCrLf '結訓人數
        sql &= " ,ISNULL(gs.JSNum,0) JSNum" & vbCrLf '離退訓人數
        sql &= " ,ISNULL(gs.CNT1M,0) CNT1M" & vbCrLf '開訓男性人數
        sql &= " ,ISNULL(gs.CNT1F,0) CNT1F" & vbCrLf '開訓女性人數
        '01公務02就安03就保04再出發
        '就安開訓人數、就保開訓人數、就安結訓人數、就保結訓人數、就安離退訓人數、就保離退訓人數、就安開訓男性人數、就保開訓男性人數、就安開訓女性人數、就保開訓女性人數
        sql &= " ,ISNULL(gs.SNum02,0) SNum02" & vbCrLf
        sql &= " ,ISNULL(gs.ESNum02,0) ESNum02" & vbCrLf
        sql &= " ,ISNULL(gs.JSNum02,0) JSNum02" & vbCrLf
        sql &= " ,ISNULL(gs.CNT1M02,0) CNT1M02" & vbCrLf
        sql &= " ,ISNULL(gs.CNT1F02,0) CNT1F02" & vbCrLf

        sql &= " ,ISNULL(gs.SNum03,0) SNum03" & vbCrLf
        sql &= " ,ISNULL(gs.ESNum03,0) ESNum03" & vbCrLf
        sql &= " ,ISNULL(gs.JSNum03,0) JSNum03" & vbCrLf
        sql &= " ,ISNULL(gs.CNT1M03,0) CNT1M03" & vbCrLf
        sql &= " ,ISNULL(gs.CNT1F03,0) CNT1F03" & vbCrLf
        '開訓15~19歲人數、開訓20~24歲人數、開訓25~29歲人數、開訓30~34歲人數、開訓35~39歲人數、
        '開訓40~44歲人數、開訓45~49歲人數、開訓50~54歲人數、開訓55~59歲人數、開訓60~64歲人數、開訓65歲以上人數。
        sql &= " ,ISNULL(gs.YOD15,0) YOD15" & vbCrLf
        sql &= " ,ISNULL(gs.YOD20,0) YOD20" & vbCrLf
        sql &= " ,ISNULL(gs.YOD25,0) YOD25" & vbCrLf
        sql &= " ,ISNULL(gs.YOD30,0) YOD30" & vbCrLf
        sql &= " ,ISNULL(gs.YOD35,0) YOD35" & vbCrLf
        sql &= " ,ISNULL(gs.YOD40,0) YOD40" & vbCrLf
        sql &= " ,ISNULL(gs.YOD45,0) YOD45" & vbCrLf
        sql &= " ,ISNULL(gs.YOD50,0) YOD50" & vbCrLf
        sql &= " ,ISNULL(gs.YOD55,0) YOD55" & vbCrLf
        sql &= " ,ISNULL(gs.YOD60,0) YOD60" & vbCrLf
        sql &= " ,ISNULL(gs.YOD65,0) YOD65" & vbCrLf

        '報名人數
        sql &= " ,ISNULL(gs3.STUDETNUM,0) STUDETNUM" & vbCrLf
        '甄試人數
        sql &= " ,ISNULL(gs3.STUDETNUM2,0) STUDETNUM2" & vbCrLf
        '2、錄取人數：為【錄訓作業】功能中的「正取」人數。
        sql &= " ,ISNULL(gs3.STUDETNUM3,0) STUDETNUM3" & vbCrLf '錄取人數
        '1、到考率(%)=甄試人數/報名人數。 Attendance rate
        sql &= " ,CASE WHEN ISNULL(gs3.STUDETNUM,0)>0 THEN concat(round(cast(isnull(gs3.STUDETNUM2,0) as float)/gs3.STUDETNUM*100,2),'%') END ATTENRATE" & vbCrLf
        '3、錄取率(%)=錄取人數(正取)/甄試人數。 ACCEPtance rate
        sql &= " ,CASE WHEN ISNULL(gs3.STUDETNUM2,0)>0 THEN concat(round(cast(isnull(gs3.STUDETNUM3,0) as float)/gs3.STUDETNUM2*100,2),'%') END ACCEPRATE" & vbCrLf
        '4、開訓人數比率(%)=開訓人數/招生人數。 Number of trainees ratio
        sql &= " ,CASE WHEN cc.CNum>0 THEN concat(round(cast(isnull(gs.SNum,0) as float)/cc.CNum*100,2),'%') END TRAINRATE" & vbCrLf
        '5、離退訓率(%)=離退訓人數/開訓人數。 Retirement rate
        sql &= " ,CASE WHEN ISNULL(gs.SNum,0)>0 THEN concat(round(cast(isnull(gs.JSNum,0) as float)/gs.SNum*100,2),'%') END RTIRERATE" & vbCrLf

        '期別、到考率(%)、錄取人數、錄取率(%)、開訓人數比率(%)、離退訓率(%)、第1部分滿意度、第2部分滿意度、第3部分滿意度、第4部分滿意度、第5部分滿意度、平均滿意度
        '第1部分滿意度/'第2部分滿意度/'第3部分滿意度/'第4部分滿意度/'第5部分滿意度/'平均滿意度
        sql &= " ,case when q4.Q1_AVERAGE is not null then concat(q4.Q1_AVERAGE,'%') end Q1_AVERAGE" & vbCrLf
        sql &= " ,case when q4.Q2_AVERAGE is not null then concat(q4.Q2_AVERAGE,'%') end Q2_AVERAGE" & vbCrLf
        sql &= " ,case when q4.Q3_AVERAGE is not null then concat(q4.Q3_AVERAGE,'%') end Q3_AVERAGE" & vbCrLf
        sql &= " ,case when q4.Q4_AVERAGE is not null then concat(q4.Q4_AVERAGE,'%') end Q4_AVERAGE" & vbCrLf
        sql &= " ,case when q4.Q5_AVERAGE is not null then concat(q4.Q5_AVERAGE,'%') end Q5_AVERAGE" & vbCrLf
        'sql &= " ,q4.Q6_AVERAGE Q6_AVERAGE" & vbCrLf
        sql &= " ,case when q4.AVERAGE is not null then concat(q4.AVERAGE,'%') end AVERAGE" & vbCrLf

        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WC2 gs on gs.OCID =cc.OCID" & vbCrLf
        sql &= " LEFT JOIN WC3 gs3 on gs3.OCID =cc.OCID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_QUESTIONARY4 q4 ON q4.OCID=cc.OCID" & vbCrLf

        Return sql
    End Function

    '顯示所選轄區
    Sub area()
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr, DistrictName As String
        '選擇轄區
        Dim itemstr As String = ViewState("itemstr")
        Dim sqlstr As String = ""
        sqlstr &= " SELECT NAME,DISTID FROM ID_DISTRICT "
        If itemstr <> "" Then sqlstr &= " WHERE DistID IN (" & itemstr & ")"
        sqlstr &= " ORDER BY DISTID"

        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
        Dim DistrictName As String = ""
        For Each dr As DataRow In dt.Rows
            DistrictName &= String.Concat(If(DistrictName <> "", ",", ""), dr("Name"))
        Next

        DistrictLabel.Text = DistrictName
        District.Visible = True
        DistrictLabel.Visible = True
    End Sub

    'LIST
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        '審核狀態和連結實作 
        'Const cst_序號 As Integer=0
        Const cst_轄區分署 As Integer = 1
        Const cst_訓練計畫 As Integer = 2
        Const cst_訓練機構名稱 As Integer = 3
        Const cst_縣市別 As Integer = 4
        Const cst_班別名稱 As Integer = 5
        'Const cst_訓練職類 As Integer=6
        'Const cst_訓練性質 As Integer=7
        'Const cst_訓練時段 As Integer=8
        Const cst_開訓日期 As Integer = 9
        'Const cst_結訓日期 As Integer=10
        'Const cst_就業人數 As Integer=15
        Select Case e.Item.ItemType
            Case ListItemType.Header
                '排序功能
                If ViewState(cst_vssort) <> "" Then
                    Dim i As Integer = -1
                    Dim mysort As New System.Web.UI.WebControls.Image
                    mysort.ImageUrl = "../../images/SortDown.gif"
                    Select Case ViewState(cst_vssort)
                        Case "DistID", "DistID DESC"
                            i = cst_轄區分署
                            If ViewState(cst_vssort) = "DistID" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "PlanID", "PlanID DESC"
                            i = cst_訓練計畫
                            If ViewState(cst_vssort) = "PlanID" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "OrgName", "OrgName DESC"
                            'mylabel="OrgName"
                            i = cst_訓練機構名稱
                            If ViewState(cst_vssort) = "OrgName" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "CityName", "CityName DESC"
                            i = cst_縣市別
                            If ViewState(cst_vssort) = "CityName" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "ClassCName2", "ClassCName2 DESC"
                            i = cst_班別名稱
                            If ViewState(cst_vssort) = "ClassCName2" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "STDate", "STDate DESC"
                            i = cst_開訓日期
                            If ViewState(cst_vssort) = "STDate" Then mysort.ImageUrl = "../../images/SortUp.gif"
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                '序號
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim mybtn As LinkButton = e.Item.Cells(cst_班別名稱).Controls(0)
                mybtn.Attributes("onclick") = "window.open('CP_04_003_01.aspx?ID=" & Request("ID") & "&Student_Data=" & drv("OCID") & "','OCID','width=500,height=500'); return false;"
                'Dim vTitle As String=""
                'If Convert.ToString(drv("SUM_WINUM")) <> "" Then
                '    If drv("SUM_WINUM") > 0 Then
                '        vTitle="含提前就業人數:" & Convert.ToString(drv("SUM_WINUM"))
                '        TIMS.Tooltip(e.Item.Cells(cst_就業人數), vTitle)
                '    Else
                '        vTitle=""
                '        TIMS.Tooltip(e.Item.Cells(cst_就業人數), "")
                '    End If
                'End If
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        '匯入排序文字。
        ViewState(cst_vssort) = If(ViewState(cst_vssort) <> e.SortExpression, e.SortExpression, e.SortExpression & " DESC")
        PageControler1.Sort = ViewState(cst_vssort)
        PageControler1.ChangeSort()
    End Sub

    '報表列印 CP_04_003_Rpt.jrxml
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Const cst_ss_newICity As String = "newICity"
        Dim s_newICity As String = If(Session(cst_ss_newICity) IsNot Nothing, Convert.ToString(Session(cst_ss_newICity)), "")

        Dim myValue As String = ""
        myValue = "prg=CP_04_003_Rpt"
        myValue += "&AppliedDate=" & Request("yearlist")
        myValue += "&DistID=" & Convert.ToString(ViewState("newDistID"))
        myValue += "&TPlanID=" & ViewState("newTPlanID")
        myValue += "&SSTDate=" & ViewState("SSTDate")
        myValue += "&ESTDate=" & ViewState("ESTDate")
        myValue += "&itemcity=" & s_newICity '報表用。
        myValue += "&NotOpen=" & ViewState("NotOpenStaus")
        myValue += "&DistrictLabel=" & Server.UrlEncode(DistrictLabel.Text)
        'myValue += "&SumClass=" &  CountLabel.Text
        myValue += "&TTNum=" & CountLabel.Text
        myValue += "&STNum=" & STNum.Text
        myValue += "&SSNum=" & SSNum.Text
        myValue += "&SESNum=" & SESNum.Text
        'myValue += "&SGSNum=" &  SGSNum.Text
        myValue += "&ClassCName=" & Server.UrlEncode(ViewState("ClassCName"))
        '機構OrgID若有輸入，則不比對OrgName
        If ViewState("OrgID") <> "" Then
            myValue += "&OrgName="
            myValue += "&OrgID=" & ViewState("OrgID")
        Else
            myValue += "&OrgName=" & Server.UrlEncode(ViewState("OrgName"))
            myValue += "&OrgID="
        End If
        myValue += "&TrainType=" & ViewState("TMID")
        'myValue += "&BudgetID=" &  ViewState("newBudgetID")
        'myValue += "&BudgetID2=" &  ViewState("newBudgetID")
        myValue += "&SFTDate=" & ViewState("SFTDate")
        myValue += "&EFTDate=" & ViewState("EFTDate")
        myValue += "&newICityName=" & Server.UrlEncode(ViewState("newICityName"))
        myValue += "&newTPlanIDName=" & Server.UrlEncode(ViewState("newTPlanIDName"))
        myValue += "&NotOpenStausStr=" & Server.UrlEncode(ViewState("NotOpenStausStr"))
        myValue += "&TMIDName=" & Server.UrlEncode(ViewState("TMIDName"))
        'myValue += "&newBudgetName=" & Server.UrlEncode( ViewState("newBudgetName"))
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "CP_04_003_R2", myValue)
    End Sub

    ''' <summary> '回上一頁</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '回上一頁
        ' Button2.Attributes.Add("onclick", "location.href='CP_04_003.aspx?ID=" & Request("ID") & "';return false;")
        If PageControler1.SSSDTRID <> "" AndAlso Session(PageControler1.SSSDTRID) IsNot Nothing Then Session(PageControler1.SSSDTRID) = Nothing
        Dim url1 As String = "CP_04_003.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class
