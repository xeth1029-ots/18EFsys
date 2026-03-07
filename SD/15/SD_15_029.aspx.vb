Public Class SD_15_029
    Inherits AuthBasePage
    '主旨:'【確定性需求】訓練單位辦訓相關報表：4-2 核定課程統計表
    '從:'張瑩珊 <sammychang@turbotech.com.tw> '日期: '2023/3/15 下午 05:05 '到:
    '確定性需求    '在職系統：產投計畫    '預計完成日期：2023/04/10
    '4-2：核定課程統計表  需求內容。
    '功能路徑： 首頁>>學員動態管理>>統計表>>核定課程統計表    (先暫定，我還是覺得路徑怪怪)
    '使用者：署 (系統管理者、承辦人) '系統介面 '鎖定產投
    '年度區間：先用你講的方式，鎖定5個半年度(即2年半)
    '如110年度上半年~112年度上半年，110年度下半年~112年度下半年.....
    '(像現在是112年，年度選單可以顯示到112下，由他們自己去切換區間)
    '計畫：不區分、產業人才投資計畫、提升勞工自主學習計畫
    '示意圖如下： '邏輯補充：
    '(1) 依年度/申請階段篩選(不含政策性階段)
    '(2) 都是審核通過的班級，不管是否停辦
    '表頭名稱：110年度上半年~112年度上半年產投方案核定課程統計 (紅色部分可直接抓【年度區間】那段文字)
    '產出報表範例格式檔，請參附件。 
    '【(原業務單位提供)_核定課程統計表】：為現階段業務單位自行人工處理的版本
    '【調整後_核定課程統計表】：是我有調整版面，讓系統比較方便產出的格式，可先用這版做
    '匯出欄位說明如下：     '序號、計畫別、統編、訓練單位名稱、分署別
    '核定課程總計 五個年度/階段的：
    '年度、階段     '等級：以單位在 該年度/階段往前半年的【審查計分區間】資料檔為母體，撈取單位【複審審核】=通過的【複審等級】
    '舉例說明：110年度上半年XXX單位的等級資料，使用109年7月那份【審查計分區間】資料檔，撈取XXX單位【複審審核】=通過的【複審等級】
    '班數：該單位/年度/階段 審核通過 的班數
    '補助費：該單位【核定補助費】加總 (可對應綜合查詢統計表的【核定補助費】欄位)
    '管制類     '管制類指的是僅篩選【課程職類】為這三類的班級：美容【05-01】、餐飲【06】、手工藝品【07】
    '細分說明： '美容【05-01】：【03-02】、【03-03】、【03-04】、【03-05】
    '餐飲【06】：【21-01】、【21-02】、【21-03】、【30-01】、【30-02】
    '手工藝品【07】：【02-04】、【07-01】、【07-03】
    '五個年度/階段的：年度、階段、等級、班數、補助費，這部分說明同上
    '理事長：訓練機構設定的【負責人姓名】欄位
    '同組辦訓代表、同組辦訓單位家數、同組辦訓總核班數、核班增減(與前年度同期相比)、風險說明(增減狀況及原因說明)、立委關心案：這幾欄顯示欄位，內容空白即可
    '最下面還有小計、總計的，小計我覺得不用管他了(它是分產投、自主小計，卡一個在中間很怪)，
    '總計你看看可不可以算，不行也就算了，我叫他們自己excel拉
    '以上說明如有疑問，再請提出討論，感謝
    '附件: '1.110-112上產投方案核定課程統計.xlsx	101 KB

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)
        TIMS.Get_TitleLab(objconn, Request("ID"), TitleLab1, TitleLab2)

        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            bt_EXPORT.Enabled = False : TIMS.Tooltip(bt_EXPORT, TIMS.cst_ErrorMsg17, True)
            Common.MessageBox(Me, TIMS.cst_ErrorMsg17)
            Return 'Exit Sub
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not Me.IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        Dim js_BtnLevOrg1_1 As String = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then js_BtnLevOrg1_1 = "javascript:openOrg('../../Common/LevOrg.aspx');"
        BtnLevOrg1.Attributes("onclick") = js_BtnLevOrg1_1

        'ddlDistID = TIMS.Get_DistID(ddlDistID, TIMS.Get_DISTIDT2(objconn))
        'If sm.UserInfo.LID <> 0 Then
        '    Common.SetListItem(ddlDistID, sm.UserInfo.DistID)
        '    ddlDistID.Enabled = False
        'End If
        'Button1.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"
        'BtnClear2.Attributes("onclick") = "BtnClear2Click();"

        msg.Text = ""

        '申請階段 表示 (1：上半年、2：下半年)
        ddlAPPSTAG1 = TIMS.Get_ddlAPPSTAG(ddlAPPSTAG1)

        'Dim where_ff3 As String = "TPLANID IN ('06','28','54','70')"
        Dim where_ff3 As String = "TPLANID='28'"
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, where_ff3, objconn)
        Common.SetListItem(TPlanlist1, sm.UserInfo.TPlanID)

        TPlanlist1.Enabled = If(TPlanlist1.SelectedValue <> "", False, True)
        If (Not TPlanlist1.Enabled) Then TIMS.Tooltip(TPlanlist1, "限定計畫")

        'Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        'TPlanlist1.Enabled = False
        'If flagS1 AndAlso flag_test Then TPlanlist1.Enabled = True

        Dim iSYears As Integer = (sm.UserInfo.Years - 3) '2021
        Dim iEYearsNb1 As Integer = sm.UserInfo.Years
        Dim iEYears As Integer = If(iEYearsNb1 > iSYears, iEYearsNb1, iSYears)
        yearlist = TIMS.GetSyear(yearlist, iSYears, iEYears, False)
        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist, iEYearsNb1)
    End Sub

    Protected Sub bt_EXPORT_Click(sender As Object, e As EventArgs) Handles bt_EXPORT.Click
        Call EXPORT1_XLS()
    End Sub

    Private Sub EXPORT1_XLS()
        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        Call ExpReport1(dtXls)
    End Sub

    ''' <summary>匯出EXCEL</summary>
    ''' <param name="dtXls"></param>
    Private Sub ExpReport1(dtXls As DataTable)
        Dim YEAR_ROC1 As String = TIMS.ClearSQM(Hid_YEAR_ROC1.Value)
        Dim YEAR_ROC3 As String = TIMS.ClearSQM(Hid_YEAR_ROC3.Value)

        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        '表示 (1：上半年、2：下半年)
        Dim v_ddlAPPSTAG1 As String = TIMS.GetListValue(ddlAPPSTAG1)
        Dim v_APP_C As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAG1) 'If(v_ddlAPPSTAG1 = "1", "上半年", "下半年")
        'RBL_CLASSSTATUS /課程狀態/1-已申請/2-已核班通過(二階審查)/3-已核定(班級審核)
        'RBL_CLASSSTATUS /課程狀態/1-已申請/2-已二階審查/3-已核定(班級審核)
        Dim v_RBL_CLASSSTATUS As String = TIMS.GetListValue(RBL_CLASSSTATUS)

        '核定課程總計
        Dim sP22 As String = ",年度,階段,等級,班數,補助費,年度,階段,等級,班數,補助費,年度,階段,等級,班數,補助費,年度,階段,等級,班數,補助費,年度,階段,等級,班數,補助費"
        '管制類
        Dim sP23 As String = ",年度,階段,班數,補助費,年度,階段,班數,補助費,年度,階段,班數,補助費,年度,階段,班數,補助費,年度,階段,班數,補助費"
        Dim sP5 As String = ",理事長,同組辦訓代表,同組辦訓單位家數,同組辦訓總核班數,核班增減(與前年度同期相比),風險說明(增減狀況及原因說明),立委關心案"

        Dim sP22A() As String = Split(sP22, ",")
        Dim sP23A() As String = Split(sP23, ",")
        Dim sP5A() As String = Split(sP5, ",")

        'Dim sPattern As String = "" '序號,
        '表示 (1：上半年、2：下半年)'核定課程總計
        Dim sC1 As String = ",YEARS_ROC11A,APPSTAGE_N11A,RLEVEL2_11A,CLSN11A,FIXSUMCOST11A,YEARS_ROC12A,APPSTAGE_N12A,RLEVEL2_12A,CLSN12A,FIXSUMCOST12A,YEARS_ROC21A,APPSTAGE_N21A,RLEVEL2_21A,CLSN21A,FIXSUMCOST21A,YEARS_ROC22A,APPSTAGE_N22A,RLEVEL2_22A,CLSN22A,FIXSUMCOST22A,YEARS_ROC31A,APPSTAGE_N31A,RLEVEL2_31A,CLSN31A,FIXSUMCOST31A"
        Dim sC12 As String = ",YEARS_ROC12A,APPSTAGE_N12A,RLEVEL2_12A,CLSN12A,FIXSUMCOST12A,YEARS_ROC21A,APPSTAGE_N21A,RLEVEL2_21A,CLSN21A,FIXSUMCOST21A,YEARS_ROC22A,APPSTAGE_N22A,RLEVEL2_22A,CLSN22A,FIXSUMCOST22A,YEARS_ROC31A,APPSTAGE_N31A,RLEVEL2_31A,CLSN31A,FIXSUMCOST31A,YEARS_ROC32A,APPSTAGE_N32A,RLEVEL2_32A,CLSN32A,FIXSUMCOST32A"
        '表示 (1：上半年、2：下半年)'管制類
        Dim sC2 As String = ",YEARS_ROC11,APPSTAGE_N11,CLSN11,FIXSUMCOST11,YEARS_ROC12,APPSTAGE_N12,CLSN12,FIXSUMCOST12,YEARS_ROC21,APPSTAGE_N21,CLSN21,FIXSUMCOST21,YEARS_ROC22,APPSTAGE_N22,CLSN22,FIXSUMCOST22,YEARS_ROC31,APPSTAGE_N31,CLSN31,FIXSUMCOST31"
        Dim sC22 As String = ",YEARS_ROC12,APPSTAGE_N12,CLSN12,FIXSUMCOST12,YEARS_ROC21,APPSTAGE_N21,CLSN21,FIXSUMCOST21,YEARS_ROC22,APPSTAGE_N22,CLSN22,FIXSUMCOST22,YEARS_ROC31,APPSTAGE_N31,CLSN31,FIXSUMCOST31,YEARS_ROC32,APPSTAGE_N32,CLSN32,FIXSUMCOST32"

        'Dim sColumn As String = ""
        Dim sPattern As String = String.Concat("序號", ",計畫別,統編,訓練單位名稱,分署別", sP22, sP23, sP5)
        Dim sColumn As String = String.Concat("SEQNO,ORGPLANNAME2,COMIDNO,ORGNAME,DISTNAME3", If(v_ddlAPPSTAG1 = "1", sC1, sC12), If(v_ddlAPPSTAG1 = "1", sC2, sC22), ",MASTERNAME,NU1,NU2,NU3,NU4,NU5,NU6")
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        Dim sFileName1 As String = String.Concat(YEAR_ROC1, v_APP_C, "-", YEAR_ROC3, v_APP_C, "產投方案核定課程統計", v_OrgKind2, v_RBL_CLASSSTATUS)
        'Dim sFileName1 As String = "110-112年度上半年產投方案核定課程統計"
        '套CSS值
        Dim strSTYLE As String = String.Concat("<style>", "td{mso-number-format:""\@"";}", ".noDecFormat{mso-number-format:""0"";}", "</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '建立輸出文字
        Dim ExportStr As String = ""
        '標題抬頭1
        ExportStr = String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFileName1) '& vbTab
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭2
        ExportStr = ""
        For i As Integer = 0 To 4
            ExportStr &= String.Format("<td rowspan=2>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= String.Format("<td colspan={0}>{1}</td>", Split(sP22, ",").Length - 1, "核定課程總計") '& vbTab
        ExportStr &= String.Format("<td colspan={0}>{1}</td>", Split(sP23, ",").Length - 1, "管制類") '& vbTab
        For i As Integer = 1 To sP5A.Length - 1
            ExportStr &= String.Format("<td rowspan=2>{0}</td>", sP5A(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭3
        ExportStr = ""
        For i As Integer = 1 To sP22A.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sP22A(i)) '& vbTab
        Next
        For i As Integer = 1 To sP23A.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sP23A(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '建立資料面
        Dim iRows As Integer = 0
        For Each dr As DataRow In dtXls.Rows
            iRows += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                Dim sCOLTXT As String = ""
                Select Case sColumnA(i)
                    Case "SEQNO"
                        sCOLTXT = iRows.ToString()
                    Case Else
                        sCOLTXT = Convert.ToString(dr(sColumnA(i)))
                End Select
                ExportStr &= String.Format("<td>{0}</td>", sCOLTXT)
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next

        'If iNum1 > 0 Then
        '    ExportStr = "<tr>"
        '    ExportStr &= "<td colspan=2>總計</td>"
        '    ExportStr &= String.Format("<td>{0}</td>", iOTN) 'OTN
        '    ExportStr &= String.Format("<td>{0}</td>", iCLSN1) 'CLSN1
        '    ExportStr &= String.Format("<td>{0}</td>", iCLSN2) 'CLSN2
        '    ExportStr &= String.Format("<td>{0}%</td>", iAP1) 'AP1
        '    ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP2 / iNum1, 2)) 'AP2
        '    ExportStr &= String.Format("<td>{0}</td>", iCLSN3) 'CLSN3
        '    ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP3 / iNum1, 2)) 'AP3
        '    ExportStr &= String.Format("<td>{0}</td>", iSTDN1) 'STDN1
        '    ExportStr &= String.Format("<td>{0}</td>", iSTDN2) 'STDN2
        '    ExportStr &= String.Format("<td>{0}</td>", iSTDN3) 'STDN3
        '    ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP4 / iNum1, 2)) 'AP4
        '    ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP5 / iNum1, 2)) 'AP5
        '    ExportStr &= String.Format("<td>{0}</td>", TIMS.VAL2N0(iSUB1)) 'SUB1
        '    ExportStr &= String.Format("<td>{0}</td>", TIMS.VAL2N0(iSUB2)) 'SUB2
        '    ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP6 / iNum1, 2)) 'AP6
        '    ExportStr &= "</tr>" & vbCrLf
        '    sbHTML.Append(ExportStr)
        'End If

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        'parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        Dim parmsExp As New Hashtable From {
            {"ExpType", "EXCEL"}, 'EXCEL/PDF/ODS
            {"FileName", sFileName1},
            {"strSTYLE", strSTYLE},
            {"strHTML", sbHTML.ToString()},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

#Region "NO USE"
    ''' <summary>查詢資料Old1</summary>
    ''' <returns></returns>
    Private Function SEARCH_DATA1_Old1_dt() As DataTable
        Hid_YEAR_ROC1.Value = TIMS.GetListText(yearlist) '民國年
        Hid_YEAR_ROC3.Value = (TIMS.VAL1(Hid_YEAR_ROC1.Value) + 2).ToString() '民國年
        Dim v_Yearlist As String = TIMS.GetListValue(yearlist) '西元年
        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)

        Dim D_YEARS1 As String = v_Yearlist
        Dim D_YEARS2 As String = CStr(TIMS.VAL1(v_Yearlist) + 1)
        Dim D_YEARS3 As String = CStr(TIMS.VAL1(v_Yearlist) + 2)
        Dim v_ddlAPPSTAG1 As String = TIMS.GetListValue(ddlAPPSTAG1) '申請階段

        'sParms.Add("YEARS", v_Yearlist)
        Dim sParms As New Hashtable From {
            {"TPlanID", sm.UserInfo.TPlanID},
            {"ORGKIND2", v_OrgKind2}
        }
        Dim sSql As String = "" 'sSql = "" & vbCrLf
        sSql &= String.Concat(" DECLARE @D_YEARS1 VARCHAR(4)='", D_YEARS1, "';") & vbCrLf
        sSql &= String.Concat(" DECLARE @D_YEARS2 VARCHAR(4)='", D_YEARS2, "';") & vbCrLf
        sSql &= String.Concat(" DECLARE @D_YEARS3 VARCHAR(4)='", D_YEARS3, "';") & vbCrLf

        '全部 WC1A 
        sSql &= " WITH WC1A AS (SELECT cc.OCID,cc.PLANID,cc.COMIDNO,cc.SEQNO ,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.DEFGOVCOST" & vbCrLf
        sSql &= "  ,cc.RID,cc.ORGKIND2,cc.GCID3,cc.CLASSCNAME2,cc.ISSUCCESS,cc.ISAPPRPAPER,cc.NOTOPEN,cc.PVR_ISAPPRPAPER" & vbCrLf
        sSql &= "  FROM dbo.VIEW2B cc WHERE cc.ISSUCCESS='Y' AND cc.ISAPPRPAPER='Y' AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf
        sSql &= "  AND cc.YEARS IN (@D_YEARS1,@D_YEARS2,@D_YEARS3) AND cc.TPLANID=@TPlanID AND cc.APPSTAGE IN (1,2) AND cc.ORGKIND2=@ORGKIND2)" & vbCrLf

        sSql &= " ,WC1 AS (SELECT cc.OCID,cc.COMIDNO,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.DEFGOVCOST,cc.RID,cc.ORGKIND2,cc.GCID3" & vbCrLf
        sSql &= "  FROM WC1A cc WHERE cc.GCID3 IN (SELECT GCID3 FROM dbo.V_GOVCLASSCAST3 WHERE GCODE31 IN ('05','06','07') AND PGCID3 IS NOT NULL AND GCID3!=2072))" & vbCrLf
        'sSql &= " ,WC2A AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.FIXSUMCOST) FIXSUMCOST FROM WC1A cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        'sSql &= " ,WC2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.FIXSUMCOST) FIXSUMCOST FROM WC1 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        sSql &= " ,WC2A AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.DEFGOVCOST) FIXSUMCOST FROM WC1A cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        sSql &= " ,WC2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.DEFGOVCOST) FIXSUMCOST FROM WC1 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf

        If v_ddlAPPSTAG1 = "1" Then
            sSql &= " ,WO11A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO12A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO31A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf

            sSql &= " ,WO11 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO12 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO31 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf

            sSql &= " SELECT o11A.YEARS,o11A.COMIDNO,o11A.DISTID" & vbCrLf
            sSql &= " ,rr.ORGPLANNAME2,rr.ORGNAME" & vbCrLf
            sSql &= " ,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=o11A.DISTID) DISTNAME3" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o11A.YEARS) YEARS_ROC11A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o11A.APPSTAGE) APPSTAGE_N11A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o11A.COMIDNO,o11A.TPLANID,o11A.DISTID,o11A.YEARS,o11A.APPSTAGE) RLEVEL2_11A" & vbCrLf
            sSql &= " ,o11A.CLSN1 CLSN11A" & vbCrLf
            sSql &= " ,o11A.FIXSUMCOST FIXSUMCOST11A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o12A.YEARS) YEARS_ROC12A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12A.APPSTAGE) APPSTAGE_N12A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12A.COMIDNO,o12A.TPLANID,o12A.DISTID,o12A.YEARS,o12A.APPSTAGE) RLEVEL2_12A" & vbCrLf
            sSql &= " ,o12A.CLSN1 CLSN12A" & vbCrLf
            sSql &= " ,o12A.FIXSUMCOST FIXSUMCOST12A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21A.YEARS) YEARS_ROC21A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21A.APPSTAGE) APPSTAGE_N21A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21A.COMIDNO,o21A.TPLANID,o21A.DISTID,o21A.YEARS,o21A.APPSTAGE) RLEVEL2_21A" & vbCrLf
            sSql &= " ,o21A.CLSN1 CLSN21A" & vbCrLf
            sSql &= " ,o21A.FIXSUMCOST FIXSUMCOST21A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22A.YEARS) YEARS_ROC22A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22A.APPSTAGE) APPSTAGE_N22A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22A.COMIDNO,o22A.TPLANID,o22A.DISTID,o22A.YEARS,o22A.APPSTAGE) RLEVEL2_22A" & vbCrLf
            sSql &= " ,o22A.CLSN1 CLSN22A" & vbCrLf
            sSql &= " ,o22A.FIXSUMCOST FIXSUMCOST22A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31A.YEARS) YEARS_ROC31A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31A.APPSTAGE) APPSTAGE_N31A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31A.COMIDNO,o31A.TPLANID,o31A.DISTID,o31A.YEARS,o31A.APPSTAGE) RLEVEL2_31A" & vbCrLf
            sSql &= " ,o31A.CLSN1 CLSN31A" & vbCrLf
            sSql &= " ,o31A.FIXSUMCOST FIXSUMCOST31A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o11.YEARS) YEARS_ROC11" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o11.APPSTAGE) APPSTAGE_N11" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o11.COMIDNO,o11.TPLANID,o11.DISTID,o11.YEARS,o11.APPSTAGE) RLEVEL2_11" & vbCrLf
            sSql &= " ,o11.CLSN1 CLSN11" & vbCrLf
            sSql &= " ,o11.FIXSUMCOST FIXSUMCOST11" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o12.YEARS) YEARS_ROC12" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12.APPSTAGE) APPSTAGE_N12" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12.COMIDNO,o12.TPLANID,o12.DISTID,o12.YEARS,o12.APPSTAGE) RLEVEL2_12" & vbCrLf
            sSql &= " ,o12.CLSN1 CLSN12" & vbCrLf
            sSql &= " ,o12.FIXSUMCOST FIXSUMCOST12" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21.YEARS) YEARS_ROC21" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21.APPSTAGE) APPSTAGE_N21" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21.COMIDNO,o21.TPLANID,o21.DISTID,o21.YEARS,o21.APPSTAGE) RLEVEL2_21" & vbCrLf
            sSql &= " ,o21.CLSN1 CLSN21" & vbCrLf
            sSql &= " ,o21.FIXSUMCOST FIXSUMCOST21" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22.YEARS) YEARS_ROC22" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22.APPSTAGE) APPSTAGE_N22" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22.COMIDNO,o22.TPLANID,o22.DISTID,o22.YEARS,o22.APPSTAGE) RLEVEL2_22" & vbCrLf
            sSql &= " ,o22.CLSN1 CLSN22" & vbCrLf
            sSql &= " ,o22.FIXSUMCOST FIXSUMCOST22" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31.YEARS) YEARS_ROC31" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31.APPSTAGE) APPSTAGE_N31" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31.COMIDNO,o31.TPLANID,o31.DISTID,o31.YEARS,o31.APPSTAGE) RLEVEL2_31" & vbCrLf
            sSql &= " ,o31.CLSN1 CLSN31" & vbCrLf
            sSql &= " ,o31.FIXSUMCOST FIXSUMCOST31" & vbCrLf

            sSql &= " ,rr.MASTERNAME,'' NU1,'' NU2,'' NU3,'' NU4,'' NU5,'' NU6" & vbCrLf
            sSql &= " FROM WO11A o11A" & vbCrLf
            sSql &= " LEFT JOIN WO12A o12A ON o12A.COMIDNO=o11A.COMIDNO AND o12A.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21A o21A ON o21A.COMIDNO=o11A.COMIDNO AND o21A.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22A o22A ON o22A.COMIDNO=o11A.COMIDNO AND o22A.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31A o31A ON o31A.COMIDNO=o11A.COMIDNO AND o31A.DISTID =o11A.DISTID" & vbCrLf

            sSql &= " LEFT JOIN WO11 o11 ON o11.COMIDNO=o11A.COMIDNO AND o11.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO12 o12 ON o12.COMIDNO=o11A.COMIDNO AND o12.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21 o21 ON o21.COMIDNO=o11A.COMIDNO AND o21.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22 o22 ON o22.COMIDNO=o11A.COMIDNO AND o22.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31 o31 ON o31.COMIDNO=o11A.COMIDNO AND o31.DISTID =o11A.DISTID" & vbCrLf
            sSql &= " JOIN dbo.VIEW_RIDNAME rr ON rr.RID=o11A.RID" & vbCrLf
            sSql &= " ORDER BY o11A.YEARS,rr.ORGPLANNAME2 desc,o11A.DISTID,o11A.COMIDNO" & vbCrLf

        ElseIf v_ddlAPPSTAG1 = "2" Then
            'sSql &= " ,WO11A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO12A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO31A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO32A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=2)" & vbCrLf

            'sSql &= " ,WO11 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO12 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO31 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO32 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=2)" & vbCrLf

            sSql &= " SELECT o12A.YEARS,o12A.COMIDNO,o12A.DISTID" & vbCrLf
            sSql &= " ,rr.ORGPLANNAME2,rr.ORGNAME" & vbCrLf
            sSql &= " ,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=o12A.DISTID) DISTNAME3" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o12A.YEARS) YEARS_ROC12A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12A.APPSTAGE) APPSTAGE_N12A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12A.COMIDNO,o12A.TPLANID,o12A.DISTID,o12A.YEARS,o12A.APPSTAGE) RLEVEL2_12A" & vbCrLf
            sSql &= " ,o12A.CLSN1 CLSN12A" & vbCrLf
            sSql &= " ,o12A.FIXSUMCOST FIXSUMCOST12A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21A.YEARS) YEARS_ROC21A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21A.APPSTAGE) APPSTAGE_N21A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21A.COMIDNO,o21A.TPLANID,o21A.DISTID,o21A.YEARS,o21A.APPSTAGE) RLEVEL2_21A" & vbCrLf
            sSql &= " ,o21A.CLSN1 CLSN21A" & vbCrLf
            sSql &= " ,o21A.FIXSUMCOST FIXSUMCOST21A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22A.YEARS) YEARS_ROC22A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22A.APPSTAGE) APPSTAGE_N22A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22A.COMIDNO,o22A.TPLANID,o22A.DISTID,o22A.YEARS,o22A.APPSTAGE) RLEVEL2_22A" & vbCrLf
            sSql &= " ,o22A.CLSN1 CLSN22A" & vbCrLf
            sSql &= " ,o22A.FIXSUMCOST FIXSUMCOST22A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31A.YEARS) YEARS_ROC31A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31A.APPSTAGE) APPSTAGE_N31A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31A.COMIDNO,o31A.TPLANID,o31A.DISTID,o31A.YEARS,o31A.APPSTAGE) RLEVEL2_31A" & vbCrLf
            sSql &= " ,o31A.CLSN1 CLSN31A" & vbCrLf
            sSql &= " ,o31A.FIXSUMCOST FIXSUMCOST31A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o32A.YEARS) YEARS_ROC32A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o32A.APPSTAGE) APPSTAGE_N32A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o32A.COMIDNO,o32A.TPLANID,o32A.DISTID,o32A.YEARS,o32A.APPSTAGE) RLEVEL2_32A" & vbCrLf
            sSql &= " ,o32A.CLSN1 CLSN32A" & vbCrLf
            sSql &= " ,o32A.FIXSUMCOST FIXSUMCOST32A" & vbCrLf
            '---
            sSql &= " ,dbo.FN_CYEAR2(o12.YEARS) YEARS_ROC12" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12.APPSTAGE) APPSTAGE_N12" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12.COMIDNO,o12.TPLANID,o12.DISTID,o12.YEARS,o12.APPSTAGE) RLEVEL2_12" & vbCrLf
            sSql &= " ,o12.CLSN1 CLSN12" & vbCrLf
            sSql &= " ,o12.FIXSUMCOST FIXSUMCOST12" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21.YEARS) YEARS_ROC21" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21.APPSTAGE) APPSTAGE_N21" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21.COMIDNO,o21.TPLANID,o21.DISTID,o21.YEARS,o21.APPSTAGE) RLEVEL2_21" & vbCrLf
            sSql &= " ,o21.CLSN1 CLSN21" & vbCrLf
            sSql &= " ,o21.FIXSUMCOST FIXSUMCOST21" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22.YEARS) YEARS_ROC22" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22.APPSTAGE) APPSTAGE_N22" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22.COMIDNO,o22.TPLANID,o22.DISTID,o22.YEARS,o22.APPSTAGE) RLEVEL2_22" & vbCrLf
            sSql &= " ,o22.CLSN1 CLSN22" & vbCrLf
            sSql &= " ,o22.FIXSUMCOST FIXSUMCOST22" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31.YEARS) YEARS_ROC31" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31.APPSTAGE) APPSTAGE_N31" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31.COMIDNO,o31.TPLANID,o31.DISTID,o31.YEARS,o31.APPSTAGE) RLEVEL2_31" & vbCrLf
            sSql &= " ,o31.CLSN1 CLSN31" & vbCrLf
            sSql &= " ,o31.FIXSUMCOST FIXSUMCOST31" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o32.YEARS) YEARS_ROC32" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o32.APPSTAGE) APPSTAGE_N32" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31.COMIDNO,o31.TPLANID,o31.DISTID,o31.YEARS,o31.APPSTAGE) RLEVEL2_31" & vbCrLf
            sSql &= " ,o32.CLSN1 CLSN32" & vbCrLf
            sSql &= " ,o32.FIXSUMCOST FIXSUMCOST32" & vbCrLf

            sSql &= " ,rr.MASTERNAME,'' NU1,'' NU2,'' NU3,'' NU4,'' NU5,'' NU6" & vbCrLf
            sSql &= " FROM WO12A o12A" & vbCrLf
            sSql &= " LEFT JOIN WO21A o21A ON o21A.COMIDNO=o12A.COMIDNO AND o21A.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22A o22A ON o22A.COMIDNO=o12A.COMIDNO AND o22A.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31A o31A ON o31A.COMIDNO=o12A.COMIDNO AND o31A.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO32A o32A ON o32A.COMIDNO=o12A.COMIDNO AND o31A.DISTID=o12A.DISTID" & vbCrLf

            sSql &= " LEFT JOIN WO12 o12 ON o12.COMIDNO=o12A.COMIDNO AND o12.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21 o21 ON o21.COMIDNO=o12A.COMIDNO AND o21.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22 o22 ON o22.COMIDNO=o12A.COMIDNO AND o22.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31 o31 ON o31.COMIDNO=o12A.COMIDNO AND o31.DISTID=o12A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO32 o32 ON o32.COMIDNO=o12A.COMIDNO AND o32.DISTID=o12A.DISTID" & vbCrLf

            sSql &= " JOIN dbo.VIEW_RIDNAME rr ON rr.RID=o12A.RID" & vbCrLf
            sSql &= " ORDER BY o12A.YEARS,rr.ORGPLANNAME2 desc,o12A.DISTID,o12A.COMIDNO" & vbCrLf
        End If

        'Dim sSql As String = ""
        'sSql = "" & vbCrLf
        'sSql &= " WITH WC1 AS (SELECT cc.OCID,cc.COMIDNO,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.RID" & vbCrLf
        'sSql &= " FROM dbo.VIEW2 cc" & vbCrLf
        'sSql &= " WHERE 1=1 AND cc.YEARS IN ('2021','2022','2023') AND cc.TPLANID='28' AND cc.APPSTAGE IN (1,2)" & vbCrLf
        'sSql &= " AND cc.GCID3 IN (SELECT GCID3 FROM dbo.V_GOVCLASSCAST3 WHERE GCODE31 IN ('05','06','07') AND PGCID3 IS NOT NULL AND GCID3!=2072))" & vbCrLf
        'sSql &= " ,WC2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.FIXSUMCOST) FIXSUMCOST FROM WC1 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        'sSql &= " SELECT cc.YEARS,rr.ORGPLANNAME2,cc.COMIDNO" & vbCrLf
        'sSql &= " ,rr.ORGNAME" & vbCrLf
        'sSql &= " ,cc.DISTID" & vbCrLf
        'sSql &= " ,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=cc.DISTID) DISTNAME3" & vbCrLf
        'sSql &= " ,dbo.FN_CYEAR2(cc.YEARS) YEARS_ROC" & vbCrLf
        'sSql &= " ,CASE cc.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        'sSql &= " ,(SELECT MIN(x.RLEVEL_2) FROM ORG_SCORING2 x WHERE x.COMIDNO=cc.COMIDNO AND x.TPLANID=cc.TPLANID AND x.DISTID=cc.DISTID AND x.SECONDCHK='Y'" & vbCrLf
        'sSql &= " AND CASE x.MONTHS WHEN '01' THEN x.YEARS WHEN '07' THEN convert(varchar(4),convert(int,x.YEARS)+1) END=cc.YEARS" & vbCrLf
        'sSql &= " AND CASE MONTHS WHEN '01' THEN 2 WHEN '07' THEN 1 END=cc.APPSTAGE) RLEVEL_2" & vbCrLf
        'sSql &= " ,cc.CLSN1" & vbCrLf
        'sSql &= " ,cc.FIXSUMCOST" & vbCrLf
        'sSql &= " ,rr.MASTERNAME" & vbCrLf
        'sSql &= " FROM WC2 cc" & vbCrLf
        'sSql &= " JOIN dbo.VIEW_RIDNAME rr ON rr.RID=cc.RID" & vbCrLf
        'sSql &= " ORDER BY cc.YEARS,rr.ORGPLANNAME2 desc,cc.DISTID,cc.COMIDNO" & vbCrLf

        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sSql, objconn)
        Call DbAccess.HashParmsChange(sCmd, sParms)
        dt.Load(sCmd.ExecuteReader())
        Call TIMS.CHG_dtReadOnly(dt)
        Return dt
    End Function

#End Region

    ''' <summary>查詢資料</summary>
    ''' <returns></returns>
    Private Function SEARCH_DATA1_dt() As DataTable
        Hid_YEAR_ROC3.Value = TIMS.GetListText(yearlist) '民國年
        Hid_YEAR_ROC1.Value = (TIMS.VAL1(Hid_YEAR_ROC3.Value) - 2).ToString() '民國年-2
        Dim v_Yearlist As String = TIMS.GetListValue(yearlist) '西元年
        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)

        Dim D_YEARS1 As String = CStr(TIMS.VAL1(v_Yearlist) - 2) '指定年度-2(前年)
        Dim D_YEARS2 As String = CStr(TIMS.VAL1(v_Yearlist) - 1) '指定年度-1(去年)
        Dim D_YEARS3 As String = v_Yearlist '指定年度
        Dim v_ddlAPPSTAG1 As String = TIMS.GetListValue(ddlAPPSTAG1) '申請階段
        'RBL_CLASSSTATUS /課程狀態/1-已申請/2-已核班通過(二階審查)/3-已核定(班級審核)
        'RBL_CLASSSTATUS /課程狀態/1-已申請/2-已二階審查/3-已核定(班級審核)
        Dim v_RBL_CLASSSTATUS As String = TIMS.GetListValue(RBL_CLASSSTATUS)

        'Hid_DISTID.Value = sm.UserInfo.DistID
        'Hid_DISTID.Value = TIMS.ClearSQM(Hid_DISTID.Value)
        'Dim v_ddlDistID As String = TIMS.GetListValue(ddlDistID)
        Dim vDISTID As String = ""
        Dim vCOMIDNO As String = ""
        Dim drRR As DataRow = Nothing
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value <> "" Then drRR = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        Dim fg_USE_COMIDNO As Boolean = (drRR IsNot Nothing AndAlso RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1)
        Select Case sm.UserInfo.LID
            Case 0
                Dim fg_USE_DISTID As Boolean = (drRR IsNot Nothing AndAlso RIDValue.Value <> "" AndAlso RIDValue.Value.Length = 1)
                If fg_USE_DISTID AndAlso Convert.ToString(drRR("DISTID")) <> "000" Then vDISTID = Convert.ToString(drRR("DISTID"))
                If fg_USE_COMIDNO Then vCOMIDNO = Convert.ToString(drRR("COMIDNO"))
            Case Else
                vDISTID = sm.UserInfo.DistID
                If fg_USE_COMIDNO Then vCOMIDNO = Convert.ToString(drRR("COMIDNO"))
        End Select

        Dim sParms As New Hashtable From {
            {"TPlanID", sm.UserInfo.TPlanID},
            {"ORGKIND2", v_OrgKind2}
        }
        Dim sSql As String = ""
        sSql &= String.Concat("DECLARE @D_YEARS1 VARCHAR(4)='", D_YEARS1, "';") & vbCrLf
        sSql &= String.Concat("DECLARE @D_YEARS2 VARCHAR(4)='", D_YEARS2, "';") & vbCrLf
        sSql &= String.Concat("DECLARE @D_YEARS3 VARCHAR(4)='", D_YEARS3, "';") & vbCrLf

        '全部課程 WC1A
        sSql &= " WITH WC1A AS (SELECT cc.OCID,cc.PLANID,cc.COMIDNO,cc.SEQNO ,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.DEFGOVCOST" & vbCrLf
        sSql &= "  ,cc.RID,cc.ORGKIND2,cc.GCID3,cc.CLASSCNAME2,cc.ISSUCCESS,cc.ISAPPRPAPER,cc.NOTOPEN,cc.RESULTBUTTON,cc.PVR_ISAPPRPAPER,cc.DATANOTSENT" & vbCrLf
        sSql &= "  ,pf.RESULT,pf.CURESULT" & vbCrLf '一階審查結果/核班結果 二階審查-核班結果功能中之通過及不通過班級
        sSql &= "  FROM dbo.VIEW2B cc" & vbCrLf
        sSql &= "  LEFT JOIN dbo.PLAN_STAFFOPIN pf WITH(NOLOCK) ON pf.PSNO28=cc.PSNO28" & vbCrLf
        sSql &= "  WHERE cc.ISAPPRPAPER='Y' AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf 'cc.ISSUCCESS='Y'轉入
        sSql &= "  AND cc.YEARS IN (@D_YEARS1,@D_YEARS2,@D_YEARS3)" & vbCrLf
        If vDISTID <> "" Then
            sParms.Add("DISTID", vDISTID)
            sSql &= " AND cc.DISTID=@DISTID" & vbCrLf
        End If
        If vCOMIDNO <> "" Then
            sParms.Add("COMIDNO", vCOMIDNO)
            sSql &= " AND cc.COMIDNO=@COMIDNO" & vbCrLf
        End If
        sSql &= "  AND cc.TPLANID=@TPlanID AND cc.ORGKIND2=@ORGKIND2 AND cc.APPSTAGE IN (1,2))" & vbCrLf

        '全部課程 GROUP WC2A (其它年度)
        sSql &= " ,WC2A AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.DEFGOVCOST) FIXSUMCOST" & vbCrLf
        sSql &= "  FROM WC1A cc WHERE cc.ISSUCCESS='Y'" & vbCrLf
        sSql &= "  GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf

        '全部課程 WC1A->WC1A2 (課程狀態過濾)
        sSql &= " ,WC1A2 AS ( SELECT cc.OCID,cc.COMIDNO,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.DEFGOVCOST,cc.RID,cc.ORGKIND2,cc.GCID3,cc.RESULT,cc.CURESULT" & vbCrLf
        sSql &= "  FROM WC1A cc WHERE cc.YEARS=@D_YEARS3" & vbCrLf
        'RBL_CLASSSTATUS /課程狀態/1-已申請/2-已核班通過(二階審查)/3-已核定(班級審核)
        'RBL_CLASSSTATUS /課程狀態/1-已申請/2-已二階審查/3-已核定(班級審核)
        sSql &= If(v_RBL_CLASSSTATUS = "2", " AND cc.CURESULT IN ('Y','N')", If(v_RBL_CLASSSTATUS = "3", " AND cc.ISSUCCESS='Y'", ""))
        sSql &= " )" & vbCrLf

        '全部課程 WC1A->WC1A2 (課程狀態過濾) GROUP WC1A2->WC2A2
        If v_RBL_CLASSSTATUS = "2" Then 'v_RBL_CLASSSTATUS : 2-已二階審查
            sSql &= " ,WC2A2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID" & vbCrLf
            sSql &= "  ,COUNT(CASE cc.CURESULT WHEN 'Y' THEN 1 END) CLSN1,SUM(CASE cc.CURESULT WHEN 'Y' THEN cc.DEFGOVCOST END) FIXSUMCOST" & vbCrLf
            sSql &= "  FROM WC1A2 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        Else
            sSql &= " ,WC2A2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.DEFGOVCOST) FIXSUMCOST FROM WC1A2 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        End If

        '管制類課程 WC1A->WC1 (其它年度)
        sSql &= " ,WC1 AS (SELECT cc.OCID,cc.COMIDNO,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.DEFGOVCOST,cc.RID,cc.ORGKIND2,cc.GCID3,cc.CURESULT" & vbCrLf
        sSql &= "  FROM WC1A cc" & vbCrLf
        'sSql &= "  WHERE cc.RESULTBUTTON IS NULL" & vbCrLf '審核送出(已送審)
        'sSql &= "  AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        'sSql &= "  AND cc.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sSql &= "  WHERE cc.ISSUCCESS='Y' AND cc.GCID3 IN (SELECT GCID3 FROM dbo.V_GOVCLASSCAST3 WHERE GCODE31 IN ('05','06','07') AND PGCID3 IS NOT NULL AND GCID3!=2072))" & vbCrLf

        '管制類課程 GROUP WC2 (其它年度) WC1->WC2
        If v_RBL_CLASSSTATUS = "2" Then 'v_RBL_CLASSSTATUS : 2-已二階審查
            sSql &= " ,WC2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID" & vbCrLf
            sSql &= "  ,COUNT(CASE cc.CURESULT WHEN 'Y' THEN 1 END) CLSN1,SUM(CASE cc.CURESULT WHEN 'Y' THEN cc.DEFGOVCOST END) FIXSUMCOST" & vbCrLf
            sSql &= "  FROM WC1 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        Else
            sSql &= " ,WC2 AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.DEFGOVCOST) FIXSUMCOST FROM WC1 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf
        End If

        '管制類課程 WC3 (課程狀態過濾) WC1A2->WC3 
        sSql &= " ,WC3 AS (SELECT cc.OCID,cc.COMIDNO,cc.DISTID,cc.YEARS,cc.TPLANID,cc.APPSTAGE,cc.FIXSUMCOST,cc.DEFGOVCOST,cc.RID,cc.ORGKIND2,cc.GCID3" & vbCrLf
        sSql &= "  FROM WC1A2 cc WHERE cc.GCID3 IN (SELECT GCID3 FROM dbo.V_GOVCLASSCAST3 WHERE GCODE31 IN ('05','06','07') AND PGCID3 IS NOT NULL AND GCID3!=2072))" & vbCrLf

        '管制類課程 GROUP WC3G (課程狀態過濾) WC3->WC3G 
        sSql &= " ,WC3G AS (SELECT cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID,COUNT(1) CLSN1,SUM(cc.DEFGOVCOST) FIXSUMCOST FROM WC3 cc GROUP BY cc.TPLANID,cc.COMIDNO,cc.YEARS,cc.APPSTAGE,cc.RID,cc.DISTID)" & vbCrLf

        '申請階段
        If v_ddlAPPSTAG1 = "1" Then
            sSql &= " ,WO31A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A2 cc WHERE cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO11A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO12A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf

            sSql &= " ,WO31 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC3G cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO11 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO12 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf

            sSql &= " SELECT o31A.YEARS,o31A.COMIDNO,o31A.DISTID" & vbCrLf
            sSql &= " ,rr.ORGPLANNAME2,rr.ORGNAME" & vbCrLf
            sSql &= " ,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=o31A.DISTID) DISTNAME3" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o11A.YEARS) YEARS_ROC11A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o11A.APPSTAGE) APPSTAGE_N11A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o11A.COMIDNO,o11A.TPLANID,o11A.DISTID,o11A.YEARS,o11A.APPSTAGE) RLEVEL2_11A" & vbCrLf
            sSql &= " ,o11A.CLSN1 CLSN11A" & vbCrLf
            sSql &= " ,o11A.FIXSUMCOST FIXSUMCOST11A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o12A.YEARS) YEARS_ROC12A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12A.APPSTAGE) APPSTAGE_N12A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12A.COMIDNO,o12A.TPLANID,o12A.DISTID,o12A.YEARS,o12A.APPSTAGE) RLEVEL2_12A" & vbCrLf
            sSql &= " ,o12A.CLSN1 CLSN12A" & vbCrLf
            sSql &= " ,o12A.FIXSUMCOST FIXSUMCOST12A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21A.YEARS) YEARS_ROC21A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21A.APPSTAGE) APPSTAGE_N21A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21A.COMIDNO,o21A.TPLANID,o21A.DISTID,o21A.YEARS,o21A.APPSTAGE) RLEVEL2_21A" & vbCrLf
            sSql &= " ,o21A.CLSN1 CLSN21A" & vbCrLf
            sSql &= " ,o21A.FIXSUMCOST FIXSUMCOST21A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22A.YEARS) YEARS_ROC22A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22A.APPSTAGE) APPSTAGE_N22A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22A.COMIDNO,o22A.TPLANID,o22A.DISTID,o22A.YEARS,o22A.APPSTAGE) RLEVEL2_22A" & vbCrLf
            sSql &= " ,o22A.CLSN1 CLSN22A" & vbCrLf
            sSql &= " ,o22A.FIXSUMCOST FIXSUMCOST22A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31A.YEARS) YEARS_ROC31A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31A.APPSTAGE) APPSTAGE_N31A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31A.COMIDNO,o31A.TPLANID,o31A.DISTID,o31A.YEARS,o31A.APPSTAGE) RLEVEL2_31A" & vbCrLf
            sSql &= " ,o31A.CLSN1 CLSN31A" & vbCrLf
            sSql &= " ,o31A.FIXSUMCOST FIXSUMCOST31A" & vbCrLf

            '管制類課程
            sSql &= " ,dbo.FN_CYEAR2(o11.YEARS) YEARS_ROC11" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o11.APPSTAGE) APPSTAGE_N11" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o11.COMIDNO,o11.TPLANID,o11.DISTID,o11.YEARS,o11.APPSTAGE) RLEVEL2_11" & vbCrLf
            sSql &= " ,o11.CLSN1 CLSN11" & vbCrLf
            sSql &= " ,o11.FIXSUMCOST FIXSUMCOST11" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o12.YEARS) YEARS_ROC12" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12.APPSTAGE) APPSTAGE_N12" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12.COMIDNO,o12.TPLANID,o12.DISTID,o12.YEARS,o12.APPSTAGE) RLEVEL2_12" & vbCrLf
            sSql &= " ,o12.CLSN1 CLSN12" & vbCrLf
            sSql &= " ,o12.FIXSUMCOST FIXSUMCOST12" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21.YEARS) YEARS_ROC21" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21.APPSTAGE) APPSTAGE_N21" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21.COMIDNO,o21.TPLANID,o21.DISTID,o21.YEARS,o21.APPSTAGE) RLEVEL2_21" & vbCrLf
            sSql &= " ,o21.CLSN1 CLSN21" & vbCrLf
            sSql &= " ,o21.FIXSUMCOST FIXSUMCOST21" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22.YEARS) YEARS_ROC22" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22.APPSTAGE) APPSTAGE_N22" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22.COMIDNO,o22.TPLANID,o22.DISTID,o22.YEARS,o22.APPSTAGE) RLEVEL2_22" & vbCrLf
            sSql &= " ,o22.CLSN1 CLSN22" & vbCrLf
            sSql &= " ,o22.FIXSUMCOST FIXSUMCOST22" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31.YEARS) YEARS_ROC31" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31.APPSTAGE) APPSTAGE_N31" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31.COMIDNO,o31.TPLANID,o31.DISTID,o31.YEARS,o31.APPSTAGE) RLEVEL2_31" & vbCrLf
            sSql &= " ,o31.CLSN1 CLSN31" & vbCrLf
            sSql &= " ,o31.FIXSUMCOST FIXSUMCOST31" & vbCrLf

            sSql &= " ,rr.MASTERNAME,'' NU1,'' NU2,'' NU3,'' NU4,'' NU5,'' NU6" & vbCrLf

            sSql &= " FROM WO31A o31A" & vbCrLf
            sSql &= " LEFT JOIN WO11A o11A ON o11A.COMIDNO=o31A.COMIDNO AND o11A.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO12A o12A ON o12A.COMIDNO=o31A.COMIDNO AND o12A.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21A o21A ON o21A.COMIDNO=o31A.COMIDNO AND o21A.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22A o22A ON o22A.COMIDNO=o31A.COMIDNO AND o22A.DISTID =o31A.DISTID" & vbCrLf

            sSql &= " LEFT JOIN WO11 o11 ON o11.COMIDNO=o31A.COMIDNO AND o11.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO12 o12 ON o12.COMIDNO=o31A.COMIDNO AND o12.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21 o21 ON o21.COMIDNO=o31A.COMIDNO AND o21.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22 o22 ON o22.COMIDNO=o31A.COMIDNO AND o22.DISTID =o31A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31 o31 ON o31.COMIDNO=o31A.COMIDNO AND o31.DISTID =o31A.DISTID" & vbCrLf

            sSql &= " JOIN dbo.VIEW_RIDNAME rr ON rr.RID=o31A.RID" & vbCrLf
            sSql &= " ORDER BY o31A.YEARS,rr.ORGPLANNAME2 desc,o31A.DISTID,o31A.COMIDNO" & vbCrLf

        ElseIf v_ddlAPPSTAG1 = "2" Then
            sSql &= " ,WO32A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A2 cc WHERE cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO12A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO31A AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2A cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf

            sSql &= " ,WO32 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC3G cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO12 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS1 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO21 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=1)" & vbCrLf
            sSql &= " ,WO22 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS2 AND cc.APPSTAGE=2)" & vbCrLf
            sSql &= " ,WO31 AS (SELECT cc.RID,cc.YEARS,cc.COMIDNO,cc.APPSTAGE,cc.TPLANID,cc.DISTID,cc.CLSN1,cc.FIXSUMCOST FROM WC2 cc WHERE cc.YEARS=@D_YEARS3 AND cc.APPSTAGE=1)" & vbCrLf

            sSql &= " SELECT o32A.YEARS,o32A.COMIDNO,o32A.DISTID" & vbCrLf
            sSql &= " ,rr.ORGPLANNAME2,rr.ORGNAME" & vbCrLf
            sSql &= " ,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=o32A.DISTID) DISTNAME3" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o12A.YEARS) YEARS_ROC12A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12A.APPSTAGE) APPSTAGE_N12A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12A.COMIDNO,o12A.TPLANID,o12A.DISTID,o12A.YEARS,o12A.APPSTAGE) RLEVEL2_12A" & vbCrLf
            sSql &= " ,o12A.CLSN1 CLSN12A" & vbCrLf
            sSql &= " ,o12A.FIXSUMCOST FIXSUMCOST12A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21A.YEARS) YEARS_ROC21A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21A.APPSTAGE) APPSTAGE_N21A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21A.COMIDNO,o21A.TPLANID,o21A.DISTID,o21A.YEARS,o21A.APPSTAGE) RLEVEL2_21A" & vbCrLf
            sSql &= " ,o21A.CLSN1 CLSN21A" & vbCrLf
            sSql &= " ,o21A.FIXSUMCOST FIXSUMCOST21A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22A.YEARS) YEARS_ROC22A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22A.APPSTAGE) APPSTAGE_N22A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22A.COMIDNO,o22A.TPLANID,o22A.DISTID,o22A.YEARS,o22A.APPSTAGE) RLEVEL2_22A" & vbCrLf
            sSql &= " ,o22A.CLSN1 CLSN22A" & vbCrLf
            sSql &= " ,o22A.FIXSUMCOST FIXSUMCOST22A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31A.YEARS) YEARS_ROC31A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31A.APPSTAGE) APPSTAGE_N31A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31A.COMIDNO,o31A.TPLANID,o31A.DISTID,o31A.YEARS,o31A.APPSTAGE) RLEVEL2_31A" & vbCrLf
            sSql &= " ,o31A.CLSN1 CLSN31A" & vbCrLf
            sSql &= " ,o31A.FIXSUMCOST FIXSUMCOST31A" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o32A.YEARS) YEARS_ROC32A" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o32A.APPSTAGE) APPSTAGE_N32A" & vbCrLf
            sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o32A.COMIDNO,o32A.TPLANID,o32A.DISTID,o32A.YEARS,o32A.APPSTAGE) RLEVEL2_32A" & vbCrLf
            sSql &= " ,o32A.CLSN1 CLSN32A" & vbCrLf
            sSql &= " ,o32A.FIXSUMCOST FIXSUMCOST32A" & vbCrLf
            '---
            sSql &= " ,dbo.FN_CYEAR2(o12.YEARS) YEARS_ROC12" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o12.APPSTAGE) APPSTAGE_N12" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o12.COMIDNO,o12.TPLANID,o12.DISTID,o12.YEARS,o12.APPSTAGE) RLEVEL2_12" & vbCrLf
            sSql &= " ,o12.CLSN1 CLSN12" & vbCrLf
            sSql &= " ,o12.FIXSUMCOST FIXSUMCOST12" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o21.YEARS) YEARS_ROC21" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o21.APPSTAGE) APPSTAGE_N21" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o21.COMIDNO,o21.TPLANID,o21.DISTID,o21.YEARS,o21.APPSTAGE) RLEVEL2_21" & vbCrLf
            sSql &= " ,o21.CLSN1 CLSN21" & vbCrLf
            sSql &= " ,o21.FIXSUMCOST FIXSUMCOST21" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o22.YEARS) YEARS_ROC22" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o22.APPSTAGE) APPSTAGE_N22" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o22.COMIDNO,o22.TPLANID,o22.DISTID,o22.YEARS,o22.APPSTAGE) RLEVEL2_22" & vbCrLf
            sSql &= " ,o22.CLSN1 CLSN22" & vbCrLf
            sSql &= " ,o22.FIXSUMCOST FIXSUMCOST22" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o31.YEARS) YEARS_ROC31" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o31.APPSTAGE) APPSTAGE_N31" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31.COMIDNO,o31.TPLANID,o31.DISTID,o31.YEARS,o31.APPSTAGE) RLEVEL2_31" & vbCrLf
            sSql &= " ,o31.CLSN1 CLSN31" & vbCrLf
            sSql &= " ,o31.FIXSUMCOST FIXSUMCOST31" & vbCrLf

            sSql &= " ,dbo.FN_CYEAR2(o32.YEARS) YEARS_ROC32" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(o32.APPSTAGE) APPSTAGE_N32" & vbCrLf
            'sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(o31.COMIDNO,o31.TPLANID,o31.DISTID,o31.YEARS,o31.APPSTAGE) RLEVEL2_31" & vbCrLf
            sSql &= " ,o32.CLSN1 CLSN32" & vbCrLf
            sSql &= " ,o32.FIXSUMCOST FIXSUMCOST32" & vbCrLf

            sSql &= " ,rr.MASTERNAME,'' NU1,'' NU2,'' NU3,'' NU4,'' NU5,'' NU6" & vbCrLf
            sSql &= " FROM WO32A o32A" & vbCrLf
            sSql &= " LEFT JOIN WO12A o12A ON o12A.COMIDNO=o32A.COMIDNO AND o12A.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21A o21A ON o21A.COMIDNO=o32A.COMIDNO AND o21A.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22A o22A ON o22A.COMIDNO=o32A.COMIDNO AND o22A.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31A o31A ON o31A.COMIDNO=o32A.COMIDNO AND o31A.DISTID=o32A.DISTID" & vbCrLf

            sSql &= " LEFT JOIN WO12 o12 ON o12.COMIDNO=o32A.COMIDNO AND o12.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO21 o21 ON o21.COMIDNO=o32A.COMIDNO AND o21.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO22 o22 ON o22.COMIDNO=o32A.COMIDNO AND o22.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO31 o31 ON o31.COMIDNO=o32A.COMIDNO AND o31.DISTID=o32A.DISTID" & vbCrLf
            sSql &= " LEFT JOIN WO32 o32 ON o32.COMIDNO=o32A.COMIDNO AND o32.DISTID=o32A.DISTID" & vbCrLf

            sSql &= " JOIN dbo.VIEW_RIDNAME rr ON rr.RID=o32A.RID" & vbCrLf
            sSql &= " ORDER BY o32A.YEARS,rr.ORGPLANNAME2 desc,o32A.DISTID,o32A.COMIDNO" & vbCrLf
        End If

        'Dim s_sSqlsParms As String = String.Concat(sSql, TIMS.GetMyValue4(sParms))
        'TIMS.WriteTraceLog(s_sSqlsParms)

        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sSql, objconn)
        Call DbAccess.HashParmsChange(sCmd, sParms)
        dt.Load(sCmd.ExecuteReader())
        Call TIMS.CHG_dtReadOnly(dt)
        Return dt
    End Function

End Class
