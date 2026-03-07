Partial Class SD_05_010_pop
    Inherits AuthBasePage

    Dim ff As String = "" '.Length > 0 
    Dim CPdt As DataTable
    'Dim FunDr As DataRow
    'https://jira.turbotech.com.tw/browse/TIMSC-56
    '參訓的訓練機構欄位資訊，不顯示。
    Dim flag_NoOrgName As Boolean = False '
    Dim dtIdentity As DataTable

    Dim but_S_Export As Boolean = False '是否使用匯出鈕功能
    Const cst_msg1 As String = "委訓機構只能查詢有參加過該機構的學員!!"

    Const CST_KD_STUDENTLIST As String = "StudentList" 'Session("IDNOArray")

    '找尋所有計畫，排除 下列計畫
    '1.接受企業委託訓練 [07]
    '2.在職進修訓練 [06]
    '3.產業人才投資方案 [28]
    '4.充電起飛計畫(補助在職勞工參訓) [54]
    '由e網審核報名傳過來的
    'Dim bln_SD01004Type As Boolean = False  '由e網審核報名傳過來的
    'Dim blnSearchType1 As Boolean = False '執行 排除在職sql
    'Dim blnSearchType1 As Boolean = True '執行 不排除在職搜尋(全部搜尋)

    'False:非在職計畫，排除搜尋 下列 在職計畫 'True:屬在職計畫，不排除搜尋任何計畫
    'SD_01_004「近二年參訓歷史」與「查詢參訓歷史」
    'SD_05_010.aspx
    '1.接受企業委託訓練 [07]  委託訓練
    '2.在職進修訓練 [06] 在職進修
    '3.產業人才投資方案 [28] 產業人才
    '4.充電起飛計畫(補助在職勞工參訓) [54] 充電起飛
    'Public Const TIMS.Cst_NONTPlanID3 As String = "'07','06','28','54'"

    Dim rqSD_01_004_Type As String = "" 'TIMS.ClearSQM(Request("SD_01_004_Type"))
    Dim rqIDNO As String = "" '
    Dim rqTwoYears As String = ""
    Dim rqBtnHistory As String = ""
    'Dim rqState As String = "" 'Request("state")
    'Dim rqPlanYear As String = "" 'Request("PlanYear")
    'Dim rqStart_date As String = "" 'Request("start_date")
    'Dim rqEnd_date As String = "" 'Request("end_date")
    'Dim rqDistID As String = "" 'Request("DistID")
    'Dim rqToken As String = "" 'TOKEN要有特定值 sid=tims

    Const cst_RTREASONID_02 As String = "02" '02:訓期已滿1/2提前就業
    Const cst_RTREASONID_02_NAME As String = "就業(提前就業)" '02:訓期已滿1/2提前就業
    Const cst_FTDateM3_NAME1 As String = "訓後3個月內"
    Const cst_FTDateM3_NAME2 As String = "無記錄"
    'Const cst_inline As String = "inline"
    Const cst_inline As String = ""

    Dim dtRejectTReason As DataTable = Nothing

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    '是否啟用西元轉民國年機制
    Dim flag_Roc As Boolean = False

    Const cst_dg2_出生日期 As Integer = 3
    Const cst_dg2_年度 As Integer = 5

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)

        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid2

        rqSD_01_004_Type = TIMS.ClearSQM(Request("SD_01_004_Type"))
        rqIDNO = TIMS.ClearSQM(Request("IDNO"))
        Dim rqENCIDNO As String = TIMS.ClearSQM(Request("ENCIDNO"))
        If rqENCIDNO <> "" Then
            '若有傳入參數 ENCIDNO，直接解密並蓋過身分證號
            rqIDNO = RSA20031.AesDecrypt2(rqENCIDNO)
        End If
        rqTwoYears = TIMS.ClearSQM(Request("TwoYears"))
        rqBtnHistory = TIMS.ClearSQM(Request("BtnHistory"))

        Dim sql As String = "SELECT RTREASONID,REASON FROM dbo.KEY_REJECTTREASON ORDER BY RTREASONID"
        dtRejectTReason = DbAccess.GetDataTable(sql, objconn)

        'bln_SD01004Type = False '自己程式呼叫
        'If Request("SD_01_004_Type") <> "" Then  '由e網審核報名傳過來的
        '    bln_SD01004Type = True '由e網審核報名傳過來的
        '    blnSearchType1 = False  '執行 排除在職sql
        '    If TIMS.Cst_NONTPlanID3.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        'True:屬在職計畫，不排除搜尋任何計畫
        '        blnSearchType1 = True
        '    End If
        'End If
        'blnSearchType1 = True '永遠搜尋全部。

        Button2.Style("display") = "none"

        flag_Roc = TIMS.CHK_REPLACE2ROC_YEARS()

        If Not IsPostBack Then
            CCreate1()
        End If

        tr01a.Style.Item("display") = cst_inline
        tr01b.Style.Item("display") = cst_inline
        tr01c.Style.Item("display") = cst_inline
        tr01d.Style.Item("display") = cst_inline
        tr01e.Style.Item("display") = cst_inline
        tr02d.Style.Item("display") = cst_inline '顯示

        Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "0"
                'tr02d.Style.Item("display") = cst_inline '顯示
                Select Case sm.UserInfo.RoleID
                    Case "0", "1" '系統管理者
                    Case Else
                        tr02d.Style.Item("display") = "none" '不顯示
                End Select
            Case "1"
                'Common.SetListItem(DistID, sm.UserInfo.DistID)
                'DistID.Enabled = False
                Select Case sm.UserInfo.RoleID
                    Case "0", "1" '系統管理者
                    Case Else
                        tr02d.Style.Item("display") = "none" '不顯示
                End Select
            Case Else
                tr01a.Style.Item("display") = "none"
                tr01b.Style.Item("display") = "none"
                tr01c.Style.Item("display") = "none"
                tr01d.Style.Item("display") = "none"
                tr01e.Style.Item("display") = "none"

                tr02d.Style.Item("display") = "none" '不顯示
        End Select

        '檢查帳號的功能權限-----------------------------------Start
        'If rqSD_01_004_Type = "" Then
        '    '如果不是由其他功能導進此頁的
        '    Button1.Enabled = False
        '    If au.blnCanSech Then Button1.Enabled = True
        'End If
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    Sub CCreate1()
        msg.Text = ""
        Button1.Attributes("onclick") = "return chkdata();"
        DistID.Attributes("onchange") = "GetMode();"
        TPlanID.Attributes("onchange") = "GetMode();"
        OCID.Attributes("onchange") = "if(this.selectedIndex!=0){document.form1.OCIDValue.value=this.value;}else{document.form1.OCIDValue.value='';}"
        but_S_Export = False '是否使用匯出鈕功能
        DistID = TIMS.Get_DistID(DistID, TIMS.dtNothing, objconn)
        TPlanID = TIMS.Get_TPlan(TPlanID, , , , , objconn)
        ShowDataTable.Style.Item("display") = "none"
        OCID.Items.Add("請選擇機構")

        '自己程式
        but_P.Visible = False '列印
        but_S.Visible = False '匯出
        trRBListExpType.Visible = False '匯出
        Btnclose.Visible = False '關閉鍵
        Button4.Visible = False '回上一頁

        If rqSD_01_004_Type = "" Then
            Button4.Visible = True  '回上一頁
        End If
        If rqSD_01_004_Type <> "" Then  '由e網審核報名傳過來的
            'Button1_Click(sender, e)
            Call Search1() '查詢(條件)
            Btnclose.Visible = True  '關閉鍵
            Select Case rqSD_01_004_Type
                Case CST_KD_STUDENTLIST '"StudentList"
                    '由e網審核報名傳過來的學員參訓歷史查詢List
                    but_P.Visible = True    '列印
                    but_S.Visible = True    '匯出
                    trRBListExpType.Visible = True '匯出
            End Select
        End If
    End Sub

    ''' <summary>
    ''' 將搜尋資料加入dt (有2個sub使用) '(含有效資料與已被刪除的資料)
    ''' </summary>
    ''' <param name="dt">將搜尋資料加入dt</param>
    ''' <param name="dt3">搜尋資料</param>
    ''' <param name="RecordCountInt">筆數限制</param>
    ''' <remarks></remarks>
    Sub SUtl_AddDt3(ByRef dt As DataTable, ByRef dt3 As DataTable, ByRef RecordCountInt As Integer)
        Dim dr As DataRow = Nothing

        For Each dr3 As DataRow In dt3.Select("", "STDate") 'For Each dr3 In dt3.Rows
            If RecordCountInt > 0 Then RecordCountInt -= 1
            If Not (RecordCountInt > 0) Then Exit For '超過 最大筆數限制

            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("VSSORT") = If(Session("IDNOArray") Is Nothing, 99, TIMS.Get_VSSORT(TIMS.ChangeIDNO(dr3("IDNO")), Session("IDNOArray")))
            dr("IDNO") = TIMS.ChangeIDNO(dr3("IDNO")) '1.身分證號
            dr("Name") = dr3("Name") '1.姓名
            dr("Sex") = dr3("Sex") '性別
            dr("Birthday") = dr3("Birthday") '2.出生年月日
            dr("DistName") = dr3("DistName") '2.分署
            dr("Years") = dr3("Years") '3.訓練年度
            dr("PlanName") = dr3("PlanName") '4.訓練計畫
            If Not flag_NoOrgName Then dr("OrgName") = dr3("OrgName") '5.訓練機構
            dr("TMID") = dr3("TMID") '6.訓練職類
            dr("CJOB_NAME") = dr3("CJOB_NAME") '7.通俗職類
            dr("ClassName") = dr3("ClassName") '8.班別名稱

            Dim strSTDate As String = If(flag_Roc, TIMS.Cdate17(dr3("STDate")), TIMS.Cdate3(dr3("STDate")))
            Dim strFTDate As String = If(flag_Roc, TIMS.Cdate17(dr3("FTDate")), TIMS.Cdate3(dr3("FTDate")))

            'THours: '9.受訓時數
            'TRound: '10.受訓期間
            Dim fg_isDate As Boolean = False
            Select Case $"{dr3("StudStatus")}" '訓練狀態，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                Case "2" '"離訓"
                    dr("THours") = String.Concat("<FONT color='Red'>", dr3("TrainHours"), "</FONT>") '參訓時數，以 Class_StudentsOfClass 為主
                    Dim strRejectTDate1 As String = TIMS.TryFormatDate(dr3, "RejectTDate1", "離訓日期異常", fg_isDate)
                    If (fg_isDate) Then strRejectTDate1 = If(flag_Roc, TIMS.Cdate17(dr3("RejectTDate1")), TIMS.Cdate3(dr3("RejectTDate1")))
                    dr("TRound") = String.Concat(strSTDate, "<BR>|<BR>", strRejectTDate1)
                Case "3" '"退訓"
                    dr("THours") = String.Concat("<FONT color='Red'>", dr3("TrainHours"), "</FONT>") '參訓時數，以 Class_StudentsOfClass 為主
                    Dim strRejectTDate2 As String = TIMS.TryFormatDate(dr3, "RejectTDate2", "退訓日期異常", fg_isDate)
                    If (fg_isDate) Then strRejectTDate2 = If(flag_Roc, TIMS.Cdate17(dr3("RejectTDate2")), TIMS.Cdate3(dr3("RejectTDate2")))
                    dr("TRound") = String.Concat(strSTDate, "<BR>|<BR>", strRejectTDate2)
                Case Else
                    dr("THours") = dr3("THours") '參訓時數，以 Class_StudentsOfClass 為優先資料顯示 Class_ClassInfo 為副
                    dr("TRound") = String.Concat(strSTDate, "<BR>|<BR>", strFTDate)
            End Select

            dr("SkillName") = dr3("ExamName") '11.技能檢定
            dr("WEEKS") = dr3("WEEKS")  '12.上課時間

            '13.訓練狀態
            dr("TFlag") = TIMS.CHG_TFLAG($"{dr3("StudStatus")}")

            Dim flagJOB23 As Boolean = False '只要 cst_RTREASONID_02 就直接定義為就業 by 20161107
            '補離退資訊 '14.遞補期限內離訓(※註)
            Select Case $"{dr3("StudStatus")}" '.ToString
                Case "2", "3" '2:離訓'3:退訓
                    Dim mykey As String = TIMS.ConvertStr(dr3("RTReasonID"))
                    If mykey = cst_RTREASONID_02 Then
                        '只要 cst_RTREASONID_02 就直接定義為就業 by 20161107
                        flagJOB23 = True
                    End If
                    Dim myMsg1 As String = ""
                    Dim myMsg2 As String = ""
                    If mykey <> "" Then mykey = Trim(mykey)
                    If mykey <> "" Then
                        ff = "RTReasonID='" & mykey & "'"
                        If dtRejectTReason.Select(ff).Length > 0 Then myMsg1 = dtRejectTReason.Select(ff)(0)("Reason")
                    End If
                    If myMsg1 <> "" Then myMsg1 = Trim(myMsg1)
                    If myMsg1 <> "" Then dr("TFlag") &= "：" & myMsg1

                    myMsg2 = TIMS.ConvertStr(dr3("RTReasoOther"))
                    If myMsg2 <> "" Then myMsg2 = Trim(myMsg2)
                    If myMsg2 <> "" Then dr("TFlag") &= "(" & myMsg2 & ")"

                    'RejectDayIn14: 遞補期限內離訓(※註)
                    Select Case $"{dr3("RejectDayIn14")}"
                        Case "Y"
                            dr("RejectDayIn14") = "是"
                            'dr("TFlag") &= "(遞補期限內離訓：是)"
                        Case "N"
                            dr("RejectDayIn14") = "否"
                            'dr("TFlag") &= "(遞補期限內離訓：否)"
                    End Select
            End Select

            '參訓身分
            ff = "IdentityID='" & dr3("MIdentityID") & "'"
            dr("Ident") = If(dtIdentity.Select(ff).Length > 0, dtIdentity.Select(ff)(0)("Name"), "無身分別")
            '電話1
            dr("Tel") = $"{dr3("PhoneD")}"
            '地址。
            dr("Address") = $"{TIMS.Get_ZipName(dr3("ZipCode1"))}{dr3("Address")}"

            '就業狀況
            'If $"{dr3("JobStatus")) <> "" Then
            '    dr("JobStatus") = dr3("JobStatus") '15.訓後就業狀況
            'Else
            '    If $"{dr3("FTDateM3")) = "Y" Then
            '        dr("JobStatus") = cst_FTDateM3_NAME1 '"訓後3個月內" '15.訓後就業狀況
            '    Else
            '        dr("JobStatus") = cst_FTDateM3_NAME2 '"無記錄" '15.訓後就業狀況
            '    End If
            'End If
            'If flagJOB23 Then
            '    '只要 cst_RTREASONID_02 就直接定義為就業 by 20161107
            '    '今年1/1之後結訓的班級 學員屬提前就業者,請在訓後就業狀況的狀態 統一調整為"就業"
            '    dr("JobStatus") = cst_RTREASONID_02_NAME '"就業"
            'End If
            '就業單位名稱(BusName)
            'dr("JobOrgName") = $"{dr3("JobOrgName"))
            'https://jira.turbotech.com.tw/browse/TIMSC-201
            '多加一個"備註"欄位，如果，參訓歷史的計畫是「托育人員職業訓練」及「照顧服務員職業訓練」，備註欄位顯示「在職者」或「待業者」。
            dr("MEMO1") = $"{dr3("MEMO1")}"
        Next
    End Sub

    '查詢 Show出資料 (含SQL)
    Sub GetStudentData(ByVal SearchStr1 As String, ByVal SearchStr2 As String, ByVal SearchStr3 As String)
        ' FROM StdAll、History_StudentInfo93、Class_StudentsOfClass
        '含有舊table資料的搜尋.

        '排除在職sql
        'Dim sSql_notPlanName1 As String = ""
        'Dim sSql_notPlanName2 As String = ""
        'Dim sSql_notPlanName3 As String = ""
        'sSql_notPlanName1 = ""
        'sSql_notPlanName2 = ""
        'sSql_notPlanName3 = ""

        '由e網審核報名傳過來的 執行排除功能
        'blnSearchType1 (false:不搜尋在職 true:搜尋全部不排除)
        'If bln_SD01004Type AndAlso Not blnSearchType1 Then
        '    '排除下列計畫名稱與計畫號。
        '    sSql_notPlanName1 = ""
        '    sSql_notPlanName1 += " AND NOT EXISTS (SELECT 'x' FROM StdAll x WHERE x.StdID=m.StdID"
        '    sSql_notPlanName1 += " AND (1!=1"
        '    sSql_notPlanName1 += " OR x.planname like '%委託訓練%'"
        '    sSql_notPlanName1 += " OR x.planname like '%在職進修%'"
        '    sSql_notPlanName1 += " OR x.planname like '%產業人才%'"
        '    sSql_notPlanName1 += " OR x.planname like '%充電起飛%'"
        '    sSql_notPlanName1 += " ) )"

        '    sSql_notPlanName2 = ""
        '    sSql_notPlanName2 += " AND NOT EXISTS (SELECT 'x' FROM History_StudentInfo93 x WHERE x.Serial=a.Serial"
        '    sSql_notPlanName2 += "  AND x.TPlanID IN (" & TIMS.Cst_NONTPlanID3 & ") "
        '    sSql_notPlanName2 += " )"

        '    sSql_notPlanName3 = ""
        '    sSql_notPlanName3 += " AND NOT EXISTS (SELECT 'x' FROM ID_Plan x WHERE x.PlanID=i.PlanID"
        '    sSql_notPlanName3 += "  AND x.TPlanID IN (" & TIMS.Cst_NONTPlanID3 & ") "
        '    sSql_notPlanName3 += " )"
        'End If

        'rqSD_01_004_Type = TIMS.ClearSQM(Request("SD_01_004_Type"))
        Dim RecordCountInt As Integer = 2000 '最大筆數限制
        dtIdentity = TIMS.Get_KeyTable("KEY_IDENTITY", "", objconn)

        Dim dt As New DataTable
        TIMS.INIT_SPECdt(dt)

        Dim dr As DataRow = Nothing
        Dim sql As String = $"SELECT * FROM dbo.StdAll m WHERE 1=1 {SearchStr1}"
        'If sSql_notPlanName1 <> "" Then
        '    sql += sSql_notPlanName1
        'End If
        Dim dt1 As DataTable
        Try
            dt1 = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try
        For Each dr1 As DataRow In dt1.Rows
            If RecordCountInt > 0 Then RecordCountInt -= 1
            If Not (RecordCountInt > 0) Then Exit For '超過 最大筆數限制

            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("VSSORT") = If(Session("IDNOArray") Is Nothing, 99, TIMS.Get_VSSORT(TIMS.ChangeIDNO(dr1("SID")), Session("IDNOArray")))
            dr("IDNO") = TIMS.ChangeIDNO(dr1("SID"))
            dr("Name") = dr1("Name")
            dr("Sex") = dr1("Sex")
            dr("Birthday") = dr1("Birth")
            dr("DistName") = dr1("DistName")
            dr("Years") = dr1("Years")
            dr("PlanName") = dr1("PlanName")
            If Not flag_NoOrgName Then dr("OrgName") = dr1("TrinUnit")
            'dr("TMID") = dr1("")
            dr("ClassName") = dr1("ClassName")
            'dr("THours") = dr1("")
            If dr1("SDate").ToString <> "" And dr1("EDate").ToString <> "" Then dr("TRound") = Common.FormatDate(dr1("SDate")) & "<BR>|<BR>" & Common.FormatDate(dr1("EDate"))
            'dr("SkillName") = dr1("")
            dr("TFlag") = "結訓." '預設結訓
            dr("Ident") = If(IsNumeric(dr1("Ident")), "無法辨別", dr1("Ident").ToString)
            dr("Tel") = dr1("Tel").ToString
            dr("Address") = dr1("Addr").ToString
        Next

        sql = ""
        sql &= " SELECT a.* ,b.TrainName"
        sql &= $" FROM (SELECT * FROM dbo.HISTORY_STUDENTINFO93 WHERE 1=1 {SearchStr2}) a "
        sql &= " LEFT JOIN dbo.KEY_TRAINTYPE b ON a.TMID=b.TMID"
        sql &= " WHERE 1=1"
        'If sSql_notPlanName2 <> "" Then
        '    sql += sSql_notPlanName2
        'End If
        Dim dt2 As DataTable
        Try
            dt2 = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try

        For Each dr2 As DataRow In dt2.Rows
            If RecordCountInt > 0 Then RecordCountInt -= 1
            If Not (RecordCountInt > 0) Then Exit For '超過 最大筆數限制

            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("VSSORT") = If(Session("IDNOArray") Is Nothing, 99, TIMS.Get_VSSORT(TIMS.ChangeIDNO(dr2("IDNO")), Session("IDNOArray")))
            dr("IDNO") = TIMS.ChangeIDNO(dr2("IDNO"))
            dr("Name") = dr2("Name")
            dr("Sex") = dr2("Sex")
            dr("Birthday") = dr2("Birth")
            dr("DistName") = dr2("DistName")
            dr("PlanName") = dr2("PlanName")
            If Not flag_NoOrgName Then dr("OrgName") = dr2("TrinUnit")
            dr("TMID") = dr2("TrainName")
            dr("ClassName") = dr2("ClassName")

            'dr("TRound") = Common.FormatDate(dr2("SDate")) & "<BR>|<BR>" & Common.FormatDate(dr2("EDate"))
            Dim strSDate As String = If(flag_Roc, TIMS.Cdate17(dr2("SDate")), TIMS.Cdate3(dr2("SDate")))
            Dim strEDate As String = If(flag_Roc, TIMS.Cdate17(dr2("EDate")), TIMS.Cdate3(dr2("EDate")))
            dr("TRound") = strSDate & "<BR>|<BR>" & strEDate

            dr("TFlag") = "結訓." '預設結訓
            ff = "IdentityID='" & dr2("Ident") & "'"
            dr("Ident") = If(dtIdentity.Select(ff).Length > 0, dtIdentity.Select(ff)(0)("Name"), "無身分別")
            dr("Tel") = dr2("Tel").ToString
            dr("Address") = dr2("Addr").ToString
        Next

        '/*
        '資料重複請刪除()
        'select * 
        'from Stud_TechExam 
        'where socid in ( SELECT SOCID FROM Stud_TechExam  group by SOCID having count(*) > 1 ) 
        'order by socid 
        'delete Stud_TechExam 
        'where socid in ( SELECT SOCID  FROM Stud_TechExam  group by SOCID having count(*) > 1 ) 
        'and steid in (SELECT max(steid) setid  FROM Stud_TechExam  group by SOCID having count(*) > 1 )
        '*/
        'Dim CJOBStr As String = "" '& CJOBStr 'MySqlStr += CJOBStr 'sql += CJOBStr & vbCrLf
        'If cjobValue.Value <> "" Then
        '    'CJOBStr += " and c.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        '    CJOBStr += " and c.CJOB_UNKEY=" & cjobValue.Value & vbCrLf
        'End If

        'Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT b.IDNO" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,b.Sex" & vbCrLf
        sql &= " ,b.Birthday" & vbCrLf
        sql &= " ,f.Name DistName" & vbCrLf
        sql &= " ,e.OrgName" & vbCrLf
        sql &= " ,ISNULL(g.TrainName,g.JOBNAME) TMID" & vbCrLf
        'sql &= " ,c.ClassCName +'第' +c.cyclType +'期' ClassName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,case when a.TrainHours is null then c.THours else a.TrainHours end THours" & vbCrLf
        '2009/08/25 改成以班級的開結訓日為開結訓日
        sql &= " ,c.STDate" & vbCrLf
        sql &= " ,c.FTDate" & vbCrLf
        'FTDateM3'訓後3個月內
        sql &= " ,CASE WHEN DATEADD(month, 3, c.FTDate) >= GETDATE() AND c.FTDate <= GETDATE() THEN 'Y' END FTDateM3" & vbCrLf
        sql &= " ,a.TrainHours" & vbCrLf
        sql &= " ,a.RejectTDate1" & vbCrLf
        sql &= " ,a.RejectTDate2" & vbCrLf
        'NVL , NVL2 ,COALESCE ,DECODE
        sql &= " ,ISNULL(h.ExamName,ISNULL(h.ExamName2,h.ExamName3)) ExamName" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,a.MIdentityID" & vbCrLf
        sql &= " ,j.PhoneD" & vbCrLf
        sql &= " ,j.ZipCode1" & vbCrLf
        sql &= " ,j.Address" & vbCrLf
        sql &= " ,k.PlanName" & vbCrLf
        sql &= " ,i.Years" & vbCrLf
        sql &= " ,s.CJOB_NAME" & vbCrLf
        sql &= " ,CASE WHEN i.TPLANID='06' THEN c.NOTE3" & vbCrLf
        sql &= " ELSE dbo.FN_GET_PLAN_ONCLASS(pp.PlanID,pp.ComIDNO,pp.SeqNo,'WEEKTIME') END WEEKS" & vbCrLf
        'dbo.fn_GET_JOBSTATUS(sg3.IsGetJob,sg3.PUBLICRESCUE)
        'sql &= " --,sg3.IsGetJob IsGetJob" & vbCrLf
        'sql &= " --,dbo.FN_GET_JOBSTATUS(sg3.IsGetJob,sg3.PUBLICRESCUE) JobStatus" & vbCrLf
        'sql &= " ,sg9.IsGetJob IsGetJob9"
        'sql &= " ,dbo.fn_GET_JOBSTATUS(sg9.IsGetJob,sg9.PUBLICRESCUE) JobStatus9" & vbCrLf 'else '未填寫'
        'sql &= " ,dbo.DECODE6(sg3.IsGetJob,'0','未就業','1','就業','2','不就業') JobStatus" & vbCrLf 'else '未填寫'
        sql &= " ,a.RTReasonID" & vbCrLf
        'else '未填寫'
        sql &= " ,a.RTReasoOther" & vbCrLf
        'RejectDayIn14'(兩週內)離退訓
        sql &= " ,a.RejectDayIn14" & vbCrLf
        'JobOrgName BusName	事業單位名稱
        'sql &= " --,ISNULL(sg3.BusName,bli.COMNAME) JobOrgName" & vbCrLf
        'https://jira.turbotech.com.tw/browse/TIMSC-201
        '多加一個"備註"欄位，如果，參訓歷史的計畫是「托育人員職業訓練」及「照顧服務員職業訓練」，備註欄位顯示「在職者」或「待業者」。
        sql &= " ,case when i.TPLANID in ('46','47','69','58') then dbo.DECODE(a.WorkSuppIdent,'Y','在職者','待業者') end MEMO1" & vbCrLf
        sql &= " FROM dbo.CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO b ON b.SID=a.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_SUBDATA j ON j.SID=a.SID" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO c ON a.OCID=c.OCID" & vbCrLf
        sql &= " JOIN dbo.Plan_PlanInfo pp on c.planid = pp.planid and pp.comidno = c.comidno and pp.seqno = c.seqno" & vbCrLf
        sql &= " JOIN dbo.Auth_Relship d ON c.RID=d.RID" & vbCrLf
        sql &= " JOIN dbo.Org_OrgInfo e ON d.OrgID=e.OrgID" & vbCrLf
        sql &= " JOIN dbo.ID_Plan i ON i.PlanID=c.PlanID" & vbCrLf
        sql &= " JOIN dbo.KEY_plan k ON k.TPlanID=i.TPlanID" & vbCrLf
        sql &= " LEFT JOIN dbo.ID_District f ON d.DistID=f.DistID" & vbCrLf
        sql &= " LEFT JOIN dbo.Key_TrainType g ON c.TMID=g.TMID" & vbCrLf
        sql &= " LEFT JOIN dbo.SHARE_CJOB s on s.CJOB_UNKEY=c.CJOB_UNKEY" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_TECHEXAM h ON a.SOCID=h.SOCID" & vbCrLf
        'sql &= " --LEFT JOIN dbo.STUD_GETJOBSTATE3 sg3 ON sg3.CPoint=1 and sg3.SOCID =a.SOCID" & vbCrLf
        'sql &= " --LEFT JOIN dbo.STUD_BLIGATEDATA bli ON bli.SBID=sg3.SBID" & vbCrLf
        'sql &= " LEFT JOIN STUD_GETJOBSTATE3 sg9 ON sg9.CPoint=9 and sg9.SOCID =a.SOCID" & vbCrLf
        sql &= $" WHERE 1=1 {SearchStr3}" & vbCrLf
        If cjobValue.Value <> "" Then
            sql &= " and c.CJOB_UNKEY=" & cjobValue.Value & vbCrLf
        End If
        'sql += CJOBStr & vbCrLf
        'If sSql_notPlanName3 <> "" Then sql += sSql_notPlanName3
        Dim dt3 As DataTable
        Try
            dt3 = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try
        '將搜尋資料加入dt
        Call SUtl_AddDt3(dt, dt3, RecordCountInt)

        'CLASS_STUDENTSOFCLASSDELDATA
        'Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT b.IDNO" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,b.Sex" & vbCrLf
        sql &= " ,b.Birthday" & vbCrLf
        sql &= " ,f.Name DistName" & vbCrLf
        sql &= " ,e.OrgName" & vbCrLf
        sql &= " ,ISNULL(g.TrainName,g.JOBNAME) TMID" & vbCrLf
        'sql &= " ,c.ClassCName +'第' +c.cyclType +'期' ClassName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,case when a.TrainHours is null then c.THours else a.TrainHours end THours" & vbCrLf
        sql &= " ,c.STDate" & vbCrLf
        sql &= " ,c.FTDate" & vbCrLf
        sql &= " ,CASE WHEN DATEADD(month, 3, c.FTDate) >= GETDATE() AND c.FTDate <= GETDATE() THEN 'Y' END FTDateM3" & vbCrLf
        sql &= " ,a.TrainHours" & vbCrLf
        sql &= " ,a.RejectTDate1" & vbCrLf
        sql &= " ,a.RejectTDate2" & vbCrLf
        sql &= " ,ISNULL(h.ExamName,ISNULL(h.ExamName2,h.ExamName3)) ExamName" & vbCrLf
        '不符參訓資格(CLASS_STUDENTSOFCLASSDELDATA)
        sql &= " ,9 StudStatus" & vbCrLf
        sql &= " ,a.MIdentityID" & vbCrLf
        sql &= " ,j.PhoneD" & vbCrLf
        sql &= " ,j.ZipCode1" & vbCrLf
        sql &= " ,j.Address" & vbCrLf
        sql &= " ,k.PlanName" & vbCrLf
        sql &= " ,i.Years" & vbCrLf
        sql &= " ,s.CJOB_NAME" & vbCrLf
        sql &= " ,CASE WHEN i.TPLANID='06' THEN c.NOTE3" & vbCrLf
        sql &= " ELSE dbo.FN_GET_PLAN_ONCLASS(pp.PlanID,pp.ComIDNO,pp.SeqNo,'WEEKTIME') END WEEKS" & vbCrLf
        'sql &= " --,sg3.IsGetJob IsGetJob" & vbCrLf
        'sql &= " --,dbo.FN_GET_JOBSTATUS(sg3.IsGetJob,sg3.PUBLICRESCUE) JobStatus" & vbCrLf
        sql &= " ,a.RTReasonID" & vbCrLf
        sql &= " ,a.RTReasoOther" & vbCrLf
        'RejectDayIn14'(兩週內)離退訓
        sql &= " ,a.RejectDayIn14" & vbCrLf
        'JobOrgName BusName	事業單位名稱
        'sql &= " --,ISNULL(sg3.BusName,bli.COMNAME) JobOrgName" & vbCrLf
        'https://jira.turbotech.com.tw/browse/TIMSC-201
        '多加一個"備註"欄位，如果，參訓歷史的計畫是「托育人員職業訓練」及「照顧服務員職業訓練」，備註欄位顯示「在職者」或「待業者」。
        sql &= " ,case when i.TPLANID in ('46','47','69','58') then dbo.DECODE (a.WorkSuppIdent,'Y','在職者','待業者') end MEMO1" & vbCrLf
        sql &= " FROM dbo.CLASS_STUDENTSOFCLASSDELDATA a" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO b ON b.SID=a.SID" & vbCrLf
        sql &= " JOIN dbo.STUD_SUBDATA j ON j.SID=a.SID" & vbCrLf
        '刪除資料有限定 dg.DELREASON='4'
        sql &= " JOIN dbo.SYS_DELLOG dg on dg.SOCID =a.SOCID and dg.IDNO=b.IDNO and dg.DELREASON='4'" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO c ON a.OCID=c.OCID" & vbCrLf
        sql &= " JOIN dbo.Plan_PlanInfo pp on c.planid = pp.planid and pp.comidno = c.comidno and pp.seqno = c.seqno" & vbCrLf
        sql &= " JOIN dbo.Auth_Relship d ON c.RID=d.RID" & vbCrLf
        sql &= " JOIN dbo.Org_OrgInfo e ON d.OrgID=e.OrgID" & vbCrLf
        sql &= " JOIN dbo.ID_Plan i ON i.PlanID=c.PlanID" & vbCrLf
        sql &= " JOIN dbo.KEY_plan k ON k.TPlanID=i.TPlanID" & vbCrLf
        sql &= " LEFT JOIN dbo.ID_District f ON d.DistID=f.DistID" & vbCrLf
        sql &= " LEFT JOIN dbo.Key_TrainType g ON c.TMID=g.TMID" & vbCrLf
        sql &= " LEFT JOIN dbo.SHARE_CJOB s on s.CJOB_UNKEY=c.CJOB_UNKEY" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_TECHEXAM h ON a.SOCID=h.SOCID" & vbCrLf
        'sql &= " --LEFT JOIN dbo.STUD_GETJOBSTATE3 sg3 ON sg3.CPoint=1 and sg3.SOCID =a.SOCID" & vbCrLf
        'sql &= " --LEFT JOIN dbo.STUD_BLIGATEDATA bli ON bli.SBID=sg3.SBID" & vbCrLf
        sql &= $" WHERE 1=1 {SearchStr3}" & vbCrLf
        If cjobValue.Value <> "" Then
            sql &= " and c.CJOB_UNKEY=" & cjobValue.Value & vbCrLf
        End If
        'sql &= " AND ROWNUM <=10" & vbCrLf
        Dim dt5 As DataTable = Nothing
        Try
            dt5 = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try
        '將搜尋資料加入dt
        Call SUtl_AddDt3(dt, dt5, RecordCountInt)

        Dim flagGs3a As Boolean = False '是否進入查詢 Y@true N@false 學員身障參訓歷史
        Dim flagGs3b As Boolean = False '查詢正常 true:正常 / false:異常
        Dim oDs As DataSet = Nothing
        If IDNO.Text <> "" OrElse Name.Text <> "" Then
            flagGs3a = True
            Try
                '身障webService.C120486938
                '學員身障參訓歷史(IDNO,生日yyyy/MM/dd,CNAME)
                'http://163.29.199.211/TIMSWS/GetStudents.asmx
                'Dim wsGs As New GetStudents.GetStudents
                'oDs = wsGs.GetStudentData3(IDNO.Text, "", Name.Text)
                flagGs3b = True
            Catch ex As Exception
                flagGs3b = False
            End Try
        Else
            '承訓單位於上方路徑欲利用「查詢歷史紀錄」查詢時，無法顯示「身心障礙者職業重建服務窗口計畫」歷史紀錄
            '，雖於報名完成後可至「首頁>>學員動態管理>>教務管理>>學員參訓歷史」查詢
            '，但有資訊不一致性的問題，故建議在「查詢歷史紀錄」中也連結「身心障礙者職業重建服務窗口計畫」
            '，讓承訓單位更方便查詢。 by AMU 201511103
            'SD_01_001 '首頁>>學員動態管理>>招生作業>>報名登錄
            'http://163.29.199.211/TIMSWS/GetStudents.asmx/GetStudentData3a?vIDNOArray=R120590402,C120486938,F122394972,Y120317346,H290115281,K121341300
            Dim IDNOStr As String = ""
            If Session("IDNOArray") IsNot Nothing Then
                Dim IDNOArray As ArrayList = Session("IDNOArray")
                For i As Integer = 0 To IDNOArray.Count - 1
                    IDNOStr &= $"{If(IDNOStr <> "", ",", "")}'{TIMS.ClearSQM(TIMS.ChangeIDNO(IDNOArray(i)))}'"
                Next
            End If

            If IDNOStr <> "" Then
                flagGs3a = True
                Try
                    '身障webService.C120486938
                    '學員身障參訓歷史(IDNO,生日yyyy/MM/dd,CNAME)
                    'http://163.29.199.211/TIMSWS/GetStudents.asmx
                    'Dim wsGs As New GetStudents.GetStudents
                    'oDs = wsGs.GetStudentData3a(IDNOStr)
                    flagGs3b = True
                Catch ex As Exception
                    flagGs3b = False
                End Try
            End If
        End If

        '查詢正常 true:正常
        If flagGs3b Then
            Dim flag_can_show As Boolean = False
            If oDs IsNot Nothing Then
                If oDs.Tables.Count > 0 Then flag_can_show = True
            End If
            '顯示查詢資訊
            If flag_can_show Then
                Dim dt3B As DataTable = Nothing
                dt3B = oDs.Tables(0)
                For Each dr3 As DataRow In dt3B.Rows 'For Each dr3 In dt3.Rows
                    If RecordCountInt > 0 Then RecordCountInt -= 1
                    If Not (RecordCountInt > 0) Then Exit For '超過 最大筆數限制

                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("VSSORT") = If(Session("IDNOArray") Is Nothing, 99, TIMS.Get_VSSORT(TIMS.ChangeIDNO(dr3("IDNO")), Session("IDNOArray")))
                    dr("IDNO") = TIMS.ChangeIDNO(dr3("IDNO")) '1.身分證號
                    dr("Name") = dr3("CNAME") '1.姓名
                    dr("Sex") = dr3("SEX") '性別
                    dr("Birthday") = dr3("Birthday") '2.出生年月日
                    dr("DistName") = dr3("DISTNAME") '2.分署
                    dr("Years") = dr3("YEARS") '3.訓練年度
                    dr("PlanName") = "<FONT color='Red'>" & dr3("PLANNAME") & "</FONT>" '4.訓練計畫
                    If Not flag_NoOrgName Then dr("OrgName") = dr3("ORGNAME") '5.訓練機構
                    dr("TMID") = dr3("TRAINNAME") '6.訓練職類
                    'dr("CJOB_NAME") = dr3("CJOB_NAME") '7.通俗職類
                    dr("ClassName") = dr3("CLASSCNAME") '8.班別名稱

                    'THours: '9.受訓時數
                    'TRound: '10.受訓期間
                    dr("THours") = dr3("THOURS")
                    'dr("TRound") = Common.FormatDate(dr3("STDate")) & "<BR>|<BR>" & Common.FormatDate(dr3("FTDATE"))

                    Dim strSTDate As String = If(flag_Roc, TIMS.Cdate17(dr3("STDate")), TIMS.Cdate3(dr3("STDate")))
                    Dim strFTDate As String = If(flag_Roc, TIMS.Cdate17(dr3("FTDate")), TIMS.Cdate3(dr3("FTDate")))
                    dr("TRound") = strSTDate & "<BR>|<BR>" & strFTDate

                    'dr("SkillName") = dr3("ExamName") '11.技能檢定
                    'dr("WEEKS") = dr3("WEEKS")  '12.上課時間
                    '13.訓練狀態
                    dr("TFlag") = $"{dr3("StudStatus")}"
                    '補離退資訊
                    'dr("JobStatus") = dr3("JobStatus") '15.訓後就業狀況
                    ''參訓身分
                    'If Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'").Length > 0 Then
                    '    dr("Ident") = Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'")(0)("Name")
                    'Else
                    '    dr("Ident") = "無身分別"
                    'End If
                    '電話1
                    dr("Tel") = dr3("Tel").ToString
                    '地址。
                    dr("Address") = dr3("Address").ToString
                Next
            End If
        End If


        Dim dt4B As DataTable = Nothing
        Dim flagGs4 As Boolean = True '查詢正常 true:正常 / false:異常
        If IDNO.Text <> "" Then
            IDNO.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(IDNO.Text))
            Try
                dt4B = TIMS.GetTrainingListS(IDNO.Text)
            Catch ex As Exception
                flagGs4 = False 'false:異常
            End Try
        Else
            Dim IDNOStr As String = "" '多筆
            If Session("IDNOArray") IsNot Nothing Then
                Dim IDNOArray As ArrayList = Session("IDNOArray")
                For i As Integer = 0 To IDNOArray.Count - 1
                    IDNOStr &= $"{If(IDNOStr <> "", ",", "")}{TIMS.ClearSQM(TIMS.ChangeIDNO(IDNOArray(i)))}"
                Next
            End If
            If IDNOStr <> "" Then
                Try
                    dt4B = TIMS.GetTrainingListS(IDNOStr)
                Catch ex As Exception
                    flagGs4 = False 'false:異常
                End Try
            End If
        End If

        If dt4B Is Nothing Then flagGs4 = False 'false:異常
        If dt4B IsNot Nothing Then
            If dt4B.Rows.Count = 0 Then flagGs4 = False 'false:異常
        End If
        If flagGs4 Then
            For Each dr4 As DataRow In dt4B.Rows 'For Each dr3 In dt3.Rows
                If RecordCountInt > 0 Then RecordCountInt -= 1
                If Not (RecordCountInt > 0) Then Exit For '超過 最大筆數限制

                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("VSSORT") = If(Session("IDNOArray") Is Nothing, 99, TIMS.Get_VSSORT(TIMS.ChangeIDNO(dr4("IDNO")), Session("IDNOArray")))
                dr("IDNO") = TIMS.ChangeIDNO(dr4("IDNO")) '1.身分證號
                dr("Name") = $"{dr4("NAME")}" '1.姓名
                dr("Sex") = " " 'dr3("SEX") '性別
                dr("Birthday") = $"{dr4("Birthday")}" '2.出生年月日
                dr("DistName") = $"{dr4("DISTNAME")}" '2.分署
                dr("Years") = $"{dr4("YEARS")}" '3.訓練年度
                dr("PlanName") = $"<FONT color='Red'>{dr4("PLANNAME")}</FONT>" '4.訓練計畫
                'dr("OrgName") = $"{dr4("ORGNAME")) '5.訓練機構
                If Not flag_NoOrgName Then dr("OrgName") = $"{dr4("ORGNAME")}" '5.訓練機構
                dr("TMID") = $"{dr4("TRAINNAME")}" '6.訓練職類
                dr("CJOB_NAME") = $"{dr4("CJOB_NAME")}" '7.通俗職類
                dr("ClassName") = $"{dr4("CLASSCNAME")}" '8.班別名稱

                'THours: '9.受訓時數
                'TRound: '10.受訓期間
                dr("THours") = $"{dr4("THOURS")}"
                'dr("TRound") = Common.FormatDate(dr3("STDate")) & "<BR>|<BR>" & Common.FormatDate(dr3("FTDATE"))
                Dim strSTDate As String = $"{dr4("TRound")}".Split("-")(0) 'If(flag_Roc, TIMS.cdate17(dr3("STDate")), Common.FormatDate(dr3("STDate")))
                Dim strFTDate As String = $"{dr4("TRound")}".Split("-")(1) 'If(flag_Roc, TIMS.cdate17(dr3("FTDate")), Common.FormatDate(dr3("FTDate")))
                If (flag_Roc) Then strSTDate = TIMS.Cdate17(CDate(strSTDate))
                If (flag_Roc) Then strFTDate = TIMS.Cdate17(CDate(strFTDate))
                dr("TRound") = $"{strSTDate}<BR>|<BR>{strFTDate}"
                'dr("SkillName") = dr3("ExamName") '11.技能檢定
                'dr("WEEKS") = dr3("WEEKS")  '12.上課時間
                '13.訓練狀態
                dr("TFlag") = TIMS.CHG_TFLAG($"{dr4("TFlag")}") 'dr4("TFlag")) 'dr4("StudStatus"))

                '補離退資訊
                'dr("JobStatus") = dr3("JobStatus") '15.訓後就業狀況
                ''參訓身分
                'If Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'").Length > 0 Then
                '    dr("Ident") = Key_Identity.Select("IdentityID='" & dr3("MIdentityID") & "'")(0)("Name")
                'Else
                '    dr("Ident") = "無身分別"
                'End If
                '電話1
                dr("Tel") = " " 'dr3("Tel").ToString
                '地址。
                dr("Address") = " " 'dr3("Address").ToString
                dr("WEEKS") = $"{dr4("WEEKS")}" '12.上課時間
                Dim T_WORKSUPPIDENT As String = If($"{dr4("WORKSUPPIDENT")}" = "Y", "在職者", "")
                dr("MEMO1") &= If((T_WORKSUPPIDENT <> $"{dr4("MEMO1")}"), $"{T_WORKSUPPIDENT}{dr4("MEMO1")}", $"{T_WORKSUPPIDENT}")
            Next
        End If

        If dt.Rows.Count = 0 Then
            msg.Text = "查無資料!"
            If rqSD_01_004_Type <> "" Then '如果是由SD_01_004E網審核報名功能的學員參訓歷史查詢
                SearchTable.Style.Item("display") = "none"  '查詢條件TABLE
                'Table2.Visible = False         '功能路徑顯示列
                'Common.MessageBox(Me, "查無資料!")

                If $"{rqBtnHistory}" <> "" Then
                    '.click();window.close();
                    Page.RegisterStartupScript("ThenHistory", "<script>opener.document.getElementById('" & rqBtnHistory & "').click();window.close();</script>")
                Else
                    Page.RegisterStartupScript("History", "<script>window.close();</script>")
                End If
            Else
                SearchTable.Style.Item("display") = cst_inline '查詢條件TABLE
                'Table2.Visible = True           '功能路徑顯示列
            End If
            ShowDataTable.Style.Item("display") = "none"     '查詢結果TABLE
        Else
            '有查詢，但異常狀況。
            If flagGs3a AndAlso Not flagGs3b Then msg.Text = "學員身障參訓歷史查詢失敗!!"
            RecordCount.Text = dt.Rows.Count
            SearchTable.Style.Item("display") = "none"
            ShowDataTable.Style.Item("display") = cst_inline
            If rqSD_01_004_Type <> "" Then '如果是由SD_01_004E網審核報名功能的學員參訓歷史查詢
                Table3.Visible = False         '說明列
                Button5.Visible = False        '回上一頁
                'Table2.Visible = False         '功能路徑顯示列
                Lab_TR.Visible = False         '說明列
            Else
                Table3.Visible = True         '說明列
                Button5.Visible = True        '回上一頁
                'Table2.Visible = True         '功能路徑顯示列
                Lab_TR.Visible = True         '說明列
            End If

            If Session("IDNOArray") IsNot Nothing Then
                dt.DefaultView.Sort = "VSSORT,IDNO,Birthday,TRound"
            Else
                dt.DefaultView.Sort = "IDNO,Birthday,TRound"
            End If
            PageControler1.PageDataTable = dt.DefaultView.Table

            If Session("IDNOArray") IsNot Nothing Then
                PageControler1.Sort = "VSSORT,IDNO,Birthday,TRound"
            Else
                PageControler1.Sort = "IDNO,Birthday,TRound"
            End If
            PageControler1.ControlerLoad()

            'dt.Dispose()
            'dt = Nothing
        End If
    End Sub

    '查詢(條件)
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub '異常離開

        Select Case $"{sm.UserInfo.LID}"
            Case "0" '署
            Case "1" '分署
            Case Else '其他委訓單位
                flag_NoOrgName = True
        End Select

        'rqSD_01_004_Type = TIMS.ClearSQM(Request("SD_01_004_Type"))
        'rqIDNO = TIMS.ClearSQM(Request("IDNO"))
        If rqSD_01_004_Type = "Student" Then
            '由SD_01_004E網審核功能的按近二年參訓歷史鍵代入的依學員
            IDNO.Text = rqIDNO 'Request("IDNO")
            '--安全性預防--
            IDNO.Text = TIMS.ClearSQM(IDNO.Text)
            If IDNO.Text.Length < 8 OrElse IDNO.Text.Length > 12 Then
                '字長只能介於8~12
                Exit Sub '異常離開
            End If
        End If

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid2)

        Dim SearchStr1 As String = ""
        Dim SearchStr2 As String = ""
        Dim SearchStr3 As String = ""

        DataGrid2.CurrentPageIndex = 0

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If Name.Text <> "" Then Name.Text = Trim(Name.Text)
        Name.Text = TIMS.ClearSQM(Name.Text)

        If (OCIDValue.Value <> "" AndAlso IDNO.Text = "" AndAlso Name.Text = "") OrElse $"{rqSD_01_004_Type}" = CST_KD_STUDENTLIST Then
            'Request("Type") = "StudentList" 由報名登錄功能的按參訓歷史查詢鍵,依所查詢出來的資料
            Dim sql As String
            Dim dt As New DataTable
            Dim dr As DataRow
            Dim IDNOStr As String = ""   '身分證字串

            '本機 且有班級ocid
            If rqSD_01_004_Type <> CST_KD_STUDENTLIST AndAlso OCIDValue.Value <> "" Then
                'TIMS.OpenDbConn(objconn)
                sql = " SELECT b.IDNO FROM CLASS_STUDENTSOFCLASS a JOIN STUD_STUDENTINFO b On a.SID=b.SID WHERE a.OCID=@OCID"
                Dim cmd As New SqlCommand(sql, objconn)
                With cmd
                    .Parameters.Clear()
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue.Value
                    dt = New DataTable
                    dt.Load(.ExecuteReader())
                End With
                IDNOStr = ""
                For Each dr In dt.Rows
                    If IDNOStr.IndexOf(TIMS.ChangeIDNO(dr("IDNO"))) = -1 Then
                        IDNOStr &= $"{If(IDNOStr <> "", ",", "")}'{TIMS.ClearSQM(TIMS.ChangeIDNO(dr("IDNO")))}'"
                    End If
                Next
                'If IDNOStr <> "" Then IDNOStr = UCase(IDNOStr)
            Else
                '由SD_01_004E網審核功能的按參訓歷史查詢鍵(依所查詢出來的資料) 'SD_01_001 '首頁>>學員動態管理>>招生作業>>報名登錄
                IDNOStr = ""
                If Session("IDNOArray") IsNot Nothing Then
                    Dim IDNOArray As ArrayList = Session("IDNOArray")
                    For i As Integer = 0 To IDNOArray.Count - 1
                        IDNOStr &= $"{If(IDNOStr <> "", ",", "")}'{TIMS.ClearSQM(TIMS.ChangeIDNO(IDNOArray(i)))}'"
                    Next
                End If
            End If

            If IDNOStr <> "" Then
                SearchStr1 = $" and SID IN ({IDNOStr})"
                SearchStr2 = $" and IDNO IN ({IDNOStr})"
                SearchStr3 = $" and b.IDNO IN ({IDNOStr})"
            Else
                '不查資料
                SearchStr1 += " and 1<>1"
                SearchStr2 += " and 1<>1"
                SearchStr3 += " and 1<>1"
            End If
        Else
            '一般狀況查詢
            'IDNO.Text = TIMS.ClearSQM(IDNO.Text)
            If IDNO.Text <> "" Then
                SearchStr1 = " and SID='" & IDNO.Text & "'"
                SearchStr2 = " and IDNO='" & IDNO.Text & "'"
                SearchStr3 = " and b.IDNO='" & IDNO.Text & "'"
            End If
            If Name.Text <> "" Then
                SearchStr1 += " and Name like '%" & Name.Text & "%'"
                SearchStr2 += " and Name like '%" & Name.Text & "%'"
                SearchStr3 += " and b.Name like '%" & Name.Text & "%'"
            End If
            If DistID.SelectedIndex <> 0 AndAlso DistID.SelectedValue <> "" Then
                SearchStr1 += " and DistID='" & DistID.SelectedValue & "'"
                SearchStr2 += " and DistID='" & DistID.SelectedValue & "'"
                SearchStr3 += " and d.DistID='" & DistID.SelectedValue & "'"
            End If

            If TPlanID.SelectedIndex <> 0 AndAlso TPlanID.SelectedValue <> "" Then
                SearchStr1 += " and TPlanID='" & TPlanID.SelectedValue & "'"
                SearchStr2 += " and TPlanID='" & TPlanID.SelectedValue & "'"
                SearchStr3 += " and i.TPlanID='" & TPlanID.SelectedValue & "'"
                TPlanName.Text = "訓練計畫:" & TPlanID.SelectedItem.Text
            Else
                TPlanName.Text = "訓練計畫:不區分"
            End If
            If center.Text <> "" Then
                SearchStr1 += " and TrinUnit like '%" & center.Text & "%'"
                SearchStr2 += " and TrinUnit like '%" & center.Text & "%'"
                If RIDValue.Value = "" Then
                    SearchStr3 += " and e.OrgName ='" & center.Text & "'"
                Else
                    SearchStr3 += " and d.RID ='" & RIDValue.Value & "'"
                End If
            End If
            If STDate.Text <> "" Then
                SearchStr1 += " and SDate >= " & TIMS.To_date(STDate.Text)
                SearchStr2 += " and SDate >= " & TIMS.To_date(STDate.Text)
                SearchStr3 += " and c.STDate >= " & TIMS.To_date(STDate.Text)
            End If
            If FTDate.Text <> "" Then
                SearchStr1 += " and SDate <= " & TIMS.To_date(FTDate.Text)
                SearchStr2 += " and SDate <= " & TIMS.To_date(FTDate.Text)
                SearchStr3 += " and c.STDate <= " & TIMS.To_date(FTDate.Text)
            End If
        End If

        'If trainValue.Value <> "" Then
        '    SearchStr2 += " and TMID='" & trainValue.Value & "'"
        '    SearchStr3 += " and c.TMID='" & trainValue.Value & "'"
        'End If

        '--------------start報名登錄,e網審核,功能的近兩年參訓資料查詢-------------------------------------------
        If rqTwoYears IsNot Nothing Then
            rqTwoYears = TIMS.ClearSQM(rqTwoYears)
            If rqTwoYears <> "" AndAlso Len(rqTwoYears) = 4 Then
                SearchStr1 += " and DATEPART(YEAR, SDate)>='" & TIMS.CINT1(rqTwoYears) & "'"
                SearchStr2 += " and DATEPART(YEAR, SDate)>='" & TIMS.CINT1(rqTwoYears) & "'"
                SearchStr3 += " and DATEPART(YEAR, c.STDate) >='" & TIMS.CINT1(rqTwoYears) & "'"
            Else
                '有輸入值但不符合4碼期待，不提供查詢資料
                SearchStr1 += " AND 1<>1"
                SearchStr2 += " AND 1<>1"
                SearchStr3 += " AND 1<>1"
            End If
        End If
        '-------------end報名登錄,e網審核,功能的近兩年參訓資料查詢----------------------------------------------

        If SearchStr1 = "" OrElse SearchStr2 = "" OrElse SearchStr3 = "" Then
            Common.MessageBox(Me, "請輸入搜尋條件!!")
            Exit Sub
        End If

        Dim noOk1 As Boolean = True '沒身分證號 b.IDNO
        Dim noOk2 As Boolean = True '沒姓名 b.Name
        Dim noOk3 As Boolean = True '輸入資料過少。
        Dim noOk4 As Boolean = True '沒開訓日 c.STDate >= <=
        If SearchStr3 <> "" Then
            If SearchStr3.IndexOf("b.IDNO") > -1 Then noOk1 = False '可以搜尋了。

            If SearchStr3.IndexOf("b.Name") > -1 Then noOk2 = False '可以搜尋了。

            If SearchStr3.IndexOf(">=") > -1 AndAlso SearchStr3.IndexOf("<=") > -1 Then noOk4 = False '可以搜尋了。
        End If
        If SearchStr1 <> "" Then
            Dim ss3 As String() = SearchStr1.Split("and")

            '若超過2項。
            If ss3.Length > 2 Then noOk3 = False '可以搜尋了。

            If noOk1 AndAlso noOk2 AndAlso noOk3 AndAlso noOk4 Then
                Common.MessageBox(Me, "請輸入詳細搜尋條件!!")
                Exit Sub
            End If
        End If

        Call GetStudentData(SearchStr1, SearchStr2, SearchStr3)
    End Sub

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '查詢(條件)
        Call Search1()
    End Sub

    '依身分證號 設定
    Sub SEARCH3(ByVal IDNOStr As String)
        Dim SearchStr1 As String = ""
        Dim SearchStr2 As String = ""
        Dim SearchStr3 As String = ""
        If IDNOStr <> "" Then
            SearchStr1 = $" and SID IN ({IDNOStr})"
            SearchStr2 = $" and IDNO IN ({IDNOStr})"
            SearchStr3 = $" and b.IDNO IN ({IDNOStr})"
        End If
        Call GetStudentData(SearchStr1, SearchStr2, SearchStr3)
    End Sub

    '班別查詢鈕
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OCID.Items.Clear()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        If RIDValue.Value = "" OrElse PlanID.Value = "" Then
            OCID.Items.Insert(0, New ListItem("此計畫、機構底下沒有任何班級", ""))
            Return
        End If
        Dim vTPlanID As String = TIMS.GetListValue(TPlanID)
        Dim vDistID As String = TIMS.GetListValue(DistID)
        Dim parms As New Hashtable From {
            {"RID", RIDValue.Value},
            {"TPlanID", vTPlanID},
            {"DistID", vDistID},
            {"PlanID", PlanID.Value}
        }
        Dim sql As String = ""
        sql &= " SELECT OCID,CLASSCNAME2 FROM VIEW2 "
        sql &= " WHERE RID=@RID AND TPlanID=@TPlanID AND DistID=@DistID AND PlanID=@PlanID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then
            OCID.Items.Insert(0, New ListItem("此計畫、機構底下沒有任何班級", ""))
            Return
        End If

        For Each dr As DataRow In dt.Rows
            OCID.Items.Add(New ListItem($"{dr("CLASSCNAME2")}", $"{dr("OCID")}"))
        Next
        OCID.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '回上一頁(下邊)鈕
    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        SearchTable.Style.Item("display") = cst_inline
        ShowDataTable.Style.Item("display") = "none"
    End Sub

    '回上一頁(上邊)鈕
    Private Sub Button5_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.ServerClick
        SearchTable.Style.Item("display") = cst_inline
        ShowDataTable.Style.Item("display") = "none"
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If ViewState("sort") IsNot Nothing Then
                    Dim i As Integer = -1
                    Select Case ViewState("sort")
                        Case "Name", "Name DESC"
                            i = 1
                        Case "IDNO", "IDNO DESC"
                            i = 2
                        Case "Birthday", "Birthday DESC"
                            i = 3
                        Case "DistName", "DistName DESC"
                            i = 4
                        Case "OrgName", "OrgName DESC"
                            i = 7
                        Case "TMID", "TMID DESC"
                            i = 8
                        Case "ClassName", "ClassName DESC"
                            i = 10
                        Case "TRound", "TRound DESC"
                            i = 12
                    End Select

                    Dim MyImage As New Web.UI.WebControls.Image
                    MyImage.ImageUrl = If(ViewState("sort").ToString.IndexOf(" DESC") = -1, "../../images/SortUp.gif", "../../images/SortDown.gif")
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(MyImage)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyTable As HtmlTable = e.Item.FindControl("Table4")

                Dim LName As Label = e.Item.FindControl("LName")
                Dim LIDNO As Label = e.Item.FindControl("LIDNO")
                Dim LBirthday As Label = e.Item.FindControl("LBirthday")
                Dim LSex As Label = e.Item.FindControl("LSex")
                Dim LIdent As Label = e.Item.FindControl("LIdent")
                Dim LTel As Label = e.Item.FindControl("LTel")
                Dim LAddress As Label = e.Item.FindControl("LAddress")

                If but_S_Export Then MyTable.Visible = False

                If Not but_S_Export Then
                    Select Case sm.UserInfo.LID
                        Case 0, 1
                            e.Item.Cells(1).Style("CURSOR") = "hand"
                            e.Item.Cells(1).Attributes("onmouseover") = "ShowPersonData('" & MyTable.ClientID & "');"
                            e.Item.Cells(1).Attributes("onmouseout") = "HidPersonData('" & MyTable.ClientID & "');"
                    End Select
                End If

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                LName.Text = drv("Name").ToString
                LIDNO.Text = TIMS.ChangeIDNO(drv("IDNO").ToString)
                LBirthday.Text = If(flag_Roc, TIMS.Cdate17(drv("Birthday")), TIMS.Cdate3(drv("Birthday")))
                LSex.Text = If($"{drv("Sex")}" = "M", "男", If($"{drv("Sex")}" = "F", "女", $"{drv("Sex")}"))
                LIdent.Text = $"{drv("Ident")}"
                LTel.Text = $"{drv("Tel")}"
                LAddress.Text = $"{drv("Address")}"

                e.Item.Cells(cst_dg2_出生日期).Text = LBirthday.Text 'TIMS.cdate18(drv("Birthday"))
                If flag_Roc AndAlso $"{drv("Years")}" <> "" Then
                    'e.Item.Cells(cst_dg2_出生日期).Text = "Then" & TIMS.cdate17(drv("Birthday"))
                    e.Item.Cells(cst_dg2_年度).Text = (CInt(drv("Years")) - 1911).ToString()
                End If

        End Select

        If rqSD_01_004_Type = "Student" Then '如果是由SD_01_004E網審核報名功能的學員參訓歷史查詢(單一查詢時)
            e.Item.Cells(1).Style.Item("display") = "none"
            e.Item.Cells(2).Style.Item("display") = "none"
            e.Item.Cells(3).Style.Item("display") = "none"
            '.Style.Item("display") = "none"
        Else
            e.Item.Cells(1).Style.Item("display") = cst_inline
            e.Item.Cells(2).Style.Item("display") = cst_inline
            e.Item.Cells(3).Style.Item("display") = cst_inline
        End If
    End Sub

    Private Sub DataGrid2_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid2.SortCommand
        ViewState("sort") = If(ViewState("sort") = e.SortExpression, String.Concat(e.SortExpression, " DESC"), e.SortExpression)
        PageControler1.Sort = ViewState("sort")
        'PageControler1.ChangeSort(ViewState("sort"))
        'Dim CPdt As DataTable
        'CPdt = dt.Copy()
        'PageControler1.DataTableCreate(CPdt, PageControler1.Sort, PageControler1.PageIndex)
        PageControler1.DataTableCreate(CPdt, PageControler1.Sort)
    End Sub

    '查詢匯入資料鈕
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Const Cst_Upload_Path As String = "~/SD/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_Upload_Path)

        Const Cst_MyFileType As String = "csv"
        'Dim MyFile As System.IO.File
        'Dim FileOCIDValue, MyFileName As String
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        Const cst_flag As String = ","

        '檢查檔案格式與大小----------   Start
        If File1.Value = "" Then
            Common.MessageBox(Me, "未選擇檔案!")
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        'FileOCIDValue = Split(Split(MyFileName, "-")(1), ".")(0)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> Cst_MyFileType Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為" & Cst_MyFileType & "檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小----------   End

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Cst_Upload_Path}{MyFileName}")
        File1.PostedFile.SaveAs(filePath1) '上傳檔案
        'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)

        '將檔案讀出放入記憶體
        Dim sr As System.IO.Stream
        Dim srr As System.IO.StreamReader
        sr = IO.File.OpenRead(filePath1)
        srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

        Dim OneRow As String            'srr.ReadLine 一行一行的資料
        Dim RowIndex As Integer = 0     '讀取行累計數
        Dim Reason As String = ""       '儲存錯誤的原因
        Dim colArray As Array

        Dim dtWrong As New DataTable    '儲存錯誤資料的DataTable
        'Dim drWrong As DataRow
        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        ViewState("searchIDNO") = ""
        Do While srr.Peek >= 0
            OneRow = srr.ReadLine
            If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈

            If RowIndex > -1 Then
                Reason = ""
                colArray = Split(OneRow, cst_flag)
                If colArray.Length = 0 Then Exit Do
                colArray(0) = TIMS.ChangeIDNO($"{colArray(0)}")

                'If Not TIMS.CheckIDNO($"{colArray(0))) Then
                '    Reason += $"{colArray(0)) & "<--身分證號有誤"
                'Else
                '    If ViewState("searchIDNO").ToString.IndexOf($"{colArray(0))) = -1 Then
                '        If ViewState("searchIDNO") = "" Then
                '            ViewState("searchIDNO") += "'" & $"{colArray(0)) & "'"
                '        Else
                '            ViewState("searchIDNO") += ",'" & $"{colArray(0)) & "'"
                '        End If
                '    End If
                'End If

                If ViewState("searchIDNO").ToString.IndexOf($"{colArray(0)}") = -1 Then
                    If ViewState("searchIDNO") <> "" Then ViewState("searchIDNO") += ","
                    ViewState("searchIDNO") += $"'{colArray(0)}'"
                End If

                'If Reason <> "" Then
                '    '錯誤資料，填入錯誤資料表
                '    drWrong = dtWrong.NewRow
                '    dtWrong.Rows.Add(drWrong)
                '    drWrong("Index") = RowIndex
                '    drWrong("Reason") = Reason
                'End If
            End If
            RowIndex += 1 '讀取行累計數
        Loop

        sr.Close()
        srr.Close()
        'IO.File.Delete(Server.MapPath(Cst_MyFileType & MyFileName))
        '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
        TIMS.MyFileDelete(filePath1)

        Call SEARCH3(ViewState("searchIDNO"))

        'If dtWrong.Rows.Count = 0 Then
        '    SEARCH3(ViewState("searchIDNO"))
        'Else
        '    'SEARCH3(ViewState("searchIDNO"))
        '    Reason = "資料匯入有誤!!" & vbCrLf
        '    For Each dr As DataRow In dtWrong.Rows
        '        Reason += dr("Reason").ToString & vbCrLf
        '    Next
        '    Common.MessageBox(Me, Reason)
        'End If
    End Sub

    '關閉鈕
    Private Sub Btnclose_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btnclose.ServerClick
        If $"{rqBtnHistory}" <> "" Then
            '對父層 按鈕 呼叫
            Page.RegisterStartupScript("ThenHistory", "<script>If(opener.document.getElementById('" & rqBtnHistory & "')) opener.document.getElementById('" & rqBtnHistory & "').click();window.close();</script>")
        Else
            Page.RegisterStartupScript("History", "<script>window.close();</script>")
        End If
    End Sub

    '列印鈕
    Private Sub But_P_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_P.Click
        DataGrid2.AllowPaging = False
        'Button1_Click(sender, e)
        Call Search1() '查詢(條件)
        PageControler1.Visible = False
        but_P.Visible = False
        but_S.Visible = False
        Btnclose.Visible = False
        RegisterStartupScript("scripprint", "<script>printDoc();document.getElementById('but_f').click();</script>")
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

#Region "NO USE"
    ''列印BUG問題解決
    'Public Overloads Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
    '    '此段為必要
    'End Sub
#End Region

    '匯出鈕
    Private Sub But_S_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_S.Click
        'Dim fileName As String = "student_old.xls"
        but_S_Export = True '使用匯出鈕功能
        DataGrid2.AllowPaging = False
        'Button1_Click(sender, e)
        Call Search1() '查詢(條件)
        'Response.Clear()
        'Response.Buffer = True
        'Response.Charset = "UTF-8" '設定字集
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        'Response.AppendHeader("Content-Disposition", "attachment;filename=assets" & ".xls")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        'Dim sFileName1 As String = ""

        DataGrid2.EnableViewState = False  '把ViewState給關了
        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        DataGrid2.RenderControl(objHtmlTextWriter)
        Dim strHTML As String = ""
        strHTML &= TIMS.sUtl_AntiXss($"{objStringWriter}")

        'Dim sFileName1 As String = ""
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        'Dim strHTML As String = ""
        'strHTML &= ("<div>")
        'strHTML &= ("<table>")
        'strHTML &= ("</table>")
        'strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        'parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        DataGrid2.AllowPaging = True
        TIMS.Utl_RespWriteEnd(Me, objconn, "")  'Response.End()
    End Sub

    '列印後動作(前端重新呼叫)
    Private Sub But_f_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_f.Click
        DataGrid2.AllowPaging = True
        'Button1_Click(sender, e)
        Call Search1() '查詢(條件)
        PageControler1.Visible = True
        but_P.Visible = True
        but_S.Visible = True
        Btnclose.Visible = True
    End Sub

#Region "NO USE"
    ''委訓機構只能查詢有參加過該機構的學員 'Trun:可查詢 False:該學員非該機構的學員
    'Function CheckLID2(ByVal OrgID As String, ByVal vIDNO As String) As Boolean
    '    If vIDNO <> "" Then vIDNO = Trim(vIDNO)
    '    If vIDNO <> "" Then vIDNO = UCase(vIDNO)
    '    If vIDNO <> "" Then vIDNO = TIMS.ChangeIDNO(vIDNO)

    '    Dim rst As Boolean = False
    '    Dim Sql As String = ""
    '    Sql = "" & vbCrLf
    '    sql &= " select 'x' x " & vbCrLf
    '    sql &= " from stud_studentinfo ss   " & vbCrLf
    '    sql &= " join class_studentsofclass cs   on cs.sid =ss.sid" & vbCrLf
    '    sql &= " join class_classinfo cc   on cc.ocid =cs.ocid" & vbCrLf
    '    sql &= " join auth_relship ar   on ar.RID =cc.RID" & vbCrLf
    '    sql &= " where ss.idno ='" & vIDNO & "'" & vbCrLf
    '    sql &= " AND ar.OrgID  ='" & OrgID & "'" & vbCrLf
    '    Dim dt As DataTable = DbAccess.GetDataTable(Sql, objconn)
    '    If dt.Rows.Count > 0 Then
    '        rst = True
    '    End If
    '    Return rst
    'End Function
#End Region

End Class