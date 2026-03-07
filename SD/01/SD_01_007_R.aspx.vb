Partial Class SD_01_007_R
    Inherits AuthBasePage

    'iReport
    'SD_01_007_R*.jrxml
    '產投28 (2014,2016使用中) (LpfBatchGet28e.vbproj--STUD_BLIGATEDATA28E)
    'SD_01_007_R2_b '署(局)、分署(中心) --Stud_BligateData28
    'SD_01_007_R2b_b '委訓單位 'Member --Stud_BligateData28

    'select socid,idno,count(1) cnt  from Stud_BligateData28 group by socid,idno having count(1) >1
    'select socid,count(1) cnt  from Stud_BligateData28 group by socid having count(1) >1
    '在職進修、接受企業委託 (lpfBatchGet02.vbproj--STUD_SELRESULTBLI)
    '06	在職進修訓練    'XXX 07	接受企業委託訓練
    '66	主題產業職業訓練(在職)
    'SD_01_007_R3_b '署(局)、分署(中心) --STUD_SELRESULTBLI, Stud_BligateData28,Stud_BligateData28e
    'SD_01_007_R3b_b '委訓單位 'Member --Stud_BligateData28

    Dim sMemo As String = "" '(查詢原因)
    '3.產業人才投資方案 [28] '4:充電起飛計畫（非在職／是產投）[54]--參訓學員投保狀況檢核表 'SD_01_007_R2*_b.jrxml
    Const cst_printFN2b As String = "SD_01_007_R2_b" 'STUD_BLIGATEDATA28[署(局)、分署(中心)]
    Const cst_printFN2bb As String = "SD_01_007_R2b_b" 'STUD_BLIGATEDATA28[委訓單位]
    '06:在職進修訓練 66:主題產業職業訓練(在職) --參訓學員投保狀況檢核表 SD_01_007_R3_b*.jrxml
    Const cst_printFN3b As String = "SD_01_007_R3_b" '開訓日 STUD_SELRESULTBLI
    Const cst_printFN3bb As String = "SD_01_007_R3_bb" '開訓日 STUD_SELRESULTBLI
    '06:在職進修訓練 66:主題產業職業訓練(在職) --甄試學員投保狀況檢核表 --LpfBatchGet28es.vbproj SD_01_007_R4_b*.jrxml
    Const cst_printFN4b As String = "SD_01_007_R4_b" '甄試日 STUD_BLIGATEDATA28E '署(局)、分署(中心)
    Const cst_printFN4bb As String = "SD_01_007_R4_bb" '甄試日 STUD_BLIGATEDATA28E '委訓單位
    '70:區域產業據點職業訓練計畫(在職)
    Const cst_printFN5b As String = "SD_01_007_R5_b" '甄試日 STUD_BLIGATEDATA28E '署(局)、分署(中心) SD_01_007_R5_b*.jrxml
    Const cst_printFN5bb As String = "SD_01_007_R5_bb" '甄試日 STUD_BLIGATEDATA28E '委訓單位

    'https://jira.turbotech.com.tw/browse/TIMSB-1081
    '開頭數字為075、175（裁減續保）、076（職災續保）、09（訓）皆為不予補助對象，並設定阻擋

    ',(vp.Years-1911)+'年度　參訓學員投保狀況檢核表' TopTitle3
    '勞保勾稽 (lpfBatchGet28.vbproj) UPDateBatch2
    'FROM STUD_SELRESULTBLI
    'FROM STUD_BLIGATEDATA28E
    'TIMS.Cst_TPlanID06AppPlan2
    'TIMS.Cst_TPlanID28AppPlan

#Region "(No Use)"

    '產投
    'SD_01_007_R2  '署(局)、分署(中心)
    'SD_01_007_R2b '委訓單位 'Member
    '在職進修、接受企業委託
    'SD_01_007_R3  '署(局)、分署(中心)
    'SD_01_007_R3b'委訓單位 'Member
    '/**NEW 2014**/
    'Dim iPYNum14 As Integer=1 'iPYNum14=TIMS.sUtl_GetPYNum14(Me) '若是登入年度為 2014年以後，則傳回2，其餘為1

    '/**NEW 2014**/
    '產投 (2014停用中)
    'SD_01_007_R2  '署(局)、分署(中心)
    'SD_01_007_R2b '委訓單位 'Member
    '在職進修、接受企業委託
    'SD_01_007_R3  '署(局)、分署(中心)
    'SD_01_007_R3b '委訓單位 'Member

    '產投 (2014使用中)
    'SD_01_007_R2_b  '署(局)、分署(中心)
    'SD_01_007_R2b_b '委訓單位 'Member
    '在職進修、接受企業委託
    'SD_01_007_R3_b  '署(局)、分署(中心)
    'SD_01_007_R3b_b '委訓單位 'Member

#End Region

    'Dim blnCanAdds As Boolean=False '新增
    'Dim blnCanMod As Boolean=False  '修改
    'Dim blnCanDel As Boolean=False  '刪除
    'Dim blnCanSech As Boolean=False '查詢
    'Dim blnCanPrnt As Boolean=False '列印

    'ORDER BY  Printtype
    Const cst_Printtype_B1 As String = "B1" '"b.IDNO" 'B1
    Const cst_Printtype_S1 As String = "S1" '"cs.StudentID" 'S1
    Const cst_Printtype3 As String = "c.IDNO" '"b.IDNO" '06 66
    Const cst_Printtype4B As String = "a.IDNO" '06 66 /70

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        ''========== (依照承辦人需求，調整功能路徑名稱，by:20180912)
        'If Not Page.IsPostBack Then
        '    If sm.UserInfo.TPlanID="28" Then  '計畫別為"產業人才投資方案",功能路徑需改變
        '        TitleLab2.Text="首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;參訓學員投保狀況檢核表"
        '    Else
        '        TitleLab2.Text="首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;參訓學員投保狀況檢核表"
        '    End If
        'End If
        ''===============================================================

        'iPYNum14=TIMS.sUtl_GetPYNum14(Me)
        trStExChoice1.Visible = False
        '匯出e網民眾投保狀況檢核表,'甄試學員投保狀況檢核表,'btnExport.Text="" 'DEF:不顯示,'06,66,
        btnExport.Visible = False
        btnExport2.Visible = False
        trPrtSortType1.Visible = False
        trStExChoice1.Visible = False
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trStExChoice1.Visible = True
            btnExport2.Visible = True '匯出投保狀況檢核表
            'btnExport2.Text="匯出甄試學員投保狀況檢核表"
        End If
        '"28"
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trPrtSortType1.Visible = True
            btnExport.Visible = True
            btnExport.Text = "匯出e網民眾投保狀況檢核表" 'DEF:28
        End If

        If Not Page.IsPostBack Then
            Call CCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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
        Print.Attributes("onclick") = "javascript:return CheckPrint();"
    End Sub

#Region "(No Use)"

    'Select Case print_type.SelectedValue
    '    Case "1"
    '        myvalue += "&Printtype=b.IDNO" '依身分證號
    '    Case "2"
    '        myvalue += "&Printtype=cs.StudentID"
    'End Select
    'myvalue += "&OrgName=" & Convert.ToString(center.Text)
    'myvalue += "&ClassCName=" & Convert.ToString(TMID1.Text + OCID1.Text)
    'If cjobValue.Value <> "" Then
    '    myvalue += "&PlanID=" & sm.UserInfo.PlanID
    '    myvalue += "&CJOB_UNKEY=" & Convert.ToString(cjobValue.Value)
    'Else
    '    myvalue += "&PlanID="
    '    myvalue += "&CJOB_UNKEY="
    'End If
    'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_01_007_R", myvalue)
    'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_01_007_R2b", myvalue)

#End Region
    Sub CCreate1()
        '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示
        'rblWorkMode.Enabled=False
        'Common.SetListItem(rblWorkMode, TIMS.cst_wmdip1)
        'TIMS.Tooltip(rblWorkMode, "全計畫-學員處分功能-身分證號-隱碼顯示")

        '取出鍵詞-查詢原因-INQUIRY
        Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        If sm.UserInfo.LID <> "2" Then
            '若只有管理一個班級，自動協助帶出班級--by andy 2009-02-25
            Call TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objconn)
        Else
            'Button12_Click(sender, e)
            Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
        End If
    End Sub

    Function GET_SQL_SD01007R2BB() As String
        Dim RSTSQL As String = "SELECT a.OCID1,ss.NAME ,b.IDNO,cs.STUDENTID ,dbo.FN_GET_MASK1(b.IDNO) IDNO_MK ,a.RELENTERDATE ,c.MDATE
,CASE c.CHANGEMODE WHEN 1 THEN '工作部門或特殊身分異動' WHEN 2 THEN '退保' WHEN 3 THEN '調薪' WHEN 4 THEN '加保' END CHANGEMODE
,c.COMNAME ,c.ACTNO 
,CASE WHEN SUBSTRING(c.ACTNO,1,2)='09' THEN '不符合' WHEN c.CHANGEMODE NOT IN (2) THEN '符合' ELSE '不符合' END 
 +CASE WHEN a.CMASTER1='Y' THEN '(負責人不適用就保)' ELSE '' END 
 +CASE WHEN dbo.FN_GET_BIEF(a.OCID1,b.IDNO)='M' OR c.BIEF='M' THEN '(多元就業計畫進用人員不適用就保)' ELSE '' END CAPMODE
FROM STUD_ENTERTYPE a
JOIN STUD_ENTERTEMP b ON a.SETID=b.SETID 
JOIN STUD_STUDENTINFO ss ON ss.IDNO=b.IDNO
JOIN CLASS_STUDENTSOFCLASS cs ON cs.SID=ss.SID AND cs.OCID=a.OCID1 AND cs.STUDSTATUS NOT IN (2,3) 
LEFT JOIN STUD_BLIGATEDATA28 c ON c.SOCID=cs.SOCID AND c.IDNO=b.IDNO WHERE a.OCID1=@OCID1"
        Return RSTSQL
    End Function

    Function GET_SQL_SD01007R3BB(VRTN As String) As String
        'VRTN = "C"欄位匯出
        If VRTN = "C" Then Return "SETID,ENTERDATE,SERNUM,RELENTERDATE,NAME,IDNO,MDATE,CHANGEMODE,COMNAME,ACTNO,RESOLDER"

        Dim RSTSQL As String = "SELECT a.SETID,a.ENTERDATE,a.SERNUM	,a.OCID1,a.RELENTERDATE	,b.NAME ,b.IDNO
,c.MDATE,c.CHANGEMODE,c.COMNAME ,c.ACTNO,dbo.FN_GET_RESOLDER(b.IDNO,cc.STDATE) RESOLDER
FROM VIEW2 cc
JOIN STUD_ENTERTYPE a ON a.OCID1=cc.OCID
JOIN STUD_ENTERTEMP b ON a.SETID=b.SETID
JOIN STUD_SELRESULT b2 ON a.SETID=b2.SETID AND a.EnterDate=b2.EnterDate AND a.SerNum=b2.SerNum AND a.OCID1=b2.OCID
LEFT JOIN STUD_SELRESULTBLI c ON a.SETID=c.SETID AND a.EnterDate=c.EnterDate AND a.SerNum=c.SerNum AND a.OCID1=c.OCID 	
WHERE b2.SelResultID IN ('01','02') AND b2.ABANDON IS NULL AND cc.OCID=@OCID"
        Return RSTSQL
    End Function

    Function GET_SQL_SD01007R4BB(VRTN As String) As String
        'VRTN = "C"欄位匯出
        If VRTN = "C" Then Return "NAME,IDNO,RELENTERDATE,MDATE,CHANGEMODE,COMNAME,ACTNO,CAPMODE"

        Dim RSTSQL As String = "
WITH WC1 AS (
	SELECT cc.OCID ,cc.CLASSCNAME ,cc.CYCLTYPE ,cc.EXAMDATE,cc.STDATE,cc.FTDATE ,cc.RID,cc.PLANID,format(cc.modifydate,'mmssdd') MSD
	FROM CLASS_CLASSINFO cc	JOIN ID_PLAN ip ON ip.planid=cc.planid	JOIN KEY_PLAN Kp ON Kp.Tplanid=ip.tplanid WHERE cc.OCID=@OCID 
)
,WET2 AS (
  SELECT c.SBEID ,a.NAME ,a.IDNO ,b.OCID1 ,b.CMASTER1,b.RELENTERDATE ,dbo.FN_GET_RESOLDER(a.IDNO,cc.STDATE) RESOLDER
  FROM WC1 cc JOIN STUD_ENTERTYPE2 b on b.OCID1=cc.OCID  JOIN STUD_ENTERTEMP2 a ON a.ESETID=b.ESETID
  LEFT JOIN STUD_BLIGATEDATA28E c ON c.ESETID=b.ESETID AND c.ESERNUM=b.ESERNUM 
  WHERE b.signUpStatus NOT IN (2) /*排除報名審核失敗的學員*/
)
,WET1 AS (
  SELECT c.SBEID ,a.NAME ,a.IDNO ,b.OCID1 ,b.CMASTER1,b.RELENTERDATE ,dbo.FN_GET_RESOLDER(a.IDNO,cc.STDATE) RESOLDER
  FROM WC1 cc JOIN STUD_ENTERTYPE b ON b.OCID1=cc.OCID  JOIN STUD_ENTERTEMP a ON a.SETID=b.SETID
  LEFT JOIN STUD_BLIGATEDATA28E c ON c.SETID=b.SETID AND c.ENTERDATE=b.ENTERDATE AND c.SERNUM=b.SERNUM 
  WHERE NOT EXISTS (SELECT 'X' FROM WET2 X WHERE X.IDNO=A.IDNO AND X.OCID1=B.OCID1)
)
SELECT  a.NAME,a.IDNO,a.RELENTERDATE ,c.MDATE
,CASE WHEN a.RESOLDER IS NOT NULL THEN '軍'
 WHEN c.CHANGEMODE=1 THEN '工作部門或特殊身分異動'
 WHEN c.CHANGEMODE=2 THEN '退保'
 WHEN c.CHANGEMODE=3 THEN '調薪'
 WHEN c.CHANGEMODE=4 THEN '加保' END CHANGEMODE 
,ISNULL(a.RESOLDER,c.COMNAME) COMNAME
,c.ACTNO
,CASE WHEN a.RESOLDER IS NOT NULL THEN '符合' ELSE 
  CASE WHEN SUBSTRING(c.ACTNO,1,2)='09' THEN '不符合' WHEN c.CHANGEMODE NOT IN (2) THEN
   CASE WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.EXAMDATE-1 THEN '尚未開訓僅供參考'
   WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.STDATE THEN '尚未開訓僅供參考' ELSE '符合' END ELSE '不符合' END END 
 +CASE WHEN a.CMASTER1='Y' THEN '(負責人不適用就保)' ELSE '' END 
 +CASE WHEN dbo.FN_GET_BIEF(cc.OCID,a.IDNO)='M' THEN '(多元就業計畫進用人員不適用就保)' ELSE '' END CAPMODE  
FROM WC1 cc
JOIN (SELECT * FROM WET2 UNION SELECT * FROM WET1 ) a ON a.OCID1=cc.OCID
JOIN VIEW_PLAN vp ON vp.PlanID=cc.PlanID
JOIN VIEW_RIDNAME rr ON rr.RID=cc.RID
LEFT JOIN STUD_BLIGATEDATA28E c ON c.SBEID=a.SBEID
"
        Return RSTSQL
    End Function

    '列印
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示
        'Dim v_rblWorkMode As String=TIMS.GetListValue(rblWorkMode)
        'ViewState(TIMS.gcst_rblWorkMode)=v_rblWorkMode
        ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip1

        Dim sMemo As String = ""
        Dim MyValue As String = ""
        Dim prtFilename As String = "" '列印表件名稱
        HidPrinttype.Value = ""
        '28'54 'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim v_print_type As String = TIMS.GetListValue(print_type)
            Select Case v_print_type 'print_type.SelectedValue
                Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                    HidPrinttype.Value = cst_Printtype_B1 '"&Printtype=b.IDNO" '依身分證號(排序)
                Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                    HidPrinttype.Value = cst_Printtype_S1 '"&Printtype=cs.StudentID" '依學號(排序)
            End Select
            'HidPrinttype.Value=cst_Printtype3 '"&Printtype=b.IDNO" '依身分證號(排序)
            '產投報表。'3.產業人才投資方案 [28] '4:充電起飛計畫（在職）[54] 
            Select Case Convert.ToString(sm.UserInfo.LID)
                Case "0", "1" '署(局)、分署(中心)
                    prtFilename = cst_printFN2b '"SD_01_007_R2_b"
                Case Else '委訓單位
                    prtFilename = cst_printFN2bb '"SD_01_007_R2b_b"
                    ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip2
            End Select

            MyValue = ""
            TIMS.SetMyValue(MyValue, "OCID", $"{drCC("OCID")}")
            TIMS.SetMyValue(MyValue, "MSD", $"{drCC("MSD")}")
            TIMS.SetMyValue(MyValue, "Printtype", HidPrinttype.Value)
            sMemo = $"{GET_SEARCH_MEMO()}{MyValue}&prt={prtFilename}"
            'TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm列印, 2, OCIDValue1.Value, sMemo)

            Dim S_PMS2 As New Hashtable From {{"OCID1", $"{drCC("OCID")}"}}
            Dim Sql2 As String = GET_SQL_SD01007R2BB()
            Dim dt2 As DataTable = DbAccess.GetDataTable(Sql2, objconn, S_PMS2)
            '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
            Dim vRESDESC2 As String = TIMS.GET_RESDESCdt(dt2, "NAME,IDNO,STUDENTID,RELENTERDATE,MDATE,CHANGEMODE,COMNAME,ACTNO,CAPMODE")
            Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm列印, ViewState(TIMS.gcst_rblWorkMode), OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt2.Rows.Count, vRESDESC2)

            If TIMS.dtNODATA(dt2) Then
                Common.MessageBox(Me, "查無班級資料!!")
                Exit Sub
            End If

            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, MyValue)
            Return
        End If

        'v_rblStExChoice1: Case "2" '開訓日,Case "4" '開訓日前4日,Case "3" '甄試日,
        Dim v_rblStExChoice1 As String = TIMS.GetListValue(rblStExChoice1)
        '06 66 70 'If TIMS.Cst_TPlanID06AppPlan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case v_rblStExChoice1'rblStExChoice1.SelectedValue
                Case "2" '開訓日
                Case "4" '開訓日前4日
                Case "3" '甄試日
                Case Else
                    Common.MessageBox(Me, "請選擇勾稽條件!!")
                    Exit Sub
            End Select

            Select Case v_rblStExChoice1'rblStExChoice1.SelectedValue
                Case "2" '開訓日 STUD_SELRESULTBLI
                    HidPrinttype.Value = cst_Printtype3 '"&Printtype=b.IDNO" '依身分證號(排序)
                    Select Case Convert.ToString(sm.UserInfo.LID)
                        Case "0", "1" '署(局)、分署(中心)
                            prtFilename = cst_printFN3b '"SD_01_007_R3_b" '開訓日 STUD_SELRESULTBLI
                        Case Else '委訓單位
                            prtFilename = cst_printFN3bb '"SD_01_007_R3_bb" '開訓日 STUD_SELRESULTBLI
                            ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip2
                    End Select

                Case "4" '開訓日前4日 STUD_SELRESULTBLI
                    HidPrinttype.Value = cst_Printtype3 '"&Printtype=b.IDNO" '依身分證號(排序)
                    Select Case Convert.ToString(sm.UserInfo.LID)
                        Case "0", "1" '署(局)、分署(中心)
                            prtFilename = cst_printFN3b '"SD_01_007_R3_b" '開訓日 STUD_SELRESULTBLI
                        Case Else '委訓單位
                            prtFilename = cst_printFN3bb '"SD_01_007_R3_bb" '開訓日 STUD_SELRESULTBLI
                            ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip2
                    End Select

                Case "3" '甄試日 STUD_BLIGATEDATA28E
                    '1.接受企業委託訓練 [07] '2.在職進修訓練 [06]
                    'myvalue += "&Printtype=b.IDNO" '依身分證號(排序)
                    HidPrinttype.Value = cst_Printtype4B '"&Printtype=a.IDNO" '依身分證號(排序)
                    Select Case Convert.ToString(sm.UserInfo.LID)
                        Case "0", "1" '甄試日 STUD_BLIGATEDATA28E '署(局)、分署(中心)
                            prtFilename = cst_printFN4b '"SD_01_007_R4_b"
                            If sm.UserInfo.TPlanID.Equals(TIMS.Cst_TPlanID70) Then prtFilename = cst_printFN5b
                        Case Else  '甄試日 STUD_BLIGATEDATA28E '委訓單位
                            prtFilename = cst_printFN4bb '"SD_01_007_R4_bb"
                            If sm.UserInfo.TPlanID.Equals(TIMS.Cst_TPlanID70) Then prtFilename = cst_printFN5bb
                            ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip2
                    End Select
            End Select
        End If
        If HidPrinttype.Value = "" Then
            Common.MessageBox(Me, "!!該計畫不提供該報表，請洽系統管理者!!")
            Exit Sub
        End If
        If prtFilename = "" Then
            Common.MessageBox(Me, "!!該計畫不提供該報表，請洽系統管理者!!")
            Exit Sub
        End If

        MyValue = ""
        TIMS.SetMyValue(MyValue, "OCID", $"{drCC("OCID")}")
        TIMS.SetMyValue(MyValue, "MSD", $"{drCC("MSD")}")
        TIMS.SetMyValue(MyValue, "Printtype", HidPrinttype.Value)
        'TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm列印, 2, OCIDValue1.Value, sMemo)
        sMemo &= $"{GET_SEARCH_MEMO()}{MyValue}&prt={prtFilename}"

        Dim S_PMS3 As New Hashtable From {{"OCID", $"{drCC("OCID")}"}}
        'Case "2" '開訓日,Case "4" '開訓日前4日,Case "3" '甄試日,
        Dim Sql3 As String = If(v_rblStExChoice1 = "3", GET_SQL_SD01007R4BB(""), GET_SQL_SD01007R3BB(""))
        Dim vCOLNM3 As String = If(v_rblStExChoice1 = "3", GET_SQL_SD01007R4BB("C"), GET_SQL_SD01007R3BB("C"))
        Dim dt3 As DataTable = DbAccess.GetDataTable(Sql3, objconn, S_PMS3)
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC3 As String = TIMS.GET_RESDESCdt(dt3, vCOLNM3)
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm列印, ViewState(TIMS.gcst_rblWorkMode), OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt3.Rows.Count, vRESDESC3)
        If TIMS.dtNODATA(dt3) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, MyValue)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '判斷機構是否只有一個班級
        Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
    End Sub

#Region "匯出e網民眾投保狀況檢核表"
    '匯出EXCEL (匯出e網民眾投保狀況檢核表)  STUD_BLIGATEDATA28E 產投 (LpfBatchGet28e.vbproj) 
    Sub ExpReport28()
        Dim sFileName1 As String = "匯出e網民眾投保狀況檢核表"

        Dim sSql As String = ""
        sSql &= " SELECT b.SIGNNO ,a.NAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK,a.IDNO" & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) ClassCName" & vbCrLf
        sSql &= " ,CASE c.ChangeMode WHEN 1 THEN '工作部門或特殊身分異動' WHEN 2 THEN '退保' WHEN 3 THEN '調薪' WHEN 4 THEN '加保' END CHANGEMODE" & vbCrLf
        sSql &= " ,c.COMNAME,c.ACTNO" & vbCrLf
        'sql += "  ,CASE WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.STDATE THEN '尚未勾稽' ELSE CASE WHEN c.ESERNUM IS NULL THEN '不符合' ELSE '符合' END END useValid" & vbCrLf
        sSql &= " ,CASE WHEN GETDATE() < cc.STDATE THEN '尚未勾稽' ELSE" & vbCrLf
        sSql &= " CASE WHEN SUBSTRING(c.ACTNO,1,2)='09' THEN '不符合' WHEN c.ESERNUM IS NULL THEN '不符合' ELSE '符合' END END USEVALID" & vbCrLf
        sSql &= " FROM STUD_ENTERTYPE2 b" & vbCrLf
        sSql &= " JOIN STUD_ENTERTEMP2 a ON a.ESETID=b.ESETID" & vbCrLf
        sSql &= " JOIN CLASS_CLASSINFO cc ON cc.OCID=b.OCID1" & vbCrLf
        sSql &= " JOIN ID_PLAN ip ON ip.PLANID=cc.PLANID" & vbCrLf
        sSql &= " JOIN KEY_PLAN Kp ON Kp.Tplanid=ip.tplanid" & vbCrLf
        sSql &= " LEFT JOIN STUD_BLIGATEDATA28E c ON c.ESERNUM=b.ESERNUM AND c.ESETID=b.ESETID" & vbCrLf
        sSql &= " WHERE ip.TPLANID='28'" & vbCrLf '產投:28
        'sql &= " AND ip.TPLANID=@TPLANID" & vbCrLf '產投:28
        sSql &= " AND cc.OCID=@OCID" & vbCrLf

        Dim v_print_type As String = TIMS.GetListValue(print_type)
        Select Case v_print_type 'print_type.SelectedValue
            Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                sSql &= " ORDER BY a.IDNO" & vbCrLf
            Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                sSql &= " ORDER BY b.SIGNNO" & vbCrLf
            Case Else
                sSql &= " ORDER BY b.SIGNNO" & vbCrLf
        End Select

        Dim sCmd As New SqlCommand(sSql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            '.Parameters.Add("TPLANID", SqlDbType.VarChar).Value=sm.UserInfo.TPlanID
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            dt.Load(.ExecuteReader())
        End With

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&ACT={sFileName1}&OCIDValue1={OCIDValue1.Value}&COUNT={dt.Rows.Count}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SIGNNO,NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,USEVALID")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If

        'Response.Clear()
        'Response.Charset="BIG5"
        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8))
        'Response.ContentType="Application/octet-stream"
        'Response.ContentEncoding=System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding=System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType="application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType="application/ms-excel;charset=utf-8"
        'Response.ContentType="application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        'Common.RespWrite(Me, "<html>")
        'Common.RespWrite(Me, "<head>")
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        '套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'sCommon.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-Number - Format():  "0" 
        'Common.RespWrite(Me, "</style>")
        'Common.RespWrite(Me, "</head>")

        '序號,身分證字號
        Const s_title1 As String = "報名序號,姓名,身分證統一編號,班級名稱,保險別,投保單位,保險證號,在保資格<br>(符合/不符合)"
        Const s_data1 As String = "SIGNNO,NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,USEVALID"
        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        For Each s_V1 As String In As_title1
            ExportStr &= String.Format("<td>{0}</td>", s_V1) & vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            For Each s_V1 As String In As_data1
                'Select Case s_V1
                '    Case "IDNO"
                '        '署(局)、分署(中心) 正常顯示／委訓單位 模糊顯示 / '署(局)、分署(中心) 不需隱碼／委訓單位 需隱碼
                '        Dim flag_N1 As Boolean=If(sm.UserInfo.LID=0, True, If(sm.UserInfo.LID=1, True, False))
                '        Dim S_IDNO_N1 As String=If(flag_N1, Convert.ToString(dr(s_V1)), TIMS.strMask(Convert.ToString(dr(s_V1)), 1))
                '        ExportStr &= String.Format("<td>{0}</td>", S_IDNO_N1) & vbTab
                '    Case Else
                '        ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(s_V1))) & vbTab
                'End Select
                ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(s_V1))) & vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        'Response.End()

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strHTML", strHTML)
        'parmsExp.Add("ResponseNoEnd", "")
        TIMS.Utl_ExportRp1(Me, parmsExp)
    End Sub

#End Region

    '匯出e網民眾投保狀況檢核表 (產投28)
    Protected Sub BtnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示
        'Dim v_rblWorkMode As String=TIMS.GetListValue(rblWorkMode)
        'ViewState(TIMS.gcst_rblWorkMode)=v_rblWorkMode
        ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip1

        '匯出e網民眾投保狀況檢核表
        '甄試學員投保狀況檢核表
        Dim flagNoExp1 As Boolean = True
        '"28"
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            flagNoExp1 = False
            'btnExport.Visible=True
            'btnExport.Text="匯出e網民眾投保狀況檢核表" 'DEF:28 'strTitle1="匯出e網民眾投保狀況檢核表"
            Call ExpReport28()
        End If
        If flagNoExp1 Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If
    End Sub


    '查詢原因-INQUIRY
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""

        center.Text = TIMS.ClearSQM(center.Text)
        OCID1.Text = TIMS.ClearSQM(OCID1.Text)
        txtCJOB_NAME.Text = TIMS.ClearSQM(txtCJOB_NAME.Text)
        Dim v_print_type As String = TIMS.GetListValue(print_type)
        Dim v_rblStExChoice1 As String = TIMS.GetListValue(rblStExChoice1)

        If center.Text <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Text)
        If OCID1.Text <> "" Then RstMemo &= String.Concat("&班級名稱=", OCID1.Text)
        If txtCJOB_NAME.Text <> "" Then RstMemo &= String.Concat("&通俗職類=", txtCJOB_NAME.Text)
        If v_print_type <> "" Then RstMemo &= String.Concat("&列印排序方式=", v_print_type)
        If v_rblStExChoice1 <> "" Then RstMemo &= String.Concat("&勾稽條件=", v_rblStExChoice1)

        Return RstMemo
    End Function



#Region "匯出投保狀況檢核表 06 66"

    '"SD_01_007_R3_b" '開訓日 STUD_SELRESULTBLI
    '匯出EXCEL (甄試學員投保狀況檢核表) (自辦)在職進修訓練 STUD_SELRESULTBLI 開訓日 
    ''' <summary>
    ''' STUD_SELRESULTBLI-開訓日
    ''' </summary>
    ''' <param name="sTPlanID"></param>
    Sub ExpReport06_66_S(ByVal sTPlanID As String)
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sTPlanID) = -1 Then Exit Sub

        Dim sFileName1 As String = "參訓學員投保狀況檢核表" '開訓日

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.OCID ,cc.CLASSCNAME ,cc.CYCLTYPE" & vbCrLf
        sql &= " 	,cc.EXAMDATE,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= "    ,format(cc.modifydate,'mmssdd') MSD" & vbCrLf
        sql &= " 	FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " 	JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= " 	JOIN KEY_PLAN Kp ON Kp.Tplanid=ip.tplanid" & vbCrLf
        'sql &= " 	AND ip.TPLANID='06' AND ip.YEARS= '2021' and cc.OCID =132789" & vbCrLf
        sql &= " 	WHERE ip.TPLANID=@TPLANID AND cc.OCID=@OCID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " , WS1 AS ( SELECT a.SETID,a.ENTERDATE,a.SERNUM" & vbCrLf
        sql &= " 	,a.OCID1,a.RELENTERDATE" & vbCrLf
        sql &= " 	,b.NAME ,b.IDNO" & vbCrLf
        sql &= " 	,c.MDATE,c.CHANGEMODE,c.COMNAME ,c.ACTNO" & vbCrLf
        sql &= " 	,dbo.FN_GET_RESOLDER(b.IDNO,cc.STDATE) RESOLDER" & vbCrLf
        sql &= " 	FROM WC1 cc" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTYPE a ON a.OCID1=cc.OCID" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTEMP b ON a.SETID=b.SETID" & vbCrLf
        sql &= " 	JOIN STUD_SELRESULT b2 ON a.SETID=b2.SETID AND a.EnterDate=b2.EnterDate AND a.SerNum=b2.SerNum AND a.OCID1=b2.OCID" & vbCrLf
        sql &= " 	LEFT JOIN STUD_SELRESULTBLI c ON a.SETID=c.SETID AND a.EnterDate=c.EnterDate AND a.SerNum=c.SerNum AND a.OCID1=c.OCID" & vbCrLf
        sql &= " 	WHERE b2.SelResultID IN ('01','02')" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT wb.NAME" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(wb.IDNO) IDNO_MK,wb.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME ,cc.CyclType) CLASSCNAME" & vbCrLf
        sql &= " ,CASE WHEN wb.RESOLDER IS NOT NULL THEN '軍'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=1 THEN '工作部門或特殊身分異動'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=2 THEN '退保'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=3 THEN '調薪'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=4 THEN '加保' ELSE '' END CHANGEMODE" & vbCrLf
        sql &= " ,ISNULL(wb.RESOLDER,wb.COMNAME) COMNAME" & vbCrLf
        sql &= " ,wb.ACTNO" & vbCrLf
        sql &= " ,wb.ENTERDATE" & vbCrLf
        sql &= " ,CASE WHEN wb.RESOLDER IS NOT NULL THEN '符合' ELSE" & vbCrLf
        sql &= "   CASE WHEN SUBSTRING(wb.ACTNO,1,2)='09' THEN '不符合' WHEN wb.CHANGEMODE NOT IN (2) THEN" & vbCrLf
        sql &= "    CASE WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.EXAMDATE-1 THEN '尚未開訓僅供參考'" & vbCrLf
        sql &= "    WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.STDATE THEN '尚未開訓僅供參考' ELSE '符合' END ELSE '不符合' END END CAPMODE" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WS1 wb on wb.OCID1=cc.OCID" & vbCrLf

        Dim v_print_type As String = TIMS.GetListValue(print_type)
        Select Case v_print_type 'print_type.SelectedValue
            Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                'HidPrinttype.Value=cst_Printtype1 '"&Printtype=b.IDNO" '依身分證號(排序)
                sql &= " ORDER BY wb.IDNO" & vbCrLf
            Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                'HidPrinttype.Value=cst_Printtype2 '"&Printtype=cs.StudentID" '依學號(排序)
                'sql &= " ORDER BY c.ESERNUM, c.ENTERDATE" & vbCrLf
                sql &= " ORDER BY wb.ENTERDATE" & vbCrLf
            Case Else
                'sql &= " ORDER BY c.ESERNUM, c.ENTERDATE" & vbCrLf
                sql &= " ORDER BY wb.ENTERDATE" & vbCrLf
        End Select

        Dim s_parms As New Hashtable From {{"TPLANID", sTPlanID}, {"OCID", OCIDValue1.Value}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, s_parms)

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&ACT={sFileName1}&TPLANID={sTPlanID}&OCIDValue1={OCIDValue1.Value}&COUNT={dt.Rows.Count}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,CAPMODE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If

        dt.DefaultView.Sort = "IDNO" '排序
        dt = TIMS.dv2dt(dt.DefaultView)

        '序號,
        Const s_title1 As String = "姓名,身分證字號,班級名稱,保險別,投保單位,保險證號,在保資格<br>(符合/不符合)"
        Const s_data1 As String = "NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,CAPMODE"
        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= String.Format("<td>{0}</td>", "序號") & vbTab
        For Each s_V1 As String In As_title1
            ExportStr &= String.Format("<td>{0}</td>", s_V1) & vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            ExportStr &= String.Format("<td>{0}</td>", CStr(iSeqno)) & vbTab
            For Each s_V1 As String In As_data1
                'Select Case s_V1
                '    Case "IDNO"
                '        '署(局)、分署(中心) 正常顯示／委訓單位 模糊顯示 / '署(局)、分署(中心) 不需隱碼／委訓單位 需隱碼
                '        Dim flag_N1 As Boolean=If(sm.UserInfo.LID=0, True, If(sm.UserInfo.LID=1, True, False))
                '        Dim S_IDNO_N1 As String=If(flag_N1, Convert.ToString(dr(s_V1)), TIMS.strMask(Convert.ToString(dr(s_V1)), 1))
                '        ExportStr &= String.Format("<td>{0}</td>", S_IDNO_N1) & vbTab
                '    Case Else
                '        ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(s_V1))) & vbTab
                'End Select
                ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(s_V1))) & vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        'Response.End()

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        'parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        'TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    ''' <summary>
    ''' 'STUD_SELRESULTBLI-開訓日4日
    ''' </summary>
    ''' <param name="sTPlanID"></param>
    Sub ExpReport06_66_S4(ByVal sTPlanID As String)
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sTPlanID) = -1 Then Exit Sub

        Dim sFileName1 As String = "參訓學員投保狀況檢核表" '開訓日4日

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.OCID ,cc.CLASSCNAME ,cc.CYCLTYPE" & vbCrLf
        sql &= " 	,cc.EXAMDATE,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= "    ,format(cc.modifydate,'mmssdd') MSD" & vbCrLf
        sql &= " 	FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " 	JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= " 	JOIN KEY_PLAN Kp ON Kp.Tplanid=ip.tplanid" & vbCrLf
        sql &= " 	WHERE ip.TPLANID=@TPLANID AND cc.OCID=@OCID" & vbCrLf
        'sql &= " 	AND ip.TPLANID='06' AND ip.YEARS= '2021' and cc.OCID =132789" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WS1 AS ( SELECT a.SETID,a.ENTERDATE,a.SERNUM" & vbCrLf
        sql &= " 	,a.OCID1,a.RELENTERDATE" & vbCrLf
        sql &= " 	,b.NAME ,b.IDNO" & vbCrLf
        sql &= " 	,c.MDATE,c.CHANGEMODE,c.COMNAME ,c.ACTNO" & vbCrLf
        sql &= " 	,dbo.FN_GET_RESOLDER(b.IDNO,cc.STDATE) RESOLDER" & vbCrLf
        sql &= " 	,cc.EXAMDATE,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= " 	FROM WC1 cc" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTYPE a ON a.OCID1=cc.OCID" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTEMP b ON a.SETID=b.SETID" & vbCrLf
        sql &= " 	JOIN STUD_SELRESULT b2 ON a.SETID=b2.SETID AND a.EnterDate=b2.EnterDate AND a.SerNum=b2.SerNum AND a.OCID1=b2.OCID" & vbCrLf
        sql &= " 	LEFT JOIN STUD_SELRESULTBLI c ON a.SETID=c.SETID AND a.EnterDate=c.EnterDate AND a.SerNum=c.SerNum AND a.OCID1=c.OCID" & vbCrLf
        sql &= " 	WHERE b2.SelResultID IN ('01','02')" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT wb.NAME" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(wb.IDNO) IDNO_MK,wb.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME ,cc.CyclType) CLASSCNAME" & vbCrLf
        sql &= " ,CASE WHEN wb.RESOLDER IS NOT NULL THEN '軍'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=1 THEN '工作部門或特殊身分異動'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=2 THEN '退保'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=3 THEN '調薪'" & vbCrLf
        sql &= " WHEN wb.CHANGEMODE=4 THEN '加保'" & vbCrLf
        sql &= " ELSE '' END CHANGEMODE" & vbCrLf
        sql &= " ,ISNULL(wb.RESOLDER,wb.COMNAME) COMNAME" & vbCrLf
        sql &= " ,wb.ACTNO" & vbCrLf
        sql &= " ,wb.ENTERDATE" & vbCrLf
        sql &= " ,CASE WHEN wb.RESOLDER IS NOT NULL THEN '符合' ELSE" & vbCrLf
        sql &= "   CASE WHEN SUBSTRING(wb.ACTNO,1,2)='09' THEN '不符合' WHEN wb.CHANGEMODE NOT IN (2) THEN" & vbCrLf
        sql &= "    CASE WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.EXAMDATE-1 THEN '尚未開訓僅供參考'" & vbCrLf
        sql &= "    WHEN dbo.TRUNC_DATETIME(GETDATE()) < cc.STDATE THEN '尚未開訓僅供參考' ELSE '符合' END ELSE '不符合' END END CAPMODE" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WS1 wb ON wb.OCID1=cc.OCID" & vbCrLf

        Dim v_print_type As String = TIMS.GetListValue(print_type)
        Select Case v_print_type 'print_type.SelectedValue
            Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                'HidPrinttype.Value=cst_Printtype1 '"&Printtype=b.IDNO" '依身分證號(排序)
                sql &= " ORDER BY wb.IDNO" & vbCrLf
            Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                'HidPrinttype.Value=cst_Printtype2 '"&Printtype=cs.StudentID" '依學號(排序)
                'sql &= " ORDER BY c.ESERNUM, c.ENTERDATE" & vbCrLf
                sql &= " ORDER BY wb.ENTERDATE" & vbCrLf
            Case Else
                'sql &= " ORDER BY c.ESERNUM, c.ENTERDATE" & vbCrLf
                sql &= " ORDER BY wb.ENTERDATE" & vbCrLf
        End Select

        Dim s_parms As New Hashtable From {{"TPLANID", sTPlanID}, {"OCID", OCIDValue1.Value}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, s_parms)

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&ACT={sFileName1}&TPLANID={sTPlanID}&OCIDValue1={OCIDValue1.Value}&COUNT={dt.Rows.Count}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,CAPMODE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If

        dt.DefaultView.Sort = "IDNO" '排序
        dt = TIMS.dv2dt(dt.DefaultView)

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>序號</td>" & vbTab
        ExportStr &= "<td>姓名</td>" & vbTab
        ExportStr &= "<td>身分證字號</td>" & vbTab
        ExportStr &= "<td>班級名稱</td>" & vbTab
        ExportStr &= "<td>保險別</td>" & vbTab
        ExportStr &= "<td>投保單位</td>" & vbTab
        ExportStr &= "<td>保險證號</td>" & vbTab
        ExportStr &= "<td>在保資格<br>(符合/不符合)</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            'ExportStr &= "<td>" & Convert.ToString(dr("SIGNNO")) & "</td>" & vbTab
            ExportStr &= "<td>" & CStr(iSeqno) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("NAME")) & "</td>" & vbTab

            '署(局)、分署(中心) 不需隱碼／委訓單位 需隱碼
            'Dim fgMASK1 As Boolean=If(sm.UserInfo.LID=0, False, If(sm.UserInfo.LID=1, False, True))
            'Dim vIDNO As String=If(fgMASK1, TIMS.strMask(Convert.ToString(dr("IDNO")), 1), Convert.ToString(dr("IDNO")))
            'ExportStr &= String.Format("<td>{0}</td>", vIDNO) & vbTab
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("IDNO_MK"))) & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CLASSCNAME")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CHANGEMODE")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("COMNAME")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ACTNO")) & "</td>" & vbTab
            'Dim S_CAPMODE As String=If(Convert.ToString(dr("USEVALID")) <> "", Convert.ToString(dr("USEVALID")), Convert.ToString(dr("CAPMODE")))
            'ExportStr &= "<td>" & S_CAPMODE & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CAPMODE")) & "</td>" & vbTab
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        'Response.End()

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strHTML", strHTML)
        'parmsExp.Add("ResponseNoEnd", "")
        TIMS.Utl_ExportRp1(Me, parmsExp)
    End Sub

    ''' <summary>
    ''' STUD_BLIGATEDATA28E-甄試日 匯出EXCEL (匯出 甄試學員投保狀況檢核表) (自辦)在職進修訓練  
    ''' </summary>
    ''' <param name="sTPlanID"></param>
    Sub ExpReport06_66_E(ByVal sTPlanID As String)
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sTPlanID) = -1 Then Exit Sub

        Dim sFileName1 As String = "甄試學員投保狀況檢核表" '甄試日

        '甄試日
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.OCID ,cc.CLASSCNAME ,cc.CYCLTYPE" & vbCrLf
        sql &= "    ,cc.EXAMDATE,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= "    ,format(cc.modifydate,'mmssdd') MSD" & vbCrLf
        sql &= "    FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= "    JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= "    JOIN KEY_PLAN Kp ON Kp.Tplanid=ip.tplanid" & vbCrLf
        'sql &= " 	AND ip.TPLANID='06' AND ip.YEARS= '2021' and cc.OCID =132789" & vbCrLf
        sql &= "    WHERE ip.TPLANID=@TPLANID AND cc.OCID=@OCID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WET2 AS ( SELECT c.SBEID ,a.NAME ,a.IDNO ,b.OCID1 ,b.CMASTER1" & vbCrLf
        sql &= "    ,cc.ClassCName ,cc.CyclType ,cc.EXAMDATE,cc.STDATE" & vbCrLf
        sql &= "    ,dbo.FN_GET_RESOLDER(a.IDNO,cc.STDATE) RESOLDER" & vbCrLf
        sql &= "    FROM WC1 cc" & vbCrLf
        sql &= "    JOIN STUD_ENTERTYPE2 b ON b.OCID1=cc.OCID" & vbCrLf
        sql &= "    JOIN STUD_ENTERTEMP2 a ON a.ESETID=b.ESETID" & vbCrLf
        sql &= "    LEFT JOIN STUD_BLIGATEDATA28E c ON c.ESETID=b.ESETID AND c.ESERNUM=b.ESERNUM" & vbCrLf
        'OJT-21012201：<系統> 在職進修訓練(自辦) - 參訓學員投保狀況檢核表：審核失敗的人不要顯示在甄試日勾稽的名單中
        sql &= "    WHERE b.signUpStatus NOT IN (2)" & vbCrLf '排除掉審核失敗的人(e網報名審核)
        sql &= " )" & vbCrLf
        sql &= " ,WET1 AS ( SELECT c.SBEID ,a.NAME ,a.IDNO ,b.OCID1 ,b.CMASTER1" & vbCrLf
        sql &= "    ,cc.ClassCName ,cc.CyclType ,cc.EXAMDATE,cc.STDATE" & vbCrLf
        sql &= "    ,dbo.FN_GET_RESOLDER(a.IDNO,cc.STDATE) RESOLDER" & vbCrLf
        sql &= "    FROM WC1 cc" & vbCrLf
        sql &= "    JOIN STUD_ENTERTYPE b ON b.OCID1=cc.OCID" & vbCrLf
        sql &= "    JOIN STUD_ENTERTEMP a ON a.SETID=b.SETID" & vbCrLf
        sql &= "    LEFT JOIN STUD_BLIGATEDATA28E c ON c.SETID=b.SETID AND c.ENTERDATE=b.ENTERDATE AND c.SERNUM=b.SERNUM" & vbCrLf
        sql &= "    WHERE NOT EXISTS (SELECT 'X' FROM WET2 X WHERE X.IDNO=A.IDNO AND X.OCID1=B.OCID1)" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT wb.SBEID ,wb.NAME" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(wb.IDNO) IDNO_MK,wb.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(wb.CLASSCNAME ,wb.CyclType) CLASSCNAME" & vbCrLf
        sql &= " ,CASE WHEN wb.RESOLDER IS NOT NULL THEN '軍'" & vbCrLf
        sql &= " WHEN c.CHANGEMODE=1 THEN '工作部門或特殊身分異動'" & vbCrLf
        sql &= " WHEN c.CHANGEMODE=2 THEN '退保'" & vbCrLf
        sql &= " WHEN c.CHANGEMODE=3 THEN '調薪'" & vbCrLf
        sql &= " WHEN c.CHANGEMODE=4 THEN '加保' END CHANGEMODE" & vbCrLf
        sql &= " ,ISNULL(wb.RESOLDER,c.COMNAME) COMNAME" & vbCrLf
        sql &= " ,c.ACTNO" & vbCrLf

        sql &= " ,CASE WHEN wb.RESOLDER IS NOT NULL THEN '符合' ELSE" & vbCrLf
        sql &= "   CASE WHEN SUBSTRING(c.ACTNO,1,2)='09' THEN '不符合' WHEN c.CHANGEMODE NOT IN (2) THEN" & vbCrLf
        sql &= "    CASE WHEN dbo.TRUNC_DATETIME(GETDATE()) < wb.EXAMDATE-1 THEN '尚未開訓僅供參考'" & vbCrLf
        sql &= "    WHEN dbo.TRUNC_DATETIME(GETDATE()) < wb.STDATE THEN '尚未開訓僅供參考' ELSE '符合' END ELSE '不符合' END END" & vbCrLf
        sql &= " +CASE WHEN wb.CMASTER1='Y' THEN '(負責人不適用就保)' ELSE '' END" & vbCrLf
        sql &= " +CASE WHEN dbo.FN_GET_BIEF(wb.OCID1,wb.IDNO)='M' THEN '(多元就業計畫進用人員不適用就保)' ELSE '' END CAPMODE" & vbCrLf

        sql &= " ,c.esetid ,c.esernum ,c.setid ,c.enterdate ,c.sernum" & vbCrLf
        sql &= " FROM (SELECT * FROM WET2 UNION SELECT * FROM WET1) wb" & vbCrLf
        sql &= " LEFT JOIN STUD_BLIGATEDATA28E c ON c.SBEID=wb.SBEID" & vbCrLf
        Dim v_print_type As String = TIMS.GetListValue(print_type)
        Select Case v_print_type 'print_type.SelectedValue
            Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                'HidPrinttype.Value=cst_Printtype1 '"&Printtype=b.IDNO" '依身分證號(排序)
                sql &= " ORDER BY wb.IDNO" & vbCrLf
            Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                'HidPrinttype.Value=cst_Printtype2 '"&Printtype=cs.StudentID" '依學號(排序)
                sql &= " ORDER BY c.ESERNUM, c.ENTERDATE" & vbCrLf
            Case Else
                sql &= " ORDER BY c.ESERNUM, c.ENTERDATE" & vbCrLf
        End Select

        Dim s_parms As New Hashtable From {{"TPLANID", sTPlanID}, {"OCID", OCIDValue1.Value}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, s_parms)

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&ACT={sFileName1}&TPLANID={sTPlanID}&OCIDValue1={OCIDValue1.Value}&COUNT={dt.Rows.Count}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,CAPMODE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Return
        End If

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        Dim ExportStr As String = ""
        '建立抬頭 '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>序號</td>" & vbTab
        ExportStr &= "<td>姓名</td>" & vbTab
        ExportStr &= "<td>身分證字號</td>" & vbTab
        ExportStr &= "<td>班級名稱</td>" & vbTab
        ExportStr &= "<td>保險別</td>" & vbTab
        ExportStr &= "<td>投保單位</td>" & vbTab
        ExportStr &= "<td>保險證號</td>" & vbTab
        ExportStr &= "<td>在保資格<br>(符合/不符合)</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            'ExportStr &= "<td>" & Convert.ToString(dr("SIGNNO")) & "</td>" & vbTab
            ExportStr &= "<td>" & CStr(iSeqno) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("NAME")) & "</td>" & vbTab

            '署(局)、分署(中心) 不需隱碼／委訓單位 需隱碼
            'Dim fgMASK1 As Boolean=If(sm.UserInfo.LID=0, False, If(sm.UserInfo.LID=1, False, True))
            'Dim vIDNO As String=If(fgMASK1, TIMS.strMask(Convert.ToString(dr("IDNO")), 1), Convert.ToString(dr("IDNO")))
            'ExportStr &= String.Format("<td>{0}</td>", vIDNO) & vbTab
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("IDNO_MK"))) & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CLASSCNAME")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CHANGEMODE")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("COMNAME")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ACTNO")) & "</td>" & vbTab
            'Dim S_CAPMODE As String=If(Convert.ToString(dr("USEVALID")) <> "", Convert.ToString(dr("USEVALID")), Convert.ToString(dr("CAPMODE")))
            'ExportStr &= "<td>" & S_CAPMODE & "</td>" & vbTab
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CAPMODE"))) & vbTab

            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        'Response.End()

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strHTML", strHTML)
        'parmsExp.Add("ResponseNoEnd", "")
        TIMS.Utl_ExportRp1(Me, parmsExp)
    End Sub

#End Region

#Region "匯出投保狀況檢核表 70"

    ''' <summary> STUD_BLIGATEDATA28E-甄試日 匯出EXCEL (匯出 甄試學員投保狀況檢核表)</summary>
    ''' <param name="sTPlanID"></param>
    Sub ExpReport70_E(ByVal sTPlanID As String)
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sTPlanID) = -1 Then Exit Sub

        Dim sFileName1 As String = "甄試學員投保狀況檢核表" '甄試日

        '甄試日
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.OCID ,cc.CLASSCNAME ,cc.CYCLTYPE" & vbCrLf
        sql &= " ,cc.EXAMDATE,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= " ,format(cc.modifydate,'mmssdd') MSD" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= " JOIN KEY_PLAN Kp ON Kp.Tplanid=ip.tplanid" & vbCrLf
        sql &= " WHERE ip.TPLANID=@TPLANID AND cc.OCID=@OCID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WET1 AS ( SELECT c.SBEID ,a.NAME ,a.IDNO ,b.OCID1 ,b.CMASTER1,b.RELENTERDATE" & vbCrLf
        sql &= " ,cc.CLASSCNAME ,cc.CyclType ,cc.EXAMDATE ,cc.STDATE" & vbCrLf
        sql &= " ,dbo.FN_GET_RESOLDER(a.IDNO,cc.STDATE) RESOLDER" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b on b.OCID1=cc.OCID" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a ON a.SETID=b.SETID" & vbCrLf
        sql &= " JOIN STUD_BLIGATEDATA28E c ON c.SETID=b.SETID AND c.ENTERDATE=b.ENTERDATE AND c.SERNUM=b.SERNUM" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT a1.SBEID ,a1.NAME" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(a1.IDNO) IDNO_MK,a1.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a1.CLASSCNAME ,a1.CyclType) CLASSCNAME" & vbCrLf
        sql &= " ,CASE WHEN a1.RESOLDER IS NOT NULL THEN '軍'" & vbCrLf
        sql &= "  WHEN c.CHANGEMODE=1 THEN '工作部門或特殊身分異動'" & vbCrLf
        sql &= "  WHEN c.CHANGEMODE=2 THEN '退保'" & vbCrLf
        sql &= "  WHEN c.CHANGEMODE=3 THEN '調薪'" & vbCrLf
        sql &= "  WHEN c.CHANGEMODE=4 THEN '加保' END CHANGEMODE" & vbCrLf
        sql &= " ,ISNULL(a1.RESOLDER,c.COMNAME) COMNAME" & vbCrLf
        sql &= " ,c.ACTNO,a1.EXAMDATE" & vbCrLf
        sql &= " ,CASE WHEN a1.RESOLDER IS NOT NULL THEN '符合' ELSE" & vbCrLf
        sql &= "   CASE WHEN SUBSTRING(c.ACTNO,1,2)='09' THEN '不符合' WHEN c.CHANGEMODE NOT IN (2) THEN" & vbCrLf
        sql &= "    CASE WHEN dbo.TRUNC_DATETIME(GETDATE()) < a1.EXAMDATE-1 THEN '尚未開訓僅供參考'" & vbCrLf
        sql &= "    WHEN dbo.TRUNC_DATETIME(GETDATE()) < a1.STDATE THEN '尚未開訓僅供參考' ELSE '符合' END ELSE '不符合' END END" & vbCrLf
        sql &= "  +CASE WHEN a1.CMASTER1='Y' THEN '(負責人不適用就保)' ELSE '' END" & vbCrLf
        sql &= "  +CASE WHEN dbo.FN_GET_BIEF(a1.OCID1,a1.IDNO)='M' THEN '(多元就業計畫進用人員不適用就保)' ELSE '' END CAPMODE" & vbCrLf
        sql &= " ,c.ESETID ,c.ESERNUM" & vbCrLf
        sql &= " ,c.SETID ,c.ENTERDATE ,c.SERNUM" & vbCrLf
        sql &= " FROM WET1 a1" & vbCrLf
        sql &= " JOIN STUD_BLIGATEDATA28E c ON c.SBEID=a1.SBEID" & vbCrLf

        Dim v_print_type As String = TIMS.GetListValue(print_type)
        Select Case v_print_type 'print_type.SelectedValue
            Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                'HidPrinttype.Value=cst_Printtype1 '"&Printtype=b.IDNO" '依身分證號(排序)
                sql &= " ORDER BY a1.IDNO" & vbCrLf
            Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                'HidPrinttype.Value=cst_Printtype2 '"&Printtype=cs.StudentID" '依學號(排序)
                sql &= " ORDER BY c.ENTERDATE,c.SETID" & vbCrLf
            Case Else
                sql &= " ORDER BY c.ENTERDATE,c.SETID" & vbCrLf
        End Select

        Dim s_parms As New Hashtable From {{"TPLANID", sTPlanID}, {"OCID", OCIDValue1.Value}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, s_parms)

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        sMemo &= $"&ACT={sFileName1}&TPLANID={sTPlanID}&OCIDValue1={OCIDValue1.Value}&COUNT={dt.Rows.Count}"
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO_MK,CLASSCNAME,CHANGEMODE,COMNAME,ACTNO,CAPMODE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Return
        End If

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        Dim ExportStr As String = ""
        '建立抬頭
        '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>序號</td>" & vbTab
        ExportStr &= "<td>姓名</td>" & vbTab
        ExportStr &= "<td>身分證字號</td>" & vbTab
        ExportStr &= "<td>班級名稱</td>" & vbTab
        ExportStr &= "<td>保險別</td>" & vbTab
        ExportStr &= "<td>投保單位</td>" & vbTab
        ExportStr &= "<td>保險證號</td>" & vbTab
        ExportStr &= "<td>在保資格<br>(符合/不符合)</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            'ExportStr &= "<td>" & Convert.ToString(dr("SIGNNO")) & "</td>" & vbTab
            ExportStr &= "<td>" & CStr(iSeqno) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("NAME")) & "</td>" & vbTab

            '署(局)、分署(中心) 不需隱碼／委訓單位 需隱碼
            'Dim fgMASK1 As Boolean=If(sm.UserInfo.LID=0, False, If(sm.UserInfo.LID=1, False, True))
            'Dim vIDNO As String=If(fgMASK1, TIMS.strMask(Convert.ToString(dr("IDNO")), 1), Convert.ToString(dr("IDNO")))
            'ExportStr &= String.Format("<td>{0}</td>", vIDNO) & vbTab
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("IDNO_MK"))) & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CLASSCNAME")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CHANGEMODE")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("COMNAME")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ACTNO")) & "</td>" & vbTab
            'Dim S_CAPMODE As String=If(Convert.ToString(dr("USEVALID")) <> "", Convert.ToString(dr("USEVALID")), Convert.ToString(dr("CAPMODE")))
            'ExportStr &= "<td>" & S_CAPMODE & "</td>" & vbTab
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("CAPMODE"))) & vbTab

            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        'Response.End()

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strHTML", strHTML)
        'parmsExp.Add("ResponseNoEnd", "")
        TIMS.Utl_ExportRp1(Me, parmsExp)
    End Sub

#End Region

    '匯出投保狀況檢核表 (06 66 70)
    Protected Sub BtnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim v_rblStExChoice1 As String = TIMS.GetListValue(rblStExChoice1)
        Select Case v_rblStExChoice1'rblStExChoice1.SelectedValue
            Case "2" '開訓日
            Case "4" '開訓日前4日
            Case "3" '甄試日
            Case Else
                Common.MessageBox(Me, "請選擇勾稽條件!!")
                Exit Sub
        End Select

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        '作業顯示模式：0:其他 1:模糊顯示 2:正常顯示
        'Dim v_rblWorkMode As String=TIMS.GetListValue(rblWorkMode)
        'ViewState(TIMS.gcst_rblWorkMode)=v_rblWorkMode
        ViewState(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip1

        '匯出e網民眾投保狀況檢核表
        '甄試學員投保狀況檢核表
        Dim flagNoExp1 As Boolean = True
        'Dim v_rblStExChoice1 As String=TIMS.GetListValue(rblStExChoice1)
        Select Case v_rblStExChoice1 'rblStExChoice1.SelectedValue
            Case "2" '開訓日
                flagNoExp1 = False
                Call ExpReport06_66_S(sm.UserInfo.TPlanID)
            Case "4" '開訓日前4日
                flagNoExp1 = False
                Call ExpReport06_66_S4(sm.UserInfo.TPlanID)
            Case "3" '甄試日 
                flagNoExp1 = False
                If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Call ExpReport70_E(sm.UserInfo.TPlanID)
                Else
                    Call ExpReport06_66_E(sm.UserInfo.TPlanID)
                End If
        End Select
        If flagNoExp1 Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub
End Class