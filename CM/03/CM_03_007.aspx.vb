Partial Class CM_03_007
    Inherits AuthBasePage

    '職前班使用
    'ReportQuery
    '增加身分別 須修改報表
    'CM_03_007_1
    'AND cs.MAKESOCID IS NULL

    'dbo.fn_GET_MIdentityID1
    '增加身分別 使用程式代入 Identity = TIMS.Get_Identity(Identity, 6)
    'CM_03_007_1 @TR
    'CM_03_007_2 @TR '年齡
    'CM_03_007_3 @TR
    'CM_03_007_4 @TR
    'CM_03_007_5 @TR '性別
    'CM_03_007_6 @TR
    'CM_03_007_7 @TR '縣市別

    'CM_03_007_*.jrxml 
    'Const cst_CM_03_007_1 As String = "CM_03_007_1" '身分別
    'Const cst_CM_03_007_2 As String = "CM_03_007_2" '年齡
    'Const cst_CM_03_007_3 As String = "CM_03_007_3" '訓練職類
    'Const cst_CM_03_007_4 As String = "CM_03_007_4" '教育程度
    'Const cst_CM_03_007_5 As String = "CM_03_007_5" '性別
    'Const cst_CM_03_007_6 As String = "CM_03_007_6" '通俗職類
    'Const cst_CM_03_007_7 As String = "CM_03_007_7" '縣市別 (目前無此選項。)

    'Const cst_CM_03_007_1 As String = "CM_03_007_1_b" '身分別
    Const cst_CM_03_007_1 As String = "CM_03_007_c_1" '身分別
    Const cst_CM_03_007_2 As String = "CM_03_007_2_b" '年齡
    Const cst_CM_03_007_3 As String = "CM_03_007_3_b" '訓練職類
    Const cst_CM_03_007_4 As String = "CM_03_007_4_b" '教育程度
    Const cst_CM_03_007_5 As String = "CM_03_007_5_b" '性別
    Const cst_CM_03_007_6 As String = "CM_03_007_6_b" '通俗職類
    'Const cst_CM_03_007_7 As String = "CM_03_007_7_b" '縣市別 (目前無此選項。)
    Const cst_CM_03_007_8 As String = "CM_03_007_8_b" '就職狀況
    Const cst_CM_03_007_11 As String = "CM_03_007_11_b" '年齡2

    Const cst_rblST1_AND As String = "AND"
    Const cst_rblST1_OR As String = "OR"

    Dim gExportStr As String = ""

    'SELECT * FROM KEY_IDENTITY WHERE IDENTITYID IN ('01','02','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42')  ORDER BY 1
    'SELECT * FROM KEY_IDENTITY WHERE NAME LIKE '%生%' ORDER BY 1
    '"'01','02','03','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36'"
    '主要特定對象統計表
    'Const cst_UseIdentityID As String = "'01','02','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42'"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()
        End If

        'IdentityTR.Style("display") = "inline"
        IdentityTR.Style("display") = TIMS.cst_inline1
        If rblMode1.SelectedIndex = 0 Then
            IdentityTR.Style("display") = "none"
        End If
    End Sub

    Sub CreateItem()
        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        '年度
        Syear = TIMS.GetSyear(Syear)
        'Common.SetListItem(Syear, Now.Year)
        Common.SetListItem(Syear, sm.UserInfo.Years)

        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))
        '計畫
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        '預算來源
        'BudgetList = TIMS.Get_Budget(BudgetList, 33, objconn)
        BudgetList = TIMS.Get_Budget(BudgetList, 3, objconn)
        '身分別  '顯示的身分別
        '03"負擔家計婦女"併入28"獨立負擔家計者"計算,並把"負擔家計婦女"項目拿掉.
        'CM_03_007 (報表)
        Identity = TIMS.Get_Identity(Identity, 66, objconn)
        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        '選擇全部身分別
        Identity.Attributes("onclick") = "SelectAll('Identity','Identity_List');"
        '列印檢查
        Print.Attributes("onclick") = "javascript:return CheckPrint();"
        '如果統計項目的選項改變
        rblMode1.Attributes("onclick") = "ChangeMode();"
    End Sub

    '列印 (iReport)
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Dim Identity1 As String = ""
        Dim DistID1 As String = ""
        Dim TPlanID1 As String = ""
        Dim BudgetID As String = ""
        '取得輸入參數
        Call Get_MySelectValue(Identity1, DistID1, TPlanID1, BudgetID)

        '報表要用的標題轄區參數
        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&STTDate=" & Me.STDate1.Text
        MyValue &= "&FTTDate=" & Me.STDate2.Text
        MyValue &= "&SFTDate=" & Me.FTDate1.Text
        MyValue &= "&FFTDate=" & Me.FTDate2.Text
        Select Case rblMode1.SelectedIndex
            Case 0
                '排除 MyValue &= "&Identity=" & Identity1
            Case Else
                MyValue &= "&Identity=" & Identity1
        End Select
        MyValue &= "&DistID=" & DistID1
        MyValue &= "&TPlanID=" & TPlanID1
        MyValue &= "&BudgetID=" & BudgetID
        MyValue &= "&Years=" & Syear.SelectedValue
        Select Case rblSchType1.SelectedValue
            Case cst_rblST1_AND '"AND"
            Case cst_rblST1_OR '"OR"
                MyValue &= "&ANDOR1=1"
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        'CM_03_007_*.jrxml
        Dim sFileName As String = ""
        sFileName = ""
        Select Case rblMode1.SelectedIndex
            Case 0 '身分別
                sFileName = cst_CM_03_007_1 '"CM_03_007_1"
            Case 1 '年齡
                sFileName = cst_CM_03_007_2 '"CM_03_007_2"
            Case 7 '7 '就職狀況(Excel匯出)
                'sFileName = cst_CM_03_007_11 'CM_03_007_8就職狀況(Excel匯出)
                sFileName = TIMS.cst_NO
            Case 2 '訓練職類
                sFileName = cst_CM_03_007_3 '"CM_03_007_3"
            Case 3 '教育程度
                sFileName = cst_CM_03_007_4 '"CM_03_007_4"
            Case 4 '性別
                sFileName = cst_CM_03_007_5 '"CM_03_007_5" '性別
            Case 5 '通俗職類
                sFileName = cst_CM_03_007_6 '"CM_03_007_6" '通俗職類
                'Case 6 '縣市別(Excel匯出)
                'sFileName = cst_CM_03_007_7 'CM_03_007_7縣市別(Excel匯出)
                'Call sUtl_ExpRptExcel(cst_CM_03_007_7, MyValue) '縣市別
                'Exit Sub
            Case 6 '7 '就職狀況(Excel匯出)
                sFileName = cst_CM_03_007_8 'CM_03_007_8就職狀況(Excel匯出)
        End Select
        If (sFileName = TIMS.cst_NO) Then
            Common.MessageBox(Me, "該統計項目-暫無報表!!")
            Exit Sub
        End If
        If (sFileName = "") Then
            Common.MessageBox(Me, "請確認統計項目!!")
            Exit Sub
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sFileName, MyValue)

        'Select Case rblMode1.SelectedIndex
        '    Case 0, 1, 2, 3, 4, 5
        'End Select

    End Sub

    '取得輸入參數
    Sub Get_MySelectValue(ByRef Identity1 As String, ByRef DistID1 As String, ByRef TPlanID1 As String, ByRef BudgetID As String)
        '報表要用的身分別參數
        'Dim Identity1 As String = ""
        Identity1 = ""
        For i As Integer = 1 To Identity.Items.Count - 1
            If Identity.Items(i).Selected Then
                If Identity1 <> "" Then Identity1 &= ","
                Identity1 &= "\'" & Me.Identity.Items(i).Value & "\'"
            End If
        Next

        If Identity1 <> "" Then
            If Identity1.IndexOf("28") > -1 Then
                '身分別  '顯示的身分別
                '03"負擔家計婦女"併入28"獨立負擔家計者"計算,並把"負擔家計婦女"項目拿掉.
                'CM_03_007 (報表)
                '補 03
                Identity1 += "," & Convert.ToString("\'03\'")
            End If
        End If

        '報表要用的轄區參數
        'Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        'Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)
        '報表要用的預算來源參數
        'Dim BudgetID As String = ""
        BudgetID = TIMS.GetCheckBoxListRptVal(BudgetList, 0)
    End Sub

    '匯出SUB (SQL)
    'Sub sUtl_ExpRptExcel(ByVal sFileName As String, ByVal strSession As String)
    '    Call Exp1(sFileName, strSession)
    'End Sub

    '鎖定計畫範圍。
    Public Shared Function Get_KeyPlanDt(ByVal ssTPlanID As String, ByVal oConn As SqlConnection) As DataTable
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT TPlanID,PlanName"
        sql &= " FROM KEY_PLAN"
        sql &= " WHERE 1=1" & vbCrLf
        If ssTPlanID <> "" Then
            sql &= " AND TPlanID IN (" & ssTPlanID.Replace("\'", "'") & ")" & vbCrLf
        End If
        sql &= " ORDER BY TPlanID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn)
        Return dt
    End Function

    'GROUP 群組分類
    Public Shared Function Get_MainDt2(ByVal MyPage As Page, ByVal vMod1 As String, ByRef strTitle2 As String, ByRef XID_NAME As String, ByVal oConn As SqlConnection) As DataTable
        Dim sqlX As String = ""
        Select Case vMod1 'rblMode1.SelectedIndex
            Case 0 '身分別 IDENTITYID
                sqlX = "" & vbCrLf
                sqlX &= " SELECT '0'+X.XID SORTX" & vbCrLf
                sqlX &= " ,X.XID,X.XNAME" & vbCrLf
                sqlX &= " FROM (" & vbCrLf
                'cst_IdentityM1
                '01','02','03','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42'
                'sqlX &= " SELECT DISTINCT CASE WHEN IDENTITYID IN ('01','02','03','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42') THEN IDENTITYID" & vbCrLf
                'sqlX &= " ,CASE WHEN IDENTITYID IN ('01','02','03','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42') THEN CONVERT(varchar, NAME)" & vbCrLf
                sqlX &= " SELECT DISTINCT CASE WHEN IDENTITYID IN (" & TIMS.cst_IdentityM1 & ") THEN IDENTITYID" & vbCrLf
                sqlX &= " ELSE '99' END XID" & vbCrLf
                sqlX &= " ,CASE WHEN IDENTITYID IN (" & TIMS.cst_IdentityM1 & ") THEN CONVERT(NVARCHAR(30), NAME)" & vbCrLf
                sqlX &= " ELSE '其他' END XNAME" & vbCrLf
                sqlX &= " FROM KEY_IDENTITY" & vbCrLf
                sqlX &= " ) X" & vbCrLf
                sqlX &= " UNION SELECT '9-1','-1','無' " & vbCrLf 'SORTX,XID,XNAME
                sqlX &= " ORDER BY 1" & vbCrLf

                XID_NAME = "MidentityId"
            Case 1 '年齡
                strTitle2 = "－依年齡" '"CM_03_007_2" 'SELECT * FROM V_YEARSOLD2B ORDER BY 1
                sqlX = "SELECT XID,XNAME FROM V_YEARSOLD2B ORDER BY XID" 'dbo.fn_YEARSOLDID2B old@dbo.fn_YEARSOLDID 
                XID_NAME = "YEARSOLDID"

            Case 7 '年齡2
                strTitle2 = "－依年齡" '"CM_03_007_2" 'SELECT * FROM V_YEARSOLD2B ORDER BY 1
                sqlX = "SELECT XID,XNAME FROM V_YEARSOLD2E ORDER BY XID" 'dbo.fn_YEARSOLDID2B old@dbo.fn_YEARSOLDID 
                XID_NAME = "YEARSOLDID2E"

            Case 2 '訓練職類
                strTitle2 = "－依訓練職類" '"CM_03_007_3"
                sqlX = "SELECT DISTINCT BUSID XID,BUSNAME XNAME FROM KEY_TRAINTYPE ORDER BY 1"
                XID_NAME = "BusID"
            Case 3 '教育程度
                strTitle2 = "－依教育程度" '"CM_03_007_4"
                sqlX = "SELECT DEGREEID XID,NAME XNAME FROM KEY_DEGREE WHERE DEGREEID<'14' ORDER BY 1"
                XID_NAME = "DegreeID"
            Case 4 '性別
                strTitle2 = "－依性別" '"CM_03_007_5" '性別
                sqlX = "select SEXID XID , CNAME XNAME FROM V_SEX ORDER BY 1 DESC"
                XID_NAME = "SEX"
            Case 5 '通俗職類
                strTitle2 = "－依通俗職類" '"CM_03_007_6" '通俗職類

                '啟用2016年通俗職類
                Dim flag_Cjob2016 As Boolean = TIMS.Get_sCjob2016_USE(MyPage)
                Dim str_SHARECJOB_YEAR As String = ""
                If flag_Cjob2016 Then str_SHARECJOB_YEAR = TIMS.cst_SHARE_CJOB_2016

                Select Case str_SHARECJOB_YEAR
                    Case TIMS.cst_SHARE_CJOB_2016
                        sqlX = ""
                        sqlX &= " SELECT dbo.LPAD(CJOB_TYPE,2,'0') SORT, CJOB_TYPE XID,CJOB_NAME XNAME FROM SHARE_CJOB WHERE CJOB_NO IS NULL "
                        sqlX &= " AND CYEARS ='2019' "
                        sqlX &= " ORDER BY 1" '排除999

                    Case Else
                        sqlX = ""
                        sqlX &= " SELECT dbo.LPAD(CJOB_TYPE,2,'0') SORT, CJOB_TYPE XID,CJOB_NAME XNAME FROM SHARE_CJOB WHERE CJOB_NO IS NULL "
                        sqlX &= " AND CYEARS ='2014' "
                        sqlX &= " ORDER BY 1" '排除999
                End Select

                XID_NAME = "CJOBTYPE"
                'Case 6 '縣市別(Excel匯出)
                '    strTitle2 = "－依縣市別" '"CM_03_007_7 縣市別(Excel匯出)
                '    sqlX = "select CTID XID,CTName XNAME FROM ID_CITY WHERE CTID <100 ORDER BY CTID" '排除999
                '    XID_NAME = "CTID"

            Case 6 '就職狀況(Excel匯出)
                strTitle2 = "－就職狀況" '"CM_03_007_8 就職狀況(Excel匯出)
                sqlX = "SELECT JOBSTATE XID,JNAME XNAME FROM V_JOBSTATE ORDER BY JOBSTATE" '排除999
                XID_NAME = "JOBSTATE"
        End Select
        Dim dtX As DataTable
        dtX = DbAccess.GetDataTable(sqlX, oConn)
        Return dtX
    End Function

    '主要查詢資料
    Public Shared Function Get_MainDt1(ByVal sFileName As String, ByVal strSession As String, ByVal vSchType1 As String, ByVal oConn As SqlConnection) As DataTable
        Dim ssSTDate1 As String = TIMS.GetMyValue(strSession, "STTDate")
        Dim ssSTDate2 As String = TIMS.GetMyValue(strSession, "FTTDate")
        Dim ssFTDate1 As String = TIMS.GetMyValue(strSession, "SFTDate")
        Dim ssFTDate2 As String = TIMS.GetMyValue(strSession, "FFTDate")
        Dim ssIdentity As String = TIMS.GetMyValue(strSession, "Identity")

        Dim ssDistID As String = TIMS.GetMyValue(strSession, "DistID")
        Dim ssTPlanID As String = TIMS.GetMyValue(strSession, "TPlanID")
        Dim ssBudgetID As String = TIMS.GetMyValue(strSession, "BudgetID")
        Dim ssYears As String = TIMS.GetMyValue(strSession, "Years")

        Dim strUseIdentityID As String = TIMS.cst_IdentityM1

        'cst_UseIdentityID 
        'Dim dt_KeyPlan As DataTable
        '身分別 CM_03_007_1
        'sql += " ,ISNULL( ky.MergeID, ky.IdentityID) MidentityId" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(ky.MergeID, ky.IdentityID) IN (" & strUseIdentityID & ") THEN ISNULL(ky.MergeID, ky.IdentityID)" & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID ,cs.SOCID ,cs.SID" & vbCrLf
        sql &= " ,ip.TPlanID" & vbCrLf
        sql &= " ,CASE WHEN cs.MidentityId IN (" & strUseIdentityID & ") THEN cs.MidentityId" & vbCrLf
        sql &= "  WHEN cs.MidentityId IS NULL THEN '-1'" & vbCrLf
        sql &= "  ELSE '99' END MidentityId" & vbCrLf

        '年齡 CM_03_007_2
        'sql += " ,dbo.fn_YEARSOLDID(cc.FTDate,ss.Birthday) YEARSOLDID" & vbCrLf
        sql &= " ,dbo.FN_YEARSOLDID2B(cc.FTDate,ss.Birthday) YEARSOLDID" & vbCrLf
        '年齡2 CM_03_007_11
        sql &= " ,dbo.FN_YEARSOLDID2E(dbo.FN_YEARSOLD(cc.FTDate,ss.Birthday)) YEARSOLDID2E" & vbCrLf

        sql &= " ,dbo.FN_YEARSOLD(cc.FTDate,ss.Birthday) YEARSOLD" & vbCrLf
        'sql &= " ,TRUNC(dbo.MONTHS_BETWEEN(cc.FTDate,ss.Birthday)/12) YEARSOLD" & vbCrLf
        '訓練職類 CM_03_007_3
        sql &= " ,vt.BUSID" & vbCrLf
        '教育程度 CM_03_007_4
        sql &= " ,ss.DEGREEID" & vbCrLf
        '性別 CM_03_007_5
        sql &= " ,ss.SEX" & vbCrLf
        '通俗職類 CM_03_007_6
        sql &= " ,V.CJOB_TYPE CJOBTYPE" & vbCrLf
        '縣市別 (目前無此選項。) CM_03_007_7
        sql &= " ,cc.CTID" & vbCrLf
        sql &= " ,cc.CTNAME" & vbCrLf
        '就職狀況
        sql &= " ,ss.JOBSTATE" & vbCrLf

        Select Case vSchType1 'rblSchType1.SelectedValue
            Case cst_rblST1_AND '"AND"
                '/*開訓人數*/
                sql &= " ,1 OPENX" & vbCrLf
                '/*結訓人數 不含 提前就業人數*/
                sql &= " ,CASE WHEN cc.FTDate < GETDATE() AND cs.STUDSTATUS NOT IN (2,3) THEN 1 END finx" & vbCrLf
                '/*就業人數 不含 提前就業人數*/
                sql &= " ,CASE WHEN ISNULL(cs.WkAheadOfSch,' ')  !='Y' AND sg3.SOCID is not null" & vbCrLf
                sql &= " 	AND sg3.IsGetJob=1 AND cs.STUDSTATUS NOT IN (2,3) then 1 end injobx " & vbCrLf
            Case cst_rblST1_OR '"OR"
                '/*開訓人數*/
                sql &= " ,CASE WHEN 1=1" & vbCrLf
                If ssSTDate1 <> "" Then
                    sql &= " AND cc.STDate >= " & TIMS.To_date(ssSTDate1) & vbCrLf '" & ssSTDate1 & "'" & vbCrLf
                End If
                If ssSTDate2 <> "" Then
                    sql &= " AND cc.STDate <=" & TIMS.To_date(ssSTDate2) & vbCrLf '" & ssSTDate2 & "'" & vbCrLf
                End If
                sql &= " THEN 1 END OPENX" & vbCrLf
                '/*結訓人數 不含 提前就業人數*/
                sql &= " ,CASE WHEN cc.FTDate < GETDATE() AND cs.STUDSTATUS NOT IN (2,3) " & vbCrLf
                If ssFTDate1 <> "" Then
                    sql &= " AND cc.FTDate >=" & TIMS.To_date(ssFTDate1) & vbCrLf '" & ssFTDate1 & "'" & vbCrLf
                End If
                If ssFTDate2 <> "" Then
                    sql &= " AND cc.FTDate <=" & TIMS.To_date(ssFTDate2) & vbCrLf '" & ssFTDate2 & "'" & vbCrLf
                End If
                sql &= " THEN 1 END finx" & vbCrLf
                '/*就業人數 不含 提前就業人數*/
                sql &= " ,CASE WHEN ISNULL(cs.WkAheadOfSch,' ')  !='Y' AND sg3.SOCID is not null" & vbCrLf
                If ssFTDate1 <> "" Then
                    sql &= " AND cc.FTDate >=" & TIMS.To_date(ssFTDate1) & vbCrLf '" & ssFTDate1 & "'" & vbCrLf
                End If
                If ssFTDate2 <> "" Then
                    sql &= " AND cc.FTDate <=" & TIMS.To_date(ssFTDate2) & vbCrLf '" & ssFTDate2 & "'" & vbCrLf
                End If
                sql &= " 	AND sg3.IsGetJob=1 AND cs.STUDSTATUS NOT IN (2,3) then 1 end injobx " & vbCrLf
        End Select

        '/*提前就業人數*/
        sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' THEN 1 end wjobx " & vbCrLf
        '/*不就業人數*/
        sql &= " ,CASE WHEN ISNULL(cs.WkAheadOfSch,' ') !='Y' AND sg3.SOCID is not null" & vbCrLf
        sql &= "  AND sg3.IsGetJob=2 AND cs.STUDSTATUS NOT IN (2,3) then 1 end nojobx " & vbCrLf
        '在職者
        sql &= " ,CASE WHEN cs.WorkSuppIdent='Y' THEN 1 END WIdent" & vbCrLf
        '就業關聯性
        sql &= " ,CASE WHEN cs.STUDSTATUS NOT IN (2,3) AND sg3.JOBRELATE='Y' THEN 1 END JobRelNum" & vbCrLf
        '公法救助人數 (PUCount) PUBLICRESCUE='Y'
        sql &= " ,CASE WHEN cs.STUDSTATUS NOT IN (2,3) AND sg3.PUBLICRESCUE='Y' AND sg3.SOCID IS NOT NULL THEN 1 END PUCount" & vbCrLf
        sql &= " ,cc.Years,cc.DistName,cc.PlanName,cc.OrgName,cc.ClassCName" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDATE, 111) STDATE" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.FTDATE, 111) FTDATE" & vbCrLf
        sql &= " ,vt.BUSNAME" & vbCrLf
        sql &= " ,ss.name stdname" & vbCrLf
        sql &= " ,ss.idno" & vbCrLf
        sql &= " ,CONVERT(varchar, ss.BIRTHDAY, 111) BIRTHDAY" & vbCrLf
        sql &= " ,ISNULL(KY.NAME,'無') MIdentityN" & vbCrLf
        sql &= " ,cs.STUDSTATUS" & vbCrLf

        sql &= " FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.SID=cs.SID" & vbCrLf
        sql &= " JOIN VIEW2 cc on cc.OCID =cs.OCID" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.PlanID = cc.PlanID" & vbCrLf
        sql &= " JOIN VIEW_TRAINTYPE vt on vt.TMID=cc.TMID" & vbCrLf
        sql &= " JOIN SHARE_CJOB v on v.CJOB_UNKEY = cc.CJOB_UNKEY" & vbCrLf
        sql &= " LEFT JOIN KEY_IDENTITY ky on cs.MidentityId = ky.IdentityID" & vbCrLf
        sql &= " LEFT JOIN STUD_GETJOBSTATE3 sg3 on sg3.SOCID =cs.SOCID AND sg3.CPoint=1" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cs.MAKESOCID IS NULL" & vbCrLf
        sql &= " AND cc.NotOpen='N'" & vbCrLf
        sql &= " AND cc.IsSuccess='Y'" & vbCrLf
        'sql &= " AND cs.MIdentityID IS NOT NULL" & vbCrLf
        'sql += " AND cs.MIdentityID <> ''" & vbCrLf
        'Select Case rblSchType1.SelectedValue
        '    Case cst_rblST1_AND '"AND"
        '    Case cst_rblST1_OR '"OR"
        'End Select

        Select Case vSchType1 'rblSchType1.SelectedValue
            Case cst_rblST1_AND '"AND"
                If ssSTDate1 <> "" Then
                    sql &= " AND cc.STDate >= " & TIMS.To_date(ssSTDate1) & vbCrLf '" & ssSTDate1 & "'" & vbCrLf
                End If
                If ssSTDate2 <> "" Then
                    sql &= " AND cc.STDate <=" & TIMS.To_date(ssSTDate2) & vbCrLf '" & ssSTDate2 & "'" & vbCrLf
                End If
                If ssFTDate1 <> "" Then
                    sql &= " AND cc.FTDate >=" & TIMS.To_date(ssFTDate1) & vbCrLf '" & ssFTDate1 & "'" & vbCrLf
                End If
                If ssFTDate2 <> "" Then
                    sql &= " AND cc.FTDate <=" & TIMS.To_date(ssFTDate2) & vbCrLf '" & ssFTDate2 & "'" & vbCrLf
                End If

            Case cst_rblST1_OR '"OR"
                sql &= " AND (1!=1 OR (1=1" & vbCrLf
                If ssSTDate1 <> "" Then
                    sql &= " AND cc.STDate >= " & TIMS.To_date(ssSTDate1) & vbCrLf '" & ssSTDate1 & "'" & vbCrLf
                End If
                If ssSTDate2 <> "" Then
                    sql &= " AND cc.STDate <=" & TIMS.To_date(ssSTDate2) & vbCrLf '" & ssSTDate2 & "'" & vbCrLf
                End If
                sql &= " ) OR (1=1" & vbCrLf
                If ssFTDate1 <> "" Then
                    sql &= " AND cc.FTDate >=" & TIMS.To_date(ssFTDate1) & vbCrLf '" & ssFTDate1 & "'" & vbCrLf
                End If
                If ssFTDate2 <> "" Then
                    sql &= " AND cc.FTDate <=" & TIMS.To_date(ssFTDate2) & vbCrLf '" & ssFTDate2 & "'" & vbCrLf
                End If
                sql &= " ))" & vbCrLf
            Case Else
                'Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                'Exit Function
        End Select

        'strUseIdentityID
        If ssIdentity <> "" Then
            'https://jira.turbotech.com.tw/browse/TIMSC-268
            '增加無的選項
            sql &= " AND ISNULL(cs.MIdentityID,'-1') IN (" & ssIdentity.Replace("\'", "'") & ")" & vbCrLf
        End If
        If ssDistID <> "" Then
            sql &= " AND ip.DistID IN (" & ssDistID.Replace("\'", "'") & ")" & vbCrLf
        End If
        If ssTPlanID <> "" Then
            sql &= " AND ip.TPlanID IN (" & ssTPlanID.Replace("\'", "'") & ")" & vbCrLf
        End If
        If ssBudgetID <> "" Then
            'https://jira.turbotech.com.tw/browse/TIMSC-268
            '增加無的選項
            sql &= " AND ISNULL(cs.BudgetID,'-1') IN (" & ssBudgetID.Replace("\'", "'") & ")" & vbCrLf
        End If
        If ssYears <> "" Then
            sql &= " AND ip.Years ='" & ssYears & "'" & vbCrLf
        End If
        'Dim sCmd As New SqlCommand(sql, objconn)
        'Dim dt As New DataTable
        'With sCmd
        '    .Parameters.Clear()
        '    dt.Load(.ExecuteReader())
        'End With

        Dim da As New SqlDataAdapter
        da.SelectCommand = New SqlCommand
        With da.SelectCommand
            .Connection = oConn
            .CommandTimeout = 100
            .CommandText = sql
            .Parameters.Clear()
        End With
        Dim dt As New DataTable
        da.Fill(dt)
        Return dt
    End Function

    '匯出(1:統計 2:明細)
    Sub ExpX(ByVal iType As Integer)
        Dim Identity1 As String = ""
        Dim DistID1 As String = ""
        Dim TPlanID1 As String = ""
        Dim BudgetID As String = ""
        '取得輸入參數
        Call Get_MySelectValue(Identity1, DistID1, TPlanID1, BudgetID)

        '報表要用的標題轄區參數
        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&STTDate=" & Me.STDate1.Text
        MyValue &= "&FTTDate=" & Me.STDate2.Text
        MyValue &= "&SFTDate=" & Me.FTDate1.Text
        MyValue &= "&FFTDate=" & Me.FTDate2.Text
        Select Case rblMode1.SelectedIndex
            Case 0
                '排除 MyValue &= "&Identity=" & Identity1
            Case Else
                MyValue &= "&Identity=" & Identity1
        End Select
        MyValue &= "&DistID=" & DistID1
        MyValue &= "&TPlanID=" & TPlanID1
        MyValue &= "&BudgetID=" & BudgetID
        MyValue &= "&Years=" & Syear.SelectedValue
        Select Case rblSchType1.SelectedValue
            Case cst_rblST1_AND '"AND"
            Case cst_rblST1_OR '"OR"
                MyValue &= "&ANDOR1=1"
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        '結束狀況有誤
        Dim okFlag As Boolean = False
        Try
            Call TIMS.OpenDbConn(objconn)

            Dim sFileName As String = ""
            sFileName = ""
            Select Case rblMode1.SelectedIndex
                Case 0 '身分別
                    sFileName = cst_CM_03_007_1 '"CM_03_007_1"
                Case 1 '年齡
                    sFileName = cst_CM_03_007_2 '"CM_03_007_2"
                Case 2 '訓練職類
                    sFileName = cst_CM_03_007_3 '"CM_03_007_3"
                Case 3 '教育程度
                    sFileName = cst_CM_03_007_4 '"CM_03_007_4"
                Case 4 '性別
                    sFileName = cst_CM_03_007_5 '"CM_03_007_5" '性別
                Case 5 '通俗職類
                    sFileName = cst_CM_03_007_6 '"CM_03_007_6" '通俗職類
                Case 6 '縣市別(Excel匯出)
                    'sFileName = cst_CM_03_007_7 '縣市別(Excel匯出)
                Case 7 '就職狀況(Excel匯出)
                    sFileName = cst_CM_03_007_8 '"CM_03_007_8" '就職狀況
                Case 11 '年齡2
                    sFileName = cst_CM_03_007_11 '"CM_03_007_11"
            End Select

            Select Case iType
                Case 1
                    '匯出SUB'SQL
                    Call sUtl_ExpRptExcel1(sFileName, MyValue)  '匯出SUB'SQL
                Case 2
                    '匯出SUB'SQL
                    Call sUtl_ExpRptExcel2(sFileName, MyValue)  '匯出SUB'SQL
            End Select

            okFlag = True '結束狀況無誤
            Call TIMS.CloseDbConn(objconn)

        Catch ex As Exception
            'If conn.State = ConnectionState.Open Then conn.Close()
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)
            Exit Sub
        End Try

        '結束狀況無誤
        If okFlag Then
            Response.End()
        End If
    End Sub

    '匯出 Excel- '匯出統計資料
    Sub sUtl_ExpRptExcel1(ByVal sFileName As String, ByVal strSession As String)
        'Dim XID_NAME As String = "" '班級學員使用欄位名稱
        'Dim sqlX As String = ""
        'Dim dtX As DataTable

        'SELECT * FROM KEY_IDENTITY WHERE IDENTITYID IN ('01','02','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42')  ORDER BY 1
        'SELECT * FROM KEY_IDENTITY WHERE NAME LIKE '%生%' ORDER BY 1
        '"'01','02','03','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36'"
        'Const cst_UseIdentityID As String = "'01','02','04','05','06','07','09','10','13','14','17','20','27','28','29','33','35','36','42'"
        '身分別問題 (其他為:99)
        'https://jira.turbotech.com.tw/browse/TIMSC-226

        'dt.DefaultView.Sort = "DistName,ClassID,orgname,OrgTypeName,ClassName,CyclType"
        'dt = TIMS.dv2dt(dt.DefaultView)
        'If dt.Rows.Count = 0 Then
        '    Common.MessageBox(Me, "目前條件查無資料!!")
        '    Exit Sub
        'End If

        Const cst_iColspanNUM As Integer = 2 '10 '每組固定數量。

        Dim ssSTDate1 As String = TIMS.GetMyValue(strSession, "STTDate")
        Dim ssSTDate2 As String = TIMS.GetMyValue(strSession, "FTTDate")
        Dim ssFTDate1 As String = TIMS.GetMyValue(strSession, "SFTDate")
        Dim ssFTDate2 As String = TIMS.GetMyValue(strSession, "FFTDate")
        Dim ssIdentity As String = TIMS.GetMyValue(strSession, "Identity")
        Dim ssDistID As String = TIMS.GetMyValue(strSession, "DistID")
        Dim ssTPlanID As String = TIMS.GetMyValue(strSession, "TPlanID")
        Dim ssBudgetID As String = TIMS.GetMyValue(strSession, "BudgetID")
        Dim ssYears As String = TIMS.GetMyValue(strSession, "Years")

        'Dim sql As String = ""
        Dim dt As DataTable = Get_MainDt1(sFileName, strSession, rblSchType1.SelectedValue, objconn)
        '鎖定計畫範圍。
        Dim dt_KeyPlan As DataTable = Get_KeyPlanDt(ssTPlanID, objconn)

        'Dim cst_sTitle As String = "主要對象統計表"
        Dim strTitle2 As String = "" '依抬頭
        Dim XID_NAME As String = "" '班級學員使用欄位名稱
        Dim dtX As DataTable = Get_MainDt2(Me, rblMode1.SelectedIndex, strTitle2, XID_NAME, objconn)

        Dim sFileName1 As String = "主要對象統計表"

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= (".noDecFormat2{mso-number-format:""0.00%"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim MyValue As String = ""
        Dim colspan As Integer = 0
        Dim ExportStr As String = ""
        ExportStr = ""
        ExportStr &= "<tr>"
        colspan += 1
        For i As Integer = 0 To dtX.Rows.Count - 1
            colspan += cst_iColspanNUM
        Next
        colspan += cst_iColspanNUM
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & sFileName1 & strTitle2 & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "年度：" & ssYears & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "開訓期間：" & ssSTDate1 & "~" & ssSTDate2 & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "查詢方式：" & rblSchType1.SelectedItem.Text & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "結訓期間：" & ssFTDate1 & "~" & ssFTDate2 & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "轄區：" & TIMS.GET_DISTNAME(objconn, ssDistID.Replace("\'", "'")) & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "預算別：" & TIMS.GET_BudgetName(ssBudgetID.Replace("\'", ""), objconn) & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=""" & CStr(colspan) & """>" & "身分別：" & TIMS.Get_IdentityName(ssIdentity.Replace("\'", ""), objconn) & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td rowspan=""2"">訓練計畫</td>" & vbTab
        For Each dr As DataRow In dtX.Rows
            ExportStr &= "<td colspan=""" & cst_iColspanNUM & """>" & dr("XNAME").ToString & "</td>" & vbTab
        Next
        ExportStr &= "<td colspan=""" & cst_iColspanNUM & """>" & "合計" & "</td>" & vbTab
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        ExportStr = ""
        ExportStr &= "<tr>"
        'ExportStr &= "<td rowspan=""2"">訓練計畫</td>" & vbTab
        For i As Integer = 0 To dtX.Rows.Count - 1
            '每組固定數量。
            'Dim dr As DataRow = dt_CITY.Rows(i)
            '開訓人數、結訓人數、提前就業、不就業、在職者、就業人數、就業率
            ExportStr &= "<td>開訓人數</td>" & vbTab '開訓人數
            ExportStr &= "<td>結訓人數</td>" & vbTab '結訓人數 不含 提前就業人數
            'ExportStr &= "<td>提前就業人數</td>" & vbTab '開訓人數
            'ExportStr &= "<td>不就業人數</td>" & vbTab '不就業人數
            'ExportStr &= "<td>在職者</td>" & vbTab '在職者
            'ExportStr &= "<td>就業人數</td>" & vbTab '就業人數 不含 提前就業人數
            'ExportStr &= "<td>公法救助人數</td>" & vbTab '公法救助人數 (PUCount) PUBLICRESCUE='Y'
            'ExportStr &= "<td>就業率</td>" & vbTab '就業率
            'ExportStr &= "<td>就業關聯人數</td>" & vbTab '就業關聯人數 不含 離退人數
            'ExportStr &= "<td>就業關聯率</td>" & vbTab '就業關聯率
        Next
        '合計
        ExportStr &= "<td>開訓人數</td>" & vbTab '開訓人數
        ExportStr &= "<td>結訓人數</td>" & vbTab '結訓人數 不含 提前就業人數
        'ExportStr &= "<td>提前就業人數</td>" & vbTab '開訓人數
        'ExportStr &= "<td>不就業人數</td>" & vbTab '不就業人數
        'ExportStr &= "<td>在職者</td>" & vbTab '在職者
        'ExportStr &= "<td>就業人數</td>" & vbTab '就業人數 不含 提前就業人數
        'ExportStr &= "<td>公法救助人數</td>" & vbTab '公法救助人數 (PUCount) PUBLICRESCUE='Y'
        'ExportStr &= "<td>就業率</td>" & vbTab '就業率
        'ExportStr &= "<td>就業關聯人數</td>" & vbTab '就業關聯人數 不含 離退人數
        'ExportStr &= "<td>就業關聯率</td>" & vbTab '就業關聯率

        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        Dim sFilter As String = ""
        Dim openx As Integer = 0 '開訓人數
        Dim finx As Integer = 0 '結訓人數
        'Dim wjobx As Integer = 0 '提前就業人數
        'Dim nojobx As Integer = 0 '不就業人數
        'Dim WIdent As Integer = 0 '在職者(WorkSuppIdent)
        'Dim injobx As Integer = 0 '就業人數
        'Dim iPUCount As Integer = 0 '公法救助人數 (PUCount) PUBLICRESCUE='Y'
        'Dim iJobRelNum As Integer = 0 '就業關聯人數

        Dim iAopenx As Integer = 0 '開訓人數(A)
        Dim iAfinx As Integer = 0 '結訓人數(A)
        'Dim iAwjobx As Integer = 0 '提前就業人數(A)
        'Dim iAnojobx As Integer = 0 '不就業人數(A)
        'Dim iAWIdent As Integer = 0 '在職者(WorkSuppIdent)(A)
        'Dim iAnjobx As Integer = 0 '就業人數(A)
        'Dim iAPUCount As Integer = 0 '公法救助人數 (PUCount) PUBLICRESCUE='Y'
        'Dim iAJobRelNum As Integer = 0 '就業關聯人數(A)

        'Dim jRate1 As Double = 0 '就業率1
        'Dim jRate2 As Double = 0 '就業率2
        'Dim irJobRelRate As Double = 0 '就業關聯率

        'else (isnull(sum(CASE WHEN g.x1=1 AND (g.injobx=1 OR g.wjobx=1) then 1 end),0)+.0) / isnull(sum(CASE WHEN g.x1=1 AND g.finx=1 then 1 end),0) end AS JOB01
        For i1 As Integer = 0 To dt_KeyPlan.Rows.Count - 1
            ExportStr = ""
            ExportStr &= "<tr>"
            Dim dr1 As DataRow = dt_KeyPlan.Rows(i1)
            ExportStr &= "<td>" & dr1("PlanName") & "</td>" & vbTab '計畫名稱
            For i2 As Integer = 0 To dtX.Rows.Count - 1
                Dim dr2 As DataRow = dtX.Rows(i2)
                sFilter = ""
                sFilter &= " TPlanID='" & Convert.ToString(dr1("TPlanID")) & "'"
                sFilter &= " AND " & XID_NAME & "='" & Convert.ToString(dr2("XID")) & "'" '各選項
                openx = dt.Select(sFilter & " AND openx=1").Length
                finx = dt.Select(sFilter & " AND finx=1").Length
                'wjobx = dt.Select(sFilter & " AND wjobx=1").Length
                'nojobx = dt.Select(sFilter & " AND nojobx=1").Length
                'injobx = dt.Select(sFilter & " AND injobx=1").Length
                'iPUCount = dt.Select(sFilter & " AND PUCount=1").Length

                'WIdent = dt.Select(sFilter & " AND WIdent=1").Length
                'iJobRelNum = dt.Select(sFilter & " AND JobRelNum=1").Length
                'jRate1 = 0
                'If (finx + wjobx - nojobx - iPUCount) <> 0 Then
                '    jRate1 = TIMS.Round(CDbl(injobx + wjobx - iPUCount) / CDbl(finx + wjobx - nojobx - iPUCount) * 100, 2)
                'End If
                'jRate2 = 0
                'If (finx + wjobx - WIdent - iPUCount) <> 0 Then
                '    jRate2 = TIMS.Round(CDbl(injobx + wjobx - iPUCount) / CDbl(finx + wjobx - WIdent - iPUCount) * 100, 2)
                'End If
                'irJobRelRate = 0
                'If (finx) <> 0 Then
                '    irJobRelRate = TIMS.Round(CDbl(iJobRelNum) / CDbl(finx) * 100, 2)
                'End If

                ExportStr &= "<td class=""noDecFormat"">" & openx & "</td>" & vbTab '開訓人數
                ExportStr &= "<td class=""noDecFormat"">" & finx & "</td>" & vbTab '結訓人數 不含 提前就業人數
                'ExportStr &= "<td class=""noDecFormat"">" & wjobx & "</td>" & vbTab '提前就業人數
                'ExportStr &= "<td class=""noDecFormat"">" & nojobx & "</td>" & vbTab '不就業人數
                'ExportStr &= "<td class=""noDecFormat"">" & WIdent & "</td>" & vbTab '在職者
                'ExportStr &= "<td class=""noDecFormat"">" & injobx & "</td>" & vbTab '就業人數 不含 提前就業人數
                'ExportStr &= "<td class=""noDecFormat"">" & iPUCount & "</td>" & vbTab '公法救助人數 (PUCount) PUBLICRESCUE='Y'

                'ExportStr &= "<td class=""noDecFormat2"">" & CStr(jRate1) & "%</td>" & vbTab  '就業率
                'ExportStr &= "<td class=""noDecFormat"">" & iJobRelNum & "</td>" & vbTab '就業關聯人數 不含 離退人數
                'ExportStr &= "<td class=""noDecFormat2"">" & CStr(irJobRelRate) & "%</td>" & vbTab  '就業關聯率
            Next
            'ExportStr &= "<td rowspan=""6"">" & "合計" & "</td>" & vbTab

            '每計畫合計 總合顯示
            sFilter = ""
            sFilter &= " TPlanID='" & Convert.ToString(dr1("TPlanID")) & "'"

            openx = dt.Select(sFilter & " AND openx=1").Length
            finx = dt.Select(sFilter & " AND finx=1").Length
            'wjobx = dt.Select(sFilter & " AND wjobx=1").Length
            'nojobx = dt.Select(sFilter & " AND nojobx=1").Length
            'WIdent = dt.Select(sFilter & " AND WIdent=1").Length
            'injobx = dt.Select(sFilter & " AND injobx=1").Length
            'iPUCount = dt.Select(sFilter & " AND PUCount=1").Length

            'iJobRelNum = dt.Select(sFilter & " AND JobRelNum=1").Length

            'jRate1 = 0
            'If (finx + wjobx - nojobx - iPUCount) <> 0 Then
            '    jRate1 = TIMS.Round(CDbl(injobx + wjobx - iPUCount) / CDbl(finx + wjobx - nojobx - iPUCount) * 100, 2)
            'End If
            'jRate2 = 0
            'If (finx + wjobx - WIdent - iPUCount) <> 0 Then
            '    jRate2 = TIMS.Round(CDbl(injobx + wjobx - iPUCount) / CDbl(finx + wjobx - WIdent - iPUCount) * 100, 2)
            'End If
            'irJobRelRate = 0
            'If (finx) <> 0 Then
            '    irJobRelRate = TIMS.Round(CDbl(iJobRelNum) / CDbl(finx) * 100, 2)
            'End If

            iAopenx += openx
            iAfinx += finx
            'iAwjobx += wjobx
            'iAnojobx += nojobx
            'iAWIdent += WIdent
            'iAnjobx += injobx
            'iAPUCount += iPUCount

            'iAJobRelNum += iJobRelNum

            ExportStr &= "<td class=""noDecFormat"">" & openx & "</td>" & vbTab '開訓人數
            ExportStr &= "<td class=""noDecFormat"">" & finx & "</td>" & vbTab '結訓人數 不含 提前就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & wjobx & "</td>" & vbTab '提前就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & nojobx & "</td>" & vbTab '不就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & WIdent & "</td>" & vbTab '在職者
            'ExportStr &= "<td class=""noDecFormat"">" & injobx & "</td>" & vbTab '就業人數 不含 提前就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & iPUCount & "</td>" & vbTab '公法救助人數 (PUCount) PUBLICRESCUE='Y'

            'ExportStr &= "<td class=""noDecFormat2"">" & CStr(jRate1) & "%</td>" & vbTab  '就業率
            'ExportStr &= "<td class=""noDecFormat"">" & iJobRelNum & "</td>" & vbTab '就業關聯人數 不含 離退人數
            'ExportStr &= "<td class=""noDecFormat2"">" & CStr(irJobRelRate) & "%</td>" & vbTab  '就業關聯率

            ExportStr &= "</tr>"
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next

        '最後1筆　合計
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td>合計</td>" & vbTab '計畫名稱
        For i2 As Integer = 0 To dtX.Rows.Count - 1
            Dim dr2 As DataRow = dtX.Rows(i2)
            sFilter = ""
            'sFilter &= " TPlanID='" & Convert.ToString(dr1("TPlanID")) & "'"
            sFilter &= " " & XID_NAME & "='" & Convert.ToString(dr2("XID")) & "'"
            openx = dt.Select(sFilter & " AND openx=1").Length
            finx = dt.Select(sFilter & " AND finx=1").Length
            'wjobx = dt.Select(sFilter & " AND wjobx=1").Length
            'nojobx = dt.Select(sFilter & " AND nojobx=1").Length
            'WIdent = dt.Select(sFilter & " AND WIdent=1").Length
            'injobx = dt.Select(sFilter & " AND injobx=1").Length
            'iPUCount = dt.Select(sFilter & " AND PUCount=1").Length

            'iJobRelNum = dt.Select(sFilter & " AND JobRelNum=1").Length

            'jRate1 = 0
            'If (finx + wjobx - nojobx - iPUCount) <> 0 Then
            '    jRate1 = TIMS.Round(CDbl(injobx + wjobx - iPUCount) / CDbl(finx + wjobx - nojobx - iPUCount) * 100, 2)
            'End If
            'jRate2 = 0
            'If (finx + wjobx - WIdent - iPUCount) <> 0 Then
            '    jRate2 = TIMS.Round(CDbl(injobx + wjobx - iPUCount) / CDbl(finx + wjobx - WIdent - iPUCount) * 100, 2)
            'End If
            'irJobRelRate = 0
            'If (finx) <> 0 Then
            '    irJobRelRate = TIMS.Round(CDbl(iJobRelNum) / CDbl(finx) * 100, 2)
            'End If

            ExportStr &= "<td class=""noDecFormat"">" & openx & "</td>" & vbTab '開訓人數
            ExportStr &= "<td class=""noDecFormat"">" & finx & "</td>" & vbTab '結訓人數 不含 提前就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & wjobx & "</td>" & vbTab '提前就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & nojobx & "</td>" & vbTab '不就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & WIdent & "</td>" & vbTab '在職者
            'ExportStr &= "<td class=""noDecFormat"">" & injobx & "</td>" & vbTab '就業人數 不含 提前就業人數
            'ExportStr &= "<td class=""noDecFormat"">" & iPUCount & "</td>" & vbTab '公法救助人數 (PUCount) PUBLICRESCUE='Y'

            'ExportStr &= "<td class=""noDecFormat2"">" & CStr(jRate1) & "%</td>" & vbTab  '就業率1
            'ExportStr &= "<td class=""noDecFormat"">" & iJobRelNum & "</td>" & vbTab '就業關聯人數 不含 離退人數
            'ExportStr &= "<td class=""noDecFormat2"">" & CStr(irJobRelRate) & "%</td>" & vbTab  '就業關聯率
        Next

        'jRate1 = 0
        'If (iAfinx + iAwjobx - iAnojobx - iAPUCount) <> 0 Then
        '    jRate1 = TIMS.Round(CDbl(iAnjobx + iAwjobx - iAPUCount) / CDbl(iAfinx + iAwjobx - iAnojobx - iAPUCount) * 100, 2)
        'End If
        'jRate2 = 0
        'If (iAfinx + iAwjobx - iAWIdent - iAPUCount) <> 0 Then
        '    jRate2 = TIMS.Round(CDbl(iAnjobx + iAwjobx - iAPUCount) / CDbl(iAfinx + iAwjobx - iAWIdent - iAPUCount) * 100, 2)
        'End If
        'irJobRelRate = 0
        'If (finx) <> 0 Then
        '    irJobRelRate = TIMS.Round(CDbl(iAJobRelNum) / CDbl(iAfinx) * 100, 2)
        'End If

        ExportStr &= "<td class=""noDecFormat"">" & iAopenx & "</td>" & vbTab '開訓人數
        ExportStr &= "<td class=""noDecFormat"">" & iAfinx & "</td>" & vbTab '結訓人數 不含 提前就業人數
        'ExportStr &= "<td class=""noDecFormat"">" & iAwjobx & "</td>" & vbTab '提前就業人數
        'ExportStr &= "<td class=""noDecFormat"">" & iAnojobx & "</td>" & vbTab '不就業人數
        'ExportStr &= "<td class=""noDecFormat"">" & iAWIdent & "</td>" & vbTab '在職者
        'ExportStr &= "<td class=""noDecFormat"">" & iAnjobx & "</td>" & vbTab '就業人數 不含 提前就業人數
        'ExportStr &= "<td class=""noDecFormat"">" & iAPUCount & "</td>" & vbTab '公法救助人數 (PUCount) PUBLICRESCUE='Y'

        'ExportStr &= "<td class=""noDecFormat2"">" & CStr(jRate1) & "%</td>" & vbTab  '就業率1
        'ExportStr &= "<td class=""noDecFormat"">" & iAJobRelNum & "</td>" & vbTab '就業關聯人數 不含 離退人數
        'ExportStr &= "<td class=""noDecFormat2"">" & CStr(irJobRelRate) & "%</td>" & vbTab  '就業關聯率

        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

    '匯出 Excel- '匯出班級明細資料
    Sub sUtl_ExpRptExcel2(ByVal sFileName As String, ByVal strSession As String)
        Dim ssSTDate1 As String = TIMS.GetMyValue(strSession, "STTDate")
        Dim ssSTDate2 As String = TIMS.GetMyValue(strSession, "FTTDate")
        Dim ssFTDate1 As String = TIMS.GetMyValue(strSession, "SFTDate")
        Dim ssFTDate2 As String = TIMS.GetMyValue(strSession, "FFTDate")
        Dim ssIdentity As String = TIMS.GetMyValue(strSession, "Identity")
        Dim ssDistID As String = TIMS.GetMyValue(strSession, "DistID")
        Dim ssTPlanID As String = TIMS.GetMyValue(strSession, "TPlanID")
        Dim ssBudgetID As String = TIMS.GetMyValue(strSession, "BudgetID")
        Dim ssYears As String = TIMS.GetMyValue(strSession, "Years")

        'Dim sql As String = ""
        Dim dt As DataTable = Get_MainDt1(sFileName, strSession, rblSchType1.SelectedValue, objconn)
        '鎖定計畫範圍。
        Dim dt_KeyPlan As DataTable = Get_KeyPlanDt(ssTPlanID, objconn)

        'Dim cst_sTitle As String = "主要對象班級明細資料"
        Dim strTitle2 As String = "" '依抬頭
        Dim XID_NAME As String = "" '班級學員使用欄位名稱
        Dim dtX As DataTable = Get_MainDt2(Me, rblMode1.SelectedIndex, strTitle2, XID_NAME, objconn)

        Dim sFileName1 As String = "主要對象班級明細資料"

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= (".noDecFormat2{mso-number-format:""0.00%"";}")
        strSTYLE &= ("</style>")

        '年度、轄區、訓練計畫、訓練單位、班級名稱、開訓日期、結訓日期、職類(大項)、學員姓名、身分證號、性別、出生年月日、主要參訓身分別、是否為在職者、是否離訓、是否退訓、是否結訓、是否提前就業、是否為不就業人數、是否為未就業、是否為就業、是否為公法救助
        'Const cst_iColspanNUM As Integer = 10 '每組固定數量。
        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim strExp As String = ""
        strExp = "<tr>"
        strExp &= "<td>年度</td>"
        strExp &= "<td>轄區</td>"
        strExp &= "<td>訓練計畫</td>"
        strExp &= "<td>訓練單位</td>"
        strExp &= "<td>班級名稱</td>"
        strExp &= "<td>開訓日期</td>"
        strExp &= "<td>結訓日期</td>"
        strExp &= "<td>職類(大項)</td>"
        strExp &= "<td>學員姓名</td>"
        strExp &= "<td>身分證號</td>"
        strExp &= "<td>性別</td>"
        strExp &= "<td>出生年月日</td>"
        strExp &= "<td>主要參訓身分別</td>"
        'strExp &= "<td>是否為在職者</td>"
        strExp &= "<td>是否離訓</td>"
        strExp &= "<td>是否退訓</td>"
        strExp &= "<td>是否結訓</td>"
        'strExp &= "<td>是否提前就業</td>"
        'strExp &= "<td>是否為不就業人數</td>"
        'strExp &= "<td>是否為未就業</td>"
        'strExp &= "<td>是否為就業</td>"
        'strExp &= "<td>是否為公法救助</td>"
        strExp &= "</tr>"
        strHTML &= (strExp)

        Dim TMP1 As String = ""
        For Each dr1 As DataRow In dt.Rows
            strExp = "<tr>"
            strExp &= "<td>" & Convert.ToString(dr1("Years")) & "</td>" '年度</td>"
            strExp &= "<td>" & Convert.ToString(dr1("DistName")) & "</td>" '轄區</td>"
            strExp &= "<td>" & Convert.ToString(dr1("PlanName")) & "</td>" '訓練計畫</td>"
            strExp &= "<td>" & Convert.ToString(dr1("OrgName")) & "</td>" '訓練單位</td>"
            strExp &= "<td>" & Convert.ToString(dr1("ClassCName")) & "</td>" '班級名稱</td>"
            strExp &= "<td>" & Convert.ToString(dr1("STDATE")) & "</td>" '開訓日期</td>"
            strExp &= "<td>" & Convert.ToString(dr1("FTDATE")) & "</td>" '結訓日期</td>"
            strExp &= "<td>" & Convert.ToString(dr1("BUSNAME")) & "</td>" '職類(大項)</td>"
            strExp &= "<td>" & Convert.ToString(dr1("stdname")) & "</td>" '學員姓名</td>"
            strExp &= "<td>" & Convert.ToString(dr1("IDNO")) & "</td>" '身分證號</td>"
            TMP1 = "女"
            If Convert.ToString(dr1("SEX")) = "M" Then TMP1 = "男"
            strExp &= "<td>" & TMP1 & "</td>" '性別</td>"
            strExp &= "<td>" & Convert.ToString(dr1("BIRTHDAY")) & "</td>" '出生年月日</td>"
            strExp &= "<td>" & Convert.ToString(dr1("MIdentityN")) & "</td>" '主要參訓身分別</td>"
            'TMP1 = "否"'If Convert.ToString(dr1("WIdent")) = "1" Then TMP1 = "是"
            'strExp &= "<td>" & TMP1 & "</td>" '是否為在職者</td>"

            TMP1 = If(Convert.ToString(dr1("STUDSTATUS")) = "2", "是", "否")
            strExp &= "<td>" & TMP1 & "</td>" '是否離訓</td>"
            TMP1 = If(Convert.ToString(dr1("STUDSTATUS")) = "3", "是", "否")
            strExp &= "<td>" & TMP1 & "</td>" '是否退訓</td>"
            TMP1 = If(Convert.ToString(dr1("finx")) = "1", "是", "否")
            strExp &= "<td>" & TMP1 & "</td>" '是否結訓</td>"

            'TMP1 = "否"'If Convert.ToString(dr1("wjobx")) = "1" Then TMP1 = "是"
            'strExp &= "<td>" & TMP1 & "</td>" '是否提前就業</td>"
            'TMP1 = "否"'If Convert.ToString(dr1("nojobx")) = "1" Then TMP1 = "是"
            'strExp &= "<td>" & TMP1 & "</td>" '是否為不就業人數</td>"
            'TMP1 = "否"'If Convert.ToString(dr1("nojobx")) <> "1" AndAlso Convert.ToString(dr1("injobx")) <> "1" Then TMP1 = "是"
            'strExp &= "<td>" & TMP1 & "</td>" '是否為未就業</td>"
            'TMP1 = "否"'If Convert.ToString(dr1("injobx")) = "1" Then TMP1 = "是"
            'strExp &= "<td>" & TMP1 & "</td>" '是否為就業</td>"
            'TMP1 = "否"'If Convert.ToString(dr1("PUCount")) = "1" Then TMP1 = "是"
            'strExp &= "<td>" & TMP1 & "</td>" '是否為公法救助</td>"

            strExp &= "</tr>"
            strHTML &= (strExp)
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
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()  
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出 統計資料-匯出鈕(EXCEL)
    Private Sub BtnExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExp.Click
        Call ExpX(1)
    End Sub

    '匯出班級明細資料-匯出鈕(EXCEL)
    Protected Sub BtnExp2_Click(sender As Object, e As EventArgs) Handles BtnExp2.Click
        Call ExpX(2)
    End Sub

End Class