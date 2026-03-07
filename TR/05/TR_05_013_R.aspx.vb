Partial Class TR_05_013_R
    Inherits AuthBasePage

    'dt 其他欄位計算
    'Function change_table(ByVal dt As DataTable) As DataTable
    '    Const cst_結訓人數 As Integer = 4
    '    'Const cst_結訓人數男 As Integer = 5
    '    'Const cst_結訓人數女 As Integer = 6
    '    'Const cst_就業人數 As Integer = 11
    '    'Const cst_就業率 As Integer = 12
    '    'For i As Integer = 0 To dt.Rows.Count - 1
    '    '    Dim dr As DataRow = dt.Rows(i)
    '    '    dr(cst_就業率) = "0.00%"
    '    '    If CInt(dr(cst_結訓人數)) <> 0 Then dr(cst_就業率) = TIMS.ROUND(CDbl(dr(cst_就業人數)) / CDbl(dr(cst_結訓人數)) * 100, 2) & "%"
    '    'Next
    '    'dt.AcceptChanges()
    '    Return dt
    'End Function

    Const cst_vsTR05013RPAR As String = "TR05013RPAR"
    Const cst_vsTR05013RCNT As String = "TR05013RCNT"
    Const cst_vsTR05013RSQL As String = "TR05013RSQL"

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        If Not IsPostBack Then
            msg.Text = ""
            PageControler1.Visible = False
            DataGrid1.Visible = False
            Call CreateItem() '關鍵字詞建立
        End If
    End Sub

    '關鍵字詞建立
    Sub CreateItem()
        Syear = TIMS.GetSyear(Syear) '年度
        TIMS.Tooltip(Syear, "請選擇年度或開結訓日期(至少要有2項填寫)", True)

        ''轄區別
        'DistID = TIMS.Get_DistID(DistID)
        'DistID.Items.Insert(0, New ListItem("全部", ""))

        '轄區別
        'DistID = TIMS.Get_DistID(DistID)
        'DistID.Items.Remove(DistID.Items.FindByValue(""))
        'DistID.Items.Insert(0, New ListItem("全部", ""))
        '轄區別
        CBLDISTID = TIMS.Get_DistID(CBLDISTID, TIMS.dtNothing, objconn)
        CBLDISTID.Items.Remove(CBLDISTID.Items.FindByValue(""))
        CBLDISTID.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y") '計畫別
        BudgetList = TIMS.Get_Budget(BudgetList, 4) '預算來源  '4:含 ECFA(協助)

        'DistID.Enabled = True
        'Common.SetListItem(DistID, sm.UserInfo.DistID)
        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(CBLDISTID, sm.UserInfo.DistID) '轄區
        End If

        Common.SetListItem(Syear, sm.UserInfo.Years)
        Syear.Enabled = True

        '97(協助基金)
        Common.SetListItem(BudgetList, "97")
        'BudgetList.SelectedValue = "97"
        'BudgetList.Enabled = False '鎖定 ('無法鎖定 無法取值)

        If sm.UserInfo.DistID <> "000" Then
            '若不是署則鎖定下列功能
            Syear.Enabled = False
            CBLDISTID.Enabled = False
            'DistID.Attributes.Add("disabled", "disabled")
        End If
        Select Case CStr(sm.UserInfo.LID) '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "1"
                Syear.Enabled = True '開放分署可選年度
        End Select

        ''選擇全部轄區
        'DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部轄區
        CBLDISTID.Attributes("onclick") = "SelectAll('CBLDISTID','HidCBLDISTID');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        'btnPrint.Attributes("onclick") = "return chkSearch();"
        btnSearch.Attributes("onclick") = "return chkSearch();"
        btnExport.Attributes("onclick") = "return chkSearch();"
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        'If Me.DistID.SelectedValue = "" Then Errmsg += "請選擇轄區中心" & vbCrLf
        ''If Me.hidDistID.Value = "" Then Errmsg += "請選擇轄區中心" & vbCrLf
        Dim v_Syear As String = TIMS.GetListValue(Syear)
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)

        Dim iMOK As Integer = 0
        If v_Syear <> "" Then iMOK += 2
        If STDate1.Text <> "" Then iMOK += 1
        If STDate2.Text <> "" Then iMOK += 1
        If FTDate1.Text <> "" Then iMOK += 1
        If FTDate2.Text <> "" Then iMOK += 1

        Dim fg_NG_ALL_NO_SEL1 As Boolean = (v_Syear = "" AndAlso STDate1.Text <> "" AndAlso STDate2.Text <> "" AndAlso FTDate1.Text <> "" AndAlso FTDate2.Text <> "")
        If fg_NG_ALL_NO_SEL1 OrElse iMOK < 2 Then Errmsg &= "請選擇年度或開結訓日期(至少要有2項填寫)" & vbCrLf

        Dim mySTDate1 As String = If(flag_ROC, TIMS.Cdate18(STDate1.Text), STDate1.Text)  'edit，by:20181022
        Dim mySTDate2 As String = If(flag_ROC, TIMS.Cdate18(STDate2.Text), STDate2.Text)  'edit，by:20181022
        Dim myFTDate1 As String = If(flag_ROC, TIMS.Cdate18(FTDate1.Text), FTDate1.Text)  'edit，by:20181022
        Dim myFTDate2 As String = If(flag_ROC, TIMS.Cdate18(FTDate2.Text), FTDate2.Text)  'edit，by:20181022

        If mySTDate1 <> "" Then
            If Not TIMS.IsDate1(mySTDate1) Then Errmsg += "開訓期間 的起始日不是正確的日期格式" & vbCrLf
            'If Errmsg = "" Then STDate1.Text = CDate(mySTDate1).ToString("yyyy/MM/dd")
        End If

        If mySTDate2 <> "" Then
            If Not TIMS.IsDate1(mySTDate2) Then Errmsg += "開訓期間 的迄止日不是正確的日期格式" & vbCrLf
            'If Errmsg = "" Then STDate2.Text = CDate(mySTDate2).ToString("yyyy/MM/dd")
        End If

        If Errmsg = "" AndAlso mySTDate1 <> "" AndAlso mySTDate2 <> "" Then
            If DateDiff(DateInterval.Day, CDate(mySTDate1), CDate(mySTDate2)) < 0 Then Errmsg += "開訓期間 日期起迄有誤，迄日需大於起日" & vbCrLf
        End If
        If myFTDate1 <> "" AndAlso Not TIMS.IsDate1(myFTDate1) Then
            Errmsg += "結訓期間 的起始日不是正確的日期格式" & vbCrLf
            'If Errmsg = "" Then FTDate1.Text = CDate(myFTDate1).ToString("yyyy/MM/dd")
        End If
        If myFTDate2 <> "" AndAlso Not TIMS.IsDate1(myFTDate2) Then
            Errmsg += "結訓期間 的迄止日不是正確的日期格式" & vbCrLf
            'If Errmsg = "" Then FTDate2.Text = CDate(myFTDate2).ToString("yyyy/MM/dd")
        End If

        If Errmsg = "" AndAlso myFTDate1 <> "" AndAlso myFTDate2 <> "" Then
            If DateDiff(DateInterval.Day, CDate(myFTDate1), CDate(myFTDate2)) < 0 Then Errmsg += "結訓期間 日期起迄有誤，迄日需大於起日" & vbCrLf
        End If

        Dim j As Integer = 0
        Dim CBLobj As CheckBoxList

        j = 0
        CBLobj = TPlanID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem IsNot Nothing AndAlso objitem.Selected Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇訓練計畫" & vbCrLf

        j = 0
        CBLobj = BudgetList
        For i As Integer = 0 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem IsNot Nothing AndAlso objitem.Selected Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇預算來源" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Function Sch_WITH_ALL(ByRef parms As Hashtable) As String
        Dim itemDist As String = TIMS.GetCblValueIn(CBLDISTID)
        Dim itemPlan As String = TIMS.GetCblValueIn(TPlanID)
        Dim v_Syear As String = TIMS.GetListValue(Syear)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT vp.DistID,vp.DISTNAME,vp.YEARS,vp.PLANNAME" & vbCrLf
        sql &= " ,cc.OCID,cc.STDATE,cc.FTDATE,cc.CLASSCNAME,cc.CYCLTYPE" & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        sql &= " FROM VIEW_PLAN vp" & vbCrLf
        sql &= " JOIN KEY_PLAN kp ON kp.TPlanID = vp.TPlanID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc ON vp.planid = cc.planid" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.planid = cc.planid AND pp.comidno = cc.comidno AND pp.seqno = cc.seqno" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo ON oo.comidno = cc.comidno" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y'" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        '依登入年度
        'sql &= " AND vp.years = '" & sm.UserInfo.Years & "'" & vbCrLf
        If v_Syear <> "" Then
            sql &= " AND vp.years = @years" & vbCrLf
            parms.Add("years", v_Syear)
        End If
        If itemDist <> "" Then sql &= String.Concat(" AND vp.DISTID IN (", itemDist, ")", vbCrLf)
        If Me.STDate1.Text <> "" Then
            sql &= " AND cc.STdate >= @STdate1" & vbCrLf
            parms.Add("STdate1", If(flag_ROC, TIMS.Cdate18(Me.STDate1.Text), Me.STDate1.Text))  'edit，by:20181022
        End If
        If Me.STDate2.Text <> "" Then
            sql &= " AND cc.STdate <= @STdate2" & vbCrLf
            parms.Add("STdate2", If(flag_ROC, TIMS.Cdate18(Me.STDate2.Text), Me.STDate2.Text))  'edit，by:20181022
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= " AND cc.ftdate >= @ftdate1" & vbCrLf
            parms.Add("ftdate1", If(flag_ROC, TIMS.Cdate18(Me.FTDate1.Text), Me.FTDate1.Text))  'edit，by:20181022
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= " AND cc.ftdate <= @ftdate2" & vbCrLf
            parms.Add("ftdate2", If(flag_ROC, TIMS.Cdate18(Me.FTDate2.Text), Me.FTDate2.Text))  'edit，by:20181022
        End If
        If itemPlan <> "" Then sql &= String.Concat(" AND vp.TPlanID IN (", itemPlan, ")", vbCrLf)
        'sql &= " AND vp.years = '2021'" & vbCrLf
        'sql &= " AND vp.DISTID IN ('001')" & vbCrLf
        'sql &= " AND vp.TPlanID IN ('28')" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " SELECT cc.OCID ,cs.SOCID" & vbCrLf
        sql &= " ,ss.NAME ,ss.IDNO ,ss.BIRTHDAY" & vbCrLf
        sql &= " ,cs.ACTNO" & vbCrLf
        sql &= " ,ISNULL(SP.ACTNAME,bb.UNAME) ACTNAME" & vbCrLf
        'sql &= " /* '增加身分別、年齡、教育程度、性別 */" & vbCrLf
        sql &= " ,cs.MIDENTITYID" & vbCrLf
        sql &= " ,ki.NAME MIDNAME" & vbCrLf
        sql &= " ,DATEDIFF(YEAR, SS.BIRTHDAY,CC.STDATE) YEARSOLD1" & vbCrLf
        sql &= " ,SS.DEGREEID" & vbCrLf
        sql &= " ,KD.NAME DEGREENAME" & vbCrLf
        sql &= " ,SS.SEX" & vbCrLf
        sql &= " ,CASE SS.SEX WHEN 'M' THEN '男' WHEN 'F' THEN '女' ELSE SS.SEX END SEX2" & vbCrLf
        sql &= " ,cs.CREDITPOINTS" & vbCrLf
        sql &= " ,cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cs.ISAPPRPAPER" & vbCrLf
        sql &= " ,dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE) I_STUDCNT14B" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.ocid = cc.ocid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.SID = cs.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA ss2 ON ss2.sid = ss.sid" & vbCrLf
        sql &= " LEFT JOIN KEY_IDENTITY ki ON ki.IdentityID= cs.MIdentityID" & vbCrLf
        sql &= " LEFT JOIN KEY_DEGREE KD ON KD.DEGREEID = SS.DEGREEID" & vbCrLf
        sql &= " LEFT JOIN STUD_SERVICEPLACE SP ON SP.SOCID = cs.SOCID" & vbCrLf
        sql &= " LEFT JOIN BUS_BASICDATA bb ON bb.Ubno = cs.ActNo" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        '預算來源-限定:97 ECFA 協助
        sql &= " AND cs.BudgetID = '97'" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WS2 AS (" & vbCrLf
        sql &= " SELECT cc.ocid" & vbCrLf
        '/*開訓人次-協助*/
        sql &= " ,COUNT(case when cs.IsApprPaper='Y'" & vbCrLf
        sql &= " AND cs.I_STUDCNT14B=1" & vbCrLf
        sql &= " then 1 end) openstudcount97" & vbCrLf
        '/*結訓-協助人次*/
        sql &= " ,COUNT(case when cs.IsApprPaper='Y'" & vbCrLf
        sql &= " and cs.CreditPoints is not NULL" & vbCrLf
        sql &= " and cs.StudStatus Not IN (2,3)" & vbCrLf
        sql &= " and cc.FTDate < GETDATE()" & vbCrLf
        sql &= " then 1 end) closestudcout97" & vbCrLf
        '/*結訓-協助人次(男)*/
        sql &= " ,COUNT(case when cs.IsApprPaper='Y'" & vbCrLf
        sql &= " and cs.CreditPoints is not NULL" & vbCrLf
        sql &= " and cs.StudStatus Not IN (2,3)" & vbCrLf
        sql &= " and cc.FTDate < GETDATE() AND cs.sex='M'" & vbCrLf
        sql &= " then 1 end) closestudcout97M" & vbCrLf
        '/*結訓-協助人次(女)*/
        sql &= " ,COUNT(case when cs.IsApprPaper='Y'" & vbCrLf
        sql &= " and cs.CreditPoints is not NULL" & vbCrLf
        sql &= " and cs.StudStatus Not IN (2,3)" & vbCrLf
        sql &= " and cc.FTDate < GETDATE() AND cs.sex='F'" & vbCrLf
        sql &= " then 1 end) closestudcout97F" & vbCrLf

        sql &= " ,COUNT(CASE WHEN cs.socid IS NOT NULL AND cs.studstatus IN (2,3) THEN 1 END) cnt2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cs.socid IS NOT NULL AND cs.studstatus NOT IN (2,3) THEN 1 END) cnt3" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN WS1 cs on cs.ocid=cc.ocid" & vbCrLf
        sql &= " GROUP BY cc.ocid" & vbCrLf
        sql &= " )" & vbCrLf
        '報名1
        sql &= " ,WE1 AS (" & vbCrLf
        sql &= " SELECT b.ocid1 ,a.idno ,b.EnterDate" & vbCrLf
        sql &= " FROM stud_enterTemp a" & vbCrLf
        sql &= " JOIN stud_enterType b ON a.SETID = b.SETID" & vbCrLf
        sql &= " JOIN WC1 cc ON cc.ocid = b.ocid1" & vbCrLf
        sql &= " )" & vbCrLf
        '報名2
        sql &= " ,WE2 AS (" & vbCrLf
        sql &= " SELECT  b.ocid1 ,a.idno ,b.EnterDate" & vbCrLf
        sql &= " FROM stud_enterTemp2 a" & vbCrLf
        sql &= " JOIN stud_enterType2 b ON a.eSETID = b.eSETID" & vbCrLf
        sql &= " JOIN WC1 cc ON cc.ocid = b.ocid1" & vbCrLf
        sql &= " )" & vbCrLf
        '統計ocid1 /cnt1
        sql &= " ,WE3 AS (" & vbCrLf
        sql &= " SELECT g.ocid1 ,COUNT(1) cnt1" & vbCrLf
        sql &= " FROM (SELECT ocid1,idno FROM WE1 UNION SELECT ocid1,idno FROM WE2) g" & vbCrLf
        sql &= " GROUP BY g.ocid1" & vbCrLf
        sql &= " )" & vbCrLf
        '統計ocid1,IDNO /ENTERDATE
        sql &= " ,WE4 AS (" & vbCrLf
        sql &= " SELECT g.ocid1,g.IDNO, MAX(g.EnterDate) EnterDate" & vbCrLf
        sql &= " FROM (SELECT ocid1,IDNO,ENTERDATE FROM WE1 UNION SELECT ocid1,IDNO,ENTERDATE FROM WE2) g" & vbCrLf
        sql &= " GROUP BY g.ocid1,g.IDNO" & vbCrLf
        sql &= " )" & vbCrLf

        Return sql
    End Function

    '統計 查詢 [SQL] ECFA人數統計資料
    Sub Search1()

#Region "(No Use)"

        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " select" & vbCrLf
        'sql &= " vp.DistID" & vbCrLf '0
        'sql &= " ,vp.distname 轄區" & vbCrLf '1
        'sql &= " ,isnull(SUM(isnull(g2.cnt1,0)),0) 報名人數" & vbCrLf '2
        'sql &= " ,isnull(SUM(isnull(g1.cnt1,0)),0) 參訓人數" & vbCrLf '3
        'sql &= " ,isnull(SUM(isnull(g1.cnt4,0)),0) 結訓人數" & vbCrLf 'cst_結訓人數:4
        'sql &= " ,isnull(SUM(isnull(g1.cnt4a,0)),0) ""結訓人數(男)""" & vbCrLf 'cst_結訓人數:4
        'sql &= " ,isnull(SUM(isnull(g1.cnt4b,0)),0) ""結訓人數(女)""" & vbCrLf 'cst_結訓人數:4

        'sql &= " ,isnull(SUM(isnull(g1.cnt5,0)),0) 活津貼申請人數" & vbCrLf '6
        'sql &= " ,isnull(SUM(isnull(g1.cnt6,0)),0) 生活津貼通過人數" & vbCrLf '7
        'sql &= " ,isnull(SUM(isnull(g1.cnt7,0)),0) ""系統判定(就業)人數""" & vbCrLf '8
        'sql &= " ,isnull(SUM(isnull(g1.cnt8,0)),0) ""人工判定(就業)人數""" & vbCrLf '9
        'sql &= " ,isnull(SUM(isnull(g1.cnt9,0)),0) 就業人數" & vbCrLf 'cst_就業人數:9
        'sql &= " ,'' 就業率" & vbCrLf 'cst_就業率:10

#End Region

        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql = Sch_WITH_ALL(parms)
        sql &= " SELECT cc.DistID 轄區代碼" & vbCrLf
        sql &= " ,cc.DISTNAME 轄區" & vbCrLf
        sql &= " ,ISNULL(SUM(ISNULL(g3.cnt1,0)),0) 報名人數" & vbCrLf

        sql &= " ,ISNULL(SUM(ISNULL(ss.openstudcount97,0)),0) 參訓人數" & vbCrLf
        sql &= " ,ISNULL(SUM(ISNULL(ss.closestudcout97,0)),0) 結訓人數" & vbCrLf
        sql &= " ,ISNULL(SUM(ISNULL(ss.closestudcout97M,0)),0) ""結訓人數(男)""" & vbCrLf
        sql &= " ,ISNULL(SUM(ISNULL(ss.closestudcout97F,0)),0) ""結訓人數(女)""" & vbCrLf

        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt1,0)),0) 參訓人數" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt4,0)),0) 結訓人數" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt4a,0)),0) ""結訓人數(男)""" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt4b,0)),0) ""結訓人數(女)""" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt5,0)),0) 生活津貼申請人數" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt6,0)),0) 生活津貼通過人數" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt7,0)),0) ""系統判定(就業)人數""" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt8,0)),0) ""人工判定(就業)人數""" & vbCrLf
        'sql &= " ,ISNULL(SUM(ISNULL(ss.cnt9,0)),0) 就業人數" & vbCrLf
        'sql &= " ,'' 就業率" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WS2 ss ON ss.ocid = cc.ocid" & vbCrLf
        sql &= " LEFT JOIN WE3 g3 ON g3.ocid1 = cc.ocid" & vbCrLf
        sql &= " GROUP BY cc.DistID ,cc.distname" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        'If dt.Rows.Count > 0 Then dt = change_table(dt) '計算(就業率)
        ViewState(cst_vsTR05013RPAR) = TIMS.GetMyValue3(parms)
        ViewState(cst_vsTR05013RCNT) = dt.Rows.Count
        ViewState(cst_vsTR05013RSQL) = sql

#Region "(No Use)"

        'Try
        'Catch ex As Exception
        '    'Common.RespWrite(Me, TIMS.sUtl_AntiXss(sql))
        '    'Common.RespWrite(Me, ex.ToString)
        '    Common.MessageBox(Me, ex.ToString)
        '    Common.MessageBox(Me, "若匯出效能太差，請選擇單1計畫或單1轄區!!")
        '    Exit Sub
        'End Try

#End Region

        'Table4.Style("display") = "inline"
        'Print.Visible = False
        'btnExport1.Visible = False
        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            'dt.DefaultView
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True
            'Print.Visible = True
            'btnExport1.Visible = True
            'PageControler1.Visible = True
            PageControler1.PageDataTable = dt
            'PageControler1.Sort = "訓練計畫,訓練機構,班級名稱,訓練起迄"
            PageControler1.Sort = "轄區代碼" '"DistID"
            'PageControler1.SqlString = Sql
            PageControler1.ControlerLoad()
        End If
        'Else
        'Common.MessageBox(Me, "查無資料")
    End Sub

    '明細資料 查詢 [SQL] ECFA學員明細資料
    Sub Search2()

        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql = Sch_WITH_ALL(parms)
        sql &= " SELECT cc.DISTNAME 轄區" & vbCrLf
        sql &= " ,cc.YEARS 年度" & vbCrLf
        sql &= " ,cc.PLANNAME 訓練計畫" & vbCrLf
        sql &= " ,cc.ORGNAME 訓練單位" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) 班級名稱" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) 開訓日" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) 結訓日" & vbCrLf
        sql &= " ,ss.name 姓名" & vbCrLf
        'sql &= " ,ss.IDNO 身分證字號" & vbCrLf
        ',dbo.FN_GET_MASK1(ss.IDNO) IDNO_MK
        sql &= " ,dbo.FN_GET_MASK1(ss.IDNO) 身分證字號" & vbCrLf
        'sql &= " ,format(ss.BIRTHDAY,'yyyy/MM/dd') 出生日期" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK2(ss.BIRTHDAY) 出生日期" & vbCrLf
        sql &= " ,ss.ACTNO 保險證號" & vbCrLf
        sql &= " ,ss.ACTNAME 最後投保單位" & vbCrLf
        sql &= " ,format(e4.EnterDate,'yyyy/MM/dd') 報名日期" & vbCrLf
        '增加身分別、年齡、教育程度、性別
        sql &= " ,ss.MIDNAME 身分別" & vbCrLf
        sql &= " ,ss.YEARSOLD1 年齡" & vbCrLf
        sql &= " ,ss.DEGREENAME 教育程度" & vbCrLf
        sql &= " ,ss.SEX2 性別" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN WS1 ss ON ss.ocid = cc.ocid" & vbCrLf
        sql &= " LEFT JOIN WE4 e4 ON e4.idno = ss.idno AND e4.ocid1 = ss.ocid" & vbCrLf
        sql &= " WHERE 1=1 AND ss.IsApprPaper='Y' AND ss.I_STUDCNT14B=1" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        ViewState(cst_vsTR05013RPAR) = TIMS.GetMyValue3(parms)
        ViewState(cst_vsTR05013RCNT) = dt.Rows.Count
        ViewState(cst_vsTR05013RSQL) = sql

        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True
            'Print.Visible = True
            'btnExport1.Visible = True
            'PageControler1.Visible = True
            PageControler1.PageDataTable = dt
            'PageControler1.Sort = "訓練計畫,訓練機構,班級名稱,訓練起迄"
            PageControler1.Sort = "轄區,年度,訓練計畫,訓練單位,班級名稱"
            'PageControler1.SqlString = Sql
            PageControler1.ControlerLoad()
        End If
        'Else
        'Common.MessageBox(Me, "查無資料")
    End Sub

    '班級別 ECFA班級統計資料
    Sub Search3()
        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql = Sch_WITH_ALL(parms)
        sql &= " SELECT cc.distname 轄區" & vbCrLf
        sql &= " ,cc.years 年度" & vbCrLf
        sql &= " ,cc.planname 訓練計畫" & vbCrLf
        sql &= " ,cc.orgname 訓練單位" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) 班級名稱" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) 開訓日" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) 結訓日" & vbCrLf
        sql &= " ,ss.openstudcount97 開訓人數" & vbCrLf
        sql &= " ,ss.closestudcout97 結訓人數" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN WS2 ss ON ss.ocid = cc.ocid" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        ViewState(cst_vsTR05013RPAR) = TIMS.GetMyValue3(parms)
        ViewState(cst_vsTR05013RCNT) = dt.Rows.Count
        ViewState(cst_vsTR05013RSQL) = sql

        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True
            'Print.Visible = True
            'btnExport1.Visible = True
            'PageControler1.Visible = True
            PageControler1.PageDataTable = dt
            'PageControler1.Sort = "訓練計畫, 訓練機構, 班級名稱, 訓練起迄"
            PageControler1.Sort = "轄區, 年度, 訓練計畫, 訓練單位, 班級名稱"
            'PageControler1.SqlString = Sql
            PageControler1.ControlerLoad()
        End If
        'Else
        'Common.MessageBox(Me, "查無資料")
    End Sub

    ''' <summary> '匯出  1:統計資料 2:明細資料 3:班級統計 </summary>
    ''' <param name="iType"></param>
    Sub sExport123(ByVal iType As Integer)
        '1:統計資料/2:明細資料/3:班級別
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim sFileName1 As String = ""
        Select Case iType
            Case 1
                Call Search1()
                sFileName1 = "ECFA人數統計資料"
            Case 2
                Call Search2()
                sFileName1 = "ECFA學員明細資料"
            Case 3
                Call Search3()
                sFileName1 = "ECFA班級統計資料"
        End Select

        If DataGrid1.Visible = False OrElse msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If

        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= String.Concat("&ACT=", sFileName1) & vbCrLf
        sMemo &= String.Concat("&PARMS=", ViewState(cst_vsTR05013RPAR)) & vbCrLf
        sMemo &= String.Concat("&COUNT=", ViewState(cst_vsTR05013RCNT)) & vbCrLf
        sMemo &= String.Concat("&SQL=", ViewState(cst_vsTR05013RSQL)) & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip2, "", sMemo, objconn)

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format: ""\@"";}")
        strSTYLE &= ("</style>")
        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        DataGrid1.AllowPaging = True '開啟分頁
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()

        'DataGrid1.Columns(8).Visible = True
    End Sub

#Region "(No Use)"

    '匯出
    'Private Sub btnExport2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim Errmsg As String = ""
    '    Call CheckData1(Errmsg)
    '    If Errmsg <> "" Then
    '        Common.MessageBox(Page, Errmsg)
    '        Exit Sub
    '    End If
    '    DataGrid1.AllowPaging = False
    '    'DataGrid1.Columns(8).Visible = False
    '    DataGrid1.EnableViewState = False  '把ViewState給關了
    '    Call Search2()
    '    If DataGrid1.Visible = False OrElse msg.Text <> "" Then
    '        Common.MessageBox(Page, msg.Text)
    '        Exit Sub
    '    End If
    '    Dim sFileName As String = ""
    '    sFileName = HttpUtility.UrlEncode("ECFA學員明細資料.xls", System.Text.Encoding.UTF8)
    '    Response.Clear()
    '    Response.Buffer = True
    '    Response.Charset = "UTF-8" '設定字集
    '    Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
    '    Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
    '    'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
    '    '文件內容指定為Excel
    '    Response.ContentType = "application/ms-excel;charset=utf-8"
    '    Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
    '    ''套CSS值
    '    Common.RespWrite(Me, "<style>")
    '    Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
    '    Common.RespWrite(Me, "</style>")
    '    DataGrid1.AllowPaging = False
    '    'DataGrid1.Columns(8).Visible = False
    '    DataGrid1.EnableViewState = False  '把ViewState給關了
    '    Dim objStringWriter As New System.IO.StringWriter
    '    Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
    '    Div1.RenderControl(objHtmlTextWriter)
    '    Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
    '    Response.End()
    '    DataGrid1.AllowPaging = True
    '    'DataGrid1.Columns(8).Visible = True
    'End Sub

#End Region

    '查詢 1.統計 2.明細資料
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim v_rblSearchStyle As String = TIMS.GetListValue(rblSearchStyle)
        Select Case v_rblSearchStyle 'Me.rblSearchStyle.SelectedValue
            Case "1" '統計
                Call Search1()
            Case "2" '明細
                Call Search2()
            Case "3" '班級別
                Call Search3()
            Case Else
                Common.MessageBox(Me, "請選擇查詢方式!!")
                Exit Sub
        End Select
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    '匯出 excel
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Dim v_rblSearchStyle As String = TIMS.GetListValue(rblSearchStyle)
        Select Case v_rblSearchStyle 'Me.rblSearchStyle.SelectedValue
            Case "1", "2", "3"
                '1:統計資料/2:明細資料/3:班級別
                Dim iType As Integer = Val(v_rblSearchStyle) 'Val(Me.rblSearchStyle.SelectedValue)
                Call sExport123(iType)
            Case Else
                Common.MessageBox(Me, "請選擇查詢方式!!")
                Exit Sub
        End Select
    End Sub
End Class