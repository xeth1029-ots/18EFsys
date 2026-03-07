Partial Class SD_13_003
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    'Const cst_是否取得結訓資格 = 3
    'Const cst_出席達3分之2 = 4
    'Const cst_出席達4分之3 = 4

    'Const cst_是否補助 = 5
    'Const cst_總費用 = 6
    'Const cst_補助費用 = 7
    'Const cst_個人支付 = 8
    'Const cst_剩餘可用餘額 = 9
    Const cst_其他申請中金額 As Integer = 10
    'Const cst_審核狀態 = 11
    'Const cst_審核備註 = 12
    'Const cst_保險證號 = 13
    'Const cst_預算別代碼 = 14
    'Dim objgConn As SqlConnection

    '年度大於2011啟用。

#Region "Functions"

    ''' <summary>
    ''' 學員補助申請資料-查詢資料 [SQL]
    ''' </summary>
    ''' <param name="iOCID"></param>
    ''' <param name="sort_Col"></param>
    ''' <returns></returns>
    Private Function GET_CLASSSTUDENTSOFCLASS(ByVal iOCID As Integer, ByVal sort_Col As String) As DataTable
        Dim rst As DataTable = Nothing
        sort_Col = TIMS.ClearSQM(sort_Col)
        Dim s_ORDERBYVAL As String = If(sort_Col <> "", sort_Col, "StudentID")
        Dim pms_s1 As New Hashtable From {{"OCID", iOCID}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID,a.SOCID ,b.IDNO,dbo.FN_CSTUDID2(a.StudentID) StudentID" & vbCrLf 'ORDER BY
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,d.TotalCost ,f.SumOfMoney ,f.PayMoney" & vbCrLf
        sql &= " ,f.AppliedStatusM ,f.AppliedNote" & vbCrLf
        sql &= " ,a.ActNO ,f.BudID" & vbCrLf
        sql &= " ,ISNULL(a.CreditPoints,0) CreditPoints" & vbCrLf
        sql &= " ,d.THours" & vbCrLf
        'sql &= " ,ISNULL(t.TOHours,0) TOHours" & vbCrLf 'sql &= " ,ISNULL(ff.COUNTHOURS,0)-ISNULL(ff.COUNTHOURS2,0) TOHours" & vbCrLf
        sql &= " ,ISNULL(ff.COUNTHOURS,0) TOHours" & vbCrLf
        sql &= " ,CONVERT(varchar, c.STDate, 111) STDate" & vbCrLf
        '政府已補助經費 'sql &= " ,0 GovCost" & vbCrLf
        '其他申請中金額 'sql &= " ,0 GovAppl2" & vbCrLf
        sql &= " ,dbo.FN_GET_GOVCOST2(b.IDNO, CONVERT(varchar, c.STDate, 111)) GovAppl2" & vbCrLf
        '檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
        'sql &= " ,'N' bliFlag" & vbCrLf
        '政府已補助經費 'sql &= " ,dbo.fn_GET_GOVCOST(b.IDNO, CONVERT(varchar, c.STDate, 111)) GovCost" & vbCrLf
        '其他申請中金額 'sql &= " ,dbo.FN_GET_GOVAPPL2(b.IDNO,c.STDate) GovAppl2" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO c" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO d ON d.ComIDNO=c.ComIDNO and d.PlanID=c.PlanID and d.SeqNO=c.SeqNO" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS a ON a.OCID=c.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b ON b.SID=a.SID" & vbCrLf
        sql &= " JOIN STUD_SUBSIDYCOST f ON f.SOCID=a.SOCID" & vbCrLf

        'select * from STUD_TURNOUT2 WHERE rownum <=10 '喪假(LEAVEID:05)。99:(使用者輸入)
        sql &= " LEFT JOIN ( SELECT t.SOCID" & vbCrLf
        sql &= " ,SUM(CASE WHEN t.LEAVEID IS NULL THEN t.Hours END) COUNTHOURS" & vbCrLf
        sql &= " FROM STUD_TURNOUT2 t" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.socid =t.socid" & vbCrLf
        sql &= " WHERE cs.OCID=@OCID" & vbCrLf
        sql &= " GROUP BY t.SOCID ) ff ON ff.SOCID=a.SOCID" & vbCrLf
        sql &= " WHERE a.OCID=@OCID" & vbCrLf
        sql &= $" ORDER BY {s_ORDERBYVAL} "
        rst = DbAccess.GetDataTable(sql, objconn, pms_s1)
        Return rst
    End Function

    '顯示查詢後的資料LIST
    Private Sub Show_DataGrid(ByVal dg As DataGrid, ByVal dt As DataTable, Optional ByVal key As String = "", Optional ByVal page As Integer = 0)
        AuditNumPanel.Visible = False
        msg.Text = "查無資料"
        DataGridTable.Style("display") = "none"

        If dt IsNot Nothing Then
            dg.DataSource = dt
            dg.CurrentPageIndex = page
            If key <> "" Then dg.DataKeyField = key
            dg.DataBind()
            AuditNum()

            AuditNumPanel.Visible = True
            msg.Text = ""
            DataGridTable.Style("display") = "" '"inline"
        End If
    End Sub

    ''' <summary>
    ''' 保留查詢參數
    ''' </summary>
    Sub KeepSearch()
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        TMID1.Text = TIMS.ClearSQM(TMID1.Text)
        OCID1.Text = TIMS.ClearSQM(OCID1.Text)
        TMIDValue1.Value = TIMS.ClearSQM(TMIDValue1.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim s_Search As String = ""
        s_Search &= "prg=SD13003"
        s_Search &= "&center=" & center.Text
        s_Search &= "&RIDValue=" & RIDValue.Value
        s_Search &= "&TMID1=" & TMID1.Text
        s_Search &= "&OCID1=" & OCID1.Text
        s_Search &= "&TMIDValue1=" & TMIDValue1.Value
        s_Search &= "&OCIDValue1=" & OCIDValue1.Value
        s_Search &= "&Button1=TRUE"
        'Dim flag_display As Boolean = False
        'If DataGridTable.Style("display") = "inline" OrElse DataGridTable.Style("display") = "" Then flag_display = True
        's_Search += If(flag_display, "&Button1=TRUE", "&Button1=FALSE")
        Session("_Search") = s_Search
    End Sub

    ''' <summary>
    ''' '帶入查詢參數
    ''' </summary>
    Sub UseKeepSearch()
        '帶入查詢參數
        Dim MyVale As String = ""
        If Session("_Search") IsNot Nothing AndAlso Convert.ToString(Session("_Search")) <> "" Then
            MyVale = TIMS.GetMyValue(Session("_Search"), "prg")
            If MyVale = "SD13003" Then
                center.Text = TIMS.GetMyValue(Session("_Search"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("_Search"), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Session("_Search"), "TMID1")
                OCID1.Text = TIMS.GetMyValue(Session("_Search"), "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session("_Search"), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Session("_Search"), "OCIDValue1")
                If TIMS.GetMyValue(Session("_Search"), "Button1") = "TRUE" Then
                    Session("_Search") = Nothing
                    'Button1_Click(sender, e)
                    sSearch1()
                End If
            End If
            Session("_Search") = Nothing
        End If
    End Sub

    '整班經費審核通過及不補助
    '儲存、經費審核確認、審核選單開放使用
    Sub OpenButtons(ByVal OCIDValue As String)

        '查看學員資料
        Dim hpms As New Hashtable From {{"OCID", Val(OCIDValue)}}
        Dim sql As String = ""
        sql &= " SELECT cs.SOCID,f.APPLIEDSTATUSM" & vbCrLf
        sql &= " FROM STUD_SUBSIDYCOST f" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.SOCID=f.SOCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.SID=cs.SID" & vbCrLf
        sql &= " WHERE cs.OCID=@OCID AND f.AppliedStatusM is null" & vbCrLf '經費審核，尚未確認
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, hpms)

        '整班經費審核通過及不補助
        '儲存、經費審核確認、審核選單開放使用
        'Dim hpms2 As New Hashtable From {{"OCID", Val(OCIDValue)}}
        'Dim sql2 As String = "SELECT APPLIEDRESULTM FROM CLASS_CLASSINFO WHERE OCID=@OCID"
        'Dim dr2 As DataRow = DbAccess.GetOneRow(sql2, objconn, hpms2)

        '還原經費審核確認" ToolTip="整班經費審核通過及不補助"
        AuditCheckR.Visible = False

        'by AMU 20220330 停用-還原經費審核確認
        'If IsDBNull(dr("AppliedResultM")) OrElse (dr("AppliedResultM").ToString = "R") Then
        '    AuditCheck.Enabled = True '經費審核確認
        '    TIMS.Tooltip(AuditCheck, "整班經費審核，尚未確認", True)
        'Else
        '    If dt.Rows.Count = 0 Then
        '        Dim flag_CanShow_AuditCheckR As Boolean = False
        '        If sm.UserInfo.RoleID = 0 Then flag_CanShow_AuditCheckR = True '還原經費審核確認 0:超級使用者
        '        If sm.UserInfo.LID = 0 Then flag_CanShow_AuditCheckR = True '還原經費審核確認
        '        If sm.UserInfo.LID = 1 Then flag_CanShow_AuditCheckR = True '還原經費審核確認

        '        '只有系統管理者開啟此功能。 '還原經費審核確認" ToolTip="整班經費審核通過及不補助"
        '        If flag_CanShow_AuditCheckR Then
        '            AuditCheckR.Visible = True
        '            TIMS.Tooltip(AuditCheckR, "提供還原整班經費審核通過(補助及不補助)", True)
        '        End If

        '        AuditCheck.Enabled = False '經費審核確認
        '        TIMS.Tooltip(AuditCheck, "整班經費審核通過(補助及不補助)", True)
        '    Else
        '        AuditCheck.Enabled = True '經費審核確認
        '        TIMS.Tooltip(AuditCheck, "整班經費審核通過,但尚有學員未審核通過，尚未確認", True)
        '    End If
        'End If
    End Sub

    Sub AuditNum()
        'Dim sql As String
        'Dim dr As DataRow
        'Dim dt As DataTable
        'sql = "SELECT Sum(Case When AppliedStatusM = 'Y' Then 1 Else 0 End) as SNum, " '審核成功筆數
        'sql &= "Sum(Case When AppliedStatusM = 'N' Then 1 Else 0 End) as FNum, " '審核失敗筆數
        'sql &= "Sum(Case When AppliedStatusM = 'R' Then 1 Else 0 End) as RNum, " '退件修正筆數
        'sql &= "Sum(Case When AppliedStatusM is Null Then 1 Else 0 End) as ANum " '未審核(請選擇、還原)筆數
        'sql &= "FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') AND Budid is not null "
        SNum.Text = "" 'dr("SNum").ToString
        FNum.Text = "" 'dr("FNum").ToString
        ANum.Text = "" 'dr("ANum").ToString
        RNum.Text = "" 'dr("RNum").ToString

        Dim hpms As New Hashtable From {{"OCID", Val(OCIDValue1.Value)}}
        Dim sql As String = ""
        sql &= " SELECT ISNULL(COUNT(Case When c.AppliedStatusM = 'Y' Then 1 End),0) SNum" & vbCrLf
        sql &= " ,ISNULL(COUNT(Case When c.AppliedStatusM = 'N' Then 1 End),0) FNum" & vbCrLf
        sql &= " ,ISNULL(COUNT(Case When c.AppliedStatusM = 'R' Then 1 End),0) RNum" & vbCrLf
        sql &= " ,ISNULL(COUNT(Case When c.AppliedStatusM is Null Then 1 End),0) ANum" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_SUBSIDYCOST c WITH(NOLOCK) on c.SOCID=cs.SOCID" & vbCrLf
        sql &= " WHERE c.Budid IS NOT NULL and cs.OCID=@OCID" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, hpms)
        If dr Is Nothing Then Exit Sub
        SNum.Text = $"{dr("SNum")}"
        FNum.Text = $"{dr("FNum")}"
        ANum.Text = $"{dr("ANum")}"
        RNum.Text = $"{dr("RNum")}"
    End Sub

#End Region

#Region "NO USE"
    'Function chkStud_SubsidyCost() As String  '檢查是否  補助費用>剩餘可用餘額  
    '    Dim errMsg As String = ""
    '    'Dim i As Int16 = 0
    '    For Each itm As DataGridItem In Datagrid2.Items
    '        Dim lab_SubSidyCost As Label = itm.FindControl("lab_SubSidyCost")   '補助費用
    '        Dim lab_Balance As Label = itm.FindControl("lab_Balance")            '剩餘可用餘額
    '        Dim Supply As Integer = 0     '補助費用
    '        Dim totLeft As Integer = 0    '剩餘可用餘額
    '        'i = i + 1
    '        If Trim(lab_SubSidyCost.Text) <> "" Then
    '            Supply = CInt(lab_SubSidyCost.Text)
    '        End If
    '        If Trim(lab_Balance.Text) <> "" Then
    '            totLeft = CInt(lab_Balance.Text)
    '        End If
    '        'errMsg = errMsg + "(" & i.ToString & ") " & Supply.ToString & "/" & totLeft.ToString & "\n"
    '        If Supply > totLeft Then
    '            errMsg = "補助費用 不可大於 剩餘可用餘額！"
    '            lab_SubSidyCost.Style("background-color") = "#FFCCFF"
    '        Else
    '            lab_SubSidyCost.Style("background-color") = "#FFFFFF"
    '        End If
    '    Next
    '    'Common.MessageBox(Me.Page, "<script>alert('" & errMsg & "'); </script>")
    '    Return errMsg
    'End Function
#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        Dim s_FUNID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        TIMS.Get_TitleLab(s_FUNID, TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'SD13003_BTN5 產投計畫-補助金審核-核銷文件審查-按鈕
        If Hid_SD13003_BTN5.Value = "" Then Hid_SD13003_BTN5.Value = TIMS.Utl_GetConfigVAL0(objconn, "SD13003_BTN5", 0)
        Hid_SD13003_BTN5.Value = TIMS.ClearSQM(Hid_SD13003_BTN5.Value)
        'Dim s_SD13003_BTN5 As String = Hid_SD13003_BTN5.Value
        If Hid_SD13003_BTN5.Value = "N" Then Button5.Visible = False

        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            Call CCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

        Button1.Attributes("onclick") = "return CheckSearch();"
        'AuditCheck.Attributes("onclick") = "return CheckDate(); return false;" '確認重複參訓是否繼續
        AuditCheck.Attributes("onclick") = "if (confirm('經費審核確認') ==true){ return chkMoney(); } else {return false;}  "
        AuditNumPanel.Visible = False

        '暫不提供 還原經費審核確認
        AuditCheckR.Visible = False
        AuditCheckR.Attributes("onclick") = "return confirm('這樣會還原經費審核確認，是否確定？');"

        '整班經費審核通過及不補助
        If OCIDValue1.Value <> "" Then Call OpenButtons(OCIDValue1.Value)

    End Sub

    Sub CCreate1()
        msg.Text = ""
        DataGridTable.Style("display") = "none"
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        'Button3.Attributes("onclick") = "return confirm('這樣會改變審核狀態，是否確定？');"

        '帶入查詢參數
        Call UseKeepSearch()

        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If center.Text <> "" Then RstMemo &= String.Concat("&center=", center.Text)
        If RIDValue.Value <> "" Then RstMemo &= String.Concat("&RID=", RIDValue.Value)
        Return RstMemo
    End Function

    Sub sSearch1()
        'Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        '職前參訓歷史查詢-WEB-SERVICE-依OCID
        TIMS.GetTrainingList2OCID(sm, objconn, OCIDValue1.Value)

        'Dim pms_s1 As New Hashtable From {{"OCID", Val(OCIDValue1.Value)}}
        'Dim sqlStr As String = "SELECT APPLIEDRESULTM FROM CLASS_CLASSINFO WHERE OCID=@OCID"
        ViewState("appRst") = Convert.ToString(drCC("APPLIEDRESULTM"))

        Dim vs_sort As String = TIMS.ClearSQM(ViewState("sort"))
        Dim odt As DataTable = GET_CLASSSTUDENTSOFCLASS(OCIDValue1.Value, vs_sort)

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(odt, "NAME,IDNO")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, odt.Rows.Count, vRESDESC)

        Call Show_DataGrid(Datagrid2, odt, "SOCID")

        '整班經費審核通過及不補助
        Call OpenButtons(OCIDValue1.Value)
    End Sub

    '查詢鈕
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        sSearch1()
    End Sub

    '審核確認鈕
    Private Sub AuditCheck_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AuditCheck.Click

        Call SaveDataAll_1()

    End Sub

    Private Sub SaveDataAll_1()
        Dim sql As String
        Dim CNT As Integer = 0 '資料總數
        Dim Y, N, R, S As Integer 'Y為審核成功數,N為審核失敗數,R為退件修正數,S為請選擇數

        '職類/班別 OCIDValue1.Value
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "職類/班別 有誤，請重新選擇／查詢！")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "職類/班別 有誤，請重新選擇／查詢！")
            Exit Sub
        ElseIf Not TIMS.Check_IsClosed(Val(OCIDValue1.Value), objconn) Then
            '未完成結訓動作
            Common.MessageBox(Me, "(本班尚未完成結訓動作)資料無法儲存!")
            Return
        ElseIf TIMS.CHK_STUDENTSOFCLASS_S1(Val(OCIDValue1.Value), objconn) Then
            Common.MessageBox(Me, "(本班尚未完成結訓動作)有學員參訓狀態仍為在訓中")
            Return
        End If

        Dim chk_overpay As Boolean = False '是否溢撥(預設值=false 否)
        Dim Overpay_Name As String = ""
        For Each eItem As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = eItem.FindControl("list_Verify")
            Dim hid_OverPay As HtmlInputHidden = eItem.FindControl("hid_OverPay")
            Dim link_Name As LinkButton = eItem.FindControl("link_Name")
            Dim lab_Balance As Label = eItem.FindControl("lab_Balance") '剩餘可用餘額
            Dim lab_SubSidyCost As Label = eItem.FindControl("lab_SubSidyCost") '補助費用(本次補助費用)

            '目前尚未審過 ／ '選擇審核通過者
            If (AppliedStatusM.Enabled = True) AndAlso AppliedStatusM.SelectedValue = "Y" Then
                '執行計算重新計算
                hid_OverPay.Value = (If(lab_Balance.Text = "", 0, CInt(lab_Balance.Text)) - If(lab_SubSidyCost.Text = "", 0, CInt(lab_SubSidyCost.Text)))
                '剩餘可用餘額 減去 本次補助費用之後的金額是否有負數
                If CInt(hid_OverPay.Value) < 0 Then
                    chk_overpay = True
                    Overpay_Name &= String.Concat(If(Overpay_Name <> "", "\n", ""), "學員:", link_Name.Text)
                End If
            End If
        Next

        If chk_overpay = True Then
            Page.RegisterStartupScript("Errmsg", "<script>alert('" & ("本班學員,於本次補助審核後,其剩餘可用額度將會變成負數,造成溢撥狀況,煩請再確認補助金額!" & "\n" & Overpay_Name).ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
            Exit Sub
        End If

        CNT = 0
        Y = 0 : N = 0 : R = 0 : S = 0
        For Each item As DataGridItem In Datagrid2.Items
            Dim AppliedStatusM As DropDownList = item.FindControl("list_Verify")

            Y += If(Convert.ToString(AppliedStatusM.SelectedValue) = "Y", 1, 0)
            N += If(Convert.ToString(AppliedStatusM.SelectedValue) = "N", 1, 0)
            R += If(Convert.ToString(AppliedStatusM.SelectedValue) = "R", 1, 0)
            Select Case Convert.ToString(AppliedStatusM.SelectedValue)
                Case "Y", "N", "R"
                Case Else '"0"
                    S += 1
            End Select
            CNT += 1
        Next

        AuditNumPanel.Visible = True '經費審核確認 開啟
        If CNT = 0 OrElse (Y + N + R) = 0 OrElse S > 0 Then
            If S > 0 Then
                '學員審核只要一筆為請選擇，就給警告
                Common.MessageBox(Me, "學員審核狀態有「請選擇」！")
                Exit Sub
            End If
            Common.MessageBox(Me, "學員資料數 異常！")
            Exit Sub
        End If

        'Button3_Click(sender, e)
        '儲存區如下順序有別
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dr2 As DataRow = Nothing
        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim objTrans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                sql = $"SELECT * FROM STUD_SUBSIDYCOST p WHERE EXISTS (SELECT 'x' FROM CLASS_STUDENTSOFCLASS c WHERE c.OCID={OCIDValue1.Value} AND c.SOCID=p.SOCID )"
                dt = DbAccess.GetDataTable(sql, da, objTrans)
                For Each item As DataGridItem In Datagrid2.Items
                    Dim list_Verify As DropDownList = item.FindControl("list_Verify") 'AppliedStatusM
                    Dim txt_VerifyNote As TextBox = item.FindControl("txt_VerifyNote") 'AppliedNote
                    Dim list_BudID As DropDownList = item.FindControl("list_BudID") 'BudID
                    'UPDATE Stud_SubsidyCost.BudgetId
                    dr = dt.Select("SOCID='" & Datagrid2.DataKeys(item.ItemIndex) & "'")(0)
                    Select Case list_Verify.SelectedIndex
                        Case 0 '請選擇
                            dr("AppliedStatusM") = Convert.DBNull
                            dr("AppliedStatus") = Convert.DBNull
                        Case 1 '審核成功
                            dr("AppliedStatusM") = "Y"
                            dr("AppliedStatus") = Convert.DBNull
                        Case 2 '審核失敗
                            dr("AppliedStatusM") = "N"
                            dr("AppliedStatus") = "0" '不撥款
                        Case 3 '退件修正
                            dr("AppliedStatusM") = "R"
                            dr("AppliedStatus") = "0" '不撥款
                        Case Else
                            dr("AppliedStatusM") = Convert.DBNull
                            dr("AppliedStatus") = Convert.DBNull
                    End Select
                    dr("BudID") = If(list_BudID.SelectedValue <> "", list_BudID.SelectedValue, Convert.DBNull) '預算別
                    dr("AppliedNote") = If(txt_VerifyNote.Text <> "", txt_VerifyNote.Text, Convert.DBNull)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                Next
                DbAccess.UpdateDataTable(dt, da, objTrans)

                sql = $"SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID={OCIDValue1.Value}"
                dt2 = DbAccess.GetDataTable(sql, da2, objTrans)
                For Each item As DataGridItem In Datagrid2.Items
                    Dim list_Verify As DropDownList = item.FindControl("list_Verify") 'AppliedStatusM
                    Dim txt_VerifyNote As TextBox = item.FindControl("txt_VerifyNote") 'AppliedNote
                    Dim list_BudID As DropDownList = item.FindControl("list_BudID") 'BudID'預算別
                    'Dim AppliedStatusM As DropDownList = item.FindControl("list_Verify")
                    'Dim AppliedNote As TextBox = item.FindControl("txt_VerifyNote")
                    'Dim BudID As DropDownList = item.FindControl("list_BudID")
                    'UPDATE Class_StudentsOfClass.BudgetId
                    dr2 = dt2.Select($"SOCID='{Datagrid2.DataKeys(item.ItemIndex)}'")(0)
                    dr2("BudgetId") = If(list_BudID.SelectedValue <> "", list_BudID.SelectedValue, Convert.DBNull) '預算別
                    dr2("ModifyAcct") = sm.UserInfo.UserID
                    dr2("ModifyDate") = Now
                Next
                DbAccess.UpdateDataTable(dt2, da2, objTrans)

                'R:若有一筆選退件修正，則該班開班基本資料的學員經費審核狀態為退件修正
                '學員只有審核成功和失敗，只要1筆為成功，則該班開班基本資料的學員經費審核狀態為審核成功
                'N:學員全部為審核失敗，則該班開班基本資料的學員經費審核狀態為審核失敗
                Dim ARMValue As String = If(R > 0, "R", If(Y > 0, "Y", "N"))

                Dim u_sql As String = " UPDATE CLASS_CLASSINFO SET APPLIEDRESULTM=@AppliedResultM WHERE OCID=@OCID "
                Using oCmd As New SqlCommand(u_sql, TransConn, objTrans)
                    With oCmd
                        .Parameters.Clear()
                        .Parameters.Add("AppliedResultM", SqlDbType.VarChar).Value = ARMValue
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                        .ExecuteNonQuery()
                    End With
                End Using
                DbAccess.CommitTrans(objTrans)
            Catch ex As Exception
                DbAccess.RollbackTrans(objTrans)

                TIMS.LOG.Error(ex.Message, ex)
                Common.MessageBox(Me, "儲存失敗，請重新執行，補助審核！")
                Exit Sub
            End Try
        End Using

        '查詢鈕
        'Button1_Click(sender, e)
        sSearch1()
        'End If

        '整班經費審核通過及不補助 LOG
        'Call UPDATE_Stud_SubsidyCostLOG(OCIDValue1.Value)

        'Try
        '    'Dim objconn As SqlConnection
        '    'TIMS.TestDbConn(Me, objconn, True)
        '    'objConn = DbAccess.GetConnection()
        '    'If objConn.State = ConnectionState.Closed Then objConn.Open()

        'Catch ex As Exception
        '    DbAccess.RollbackTrans(objTrans)
        '    Common.MessageBox(Me, "儲存失敗，請重新執行，補助審核！")
        '    Exit Sub
        'End Try
    End Sub

    '還原審核確認鈕
    Protected Sub AuditCheckR_Click(sender As Object, e As EventArgs) Handles AuditCheckR.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Return

        '學員經費審核結果 (Null未審核)
        Dim sqlStr As String = $"UPDATE CLASS_CLASSINFO SET AppliedResultM=null where OCID={OCIDValue1.Value}"
        DbAccess.ExecuteNonQuery(sqlStr, objconn)

        sSearch1()
    End Sub

    '單一班級查詢1
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Style("display") = "none"
    End Sub

    '單一班級查詢2
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        If e Is Nothing Then Return
        If e.CommandName Is Nothing Then Return
        If e.CommandArgument Is Nothing Then Return

        Select Case e.CommandName
            Case "back" '還原功能
                If e.CommandArgument = "" Then Return
                Dim strCmdArg As String = e.CommandArgument
                Hid_SOCID.Value = TIMS.GetMyValue(strCmdArg, "SOCID")
                Hid_OCID.Value = TIMS.GetMyValue(strCmdArg, "OCID")
                Hid_SOCID.Value = TIMS.ClearSQM(Hid_SOCID.Value)
                Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
                If Hid_SOCID.Value = "" OrElse Hid_OCID.Value = "" Then Return

                Dim sqlStr As String = ""
                If Convert.ToString(Hid_SOCID.Value) <> "" Then
                    '學員經費審核狀態-申請 (NULL未審核)
                    sqlStr = $"UPDATE STUD_SUBSIDYCOST SET APPLIEDSTATUSM=null where SOCID={Hid_SOCID.Value}"
                    DbAccess.ExecuteNonQuery(sqlStr, objconn)
                End If

                If OCIDValue1.Value <> "" Then
                    '學員經費審核結果 (Null未審核)
                    sqlStr = $"UPDATE CLASS_CLASSINFO SET APPLIEDRESULTM=null where OCID={Hid_OCID.Value}"
                    DbAccess.ExecuteNonQuery(sqlStr, objconn)
                End If

                Common.MessageBox(Me, "還原成功")
                '重新執行查詢
                'Button1_Click(Me, e)
                sSearch1()

            Case "Link" 'linkName (link_Name)
                'Dim strCmdArg As String = String.Empty
                'strCmdArg &= "&IDNO=" & Convert.ToString(drData("IDNO"))
                'strCmdArg &= "&Name=" & Convert.ToString(drData("Name"))
                'strCmdArg &= "&STDate=" & Common.FormatDate(Convert.ToString(drData("STDate")))
                'strCmdArg &= "&ActNo=" & Convert.ToString(drData("ActNO"))
                'strCmdArg &= "&SOCID=" & Convert.ToString(drData("SOCID"))
                'linkName.CommandArgument = strCmdArg
                Dim s_CCmdArg1 As String = e.CommandArgument
                If s_CCmdArg1 = "" Then Return

                Dim sCmdArg As String = ""
                sCmdArg &= "&IDNO=" & TIMS.GetMyValue(s_CCmdArg1, "IDNO")
                sCmdArg &= "&Name=" & TIMS.GetMyValue(s_CCmdArg1, "Name")
                sCmdArg &= "&STDate=" & TIMS.GetMyValue(s_CCmdArg1, "STDate")
                sCmdArg &= "&ActNo=" & TIMS.GetMyValue(s_CCmdArg1, "ActNo")
                sCmdArg &= "&SOCID=" & TIMS.GetMyValue(s_CCmdArg1, "SOCID")
                Call KeepSearch()
                Dim s_FUNID As String = TIMS.Get_MRqID(Me) 'Request("ID")
                TIMS.Utl_Redirect1(Me, "../13/SD_13_003_Bligate.aspx?ID=" & s_FUNID & sCmdArg)

        End Select
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Const Msg_目前申請總額 As String = "學員目前已申請未核准補助金總額(超過剩餘可用餘額以紅字表示)"
        Dim vs_appRst As String = TIMS.ClearSQM(ViewState("appRst"))
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim list_VerifyAll As DropDownList = e.Item.FindControl("list_VerifyAll")
                list_VerifyAll.Enabled = False
                Select Case vs_appRst
                    Case "Y"
                        'list_VerifyAll.Enabled = False
                        TIMS.Tooltip(list_VerifyAll, "班級，學員已審核確認!全選功能取消!!")
                    Case "N"
                        '學員全部為審核失敗
                        'list_VerifyAll.Enabled = False
                        TIMS.Tooltip(list_VerifyAll, "班級，學員已審核失敗!全選功能取消!!")
                    Case Else
                        list_VerifyAll.Enabled = True
                        'listVerifyAll.Attributes.Add("onChange", "SelectAll();")
                        list_VerifyAll.Attributes.Add("onChange", "SelectAll_J();")
                End Select

                Dim mysort As New System.Web.UI.WebControls.Image
                Dim vs_sort As String = Convert.ToString(ViewState("sort"))
                Select Case vs_sort 'Me.ViewState("sort")
                    Case "StudentID"
                        mysort.ImageUrl = "../../images/SortUp.gif"
                        e.Item.Cells(cst_學號).Controls.Add(mysort)
                    Case "StudentID DESC"
                        mysort.ImageUrl = "../../images/SortDown.gif"
                        e.Item.Cells(cst_學號).Controls.Add(mysort)
                    Case "IDNO"
                        mysort.ImageUrl = "../../images/SortUp.gif"
                        e.Item.Cells(cst_身分證號碼).Controls.Add(mysort)
                    Case "IDNO DESC"
                        mysort.ImageUrl = "../../images/SortDown.gif"
                        e.Item.Cells(cst_身分證號碼).Controls.Add(mysort)
                End Select

            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drData As DataRowView = e.Item.DataItem
                Dim labStudentID As Label = e.Item.FindControl("lab_StudentID")
                Dim linkName As LinkButton = e.Item.FindControl("link_Name")
                Dim labIDNO As Label = e.Item.FindControl("lab_IDNO")
                Dim listVerify As DropDownList = e.Item.FindControl("list_Verify")
                Dim btn_BackVerify As LinkButton = e.Item.FindControl("btn_BackVerify") '還原鈕
                Dim txtVerifyNote As TextBox = e.Item.FindControl("txt_VerifyNote")
                Dim labActNO As Label = e.Item.FindControl("lab_ActNO")

                Dim listBudID As DropDownList = e.Item.FindControl("list_BudID")
                Dim labEndClass As Label = e.Item.FindControl("lab_EndClass")
                Dim lab_OnClassRate As Label = e.Item.FindControl("lab_OnClassRate") '出席達3/4
                Dim lab_IsSubSidy As Label = e.Item.FindControl("lab_IsSubSidy") '是否補助
                Dim labTotal As Label = e.Item.FindControl("lab_Total") '總費用
                Dim lab_SubSidyCost As Label = e.Item.FindControl("lab_SubSidyCost") '補助費用
                Dim labPersonalCost As Label = e.Item.FindControl("lab_PersonalCost") '個人支付
                Dim lab_Balance As Label = e.Item.FindControl("lab_Balance") '剩餘可用餘額
                Dim labOtherGovApply As Label = e.Item.FindControl("lab_OtherGovApply") '其他申請中金額 labOtherGovApply  
                'Dim lab_Star As Label = e.Item.FindControl("lab_Star")
                'Dim totalSubSidy, sumSubSidy As Integer
                'Dim sDate, eDate As String
                Dim hid_OverPay As HtmlInputHidden = e.Item.FindControl("hid_OverPay")

                listBudID.Enabled = False '預算別鎖定
                labStudentID.Text = Convert.ToString(drData("StudentID"))
                linkName.Text = Convert.ToString(drData("Name"))
                labIDNO.Text = Convert.ToString(drData("IDNO"))

                Dim strCmdArg As String = String.Empty
                strCmdArg &= "&IDNO=" & Convert.ToString(drData("IDNO"))
                strCmdArg &= "&Name=" & Convert.ToString(drData("Name"))
                strCmdArg &= "&STDate=" & Common.FormatDate(Convert.ToString(drData("STDate")))
                strCmdArg &= "&ActNo=" & Convert.ToString(drData("ActNO"))
                strCmdArg &= "&SOCID=" & Convert.ToString(drData("SOCID"))
                linkName.CommandArgument = strCmdArg

                '總費用、補助費用、個人支付
                lab_SubSidyCost.Text = drData("SumOfMoney") '本次補助費用。

                labPersonalCost.Text = drData("PayMoney")
                labTotal.Text = drData("SumOfMoney") + drData("PayMoney")
                '餘額、其他補助
                'If sm.UserInfo.Years < 2008 Then totalSubSidy = 30000 Else totalSubSidy = 50000
                '可用補助額
                Dim iTotalSubSidy As Integer = TIMS.Get_3Y_SupplyMoney()
                '含職前webservice
                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drData("IDNO").ToString(), drData("STDate").ToString(), "", "Y", objconn)
                iTotalSubSidy -= SubsidyCost
                'e.Item.Cells(cst_剩餘可用餘額).Text = iTotalSubSidy
                lab_Balance.Text = iTotalSubSidy

                'Dim iTotalSubSidy As Integer = TIMS.Get_3Y_SupplyMoney(Me)
                'labBalance.Text = totalSubSidy - Get_SubSidyCost(drData("IDNO"), drData("STDate"), "Y", sDate, eDate, objConn)
                '剩餘可用餘額= '可用補助額 - '政府已補助經費
                'labBalance.Text = totalSubSidy - TIMS.Get_SubsidyCost(drData("IDNO"), drData("STDate"), , , objconn)
                '剩餘可用餘額= '可用補助額 - '政府已補助經費
                'lab_Balance.Text = totalSubSidy - Val(drData("GovCost"))

                '尚未審核者需檢查剩餘可用餘額 減去 本次補助費用之後的金額 (是否有負數)
                hid_OverPay.Value = ""
                If Convert.ToString(drData("AppliedStatusM")) <> "N" Then
                    '除了不通過之外，其餘金額檢查一次
                    hid_OverPay.Value = (If(lab_Balance.Text = "", 0, CInt(lab_Balance.Text)) - If(lab_SubSidyCost.Text = "", 0, CInt(lab_SubSidyCost.Text)))
                End If

                If lab_Balance.Text <> "" AndAlso iTotalSubSidy < 0 Then
                    lab_Balance.ForeColor = Color.Red
                End If

                'labOtherGovApply.Text = Get_SubSidyCost(drData("IDNO"), drData("STDate"), "N", sDate, eDate, objConn)
                '其他申請中金額
                'labOtherGovApply.Text = TIMS.Get_SubsidyCost(drData("IDNO"), drData("STDate"), drData("SOCID"), "N", objconn)
                'labOtherGovApply.Text = drData("GovAppl2") 'Get_SubSidyCost(drData("IDNO"), drData("STDate"), "N", sDate, eDate, objConn)
                '其他申請中金額
                labOtherGovApply.Text = Val(drData("GovAppl2"))

                Dim sDate As String = ""
                Dim eDate As String = ""
                e.Item.Cells(cst_其他申請中金額).ToolTip = Msg_目前申請總額 ' "學員目前已申請未核准補助金總額(超過剩餘可用餘額以紅字表示)"
                e.Item.Cells(cst_學號).ToolTip = TIMS.Get_StudSubSidyCostTooltip(Convert.ToString(drData("IDNO")), sDate, eDate, objconn)
                e.Item.Cells(cst_姓名).ToolTip = e.Item.Cells(cst_學號).ToolTip
                e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_學號).ToolTip
                '審核
                txtVerifyNote.Text = Convert.ToString(drData("AppliedNote"))
                labActNO.Text = Convert.ToString(drData("ActNO"))
                listBudID = TIMS.Get_Budget(listBudID, 2, objconn)

                btn_BackVerify.Visible = False '還原鈕
                If IsDBNull(drData("AppliedStatusM")) = False Then
                    listVerify.SelectedValue = drData("AppliedStatusM")
                    btn_BackVerify.Visible = True
                    listVerify.Enabled = False
                    txtVerifyNote.Enabled = False
                    listBudID.Enabled = False '預算別鎖定
                End If
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drData("SOCID")))
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drData("OCID")))
                btn_BackVerify.CommandArgument = sCmdArg ' Convert.ToString(drData("SOCID"))

                If listBudID.Items.FindByValue(Convert.ToString(drData("BudID"))) IsNot Nothing Then
                    Common.SetListItem(listBudID, Convert.ToString(drData("BudID")))
                End If
                'If listBudID.Items.FindByValue(Convert.ToString(drData("BudID"))) IsNot Nothing Then listBudID.SelectedValue = drData("BudID")
                '是否結訓
                'If drData("CreditPoints") = True Then labEndClass.Text = "是" Else labEndClass.Text = "否"

                labEndClass.Text = If(Not Convert.IsDBNull(drData("CreditPoints")) AndAlso Convert.ToInt32(drData("CreditPoints")) = 1, "是", "否")
                '出席達2/3
                '出席達3/4 2012 
                '缺席未超過1/5 2016
                'e.Item.Cells(cst_出席達4分之3).Text = "否"
                lab_OnClassRate.Text = "是"
                If drData("TOHours") > 0 Then '出缺勤檔 請假時數
                    If CDbl(drData("TOHours")) > CDbl(drData("THours") / 5) Then
                        lab_OnClassRate.Text = "否"
                    End If
                End If
                '是否補助
                If lab_OnClassRate.Text = "是" AndAlso labEndClass.Text = "是" Then lab_IsSubSidy.Text = "是" Else lab_OnClassRate.Text = "<font color='red'>否</font>"

                '判斷開訓日是否有落在加退保期間，沒有的話打上星號。
                'labStar.Visible = Get_StudBligateData(labIDNO.Text, drData("STDate"), objconn).Equals(False)
                '判斷開訓日是否有落在加退保期間，沒有的話打上星號。
                'lab_Star.Visible = False
                'If Convert.ToString(drData("bliFlag")) = "N" Then lab_Star.Visible = True
                '20091230 andy add
                Dim hid_SubSidyCost As HtmlInputHidden = e.Item.FindControl("hid_SubSidyCost")
                'Dim hid_Name As HtmlInputHidden = e.Item.FindControl("hid_Name")
                'Dim hid_Balance As HtmlInputHidden = e.Item.FindControl("hid_Balance")
                Dim hid_vstatus As HtmlInputHidden = e.Item.FindControl("hid_vstatus")
                hid_SubSidyCost.Value = Val(lab_SubSidyCost.Text) '轉為數字

                'hid_Balance.Value = labBalance.Text
                'hid_Name.Value = Convert.ToString(drData("Name"))
                hid_vstatus.Value = If(listVerify.Enabled, "1", "2")
        End Select

    End Sub

    Private Sub Datagrid2_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles Datagrid2.SortCommand
        Dim vs_sort As String = TIMS.ClearSQM(ViewState("sort"))
        ViewState("sort") = If(vs_sort <> e.SortExpression, e.SortExpression, String.Concat(e.SortExpression, " DESC"))
        '重新執行查詢
        'Button1_Click(Me, e)
        sSearch1()
    End Sub

    '整班經費審核通過及不補助 LOG
    Sub UPDATE_Stud_SubsidyCostLOG(ByVal OCIDValue As String)
        If OCIDValue = "" Then Exit Sub

        'Dim oCmd As SqlCommand = Nothing
        Dim sql As String = " SELECT * FROM V_STUDENTINFO WHERE OCID =@OCID" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue
            dt1.Load(.ExecuteReader())
        End With

        Dim sql2 As String = ""
        sql2 &= " SELECT f.SOCID" & vbCrLf
        sql2 &= " ,f.SUMOFMONEY" & vbCrLf
        sql2 &= " ,f.PAYMONEY" & vbCrLf
        sql2 &= " ,f.APPLIEDSTATUS" & vbCrLf
        sql2 &= " ,f.APPLIEDNOTE" & vbCrLf
        sql2 &= " ,f.SUPPLYID" & vbCrLf
        sql2 &= " ,f.BUDID" & vbCrLf
        sql2 &= " ,f.MODIFYACCT" & vbCrLf
        sql2 &= " ,f.MODIFYDATE" & vbCrLf
        sql2 &= " ,f.APPLIEDSTATUSM" & vbCrLf
        sql2 &= " ,f.ALLOTDATE" & vbCrLf
        sql2 &= " ,cs.OCID" & vbCrLf
        sql2 &= " ,ss.IDNO" & vbCrLf
        sql2 &= " FROM STUD_SUBSIDYCOST f" & vbCrLf
        sql2 &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.SOCID=f.SOCID" & vbCrLf
        sql2 &= " JOIN STUD_STUDENTINFO ss on ss.SID=cs.SID" & vbCrLf
        sql2 &= " where cs.OCID=@OCID" & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt2 As New DataTable
        Dim oCmd2 As New SqlCommand(sql2, objconn)
        With oCmd2
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue
            dt2.Load(.ExecuteReader())
        End With
    End Sub

    ''' <summary>
    ''' 核銷文件審查
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!")
            Exit Sub
        End If
        'Dim vs_sort As String = TIMS.ClearSQM(ViewState("sort"))
        Dim odt As DataTable = GET_CLASSSTUDENTSOFCLASS(OCIDValue1.Value, "")
        If odt Is Nothing OrElse odt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無 學員補助申請資料!")
            Exit Sub
        End If

        Call KeepSearch()
        Dim s_FUNID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        TIMS.Utl_Redirect1(Me, "../../TC/13/TC_13_001.aspx?ID=" & s_FUNID)

    End Sub

#Region "NO USE"

    'Dim work2016 As String = TIMS.Utl_GetConfigSet("work2016")
    'If work2016 <> "Y" Then
    '    Select Case sm.UserInfo.Years
    '        Case Is <= "2011"
    '            Call TIMS.CloseDbConn(objconn)
    '            Server.Transfer("SD_13_003_00.aspx?ID=" & Request("ID"))
    '            Exit Sub
    '        Case Is <= "2015"
    '            Call TIMS.CloseDbConn(objconn)
    '            Server.Transfer("SD_13_003_15.aspx?ID=" & Request("ID"))
    '            Exit Sub
    '    End Select
    'End If

    '檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
    'Function Get_StudBligateData(ByVal idno As String,
    '                             ByVal sdate As DateTime,
    '                             ByVal tmpConn As SqlConnection) As Boolean
    '    Dim rst As Boolean = False  '沒有投保
    '    Dim sqlStr As String = String.Empty
    '    idno = TIMS.ChangeIDNO(idno)
    '    Call TIMS.OpenDbConn(tmpConn)

    '    '先檢查是否存在只有加保沒有退保的紀錄，有的話代表加保中。
    '    'sqlStr = " select count(1) cnt" & vbCrLf
    '    'sqlStr &= " from (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=4 and upper(IDNO)='" & idno & "' group by ActNo,ChangeMode) a" & vbCrLf
    '    'sqlStr &= " left join (select ActNo,ChangeMode,Max(MDate) as MDate from Stud_BligateData where ChangeMode=2 and upper(IDNO)='" & idno & "' group by ActNo,ChangeMode) b on b.ActNo=a.ActNo and b.MDate>=a.MDate" & vbCrLf
    '    'sqlStr &= " where b.MDate is null and a.MDate<=" & TIMS.to_date(sdate)
    '    'Dim cntBli As Integer = DbAccess.ExecuteScalar(sqlStr, tmpConn)

    '    sqlStr = ""
    '    sqlStr &= " select a.ActNo,a.ChangeMode,a.MDate,b.ActNo as ActNo2,b.ChangeMode as ChangeMode2,b.MDate as MDate2" & vbCrLf
    '    sqlStr &= " from (select ActNo,ChangeMode,Max(MDate) MDate from Stud_BligateData where ChangeMode=4 and IDNO='" & idno & "' and MDate<= " & TIMS.to_date(sdate) & " group by ActNo,ChangeMode ) a "
    '    sqlStr &= " join (select ActNo,ChangeMode,Max(MDate) MDate from Stud_BligateData where ChangeMode=2 and IDNO='" & idno & "' group by ActNo,ChangeMode) b on b.ActNo=a.ActNo and b.MDate>=a.MDate "
    '    sqlStr &= " order by a.MDate desc "
    '    Dim dt As New DataTable
    '    Dim oCmd As New SqlCommand(sqlStr, tmpConn)
    '    With oCmd
    '        .Parameters.Clear()
    '        dt.Load(.ExecuteReader())
    '    End With
    '    If dt.Rows.Count = 0 Then Return False
    '    'If dt.Rows.Count > 0 Then rst = True
    '    '開訓前不存在只有加保紀錄的資料時，檢查開訓前最後一筆加保的退保紀錄是否在開訓後
    '    If Not rst Then
    '        If dt.Rows.Count > 0 Then
    '            For Each dr As DataRow In dt.Rows
    '                If CDate(dr("MDate2")) >= sdate Then
    '                    rst = False
    '                    Exit For
    '                Else
    '                    rst = True
    '                End If
    '            Next
    '        End If
    '    End If
    '    Return rst
    'End Function
#End Region

End Class