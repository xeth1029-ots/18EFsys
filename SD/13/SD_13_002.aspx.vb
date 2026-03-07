Partial Class SD_13_002
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    Const cst_是否取得結訓資格 As Integer = 3
    'Const cst_出席達2分之3 = 4
    'Const cst_出席達3分之4 = 4
    Const cst_缺席未超過5分之1 As Integer = 4
    Const cst_是否補助 As Integer = 5
    'Const cst_總費用 = 6
    'Const cst_補助費用 = 7
    'Const cst_個人支付 = 8
    Const cst_剩餘可用餘額 As Integer = 9
    Const cst_其他申請中金額 As Integer = 10
    'Const cst_撥款 As Integer = 11
    Const cst_撥款日期 As Integer = 12
    'Const cst_撥款備註 As Integer = 13
    'Const cst_預算別 As Integer = 14

    Dim objconn As SqlConnection
    'Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()
    Dim flag_ROC As Boolean = False

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        flag_ROC = TIMS.CHK_REPLACE2ROC_YEARS()

        'Dim work2016 As String = TIMS.Utl_GetConfigSet("work2016")
        'If work2016 <> "Y" Then
        '    Select Case sm.UserInfo.Years.ToString()
        '        Case Is <= "2011"
        '            Call TIMS.CloseDbConn(objconn)
        '            Server.Transfer("SD_13_002_00.aspx?ID=" & Request("ID"))
        '            Exit Sub
        '        Case Is <= "2015"
        '            Call TIMS.CloseDbConn(objconn)
        '            Server.Transfer("SD_13_002_15.aspx?ID=" & Request("ID"))
        '            Exit Sub
        '    End Select
        'End If
        If Not IsPostBack Then
            Call Create1()
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

    End Sub

    Sub Create1()
        msg.Text = ""
        DataGridTable.Style("display") = "none"
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        '經費審核 AuditList  Y/N/NULL-已通過/不通過/審核中
        With AuditList
            .Items.Clear()
            .Items.Insert(0, New ListItem("已通過", "Y"))
            .Items.Insert(1, New ListItem("不通過", "N"))
            .Items.Insert(2, New ListItem("審核中", "Null"))
        End With
        'AuditList.SelectedIndex = 0
        Common.SetListItem(AuditList, "Y")

        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)  '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級

        btnSch.Attributes("onclick") = "return CheckSearch();"
        btnSave.Attributes("onclick") = "return chkSave();"

        SCB_Budget = TIMS.Get_Budget(SCB_Budget, 22, objconn)  '預算來源  '4:含 ECFA(協助)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        ''經費審核 AuditList  Y/N/NULL-已通過/不通過/審核中
        'Dim gv_AuditList As String = TIMS.GetListValue(Me.AuditList)
        Select Case e.Item.ItemType
            Case ListItemType.Header
                '撥款' 請選擇/已撥款
                Dim DropDownList1 As DropDownList = e.Item.FindControl("DropDownList1") '撥款
                Dim chkAll As CheckBox = e.Item.FindControl("chkAll")
                DropDownList1.Attributes("onchange") = "SelectAll(0,this)"
                chkAll.Attributes("onclick") = "SelectAll(1,this)"
                e.Item.CssClass = "head_navy"
                e.Item.Cells(cst_其他申請中金額).ToolTip = "學員目前已申請未核准補助金總額(超過剩餘可用餘額以紅字表示)"
                If Me.ViewState("sort") <> "" Then
                    Dim mylabel As String = ""
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "StudentID", "StudentID DESC"
                            mylabel = "StudentID"
                            i = 0
                            mysort.ImageUrl = "../../images/SortDown.gif"
                            If Me.ViewState("sort") = "StudentID" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "IDNO", "IDNO DESC"
                            mylabel = "StudentID"
                            i = 2
                            mysort.ImageUrl = "../../images/SortDown.gif"
                            If Me.ViewState("sort") = "IDNO" Then mysort.ImageUrl = "../../images/SortUp.gif"
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)

                    'If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then DropDownList1.Enabled = False
                    '經費審核 AuditList  Y/N/NULL-已通過/不通過/審核中
                    Select Case Hid_AuditList.Value
                        Case "Y"
                            DropDownList1.Enabled = True
                        Case Else '"N"
                            DropDownList1.Enabled = False
                    End Select
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim iFlag As Integer = 0      '是否補助
                Dim iFlagStudy As Integer = 0 '出勤flag,未滿2/3=0；達2/3=1
                Dim drv As DataRowView = e.Item.DataItem
                '撥款' 0:請選擇/1:已撥款
                Dim AppliedStatus As DropDownList = e.Item.FindControl("AppliedStatus")
                Dim AppliedNote As TextBox = e.Item.FindControl("AppliedNote")
                Dim txtAllotDate As TextBox = e.Item.FindControl("txtAllotDate")

                Dim ibtDate As ImageButton = e.Item.FindControl("ibtDate")
                Dim iTotal As Integer = 0  '可用補助額(2007年3年3萬)(2008年3年5萬)
                Dim LabBudID As Label = e.Item.FindControl("LabBudID")
                'sqls1 &= " ,g.BudId" & vbCrLf 'GET_BudgetName
                LabBudID.Text = TIMS.GET_BudgetName(Convert.ToString(drv("BudID")), objconn)
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                e.Item.Cells(cst_是否取得結訓資格).Text = "否"
                If Convert.ToString(drv("CreditPoints")) = "1" Then
                    e.Item.Cells(cst_是否取得結訓資格).Text = "是"
                    iFlag += 1
                End If
                e.Item.Cells(cst_缺席未超過5分之1).Text = "否"
                If Convert.ToString(drv("THours")) <> "" Then
                    If drv("THours") > 0 Then
                        Dim iVal1 As Double = Val(drv("COUNTHOURS")) '- Val(drv("COUNTHOURS2"))
                        TIMS.Tooltip(e.Item.Cells(cst_缺席未超過5分之1), "出席時數:" & (drv("THours") - iVal1) & "/" & drv("THours"))
                        If iVal1 / drv("THours") <= 1 / 5 Then
                            e.Item.Cells(cst_缺席未超過5分之1).Text = "是"
                            iFlagStudy = 1
                        End If
                    End If
                End If
                'e.Item.Cells(cst_出席達3分之4).Text = "否"
                'If drv("THours") > 0 Then
                '    If (drv("THours") - drv("CountHours")) / drv("THours") >= 3 / 4 Then
                '        e.Item.Cells(cst_出席達3分之4).Text = "是"
                '        iFlagStudy = 1
                '    End If
                'End If
                '可用補助額
                iTotal = TIMS.Get_3Y_SupplyMoney()
                '含職前webservice
                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)
                iTotal -= SubsidyCost
                e.Item.Cells(cst_剩餘可用餘額).Text = iTotal
                'If iTotal - Val(drv("SumOfMoney")) >= 0 Then
                '    e.Item.Cells(cst_剩餘可用餘額).Text = iTotal
                'Else
                '    e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & iTotal & "</font>"
                '    TIMS.Tooltip(e.Item.Cells(cst_剩餘可用餘額), "剩餘可用餘額 不足補助金!")
                'End If
                'Dim mm As String = ""
                'mm &= ",SubsidyCost:" & CStr(SubsidyCost)
                'mm &= ",iTotal:" & CStr(iTotal)
                'mm &= ",GovAppl2:" & CStr(drv("GovAppl2"))
                'mm &= ",SumOfMoney:" & CStr(drv("SumOfMoney"))
                If iTotal - drv("SumOfMoney") >= 0 AndAlso drv("GovAppl2") > iTotal - drv("SumOfMoney") Then
                    e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_其他申請中金額), "其他申請中金額 不足補助金!") ' & mm)
                End If
                'If drv("GovAppl2") > (iTotal - CInt(drv("SumOfMoney"))) Then
                '    e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                'End If
                e.Item.Cells(cst_是否補助).Text = "否"
                If iFlag = 1 Then e.Item.Cells(cst_是否補助).Text = "是"
                If iFlag = 0 OrElse iFlagStudy = 0 Then
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_是否補助), " 學分為 0 或 出勤未滿2/3 時不補助!")
                End If

                '撥款' 0:請選擇/1:已撥款
                If Convert.ToString(drv("AppliedStatus")) = "1" Then
                    AppliedStatus.SelectedIndex = 1 '已撥款
                Else
                    AppliedStatus.SelectedIndex = 0 '請選擇
                End If
                AppliedNote.Text = Convert.ToString(drv("AppliedNote"))

                '經費審核
                'Dim v_AuditList As String = TIMS.GetListValue(Me.AuditList)
                '經費審核 AuditList  Y/N/NULL-已通過/不通過/審核中
                Select Case Hid_AuditList.Value'v_AuditList
                    Case "Y"
                        AppliedStatus.Enabled = True '撥款
                        AppliedNote.Enabled = True '撥款備註
                    Case Else
                        AppliedStatus.Enabled = False '撥款
                        AppliedNote.Enabled = False '撥款備註
                End Select

                If Convert.ToString(drv("IDNO")) <> "" Then
                    e.Item.Cells(cst_姓名).ToolTip = TIMS.Search_Stud_SubsidyCost(Convert.ToString(drv("IDNO")), objconn)
                    e.Item.Cells(cst_學號).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                    e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                End If
                txtAllotDate.Text = ""
                If Convert.ToString(drv("AllotDate")) <> "" Then
                    txtAllotDate.Text = If(flag_ROC, TIMS.Cdate17(drv("AllotDate")), TIMS.Cdate3(drv("AllotDate"))) 'edit，by:20181001
                End If

                ibtDate.Attributes.Add("onclick", "selDate(" & e.Item.ItemIndex + 1 & ");return false;")
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If
        'btnSch_Click(Me, e)
        Call Search1()
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If center.Text <> "" Then RstMemo &= String.Concat("&center=", center.Text)
        If RIDValue.Value <> "" Then RstMemo &= String.Concat("&RID=", RIDValue.Value)
        Return RstMemo
    End Function

    '查詢 SQL 
    Sub Search1()
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

        Dim cPlanid As String = CStr(drCC("planid"))
        Dim cComIDNO As String = CStr(drCC("comidno"))
        Dim cSeqNo As String = CStr(drCC("seqno"))

        Dim vSBudget As String = TIMS.GetCblValue(SCB_Budget)
        vSBudget = TIMS.CombiSQM2IN(vSBudget)
        If vSBudget.IndexOf(TIMS.cst_Budget_不拘id) > -1 Then vSBudget = "" '(含有不拘)清理查詢資料

        Dim sqls1 As String = ""
        sqls1 &= " SELECT d.SOCID" & vbCrLf
        sqls1 &= " ,dbo.FN_CSTUDID2(d.StudentID) StudentID" & vbCrLf
        sqls1 &= " ,e.Name" & vbCrLf
        sqls1 &= " ,e.IDNO" & vbCrLf
        '除數可能有溢位問題，無條件捨去餘2位數。
        sqls1 &= " ,CASE WHEN b.TotalCost>=ISNULL(c.Total2,0) THEN FLOOR(ISNULL(b.TotalCost,0)/ISNULL(b.TNum,1)) ELSE FLOOR(ISNULL(c.Total2,0)/ISNULL(b.TNum,1)) END Total" & vbCrLf
        sqls1 &= " ,d.CreditPoints ,a.THours" & vbCrLf
        'sqls1 += " ,ISNULL(f.CountHours,0) CountHours" & vbCrLf
        sqls1 &= " ,ISNULL(f.COUNTHOURS,0) COUNTHOURS" & vbCrLf
        'sqls1 += " ,ISNULL(f.COUNTHOURS2,0) COUNTHOURS2" & vbCrLf ' 扣除「喪假」時數
        sqls1 &= " ,e.DegreeID ,d.StudStatus ,d.MIdentityID ,a.STDate ,g.SOCID Exist ,g.SumOfMoney" & vbCrLf
        sqls1 &= " ,g.PayMoney ,g.AppliedStatus" & vbCrLf
        sqls1 &= " ,g.AllotDate ,g.AppliedNote" & vbCrLf
        'sqls1 &= " ,ISNULL(g.BudID,d.BudgetID) BudgetID" & vbCrLf
        sqls1 &= " ,g.BudId" & vbCrLf 'GET_BudgetName
        'sqls1 &= " ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2" & vbCrLf
        'sqls1 &= " ,dbo.FN_GET_GOVCOST(e.IDNO,convert(varchar,a.STDate,111)) GovAppl2" & vbCrLf
        sqls1 &= " ,dbo.FN_GET_GOVCOST2(e.IDNO, convert(varchar,a.STDate,111)) GovAppl2"

        sqls1 &= " FROM CLASS_CLASSINFO a WITH(NOLOCK)" & vbCrLf
        sqls1 &= " JOIN PLAN_PLANINFO b WITH(NOLOCK) ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo" & vbCrLf
        sqls1 &= " JOIN CLASS_STUDENTSOFCLASS d WITH(NOLOCK) ON a.OCID=d.OCID" & vbCrLf
        sqls1 &= " JOIN STUD_STUDENTINFO e WITH(NOLOCK) ON d.SID=e.SID" & vbCrLf
        sqls1 &= " JOIN STUD_SUBSIDYCOST g WITH(NOLOCK) ON d.SOCID=g.SOCID" & vbCrLf
        'c: COSTITEM (G)
        sqls1 &= " LEFT JOIN (" & vbCrLf
        sqls1 &= "   SELECT PlanID" & vbCrLf
        sqls1 &= "   ,ComIDNO" & vbCrLf
        sqls1 &= "   ,SeqNo" & vbCrLf
        sqls1 &= "   ,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)) Total" & vbCrLf
        sqls1 &= "   ,Sum(ISNULL(OPrice,1)*ISNULL(Itemage,1)) Total2" & vbCrLf
        sqls1 &= "   FROM PLAN_COSTITEM WITH(NOLOCK)" & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sqls1 &= " WHERE COSTMODE = 5" & vbCrLf
        Else
            sqls1 &= " WHERE COSTMODE <> 5" & vbCrLf
        End If
        sqls1 &= " and planid ='" & cPlanid & "'" & vbCrLf
        sqls1 &= " and ComIDNO ='" & cComIDNO & "'" & vbCrLf
        sqls1 &= " and SeqNo ='" & cSeqNo & "'" & vbCrLf
        sqls1 &= " Group By PlanID,ComIDNO,SeqNo" & vbCrLf
        sqls1 &= " ) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo" & vbCrLf
        'f: STUD_TURNOUT2,CLASS_STUDENTSOFCLASS (G)
        sqls1 &= " LEFT JOIN (" & vbCrLf
        'select * from STUD_TURNOUT2 WHERE rownum <=10
        '喪假(LEAVEID:05)。99:(使用者輸入)
        sqls1 &= " SELECT t.SOCID" & vbCrLf
        sqls1 &= " ,SUM(CASE WHEN t.LEAVEID IS NULL THEN t.Hours END) COUNTHOURS" & vbCrLf
        sqls1 &= " FROM STUD_TURNOUT2 t WITH(NOLOCK)" & vbCrLf
        sqls1 &= " JOIN CLASS_STUDENTSOFCLASS cs WITH(NOLOCK) on cs.socid =t.socid" & vbCrLf
        sqls1 &= " WHERE cs.OCID ='" & OCIDValue1.Value & "'" & vbCrLf
        sqls1 &= " Group By t.SOCID" & vbCrLf
        sqls1 &= " ) f ON f.SOCID=d.SOCID" & vbCrLf
        sqls1 &= " WHERE 1=1" & vbCrLf
        sqls1 &= " AND a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        sqls1 &= " and g.Budid IS NOT NULL" & vbCrLf
        If vSBudget <> "" Then sqls1 &= " and g.Budid IN (" & vSBudget & ")" & vbCrLf '預算別有選擇
        sqls1 &= " and a.AppliedResultR IN ('Y','C')" & vbCrLf 'Y 通過 'C 全班學員資料確認
        Select Case AuditList.SelectedIndex
            Case 0
                sqls1 &= " AND a.AppliedResultM='Y' AND g.AppliedStatusM='Y'" & vbCrLf
            Case 1
                sqls1 &= " AND a.AppliedResultM='N'" & vbCrLf
            Case 2
                sqls1 &= " AND a.AppliedResultM IS Null" & vbCrLf
        End Select
        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sqls1 &= " order by e." & Me.ViewState("sort").ToString
        ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
            sqls1 &= " order by dbo.FN_CSTUDID2(d.StudentID) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
        Else
            sqls1 &= " order by dbo.FN_CSTUDID2(d.StudentID)" & vbCrLf
        End If
        Dim sCmd As New SqlCommand(sqls1, objconn)
        'dt = DbAccess.GetDataTable(sqlstr, objconn)
        'SELECT SOCID,Sum(Hours) CountHours FROM STUD_TURNOUT2 Group By SOCID
        'SELECT *  FROM STUD_TURNOUT2 WHERE ROWNUM <=10
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable

        Try
            'SQL
            With sCmd
                .Parameters.Clear()
                dt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sqls1: */" & vbCrLf
            strErrmsg += sqls1 & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢")
            Exit Sub
        End Try

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then

            DataGridTable.Style("display") = "inline"
            msg.Text = ""

            '經費審核 AuditList  Y/N/NULL-已通過/不通過/審核中
            'Dim gv_AuditList As String = TIMS.GetListValue(Me.AuditList)
            Hid_AuditList.Value = TIMS.GetListValue(Me.AuditList)

            If ViewState("sort") = "" Then ViewState("sort") = "StudentID"
            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()

            btnSave.Enabled = True
            If Me.AuditList.SelectedIndex = 1 OrElse Me.AuditList.SelectedIndex = 2 Then btnSave.Enabled = False
        End If
    End Sub

    '查詢鈕
    Private Sub btnSch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSch.Click
        Call Search1()
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    Sub SAVEDATA1()
        Dim intCnt As Integer = 0
        Dim sql As String = ""
        Dim s_uParms As String = ""

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)

        sql = ""
        sql &= " UPDATE Stud_SubsidyCost "
        sql &= " set AppliedStatus=@AppliedStatus "
        sql &= " ,AllotDate= convert(date, @AllotDate) "
        sql &= " ,AppliedNote=@AppliedNote "
        sql &= " ,MODIFYDATE=GETDATE() "
        sql &= " ,MODIFYACCT=@MODIFYACCT "
        sql &= " where SOCID=@SOCID "
        Dim uCmd As New SqlCommand(sql, conn, trans)

        Try
            For i As Integer = 0 To DataGrid1.Items.Count - 1
                '撥款' 0:請選擇/1:已撥款
                Dim AppliedStatus As DropDownList = DataGrid1.Items(i).FindControl("AppliedStatus")
                Dim AllotDate As TextBox = DataGrid1.Items(i).FindControl("txtAllotDate")
                Dim AppliedNote As TextBox = DataGrid1.Items(i).FindControl("AppliedNote")
                'edit，by:20181001
                Dim myAllotDate As String = If(flag_ROC, TIMS.Cdate18(AllotDate.Text), TIMS.Cdate3(AllotDate.Text))
                With uCmd
                    .Parameters.Clear()
                    '撥款' 0:請選擇/1:已撥款
                    .Parameters.Add("AppliedStatus", SqlDbType.VarChar).Value = If(AppliedStatus.SelectedIndex = 1, "1", Convert.DBNull)
                    .Parameters.Add("AllotDate", SqlDbType.DateTime).Value = If(myAllotDate <> "", myAllotDate, Convert.DBNull) 'myAllotDate  'edit，by:20181001
                    .Parameters.Add("AppliedNote", SqlDbType.VarChar).Value = If(AppliedNote.Text <> "", AppliedNote.Text, Convert.DBNull)
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = DataGrid1.DataKeys(i) 'SOCID
                    s_uParms = TIMS.GetMyValue3(uCmd.Parameters)
                    .ExecuteNonQuery()
                End With
            Next
            intCnt = 1
            DbAccess.CommitTrans(trans) 'trans.Commit()
            'If trans IsNot Nothing Then trans.Dispose()
            'If Not sda Is Nothing Then sda.Dispose()
            'Call TIMS.CloseDbConn(conn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "sql:" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "s_uParms:" & vbCrLf
            strErrmsg += s_uParms & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            DbAccess.RollbackTrans(trans) 'trans.Rollback()
            Call TIMS.CloseDbConn(conn)

            Common.MessageBox(Page, ex.Message)
            Return
        End Try
        Call TIMS.CloseDbConn(conn)
        If trans IsNot Nothing Then trans.Dispose()

        If intCnt = 1 Then
            Common.MessageBox(Me, "儲存成功")
            'btnSch_Click(sender, e)
            Call Search1()
        End If
    End Sub

    '儲存鈕
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        ElseIf Not TIMS.Check_IsClosed(Val(OCIDValue1.Value), objconn) Then
            '未完成結訓動作
            Common.MessageBox(Me, "(本班尚未完成結訓動作)資料無法儲存!")
            Return
        ElseIf TIMS.CHK_STUDENTSOFCLASS_S1(Val(OCIDValue1.Value), objconn) Then
            Common.MessageBox(Me, "(本班尚未完成結訓動作)有學員參訓狀態仍為在訓中")
            Return
        End If

        TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql &= " SELECT c.SOCID" & vbCrLf
        sql &= " FROM Stud_SubsidyCost c" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs on cs.socid =c.socid" & vbCrLf
        sql &= " WHERE OCID= @OCID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "此班學員未建立補助撥款，請重新建立")
            Exit Sub
        End If

        For i As Integer = 0 To DataGrid1.Items.Count - 1
            '撥款' 0:請選擇/1:已撥款
            Dim AppliedStatus As DropDownList = DataGrid1.Items(i).FindControl("AppliedStatus")
            Dim AllotDate As TextBox = DataGrid1.Items(i).FindControl("txtAllotDate")
            Dim AppliedNote As TextBox = DataGrid1.Items(i).FindControl("AppliedNote")
            'edit，by:20181001
            Dim myAllotDate As String = If(flag_ROC, TIMS.Cdate18(AllotDate.Text), TIMS.Cdate3(AllotDate.Text))
            If myAllotDate <> "" AndAlso Not TIMS.IsDate1(myAllotDate) Then
                Common.MessageBox(Me, "撥款日期有誤，請重新輸入!")
                Exit Sub
            End If
        Next

        '儲存
        SAVEDATA1()
    End Sub

    '單一班級查詢1
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Style("display") = "none"
    End Sub

    '單一班級查詢2
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class