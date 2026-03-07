Partial Class SD_13_002_16
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
    'Const cst_撥款 = 11
    'Const cst_撥款日期 = 12
    'Const cst_撥款備註 = 13
    'Const cst_預算別 = 14

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
        '檢查Session是否存在 End

        Dim work2016 As String = TIMS.Utl_GetConfigSet("work2016")
        If work2016 <> "Y" Then
            Select Case sm.UserInfo.Years.ToString()
                Case Is <= "2011"
                    Call TIMS.CloseDbConn(objconn)
                    Server.Transfer("SD_13_002_00.aspx?ID=" & Request("ID"))
                    Exit Sub
                Case Is <= "2015"
                    Call TIMS.CloseDbConn(objconn)
                    Server.Transfer("SD_13_002_15.aspx?ID=" & Request("ID"))
                    Exit Sub
            End Select
        End If

        If Not IsPostBack Then Call Create1()

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

        With AuditList
            .Items.Insert(0, New ListItem("已通過", "Y"))
            .Items.Insert(1, New ListItem("不通過", "N"))
            .Items.Insert(2, New ListItem("審核中", "Null"))
        End With
        AuditList.SelectedIndex = 0

        '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        btnSch.Attributes("onclick") = "return CheckSearch();"
        btnSave.Attributes("onclick") = "return chkSave();"

        '預算來源  '4:含 ECFA(協助)
        SCB_Budget = TIMS.Get_Budget(SCB_Budget, 22, objconn)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim DropDownList1 As DropDownList = e.Item.FindControl("DropDownList1")
                Dim chkAll As CheckBox = e.Item.FindControl("chkAll")
                DropDownList1.Attributes("onchange") = "SelectAll(0,this)"
                chkAll.Attributes("onclick") = "SelectAll(1,this)"
                e.Item.CssClass = "SD_TD1"
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
                            If Me.ViewState("sort") = "StudentID" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            End If
                        Case "IDNO", "IDNO DESC"
                            mylabel = "StudentID"
                            i = 2
                            mysort.ImageUrl = "../../images/SortDown.gif"
                            If Me.ViewState("sort") = "IDNO" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            End If
                    End Select

                    If i <> -1 Then
                        e.Item.Cells(i).Controls.Add(mysort)
                    End If

                    DropDownList1.Enabled = True
                    If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then
                        DropDownList1.Enabled = False
                    End If
                End If
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim iFlag As Integer = 0      '是否補助
                Dim iFlagStudy As Integer = 0 '出勤flag,未滿2/3=0；達2/3=1
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedStatus As DropDownList = e.Item.FindControl("AppliedStatus")
                Dim AppliedNote As TextBox = e.Item.FindControl("AppliedNote")
                Dim txtAllotDate As TextBox = e.Item.FindControl("txtAllotDate")
                Dim ibtDate As ImageButton = e.Item.FindControl("ibtDate")
                Dim iTotal As Integer = 0  '可用補助額(2007年3年3萬)(2008年3年5萬)
                Dim LabBudID As Label = e.Item.FindControl("LabBudID")
                ' sqls1 &= " ,g.BudId" & vbCrLf 'GET_BudgetName
                LabBudID.Text = TIMS.GET_BudgetName(Convert.ToString(drv("BudID")), objconn)
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
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

                '可用補助額(2007年3年3萬)
                '可用補助額(2008年3年5萬)
                '可用補助額(2012年3年7萬)
                iTotal = TIMS.Get_3Y_SupplyMoney(Me)

                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)
                iTotal -= SubsidyCost

                e.Item.Cells(cst_剩餘可用餘額).Text = iTotal
                'If iTotal - Val(drv("SumOfMoney")) >= 0 Then
                '    e.Item.Cells(cst_剩餘可用餘額).Text = iTotal
                'Else
                '    e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & iTotal & "</font>"
                '    TIMS.Tooltip(e.Item.Cells(cst_剩餘可用餘額), "剩餘可用餘額 不足補助金!")
                'End If

                If iTotal - drv("SumOfMoney") >= 0 AndAlso drv("GovAppl2") > iTotal - drv("SumOfMoney") Then
                    e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_其他申請中金額), "其他申請中金額 不足補助金!")
                End If

                e.Item.Cells(cst_是否補助).Text = "否"
                If iFlag = 1 Then
                    e.Item.Cells(cst_是否補助).Text = "是"
                End If

                If iFlag = 0 OrElse iFlagStudy = 0 Then
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_是否補助), " 學分為 0 或 出勤未滿2/3 時不補助!")
                End If

                If IsDBNull(drv("AppliedStatus")) Then
                    AppliedStatus.SelectedIndex = 0 '請選擇
                Else
                    If drv("AppliedStatus") = 1 Then
                        AppliedStatus.SelectedIndex = 1 '已撥款
                    End If
                End If
                AppliedNote.Text = Convert.ToString(drv("AppliedNote"))

                AppliedStatus.Enabled = True '撥款
                AppliedNote.Enabled = True '撥款備註
                If Me.AuditList.SelectedIndex = 1 _
                    OrElse Me.AuditList.SelectedIndex = 2 Then

                    AppliedStatus.Enabled = False '撥款
                    AppliedNote.Enabled = False '撥款備註
                End If

                If Convert.ToString(drv("IDNO")) <> "" Then
                    e.Item.Cells(cst_姓名).ToolTip = TIMS.Search_Stud_SubsidyCost(Convert.ToString(drv("IDNO")), objconn)
                    e.Item.Cells(cst_學號).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                    e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                End If

                If Convert.ToString(drv("AllotDate")) <> "" Then txtAllotDate.Text = Convert.ToDateTime(Convert.ToString(drv("AllotDate"))).ToString("yyyy/MM/dd")
                If flag_ROC Then txtAllotDate.Text = TIMS.cdate17(drv("AllotDate"))  'edit，by:20181001
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

    '查詢 SQL 
    Sub Search1()
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If
        Dim cPlanid As String = CStr(drCC("planid"))
        Dim cComIDNO As String = CStr(drCC("comidno"))
        Dim cSeqNo As String = CStr(drCC("seqno"))

        Dim vSBudget As String = TIMS.GetCblValue(SCB_Budget)
        vSBudget = TIMS.CombiSQM2IN(vSBudget)
        If vSBudget.IndexOf(TIMS.cst_Budget_不拘id) > -1 Then vSBudget = "" '(含有不拘)清理查詢資料

        Dim sqls1 As String = ""
        sqls1 = "" & vbCrLf
        sqls1 &= " SELECT d.SOCID ,dbo.FN_CSTUDID2(d.StudentID) StudentID ,e.Name ,e.IDNO " & vbCrLf
        '除數可能有溢位問題，無條件捨去餘2位數。
        sqls1 &= "       ,CASE WHEN b.TotalCost>=ISNULL(c.Total2,0) THEN FLOOR(ISNULL(b.TotalCost,0)/ISNULL(b.TNum,1)) ELSE FLOOR(ISNULL(c.Total2,0)/ISNULL(b.TNum,1)) END Total " & vbCrLf
        sqls1 &= "       ,d.CreditPoints ,a.THours " & vbCrLf
        'sqls1 += "      ,ISNULL(f.CountHours,0) CountHours " & vbCrLf
        sqls1 &= "       ,ISNULL(f.COUNTHOURS,0) COUNTHOURS " & vbCrLf
        'sqls1 += "      ,ISNULL(f.COUNTHOURS2,0) COUNTHOURS2 " & vbCrLf ' 扣除「喪假」時數
        sqls1 &= "       ,e.DegreeID ,d.StudStatus ,d.MIdentityID ,a.STDate ,g.SOCID Exist ,g.SumOfMoney ,g.PayMoney " & vbCrLf
        sqls1 &= "       ,g.AppliedStatus ,g.AllotDate ,g.AppliedNote " & vbCrLf
        'sqls1 &= "      ,ISNULL(g.BudID,d.BudgetID) BudgetID" & vbCrLf
        sqls1 &= "       ,g.BudId " & vbCrLf 'GET_BudgetName
        sqls1 &= "       ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2 " & vbCrLf
        sqls1 &= " FROM CLASS_CLASSINFO a " & vbCrLf
        sqls1 &= " JOIN PLAN_PLANINFO b ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo " & vbCrLf
        sqls1 &= " JOIN CLASS_STUDENTSOFCLASS d ON a.OCID=d.OCID " & vbCrLf
        sqls1 &= " JOIN STUD_STUDENTINFO e ON d.SID=e.SID " & vbCrLf
        sqls1 &= " JOIN STUD_SUBSIDYCOST g ON d.SOCID=g.SOCID " & vbCrLf
        'c: COSTITEM (G)
        sqls1 &= " LEFT JOIN (" & vbCrLf
        sqls1 &= "   SELECT PlanID ,ComIDNO ,SeqNo ,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)) Total ,Sum(ISNULL(OPrice,1)*ISNULL(Itemage,1)) Total2 " & vbCrLf
        sqls1 &= "   FROM PLAN_COSTITEM " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sqls1 &= " WHERE COSTMODE = 5 " & vbCrLf
        Else
            sqls1 &= " WHERE COSTMODE <> 5 " & vbCrLf
        End If
        sqls1 &= " and planid ='" & cPlanid & "'" & vbCrLf
        sqls1 &= " and ComIDNO ='" & cComIDNO & "'" & vbCrLf
        sqls1 &= " and SeqNo ='" & cSeqNo & "'" & vbCrLf
        sqls1 &= " Group By PlanID,ComIDNO,SeqNo" & vbCrLf
        sqls1 &= " ) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo" & vbCrLf
        'f: STUD_TURNOUT2,CLASS_STUDENTSOFCLASS (G)
        sqls1 &= " LEFT JOIN ( " & vbCrLf
        'select * from STUD_TURNOUT2 WHERE rownum <=10
        '喪假(LEAVEID:05)。99:(使用者輸入)
        sqls1 &= " SELECT t.SOCID ,SUM(CASE WHEN t.LEAVEID IS NULL THEN t.Hours END) COUNTHOURS " & vbCrLf
        sqls1 &= " FROM STUD_TURNOUT2 t" & vbCrLf
        sqls1 &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.socid =t.socid " & vbCrLf
        sqls1 &= " WHERE 1=1 and cs.OCID ='" & OCIDValue1.Value & "' " & vbCrLf
        sqls1 &= " Group By t.SOCID" & vbCrLf
        sqls1 &= " ) f ON f.SOCID=d.SOCID" & vbCrLf
        sqls1 &= " WHERE 1=1" & vbCrLf
        sqls1 &= "    AND a.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        sqls1 &= "    and g.Budid IS NOT NULL " & vbCrLf
        '預算別有選擇
        If vSBudget <> "" Then sqls1 &= " and g.Budid IN (" & vSBudget & ")" & vbCrLf
        sqls1 &= " and a.AppliedResultR IN ('Y','C') " & vbCrLf 'Y 通過 'C 全班學員資料確認
        Select Case AuditList.SelectedIndex
            Case 0
                sqls1 &= " AND a.AppliedResultM='Y' AND g.AppliedStatusM='Y' " & vbCrLf
            Case 1
                sqls1 &= " AND a.AppliedResultM='N' " & vbCrLf
            Case 2
                sqls1 &= " AND a.AppliedResultM IS Null " & vbCrLf
        End Select
        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sqls1 &= " order by e." & Me.ViewState("sort").ToString
        ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
            sqls1 &= " order by dbo.FN_CSTUDID2(d.StudentID) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
        Else
            sqls1 &= " order by dbo.FN_CSTUDID2(d.StudentID) " & vbCrLf
        End If
        Dim sCmd As New SqlCommand(sqls1, objconn)
        'dt = DbAccess.GetDataTable(sqlstr, objconn)
        'SELECT SOCID,Sum(Hours) CountHours FROM STUD_TURNOUT2 Group By SOCID
        'SELECT *  FROM STUD_TURNOUT2 WHERE ROWNUM <=10
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Style("display") = "inline"
            msg.Text = ""
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

    '儲存鈕
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim intCnt As Integer = 0

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT c.SOCID " & vbCrLf
        sql += " FROM Stud_SubsidyCost c" & vbCrLf
        sql += " JOIN Class_StudentsOfClass cs on cs.socid =c.socid" & vbCrLf
        sql += " where 1=1 and OCID= @OCID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        DbAccess.Open(objconn)
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

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            sql = ""
            sql &= " UPDATE Stud_SubsidyCost "
            sql += " set AppliedStatus=@AppliedStatus"
            sql += " ,AllotDate= convert(datetime, @AllotDate, 111)"
            sql += " ,AppliedNote=@AppliedNote "
            sql += " ,ModifyAcct=@ModifyAcct,ModifyDate=getdate() "
            sql += " where SOCID=@SOCID"
            Dim uCmd As New SqlCommand(sql, conn, trans)

            'sda.UpdateCommand = New SqlCommand(sql, conn, trans)
            For i As Integer = 0 To DataGrid1.Items.Count - 1
                Dim AppliedStatus As DropDownList = Nothing
                Dim AllotDate As TextBox = Nothing
                Dim AppliedNote As TextBox = Nothing
                AppliedStatus = DataGrid1.Items(i).FindControl("AppliedStatus")
                AllotDate = DataGrid1.Items(i).FindControl("txtAllotDate")
                AppliedNote = DataGrid1.Items(i).FindControl("AppliedNote")
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("AppliedStatus", SqlDbType.VarChar).Value = IIf(AppliedStatus.SelectedIndex = 0, Convert.DBNull, "1")
                    '.Parameters.Add("AllotDate", SqlDbType.VarChar).Value = TIMS.cdate2(AllotDate.Text, "yyyy/MM/dd")
                    Dim myAllotDate As String = ""  'edit，by:20181001
                    If flag_ROC Then
                        myAllotDate = TIMS.cdate18(AllotDate.Text)  'edit，by:20181001
                    Else
                        myAllotDate = TIMS.cdate2(AllotDate.Text, "yyyy/MM/dd")  'edit，by:20181001
                    End If
                    .Parameters.Add("AllotDate", SqlDbType.VarChar).Value = myAllotDate  'edit，by:20181001
                    .Parameters.Add("AppliedNote", SqlDbType.VarChar).Value = IIf(AppliedNote.Text = "", Convert.DBNull, AppliedNote.Text)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = DataGrid1.DataKeys(i)
                    .ExecuteNonQuery()
                End With
            Next
            intCnt = 1

            trans.Commit()
            If Not trans Is Nothing Then trans.Dispose()
            'If Not sda Is Nothing Then sda.Dispose()
            Call TIMS.CloseDbConn(conn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            trans.Rollback()
            Call TIMS.CloseDbConn(conn)
            Common.MessageBox(Page, ex.ToString)
        End Try
        Call TIMS.CloseDbConn(conn)
        If intCnt = 1 Then
            Common.MessageBox(Me, "儲存成功")
            'btnSch_Click(sender, e)
            Call Search1()
        End If
    End Sub

    '單一班級查詢1
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = Nothing
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Style("display") = "none"
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                DataGridTable.Style("display") = "none"
            End If
        End If
    End Sub

    '單一班級查詢2
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class