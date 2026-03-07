Partial Class SD_13_002_15
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    Const cst_是否取得結訓資格 As Integer = 3
    'Const cst_出席達2分之3 = 4
    Const cst_出席達3分之4 As Integer = 4
    Const cst_是否補助 As Integer = 5
    'Const cst_總費用 As Integer = 6
    'Const cst_補助費用 As Integer = 7
    'Const cst_個人支付 As Integer = 8
    Const cst_剩餘可用餘額 As Integer = 9
    Const cst_其他申請中金額 As Integer = 10
    'Const cst_撥款 As Integer = 11
    Const cst_撥款日期 As Integer = 12
    'Const cst_撥款備註 As Integer = 13
    'Const cst_預算別 As Integer = 14

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

        Select Case sm.UserInfo.Years.ToString()
            Case Is <= "2011"
                Call TIMS.CloseDbConn(objconn)
                Server.Transfer("SD_13_002_00.aspx?ID=" & Request("ID"))
                Exit Sub
        End Select

        If Not IsPostBack Then
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
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)  '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            btnSch.Attributes("onclick") = "return CheckSearch();"
            btnSave.Attributes("onclick") = "return chkSave();"
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
                            If Me.ViewState("sort") = "StudentID" Then mysort.ImageUrl = "../../images/SortUp.gif"
                        Case "IDNO", "IDNO DESC"
                            mylabel = "StudentID"
                            i = 2
                            mysort.ImageUrl = "../../images/SortDown.gif"
                            If Me.ViewState("sort") = "IDNO" Then mysort.ImageUrl = "../../images/SortUp.gif"
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                    DropDownList1.Enabled = True
                    If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then DropDownList1.Enabled = False
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
                'sqlstr += " ,g.BudId" & vbCrLf 'GET_BudgetName
                LabBudID.Text = TIMS.GET_BudgetName(Convert.ToString(drv("BudID")), objconn)
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                e.Item.Cells(cst_是否取得結訓資格).Text = "否"
                If Convert.ToString(drv("CreditPoints")) = "1" Then
                    e.Item.Cells(cst_是否取得結訓資格).Text = "是"
                    iFlag += 1
                End If
                e.Item.Cells(cst_出席達3分之4).Text = "否"
                If drv("THours") > 0 Then
                    If (drv("THours") - drv("CountHours")) / drv("THours") >= 3 / 4 Then
                        e.Item.Cells(cst_出席達3分之4).Text = "是"
                        iFlagStudy = 1
                    End If
                End If

                '可用補助額(2007年3年3萬)
                '可用補助額(2008年3年5萬)
                '可用補助額(2012年3年7萬)
                iTotal = TIMS.Get_3Y_SupplyMoney(Me)

                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)
                iTotal -= SubsidyCost

                If iTotal - Val(drv("SumOfMoney")) >= 0 Then
                    e.Item.Cells(cst_剩餘可用餘額).Text = iTotal
                Else
                    e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & iTotal & "</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_剩餘可用餘額), "剩餘可用餘額 不足補助金!")
                End If

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
                    If drv("AppliedStatus") = 1 Then AppliedStatus.SelectedIndex = 1 '已撥款
                End If
                AppliedNote.Text = Convert.ToString(drv("AppliedNote"))

                AppliedStatus.Enabled = True '撥款
                AppliedNote.Enabled = True '撥款備註
                If Me.AuditList.SelectedIndex = 1 OrElse Me.AuditList.SelectedIndex = 2 Then
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

        btnSch_Click(Me, e)
    End Sub

    '查詢 SQL 
    Sub Search1()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If

        Dim parms As New Hashtable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT d.SOCID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(d.StudentID) StudentID" & vbCrLf
        sql &= " ,e.Name" & vbCrLf
        sql &= " ,e.IDNO" & vbCrLf
        '除數可能有溢位問題，無條件捨去餘2位數。
        sql &= " ,CASE WHEN b.TotalCost>=ISNULL(c.Total2,0) THEN dbo.TRUNC(ISNULL(b.TotalCost,0)/ISNULL(b.TNum,1) ,0)" & vbCrLf
        sql &= " ELSE dbo.TRUNC(ISNULL(c.Total2,0)/ISNULL(b.TNum,1) ,0) END Total" & vbCrLf
        sql &= " ,d.CreditPoints" & vbCrLf
        sql &= " ,a.THours" & vbCrLf
        sql &= " ,ISNULL(f.CountHours,0) CountHours" & vbCrLf
        sql &= " ,e.DegreeID" & vbCrLf
        sql &= " ,d.StudStatus ,d.MIdentityID ,a.STDate" & vbCrLf
        sql &= " ,g.SOCID Exist" & vbCrLf
        sql &= " ,g.SumOfMoney" & vbCrLf
        sql &= " ,g.PayMoney ,g.AppliedStatus ,g.AllotDate ,g.AppliedNote" & vbCrLf
        sql &= " ,g.BudId" & vbCrLf
        sql &= " ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2" & vbCrLf
        sql &= " FROM Class_ClassInfo a" & vbCrLf
        sql &= " JOIN Plan_PlanInfo b ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= " 	SELECT PlanID ,ComIDNO ,SeqNo ,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)) Total ,Sum(ISNULL(OPrice,1)*ISNULL(Itemage,1)) Total2" & vbCrLf
        sql &= " 	FROM Plan_CostItem" & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql &= " WHERE COSTMODE = 5 " & vbCrLf
        Else
            sql &= " WHERE COSTMODE <> 5 " & vbCrLf
        End If
        sql &= " 	AND PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNo=@SeqNo" & vbCrLf
        sql &= " 	Group By PlanID,ComIDNO,SeqNo" & vbCrLf
        sql &= " ) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS d ON a.OCID=d.OCID" & vbCrLf
        sql &= " JOIN Stud_StudentInfo e ON d.SID=e.SID" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= " SELECT st2.SOCID,Sum(st2.Hours) CountHours" & vbCrLf
        sql &= " FROM STUD_TURNOUT2 st2" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on st2.socid =cs.socid" & vbCrLf
        sql &= " WHERE 1=1 AND cs.OCID=@OCID" & vbCrLf
        sql &= " Group By st2.SOCID" & vbCrLf
        sql &= " ) f ON d.SOCID=f.SOCID" & vbCrLf
        sql &= " JOIN Stud_SubsidyCost g ON d.SOCID=g.SOCID" & vbCrLf
        sql &= " WHERE a.OCID=@OCID" & vbCrLf

        sql &= " and g.Budid is not null "
        sql &= " and a.AppliedResultR IN ('Y','C') " 'Y 通過 'C 全班學員資料確認
        Select Case AuditList.SelectedIndex
            Case 0
                sql &= " AND a.AppliedResultM='Y' AND g.AppliedStatusM='Y' " & vbCrLf
            Case 1
                sql &= " AND a.AppliedResultM='N' " & vbCrLf
            Case 2
                sql &= " AND a.AppliedResultM IS Null " & vbCrLf
        End Select

        sql &= " AND g.AppliedStatusM='Y'" & vbCrLf
        'sql &= " order by dbo.FN_CSTUDID2(d.StudentID)" & vbCrLf
        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sql &= " order by e." & Me.ViewState("sort").ToString
        ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
            sql &= " order by dbo.FN_CSTUDID2(d.StudentID) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
        Else
            sql &= " order by dbo.FN_CSTUDID2(d.StudentID) " & vbCrLf
        End If

        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試
        'If flag_test Then
        '    Dim slogMsg1 As String = ""
        '    slogMsg1 &= "##SD_13_002_15.aspx, sqlstr: " & sqlstr & vbCrLf
        '    'slogMsg1 &= "##SYS_01_001, myParam: " & TIMS.GetMyValue3(myParam) & vbCrLf
        '    TIMS.writeLog(Me, slogMsg1)
        'End If

        'Dim sCmd As New SqlCommand(sqlstr, objconn)
        ''dt = DbAccess.GetDataTable(sqlstr, objconn)
        ''SELECT SOCID,Sum(Hours) CountHours FROM Stud_Turnout2 Group By SOCID
        ''SELECT *  FROM Stud_Turnout2 WHERE ROWNUM <=10
        'Call TIMS.OpenDbConn(objconn)
        'Dim dt As New DataTable
        'With sCmd
        '    .Parameters.Clear()
        '    dt.Load(.ExecuteReader())
        'End With

        'sql &= " 	AND PlanID=@PlanID and ComIDNO=@ComIDNO and SeqNo=@SeqNo" & vbCrLf
        parms.Clear()
        parms.Add("PlanID", drCC("PlanID"))
        parms.Add("ComIDNO", drCC("ComIDNO"))
        parms.Add("SeqNo", drCC("SeqNo"))
        parms.Add("OCID", OCIDValue1.Value)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

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
        Dim conn As SqlConnection
        conn = DbAccess.GetConnection()
        Dim trans As SqlTransaction = Nothing

        Try
            Dim sda As New SqlDataAdapter
            Dim dt As New DataTable
            Dim dr As DataRow = Nothing
            Call TIMS.OpenDbConn(conn)
            trans = DbAccess.BeginTrans(conn)
            sql = "" & vbCrLf
            sql &= " SELECT c.SOCID " & vbCrLf
            sql &= " FROM Stud_SubsidyCost c" & vbCrLf
            sql &= " JOIN Class_StudentsOfClass cs on cs.socid =c.socid and OCID= @OCID" & vbCrLf
            With sda
                .SelectCommand = New SqlCommand(sql, conn, trans)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                .Fill(dt)
            End With

            If dt.Rows.Count > 0 Then
                sql = ""
                sql = " UPDATE Stud_SubsidyCost "
                sql &= " set AppliedStatus=@AppliedStatus"
                sql &= " ,AllotDate= convert(datetime, @AllotDate, 111)"
                sql &= " ,AppliedNote=@AppliedNote "
                sql &= " ,ModifyAcct=@ModifyAcct,ModifyDate=getdate() "
                sql &= " where SOCID=@SOCID"
                sda.UpdateCommand = New SqlCommand(sql, conn, trans)
                Dim AppliedStatus As DropDownList = Nothing
                Dim AllotDate As TextBox = Nothing
                Dim AppliedNote As TextBox = Nothing
                For i As Integer = 0 To DataGrid1.Items.Count - 1
                    AppliedStatus = DataGrid1.Items(i).FindControl("AppliedStatus")
                    AllotDate = DataGrid1.Items(i).FindControl("txtAllotDate")
                    AppliedNote = DataGrid1.Items(i).FindControl("AppliedNote")
                    With sda
                        .UpdateCommand.Parameters.Clear()
                        .UpdateCommand.Parameters.Add("AppliedStatus", SqlDbType.VarChar).Value = IIf(AppliedStatus.SelectedIndex = 0, Convert.DBNull, "1")
                        '.UpdateCommand.Parameters.Add("AllotDate", SqlDbType.VarChar).Value = TIMS.cdate2(AllotDate.Text, "yyyy/MM/dd")
                        Dim myAllotDate As String = ""  'edit，by:20181001
                        If flag_ROC Then
                            myAllotDate = TIMS.cdate18(AllotDate.Text)  'edit，by:20181001
                        Else
                            myAllotDate = TIMS.cdate2(AllotDate.Text, "yyyy/MM/dd")  'edit，by:20181001
                        End If
                        .UpdateCommand.Parameters.Add("AllotDate", SqlDbType.VarChar).Value = myAllotDate  'edit，by:20181001
                        .UpdateCommand.Parameters.Add("AppliedNote", SqlDbType.VarChar).Value = IIf(AppliedNote.Text = "", Convert.DBNull, AppliedNote.Text)
                        .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .UpdateCommand.Parameters.Add("SOCID", SqlDbType.VarChar).Value = DataGrid1.DataKeys(i)
                        .UpdateCommand.ExecuteNonQuery()
                    End With
                Next
                intCnt = 1
            Else
                Common.MessageBox(Me, "此班學員未建立補助撥款，請重新建立")
            End If
            trans.Commit()
            If Not trans Is Nothing Then trans.Dispose()
            If Not sda Is Nothing Then sda.Dispose()
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
            btnSch_Click(sender, e)
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