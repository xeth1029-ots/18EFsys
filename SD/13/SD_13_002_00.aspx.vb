Partial Class SD_13_002_00
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    Const cst_是否取得結訓資格 As Integer = 3
    Const cst_出席達3分之2 As Integer = 4
    Const cst_是否補助 As Integer = 5
    'Const cst_總費用 As Integer = 6
    'Const cst_補助費用 As Integer = 7
    'Const cst_個人支付 As Integer = 8
    Const cst_剩餘可用餘額 As Integer = 9
    Const cst_其他申請中金額 As Integer = 10
    Const cst_撥款日期 As Integer = 12

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
            '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
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

        btnSch.Attributes("onclick") = "return CheckSearch();"
        btnSave.Attributes("onclick") = "return chkSave();"
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
                            If Me.ViewState("sort") = "StudentID" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "IDNO", "IDNO DESC"
                            mylabel = "StudentID"
                            i = 2
                            If Me.ViewState("sort") = "IDNO" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                    If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then
                        DropDownList1.Enabled = False
                    Else
                        DropDownList1.Enabled = True
                    End If
                End If
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Flag As Integer = 0      '是否補助
                Dim FlagStudy As Integer = 0 '出勤flag,未滿2/3=0；達2/3=1
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedStatus As DropDownList = e.Item.FindControl("AppliedStatus")
                Dim AppliedNote As TextBox = e.Item.FindControl("AppliedNote")
                Dim txtAllotDate As TextBox = e.Item.FindControl("txtAllotDate")
                Dim ibtDate As ImageButton = e.Item.FindControl("ibtDate")
                Dim Total As Integer = 0  '可用補助額(2007年3年3萬)(2008年3年5萬)
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                If IsDBNull(drv("CreditPoints")) Then
                    e.Item.Cells(cst_是否取得結訓資格).Text = "否"
                Else
                    If drv("CreditPoints") Then
                        e.Item.Cells(cst_是否取得結訓資格).Text = "是"
                        Flag += 1
                    Else
                        e.Item.Cells(cst_是否取得結訓資格).Text = "否"
                    End If
                End If
                e.Item.Cells(cst_出席達3分之2).Text = "否"
                If drv("THours") > 0 Then
                    If (drv("THours") - drv("CountHours")) / drv("THours") >= 2 / 3 Then
                        e.Item.Cells(cst_出席達3分之2).Text = "是"
                        FlagStudy = 1
                    End If
                End If
                'If sm.UserInfo.Years < 2008 Then
                '    ''2007年前(含2007)
                '    Total = 30000
                'Else
                '    ''2008年後(含2008)
                '    Total = 50000
                'End If
                '可用補助額(2007年3年3萬)
                '可用補助額(2008年3年5萬)
                '可用補助額(2012年3年7萬)
                Total = TIMS.Get_3Y_SupplyMoney(Me)
                Dim SubsidyCost As Double = TIMS.Get_SubsidyCost(drv("IDNO").ToString(), drv("STDate").ToString(), "", "Y", objconn)
                Total -= SubsidyCost
                If Total - drv("SumOfMoney") >= 0 Then
                    e.Item.Cells(cst_剩餘可用餘額).Text = Total
                Else
                    e.Item.Cells(cst_剩餘可用餘額).Text = "<font color=Red>" & Total & "</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_剩餘可用餘額), "剩餘可用餘額 不足補助金!")
                End If
                If Total - drv("SumOfMoney") >= 0 AndAlso drv("GovAppl2") > Total - drv("SumOfMoney") Then
                    e.Item.Cells(cst_其他申請中金額).Text = "<font color=Red>" & drv("GovAppl2").ToString & "</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_其他申請中金額), "其他申請中金額 不足補助金!")
                End If
                If Flag = 1 Then
                    e.Item.Cells(cst_是否補助).Text = "是"
                Else
                    e.Item.Cells(cst_是否補助).Text = "否"
                End If
                If Flag = 0 Or FlagStudy = 0 Then
                    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_是否補助), " 學分為 0 或 出勤未滿2/3 時不補助!")
                End If
                If IsDBNull(drv("AppliedStatus")) Then
                    AppliedStatus.SelectedIndex = 0 '請選擇
                Else
                    If drv("AppliedStatus") = 1 Then AppliedStatus.SelectedIndex = 1 '已撥款
                End If
                AppliedNote.Text = drv("AppliedNote").ToString
                If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then
                    AppliedStatus.Enabled = False
                    AppliedNote.Enabled = False
                Else
                    AppliedStatus.Enabled = True
                    AppliedNote.Enabled = True
                End If
                If drv("IDNO").ToString <> "" Then
                    e.Item.Cells(cst_姓名).ToolTip = TIMS.Search_Stud_SubsidyCost(drv("IDNO").ToString, objconn)
                    e.Item.Cells(cst_學號).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                    e.Item.Cells(cst_身分證號碼).ToolTip = e.Item.Cells(cst_姓名).ToolTip
                End If
                If Convert.ToString(drv("AllotDate")) <> "" Then txtAllotDate.Text = Convert.ToDateTime(Convert.ToString(drv("AllotDate"))).ToString("yyyy/MM/dd")
                ibtDate.Attributes.Add("onclick", "selDate(" & e.Item.ItemIndex + 1 & ");return false;")
                If flag_ROC Then txtAllotDate.Text = TIMS.cdate17(drv("AllotDate"))  'edit，by:20181001
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

    '查詢鈕
    Private Sub btnSch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSch.Click
        Dim dt As DataTable

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " " & vbCrLf
        sqlstr &= " SELECT d.SOCID" & vbCrLf
        sqlstr &= " ,d.OCID" & vbCrLf
        sqlstr &= " ,dbo.FN_CSTUDID2(d.StudentID) StudentID" & vbCrLf
        sqlstr &= " ,e.Name" & vbCrLf
        sqlstr &= " ,e.IDNO" & vbCrLf
        '除數可能有溢位問題，無條件捨去餘2位數。
        sqlstr &= " ,case when b.TotalCost>=ISNULL(c.Total2,0) then ROUND(ISNULL(b.TotalCost,0)/ISNULL(b.TNum,1),2) else ROUND(ISNULL(c.Total2,0)/ISNULL(b.TNum,1),2) end Total" & vbCrLf
        sqlstr &= " ,d.CreditPoints" & vbCrLf
        sqlstr &= " ,a.THours" & vbCrLf
        sqlstr &= " ,ISNULL(f.CountHours,0) CountHours" & vbCrLf
        sqlstr &= " ,e.DegreeID" & vbCrLf
        sqlstr &= " ,d.StudStatus" & vbCrLf
        sqlstr &= " ,d.MIdentityID" & vbCrLf
        sqlstr &= " ,a.STDate" & vbCrLf
        sqlstr &= " ,g.SOCID Exist" & vbCrLf
        sqlstr &= " ,g.SumOfMoney" & vbCrLf
        sqlstr &= " ,g.PayMoney ,g.AppliedStatus" & vbCrLf
        sqlstr &= " ,g.AllotDate ,g.AppliedNote" & vbCrLf
        sqlstr &= " ,g.BudId" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_GOVAPPL2(e.IDNO,a.STDate) GovAppl2" & vbCrLf
        sqlstr &= " FROM Class_ClassInfo a" & vbCrLf
        sqlstr &= " JOIN Plan_PlanInfo b ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo" & vbCrLf
        sqlstr &= " LEFT JOIN (" & vbCrLf

        sqlstr &= " SELECT PlanID ,ComIDNO ,SeqNo" & vbCrLf
        sqlstr &= " ,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)) Total" & vbCrLf
        sqlstr &= " ,Sum(ISNULL(OPrice,1)*ISNULL(Itemage,1)) Total2" & vbCrLf
        sqlstr &= " FROM Plan_CostItem" & vbCrLf
        sqlstr += If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, " WHERE COSTMODE=5 ", " WHERE COSTMODE<>5 ") & vbCrLf
        sqlstr &= " GROUP BY PlanID,ComIDNO,SeqNo" & vbCrLf

        sqlstr &= " ) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo" & vbCrLf
        sqlstr &= " join ID_PLAN ip on ip.PLANID=a.PLANID" & vbCrLf
        sqlstr &= " JOIN Class_StudentsOfClass d ON a.OCID=d.OCID" & vbCrLf
        sqlstr &= " JOIN Stud_StudentInfo e ON d.SID=e.SID" & vbCrLf

        sqlstr &= " LEFT JOIN (" & vbCrLf
        sqlstr &= " SELECT SOCID,Sum(Hours) as CountHours" & vbCrLf
        sqlstr &= " FROM Stud_Turnout2" & vbCrLf
        sqlstr &= " Group By SOCID) f ON d.SOCID=f.SOCID" & vbCrLf

        sqlstr &= " JOIN Stud_SubsidyCost g ON d.SOCID=g.SOCID" & vbCrLf
        sqlstr &= " WHERE 1=1" & vbCrLf
        sqlstr &= " and g.Budid is not null" & vbCrLf
        sqlstr &= " and a.AppliedResultR='Y'" & vbCrLf
        Select Case AuditList.SelectedIndex
            Case 0
                sqlstr += " AND a.AppliedResultM='Y' AND g.AppliedStatusM='Y' " & vbCrLf
            Case 1
                sqlstr += " AND a.AppliedResultM='N' " & vbCrLf
            Case 2
                sqlstr += " AND a.AppliedResultM IS Null " & vbCrLf
        End Select
        sqlstr &= " AND ip.TPLANID ='28'" & vbCrLf
        'sqlstr &= " AND ip.YEARS='2018'" & vbCrLf
        'sqlstr &= " AND a.OCID =114350" & vbCrLf
        sqlstr += " and a.OCID='" & OCIDValue1.Value & "' " & vbCrLf
        If InStr(Me.ViewState("sort"), "IDNO") > 0 Then
            sqlstr &= " order by e." & Me.ViewState("sort").ToString
        ElseIf InStr(Me.ViewState("sort"), "StudentID") > 0 Then
            sqlstr &= " order by dbo.FN_CSTUDID2(d.StudentID) " & Replace(Me.ViewState("sort").ToString, "StudentID", "") & vbCrLf
        Else
            sqlstr &= " order by dbo.FN_CSTUDID2(d.StudentID) " & vbCrLf
        End If
        dt = DbAccess.GetDataTable(sqlstr, objconn)

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
            If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then
                btnSave.Enabled = False
            End If
        End If
    End Sub

    '儲存鈕
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim conn As SqlConnection = DbAccess.GetConnection
        Dim trans As SqlTransaction = Nothing
        Dim sda As New SqlDataAdapter
        Dim sql As String = ""

        Dim dt As New DataTable
        Dim dr As DataRow = Nothing
        Dim intCnt As Integer = 0

        Try
            'conn.Open()
            Call TIMS.OpenDbConn(conn)
            trans = DbAccess.BeginTrans(conn)
            sql = " SELECT SOCID FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID= @OCID) "
            With sda
                .SelectCommand = New SqlCommand(sql, conn, trans)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                .Fill(dt)
            End With
            If dt.Rows.Count > 0 Then
                sql = ""
                sql &= " update Stud_SubsidyCost "
                sql &= " set AppliedStatus= @AppliedStatus "
                sql &= "  ,AllotDate= @AllotDate "
                sql &= "  ,AppliedNote= @AppliedNote "
                sql &= "  ,ModifyAcct= @ModifyAcct "
                sql &= "  ,ModifyDate=getdate() "
                sql += " where SOCID= @SOCID "
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
                        '.UpdateCommand.Parameters.Add("AllotDate", SqlDbType.VarChar).Value = IIf(AllotDate.Text = "", Convert.DBNull, AllotDate.Text)
                        Dim myAllotDate As String = ""  'edit，by:20181001
                        If flag_ROC Then
                            myAllotDate = TIMS.cdate18(AllotDate.Text)  'edit，by:20181001
                        Else
                            myAllotDate = TIMS.cdate2(AllotDate.Text, "yyyy/MM/dd")  'edit，by:20181001
                        End If
                        .UpdateCommand.Parameters.Add("AllotDate", SqlDbType.VarChar).Value = IIf(myAllotDate = "", Convert.DBNull, myAllotDate)  'edit，by:20181001
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
        '判斷機構是否只有一個班級
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
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