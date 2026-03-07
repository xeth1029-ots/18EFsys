Partial Class SD_17_002
    Inherits AuthBasePage

    Const cst_學號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    Const cst_是否取得結訓資格 As Integer = 3
    'Const cst_出席達2分之3 = 4
    Const cst_出席達80 As Integer = 4
    Const cst_是否為在職者 As Integer = 5

    Const cst_是否補助 As Integer = 6
    Const cst_總費用 As Integer = 7
    Const cst_補助費用 As Integer = 8
    Const cst_個人支付 As Integer = 9
    Const cst_剩餘可用餘額 As Integer = 10
    Const cst_其他申請中金額 As Integer = 11

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
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
                    Select Case Convert.ToString(Me.ViewState("sort"))
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

                    If i <> -1 Then
                        e.Item.Cells(i).Controls.Add(mysort)
                    End If

                    If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then
                        DropDownList1.Enabled = False
                    Else
                        DropDownList1.Enabled = True
                    End If
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                'Dim Flag As Integer = 0      '是否補助
                'Dim FlagStudy As Integer = 0 '出勤flag,未滿2/3=0；達2/3=1
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedStatus As DropDownList = e.Item.FindControl("AppliedStatus")
                Dim AppliedNote As TextBox = e.Item.FindControl("AppliedNote")
                Dim txtAllotDate As TextBox = e.Item.FindControl("txtAllotDate")
                Dim ibtDate As ImageButton = e.Item.FindControl("ibtDate")
                Dim labTotal As Label = e.Item.FindControl("lab_Total") '總費用
                Dim hid_SOCID As HtmlInputHidden = e.Item.FindControl("hid_SOCID")
                Dim labEndClass As Label = e.Item.FindControl("lab_EndClass") '是否取得結訓資格
                Dim labOnClassRate As Label = e.Item.FindControl("lab_OnClassRate") '出席達80%
                Dim labWorkSuppIdent As Label = e.Item.FindControl("lab_WorkSuppIdent") '是否為在職者
                Dim labIsSubSidy As Label = e.Item.FindControl("lab_IsSubSidy") '是否補助

                Dim labOtherGovApply As Label = e.Item.FindControl("lab_OtherGovApply") '其他申請中的金額。

                Dim Total As Integer = 0  '可用補助額(2007年3年3萬)(2008年3年5萬)

                hid_SOCID.Value = drv("SOCID")
                '總費用
                labTotal.Text = drv("SumOfMoney") + drv("PayMoney")
                '是否結訓
                If CBool(drv("CreditPoints")) = True Then labEndClass.Text = "是" Else labEndClass.Text = "否"
                '出席達80%
                If drv("TOHours") > 0 Then
                    If drv("TOHours") > (drv("THours") * 2 / 10) Then labOnClassRate.Text = "否" Else labOnClassRate.Text = "是"
                Else
                    labOnClassRate.Text = "是"
                End If
                '是否為在職者
                Select Case Convert.ToString(drv("WorkSuppIdent"))
                    Case "Y"
                        labWorkSuppIdent.Text = "是"
                    Case Else
                        labWorkSuppIdent.Text = "否"
                End Select
                '是否補助
                If labOnClassRate.Text = "是" And labEndClass.Text = "是" Then labIsSubSidy.Text = "是" Else labIsSubSidy.Text = "<font color='red'>否</font>"

                '可用補助額
                Total = TIMS.Get_3Y_SupplyMoney()
                '含職前webservice
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

                'If Flag = 1 Then
                '    e.Item.Cells(cst_是否補助).Text = "是"
                'Else
                '    e.Item.Cells(cst_是否補助).Text = "否"
                'End If

                'If Flag = 0 Or FlagStudy = 0 Then
                '    e.Item.Cells(cst_是否補助).Text = "<font color='RED'>否</font>"
                '    TIMS.Tooltip(e.Item.Cells(cst_是否補助), " 學分為 0 或 出勤未滿2/3 時不補助!")
                'End If

                If IsDBNull(drv("AppliedStatus")) Then
                    AppliedStatus.SelectedIndex = 0 '請選擇
                Else
                    If drv("AppliedStatus") = 1 Then
                        AppliedStatus.SelectedIndex = 1 '已撥款
                    End If
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

                If Convert.ToString(drv("AllotDate")) <> "" Then
                    txtAllotDate.Text = Convert.ToString(drv("AllotDate"))
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

        btnSch_Click(Me, e)
    End Sub

    '查詢鈕
    Private Sub btnSch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSch.Click

        Dim dt As DataTable
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " select a.SOCID"
        sqlstr += " , dbo.SUBSTR(a.StudentID,-2) StudentID" & vbCrLf
        sqlstr += " ,b.IDNO,b.Name,d.TotalCost" & vbCrLf
        sqlstr += " ,f.SumOfMoney,f.PayMoney ,a.WorkSuppIdent" & vbCrLf
        sqlstr += " ,f.AppliedStatusM, a.ActNO,f.BudID" & vbCrLf
        sqlstr += " ,f.AppliedStatus " & vbCrLf
        sqlstr += " ,CONVERT(varchar, f.AllotDate, 111) AllotDate" & vbCrLf
        sqlstr += " ,f.AppliedNote " & vbCrLf
        sqlstr += " ,dbo.NVL(a.CreditPoints,0) CreditPoints" & vbCrLf
        sqlstr += " ,d.THours" & vbCrLf
        sqlstr += " ,dbo.NVL(tt.TOHours,0) TOHours" & vbCrLf
        sqlstr += " ,CONVERT(varchar, c.STDate, 111)  STDate" & vbCrLf
        sqlstr += " ,dbo.FN_GET_GOVAPPL2(b.IDNO,c.STDate) GovAppl2" & vbCrLf
        sqlstr += " from Class_StudentsOfClass a" & vbCrLf
        sqlstr += " join Stud_StudentInfo b on b.SID=a.SID" & vbCrLf
        sqlstr += " join Class_ClassInfo c on c.OCID=a.OCID" & vbCrLf
        sqlstr += " join Plan_PlanInfo d on d.ComIDNO=c.ComIDNO and d.PlanID=c.PlanID and d.SeqNO=c.SeqNO" & vbCrLf
        sqlstr += " join id_plan ip on ip.planid =c.planid" & vbCrLf
        sqlstr += " join Stud_SubSidyCost f on f.SOCID=a.SOCID" & vbCrLf '經費審核
        sqlstr += " left join (" & vbCrLf
        sqlstr += "  select m.socid ,sum(dbo.NVL(m.Hours,0)) TOHours" & vbCrLf
        sqlstr += "  from Stud_Turnout m" & vbCrLf
        sqlstr += "  group by m.socid" & vbCrLf
        sqlstr += " ) tt on tt.socid =a.socid" & vbCrLf
        sqlstr += " where 1=1" & vbCrLf
        sqlstr += " and c.OCID='" & OCIDValue1.Value & "' "
        '經費審核。'AppliedStatusM
        Select Case AuditList.SelectedValue
            Case "Y"
                sqlstr += " AND f.AppliedStatusM='Y' " & vbCrLf
            Case "N"
                sqlstr += " AND f.AppliedStatusM!='Y' " & vbCrLf
            Case Else 'Null '審核中
                sqlstr += " AND dbo.NVL(f.AppliedStatusM , ' ') NOT IN ('Y','N') " & vbCrLf
        End Select
        Dim ORDER_BY As String
        ORDER_BY = UCase(Me.ViewState("sort"))
        If ORDER_BY <> "" Then
            Select Case UCase(ORDER_BY)
                Case UCase("StudentID")
                    sqlstr += " order by a.StudentID "
                Case UCase("StudentID DESC")
                    sqlstr += " order by a.StudentID desc"
                Case UCase("IDNO")
                    sqlstr += " order by b.idno "
                Case UCase("IDNO DESC")
                    sqlstr += " order by b.idno desc"
            End Select
        Else
            sqlstr += " order by dbo.SUBSTR(a.StudentID,-2) "
        End If
        'sqlstr += " and ip.tplanid ='54'" & vbCrLf
        'sqlstr += " and ip.distid ='001'" & vbCrLf
        'sqlstr += " and ip.years='2013'" & vbCrLf
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        DataGridTable.Style("display") = "none"
        msg.Text = "查無資料"
        btnSave.Enabled = False
        If dt.Rows.Count > 0 Then
            DataGridTable.Style("display") = "inline"
            msg.Text = ""
            btnSave.Enabled = True '可儲存。

            If ViewState("sort") = "" Then ViewState("sort") = "StudentID"

            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()

            'Select Case AuditList.SelectedValue
            '    Case "Y", "N" '已通過'不通過
            '        btnSave.Enabled = False
            '    Case Else
            '        btnSave.Enabled = True
            'End Select
            'If Me.AuditList.SelectedIndex = 1 Or Me.AuditList.SelectedIndex = 2 Then
            '    btnSave.Enabled = False
            'Else
            '    btnSave.Enabled = True
            'End If
        End If
    End Sub

    '儲存鈕
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim intCnt As Integer = 0

        Dim conn As SqlConnection
        conn = DbAccess.GetConnection()
        Dim trans As SqlTransaction = Nothing
        Try
            Dim sda As New SqlDataAdapter
            Dim sql As String = ""
            Dim dt As New DataTable
            Dim dr As DataRow = Nothing

            Call TIMS.OpenDbConn(conn)
            trans = DbAccess.BeginTrans(conn)

            sql = "" & vbCrLf
            sql += " SELECT c.SOCID " & vbCrLf
            sql += " FROM Stud_SubsidyCost c" & vbCrLf
            sql += " JOIN Class_StudentsOfClass cs on cs.socid =c.socid and OCID= @OCID" & vbCrLf

            With sda
                .SelectCommand = New SqlCommand(sql, conn, trans)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                .Fill(dt)
            End With

            If dt.Rows.Count > 0 Then
                sql = ""
                sql &= " update Stud_SubsidyCost "
                sql += " set AppliedStatus=@AppliedStatus"
                sql += " ,AllotDate= convert(datetime, @AllotDate, 111)"
                sql += " ,AppliedNote=@AppliedNote "
                sql += " ,ModifyAcct=@ModifyAcct"
                sql += " ,ModifyDate=getdate() "
                sql += " where SOCID=@SOCID"

                sda.UpdateCommand = New SqlCommand(sql, conn, trans)
                Dim AppliedStatus As DropDownList = Nothing
                Dim AllotDate As TextBox = Nothing
                Dim AppliedNote As TextBox = Nothing
                Dim hid_SOCID As HtmlInputHidden = Nothing

                For Each eitem As DataGridItem In DataGrid1.Items
                    AppliedStatus = eitem.FindControl("AppliedStatus")
                    AllotDate = eitem.FindControl("txtAllotDate")
                    AppliedNote = eitem.FindControl("AppliedNote")
                    hid_SOCID = eitem.FindControl("hid_SOCID")

                    With sda
                        .UpdateCommand.Parameters.Clear()
                        .UpdateCommand.Parameters.Add("AppliedStatus", SqlDbType.VarChar).Value = IIf(AppliedStatus.SelectedIndex = 0, Convert.DBNull, "1")
                        .UpdateCommand.Parameters.Add("AllotDate", SqlDbType.VarChar).Value = TIMS.Cdate2(AllotDate.Text, "yyyy/MM/dd")
                        .UpdateCommand.Parameters.Add("AppliedNote", SqlDbType.VarChar).Value = IIf(AppliedNote.Text = "", Convert.DBNull, AppliedNote.Text)
                        .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .UpdateCommand.Parameters.Add("SOCID", SqlDbType.VarChar).Value = hid_SOCID.Value
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
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class