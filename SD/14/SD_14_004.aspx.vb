Partial Class SD_14_004
    Inherits AuthBasePage

    'ReportQuery
    'SQControl.aspx
    'SD_14_004*.JRXML
    Const cst_printFN1 As String = "SD_14_004" '非產投 師資基本資料表
    Const cst_printFN2 As String = "SD_14_004B" '產投 師資基本資料表

    Dim iPYNum As Integer = 1  'iPYNum = TIMS.sUtl_GetPYNum(Me)  '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

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

        PageControler1.PageDataGrid = DataGrid1
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        Years.Value = sm.UserInfo.Years - 1911

        'Button3.Attributes("onclick") = "return CheckPrint();"
    End Sub

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim sql As String = ""
        Dim dt As DataTable
        Dim sTmp As String = ""

        Dim SearchStr As String = ""
        TechIDValue.Value = ""
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        SearchStr &= " AND a.RID = '" & RIDValue.Value & "' " & vbCrLf

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text <> "" Then SearchStr &= " AND a.IDNO = '" & IDNO.Text & "' " & vbCrLf

        TeachCName.Text = TIMS.ClearSQM(TeachCName.Text)
        If TeachCName.Text <> "" Then SearchStr &= " AND (a.TeachCName LIKE '%" & Replace(TeachCName.Text, " ", "%") & "%' OR a.TeachEName LIKE '%" & Replace(TeachCName.Text, " ", "%") & "%') " & vbCrLf

        TeacherID.Text = TIMS.ClearSQM(TeacherID.Text)
        If TeacherID.Text <> "" Then SearchStr &= " AND a.TeacherID = '" & TeacherID.Text & "' " & vbCrLf

        sTmp = TIMS.ClearSQM(KindEngage.SelectedValue)
        If KindEngage.SelectedIndex <> 0 AndAlso sTmp <> "" Then SearchStr &= " AND a.KindEngage = '" & sTmp & "' " & vbCrLf

        sTmp = TIMS.ClearSQM(WorkStatus.SelectedValue)
        If WorkStatus.SelectedIndex <> 0 AndAlso sTmp <> "" Then SearchStr &= " AND a.WorkStatus = '" & sTmp & "' " & vbCrLf

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                SearchStr &= " AND c.OrgKind <> 10 " & vbCrLf '產業人才投資計畫
            Else
                SearchStr &= " AND c.OrgKind = 10 " & vbCrLf '提升勞工自主學習計畫
            End If
        End If

        sql = "" & vbCrLf
        sql &= " SELECT a.RID" & vbCrLf
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " ,a.TeachCName" & vbCrLf
        sql &= " ,a.TeachEName" & vbCrLf
        sql &= " ,a.TeacherID" & vbCrLf
        'sql &= " /*內外聘*/" & vbCrLf
        sql &= " ,a.KindEngage" & vbCrLf
        sql &= " ,case a.KindEngage when '1' then '內聘' else '外聘' end KindEngage_N" & vbCrLf
        'sql &= " /*排課使用*/" & vbCrLf
        sql &= " ,a.WorkStatus" & vbCrLf
        sql &= " ,case a.WorkStatus when '1' then '是' else '否' end WorkStatus_N" & vbCrLf
        sql &= " ,a.TechID" & vbCrLf
        sql &= " FROM Teach_TeacherInfo a" & vbCrLf
        sql &= " JOIN Auth_Relship b ON a.RID = b.RID" & vbCrLf
        sql &= " JOIN Org_OrgInfo c ON b.OrgID = c.OrgID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= SearchStr
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sSearch1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim TechID As HtmlInputCheckBox = e.Item.FindControl("TechID")
                TechID.Value = Convert.ToString(drv("TechID"))
                TechID.Attributes("onclick") = "SelectItem(this.checked,this.value);"
                Dim TechIDArray As Array = Split(TechIDValue.Value, ",")
                For i As Integer = 0 To TechIDArray.Length - 1
                    If drv("TechID").ToString = TechIDArray(i) Then TechID.Checked = True
                Next
                'If drv("KindEngage").ToString = "1" Then
                '    e.Item.Cells(4).Text = "內聘"
                'Else
                '    e.Item.Cells(4).Text = "外聘"
                'End If
                'If drv("WorkStatus").ToString = "1" Then '2010/04/27
                '    e.Item.Cells(5).Text = "是"
                'Else
                '    e.Item.Cells(5).Text = "否"
                'End If
        End Select
    End Sub

    '列印
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        Dim TechIDStr As String = ""
        For i As Integer = 0 To DataGrid1.Items.Count - 1
            Dim TechID As HtmlInputCheckBox = CType(DataGrid1.Items(i).FindControl("TechID"), HtmlInputCheckBox)
            If TechID.Checked AndAlso TechID.Value <> "" Then
                '當被選取要做的事情 
                TechIDStr &= String.Concat(If(TechIDStr <> "", ",", ""), "\'", TechID.Value, "\'")
            End If
        Next

        '28:產業人才投資方案
        Dim sTitle As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim PNAME As String = PlanPoint.SelectedItem.Text
            Select Case PlanPoint.SelectedValue
                Case "1", "2"
                    sTitle += "（" & PNAME & "）"
            End Select
        End If
        If TechIDStr = "" Then
            Common.MessageBox(Me, "請選擇要列印的師資")
            Exit Sub
        End If

        Dim vFilenmae1 As String = cst_printFN1
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso iPYNum >= 3 Then vFilenmae1 = cst_printFN2
        Dim MyValue As String = ""
        MyValue = "kjk=kjk"
        MyValue += "&TechID=" & TechIDStr
        MyValue += "&Years=" & Years.Value
        MyValue += "&Title=" & sTitle
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vFilenmae1, MyValue)
    End Sub
End Class