Partial Class SD_14_022
    Inherits AuthBasePage

    '//已轉班 (vp.Years-1911)
    Const cst_printFN1 As String = "SD_14_022"

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
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        hidYears.Value = sm.UserInfo.Years - 1911 '設定登入民國年

        If Not IsPostBack Then
            msg.Text = "" '每次 清空
            DataGridTable.Visible = False '預設 隱藏
            'ClassTR.Visible = False '預設 隱藏
            hidOCIDValue.Value = ""
            'hidPCSValue.Value = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'Common.SetListItem(Radio1, "0") '預設 未轉班
            'Me.Radio1.SelectedIndex = 0
            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2

            '列印
            btnPrint1.Attributes("onclick") = "return CheckPrint();"
            Button4.Attributes("onclick") = "ClearData();"
        End If

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
        'Select Case Radio1.SelectedValue
        '    Case "1" '已轉班
        'End Select
    End Sub

#Region "(No Use)"

    'Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
    '    hidOCIDValue.Value = ""
    '    hidPCSValue.Value = ""
    '    DataGridTable.Visible = False

    '    ClassTR.Visible = False
    '    Select Case Radio1.SelectedValue
    '        Case "0" '未轉班
    '            ClassTR.Visible = False
    '        Case "1" '已轉班 顯示班別 
    '            ClassTR.Visible = True
    '    End Select
    'End Sub

#End Region

    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = ""
        sql += " SELECT cc.OCID ,CONVERT(VARCHAR,pp.PlanID) + '_' + CONVERT(VARCHAR, pp.ComIDNO) + '_' + CONVERT(VARCHAR,pp.SeqNo) PCSValue" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) ClassCName" & vbCrLf
        sql += " ,CONVERT(VARCHAR, cc.STDate, 111) STDate ,CONVERT(VARCHAR, cc.FTDate, 111) FTDate ,v1.OrgName " & vbCrLf
        sql += " FROM PLAN_PLANINFO pp " & vbCrLf
        sql += " JOIN CLASS_CLASSINFO cc ON cc.PlanID = pp.PlanID AND cc.comidno = pp.comidno AND cc.seqno = pp.seqno " & vbCrLf
        sql += " JOIN ID_PLAN ip ON ip.PlanID = pp.PlanID " & vbCrLf
        sql += " JOIN AUTH_RELSHIP ar ON ar.RID = cc.RID " & vbCrLf
        sql += " JOIN VIEW_RIDNAME v1 ON pp.RID = v1.RID " & vbCrLf
        sql += " WHERE pp.IsApprPaper = 'Y' " & vbCrLf  '限制為只有正式儲存之班級
        sql += "    AND pp.TransFlag = 'Y' " & vbCrLf  '已轉班
        If sm.UserInfo.RID = "A" Then
            sql += " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql += " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            sql += " AND pp.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        If RelShip <> "" Then sql += " AND ar.RelShip LIKE '" & RelShip & "%' " & vbCrLf
        If OCIDValue1.Value <> "" Then sql += " AND cc.OCID = '" & OCIDValue1.Value & "' " & vbCrLf

        '28:產業人才投資方案
        hidorgid.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sql += " AND v1.OrgKind <> '10' " & vbCrLf
                    hidorgid.Value = "G"
                Case Else
                    '提升勞工自主學習計畫
                    sql += " AND v1.OrgKind = '10' " & vbCrLf
                    hidorgid.Value = "W"
            End Select
        End If

        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            PageControler1.Visible = True
            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call Search1() '1:已轉班
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim drv As DataRowView = e.Item.DataItem
                Dim checkbox1 As HtmlInputCheckBox = e.Item.FindControl("checkbox1")
                Dim OCID As HiddenField = e.Item.FindControl("OCID")
                'Dim PCSValue As HiddenField = e.Item.FindControl("PCSValue")
                OCID.Value = Convert.ToString(drv("OCID"))
                'PCSValue.Value = Convert.ToString(drv("PCSValue"))
                If hidOCIDValue.Value.IndexOf(CStr(drv("OCID"))) > -1 Then checkbox1.Checked = True
                'If hidPCSValue.Value.IndexOf(CStr(drv("PCSValue"))) > -1 Then checkbox1.Checked = True
        End Select
    End Sub

    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""

#Region "(No Use)"

        'Dim vsFileName1 As String = ""
        'Select Case hidorgid.Value
        '    Case "G"
        '        vsFileName1 = cst_print_G1
        '    Case "W"
        '        vsFileName1 = cst_print_W1
        'End Select
        'If vsFileName1 = "" Then
        '    Errmsg += "請選擇計畫種類 !" & vbCrLf
        'End If
        'Dim vsFileName1 As String = ""
        'vsFileName1 = cst_print_1

#End Region

        hidOCIDValue.Value = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            Dim OCID As HiddenField = eItem.FindControl("OCID")
            If checkbox1.Checked And OCID.Value <> "" Then
                If hidOCIDValue.Value <> "" Then hidOCIDValue.Value &= ","
                hidOCIDValue.Value &= "\'" & OCID.Value & "\'"
            End If
        Next
        If Trim(hidOCIDValue.Value) = "" Then Errmsg += "請選擇 職類/班別 !" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Dim prtstr As String = ""
        prtstr = ""
        'prtstr += "&Years=" & hidYears.Value
        prtstr += "&TPlanID=" & sm.UserInfo.TPlanID
        prtstr += "&OCID=" & hidOCIDValue.Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, prtstr)
    End Sub
End Class