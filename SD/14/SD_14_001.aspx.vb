Partial Class SD_14_001
    Inherits AuthBasePage

    'SD_14_001_2010_G
    'SD_14_001_2010_W2
    'SD_14_001_2010_54

    'SD_14_001_18*.jrxml
    Const cst_printFN2g4 As String = "SD_14_001_18G" '2018-AppStage
    Const cst_printFN2w4 As String = "SD_14_001_18W" '2018-AppStage
    Const cst_printFN54 As String = "SD_14_001_2010_54" 'TPLANID:54

    'Old
    'Const cst_printFN2g1 As String = "SD_14_001"
    'Const cst_printFN2g2 As String = "SD_14_001_2009"
    'Const cst_printFN2g3 As String = "SD_14_001_2010_G"
    'Const cst_printFN2w1 As String = "SD_14_001_1"
    'Const cst_printFN2w2 As String = "SD_14_001_1_2009"
    'Const cst_printFN2w3 As String = "SD_14_001_2010_W2"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) 'TRPlanPoint28
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        tr_rbl_AppliedResult54.Visible = If(TIMS.Cst_TPlanID54AppPlan = sm.UserInfo.TPlanID, True, False)

        PageControler1.PageDataGrid = DataGrid1
        If Not IsPostBack Then Call Create1()
        PlanPoint.Attributes("onclick") = "return clearSelectValue();"
        'Button3.Attributes("onclick") = "return CheckPrint('" & ReportQuery.GetSmartQueryPath & "');"
    End Sub

    Sub Create1()
        PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
        Common.SetListItem(PlanPoint, "1")

        '依登入者機構判斷計畫種類
        Dim sql As String
        Dim dr As DataRow

        '依登入者 LID 判斷是否可自由輸入
        If sm.UserInfo.LID = 2 Then '委訓單位動作
            sql = "select OrgName,ComIDNO,OrgKind2 From Org_OrgInfo where OrgID = '" & sm.UserInfo.OrgID & "' "
            dr = DbAccess.GetOneRow(sql, objconn)
            OrgName.Text = dr("OrgName")
            ComIDNO.Text = dr("ComIDNO")
            If dr("OrgKind2") = "W" Then
                Common.SetListItem(PlanPoint, "2")
                'PlanPoint.SelectedValue = 2
            ElseIf dr("OrgKind2") = "G" Then
                Common.SetListItem(PlanPoint, "1")
                'PlanPoint.SelectedValue = 1
            End If
            OrgName.Enabled = False
            ComIDNO.Enabled = False
            PlanPoint.Enabled = False
        End If

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
            Call TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage)
        End If

        ROC_Years.Value = (sm.UserInfo.Years - 1911) '登入年度轉民國年份
        SelectValue.Value = "" '選擇清除工作
        DataGridTable.Visible = False
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim v_AppStage As String = "" 'TIMS.GetListValue(AppStage)
        If tr_AppStage_TP28.Visible Then v_AppStage = TIMS.GetListValue(AppStage)
        If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage

        '選擇清除工作
        SelectValue.Value = ""

        '依申請階段 重複資料過濾
        Dim sql_DISTINCT As String = If(v_AppStage <> "", "SELECT DISTINCT b.RSID", "SELECT b.RSID")
        Dim sql_APPSTAGE As String = If(v_AppStage <> "", ",pp.AppStage ", ",NULL AppStage")

        Dim sql As String = "" 'sql = "" & vbCrLf
        sql &= sql_DISTINCT & vbCrLf
        sql &= sql_APPSTAGE & vbCrLf
        sql &= " ,b.RID " & vbCrLf
        sql &= " ,d.DistID ,d.Name DistName" & vbCrLf
        sql &= " ,a.OrgName ,a.ComIDNO " & vbCrLf
        sql &= " ,c.PlanMaster ,c.PlanMasterPhone " & vbCrLf
        sql &= " ,a.OrgID " & vbCrLf
        sql &= " ,ip.PlanID " & vbCrLf
        sql &= " FROM ORG_ORGINFO a " & vbCrLf
        sql &= " JOIN AUTH_RELSHIP b ON a.OrgID = b.OrgID " & vbCrLf
        sql &= " JOIN ORG_ORGPLANINFO c ON b.RSID = c.RSID " & vbCrLf
        sql &= " JOIN ID_DISTRICT d ON b.DistID = d.DistID " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PlanID=b.PlanID" & vbCrLf
        '依申請階段 重複資料過濾
        If v_AppStage <> "" Then sql &= " JOIN PLAN_PLANINFO pp on a.ComIDNO = pp.ComIDNO  AND b.PlanID = pp.PlanID " & vbCrLf '依申請階段 
        sql &= " WHERE 1=1 " & vbCrLf

        hid_planid.Value = "" & sm.UserInfo.PlanID & ""
        If sm.UserInfo.RID = "A" Then
            hid_planid.Value = ""
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
            'sql &= " AND b.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' AND Years = '" & sm.UserInfo.Years & "') " & vbCrLf
        Else
            sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        OrgName.Text = TIMS.ClearSQM(OrgName.Text)
        ComIDNO.Text = TIMS.ClearSQM(ComIDNO.Text)
        If OrgName.Text <> "" Then sql &= " AND a.OrgName LIKE N'%" & OrgName.Text & "%' " & vbCrLf '有UNICODE字元問題
        If ComIDNO.Text <> "" Then sql &= " AND a.ComIDNO = '" & ComIDNO.Text & "'" & vbCrLf

        '28:產業人才投資方案
        KindValue.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                '產業人才投資計畫
                KindValue.Value = "G"
                sql &= " AND a.OrgKind <> 10 " & vbCrLf
            Else
                '提升勞工自主學習計畫
                KindValue.Value = "W"
                sql &= " AND a.OrgKind = 10 " & vbCrLf
            End If
        End If

        'Dim v_AppStage As String = TIMS.ClearSQM(AppStage.SelectedValue)
        If v_AppStage <> "" Then sql &= " AND pp.AppStage='" & v_AppStage & "' " & vbCrLf '依申請階段

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg.Text = ""

        Dim dr1 As DataRow = dt.Rows(0)
        hid_planid.Value = Convert.ToString(dr1("PlanID"))
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim RSID As HtmlInputCheckBox = e.Item.FindControl("RSID")
                Dim RID As HtmlInputHidden = e.Item.FindControl("RID")
                RID.Value = Convert.ToString(drv("RID"))
                RSID.Value = Convert.ToString(drv("RSID"))
                'RSID.Attributes("onclick") = "SelectItem(this.checked,this.value);"
                Dim strValueAry As String() = Split(SelectValue.Value, ",") '繼續勾選
                For Each sVal1 As String In strValueAry
                    If "'" & Convert.ToString(drv("RSID")) & "'" = sVal1 Then RSID.Checked = True
                Next
                'For i As Integer = 0 To ValueArray.Length - 1
                '    If drv("RSID").ToString = ValueArray(i) Then RSID.Checked = True
                'Next
        End Select
    End Sub

    Public Shared Function Get_DG_Value1(ByRef DG1 As DataGrid, ByVal HtmlInputCheckBoxFindCtrlID As String) As String
        Dim RST As String = ""
        If DG1 Is Nothing Then Return RST

        For Each eItem As DataGridItem In DG1.Items
            'Dim chkRSID As HtmlInputCheckBox = eItem.FindControl("RSID")
            'Dim hRID As HtmlInputHidden = Item.FindControl("RID")
            Dim chkRSID As HtmlInputCheckBox = eItem.FindControl(HtmlInputCheckBoxFindCtrlID)
            If chkRSID Is Nothing Then Return RST

            If chkRSID.Checked AndAlso chkRSID.Value <> "" Then
                RST &= String.Concat(If(RST <> "", ",", ""), chkRSID.Value)
            End If
        Next
        Return RST
    End Function

    Public Shared Function Get_DG_Value2(ByRef DG1 As DataGrid, ByVal HtmlInputCheckBoxFindCtrlID As String, ByVal HtmlInputHiddenFindCtrlID As String) As String
        Dim RST As String = ""
        If DG1 Is Nothing Then Return RST

        For Each Item As DataGridItem In DG1.Items
            'Dim chkRSID As HtmlInputCheckBox = Item.FindControl("RSID")
            'Dim hRID As HtmlInputHidden = Item.FindControl("RID")
            Dim chkRSID As HtmlInputCheckBox = Item.FindControl(HtmlInputCheckBoxFindCtrlID)
            Dim hRID As HtmlInputHidden = Item.FindControl(HtmlInputHiddenFindCtrlID)
            If chkRSID Is Nothing OrElse hRID Is Nothing Then Return RST

            If chkRSID.Checked AndAlso hRID.Value <> "" Then
                RST &= String.Concat(If(RST <> "", ",", ""), "\'", hRID.Value, "\'")
            End If
        Next
        Return RST
    End Function

    '列印
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim vsFileName1 As String = ""
        Dim Errmsg As String = ""

        SelectValue.Value = Get_DG_Value1(DataGrid1, "RSID")
        'SelectValue.Value = TIMS.ClearSQM(SelectValue.Value)
        If SelectValue.Value = "" Then Errmsg += "請選擇訓練機構 !" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Dim RIDvalue1 As String = Get_DG_Value2(DataGrid1, "RSID", "RID")
        If RIDvalue1 = "" Then Errmsg += "請選擇訓練機構 !" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Dim i_ROC_YearsV As Integer = Val(ROC_Years.Value)
        Select Case Convert.ToString(sm.UserInfo.TPlanID)
            Case "54"
                vsFileName1 = cst_printFN54 '"SD_14_001_2010_54"

            Case "28"
                '產業人才投資計畫/提升勞工自主學習計畫
                vsFileName1 = If(KindValue.Value = "G", cst_printFN2g4, cst_printFN2w4)
                'Select Case KindValue.Value
                '    Case "G" '產業人才投資計畫
                '        vsFileName1 = cst_printFN2g4 '2018使用
                '        If i_ROC_YearsV < 98 Then
                '            vsFileName1 = cst_printFN2g1 '"SD_14_001"
                '        ElseIf i_ROC_YearsV = 98 Then
                '            vsFileName1 = cst_printFN2g2 '"SD_14_001_2009"
                '        ElseIf i_ROC_YearsV < 107 Then '>99
                '            vsFileName1 = cst_printFN2g3 '"SD_14_001_2010_G"
                '        End If
                '    Case "W" '//提升勞工自主學習計畫
                '        vsFileName1 = cst_printFN2w4 '2018使用
                '        If i_ROC_YearsV < 98 Then
                '            vsFileName1 = cst_printFN2w1 '"SD_14_001_1"
                '        ElseIf i_ROC_YearsV = 98 Then
                '            vsFileName1 = cst_printFN2w2 '"SD_14_001_1_2009"
                '        ElseIf i_ROC_YearsV < 107 Then '>99
                '            vsFileName1 = cst_printFN2w3 '"SD_14_001_2010_W2"
                '        End If
                '    Case Else
                '        Errmsg += "請選擇計畫種類 !" & vbCrLf
                'End Select
        End Select

        If vsFileName1 = "" Then
            Errmsg += "請選擇計畫種類 !" & vbCrLf
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        'For Each Item As DataGridItem In DataGrid1.Items
        '    Dim chkRSID As HtmlInputCheckBox = Item.FindControl("RSID")
        '    Dim hRID As HtmlInputHidden = Item.FindControl("RID")
        '    If chkRSID.Checked AndAlso hRID.Value <> "" Then
        '        If RIDvalue1 <> "" Then RIDvalue1 &= ","
        '        RIDvalue1 &= "\'" & hRID.Value & "\'"
        '    End If
        'Next

        Dim prtstr As String = ""
        prtstr &= "&Years=" & CStr(i_ROC_YearsV)
        prtstr &= "&RSID=" & SelectValue.Value
        prtstr &= "&planid=" & TIMS.ClearSQM(hid_planid.Value)
        prtstr &= "&rid=" & RIDvalue1
        '依申請階段 
        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If v_AppStage <> "" AndAlso v_AppStage > "0" Then prtstr &= "&AppStage=" & v_AppStage

        Dim v_rbl_AppliedResult54 As String = TIMS.GetListValue(rbl_AppliedResult54)
        If tr_rbl_AppliedResult54.Visible AndAlso v_rbl_AppliedResult54 <> "" Then
            '<%-- A:不區分,Y:審核通過,M:審核中,N:不通過--%>
            '<%-- 01:不區分,02:審核通過,03:審核中,04:不通過--%>
            prtstr &= "&APPRST=" & v_rbl_AppliedResult54
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vsFileName1, prtstr)
    End Sub

End Class
