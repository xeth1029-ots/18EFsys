Partial Class SD_14_003
    Inherits AuthBasePage

    '2011報表
    'SQControl.aspx
    '//未轉班
    'W BussinessTrain: SD_14_003_5_2 '提升勞工自主學習計畫
    'G BussinessTrain: SD_14_003_1_2009 '產業人才投資計畫
    '//已轉班
    'W BussinessTrain: SD_14_003_4_2009 '提升勞工自主學習計畫
    'G BussinessTrain: SD_14_003_2009 '產業人才投資計畫 
    '(vp.Years-1911)

#Region "(No Use)"

    '2013報表(未設計) 'SQControl.aspx
    '//未轉班
    'W BussinessTrain: SD_14_003_W1_2013 '提升勞工自主學習計畫
    'G BussinessTrain: SD_14_003_G1_2013 '產業人才投資計畫
    '//已轉班
    'W BussinessTrain: SD_14_003_W2_2013 '提升勞工自主學習計畫
    'G BussinessTrain: SD_14_003_G2_2013 '產業人才投資計畫

#End Region

    'OLD
    'Const cst_print_2009W As String = "SD_14_003_4_2009"
    'Const cst_print_2009G As String = "SD_14_003_2009"
    'Const cst_print_2018W As String = "SD_14_003_2018W"
    'Const cst_print_2018G As String = "SD_14_003_2018G"
    'Const cst_print_09W As String = "SD_14_003_5_2"
    'Const cst_print_09G As String = "SD_14_003_1_2009"
    'Const cst_print_18W As String = "SD_14_003_18W"
    'Const cst_print_18G As String = "SD_14_003_18G"

    '(NEW)
    'SD_14_003_2019*.jrxml
    'SD_14_003_19*.jrxml
    Const cst_print_2019W As String = "SD_14_003_2019W" 'CC '已轉班
    Const cst_print_2019G As String = "SD_14_003_2019G" 'CC '已轉班
    Const cst_print_19W As String = "SD_14_003_19W" 'PP '未轉班
    Const cst_print_19G As String = "SD_14_003_19G" 'PP '未轉班

    '(NEW)
    '2019-'政策性產業課程可辦理班數-PLAN_PRECLASS
    'Const cst_print_2019W2 As String = "SD_14_003_2019W2" 'CC
    'Const cst_print_2019G2 As String = "SD_14_003_2019G2" 'CC
    'Const cst_print_19W2 As String = "SD_14_003_19W2" 'PP
    'Const cst_print_19G2 As String = "SD_14_003_19G2" 'PP

    '政策性產業課程可辦理班數-PLAN_PRECLASS
    'Dim flag_SHOW_2019_3 As Boolean = False

    'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
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
        'iPYNum = TIMS.sUtl_GetPYNum(Me)
        '政策性產業課程可辦理班數-PLAN_PRECLASS
        'flag_SHOW_2019_3 = TIMS.SHOW_2019_3() 'work2019x03

        If Not IsPostBack Then
            msg.Text = "" '每次 清空
            ROC_Years.Value = sm.UserInfo.Years - 1911 '設定登入民國年
            DataGridTable.Visible = False '預設 隱藏
            ClassTR.Visible = False '預設 隱藏
            Hid_OCIDValue.Value = ""
            Hid_PCSVALUE.Value = ""
            'SeqNoValue.Value = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Common.SetListItem(Radio1, "0") '預設 未轉班
            'Me.Radio1.SelectedIndex = 0
            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")

            '依申請階段'表示 (1：上半年、2：下半年、3：政策性產業)
            If tr_AppStage_TP28.Visible Then
                AppStage2 = TIMS.Get_APPSTAGE2(AppStage2)
                TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage2) 'Common.SetListItem(AppStage2, "3")
            End If
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '班級狀態 0:未轉班 /1:已轉班
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Select Case v_Radio1'Radio1.SelectedValue
            Case "1" '已轉班
                TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
                If HistoryTable.Rows.Count <> 0 Then
                    OCID1.Attributes("onclick") = "showObj('HistoryList');"
                    OCID1.Style("CURSOR") = "hand"
                End If
        End Select

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        '列印
        BtnPrint3.Attributes("onclick") = "return CheckPrint();"
        Button4.Attributes("onclick") = "ClearData();"

        '隱藏 全部列印
        'AllPrint.Visible = False
        'AllPrint.Attributes("onclick") = "SelectAll3(this.checked);"
    End Sub

    Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        Hid_OCIDValue.Value = ""
        Hid_PCSVALUE.Value = ""
        ClassTR.Visible = False
        DataGridTable.Visible = False
        '班級狀態 0:未轉班 /1:已轉班
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Select Case v_Radio1'Radio1.SelectedValue
            Case "0" '未轉班
                ClassTR.Visible = False
            Case "1" '已轉班 顯示班別 
                ClassTR.Visible = True
        End Select
    End Sub

    '已轉班
    Private Sub CreateClass(ByVal sTypePC As String)
        Select Case sTypePC
            Case "C" 'CLASS
            Case "P" 'PLAN
        End Select

        PageControler1.PageDataGrid = DataGrid1
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim drN As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If drN Is Nothing Then Exit Sub
        Dim vRelShip As String = Convert.ToString(drN("RelShip"))
        Hid_PlanID1.Value = Convert.ToString(drN("PlanID"))

        Dim v_AppStage2 As String = "" 'TIMS.GetListValue(AppStage2)
        If tr_AppStage_TP28.Visible Then v_AppStage2 = TIMS.GetListValue(AppStage2)
        If (v_AppStage2 <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage2

        'IsApprPaper='Y':限制為只有正式儲存之班級
        'TransFlag ='Y':已轉班 'TransFlag ='N':未轉班,
        Dim SQL_1 As String = "
SELECT pp.PlanID , pp.ComIDNO ,pp.SeqNo,CONCAT(pp.PlanID,'x',pp.ComIDNO,'x',pp.SeqNo) PCSVALUE,cc.OCID 
,pp.AppStage,pp.RESULTBUTTON
,CONCAT(dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE),dbo.FN_GET_RESULTBUTTON_YR(pp.RESULTBUTTON)) CLASSCNAME
,CONVERT(varchar, pp.STDate, 111) STDate,CONVERT(varchar, pp.FDDate, 111) FTDate
,v1.OrgName
FROM PLAN_PLANINFO pp
LEFT JOIN CLASS_CLASSINFO cc ON cc.PlanID=pp.PlanID AND cc.comidno=pp.comidno AND cc.seqno=pp.seqno
JOIN ID_PLAN ip ON ip.PlanID=pp.PlanID
JOIN VIEW_RIDNAME v1 ON v1.RID=pp.RID
WHERE pp.IsApprPaper='Y'
"
        SQL_1 &= $" AND ip.TPlanID='{sm.UserInfo.TPlanID}' AND ip.YEARS='{sm.UserInfo.Years}'" & vbCrLf
        Select Case sTypePC
            Case "P" 'PLAN
                SQL_1 &= " AND pp.TransFlag='N'" & vbCrLf '未轉班
            Case "C" 'CLASS
                SQL_1 &= " AND pp.TransFlag='Y'" & vbCrLf '已轉班
        End Select
        '依申請階段
        If v_AppStage2 <> "" Then SQL_1 &= $" AND pp.AppStage='{v_AppStage2}'" & vbCrLf
        If sm.UserInfo.LID <> 0 Then
            SQL_1 &= $" AND ip.DistID='{sm.UserInfo.DistID}' AND pp.PlanID={sm.UserInfo.PlanID}" & vbCrLf
        End If
        If vRelShip <> "" Then SQL_1 &= $" AND v1.RelShip LIKE '{vRelShip}%'" & vbCrLf

        '28:產業人才投資方案
        orgid.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投-計畫別 1:產業人才投資計畫 /2:提升勞工自主學習計畫
            Dim v_PlanPoint As String = TIMS.GetListValue(PlanPoint)
            Select Case v_PlanPoint'PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    SQL_1 &= " AND v1.OrgKind <> '10'" & vbCrLf
                    orgid.Value = "G"
                Case Else
                    '提升勞工自主學習計畫
                    SQL_1 &= " AND v1.OrgKind='10'" & vbCrLf
                    orgid.Value = "W"
            End Select
        End If

        PageControler1.Visible = False
        DataGrid1.Visible = False
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(SQL_1, objconn)
        If TIMS.dtNODATA(dt) Then Return

        DataGrid1.Visible = True
        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        '班級狀態 0:未轉班 /1:已轉班
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Select Case v_Radio1'Radio1.SelectedValue
            Case "0" '未轉班
                Call CreateClass("P")
            Case "1" '已轉班
                Call CreateClass("C")
            Case Else
                Common.MessageBox(Me, "請選擇班級狀態!!")
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim chkbox1 As HtmlInputCheckBox = e.Item.FindControl("chkbox1")
                Dim PCSValue As HtmlInputHidden = e.Item.FindControl("PCSValue")
                Dim OCIDValue As HtmlInputHidden = e.Item.FindControl("OCIDValue")
                PCSValue.Value = Convert.ToString(drv("PCSValue"))
                OCIDValue.Value = Convert.ToString(drv("OCID"))
        End Select
    End Sub

    Protected Sub BtnPrint3_Click(sender As Object, e As EventArgs) Handles BtnPrint3.Click
        Hid_PCSVALUE.Value = ""
        Hid_OCIDValue.Value = ""
        Dim v_OCID As String = ""
        Dim v_PCSVALUE As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim chkbox1 As HtmlInputCheckBox = eItem.FindControl("chkbox1")
            Dim PCSValue As HtmlInputHidden = eItem.FindControl("PCSValue")
            Dim OCIDValue As HtmlInputHidden = eItem.FindControl("OCIDValue")
            If chkbox1.Checked AndAlso OCIDValue.Value <> "" Then
                If v_OCID = "" Then v_OCID = OCIDValue.Value
                Hid_OCIDValue.Value &= String.Concat(If(Hid_OCIDValue.Value <> "", ",", ""), TIMS.ClearSQM(OCIDValue.Value))
            End If
            If chkbox1.Checked AndAlso PCSValue.Value <> "" Then
                If v_PCSVALUE = "" Then v_PCSVALUE = Hid_PCSVALUE.Value
                Hid_PCSVALUE.Value &= String.Concat(If(Hid_PCSVALUE.Value <> "", ",", ""), TIMS.ClearSQM(PCSValue.Value))
            End If
        Next

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)
        Dim v_AppStage2 As String = TIMS.GetListValue(AppStage2)
        '當查詢條件之【申請階段】選擇： 1上半年 或 2下半年  >> 則訓練計畫總表不要印出「政策性產業課程可辦理班數」之欄位。
        '當查詢條件之【申請階段】選擇： 3政策性產業，則會印出「政策性產業課程可辦理班數」之欄位。
        Dim flag_AppStage2_val_3 As Boolean = False
        flag_AppStage2_val_3 = If(tr_AppStage_TP28.Visible, If(v_AppStage2 = "3", True, False), False)

        'SD_14_003_4_2009
        Dim MyValue As String = ""
        Dim vFileN As String = ""
        '班級狀態 0:未轉班 /1:已轉班
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Select Case v_Radio1'Radio1.SelectedValue
            Case "0" '未轉 'PP
                'If sm.UserInfo.Years >= 2019 Then
                '    vFileN = cst_print_19W
                '    If orgid.Value = "G" Then vFileN = cst_print_19G
                'ElseIf sm.UserInfo.Years >= 2018 Then
                '    vFileN = cst_print_18W
                '    If orgid.Value = "G" Then vFileN = cst_print_18G
                'Else
                '    vFileN = cst_print_09W
                '    If orgid.Value = "G" Then vFileN = cst_print_09G
                'End If

                vFileN = If(orgid.Value = "G", cst_print_19G, cst_print_19W)

                'If (flag_SHOW_2019_3 AndAlso flag_AppStage2_val_3) Then
                '    vFileN = cst_print_19W2
                '    If orgid.Value = "G" Then vFileN = cst_print_19G2
                'End If

            Case "1" '已轉
                'If sm.UserInfo.Years >= 2019 Then
                '    vFileN = cst_print_2019W
                '    If orgid.Value = "G" Then vFileN = cst_print_2019G
                'ElseIf sm.UserInfo.Years >= 2018 Then
                '    vFileN = cst_print_2018W
                '    If orgid.Value = "G" Then vFileN = cst_print_2018G
                'Else
                '    vFileN = cst_print_2009W
                '    If orgid.Value = "G" Then vFileN = cst_print_2009G
                'End If

                vFileN = If(orgid.Value = "G", cst_print_2019G, cst_print_2019W)

                'If (flag_SHOW_2019_3 AndAlso flag_AppStage2_val_3) Then
                '    vFileN = cst_print_2019W2
                '    If orgid.Value = "G" Then vFileN = cst_print_2019G2
                'End If

            Case Else
                Common.MessageBox(Me, "請選擇 班級狀態")
                Exit Sub
        End Select

        '班級狀態 0:未轉班 /1:已轉班
        'Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Select Case v_Radio1'Radio1.SelectedValue
            Case "0" '未轉 'PP
                If Hid_PCSVALUE.Value = "" Then
                    Common.MessageBox(Me, "請選擇計畫")
                    Exit Sub
                End If
                MyValue &= "&PCSVALUE=" & Hid_PCSVALUE.Value
            Case "1" '已轉
                If Hid_OCIDValue.Value = "" Then
                    Common.MessageBox(Me, "請選擇班級")
                    Exit Sub
                End If
                MyValue &= "&OCID=" & Hid_OCIDValue.Value
            Case Else
                Common.MessageBox(Me, "請選擇 班級狀態")
                Exit Sub
        End Select

        Select Case sm.UserInfo.LID
            Case 0
                If v_PCSVALUE <> "" Then
                    Dim drPP As DataRow = TIMS.GetPCSDate(v_PCSVALUE, objconn)
                    If drPP IsNot Nothing Then Hid_PlanID1.Value = Convert.ToString(drPP("PLANID"))
                ElseIf v_OCID <> "" Then
                    Dim drCC As DataRow = TIMS.GetOCIDDate(v_OCID, objconn)
                    If drCC IsNot Nothing Then Hid_PlanID1.Value = Convert.ToString(drCC("PLANID"))
                End If
                MyValue &= "&PlanID=" & Hid_PlanID1.Value
            Case Else
                MyValue &= "&PlanID=" & Convert.ToString(sm.UserInfo.PlanID)
        End Select
        MyValue &= "&Years=" & ROC_Years.Value

        '依申請階段 
        '表示 (1：上半年、2：下半年、3：政策性產業)
        If tr_AppStage_TP28.Visible Then
            'Dim v_AppStage2 As String = TIMS.GetListValue(AppStage2)
            If v_AppStage2 <> "" AndAlso v_AppStage2 > "0" Then MyValue &= "&AppStage=" & v_AppStage2
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vFileN, MyValue)
    End Sub
End Class