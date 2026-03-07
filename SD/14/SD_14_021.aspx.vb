Partial Class SD_14_021
    Inherits AuthBasePage

    '//已轉班
    'G BussinessTrain: SD_14_021 '產業人才投資計畫 
    'W BussinessTrain: SD_14_021_b '提升勞工自主學習計畫    '(vp.Years-1911)    
    'SD_14_021*.jrxml
    Const cst_printFN_G1 As String = "SD_14_021" '1:2017前
    Const cst_printFN_W1 As String = "SD_14_021_b" '1:2017前
    Const cst_printFN_G2 As String = "SD_14_021G" '2:2017
    Const cst_printFN_W2 As String = "SD_14_021W_b" '2:2017
    '報表分為兩頁印
    Const cst_printFN_G1_2 As String = "SD_14_021_2page" '1:2017前 / SD_14_021
    Const cst_printFN_W1_2 As String = "SD_14_021_b_2page" '1:2017前 / SD_14_021_b
    Const cst_printFN_G2_2 As String = "SD_14_021G_2page" '2:2017 / SD_14_021G
    Const cst_printFN_W2_2 As String = "SD_14_021W_b_2page" '2:2017 / SD_14_021W_b

    'NEW 補助學員參訓契約書 參訓學員簽訂之契約書
    Const cst_printFN_G3 As String = "SD_14_021G3" '3:2018之後
    Const cst_printFN_W3 As String = "SD_14_021W3" '3:2018之後
    Const cst_printFN_G3_2 As String = "SD_14_021G3_2page" '3:2018之後 / SD_14_021G3
    Const cst_printFN_W3_2 As String = "SD_14_021W3_2page" '3:2018之後 / SD_14_021W3
    'Const cst_printFN_G3O As String = "OJTSD1421G3" '3:2018之後(報名網空白)
    'Const cst_printFN_W3O As String = "OJTSD1421W3" '3:2018之後(報名網空白)
    'Const cst_printFN_G3S As String = "OJTSD1421G3S" '3:2018之後(報名網簽名)
    'Const cst_printFN_W3S As String = "OJTSD1421W3S" '3:2018之後(報名網簽名)

    Dim prtFilename As String = "" '列印表件名稱
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        iPYNum = TIMS.sUtl_GetPYNum(Me)
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

    Sub Search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = ""
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,Convert(nvarchar, pp.PlanID) + '_' +Convert(nvarchar, pp.ComIDNO)" & vbCrLf
        sql &= " + '_' + Convert(nvarchar,pp.SeqNo) PCSValue" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,pp.CYCLTYPE) ClassCName" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.FTDate, 111) FTDate" & vbCrLf
        sql &= " ,v1.OrgName " & vbCrLf
        sql &= " FROM dbo.PLAN_PLANINFO pp" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO cc on cc.PlanID=pp.PlanID AND cc.comidno=pp.comidno AND cc.seqno=pp.seqno" & vbCrLf
        sql &= " JOIN dbo.ID_Plan ip on ip.PlanID=pp.PlanID" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP ar on ar.RID=cc.RID" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME v1 on pp.RID=v1.RID" & vbCrLf
        '限制為只有正式儲存之班級 '已轉班
        sql &= " WHERE pp.IsApprPaper='Y' AND pp.TransFlag ='Y' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " and ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            'sql += " and pp.PlanID='177'" & vbCrLf
            sql &= " and pp.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        If RelShip <> "" Then
            sql &= " AND ar.RelShip like '" & RelShip & "%'" & vbCrLf
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value <> "" Then
            'sql += " AND cc.OCID='924'" & vbCrLf
            sql &= " AND cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If

        '28:產業人才投資方案
        hidorgid.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sql &= " AND v1.OrgKind <> '10'" & vbCrLf
                    hidorgid.Value = "G"
                Case Else
                    '提升勞工自主學習計畫
                    sql &= " AND v1.OrgKind  = '10'" & vbCrLf
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
        '1:已轉班
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"

            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""

                Dim drv As DataRowView = e.Item.DataItem
                Dim checkbox1 As HtmlInputCheckBox = e.Item.FindControl("checkbox1")
                Dim OCID As HiddenField = e.Item.FindControl("OCID")
                'Dim PCSValue As HiddenField = e.Item.FindControl("PCSValue")
                OCID.Value = Convert.ToString(drv("OCID"))

                Dim f_OCID As String = String.Concat("'", OCID.Value, "'")
                checkbox1.Checked = (hidOCIDValue.Value.IndexOf(f_OCID) > -1)
        End Select
    End Sub

    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""

        '20180723 加入 列印方式(可分兩頁印)
        Dim v_Print_Option As String = TIMS.GetListValue(Print_Option)
        'printtype.Value = Print_Option.SelectedValue
        Select Case v_Print_Option
            Case "1", "2"
            Case Else
                Common.MessageBox(Me, "請選擇 列印選項!")
                Return
        End Select

        hidorgid.Value = TIMS.ClearSQM(hidorgid.Value)
        'Dim vsFileName1 As String = ""
        If v_Print_Option = "1" Then
            If iPYNum >= 3 Then
                prtFilename = If(hidorgid.Value = "G", cst_printFN_G3, If(hidorgid.Value = "W", cst_printFN_W3, "")) 'cst_printFN_G3 '3:2018
            ElseIf iPYNum = 2 Then
                prtFilename = If(hidorgid.Value = "G", cst_printFN_G2, If(hidorgid.Value = "W", cst_printFN_W2, "")) 'cst_printFN_G2 '2:2017
            Else
                prtFilename = If(hidorgid.Value = "G", cst_printFN_G1, If(hidorgid.Value = "W", cst_printFN_W1, ""))  '1:2017前
            End If
        Else
            '20180723 加入 列印方式(可分兩頁印)
            If iPYNum >= 3 Then
                prtFilename = If(hidorgid.Value = "G", cst_printFN_G3_2, If(hidorgid.Value = "W", cst_printFN_W3_2, ""))  'cst_printFN_W3_2 '3:2018
            ElseIf iPYNum = 2 Then
                prtFilename = If(hidorgid.Value = "G", cst_printFN_G2_2, If(hidorgid.Value = "W", cst_printFN_W2_2, ""))  'cst_printFN_W2_2 '2:2017
            Else
                prtFilename = If(hidorgid.Value = "G", cst_printFN_G1_2, If(hidorgid.Value = "W", cst_printFN_W1_2, ""))  'cst_printFN_W1_2'1:2017前
            End If
        End If

        If sm.UserInfo.TPlanID = TIMS.Cst_TPlanID54AppPlan Then
            If v_Print_Option = "1" Then
                prtFilename = cst_printFN_G1
                If iPYNum = 2 Then prtFilename = cst_printFN_G2
                If iPYNum >= 3 Then prtFilename = cst_printFN_G3
            Else
                '20180723 加入 列印方式(可分兩頁印)
                prtFilename = cst_printFN_G1_2
                If iPYNum = 2 Then prtFilename = cst_printFN_G2_2
                If iPYNum >= 3 Then prtFilename = cst_printFN_G3_2
            End If
        End If

        If prtFilename = "" Then
            Errmsg += "請選擇計畫種類 !" & vbCrLf
        End If

        hidOCIDValue.Value = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            Dim OCID As HiddenField = eItem.FindControl("OCID")
            If checkbox1.Checked And OCID.Value <> "" Then
                hidOCIDValue.Value &= String.Concat(If(hidOCIDValue.Value <> "", ",", ""), "'", OCID.Value, "'")
            End If
        Next

        If Trim(hidOCIDValue.Value) = "" Then
            Errmsg += "請選擇 職類/班別 !" & vbCrLf
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Dim v_OCIDValue1 As String = TIMS.ClearSQM(hidOCIDValue.Value)
        Dim prtstr As String = ""
        'prtstr += "&Years=" & hidYears.Value
        prtstr += "&TPlanID=" & sm.UserInfo.TPlanID
        prtstr += "&OCID=" & v_OCIDValue1
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, prtstr)
    End Sub
End Class
