Partial Class SD_14_023
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_14_023" 'SD_14_023.jrxml '(原結訓證書)'(含技檢訓練時數 顯示)
    'Const cst_printFN1b As String = "SD_14_023b" '(含技檢訓練時數 顯示)
    Const cst_printFN1c As String = "SD_14_023c" '技檢訓練時數清單

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

        'PageControler1.PageDataGrid = DataGrid1

        hidYears.Value = sm.UserInfo.Years - 1911 '設定登入民國年

        If Not IsPostBack Then
            '列印
            btnPrint1.Attributes("onclick") = "return CheckPrint();"
            Button4.Attributes("onclick") = "ClearData();"
            'msg.Text = ""                  '每次 清空
            'DataGridTable.Visible = False  '預設 隱藏
            'ClassTR.Visible = False        '預設 隱藏
            hidOCIDValue.Value = ""
            'hidPCSValue.Value = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'Common.SetListItem(Radio1, "0")  '預設 未轉班
            'Me.Radio1.SelectedIndex = 0
            'PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint)
            'Common.SetListItem(PlanPoint, "1")

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2
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

    ''' <summary>列印按鈕</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Errmsg += "請選擇 職類/班別 !" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        'Dim sTEHOURS As String = TIMS.GetPCS_TEHOURS(drCC("PLANID"), drCC("COMIDNO"), drCC("SEQNO"), objconn)
        'Dim s_PrintNM1 As String = If(sTEHOURS <> "" AndAlso Val(sTEHOURS) > 0, cst_printFN1b, cst_printFN1)
        'prtstr = "" 'prtstr += "&Years=" & hidYears.Value
        '(add『證書編號』欄位，by:20180803)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, s_PrintNM1, prtstr)

        txtCert.Text = TIMS.ClearSQM(txtCert.Text)
        Dim prtstr As String = ""
        prtstr &= String.Concat("&TPlanID=", sm.UserInfo.TPlanID)
        prtstr &= String.Concat("&OCID=", drCC("OCID"))
        prtstr &= String.Concat("&MYCERT=", txtCert.Text)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, prtstr)
    End Sub

    ''' <summary>技檢訓練時數清單</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnPrint2_Click(sender As Object, e As EventArgs) Handles btnPrint2.Click
        Dim Errmsg As String = ""
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Errmsg += "請選擇 職類/班別 !" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim sTEHOURS As String = TIMS.GetPCS_TEHOURS(drCC("PLANID"), drCC("COMIDNO"), drCC("SEQNO"), objconn)
        If (sTEHOURS = "" OrElse Val(sTEHOURS) <= 0) Then
            Errmsg = "此班課程無符合技能檢定訓練時數!"
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Dim prtstr As String = ""
        prtstr &= String.Concat("&TPlanID=", sm.UserInfo.TPlanID)
        prtstr &= String.Concat("&OCID=", drCC("OCID"))
        prtstr &= String.Concat("&RID=", drCC("RID"))
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1c, prtstr)
    End Sub
End Class