Partial Class SD_04_015
    Inherits AuthBasePage

    'SD_04_015
    'SD_04_015_2.jrxml
    'SD_04_015_2_ds
    'SD_04_015_3
    'SD_04_015_3_ds
    Const cst_printFN1 As String = "SD_04_015"
    Const cst_printFN2 As String = "SD_04_015_2"
    Const cst_printFN3 As String = "SD_04_015_3"
    Const cst_prt2txt70 As String = "列印助教"

    'Dim objconn As SqlConnection
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        'objconn = DbAccess.GetConnection()
        If sm.UserInfo.TPlanID = TIMS.Cst_TPlanID70 Then
            btnPrint2.Text = cst_prt2txt70
            btnPrint3.Visible = False
        End If

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

        btnPrint1.Attributes("onclick") = "return check_data();"

        '若只有管理一個班級，自動協助帶出班級--by AMU 2009-02
        'TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '列印教師
    Private Sub btnPrint1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint1.Click
        Dim RID As String = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(myValue, "RID", RID)
        TIMS.SetMyValue(myValue, "TPlanID", sm.UserInfo.TPlanID)
        ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub

    '列印第二教師1
    Private Sub btnPrint2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint2.Click
        Dim RID As String = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(myValue, "RID", RID)
        TIMS.SetMyValue(myValue, "TPlanID", sm.UserInfo.TPlanID)
        'Dim str_printFN2 As String = cst_printFN2
        'If sm.UserInfo.TPlanID = TIMS.Cst_TPlanID70 Then str_printFN2 = cst_printFN2b
        ReportQuery.PrintReport(Me, cst_printFN2, myValue)
    End Sub

    '列印第二教師2
    Private Sub btnPrint3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint3.Click
        Dim RID As String = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(myValue, "RID", RID)
        TIMS.SetMyValue(myValue, "TPlanID", sm.UserInfo.TPlanID)
        ReportQuery.PrintReport(Me, cst_printFN3, myValue)
    End Sub

End Class

