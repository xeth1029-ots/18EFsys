Partial Class CM_03_006_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "CM_03_006_R"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'TIMS.TestDbConn(Me, objConn)
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            CreateItem()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Sub CreateItem()
        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)
        'Print.Attributes("onclick") = "CheckPrint();return false;"

        'Dim dtCity As DataTable = Nothing
        'Dim dr As DataRow
        'Dim sqlstr As String
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        '縣市
        Dim dtCity As DataTable = TIMS.Get_dtCity(objconn)
        CityID = TIMS.Get_CityName(CityID, dtCity)

        Print.Attributes("onclick") = "return CheckPrint();"
        '選擇全部縣市
        CityID.Attributes("onclick") = "SelectAll('CityID','CityHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '報表要用的縣市
        Dim ICityID1 As String = ""
        For i As Integer = 1 To Me.CityID.Items.Count - 1
            If Me.CityID.Items(i).Selected AndAlso CityID.Items(i).Value <> "" Then
                If ICityID1 <> "" Then ICityID1 &= ","
                ICityID1 &= CityID.Items(i).Value
            End If
        Next

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= TPlanID.Items(i).Value
            End If
        Next

        Dim ICityIDa As String = "x"
        If ICityID1 = "" Then ICityIDa = "a" '全部
        Dim TPlanIDa As String = "x"
        If TPlanID1 = "" Then TPlanIDa = "a" '全部

        Dim MyValue As String = ""
        MyValue = ""
        'MyValue &= "&yearlist=" & Mid(Me.yearlist.SelectedValue, 3, 2)
        MyValue &= "&SYear=" & Me.yearlist.SelectedValue
        If ICityID1 <> "" Then MyValue &= "&ICityID=" & ICityID1
        If TPlanID1 <> "" Then MyValue &= "&TPlanID=" & TPlanID1
        MyValue &= "&ICityIDa=" & ICityIDa
        MyValue &= "&TPlanIDa=" & TPlanIDa
        MyValue &= "&RID=" & RIDValue.Value
        MyValue &= "&OrgName=" & TIMS.UrlEncode1(center.Text)
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

        'ReportQuery.PrintReport(Me, "TR", "CM_03_006_R", "yearlist=" & Mid(Me.yearlist.SelectedValue, 3, 2) & _
        '                                                    "&SYear=" & Me.yearlist.SelectedValue & _
        '                                                    "&ICityID=" & ICityID1 & "&ICityName=" & newICityName & _
        '                                                    "&TPlanID=" & TPlanID1 & "&RID=" & RIDValue.Value & _
        '                                                    "&PlanName=" & Convert.ToString(newTPlanIDName) & _
        '                                                    "&OrgName=" & Convert.ToString(center.Text))
    End Sub
End Class
