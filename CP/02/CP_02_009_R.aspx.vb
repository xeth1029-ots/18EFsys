Partial Class CP_02_009_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "Train_list"
    'Train_list
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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Create1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button1.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button2.Attributes("onclick") = "javascript:return print();"

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Sub Create1()
        syear = TIMS.GetSyear(syear)
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '預算別 BUDID IN ('02','03')
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            BudID = TIMS.Get_Budget(BudID, 39, objconn)
        Else
            BudID = TIMS.Get_Budget(BudID, 37, objconn)
        End If

        Dim t_title As String = "依使用者登入之計畫，僅統計該計畫的數據"
        TIMS.Tooltip(Button2, t_title)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim sql As String = "select Relship  from Auth_Relship where RID='" & RIDValue.Value & "'"
        'Dim dr As DataRow
        'dr = DbAccess.GetOneRow(sql)
        'Dim relship_str As String = dr("Relship")
        'Dim strScript As String
        'Dim MyValue As String = ""
        Dim sRelship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
        If sRelship = "" Then
            Common.MessageBox(Me, "查詢資料有誤!!")
            Exit Sub
        End If
        Dim s_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        If s_DistID = "" Then
            Common.MessageBox(Me, "查詢資料有誤!!")
            Exit Sub
        End If

        'If Me.CheckBox1.Checked Then
        '    MyValue = "Relship=" & sRelship & "&Years=" & syear.SelectedValue
        'Else
        '    MyValue = "RID=" & Me.RIDValue.Value & "&Years=" & syear.SelectedValue
        'End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim v_BudID As String = TIMS.GetCblValue(BudID)
        If v_BudID = "" Then
            '預算別 BUDID IN ('02','03')
            v_BudID = "02"
            If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then v_BudID = "02,03"
        End If

        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "RID", RIDValue.Value)
        TIMS.SetMyValue(MyValue, "DistID", s_DistID)
        TIMS.SetMyValue(MyValue, "Years", TIMS.GetListValue(syear))
        TIMS.SetMyValue(MyValue, "BUDID", v_BudID)
        TIMS.SetMyValue(MyValue, "TPlanID", sm.UserInfo.TPlanID)
        'MyValue = "RID=" & RIDValue.Value & "&Years=" & syear.SelectedValue & "&TPlanID=" & sm.UserInfo.TPlanID
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

    End Sub

End Class
