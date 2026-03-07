Partial Class TR_04_004_R
    Inherits AuthBasePage

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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Dim dt As DataTable
            Dim sqlstr As String
            sqlstr = "SELECT NAME,DISTID FROM ID_DISTRICT ORDER BY DISTID"
            dt = DbAccess.GetDataTable(sqlstr, objconn)
            Me.DistrictList.DataSource = dt
            Me.DistrictList.DataTextField = "Name"
            Me.DistrictList.DataValueField = "DistID"
            Me.DistrictList.DataBind()
            Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

            Syear = TIMS.GetSyear(Syear)
            Common.SetListItem(Syear, Now.Year)

            TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
            TPlanID.Items(0).Selected = True

        End If

        Button1.Attributes("OnClick") = "javascript:return chk()"

        If sm.UserInfo.DistID = "000" Then
            DistrictList.Enabled = True
        Else
            DistrictList.SelectedValue = sm.UserInfo.DistID
            DistrictList.Enabled = False
        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim sdate, edate, stitle, etitle, DistID, newDistID As String
        'Dim i As Integer

        '報表要用的轄區參數
        Dim DistID As String = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected AndAlso Me.DistrictList.Items(i).Value <> "" Then
                If DistID <> "" Then DistID &= ","
                DistID &= "\'" & Me.DistrictList.Items(i).Value & "\'"
            End If
        Next

        'If DistID <> "" Then
        '    newDistID = Mid(DistID, 1, DistID.Length - 1)
        'End If

        Dim sdate As String = Syear.SelectedValue + "/1/1"
        Dim edate As String = Syear.SelectedValue + "/12/31"
        Dim stitle As String = ""
        Dim etitle As String = ""
        If STDate1.Text <> "" Or STDate2.Text <> "" Then
            stitle = STDate1.Text + " ~ " + STDate2.Text
        End If
        If FTDate1.Text <> "" Or FTDate2.Text <> "" Then
            etitle = FTDate1.Text + " ~ " + FTDate2.Text
        End If
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_004_R&path=TIMS&TPlanID=" & TPlanID.SelectedValue & "&PYear=" & Syear.SelectedValue & "&PlanName='+escape('" & TPlanID.SelectedItem.Text & "')+'&AppliedDate=" & sdate & "&AppliedDate1=" & edate & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle & "&DistID=" & newDistID & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

        Dim MyValue As String = ""
        MyValue = "TPlanID=" & TPlanID.SelectedValue
        MyValue &= "&PYear=" & Syear.SelectedValue
        MyValue &= "&PlanName=" & TPlanID.SelectedItem.Text
        MyValue &= "&AppliedDate=" & sdate
        MyValue &= "&AppliedDate1=" & edate
        MyValue &= "&STDate1=" & STDate1.Text
        MyValue &= "&STDate2=" & STDate2.Text
        MyValue &= "&FTDate1=" & FTDate1.Text
        MyValue &= "&FTDate2=" & FTDate2.Text
        MyValue &= "&stitle=" & stitle
        MyValue &= "&etitle=" & etitle
        If DistID <> "" Then
            MyValue &= "&DistID=" & DistID
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "TR_04_004_R", MyValue)
    End Sub
End Class
