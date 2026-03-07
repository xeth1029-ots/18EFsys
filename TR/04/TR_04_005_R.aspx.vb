Partial Class TR_04_005_R
    Inherits AuthBasePage

    'ReportQuery
    'TR_04_005_R_1
    'TR_04_005_R_2
    'TR_04_005_R_3
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
            sqlstr = "SELECT Name,DistID FROM ID_District"
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
        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '報表要用的轄區參數
        Dim DistID As String = ""
        DistID = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected Then
                If DistID <> "" Then DistID += ","
                DistID += Convert.ToString("\'" & Me.DistrictList.Items(i).Value & "\'")
            End If
        Next

        Dim stitle As String = ""
        If STDate1.Text <> "" OrElse STDate2.Text <> "" Then
            stitle = STDate1.Text + " ~ " + STDate2.Text
        End If
        Dim etitle As String = ""
        If FTDate1.Text <> "" OrElse FTDate2.Text <> "" Then
            etitle = FTDate1.Text + " ~ " + FTDate2.Text
        End If

        'Const Cst_MIdentityID As String = "'02','03','04','05','06','07','10','13','14','27','28','35'"
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)

        Dim sMyValue As String = ""
        sMyValue = ""
        sMyValue &= "&path=TIMS"
        sMyValue &= "&TPlanID=" & TPlanID.SelectedValue
        sMyValue &= "&PYear=" & Syear.SelectedValue
        If DistID <> "" Then
            sMyValue &= "&DistID=" & DistID
        End If
        sMyValue &= "&STDate1=" & STDate1.Text
        sMyValue &= "&STDate2=" & STDate2.Text
        sMyValue &= "&FTDate1=" & FTDate1.Text
        sMyValue &= "&FTDate2=" & FTDate2.Text
        sMyValue &= "&stitle=" & stitle
        sMyValue &= "&etitle=" & etitle
        sMyValue &= "&PlanName=" & TPlanID.SelectedItem.Text & "" '中文跳脫字
        'sMyValue += "&MIdentityID=" & Replace(Cst_MIdentityID, "'", "\'")

        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        ''strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R&path=TIMS&TPlanID=" & TPlanID.SelectedValue & "&PYear=" & Syear.SelectedValue & "&DistID=" & newDistID & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle & "&PlanName='+escape('" & TPlanID.SelectedItem.Text & "')+'');" + vbCrLf
        ''開訓人數
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R_1" & sMyValue & "');" + vbCrLf
        ''結訓人數
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R_2" & sMyValue & "');" + vbCrLf
        ''就業人數
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R_3" & sMyValue & "');" + vbCrLf
        ' ''開訓人數
        ''strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R_1&path=TIMS&TPlanID=" & TPlanID.SelectedValue & "&PYear=" & Syear.SelectedValue & "&DistID=" & DistID & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle & "&PlanName='+escape('" & TPlanID.SelectedItem.Text & "')+'');" + vbCrLf
        ' ''結訓人數
        ''strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R_2&path=TIMS&TPlanID=" & TPlanID.SelectedValue & "&PYear=" & Syear.SelectedValue & "&DistID=" & DistID & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle & "&PlanName='+escape('" & TPlanID.SelectedItem.Text & "')+'');" + vbCrLf
        ' ''就業人數
        ''strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=TR&filename=TR_04_005_R_3&path=TIMS&TPlanID=" & TPlanID.SelectedValue & "&PYear=" & Syear.SelectedValue & "&DistID=" & DistID & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle & "&PlanName='+escape('" & TPlanID.SelectedItem.Text & "')+'');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "TR_04_005_R_1", sMyValue)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "TR_04_005_R_2", sMyValue)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "TR_04_005_R_3", sMyValue)
    End Sub
End Class
