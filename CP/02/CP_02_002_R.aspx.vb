Partial Class CP_02_002_R
    Inherits AuthBasePage

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

        If Not Me.IsPostBack Then
            'Dim Sqlstr As String
            Button1.Attributes("onclick") = "javascript:return search()"
            '取得訓練計畫
            'TPlan = TIMS.Get_TPlan(TPlan, , 1)
            Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn, 1)

            chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
            chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
            chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"
        End If
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = Trim(start_date.Text)
        end_date.Text = Trim(end_date.Text)

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "結訓日期 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            Errmsg += "結訓日期 起始日期 為必填" & vbCrLf
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "結訓日期 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            Errmsg += "結訓日期 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then
                    Errmsg += "【結訓日期】的起日不得大於【結訓日期】的迄日!!" & vbCrLf
                End If
            End If
        End If

        Dim Y1, Y2, SYMD, EYMD As String
        Try
            Y1 = CInt(Convert.ToString(start_date.Text).Substring(0, 4)) - 1911
            SYMD = Y1 & "年" & Convert.ToString(start_date.Text).Substring(5, 2) & "月"
            Y2 = CInt(Convert.ToString(end_date.Text).Substring(0, 4)) - 1911
            EYMD = Y2 & "年" & Convert.ToString(end_date.Text).Substring(5, 2) & "月"
        Catch ex As Exception
            Errmsg += "開、結訓日期 格式有誤" & vbCrLf
        End Try

        If Convert.ToString(OCID.SelectedValue) = "" _
            OrElse Not IsNumeric(OCID.SelectedValue) Then
            Errmsg += "統計對象 為必選" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim OC As String
        'Dim OC1 As String
        'Dim OC_TYPE As String
        'Dim TPlanID As String
        'Dim i As Integer
        'Dim newTPlanID As String

        '判斷統計對象
        'OC1 = Convert.ToString(OCID.SelectedValue)

        Dim OC_TYPE As String = "" '1:局屬/0:非局屬
        OC_TYPE = "" '署(局)屬/非署(局)屬
        Select Case Convert.ToString(OCID.SelectedValue)
            Case "1"
                OC_TYPE = "1" '署(局)屬
            Case "2"
                OC_TYPE = "0" '非署(局)屬
        End Select

        Dim TPlanID As String = "" '選擇計畫
        'TPlanID = ""
        'For i As Integer = 0 To Me.TPlan.Items.Count - 1
        '    If Me.TPlan.Items(i).Selected Then
        '        If TPlanID <> "" Then TPlanID += ","
        '        TPlanID += Convert.ToString("\'" & Me.TPlan.Items(i).Value & "\'")
        '    End If
        'Next
        TPlanID = ""
        TPlanID = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 3)

        Dim Y1, Y2, SYMD, EYMD As String
        Y1 = CInt(Convert.ToString(start_date.Text).Substring(0, 4)) - 1911
        SYMD = Y1 & "年" & Convert.ToString(start_date.Text).Substring(5, 2) & "月"
        Y2 = CInt(Convert.ToString(end_date.Text).Substring(0, 4)) - 1911
        EYMD = Y2 & "年" & Convert.ToString(end_date.Text).Substring(5, 2) & "月"

        Dim MyValue As String = ""
        Dim strScript As String = ""

        MyValue = "jkl=jkl"
        MyValue += "&OC_TYPE=" & OC_TYPE
        MyValue += "&start_date=" & start_date.Text
        MyValue += "&end_date=" & end_date.Text
        MyValue += "&TPlan=" & TPlanID
        MyValue += "&SYMD=" & SYMD
        MyValue += "&EYMD=" & EYMD
        MyValue += "&start_date2=" & start_date.Text
        MyValue += "&end_date2=" & end_date.Text
        MyValue += "&TPlan2=" & TPlanID

        'Select Case OC_TYPE
        '    Case "1" '署(局)屬(依結訓日期統計)
        '        'by AMU 2011
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "CP_02_002_R2", MyValue)

        '    Case Else '2:非署(局)屬 或 全選(依結訓學員資料卡填寫日)
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "CP_02_002_Rpt", MyValue)

        'End Select

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "window.open('CP_02_002_R_Prt.aspx?" + MyValue + "');" + vbCrLf
        strScript += "</script>"
        TIMS.RegisterStartupScript(Me, "window_onload", strScript)

    End Sub

End Class