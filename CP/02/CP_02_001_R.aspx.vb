Partial Class CP_02_001_R
    Inherits AuthBasePage

    Dim objstr As String
    Dim objtable As DataTable

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
            Button1.Attributes("onclick") = "javascript:return search()"
            '取得訓練計畫
            'TPlan = TIMS.Get_TPlan(TPlan, , 1)
            Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn, 1)
        End If

        chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
        chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
        chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"

    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(start_date.Text) <> "" Then start_date.Text = Trim(start_date.Text) Else start_date.Text = ""
        If Trim(end_date.Text) <> "" Then end_date.Text = Trim(end_date.Text) Else end_date.Text = ""

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

        'Dim OC1 As String
        'Dim OC_TYPE As String
        'Dim TPlanID As String
        'Dim a As Integer
        'Dim newTPlanID As String
        '判斷統計對象
        Dim OC_TYPE As String = ""
        Dim OC1 As String = Convert.ToString(OCID.SelectedValue)
        If OC1 <> "" Then
            If OC1 = "1" Then
                OC_TYPE = "1"
            End If
            If OC1 = "2" Then
                OC_TYPE = "0"
            End If
        End If

        Dim TPlanID As String = ""
        TPlanID = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 3)
        'newTPlanID = TPlanID
        'For a = 0 To Me.TPlan.Items.Count - 1
        '    If Me.TPlan.Items(a).Selected Then
        '        If TPlanID = "" Then
        '            TPlanID = Convert.ToString("\'" & Me.TPlan.Items(a).Value & "\'" & ",")
        '        Else
        '            TPlanID = TPlanID & Convert.ToString("\'" & Me.TPlan.Items(a).Value & "\'" & ",")
        '        End If
        '    End If
        'Next
        'If TPlanID <> "" Then
        '    newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        'End If

        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        Dim Y1, Y2, SYMD, EYMD As String
        Y1 = Convert.ToInt16(Convert.ToString(start_date.Text).Substring(0, 4)) - 1911
        SYMD = Y1 & "年" & Convert.ToString(start_date.Text).Substring(5, 2) & "月"
        Y2 = Convert.ToInt16(Convert.ToString(end_date.Text).Substring(0, 4)) - 1911
        EYMD = Y2 & "年" & Convert.ToString(end_date.Text).Substring(5, 2) & "月"

        Dim strListX As String = ""
        Dim strListY As String = ""
        strListX = ""
        For i As Integer = 0 To Me.SortX.Items.Count - 1
            If Me.SortX.Items(i).Selected Then
                If strListX <> "" Then strListX &= ","
                strListX &= Me.SortX.Items(i).Value
            End If
        Next
        strListY = ""
        For j As Integer = 0 To Me.SortY.Items.Count - 1
            If Me.SortY.Items(j).Selected Then
                If strListY <> "" Then strListY &= ","
                strListY &= Me.SortY.Items(j).Value
            End If
        Next

        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=CP_02_001_Rpt&path=TIMS&start_date=" & start_date.Text & "&end_date=" & end_date.Text & "&OC_TYPE=" & OC_TYPE & "&TPlan=" & newTPlanID & "&P1='+escape('" & strListX & "')+'&P2='+escape('" & strListY & "')+ '&SYMD='+escape('" & SYMD & "')+'&EYMD='+escape('" & EYMD & "'));" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

        Dim sUrl As String = "CP_02_001_R_Rpt.aspx"
        Dim MyValue As String = ""
        MyValue &= $"&start_date={start_date.Text}&end_date={end_date.Text}"
        MyValue &= $"&OC_TYPE={OC_TYPE}&TPlan={TPlanID}"
        MyValue &= $"&P1={strListX}&P2={strListY}"
        MyValue &= $"&SYMD={SYMD}&EYMD={EYMD}"
        ReportQuery.Redirect(Me, sUrl, MyValue)

    End Sub

End Class
