Partial Class Translation
    Inherits AuthBasePage

    Dim vMsg1 As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then Call CreateItem1()
    End Sub

    Sub CreateItem1()
        '將傳值代入
        hidSN.Value = Server.UrlDecode(Request("sn"))
        hidName.Value = Server.UrlDecode(Request("name"))
        Select Case hidName.Value.Length
            Case 1
                txtName1.Text = hidName.Value
            Case 2, 3
                txtName1.Text = Mid(hidName.Value, 1, 1)
                txtName2.Text = Mid(hidName.Value, 2)
            Case Else
                txtName1.Text = Mid(hidName.Value, 1, 2)
                txtName2.Text = Mid(hidName.Value, 3)
        End Select

        If hidSN.Value = "stud" Then
            If Split(Server.UrlDecode(Request("field")), ",").Length > 1 Then
                hidRtnID1.Value = Split(Server.UrlDecode(Request("field")), ",")(0)
                hidRtnID2.Value = Split(Server.UrlDecode(Request("field")), ",")(1)
            End If
        Else
            hidRtnID1.Value = Server.UrlDecode(Request("field"))
        End If
        btnSent.Attributes.Add("onclick", "sentValue();")
        btnLev.Attributes.Add("onclick", "window.close();")
        btnSch.Attributes.Add("onclick", "return chkSch();")
        'btnSch_Click(sender, e)
        Call Search1()

        '如果是系統管理者開啟功能。
        btnAdd.Visible = False
        If TIMS.IsSuperUser(Me, 1) Then btnAdd.Visible = True
    End Sub

    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        Dim strWord As String = "" '中文姓名
        Dim strOWord As String = "" '記錄將比對中文(單一)
        Dim strCname As String = "" '記錄查詢資料中文清單
        'Dim strOCname As String = "" '記錄查詢資料中文(單一)
        Dim strEng As String = "" '記錄比對結果取得之英文
        Dim intCnt As Integer = 0 '控制是否查到資料(0=>否, 1=>有)
        Dim msg As String = ""
        Dim sql As String = ""

        '查詢中翻英比對資料
        sql = " SELECT * FROM ID_TRANSLATION ORDER BY ENGID "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        For i As Integer = 1 To 2
            strEng = ""
            If i = 1 Then
                strWord = txtName1.Text
            Else
                strWord = txtName2.Text
            End If
            '比對中文
            For j As Integer = 1 To strWord.Length
                strOWord = Mid(strWord, j, 1) '取中文姓名一個字
                For k As Integer = 0 To dt.Rows.Count - 1
                    Dim dr As DataRow = Nothing
                    dr = dt.Rows(k)
                    strCname = dr("cname") '取資料庫比對字串
                    If strCname.IndexOf(strOWord) > -1 Then
                        intCnt = 1
                        If strEng = "" Then
                            strEng = Replace(dr("engid"), " ", "")
                        Else
                            strEng += "," + Replace(dr("engid"), " ", "")
                        End If
                        Exit For
                    End If
                Next
                If intCnt = 0 Then
                    If strEng = "" Then
                        strEng = strOWord
                    Else
                        strEng += "," + strOWord
                    End If
                    msg += "「" & strOWord & "」"
                End If
                intCnt = 0
            Next
            '顯示
            If i = 1 Then
                txtEng1.Text = ""
                Select Case Split(strEng, ",").Length
                    Case 1
                        txtEng1.Text = Split(strEng, ",")(0)
                    Case Else
                        txtEng1.Text = Replace(strEng, ",", " ")
                End Select
            Else
                txtEng2.Text = ""
                Select Case Split(strEng, ",").Length
                    Case 1
                        txtEng2.Text = Split(strEng, ",")(0)
                    Case Else
                        txtEng2.Text = Replace(strEng, ",", "-")
                End Select
            End If
        Next

        If msg <> "" Then
            vMsg1 = msg & "無法翻譯英文!"
            Common.MessageBox(Me, vMsg1)
            '偵錯用儲存欄
            'Dim strErrmsg As String = ""
            'strErrmsg += "[Translation]" & vbCrLf
            'strErrmsg += vMsg1 & vbCrLf
            'strErrmsg += "txtName1.Text:" & txtName1.Text & vbCrLf
            'strErrmsg += "txtName2.Text:" & txtName2.Text & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            'Call TIMS.SendMailTest(strErrmsg)
        End If
    End Sub

    Private Sub btnSch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSch.Click
        Call Search1()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        panSch.Visible = False
        panAdd.Visible = True
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Dim sql As String = ""
        Dim strWord As String = ""

        txtAWord.Text = TIMS.ClearSQM(txtAWord.Text)
        txtAEng.Text = TIMS.ClearSQM(txtAEng.Text)
        sql = " UPDATE ID_TRANSLATION SET cname = @cname WHERE engid = @engid "
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " INSERT INTO ID_TRANSLATION (engid ,cname) "
        sql &= " VALUES(@engid ,@cname) "
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = " SELECT * FROM ID_TRANSLATION WHERE engid = @engid "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("engid", SqlDbType.VarChar).Value = txtAEng.Text
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            '修改(加新增的字上去)
            strWord = dt.Rows(0)("cname")
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("cname", SqlDbType.VarChar).Value = strWord & txtAWord.Text
                .Parameters.Add("engid", SqlDbType.VarChar).Value = txtAEng.Text
                .ExecuteNonQuery()
            End With
            Common.MessageBox(Me, "修改完成!")
        Else
            '新增(新增一英文發音)
            strWord = txtAWord.Text
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("engid", SqlDbType.VarChar).Value = txtAEng.Text
                .Parameters.Add("cname", SqlDbType.VarChar).Value = strWord
                .ExecuteNonQuery()
            End With
            Common.MessageBox(Me, "新增完成!")
        End If

        Call goBack1()
    End Sub

    Sub goBack1()
        txtAEng.Text = ""
        txtAWord.Text = ""
        panAdd.Visible = False
        panSch.Visible = True
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Call goBack1()
    End Sub
End Class