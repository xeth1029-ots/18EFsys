Partial Class SYS_04_013_add
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

        ' Request("CalID") ==> Me.hidcalID.Value 
        If Request("CalID") <> "" Then Me.hidcalID.Value = Request("CalID")

        If Not IsPostBack Then
            btnSave1.Attributes.Add("OnClick", "return checkSave1();")

            Call ClearList()

            If Me.hidcalID.Value <> "" Then
                Call Create1(Me.hidcalID.Value)
            End If
        End If

    End Sub

    '清除
    Sub ClearList()
        Me.subject.Text = ""
        Me.OSDate.Text = ""
        Me.OFDate.Text = ""
        Me.txtcontext.Text = ""
        'CalID
    End Sub

    '取得 顯示
    Sub Create1(ByVal CalID As String)
        Dim dr As DataRow
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = ""
        sql += " select * from Auth_AccCal WHERE calID=" & Me.hidcalID.Value & vbCrLf
        sql += " and Account='" & sm.UserInfo.UserID & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)

            Me.subject.Text = Convert.ToString(dr("subject"))
            If Convert.ToString(dr("OSDate")) <> "" Then
                Me.OSDate.Text = CDate(dr("OSDate")).ToString("yyyy/MM/dd")
            End If
            If Convert.ToString(dr("OFDate")) <> "" Then
                Me.OFDate.Text = CDate(dr("OFDate")).ToString("yyyy/MM/dd")
            End If
            Me.txtcontext.Text = Convert.ToString(dr("context"))
        End If
    End Sub

    Private Sub back_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles back.ServerClick
        TIMS.Utl_Redirect1(Me, "SYS_04_013.aspx?ID=" & Request("ID") & "&s=1")
    End Sub

    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        Dim vsMsg As String = ""

        Dim saveOk As Boolean = False '儲存失敗!!!
        Dim sActionType As String = ""
        Const Cst_Insert As String = "Insert"
        Const Cst_Update As String = "Update" '修改
        sActionType = Cst_Update '修改
        If Me.hidcalID.Value = "" Then sActionType = Cst_Insert '新增 

        Try
            Dim da As SqlDataAdapter = Nothing
            Dim dr As DataRow = Nothing
            Dim dt As DataTable = Nothing
            Dim sql As String = ""
            Select Case sActionType
                Case Cst_Insert
                    sql = "select * from Auth_AccCal WHERE 1<>1 "
                    dt = DbAccess.GetDataTable(sql, da, objconn)
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("Account") = sm.UserInfo.UserID
                    dr("CreateDate") = Now
                    dr("ModifyDate") = Now
                Case Cst_Update
                    sql = ""
                    sql += " select * from Auth_AccCal WHERE calID=" & Me.hidcalID.Value & vbCrLf
                    sql += " and Account='" & sm.UserInfo.UserID & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(sql, da, objconn)
                    If dt.Rows.Count > 0 Then
                        dr = dt.Rows(0)
                        dr("Account") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    End If
            End Select
            If Not dr Is Nothing Then
                dr("subject") = Me.subject.Text
                dr("OSDate") = Common.FormatDate(Me.OSDate.Text)
                dr("OFDate") = Common.FormatDate(Me.OFDate.Text)
                dr("context") = Me.txtcontext.Text
                DbAccess.UpdateDataTable(dt, da)
            End If

            saveOk = True '儲存成功!!!
        Catch ex As Exception
            vsMsg = "!!儲存時發生錯誤!!"
            Common.MessageBox(Me, vsMsg)
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
        End Try

        If saveOk Then
            Common.RespWrite(Me, "<script> alert('儲存成功!');")
            Common.RespWrite(Me, "location.href='SYS_04_013.aspx?ID=" & Request("ID") & "&s=1';</script>")
        End If

    End Sub

End Class
