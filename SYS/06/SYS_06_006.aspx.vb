Partial Class SYS_06_006
    Inherits AuthBasePage

    'AUTH_PASSWORD
    Const cst_PlanYears As String = "PlanYears"
    Const cst_DistID As String = "DistID"
    Const cst_PlanID As String = "PlanID"

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
            Me.ViewState("years") = sm.UserInfo.Years
            Me.ViewState("distid") = sm.UserInfo.DistID

            Show_DropDownList(cst_PlanYears, list_Years, "Years", "Years")
            Show_DropDownList(cst_DistID, list_DistID, "Name", "DistID")
            Show_DropDownList(cst_PlanID, list_PlanID, "PlanName", "PlanID")

            'list_Years.SelectedValue = sm.UserInfo.Years
            'list_DistID.SelectedValue = sm.UserInfo.DistID
            'list_PlanID.SelectedValue = sm.UserInfo.PlanID
            Common.SetListItem(list_Years, sm.UserInfo.Years)
            Common.SetListItem(list_DistID, sm.UserInfo.DistID)
            Common.SetListItem(list_PlanID, sm.UserInfo.PlanID)

            Dim v_list_PlanID As String = TIMS.GetListValue(list_PlanID)
            If v_list_PlanID <> "" Then
                Call Show_list_PlanID(v_list_PlanID)
            End If
        End If

    End Sub

    Private Sub list_Years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_Years.SelectedIndexChanged
        '依年度 轄區 顯示可用計畫
        Call Show_list_YearsDistID()
    End Sub

    Private Sub list_DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_DistID.SelectedIndexChanged
        '依年度 轄區 顯示可用計畫
        Call Show_list_YearsDistID()
    End Sub

    '依年度 轄區 顯示可用計畫
    Sub Show_list_YearsDistID()
        Me.ViewState("years") = list_Years.SelectedValue
        Me.ViewState("distid") = list_DistID.SelectedValue

        '依年度 轄區 顯示可用計畫
        Show_DropDownList(cst_PlanID, list_PlanID, "PlanName", "PlanID")

        If list_PlanID.SelectedValue <> "" Then
            Call Show_list_PlanID(list_PlanID.SelectedValue)
        End If
    End Sub

    '顯示DropDownList資料
    Private Sub Show_DropDownList(ByVal strFlag As String, ByVal objDDL As DropDownList, ByVal textField As String, ByVal valueField As String)
        Me.tPXssXArd1.Text = ""
        Me.tPXssXArd2.Text = ""
        msg.Text = "密碼尚未設定!!"
        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try

        Dim sql As String = ""
        Select Case strFlag
            Case cst_PlanYears
                sql = "select distinct Years from ID_Plan WHERE Years!=' ' order by Years"

            Case cst_DistID
                sql = "select DistID,Name from ID_District order by DistID Asc "

            Case cst_PlanID
                sql = "" & vbCrLf
                sql += " select distinct a.PlanID" & vbCrLf
                sql += " ,a.Years+c.Name+b.PlanName+a.Seq" & vbCrLf
                sql += " +case when dbo.NVL(a.SubTitle,' ')!=' ' then '(' + a.SubTitle + ')' end PlanName" & vbCrLf
                sql += " ,dbo.FN_GET_PXSSXARD(a.PLANID,1) PXssXArd" & vbCrLf
                sql += " ,dbo.FN_GET_PXSSXARD(a.PLANID,2) HASHPWD1" & vbCrLf
                sql += " FROM ID_PLAN a" & vbCrLf
                sql += " join Key_Plan b on b.TPlanID=a.TPlanID" & vbCrLf
                sql += " join ID_District c on c.DistID=a.DistID" & vbCrLf
                sql += " where a.Years='" & Convert.ToString(Me.ViewState("years")) & "'" & vbCrLf
                sql += " and a.DistID='" & Convert.ToString(Me.ViewState("distid")) & "'" & vbCrLf
                sql += " order by a.PlanID asc" & vbCrLf

        End Select

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        objDDL.Items.Clear()
        If dt.Rows.Count > 0 Then
            objDDL.DataSource = dt
            objDDL.DataTextField = textField
            objDDL.DataValueField = valueField
            objDDL.DataBind()
        End If
    End Sub

    Sub Show_list_PlanID(ByVal planid As String)
        Me.tPXssXArd1.Text = ""
        Me.tPXssXArd2.Text = ""
        msg.Text = "密碼尚未設定!!"

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select distinct a.PlanID" & vbCrLf
        sql += " ,a.Years+c.Name+b.PlanName+a.Seq" & vbCrLf
        sql += " +case when dbo.NVL(a.SubTitle,' ')!=' ' then '(' + a.SubTitle + ')' end PlanName" & vbCrLf
        sql += " ,dbo.FN_GET_PXSSXARD(a.PLANID,1) PXssXArd" & vbCrLf
        sql += " ,dbo.FN_GET_PXSSXARD(a.PLANID,2) HASHPWD1" & vbCrLf
        sql += " from ID_Plan a " & vbCrLf
        sql += " JOIN KEY_PLAN b on b.TPlanID=a.TPlanID " & vbCrLf
        sql += " JOIN ID_DISTRICT c on c.DistID=a.DistID " & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and a.Years=" & Convert.ToString(Me.ViewState("years")) & " " & vbCrLf
        sql += " and a.DistID='" & Convert.ToString(Me.ViewState("distid")) & "' " & vbCrLf
        sql += " and a.PlanID ='" & planid & "' " & vbCrLf
        'sql += "  order by a.PlanID asc " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            hidPXwSd.Value = Convert.ToString(dr("PXssXArd"))
            If hidPXwSd.Value <> "" Then
                Me.tPXssXArd1.Text = "********"
                Me.tPXssXArd2.Text = "********"

                Dim vs_txtpass As String = hidPXwSd.Value
                Dim str_txtpass As String = ""
                str_txtpass = "tPXssXArd1"
                Me.RegisterStartupScript("key_" & str_txtpass, "<script>if (document.getElementById('" & str_txtpass & "')) { if (document.getElementById('" & str_txtpass & "').value=='') {document.getElementById('" & str_txtpass & "').value = '" & vs_txtpass & "';}}</script>")
                str_txtpass = "tPXssXArd2"
                Me.RegisterStartupScript("key_" & str_txtpass, "<script>if (document.getElementById('" & str_txtpass & "')) { if (document.getElementById('" & str_txtpass & "').value=='') {document.getElementById('" & str_txtpass & "').value = '" & vs_txtpass & "';}}</script>")
                msg.Text = "密碼已經設定!!"
            End If
        End If

    End Sub

    Function sUtl_Savedata() As Integer
        Dim rst As Integer = 0 '儲存異常

        Dim v_list_PlanID As String = TIMS.GetListValue(list_PlanID)

        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " DELETE AUTH_PASSWORD " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        sql += " AND PlanID= @PlanID" & vbCrLf
        Dim parms As New Hashtable
        parms.Add("PlanID", v_list_PlanID)
        DbAccess.ExecuteNonQuery(sql, objconn, parms)

        sql = "" & vbCrLf
        'sql += "  /* IDENTITY(1,1):  */ " & vbCrLf
        sql += " INSERT INTO AUTH_PASSWORD(" & vbCrLf
        sql += " PlanID" & vbCrLf
        'sql &= " ,Password" & vbCrLf
        sql &= " ,HASHPWD1" & vbCrLf
        sql &= " ,ModifyAcct" & vbCrLf
        sql &= " ,ModifyDate" & vbCrLf
        sql += " ) VALUES (" & vbCrLf
        sql += " @PlanID" & vbCrLf
        'sql &= " ,@Password" & vbCrLf
        sql &= " ,@HASHPWD1" & vbCrLf
        sql &= " ,@ModifyAcct" & vbCrLf
        sql &= " ,GETDATE()" & vbCrLf
        sql += " ) " & vbCrLf
        Dim i_parms As New Hashtable
        i_parms.Add("PlanID", v_list_PlanID)
        i_parms.Add("PASSWORD", tPXssXArd1.Text)
        i_parms.Add("HASHPWD1", TIMS.CreateHash(tPXssXArd1.Text))
        i_parms.Add("ModifyAcct", sm.UserInfo.UserID)
        rst = DbAccess.ExecuteNonQuery(sql, objconn, i_parms)

        'Try
        '    da.SelectCommand.CommandText = sql
        '    da.SelectCommand.Parameters.Clear()
        '    da.SelectCommand.Parameters.Add("PlanID", SqlDbType.VarChar).Value = v_list_PlanID 'list_PlanID.SelectedValue
        '    'da.SelectCommand.Parameters.Add("Password", SqlDbType.VarChar).Value = TIMS.ClearSQM(tPXssXArd1.Text)
        '    da.SelectCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
        '    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
        '    da.SelectCommand.ExecuteNonQuery()
        '    If da.SelectCommand.Connection.State = ConnectionState.Open Then da.SelectCommand.Connection.Close()
        '    rst = 1 '儲存正常
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
        Return rst
    End Function

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Dim sErrmsg As String = ""
        sErrmsg = ""
        Dim strPXssXArd1 As String = tPXssXArd1.Text
        tPXssXArd1.Text = TIMS.ClearSQM(tPXssXArd1.Text)
        tPXssXArd2.Text = TIMS.ClearSQM(tPXssXArd2.Text)
        If strPXssXArd1 <> tPXssXArd1.Text Then
            sErrmsg += "密碼" & TIMS.cst_ErrorMsg10 & vbCrLf
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        If Me.tPXssXArd1.Text = "" Then
            sErrmsg += "密碼不可為空白" & vbCrLf
        End If
        If Me.tPXssXArd2.Text = "" Then
            sErrmsg += "重key密碼不可為空白" & vbCrLf
        End If

        If sErrmsg = "" Then
            If Me.tPXssXArd1.Text <> Me.tPXssXArd2.Text Then
                sErrmsg += "密碼與重key密碼不合" & vbCrLf
            End If
        End If

        If sErrmsg = "" Then
            If tPXssXArd1.Text = hidPXwSd.Value Then
                sErrmsg += "密碼與舊密碼相同" & vbCrLf
            End If
        End If
        If sErrmsg = "" Then
            '密碼設計長度過短
            If Not TIMS.CheckPxssword(tPXssXArd1.Text) Then
                sErrmsg += "* 為保障個人帳號資料安全，密碼必須為 " & vbCrLf
                sErrmsg += "2位英文字母與2位阿拉伯數字以上的組合，謝謝!" & vbCrLf
            End If
        End If
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        If sUtl_Savedata() = 1 Then
            Common.MessageBox(Me, "設定完成!!")
            Dim v_list_PlanID As String = TIMS.GetListValue(list_PlanID)
            Me.ViewState("planid") = v_list_PlanID 'list_PlanID.SelectedValue
            '依計畫 顯示可用密碼
            Call Show_list_PlanID(Me.ViewState("planid"))
        End If
    End Sub

    Private Sub list_PlanID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_PlanID.SelectedIndexChanged
        Dim v_list_PlanID As String = TIMS.GetListValue(list_PlanID)
        Me.ViewState("planid") = v_list_PlanID 'list_PlanID.SelectedValue

        '依計畫 顯示可用密碼
        Call Show_list_PlanID(Me.ViewState("planid"))
    End Sub

End Class
