Partial Class OB_01_002_add
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then

            Me.ViewState("Action") = UCase(Request("Action"))

            Dim sql As String = ""
            Dim dr As DataRow = Nothing

            If Request("MSN") <> "" And Me.ViewState("Action") = "EDIT" Then

                sql = ""
                sql += " SELECT b.OrgName, a.* " & vbCrLf
                sql += " FROM OB_Member a " & vbCrLf
                sql += " LEFT JOIN OB_Org b on a.OrgSN=b.OrgSN " & vbCrLf
                sql += " WHERE MSN='" & Request("MSN") & "' " & vbCrLf
                If sm.UserInfo.DistID <> "000" Then
                    sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                End If

                dr = DbAccess.GetOneRow(sql, objconn)
                Me.ViewState("MSN") = Request("MSN")
                Me.ViewState("MSN") = TIMS.ClearSQM(Me.ViewState("MSN"))

                Me.center.Text = dr("OrgName").ToString
                Me.orgid_value.Value = dr("OrgID").ToString

                If Me.orgid_value.Value <> "" Then
                    rb1.Checked = True
                Else
                    rb2.Checked = True
                End If

                Me.DeptName.Text = dr("DeptName").ToString
                Me.memName.Text = dr("memName").ToString

                If dr("Qualified").ToString <> "" Then
                    Common.SetListItem(Me.rblQualified, dr("Qualified"))
                End If

                'If dr("Qualified").ToString = "Y" Then
                '    radio1.Checked = True
                'End If
                'If dr("Qualified").ToString = "N" Then
                '    radio2.Checked = True
                'End If
            End If

        End If

        Me.ViewState("center") = Me.center.Text
        Me.ViewState("orgid_value") = Me.orgid_value.Value
        PageLoadSetLast1()

    End Sub

    Sub PageLoadSetLast1()
        rb1.Attributes("onclick") = "set_Orgname1('rb1');"
        rb2.Attributes("onclick") = "set_Orgname1('rb2');"

        If rb1.Checked Then
            Page.RegisterStartupScript("rb_checked", "<script>set_Orgname1('rb1');</script>")
        End If
        If rb2.Checked Then
            Page.RegisterStartupScript("rb_checked", "<script>set_Orgname1('rb2');</script>")
        End If

        '因有傳入值 yearlist.SelectedValue.ToString 故放此位置，才可讀到值
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Org.Attributes("onclick") += "javascript:rb1_checked();"

        'btnSave.Attributes("onclick") = "return CheckData1();"
        If Me.ViewState("center") <> "" Then
            Dim strscript As String = "" & vbCrLf
            strscript = "<script language='javascript'>" & vbCrLf
            strscript += "document.all('center').value='" & Me.ViewState("center") & "';" & vbCrLf
            strscript += "document.all('orgid_value').value='" & Me.ViewState("orgid_value") & "';" & vbCrLf
            strscript += "</script>" & vbCrLf
            Me.ViewState("center") = Nothing
            Me.ViewState("orgid_value") = Nothing
            Page.RegisterStartupScript("test1", strscript)
        End If

    End Sub

    Function CheckData(ByRef Errmag As String) As Boolean
        Errmag = ""
        CheckData = False

        If Not (rb1.Checked Or rb2.Checked) Then
            Errmag += "請選擇服務單位型態" & vbCrLf
        End If

        If rb1.Checked Then
            If orgid_value.Value = "" Then
                Errmag += "請選擇服務單位" & vbCrLf
            End If
        End If

        If rb2.Checked Then
            If center.Text.Trim = "" Then
                Errmag += "請輸入服務單位名稱" & vbCrLf
            End If
            center.Text = center.Text.Trim
        End If

        'DeptName
        If DeptName.Text.Trim = "" Then
            Errmag += "請輸入服務部門" & vbCrLf
        End If
        DeptName.Text = DeptName.Text.Trim

        If memName.Text.Trim = "" Then
            Errmag += "請輸入成員姓名" & vbCrLf
        End If
        memName.Text = memName.Text.Trim

        If Me.rblQualified.SelectedValue = "" Then
            Errmag += "請選擇具備採購法證照" & vbCrLf
        End If
        'If Not (radio1.Checked Or radio2.Checked) Then
        '    Errmag += "請選擇具備採購法證照" & vbCrLf
        'End If

        'Select Case Me.ViewState("Action")     2009/06/29 拿掉
        '    Case "ADD"
        '        If Chk_Double_OB_Member(center.Text, DeptName.Text) Then
        '            Errmag += "此服務單位名稱、服務部門已經存在" & vbCrLf
        '        End If
        '    Case "EDIT"
        '        If Chk_Double_OB_Member(center.Text, DeptName.Text, Me.ViewState("MSN")) Then
        '            Errmag += "此服務單位名稱、服務部門已經存在" & vbCrLf
        '        End If
        'End Select

        If Errmag = "" Then
            CheckData = True
        End If

    End Function

    Function Chk_Double_OB_Member(ByVal OrgName As String, ByVal DeptName As String, Optional ByVal msn As Integer = 0) As Boolean
        Dim str_flag As String = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select b.OrgName, a.* " & vbCrLf
        sql += " from OB_Member a " & vbCrLf
        sql += " JOIN OB_Org b on a.OrgSN=b.OrgSN " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If
        sql += " AND b.OrgName='" & OrgName & "'" & vbCrLf
        sql += " AND a.DeptName='" & DeptName & "'" & vbCrLf
        If msn <> 0 Then
            '排除本身
            sql += " AND a.msn != " & msn & vbCrLf
        End If

        str_flag = DbAccess.ExecuteScalar(sql, objconn)
        If str_flag Is Nothing Then
            Chk_Double_OB_Member = False
        Else
            Chk_Double_OB_Member = True
        End If
    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim Errmsg As String = ""
        If Not CheckData(Errmsg) Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Select Case Me.ViewState("Action")
            Case "ADD"
                SAVE_Member()
            Case "EDIT"
                SAVE_Member(Me.ViewState("MSN"))
        End Select

    End Sub

    Sub SAVE_Member(Optional ByVal MSN As Integer = 0)
        'objconn = DbAccess.GetConnection
        Call TIMS.OpenDbConn(objconn)
        Dim vsOrgSN As Integer = 0
        vsOrgSN = Get_OBOrgSN(Me, Me.center.Text, objconn)
        If vsOrgSN = 0 Then
            Common.MessageBox(Me, "未輸入有效機構名稱!!")
            Exit Sub
        End If

        Dim sqlStr As String = ""
        Select Case Me.ViewState("Action")
            Case "ADD"
                sqlStr = "" & vbCrLf
                sqlStr += " INSERT INTO OB_Member(OrgID, OrgSN, DeptName, memName, Qualified " & vbCrLf
                sqlStr += " , DistID, CreateAcct, CreateTime, ModifyAcct, ModifyTime) " & vbCrLf
                sqlStr += " VALUES( @OrgID, @OrgSN , @DeptName , @memName, @Qualified  " & vbCrLf
                sqlStr += " , @DistID, @CreateAcct, getdate(), @ModifyAcct, getdate()) " & vbCrLf
                Dim iCmd As New SqlCommand(sqlStr, objconn)

                With iCmd
                    .Parameters.Clear()
                    If rb2.Checked Then
                        .Parameters.Add("OrgID", SqlDbType.Int).Value = Convert.DBNull
                    Else
                        .Parameters.Add("OrgID", SqlDbType.Int).Value = orgid_value.Value
                    End If
                    .Parameters.Add("OrgSN", SqlDbType.Decimal).Value = vsOrgSN
                    .Parameters.Add("DeptName", SqlDbType.NVarChar, 20).Value = DeptName.Text
                    .Parameters.Add("memName", SqlDbType.NVarChar, 20).Value = memName.Text
                    '.Parameters.Add("Qualified", SqlDbType.Char, 1).Value = IIf(radio1.Checked, "Y", "N")
                    .Parameters.Add("Qualified", SqlDbType.Char, 1).Value = Me.rblQualified.SelectedValue
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                    .Parameters.Add("DistID", SqlDbType.VarChar, 3).Value = sm.UserInfo.DistID
                    .Parameters.Add("CreateAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With

                Dim strScript As String
                strScript = "<script language=""javascript"">" & vbCrLf
                strScript += "alert('工作小組成員資料建檔-新增成功!!');" & vbCrLf
                strScript += "location.href='OB_01_002.aspx?ID=" & Request("ID") & "';" & vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript("", strScript)

            Case "EDIT"
                sqlStr = "" & vbCrLf
                sqlStr += " UPDATE OB_Member" & vbCrLf
                sqlStr += " SET  OrgID=@OrgID " & vbCrLf
                sqlStr += " , OrgSN=@OrgSN " & vbCrLf
                sqlStr += " , DeptName=@DeptName " & vbCrLf
                sqlStr += " , memName=@memName " & vbCrLf
                sqlStr += " , Qualified=@Qualified " & vbCrLf
                sqlStr += " , ModifyAcct=@ModifyAcct" & vbCrLf
                sqlStr += " , ModifyTime=getdate()" & vbCrLf
                sqlStr += " WHERE MSN=@MSN " & vbCrLf
                Dim uCmd As New SqlCommand(sqlStr, objconn)

                With uCmd
                    .Parameters.Clear()
                    If rb2.Checked Then
                        .Parameters.Add("OrgID", SqlDbType.Int).Value = Convert.DBNull
                    Else
                        .Parameters.Add("OrgID", SqlDbType.Int).Value = orgid_value.Value
                    End If
                    .Parameters.Add("OrgSN", SqlDbType.Decimal).Value = vsOrgSN
                    .Parameters.Add("DeptName", SqlDbType.NVarChar, 20).Value = DeptName.Text
                    .Parameters.Add("memName", SqlDbType.NVarChar, 20).Value = memName.Text
                    '.Parameters.Add("Qualified", SqlDbType.Char, 1).Value = IIf(radio1.Checked, "Y", "N")
                    .Parameters.Add("Qualified", SqlDbType.Char, 1).Value = Me.rblQualified.SelectedValue
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                    .Parameters.Add("MSN", SqlDbType.Decimal).Value = MSN
                    .ExecuteNonQuery()
                End With

                Dim strScript As String
                strScript = "<script language=""javascript"">" & vbCrLf
                strScript += "alert('工作小組成員資料建檔-修改成功!!');" & vbCrLf
                strScript += "location.href='OB_01_002.aspx?ID=" & Request("ID") & "';" & vbCrLf
                strScript += "</script>"
                Page.RegisterStartupScript("", strScript)

        End Select


    End Sub

    Public Shared Function Get_OBOrgSN(ByRef MyPage As Page, ByVal OrgName As String, ByVal tConn As SqlConnection) As Integer
        Dim rst As Integer = 0
        'Dim objCmd As SqlCommand
        Dim sqlStr As String
        Dim sm As SessionModel = SessionModel.Instance()

        OrgName = Trim(OrgName)
        If OrgName = "" Then Return rst '離開

        Dim dt As New DataTable
        Call TIMS.OpenDbConn(tConn)
        sqlStr = "SELECT OrgSN FROM OB_Org WHERE OrgName= @OrgName "
        Dim sCmd As New SqlCommand(sqlStr, tConn)

        sqlStr = "" & vbCrLf
        sqlStr += " INSERT INTO OB_Org(ORGSN, OrgName, CreateAcct, CreateTime)" & vbCrLf
        sqlStr += " VALUES(@ORGSN, @OrgName, @CreateAcct, getdate() )  " & vbCrLf
        Dim iCmd As New SqlCommand(sqlStr, tConn)

        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OrgName", SqlDbType.NVarChar).Value = OrgName
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            rst = dt.Rows(0)("OrgSN")
        Else
            Dim iORGSN As Integer = DbAccess.GetNewId(tConn, "OB_ORG_ORGSN_SEQ,OB_ORG,ORGSN")
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("ORGSN", SqlDbType.Int).Value = iORGSN
                .Parameters.Add("OrgName", SqlDbType.NVarChar, 40).Value = Trim(OrgName)
                .Parameters.Add("CreateAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                .ExecuteNonQuery()
            End With
            rst = iORGSN
        End If
        Return rst
    End Function

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        'Response.Redirect("OB_01_002.aspx?ID=" & Request("ID"))
        Dim url1 As String = "OB_01_002.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class
