Partial Class CP_06_001_detail_add
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            Me.ViewState("Add_SearchStr") = Session("Add_SearchStr")
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("Add_SearchStr") = Nothing
            Session("_SearchStr") = Nothing

            create()
        End If

        If Path.SelectedValue = 1 Then
            PanelAns.Visible = False
            PathOGHID_List.Visible = False
        Else
            PanelAns.Visible = True
            PathOGHID_List.Visible = True
        End If

        AnsType.Attributes("onclick") = "ChangeAnsType();"
        Multiline.Attributes("onclick") = "ChangeMultiline();"
        Save.Attributes("onclick") = "return ChkData();"
    End Sub

    Sub create()
        Dim dt As DataTable
        Dim sql As String
        Dim dr As DataRow

        PanelAns.Visible = False
        MultilineTable.Style("display") = "none"
        Table_Multiline.Style("display") = "none"
        Table_AnsChooseN.Style("display") = "none"
        Table_AnsList.Style("display") = "none"

        If Request("OGQID") <> "" Then
            sql = "select '('+CONVERT(varchar, Seq)+')'+Heading as Title,OGHID from Org_GradedHeading where Path=1 and OGQID ='" & Request("OGQID") & "' "
            dt = DbAccess.GetDataTable(sql)

            If dt.Rows.Count > 0 Then
                With PathOGHID_List
                    .DataSource = dt
                    .DataTextField = "Title"
                    .DataValueField = "OGHID"
                    .DataBind()
                End With
            End If

            PathOGHID_List.Items.Insert(0, New ListItem("===請選擇===", ""))
        End If

        If Request("OGHID") <> "" Then
            sql = "select * from Org_GradedHeading where OGHID='" & Request("OGHID") & "' "
            dr = DbAccess.GetOneRow(sql)

            If Not dr Is Nothing Then
                'If dr("PathOGHID").ToString <> "" Then
                '    PathOGHID_List.SelectedValue = dr("PathOGHID")
                'End If

                Common.SetListItem(Path, dr("Path"))
                Common.SetListItem(PathOGHID_List, dr("PathOGHID"))
                Common.SetListItem(AnsType, dr("AnsType"))

                Heading.Text = dr("Heading")
                Seq.Text = dr("Seq")

                If dr("Useing") = "Y" Then Useing.Checked = True Else Useing.Checked = False

                If dr("AnsType").ToString = "" Then
                    AnsType.SelectedIndex = -1
                Else
                    AnsType.SelectedValue = dr("AnsType").ToString
                End If

                If dr("Multiline").ToString = "" Then
                    Multiline.SelectedIndex = -1
                    MultilineTable.Style("display") = "none"
                Else
                    Multiline.SelectedValue = dr("Multiline").ToString
                    If Multiline.SelectedValue = "N" Then
                        MultilineTable.Style("display") = "none"
                    Else
                        MultilineTable.Style("display") = "inline"
                    End If
                End If

                If dr("Rows").ToString <> "" Then Rows.Text = dr("Rows") Else Rows.Text = ""
                If dr("AnsChooseN").ToString <> "" Then AnsChooseN.Text = dr("AnsChooseN") Else AnsChooseN.Text = ""
                If dr("AnsList").ToString <> "" Then AnsList.Text = dr("AnsList") Else AnsList.Text = ""

                Select Case dr("AnsType").ToString
                    Case "01"
                        PanelAns.Visible = True
                        Table_Multiline.Style("display") = "none"
                        Table_AnsChooseN.Style("display") = "none"
                        Table_AnsList.Style("display") = "inline"
                    Case "02"
                        PanelAns.Visible = True
                        Table_Multiline.Style("display") = "none"
                        Table_AnsChooseN.Style("display") = "inline"
                        Table_AnsList.Style("display") = "inline"
                    Case "03"
                        PanelAns.Visible = True
                        Table_Multiline.Style("display") = "inline"
                        Table_AnsChooseN.Style("display") = "inline"
                        Table_AnsList.Style("display") = "none"
                End Select
            End If
        End If
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        'Dim da As SqlDataAdapter = Nothing
        'Dim dt As DataTable
        'Dim Trans As SqlTransaction
        'Dim sql As String

        Dim conn As SqlConnection = DbAccess.GetConnection

        If AnsList.Text <> "" Then
            If ChkAnsList(AnsList.Text) = False Then
                Select Case AnsType.SelectedValue
                    Case "01"
                        PanelAns.Visible = True
                        Table_Multiline.Style("display") = "none"
                        Table_AnsChooseN.Style("display") = "none"
                        Table_AnsList.Style("display") = "inline"
                    Case "02"
                        PanelAns.Visible = True
                        Table_Multiline.Style("display") = "none"
                        Table_AnsChooseN.Style("display") = "inline"
                        Table_AnsList.Style("display") = "inline"
                    Case "03"
                        PanelAns.Visible = True
                        Table_Multiline.Style("display") = "inline"
                        Table_AnsChooseN.Style("display") = "inline"
                        Table_AnsList.Style("display") = "none"
                End Select
                Exit Sub
            End If
        End If

        Dim sql As String = ""
        Select Case UCase(Request("Process"))
            Case "ADD"
                sql = "select * from Org_GradedHeading where OGQID='" & Request("OGQID") & "' "
                sql += "and path='" & Path.SelectedValue & "' "
                If Seq.Text <> "" Then
                    sql += "and Seq='" & Seq.Text & "' "
                End If
                Dim dt As DataTable = Nothing
                dt = DbAccess.GetDataTable(sql, conn)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "排序重覆")
                    Exit Sub
                End If
        End Select

        Dim Trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            Dim dr As DataRow = Nothing

            Select Case UCase(Request("Process"))
                Case "ADD"
                    sql = "select * from Org_GradedHeading where 1<>1 "
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    dr = dt.NewRow
                    dt.Rows.Add(dr)

                    dr("OGQID") = Request("OGQID")
                    dr("RID") = Request("RID")
                    dr("CreateDate") = Now()
                    dr("CreateAcct") = sm.UserInfo.UserID
                Case "UPDATE"
                    sql = "select * from Org_GradedHeading where OGHID='" & Request("OGHID") & "' "
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    dr = dt.Rows(0)
            End Select

            dr("Heading") = Heading.Text
            dr("Path") = Path.SelectedValue
            dr("Seq") = Seq.Text

            If Useing.Checked = True Then dr("Useing") = "Y" Else dr("Useing") = "N"

            If AnsType.SelectedIndex = -1 Then
                dr("AnsType") = ""
            Else
                dr("AnsType") = AnsType.SelectedValue
            End If

            If Multiline.SelectedIndex = -1 Then
                dr("Multiline") = Convert.DBNull
            Else
                dr("Multiline") = Multiline.SelectedValue
                If dr("Multiline") = "N" Then
                    dr("Rows") = Convert.DBNull
                Else
                    dr("Rows") = IIf(Rows.Text = "", Convert.DBNull, Rows.Text)
                End If
            End If

            If Me.PathOGHID_List.SelectedValue <> "" Then
                dr("PathOGHID") = Me.PathOGHID_List.SelectedValue
            End If

            dr("AnsChooseN") = IIf(AnsChooseN.Text = "", Convert.DBNull, AnsChooseN.Text)
            dr("AnsList") = IIf(AnsList.Text = "", Convert.DBNull, AnsList.Text)
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now()

            If Path.SelectedValue = "1" Then
                dr("AnsType") = ""
                dr("Multiline") = Convert.DBNull
                dr("Rows") = Convert.DBNull
                dr("AnsList") = Convert.DBNull
                dr("AnsChooseN") = Convert.DBNull
            End If

            Select Case AnsType.SelectedValue
                Case "01"
                    dr("Multiline") = Convert.DBNull
                    dr("Rows") = Convert.DBNull
                    dr("AnsChooseN") = Convert.DBNull
                Case "02"
                    dr("Multiline") = Convert.DBNull
                    dr("Rows") = Convert.DBNull
                Case "03"
                    dr("AnsList") = Convert.DBNull
            End Select

            DbAccess.UpdateDataTable(dt, da, Trans)
            DbAccess.CommitTrans(Trans)

            Session("Add_SearchStr") = Me.ViewState("Add_SearchStr")
            Session("_SearchStr") = Me.ViewState("_SearchStr")
            Common.RespWrite(Me, "<script> ")
            Common.RespWrite(Me, " alert('儲存成功');")
            Common.RespWrite(Me, "location.href='CP_06_001_add.aspx?ID=" & Request("ID") & "'")
            Common.RespWrite(Me, "</script> ")
        Catch ex As Exception
            DbAccess.RollbackTrans(Trans)
            Call TIMS.CloseDbConn(conn)
            Throw ex
        End Try
        Call TIMS.CloseDbConn(conn)
    End Sub

    Function ChkAnsList(ByVal AList As String) As Boolean
        Dim myArray As Array = Split(AList, ",")
        Dim i, j As Integer
        j = 0

        Select Case Me.AnsType.SelectedValue
            Case "01"
                For i = 0 To myArray.Length - 1
                    If myArray(i) = "" Then j += 1
                Next

                If j > 0 Then
                    Common.RespWrite(Me, "<script> ")
                    Common.RespWrite(Me, " alert('答案提示不可為空白');")
                    Common.RespWrite(Me, "</script> ")
                    Return False
                ElseIf myArray.Length < 2 Then
                    Common.RespWrite(Me, "<script> ")
                    Common.RespWrite(Me, " alert('答案列示至少要1個半型逗點');")
                    Common.RespWrite(Me, "</script> ")
                    Return False
                Else
                    Return True
                End If
            Case "02"
                If CInt(AnsChooseN.Text) > myArray.Length Then
                    Common.RespWrite(Me, "<script> ")
                    Common.RespWrite(Me, " alert('複選答案數量限制(問答長度)不可大於答案列示的長度');")
                    Common.RespWrite(Me, "</script> ")
                    Return False
                Else
                    Return True
                End If
            Case Else
                Return True
        End Select
    End Function

    Private Sub return_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles return_btn.Click
        Session("Add_SearchStr") = Me.ViewState("Add_SearchStr")
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        'Response.Redirect("CP_06_001_add.aspx?ID=" & Request("ID"))
        Dim url1 As String = "CP_06_001_add.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub
End Class
