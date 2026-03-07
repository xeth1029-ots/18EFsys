Partial Class CP_06_001_ADD
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

        '分頁設定---------------Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定---------------End
        PageControler1.Visible = False

        'Dim da As SqlDataAdapter = nothing
        'Dim dt As DataTable
        'Dim Trans As SqlTransaction
        'Dim sql As String
        'Dim dr As DataRow
        'Dim conn As SqlConnection = DbAccess.GetConnection

        If Not IsPostBack Then
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing

            If Not Session("Add_SearchStr") Is Nothing Then
                Dim UseingStr As String

                QuestNum.Text = TIMS.GetMyValue(Session("Add_SearchStr"), "QuestNum")
                QuestName.Text = TIMS.GetMyValue(Session("Add_SearchStr"), "QuestName")
                PathAttach.SelectedValue = TIMS.GetMyValue(Session("Add_SearchStr"), "PathAttach")
                UseingStr = TIMS.GetMyValue(Session("Add_SearchStr"), "Useing")
                If UseingStr = "Y" Then Useing.Checked = True Else Useing.Checked = False
                Order.SelectedValue = TIMS.GetMyValue(Session("Add_SearchStr"), "Order")

                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("Add_SearchStr"), "PageIndex")
                Dim MyValue As String = TIMS.GetMyValue(Session("Add_SearchStr"), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If

                If TIMS.GetMyValue(Session("Add_SearchStr"), "submit") = "1" Then
                    Query_Click(sender, e)
                End If

                Session("Add_SearchStr") = Nothing
            End If

            If Request("OGQID") <> "" Then
                Dim dr As DataRow
                Dim sql As String
                sql = "select * from Org_GradedQuest where OGQID='" & Request("OGQID") & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    QuestNum.Text = dr("QuestNum").ToString
                    QuestName.Text = dr("QuestName").ToString
                    Common.SetListItem(PathAttach, dr("PathAttach"))
                    If dr("Useing").ToString = "Y" Then Useing.Checked = True Else Useing.Checked = False
                    If dr("Used").ToString <> "" Then
                        QuestNum.Enabled = False
                        QuestName.Enabled = False
                        PathAttach.Enabled = False
                        Useing.Enabled = False
                    End If
                End If
            End If

            Select Case UCase(Request("Process"))
                Case "VIEW"
                    Query.Visible = False
                    Save.Visible = False
                    Panel1.Visible = False
                    Q_Set.Visible = False
                Case "ADD"
                    Query.Visible = False
                    Panel1.Visible = False
                    Q_Set.Visible = False

                    Save.Visible = True
                Case "UPDATE"
                    Query.Visible = True
                    Panel1.Visible = True
                    Q_Set.Visible = True

                    Save.Visible = True
                Case Else
                    Query.Visible = True
                    Panel1.Visible = True
                    Q_Set.Visible = True

                    Save.Visible = False
            End Select
        End If

        Q_Set.Attributes("onclick") = "return ChkDataQ();"
        Save.Attributes("onclick") = "return ChkData();"
        Query.Attributes("onclick") = "return ChkData();"
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        'Dim conn As SqlConnection = DbAccess.GetConnection
        'Dim Trans As SqlTransaction
        ''Try
        'Trans = DbAccess.BeginTrans(conn)

        Select Case UCase(Request("Process"))
            Case "ADD"
                sql = "select 'x' from Org_GradedQuest where QuestNum='" & QuestNum.Text & "' "
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "問卷代號不可重複")
                    Exit Sub
                End If
            Case "UPDATE"
                sql = "select 'x' from Org_GradedQuest where QuestNum='" & QuestNum.Text & "' AND OGQID<>'" & Request("OGQID") & "' "
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "問卷代號不可重複")
                    Exit Sub
                End If
        End Select

        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
        Try
            Select Case UCase(Request("Process"))
                Case "ADD"
                    sql = "select * from Org_GradedQuest WHERE 1<>1"
                    dt = DbAccess.GetDataTable(sql, da, oTrans)
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("CreateDate") = Now()
                    dr("CreateAcct") = sm.UserInfo.UserID
                Case "UPDATE"
                    sql = "select * from Org_GradedQuest where OGQID='" & Request("OGQID") & "' "
                    dt = DbAccess.GetDataTable(sql, da, oTrans)
                    dr = dt.Rows(0)
            End Select

            dr("QuestNum") = QuestNum.Text
            dr("QuestName") = QuestName.Text
            dr("PathAttach") = PathAttach.SelectedValue
            If Useing.Checked Then dr("Useing") = "Y" Else dr("Useing") = "N"
            dr("RID") = Request("RID")
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now()

            DbAccess.UpdateDataTable(dt, da, oTrans)
            DbAccess.CommitTrans(oTrans)
        Catch ex As Exception
            DbAccess.RollbackTrans(oTrans)
            Exit Sub
        End Try
        Call TIMS.CloseDbConn(oConn)

        Session("_SearchStr") = Me.ViewState("_SearchStr")
        Common.RespWrite(Me, "<script> ")
        Common.RespWrite(Me, " alert('儲存成功');")
        Common.RespWrite(Me, "location.href='CP_06_001.aspx?ID=" & Request("ID") & "'")
        Common.RespWrite(Me, "</script> ")

        'Common.MessageBox(Me, "儲存成功")
        'Common.RespWrite(Me, "<SCRIPT>if(confirm('是否要「題目設定」'))location.href='../06/CP_06_001_detail_add.aspx?RID=" & Request("RID") & "&QuestNum=" & QuestNum.Text & "&Process=add' ;else location.href='../06/CP_06_001.aspx'</SCRIPT>")
        'Catch ex As Exception
        '    DbAccess.RollbackTrans(Trans)
        '    Throw ex
        'End Try
    End Sub

    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Query.Click
        Call create()
    End Sub

    Sub create(Optional ByVal sql As String = "")
        Dim dt As DataTable
        If sql = "" Then
            sql = "" & vbCrLf
            sql += " select b.Path,b.Heading,b.Seq,b.OGHID, b.OGQID,b.Used " & vbCrLf
            sql += " from Org_GradedQuest a" & vbCrLf
            sql += " join Org_GradedHeading b on a.OGQID=b.OGQID" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and a.QuestNum='" & QuestNum.Text & "' " & vbCrLf
            sql += " and a.QuestName='" & QuestName.Text & "'" & vbCrLf
            sql += " and a.PathAttach='" & PathAttach.SelectedValue & "'" & vbCrLf
            If Useing.Checked = True Then
                sql += " and a.Useing='Y' "
            Else
                sql += " and a.Useing='N' "
            End If
        End If
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGrid1.Visible = False
        PageControler1.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            PageControler1.Visible = True
            msg.Text = ""

            'PageControler1.SqlString = sql
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OGHID"

            If Order.SelectedValue = 1 Then
                PageControler1.Sort = "Seq"
            Else
                PageControler1.Sort = "Heading"
            End If
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                GetSearchStr()
                Session("_SearchStr") = Me.ViewState("_SearchStr")

                Dim url1 As String = "CP_06_001_detail_add.aspx?Process=update&OGHID=" & e.CommandArgument & "&OGQID=" & Me.ViewState("OGQID") & "&ID=" & Request("ID")
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                Dim sql As String

                sql = "delete Org_GradedHeading where OGHID='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功！")

                Call create()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim drv As DataRowView
        drv = e.Item.DataItem
        Dim btn_edit As Button = e.Item.FindControl("Edit")
        Dim btn_del As Button = e.Item.FindControl("Del")

        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            btn_edit.CommandArgument = drv("OGHID")
            btn_del.CommandArgument = drv("OGHID")
            Me.ViewState("OGQID") = e.Item.Cells(4).Text

            If drv("Used").ToString = "Y" Then
                btn_edit.Enabled = False
                btn_del.Enabled = False
            Else
                btn_edit.Enabled = True
                btn_del.Enabled = True
            End If

            btn_del.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"
        End If
    End Sub

    Private Sub Q_Set_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Q_Set.Click
        Dim sql As String
        Dim dr As DataRow

        sql = "select * from Org_GradedQuest where QuestNum='" & QuestNum.Text & "' "
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, "無此問卷代號")
            Exit Sub
        End If

        GetSearchStr()
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        Dim url1 As String = "CP_06_001_detail_add.aspx?Process=add&QuestNum=" & QuestNum.Text & "&RID=" & dr("RID") & "&OGQID=" & dr("OGQID") & "&ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub

    Private Sub return_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles return_btn.Click
        'Common.RespWrite(Me, "<script> ")
        'Common.RespWrite(Me, " alert('儲存成功');")
        'Common.RespWrite(Me, "location.href='CP_06_001.aspx?ID=" & Request("ID") & "'")
        'Common.RespWrite(Me, "</script> ")
        Session("_SearchStr") = Me.ViewState("_SearchStr")

        Dim url1 As String = TIMS.GetFunIDUrl(Request("ID"), 1, objconn)
        Call TIMS.Utl_Redirect(Me, objconn, url1)

        'call TIMS.Utl_Redirect(Me, objconn,url & "?ID=" & Request("ID"))
    End Sub

    Sub GetSearchStr()
        Dim SearchStr1 As String = ""
        TIMS.SetMyValue(SearchStr1, "QuestNum", QuestNum.Text)
        TIMS.SetMyValue(SearchStr1, "QuestName", QuestName.Text)
        TIMS.SetMyValue(SearchStr1, "PathAttach", PathAttach.SelectedValue)

        Dim UseingStr As String = "N"
        If Useing.Checked Then UseingStr = "Y"
        TIMS.SetMyValue(SearchStr1, "Useing", UseingStr)
        TIMS.SetMyValue(SearchStr1, "Order", Order.SelectedValue)
        TIMS.SetMyValue(SearchStr1, "PageIndex", CStr(DataGrid1.CurrentPageIndex + 1))

        Dim submit1 As String = "0"
        If DataGrid1.Visible Then submit1 = "1"
        TIMS.SetMyValue(SearchStr1, "submit", submit1)

        Session("Add_SearchStr") = SearchStr1
    End Sub
End Class
