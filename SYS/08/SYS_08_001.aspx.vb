Partial Class SYS_08_001
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

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
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End
        '分頁設定 Start
        'PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'insert.Disabled = True
        'If blnCanAdds Then insert.Disabled = False
        'save.Disabled = True
        'If blnCanAdds Then save.Disabled = False
        'search.Disabled = True
        'If blnCanSech Then search.Disabled = False

        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else

        '            FunDr = FunDrArray(0)

        '            If FunDr("Adds") = "1" Then
        '                insert.Disabled = False
        '                save.Disabled = False
        '            Else
        '                insert.Disabled = True
        '                save.Disabled = True
        '            End If

        '            If FunDr("Sech") = "1" Then
        '                search.Disabled = False
        '            Else
        '                search.Disabled = True
        '                save.Disabled = True
        '            End If
        '        End If
        '    End If
        'End If


        If Not IsPostBack Then
            table_F.Visible = True
            table_I.Visible = False
            PageControler1.Visible = False
        End If

    End Sub

    Private Sub search_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.ServerClick, return1.ServerClick
        Call dt_search()
    End Sub

    Sub dt_search()
        'Dim str As String
        'Dim sql As String
        'Dim dt As DataTable

        Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)
        Dim sql As String = ""
        sql = ""
        sql &= " select SVID"
        sql &= " ,Name"
        sql &= " ,case Avail when 'Y' then '啟用' else '不啟用' end Avail"
        sql &= " from ID_Survey"
        sql &= " where 1=1"
        If Ipt_Name.Value <> "" Then   '搜尋條件
            sql &= " and Name like '%" & Ipt_Name.Value & "%' "
        End If
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        msg.Visible = True
        table_F.Visible = True
        table_I.Visible = False
        DataGrid1.Visible = False
        PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            'msg.Visible = False
            msg.Text = ""
            table_F.Visible = True
            table_I.Visible = False
            DataGrid1.Visible = True
            PageControler1.Visible = True
            'PageControler1.SqlString = sql
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub insert_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles insert.ServerClick

        '新增
        table_F.Visible = False
        table_I.Visible = True
        DataGrid1.Visible = False
        msg.Visible = False
        PageControler1.Visible = False
        IputQName.Value = ""
        Ipt_Name.Value = ""
        ISUSE.Checked = True
        Mode.Value = "I"
    End Sub

    Private Sub save_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles save.ServerClick
        Dim sql As String = ""
        sql = ""
        sql &= " INSERT INTO ID_Survey (SVID,Name,Avail,ModifyAcct,ModifyDate) "
        sql &= " values (@SVID,@Name,@Avail,@ModifyAcct,getdate())"
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " UPDATE ID_Survey"
        sql &= " SET Name=@Name"
        sql &= " ,Avail=Avail"
        sql &= " ,ModifyAcct=ModifyAcct"
        sql &= " ,ModifyDate=getdate()"
        sql &= " WHERE SVID =@SVID"
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " SELECT * FROM ID_Survey"
        sql &= " WHERE SVID =@SVID"
        Dim sCmd As New SqlCommand(sql, objconn)

        Dim ISUSE2 As String = "N"
        If ISUSE.Checked = True Then '是否啟用
            ISUSE2 = "Y"
        End If

        'Dim dr As DataRow
        Call TIMS.OpenDbConn(objconn)
        If Mode.Value = "I" Then           '新增
            If IputQName.Value <> "" Then  '問卷名稱不是空值時

                Dim iSVID As Integer = DbAccess.GetNewId(objconn, "ID_SURVEY_SVID_SEQ,ID_SURVEY,SVID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("SVID", SqlDbType.Int).Value = iSVID
                    .Parameters.Add("Name", SqlDbType.VarChar).Value = IputQName.Value
                    .Parameters.Add("Avail", SqlDbType.VarChar).Value = ISUSE2
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With

                'sql = ""
                'sql &= " INSERT INTO ID_Survey (Name,Avail,ModifyAcct) "
                'sql &= " values('" & IputQName.Value & "', '" & ISUSE2 & "','" & sm.UserInfo.UserID & "')"
                'DbAccess.ExecuteNonQuery(sql)
                Common.MessageBox(Me, "新增成功")
                Call dt_search()
            Else
                Common.MessageBox(Me, "請輸入問卷名稱")

            End If

        Else                                '修改
            If Mode.Value = "E" And SVID.Value <> "" Then
                If IputQName.Value <> "" Then

                    'sql = "Select * from ID_Survey Where SVID = '" & SVID.Value & "'"
                    'dr = DbAccess.GetOneRow(sql)
                    With uCmd
                        .Parameters.Clear()
                        .Parameters.Add("Name", SqlDbType.VarChar).Value = IputQName.Value
                        .Parameters.Add("Avail", SqlDbType.VarChar).Value = ISUSE2
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                        .Parameters.Add("SVID", SqlDbType.Int).Value = Val(SVID.Value)
                        .ExecuteNonQuery()
                    End With

                    'sql = "update ID_Survey "
                    'sql += "set Name ='" & IputQName.Value & "',Avail ='" & ISUSE2 & "',ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                    'sql += "where SVID = '" & SVID.Value & "' "
                    'DbAccess.ExecuteNonQuery(sql)
                    Common.MessageBox(Me, "修改成功")
                    Call dt_search()
                Else
                    Common.MessageBox(Me, "請輸入問卷名稱")
                End If

            End If
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                Dim drv As DataRowView = e.Item.DataItem
                Dim btn_edit As Button = e.Item.FindControl("Btn_edit")
                Dim btn_del As Button = e.Item.FindControl("Btn_del")

                btn_edit.CommandArgument = drv("SVID").ToString
                btn_del.Attributes("onclick") = "return confirm('確定是否要刪除?');"
                btn_del.CommandArgument = drv("SVID").ToString
        End Select

    End Sub

    Function GetSurvey1(ByVal iSVID As Integer) As DataRow
        Dim rst As DataRow = Nothing
        Dim sql As String = " SELECT * FROM ID_Survey WHERE SVID =@SVID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SVID", SqlDbType.Int).Value = iSVID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Return rst
        rst = dt.Rows(0)
        Return rst
    End Function

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand

        'Dim sql As String
        'Dim dt As DataTable
        'Dim dr As DataRow

        Select Case e.CommandName
            Case "edit"        '修改
                Mode.Value = "E"

                table_I.Visible = True
                table_F.Visible = False
                DataGrid1.Visible = False
                PageControler1.Visible = False

                SVID.Value = e.CommandArgument
                Dim dr As DataRow = GetSurvey1(Val(SVID.Value))
                'Sql = "Select * from ID_Survey Where SVID = '" & e.CommandArgument & "'"
                'dr = DbAccess.GetOneRow(Sql)
                IputQName.Value = dr("Name")

            Case "del"          '刪除
                SVID.Value = e.CommandArgument
                Dim dr As DataRow = GetSurvey1(Val(SVID.Value))
                If dr Is Nothing Then Exit Sub

                Dim sql As String = ""
                sql = "DELETE ID_Survey WHERE SVID='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)

                Common.MessageBox(Me, "刪除成功")
                Call dt_search()

        End Select

    End Sub
End Class
