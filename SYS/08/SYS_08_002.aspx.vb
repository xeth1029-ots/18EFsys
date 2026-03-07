Partial Class SYS_08_002
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
        '分頁設定
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = Datagrid2

        'Save.Disabled = True
        'If blnCanAdds Then Save.Disabled = False
        'search.Disabled = True
        'If blnCanSech Then search.Disabled = False
        'If blnCanSech Then Save.Disabled = False

        '分頁設定
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
        '                Save.Disabled = False
        '            Else

        '            End If

        '            If FunDr("Sech") = "1" Then
        '                search.Disabled = False
        '            Else
        '                search.Disabled = True
        '                Save.Disabled = True
        '            End If
        '        End If
        '    End If
        'End If

        If Not IsPostBack Then

            table_F.Visible = True
            Table3.Visible = False
            PageControler1.Visible = False
            PageControler2.Visible = False

            Save.Attributes("onclick") = "Check();"
        End If



    End Sub

    Private Sub search_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.ServerClick, Return1.ServerClick
        dt_search()
    End Sub

    Sub dt_search()
        'Dim str As String
        'Dim dt As DataTable
        Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)

        Dim sql As String = ""
        sql = ""
        sql &= " select SVID, Name"
        sql &= " ,case Avail when 'Y' then '啟用' else '不啟用' end as Avail"
        sql &= " from ID_Survey"
        sql &= " where 1=1"
        sql &= " and Avail <> 'N'"
        If Ipt_Name.Value <> "" Then   '搜尋條件
            sql &= " and Name like '%" & Ipt_Name.Value & "%' "
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            msg.Text = "查無資料"
            msg.Visible = True
            table_F.Visible = True
            Table2.Visible = False
            Table3.Visible = False
            Table4.Visible = False
            DataGrid1.Visible = False
            PageControler1.Visible = False
            Exit Sub
        End If

        msg.Visible = False
        table_F.Visible = True
        Table2.Visible = True
        Table3.Visible = False
        Table4.Visible = False
        DataGrid1.Visible = True
        PageControler1.Visible = True
        'PageControler1.SqlString = sql
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

        'If TIMS.Get_SQLRecordCount(sql, objconn) = 0 Then
        'Else
        'End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim SVID As String = TIMS.GetMyValue(sCmdArg, "SVID")
        Dim cName As String = TIMS.GetMyValue(sCmdArg, "Name")
        If SVID = "" Then Exit Sub

        Select Case e.CommandName
            Case "edit1"
                Call Etid_TitleQ(SVID, cName)
        End Select
        'If e.Item.ItemType = ListItemType.AlternatingItem Or e.Item.ItemType = ListItemType.Item Then
        '    Dim drv As DataRowView = e.Item.DataItem
        '    Etid_TitleQ(e.Item.Cells(4).Text.ToString, e.Item.Cells(1).Text.ToString)
        'End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                Dim btn_E As Button = e.Item.FindControl("Btn_edit")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SVID", Convert.ToString(drv("SVID")))
                TIMS.SetMyValue(sCmdArg, "Name", Convert.ToString(drv("Name")))
                btn_E.CommandArgument = sCmdArg
        End Select
        'If e.Item.ItemType <> ListItemType.Footer And e.Item.ItemType <> ListItemType.Header Then
        'End If
    End Sub

    Sub Etid_TitleQ(ByVal SVID As String, ByVal Name As String)

        table_F.Visible = False
        Table2.Visible = False
        Table3.Visible = True
        Table4.Visible = True
        QName.InnerText = Name
        SVID2.Value = SVID
        MODE.Value = "A"

        bt_search2()

    End Sub

    Sub bt_search2()

        'Dim sql As String = ""

        SVID2.Value = TIMS.ClearSQM(SVID2.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql = " SELECT SKID, TOPIC, SERIAL FROM KEY_SURVEYKIND" & vbCrLf
        sql = " where SVID =@SVID" & vbCrLf
        sql = " order by Serial"
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SVID", SqlDbType.VarChar).Value = SVID2.Value
            dt.Load(.ExecuteReader())
        End With

        'Call CloseDbConn(conn)
        If dt.Rows.Count = 0 Then
            table_F.Visible = False
            Table2.Visible = False
            Table3.Visible = True
            Table4.Visible = False
            Exit Sub
        End If

        table_F.Visible = False
        Table2.Visible = False
        Table3.Visible = True
        Table4.Visible = True
        PageControler2.Visible = True
        PageControler2.PageDataTable = dt
        'Pagecontroler2.SqlString = sql
        PageControler2.ControlerLoad()

    End Sub

    Private Sub Save_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.ServerClick

        Dim sql As String
        Dim dt As DataTable
        Dim i As Integer
        Dim Array1 As Array
        Dim Serailnew As Integer


        If Name3.Value <> "" And Serial.Value <> "" Then '名稱及排序都有填


            If MODE.Value = "E" Then '修改 sql 排除正在修改的那一筆資料,取出 >= 欲修改的排序

                sql = " select SKID, Serial from Key_SurveyKind where SVID = " & SVID2.Value & " and Serial >= " & Serial.Value & " and SKID <> " & SKID2.Value & "  order by Serial"
                dt = DbAccess.GetDataTable(sql)
            Else      '新增 取出>= 欲修改的排序
                sql = " select SKID, Serial from Key_SurveyKind where SVID = " & SVID2.Value & " and Serial >= " & Serial.Value & " order by Serial"
                dt = DbAccess.GetDataTable(sql)

            End If

            If dt.Select("Serial = '" & Serial.Value & "'").Length > 0 Then '有重複的排序時

                If MODE.Value = "E" Then        '修改

                    For i = 0 To dt.Rows.Count - 1
                        Array1 = dt.Rows(i).ItemArray
                        If Array1(1).ToString = Serial.Value Then '如果己經有的排序跟欲輸入的排序相同時

                            Serailnew = CInt(Array1(1)) + 1 '新的排序 = 欲輸入的排序 +1

                            sql = "Update Key_SurveyKind "
                            sql += "Set Topic = '" & Name3.Value & "',Serial = '" & Serial.Value & "'"
                            sql += "where SKID = '" & SKID2.Value & "' "
                            DbAccess.ExecuteNonQuery(sql)

                        End If

                        sql = "Update Key_SurveyKind "
                        sql += "Set Serial = '" & Serailnew & "'"
                        sql += "where SKID = '" & Array1(0).ToString & "' "
                        DbAccess.ExecuteNonQuery(sql)

                        Serailnew = Serailnew + 1  '排序 + 1

                    Next

                    Common.MessageBox(Me, "修改成功")
                    Name3.Value = ""
                    Serial.Value = ""
                    Etid_TitleQ(SVID2.Value.ToString, QName.InnerText.ToString)

                Else                              '新增

                    For i = 0 To dt.Rows.Count - 1
                        Array1 = dt.Rows(i).ItemArray
                        If Array1(1).ToString = Serial.Value Then  ' 如果己經有的排序跟輸入的排序相同時

                            Serailnew = CInt(Array1(1)) + 1   '新的排序 = 欲輸入的排序 +1

                            sql = "Insert Into Key_SurveyKind(Topic,SVID,Serial,ModifyAcct)"
                            sql += "values('" & Name3.Value & "', '" & SVID2.Value & "','" & Serial.Value & "' ,'" & sm.UserInfo.UserID & "')"
                            DbAccess.ExecuteNonQuery(sql)

                        End If
                        sql = "Update Key_SurveyKind "
                        sql += "Set Serial = '" & Serailnew & "'"
                        sql += "where SKID = '" & Array1(0).ToString & "' "
                        DbAccess.ExecuteNonQuery(sql)
                        Serailnew = Serailnew + 1 '排序 + 1
                    Next
                    Common.MessageBox(Me, "新增成功")
                    Name3.Value = ""
                    Serial.Value = ""
                    Etid_TitleQ(SVID2.Value.ToString, QName.InnerText.ToString)
                End If

            Else '沒有重複的排序號碼

                If MODE.Value = "E" Then '修改

                    sql = "Update Key_SurveyKind "
                    sql += "Set Topic = '" & Name3.Value & "',Serial = '" & Serial.Value & "'"
                    sql += "where SKID = '" & SKID2.Value & "' "

                    DbAccess.ExecuteNonQuery(sql)
                    Common.MessageBox(Me, "修改成功")
                    Name3.Value = ""
                    Serial.Value = ""
                    Etid_TitleQ(SVID2.Value.ToString, QName.InnerText.ToString)

                Else  '新增
                    sql = "Insert Into Key_SurveyKind(Topic,SVID,Serial,ModifyAcct)"
                    sql += "values('" & Name3.Value & "', '" & SVID2.Value & "','" & Serial.Value & "' ,'" & sm.UserInfo.UserID & "')"
                    DbAccess.ExecuteNonQuery(sql)
                    Common.MessageBox(Me, "新增成功")
                    Name3.Value = ""
                    Serial.Value = ""
                    Etid_TitleQ(SVID2.Value.ToString, QName.InnerText.ToString)

                End If

            End If

        End If

    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound

        If e.Item.ItemType <> ListItemType.Footer And e.Item.ItemType <> ListItemType.Header Then

            Dim drv As DataRowView = e.Item.DataItem
            Dim btn_E As Button = e.Item.FindControl("edit")
            Dim btn_D As Button = e.Item.FindControl("del")

            btn_E.CommandArgument = drv("SKID").ToString
            btn_D.CommandArgument = drv("SKID").ToString
            btn_D.Attributes("onclick") = "confirm('是否確定要刪除?');"

        End If

    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand

        Dim sql As String
        'Dim dt As DataTable
        Dim dr As DataRow

        Select Case e.CommandName

            Case "edit"   '修改

                MODE.Value = "E"

                sql = "select * from Key_SurveyKind where SKID = '" & e.CommandArgument & "' "
                dr = DbAccess.GetOneRow(sql)

                SKID2.Value = e.CommandArgument
                Name3.Value = dr("Topic").ToString
                Serial.Value = dr("Serial").ToString

            Case "del" '刪除

                sql = "Delete Key_SurveyKind where SKID = '" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql)
                Common.MessageBox(Me, "刪除成功")
                Etid_TitleQ(SVID2.Value.ToString, QName.InnerText.ToString)

        End Select
    End Sub

End Class
