Partial Class SV_01_002
    Inherits AuthBasePage

    'Dim FunDr As DataRow
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
        '分頁設定
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = Datagrid2
        '分頁設定
        'If sm.UserInfo.RoleID <> 0 Then
        'End If
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>Top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '    End If
        'End If

        If Not IsPostBack Then

            table_F.Visible = True
            Table3.Visible = False
            PageControler1.Visible = False
            PageControler2.Visible = False

        End If

        btnSave1.Attributes("onclick") = "Check();"

    End Sub

    Private Sub search_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.ServerClick, Return1.ServerClick
        Call dt_search()
    End Sub

    Sub dt_search()
        Dim sql As String = ""
        'Dim dt As DataTable
        'If Ipt_Name.Value <> "" Then   '搜尋條件
        '    str = " and Name like '%" & Ipt_Name.Value & "%' "
        'End If

        sql = ""
        sql &= " select a.SVID" & vbCrLf
        sql &= " , a.Name" & vbCrLf
        sql &= " ,case a.Avail when 'Y' then '啟用' else '不啟用' end Avail" & vbCrLf
        sql += " ,a.Avail ISUSE" & vbCrLf
        sql += " ,a.internal " & vbCrLf
        sql &= " from ID_Survey a" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.Avail <> 'N'" & vbCrLf
        If Ipt_Name.Value <> "" Then   '搜尋條件
            sql &= " and a.Name like '%" & Ipt_Name.Value & "%' "
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        msg.Visible = True
        table_F.Visible = True
        Table2.Visible = True
        Table3.Visible = False
        Table4.Visible = False
        DataGrid1.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count > 0 Then
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
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

        If e.Item.ItemType <> ListItemType.Footer And e.Item.ItemType <> ListItemType.Header Then

            e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
        End If

        Dim btn_E As Button = e.Item.FindControl("Btn_edit")

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand


        If e.Item.ItemType = ListItemType.AlternatingItem Or e.Item.ItemType = ListItemType.Item Then

            Dim drv As DataRowView = e.Item.DataItem

            Etid_TitleQ(e.Item.Cells(4).Text.ToString, e.Item.Cells(1).Text.ToString) '導到第二個畫面

        End If
    End Sub

    Sub Etid_TitleQ(ByVal SVID As String, ByVal Name As String) '重新整理頁面
        table_F.Visible = False
        Table2.Visible = False
        Table3.Visible = True
        Table4.Visible = True

        'QName.InnerText = Name
        QName.Text = Name
        SVID2.Value = SVID
        MODE.Value = "A"  '表新增

        Call bt_search2()
    End Sub

    Sub bt_search2()
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT SKID, TOPIC, SERIAL,SVID"
        sql &= " FROM KEY_SURVEYKIND"
        sql &= " where 1=1"
        sql &= " AND SVID = " & SVID2.Value
        sql &= " ORDER BY SERIAL"

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        table_F.Visible = False
        Table2.Visible = False
        Table3.Visible = True
        Table4.Visible = False
        If dt.Rows.Count > 0 Then
            Dim dt2 As DataTable = Nothing
            Dim dr As DataRow = dt.Rows(0)
            dt2 = TIMS.Get_dtSdS(dr("SVID"), objconn)
            If dt2.Rows.Count > 0 Then
                btnSave1.Enabled = False
                btnSave1.ToolTip = "此【問卷分類標題設定】底下的【問卷資料填寫】已有學員填寫的資料，不能新增修改，若執意要修改，請先刪除【問卷資料填寫】!!"
            End If

            table_F.Visible = False
            Table2.Visible = False
            Table3.Visible = True
            Table4.Visible = True
            PageControler2.Visible = True

            PageControler2.PageDataTable = dt '.SqlString = sql
            PageControler2.ControlerLoad()
        End If

    End Sub

    Function MaxSerialValue(ByVal SVID As String) As Integer
        Dim rst As Integer = 0
        Dim sql As String = ""
        sql = "SELECT dbo.NVL(MAX(SERIAL),0)+1 MAXSERIAL FROM KEY_SURVEYKIND WHERE SVID='" & SVID & "' "
        rst = DbAccess.ExecuteScalar(sql, objconn)
        Return rst
    End Function

    Sub SaveData1()
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        'Dim i As Integer
        'Dim Array1 As Array
        Dim Serailnew As Integer = 0

        Serial.Value = Trim(Serial.Value)
        If Serial.Value = "0" Then
            Serial.Value = "1"
        End If

        If Name3.Value <> "" And Serial.Value <> "" Then '名稱及排序都有填
            Select Case MODE.Value
                Case "A"  '新增
                    '新增 取出>= 欲修改的排序
                    sql = " select SKID, Serial from Key_SurveyKind where SVID = " & SVID2.Value & " and Serial >= " & Serial.Value & " order by Serial"
                    dt = DbAccess.GetDataTable(sql, objconn)
                Case "E" '修改
                    '修改 sql 排除正在修改的那一筆資料,取出 >= 欲修改的排序
                    sql = " select SKID, Serial from Key_SurveyKind where SVID = " & SVID2.Value & " and Serial >= " & Serial.Value & " and SKID != " & SKID2.Value & "  order by Serial"
                    dt = DbAccess.GetDataTable(sql, objconn)
            End Select
            If dt Is Nothing Then Exit Sub
            If dt.Select("Serial = '" & Serial.Value & "'").Length > 0 Then '有重複的排序時
                Select Case MODE.Value
                    Case "A" '新增
                        'For i = 0 To dt.Rows.Count - 1
                        '    Array1 = dt.Rows(i).ItemArray
                        '    If Array1(1).ToString = Serial.Value Then  ' 如果己經有的排序跟輸入的排序相同時

                        '        Serailnew = CInt(Array1(1)) + 1   '新的排序 = 欲輸入的排序 +1

                        '        sql = "Insert Into Key_SurveyKind(Topic,SVID,Serial,ModifyAcct)" '新增欲輸入的排序
                        '        sql += "values('" & Name3.Value & "', '" & SVID2.Value & "','" & Serial.Value & "' ,'" & sm.UserInfo.UserID & "')"
                        '        DbAccess.ExecuteNonQuery(sql)

                        '    End If
                        '    sql = "Update Key_SurveyKind "    '原相同的排序 則updat成新排序(原排序+1)
                        '    sql += "Set Serial = '" & Serailnew & "',ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                        '    sql += "where SKID = '" & Array1(0).ToString & "' "
                        '    DbAccess.ExecuteNonQuery(sql)
                        '    Serailnew = Serailnew + 1 '排序 + 1 之後的筆數資料則是重新排序,都+1
                        'Next

                        Serailnew = MaxSerialValue(SVID2.Value)

                        sql = ""
                        sql &= " Insert Into Key_SurveyKind(Topic,SVID,Serial,ModifyAcct)" '新增欲輸入的排序
                        sql += " values('" & Name3.Value & "', '" & SVID2.Value & "','" & Serailnew & "' ,'" & sm.UserInfo.UserID & "')"
                        DbAccess.ExecuteNonQuery(sql, objconn)

                        Common.MessageBox(Me, "新增成功")
                        Name3.Value = ""
                        Serial.Value = "0"

                        Call Etid_TitleQ(SVID2.Value.ToString, QName.Text) '重新整理頁面

                    Case "E" '修改
                        'For i = 0 To dt.Rows.Count - 1
                        '    Array1 = dt.Rows(i).ItemArray
                        '    If Array1(1).ToString = Serial.Value Then '如果己經有的排序跟欲輸入的排序相同時

                        '        Serailnew = CInt(Array1(1)) + 1 '新的排序 = 欲輸入的排序 +1

                        '        sql = "Update Key_SurveyKind "    'update 欲修改的那一筆資料,排序為原輸入值
                        '        sql += "Set Topic = '" & Name3.Value & "',Serial = '" & Serial.Value & "',ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                        '        sql += "where SKID = '" & SKID2.Value & "' "
                        '        DbAccess.ExecuteNonQuery(sql)

                        '    End If

                        '    sql = "Update Key_SurveyKind "         'update 相同排序的那筆資料為 新排序(原排序 +1)
                        '    sql += "Set Serial = '" & Serailnew & "',ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                        '    sql += "where SKID = '" & Array1(0).ToString & "' "
                        '    DbAccess.ExecuteNonQuery(sql)

                        '    Serailnew = Serailnew + 1  '排序 + 1  將之後的資料排序重新排列,都+1

                        'Next
                        Serailnew = MaxSerialValue(SVID2.Value)

                        sql = ""
                        sql &= " Update Key_SurveyKind "    'update 欲修改的那一筆資料,排序為原輸入值
                        sql += " Set Topic = '" & Name3.Value & "',Serial = '" & Serailnew & "',ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                        sql += " where SKID = '" & SKID2.Value & "' "
                        DbAccess.ExecuteNonQuery(sql, objconn)

                        Common.MessageBox(Me, "修改成功")
                        Name3.Value = ""
                        Serial.Value = "0"
                        Call Etid_TitleQ(SVID2.Value.ToString, QName.Text) '重新整理頁面
                End Select

            Else
                '沒有重複的排序號碼
                Select Case MODE.Value
                    Case "A" '新增
                        sql = " Insert Into Key_SurveyKind(Topic,SVID,Serial,ModifyAcct)"
                        sql += " values('" & Name3.Value & "', '" & SVID2.Value & "','" & Serial.Value & "' ,'" & sm.UserInfo.UserID & "')"
                        DbAccess.ExecuteNonQuery(sql, objconn)
                        Common.MessageBox(Me, "新增成功")
                        Name3.Value = ""
                        Serial.Value = "0"

                        Call Etid_TitleQ(SVID2.Value.ToString, QName.Text) '重新整理頁面
                    Case "E" '修改
                        sql = " Update Key_SurveyKind "
                        sql += " Set Topic = '" & Name3.Value & "',Serial = '" & Serial.Value & "',ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                        sql += " where SKID = '" & SKID2.Value & "' "

                        DbAccess.ExecuteNonQuery(sql, objconn)
                        Common.MessageBox(Me, "修改成功")
                        Name3.Value = ""
                        Serial.Value = "0"

                        Call Etid_TitleQ(SVID2.Value.ToString, QName.Text) '重新整理頁面
                End Select
            End If

        End If
    End Sub

    'Private Sub Save_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.ServerClick

    'End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn_E As Button = e.Item.FindControl("edit")
                Dim btn_D As Button = e.Item.FindControl("del")
                btn_D.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                btn_E.CommandArgument = drv("SKID").ToString
                btn_D.CommandArgument = drv("SKID").ToString

                Dim dt As DataTable = Nothing
                Dim dt2 As DataTable = Nothing
                Dim sql As String = ""
                Sql = "Select * from ID_SurveyQuestion where SKID ='" & drv("SKID").ToString & "'"
                dt = DbAccess.GetDataTable(Sql, objconn)
                If dt.Rows.Count <> 0 Then
                    btn_D.Enabled = False
                    btn_D.ToolTip = "此【問卷分類標題設定】底下的【問卷題目設定】有資料，不能刪除，若執意要刪除，請先刪除【問卷題目設定】!!"
                    'btn_D.Attributes("onclick") = "alert('此【問卷分類標題設定】底下的【問卷題目設定】有資料，不能刪除，若執意要刪除，請先刪除【問卷題目設定】!!');"
                End If

                'sql2 = ""
                'sql2 += " select SSID,SOCID,DONEDATE,SVID,SKID,SQID,SAID,dbo.SUBSTR(SANOTE, 1, 4000) SANOTE,  " & vbCrLf
                'sql2 += " dbo.SUBSTR(SQNOTE, 1, 4000) SQNOTE,MODIFYACCT,MODIFYDATE  " & vbCrLf
                'sql2 += " from Stud_Survey where SKID = '" & drv("SKID").ToString & "'"
                'dt2 = DbAccess.GetDataTable(sql2, objconn)
                dt2 = TIMS.Get_dtSdS(drv("SVID"), objconn)
                If dt2.Rows.Count <> 0 Then
                    btn_E.Enabled = False
                    btn_E.ToolTip = "此【問卷分類標題設定】底下的【問卷資料填寫】已有學員填寫的資料，不能修改，若執意要修改，請先刪除【問卷資料填寫】!!"
                End If

        End Select


    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit"   '修改
                MODE.Value = "E" '修改
                sql = "select * from Key_SurveyKind where SKID = '" & e.CommandArgument & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    SKID2.Value = e.CommandArgument
                    Name3.Value = dr("Topic").ToString
                    Serial.Value = dr("Serial").ToString
                End If
            Case "del" '刪除
                sql = "Select * from ID_SurveyQuestion where SKID ='" & e.CommandArgument & "'"
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count = 0 Then
                    sql = "Delete Key_SurveyKind where SKID = '" & e.CommandArgument & "'"
                    DbAccess.ExecuteNonQuery(sql, objconn)
                    Common.MessageBox(Me, "刪除成功")

                    Call Etid_TitleQ(SVID2.Value.ToString, QName.Text) '重新整理頁面
                End If
        End Select

    End Sub

    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Call SaveData1()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
