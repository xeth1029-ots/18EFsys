Partial Class SYS_02_006
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim Auth_AccRWClass As DataTable

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
            msg.Text = ""
            Button1.Attributes("onclick") = "return search();"
            Account.Attributes("onchange") = "document.getElementById('Table5').style.display='none';"

            Call TIMS.Get_LoginYear(DDLYears, sm.UserInfo.UserID, objconn) '順序1.
            If sm.UserInfo.Years <> Nothing Then Common.SetListItem(DDLYears, sm.UserInfo.Years)
            If DDLYears.SelectedValue <> "" Then
                Call TIMS.Get_LoginPlan(DDLPlan, sm.UserInfo.UserID, sm.UserInfo.LID, DDLYears.SelectedValue, objconn) '順序2.
            End If
            Account = TIMS.Get_Account(Account, sm.UserInfo.PlanID, sm.UserInfo.RID, objconn) '順序3.
            Table5.Style.Item("display") = "none"
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    If sm.UserInfo.RoleID <> 0 Then
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            If FunDr("Sech") = "1" Then
        '                Button1.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    '查詢 SQL
    Sub sUtl_Search1()
        If DDLPlan.SelectedValue = "" Then
            Common.MessageBox(Me, "未選擇有效計畫請重新選擇!!")
            Exit Sub
        End If

        Dim sql As String
        Dim dt As DataTable

        Checkbox1.Checked = False
        '查詢帳號 有權限使用的班級資料。
        sql = "SELECT * FROM Auth_AccRWClass WHERE Account='" & Account.SelectedValue & "'"
        Auth_AccRWClass = DbAccess.GetDataTable(sql, objconn)

        sql = ""
        sql &= " SELECT c.OCID"
        sql &= " ,c.ClassCName"
        sql &= " ,c.CyclType"
        sql &= " ,c.LevelType"
        sql &= " ,c.CTName "
        sql &= " FROM Class_ClassInfo c"
        sql &= " WHERE 1=1"
        '是否轉入成功
        sql &= " AND c.IsSuccess='Y'"
        Select Case sm.UserInfo.TPlanID 'select * from key_plan where tplanid ='02'
            Case "02" '自辦職前訓練。業務。
                sql &= " and c.RID LIKE '" & sm.UserInfo.RID & "%'"
            Case Else
                sql &= " and c.RID = '" & sm.UserInfo.RID & "'"
        End Select
        '依年度
        '依計畫
        'sql &= " and PlanID='" & sm.UserInfo.PlanID & "'"
        sql &= " and PlanID='" & DDLPlan.SelectedValue & "'"

        Select Case Situation.SelectedIndex
            Case 0 '有賦予
                sql &= " and OCID IN (SELECT OCID FROM Auth_AccRWClass WHERE Account='" & Account.SelectedValue & "')"
            Case 1 '未賦予
                sql &= " and OCID NOT IN (SELECT OCID FROM Auth_AccRWClass WHERE Account='" & Account.SelectedValue & "')"
        End Select
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        Table5.Style.Item("display") = "none"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table5.Style.Item("display") = ""
            hidMYChkBoxValue.Value = ""
            DataList1.DataSource = dt
            DataList1.DataBind()

            Checkbox1.Attributes("onclick") = "SelectAll(this.checked," & dt.Rows.Count & ");"
        End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sUtl_Search1()
    End Sub

    Private Sub DataList1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles DataList1.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            If e.Item.ItemIndex Mod 6 > 2 Then e.Item.BackColor = Color.White
            'e.Item.BackColor = Color.FromName("#FFF8F0")

            Dim drv As DataRowView = e.Item.DataItem
            Dim MyCheckBox As CheckBox = e.Item.FindControl("ClassName")
            Dim Teacher As Label = e.Item.FindControl("Teacher")
            Dim OCID As HtmlInputHidden = e.Item.FindControl("OCID")

            MyCheckBox.Attributes("onclick") = "SelectRtn(this.checked, '" & Checkbox1.ClientID & "');"

            If hidMYChkBoxValue.Value <> "" Then hidMYChkBoxValue.Value &= ","
            hidMYChkBoxValue.Value &= MyCheckBox.ClientID

            MyCheckBox.Text = TIMS.GET_CLASSNAME(Convert.ToString(drv("ClassCName")), Convert.ToString(drv("CyclType")))

            If Not IsDBNull(drv("LevelType")) Then
                If Int(drv("LevelType")) <> 0 Then
                    MyCheckBox.Text += "第" & TIMS.GetChtNum(Int(drv("LevelType"))) & "階段"
                End If
            End If

            If drv("CTName").ToString <> "" Then Teacher.Text = "老師：" & drv("CTName").ToString
            OCID.Value = drv("OCID").ToString()
            If Auth_AccRWClass.Select("OCID='" & drv("OCID").ToString & "'").Length <> 0 Then
                MyCheckBox.Checked = True
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        Dim da As SqlDataAdapter = Nothing
        'Dim conn As SqlConnection
        Dim dr As DataRow
        Dim dt As DataTable

        Call TIMS.OpenDbConn(objconn)
        sql = "SELECT * FROM Auth_AccRWClass WHERE Account='" & Account.SelectedValue & "'"
        dt = DbAccess.GetDataTable(sql, da, objconn)

        Dim iACID As Int64 = 0

        'For Each item As DataListItem In DataList1.Items
        '    Dim MyCheckBox As CheckBox = item.FindControl("ClassName")
        '    Dim OCID As HtmlInputHidden = item.FindControl("OCID")

        '    If MyCheckBox.Checked Then
        '        If dt.Select("OCID='" & OCID.Value & "'").Length = 0 Then
        '            dr = dt.NewRow
        '            dt.Rows.Add(dr)

        '            '改由程式產生Pk值
        '            iACID = DbAccess.GetNewId(objconn, "Auth_AccRWClass_ACID_SEQ,Auth_AccRWClass,ACID")
        '            dr("ACID") = iACID
        '        Else
        '            dr = dt.Select("OCID='" & OCID.Value & "'")(0)
        '        End If
        '        dr("Account") = Account.SelectedValue
        '        dr("OCID") = OCID.Value
        '        dr("ModifyAcct") = sm.UserInfo.UserID
        '        dr("ModifyDate") = Now
        '    Else
        '        If dt.Select("OCID='" & OCID.Value & "'").Length <> 0 Then
        '            dr = dt.Select("OCID='" & OCID.Value & "'")(0)
        '            dr.Delete()
        '        End If
        '    End If
        'Next

        'DbAccess.UpdateDataTable(dt, da)

        '新增
        sql = " insert into AUTH_ACCRWCLASS (ACID,ACCOUNT,OCID,MODIFYACCT,MODIFYDATE) "
        sql += " values (@ACID,@ACCOUNT,@OCID,@MODIFYACCT,@MODIFYDATE) "
        Dim iCmd As New SqlCommand(sql, objconn)

        '修改
        sql = ""
        sql &= " UPDATE AUTH_ACCRWCLASS"
        sql += " SET MODIFYACCT=@MODIFYACCT "
        sql += " ,MODIFYDATE=@MODIFYDATE "
        sql += " WHERE 1=1 "
        sql += " AND ACID=@ACID "
        Dim uCmd As New SqlCommand(sql, objconn)

        '刪除
        sql = " delete from AUTH_ACCRWCLASS where ACID=@ACID "
        Dim dCmd As New SqlCommand(sql, objconn)

        For Each item As DataListItem In DataList1.Items
            Dim MyCheckBox As CheckBox = item.FindControl("ClassName")
            Dim OCID As HtmlInputHidden = item.FindControl("OCID")

            If MyCheckBox.Checked Then
                If dt.Select("OCID='" & OCID.Value & "'").Length = 0 Then
                    '改由程式產生Pk值
                    iACID = DbAccess.GetNewId(objconn, "Auth_AccRWClass_ACID_SEQ,Auth_AccRWClass,ACID")

                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("ACID", SqlDbType.Int).Value = iACID
                        .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = Account.SelectedValue
                        .Parameters.Add("OCID", SqlDbType.Int).Value = OCID.Value
                        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value = Now
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
                    End With
                Else
                    'dr = dt.Select("OCID='" & OCID.Value & "'")(0)
                    'With uCmd
                    '    .Parameters.Clear()
                    '    .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = Account.SelectedValue
                    '    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '    .Parameters.Add("ACID", SqlDbType.Int).Value = Convert.ToInt64(dr("ACID"))
                    '    .Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value = Now
                    '    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    '    DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                    'End With
                End If
            Else
                If dt.Select("OCID='" & OCID.Value & "'").Length <> 0 Then
                    dr = dt.Select("OCID='" & OCID.Value & "'")(0)
                    With dCmd
                        .Parameters.Clear()
                        .Parameters.Add("ACID", SqlDbType.Int).Value = Convert.ToInt64(dr("ACID"))
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(dCmd.CommandText, objconn, dCmd.Parameters)
                    End With
                End If
            End If
        Next

        Common.MessageBox(Me, "儲存成功!")

        '查詢
        Call sUtl_Search1()
    End Sub

    '年度選擇。
    Protected Sub DDLYears_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDLYears.SelectedIndexChanged
        If DDLYears.SelectedValue <> "" Then
            Call TIMS.Get_LoginPlan(DDLPlan, sm.UserInfo.UserID, sm.UserInfo.LID, DDLYears.SelectedValue, objconn) '順序2.
        End If
    End Sub



End Class
