Partial Class SD_05_011_add
    Inherits AuthBasePage

    'Dim ProcessType, Stdid, source As String
    'Dim objconn As SqlConnection
    'Dim objreader As SqlDataReader
    'Dim FunDr As DataRow

    Dim ProcessType As String = ""
    Dim Stdid As String = ""
    Dim source As String = ""

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        ProcessType = TIMS.ClearSQM(Request("ProcessType"))
        Stdid = TIMS.ClearSQM(Request("stdid"))
        source = TIMS.ClearSQM(Request("source"))


        If Not Page.IsPostBack Then

            Call cCreate1()

            Call show_data1()

        End If


    End Sub

    Sub cCreate1()
        Button1.Attributes("onclick") = "history.go(-1);"
        bt_save.Attributes("onclick") = "javascript:return chkdata();"

        If ProcessType = "Insert" Then
            '20100208 按新增時代查詢之 身分證號碼 & 姓名
            SID.Text = TIMS.ClearSQM(Request("StuID"))
            Name.Text = TIMS.ClearSQM(Request("StuName"))

            If Not YearList.Items.FindByText("2004") Is Nothing Then
                YearList.Items.Remove(YearList.Items.FindByText("2004"))
            End If
        End If

        Plan_List = TIMS.Get_TPlan(Plan_List,,,,, objconn)

        DistrictList = TIMS.Get_DistID(DistrictList, Nothing, objconn)

        'YearList = TIMS.GetSyear(YearList)
        'Dim sqlstr_plan As String = "select * from Key_Plan"
        'objreader = DbAccess.GetReader(sqlstr_plan, objconn)
        'Me.Plan_List.DataSource = objreader
        'Me.Plan_List.DataTextField = "PlanName"
        'Me.Plan_List.DataValueField = "TPlanID"
        'Me.Plan_List.DataBind()
        'Me.Plan_List.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'objreader.Close()
        'objconn.Close()

        'Dim sqlstr_id As String = "select * from ID_District"
        'objreader = DbAccess.GetReader(sqlstr_id, objconn)
        'Me.DistrictList.DataSource = objreader
        'Me.DistrictList.DataTextField = "name"
        'Me.DistrictList.DataValueField = "DistID"
        'Me.DistrictList.DataBind()
        'Me.DistrictList.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'objreader.Close()
        'objconn.Close()

    End Sub

    Sub show_data1()
        If ProcessType = "Update" Then
            If sm.UserInfo.DistID = "000" Then bt_save.Visible = False

            Dim row_list As DataRow
            Dim sqlstr_list As String
            If source.ToUpper = "STDALL" Then

                sqlstr_list = "select * from  StdAll where StdID=" & Stdid
                row_list = DbAccess.GetOneRow(sqlstr_list)

                YearList.SelectedValue = row_list("Years")
                Common.SetListItem(DistrictList, row_list("DistID").ToString)

                Plan_List.SelectedValue = row_list("TPlanID")
                If Convert.IsDBNull(row_list("ClassName")) Then
                    ClassName.Text = ""
                Else
                    ClassName.Text = row_list("ClassName")
                End If
                If Convert.IsDBNull(row_list("CosUnit")) Then
                    CosUnit.Text = ""
                Else
                    CosUnit.Text = row_list("CosUnit")
                End If
                If Convert.IsDBNull(row_list("TrinUnit")) Then
                    trinUnit.Text = ""
                Else
                    trinUnit.Text = row_list("TrinUnit")
                End If
                If Convert.IsDBNull(row_list("SDate")) Then
                    SDate.Text = ""
                Else
                    SDate.Text = row_list("SDate")
                End If
                If Convert.IsDBNull(row_list("EDate")) Then
                    EDate.Text = ""
                Else
                    EDate.Text = row_list("EDate")
                End If
                If Convert.IsDBNull(row_list("Name")) Then
                    Name.Text = ""
                Else
                    Name.Text = row_list("Name")
                End If
                If Convert.IsDBNull(row_list("SID")) Then
                    SID.Text = ""
                Else
                    SID.Text = row_list("SID")
                End If
                If Convert.IsDBNull(row_list("Birth")) Then
                    birthday.Text = ""
                Else
                    birthday.Text = row_list("Birth")
                End If
                Common.SetListItem(Sex_List, row_list("Sex").ToString)
                If Convert.IsDBNull(row_list("Ident")) Then
                    Ident.Text = ""
                Else
                    Ident.Text = row_list("Ident")
                End If
                If Convert.IsDBNull(row_list("Tel")) Then
                    Tel.Text = ""
                Else
                    Tel.Text = row_list("Tel")
                End If
                If Convert.IsDBNull(row_list("Addr")) Then
                    Addr.Text = ""
                Else
                    Addr.Text = row_list("Addr")
                End If
            ElseIf source.ToUpper = "HISTORY_STUDENTINFO93" Then
                sqlstr_list = "select * from  History_StudentInfo93  where Serial=" & Stdid
                row_list = DbAccess.GetOneRow(sqlstr_list)

                Common.SetListItem(YearList, "2004")
                Common.SetListItem(DistrictList, row_list("DistID").ToString)
                Common.SetListItem(Plan_List, row_list("TPlanID").ToString)
                If Convert.IsDBNull(row_list("ClassName")) Then
                    ClassName.Text = ""
                Else
                    ClassName.Text = row_list("ClassName")
                End If
                If Convert.IsDBNull(row_list("CosUnit")) Then
                    CosUnit.Text = ""
                Else
                    CosUnit.Text = row_list("CosUnit")
                End If
                If Convert.IsDBNull(row_list("TrinUnit")) Then
                    trinUnit.Text = ""
                Else
                    trinUnit.Text = row_list("TrinUnit")
                End If
                If Convert.IsDBNull(row_list("SDate")) Then
                    SDate.Text = ""
                Else
                    SDate.Text = row_list("SDate")
                End If
                If Convert.IsDBNull(row_list("EDate")) Then
                    EDate.Text = ""
                Else
                    EDate.Text = row_list("EDate")
                End If
                If Convert.IsDBNull(row_list("Name")) Then
                    Name.Text = ""
                Else
                    Name.Text = row_list("Name")
                End If
                If Convert.IsDBNull(row_list("IDNO")) Then
                    SID.Text = ""
                Else
                    SID.Text = row_list("IDNO")
                End If
                If Convert.IsDBNull(row_list("Birth")) Then
                    birthday.Text = ""
                Else
                    birthday.Text = row_list("Birth")
                End If
                If Convert.IsDBNull(row_list("Sex")) Then
                Else
                    Sex_List.SelectedValue = row_list("Sex")
                End If
                If Convert.IsDBNull(row_list("Ident")) Then
                    Ident.Text = ""
                Else
                    Ident.Text = row_list("Ident")
                End If
                If Convert.IsDBNull(row_list("Tel")) Then
                    Tel.Text = ""
                Else
                    Tel.Text = row_list("Tel")
                End If
                If Convert.IsDBNull(row_list("Addr")) Then
                    Addr.Text = ""
                Else
                    Addr.Text = row_list("Addr")
                End If
                '可以查看,不能修改
                YearList.Enabled = False
                DistrictList.Enabled = False
                Plan_List.Enabled = False
                ClassName.Enabled = False
                CosUnit.Enabled = False
                trinUnit.Enabled = False
                SDate.Enabled = False
                EDate.Enabled = False
                Name.Enabled = False
                SID.Enabled = False
                birthday.Enabled = False
                Sex_List.Enabled = False
                Ident.Enabled = False
                Tel.Enabled = False
                Addr.Enabled = False
                bt_save.Enabled = False

            End If

        End If

    End Sub
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim sqlTable As New DataTable
        Dim sqldr As DataRow = Nothing
        Dim sqlStr As String = ""

        If ProcessType = "Insert" Then

            sqlStr = "SELECT * FROM StdAll WHERE 1<>1"
            sqlTable = DbAccess.GetDataTable(sqlStr, sqlAdapter, objconn)
            sqldr = sqlTable.NewRow

            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "alert('歷史資料儲存成功!!');" + vbCrLf
            strScript += "location.href='SD_05_011.aspx?ID='+document.getElementById('Re_ID').value;" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)
        ElseIf ProcessType = "Update" Then

            sqlStr = "select * from StdAll where StdID=" & Stdid
            sqldr = DbAccess.GetUpdateRow(sqlStr, sqlTable, sqlAdapter, objconn)

            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "alert('歷史資料修改成功!!');" + vbCrLf
            strScript += "location.href='SD_05_011.aspx?ID='+document.getElementById('Re_ID').value;" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)
        End If

        sqldr("Years") = YearList.SelectedValue
        sqldr("DistID") = DistrictList.SelectedValue
        sqldr("TPlanID") = Plan_List.SelectedValue
        sqldr("ClassName") = ClassName.Text
        sqldr("CosUnit") = CosUnit.Text
        If trinUnit.Text = "" Then
            sqldr("TrinUnit") = Convert.DBNull
        Else
            sqldr("TrinUnit") = trinUnit.Text
        End If
        If SDate.Text = "" Then
            sqldr("SDate") = Convert.DBNull
        Else
            sqldr("SDate") = SDate.Text
        End If
        If EDate.Text = "" Then
            sqldr("EDate") = Convert.DBNull
        Else
            sqldr("EDate") = EDate.Text
        End If
        If Name.Text = "" Then
            sqldr("Name") = Convert.DBNull
        Else
            sqldr("Name") = Name.Text
        End If
        If SID.Text = "" Then
            sqldr("SID") = Convert.DBNull
        Else
            sqldr("SID") = SID.Text
        End If
        If birthday.Text = "" Then
            sqldr("Birth") = Convert.DBNull
        Else
            sqldr("Birth") = birthday.Text
        End If
        If Sex_List.SelectedValue = "" Then
            sqldr("Sex") = Convert.DBNull
        Else
            sqldr("Sex") = Sex_List.SelectedValue
        End If
        If Ident.Text = "" Then
            sqldr("Ident") = Convert.DBNull
        Else
            sqldr("Ident") = Ident.Text
        End If
        If Tel.Text = "" Then
            sqldr("Tel") = Convert.DBNull
        Else
            sqldr("Tel") = Tel.Text
        End If
        If Addr.Text = "" Then
            sqldr("Addr") = Convert.DBNull
        Else
            sqldr("Addr") = Addr.Text
        End If

        sqldr("ModifyAcct") = sm.UserInfo.UserID
        sqldr("ModifyDate") = Now()
        If ProcessType = "Insert" Then
            sqlTable.Rows.Add(sqldr)
            DbAccess.UpdateDataTable(sqlTable, sqlAdapter)
        Else
            sqlAdapter.Update(sqlTable)
        End If


    End Sub
End Class
