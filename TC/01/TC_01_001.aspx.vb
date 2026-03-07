Partial Class TC_01_001
    Inherits AuthBasePage

    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not Page.IsPostBack Then
            DistID = TIMS.Get_DistID(DistID)
            DistID.Enabled = True
            trDistid.Visible = True
            Common.SetListItem(DistID, sm.UserInfo.DistID) '預設為自己轄區分署(轄區中心)
            If sm.UserInfo.DistID <> "000" Then
                trDistid.Visible = False '非署(局)不顯示
                DistID.Items.Remove(DistID.Items.FindByValue("000")) '移除署(局)
                DistID.Enabled = False
            End If
            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)
            Call Show_KeyPlan() '年度選擇後重新搜尋計畫 SQL
            Common.SetListItem(planlist, sm.UserInfo.TPlanID)
            PageControler1.Visible = False
        End If

        'DistID.Enabled = True
        If sm.UserInfo.DistID <> "000" Then
            trDistid.Visible = False '非署(局)不顯示
            Common.SetListItem(DistID, sm.UserInfo.DistID) '永遠顯示自己轄區分署(轄區中心)
            'DistID.Enabled = False
        End If

#Region "(No Use)"

        '檢查帳號的功能權限-----------------------------------Start
        'Button3.Disabled = True
        'bt_add.Enabled = False
        'If au.blnCanAdds Then bt_add.Enabled = True
        'If au.blnCanAdds Then Button3.Disabled = False
        'bt_search.Enabled = False
        'If au.blnCanSech Then bt_search.Enabled = True
        '檢查帳號的功能權限-----------------------------------End
        '匯入年度計畫
        'Button3.Attributes("onclick") = "window.open('TC_01_001_import.aspx','ZipWin','width=900,height=900,toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no');"

#End Region

        If Not Session("_search") Is Nothing Then
            Common.SetListItem(yearlist, TIMS.GetMyValue(Session("_search"), "yearlist"))
            Call Show_KeyPlan() '年度選擇後重新搜尋計畫 SQL
            Common.SetListItem(planlist, TIMS.GetMyValue(Session("_search"), "planlist"))
            PageControler1.PageIndex = 0
            'PageControler1.PageIndex = TIMS.GetMyValue(Session("_search"), "PageIndex")
            Dim MyValue As String = TIMS.GetMyValue(Session("_search"), "PageIndex")
            If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                MyValue = CInt(MyValue)
                PageControler1.PageIndex = MyValue
            End If
            'If TIMS.GetMyValue(Session("_search"), "submit") = "1" Then bt_search_Click(sender, e)
            If TIMS.GetMyValue(Session("_search"), "submit") = "1" Then Search1()
            Session("_search") = Nothing
        End If

        'DataGrid1.Visible = False
        'PageControler1.Visible = False
        'Table2.Visible = False
        Dim sql As String = ""
        sql = " SELECT ITEMVALUE FROM SYS_VAR WHERE SPAGE = 'TC_01_001' AND ITEMNAME = 'Add' "
        Dim ItemVALUE As String = ""
        ItemVALUE = DbAccess.ExecuteScalar(sql, objconn)
        If Not ItemVALUE Is Nothing Then
            If ItemVALUE = "N" Then
                bt_add.Enabled = False
                'bt_add.ToolTip = "職訓局整併計畫中，暫時無法新增／匯入"
                bt_add.ToolTip = "勞動力發展署整併計畫中，暫時無法新增／匯入"
                btn_ImpYears.Enabled = False
                TIMS.Tooltip(btn_ImpYears, bt_add.ToolTip)
                'Button3.Disabled = True
                'TIMS.Tooltip(Button3, bt_add.ToolTip)
            End If
        End If
    End Sub

    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        KeepSearch()
        TIMS.Utl_Redirect1(Me, "TC_01_001_add.aspx?ID=" & Request("ID"))
    End Sub

    Sub KeepSearch()
        Session("_search") = "yearlist=" & yearlist.SelectedValue
        Session("_search") += "&planlist=" & planlist.SelectedValue
        Session("_search") += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        If DataGrid1.Visible = True Then
            Session("_search") += "&submit=1"
        Else
            Session("_search") += "&submit=0"
        End If
    End Sub

    Sub Search1()
        If Me.yearlist.SelectedValue = "" Then
            Common.MessageBox(Me, "請輸入年度")
            Exit Sub
        End If

        'sqlstr += " ,a.Years+c.Name+b.PlanName+a.Seq" & vbCrLf '★
        'sqlstr += " +case when b.Clsyear is null or b.Clsyear > a.Years then '' else '…已停用'+ convert(varchar,Clsyear) end as PlanName" '★
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " SELECT a.TPlanID ,a.PlanID ,a.Years + c.Name + b.PlanName + a.Seq + CASE WHEN b.Clsyear IS NULL OR b.Clsyear > a.Years THEN '' ELSE '…已停用' + Clsyear END PlanName ,a.SDate ,a.EDate " & vbCrLf
        sqlstr += " FROM ID_PLAN a " & vbCrLf
        sqlstr += " JOIN Key_Plan b ON a.TPlanID = b.TPlanID " & vbCrLf
        sqlstr += " JOIN ID_District c ON a.DistID = c.DistID " & vbCrLf
        sqlstr += " WHERE 1=1 " & vbCrLf
        sqlstr += "    AND a.YEARS = '" & Me.yearlist.SelectedValue & "' " & vbCrLf

        If trDistid.Visible Then
            '署(局)顯示
            Me.ViewState("DistIDValue") = GetDistIDValue()
            If Me.ViewState("DistIDValue") <> "" Then
                sqlstr += " AND a.DistID IN (" & Me.ViewState("DistIDValue") & ") " & vbCrLf
            Else
                sqlstr += " AND a.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf
            End If
        Else
            '非署(局)不顯示
            sqlstr += "AND a.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf
        End If

        If Me.planlist.SelectedValue <> "" Then sqlstr += " AND a.TPlanID = '" & Me.planlist.SelectedValue & "' "

        'sqlstr += " order by a.PlanID"
        'Dim dt As DataTable
        'Dim sql As String
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        DataGrid2.Visible = False
        DataGrid1.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

        DataGrid1.Visible = True
        PageControler1.Visible = True
        'PageControler1.SqlString = sqlstr
        'PageControler1.PrimaryKey = "PlanID"
        'PageControler1.Sort = "PlanID"
        'PageControler1.ControlerLoad()
        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "PlanID"
        PageControler1.Sort = "PlanID"
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call Search1()
    End Sub


    '查詢 賦予的帳號列表 / 刪除
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Select Case e.CommandName
            Case "view"
                '檢查計劃是否有賦予給任何的帳號-----------------------------------Start
                sql = "" & vbCrLf
                sql += " SELECT b.Name + ' [' + a.Account + ']' 賦予的帳號列表 ,o.orgname AS 機構名稱 ,c.RSID AS 機構業務ID ,c.RID AS RID " & vbCrLf
                sql += " FROM Auth_AccRWPlan a " & vbCrLf
                sql += " JOIN auth_account b ON a.Account = b.Account " & vbCrLf
                sql += " JOIN Auth_Relship c ON a.rid = c.rid " & vbCrLf
                sql += " JOIN org_orginfo o ON c.orgid = o.orgid " & vbCrLf
                sql += " WHERE a.PlanID = '" & e.CommandArgument & "' " & vbCrLf
                sql += " ORDER BY c.RID " & vbCrLf
                dt = DbAccess.GetDataTable(sql, objconn)
                DataGrid2.Visible = False
                If Not dt Is Nothing Then
                    DataGrid1.Visible = False
                    PageControler1.Visible = False
                    Table2.Visible = True
                    DataGrid2.Visible = True
                    DataGrid2.DataSource = dt
                    DataGrid2.DataBind()
                    Exit Sub
                Else
                    Common.MessageBox(Me, "此計劃沒有賦予帳號")
                    Exit Sub
                End If
                '檢查計劃是否有賦予給任何的帳號-----------------------------------End
            Case "edit"
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_001_add.aspx?ID=" & Request("ID") & "&editid=" & e.CommandArgument)
            Case "del"
                '檢查計劃是否有賦予給任何的帳號-----------------------------------Start
                sql = " SELECT * FROM Auth_AccRWPlan WHERE PlanID = '" & e.CommandArgument & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    Common.MessageBox(Me, "此計劃有賦予帳號不能刪除")
                    Exit Sub
                End If
                '檢查計劃是否有賦予給任何的帳號-----------------------------------End

                '檢查計劃是否有賦予給任何的機構-----------------------------------Start
                sql = "SELECT * FROM Auth_AcctOrg WHERE PlanID = '" & e.CommandArgument & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    Common.MessageBox(Me, "此計劃有賦予機構不能刪除")
                    Exit Sub
                End If
                '檢查計劃是否有賦予給任何的機構-----------------------------------End

                '檢查計畫是否與Plan_PlanInfo 參照---1/31 ( Melody)
                sql = "SELECT * FROM  Plan_PlanInfo WHERE PlanID = '" & e.CommandArgument & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    Common.MessageBox(Me, "此計劃與計畫主檔參照不能刪除")
                    Exit Sub
                End If

#Region "(No Use)"

                '檢查計劃是否為青年職涯啟動計劃---2007/09/11 by AMU--------------End
                'Dim flag_YoungPlan As Boolean = False
                'sql = "SELECT seq FROM ID_Plan WHERE TPlanID=36 AND PlanID='" & e.CommandArgument & "'"
                'dr = DbAccess.GetOneRow(sql, objconn)
                'Dim YoungPlan_Planid As String = ""
                'If Not dr Is Nothing Then
                '    flag_YoungPlan = True
                '    sql = "SELECT planid FROM ID_Plan WHERE TPlanID=36 AND seq='" & dr("seq") & "'"
                '    dt = DbAccess.GetDataTable(sql, objconn)
                '    For i As Integer = 0 To dt.Rows.Count - 1
                '        If i = dt.Rows.Count - 1 Then
                '            YoungPlan_Planid += "'" & dt.Rows(i).Item("Planid").ToString() & "' " & vbCrLf
                '        Else
                '            YoungPlan_Planid += "'" & dt.Rows(i).Item("Planid").ToString() & "', " & vbCrLf
                '        End If
                '    Next
                '    sql = "SELECT * FROM  Plan_PlanInfo WHERE PlanID in (" & YoungPlan_Planid & ") "
                '    dr = DbAccess.GetOneRow(sql, objconn)
                '    If Not dr Is Nothing Then
                '        Dim msg As String = ""
                '        msg += "此計劃已與其它轄區計畫參照無法刪除" & vbCrLf
                '        msg += "若要刪除請連絡系統管理者"
                '        Common.MessageBox(Me, msg)
                '        Exit Sub
                '    End If
                'End If

#End Region

                '開始刪除計畫----------   Start
                sql = "DELETE ID_Plan WHERE PlanID = '" & e.CommandArgument & "' "
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功")
                Call Search1()
                'bt_search_Click(bt_search, e)
                '開始刪除計畫----------   End
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button1") '修改
                Dim btn2 As Button = e.Item.FindControl("Button2") '刪除
                Dim btnView As Button = e.Item.FindControl("btnView") '查詢賦予帳號

                If flag_ROC Then e.Item.Cells(2).Text = TIMS.Cdate17(drv("SDate"))  '(將原先的西年日期改為民國日期，by:20180928、20181001)
                If flag_ROC Then e.Item.Cells(3).Text = TIMS.Cdate17(drv("EDate"))  '(將原先的西年日期改為民國日期，by:20180928、20181001)

                btn1.Attributes("onclick") = "but_edit(" & drv("PlanID").ToString & ");"
                btn2.Attributes("onclick") = "return confirm('確定是否要刪除?');"
                btn1.CommandArgument = drv("PlanID").ToString
                btn2.CommandArgument = drv("PlanID").ToString
                btnView.CommandArgument = drv("PlanID").ToString
                btn1.Enabled = True
                btn2.Enabled = True
                btnView.Enabled = True
        End Select

#Region "(No Use)"

        'btn1.Enabled = False
        'If au.blnCanMod Then btn1.Enabled = True
        'btn2.Enabled = False
        'If au.blnCanDel Then btn2.Enabled = True
        'btnView.Enabled = False
        'If au.blnCanDel Then btnView.Enabled = True

        '青年職涯啟動計劃，只有泰山能修改、刪除
        'If drv("TPlanID") = 36 And sm.UserInfo.DistID <> "002" Then
        '    btn1.Enabled = False
        '    btn1.ToolTip = "青年職涯啟動計劃，只有轄區泰山能修改"
        '    btn2.Enabled = False
        '    btn2.ToolTip = "青年職涯啟動計劃，只有轄區泰山能刪除"
        'End If

#End Region
        'If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then End If
    End Sub

    Function GetDistIDValue() As String
        Dim rst As String = ""
        For i As Integer = 0 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If rst <> "" Then rst += ","
                rst += "'" & Me.DistID.Items(i).Value & "'"
            End If
        Next
        Return rst
    End Function

    '年度選擇後重新搜尋計畫 SQL
    Sub Show_KeyPlan()
        If yearlist.SelectedValue = "" Then
            Common.MessageBox(Me, "請選擇有效年度!!")
            Exit Sub
        End If
        Dim sqlstr As String = ""
        '含不啟用的計畫
        If cbk1.Checked Then
            sqlstr = " SELECT TPLANID, PLANNAME + CASE WHEN CLSYEAR IS NULL OR CLSYEAR > '" & yearlist.SelectedValue & "' then '' else '…已停用' + CLSYEAR END PLANNAME FROM KEY_PLAN ORDER BY TPLANID " & vbCrLf
        Else
            sqlstr = " SELECT TPLANID, PLANNAME FROM KEY_PLAN WHERE (Clsyear is null or Clsyear > '" & yearlist.SelectedValue & "') ORDER BY TPlanID "
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
        With planlist
            .DataSource = dt
            .DataTextField = "PlanName"
            .DataValueField = "TPlanID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        Call Show_KeyPlan()  '年度選擇後重新搜尋計畫 SQL
        Common.SetListItem(planlist, sm.UserInfo.TPlanID)
        Call ClearDataGrid2()
    End Sub

    Sub ClearDataGrid2()
        Dim dt As DataTable
        Dim sql As String
        sql = "" & vbCrLf
        sql += " SELECT b.Name + '['+a.Account + ']' 賦予的帳號列表 ,o.orgname 機構名稱 ,c.RSID 機構業務ID ,c.RID RID " & vbCrLf
        sql += " FROM Auth_AccRWPlan a " & vbCrLf
        sql += " JOIN auth_account b ON a.Account = b.Account " & vbCrLf
        sql += " JOIN Auth_Relship c ON a.rid = c.rid " & vbCrLf
        sql += " JOIN org_orginfo o ON c.orgid = o.orgid " & vbCrLf
        sql += " WHERE 1<>1 " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid2.DataSource = dt.DefaultView
        DataGrid2.DataBind()
        If dt.Rows.Count = 0 Then DataGrid2.Visible = False
    End Sub

    Private Sub bt_search2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search2.Click
        Call Show_KeyPlan()
    End Sub

    '匯入年度計畫
    Protected Sub btn_ImpYears_Click(sender As Object, e As EventArgs) Handles btn_ImpYears.Click
        KeepSearch()
        TIMS.Utl_Redirect1(Me, "TC_01_001_import.aspx?ID=" & Request("ID"))
    End Sub
End Class