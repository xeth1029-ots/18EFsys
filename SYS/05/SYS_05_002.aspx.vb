Partial Class SYS_05_002
    Inherits AuthBasePage

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
        PageControler1.PageDataGrid = DataGrid1

        '檢查日期格式
        Me.SPostDate.Attributes("onchange") = "check_date();"
        Me.EPostDate.Attributes("onchange") = "check_date();"

        If Not Session("_search") Is Nothing Then
            Dim MyValue As String = ""
            Dim ssSsearch As String = Convert.ToString(Session("_search"))
            MyValue = TIMS.GetMyValue(ssSsearch, "ItemList")
            If MyValue <> "" Then Common.SetListItem(ItemList, MyValue)
            MyValue = TIMS.GetMyValue(ssSsearch, "SPostDate")
            If MyValue <> "" Then Me.SPostDate.Text = MyValue
            MyValue = TIMS.GetMyValue(ssSsearch, "EPostDate")
            If MyValue <> "" Then Me.EPostDate.Text = MyValue
            MyValue = TIMS.GetMyValue(ssSsearch, "PageIndex")
            If MyValue <> "" Then Me.PageControler1.PageIndex = Val(MyValue)
            Session("_search") = Nothing
        End If

        If Not Page.IsPostBack Then
            'Dim dt As DataTable
            'Dim dr As DataRow
            'Dim sqlstr As String
            'sqlstr = "SELECT Type FROM Home_News "
            'dt = DbAccess.GetDataTable(sqlstr, objconn)
            Call create1()
            Call search1()
        End If
    End Sub

    Sub search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim v_RBL_ORDERBY As String = TIMS.GetListValue(RBL_ORDERBY)
        Dim v_RBL_SHOWTYPE As String = TIMS.GetListValue(RBL_SHOWTYPE)

        Dim v_ItemList As String = TIMS.GetListValue(ItemList)
        SPostDate.Text = TIMS.Cdate3(TIMS.ClearSQM(SPostDate.Text))
        EPostDate.Text = TIMS.Cdate3(TIMS.ClearSQM(EPostDate.Text))

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT TOP 300 a.HNID" & vbCrLf
        sql &= " ,a.TYPE" & vbCrLf
        'TYPE_NAME
        sql &= " ,case when a.TYPE=1 then 'News'" & vbCrLf
        sql &= "  when a.TYPE=2 then '新功能'" & vbCrLf
        sql &= "  when a.TYPE=3 then '文件下載'" & vbCrLf
        sql &= "  when a.TYPE=4 then '影音教學'" & vbCrLf
        sql &= "  else '(未設計)' END TYPE_NAME" & vbCrLf
        sql &= " ,a.POSTDATE" & vbCrLf
        sql &= " ,a.SUBJECT" & vbCrLf
        sql &= " ,a.ISSHOW" & vbCrLf
        'SHOW_NAME
        sql &= " ,case when a.ISSHOW='Y' then '是'" & vbCrLf
        sql &= "  when a.ISSHOW='N' then '否'" & vbCrLf
        sql &= "  else '否' END SHOW_NAME" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.POSTFDATE" & vbCrLf
        sql &= " ,a.MSGWEEK" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,case when a.type=3 then replace(a.Subject,'Doc/','../../Doc/') " & vbCrLf
        sql &= "  when a.type=4 then replace(a.Subject,'media/','../../media/') " & vbCrLf
        sql &= "  else a.Subject end Subject2" & vbCrLf
        sql &= " FROM HOME_NEWS a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN AUTH_ACCOUNT b WITH(NOLOCK) ON a.MODIFYACCT=b.ACCOUNT" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If v_RBL_SHOWTYPE <> "" Then sql &= " AND a.ISSHOW='" & v_RBL_SHOWTYPE & "'" & vbCrLf
        'sql &= " AND a.MODIFYDATE >= GETDATE()-1000" & vbCrLf
        If ItemList.SelectedIndex > 0 AndAlso v_ItemList <> "" Then sql &= " AND a.Type = " & v_ItemList & " " & vbCrLf

        If Me.SPostDate.Text <> "" Then sql &= " AND a.PostDate >= " & TIMS.To_date(Me.SPostDate.Text) & vbCrLf

        If Me.EPostDate.Text <> "" Then sql &= " AND a.PostDate <= " & TIMS.To_date(Me.EPostDate.Text) & vbCrLf

        Select Case v_RBL_ORDERBY
            Case "M"
                sql &= " ORDER BY a.MODIFYDATE DESC,a.POSTDATE DESC" & vbCrLf
            Case Else
                'Case "P"
                sql &= " ORDER BY a.POSTDATE DESC,a.MODIFYDATE DESC" & vbCrLf
        End Select

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        Me.msg.Text = "查無資料"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.msg.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True

            Me.DataGrid1.DataKeyField = "HNID"

            'PageControler1.SqlPrimaryKeyDataCreate(sqlstr, "HNID", "PostDate DESC,Type")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "HNID"
            Select Case v_RBL_ORDERBY
                Case "M"
                    PageControler1.Sort = "MODIFYDATE DESC,POSTDATE DESC"
                Case Else
                    PageControler1.Sort = "POSTDATE DESC,MODIFYDATE DESC"
            End Select
            PageControler1.ControlerLoad()
        End If
    End Sub

    Sub create1()
        ItemList.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    ''' <summary>
    ''' 刪除 
    ''' </summary>
    ''' <param name="iHNID"></param>
    Sub DEL_1(ByVal iHNID As Integer)
        Dim Parms As New Hashtable
        Parms.Clear()
        Parms.Add("HNID", iHNID)
        'U/I/D

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE HOME_NEWS" & vbCrLf
        sql &= " SET MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE HNID=@HNID" & vbCrLf 'PK
        DbAccess.ExecuteNonQuery(sql, objconn, Parms)

        sql = "" & vbCrLf
        sql &= " INSERT INTO HOME_NEWS_BAK1(" & vbCrLf
        sql &= " HNID,TYPE,POSTDATE,SUBJECT,ISSHOW,MODIFYACCT,MODIFYDATE,POSTFDATE,MSGWEEK,Doc0,Doc1" & vbCrLf
        sql &= " ) SELECT " & vbCrLf
        sql &= " HNID,TYPE,POSTDATE,SUBJECT,ISSHOW,MODIFYACCT,MODIFYDATE,POSTFDATE,MSGWEEK,Doc0,Doc1" & vbCrLf
        sql &= " FROM HOME_NEWS" & vbCrLf
        sql &= " WHERE HNID=@HNID" & vbCrLf 'PK
        DbAccess.ExecuteNonQuery(sql, objconn, Parms)

        sql = "" & vbCrLf
        sql &= " DELETE HOME_NEWS" & vbCrLf
        sql &= " WHERE HNID=@HNID" & vbCrLf 'PK
        DbAccess.ExecuteNonQuery(sql, objconn, Parms)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "UCmd1"
                Session("_search") = "ItemList=" & Me.ItemList.SelectedValue
                Session("_search") += "&SPostDate=" & Me.SPostDate.Text
                Session("_search") += "&EPostDate=" & Me.EPostDate.Text
                Session("_search") += "&PageIndex=" & Me.DataGrid1.CurrentPageIndex + 1

                Dim url1 As String = ""
                url1 = "SYS_05_002_add.aspx?ID=" & TIMS.Get_MRqID(Me)
                url1 &= "&gptodo=update&HNID=" & DataGrid1.DataKeys(e.Item.ItemIndex)
                '修改 按鈕
                TIMS.Utl_Redirect1(Me, url1)

            Case "DCmd1" '刪除 
                Dim oHNID As Object = DataGrid1.DataKeys(e.Item.ItemIndex)
                If oHNID Is Nothing Then Exit Sub

                Dim iHNID As Integer = Val(oHNID)
                Call DEL_1(iHNID)

                '查詢
                Me.msg.Text = ""
                Call search1()

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        'Dim myTableCell As TableCell
        Dim mylbtUpdate As LinkButton = e.Item.FindControl("lbtUpdate")
        Dim mylbtDelete As LinkButton = e.Item.FindControl("lbtDelete")
        Dim LabType2 As Label = e.Item.FindControl("LabType2")
        Dim LabisShow As Label = e.Item.FindControl("LabisShow")

        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = ""
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""

                '序號
                e.Item.Cells(0).Text = (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + e.Item.ItemIndex + 1
                '項目
                LabType2.Text = Convert.ToString(drv("TYPE_NAME"))

                '永遠顯示
                LabisShow.Text = Convert.ToString(drv("SHOW_NAME"))

                'myTableCell = e.Item.Cells(8)
                'myDeleteButton = myTableCell.Controls(0)
                mylbtDelete.Attributes.Add("onclick", "return confirm('您確定要刪除嗎?');")
        End Select
    End Sub

    Private Sub reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles reset.Click
        'RESET
        Me.SPostDate.Text = ""
        Me.EPostDate.Text = ""
        Me.ItemList.SelectedIndex = 0
        '查詢
        Me.msg.Text = ""
        Call search1()
    End Sub

    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        Dim url1 As String = ""
        url1 = "SYS_05_002_add.aspx?ID=" & TIMS.Get_MRqID(Me)
        url1 &= "&gptodo=add"
        '新增 按鈕
        TIMS.Utl_Redirect1(Me, url1)
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        '查詢
        Me.msg.Text = ""
        Call search1()
    End Sub

End Class
