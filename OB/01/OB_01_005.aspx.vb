Partial Class OB_01_005
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        PageControler1 = Me.FindControl("PageControler1")

        '檢查Session是否存在 Start
        'If sm.UserInfo.UserID Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');top.location.href='../../MOICA_Login.aspx';</script>")
        '    Response.End()
        'End If
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) '☆
        '檢查Session是否存在 End

        If Not IsPostBack Then
            '載入分署(中心)
            Dim sqlstr As String
            sqlstr = "select * from ID_District "
            If sm.UserInfo.DistID <> "000" Then '系統管理者可查全部
                sqlstr += "where DistID = '" & sm.UserInfo.DistID & "'"
            End If
            Dim dt As DataTable = DbAccess.GetDataTable(sqlstr)
            With ddl_DistID
                .DataSource = dt
                .DataTextField = "name"
                .DataValueField = "DistID"
                .DataBind()
                .SelectedValue = sm.UserInfo.DistID
            End With
            txt_DistID.Text = dt.Rows(0)("name")
        End If
        '載入選擇頁
        PageControler1.PageDataGrid = dg_Sch
        If Not IsPostBack Then
            If Not Session("_SearchStr") Is Nothing Then
                Dim MyArray As Array
                Dim MyItem As String
                Dim MyValue As String
                MyArray = Split(Session("_SearchStr"), "&")
                For i As Integer = 0 To MyArray.Length - 1
                    MyItem = Split(MyArray(i), "=")(0)
                    MyValue = Split(MyArray(i), "=")(1)
                    Select Case MyItem
                        Case "PageIndex"
                            PageControler1.PageIndex = MyValue
                    End Select
                Next
                Panel_View.Visible = False
                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        search()
    End Sub
    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        Dim sql As String
        Dim dt As DataTable
        sql = "select a.ORSN,a.ORName,(select count(ORSN) from OB_ReviewItem b where a.ORSN=b.ORParent) num,ORAvail from OB_ReviewItem "
        sql += "a where a.ORLevels=0 and a.DistID='" & ddl_DistID.SelectedValue & "' "
        If Len(txt_ORName.Text) > 0 Then
            sql += " and ORName like '%" & txt_ORName.Text & "%'"
        End If
        dt = DbAccess.GetDataTable(sql)
        If TIMS.Get_SQLRecordCount(sql) > 0 Then
            msg.Visible = False
            Panel_View.Visible = True
            PageControler1.SqlString = sql
            PageControler1.ControlerLoad()
        Else
            msg.Visible = True
            Panel_View.Visible = False
        End If
    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound

        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            Dim btnview As Button = e.Item.FindControl("btn_view")
            Dim btnedit As Button = e.Item.FindControl("btn_edit")


            e.Item.Cells(0).Text = e.Item.ItemIndex + 1
            If drv("ORAvail") = "Y" Then
                e.Item.Cells(3).Text = "是"
            Else
                e.Item.Cells(3).Text = "否"
            End If

            btnview.CommandArgument = drv("ORSN")
            btnedit.CommandArgument = drv("ORSN")

            'Dim btndel As Button = e.Item.FindControl("btn_del")
            'If drv("num") <> "0" Then
            '    btndel.Enabled = False
            'End If
            'btndel.CommandArgument = drv("ORSN")
            'btndel.Attributes("onclick") = "return confirm('確定要刪除第" & e.Item.Cells(0).Text & "筆-" & drv("num") & "筆資料?');"

        End If
    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Dim sql As String
        'Dim dt As DataTable
        Dim dr As DataRow
        Select Case e.CommandName
            Case "view"
                search2(e.CommandArgument)
                Panel_View2.Visible = True
                Panel_View.Visible = False
                Panel_Sch.Visible = False
            Case "edit"
                Me.ViewState("flag") = "edit"
                Me.ViewState("ORSN") = e.CommandArgument
                Panel_Add_Edit.Visible = True
                Panel_View.Visible = False
                Panel_Sch.Visible = False
                sql = "select ORName,ORAvail from OB_ReviewItem where ORSN=" & e.CommandArgument
                dr = DbAccess.GetOneRow(sql)
                txt_Name.Text = dr("ORName")
                If dr("ORAvail") = "Y" Then
                    ckb_ORAvail.Checked = True
                Else
                    ckb_ORAvail.Checked = False
                End If
            Case "del"
                sql = "delete OB_ReviewItem where ORSN=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql)
                Page.RegisterStartupScript("del", "<script>alert('刪除成功!');</script>")
                search()
        End Select
    End Sub

    Sub search2(ByVal ORParent As String)
        Dim sql As String
        Dim dt As DataTable
        Me.ViewState("ORParent") = ORParent
        sql = "select ORSN,ORName,ORAvail from OB_ReviewItem where ORParent=" & ORParent
        dt = DbAccess.GetDataTable(sql)
        If dt.Rows.Count = 0 Then
            msg2.Visible = True
            dg_Sch2.Visible = False
        Else
            msg2.Visible = False
            dg_Sch2.Visible = True
            dg_Sch2.DataSource = dt
            dg_Sch2.DataBind()
        End If

    End Sub

    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        Me.ViewState("flag") = "add"
        Panel_Add_Edit.Visible = True
        Panel_View.Visible = False
        Panel_Sch.Visible = False
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If txt_Name.Text = "" Then
            Common.MessageBox(Me, "請填寫【評選項目名稱】")
            Exit Sub
        Else
            Dim sql As String
            Dim dt As DataTable
            sql = "select ORName from OB_ReviewItem where ORName='" & txt_Name.Text & "'"
            If Me.ViewState("ORSN") <> "" Then
                sql += " and ORSN<>" & Me.ViewState("ORSN")
            End If

            dt = DbAccess.GetDataTable(sql)
            If dt.Rows.Count > 0 Then
                Common.MessageBox(Me, "【評選項目名稱】重複")
                dt.Clear()
            Else
                Dim ORAvail As String
                If ckb_ORAvail.Checked = True Then
                    ORAvail = "Y"
                Else
                    ORAvail = "N"
                End If
                If Me.ViewState("flag") = "add" Then
                    sql = "insert into OB_ReviewItem(ORName,ORLevels,ORAvail,DistID,ModifyAcct,ModifyDate) values('" & txt_Name.Text & "',0,'"
                    sql += ORAvail & "','" & sm.UserInfo.DistID & "','" & sm.UserInfo.UserID & "' ,getdate() ) "
                    ' & FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4) '★ 改成取getdate()
                    DbAccess.ExecuteNonQuery(sql)
                    Common.MessageBox(Me, "存檔成功!")
                    txt_Name.Text = ""
                    ckb_ORAvail.Checked = True

                    Dim strScript As String
                    strScript = "<script language=""javascript"">" + vbCrLf
                    strScript += "if (window.confirm('資料新增成功!\n請問是否繼續新增？')){" + vbCrLf
                    strScript += " " + vbCrLf
                    strScript += "} else {;" + vbCrLf
                    strScript += "  document.getElementById('btn_lev').click();" + vbCrLf
                    strScript &= "}" & vbCrLf
                    strScript &= "</script>"
                    Page.RegisterStartupScript("ring", strScript)
                Else
                    Dim da As SqlDataAdapter = Nothing
                    Dim conn As New SqlConnection
                    conn = DbAccess.GetConnection
                    sql = "update OB_ReviewItem set ORName='" & txt_Name.Text & "',ORAvail='" & ORAvail & "',ModifyAcct='"
                    sql += sm.UserInfo.UserID & "',ModifyDate=getdate() " 'FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4) '★ 改成取getdate()
                    sql += "where ORSN=" & Me.ViewState("ORSN")
                    dt = DbAccess.GetDataTable(sql, da, conn)
                    DbAccess.UpdateDataTable(dt, da)
                    Common.MessageBox(Me, "修改成功!")
                    btn_lev_Click(Me, e)
                    search()
                End If
            End If
        End If
    End Sub

    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        Panel_Add_Edit.Visible = False
        Panel_Sch.Visible = True
        If Me.ViewState("flag") = "edit" Then
            Panel_View.Visible = True
        End If
        Me.ViewState("flag") = ""
        Me.ViewState("ORSN") = ""
        txt_Name.Text = ""
        ckb_ORAvail.Checked = True
    End Sub

    Private Sub dg_Sch2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch2.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            Dim btnedit2 As Button = e.Item.FindControl("btn_edit2")
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1
            If drv("ORAvail") = "Y" Then
                e.Item.Cells(2).Text = "是"
            Else
                e.Item.Cells(2).Text = "否"
            End If
            btnedit2.CommandArgument = drv("ORSN")

            'Dim btndel2 As Button = e.Item.FindControl("btn_del2")
            'btndel2.CommandArgument = drv("ORSN")
            'btndel2.Attributes("onclick") = "return confirm('確定要刪除第" & e.Item.Cells(0).Text & "筆資料?');"
        End If
    End Sub

    Private Sub dg_Sch2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch2.ItemCommand
        Dim sql As String
        'Dim dt As DataTable
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit"
                Panel_View2.Visible = False
                Panel_Add_Edit2.Visible = True
                sql = "select ORName,ORAvail from OB_ReviewItem where ORSN=" & e.CommandArgument
                dr = DbAccess.GetOneRow(sql)
                txt_Name2.Text = dr("ORName")
                If dr("ORAvail") = "Y" Then
                    ckb_ORAvail2.Checked = True
                Else
                    ckb_ORAvail2.Checked = False
                End If
                Me.ViewState("ORSN") = e.CommandArgument
                Me.ViewState("flag") = "edit"
            Case "del"
                sql = "delete OB_ReviewItem where ORSN=" & e.CommandArgument
                DbAccess.ExecuteNonQuery(sql)
                Page.RegisterStartupScript("del", "<script>alert('刪除成功!');</script>")
                search2(Me.ViewState("ORParent"))
        End Select
    End Sub

    Private Sub btn_back2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back2.Click
        Me.ViewState("ORParent") = ""
        Panel_View2.Visible = False
        Panel_View.Visible = True
        Panel_Sch.Visible = True
        search()
    End Sub

    Private Sub btn_Add2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add2.Click
        Panel_View2.Visible = False
        Panel_Add_Edit2.Visible = True
        Me.ViewState("flag") = "add"
    End Sub

    Private Sub btn_save2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save2.Click
        If txt_Name2.Text = "" Then
            Common.MessageBox(Me, "請填寫【評選細項名稱】")
            Exit Sub
        Else
            Dim sql As String
            Dim dt As DataTable
            sql = "select ORName from OB_ReviewItem where ORName='" & txt_Name2.Text & "'"
            If Me.ViewState("ORSN") <> "" Then
                sql += " and ORSN<>" & Me.ViewState("ORSN")
            End If

            dt = DbAccess.GetDataTable(sql)
            If dt.Rows.Count > 0 Then
                Common.MessageBox(Me, "【評選項目名稱】重複")
                dt.Clear()
            Else
                Dim ORAvail2 As String
                If ckb_ORAvail2.Checked = True Then
                    ORAvail2 = "Y"
                Else
                    ORAvail2 = "N"
                End If
                If Me.ViewState("flag") = "add" Then
                    sql = "insert into OB_ReviewItem(ORName,ORParent,ORLevels,ORAvail,DistID,ModifyAcct,ModifyDate) values('"
                    sql += txt_Name2.Text & "'," & Me.ViewState("ORParent") & ",1,'" & ORAvail2 & "','" & sm.UserInfo.DistID & "','"
                    sql += sm.UserInfo.UserID & "',getdate())"
                    DbAccess.ExecuteNonQuery(sql)
                    Common.MessageBox(Me, "存檔成功!")
                    txt_Name2.Text = ""
                    ckb_ORAvail2.Checked = True
                    search2(Me.ViewState("ORParent"))
                    Dim strScript As String
                    strScript = "<script language=""javascript"">" + vbCrLf
                    strScript += "if (window.confirm('資料新增成功!\n請問是否繼續新增？')){" + vbCrLf
                    strScript += " " + vbCrLf
                    strScript += "} else {;" + vbCrLf
                    strScript += "  document.getElementById('btn_lev2').click();" + vbCrLf
                    strScript &= "}" & vbCrLf
                    strScript &= "</script>"
                    Page.RegisterStartupScript("ring", strScript)
                Else
                    Dim da As SqlDataAdapter = Nothing
                    Dim conn As New SqlConnection
                    conn = DbAccess.GetConnection
                    sql = "update OB_ReviewItem set ORName='" & txt_Name2.Text & "',ORAvail='" & ORAvail2 & "',ModifyAcct='"
                    sql += sm.UserInfo.UserID & "',ModifyDate=getdate() "
                    sql += "where ORSN=" & Me.ViewState("ORSN")
                    dt = DbAccess.GetDataTable(sql, da, conn)
                    DbAccess.UpdateDataTable(dt, da)
                    Common.MessageBox(Me, "修改成功!")
                    btn_lev2_Click(Me, e)
                    search2(Me.ViewState("ORParent"))
                End If
            End If
        End If
    End Sub

    Private Sub btn_lev2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev2.Click
        Me.ViewState("flag") = ""
        Me.ViewState("ORSN") = ""
        Panel_View2.Visible = True
        Panel_Add_Edit2.Visible = False
        txt_Name2.Text = ""
        ckb_ORAvail2.Checked = True
    End Sub
End Class
