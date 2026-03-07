Public Class SYS_05_003
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not Page.IsPostBack Then
            labmsg.Text = ""
            Call sUtl_Cancel1()
            tbSch.Visible = True

            Call initObj()
        End If
    End Sub

    '頁籤說明及輸入
    Function Get_ddlTabNum(ByVal obj As DropDownList) As DropDownList
        '<asp@ListItem Value="6">報名資料維護</asp@ListItem>
        '<asp@ListItem Value="1">開班資料查詢</asp@ListItem>
        '<asp@ListItem Value="2">線上報名</asp@ListItem>
        '<asp@ListItem Value="3">線上報名查詢</asp@ListItem>
        '<asp@ListItem Value="4">補助金申請查詢</asp@ListItem>
        With obj
            .Items.Clear()
            .Items.Add(New ListItem("==請選擇==", ""))

            .Items.Add(New ListItem("報名資料維護", "6"))
            .Items.Add(New ListItem("開班資料查詢", "1"))
            .Items.Add(New ListItem("線上報名", "2"))
            .Items.Add(New ListItem("線上報名查詢", "3"))
            .Items.Add(New ListItem("補助金申請查詢", "4"))
        End With
        Return obj
    End Function

    '功能第一次載入初始化
    Sub initObj()
        'Call ListClass.crtDropDownList("org_master", ddlQoyc_status)
        ddlQTabNum = Get_ddlTabNum(ddlQTabNum)
        TabNum = Get_ddlTabNum(TabNum)
    End Sub

    '取消
    Sub sUtl_Cancel1()
        tbSch.Visible = False
        tbList.Visible = False
        tbEdit.Visible = False
    End Sub

    '清除值(及狀態設定)
    Sub clsValue()
        TabNum.Enabled = True

        HidHN2ID.Value = ""

        TabNum.SelectedIndex = -1
        Seqno.Text = ""
        Subject.Text = ""
        PostDate.Text = ""
        'ShowNews.SelectedValue = "N"
        Common.SetListItem(ShowNews, "N")
    End Sub

    '記錄查詢條件 
    Sub Search1Value()
        '記錄查詢條件
        If txtQSeqno.Text <> "" Then txtQSeqno.Text = Trim(txtQSeqno.Text)

        ViewState("TABNUM") = ddlQTabNum.SelectedValue
        ViewState("SEQNO") = txtQSeqno.Text
    End Sub

    '查詢
    Sub Search1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.HN2ID  /*PK*/ " & vbCrLf
        sql += " ,a.TABNUM" & vbCrLf
        sql += " ,dbo.DECODE12(a.TABNUM,6,'報名資料維護',1,'開班資料查詢',2,'線上報名',3,'線上報名查詢',4,'補助金申請查詢','') TABNUM2" & vbCrLf
        sql += " ,a.SEQNO" & vbCrLf
        sql += " ,a.SUBJECT" & vbCrLf
        sql += " ,CONVERT(varchar, a.POSTDATE, 111) POSTDATE" & vbCrLf
        sql += " ,a.SHOWNEWS" & vbCrLf
        sql += " ,a.CREATEACCT" & vbCrLf
        sql += " ,a.CREATEDATE" & vbCrLf
        sql += " ,a.MODIFYACCT" & vbCrLf
        sql += " ,a.MODIFYDATE" & vbCrLf
        sql += " ,a.ISDELETE" & vbCrLf
        sql += " ,a.STOPACCT" & vbCrLf
        sql += " ,a.STOPDATE" & vbCrLf
        sql += " ,ac.NAME MODIFYNAME" & vbCrLf
        sql += " FROM HOME_NEWS2 a" & vbCrLf
        sql += " JOIN AUTH_ACCOUNT ac on ac.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND a.ISDELETE IS NULL" & vbCrLf '顯示未刪除的資料
        If Convert.ToString(ViewState("TABNUM")) <> "" Then
            sql += " AND a.TABNUM=@TABNUM" & vbCrLf
        End If
        If Convert.ToString(ViewState("SEQNO")) <> "" Then
            sql += " AND a.SEQNO=@SEQNO" & vbCrLf
        End If
        sql += " ORDER BY a.TABNUM,a.SEQNO" & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            If Convert.ToString(ViewState("TABNUM")) <> "" Then
                .Parameters.Add("TABNUM", SqlDbType.VarChar).Value = ViewState("TABNUM")
            End If
            If Convert.ToString(ViewState("SEQNO")) <> "" Then
                .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = ViewState("SEQNO")
            End If
            dt.Load(.ExecuteReader())
        End With

        labmsg.Text = "查無資料"
        tbList.Visible = False
        If dt.Rows.Count > 0 Then
            'CPdt = dt.Copy()
            labmsg.Text = ""
            tbList.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Sub loadData()
        If HidHN2ID.Value = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT " & vbCrLf
        sql += " a.HN2ID  /*PK*/ " & vbCrLf
        sql += " ,a.TABNUM" & vbCrLf
        sql += " ,dbo.DECODE12(a.TABNUM,6,'報名資料維護',1,'開班資料查詢',2,'線上報名',3,'線上報名查詢',4,'補助金申請查詢','') TABNUM2" & vbCrLf
        sql += " ,a.SEQNO" & vbCrLf
        sql += " ,a.SUBJECT" & vbCrLf
        sql += " ,CONVERT(varchar, a.POSTDATE, 111) POSTDATE" & vbCrLf
        sql += " ,a.SHOWNEWS" & vbCrLf
        sql += " ,a.CREATEACCT" & vbCrLf
        sql += " ,a.CREATEDATE" & vbCrLf
        sql += " ,a.MODIFYACCT" & vbCrLf
        sql += " ,a.MODIFYDATE" & vbCrLf
        'sql += " ,a.ISDELETE" & vbCrLf
        'sql += " ,a.STOPACCT" & vbCrLf
        'sql += " ,a.STOPDATE" & vbCrLf
        'sql += " ,ac.NAME MODIFYNAME" & vbCrLf
        sql += " FROM HOME_NEWS2 a" & vbCrLf
        'sql += " JOIN AUTH_ACCOUNT ac on ac.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        'sql += " AND a.ISDELETE IS NULL" & vbCrLf
        sql += " AND a.HN2ID=@HN2ID " & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("HN2ID", SqlDbType.Int).Value = Val(HidHN2ID.Value)
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)

            Me.HidHN2ID.Value = Convert.ToString(dr("HN2ID"))
            If Convert.ToString(dr("TabNum")) <> "" Then
                Common.SetListItem(TabNum, dr("TabNum"))
            End If
            Seqno.Text = Convert.ToString(dr("Seqno"))
            Subject.Text = Convert.ToString(dr("Subject"))
            PostDate.Text = TIMS.Cdate3(dr("PostDate"))
            Select Case Convert.ToString(dr("ShowNews"))
                Case "Y"
                    Common.SetListItem(ShowNews, dr("ShowNews"))
                Case Else
                    Common.SetListItem(ShowNews, "N")
            End Select

        End If
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Seqno.Text = TIMS.ClearSQM(Seqno.Text)
        PostDate.Text = TIMS.ClearSQM(PostDate.Text)

        If TabNum.SelectedValue = "" Then
            Errmsg += "請選擇 頁籤項目" & vbCrLf
        End If
        If Seqno.Text = "" Then
            Errmsg += "請輸入 內部序號" & vbCrLf
        End If
        If PostDate.Text = "" Then
            Errmsg += "請輸入 發布日期" & vbCrLf
        End If
        If Subject.Text = "" Then
            Errmsg += "請輸入 發布主題" & vbCrLf
        End If

        If Seqno.Text <> "" Then
            If Not TIMS.IsInt(Seqno.Text) Then
                Errmsg += "內部序號 請輸入正整數" & vbCrLf
            End If
        End If

        If PostDate.Text <> "" Then
            If Not TIMS.IsDate1(PostDate.Text) Then
                Errmsg += "發布日期 請輸入正確日期格式" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Sub SaveData1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim rst As Integer = 0

        Call TIMS.OpenDbConn(objconn)
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        'Dim sql As String = ""
        'Dim dt As DataTable = Nothing
        'Dim dr As DataRow = Nothing

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " INSERT INTO HOME_NEWS2 ( " & vbCrLf
        sql += " HN2ID" & vbCrLf '/*PK*/ 
        sql += " ,TABNUM" & vbCrLf
        sql += " ,SEQNO" & vbCrLf
        sql += " ,SUBJECT" & vbCrLf
        sql += " ,POSTDATE" & vbCrLf
        sql += " ,SHOWNEWS" & vbCrLf
        sql += " ,CREATEACCT" & vbCrLf
        sql += " ,CREATEDATE" & vbCrLf
        sql += " ,MODIFYACCT" & vbCrLf
        sql += " ,MODIFYDATE" & vbCrLf
        'sql += " ,ISDELETE" & vbCrLf
        'sql += " ,STOPACCT" & vbCrLf
        'sql += " ,STOPDATE" & vbCrLf
        sql += " ) VALUES (" & vbCrLf
        sql += " @HN2ID" & vbCrLf
        sql += " ,@TABNUM" & vbCrLf
        sql += " ,@SEQNO" & vbCrLf
        sql += " ,@SUBJECT" & vbCrLf
        sql += " ,@POSTDATE" & vbCrLf
        sql += " ,@SHOWNEWS" & vbCrLf
        sql += " ,@CREATEACCT" & vbCrLf
        sql += " ,getdate()" & vbCrLf
        sql += " ,@MODIFYACCT" & vbCrLf
        sql += " ,getdate()" & vbCrLf
        'sql += " ,ISDELETE" & vbCrLf
        'sql += " ,STOPACCT" & vbCrLf
        'sql += " ,STOPDATE" & vbCrLf
        sql += " ) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " UPDATE HOME_NEWS2" & vbCrLf
        'sql += " SET HN2ID" & vbCrLf '/*PK*/ 
        'sql += " ,TABNUM" & vbCrLf
        sql += " SET SEQNO=@SEQNO" & vbCrLf
        sql += " ,SUBJECT=@SUBJECT" & vbCrLf
        sql += " ,POSTDATE=@POSTDATE" & vbCrLf
        sql += " ,SHOWNEWS=@SHOWNEWS" & vbCrLf
        'sql += " ,CREATEACCT" & vbCrLf
        'sql += " ,CREATEDATE" & vbCrLf
        sql += " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql += " ,MODIFYDATE=getdate()" & vbCrLf
        'sql += " ,ISDELETE" & vbCrLf
        'sql += " ,STOPACCT" & vbCrLf
        'sql += " ,STOPDATE" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND HN2ID=@HN2ID" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        '新增重複判斷
        sql = "" & vbCrLf
        sql += " SELECT 'X' FROM HOME_NEWS2" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND ISDELETE IS NULL" & vbCrLf '使用中。
        sql += " AND TABNUM=@TABNUM" & vbCrLf
        sql += " AND SEQNO=@SEQNO" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        '修改重複判斷
        sql = "" & vbCrLf
        sql += " SELECT 'X' FROM HOME_NEWS2" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND ISDELETE IS NULL" & vbCrLf '使用中。
        sql += " AND TABNUM=@TABNUM" & vbCrLf
        sql += " AND SEQNO=@SEQNO" & vbCrLf
        sql += " AND HN2ID!=@HN2ID" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)

        If HidHN2ID.Value = "" Then
            '新增
            Dim dt1 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("TABNUM", SqlDbType.Int).Value = Val(TabNum.SelectedValue)
                .Parameters.Add("SEQNO", SqlDbType.Int).Value = Val(Seqno.Text)
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該頁籤序號已新增，請使用修改功能!!")
                Exit Sub
            End If
        Else
            '修改
            Dim dt1 As New DataTable
            With sCmd2
                .Parameters.Clear()
                .Parameters.Add("TABNUM", SqlDbType.Int).Value = Val(TabNum.SelectedValue)
                .Parameters.Add("SEQNO", SqlDbType.Int).Value = Val(Seqno.Text)
                .Parameters.Add("HN2ID", SqlDbType.Int).Value = Val(HidHN2ID.Value)
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該頁籤序號已存在，請重新輸入!!")
                Exit Sub
            End If
        End If


        If HidHN2ID.Value = "" Then
            '新增
            Dim iHN2ID As Integer = DbAccess.GetNewId(objconn, " HOME_NEWS2_HN2ID_SEQ,HOME_NEWS2,HN2ID")
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("HN2ID", SqlDbType.Int).Value = iHN2ID
                .Parameters.Add("TABNUM", SqlDbType.Int).Value = Val(TabNum.SelectedValue)
                .Parameters.Add("SEQNO", SqlDbType.Int).Value = Val(Seqno.Text)
                .Parameters.Add("SUBJECT", SqlDbType.NVarChar).Value = Subject.Text
                .Parameters.Add("POSTDATE", SqlDbType.DateTime).Value = CDate(PostDate.Text)
                .Parameters.Add("SHOWNEWS", SqlDbType.VarChar).Value = ShowNews.SelectedValue
                .Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                rst = .ExecuteNonQuery()
            End With
        Else
            '修改
            With uCmd
                .Parameters.Clear()
                '.Parameters.Add("TABNUM", SqlDbType.Int).Value = Val(TabNum.SelectedValue)
                .Parameters.Add("SEQNO", SqlDbType.Int).Value = Val(Seqno.Text)
                .Parameters.Add("SUBJECT", SqlDbType.NVarChar).Value = Subject.Text
                .Parameters.Add("POSTDATE", SqlDbType.DateTime).Value = CDate(PostDate.Text)
                .Parameters.Add("SHOWNEWS", SqlDbType.VarChar).Value = ShowNews.SelectedValue
                '.Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                .Parameters.Add("HN2ID", SqlDbType.Int).Value = Val(HidHN2ID.Value)
                rst = .ExecuteNonQuery()
            End With
        End If

        If rst > 0 Then
            Call sUtl_Cancel1()
            tbSch.Visible = True

            Call Search1()
        Else
            Common.MessageBox(Page, "執行完畢，無資料更動!")
        End If
    End Sub

    '刪除
    Sub Delete1()
        If HidHN2ID.Value = "" Then
            Common.MessageBox(Page, "未輸入刪除序號，請重新查詢!")
            Exit Sub
        End If

        Dim rst As Integer = 0
        Dim sql As String = ""
        sql = "DELETE HOME_NEWS2 WHERE HN2ID=@HN2ID "
        Call TIMS.OpenDbConn(objconn)

        Dim dCmd As New SqlCommand(sql, objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("HN2ID", SqlDbType.Int).Value = Val(HidHN2ID.Value)
            rst = .ExecuteNonQuery()
        End With
        If rst = 1 Then
            Common.MessageBox(Page, "刪除成功!")
            Call Search1()
        End If
    End Sub

    '查詢鈕
    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        '記錄查詢條件
        Call Search1Value()

        Call Search1()
    End Sub

    '新增鈕
    Protected Sub btnAdd1_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click
        clsValue()

        Call sUtl_Cancel1()
        tbEdit.Visible = True
    End Sub

    '儲存
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Call SaveData1()
    End Sub

    '回上頁
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Call sUtl_Cancel1()
        tbSch.Visible = True
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "UPD" '修改
                Call sUtl_Cancel1()
                tbEdit.Visible = True

                Call clsValue()
                TabNum.Enabled = False

                Dim sCmdArg As String = Convert.ToString(e.CommandArgument)
                HidHN2ID.Value = TIMS.GetMyValue(sCmdArg, "HN2ID")

                Call loadData()

            Case "DEL" '刪除
                Dim sCmdArg As String = Convert.ToString(e.CommandArgument)
                HidHN2ID.Value = TIMS.GetMyValue(sCmdArg, "HN2ID")

                Call Delete1()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SYS_TD2"
                End If
                '序號
                e.Item.Cells(0).Text = (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + e.Item.ItemIndex + 1

                Dim labTABNUM2 As Label = e.Item.FindControl("labTABNUM2")
                Dim labSeqno As Label = e.Item.FindControl("labSeqno")
                Dim labShowNews As Label = e.Item.FindControl("labShowNews")
                Dim lbtUpdate As LinkButton = e.Item.FindControl("lbtUpdate")
                Dim lbtDelete As LinkButton = e.Item.FindControl("lbtDelete")

                labTABNUM2.Text = Convert.ToString(drv("TABNUM2"))
                labSeqno.Text = Convert.ToString(drv("Seqno"))

                labShowNews.Text = ""
                Select Case Convert.ToString(drv("ShowNews"))
                    Case "Y"
                        labShowNews.Text = "是"
                End Select

                lbtDelete.Attributes.Add("onclick", "return confirm('您確定要刪除第" & e.Item.Cells(0).Text & "筆資料嗎?');")

                Dim sCmdArg As String = ""
                Call TIMS.SetMyValue(sCmdArg, "HN2ID", drv("HN2ID"))
                lbtUpdate.CommandArgument = sCmdArg
                lbtDelete.CommandArgument = sCmdArg

        End Select
    End Sub

    '預覽。
    Protected Sub btnPreview_Click(sender As Object, e As EventArgs) Handles btnPreview.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim sUrl As String = "http://localhost:8651/TIMSonline40f/ShowPreview.aspx?k=j"
        'Dim sUrl As String = "http://tims.etraining.gov.tw/TIMSOnline/ShowPreview.aspx"
        '<add key="ShowOnline" value="http://localhost:8651/TIMSonline40f/ShowPreview.aspx?k=j" />

        Dim Cst_ShowOnline As String = "http://tims.etraining.gov.tw/TIMSOnline/ShowPreview.aspx?k=j"
        Dim sShowOnline As String = TIMS.Utl_GetConfigSet("ShowOnline")
        Dim sUrl As String = ""

        Dim sTmpSubject1 As String = Me.Subject.Text
        If sTmpSubject1.ToLower.Substring(0, 3) = "<p>" Then
            sTmpSubject1 = sTmpSubject1.Substring(3)
        End If
        Dim iLen1 As Integer = sTmpSubject1.Length - 1 - 5
        If iLen1 >= 0 Then '</p>
            If Server.UrlEncode(sTmpSubject1.ToLower.Substring(iLen1)) = "%3c%2fp%3e%0d%0a" Then
                sTmpSubject1 = sTmpSubject1.Substring(0, iLen1)
            End If
        End If

        sUrl = Cst_ShowOnline
        If sShowOnline <> "" Then
            '有設定web config ShowOnline
            sUrl = sShowOnline
        End If
        sUrl &= "&sType2=" & Server.UrlEncode(Me.TabNum.SelectedValue)
        sUrl &= "&ShowMsg=" & Server.UrlEncode(sTmpSubject1)
        sUrl &= "&ShowNews=" & Server.UrlEncode(Me.ShowNews.SelectedValue)

        Common.RespWrite(Me, "<script>window.open('" & sUrl & "','ShowPreview','width=800,height=550,resizable=1,scrollbars=1,status=1');</script>")
    End Sub
End Class