Partial Class RWB_01_004
    Inherits AuthBasePage

    ''TB_QA
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        '處理[分頁設定元件]出現的時機
        If PageControler1.PageDataGrid.Items.Count > 0 Then
            PageControler1.Visible = True
        Else
            PageControler1.Visible = False
        End If

        If Not IsPostBack Then
            Call sCreate1() '頁面初始化

            '20190103 若有先前查詢條件記錄，則將資料重新讀取到頁面中
            If Not Session("_SearchStr") Is Nothing Then
                schC_SDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE1")
                schC_SDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE2")
                rblType.SelectedIndex = rblType.Items.IndexOf(rblType.Items.FindByValue(TIMS.GetMyValue(Session("_SearchStr"), "rblType")))
                schC_EDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE1")
                schC_EDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE2")
                rblUse.SelectedIndex = rblUse.Items.IndexOf(rblUse.Items.FindByValue(TIMS.GetMyValue(Session("_SearchStr"), "rblUse")))
                txtKeyword.Text = TIMS.GetMyValue(Session("_SearchStr"), "txtKeyword")
                sSearch1()
                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    '頁面初始化
    Sub sCreate1()
        schC_SDATE1.Text = ""
        schC_SDATE2.Text = ""
        rblType.SelectedValue = "1"
        schC_EDATE1.Text = ""
        schC_EDATE2.Text = ""
        rblUse.SelectedValue = "1"
        txtKeyword.Text = ""
    End Sub

    '資料查詢
    Sub sSearch1()
        schC_SDATE1.Text = TIMS.ClearSQM(schC_SDATE1.Text).Trim
        schC_SDATE2.Text = TIMS.ClearSQM(schC_SDATE2.Text).Trim
        schC_EDATE1.Text = TIMS.ClearSQM(schC_EDATE1.Text).Trim
        schC_EDATE2.Text = TIMS.ClearSQM(schC_EDATE2.Text).Trim
        Dim vITEM1 As String = TIMS.ClearSQM(rblType.SelectedValue)
        Dim vITEM2 As String = TIMS.ClearSQM(rblUse.SelectedValue)
        txtKeyword.Text = TIMS.ClearSQM(txtKeyword.Text).Trim

        Dim schCSDATE1 As String = schC_SDATE1.Text.Trim
        Dim schCSDATE2 As String = schC_SDATE2.Text.Trim
        Dim schCEDATE1 As String = schC_EDATE1.Text.Trim
        Dim schCEDATE2 As String = schC_EDATE2.Text.Trim

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.QAID DESC) ROWNUM " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.START_DATE, 111) CSDATE " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.END_DATE, 111) CEDATE " & vbCrLf
        sql &= "        ,a.QAID " & vbCrLf
        sql &= "        ,a.TYPEID " & vbCrLf
        sql &= "        ,CASE WHEN a.TYPEID = '1' THEN '產業人才投資方案' " & vbCrLf
        sql &= "              WHEN a.TYPEID = '2' THEN '自辦在職訓練' " & vbCrLf
        sql &= "              WHEN a.TYPEID = '3' THEN '企業委託訓練' " & vbCrLf
        sql &= "              WHEN a.TYPEID = '4' THEN '充電起飛' " & vbCrLf
        sql &= "              WHEN a.TYPEID = '5' THEN '網站操作問題' " & vbCrLf
        sql &= "              ELSE '' END C_TYPE " & vbCrLf
        sql &= "        ,a.START_DATE " & vbCrLf
        sql &= "        ,a.END_DATE " & vbCrLf
        sql &= "        ,a.QUESTION " & vbCrLf
        sql &= "        ,CASE WHEN LEN(a.QUESTION) > 15 THEN SUBSTRING(a.QUESTION, 1, 15) + '...' ELSE a.QUESTION END QUESTION1 " & vbCrLf
        sql &= "        ,a.ANSWER " & vbCrLf
        sql &= "        ,a.ISUSED " & vbCrLf
        sql &= "        ,CASE WHEN a.ISUSED = 'Y' THEN '啟用' " & vbCrLf
        sql &= "              WHEN a.ISUSED = 'N' THEN '停用' " & vbCrLf
        sql &= "              ELSE '' END C_ISUSED " & vbCrLf
        sql &= "        ,a.MODIFYACCT " & vbCrLf
        sql &= "        ,a.MODIFYDATE " & vbCrLf
        sql &= " FROM TB_QA a " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf

        If schCSDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.START_DATE, 111) >= @schCSDATE1 " & vbCrLf
        If schCSDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.START_DATE, 111) <= @schCSDATE2 " & vbCrLf
        If schCEDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.END_DATE, 111) >= @schCEDATE1 " & vbCrLf
        If schCEDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.END_DATE, 111) <= @schCEDATE2 " & vbCrLf
        If vITEM1 <> "" Then sql &= " AND a.TYPEID = @TYPEID " & vbCrLf
        If vITEM2 <> "" Then sql &= " AND a.ISUSED = @ISUSED " & vbCrLf
        If txtKeyword.Text <> "" Then sql &= " AND (a.QUESTION LIKE @keyword OR a.ANSWER LIKE @keyword) " & vbCrLf

        sql &= " ORDER BY a.MODIFYDATE DESC " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If schCSDATE1 <> "" Then parms.Add("schCSDATE1", IIf(flag_ROC, TIMS.Cdate18(schCSDATE1), schCSDATE1))  'edit，by:20181019
        If schCSDATE2 <> "" Then parms.Add("schCSDATE2", IIf(flag_ROC, TIMS.Cdate18(schCSDATE2), schCSDATE2))  'edit，by:20181019
        If schCEDATE1 <> "" Then parms.Add("schCEDATE1", IIf(flag_ROC, TIMS.Cdate18(schCEDATE1), schCEDATE1))  'edit，by:20181019
        If schCEDATE2 <> "" Then parms.Add("schCEDATE2", IIf(flag_ROC, TIMS.Cdate18(schCEDATE2), schCEDATE2))  'edit，by:20181019
        If vITEM1 <> "" Then parms.Add("TYPEID", vITEM1)
        If vITEM2 <> "" Then parms.Add("ISUSED", vITEM2)
        If txtKeyword.Text <> "" Then parms.Add("keyword", ("%" + txtKeyword.Text + "%"))

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料"
        tb_Sch.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            tb_Sch.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '進行資料查詢作業
    Protected Sub bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Call sSearch1() '查詢
    End Sub

    '資料新增
    Protected Sub bt_add_Click(sender As Object, e As EventArgs) Handles bt_add.Click
        Dim url1 As String = "RWB_01_004_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(2).Text = IIf(flag_ROC, TIMS.Cdate17(drv("CSDATE")), drv("CSDATE"))  'edit，by:20181019
                e.Item.Cells(3).Text = IIf(flag_ROC, TIMS.Cdate17(drv("CEDATE")), drv("CEDATE"))  'edit，by:20181019
                Dim btnEDIT1 As Button = e.Item.FindControl("btnEDIT1")
                Dim btnDEL1 As Button = e.Item.FindControl("btnDEL1")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "QAID", TIMS.CStr1(drv("QAID")))
                btnEDIT1.CommandArgument = sCmdArg
                btnDEL1.CommandArgument = sCmdArg
                btnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg2
        End Select
    End Sub

    'DataGrid1功能事件
    Protected Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        hid_V.Value = TIMS.GetMyValue(sCmdArg, "QAID")
        If hid_V.Value = "" Then Exit Sub
        hid_V.Value = TIMS.ClearSQM(hid_V.Value)

        Select Case e.CommandName
            Case "edit"
                GetSearchStr()  'edit，by:20190103
                Dim url1 As String = "RWB_01_004_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
                url1 &= "&A=E&QAID=" & hid_V.Value
                url1 &= "&SEQNO_E=" & TIMS.EncryptAes(hid_V.Value)
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " UPDATE TB_QA " & vbCrLf
                sql &= " SET ISUSED = 'N' " & vbCrLf
                sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
                sql &= " ,MODIFYACCT = @UACCT " & vbCrLf
                sql &= " WHERE QAID = @QAID " & vbCrLf
                Dim parms As Hashtable = New Hashtable()
                parms.Add("UACCT", sm.UserInfo.UserID)
                parms.Add("QAID", hid_V.Value)
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
                Common.MessageBox(Me, "刪除成功")
                Call sSearch1()
            Case Else
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
        End Select
    End Sub

    '20190103 將目前的查詢條件儲存起來
    Sub GetSearchStr()
        Session("_SearchStr") = "prg=RWB_01_004"
        Session("_SearchStr") &= "&schC_SDATE1=" & schC_SDATE1.Text
        Session("_SearchStr") += "&schC_SDATE2=" & schC_SDATE2.Text
        Session("_SearchStr") += "&rblType=" & rblType.SelectedValue
        Session("_SearchStr") += "&schC_EDATE1=" & schC_EDATE1.Text
        Session("_SearchStr") += "&schC_EDATE2=" & schC_EDATE2.Text
        Session("_SearchStr") += "&rblUse=" & rblUse.SelectedValue
        Session("_SearchStr") += "&txtKeyword=" & txtKeyword.Text
    End Sub
End Class