Partial Class RWB_01_005
    Inherits AuthBasePage

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
                schC_CDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_CDATE1")
                schC_CDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_CDATE2")
                schC_SDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE1")
                schC_SDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE2")
                schCONTENT1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schCONTENT1")
                schC_EDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE1")
                schC_EDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE2")
                sSearch1()
                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    '頁面初始化
    Sub sCreate1()
        schC_CDATE1.Text = ""
        schC_CDATE2.Text = ""
        schC_SDATE1.Text = ""
        schC_SDATE2.Text = ""
        schCONTENT1.Text = ""
        schC_EDATE1.Text = ""
        schC_EDATE2.Text = ""
    End Sub

    '資料查詢
    Sub sSearch1()
        schC_CDATE1.Text = TIMS.ClearSQM(schC_CDATE1.Text).Trim
        schC_CDATE2.Text = TIMS.ClearSQM(schC_CDATE2.Text).Trim
        schC_SDATE1.Text = TIMS.ClearSQM(schC_SDATE1.Text).Trim
        schC_SDATE2.Text = TIMS.ClearSQM(schC_SDATE2.Text).Trim
        schC_EDATE1.Text = TIMS.ClearSQM(schC_EDATE1.Text).Trim
        schC_EDATE2.Text = TIMS.ClearSQM(schC_EDATE2.Text).Trim
        schCONTENT1.Text = TIMS.ClearSQM(schCONTENT1.Text).Trim

        Dim schCCDATE1 As String = schC_CDATE1.Text.Trim
        Dim schCCDATE2 As String = schC_CDATE2.Text.Trim
        Dim schCSDATE1 As String = schC_SDATE1.Text.Trim
        Dim schCSDATE2 As String = schC_SDATE2.Text.Trim
        Dim schCEDATE1 As String = schC_EDATE1.Text.Trim
        Dim schCEDATE2 As String = schC_EDATE2.Text.Trim

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.SEQNO DESC) AS ROWNUM " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.C_SDATE, 111) CSDATE " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.C_EDATE, 111) CEDATE " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.C_CDATE, 111) CCDATE " & vbCrLf
        sql &= "        ,a.SEQNO " & vbCrLf
        sql &= "        ,a.FUNID " & vbCrLf
        sql &= "        ,a.C_SDATE " & vbCrLf
        sql &= "        ,a.C_EDATE " & vbCrLf
        sql &= "        ,a.C_TITLE " & vbCrLf
        sql &= "        ,CASE WHEN LEN(a.C_CONTENT1) > 20 THEN SUBSTRING(a.C_CONTENT1, 1, 20) + '...' ELSE a.C_CONTENT1 END CCONTENT1 " & vbCrLf
        sql &= "        ,a.C_CONTENT1 " & vbCrLf
        sql &= "        ,a.C_CONTENT2 " & vbCrLf
        sql &= "        ,a.C_CONTENT3 " & vbCrLf
        sql &= "        ,a.C_CDATE " & vbCrLf
        sql &= "        ,a.C_CACCT " & vbCrLf
        sql &= "        ,a.C_UDATE " & vbCrLf
        sql &= "        ,a.C_UACCT " & vbCrLf
        sql &= "        ,a.C_STATUS " & vbCrLf
        sql &= " FROM TB_CONTENT a " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "       AND a.C_STATUS <> 'D' " & vbCrLf
        sql &= "       AND a.FUNID = '004' " & vbCrLf

        If schCCDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.C_CDATE, 111) >= @schCCDATE1 " & vbCrLf
        If schCCDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.C_CDATE, 111) <= @schCCDATE2 " & vbCrLf
        If schCSDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.C_SDATE, 111) >= @schCSDATE1 " & vbCrLf
        If schCSDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.C_SDATE, 111) <= @schCSDATE2 " & vbCrLf
        If schCEDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.C_EDATE, 111) >= @schCEDATE1 " & vbCrLf
        If schCEDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.C_EDATE, 111) <= @schCEDATE2 " & vbCrLf
        If schCONTENT1.Text <> "" Then sql &= (" AND a.C_CONTENT1 LIKE @schCONTENT1 ") & vbCrLf

        sql &= " ORDER BY a.C_UDATE DESC " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If schCCDATE1 <> "" Then parms.Add("schCCDATE1", IIf(flag_ROC, TIMS.Cdate18(schCCDATE1), schCCDATE1))  'edit，by:20181019
        If schCCDATE2 <> "" Then parms.Add("schCCDATE2", IIf(flag_ROC, TIMS.Cdate18(schCCDATE2), schCCDATE2))  'edit，by:20181019
        If schCSDATE1 <> "" Then parms.Add("schCSDATE1", IIf(flag_ROC, TIMS.Cdate18(schCSDATE1), schCSDATE1))  'edit，by:20181019
        If schCSDATE2 <> "" Then parms.Add("schCSDATE2", IIf(flag_ROC, TIMS.Cdate18(schCSDATE2), schCSDATE2))  'edit，by:20181019
        If schCEDATE1 <> "" Then parms.Add("schCEDATE1", IIf(flag_ROC, TIMS.Cdate18(schCEDATE1), schCEDATE1))  'edit，by:20181019
        If schCEDATE2 <> "" Then parms.Add("schCEDATE2", IIf(flag_ROC, TIMS.Cdate18(schCEDATE2), schCEDATE2))  'edit，by:20181019
        If schCONTENT1.Text <> "" Then parms.Add("schCONTENT1", ("%" + schCONTENT1.Text + "%"))

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
        Dim url1 As String = "RWB_01_005_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(1).Text = IIf(flag_ROC, TIMS.Cdate17(drv("CCDATE")), drv("CCDATE"))  'edit，by:20181019
                e.Item.Cells(2).Text = IIf(flag_ROC, TIMS.Cdate17(drv("CSDATE")), drv("CSDATE"))  'edit，by:20181019
                e.Item.Cells(3).Text = IIf(flag_ROC, TIMS.Cdate17(drv("CEDATE")), drv("CEDATE"))  'edit，by:20181019
                Dim btnEDIT1 As Button = e.Item.FindControl("btnEDIT1")
                Dim btnDEL1 As Button = e.Item.FindControl("btnDEL1")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SEQNO", TIMS.CStr1(drv("SEQNO")))
                btnEDIT1.CommandArgument = sCmdArg
                btnDEL1.CommandArgument = sCmdArg
                btnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg2
        End Select
    End Sub

    'DataGrid1功能事件
    Protected Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        hid_V.Value = TIMS.GetMyValue(sCmdArg, "SEQNO")
        If hid_V.Value = "" Then Exit Sub
        hid_V.Value = TIMS.ClearSQM(hid_V.Value)

        Select Case e.CommandName
            Case "edit"
                GetSearchStr()  'edit，by:20190103
                Dim url1 As String = "RWB_01_005_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
                url1 &= "&A=E&SEQNO=" & hid_V.Value
                url1 &= "&SEQNO_E=" & TIMS.EncryptAes(hid_V.Value)
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " UPDATE TB_CONTENT " & vbCrLf
                sql &= " SET C_STATUS = 'D' " & vbCrLf
                sql &= " ,C_UDATE = GETDATE() " & vbCrLf
                sql &= " ,C_UACCT = @C_UACCT " & vbCrLf
                sql &= " WHERE SEQNO = @SEQNO " & vbCrLf
                Dim parms As Hashtable = New Hashtable()
                parms.Add("C_UACCT", sm.UserInfo.UserID)
                parms.Add("SEQNO", hid_V.Value)
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
        Session("_SearchStr") = "prg=RWB_01_005"
        Session("_SearchStr") &= "&schC_CDATE1=" & schC_CDATE1.Text
        Session("_SearchStr") += "&schC_CDATE2=" & schC_CDATE2.Text
        Session("_SearchStr") += "&schC_SDATE1=" & schC_SDATE1.Text
        Session("_SearchStr") += "&schC_SDATE2=" & schC_SDATE2.Text
        Session("_SearchStr") += "&schCONTENT1=" & schCONTENT1.Text
        Session("_SearchStr") += "&schC_EDATE1=" & schC_EDATE1.Text
        Session("_SearchStr") += "&schC_EDATE2=" & schC_EDATE2.Text
    End Sub
End Class