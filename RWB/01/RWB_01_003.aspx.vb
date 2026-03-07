Partial Class RWB_01_003
    Inherits AuthBasePage

    ''TB_DLFILE
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
                ddlType.SelectedIndex = ddlType.Items.IndexOf(ddlType.Items.FindByValue(TIMS.GetMyValue(Session("_SearchStr"), "ddlType")))
                change_ddlType()
                ddlPlan.SelectedIndex = ddlPlan.Items.IndexOf(ddlPlan.Items.FindByValue(TIMS.GetMyValue(Session("_SearchStr"), "ddlPlan")))
                schC_CDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_CDATE1")
                schC_CDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_CDATE2")
                schC_SDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE1")
                schC_SDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE2")
                schKeyword1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schKeyword1")
                schC_EDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE1")
                schC_EDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE2")
                sSearch1()
                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    '頁面初始化
    Sub sCreate1()
        ddlType.SelectedValue = "1"

        ddlPlan.Items.Clear()
        ddlPlan.Items.Add(New ListItem("產業人才投資方案", "1"))
        ddlPlan.Items.Add(New ListItem("自辦在職訓練", "2"))
        ddlPlan.Items.Add(New ListItem("企業委託訓練", "3"))
        ddlPlan.Items.Add(New ListItem("充電起飛", "4"))
        ddlPlan.SelectedIndex = 0
        ddlPlan.Enabled = True

        schC_CDATE1.Text = ""
        schC_CDATE2.Text = ""
        schC_SDATE1.Text = ""
        schC_SDATE2.Text = ""
        schKeyword1.Text = ""
        schC_EDATE1.Text = ""
        schC_EDATE2.Text = ""
    End Sub

    '[(資料下載)類別]異動事件
    Protected Sub ddlType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlType.SelectedIndexChanged
        change_ddlType()
    End Sub

    '[(資料下載)類別]項目進行異動
    Sub change_ddlType()
        If ddlType.SelectedValue = "1" Then
            ddlPlan.Items.Clear()
            ddlPlan.Items.Add(New ListItem("產業人才投資方案", "1"))
            ddlPlan.Items.Add(New ListItem("自辦在職訓練", "2"))
            ddlPlan.Items.Add(New ListItem("企業委託訓練", "3"))
            ddlPlan.Items.Add(New ListItem("充電起飛", "4"))
            ddlPlan.SelectedIndex = 0
            ddlPlan.Enabled = True
        Else
            ddlPlan.Items.Clear()
            ddlPlan.Items.Add(New ListItem("[無計畫內容]", ""))
            ddlPlan.SelectedIndex = 0
            ddlPlan.Enabled = False
        End If
    End Sub

    '資料查詢
    Sub sSearch1()
        schC_CDATE1.Text = TIMS.ClearSQM(schC_CDATE1.Text).Trim
        schC_CDATE2.Text = TIMS.ClearSQM(schC_CDATE2.Text).Trim
        schC_SDATE1.Text = TIMS.ClearSQM(schC_SDATE1.Text).Trim
        schC_SDATE2.Text = TIMS.ClearSQM(schC_SDATE2.Text).Trim
        schC_EDATE1.Text = TIMS.ClearSQM(schC_EDATE1.Text).Trim
        schC_EDATE2.Text = TIMS.ClearSQM(schC_EDATE2.Text).Trim
        schKeyword1.Text = TIMS.ClearSQM(schKeyword1.Text).Trim

        Dim myCCDATE1 As String = schC_CDATE1.Text.Trim
        Dim myCCDATE2 As String = schC_CDATE2.Text.Trim
        Dim myCSDATE1 As String = schC_SDATE1.Text.Trim
        Dim myCSDATE2 As String = schC_SDATE2.Text.Trim
        Dim myCEDATE1 As String = schC_EDATE1.Text.Trim
        Dim myCEDATE2 As String = schC_EDATE2.Text.Trim
        Dim myTYPE As String = ddlType.SelectedValue.Trim
        Dim myPLAN As String = ddlPlan.SelectedValue.Trim
        Dim myTITLE As String = schKeyword1.Text.Trim

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.DLID DESC) AS ROWNUM " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.START_DATE, 111) CSDATE " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.END_DATE, 111) CEDATE " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.UPLOADDATE, 111) CCDATE " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.MODIFYDATE, 111) CMDATE " & vbCrLf
        sql &= "        ,a.DLID " & vbCrLf
        sql &= "        ,a.KINDID " & vbCrLf
        sql &= "        ,a.PLANID " & vbCrLf
        sql &= "        ,a.START_DATE " & vbCrLf
        sql &= "        ,a.END_DATE " & vbCrLf
        sql &= "        ,a.DLTITLE " & vbCrLf
        sql &= "        ,CASE WHEN LEN(a.DLTITLE) > 20 THEN SUBSTRING(a.DLTITLE, 1, 20) + '...' ELSE a.DLTITLE END TITLE1 " & vbCrLf
        sql &= "        ,a.UPLOADDATE " & vbCrLf
        sql &= "        ,a.ISUSED " & vbCrLf
        sql &= "        ,a.MEMO " & vbCrLf
        sql &= "        ,a.MODIFYACCT " & vbCrLf
        sql &= "        ,a.MODIFYDATE " & vbCrLf
        sql &= " FROM TB_DLFILE a " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.ISUSED = 'Y' " & vbCrLf

        If myCCDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.UPLOADDATE, 111) >= @CCDATE1 " & vbCrLf
        If myCCDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.UPLOADDATE, 111) <= @CCDATE2 " & vbCrLf
        If myCSDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.START_DATE, 111) >= @CSDATE1 " & vbCrLf
        If myCSDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.START_DATE, 111) <= @CSDATE2 " & vbCrLf
        If myCEDATE1 <> "" Then sql &= " AND CONVERT(VARCHAR, a.END_DATE, 111) >= @CEDATE1 " & vbCrLf
        If myCEDATE2 <> "" Then sql &= " AND CONVERT(VARCHAR, a.END_DATE, 111) <= @CEDATE2 " & vbCrLf
        If myTYPE <> "" Then sql &= " AND a.KINDID = @TYPE1 " & vbCrLf
        If myPLAN <> "" Then sql &= " AND a.PLANID = @PLAN1 " & vbCrLf
        If myTITLE <> "" Then sql &= (" AND a.DLTITLE LIKE @TITLE1 ") & vbCrLf

        sql &= " ORDER BY a.MODIFYDATE DESC " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If myCCDATE1 <> "" Then parms.Add("CCDATE1", IIf(flag_ROC, TIMS.Cdate18(myCCDATE1), myCCDATE1))  'edit，by:20181019
        If myCCDATE2 <> "" Then parms.Add("CCDATE2", IIf(flag_ROC, TIMS.Cdate18(myCCDATE2), myCCDATE2))  'edit，by:20181019
        If myCSDATE1 <> "" Then parms.Add("CSDATE1", IIf(flag_ROC, TIMS.Cdate18(myCSDATE1), myCSDATE1))  'edit，by:20181019
        If myCSDATE2 <> "" Then parms.Add("CSDATE2", IIf(flag_ROC, TIMS.Cdate18(myCSDATE2), myCSDATE2))  'edit，by:20181019
        If myCEDATE1 <> "" Then parms.Add("CEDATE1", IIf(flag_ROC, TIMS.Cdate18(myCEDATE1), myCEDATE1))  'edit，by:20181019
        If myCEDATE2 <> "" Then parms.Add("CEDATE2", IIf(flag_ROC, TIMS.Cdate18(myCEDATE2), myCEDATE2))  'edit，by:20181019
        If myTYPE <> "" Then parms.Add("TYPE1", myTYPE)
        If myPLAN <> "" Then parms.Add("PLAN1", myPLAN)
        If myTITLE <> "" Then parms.Add("TITLE1", ("%" + myTITLE + "%"))

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
        Dim url1 As String = "RWB_01_003_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
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
                TIMS.SetMyValue(sCmdArg, "DLID", TIMS.CStr1(drv("DLID")))
                btnEDIT1.CommandArgument = sCmdArg
                btnDEL1.CommandArgument = sCmdArg
                btnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg2
        End Select
    End Sub

    'DataGrid1功能事件
    Protected Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        hid_V.Value = TIMS.GetMyValue(sCmdArg, "DLID")
        If hid_V.Value = "" Then Exit Sub
        hid_V.Value = TIMS.ClearSQM(hid_V.Value)

        Select Case e.CommandName
            Case "edit"
                GetSearchStr()  'edit，by:20190103
                Dim url1 As String = "RWB_01_003_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
                url1 &= "&A=E&DLID=" & hid_V.Value
                url1 &= "&SEQNO_E=" & TIMS.EncryptAes(hid_V.Value)
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " UPDATE TB_DLFILE " & vbCrLf
                sql &= " SET ISUSED = 'N' " & vbCrLf
                sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
                sql &= " ,MODIFYACCT = @UACCT " & vbCrLf
                sql &= " WHERE DLID = @DLID " & vbCrLf
                Dim parms As Hashtable = New Hashtable()
                parms.Add("UACCT", sm.UserInfo.UserID)
                parms.Add("DLID", hid_V.Value)
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
        Session("_SearchStr") = "prg=RWB_01_003"
        Session("_SearchStr") &= "&ddlType=" & ddlType.SelectedValue
        Session("_SearchStr") += "&ddlPlan=" & ddlPlan.SelectedValue
        Session("_SearchStr") += "&schC_CDATE1=" & schC_CDATE1.Text
        Session("_SearchStr") += "&schC_CDATE2=" & schC_CDATE2.Text
        Session("_SearchStr") += "&schC_SDATE1=" & schC_SDATE1.Text
        Session("_SearchStr") += "&schC_SDATE2=" & schC_SDATE2.Text
        Session("_SearchStr") += "&schKeyword1=" & schKeyword1.Text
        Session("_SearchStr") += "&schC_EDATE1=" & schC_EDATE1.Text
        Session("_SearchStr") += "&schC_EDATE2=" & schC_EDATE2.Text
    End Sub
End Class