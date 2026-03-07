Partial Class RWB_01_001
    Inherits AuthBasePage

    ''TB_CONTENT TB_CONTENT_SECTION
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
        'If PageControler1.PageDataGrid.Items.Count > 0 Then
        '    PageControler1.Visible = True
        'Else
        '    PageControler1.Visible = False
        'End If

        If Not IsPostBack Then
            Call sCreate1() '頁面初始化


        End If

    End Sub

    '頁面初始化
    Sub sCreate1()
        ddlType = TIMS.Get_RWBFUNTYPE(ddlType, 4)
        'ddlType.SelectedValue = "1"
        Common.SetListItem(ddlType, "1")

        schC_CDATE1.Text = ""
        schC_CDATE2.Text = ""
        schC_SDATE1.Text = ""
        schC_SDATE2.Text = ""
        schKeyword1.Text = ""
        schC_EDATE1.Text = ""
        schC_EDATE2.Text = ""

        '20190103 若有先前查詢條件記錄，則將資料重新讀取到頁面中
        If Session("_SearchStr") IsNot Nothing Then
            Dim myValue As String = ""
            myValue = TIMS.GetMyValue(Session("_SearchStr"), "prg")
            If myValue = "RWB_01_001" Then
                'ddlType.SelectedIndex = ddlType.Items.IndexOf(ddlType.Items.FindByValue(TIMS.GetMyValue(Session("_SearchStr"), "ddlType")))
                myValue = TIMS.GetMyValue(Session("_SearchStr"), "ddlType")
                Common.SetListItem(ddlType, myValue)
                schC_CDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_CDATE1")
                schC_CDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_CDATE2")
                schC_SDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE1")
                schC_SDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_SDATE2")
                schKeyword1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schKeyword1")
                schC_EDATE1.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE1")
                schC_EDATE2.Text = TIMS.GetMyValue(Session("_SearchStr"), "schC_EDATE2")
                sSearch1()
            End If
            Session("_SearchStr") = Nothing
        End If

    End Sub

    '資料查詢
    Sub sSearch1()
        schC_CDATE1.Text = TIMS.ClearSQM(schC_CDATE1.Text)
        schC_CDATE2.Text = TIMS.ClearSQM(schC_CDATE2.Text)
        schC_SDATE1.Text = TIMS.ClearSQM(schC_SDATE1.Text)
        schC_SDATE2.Text = TIMS.ClearSQM(schC_SDATE2.Text)
        schC_EDATE1.Text = TIMS.ClearSQM(schC_EDATE1.Text)
        schC_EDATE2.Text = TIMS.ClearSQM(schC_EDATE2.Text)
        schKeyword1.Text = TIMS.ClearSQM(schKeyword1.Text)

        Dim myCCDATE1 As String = schC_CDATE1.Text
        Dim myCCDATE2 As String = schC_CDATE2.Text
        Dim myCSDATE1 As String = schC_SDATE1.Text
        Dim myCSDATE2 As String = schC_SDATE2.Text
        Dim myCEDATE1 As String = schC_EDATE1.Text
        Dim myCEDATE2 As String = schC_EDATE2.Text

        'Dim myTYPE As String = TIMS.ClearSQM(ddlType.SelectedValue)
        Dim v_ddlType As String = TIMS.GetListValue(ddlType)
        Dim myKEYWORD1 As String = schKeyword1.Text

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY source.C_UDATE DESC) AS ROWNUM, * " & vbCrLf
        sql &= " FROM ( " & vbCrLf
        'sql &= "  SELECT DISTINCT CONVERT(VARCHAR, a.C_SDATE, 111) CSDATE " & vbCrLf
        sql &= "  SELECT CONVERT(VARCHAR, a.C_SDATE, 111) CSDATE " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, a.C_EDATE, 111) CEDATE " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, a.C_CDATE, 111) CCDATE " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, a.C_UDATE, 111) CUDATE " & vbCrLf
        sql &= "  ,a.SEQNO " & vbCrLf
        sql &= "  ,a.FUNID " & vbCrLf
        sql &= "  ,a.SUB_FUNID " & vbCrLf
        sql &= "  ,a.C_SDATE " & vbCrLf
        sql &= "  ,a.C_EDATE " & vbCrLf
        sql &= "  ,ISNULL(a.C_TITLE, '') C_TITLE " & vbCrLf
        sql &= "  ,CASE WHEN LEN(a.C_TITLE) > 20 THEN SUBSTRING(a.C_TITLE, 1, 20) + '...' ELSE ISNULL(a.C_TITLE, '') END TITLE1 " & vbCrLf
        sql &= "  ,a.C_CDATE " & vbCrLf
        sql &= "  ,a.C_STATUS " & vbCrLf
        sql &= "  ,a.C_UDATE " & vbCrLf
        sql &= "  ,a.C_UACCT " & vbCrLf
        sql &= "  FROM TB_CONTENT a " & vbCrLf
        sql &= "  LEFT JOIN TB_CONTENT_SECTION b ON a.SEQNO = b.CONTENTID AND b.SEC_NO='1'" & vbCrLf
        sql &= "  WHERE 1=1 " & vbCrLf
        sql &= "  AND a.C_STATUS <> 'D' " & vbCrLf

        If myCCDATE1 <> "" Then sql &= " AND a.C_CDATE >= CONVERT(date, @CCDATE1)" & vbCrLf
        If myCCDATE2 <> "" Then sql &= " AND a.C_CDATE <= CONVERT(date, @CCDATE2)" & vbCrLf
        If myCSDATE1 <> "" Then sql &= " AND a.C_SDATE >= CONVERT(date, @CSDATE1)" & vbCrLf
        If myCSDATE2 <> "" Then sql &= " AND a.C_SDATE <= CONVERT(date, @CSDATE2)" & vbCrLf
        If myCEDATE1 <> "" Then sql &= " AND ISNULL(a.C_EDATE, GETDATE()+(365*3)) >= CONVERT(date, @CEDATE1) " & vbCrLf
        If myCEDATE2 <> "" Then sql &= " AND ISNULL(a.C_EDATE, GETDATE()+(365*3)) <= CONVERT(date, @CEDATE2) " & vbCrLf

        Select Case v_ddlType'1:焦點消息:/2:計畫公告/3:成果集錦/011:宣導影片
            Case "011"
                sql &= " AND a.FUNID = '011' " & vbCrLf
            Case Else
                sql &= " AND a.FUNID = '001' " & vbCrLf
                If v_ddlType <> "" Then sql &= " AND a.SUB_FUNID = @TYPE1 " & vbCrLf
        End Select

        If myKEYWORD1 <> "" Then sql &= (" AND (ISNULL(a.C_TITLE, '') LIKE @TITLE1 OR ISNULL(b.SEC_CONTENT, '') LIKE @TITLE1) ") & vbCrLf

        sql &= "      ) source " & vbCrLf
        sql &= " ORDER BY source.C_UDATE DESC " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If flag_ROC Then
            'edit，by:20181018-由民國日期 改為 西元日期
            myCCDATE1 = If(myCCDATE1 <> "", TIMS.Cdate18(myCCDATE1), "")
            myCCDATE2 = If(myCCDATE2 <> "", TIMS.Cdate18(myCCDATE2), "")
            myCSDATE1 = If(myCSDATE1 <> "", TIMS.Cdate18(myCSDATE1), "")
            myCSDATE2 = If(myCSDATE2 <> "", TIMS.Cdate18(myCSDATE2), "")
            myCEDATE1 = If(myCEDATE1 <> "", TIMS.Cdate18(myCEDATE1), "")
            myCEDATE2 = If(myCEDATE2 <> "", TIMS.Cdate18(myCEDATE2), "")
        End If
        'edit，by:20181018
        parms.Add("CCDATE1", If(myCCDATE1 <> "", myCCDATE1, ""))
        parms.Add("CCDATE2", If(myCCDATE2 <> "", myCCDATE2, ""))
        parms.Add("CSDATE1", If(myCSDATE1 <> "", myCSDATE1, ""))
        parms.Add("CSDATE2", If(myCSDATE2 <> "", myCSDATE2, ""))
        parms.Add("CEDATE1", If(myCEDATE1 <> "", myCEDATE1, ""))
        parms.Add("CEDATE2", If(myCEDATE2 <> "", myCEDATE2, ""))

        Select Case v_ddlType'1:焦點消息:/2:計畫公告/3:成果集錦/011:宣導影片
            Case "011"
                'sql &= " AND a.FUNID = '011' " & vbCrLf
            Case Else
                'sql &= " AND a.FUNID = '001' " & vbCrLf
                If v_ddlType <> "" Then parms.Add("TYPE1", v_ddlType)
        End Select

        If myKEYWORD1 <> "" Then parms.Add("TITLE1", ("%" + myKEYWORD1 + "%"))

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料"
        tb_Sch.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            tb_Sch.Visible = True
            PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '進行資料查詢作業
    Protected Sub bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Call sSearch1() '查詢
    End Sub

    ''' <summary>
    ''' 資料新增
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_add_Click(sender As Object, e As EventArgs) Handles bt_add.Click
        Dim strSession As String = ""
        Dim v_ddlType As String = TIMS.GetListValue(ddlType)
        TIMS.SetMyValue(strSession, "ddlType", v_ddlType)

        Const cst_rwb01001_add As String = "rwb01001_add"
        Session(cst_rwb01001_add) = strSession

        Dim url1 As String = "RWB_01_001_edit.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    'DataGrid1功能事件
    Protected Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        hid_V_SEQNO.Value = TIMS.GetMyValue(sCmdArg, "SEQNO")
        hid_V_FUNID.Value = TIMS.GetMyValue(sCmdArg, "FUNID")

        If hid_V_SEQNO.Value = "" Then Exit Sub
        hid_V_SEQNO.Value = TIMS.ClearSQM(hid_V_SEQNO.Value)

        Select Case e.CommandName
            Case "edit"
                GetSearchStr()  'edit，by:20190103
                Dim myEditPage As String = "RWB_01_001_edit.aspx"
                'If Not ddlType.SelectedValue.Equals("3") Then myEditPage = "RWB_01_001_edit.aspx" Else myEditPage = "RWB_01_001_edit2.aspx"
                Dim url1 As String = myEditPage + "?id1=" & TIMS.Get_MRqID(Me)
                url1 &= "&A=E&SEQNO=" & hid_V_SEQNO.Value 'Request("A")
                url1 &= "&SEQNO_E=" & TIMS.EncryptAes(hid_V_SEQNO.Value)
                url1 &= "&FUNID_E=" & TIMS.EncryptAes(hid_V_FUNID.Value)

                TIMS.Utl_Redirect(Me, objconn, url1)

            Case "del"
                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " UPDATE TB_CONTENT " & vbCrLf
                sql &= " SET C_STATUS = 'D' " & vbCrLf
                sql &= " ,C_UDATE = GETDATE() " & vbCrLf
                sql &= " ,C_UACCT = @UACCT " & vbCrLf
                sql &= " WHERE SEQNO = @SEQNO " & vbCrLf
                Dim parms As Hashtable = New Hashtable()
                parms.Add("UACCT", sm.UserInfo.UserID)
                parms.Add("SEQNO", hid_V_SEQNO.Value)
                DbAccess.ExecuteNonQuery(sql, objconn, parms)

                Common.MessageBox(Me, "刪除成功")
                Call sSearch1()
            Case Else
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(1).Text = If(flag_ROC, TIMS.Cdate17(drv("CCDATE")), drv("CCDATE"))  'edit，by:20181019
                e.Item.Cells(2).Text = If(flag_ROC, TIMS.Cdate17(drv("CSDATE")), drv("CSDATE"))  'edit，by:20181019
                e.Item.Cells(3).Text = If(flag_ROC, TIMS.Cdate17(drv("CEDATE")), drv("CEDATE"))  'edit，by:20181019

                Dim btnEDIT1 As Button = e.Item.FindControl("btnEDIT1")
                Dim btnDEL1 As Button = e.Item.FindControl("btnDEL1")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SEQNO", TIMS.CStr1(drv("SEQNO")))
                TIMS.SetMyValue(sCmdArg, "FUNID", TIMS.CStr1(drv("FUNID")))
                btnEDIT1.CommandArgument = sCmdArg
                btnDEL1.CommandArgument = sCmdArg
                btnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg2
        End Select
    End Sub


    '下拉式選單異動事件
    'Protected Sub ddlType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlType.SelectedIndexChanged
    '    If ddlType.SelectedValue = "3" Then bt_add.Visible = False Else bt_add.Visible = True
    'End Sub

    '20190103 將目前的查詢條件儲存起來
    Sub GetSearchStr()
        Dim v_ddlType As String = TIMS.GetListValue(ddlType)
        Session("_SearchStr") = "prg=RWB_01_001"
        Session("_SearchStr") &= "&ddlType=" & v_ddlType 'ddlType.SelectedValue
        Session("_SearchStr") += "&schC_CDATE1=" & schC_CDATE1.Text
        Session("_SearchStr") += "&schC_CDATE2=" & schC_CDATE2.Text
        Session("_SearchStr") += "&schC_SDATE1=" & schC_SDATE1.Text
        Session("_SearchStr") += "&schC_SDATE2=" & schC_SDATE2.Text
        Session("_SearchStr") += "&schKeyword1=" & schKeyword1.Text
        Session("_SearchStr") += "&schC_EDATE1=" & schC_EDATE1.Text
        Session("_SearchStr") += "&schC_EDATE2=" & schC_EDATE2.Text
    End Sub
End Class