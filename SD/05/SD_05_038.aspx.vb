Partial Class SD_05_038
    Inherits AuthBasePage

    Dim dtDIST As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '在這裡放置使用者程式碼以初始化網頁
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
#Region "(放置使用者程式碼以初始化網頁)"

        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        PageControler1.PageDataGrid = DataGrid1
        If DataGrid1.Items.Count > 0 Then PageControler1.Visible = True Else PageControler1.Visible = False

        Dim sql As String
        sql = "SELECT DISTID,NAME FROM ID_DISTRICT WHERE DISTID!='000' ORDER BY DISTID"
        dtDIST = DbAccess.GetDataTable(sql, objconn)

        If Not Page.IsPostBack Then
            labmsg.Text = ""
            tbSch.Visible = True
            Call initObj()
        End If

#End Region
    End Sub

    '功能第一次載入初始化
    Sub initObj()
#Region "功能第一次載入初始化"

        sRECOMMDISTID = TIMS.Get_DistID(sRECOMMDISTID, dtDIST)
        sRECOMMDISTID.Enabled = False
        Common.SetListItem(sRECOMMDISTID, sm.UserInfo.DistID)
        If sm.UserInfo.DistID = "000" Then sRECOMMDISTID.Enabled = True

#End Region
    End Sub

    '記錄查詢條件 
    Sub Search1Value()
#Region "記錄查詢條件"

        sIDNO.Text = TIMS.ChangeIDNO(sIDNO.Text)
        sCNAME.Text = TIMS.ClearSQM(sCNAME.Text)
        sIDNO.Text = TIMS.ClearSQM(sIDNO.Text)
        sBIRTHDAY1.Text = TIMS.ClearSQM(sBIRTHDAY1.Text)
        sBIRTHDAY2.Text = TIMS.ClearSQM(sBIRTHDAY2.Text)
        sPREEXDATE1.Text = TIMS.ClearSQM(sPREEXDATE1.Text)
        sPREEXDATE2.Text = TIMS.ClearSQM(sPREEXDATE2.Text)
        sBIRTHDAY1.Text = TIMS.cdate3(sBIRTHDAY1.Text)
        sBIRTHDAY2.Text = TIMS.cdate3(sBIRTHDAY2.Text)
        sPREEXDATE1.Text = TIMS.cdate3(sPREEXDATE1.Text)
        sPREEXDATE2.Text = TIMS.cdate3(sPREEXDATE2.Text)

        ViewState("sCNAME") = sCNAME.Text
        ViewState("sIDNO") = sIDNO.Text
        ViewState("sBIRTHDAY1") = sBIRTHDAY1.Text
        ViewState("sBIRTHDAY2") = sBIRTHDAY2.Text
        ViewState("sPREEXDATE1") = sPREEXDATE1.Text
        ViewState("sPREEXDATE2") = sPREEXDATE2.Text
        If sm.UserInfo.DistID = "000" Then ViewState("sRECOMMDISTID") = sRECOMMDISTID.SelectedValue
        If sm.UserInfo.DistID <> "000" Then ViewState("sRECOMMDISTID") = sm.UserInfo.DistID

#End Region
    End Sub

    '查詢
    Sub Search1()
#Region "查詢"

        'ADP_RESOLDER
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
            Me.TxtPageSize.Text = 10
        End If
        If Me.TxtPageSize.Text <> Me.DataGrid1.PageSize Then Me.DataGrid1.PageSize = Me.TxtPageSize.Text

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.IDNO, a.NAME, a.DISTNAME, " & vbCrLf
        sql += "        ISNULL(CONVERT(varchar, a.BIRTHDAY, 111), '') BIRTHDAY, " & vbCrLf
        sql += "        ISNULL(CONVERT(varchar, b.FServiceDate, 111), '') PREEXDATE " & vbCrLf
        sql += " FROM V_STUDENTINFO a " & vbCrLf
        sql += " JOIN STUD_SUBDATA b ON b.SID = a.SID " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        sql += "       AND a.IDENTITYID = '43' " & vbCrLf

        If Convert.ToString(ViewState("sCNAME")) <> "" Then sql += " AND a.NAME like '%' + @sCNAME + '%' " & vbCrLf
        If Convert.ToString(ViewState("sIDNO")) <> "" Then sql += " AND a.IDNO = @sIDNO " & vbCrLf
        If Convert.ToString(ViewState("sBIRTHDAY1")) <> "" Then sql += " AND a.BIRTHDAY >= @sBIRTHDAY1 " & vbCrLf
        If Convert.ToString(ViewState("sBIRTHDAY2")) <> "" Then sql += " AND a.BIRTHDAY <= @sBIRTHDAY2 " & vbCrLf
        If Convert.ToString(ViewState("sPREEXDATE1")) <> "" Then sql += " AND b.FServiceDate >= @sPREEXDATE1 " & vbCrLf
        If Convert.ToString(ViewState("sPREEXDATE2")) <> "" Then sql += " AND b.FServiceDate <= @sPREEXDATE2 " & vbCrLf
        If Convert.ToString(ViewState("sRECOMMDISTID")) <> "" Then sql += " AND a.DISTID = @sRECOMMDISTID " & vbCrLf
        sql += " ORDER BY a.DISTID, a.NAME, a.BIRTHDAY " & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim parms As Hashtable = New Hashtable()
        If Convert.ToString(ViewState("sCNAME")) <> "" Then parms.Add("sCNAME", ViewState("sCNAME"))
        If Convert.ToString(ViewState("sIDNO")) <> "" Then parms.Add("sIDNO", ViewState("sIDNO"))
        If Convert.ToString(ViewState("sBIRTHDAY1")) <> "" Then parms.Add("sBIRTHDAY1", CDate(ViewState("sBIRTHDAY1")))
        If Convert.ToString(ViewState("sBIRTHDAY2")) <> "" Then parms.Add("sBIRTHDAY2", CDate(ViewState("sBIRTHDAY2")))
        If Convert.ToString(ViewState("sPREEXDATE1")) <> "" Then parms.Add("sPREEXDATE1", CDate(ViewState("sPREEXDATE1")))
        If Convert.ToString(ViewState("sPREEXDATE2")) <> "" Then parms.Add("sPREEXDATE2", CDate(ViewState("sPREEXDATE2")))
        If Convert.ToString(ViewState("sRECOMMDISTID")) <> "" Then parms.Add("sRECOMMDISTID", ViewState("sRECOMMDISTID"))
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        labmsg.Text = "查無資料"
        tbList.Visible = False
        If dt.Rows.Count > 0 Then
            labmsg.Text = ""
            tbList.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

#End Region
    End Sub

    '查詢鈕
    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call Search1Value()  '記錄查詢條件
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
#Region "(表格上的元件配置)"

        Dim objDG1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                e.Item.Cells(0).Text = (objDG1.PageSize * objDG1.CurrentPageIndex) + e.Item.ItemIndex + 1  '序號
        End Select

#End Region
    End Sub
End Class