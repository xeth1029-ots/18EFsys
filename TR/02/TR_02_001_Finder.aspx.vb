Partial Class TR_02_001_Finder
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9) '☆
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        msg.Text = ""

        If Not IsPostBack Then
            DataGridTable.Visible = False
        End If
        PageControler1.PageDataGrid = DataGrid1
        Button3.Attributes("onclick") = "getZip('../../js/Openwin/zipcode_search.aspx', 'TBCity', 'zip_code','city_code')"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Me.ViewState("search1") = ""
        Me.ViewState("search1") += Convert.ToString(city_code.Value).Trim
        Me.ViewState("search1") += Convert.ToString(zip_code.Value).Trim
        Me.ViewState("search1") += Convert.ToString(Uname.Text).Trim
        Me.ViewState("search1") += Convert.ToString(Ubno.Text).Trim
        Me.ViewState("search1") += Convert.ToString(Intaxno.Text).Trim
        If Me.ViewState("search1") = "" Then
            Common.MessageBox(Me, "避免搜尋結果資料過多，浪費系統資源，請輸入一查詢條件")
            Exit Sub
        End If

        'Dim sql As String = ""
        'Dim dt As DataTable = Nothing
        'Dim NewSql As String = ""
        Dim ZipStr As String = ""
        If city_code.Value <> "" Then
            ZipStr += " and Zip IN (SELECT ZipCode FROM ID_ZIP WHERE CTID IN (SELECT CTID FROM ID_City WHERE CTID='" & city_code.Value & "'))"
        End If
        If zip_code.Value <> "" Then
            ZipStr += " and Zip IN (SELECT ZipCode FROM ID_ZIP WHERE ZipCode='" & zip_code.Value & "')"
        End If

        Dim sql As String = ""
        'sql = "SELECT * FROM Bus_BasicData WHERE Uname like '%" & Uname.Text & "%' and Ubno like '%" & Ubno.Text & "%' and Intaxno like '%" & Intaxno.Text & "%' and BDID Not IN (SELECT BDID FROM Bus_VisitInfo WHERE DistID='" & sm.UserInfo.DistID & "')" & ZipStr
        sql = "SELECT * FROM Bus_BasicData WHERE 1=1 and BDID Not IN (SELECT BDID FROM Bus_VisitInfo WHERE DistID='" & sm.UserInfo.DistID & "')" & ZipStr
        If Convert.ToString(Uname.Text).Trim <> "" Then
            sql += " and Uname like '%" & Uname.Text & "%' "
        End If
        If Convert.ToString(Ubno.Text).Trim <> "" Then
            sql += " and Ubno like '%" & Ubno.Text & "%' "
        End If
        If Convert.ToString(Intaxno.Text).Trim <> "" Then
            sql += " and Intaxno like '%" & Intaxno.Text & "%' "
        End If
        '已經被新增過訪視紀錄
        sql += " and BDID Not IN (SELECT BDID FROM Bus_VisitInfo WHERE DistID='" & sm.UserInfo.DistID & "') "
        If Convert.ToString(ZipStr).Trim <> "" Then
            sql += ZipStr
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGridTable.Visible = False
        msg.Text = "查無資料<BR><DIV align='LEFT'>(備註：查無資料的原因可能為事業單位已經被新增過訪視紀錄或者勞保局沒有此事業單位的資料。)</DIV>"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""

            'PageControler1.SqlPrimaryKeyDataCreate(sql, "BDID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "BDID"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e) '序號 
                Radio1.Attributes("onclick") = "checkRadio(" & e.Item.ItemIndex + 1 & ");"
                Radio1.Value = drv("BDID")
        End Select
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim BDID As String = ""
        For Each item As DataGridItem In DataGrid1.Items
            Dim Radio1 As HtmlInputRadioButton = item.FindControl("Radio1")
            If Radio1.Checked = True Then
                BDID = Radio1.Value
            End If
        Next

        sql = "SELECT * FROM Bus_BasicData WHERE BDID='" & BDID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)

        If Not dr Is Nothing Then
            Dim ScriptStr As String

            ScriptStr = "<script>"
            ScriptStr += "opener.document.getElementById('Uname').value='" & dr("Uname").ToString & "';" & vbCrLf
            ScriptStr += "opener.document.getElementById('BDID').value='" & dr("BDID").ToString & "';" & vbCrLf
            ScriptStr += "opener.document.getElementById('Intaxno').value='" & dr("Intaxno").ToString & "';" & vbCrLf
            ScriptStr += "opener.document.getElementById('Uname').value='" & dr("Uname").ToString & "';" & vbCrLf
            If dr("Zip").ToString <> "" Then
                ScriptStr += "opener.document.getElementById('City').value='(" & dr("Zip").ToString & ")" & TIMS.Get_ZipName(dr("Zip").ToString) & "';" & vbCrLf
                ScriptStr += "opener.document.getElementById('Zip').value='" & dr("Zip").ToString & "';" & vbCrLf
            End If
            ScriptStr += "opener.document.getElementById('Addr').value='" & dr("Addr").ToString & "';" & vbCrLf
            ScriptStr += "opener.document.getElementById('TradeID').value='" & dr("TradeID").ToString & "';" & vbCrLf
            'ScriptStr += "setRadioValue(opener.document.getElementsByName('KEID'),'" & dr("KEID").ToString & "');" & vbCrLf
            'ScriptStr += "setRadioValue(opener.document.getElementsByName('Labor'),'" & IIf(dr("Labor"), "1", "0") & "');" & vbCrLf
            ScriptStr += "setRadioValue(opener.document.getElementsByName('KEID'),'" & dr("KEID").ToString & "');" & vbCrLf
            ScriptStr += "setRadioValue(opener.document.getElementsByName('Labor'),'" & IIf(dr("Labor"), "1", "0") & "');" & vbCrLf

            ScriptStr += "window.close();"
            ScriptStr += "</script>"

            Me.Page.RegisterStartupScript("return", ScriptStr)
        Else
            Common.MessageBox(Me, "查無資料")
        End If
    End Sub
End Class
