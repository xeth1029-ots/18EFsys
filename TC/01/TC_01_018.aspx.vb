Public Class TC_01_018
    Inherits AuthBasePage

    Const cst_title1 As String = "材料品項資料匯出"
    Dim sql As String
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not Page.IsPostBack Then
            msg.Text = ""
            PageControler1.Visible = False
            DataGrid1.Visible = False

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            DistrictList = TIMS.Get_DistID(DistrictList)
            Me.DistrictList.Items.Remove(Me.DistrictList.Items.FindByValue(""))
            Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        '當分署(中心)使用者使用時,轄區應該都要鎖死該轄區,不可選擇其它轄區
        Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "0"
            Case "1"
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
            Case Else
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
                'DistrictList.Style.Item("display") = "none"
        End Select

    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Protected Sub bt_export_Click(sender As Object, e As EventArgs) Handles bt_export.Click
        Call sExport1()
    End Sub

    '資料 查詢 [SQL]
    Sub Search1()
        'DistID
        'tmpStr = ""
        '轄區
        Dim itemDist As String = ""
        itemDist = ""
        If DistrictList.SelectedValue <> "" Then
            itemDist = "'" & DistrictList.SelectedValue & "'"
        Else
            For Each objitem As ListItem In DistrictList.Items
                If objitem.Value <> "" AndAlso objitem.Selected = True Then
                    If itemDist <> "" Then itemDist += ","
                    itemDist += "'" & objitem.Value & "'"
                End If
            Next
        End If

        If CNAME.Text <> "" Then CNAME.Text = TIMS.ClearSQM(CNAME.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT  IP.DISTNAME 轄區" & vbCrLf
        sql += " ,'一人份材料' 類別" & vbCrLf
        sql += " ,A.CNAME 品名" & vbCrLf
        sql += " ,A.STANDARD 規格" & vbCrLf
        sql += " ,A.UNIT	單位" & vbCrLf
        sql += " ,A.PRICE	單價" & vbCrLf
        sql += " FROM PLAN_PERSONCOST A" & vbCrLf
        sql += " JOIN VIEW_PLAN IP ON IP.PLANID =A.PLANID " & vbCrLf
        sql += " WHERE 0=0" & vbCrLf
        sql += " and IP.YEARS =@YEARS" & vbCrLf
        If itemDist <> "" Then sql += " and IP.DISTID =@DISTID" & vbCrLf

        If CNAME.Text <> "" Then sql += " and A.CNAME LIKE '%'+@CNAME+'%'" & vbCrLf

        sql += " UNION" & vbCrLf
        sql += " SELECT  IP.DISTNAME 轄區" & vbCrLf
        sql += " ,'共同材料' 類別" & vbCrLf
        sql += " ,A.CNAME 品名" & vbCrLf
        sql += " ,A.STANDARD 規格" & vbCrLf
        sql += " ,A.UNIT	單位" & vbCrLf
        sql += " ,A.PRICE	單價" & vbCrLf
        sql += " FROM PLAN_COMMONCOST A" & vbCrLf
        sql += " JOIN VIEW_PLAN IP ON IP.PLANID =A.PLANID " & vbCrLf
        sql += " WHERE 0=0" & vbCrLf
        sql += " and IP.YEARS =@YEARS" & vbCrLf
        If itemDist <> "" Then sql += " and IP.DISTID =@DISTID" & vbCrLf

        If CNAME.Text <> "" Then sql += " and A.CNAME LIKE '%'+@CNAME+'%'" & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)

        Try
            With oCmd
                .Parameters.Clear()
                .Parameters.Add("YEARS", SqlDbType.VarChar).Value = Me.yearlist.SelectedValue
                If itemDist <> "" Then
                    .Parameters.Add("DISTID", SqlDbType.VarChar).Value = itemDist
                End If
                If CNAME.Text <> "" Then
                    .Parameters.Add("CNAME", SqlDbType.VarChar).Value = CNAME.Text
                End If
                dt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            'Common.RespWrite(Me, Sql)
            'Common.RespWrite(Me, ex.ToString)
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
        End Try


        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    '匯出 資料 1
    Sub sExport1()
        'Dim Errmsg As String = ""
        'Call CheckData1(Errmsg)
        'If Errmsg <> "" Then
        '    Common.MessageBox(Page, Errmsg)
        '    Exit Sub
        'End If

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1()

        If msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If
        If DataGrid1.Visible = False Then
            Common.MessageBox(Page, "查詢資料未正確顯示。")
            Exit Sub
        End If

        Dim sFileName As String = ""
        sFileName = HttpUtility.UrlEncode(cst_title1 & ".xls", System.Text.Encoding.UTF8)

        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集

        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"

        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")

        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid1.AllowPaging = True
        'DataGrid1.Columns(8).Visible = True

    End Sub

End Class