Partial Class SYS_04_007
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    '修正Sys_VisitRate
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
        '檢查Session是否存在 End

        If Not IsPostBack Then
            DataGridTable.Visible = False
            Syear = TIMS.GetSyear(Syear)
            Common.SetListItem(Syear, Now.Year)
            Me.ViewState("Syear") = Now.Year.ToString

            If Me.ViewState("Syear") >= 2009 Then
                Button1.Attributes.Remove("onclick")
                'Button1.Attributes("onclick") = "return Check_Data2();"
            Else
                Button1.Attributes("onclick") = "return Check_Data();"
            End If

            Syear_SelectedIndexChanged(sender, Nothing)
        End If

        Syear.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
    End Sub

    '年度選擇
    Private Sub Syear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Syear.SelectedIndexChanged
        Dim sql As String
        Dim dt As DataTable

        DataGrid1.Visible = False
        DataGrid2.Visible = False

        Me.ViewState("Syear") = Syear.SelectedValue.ToString
        If Me.ViewState("Syear") >= 2009 Then
            Button1.Attributes.Remove("onclick")
            'Button1.Attributes("onclick") = "return Check_Data2();"
        Else
            Button1.Attributes("onclick") = "return Check_Data();"
        End If

        If Syear.SelectedValue >= 2009 Then
            sql = "SELECT * FROM "
            sql += "Key_Plan a "
            sql += "LEFT JOIN (SELECT * FROM Sys_VisitRate WHERE Years='" & Syear.SelectedValue & "') b ON a.TPlanID=b.TPlanID "

            dt = DbAccess.GetDataTable(sql, objconn)
            DataGridTable.Visible = True
            DataGrid2.Visible = True
            DataGrid2.DataSource = dt
            DataGrid2.DataKeyField = "TPlanID"
            DataGrid2.DataBind()
        Else
            sql = "SELECT * FROM "
            sql += "Key_Plan a "
            sql += "LEFT JOIN (SELECT * FROM Sys_VisitRate WHERE Years='" & Syear.SelectedValue & "') b ON a.TPlanID=b.TPlanID "

            dt = DbAccess.GetDataTable(sql, objconn)
            DataGridTable.Visible = True
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "TPlanID"
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lab1 As Label = e.Item.FindControl("Lab1")
                Dim Radio1 As RadioButton = e.Item.FindControl("RadioButton1")
                Dim Radio2 As RadioButton = e.Item.FindControl("RadioButton2")
                Dim Radio3 As RadioButton = e.Item.FindControl("RadioButton3") '新增依機構數
                Dim Radio4 As RadioButton = e.Item.FindControl("RadioButton4") '新增依機構數
                Dim Text1 As TextBox = e.Item.FindControl("TextBox1")
                'Dim Text2 As TextBox = e.Item.FindControl("TextBox2")
                'Dim Text3 As TextBox = e.Item.FindControl("TextBox3") '新增依機構數
                Dim Note As TextBox = e.Item.FindControl("Note")

                'e.Item.ToolTip = "96年度前開班數分署(中心)訪視率依次數, 年度後開班數分署(中心)訪視率改依百分比"
                If drv("Years").ToString <> "" Then
                    If CInt(drv("Years").ToString) <= 2007 Then
                        Radio1.Attributes("onclick") = "set_lab(" & e.Item.ItemIndex + 1 & ");"
                        Radio2.Attributes("onclick") = "set_lab(" & e.Item.ItemIndex + 1 & ");"
                        Radio3.Attributes("onclick") = "set_lab(" & e.Item.ItemIndex + 1 & ");"
                        Radio4.Attributes("onclick") = "set_lab(" & e.Item.ItemIndex + 1 & ");"
                    End If
                End If
                Radio1.Attributes("onclick") += "set_lab2(" & e.Item.ItemIndex + 1 & ");"
                Radio2.Attributes("onclick") += "set_lab2(" & e.Item.ItemIndex + 1 & ");"
                Radio3.Attributes("onclick") += "set_lab2(" & e.Item.ItemIndex + 1 & ");"
                Radio4.Attributes("onclick") += "set_lab2(" & e.Item.ItemIndex + 1 & ");"

                Radio1.Style("display") = "none"
                Radio2.Style("display") = "none"
                Radio3.Style("display") = "none"
                Radio4.Style("display") = "none"
                Select Case drv("Mode1").ToString
                    Case "1"
                        Radio1.Style("display") = "inline"
                        Radio1.Checked = True
                        Text1.Text = drv("Mode1Rate").ToString
                        If drv("Years").ToString <> "" Then
                            If CInt(drv("Years").ToString) <= 2007 Then
                                'lab1.Style("display") = "inline"
                                lab1.Style("display") = "none"
                            End If
                        End If
                    Case "2"
                        Radio2.Style("display") = "inline"
                        Radio2.Checked = True
                        Text1.Text = drv("Mode1Rate").ToString
                    Case "3"
                        Radio3.Style("display") = "inline"
                        Radio3.Checked = True
                        Text1.Text = drv("Mode1Rate").ToString
                    Case Else
                        Radio4.Style("display") = "inline"
                        Radio4.Checked = True
                End Select
                Note.Text = drv("Note").ToString
        End Select
    End Sub

    '2009年訪視計畫用，儲存規則改變
    '2009新版專用 訪視比率設定
    Const Cst_尚未選擇 As Integer = 14 'rb14
    Const Cst_分署訪視item As Integer = 8 'rbc8

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim Checkbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
                Dim d2lab1 As HtmlGenericControl = e.Item.FindControl("d2lab1")
                Dim d2lab2 As HtmlGenericControl = e.Item.FindControl("d2lab2")

                Dim txtNote As TextBox = e.Item.FindControl("txtNote")

                Dim OkFlag As Boolean = False
                OkFlag = False
                d2lab1.Style("display") = "none"
                d2lab2.Style("display") = "none"
                'value2b
                For i As Integer = 1 To Cst_尚未選擇
                    Dim RadioObj As RadioButton = e.Item.FindControl("rb" & CStr(i))
                    RadioObj.Attributes("onclick") += "set_lab2b(" & e.Item.ItemIndex + 1 & ");" 'ROWS
                    RadioObj.Style("display") = "none"
                    If CStr(i) = Convert.ToString(Cst_尚未選擇) And OkFlag = False Then
                        RadioObj.Style("display") = "inline"
                        RadioObj.Checked = True
                    Else
                        If CStr(i) = drv("Mode1").ToString Then
                            If CStr(i) <= 4 Then
                                d2lab1.Style("display") = "inline"
                            End If
                            If CStr(i) > 4 And CStr(i) <= 8 Then
                                d2lab2.Style("display") = "inline"
                            End If
                            RadioObj.Style("display") = "inline"
                            RadioObj.Checked = True
                            OkFlag = True
                        End If
                    End If
                Next

                'value2c
                For i As Integer = 1 To Cst_分署訪視item
                    Dim RadioObj As RadioButton = e.Item.FindControl("rbc" & CStr(i))
                    'RadioObj.Attributes("onclick") += "set_lab2c(" & e.Item.ItemIndex + 1 & ");" 'ROWS
                    'RadioObj.Style("display") = "none"
                    If CStr(i) = drv("Mode1Rate").ToString Then
                        'RadioObj.Style("display") = "inline"
                        RadioObj.Checked = True
                    End If
                Next

                txtNote.Text = drv("Note").ToString

                If drv("Visitor1").ToString <> "" Then
                    Checkbox1.Checked = True
                End If

                If drv("Visitor2").ToString <> "" Then
                    Checkbox2.Checked = True
                End If

        End Select
    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        sql = "SELECT * FROM Sys_VisitRate WHERE Years='" & Me.ViewState("Syear") & "'"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)

        If DataGrid1.Visible Then
            For Each item As DataGridItem In DataGrid1.Items
                Dim Radio1 As RadioButton = item.FindControl("RadioButton1")
                Dim Radio2 As RadioButton = item.FindControl("RadioButton2")
                Dim Radio3 As RadioButton = item.FindControl("RadioButton3")
                Dim Radio4 As RadioButton = item.FindControl("RadioButton4")

                Dim Text1 As TextBox = item.FindControl("TextBox1")
                'Dim Text2 As TextBox = item.FindControl("TextBox2")
                'Dim Text3 As TextBox = item.FindControl("TextBox3")
                Dim Note As TextBox = item.FindControl("Note")

                If (Radio1.Checked Or Radio2.Checked Or Radio3.Checked) Then
                    If dt.Select("TPlanID='" & DataGrid1.DataKeys(item.ItemIndex) & "'").Length = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)

                        dr("Years") = Me.ViewState("Syear")
                        dr("TPlanID") = DataGrid1.DataKeys(item.ItemIndex)
                    Else
                        'dr = dt.Rows(0)
                        dr = dt.Select("TPlanID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0)
                    End If

                    If Radio1.Checked Then
                        dr("Mode1") = 1
                        dr("Mode1Rate") = Text1.Text
                    ElseIf Radio2.Checked Then
                        dr("Mode1") = 2
                        dr("Mode1Rate") = Text1.Text
                    ElseIf Radio3.Checked Then
                        dr("Mode1") = 3
                        dr("Mode1Rate") = Text1.Text
                    End If
                    dr("Note") = IIf(Note.Text = "", Convert.DBNull, Note.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If

                If Radio4.Checked Then
                    If dt.Select("TPlanID='" & DataGrid1.DataKeys(item.ItemIndex) & "'").Length > 0 Then
                        dt.Select("TPlanID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0).Delete()
                    End If
                End If
            Next
        ElseIf DataGrid2.Visible Then
            For Each Item As DataGridItem In DataGrid2.Items
                '所有值為空
                Me.ViewState("Mode1") = ""
                Me.ViewState("Mode1Rate") = ""
                Me.ViewState("Visitor1") = ""
                Me.ViewState("Visitor2") = ""

                '尋找物件，順便給值
                Dim txtNote As TextBox = Item.FindControl("txtNote")
                Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
                Dim Checkbox2 As HtmlInputCheckBox = Item.FindControl("Checkbox2")
                If Checkbox1.Checked Then
                    Me.ViewState("Visitor1") = "1"
                End If
                If Checkbox2.Checked Then
                    Me.ViewState("Visitor2") = "1"
                End If

                'value2b
                For i As Integer = 1 To Cst_尚未選擇 - 1
                    Dim RadioObj As RadioButton = Item.FindControl("rb" & CStr(i))
                    If RadioObj.Checked = True Then
                        Me.ViewState("Mode1") = CStr(i)
                        Exit For
                    End If
                Next

                'value2c
                For i As Integer = 1 To Cst_分署訪視item
                    Dim RadioObj As RadioButton = Item.FindControl("rbc" & CStr(i))
                    If RadioObj.Checked = True Then
                        Me.ViewState("Mode1Rate") = CStr(i)
                        Exit For
                    End If
                Next

                If Me.ViewState("Mode1") <> "" Then
                    If dt.Select("TPlanID='" & DataGrid2.DataKeys(Item.ItemIndex) & "'").Length = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)

                        dr("Years") = Me.ViewState("Syear")
                        dr("TPlanID") = DataGrid2.DataKeys(Item.ItemIndex)
                    Else
                        'dr = dt.Rows(0)
                        dr = dt.Select("TPlanID='" & DataGrid2.DataKeys(Item.ItemIndex) & "'")(0)
                    End If

                    dr("Visitor1") = IIf(Me.ViewState("Visitor1") = "", Convert.DBNull, Me.ViewState("Visitor1"))
                    dr("Visitor2") = IIf(Me.ViewState("Visitor2") = "", Convert.DBNull, Me.ViewState("Visitor2"))
                    dr("Mode1") = Me.ViewState("Mode1")
                    dr("Mode1Rate") = IIf(Me.ViewState("Mode1Rate") = "", 0, Me.ViewState("Mode1Rate")) '可能未選擇
                    dr("Note") = IIf(txtNote.Text = "", Convert.DBNull, txtNote.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                If Me.ViewState("Mode1") = "" Then
                    If dt.Select("TPlanID='" & DataGrid2.DataKeys(Item.ItemIndex) & "'").Length > 0 Then
                        dt.Select("TPlanID='" & DataGrid2.DataKeys(Item.ItemIndex) & "'")(0).Delete()
                    End If
                End If
            Next
        End If

        DbAccess.UpdateDataTable(dt, da)
        'Common.MessageBox(Me, "儲存成功")

        Dim strScript As String
        strScript = "<script language=""javascript"">" & vbCrLf
        strScript += "alert('儲存成功!!');" & vbCrLf
        strScript += "location.href='SYS_04_007.aspx?ID=" & Request("ID") & "';" & vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)

    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Dim dtOld As DataTable
        Dim Sql As String = ""
        Dim oldSYears As String = ""
        Try
            Me.ViewState("Syear") = Syear.SelectedValue.ToString
            oldSYears = CStr(CInt(Me.ViewState("Syear")) - 1)
        Catch ex As Exception
            Errmsg += "上年度資訊有誤，請重新選擇年度" & vbCrLf
        End Try
        Sql = "SELECT * FROM Sys_VisitRate WHERE Years='" & oldSYears & "'"
        dtOld = DbAccess.GetDataTable(Sql, objconn)
        If dtOld.Rows.Count = 0 Then
            Errmsg += "上年度未設定訪視率，請重新選擇年度" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '套用上年度資料
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim dtOld As DataTable '舊年度資料

        Dim dr As DataRow
        Dim dt As DataTable '要填寫新年度資料
        Dim da As SqlDataAdapter = Nothing
        Dim Sql As String = ""

        Dim oldSYears As String = "" '舊年度
        oldSYears = CStr(CInt(Me.ViewState("Syear")) - 1) '舊年度

        Sql = "SELECT * FROM Sys_VisitRate WHERE Years='" & oldSYears & "'" '舊年度
        dtOld = DbAccess.GetDataTable(Sql, objconn) '舊年度資料
        If dtOld.Rows.Count > 0 Then
            Sql = "SELECT * FROM Sys_VisitRate WHERE Years='" & Me.ViewState("Syear") & "'" '新年度
            dt = DbAccess.GetDataTable(Sql, da, objconn) '要填寫新年度資料
            For Each drOld As DataRow In dtOld.Rows
                Dim filter As String = ""
                filter = ""
                filter += "Years='" & Me.ViewState("Syear") & "'" '新年度
                filter += " AND TPlanID='" & drOld("TPlanID") & "'" '舊計畫
                If dt.Select(filter).Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)

                    dr("Years") = Me.ViewState("Syear") '新年度
                    dr("TPlanID") = drOld("TPlanID") '舊計畫

                    dr("Visitor1") = drOld("Visitor1")
                    dr("Visitor2") = drOld("Visitor2")

                    dr("Mode1") = drOld("Mode1")
                    dr("Mode1Rate") = drOld("Mode1Rate")
                    dr("Note") = drOld("Note")
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                'SVRID,Years,TPlanID,Mode1,Mode1Rate,Note,ModifyAcct,ModifyDate,Visitor1,Visitor2
            Next
            DbAccess.UpdateDataTable(dt, da)
        End If

        Dim strScript As String
        strScript = "<script language=""javascript"">" & vbCrLf
        strScript += "alert('套用完成!!');" & vbCrLf
        strScript += "location.href='SYS_04_007.aspx?ID=" & Request("ID") & "';" & vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub
End Class
