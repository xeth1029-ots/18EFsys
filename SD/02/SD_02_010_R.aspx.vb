Partial Class SD_02_010_R
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        TIMS.Tooltip(Button1, "繳費通知單內容by機構，列印請先確認繳費通知單內容是否已儲存!!")

        Dim sql As String = ""
        Dim dt As DataTable

        If Not IsPostBack Then
            '3:只取出正取及備取-KEY_SELRESULT
            SelResult = TIMS.Get_SelResult(SelResult, 3, objconn)
        End If

        Button1.Attributes("onclick") = "javascript:return ReportPrint();"

        TIMS.ShowHistoryClass(Me, historytable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", True)
        If historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Me.table11.Style("display") = "none"
        Me.Button5.Visible = False

        sql = " SELECT * FROM Sys_OrgVar WHERE RID = '" & sm.UserInfo.RID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' AND Itemvar_4 IS NOT NULL "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow
            dr = dt.Rows(0)
            If Not IsDBNull(dr("ItemVar_4")) Then
                Me.msg.Visible = False
                Me.Button1.Enabled = True
            Else
                Me.msg.Visible = True
                Me.Button1.Enabled = False
            End If
        Else
            Me.msg.Visible = True
            Me.Button1.Enabled = False
        End If
    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim M1(5) As Integer

        For i As Integer = 0 To Mailtype1.Items.Count - 1
            M1(i) = If(Mailtype1.Items.Item(i).Selected, 1, 0)
        Next

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "SD_02_010_R", "OCID1='+document.getElementById('OCIDValue1').value+'&DistID=" & sm.UserInfo.DistID & "&Mailtype1=" & M1(0).ToString & "&Mailtype2=" & M1(1).ToString & "&Mailtype3=" & M1(2).ToString & "&Mailtype4=" & M1(3).ToString & "&Mailtype5=" & M1(4).ToString & "&SelResultID='+getRadioValue(document.getElementsByName('SelResult'))+'")
    End Sub

    '設定通知單內容
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.table11.Style("display") = "inline"
        Me.Button5.Visible = True
        Me.msg.Visible = False

        Dim sql As String = ""
        Dim dt As DataTable

        '將參數設定-繳費內容代入----start
        sql = " SELECT * FROM Sys_OrgVar WHERE RID = '" & sm.UserInfo.RID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' AND (Itemvar_4 IS NOT NULL OR Itemvar_5 IS NOT NULL) "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow
            dr = dt.Rows(0)
            If Not IsDBNull(dr("ItemVar_4")) Then
                Me.ItemVar_1.Text = Convert.ToString(dr("ItemVar_4"))
#Region "(No Use)"

                'Else
                '    str1 = " SELECT * FROM Sys_GlobalVar WHERE GVID = '7' AND DistID = '" & sm.UserInfo.DistID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
                '    objtable = DbAccess.GetDataTable(str1)
                '    If objtable.Rows.Count <> 0 Then
                '        dr1 = objtable.Rows(0)
                '        If Not IsDBNull(dr1("ItemVar1")) Then Me.ItemVar_1.Text = Convert.ToString(dr1("ItemVar1"))
                '    End If

#End Region
            End If
            If Not IsDBNull(dr("ItemVar_5")) Then Me.itemvar_2.Text = Convert.ToString(dr("ItemVar_5"))
        Else
            ItemVar_1.Text = ""
            itemvar_2.Text = ""
#Region "(No Use)"

            'str1 = "select * from Sys_GlobalVar where GVID = '7' and DistID='" & sm.UserInfo.DistID & "' and TPlanID='" & sm.UserInfo.TPlanID & "'"
            'objtable = DbAccess.GetDataTable(str1)
            'If objtable.Rows.Count <> 0 Then
            '    dr1 = objtable.Rows(0)
            '    If Not IsDBNull(dr1("ItemVar1")) Then Me.ItemVar_1.Text = Convert.ToString(dr1("ItemVar1"))
            'End If

#End Region
        End If
    End Sub

    '儲存
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim objadapter As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        sql = " SELECT * FROM SYS_ORGVAR WHERE RID = '" & sm.UserInfo.RID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "請先至[首頁>>學員動態管理>>招生作業>> 甄試結果通知單]功能設定[甄試通知單內容],才可使用此功能設定[繳費通知單內容]")
            Exit Sub
        End If

        '檢查
        Dim Errmsg As String = ""
        Errmsg = ""
        If Me.ItemVar_1.Text.Trim <> "" Then Me.ItemVar_1.Text = Me.ItemVar_1.Text.Trim Else Me.ItemVar_1.Text = ""
        If Me.itemvar_2.Text.Trim <> "" Then Me.itemvar_2.Text = Me.itemvar_2.Text.Trim Else Me.itemvar_2.Text = ""
        If Len(Me.ItemVar_1.Text) > 255 Then Errmsg += "正取繳費內容 長度超過系統範圍(255)" & vbCrLf
        If Len(Me.itemvar_2.Text) > 255 Then Errmsg += "備取繳費內容 長度超過系統範圍(255)" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Try
            sql = " SELECT * FROM Sys_OrgVar WHERE RID = '" & sm.UserInfo.RID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
            dt = DbAccess.GetDataTable(sql, objadapter, objconn)
            '儲存
            If dt.Rows.Count > 0 Then
                dr = dt.Rows(0)
                If Me.ItemVar_1.Text <> "" Then dr("ItemVar_4") = Me.ItemVar_1.Text.ToString
                If Me.itemvar_2.Text <> "" Then dr("ItemVar_5") = Me.itemvar_2.Text.ToString
                dr("TPlanID") = sm.UserInfo.TPlanID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now()
                DbAccess.UpdateDataTable(dt, objadapter)
                Me.Button5.Visible = False
                Me.msg.Visible = False
                Me.Button1.Enabled = True
                Common.MessageBox(Me, "儲存成功!!")
            Else
                Common.MessageBox(Me, "查無資料!!")
            End If
        Catch ex As Exception
            Common.MessageBox(Me, "儲存有誤!!")
            Common.MessageBox(Me, ex.ToString)
        End Try
    End Sub
End Class