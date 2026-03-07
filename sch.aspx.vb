Public Class sch
    Inherits AuthBasePage

    Dim ff33 As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        ''分頁設定 Start
        'PageControler1.PageGridView = GridView1
        ''分頁設定 End

        If Not IsPostBack Then ddlFun = TIMS.Get_ddlFunction(ddlFun)
    End Sub

    Protected Sub btnSch_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSch.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim strChkInput As String = ""
        Dim rtnPath As String = Request.FilePath

        strChkInput = ddlFun.SelectedValue.Trim
        If strChkInput = "SYS" Then strChkInput = "[SYS]"  '(避開系統認為的危險性字元，by:20181030)
        If TIMS.CheckInput(strChkInput) And strChkInput <> "[SYS]" Then  'edit，by:20181030
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)  'edit，by:20181030
            Exit Sub
        End If

        strChkInput = txtFunName.Text.Trim
        If TIMS.CheckInput(strChkInput) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)  'edit，by:20181030
            Exit Sub
        End If

        Call search1()
    End Sub

    Sub search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim vFun As String = TIMS.ClearSQM(ddlFun.SelectedValue) '功能種類
        Dim vFunName As String = TIMS.ClearSQM(txtFunName.Text) '功能名稱
        If vFun <> "" Then vFun = UCase(vFun)
        'txtFunName.Text >> vFunName 
        If vFunName <> "" Then vFunName = UCase(vFunName) '轉換大寫

        'Dim dt As New DataTable '所有功能(比對使用)
        Dim dt2 As New DataTable '刪除用(顯示用)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.FunID" & vbCrLf
        sql &= " ,REPLACE(a.funPath3,'>>','>') sysKindName" & vbCrLf
        sql &= " ,a.Name aName" & vbCrLf
        sql &= " ,a.Kind" & vbCrLf
        sql &= " ,a.SPage" & vbCrLf
        sql &= " ,a.PARENT PARENT1" & vbCrLf
        sql &= " ,a.LEVELS" & vbCrLf
        sql &= " FROM VIEW_FUNCTION a" & vbCrLf
        sql &= " WHERE a.Valid = 'Y'" & vbCrLf
        sql &= " AND a.SPage IS NOT NULL" & vbCrLf
        If vFun <> "" Then sql &= " AND a.Kind = @Kind " & vbCrLf '功能種類
        If vFunName <> "" Then '功能名稱
            'Dim ssfunName As String = UCase(vFunName) '轉換大寫
            sql &= " AND (1!=1 " & vbCrLf
            sql &= "  OR UPPER(a.SPage) LIKE N'%" & vFunName & "%'" & vbCrLf
            sql &= "  OR UPPER(a.pName) LIKE N'%" & vFunName & "%'" & vbCrLf
            sql &= "  OR UPPER(a.Name) LIKE N'%" & vFunName & "%'" & vbCrLf
            sql &= " ) " & vbCrLf
            'sql &= " AND (a.pName LIKE '%' + @pName+'%' OR a.Name LIKE '%' + @Name + '%') " & vbCrLf
        End If
        sql &= " ORDER BY a.kind ,psort ,a.levels ,a.sort " & vbCrLf
        'TIMS.writeLog(Me, sql)
        Dim oCmd2 As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        With oCmd2
            .Parameters.Clear()
            If vFun <> "" Then .Parameters.Add("Kind", SqlDbType.VarChar).Value = vFun
            dt2.Load(.ExecuteReader()) '刪除用
        End With
        If dt2.Rows.Count = 0 Then
            labMsg.Text = "查無資料"
            GridView1.Visible = False
            Exit Sub
        End If

        Dim dt As New DataTable '所有功能(比對使用)
        With oCmd2
            .Parameters.Clear()
            If vFun <> "" Then .Parameters.Add("Kind", SqlDbType.VarChar).Value = vFun
            dt.Load(.ExecuteReader()) '所有功能
        End With

        '有權使用的功能
        Dim dtCanUseDt As DataTable = TIMS.sGet_CanUseSchDt(objconn)
        If dtCanUseDt Is Nothing Then
            labMsg.Text = "查無資料"
            GridView1.Visible = False
            Exit Sub
        End If
        '父層/子層過濾
        dtCanUseDt.DefaultView.RowFilter = "Valid='Y'"
        dtCanUseDt = TIMS.dv2dt(dtCanUseDt.DefaultView)
        If dtCanUseDt.Rows.Count = 0 Then
            labMsg.Text = "查無資料"
            GridView1.Visible = False
            Exit Sub
        End If

        For Each dr1 As DataRow In dt.Rows
            ff33 = "FunID='" & dr1("FunID") & "'"
            '該功能 無權使用。執行刪除。
            If dtCanUseDt.Select(ff33).Length = 0 Then
                If dt2.Select(ff33).Length > 0 Then dt2.Select(ff33)(0).Delete()
            End If

            '父層權限確認
            Dim tmpPARENT As String = Convert.ToString(dr1("PARENT1"))
            If tmpPARENT <> "" AndAlso tmpPARENT <> "0" Then
                ff33 = "FunID='" & tmpPARENT & "'" '查無子項功能權限
                If dtCanUseDt.Select(ff33).Length = 0 Then
                    ff33 = "PARENT1='" & tmpPARENT & "'" '刪除所有父層功能 
                    If dt2.Select(ff33).Length > 0 Then
                        For Each dr2 As DataRow In dt2.Select(ff33)
                            dr2.Delete() '該功能 無權使用。執行刪除。
                        Next
                    End If
                End If
            End If
        Next
        '確認後輸出
        dt2.AcceptChanges()

        labMsg.Text = "查無資料"
        GridView1.Visible = False
        If dt2.Rows.Count > 0 Then
            labMsg.Text = ""
            GridView1.Visible = True
            GridView1.DataSource = dt2
            GridView1.DataBind()
        End If
    End Sub

    Private Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        If e.CommandArgument = "" Then Exit Sub
        Select Case LCase(e.CommandName)
            Case "viewcmd"
                'dr("SPage").ToString & "?ID=" & dr("FunID").ToString
                Dim sFunID As String = TIMS.GetMyValue(e.CommandArgument, "FunID")
                Dim sSPage As String = TIMS.GetMyValue(e.CommandArgument, "SPage")
                'Response.Redirect(sSPage & "?ID=" & sFunID)
                Dim flag_new_windows As Boolean = TIMS.CHK_FUNCSPAGE2(sm, sSPage)
                If flag_new_windows Then
                    TIMS.OpenWin1(Me, sSPage)
                Else
                    If sSPage.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then sSPage = sSPage.Substring(0, sSPage.Length - 5)
                    Dim url1 As String = sSPage & "?ID=" & sFunID
                    TIMS.Utl_Redirect(Me, objconn, url1)
                End If
        End Select
    End Sub

    Private Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.Header
                'e.Row.Attributes.Add("style", "display:none")
            Case DataControlRowType.DataRow
                Dim drv As DataRowView = e.Row.DataItem
                Dim lbsysKindName As LinkButton = e.Row.FindControl("lbsysKindName")
                lbsysKindName.Text = Convert.ToString(drv("sysKindName"))
                Dim cmdArg As String = ""
                cmdArg += "&FunID=" & Convert.ToString(drv("FunID"))
                cmdArg += "&SPage=" & Convert.ToString(drv("SPage"))
                lbsysKindName.CommandArgument = cmdArg
        End Select
    End Sub
End Class