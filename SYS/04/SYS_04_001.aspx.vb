Partial Class SYS_04_001
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    'Dim objconn As SqlConnection
    'Dim objtable As DataTable
    'Dim blnCanAdds As Boolean=False '新增
    'Dim blnCanMod As Boolean=False '修改
    'Dim blnCanDel As Boolean=False '刪除
    'Dim blnCanSech As Boolean=False '查詢
    'Dim blnCanPrnt As Boolean=False '列印

    'SYS_Holiday 
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        PageControler1.PageDataGrid = ShowHDay

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            Dim sdate As String = TIMS.GetSysDate(objconn)
            Start_Date.Text = sdate
            End_Date.Text = sdate
        End If

        'but_seach.Enabled=False
        'If blnCanSech Then but_seach.Enabled=True
        'but_submit.Enabled=False
        'If blnCanAdds Then but_submit.Enabled=True

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    'If sm.UserInfo.RoleID <> 0 Then
        '    'End If
        '    Dim FunDt As DataTable=sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow=FunDt.Select("FunID='" & Request("ID") & "'")
        '    If FunDrArray.Length=0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr=FunDrArray(0)
        '        If FunDr("Sech")="1" Then
        '            but_seach.Enabled=True
        '        Else
        '            but_seach.Enabled=False
        '        End If

        '        If FunDr("Adds")="1" Then
        '            but_submit.Enabled=True
        '        Else
        '            but_submit.Enabled=False
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    'Public Sub CL_Delete(ByVal sender As System.Object, ByVal e As DataGridCommandEventArgs)
    '    Dim delstr As String
    '    Dim literal As DataBoundLiteralControl
    '    literal=e.Item.Cells(0).Controls(0)
    '    delstr="delete SYS_Holiday where HolDate ='" & literal.Text.Trim() & "' and RID='" & sm.UserInfo.RID & "'"
    '    DbAccess.ExecuteNonQuery(delstr)
    '    ShowDayBind()
    'End Sub

    Private Sub but_seach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_seach.Click
        ShowDayBind()
    End Sub

    Sub ShowDayBind()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, ShowHDay)

        Dim objstr As String = ""
        objstr += " SELECT RID" & vbCrLf
        objstr += " ,HOLDATE" & vbCrLf
        objstr += " ,REASON" & vbCrLf
        objstr += " FROM SYS_Holiday" & vbCrLf
        objstr += " WHERE RID='" & sm.UserInfo.RID & "'" & vbCrLf
        objstr += " AND HolDate >=" & TIMS.To_date(Me.Start_Date.Text) & vbCrLf
        objstr += " AND HolDate <=" & TIMS.To_date(Me.End_Date.Text) & vbCrLf

        msg.Text = "查無資料"
        DataGridTable.Visible = False

        Dim dt As DataTable = DbAccess.GetDataTable(objstr, objconn)

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            'PageControler1.SqlDataCreate(objstr, "HolDate")
            PageControler1.PageDataTable = dt
            PageControler1.Sort = "HolDate"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub but_submit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_submit.Click
        Dim url1 As String = ""
        url1 = TIMS.Get_Url1(Me, "SYS_04_001_add.aspx")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub ShowHDay_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles ShowHDay.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument

        Select Case e.CommandName
            Case "Del"
                Dim sHolDate As String = TIMS.GetMyValue(sCmdArg, "HolDate")
                Dim sRID As String = TIMS.GetMyValue(sCmdArg, "RID")
                If sHolDate = "" Then Exit Sub
                If sRID = "" Then Exit Sub

                Dim delstr As String = ""
                delstr = "delete SYS_Holiday where HolDate=" & TIMS.To_date(sHolDate) & " and RID='" & sRID & "'"
                DbAccess.ExecuteNonQuery(delstr, objconn)

                ShowDayBind()
        End Select
    End Sub

    Private Sub ShowHDay_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles ShowHDay.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim drv As DataRowView = e.Item.DataItem

            Dim sCmdArg As String = ""
            TIMS.SetMyValue(sCmdArg, "HolDate", TIMS.Cdate3(drv("HolDate")))
            TIMS.SetMyValue(sCmdArg, "RID", Convert.ToString(drv("RID")))

            Dim labHolDate As Label = e.Item.FindControl("labHolDate")
            labHolDate.Text = TIMS.Cdate3(drv("HolDate")) 'Common.FormatDate(drv("HolDate"))

            Dim LabReason As Label = e.Item.FindControl("LabReason")
            LabReason.Text = Convert.ToString(drv("Reason"))

            Dim myDeleteButton As LinkButton = e.Item.FindControl("lbtDel")
            myDeleteButton.Attributes.Add("onclick", "return confirm('您確定要刪除嗎?');")
            myDeleteButton.CommandArgument = sCmdArg 'drv("HolDate")

            'myDeleteButton.Enabled=False
            'If blnCanDel Then myDeleteButton.Enabled=True

            'If FunDr("Del")="1" Then
            '    myDeleteButton.Enabled=True
            'Else
            '    myDeleteButton.Enabled=False
            'End If
        End If
    End Sub

End Class
