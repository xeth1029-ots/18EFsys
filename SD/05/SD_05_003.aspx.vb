Partial Class SD_05_003
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End
        '分頁設定---------------Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定---------------End

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = 1 Then
        '            Button3.Enabled = True
        '        Else
        '            Button3.Enabled = False
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            Button2.Enabled = True
        '        Else
        '            Button2.Enabled = False
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

        If Not IsPostBack Then
            msg.Text = ""
            Button2.Attributes("onclick") = "javascript:return search()"

            Table4.Visible = False
            start_date.Text = Now.AddMonths(-3).Date
            end_date.Text = Now.Date
        End If


    End Sub

    Sub search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.STCID" & vbCrLf
        sql &= " ,c.Name" & vbCrLf
        sql &= " ,d.CyclType CyclType1" & vbCrLf
        sql &= " ,d.LevelType LevelType1" & vbCrLf
        sql &= " ,a.Reason" & vbCrLf
        sql &= " ,a.ApplyDate" & vbCrLf
        sql &= " ,a.NewClassID" & vbCrLf
        sql &= " ,a.OrigClassID" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,b.StudentID" & vbCrLf
        sql &= " ,e.ClassCName ClassCName2" & vbCrLf
        sql &= " ,e.LevelType LevelType2" & vbCrLf
        sql &= " ,e.CyclType CyclType2" & vbCrLf
        sql &= " ,d.ClassCName ClassCName1" & vbCrLf
        sql &= " ,e.IsClosed" & vbCrLf
        sql &= " ,f.ClassID" & vbCrLf
        sql &= " FROM Stud_TranClassRecord a" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass b ON a.SOCID=b.SOCID" & vbCrLf
        sql &= " JOIN Stud_StudentInfo c ON b.SID=c.SID" & vbCrLf
        sql &= " JOIN Class_ClassInfo d ON d.OCID=a.OrigClassID" & vbCrLf
        sql &= " JOIN Class_ClassInfo e ON e.OCID=a.NewClassID" & vbCrLf
        sql &= " join ID_Class f ON e.CLSID=f.CLSID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND e.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        sql &= " and e.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql &= " AND a.ApplyDate>=" & TIMS.to_date(start_date.Text) & vbCrLf
        sql &= " AND a.ApplyDate<=" & TIMS.to_date(end_date.Text) & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!"
        Table4.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table4.Visible = True

            DataGrid1.DataKeyField = "STCID"

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "STCID"
            PageControler1.Sort = "ClassID,CyclType1,StudentID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "back"
                If e.CommandArgument = "" Then Exit Sub
                Dim STCID As String = e.CommandArgument

                '抓取原始OCID、新的OCID、還有學生的SCID
                Dim sql As String = ""
                sql = "SELECT * FROM Stud_TranClassRecord WHERE STCID='" & STCID & "'"
                Dim drResult As DataRow = DbAccess.GetOneRow(sql, objconn)
                If drResult Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                Dim Old_OCID As String = drResult("OrigClassID")
                Dim New_OCID As String = drResult("NewClassID")
                Dim Stu_SOCID As String = drResult("SOCID")

                '檢查要回復的班級是否已經結訓
                sql = "SELECT IsClosed FROM Class_ClassInfo WHERE OCID='" & Old_OCID & "'"
                Dim drC As DataRow = DbAccess.GetOneRow(sql, objconn)
                If Convert.ToString(drC("IsClosed")) <> "N" Then
                    Common.MessageBox(Me, "此學員的原來班級已經結訓，不可回復")
                    Exit Sub
                End If

                '更新Class_StudentsOfClass的OCID
                Dim da As SqlDataAdapter = Nothing
                Dim tConn As SqlConnection = DbAccess.GetConnection()
                Dim oTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
                Try

                    'TIMS.TestDbConn(Me, conn)
                    'objTrans = DbAccess.BeginTrans(conn)

                    sql = "SELECT SOCID,OCID,ModifyAcct,ModifyDate FROM Class_StudentsOfClass WHERE SOCID='" & Stu_SOCID & "'"
                    Dim dtResult As DataTable = DbAccess.GetDataTable(sql, da, oTrans)
                    Dim dr As DataRow = dtResult.Rows(0)
                    dr("OCID") = Old_OCID
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dtResult, da, oTrans)

                    sql = "DELETE Stud_TranClassRecord WHERE STCID='" & STCID & "'"
                    DbAccess.ExecuteNonQuery(sql, oTrans)
                    DbAccess.CommitTrans(oTrans)

                    Common.MessageBox(Me, "回復成功")

                    Call search1()
                    'Button2_Click(Button2, e)
                Catch ex As Exception
                    DbAccess.RollbackTrans(oTrans)
                    TIMS.CloseDbConn(tConn)
                    Throw ex
                End Try
                TIMS.CloseDbConn(tConn)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DataGrid1.PageSize * DataGrid1.CurrentPageIndex
            Dim drv As DataRowView = e.Item.DataItem

            If CInt(e.Item.Cells(7).Text) <> 0 Then
                e.Item.Cells(2).Text += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(7).Text)) & "期"
            End If
            If CInt(e.Item.Cells(9).Text) <> 0 Then
                e.Item.Cells(4).Text += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(9).Text)) & "期"
            End If

            Dim Button1 As Button = e.Item.FindControl("Button1") 'back Button1
            Button1.CommandArgument = Convert.ToString(drv("STCID"))

            If drv("IsClosed").ToString = "Y" Then
                Button1.Enabled = False
                TIMS.Tooltip(Button1, "班級已結訓,不提供回復功能")
            End If

            Button1.Attributes("onclick") = "return confirm('確定要回復此學員的班級?');"
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        KeepSearch()
        'Response.Redirect("SD_05_003_add.aspx?ID=" & Request("ID") & "")
        Dim url1 As String = "SD_05_003_add.aspx?ID=" & Request("ID") & ""
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Sub KeepSearch()
        Session("_search") = "start_date=" & start_date.Text
        Session("_search") = "&end_date=" & end_date.Text
        Session("_search") = "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        If DataGrid1.Visible = True Then
            Session("_search") = "&submit=1"
        Else
            Session("_search") = "&submit=0"
        End If
    End Sub
End Class
