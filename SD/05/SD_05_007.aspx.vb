Partial Class SD_05_007
    Inherits AuthBasePage

   'Dim FunDr As DataRow
    Dim Days1 As Integer
    Dim Days2 As Integer

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        msg.Text = ""
        If Not IsPostBack Then
            SanID = TIMS.Get_Sanction(SanID)
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            end_date.Text = Now.Date
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
        Button1.Attributes("onclick") = "javascript:return search();"

        '取出設定天數檔------------------------------Start
        Dim sql As String
        Dim dr As DataRow
        sql = "SELECT * FROM Sys_Days"
        dr = DbAccess.GetOneRow(sql)
        If Not dr Is Nothing Then
            Days1 = dr("Days1")
            Days2 = dr("Days2")
        End If
        '取出設定天數檔------------------------------End

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        'End If
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
        '            Button2.Enabled = True
        '        Else
        '            Button2.Enabled = False
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

        '分頁設定---------------Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定---------------End

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button6.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');SetOneOCID();"
        Else
            Button6.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');SetOneOCID();"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim RIDStr As String = ""
        Dim OCIDStr As String = ""
        Dim DateStr As String = ""
        Dim SanStr As String = ""

        If RIDValue.Value <> "" Then
            RIDStr = " and RID='" & RIDValue.Value & "'"
        Else
            RIDStr = " and RID='" & sm.UserInfo.RID & "'"
        End If

        If OCIDValue1.Value <> "" Then
            OCIDStr = " and OCID='" & OCIDValue1.Value & "'"
        End If
        Dim CJOBStr As String = "" '& CJOBStr
        If cjobValue.Value <> "" Then
            CJOBStr += " and CJOB_UNKEY=" & cjobValue.Value & vbCrLf
        End If

        If start_date.Text <> "" And end_date.Text <> "" Then
            DateStr = " and SanDate>=convert(datetime, '" & start_date.Text & "', 111) and SanDate<=convert(datetime, '" & end_date.Text & "', 111)"
        End If

        If SanID.SelectedIndex <> 0 Then
            SanStr = " and SanID='" & SanID.SelectedValue & "'"
        End If

        Dim sql As String
        'Dim MySqlStr As String
        sql = ""
        sql &= " SELECT   c.Name, d.SanID, d.SeqNo, d.SanDate, d.SOCID, d.Times, b.StudentID, b.StudStatus"
        sql &= " , a.OCID, a.ClassCName, a.CyclType, a.LevelType,a.IsClosed,a.FTDate,e.ClassID "
        sql += " FROM     Stud_StudentInfo c, "
        sql += "         (SELECT * FROM Class_StudentsOfClass WHERE 1=1" & OCIDStr & ") b, "
        sql += "         (SELECT * FROM Class_ClassInfo WHERE PlanID='" & sm.UserInfo.PlanID & "'" & OCIDStr & RIDStr & CJOBStr & ") a, "
        sql += "         (SELECT * FROM Stud_Sanction WHERE 1=1" & DateStr & SanStr & ") d, "
        sql += "         ID_Class e "
        sql += " WHERE    c.SID=b.SID AND b.OCID=a.OCID AND d.SOCID=b.SOCID and a.CLSID=e.CLSID"

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql)

        Table4.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            If OCIDValue1.Value = "" Then
                DataGrid1.Columns(1).Visible = True
            Else
                DataGrid1.Columns(1).Visible = False
            End If
            Table4.Visible = True
            msg.Text = ""

            PageControler1.PageDataTable = dt
            PageControler1.Sort = "ClassID,CyclType,StudentID"
            PageControler1.ControlerLoad()
        End If

        'If TIMS.Get_SQLRecordCount(sql) = 0 Then
        '    Table4.Visible = False
        '    msg.Text = "查無資料!!"
        'Else
        '    If OCIDValue1.Value = "" Then
        '        DataGrid1.Columns(1).Visible = True
        '    Else
        '        DataGrid1.Columns(1).Visible = False
        '    End If
        '    Table4.Visible = True

        '    PageControler1.SqlString = sql
        '    PageControler1.Sort = "ClassID,CyclType,StudentID"
        '    PageControler1.ControlerLoad()
        'End If


        'Dim dt As DataTable
        'dt = DbAccess.GetDataTable(sql)
        'If dt.Rows.Count = 0 Then
        '    Table4.Visible = False
        '    msg.Text = "查無資料!!"
        'Else
        '    If OCIDValue1.Value = "" Then
        '        DataGrid1.Columns(1).Visible = True
        '    Else
        '        DataGrid1.Columns(1).Visible = False
        '    End If
        '    Table4.Visible = True
        '    DataGrid1.DataSource = dt
        '    DataGrid1.DataBind()


        '    '分頁用--------------------------------------------Start
        '    DataGridPage1.MyDataTable = dt
        '    DataGridPage1.FirstTime()
        '    '分頁用--------------------------------------------End
        'End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = (DataGrid1.CurrentPageIndex * DataGrid1.PageSize) + e.Item.ItemIndex + 1

            Dim dr As DataRowView = e.Item.DataItem

            e.Item.Cells(1).Text = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))

            For i As Integer = 0 To SanID.Items.Count - 1
                If e.Item.Cells(4).Text = SanID.Items(i).Value Then
                    e.Item.Cells(4).Text = SanID.Items(i).Text
                End If
            Next
            Dim mybut1 As Button = e.Item.FindControl("Button3")
            Dim mybut2 As Button = e.Item.FindControl("Button4")
            Dim mybut3 As Button = e.Item.FindControl("Button5")

            mybut1.CommandArgument = "SD_05_007_add.aspx?ID=" & Request("ID") & "&Proecess=edit&SOCID=" & dr("SOCID") & "&SanDate=" & dr("SanDate") & "&SeqNo=" & dr("SeqNo") & "&OCID=" & dr("OCID")
            mybut2.CommandArgument = "SOCID=" & dr("SOCID") & "&SanDate" & dr("SanDate") & "&SeqNo=" & dr("SeqNo")
            mybut3.CommandArgument = "SD_05_007_add.aspx?ID=" & Request("ID") & "&Proecess=view&SOCID=" & dr("SOCID") & "&SanDate=" & dr("SanDate") & "&SeqNo=" & dr("SeqNo") & "&OCID=" & dr("OCID")
            mybut2.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"

            Select Case dr("StudStatus")
                Case 2, 3
                    mybut1.Visible = False
                    mybut2.Visible = False
                    'If FunDr("Sech") = 1 Then
                    '    mybut3.Enabled = True
                    'Else
                    '    mybut3.Enabled = False
                    'End If
                Case Else
                    If dr("IsClosed") = "Y" Then
                        Select Case sm.UserInfo.RoleID
                            Case 0, 1

                                '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天
                                '暫時先改這樣,以後還會再改
                                If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                    If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 75 Then
                                        mybut1.Visible = False
                                        mybut2.Visible = False
                                        mybut3.Visible = True

                                        'If FunDr("Sech") = 1 Then
                                        '    mybut3.Enabled = True
                                        'Else
                                        '    mybut3.Enabled = False
                                        'End If
                                    Else
                                        mybut1.Visible = True
                                        mybut2.Visible = True
                                        mybut3.Visible = False

                                        'If FunDr("Mod") = 1 Then
                                        '    mybut1.Enabled = True
                                        'Else
                                        '    mybut1.Enabled = False
                                        'End If

                                        'If FunDr("Del") = 1 Then
                                        '    mybut2.Enabled = True
                                        'Else
                                        '    mybut2.Enabled = False
                                        'End If
                                    End If
                                Else
                                    If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days2 Then
                                        mybut1.Visible = False
                                        mybut2.Visible = False
                                        mybut3.Visible = True

                                        'If FunDr("Sech") = 1 Then
                                        '    mybut3.Enabled = True
                                        'Else
                                        '    mybut3.Enabled = False
                                        'End If
                                    Else
                                        mybut1.Visible = True
                                        mybut2.Visible = True
                                        mybut3.Visible = False

                                        'If FunDr("Mod") = 1 Then
                                        '    mybut1.Enabled = True
                                        'Else
                                        '    mybut1.Enabled = False
                                        'End If

                                        'If FunDr("Del") = 1 Then
                                        '    mybut2.Enabled = True
                                        'Else
                                        '    mybut2.Enabled = False
                                        'End If
                                    End If
                                End If


                            Case Else
                                '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                                '暫時先改這樣,以後還會再改
                                If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                    If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 60 Then
                                        mybut1.Visible = False
                                        mybut2.Visible = False
                                        mybut3.Visible = True

                                        'If FunDr("Sech") = 1 Then
                                        '    mybut3.Enabled = True
                                        'Else
                                        '    mybut3.Enabled = False
                                        'End If
                                    Else
                                        mybut1.Visible = True
                                        mybut2.Visible = True
                                        mybut3.Visible = False

                                        'If FunDr("Mod") = 1 Then
                                        '    mybut1.Enabled = True
                                        'Else
                                        '    mybut1.Enabled = False
                                        'End If

                                        'If FunDr("Del") = 1 Then
                                        '    mybut2.Enabled = True
                                        'Else
                                        '    mybut2.Enabled = False
                                        'End If
                                    End If

                                Else
                                    If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days1 Then
                                        mybut1.Visible = False
                                        mybut2.Visible = False
                                        mybut3.Visible = True

                                        'If FunDr("Sech") = 1 Then
                                        '    mybut3.Enabled = True
                                        'Else
                                        '    mybut3.Enabled = False
                                        'End If
                                    Else
                                        mybut1.Visible = True
                                        mybut2.Visible = True
                                        mybut3.Visible = False

                                        'If FunDr("Mod") = 1 Then
                                        '    mybut1.Enabled = True
                                        'Else
                                        '    mybut1.Enabled = False
                                        'End If

                                        'If FunDr("Del") = 1 Then
                                        '    mybut2.Enabled = True
                                        'Else
                                        '    mybut2.Enabled = False
                                        'End If
                                    End If
                                End If

                        End Select

                    Else
                        mybut1.Visible = True
                        mybut2.Visible = True
                        mybut3.Visible = False

                        'If FunDr("Mod") = 1 Then
                        '    mybut1.Enabled = True
                        'Else
                        '    mybut1.Enabled = False
                        'End If

                        'If FunDr("Del") = 1 Then
                        '    mybut2.Enabled = True
                        'Else
                        '    mybut2.Enabled = False
                        'End If
                    End If
            End Select

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TIMS.Utl_Redirect1(Me, "SD_05_007_add.aspx?ID=" & Request("ID") & "&Proecess=add")
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "edit" Or e.CommandName = "view" Then
            TIMS.Utl_Redirect1(Me, e.CommandArgument)
        ElseIf e.CommandName = "del" Then
            Dim sql As String
            sql = "DELETE Stud_Sanction WHERE SOCID='" & DataGrid1.Items(e.Item.ItemIndex).Cells(8).Text & "' and SanDate=convert(datetime, '" & DataGrid1.Items(e.Item.ItemIndex).Cells(6).Text & "', 111) and SeqNo='" & DataGrid1.Items(e.Item.ItemIndex).Cells(9).Text & "'"
            DbAccess.ExecuteNonQuery(sql)
            Button1_Click(Button1, e)
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        Dim dr As DataRow

        '判斷機構是否只有一個班級
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value)

        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGrid1.Visible = False
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                DataGrid1.Visible = False
            End If
        End If

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
