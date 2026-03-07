Partial Class SD_09_016_R
    Inherits AuthBasePage

    'from Stud_TrainCostD
    'SD_09_016_R2
    'SD_09_016_R2_subreport1
    'SD_09_016_R
    'SD_09_016_R_subreport1

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'Me.ClassPanel.Visible = True
            PageControler1.Visible = False
            Print.Visible = True
            Save.Visible = False
            msg.Visible = False

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button12_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, , True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '列印
        Print.Attributes("onclick") = "return CheckSPrint();"
        '查詢
        Query.Attributes("onclick") = "return CheckSPrint();" '"return CheckSPrint('Query');"
        '儲存
        Save.Attributes("onclick") = "return Check_date();"
    End Sub

    '查詢
    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Query.Click '查詢
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        Dim sql As String = ""
        sql += " select cs.socid" & vbCrLf
        sql += " ,cs.StudentID" & vbCrLf
        sql += " ,ss.sid" & vbCrLf
        sql += " ,st.OtherJobCost" & vbCrLf
        sql += " ,ss.Name" & vbCrLf
        'sql += " ,case when ss.Sex = 'F' then '女' when ss.sex = 'M' then '男' end as Sex" & vbCrLf
        sql += " ,dbo.DECODE6(ss.Sex,'F','女','M','男',ss.Sex) Sex" & vbCrLf
        sql += " ,ki.Name MIdentityID" & vbCrLf
        sql += " ,ss.IDNO" & vbCrLf
        sql += " ,CONVERT(varchar, ss.birthday, 111) birthday" & vbCrLf
        sql += " ,cs.WorkSuppIdent" & vbCrLf
        sql += " ,cs.PMode " & vbCrLf
        sql += " ,st2.SumOfMoney" & vbCrLf
        sql += " ,st2.PayMoney" & vbCrLf
        sql += " ,st2.BudID" & vbCrLf
        sql += " ,st2.AppliedStatusM " & vbCrLf
        sql += " from Stud_StudentInfo ss " & vbCrLf
        sql += " join Class_StudentsOfClass cs on ss.sid = cs.sid " & vbCrLf
        sql += " join class_classinfo cc on cs.ocid = cc.ocid " & vbCrLf
        sql += " join Plan_PlanInfo pp  on cc.planid = pp.planid and cc.rid = pp.rid and cc.seqno = pp.seqno " & vbCrLf
        sql += " join id_Plan ip on ip.planid =pp.planid " & vbCrLf
        sql += " join Key_Identity ki ON ki.IdentityID = cs.MIdentityID " & vbCrLf
        sql += " left join Stud_TrainCostD st on st.socid = cs.socid " & vbCrLf
        sql += " LEFT JOIN Stud_SubsidyCost st2 ON st2.SOCID=cs.SOCID  " & vbCrLf
        sql += " where 1=1" & vbCrLf
        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        '採自費/公費欄位判斷
        'and (cs.PMode = 1 or cs.PMode is NUll) '改為 自費生 永遠排除。20141002 BY AMU
        If Not TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql += " and (cs.PMode = 1 or cs.PMode is NUll) " & vbCrLf
        End If
        sql += " and cs.studstatus not in (2,3) " & vbCrLf
        sql += " and ip.TPlanID = '" & sm.UserInfo.TPlanID & "'" & vbCrLf
        sql += " and ip.Years = '" & sm.UserInfo.Years & "'" & vbCrLf
        If RIDValue.Value <> "" Then
            sql += " and cc.RID like '" & RIDValue.Value & "%'" & vbCrLf
        Else
            sql += " and cc.RID like '" & sm.UserInfo.RID & "%'" & vbCrLf
        End If
        If OCIDValue1.Value.ToString <> "" Then
            sql += " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If
        sql += " ORDER BY cs.StudentID " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!!"
        Table4.Visible = False
        PageControler1.Visible = False
        Save.Visible = False
        msg.Visible = True

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table4.Visible = True
            PageControler1.Visible = True
            Save.Visible = True
            msg.Visible = False

            'PageControler1.SqlString = sql
            'PageControler1.PrimaryKey = "socid"
            'PageControler1.Sort = "sid"
            PageControler1.PageDataTable = dt
            PageControler1.Sort = "StudentID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Footer
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"

            Case Else
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim OtherJobCost As TextBox = e.Item.FindControl("OtherJobCost")
                Dim Pmode2 As Label = e.Item.FindControl("Pmode2")

                Pmode2.Visible = False
                OtherJobCost.Text = ""
                If Convert.ToString(drv("OtherJobCost")) <> "" Then
                    OtherJobCost.Text = drv("OtherJobCost")
                    Checkbox1.Checked = True
                End If

                '自費
                If Convert.ToString(drv("PMode")) = "2" Then
                    Const cst_msg1 As String = "該學員使用*自費"
                    Pmode2.Visible = True
                    TIMS.Tooltip(OtherJobCost, cst_msg1, True)
                    TIMS.Tooltip(Pmode2, cst_msg1, True)
                End If

                OtherJobCost.Enabled = False
                If Checkbox1.Checked Then
                    OtherJobCost.Enabled = True '可輸入補助金額喔
                End If

                '學員經費審核狀態-申請 Y:成功
                If Convert.ToString(drv("AppliedStatusM")) = "Y" Then
                    '查看補助申請
                    If Convert.ToString(drv("SumOfMoney")) <> "" Then
                        OtherJobCost.Text = drv("SumOfMoney")

                        Checkbox1.Checked = True
                        Checkbox1.Disabled = True

                        TIMS.Tooltip(Checkbox1, "補助申請", True)
                        TIMS.Tooltip(OtherJobCost, "補助申請", True)
                    End If
                End If

                Checkbox1.Attributes("onclick") = "Check1();" '如果勾選學員
        End Select

    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        PageControler1.Visible = False
        Save.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If OCIDValue1.Value = "" Then
            Errmsg += "班別代碼有誤，請確認點選職類/班別" & vbCrLf
        Else
            Dim i As Integer = 0
            i = 0
            For Each Item As DataGridItem In DataGrid1.Items
                i += 1
                Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
                Dim OtherJobCost As TextBox = Item.FindControl("OtherJobCost")

                Dim IDNOv As String = Convert.ToString(Item.Cells(4).Text)    'IDNO
                Dim SOCIDv As String = Convert.ToString(Checkbox1.Value)   'socid
                If Checkbox1.Checked = True Then
                    If IDNOv <> "" AndAlso SOCIDv <> "" Then
                    Else
                        Errmsg += "第" & CStr(i) & "筆:學員選擇有誤，請重新選擇!!" & vbCrLf
                        Exit For
                    End If
                    If OtherJobCost.Text <> "" AndAlso IsNumeric(OtherJobCost.Text) Then
                        Try
                            OtherJobCost.Text = CInt(OtherJobCost.Text)
                        Catch ex As Exception
                            Errmsg += "第" & CStr(i) & "筆:金額輸入有誤，應輸入數字!!" & vbCrLf
                            Exit For
                        End Try
                    Else
                        Errmsg += "第" & CStr(i) & "筆:金額輸入有誤，應輸入數字!!" & vbCrLf
                        Exit For
                    End If
                Else
                    If Trim(OtherJobCost.Text) <> "" Then
                        OtherJobCost.Text = Trim(OtherJobCost.Text)
                        Errmsg += "第" & CStr(i) & "筆:金額輸入有誤，不應輸入文字!!" & vbCrLf
                        Exit For
                    Else
                        OtherJobCost.Text = ""
                    End If

                End If
            Next
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        '儲存狀態正常 預設true:正常
        Dim flagSaveOk As Boolean = True
        flagSaveOk = True

        'Dim sql1 As String = ""
        Dim sql As String = ""
        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        sql = "" & vbCrLf
        sql += "  INSERT INTO Stud_TrainCostD(" & vbCrLf
        sql += "  SOCID ,IDNO ,OCID ,OtherJobCost ,ModifyAcct ,ModifyDate" & vbCrLf
        sql += " ) VALUES (" & vbCrLf
        sql += "  @SOCID ,@IDNO ,@OCID ,@OtherJobCost ,@ModifyAcct ,getdate()" & vbCrLf
        sql += " ) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)
        'da.SelectCommand.CommandText = Sql

        'Dim da2 As SqlDataAdapter = TIMS.GetOneDA(objconn)
        sql = "" & vbCrLf
        sql += " UPDATE Stud_TrainCostD " & vbCrLf
        sql += " SET OtherJobCost= @OtherJobCost " & vbCrLf
        sql += " ,IDNO= @IDNO" & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        sql += " AND SOCID= @SOCID" & vbCrLf
        sql += " AND OCID= @OCID" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)
        'da2.SelectCommand.CommandText = Sql

        '-------為避免使用者先存學員補助金額,之後再將學員作離退訓處理,所以將離退訓學員的補助金額update成null----
        If OCIDValue1.Value <> "" Then
            Dim sql1 As String = ""
            sql1 = "select * from Stud_TrainCostD where OCID = '" & OCIDValue1.Value & "'" '如果此學員沒有記錄就新增
            Dim dt As DataTable = DbAccess.GetDataTable(sql1, objconn)

            For Each Item As DataGridItem In DataGrid1.Items
                Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
                Dim OtherJobCost As TextBox = Item.FindControl("OtherJobCost")

                Dim IDNOv As String = Convert.ToString(Item.Cells(4).Text)    'IDNO
                Dim SOCIDv As String = Convert.ToString(Checkbox1.Value)   'socid

                If Checkbox1.Checked = True Then
                    '補助
                    If IDNOv <> "" AndAlso SOCIDv <> "" Then
                        '如果此學員沒有記錄就新增
                        If dt.Select("Socid = '" & SOCIDv & "'").Length = 0 Then
                            Try
                                Call TIMS.OpenDbConn(objconn)
                                With iCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCIDv
                                    .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNOv
                                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                                    .Parameters.Add("OtherJobCost", SqlDbType.Int).Value = CInt(OtherJobCost.Text)
                                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                    .ExecuteNonQuery()
                                End With
                            Catch ex As Exception
                                flagSaveOk = False '儲存異常
                                Common.MessageBox(Me, ex.ToString)
                                Exit For
                                'Throw ex
                            End Try
                        Else
                            Try
                                Call TIMS.OpenDbConn(objconn)
                                With uCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("OtherJobCost", SqlDbType.Int).Value = CInt(OtherJobCost.Text)
                                    .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNOv
                                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCIDv
                                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                                    '.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                    .ExecuteNonQuery()
                                End With
                            Catch ex As Exception
                                flagSaveOk = False '儲存異常
                                Common.MessageBox(Me, ex.ToString)
                                Exit For
                                'Throw ex
                            End Try
                        End If
                    Else
                        flagSaveOk = False '儲存異常
                        Common.MessageBox(Me, "學員選擇有誤，請重新選擇!!")
                        Exit For
                    End If
                Else
                    '不補助
                    If IDNOv <> "" AndAlso SOCIDv <> "" Then
                        '如果此學員沒有記錄就新增
                        '有資料
                        If dt.Select("Socid = '" & SOCIDv & "'").Length <> 0 Then
                            Try
                                '不補助金額清除
                                Call TIMS.OpenDbConn(objconn)
                                With uCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("OtherJobCost", SqlDbType.Int).Value = Convert.DBNull
                                    .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNOv
                                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCIDv
                                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                                    '.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                    .ExecuteNonQuery()
                                End With
                            Catch ex As Exception
                                flagSaveOk = False '儲存異常
                                Common.MessageBox(Me, ex.ToString)
                                Exit For
                                'Throw ex
                            End Try
                        End If
                    End If
                End If
            Next

            'If da.SelectCommand.Connection.State = ConnectionState.Open Then da.SelectCommand.Connection.Close()
            'If da2.SelectCommand.Connection.State = ConnectionState.Open Then da2.SelectCommand.Connection.Close()

            If flagSaveOk Then
                'Dim dr As DataRow = Nothing
                '清除 離退訓學員 的金額
                sql = "" & vbCrLf
                sql += " select st.socid " & vbCrLf
                sql += " from Class_StudentsOfClass cs " & vbCrLf
                sql += " join Stud_TrainCostD st on cs.socid = st.socid " & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and cs.studstatus in (2,3) " & vbCrLf
                sql += " and cs.OCID=@OCID" & vbCrLf
                Call TIMS.OpenDbConn(objconn)
                Dim dtX As New DataTable
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                    dtX.Load(.ExecuteReader())
                End With
                '有資料 執行 清除 離退訓學員 的金額
                If dtX.Rows.Count > 0 Then
                    sql = "" & vbCrLf
                    sql += " UPDATE Stud_TrainCostD" & vbCrLf
                    sql += " set OtherJobCost = NULL" & vbCrLf
                    sql += " where 1=1" & vbCrLf
                    sql += " AND SOCID IN  (" & vbCrLf
                    sql += "   select st.socid" & vbCrLf
                    sql += "   from Class_StudentsOfClass cs " & vbCrLf
                    sql += "   join Stud_TrainCostD st on cs.socid = st.socid " & vbCrLf
                    sql += "   WHERE 1=1" & vbCrLf
                    sql += "   and cs.studstatus in (2,3)" & vbCrLf
                    sql += "   and cs.OCID =@OCID" & vbCrLf
                    sql += " )" & vbCrLf
                    Call TIMS.OpenDbConn(objconn)
                    oCmd = New SqlCommand(sql, objconn)
                    With oCmd
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                        .ExecuteNonQuery()
                    End With
                End If
            End If
        End If
        '-----------------------------end-----------------------------------------------------------------

        Common.MessageBox(Me, "儲存成功")
        Query_Click(sender, e)

    End Sub

    '列印
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Me.ViewState("MyValue") = ""
        Me.ViewState("MyValue") += "&OCID=" & OCIDValue1.Value
        Me.ViewState("MyValue") += "&RID=" & RIDValue.Value

        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        '採自費/公費欄位判斷
        'and (cs.PMode = 1 or cs.PMode is NUll) '改為 自費生 永遠排除。20141002 BY AMU
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_09_016_R2", Me.ViewState("MyValue"))
        'If Not TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    'and (cs.PMode = 1 or cs.PMode is NUll) '改為 自費生 永遠排除。20141002 BY AMU
        '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_09_016_R2", Me.ViewState("MyValue"))
        'Else
        '    'and (cs.PMode = 1 or cs.PMode is NUll) '改為 自費生 永遠排除。20141002 BY AMU
        '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_09_016_R", Me.ViewState("MyValue"))
        'End If
    End Sub

End Class
