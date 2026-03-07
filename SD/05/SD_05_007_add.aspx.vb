Partial Class SD_05_007_add
    Inherits AuthBasePage

    Dim Days1 As Integer
    Dim Days2 As Integer

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

        Button4.Attributes("onclick") = "location.href='SD_05_007.aspx?ID=" & Request("ID") & "';"
        msg.Text = ""
        '取出設定天數檔------------------------------Start
        Dim sql As String
        Dim dr As DataRow
        sql = "SELECT * FROM Sys_Days   "
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            Days1 = dr("Days1")
            Days2 = dr("Days2")
        End If
        '取出設定天數檔------------------------------End

        If Request("Proecess") = "add" Then
            If Not IsPostBack Then
                Button1.Visible = False
                center.Text = sm.UserInfo.OrgName
                RIDValue.Value = sm.UserInfo.RID
                SanDate.Text = Now.Date
            Else
                GetStudent()
            End If
            If Hidden1.Value = "1" Then
                'Hidden1.Value = 0
            End If
        ElseIf Request("Proecess") = "edit" Or Request("Proecess") = "view" Then
            center.Enabled = False
            Button3.Disabled = True
            TMID1.Enabled = False
            OCID1.Enabled = False
            Button2.Disabled = True
            SanDate.Enabled = False
            IMG1.Style.Item("display") = "none"

            GetStudent()
            If Request("Proecess") = "view" Then
                Button1.Visible = False
            End If
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , , True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "javascript:return chkdata();"
        Button2.Attributes("onclick") = "javascript:choose_class();"
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button3.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button3.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If
    End Sub

    Sub GetStudent()
        '取出所有學生資料
        Dim sql As String
        Dim i As Integer
        Dim j As Integer
        Dim dt As DataTable


        If Request("Proecess") = "add" Then
            sql = "SELECT   b.SOCID, b.StudentID, a.Name,b.StudStatus,c.FTDate "
            sql += "FROM     Stud_StudentInfo a, "
            sql += "         (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') b, "
            sql += "         (SELECT * FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "') c "
            sql += "WHERE    a.SID=b.SID and b.OCID=c.OCID "
            sql += "Order By b.StudentID"
        Else
            sql = "SELECT   b.SOCID, b.StudentID, a.Name,b.StudStatus,c.FTDate "
            sql += "FROM     Stud_StudentInfo a, "
            sql += "         (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "' and SOCID='" & Request("SOCID") & "') b, "
            sql += "         (SELECT * FROM Class_ClassInfo WHERE OCID='" & Request("OCID") & "') c "
            sql += "WHERE    a.SID=b.SID and b.OCID=c.OCID "
            sql += "Order By b.StudentID"
        End If
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            msg.Text = "查無此班學生資料!!"
            Table4.Visible = False
            Button1.Visible = False
        Else
            Table4.Visible = True
            Table4.BorderWidth = Unit.Pixel(1)
            Button1.Visible = True
            '取出獎懲鍵詞
            sql = "SELECT * FROM Key_Sanction   "
            Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)

            Dim cell As TableCell
            Dim row As TableRow

            '先建立表頭
            row = New TableRow

            cell = New TableCell
            cell.BorderWidth = Unit.Pixel(1)
            cell.Text = "學號"
            row.Cells.Add(cell)

            cell = New TableCell
            cell.BorderWidth = Unit.Pixel(1)
            cell.Text = "姓名"
            row.Cells.Add(cell)

            For i = 0 To dt1.Rows.Count - 1
                cell = New TableCell
                cell.BorderWidth = Unit.Pixel(1)
                cell.Text = dt1.Rows(i).Item("Name")
                row.Cells.Add(cell)
            Next
            row.BackColor = Color.FromName("#2aafc0")
            row.ForeColor = Color.White
            row.HorizontalAlign = HorizontalAlign.Center
            Table4.Rows.Add(row)

            '建立資料行
            Dim dr As DataRow
            dt.DefaultView.Sort = "StudentID"
            For i = 0 To dt.Rows.Count - 1
                row = New TableRow

                cell = New TableCell
                cell.BorderWidth = Unit.Pixel(1)
                Dim SID As New Label
                SID.Text = Right(dt.Rows(i).Item("StudentID"), 2)
                Dim SOCID As New HtmlInputHidden
                SOCID.Value = dt.Rows(i).Item("SOCID")
                SOCID.ID = "SOCID_" & i
                cell.Controls.Add(SID)
                cell.Controls.Add(SOCID)
                row.Cells.Add(cell)

                cell = New TableCell
                cell.BorderWidth = Unit.Pixel(1)
                cell.Text = dt.Rows(i).Item("Name")
                Select Case dt.Rows(i).Item("StudStatus").ToString
                    Case "1"
                        cell.Text += "(在訓)"
                    Case "2"
                        cell.Text += "(離訓)"
                    Case "3"
                        cell.Text += "(退訓)"
                    Case "4"
                        cell.Text += "(續訓)"
                    Case "5"
                        cell.Text += "(結訓)"
                End Select
                row.Cells.Add(cell)

                For j = 0 To dt1.Rows.Count - 1
                    cell = New TableCell
                    cell.BorderWidth = Unit.Pixel(1)
                    Dim text As New TextBox
                    text.ID = dt1.Rows(j).Item("SanID") & "_" & i
                    text.Width = Unit.Pixel(45)
                    cell.Controls.Add(text)
                    row.Cells.Add(cell)
                Next
                '設定樣式
                If i = 0 Then
                    row.BackColor = Color.FromName("#ecf7ff")
                ElseIf i Mod 2 = 0 Then
                    row.BackColor = Color.FromName("#ecf7ff")
                ElseIf i Mod 2 = 1 Then
                    row.BackColor = Color.White
                End If
                row.HorizontalAlign = HorizontalAlign.Center

                If Request("Proecess") <> "view" Then
                    Select Case dt.Rows(i).Item("StudStatus").ToString
                        Case "5"
                            Select Case sm.UserInfo.RoleID
                                Case 0, 1
                                    '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天
                                    '暫時先改這樣,以後還會再改
                                    If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                        If DateDiff(DateInterval.Day, dt.Rows(i)("FTDate"), Now.Date) <= 75 Then
                                            Table4.Rows.Add(row)
                                        End If
                                    Else
                                        If DateDiff(DateInterval.Day, dt.Rows(i)("FTDate"), Now.Date) <= Days2 Then
                                            Table4.Rows.Add(row)
                                        End If
                                    End If

                                Case Else
                                    '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                                    '暫時先改這樣,以後還會再改
                                    If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                        If DateDiff(DateInterval.Day, dt.Rows(i)("FTDate"), Now.Date) <= 60 Then
                                            Table4.Rows.Add(row)
                                        End If
                                    Else
                                        If DateDiff(DateInterval.Day, dt.Rows(i)("FTDate"), Now.Date) <= Days1 Then
                                            Table4.Rows.Add(row)
                                        End If
                                    End If
                            End Select
                        Case Else
                            Table4.Rows.Add(row)
                    End Select
                Else
                    Table4.Rows.Add(row)
                End If
            Next
            If Table4.Rows.Count = 1 Then
                Table4.Visible = False
                Button1.Visible = False
                msg.Text = "查無此班在訓或續訓的學生!!"
            End If


            '檢視跟編輯狀態要另外填值進去
            If Request("Proecess") = "edit" Or Request("Proecess") = "view" Then
                sql = "SELECT SanID,sum(Times) as num FROM Stud_Sanction   WHERE SOCID='" & Request("SOCID") & "' and SanDate=convert(datetime, '" & Request("SanDate") & "', 111) GROUP BY SanID"
                dt = DbAccess.GetDataTable(sql, objconn)
                For Each dr In dt.Rows
                    Dim mytext As TextBox = Table4.Rows(0).FindControl(dr("SanID") & "_0")
                    If Not mytext Is Nothing Then
                        mytext.Text = dr("num")
                    Else
                        Common.MessageBox(Me, "查無TextBox" & dr("SanID") & "_0")
                    End If
                Next

                '找出所屬的班級職類
                dr = TIMS.GetOCIDDate(Request("OCID"))

                If Not dr Is Nothing Then
                    TMID1.Text = "[" & dr("TrainID") & "]" & dr("TrainName")
                    TMIDValue1.Value = dr("TMID")

                    OCID1.Text = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))
                    OCIDValue1.Value = Request("OCID")

                    '找出機構
                    center.Text = dr("OrgName")
                    RIDValue.Value = dr("RID")
                End If
                SanDate.Text = Request("SanDate")
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim drResult As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim mesbox As String = ""

        '取出獎懲鍵詞
        sql = "SELECT * FROM Key_Sanction   "
        Dim dtKey As DataTable = DbAccess.GetDataTable(sql, objconn)

        If Request("Proecess") = "edit" Or Request("Proecess") = "add" Then
            If Request("Proecess") = "add" Then
                mesbox = "新增成功!"
            Else
                mesbox = "修改成功!"
                '修改狀態時要先刪除所有的同一天資料
                sql = "DELETE Stud_Sanction WHERE SOCID='" & Request("SOCID") & "' and SanDate=convert(datetime, '" & Request("SanDate") & "', 111)"
                Try
                    DbAccess.ExecuteNonQuery(sql, objconn)
                Catch ex As Exception
                    'Common.RespWrite(Me, ex)
                    Throw ex
                End Try
            End If

            '先取出空白資料表
            sql = "SELECT * FROM Stud_Sanction WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, objconn)

            For i As Integer = 1 To Table4.Rows.Count - 1
                '先抓出資料表內是否有重複的資料
                Dim SeqNo As Integer = 0
                Dim mytext As TextBox
                Dim myhid As HtmlInputHidden = Table4.Rows(i).FindControl("SOCID_" & (i - 1))
                sql = "SELECT MAX(SeqNo) as num FROM Stud_Sanction WHERE SOCID='" & myhid.Value & "' and SanDate=convert(datetime, '" & SanDate.Text & "', 111)"
                drResult = DbAccess.GetOneRow(sql)
                If IsDBNull(drResult("num")) Then
                    SeqNo = 1
                Else
                    SeqNo = drResult("num") + 1
                End If

                For j As Integer = 0 To dtKey.Rows.Count - 1
                    mytext = Table4.Rows(i).FindControl(dtKey.Rows(j).Item("SanID") & "_" & (i - 1))
                    If mytext.Text <> "" Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)

                        dr("SOCID") = myhid.Value
                        dr("SanDate") = SanDate.Text
                        dr("SeqNo") = SeqNo
                        dr("SanID") = dtKey.Rows(j).Item("SanID")
                        dr("Times") = mytext.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    End If
                    SeqNo += 1
                Next
            Next
        End If


        Try
            DbAccess.UpdateDataTable(dt, da)
            Common.RespWrite(Me, "<script language='javascript'>alert('" & mesbox & "');</script>")
            Common.RespWrite(Me, "<script language='javascript'>location.href='SD_05_007.aspx?ID=" & Request("ID") & "';</script>")
        Catch ex As Exception
            'Common.RespWrite(Me, ex)
            Throw ex
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'Dim dr As DataRow
        ''判斷機構是否只有一個班級
        'dr = TIMS.GET_ONLYONE_OCID(RIDValue.Value)
        'If Not dr Is Nothing Then
        '    If dr("total") = "1" Then '如果只有一個班級
        '        TMID1.Text = dr("trainname")
        '        OCID1.Text = dr("classname")
        '        TMIDValue1.Value = dr("trainid")
        '        OCIDValue1.Value = dr("ocid")
        '    Else '不只一個班級
        '        TMID1.Text = ""
        '        OCID1.Text = ""
        '        TMIDValue1.Value = ""
        '        OCIDValue1.Value = ""
        '    End If
        'End If
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        Button1.Visible = False
        msg.Visible = False
    End Sub


End Class
