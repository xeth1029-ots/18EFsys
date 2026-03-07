Partial Class SV_01_004_Insert
    Inherits System.Web.UI.Page


    Dim objconn As OracleConnection

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


        If Request("inline") <> 1 Then '判斷是否是線上填寫 1是線上填寫
            '檢查Session是否存在--------------------------Start
            ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
            '檢查Session是否存在--------------------------End
        End If

        Dim SQL As String
        Dim SQL2 As String
        Dim SQL3 As String
        Dim SQL4 As String
        Dim SQL5 As String
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim dt3 As DataTable
        Dim dt4 As DataTable
        Dim dt5 As DataTable
        Dim dt6 As DataTable
        'Dim i As Integer ' 標題的筆數
        Dim j As Integer ' 題目的筆數
        Dim z As Integer ' 答案的筆數
        Dim y As Integer ' 填答的筆數
        Dim TBT As HtmlTable
        'Dim TB2 As HtmlTable
        Dim Row As HtmlTableRow
        Dim Cell As HtmlTableCell
        Dim SKID As String
        Dim SQID As String
        Dim RadioButtonList As RadioButtonList
        Dim CheckBoxList As CheckBoxList

        If Not IsPostBack Then
            SVID.Value = Request("SVID")
            SOCID.Value = Request("SOCID")
            Type.Value = Request("Type")
            Me.ViewState("sysAnswer") = Nothing
        End If

        
        SQL = "Select * from KEY_SURVEYKIND where SVID = " & SVID.Value & " order by serial " '取出標題
        dt = DbAccess.GetDataTable(sql, objconn)
        For i As Integer = 0 To dt.Rows.Count - 1 '利用迴圈新增標題

            SKID = dt.Rows(i).Item(0).ToString '　取出SKID
            TBT = New HtmlTable
            PlaceHolder1.Controls.Add(TBT) '新增TABLE
            TBT.ID = "TBT" & i
            TBT.Attributes.Add("width", "100%")

            Row = New HtmlTableRow         '新增TR
            TBT.Controls.Add(Row)

            Cell = New HtmlTableCell       '新增TD
            Cell.InnerText = dt.Rows(i).Item(1).ToString   '代入標題
            Cell.Attributes.Add("Style", "background-color: #2aafc0")   '設定顏色
            Cell.Attributes.Add("width", "100%")
            Row.Controls.Add(Cell)

            SQL2 = "Select * from ID_SurveyQuestion Where SKID = " & SKID & " order by Serial" '取出問題的題目
            dt2 = DbAccess.GetDataTable(SQL2)

            For j = 0 To dt2.Rows.Count - 1 '利用迴圈新增題目

                SQID = dt2.Rows(j).Item(0).ToString '取出SQID

                TBT = New HtmlTable
                PlaceHolder1.Controls.Add(TBT) '新增TABLE
                TBT.Attributes.Add("width", "100%")
                TBT.Attributes.Add("Style", "background-color: #e9f2fc")

                Row = New HtmlTableRow         '新增TR
                TBT.Controls.Add(Row)

                Cell = New HtmlTableCell                        '新增td
                Cell.Attributes.Add("width", "100%")
                Cell.InnerText = dt2.Rows(j).Item(1).ToString '題目內容
                Row.Controls.Add(Cell)

                SQL3 = "Select * from ID_SurveyAnswer where SQID = " & SQID & " order by Serial" '取出答案內容
                dt3 = DbAccess.GetDataTable(SQL3)

                If Type.Value = "E" Then   '如果是修改
                    SQL4 = "Select SAID from Stud_Survey where SQID = " & SQID & " and SOCID = " & SOCID.Value & " "
                    dt4 = DbAccess.GetDataTable(SQL4)
                End If

                If dt3.Rows.Count <> 0 Then     '如果答案選項不是零

                    Row = New HtmlTableRow               '新增TR
                    TBT.Controls.Add(Row)

                    Cell = New HtmlTableCell
                    Row.Controls.Add(Cell)               '新增TD

                    If dt2.Rows(j).Item(2).ToString = 1 Then    '判斷是那程型態的題目 1是radio2,2是checkbox

                        RadioButtonList = New RadioButtonList   '新增 RadioButtonList
                        RadioButtonList.ID = dt2.Rows(j).Item(0).ToString   'id = SQID 的值
                        RadioButtonList.Attributes.Add("runat", "server")
                        Cell.Controls.Add(RadioButtonList)

                        For z = 0 To dt3.Rows.Count - 1  '計算出有幾個答案選項

                            RadioButtonList.Items.Add(dt3.Rows(z).Item(1).ToString)    '新增RadioButton的TEXT 為答案的內容ANSWER
                            RadioButtonList.Items.Item(z).Value = dt3.Rows(z).Item(0).ToString 'RadioButton的值為SAID

                            If Type.Value = "E" And Me.ViewState("sysAnswer") Is Nothing Then  '如果是修改還有答案都有作答

                                For y = 0 To dt4.Rows.Count - 1    '計算出有幾個學員作答結果

                                    If RadioButtonList.Items.Item(z).Value = dt4.Rows(y).Item(0).ToString Then

                                        RadioButtonList.Items.Item(z).Selected = True

                                    End If

                                Next

                            End If

                            If Not Me.ViewState("sysAnswer") Is Nothing Then  '如果有答案沒有作答的
                                dt6 = Me.ViewState("sysAnswer")
                                dt6.Select("SQID = SQID")
                                For y = 0 To dt6.Rows.Count - 1
                                    If RadioButtonList.Items.Item(z).Value = dt6.Rows(y).Item(1).ToString Then

                                        RadioButtonList.Items.Item(z).Selected = True

                                    End If

                                Next

                            End If


                        Next

                    Else                                          '2是checkbox

                        CheckBoxList = New CheckBoxList           '新增 CheckBoxList
                        CheckBoxList.ID = dt2.Rows(j).Item(0).ToString   'id = SQID 的值
                        CheckBoxList.Attributes.Add("runat", "server")
                        Cell.Controls.Add(CheckBoxList)

                        For z = 0 To dt3.Rows.Count - 1 '計算出有幾個答案選項

                            CheckBoxList.Items.Add(dt3.Rows(z).Item(1).ToString)    '新增 CheckBoxList 的TEXT內容
                            CheckBoxList.Items.Item(z).Value = dt3.Rows(z).Item(0).ToString 'CheckBoxList 的值 = SAID

                            If Type.Value = "E" And Me.ViewState("sysAnswer") Is Nothing Then    '如果是修改及沒有答案是沒有填的
                                For y = 0 To dt4.Rows.Count - 1 '計算出有幾個學員作答結果
                                    If CheckBoxList.Items.Item(z).Value = dt4.Rows(y).Item(0).ToString Then

                                        CheckBoxList.Items.Item(z).Selected = True

                                    End If
                                Next

                            End If

                            If Not Me.ViewState("sysAnswer") Is Nothing Then '如果有問題答案沒有填
                                dt6 = Me.ViewState("sysAnswer")
                                dt6.Select("SQID = SQID")
                                For y = 0 To dt6.Rows.Count - 1
                                    If CheckBoxList.Items.Item(z).Value = dt6.Rows(y).Item(1).ToString Then

                                        CheckBoxList.Items.Item(z).Selected = True

                                    End If

                                Next

                            End If

                        Next

                    End If

                End If

            Next

        Next
        SQL5 = " select right(StudentID,2) as StudentID,ss.Name as Sname"
        SQL5 += " from Class_StudentsOfClass cs join stud_studentinfo ss on cs.sid = ss.sid "
        SQL5 += " where cs.Socid = " & SOCID.Value & " "
        dt5 = DbAccess.GetDataTable(SQL5)
        StudentIDL.Text = dt5.Rows(0).Item(0)   '學號
        SnameL.Text = dt5.Rows(0).Item(1).ToString '學員姓名
        Title.Visible = True
        PlaceHolder1.Visible = True
        Save.Visible = True   '存檔
        returnQ.Visible = True '回上一頁,關閉視窗

        If Request("inline") <> 1 Then '判斷是否是線上填寫,1是線上填寫

            returnQ.Text = "回上一頁"
        Else
            returnQ.Text = "關閉視窗" '線上填寫
        End If


        If Type.Value = "I" Or Type.Value = "R" Then   '若新增則重設鍵顯示

            reset.Visible = True '重填
        Else
            reset.Visible = False '否則就隱藏
        End If


    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click

        Dim SQL4 As String
        Dim SQL As String
        Dim SQL2 As String
        Dim SQL3 As String
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim dt3 As DataTable
        Dim dt4 As DataTable
        Dim a As Integer '標題的計數
        Dim b As Integer '題目的計數
        Dim c As Integer '答案的計數
        Dim SKID2 As String
        Dim SQID2 As String
        Dim SAID As String
        Dim RBL As RadioButtonList
        Dim CHL As CheckBoxList

        If checkdata(sender, e) = True Then    '如果問題的答案都有填寫

            SQL = "Select * from KEY_SURVEYKIND where SVID = " & Request("SVID") & " order by serial " '取出標題
            dt = DbAccess.GetDataTable(SQL)
            For a = 0 To dt.Rows.Count - 1

                SKID2 = dt.Rows(a).Item(0).ToString

                SQL2 = "Select * from ID_SurveyQuestion Where SKID = " & SKID2 & " order by Serial" '取出問題的題目
                dt2 = DbAccess.GetDataTable(SQL2)

                For b = 0 To dt2.Rows.Count - 1

                    SQID2 = dt2.Rows(b).Item(0).ToString

                    SQL3 = "Select * from ID_SurveyAnswer where SQID = " & SQID2 & " order by Serial" '取出答案內容
                    dt3 = DbAccess.GetDataTable(SQL3)

                    If dt2.Rows(b).Item(2).ToString = 1 Then    '如果是RadioButtonList

                        RBL = DirectCast(Me.Panel1.FindControl(dt2.Rows(b).Item(0).ToString), RadioButtonList) '取得 RadioButtonList

                        For c = 0 To dt3.Rows.Count - 1

                            If RBL.Items.Item(c).Selected Then

                                SAID = RBL.Items.Item(c).Value

                                If Type.Value = "I" Or Type.Value = "R" Then '如果是新增或是重填

                                    SQL = "Insert Into Stud_Survey(SOCID,SVID,SKID,SQID,SAID,ModifyAcct,ModifyDate)"
                                    SQL += "values('" & SOCID.Value & "', '" & SVID.Value & "'," & SKID2 & "," & SQID2 & "," & SAID & ",'" & sm.UserInfo.UserID & "', getdate())"
                                    DbAccess.ExecuteNonQuery(SQL)

                                ElseIf Type.Value = "E" Then  '如果是修改

                                    SQL4 = "Select * from Stud_Survey where SQID = " & SQID2 & " and SOCID = '" & SOCID.Value & "' "  '找出己新增的那筆資料
                                    dt4 = DbAccess.GetDataTable(SQL4)
                                    If dt4.Rows.Count <> 0 Then       'update 資料

                                        SQL = "Update Stud_Survey "
                                        SQL += "Set SOCID = '" & SOCID.Value & "',SVID ='" & SVID.Value & "',SKID = " & SKID2 & ",SQID = " & SQID2 & ",SAID = " & SAID & ",ModifyAcct ='" & sm.UserInfo.UserID & "',ModifyDate = getdate() "
                                        SQL += "where SQID = '" & SQID2 & "' and SOCID = '" & SOCID.Value & "' "
                                        DbAccess.ExecuteNonQuery(SQL)

                                    End If

                                End If
                            End If

                        Next

                    Else              '如果是checkboxlist

                        CHL = DirectCast(Me.Panel1.FindControl(dt2.Rows(b).Item(0).ToString), CheckBoxList) '取得checkboxlist

                        If Type.Value = "E" Then  '如果是修改

                            SQL4 = "Delete Stud_Survey where SQID = " & SQID2 & " and SOCID = '" & SOCID.Value & "'" '先刪除那一筆問題的學員作答,再重新新增
                            DbAccess.ExecuteNonQuery(SQL4)

                        End If

                        For c = 0 To dt3.Rows.Count - 1

                            If CHL.Items.Item(c).Selected Then

                                SAID = CHL.Items.Item(c).Value
                                SQL = "Insert Into Stud_Survey(SOCID,SVID,SKID,SQID,SAID,ModifyAcct,ModifyDate)"
                                SQL += "values('" & SOCID.Value & "', '" & SVID.Value & "'," & SKID2 & "," & SQID2 & "," & SAID & ",'" & sm.UserInfo.UserID & "', getdate())"
                                DbAccess.ExecuteNonQuery(SQL)

                            End If

                        Next

                    End If

                Next
            Next

            If Request("inline") <> 1 Then '判斷是否是線上填寫,1是線上填寫
                Common.RespWrite(Me, "<script>alert('儲存成功!!');location.href='SV_08_004.aspx?ID=" & Request("ID") & "&OCID=" & Request("OCID") & "&SVID=" & Request("SVID") & "&IptName=" & Request("IptName") & "&RID=" & Request("RID") & "&OCIDValue1=" & Request("OCIDValue1") & "&PG=" & Request("PG") & "'</script>")
            Else

                If Request("BtnName") <> "" Then  '
                    Common.RespWrite(Me, "<script language=javascript>")
                    Common.RespWrite(Me, "function GetValue(){")
                    Common.RespWrite(Me, "window.opener.document.form1." & Request("BtnName") & ".click();")
                    Common.RespWrite(Me, "}")
                    Common.RespWrite(Me, "GetValue();")
                    Common.RespWrite(Me, "</script>")
                End If

                Common.RespWrite(Me, "<script>alert('儲存成功!!');window.close();</script>")
            End If


        End If

    End Sub

    Function checkdata(ByVal sender As System.Object, ByVal e As System.EventArgs) '檢查問題是否都有填答案

        Dim flag As Boolean = False '預設有答案沒有填
        Dim SQL4 As String
        Dim SQL As String
        Dim SQL2 As String
        Dim SQL3 As String
        Dim dt As DataTable
        Dim dt2 As DataTable
        Dim dt3 As DataTable
        Dim dt4 As DataTable
        Dim dt5 As DataTable
        Dim a As Integer '標題的計數
        Dim b As Integer '問題的計數
        Dim c As Integer '答案的計數
        Dim SKID2 As String
        Dim SQID2 As String
        Dim SAID As String
        Dim RBL As RadioButtonList
        Dim CHL As CheckBoxList
        Dim cknum As Integer
        Dim msg As String
        Dim dr As DataRow

        dt5 = New DataTable    '新增暫存的TABLE 放己選取的答案
        dt5.Columns.Add("SQID")
        dt5.Columns.Add("SAID")

        SQL = "Select * from KEY_SURVEYKIND where SVID = " & Request("SVID") & " order by serial " '取出標題
        dt = DbAccess.GetDataTable(SQL)
        For a = 0 To dt.Rows.Count - 1

            SKID2 = dt.Rows(a).Item(0).ToString

            SQL2 = "Select * from ID_SurveyQuestion Where SKID = " & SKID2 & " order by Serial" '取出問題的題目
            dt2 = DbAccess.GetDataTable(SQL2)

            For b = 0 To dt2.Rows.Count - 1

                SQID2 = dt2.Rows(b).Item(0).ToString

                SQL3 = "Select * from ID_SurveyAnswer where SQID = " & SQID2 & " order by Serial" '取出答案內容
                dt3 = DbAccess.GetDataTable(SQL3)

                If dt2.Rows(b).Item(2).ToString = 1 Then    '如果是RadioButtonList

                    RBL = DirectCast(Me.Panel1.FindControl(dt2.Rows(b).Item(0).ToString), RadioButtonList) '取得 RadioButtonList
                    cknum = dt3.Rows.Count
                    For c = 0 To dt3.Rows.Count - 1

                        If RBL.Items.Item(c).Selected Then
                            cknum = cknum - 1  '檢查答案是否有選取
                            SAID = RBL.Items.Item(c).Value
                            dr = dt5.NewRow        '若有選答案就將SAID跟SAID存到dt5
                            dt5.Rows.Add(dr)
                            dr("SQID") = SQID2
                            dr("SAID") = SAID
                        End If
                    Next
                    If cknum = dt3.Rows.Count Then  '表示沒有填答案
                        msg += "請輸入'" & dt2.Rows(b).Item(1).ToString & "'的答案!!" & vbCrLf
                    End If
                Else              '如果是checkboxlist

                    CHL = DirectCast(Me.Panel1.FindControl(dt2.Rows(b).Item(0).ToString), CheckBoxList) '取得checkboxlist
                    cknum = dt3.Rows.Count

                    For c = 0 To dt3.Rows.Count - 1

                        If CHL.Items.Item(c).Selected Then
                            cknum = cknum - 1    '檢查答案是否有選取
                            SAID = CHL.Items.Item(c).Value

                            dr = dt5.NewRow
                            dt5.Rows.Add(dr)
                            dr("SQID") = SQID2
                            dr("SAID") = SAID

                        End If

                    Next

                    If cknum = dt3.Rows.Count Then  ''表示沒有填答案
                        msg += "請輸入'" & dt2.Rows(b).Item(1).ToString & "'的答案!!" & vbCrLf

                    End If

                End If

            Next
        Next

        If msg <> "" Then '如果答案有沒填的

            If dt5.Rows.Count = 0 Then  '假如都沒有選答案,直接按存檔
                Common.MessageBox(Me, msg)
                PlaceHolder1.Controls.Clear()
                Page_Load(sender, e)
            Else                         '假如有選答案

                Me.ViewState("sysAnswer") = dt5
                Common.MessageBox(Me, msg)
                PlaceHolder1.Controls.Clear()
                Page_Load(sender, e)
                Me.ViewState("sysAnswer") = Nothing
            End If

        Else   '如果答案都有填
            flag = True
            Return flag
        End If

    End Function

    Private Sub reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Reset.Click
        PlaceHolder1.Controls.Clear() '清除表格
        Type.Value = "R" '重填
        Page_Load(sender, e)
    End Sub

    Private Sub returnQ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles returnQ.Click
        '回上一頁

        If Request("inline") <> 1 Then '判斷是否是線上填寫,1是線上填寫
            'Response.Redirect("SV_08_004.aspx?ID=" & Request("ID") & "&OCID=" & Request("OCID") & "&SVID=" & Request("SVID") & "&IptName=" & Request("IptName") & "&RID=" & Request("RID") & "&OCIDValue1=" & Request("OCIDValue1") & "&PG=" & Request("PG") & "")
        Else
            'Common.RespWrite(Me, "<script>window.close();</script>")
        End If

    End Sub

End Class
