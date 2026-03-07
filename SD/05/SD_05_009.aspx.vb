Partial Class SD_05_009
    Inherits AuthBasePage

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            years = TIMS.GetSyear(years)
            months.Items.Add(New ListItem("==請選擇==", 0))
            For i As Integer = 1 To 12
                months.Items.Add(i)
            Next
        End If
        calculate_button.Attributes("onclick") = "javascript:return print();"
        count_Button.Attributes("onclick") = "javascript:return check();"
        CB1.Attributes("onclick") = "javascript:return check_choice(this.checked);"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.Utl_Redirect1(Me, "SD_05_009_List.aspx")
    End Sub

    Private Sub calculate_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles calculate_button.Click

        Dim sqlstr_TeacherList As String = "select a.TechID,a.TeachCName from "
        sqlstr_TeacherList += "(select Teacher1 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' union "
        sqlstr_TeacherList += "select Teacher2 from Class_Schedule where  OCID='" & OCIDValue1.Value & "'  and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' union "
        sqlstr_TeacherList += "select Teacher3 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher4 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher5 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher6 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher7 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher8 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher9 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher10 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher11 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher12 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher13 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher14 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher15 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher16 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher17 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher18 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher19 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher20 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher21 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher22 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher23 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "'  union "
        sqlstr_TeacherList += "select Teacher24 from Class_Schedule where  OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' ) as teacher  "
        sqlstr_TeacherList += " join Teach_TeacherInfo a on a.TechID=teacher.Teacher1"
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr_TeacherList, objConn)

        Me.Panel.Visible = False
        msg.Text = "查無資料!!"
        Me.Panel1.Visible = True

        If dt.Rows.Count > 0 Then
            Me.Panel.Visible = True
            msg.Text = ""
            Me.Panel1.Visible = False

            Me.CB_teacherList.DataSource = dt ' objreader
            Me.CB_teacherList.DataTextField = "TeachCName"
            Me.CB_teacherList.DataValueField = "TechID"
            Me.CB_teacherList.DataBind()
        End If

    End Sub

    Private Sub count_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles count_Button.Click
        'Dim sqlAdapter As SqlDataAdapter
        'Dim sqlTable As DataTable
        Dim sqldr As DataRow
        Dim TechIDstr As String
        Dim strMessage As String = ""
        Dim strMessage1 As String = ""
        For Each item As ListItem In CB_teacherList.Items
            If item.Selected Then
                TechIDstr = item.Text
                Dim total_hours As String = "select count(*) as count_hour,SchoolDate  from (select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and Teacher1='" & item.Value & "' "
                total_hours += "union  all select * from Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and Teacher2='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher3='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher4='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and Teacher5='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and   Teacher6='" & item.Value & "'  "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and   Teacher7='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher8='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher9='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher10='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher11='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher12='" & item.Value & "'"
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher13='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher14='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher15='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher16='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher17='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher18='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher19='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher20='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher21='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher22='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher23='" & item.Value & "' "
                total_hours += "union all select * from  Class_Schedule where OCID='" & OCIDValue1.Value & "' and DATEPART(YEAR, SchoolDate)='" & years.SelectedValue & "' and DATEPART(MONTH, SchoolDate)='" & months.SelectedValue & "' and  Teacher24='" & item.Value & "' ) as count_class group by SchoolDate "
                Dim hours As DataTable = DbAccess.GetDataTable(total_hours, objConn)
                Dim Price_sql As String = "select b.OverCharge from Teach_TeacherInfo a join ID_KindOfTeacher b on a.KindID=b.KindID where a.TechID='" & item.Value & "'"
                Dim Price As String = Convert.ToString(DbAccess.ExecuteScalar(Price_sql, objConn))

                Dim hourRow As DataRow = Nothing
                Dim daPayHour As SqlDataAdapter = Nothing
                Dim dtPayHour As DataTable = DbAccess.GetDataTable("select * from Teach_PayHour where 1<>1", daPayHour, objConn)
                For Each hourRow In hours.Rows
                    Dim strsql_check As String = "select * from Teach_PayHour where OCID='" & OCIDValue1.Value & "' and TechID='" & item.Value & "' and Month(TeachDate)=Month('" & hourRow("SchoolDate") & "')"
                    '判斷是否新增重複的資料
                    If DbAccess.GetCount(strsql_check) > 0 Then '>0有重複的資料
                        If InStr(strMessage & ",", "," & item.Text & ",") = 0 Then
                            strMessage &= "," & item.Text
                        End If
                    Else '沒有重複
                        sqldr = dtPayHour.NewRow
                        sqldr("ModifyAcct") = sm.UserInfo.UserID
                        sqldr("ModifyDate") = Now()
                        sqldr("OCID") = OCIDValue1.Value
                        sqldr("TechID") = item.Value
                        sqldr("TeachDate") = hourRow("SchoolDate")
                        sqldr("UnitPrice") = Price
                        sqldr("UnitHour") = hourRow("count_hour")
                        dtPayHour.Rows.Add(sqldr)
                        strMessage1 &= "," & item.Text
                    End If
                Next
                DbAccess.UpdateDataTable(dtPayHour, daPayHour)

                'count_Button.Enabled = False
                If Request("save_type") = "replace" Then '若點選覆蓋
                    Dim sql_del As String = "delete Teach_PayHour where OCID='" & OCIDValue1.Value & "' and TechID='" & item.Value & "' and Month(TeachDate)=Month('" & hourRow("SchoolDate") & "')"
                    DbAccess.ExecuteNonQuery(sql_del, objConn)

                    sqldr = dtPayHour.NewRow
                    sqldr("ModifyAcct") = sm.UserInfo.UserID
                    sqldr("ModifyDate") = Now()
                    sqldr("OCID") = OCIDValue1.Value
                    sqldr("TechID") = item.Value
                    sqldr("TeachDate") = hourRow("SchoolDate")
                    sqldr("UnitPrice") = Price
                    sqldr("UnitHour") = hourRow("count_hour")
                    dtPayHour.Rows.Add(sqldr)
                    If InStr(strMessage1 & ",", "," & item.Text & ",") = 0 Then
                        strMessage1 &= "," & item.Text
                    End If
                    DbAccess.UpdateDataTable(dtPayHour, daPayHour)
                ElseIf Request("save_type") = "stop" Then '點選取消
                    Exit Sub
                End If
            End If
        Next
        If strMessage <> "" AndAlso Request("save_type") Is Nothing Then '確認的錯誤訊息
            Dim totalstring As String = strMessage.Substring(1)
            Page.RegisterHiddenField("save_type", "")
            Common.AddClientScript(Page, "if (window.confirm('" & totalstring & "鍾點費試算重複,是否覆蓋?')) {")
            Common.AddClientScript(Page, "  form1.save_type.value='replace';")
            Common.AddClientScript(Page, "  form1.count_Button.click();")
            Common.AddClientScript(Page, " } else { ")
            Common.AddClientScript(Page, "  form1.save_type.value='stop';")
            Common.AddClientScript(Page, "  form1.count_Button.click();")
            Common.AddClientScript(Page, "}")
            Exit Sub
        End If
        If strMessage1 <> "" Then
            Common.MessageBox(Page, strMessage1.Substring(1) & "講師鐘點費試算成功!!" & vbCrLf)
        End If
    End Sub
End Class
