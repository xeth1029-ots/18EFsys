Partial Class SD_04_009_R
    Inherits AuthBasePage

    'SD_04_009_R
    'SD_04_009_R_2
    Const Cst_TechPrintCount As Integer = 100 '能列印老師選擇數量

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

        Button2.Attributes("onclick") = "javascript:return ReportPrint();"

        If Not IsPostBack Then

            Years = TIMS.GetSyear(Years)
            Common.SetListItem(Years, Now.Year)
            For i As Integer = 1 To 12
                Months.Items.Add(New ListItem(i & "月份", i))
            Next
            Months.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Common.SetListItem(Months, Now.Month)

            Allday_TR.Style("display") = "none" '全期的隱藏

            Call CreateItem(sm.UserInfo.RID)

            '能列印老師選擇數量
            Me.Label1.Text = "老師數量選擇，請不要超過" & Cst_TechPrintCount & "筆，避免資料遺失(有資安問題)"
            Me.Label2.Text = "自動由勾選處再勾取 " & Cst_TechPrintCount & "筆資料!"
            Me.Label3.Text = "自動由勾選處再勾取 " & Cst_TechPrintCount & "筆資料!"
        End If
        '
        InSelectAll.Attributes("onclick") = "GetAllTeach(1,this.checked);"
        OutSelectAll.Attributes("onclick") = "GetAllTeach(2,this.checked);"

    End Sub

    '建立物件
    Sub CreateItem(ByVal RID As String)
        Dim sql As String
        Dim dt As DataTable
        Dim dv As New DataView

        sql = "SELECT * FROM Teach_TeacherInfo WHERE RID='" & RID & "'"
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.TableName = "Teach_TeacherInfo"
        dv.Table = dt
        Dim SORT_1 As String = "TeachCName,TechID"

        dv.RowFilter = "KindEngage=1"
        dv.Sort = SORT_1
        With InTeach
            .DataSource = dv
            .DataTextField = "TeachCName"
            .DataValueField = "TechID"
            .DataBind()
        End With

        dv.RowFilter = "KindEngage=2"
        dv.Sort = SORT_1
        With OutTeach
            .DataSource = dv
            .DataTextField = "TeachCName"
            .DataValueField = "TechID"
            .DataBind()
        End With
    End Sub

    Function Get_PrintCount(ByVal iRow As Integer) As Boolean
        Dim Rst As Boolean = True
        If iRow > Cst_TechPrintCount Then
            Rst = False
        End If
        Return Rst
    End Function

    '查詢檢核
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Dim FlagCanPrint As Boolean = True '能列印
        Dim iRow As Integer = 0
        Dim TeacherCon As String = ""
        For i As Integer = 0 To InTeach.Items.Count - 1
            If InTeach.Items(i).Selected = True And InTeach.Items(i).Value <> "" Then
                iRow += 1
                If TeacherCon <> "" Then TeacherCon += ","
                TeacherCon += InTeach.Items(i).Value

                FlagCanPrint = Get_PrintCount(iRow)
                If Not FlagCanPrint Then
                    Exit For
                End If
            End If
        Next

        For i As Integer = 0 To OutTeach.Items.Count - 1
            If OutTeach.Items(i).Selected = True And OutTeach.Items(i).Value <> "" Then
                iRow += 1
                If TeacherCon <> "" Then TeacherCon += ","
                TeacherCon += OutTeach.Items(i).Value

                FlagCanPrint = Get_PrintCount(iRow)
                If Not FlagCanPrint Then
                    Exit For
                End If
            End If
        Next

        If Not FlagCanPrint Then '數量超過範圍
            Errmsg += "選取老師姓名太多超過系統可列印數量，請減少選擇數量至 " & Cst_TechPrintCount & " 以下!!" & vbCrLf
            'Common.MessageBox(Me, "選取老師姓名太多超過系統可列印數量，請減少選擇數量至 " & Cst_TechPrintCount & " 以下!!")
            'Exit Function
        End If

        Select Case Printtype.SelectedIndex
            Case 0 '年月
                If Years.SelectedValue = "" Then
                    Errmsg += "統計月份 年度 為必填" & vbCrLf
                End If
                If Months.SelectedValue = "" Then
                    Errmsg += "統計月份 月份 為必填" & vbCrLf
                End If
                Dim start_month As String = ""
                If Years.SelectedValue <> "" AndAlso Months.SelectedValue <> "" Then
                    s_date1.Text = Years.SelectedValue & "/" & Months.SelectedValue & "/1"
                    start_month = s_date1.Text
                    If Not TIMS.IsDate1(s_date1.Text) Then
                        Errmsg += "統計月份 起始日期格式有誤" & vbCrLf
                    End If
                    s_date2.Text = Common.FormatDate(DateAdd(DateInterval.Month, 1, CDate(start_month)))
                    If Not TIMS.IsDate1(s_date2.Text) Then
                        Errmsg += "統計月份 起始日期格式有誤" & vbCrLf
                    End If
                End If

            Case 1 '依全期
                If Trim(s_date1.Text) <> "" Then s_date1.Text = Trim(s_date1.Text) Else s_date1.Text = ""
                If Trim(s_date2.Text) <> "" Then s_date2.Text = Trim(s_date2.Text) Else s_date2.Text = ""

                If s_date1.Text <> "" Then
                    If Not TIMS.IsDate1(s_date1.Text) Then
                        Errmsg += "排課期間 起始日期格式有誤" & vbCrLf
                    End If
                    If Errmsg = "" Then
                        s_date1.Text = CDate(s_date1.Text).ToString("yyyy/MM/dd")
                    End If
                Else
                    Errmsg += "排課期間 起始日期 為必填" & vbCrLf
                End If

                If s_date2.Text <> "" Then
                    If Not TIMS.IsDate1(s_date2.Text) Then
                        Errmsg += "排課期間 迄止日期格式有誤" & vbCrLf
                    End If
                    If Errmsg = "" Then
                        s_date2.Text = CDate(s_date2.Text).ToString("yyyy/MM/dd")
                    End If
                Else
                    Errmsg += "排課期間 迄止日期 為必填" & vbCrLf
                End If

                If Errmsg = "" Then
                    If s_date1.Text.ToString <> "" AndAlso s_date2.Text.ToString <> "" Then
                        If CDate(s_date1.Text) > CDate(s_date2.Text) Then
                            Errmsg += "【開訓期間】的起日不得大於迄日!!" & vbCrLf
                        End If
                    End If
                End If

        End Select

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '列印
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim FlagCanPrint As Boolean = True '能列印
        Dim iRow As Integer = 0
        Dim TeacherCon As String = ""
        For i As Integer = 0 To InTeach.Items.Count - 1
            If InTeach.Items(i).Selected = True And InTeach.Items(i).Value <> "" Then
                iRow += 1
                If TeacherCon <> "" Then TeacherCon += ","
                TeacherCon += InTeach.Items(i).Value

                FlagCanPrint = Get_PrintCount(iRow)
                If Not FlagCanPrint Then
                    Exit For
                End If
            End If
        Next
        For i As Integer = 0 To OutTeach.Items.Count - 1
            If OutTeach.Items(i).Selected = True And OutTeach.Items(i).Value <> "" Then
                iRow += 1
                If TeacherCon <> "" Then TeacherCon += ","
                TeacherCon += OutTeach.Items(i).Value

                FlagCanPrint = Get_PrintCount(iRow)
                If Not FlagCanPrint Then
                    Exit For
                End If
            End If
        Next

        Dim start_month As String = ""
        Dim end_month As String = ""
        Dim old_SchoolDate_end As String = ""
        Select Case Printtype.SelectedIndex
            Case 0 '年月
                start_month = Years.SelectedValue & "/" & Months.SelectedValue & "/1" '標準日期 某年某月1號
                end_month = Common.FormatDate(DateAdd(DateInterval.Month, 1, CDate(start_month))) '標準日期 (下個月) 某年某月1號
                old_SchoolDate_end = DateAdd(DateInterval.Day, -1, CDate(end_month)) '-1
                end_month = old_SchoolDate_end '同值
            Case 1 '依全期
                start_month = s_date1.Text
                end_month = s_date2.Text
                old_SchoolDate_end = end_month '同值
        End Select

        Dim myValue As String = ""
        Dim fileName As String = ""

        myValue = "RID=" & sm.UserInfo.RID
        myValue &= "&TPlan=" & sm.UserInfo.TPlanID
        'myValue &= "&SDate1=" & s_date1.Text
        'myValue &= "&SDate2=" & s_date2.Text
        myValue &= "&TeachName=" & Trim(tTeachName.Text) '教師姓名
        Select Case Mode.SelectedValue
            Case "1" '依班別統計
                fileName = "SD_04_009_R"

                myValue &= "&old_SchoolDate_start=" & start_month
                myValue &= "&old_SchoolDate_end=" & old_SchoolDate_end

                myValue &= "&start_date=" & start_month
                myValue &= "&end_date=" & end_month

                myValue &= "&TechID=" & TeacherCon
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_04_009_R", "RID=" & sm.UserInfo.RID & "&TPlan=" & sm.UserInfo.TPlanID & "&old_SchoolDate_start=" & start_month & "&old_SchoolDate_end=" & DateAdd(DateInterval.Day, -1, CDate(end_month)) & "&start_date=" & start_month & "&end_date=" & end_month & "&TechID=" & TeacherCon & "")
            Case Else '"2"依課程統計
                fileName = "SD_04_009_R_2"

                myValue &= "&old_SchoolDate_end=" & old_SchoolDate_end

                myValue &= "&SDate=" & start_month
                myValue &= "&EDate=" & end_month
                myValue &= "&OrgID=" & sm.UserInfo.OrgID '含機構

                myValue &= "&TechID=" & TeacherCon
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_04_009_R_2", "RID=" & sm.UserInfo.RID & "&TPlanID=" & sm.UserInfo.TPlanID & "&SDate=" & start_month & "&EDate=" & end_month & "&OrgID=" & sm.UserInfo.OrgID & "&TechID=" & TeacherCon & "&old_SchoolDate_end=" & DateAdd(DateInterval.Day, -1, CDate(end_month)))
        End Select

        'Call TIMS.SaveSqlCommon(fileName,  myValue)

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", fileName, myValue)
    End Sub

    Sub sUtl_Tech_Selected()
        Dim FlagCanPrint As Boolean = True '能列印
        Dim iRow As Integer = 0
        'Dim TeacherCon As String
        Dim blnTechSelected As Boolean = False
        iRow = 0

        For i As Integer = 0 To InTeach.Items.Count - 1
            If InTeach.Items(i).Selected Then
                iRow += 1
                FlagCanPrint = Get_PrintCount(iRow)
                blnTechSelected = InTeach.Items(i).Selected
            End If
            If blnTechSelected Then
                iRow += 1
                FlagCanPrint = Get_PrintCount(iRow)
                If Not FlagCanPrint Then
                    blnTechSelected = False
                    Exit For
                Else
                    InTeach.Items(i).Selected = blnTechSelected
                End If
            End If
        Next
        For i As Integer = 0 To OutTeach.Items.Count - 1
            If OutTeach.Items(i).Selected Then
                iRow += 1
                FlagCanPrint = Get_PrintCount(iRow)
                blnTechSelected = OutTeach.Items(i).Selected
            End If
            If blnTechSelected Then
                iRow += 1
                FlagCanPrint = Get_PrintCount(iRow)
                If Not FlagCanPrint Then
                    blnTechSelected = False
                    Exit For
                Else
                    OutTeach.Items(i).Selected = blnTechSelected
                End If
            End If
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call sUtl_Tech_Selected()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sUtl_Tech_Selected()
    End Sub

    Private Sub Printtype_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Printtype.SelectedIndexChanged

        Select Case Printtype.SelectedIndex
            Case 0 '年月
                'Months_TR.Style("display") = "inline"
                '上面是原寫法
                Months_TR.Style("display") = ""
                Allday_TR.Style("display") = "none"
                s_date1.Text = ""
                s_date2.Text = ""
            Case 1 '依全期
                'Dim SQL As String
                'Dim dr As DataRow
                'SQL = "Select STDate,FTDate From class_classinfo where OCID = '" & OCIDValue1.Value & "' "
                'dr = DbAccess.GetOneRow(SQL)
                'If Not dr Is Nothing Then
                '    s_date1.Text = dr("STDate")
                '    s_date2.Text = dr("FTDate")
                'End If
                Months_TR.Style("display") = "none"
                'Allday_TR.Style("display") = "inline"
                '上面是原寫法
                Allday_TR.Style("display") = ""
                Years.SelectedValue = ""
                Months.SelectedValue = ""
        End Select

    End Sub
End Class
