Partial Class CP_07_002_add
    Inherits AuthBasePage

    'select top 10 * from Stud_QuesTraining
    'select top 10 * from Stud_ForumRecord
    'Dim rid As String = ""
    'Dim ocid As String = ""
    'Dim PlanID As String = ""
    'Dim Qstatus As String = ""
    'Dim socid As String = ""
    'Dim SearchStr As String = "Search" & "CP_07_002_aspx"
    Dim cst_pSearchStr As String = "Searchcp_07_cp_07_002_aspx"

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        RIDValue.Value = Request("rid")
        OCID1.Value = Request("ocid")
        hid_socid.Value = Request("socid")
        hid_planid.Value = Request("PlanID")
        hidQstatus.Value = Request("status") 'add 'edit
        If Not Session(cst_pSearchStr) Is Nothing Then Session(cst_pSearchStr) = Session(cst_pSearchStr)
        If Not IsPostBack Then

            'Me.ViewState("SearchStr") = Session(SearchStr)
            'Dim MyArray As Array = Split(Session(SearchStr), "&")
            'Dim MyItem As String
            'Dim MyValue As String
            'For i As Integer = 0 To MyArray.Length - 1
            '    MyItem = Split(MyArray(i), "=")(0)
            '    MyValue = Split(MyArray(i), "=")(1)
            '    Select Case MyItem
            '        Case "center"
            '            center.Value = MyValue
            '        Case "RIDValue"
            '            RIDValue.Value = MyValue
            '        Case "TMID1"
            '            TMID1.Value = MyValue
            '        Case "OCID1"
            '            OCID1.Value = MyValue
            '        Case "TMIDValue1"
            '            TMIDValue1.Value = MyValue
            '        Case "OCIDValue1"
            '            OCIDValue1.Value = MyValue
            '        Case "start_date"
            '            STDate1.Value = MyValue
            '        Case "end_date"
            '            STDate2.Value = MyValue
            '    End Select
            'Next
            'Session(SearchStr) = Nothing
            ChkBox01.Checked = False
            If Request("ocid") <> "" AndAlso Request("socid") <> "" Then
                Call sUtl_Get_SOCID_Students()
                Call sUtl_LoadData1(Request("socid"))
            End If
        End If
        'Me.ViewState("_SearchStr") = Session("_" & SearchStr)
        'Session("_" & SearchStr) = Nothing

        bt_save.Attributes.Add("onclick", "return confirm('確定儲存?');")
        'txtFillDate.Attributes.Add("onchange", "ChgFillDate();")
        'End If
    End Sub

    '取得所有學員並設定
    Private Sub sUtl_Get_SOCID_Students()
        'Dim dt As New DataTable
        'Dim sql As String = ""
        'Dim conn As New SqlConnection
        'conn = DbAccess.GetConnection()
        'conn.Open()
        Try
            Dim sql As String = ""
            Dim dt As DataTable
            sql = "" & vbCrLf
            sql += " SELECT a.StudentID" & vbCrLf
            sql += " , case when Len(a.StudentID)=12 then b.Name+'('+substr(a.StudentID,-3)+')'" & vbCrLf
            sql += "  else b.Name+'('+substr(a.StudentID,-2)+')' end as Name" & vbCrLf
            sql += " , a.SOCID" & vbCrLf
            sql += " FROM Class_StudentsOfClass a" & vbCrLf
            sql += " JOIN Stud_StudentInfo b  ON a.SID=b.SID" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and a.OCID='" & Request("OCID") & "'" & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)

            dt.DefaultView.Sort = "StudentID"
            With ddl_SOCID
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "SOCID"
                .DataBind()
                '.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End With
            'If Request("socid") <> "" Then
            '    Common.SetListItem(ddl_SOCID, Request("socid"))
            'End If
            Common.SetListItem(ddl_SOCID, Request("socid"))
            'dt.Dispose()
        Catch ex As Exception
            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "alert('發生錯誤!! " & ex.Message.ToString() & "');" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)
        End Try

    End Sub

    Private Sub sUtl_LoadData1(ByVal SEL_SOCID As String)
        'Dim da As New SqlDataAdapter
        'Dim dt1 As New DataTable
        'Dim dt2 As New DataTable
        'Dim dr As DataRow
        'Dim sql As String = ""
        'Dim strScript As String = ""
        'Dim conn As New SqlConnection
        'conn = DbAccess.GetConnection()
        'conn.Open()
        Try
            Dim dr As DataRow
            Dim dt2 As DataTable
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql += " select g.*" & vbCrLf
            sql += " ,i.PlanName" & vbCrLf
            sql += " ,j.Name as DistIDName" & vbCrLf
            sql += " ,a.OCID,a.STDate,a.FTDate,a.RID,b.StudentID,c.Name,e.OrgName,b.socid" & vbCrLf
            sql += " ,a.ClassCName+'(第'+a.CyclType+'期)' ClassCName" & vbCrLf
            sql += " ,g.CREATEDATE " & vbCrLf
            sql += " ,g.DISCDATE " & vbCrLf
            sql += " from Class_ClassInfo a" & vbCrLf
            sql += " join Class_StudentsOfClass b on a.OCID=b.OCID" & vbCrLf
            sql += " join Stud_StudentInfo c on b.SID=c.SID" & vbCrLf
            sql += " join Auth_Relship f  on a.RID=f.RID" & vbCrLf
            sql += " join Org_OrgInfo e on f.OrgID=e.Orgid" & vbCrLf
            sql += " left join  Stud_ForumRecord g on b.socid=g.socid" & vbCrLf
            sql += " left join id_plan  h on  a.planid=h.planid" & vbCrLf
            sql += " left join key_plan i on  h.TPlanID=i.TPlanID" & vbCrLf
            sql += " left join ID_District  j  on  h.DistID=j.DistID" & vbCrLf
            sql += " where 1=1"
            sql += " and a.PlanID='" & hid_planid.Value & "' "
            sql += " and h.TPlanID<>'28'  " '非產投計畫
            sql += " and a.ocid='" & OCID1.Value & "'"
            If SEL_SOCID <> "" Then
                sql += " and  b.socid='" & SEL_SOCID & "'"
            End If
            sql += " and  a.rid='" & RIDValue.Value & "'"
            dt2 = DbAccess.GetDataTable(sql, objconn)


            If dt2.Rows.Count > 0 Then
                dr = dt2.Rows(0)

                If Not IsDBNull(dr("CREATEDATE")) Then
                    hidQstatus.Value = "edit"
                    txt_Content.Text = Convert.ToString(dr("Content"))

                    If Convert.ToString(dr("DISCDATE")) <> "" Then
                        txtFillDate.Value = Common.FormatDate(dr("DISCDATE"))
                        lb_FillDate.Text = Convert.ToString(CDate(txtFillDate.Value).Year()) & "年" & Convert.ToString(CDate(txtFillDate.Value).Month()) & "月" & Convert.ToString(CDate(txtFillDate.Value).Day()) & "日"
                        WeekDay.Text = TIMS.ConvertNum(Convert.ToString(CDate(txtFillDate.Value).DayOfWeek()))
                    End If

                    msg.Text = ""
                    If IsDBNull(dr("ReferSignIn")) Then
                        ChkBox01.Checked = False
                    Else
                        ChkBox01.Checked = True
                    End If
                Else
                    hidQstatus.Value = "add"
                    msg.Text = "尚未填寫！"
                    txt_Content.Text = ""
                    'lb_FillDate.Text = Convert.ToString(DateTime.Now.Year()) & "年" & Convert.ToString(DateTime.Now.Month()) & "月" & Convert.ToString(DateTime.Now.Day()) & "日"
                    'WeekDay.Text = ConvertNum(Convert.ToString(DateTime.Now.DayOfWeek()))
                    Dim aDate As String = TIMS.GetSysDate(objconn)
                    txtFillDate.Value = Common.FormatDate(aDate)
                    lb_FillDate.Text = Convert.ToString(CDate(aDate).Year()) & "年" & Convert.ToString(CDate(aDate).Month()) & "月" & Convert.ToString(CDate(aDate).Day()) & "日"
                    WeekDay.Text = TIMS.ConvertNum(Convert.ToString(CDate(aDate).DayOfWeek()))
                    'lb_FillDate.Text = ""
                    'WeekDay.Text = ""                 
                End If

                lb_STDate.Text = Convert.ToString(Convert.ToDateTime(dr("STDate")).Year()) & "年" & Convert.ToString(Convert.ToDateTime(dr("STDate")).Month()) & "月" & Convert.ToString(Convert.ToDateTime(dr("STDate")).Day()) & "日"
                lb_FTDate.Text = Convert.ToString(Convert.ToDateTime(dr("FTDate")).Year()) & "年" & Convert.ToString(Convert.ToDateTime(dr("FTDate")).Month()) & "月" & Convert.ToString(Convert.ToDateTime(dr("FTDate")).Day()) & "日"
                lb_OrgName.Text = Convert.ToString(dr("OrgName"))
                lb_OCID.Text = Convert.ToString(dr("ClassCName"))
                studID.Text = Convert.ToString(dr("StudentID")).ToUpper()
                lb_PlanName.Text = Convert.ToString(dr("PlanName"))
            End If

            'dt2.Dispose()
        Catch ex As Exception
            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "alert('發生錯誤!! " & ex.Message.ToString() & "');" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)

        End Try
    End Sub

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        sUtl_SaveData()
    End Sub

    '儲存sql
    Private Sub sUtl_SaveData()
        Dim errMsg As String = ""
        If Not chkInputAnswer(errMsg) Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If

        Try
            'If Qstatus = "add" Then
            Dim sql As String = ""
            Call TIMS.OpenDbConn(objconn)
            Dim oCmd As SqlCommand
            Select Case hidQstatus.Value
                Case "add"
                    sql = ""
                    sql &= " INSERT INTO Stud_ForumRecord (" & vbCrLf
                    sql += " SOCID,OCID,Content " & vbCrLf
                    sql += " ,CreateAcct,CreateDate " & vbCrLf
                    sql += " ,ModifyAcct,ModifyDate " & vbCrLf
                    sql += " ,DISCDATE " & vbCrLf
                    sql += " ,ReferSignIn " & vbCrLf
                    sql += " ) VALUES ( " & vbCrLf
                    sql += " @SOCID,@OCID,@Content " & vbCrLf
                    sql += " ,@CreateAcct, getdate()" & vbCrLf
                    sql += " ,@ModifyAcct, getdate()" & vbCrLf
                    sql += " ,@DISCDATE " & vbCrLf
                    sql += " ,@ReferSignIn " & vbCrLf
                    sql += " )" & vbCrLf
                    oCmd = New SqlCommand(sql, objconn)
                    With oCmd
                        .Parameters.Clear()
                        .Parameters.Add("SOCID", SqlDbType.VarChar).Value = ddl_SOCID.SelectedValue
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID1.Value
                        .Parameters.Add("Content", SqlDbType.VarChar).Value = IIf(Trim(txt_Content.Text) = "", Convert.DBNull, Trim(txt_Content.Text))
                        .Parameters.Add("CreateAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        .Parameters.Add("DISCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(txtFillDate.Value)
                        .Parameters.Add("ReferSignIn", SqlDbType.VarChar).Value = IIf(ChkBox01.Checked = True, "Y", Convert.DBNull)
                        .ExecuteNonQuery()
                    End With
                Case Else
                    sql = ""
                    sql &= " UPDATE Stud_ForumRecord "
                    sql += " set Content= @Content"
                    sql += " ,DISCDATE= @DISCDATE "
                    sql += " ,ModifyAcct= @ModifyAcct "
                    sql += " ,ModifyDate=getdate()"
                    sql += " ,ReferSignIn= @ReferSignIn"
                    sql += " where  1=1 "
                    sql += " and  ocid=@OCID "
                    sql += " and  socid=@SOCID"
                    oCmd = New SqlCommand(sql, objconn)
                    With oCmd
                        .Parameters.Clear()
                        .Parameters.Add("Content", SqlDbType.VarChar).Value = IIf(Trim(txt_Content.Text) = "", Convert.DBNull, Trim(txt_Content.Text))
                        .Parameters.Add("DISCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(txtFillDate.Value)
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        .Parameters.Add("ReferSignIn", SqlDbType.VarChar).Value = IIf(ChkBox01.Checked = True, "Y", Convert.DBNull)
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID1.Value
                        .Parameters.Add("SOCID", SqlDbType.VarChar).Value = ddl_SOCID.SelectedValue
                        .ExecuteNonQuery()
                    End With
            End Select

            Common.RespWrite(Me, "<Script language='javascript'>alert('儲存成功！');</script>")

        Catch ex As Exception

            Common.RespWrite(Me, "<Script language='javascript'>alert('儲存失敗，發生錯誤!! " & ex.Message.ToString() & "');</script>")

        End Try

    End Sub

    Private Function chkInputAnswer(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True
        errMsg = ""
        If txt_Content.Text = "" Then
            errMsg += "【意見欄位】尚未填寫任何內容！" & vbCrLf
        End If

        If CheckInput(txt_Content.Text) <> "" Then
            errMsg += "【意見欄位】輸入字串中含有不合法字元" & CheckInput(txt_Content.Text) & "\n" & vbCrLf
        End If
        If txtFillDate.Value <> "" Then
            If CDate(CDate(txtFillDate.Value).ToString("yyyy-MM-dd")) > CDate(Date.Now.ToString("yyyy-MM-dd")) Then
                errMsg += "填表日期不得大於今日！" & vbCrLf
            End If
        End If

        If errMsg <> "" Then rst = False
        Return rst
    End Function

    'Private Function ReplaceStr(ByVal InputStr)
    '    InputStr = Replace(InputStr, "'", "''")
    '    Return InputStr
    'End Function

    Function CheckInput(ByVal parameter As String) As String
        Dim rst As String = ""
        Dim blackList As String() = {"'", "--", ";--", ";", "/*", "*/", "@@", _
                                    "@", "char", "nchar", "varchar", "nvarchar", "alter", _
                                    "begin", "cast", "create", "cursor", "declare", "delete", _
                                    "drop", "end", "exec", "execute", "fetch", "insert", _
                                    "kill", "open", "select", "sys", "sysobjects", "syscolumns", "table", _
                                    "update"}
        Dim strPos As Integer = 0
        Dim blackListlen As Integer = 0
        Dim InputStr As String = parameter
        'Dim errMsg As String = ""
        For i As Integer = 0 To blackList.Length - 1
            blackListlen = blackList(i).Length()
            strPos = InStr(1, InputStr, blackList(i))
            If strPos <> 0 Then
                Select Case blackList(i)
                    Case "'"
                        If rst <> "" Then rst &= ","
                        rst += "「 單引號‘」"
                    Case Else
                        If rst <> "" Then rst &= ","
                        rst += "「 " & blackList(i) & " 」"
                End Select
            End If
        Next

        Return rst
    End Function

    'Private Sub PrePage()
    '    Dim parmstr As String = ""
    '    Session(SearchStr) = Me.ViewState("SearchStr")
    '    Session("_" & SearchStr) = Me.ViewState("_SearchStr")

    '    parmstr = "'CP_07_002.aspx"
    '    If ocid <> "" Then
    '        parmstr += "?ocid=" & ocid
    '    End If
    '    If socid <> "" Then
    '        parmstr += "&socid=" & socid
    '    End If
    '    If rid <> "" Then
    '        parmstr += "&rid=" & rid
    '    End If
    '    parmstr += "';"

    '    If Me.ViewState("parmstr") <> "" Then
    '        Common.RespWrite(Me, "<script>location.href=" & Me.ViewState("parmstr") & "</script>")
    '    Else
    '        Common.RespWrite(Me, "<script>location.href=" & parmstr & "</script>")
    '    End If
    'End Sub

    'Private Sub bt_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_back.Click
    '    PrePage()
    'End Sub

    Private Sub ddl_SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_SOCID.SelectedIndexChanged
        hid_socid.Value = ddl_SOCID.SelectedValue
        Call sUtl_LoadData1(hid_socid.Value)
    End Sub

    Private Sub LinkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
        If txtFillDate.Value <> "" Then
            lb_FillDate.Text = Convert.ToString(CDate(txtFillDate.Value).Year()) & "年" & Convert.ToString(CDate(txtFillDate.Value).Month()) & "月" & Convert.ToString(CDate(txtFillDate.Value).Day()) & "日"
            WeekDay.Text = TIMS.ConvertNum(Convert.ToString(CDate(txtFillDate.Value).DayOfWeek()))
        Else
            lb_FillDate.Text = ""
            WeekDay.Text = ""
            Common.MessageBox(Me, "未輸入日期。")
            Exit Sub
        End If
    End Sub

    Protected Sub bt_back_Click(sender As Object, e As EventArgs) Handles bt_back.Click
        Dim url1 As String = String.Concat("CP_07_002.aspx?ID=", TIMS.Get_MRqID(Me))
        TIMS.Utl_Redirect1(Me, url1)
    End Sub

End Class
