Partial Class TC_01_005
    Inherits AuthBasePage

    Dim ProcessType As String = ""
    'Dim objreader As SqlDataReader
    'Dim FunDr As DataRow
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    'Dim Gdt2 As DataTable
    Dim vsTB_CourseID As String = ""
    Dim vsTB_CourseName As String = ""

    'Dim MainDt As DataTable
    Dim GsCmd1 As SqlCommand
    Dim GsCmd2 As SqlCommand
    Dim GsCmd3 As SqlCommand
    Dim GsCmd4 As SqlCommand

    '課程代碼是有區分年度 轄區 業務ID的 2009年啟用吧??
    '2015
    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DG_Course
        '分頁設定 End

        Dim sql As String = ""
        Select Case sm.UserInfo.LID
            Case "2"
                sql = "" & vbCrLf
                sql &= " SELECT PLANNAME+'-'+ORGNAME PO FROM VIEW_RWPLANRID" & vbCrLf
                sql &= " WHERE RID=@RID" & vbCrLf
                GsCmd1 = New SqlCommand(sql, objconn)
            Case Else
                sql = "" & vbCrLf
                sql &= " SELECT PLANNAME+'-'+ORGNAME PO FROM VIEW_RWPLANRID" & vbCrLf
                sql &= " WHERE RID=@RID" & vbCrLf
                sql &= " AND RWPLANID=@RWPLANID" & vbCrLf
                GsCmd1 = New SqlCommand(sql, objconn)
        End Select

        sql = "SELECT 'X' COURID FROM COURSE_COURSEINFO WHERE MAINCOURID=@CourID" '" & drv("CourID") & "'"
        GsCmd2 = New SqlCommand(sql, objconn)
        sql = "SELECT 'X' CTSID FROM CLASS_TMPSCHEDULE WHERE COURSEID=@CourID" '" & drv("CourID") & "'"
        GsCmd3 = New SqlCommand(sql, objconn)
        Dim check_schedule As String = "" '"select CSID from Class_Schedule  where Class1='" & drv("CourID") & "' or Class2='" & drv("CourID") & "' or Class3='" & drv("CourID") & "' or Class4='" & drv("CourID") & "' or Class5='" & drv("CourID") & "' or Class6='" & drv("CourID") & "' or Class7='" & drv("CourID") & "' or Class8='" & drv("CourID") & "' or Class9='" & drv("CourID") & "' or Class10='" & drv("CourID") & "' or Class11='" & drv("CourID") & "' or Class12='" & drv("CourID") & "'"
        check_schedule = "" & vbCrLf
        check_schedule += " SELECT 'X' CSID FROM CLASS_SCHEDULE" & vbCrLf
        check_schedule += " WHERE @CourID in (" & vbCrLf
        check_schedule += " Class1" & vbCrLf
        check_schedule += " ,Class2" & vbCrLf
        check_schedule += " ,Class3" & vbCrLf
        check_schedule += " ,Class4" & vbCrLf
        check_schedule += " ,Class5" & vbCrLf
        check_schedule += " ,Class6" & vbCrLf
        check_schedule += " ,Class7" & vbCrLf
        check_schedule += " ,Class8" & vbCrLf
        check_schedule += " ,Class9" & vbCrLf
        check_schedule += " ,Class10" & vbCrLf
        check_schedule += " ,Class11" & vbCrLf
        check_schedule += " ,Class12) " & vbCrLf
        GsCmd4 = New SqlCommand(check_schedule, objconn)

        If Not IsPostBack Then
            tb_DG_Course.Visible = False
            bt_search.Attributes("onclick") = "aloader2on();"
            'DrpYear = TIMS.GetSyear(DrpYear, 2005)
            'DrpYear.SelectedValue = sm.UserInfo.Years
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PageControler1.Visible = False
            'Button1.Enabled = False
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame('inline');"
            center.Style("CURSOR") = "hand"
            HistoryRID.Attributes("onclick") = "ShowFrame('none');"
        End If

        'FunDr = FunDrArray(0)
        'bt_add.Enabled = False
        'Button1.Enabled = False
        'If au.blnCanAdds Then bt_add.Enabled = True
        'If au.blnCanAdds Then Button1.Enabled = True
        'bt_search.Enabled = False
        'print.Enabled = False
        'If au.blnCanSech Then bt_search.Enabled = True
        'If au.blnCanSech Then print.Enabled = True

        check_del.Value = "1"
        check_mod.Value = "1"

        'check_del.Value = "0"
        'If au.blnCanDel Then check_del.Value = "1"
        'check_mod.Value = "0"
        'If au.blnCanMod Then check_mod.Value = "1"

        ProcessType = TIMS.ClearSQM(Request("ProcessType"))
        If Not Page.IsPostBack Then
            CB_Valid.Checked = True '是否有效
            'Session("MySreach") 沒有搜尋值還是做一次 Search1
            If Session("MySreach") Is Nothing Then
                Select Case ProcessType
                    Case "del"
                        Call Search1()
                End Select
            End If

            '取出搜尋值
            If Not Session("MySreach") Is Nothing Then
                'Me.ViewState("PageIndex") = ""
                Dim MyValue As String = ""
                Dim str1 As String = Convert.ToString(Session("MySreach"))
                'Me.ViewState("MySreach") = Session("MySreach")
                center.Text = TIMS.GetMyValue(str1, "center")
                RIDValue.Value = TIMS.GetMyValue(str1, "RIDValue")

                TB_CourseID.Text = TIMS.GetMyValue(str1, "TB_CourseID")
                TB_CourseName.Text = TIMS.GetMyValue(str1, "TB_CourseName")

                MyValue = TIMS.GetMyValue(str1, "Classification1_List")
                If MyValue <> "" Then
                    Common.SetListItem(Classification1_List, MyValue)
                End If
                MyValue = TIMS.GetMyValue(str1, "Classification2_List")
                If MyValue <> "" Then
                    Common.SetListItem(Classification2_List, MyValue)
                End If

                CB_Valid.Checked = TIMS.GetMyValue(str1, "CB_Valid") '是否有效
                TB_career_id.Text = TIMS.GetMyValue(str1, "TB_career_id")
                Classid.Text = TIMS.GetMyValue(str1, "Classid")
                Classid_Hid.Value = TIMS.GetMyValue(str1, "Classid_Hid")
                trainValue.Value = TIMS.GetMyValue(str1, "trainValue")

                MyValue = TIMS.GetMyValue(str1, "submit")
                If MyValue = "1" Then
                    MyValue = TIMS.GetMyValue(str1, "PageIndex")
                    'Me.ViewState("PageIndex") = TIMS.GetMyValue(Convert.ToString(Me.ViewState("_SearchStr")), "PageIndex")
                    If IsNumeric(MyValue) Then PageControler1.PageIndex = Val(MyValue)
                    Call Search1()
                    'bt_search_Click(sender, e)
                    'If IsNumeric(Me.ViewState("PageIndex")) Then
                    '    PageControler1.PageIndex = Me.ViewState("PageIndex")
                    '    PageControler1.CreateData()
                    'End If
                End If

                Session("MySreach") = Nothing
            End If

        End If

    End Sub

    Function GetFields(ByVal dr As DataRow) As String() ' creates comma delineated string of column names
        Dim strFields() As String = Nothing
        Dim a As Integer = 0
        ReDim strFields(dr.Table.Columns.Count - 1)
        For Each col As DataColumn In dr.Table.Columns
            strFields(a) = col.ColumnName
            a += 1
        Next
        Return strFields
    End Function

    Function RowsToTable(ByVal drs As DataRow()) As DataTable 'converts dr() to datatable
        Dim dt As New DataTable
        'Dim _Fields As String() = Nothing
        'If _Fields Is Nothing Then _Fields = GetFields(drs(0))
        Dim _Fields As String() = GetFields(drs(0))
        For Each field As String In _Fields
            dt.Columns.Add(field.Trim, drs(0).Table.Columns(field.Trim).DataType)
        Next
        For Each dr As DataRow In drs
            dt.ImportRow(dr)
        Next
        Return dt
    End Function

    '查詢
    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, DG_Course)

        Dim searchstr2_C As String = "" ' a.Master條件
        Dim searchstr3_P As String = "" ' p.父層條件
        searchstr2_C = ""
        searchstr3_P = "" '父層條件

        Me.TB_CourseID.Text = TIMS.ClearSQM(Me.TB_CourseID.Text)
        Me.TB_CourseName.Text = TIMS.ClearSQM(Me.TB_CourseName.Text)
        vsTB_CourseID = Me.TB_CourseID.Text '.Replace("'", "''")
        vsTB_CourseName = Me.TB_CourseName.Text '.Replace("'", "''")

        If vsTB_CourseID <> "" Then searchstr2_C += " and a.CourseID='" & vsTB_CourseID & "'" & vbCrLf
        If vsTB_CourseName <> "" Then searchstr2_C += " and a.CourseName like '" & vsTB_CourseName & "%'" & vbCrLf

        Dim v_Classification1 As String = TIMS.GetListValue(Classification1_List)
        Dim v_Classification2 As String = TIMS.GetListValue(Classification2_List)
        Select Case v_Classification1'Classification1_List.SelectedValue
            Case 0
            Case Else
                searchstr2_C += " and a.Classification1='" & v_Classification1 & "'" & vbCrLf
        End Select
        If Classification2_List.SelectedValue <> "3" Then
            searchstr2_C += " and a.Classification2='" & v_Classification2 & "'" & vbCrLf
        End If
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If trainValue.Value <> "" Then
            searchstr2_C += " and a.TMID ='" & trainValue.Value & "'" & vbCrLf
        End If
        Classid_Hid.Value = TIMS.ClearSQM(Classid_Hid.Value)
        If Classid_Hid.Value <> "" Then
            searchstr2_C += " and a.CLSID ='" & Classid_Hid.Value & "'" & vbCrLf
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        '是否有效 
        If CB_Valid.Checked Then
            '有效
            searchstr2_C += " and a.Valid='Y'" & vbCrLf
            searchstr2_C += " and a.RID='" & RIDValue.Value & "' " & vbCrLf '有效RID --

            searchstr3_P += " and p.Valid='Y'" & vbCrLf '父層條件
            searchstr3_P += " and p.RID='" & RIDValue.Value & "' " & vbCrLf '有效RID --
        Else
            '有效含無效 (同機構不同業務)
            searchstr2_C += " AND EXISTS (SELECT 'x' FROM Auth_Relship c WHERE c.RID='" & RIDValue.Value & "' and c.OrgID =a.OrgID)" & vbCrLf
            searchstr2_C += " AND (1!=1 " & vbCrLf
            searchstr2_C += " or a.Valid='N' " & vbCrLf '失效
            '(同機構不同業務)
            searchstr2_C += " or EXISTS (SELECT 'x' FROM Auth_Relship c WHERE c.RID='" & RIDValue.Value & "' and c.OrgID =a.OrgID) " & vbCrLf
            searchstr2_C += " )" & vbCrLf

            '父層條件'有效含無效 (同機構不同業務)
            searchstr3_P += " AND EXISTS (SELECT 'x' FROM Auth_Relship c WHERE c.RID='" & RIDValue.Value & "' and c.OrgID =p.OrgID)" & vbCrLf
            searchstr3_P += " AND (1!=1 " & vbCrLf
            searchstr3_P += " or p.Valid='N' " & vbCrLf '父層條件
            '(同機構不同業務)
            searchstr3_P += " or EXISTS (SELECT 'x' FROM Auth_Relship c WHERE c.RID='" & RIDValue.Value & "' and c.OrgID =p.OrgID) " & vbCrLf
            searchstr3_P += " )" & vbCrLf
        End If

        Dim sqlstr As String = "" & vbCrLf
        sqlstr = "" & vbCrLf
        sqlstr &= " select * from ( " & vbCrLf

        sqlstr &= " select case when a.MainCourID is null then a.CourID " & vbCrLf
        sqlstr &= "  else a.MainCourID end CourGroup" & vbCrLf
        sqlstr &= " ,a.CourID" & vbCrLf
        sqlstr &= " ,a.CourseID" & vbCrLf
        sqlstr &= " ,a.CourseName" & vbCrLf
        sqlstr &= " ,a.Hours" & vbCrLf
        sqlstr &= " ,a.Classification1" & vbCrLf
        sqlstr &= " ,a.Classification2" & vbCrLf
        sqlstr &= " ,a.MainCourID  " & vbCrLf
        'sqlstr += " ,p.CourID MainCourID  " & vbCrLf
        sqlstr &= " ,case when (a.MainCourID is not null) and (p.CourID is null) then '1' end ErrorData" & vbCrLf
        sqlstr &= " ,a.OrgID" & vbCrLf
        sqlstr &= " ,a.RID" & vbCrLf
        sqlstr &= " ,p.OrgID pOrgID" & vbCrLf
        sqlstr &= " ,p.RID pRID" & vbCrLf
        sqlstr &= " ,p.CourseName pCourseName" & vbCrLf
        sqlstr &= " FROM COURSE_COURSEINFO a " & vbCrLf
        sqlstr &= " LEFT JOIN COURSE_COURSEINFO p on p.CourID=a.MainCourID " & vbCrLf
        sqlstr += searchstr3_P & vbCrLf
        'sqlstr += " 	LEFT join ID_Class ic on ic.CLSID=a.CLSID " & vbCrLf
        'sqlstr += "  and a.RID='B1605004' " & vbCrLf 'sqlstr += "  and a.Valid='Y'" & vbCrLf
        sqlstr &= " WHERE 1=1 " & vbCrLf
        sqlstr &= searchstr2_C & vbCrLf
        sqlstr &= " ) g" & vbCrLf
        sqlstr &= " ORDER BY CourGroup ,MainCourID ,CourID" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        'Panel1.Visible = True
        'Panel.Visible = False
        msg.Text = "查無資料!!"
        tb_DG_Course.Visible = False
        DG_Course.Visible = False
        If dt.Rows.Count = 0 Then Return

        'For i As Integer = 0 To dt.Rows.Count - 1
        '    couid+=
        'Next
        'MainDt = DbAccess.GetDataTable(sqlstr, objconn)
        'Panel1.Visible = False
        'Panel.Visible = True
        msg.Text = ""
        tb_DG_Course.Visible = True
        DG_Course.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    'SQL 查詢 
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call Search1()
    End Sub

    '新增
    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        Dim RqID As String = TIMS.Get_MRqID(Me)
        TB_CourseID.Text = TIMS.ClearSQM(TB_CourseID.Text)
        TB_CourseName.Text = TIMS.ClearSQM(TB_CourseName.Text)
        Call GetSearchStr()
        'Response.Redirect("TC_01_005_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "")
        '20100208 按新增時代查詢之 課程代碼 & 課程名稱
        'Response.Redirect("TC_01_005_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&ClassID=" & TB_CourseID.Text & "&ClassName=" & TB_CourseName.Text & "")
        Dim url1 As String = "TC_01_005_add.aspx?ProcessType=Insert&ID=" & RqID & "&ClassID=" & TB_CourseID.Text & "&ClassName=" & TB_CourseName.Text
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '列印-課程代碼表
    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click
        Dim RqID As String = TIMS.Get_MRqID(Me)
        GetSearchStr()
        'Response.Redirect("TC_01_005_print.aspx?ID=" & Request("ID"))
        Dim url1 As String = "TC_01_005_print.aspx?ID=" & RqID 'TIMS.ClearSQM(Request("ID"))
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '匯入
    Private Sub BTN_IMP1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_IMP1.Click
        Dim RqID As String = TIMS.Get_MRqID(Me)
        GetSearchStr() 'Button1_Click/Button1
        'Response.Redirect("TC_01_005_import.aspx?ID=" & Request("ID") & "")
        Dim url1 As String = "TC_01_005_import.aspx?ID=" & RqID 'TIMS.ClearSQM(Request("ID"))
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '清除
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim RqID As String = TIMS.Get_MRqID(Me)
        Classid_Hid.Value = ""
        Classid.Text = ""
    End Sub

    Private Sub DG_Course_ItemCommand1(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Course.ItemCommand
        Select Case e.CommandName
            Case "edit"
                Dim RqID As String = TIMS.Get_MRqID(Me)
                Call GetSearchStr()
                'Response.Redirect("TC_01_005_add.aspx?ProcessType=Update" & e.CommandArgument & "&ID=" & Request("ID"))
                Dim url1 As String = "TC_01_005_add.aspx?ProcessType=Update" & e.CommandArgument & "&ID=" & RqID 'TIMS.ClearSQM(Request("ID"))
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select

    End Sub

    Private Sub DG_Course_ItemDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Course.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.Cells(0).Text = "序號"
                If Me.cb_CourID.Checked Then
                    e.Item.Cells(0).Text = "匯入用<BR>代碼"
                End If
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim strTmp1 As String = ""
                strTmp1 = e.Item.ItemIndex + 1 + (DG_Course.CurrentPageIndex * DG_Course.PageSize)
                If Me.cb_CourID.Checked Then strTmp1 = Convert.ToString(drv("CourID"))
                e.Item.Cells(0).Text = strTmp1

                strTmp1 = ""
                Select Case Convert.ToString(drv("Classification1"))
                    Case "1"
                        strTmp1 = "學科"
                    Case "2"
                        strTmp1 = "術科"
                End Select
                e.Item.Cells(4).Text = strTmp1

                strTmp1 = ""
                Select Case Convert.ToString(drv("Classification2"))
                    Case "0"
                        strTmp1 = "共同"
                    Case "1"
                        strTmp1 = "一般"
                    Case "2"
                        strTmp1 = "專業"
                End Select
                e.Item.Cells(5).Text = strTmp1

                Dim but As Button = e.Item.FindControl("Button4") '修改
                Dim but_del As Button = e.Item.FindControl("Button5") '刪除

                Dim sCmdArg As String = ""
                sCmdArg = ""
                sCmdArg &= "&courid=" & Convert.ToString(drv("CourID"))
                sCmdArg &= "&ErrorData=" & Convert.ToString(drv("ErrorData"))
                but.CommandArgument = sCmdArg

                but_del.CommandArgument = Convert.ToString(drv("CourID"))

                but_del.Enabled = If(check_del.Value = "1", True, False)
                but.Enabled = If(check_mod.Value = "1", True, False)

                'dr1 = DbAccess.GetOneRow(check_Class_tmp, objconn) '是否有排課(Class_TmpSchedule) TOP 1
                Dim is_parent As String = ""
                Dim ch_courid As Integer = 0
                Dim dr1 As DataRow = Nothing
                Dim dt As New DataTable
                With GsCmd3
                    .Parameters.Clear()
                    .Parameters.Add("CourID", SqlDbType.VarChar).Value = Convert.ToString(drv("CourID"))
                    dt.Load(.ExecuteReader())
                    If dt.Rows.Count > 0 Then dr1 = dt.Rows(0)
                End With
                'dr2 = DbAccess.GetOneRow(check_schedule, objconn) '是否有排課(Class_Schedule) TOP 1
                Dim dr2 As DataRow = Nothing
                Dim dt2 As New DataTable
                With GsCmd4
                    .Parameters.Clear()
                    .Parameters.Add("CourID", SqlDbType.VarChar).Value = Convert.ToString(drv("CourID"))
                    dt2.Load(.ExecuteReader())
                    If dt2.Rows.Count > 0 Then dr2 = dt2.Rows(0)
                End With

                ' ElseIf Gdt2.Select("CourID= '" & drv("CourID") & "'").Length > 0 Then  '是否有排課(Class_Schedule)
                If Not dr1 Is Nothing Then '是否有排課(Class_TmpSchedule)
                    ch_courid = drv("CourID")
                    is_parent = 0
                ElseIf Not dr2 Is Nothing Then  '是否有排課(Class_Schedule)
                    ch_courid = drv("CourID")
                    is_parent = 0
                Else
                    ch_courid = 0
                    If Convert.IsDBNull(drv("MainCourID")) Then '為主課程
                        Dim dr3 As DataRow = Nothing
                        Dim dt3 As New DataTable
                        With GsCmd2
                            .Parameters.Clear()
                            .Parameters.Add("CourID", SqlDbType.VarChar).Value = Convert.ToString(drv("CourID"))
                            dt3.Load(.ExecuteReader())
                            If dt3.Rows.Count > 0 Then dr3 = dt3.Rows(0)
                        End With
                        'dr3 = DbAccess.GetOneRow(course_list, objconn)
                        If dr3 Is Nothing Then  '是否有子課程
                            is_parent = "false"
                        Else
                            is_parent = "true"
                        End If
                    Else '非主課程
                        is_parent = "false"
                    End If
                End If
                Dim RqID As String = TIMS.Get_MRqID(Me)
                but_del.Attributes.Add("onclick", "but_del(" & drv("CourID") & ",'" & ch_courid & "'," & is_parent & "," & RqID & "); return false;")

                'If DbAccess.GetCount(check_Class_tmp) > 0 Then '是否有排課(Class_TmpSchedule)
                '    ch_courid = drv("CourID")
                '    is_parent = 0
                'ElseIf DbAccess.GetCount(check_schedule) > 0 Then    '是否有排課(Class_Schedule)
                '    ch_courid = drv("CourID")
                '    is_parent = 0
                'Else
                '    ch_courid = 0
                '    If Convert.IsDBNull(drv("MainCourID")) Then '為主課程
                '        Dim course_list As String = "select * from Course_CourseInfo  where MainCourID='" & drv("CourID") & "'"
                '        If DbAccess.GetCount(course_list) > 0 Then '是否有子課程
                '            is_parent = "true"
                '        Else
                '            is_parent = "false"
                '        End If
                '    Else '非主課程
                '        is_parent = "false"
                '    End If
                'End If

                'Me.ViewState("Result") = ""
                'If Convert.IsDBNull(drv("MainCourID")) Then
                '    Me.ViewState("Result") = ""
                'Else
                '    Me.ViewState("strsql_A") = "select CourseName from Course_CourseInfo   where CourID='" & drv("MainCourID") & "'"
                '    Me.ViewState("Result") = Convert.ToString(DbAccess.ExecuteScalar(Me.ViewState("strsql_A"), objconn))
                'End If

                e.Item.Cells(6).Text = Convert.ToString(drv("pCourseName")) ' Me.ViewState("Result")

                If Convert.ToString(drv("ErrorData")) = "1" Then
                    TIMS.Tooltip(e.Item.Cells(6), "主課程資料異常，請重新設定!!!")
                    e.Item.Cells(6).Text += "(主課程資料異常，請重新設定!!!)"
                    e.Item.Cells(6).ForeColor = Color.Red 'e.Item.Cells(6).ForeColor.Red
                End If

                If Convert.ToString(drv("RID")) <> Convert.ToString(RIDValue.Value) Then
                    but.Enabled = False
                    but_del.Enabled = False

                    Dim sPlanName As String = Get_PLANNAME(sm.UserInfo.LID, Convert.ToString(drv("RID")))
                    Dim sTitle As String = "非登入或搜尋的計畫權限RID，請重新增加該計畫之課程代碼!!!"
                    sTitle &= sPlanName

                    TIMS.Tooltip(but, sTitle)
                    TIMS.Tooltip(but_del, sTitle)
                    TIMS.Tooltip(e.Item.Cells(2), sTitle)
                    e.Item.Cells(2).ForeColor = Color.Red 'e.Item.ForeColor.Red
                End If
        End Select

    End Sub

    '取得計畫名稱
    Function Get_PLANNAME(ByVal LID As String, ByVal RID As String) As String
        Dim rst As String = ""
        Select Case sm.UserInfo.LID
            Case "2"
                With GsCmd1
                    .Parameters.Clear()
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RID
                    rst = Convert.ToString(.ExecuteScalar)
                End With
            Case Else
                With GsCmd1
                    .Parameters.Clear()
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RID
                    .Parameters.Add("RWPLANID", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
                    rst = Convert.ToString(.ExecuteScalar)
                End With
        End Select
        Return rst
    End Function

    '保留搜尋值
    Sub GetSearchStr()
        vsTB_CourseID = ""
        vsTB_CourseName = ""
        'If Me.TB_CourseID.Text <> "" Then vsTB_CourseID = Me.TB_CourseID.Text.Replace("'", "''")
        'If Me.TB_CourseName.Text <> "" Then vsTB_CourseName = Me.TB_CourseName.Text.Replace("'", "''")
        vsTB_CourseID = TIMS.ClearSQM(vsTB_CourseID)
        vsTB_CourseName = TIMS.ClearSQM(vsTB_CourseName)
        Dim sMySreach As String = ""
        TIMS.SetMyValue(sMySreach, "center", center.Text)
        TIMS.SetMyValue(sMySreach, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(sMySreach, "TB_CourseID", vsTB_CourseID)
        TIMS.SetMyValue(sMySreach, "TB_CourseName", vsTB_CourseName)
        TIMS.SetMyValue(sMySreach, "Classification1_List", Classification1_List.SelectedValue)
        TIMS.SetMyValue(sMySreach, "Classification2_List", Classification2_List.SelectedValue)
        TIMS.SetMyValue(sMySreach, "CB_Valid", Convert.ToString(CB_Valid.Checked))
        TIMS.SetMyValue(sMySreach, "TB_career_id", TB_career_id.Text)
        TIMS.SetMyValue(sMySreach, "Classid", Classid.Text)
        TIMS.SetMyValue(sMySreach, "Classid_Hid", Classid_Hid.Value)
        TIMS.SetMyValue(sMySreach, "trainValue", trainValue.Value)
        TIMS.SetMyValue(sMySreach, "PageIndex", CStr(DG_Course.CurrentPageIndex + 1))
        If tb_DG_Course.Visible Then
            TIMS.SetMyValue(sMySreach, "submit", "1")
        Else
            TIMS.SetMyValue(sMySreach, "submit", "0")
        End If
        Session("MySreach") = sMySreach
    End Sub

    '匯出
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        'Dim Table As DataTable

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        Dim sql As String = ""
        sql &= " select cc.CourseID" & vbCrLf
        sql &= " ,cc.CourseName" & vbCrLf
        sql &= " ,cc.Hours" & vbCrLf
        sql &= " ,cc.Classification1" & vbCrLf
        sql &= " ,cc.Classification2" & vbCrLf
        sql &= " ,mc.CourseID as MainCourseID" & vbCrLf
        sql &= " ,cc.CLSID" & vbCrLf
        'Sql += " --,case when cc.CLSID is not null then isnull(convert(varchar,ic.Years),'2009年以前舊資料') end Years" & vbCrLf
        sql &= " ,ic.Years" & vbCrLf
        sql &= " ,ic.ClassID" & vbCrLf
        sql &= " ,cc.TMID" & vbCrLf
        sql &= " ,kt.BusID" & vbCrLf
        sql &= " ,kt.TrainID" & vbCrLf
        sql &= " ,cc.Valid" & vbCrLf
        sql &= " ,cc.RID" & vbCrLf
        sql &= " ,ar.DistID" & vbCrLf
        sql &= " ,ar.PlanID" & vbCrLf
        sql &= " ,ar.OrgID" & vbCrLf
        'Sql += " --SELECT COUNT(*) CNT " & vbCrLf
        sql &= " FROM COURSE_COURSEINFO cc " & vbCrLf
        sql &= " JOIN AUTH_RELSHIP ar on ar.RID=cc.RID" & vbCrLf
        sql &= " LEFT JOIN COURSE_COURSEINFO mc on mc.CourID=cc.MainCourID and mc.RID=cc.RID" & vbCrLf
        sql &= " LEFT JOIN ID_CLASS ic on ic.CLSID=cc.CLSID AND ic.DistID =ar.DistID " & vbCrLf
        sql &= " LEFT JOIN VIEW_TRAINTYPE kt on kt.TMID = cc.TMID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        '先判斷 RIDValue.Value 
        sql &= " AND cc.RID = @RID" & vbCrLf
        sql &= " AND ar.DistID =@DISTID" & vbCrLf

        Dim pParms As New Hashtable
        pParms.Clear()
        Select Case sm.UserInfo.LID
            Case 0 '署使用-RIDValue-取得DISTID
                Dim vDISTID As String = sm.UserInfo.DistID
                vDISTID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                pParms.Add("RID", RIDValue.Value)
                pParms.Add("DISTID", vDISTID)
            Case 1 '分署
                pParms.Add("RID", RIDValue.Value)
                pParms.Add("DISTID", sm.UserInfo.DistID)
            Case 2 '委訓單位
                pParms.Add("RID", sm.UserInfo.RID)
                pParms.Add("DISTID", sm.UserInfo.DistID)
        End Select

        '卡轄區 不卡機構 卡RID
        'If sm.UserInfo.OrgID <> "" Then
        '    sql &= " AND ar.OrgID ='" & sm.UserInfo.OrgID & "'" & vbCrLf
        'End If
        '卡大計畫 分署(中心)不卡小計畫 卡RID
        'If sm.UserInfo.PlanID <> "" Then
        '    sql &= " AND ar.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf
        'End If
        'If sm.UserInfo.Years < 2010 Then
        '    sql &= " AND ic.Years IS NULL " & vbCrLf
        'Else
        '    sql &= " AND ic.Years ='" & sm.UserInfo.Years & "'" & vbCrLf
        'End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pParms)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        ExportX1(dt)
    End Sub

    Sub ExportX1(ByRef dt As DataTable)
        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("CourseInfo", System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim sFileName1 As String = "CourseInfo"
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= "<table>"
        Dim num As Integer = 0

        Dim ExportStr As String = ""    '建立輸出文字
        ExportStr = "課程代碼"
        ExportStr &= ",課程名稱"
        ExportStr &= ",小時數"
        ExportStr &= ",學/術科"
        ExportStr &= ",共同/一般/專業"

        ExportStr &= ",主課程代碼"
        ExportStr &= ",計畫年度"
        ExportStr &= ",歸屬班別代碼"
        ExportStr &= ",行業別代碼"
        ExportStr &= ",訓練職類"
        ExportStr &= ",是否有效"
        'ExportStr += vbTab & vbCrLf
        strHTML &= TIMS.Get_TABLETR(ExportStr)

        '建立資料面
        For Each dr As DataRow In dt.Rows
            num += 1
            ExportStr = dr("CourseID").ToString  '課程代碼
            ExportStr &= "," & dr("CourseName").ToString  '課程名稱
            ExportStr &= "," & dr("Hours").ToString  '小時數
            ExportStr &= "," & dr("Classification1").ToString  '學/術科
            ExportStr &= "," & dr("Classification2").ToString  '共同/一般/專業
            ExportStr &= "," & dr("MainCourseID").ToString  '主課程代碼
            ExportStr &= "," & dr("Years").ToString  '計畫年度
            ExportStr &= "," & dr("ClassID").ToString  '歸屬班別代碼
            ExportStr &= "," & dr("BusID").ToString  '行業別代碼
            ExportStr &= "," & dr("TrainID").ToString  '訓練職類
            ExportStr &= "," & dr("Valid").ToString  '是否有效
            strHTML &= TIMS.Get_TABLETR(ExportStr)
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

End Class

