Partial Class SD_04_014
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
#Region "在這裡放置使用者程式碼以初始化網頁"
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            RIDValue.Value = sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName
            IOrg.Text = sm.UserInfo.OrgName
            Me.ViewState("Type") = "Search"
            KindID = TIMS.Get_KindOfTeacher(KindID, 2, "", objconn)
            IVID = TIMS.Get_Invest(IVID, objconn)
            msg.Text = ""
            TableDataGrid1.Style("display") = "none"
            PageControler1.Visible = False
            'Tsearch.Style("display") = "inline"
            Tsearch.Style("display") = ""
            TInsert.Style("display") = "none"
        End If

        'Save.Attributes("onclick") = "search()"
        'IteachName.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');" '''Class TIMS.CreateTeacherScript
        IteachName.Attributes("ondblclick") = "Get_Teah('IteachName','TeahValue','TMID');"
        IteachName.Style.Item("CURSOR") = "hand"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Style("CURSOR") = "hand"
        End If
#End Region
    End Sub

    Private Sub Search1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Search1.Click
        search()
    End Sub

    Sub search()
        Dim sql As String = ""
        sql = ""
        sql += " SELECT te.TEID, oo.orgName, tt.TeacherID, tt.TeachCName, tt.IDNO, ik.KindName as KindID, " & vbCrLf
        sql += " tt.TMID,ty.JobID,ty.JobName,ty.TrainID,ty.TrainName, te.TEDate1,te.TEDate2 " & vbCrLf
        sql += " ,CONVERT(varchar, te.TEDate1, 111) + '~' + CONVERT(varchar, te.TEDate2, 111) as TEDate ,tt.RID " & vbCrLf
        sql += " FROM Teach_TeacherInfo tt " & vbCrLf
        sql += " JOIN Auth_Relship ar ON tt.RID = ar.RID " & vbCrLf
        sql += " JOIN Org_OrgInfo oo ON oo.orgID = ar.orgID " & vbCrLf
        sql += " JOIN Teacher_Employs te ON tt.TechID = te.TechID " & vbCrLf
        sql += " LEFT join ID_KindOfTeacher ik ON ik.KindID = tt.KindID " & vbCrLf
        sql += " LEFT join view_TrainType ty ON tt.TMID = ty.TMID " & vbCrLf
        sql += " WHERE tt.KindEngage = 2 " & vbCrLf
        sql += " AND tt.WorkStatus = 1 " & vbCrLf

        Select Case Convert.ToString(Me.ViewState("Type"))
            Case "ADD"
                If RIDValue.Value <> "" Then sql += " AND tt.RID = '" & RIDValue.Value & "' "
                If TeahValue.Value <> "" Then sql += " AND tt.TechID = '" & TeahValue.Value & "' "
            Case Else ' "Search"
                If KindID.SelectedIndex <> 0 Then sql += " AND tt.KindID = '" & KindID.SelectedValue & "' "
                If TeachCName.Text <> "" Then sql += " AND tt.TeachCName LIKE '%" & TeachCName.Text & "%' "
                If IDNO.Text <> "" Then sql += " AND tt.IDNO = '" & IDNO.Text & "' "
                If TeacherID.Text <> "" Then sql += " AND tt.TeacherID = '" & TeacherID.Text & "' "
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If jobValue.Value <> "" Then
                        sql += " AND tt.TMID = '" & jobValue.Value & "' "
                    ElseIf trainValue.Value <> "" Then
                        sql += " AND tt.TMID = '" & trainValue.Value & "' "
                    End If
                Else
                    If trainValue.Value <> "" Then sql += " AND tt.TMID = '" & trainValue.Value & "' "
                End If
                If IVID.SelectedIndex <> 0 Then sql += " AND tt.IVID = '" & IVID.SelectedValue & "' "
                If start_date.Text <> "" Then sql += " AND te.TEDate1 >= " & TIMS.To_date(start_date.Text)
                If end_date.Text <> "" Then sql += " AND te.TEDate2 <= " & TIMS.To_date(end_date.Text)
        End Select

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--##SD_04_014.aspx ,sql:", vbCrLf, sql))
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!!!"
        TableDataGrid1.Style("display") = "none"
        PageControler1.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            TableDataGrid1.Style("display") = ""
            PageControler1.Visible = True

            If Me.ViewState("Type") = "Search" Then
                Tsearch.Style("display") = "inline"
                TInsert.Style("display") = "none"
            Else
                Tsearch.Style("display") = "none"
            End If

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "TEID"
            PageControler1.Sort = "TeacherID"
            PageControler1.ControlerLoad()

#Region "(No Use)"

            'PageControler1.SqlString = sql
            'PageControler1.PrimaryKey = "TEID"
            'PageControler1.Sort = "TeacherID"
            'PageControler1.ControlerLoad()

#End Region
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
#Region "(No Use)"

        'Dim sql2 As String
        'Dim dr As DataRow
        'Dim Array1 As Array

#End Region
        Dim vsTEID As String = ""
        Dim vsTMID As String = ""
        Dim vsTeachCName As String = ""
        Dim vsTEDate1 As String = ""
        Dim vsTEDate2 As String = ""
        Dim vsRID As String = ""
        Dim vsOrgName As String = ""
        Dim sql As String = ""

        Select Case e.CommandName
            Case "Edit"
                vsTEID = TIMS.GetMyValue(e.CommandArgument, "TEID")
                vsTMID = TIMS.GetMyValue(e.CommandArgument, "TMID")
                vsTeachCName = TIMS.GetMyValue(e.CommandArgument, "TeachCName")
                vsTEDate1 = TIMS.GetMyValue(e.CommandArgument, "TEDate1")
                vsTEDate2 = TIMS.GetMyValue(e.CommandArgument, "TEDate2")
                vsRID = TIMS.GetMyValue(e.CommandArgument, "RID")
                vsOrgName = TIMS.GetMyValue(e.CommandArgument, "OrgName")
                'Array1 = Split(e.CommandArgument, ",")
                Me.ViewState("Type") = "Edit"
                Me.ViewState("TEID") = vsTEID
                'TInsert.Style("display") = "inline"
                TInsert.Style("display") = ""   '==== by:20180824
                IteachName.Style("display") = "none" '新增
                'IteachName2.Style("display") = "inline" '修改
                IteachName2.Style("display") = ""   '==== by:20180824
                Tsearch.Style("display") = "none"
                msg.Text = ""
                TableDataGrid1.Style("display") = "none"
                PageControler1.Visible = False
                RIDValue.Value = vsRID
                IOrg.Text = vsOrgName
                TMID.Text = vsTMID
                IteachName2.Text = vsTeachCName
                IDate1.Text = vsTEDate1
                IDate2.Text = vsTEDate2
            Case "Del"
                vsTEID = TIMS.GetMyValue(e.CommandArgument, "TEID")
                sql = "Delete Teacher_Employs where TEID = " & vsTEID
                DbAccess.ExecuteNonQuery(sql)
                Common.MessageBox(Me, "刪除成功!!!")
                search()
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.EditItem, ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LTMID As Label = e.Item.FindControl("LTMID")
                Dim BtnEdit As Button = e.Item.FindControl("Edit")
                Dim BtnDel As Button = e.Item.FindControl("Del")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                LTMID.Text = ""
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If drv("TMID").ToString <> "" Then LTMID.Text = "[" & drv("JobID").ToString & "]" & drv("JobName").ToString
                Else
                    If drv("TMID").ToString <> "" Then LTMID.Text = "[" & drv("TrainID").ToString & "]" & drv("TrainName").ToString
                End If

                Dim vsArg As String = ""
                vsArg = ""
                vsArg += "&TEID=" & drv("TEID")
                vsArg += "&TMID=" & LTMID.Text
                vsArg += "&TeachCName=" & drv("TeachCName")
                vsArg += "&TEDate1=" & drv("TEDate1")
                vsArg += "&TEDate2=" & drv("TEDate2")
                vsArg += "&RID=" & drv("RID")
                vsArg += "&OrgName=" & drv("OrgName")

                BtnEdit.CommandArgument = vsArg
                BtnDel.CommandArgument = vsArg

                If IsDBNull(drv("TEDate2")) = False Then
                    Dim myNowTime As String = DateTime.Now.ToString("yyyy/MM/dd")  '=== by:20180824
                    Dim myTEdate2 As String = Convert.ToDateTime(drv("TEDate2")).ToString("yyyy/MM/dd")  '=== by:20180824

                    'If Now() > drv("TEDate2") Then
                    If myNowTime > myTEdate2 Then  '=== by:20180824
                        BtnEdit.Enabled = False
                        BtnEdit.Visible = False  '=== by:20180824
                    Else
                        BtnEdit.Enabled = True
                        BtnEdit.Visible = True   '=== by:20180824
                    End If
                End If
        End Select
    End Sub

    Private Sub Add2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Add2.Click
        TInsert.Style("display") = "inline"
        'IteachName.Style("display") = "inline" '新增
        IteachName.Style("display") = ""   '==== by:20180824
        IteachName2.Style("display") = "none"
        Tsearch.Style("display") = "none"
        msg.Text = ""
        TableDataGrid1.Style("display") = "none"
        PageControler1.Visible = False
        RIDValue.Value = sm.UserInfo.RID
        IOrg.Text = sm.UserInfo.OrgName
        IteachName.Text = ""
        IteachName2.Text = ""
        TMID.Text = ""
        TeahValue.Value = ""
        IDate1.Text = ""
        IDate2.Text = ""
        Me.ViewState("Type") = "ADD"
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click
        'comm = DbAccess.GetConnection()
        'comm.Open()

        'Dim sql3 As String = ""
        'Dim strScript As String = ""
        Dim msg As String = ""
        If Me.ViewState("Type") = "ADD" Then
            If TeahValue.Value = "" Then msg += "請選擇講師姓名!!!" & vbCrLf
        End If
        If IDate1.Text = "" Then msg += "輸入[約聘期限起日]!!!" & vbCrLf
        If IDate2.Text = "" Then msg += "輸入[約聘期限迄日]!!!" & vbCrLf
        If IDate1.Text <> "" And IDate2.Text <> "" Then
            If CDate(IDate1.Text) >= CDate(IDate2.Text) Then msg += "[聘約期間起日]不能大於[聘約期間迄日]或等於[聘約期間迄日]!!!" & vbCrLf
        End If
        '======================================================= 取得講師的師資別 20180904
        Dim tmpKindEngage As Integer = 0
        If TeahValue.Value.Length > 0 Then
            Dim tSql As String = " SELECT ISNULL(CONVERT(INT, KINDENGAGE), 0) AS tmpKindEngage FROM TEACH_TEACHERINFO WHERE TECHID = @TECHID "
            Dim tParam As Hashtable = New Hashtable
            tParam.Add("TECHID", TeahValue.Value.Trim)
            Call TIMS.OpenDbConn(objconn)
            Dim dr As DataRow = DbAccess.GetOneRow(tSql, objconn, tParam)
            tmpKindEngage = Convert.ToInt32(dr("tmpKindEngage"))
        End If
        If tmpKindEngage = 0 Or tmpKindEngage = 1 Then msg += "講師的[師資別]必須是屬於[外聘]!!!" & vbCrLf
        '=======================================================

        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        Select Case Convert.ToString(Me.ViewState("Type"))
            Case "ADD"
                'Dim sql2 As String = ""
                'sql2 = ""
                'sql2 &= "select max(TEID)+1 as TEID from Teacher_Employs"
                Dim sql3 As String = ""
                sql3 = ""
                sql3 &= " INSERT INTO Teacher_Employs(TEID,TechID,TEDate1,TEDate2,ModifyAcct,ModifyDate) " & vbCrLf
                sql3 += " VALUES((select MAX(TEID)+1 AS TEID FROM Teacher_Employs),@TechID,@TEDate1,@TEDate2,@ModifyAcct,GETDATE()) " & vbCrLf
                Call TIMS.OpenDbConn(objconn)
                'Dim sql As New SqlCommand(sql2, objconn)
                Dim sqlcom As New SqlCommand(sql3, objconn)
                'sqlcom = New SqlCommand(sql3, objconn)
                With sqlcom
                    .Parameters.Clear()
                    .Parameters.Add("TechID", SqlDbType.Int).Value = Val(TeahValue.Value)
                    .Parameters.Add("TEDate1", SqlDbType.DateTime).Value = TIMS.Cdate2(IDate1.Text)
                    .Parameters.Add("TEDate2", SqlDbType.DateTime).Value = TIMS.Cdate2(IDate2.Text)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(sqlcom.CommandText, objconn, sqlcom.Parameters)
                End With
                search()
            Case "Edit"
                Dim sql3 As String = ""
                sql3 = ""
                sql3 &= " UPDATE Teacher_Employs " & vbCrLf
                sql3 += " SET TEDate1 = @TEDate1 ,TEDate2 = @TEDate2 ,ModifyAcct = @ModifyAcct ,ModifyDate = GETDATE() " & vbCrLf
                sql3 += " WHERE TEID = @TEID " '& Me.ViewState("TEID")
                Call TIMS.OpenDbConn(objconn)
                Dim sqlcom As New SqlCommand(sql3, objconn)
                'sqlcom = New SqlCommand(sql3, objconn)
                With sqlcom
                    .Parameters.Clear()
                    .Parameters.Add("TEDate1", SqlDbType.DateTime).Value = TIMS.Cdate2(IDate1.Text)
                    .Parameters.Add("TEDate2", SqlDbType.DateTime).Value = TIMS.Cdate2(IDate2.Text)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("TEID", SqlDbType.VarChar).Value = Me.ViewState("TEID")
                    '.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(sqlcom.CommandText, objconn, sqlcom.Parameters)
                End With
                Me.ViewState("Type") = "Search"
                search()
        End Select
    End Sub

    Private Sub ReFist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReFist.Click
        TInsert.Style("display") = "none"
        IteachName.Style("display") = "none"
        IteachName2.Style("display") = "none"
        'Tsearch.Style("display") = "inline"
        Tsearch.Style("display") = ""   '==== by:20180824
        IteachName.Text = ""
        IteachName2.Text = ""
        TeahValue.Value = ""
        IDate1.Text = ""
        IDate2.Text = ""
        Me.ViewState("Type") = "Search"
        search()
    End Sub
End Class