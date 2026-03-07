Imports System.Threading

Partial Class SD_04_008
    Inherits AuthBasePage

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
        PageControler2.PageDataGrid = DataGrid2

        msg.Text = ""

        If Not IsPostBack Then
            DataGridTable.Visible = False
            DataGridTable2.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            center2.Text = sm.UserInfo.OrgName
            RIDValue2.Value = sm.UserInfo.RID
            CreateItem()
            Page1.Visible = True
            Page2.Visible = False

            If sm.UserInfo.LID > 1 Then
                KindEngage.AutoPostBack = False
            Else
                KindEngage.AutoPostBack = True
            End If
        End If

        If TechID.Value <> "" Then
            Dim MyArray As Array = Split(TechID.Value, ",")
            For i As Integer = 0 To MyArray.Length - 1
                If Not ListBox1.Items.FindByValue(MyArray(i)) Is Nothing Then
                    ListBox2.Items.Add(ListBox1.Items.FindByValue(MyArray(i)))
                    ListBox1.Items.Remove(ListBox1.Items.FindByValue(MyArray(i)))
                End If
            Next
        End If

        Button2.Attributes("onclick") = "return CheckData();"
        Button4.Attributes("onclick") = "ChangeItem(1);"
        Button5.Attributes("onclick") = "ChangeItem(2);"
        Button7.Attributes("onclick") = "return CheckData2();"
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button3.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button3.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button9.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?RIDField=RIDValue2&OrgField=center2');"
        Else
            Button9.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx?RIDField=RIDValue2&OrgField=center2');"
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList1", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList1');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID2, "HistoryList2", "RIDValue2", "center2")
        If HistoryRID2.Rows.Count <> 0 Then
            center2.Attributes("onclick") = "showObj('HistoryList2');"
            center2.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub CreateItem()
        Dim sql As String
        Dim dt As DataTable

        sql = "select * from ID_Invest ORDER BY IVID"
        dt = DbAccess.GetDataTable(sql, objconn)
        With IVID
            .DataSource = dt
            .DataTextField = "InvestName"
            .DataValueField = "IVID"
            .DataBind()
            .Items.Insert(0, New ListItem("不區分", ""))
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        Dim SearchStr As String = ""

        TechID.Value = ""

        If RIDValue.Value <> "" Then
            SearchStr += " and RID='" & RIDValue.Value & "'"
        End If
        If TeachName.Text <> "" Then
            SearchStr += " and (TeachName like '%" & Replace(TeachName.Text, " ", "%") & "%' or TeachEName like '%" & Replace(TeachName.Text, " ", "%") & "%')"
        End If
        If IDNO.Text <> "" Then
            SearchStr += " and IDNO='" & IDNO.Text & "'"
        End If
        If TeacherID.Text <> "" Then
            SearchStr += " and TeacherID like '%" & Replace(TeacherID.Text, " ", "%") & "%'"
        End If
        If IVID.SelectedIndex <> 0 Then
            SearchStr += " and IVID='" & IVID.SelectedValue & "'"
        End If
        If trainValue.Value <> "" Then
            SearchStr += " and TMID='" & trainValue.Value & "'"
        End If
        SearchStr += " and WorkStatus='" & WorkStatus.SelectedValue & "'"
        If KindEngage.SelectedIndex <> 0 Then
            SearchStr += " and KindEngage='" & KindEngage.SelectedValue & "'"
        End If
        If KindID.SelectedIndex <> 0 Then
            SearchStr += " and KindID='" & KindID.SelectedValue & "'"
        End If

        sql = "SELECT * FROM Teach_TeacherInfo WHERE 1=1" & SearchStr
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            DataGridTable.Visible = False
            Common.MessageBox(Me, "查無資料")
        Else
            DataGridTable.Visible = True
            With ListBox1
                .DataSource = dt
                .DataTextField = "TeachCName"
                .DataValueField = "TechID"
                .DataBind()
            End With
            ListBox2.Items.Clear()
        End If
    End Sub

    Dim ProcessID As String

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'ViewCourse()
        Dim myThreadDelegate As New ThreadStart(AddressOf ViewCourse)
        Dim myThread As New Thread(myThreadDelegate)

        ProcessID = TIMS.GetRnd16Eng()
        TIMS.InsertLog(ProcessID, 0, "執行緒啟動中")
        Common.RespWrite(Me, "<script>alert('程式處理中,將開啟小視窗提是目前的進度!\n請勿重複發送查詢,以免主機大量負荷!');window.open('SD_04_008_c.aspx?ID=" & Request("ID") & "&ProcessID=" & ProcessID & "','Process','width=180,height=120,resizable=0,scrollbars=0,status=0')</script>")

        myThread.Start()
    End Sub

    Public Sub ViewCourse()
        Dim sql As String
        Dim dt1 As DataTable
        Dim dr1 As DataRow
        Dim dt2 As DataTable
        Dim dr2 As DataRow
        Dim dt3 As DataTable
        Dim dr3 As DataRow
        Dim dt4 As DataTable
        Dim dr4 As DataRow
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow
        'Dim IDNO As String

        Dim RunCount As Integer
        '2006/03/ add conn by matt
        Dim conn As SqlConnection
        conn = DbAccess.GetConnection

        Try
            sql = "DELETE Class_TeachDupResult WHERE TechID IN (" & Me.TechID.Value & ") and DupDate>='" & Start_Date.Text & "' and DupDate<='" & End_Date.Text & "'"
            DbAccess.ExecuteNonQuery(sql, objconn)

            'sql = "SELECT TechID,IDNO FROM Teach_TeacherInfo WHERE IDNO IN (SELECT IDNO FROM Teach_TeacherInfo WHERE TechID IN (" & Me.TechID.Value & ")) Order By IDNO"

            sql = "SELECT TechID,IDNO FROM Teach_TeacherInfo WHERE TechID IN (" & Me.TechID.Value & ") Order By IDNO"
            dt1 = DbAccess.GetDataTable(sql, objconn)
            For Each dr1 In dt1.Rows
                sql = "SELECT * FROM Class_TeachDupResult WHERE 1<>1"
                dt = DbAccess.GetDataTable(sql, da, objconn)

                RunCount += 1

                sql = "SELECT a.*,b.CyclType,b.ClassCName,d.OrgName,e.PlanName FROM "
                sql += "(SELECT * FROM Class_Schedule WHERE SchoolDate IN (SELECT SchoolDate FROM Class_Schedule WHERE SchoolDate>='" & Start_Date.Text & "' and SchoolDate<='" & End_Date.Text & "' Group By SchoolDate Having Count(SchoolDate)>1) and OCID IN (SELECT OCID FROM Class_ClassInfo WHERE RID IN (SELECT RID FROM Auth_Relship WHERE OrgID =(SELECT OrgID FROM Auth_Relship WHERE RID='" & RIDValue.Value & "')))) a "
                sql += "JOIN Class_ClassInfo b ON a.OCID=b.OCID "
                sql += "JOIN Auth_Relship c ON b.RID=c.RID "
                sql += "JOIN Org_OrgInfo d ON c.OrgID=d.OrgID "
                sql += "JOIN view_LoginPlan e ON b.PlanID=e.PlanID "
                sql += "Order By SchoolDate "
                dt2 = DbAccess.GetDataTable(sql, objconn)
                dt3 = DbAccess.GetDataTable(sql, objconn)

                For Each dr2 In dt2.Rows
                    For i As Integer = 1 To 24
                        If IsNumeric(dr2("Teacher" & i)) Then
                            If dr2("Teacher" & i) = dr1("TechID") Then
                                For Each dr3 In dt3.Select("SchoolDate='" & dr2("SchoolDate") & "' and CSID<>'" & dr2("CSID") & "'")
                                    If IsNumeric(dr3("Teacher" & i)) Then
                                        sql = "SELECT TechID,IDNO FROM Teach_TeacherInfo WHERE IDNO IN (SELECT IDNO FROM Teach_TeacherInfo WHERE TechID ='" & dr1("TechID") & "') Order By IDNO"
                                        dt4 = DbAccess.GetDataTable(sql, objconn)
                                        For Each dr4 In dt4.Rows
                                            If dr3("Teacher" & i) = dr4("TechID") Then
                                                If dt.Select("TechID='" & dr1("TechID") & "' and DupDate='" & dr2("SchoolDate") & "' and DupPart='" & IIf(i Mod 12 = 12, 12, i Mod 12) & "'").Length = 0 Then
                                                    sql = "SELECT * FROM Class_TeachDupResult WHERE TechID='" & dr1("TechID") & "' and DupDate='" & dr2("SchoolDate") & "' and DupPart='" & IIf(i Mod 12 = 12, 12, i Mod 12) & "'"
                                                    dr = DbAccess.GetOneRow(sql, objconn)
                                                    If dr Is Nothing Then
                                                        dr = dt.NewRow
                                                        dt.Rows.Add(dr)
                                                        dr("DupDate") = dr2("SchoolDate")
                                                        dr("TechID") = dr1("TechID")
                                                        dr("DupPart") = IIf(i Mod 12 = 12, 12, i Mod 12)

                                                        Dim DupDesc As String = ""
                                                        For Each dr5 As DataRow In dt4.Rows
                                                            For Each dr6 As DataRow In dt2.Select("SchoolDate='" & dr2("SchoolDate") & "' and Teacher" & IIf(i Mod 12 = 12, 12, i Mod 12) & "='" & dr5("TechID") & "'")
                                                                If DupDesc <> "" Then
                                                                    DupDesc += vbCrLf
                                                                End If

                                                                DupDesc += dr6("PlanName") & "-" & dr6("OrgName") & "-" & dr6("ClassCName")
                                                                If IsNumeric(dr6("CyclType")) Then
                                                                    If Int(dr6("CyclType")) <> 0 Then
                                                                        DupDesc += "第" & Int(dr6("CyclType")) & "期"
                                                                    End If
                                                                End If
                                                            Next
                                                        Next

                                                        dr("DupDesc") = DupDesc
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        End If
                    Next
                Next

                TIMS.InsertLog(ProcessID, RunCount * 100 / dt1.Rows.Count, "總計" & dt1.Rows.Count & "人,目前已經處理" & RunCount & "人")

                DbAccess.UpdateDataTable(dt, da)
            Next
            TIMS.InsertLog(ProcessID, 100, "執行緒已經結束")
        Catch ex As Exception
            TIMS.InsertLog(ProcessID, -1, "執行緒發生意外,程式結束,錯誤代碼:" & vbCrLf & ex.ToString)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Page1.Visible = False
        Page2.Visible = True

        center2.Text = sm.UserInfo.OrgName
        RIDValue2.Value = sm.UserInfo.RID
        TeacherID2.Text = ""
        IDNO2.Text = ""
        TeachCName2.Text = ""
        Start_Date2.Text = ""
        End_Date2.Text = ""
        DataGridTable2.Visible = False
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Page1.Visible = True
        Page2.Visible = False
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim sql As String = ""
        Dim SearchStr As String = ""
        Dim SearchStr2 As String = ""

        If TeachCName2.Text <> "" Then
            SearchStr += " and (TeachCName like '%" & Replace(TeachCName2.Text, " ", "%") & "%' or TeachEName like '%" & Replace(TeachCName2.Text, " ", "%") & "%')"
        End If
        If IDNO2.Text <> "" Then
            SearchStr += " and IDNO='" & IDNO2.Text & "'"
        End If
        If TeacherID2.Text <> "" Then
            SearchStr += " and TeacherID='" & TeacherID2.Text & "'"
        End If
        SearchStr += " and RID IN (SELECT RID FROM Auth_Relship WHERE OrgID=(SELECT OrgID FROM Auth_Relship WHERE RID='" & RIDValue2.Value & "'))"

        If Start_Date2.Text <> "" Then
            SearchStr2 += " and DupDate>='" & Start_Date2.Text & "'"
        End If
        If End_Date2.Text <> "" Then
            SearchStr2 += " and DupDate<='" & End_Date2.Text & "'"
        End If

        sql = "SELECT a.TDRID,a.DupDate,a.DupPart,a.DupDesc,b.TeacherID,b.TeachCName FROM "
        sql += "(SELECT * FROM Class_TeachDupResult WHERE 1=1" & SearchStr2 & " and TechID IN (SELECT TechID FROM Teach_TeacherInfo WHERE 1=1" & SearchStr & ")) a "
        sql += "JOIN Teach_TeacherInfo b ON a.TechID=b.TechID "
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGridTable2.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable2.Visible = True
            msg.Text = ""

            PageControler2.PageDataTable = dt
            PageControler2.PrimaryKey = "TDRID"
            PageControler2.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DataGrid2.CurrentPageIndex * DataGrid2.PageSize
                e.Item.Cells(5).Text = Replace(drv("DupDesc"), vbCrLf, "<BR>")
        End Select
    End Sub

    Private Sub KindEngage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KindEngage.SelectedIndexChanged
        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow

        sql = "SELECT * FROM ID_KindOfTeacher WHERE KindEngage='" & KindEngage.SelectedValue & "'"
        dt = DbAccess.GetDataTable(sql, objconn)

        With KindID
            .DataSource = dt
            .DataTextField = "KindName"
            .DataValueField = "KindID"
            .DataBind()
            .Items.Insert(0, New ListItem("不區分", ""))
        End With
    End Sub
End Class
