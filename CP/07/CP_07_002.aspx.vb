Partial Class CP_07_002
    Inherits AuthBasePage

    'CP_07_002_blank,'CP_07_002

    'select top 10 * from Stud_QuesTraining
    'select top 10 * from Stud_ForumRecord

    'Dim strRid As String = ""
    'Dim strOCID As String = ""
    'Dim strSOCID As String = ""
    'Dim isloaded As Boolean
    'Dim RelshipTable As DataTable
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
        '檢查Session是否存在 End 'PageControler1 = FindControl("PageControler1")
        PageControler1.PageDataGrid = DG_ClassInfo

        If Not IsPostBack Then
            PageControler1.Visible = False
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'isloaded = False
            bt_search.Attributes.Add("onclick", "return chkOrg();")
        End If

        'If strRid <> "" Then RIDValue.Value = strRid
        'If strOCID <> "" Then OCIDValue1.Value = strOCID

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "bt_search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        If Not IsPostBack Then
            'TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            years.Value = sm.UserInfo.Years
            PlanID.Value = sm.UserInfo.PlanID

            If Session(cst_pSearchStr) IsNot Nothing Then
                Dim strSession As String = Session(cst_pSearchStr)
                Dim MyValue As String = ""
                center.Text = TIMS.GetMyValue(strSession, "center")
                RIDValue.Value = TIMS.GetMyValue(strSession, "RIDValue")
                TMID1.Text = TIMS.GetMyValue(strSession, "TMID1")
                OCID1.Text = TIMS.GetMyValue(strSession, "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(strSession, "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(strSession, "OCIDValue1")
                STDate1.Text = TIMS.GetMyValue(strSession, "start_date")
                STDate2.Text = TIMS.GetMyValue(strSession, "end_date")
                MyValue = TIMS.GetMyValue(strSession, "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If
                Session(cst_pSearchStr) = Nothing
                Call sUtl_Search1()
                'bt_search_Click(sender, e)
                'Session(SearchStr) = Nothing
            End If
        End If
        'CP_07_002_blank
        'CP_07_002
        bt_blankRpt.Attributes.Add("onclick", "return chkprint(1);")
        bt_PrintRpt.Attributes.Add("onclick", "return chkprint(2);")

        'If strRid <> "" And strOCID <> "" And strSOCID <> "" And isloaded = False Then
        '    loadData()
        '    isloaded = True
        'End If
        'strRid = ""
        'strOCID = ""
        'strSOCID = ""
    End Sub

    '查詢
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        sUtl_Search1()
    End Sub

    Private Sub sUtl_Search1()
        If RIDValue.Value = "" Then Exit Sub
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_ClassInfo) '顯示列數不正確

        Dim sql As String = ""
        sql &= " select a.QaySDate,a.QayFDate, a.PlanID, a.OCID" & vbCrLf
        sql &= " ,CONVERT(varchar, a.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(varchar, a.FTDate, 111) FTDate" & vbCrLf
        sql &= " ,a.RID ,upper(b.StudentID) StudentID " & vbCrLf
        sql &= " ,c.Name,e.OrgName,b.socid" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) ClassCName" & vbCrLf
        sql &= " ,CONVERT(varchar, g.CreateDate, 111) CreateDate" & vbCrLf
        sql &= " ,f23.OrgName2" & vbCrLf
        sql &= " from Class_ClassInfo a" & vbCrLf
        sql &= " JOIN ID_Plan h on a.PlanID=h.PlanID" & vbCrLf
        sql &= " join Class_StudentsOfClass b on a.OCID=b.OCID" & vbCrLf
        sql &= " join Stud_StudentInfo c on b.SID=c.SID" & vbCrLf
        sql &= " join Auth_Relship f on a.RID=f.RID" & vbCrLf
        sql &= " join Org_OrgInfo e on f.OrgID=e.Orgid" & vbCrLf
        sql &= " left join MVIEW_RELSHIP23 f23 on f23.RID3=f.RID" & vbCrLf
        sql &= " left join Stud_ForumRecord g on b.socid=g.socid" & vbCrLf
        sql &= " where h.Years='" & sm.UserInfo.Years & "' " & vbCrLf
        sql &= " and a.IsSuccess='Y' " & vbCrLf
        sql &= " and h.TPlanID<>'28' " '非產投計畫
        If sm.UserInfo.RID = "A" Then
            sql &= " and  a.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Years='" & sm.UserInfo.Years & "') " & vbCrLf
        Else
            sql &= " and  a.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        If StudentID.Text <> "" Then StudentID.Text = Trim(StudentID.Text)
        If StudentID.Text <> "" Then
            sql &= " and  b.StudentID like '%'+@StudentID+'%'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            If OCIDValue1.Value = "" Then
                sql &= " and a.RID like '" & RIDValue.Value & "%'"
            End If
        End If
        If OCIDValue1.Value <> "" Then
            sql &= " and a.OCID='" & OCIDValue1.Value & "' "
        End If

        If STDate1.Text <> "" Then
            sql &= " and a.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql &= " and a.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'"
        End If
        sql &= " order by a.OCID "
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            If StudentID.Text <> "" Then
                .Parameters.Add("StudentID", SqlDbType.VarChar).Value = StudentID.Text
            End If
            dt.Load(.ExecuteReader())
        End With

        Table4.Visible = False
        'Panel1.Visible = False
        DG_ClassInfo.Visible = False
        PageControler1.Visible = False

        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無資料!!")
            Return
        End If

        Table4.Visible = True
        'Panel1.Visible = True
        DG_ClassInfo.Visible = True
        PageControler1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Sub GetSearchStr()
        Dim strSession As String = ""
        TIMS.SetMyValue(strSession, "center", center.Text)
        TIMS.SetMyValue(strSession, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(strSession, "TMID1", TMID1.Text)
        TIMS.SetMyValue(strSession, "OCID1", OCID1.Text)
        TIMS.SetMyValue(strSession, "TMIDValue1", TMIDValue1.Value)
        TIMS.SetMyValue(strSession, "OCIDValue1", OCIDValue1.Value)
        TIMS.SetMyValue(strSession, "start_date", STDate1.Text)
        TIMS.SetMyValue(strSession, "end_date", STDate2.Text)
        TIMS.SetMyValue(strSession, "PageIndex", DG_ClassInfo.CurrentPageIndex + 1)
        Session(cst_pSearchStr) = strSession
    End Sub

    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand
        Select Case e.CommandName
            Case "edit"
                GetSearchStr()
                Dim url1 As String = String.Concat("CP_07_002_add.aspx?ID=", TIMS.Get_MRqID(Me), e.CommandArgument)
                TIMS.Utl_Redirect1(Me, url1)
                'Common.RespWrite(Me, "<script language=javascript> document.location.href='" & e.CommandArgument & "';</script>")
        End Select
    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        'Const cst_管控單位 As Integer = 3
        Const cst_訓練機構 As Integer = 4
        Const cst_填寫日期 As Integer = 8

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim CB_SOCID As HtmlInputCheckBox = e.Item.FindControl("CB_SOCID")
                Dim btnedit As Button = e.Item.FindControl("btnedit")
                Dim lOrgName2 As Label = e.Item.FindControl("lOrgName2")
                Dim lQaySDate As Label = e.Item.FindControl("lQaySDate")
                Dim lQayFDate As Label = e.Item.FindControl("lQayFDate")

                CB_SOCID.Value = drv("SOCID")

                If drv("RID").ToString.Length <> 1 Then
                    If Convert.ToString(drv("OrgName2")) <> "" Then
                        lOrgName2.Text = Convert.ToString(drv("OrgName2"))
                    End If

                    'If RelshipTable.Select("RID='" & drv("RID") & "'").Length <> 0 Then
                    '    Dim Relship As String
                    '    Dim Parent As String
                    '    Relship = RelshipTable.Select("RID='" & drv("RID") & "'")(0)("Relship")
                    '    Parent = Split(Relship, "/")(Split(Relship, "/").Length - 3)
                    '    If RelshipTable.Select("RID='" & Parent & "'").Length <> 0 Then
                    '        e.Item.Cells(4).Text = RelshipTable.Select("RID='" & Parent & "'")(0)("OrgName")
                    '    End If
                    'End If
                End If
                Dim cmdArg As String = ""
                cmdArg &= "&rid=" & drv("rid")
                cmdArg &= "&ocid=" & drv("ocid")
                cmdArg &= "&socid=" & drv("socid")
                cmdArg &= "&PlanID=" & drv("PlanID")
                '填寫日期
                If IsDBNull(drv("CreateDate")) Then
                    'btnedit.CommandName = "add"
                    btnedit.Text = "新增"
                    e.Item.Cells(cst_填寫日期).Text = "尚未填寫"
                    cmdArg &= "&status=add"
                    'edit.CommandArgument = "CP_07_002_add.aspx?ocid=" & drv("ocid") & "&socid=" & drv("socid") & "&rid=" & drv("rid") & "&PlanID=" & drv("PlanID") & "&status=add"
                Else
                    btnedit.Text = "修改"
                    cmdArg &= "&status=edit"
                    'edit.CommandArgument = "CP_07_002_add.aspx?ocid=" & drv("ocid") & "&socid=" & drv("socid") & "&rid=" & drv("rid") & "&PlanID=" & drv("PlanID") & "&status=edit"
                End If
                btnedit.CommandArgument = cmdArg

                '問卷期間起日
                If drv("QaySDate").ToString <> "" Then lQaySDate.Text = Common.FormatDate(drv("QaySDate").ToString)
                '問卷期間迄日
                If drv("QayFDate").ToString <> "" Then lQayFDate.Text = Common.FormatDate(drv("QayFDate").ToString)

                btnedit.Enabled = False
                If (drv("QaySDate").ToString <> "") AndAlso (drv("QayFDate").ToString <> "") Then
                    If (CDate(drv("QaySDate")) <= CDate(Common.FormatDate(Now()))) AndAlso (CDate(drv("QayFDate")) >= CDate(Common.FormatDate(Now()))) Then
                        btnedit.Enabled = True
                    End If
                End If

                If TIMS.sUtl_ChkTest() Then
                    If btnedit.Enabled = False Then btnedit.Enabled = True '測試用
                End If

            Case ListItemType.Header
                If Not ViewState("sort") Is Nothing Then
                    Dim img As New UI.WebControls.Image
                    Dim i As Integer
                    Select Case ViewState("sort")
                        Case "OrgName", "OrgName desc"
                            i = cst_訓練機構
                    End Select

                    If ViewState("sort").ToString.IndexOf("desc") = -1 Then
                        img.ImageUrl = "../../images/SortUp.gif"
                    Else
                        img.ImageUrl = "../../images/SortDown.gif"
                    End If
                    e.Item.Cells(i).Controls.Add(img)
                End If

        End Select

    End Sub

    Private Sub DG_ClassInfo_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_ClassInfo.SortCommand
        Dim strADORN As String = If(e.SortExpression = ViewState("sort"), " desc", "")
        ViewState("sort") = $"{e.SortExpression}{strADORN}"
        PageControler1.Sort = ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    'Protected Sub bt_blankRpt_Click(sender As Object, e As EventArgs) Handles bt_blankRpt.Click
    'End Sub
End Class