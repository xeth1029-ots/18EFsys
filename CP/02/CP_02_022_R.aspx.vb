Partial Class CP_02_022_R
    Inherits AuthBasePage

    Dim dr As DataRow

    Dim DistID As String
    Dim LID As String
    Dim RID As String
    Dim OrgID As String
    Dim PlanID As String

   'Dim FunDr As DataRow
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        'Dim sql As String
        'Dim PlanKind As String
        'Dim Sqlstr As String

        'Sqlstr = "SELECT a.RID,a.Relship,b.OrgName FROM "
        'Sqlstr += "Auth_Relship a "
        'Sqlstr += "JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "
        ''  Common.RespWrite(Me, sql)
        'RelshipTable = DbAccess.GetDataTable(Sqlstr, objconn)

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?selected_year=" & yearlist.SelectedValue.ToString & "');"
        Else
            Org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx?selected_year=" & yearlist.SelectedValue.ToString & "');"
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        Dim PlanKind As String
        '依sm.UserInfo.PlanID取得PlanKind
        PlanKind = TIMS.Get_PlanKind(Me, objconn)
        If PlanKind = "1" Then
            dtPlan.Columns(6).Visible = False
        Else
            dtPlan.Columns(6).Visible = True
        End If

        Me.msg.Text = ""
        '分頁設定---------------Start
        PageControler1.PageDataGrid = dtPlan
        '分頁設定---------------End

        DistID = sm.UserInfo.DistID
        LID = sm.UserInfo.LID
        RID = sm.UserInfo.RID
        OrgID = sm.UserInfo.OrgID
        PlanID = sm.UserInfo.PlanID

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    Re_ID.Value = Request("ID")
        '    FunDr = FunDrArray(0)
        '    If FunDr("Sech") = 1 Then
        '        btnQuery.Enabled = True
        '    Else
        '        btnQuery.Enabled = False
        '    End If
        'End If

        Dim Sqlstr As String = ""
        Sqlstr = "select TPlanID,PlanKind From ID_Plan where PlanID='" & PlanID & "'"
        dr = DbAccess.GetOneRow(Sqlstr, objconn)

        If Not Me.IsPostBack Then
            DataGridTable.Visible = False
            yearlist = TIMS.GetSyear(yearlist)

            '2005/4/1--Melody年度帶預設值
            Common.SetListItem(yearlist, sm.UserInfo.Years)
            '(加強操作便利性)
            RIDValue.Value = RID
            Dim sqlstring As String = "select orgname from Auth_Relship a join Org_orginfo b on  a.orgid=b.orgid where a.RID='" & RID & "'"
            Dim orgname As String = DbAccess.ExecuteScalar(sqlstring, objconn)
            center.Text = orgname

            '取得訓練計畫
            Sqlstr = "select TPlanID  from ID_Plan where PlanID=" & sm.UserInfo.PlanID & ""
            TPlanid.Value = DbAccess.ExecuteScalar(Sqlstr, objconn)
        End If

        '帶入查詢參數
        If Not IsPostBack Then
            If Not Session("search") Is Nothing Then
                Dim MyValue As String = ""
                Dim strSession As String = Session("search")
                MyValue = TIMS.GetMyValue(strSession, "yearlist")
                Common.SetListItem(yearlist, MyValue)
                TB_career_id.Text = TIMS.GetMyValue(strSession, "TB_career_id")
                trainValue.Value = TIMS.GetMyValue(strSession, "trainValue")
                center.Text = TIMS.GetMyValue(strSession, "center")
                RIDValue.Value = TIMS.GetMyValue(strSession, "RIDValue")
                ClassName.Text = TIMS.GetMyValue(strSession, "ClassName")
                MyValue = TIMS.GetMyValue(strSession, "PageIndex")
                If MyValue <> "" Then
                    PageControler1.PageIndex = Val(MyValue)
                End If

                'btnQuery_Click(sender, e)
                Call Search1()
                Session("search") = Nothing
            End If
        End If
    End Sub

    Sub Search1()
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
            Me.TxtPageSize.Text = 10
        End If
        Me.dtPlan.PageSize = Me.TxtPageSize.Text

        dtPlan.CurrentPageIndex = 0
        Call LoadData()
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call Search1()
    End Sub

    Private Sub dtPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim btn1 As Button = e.Item.FindControl("Button1")
                Dim drv As DataRowView = e.Item.DataItem
                btn1.Visible = True
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + dtPlan.PageSize * dtPlan.CurrentPageIndex

                e.Item.Cells(7).Text = drv("ClassName").ToString
                If drv("CyclType").ToString <> "" Then
                    If Int(drv("CyclType")) <> 0 Then
                        e.Item.Cells(7).Text += "第" & Int(drv("CyclType")) & "期"
                    End If
                End If

                '列印訓練計畫but
                'btn1.CommandArgument = "PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNO=" & drv("SeqNO")
                btn1.Attributes("onclick") =     ReportQuery.ReportScript(Me, "list", "CP_02_022_R", "OCID=" & drv("OCID") & "&YEARS=" & Val(drv("PlanYear").ToString) - 1911 & "")

            Case ListItemType.Header
                If Me.ViewState("sort") <> "" Then
                    Dim mylabel As String
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "ClassName", "ClassName DESC"
                            mylabel = "ComName"
                            i = 7
                            If Me.ViewState("sort") = "ClassName" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "AppliedDate", "AppliedDate DESC"
                            mylabel = "ComName"
                            i = 2
                            If Me.ViewState("sort") = "AppliedDate" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "STDate", "STDate DESC"
                            mylabel = "ComName"
                            i = 4
                            If Me.ViewState("sort") = "STDate" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "FDDate", "FDDate DESC"
                            mylabel = "ComName"
                            i = 5
                            If Me.ViewState("sort") = "FDDate" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "OrgName", "OrgName DESC"
                            mylabel = "ComName"
                            i = 6
                            If Me.ViewState("sort") = "OrgName" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                    End Select
                    If i <> -1 Then
                        e.Item.Cells(i).Controls.Add(mysort)
                    End If
                End If
        End Select
    End Sub

    Public Sub dtPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dtPlan.ItemCommand
        Page.RegisterStartupScript("Londing", "<script>Layer_change(6);</script>")
    End Sub

    'Dim RelshipTable As DataTable

    Sub LoadData()
        '取出訓練計畫名稱
        Dim StrSql As String = ""
        Dim relshipstr As String = "select relship from Auth_Relship where RID='" & RIDValue.Value & "'"
        Dim relship As String = DbAccess.ExecuteScalar(relshipstr, objconn)
        Dim drTemp As DataRow
        StrSql = "SELECT PlanName FROM Key_Plan WHERE TPlanID IN (SELECT TPlanID FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "')"
        drTemp = DbAccess.GetOneRow(StrSql, objconn)
        If Not drTemp Is Nothing Then
            TPlanName.Text = drTemp("PlanName").ToString
        End If

        '建立訓練的DATATABLE
        StrSql = "select P1.PlanID,O1.OrgName,P1.STDate,P1.FDDate, "
        StrSql += "P1.ComIDNO,P1.PlanYear,P1.TMID,P1.AppliedResult,"
        StrSql += "P1.SeqNO,K1.PlanName,K2.TrainName,P1.ClassName, "
        StrSql += "P1.CyclType,P1.AppliedDate,P1.RID,C1.OCID "
        If sm.UserInfo.RID = "A" Then
            StrSql += "From (SELECT * FROM Plan_PlanInfo WHERE PlanID IN (SELECT PlanID From ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Years=(SELECT Years FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'))) P1 "
        Else
            StrSql += "From (SELECT * FROM Plan_PlanInfo WHERE PlanID='" & sm.UserInfo.PlanID & "') P1 "
        End If
        StrSql += "JOIN Key_Plan K1 ON P1.TPlanID=K1.TPlanID "
        StrSql += "LEFT JOIN Key_TrainType K2 ON P1.TMID=K2.TMID "
        StrSql += "JOIN ID_Plan I1 ON P1.PlanID=I1.PlanID "
        StrSql += "JOIN Org_OrgInfo O1 ON P1.ComIDNO=O1.ComIDNO "
        StrSql += "JOIN (SELECT * FROM Class_ClassInfo WHERE IsSuccess='Y') C1 ON P1.PlanID=C1.PlanID and P1.ComIDNO=C1.ComIDNO and P1.SeqNo=C1.SeqNo "
        StrSql += "WHERE P1.TPlanID='" & TPlanid.Value & "' "

        If sm.UserInfo.RID = "A" Then
            StrSql += " and P1.RID IN (SELECT RID FROM Auth_Relship WHERE relship like '" & relship & "%')"
        Else
            If dr("PlanKind") = 2 Then
                StrSql += " and P1.RID IN (SELECT RID FROM Auth_Relship WHERE relship like '" & relship & "%')"
            End If
            If LID = 0 Or LID = 1 Then
                StrSql = StrSql & " and I1.DistID='" & DistID & "'"
            End If
        End If

        If Me.yearlist.SelectedValue <> "" Then
            StrSql = StrSql & " and P1.PlanYear='" & Me.yearlist.SelectedValue & "'"
        End If

        If Me.trainValue.Value <> "" Then
            StrSql = StrSql & " and P1.TMID='" & Me.trainValue.Value & "'"
        End If
        If ClassName.Text <> "" Then
            StrSql += " and ClassName like'%" & ClassName.Text & "%'"
        End If
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    StrSql += " and P1.CyclType='0" & Int(CyclType.Text) & "'"
                Else
                    StrSql += " and P1.CyclType='" & CyclType.Text & "'"
                End If
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(StrSql, objconn)
        DataGridTable.Visible = False
        Me.msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            Me.msg.Text = ""

            'DataGridTable.Visible = True
            If ViewState("sort") = "" Then
                ViewState("sort") = "STDate,ClassName"
            End If

            'PageControler1.SqlString = StrSql
            PageControler1.PageDataTable = dt '.SqlString = StrSql
            PageControler1.Sort = ViewState("sort")
            PageControler1.ControlerLoad()
        End If

        'If TIMS.Get_SQLRecordCount(StrSql, objconn) = 0 Then
        '    DataGridTable.Visible = False
        '    Me.msg.Text = "查無資料"
        'Else
        '    DataGridTable.Visible = True
        '    If ViewState("sort") = "" Then
        '        ViewState("sort") = "STDate,ClassName"
        '    End If

        '    PageControler1.SqlString = StrSql
        '    PageControler1.Sort = ViewState("sort")
        '    PageControler1.ControlerLoad()
        'End If
    End Sub

    Private Sub dtPlan_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dtPlan.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If
        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    'Sub GetSearchStr()
    '    Dim strSession As String = ""
    '    strSession = "yearlist=" & yearlist.SelectedValue
    '    strSession &= "&TB_career_id=" & TB_career_id.Text
    '    strSession &= "&trainValue=" & trainValue.Value
    '    strSession &= "&center=" & center.Text
    '    strSession &= "&RIDValue=" & RIDValue.Value
    '    strSession &= "&ClassName=" & ClassName.Text
    '    strSession &= "&PageIndex=" & dtPlan.CurrentPageIndex + 1
    '    Session("search") = strSession
    '    'Session("search") = "yearlist=" & yearlist.SelectedValue & "&"
    '    'Session("search") += "TB_career_id=" & TB_career_id.Text
    '    ''& "&trainValue=" & trainValue.Value & "&"
    '    'Session("search") += "center=" & center.Text & "&RIDValue=" & RIDValue.Value & "&"
    '    'Session("search") += "ClassName=" & ClassName.Text & "&"
    '    'Session("search") += "PageIndex=" & dtPlan.CurrentPageIndex + 1
    'End Sub
End Class

