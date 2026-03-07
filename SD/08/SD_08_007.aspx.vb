Partial Class SD_08_007
    Inherits AuthBasePage

   'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
       'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線

        PageControler1.PageDataGrid = DataGrid1

        '檢查帳號的功能權限
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    End If
        'End If

        If Not IsPostBack Then
            If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
                Button7.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx')"
            Else
                Button7.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx')"
            End If

            '設定身分別下拉選單
            Dim sIdentityID2 As String = TIMS.Get_SubsidyID(objconn)
            Dim dt As DataTable
            Dim sql As String = ""
            sql = "SELECT * FROM Key_Identity WHERE IdentityID IN   (" & sIdentityID2 & ")"
            dt = DbAccess.GetDataTable(sql, objconn)
            With SCHIdentityID
                .DataSource = dt
                .DataValueField = "IdentityID"
                .DataTextField = "name"
                .DataBind()
                .Items.Insert(0, New ListItem("請選擇", ""))
            End With

            '設定 Client 端屬性
            Search.Attributes("onclick") = "return chkdata();"

            IMG1.Attributes("onclick") = "show_calendar('" & SCHASDate.ClientID & "','','','CY/MM/DD');"

            IMG2.Attributes("onclick") = "show_calendar('" & SCHAEDate.ClientID & "','','','CY/MM/DD');"

            TablePage.Visible = False

            searchPanel.Visible = True
            historyPanel.Visible = False
        End If


    End Sub

    Private Sub Search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Search.Click
        TablePage.Visible = False
        DataGrid1.CurrentPageIndex = 0

        Dim dt As DataTable
        Dim sql As String = ""
        '查詢資料
        sql = "" & vbCrLf
        sql += " SELECT a.*, b.OrgName, c.name stdName, c.sex, c.birthday stdBirthday " & vbCrLf
        sql += " FROM Sub_SubsidyApply a" & vbCrLf
        sql += " join Org_OrgInfo b on  a.OrgId=b.OrgID" & vbCrLf
        sql += " join (" & vbCrLf
        sql += "  select sid, name, sex, Birthday from stud_studentinfo ss" & vbCrLf
        sql += "  where exists ( select 'x' from sub_subsidyapply x where x.sid =ss.sid)" & vbCrLf
        sql += " ) c on a.sid=c.sid" & vbCrLf
        sql += " where 1=1" & vbCrLf

        '訓練機構
        If center.Text <> "" Then
            sql += " And a.OrgId = '" & orgid_value.Value & "'" & vbCrLf
        Else
            '僅能選擇自已及底下的單位資料
            sql += " And a.OrgId in (SELECT a. orgid FROM Org_OrgInfo a " & vbCrLf
            sql += " JOIN (SELECT * FROM Auth_Relship WHERE PlanID=0 or (PlanID = " & sm.UserInfo.PlanID & " and PlanID <> 0)) b " & vbCrLf
            sql += " ON a.OrgID=b.OrgID JOIN Org_OrgPlanInfo c ON b.RSID=c.RSID " & vbCrLf
            sql += " where b.rid like '" & sm.UserInfo.RID & "%') " & vbCrLf
        End If

        '身分別
        If SCHIdentityID.SelectedValue <> "" Then
            sql += " and a.IdentityID like '" & SCHIdentityID.SelectedValue & "%'" & vbCrLf
        End If

        '身分證號
        If SCHIDNO.Text <> "" Then
            sql += " and a.IDNO like '%" & SCHIDNO.Text & "%'" & vbCrLf
        End If

        '性別
        If SCHSex1.Checked And SCHSex2.Checked Then
            sql += " and c.sex in ('F','M')" & vbCrLf
        ElseIf SCHSex1.Checked And Not SCHSex2.Checked Then
            sql += " and c.sex in ('M')" & vbCrLf
        ElseIf Not SCHSex1.Checked And SCHSex2.Checked Then
            sql += " and c.sex in ('F')" & vbCrLf
        End If

        '申請日期
        If SCHASDate.Text <> "" Then
            sql += " and a.applydate >= " & TIMS.to_date(SCHASDate.Text) & vbCrLf
            'sql += " and convert(varchar,a.applydate,111) >= '" & SCHASDate.Text & "'"
        End If

        If SCHAEDate.Text <> "" Then
            sql += " and a.applydate <= " & TIMS.to_date(SCHAEDate.Text) & vbCrLf
            'sql += " and convert(varchar,a.applydate,111) <= '" & SCHAEDate.Text & "'"
        End If

        '年齡
        If SCHSAge.Text <> "" Then
            sql += " and trunc((dbo.TRUNC_DATETIME(getdate())-c.birthday) /365) <=" & SCHSAge.Text & vbCrLf
            'sql += " and c.birthday <= dateadd(year,-" & SCHSAge.Text & ",getdate()) "
        End If

        If SCHEAge.Text <> "" Then
            sql += " and trunc((dbo.TRUNC_DATETIME(getdate())-c.birthday) /365) >" & SCHEAge.Text & vbCrLf
            'sql += " and c.birthday > dateadd(year,-" & SCHEAge.Text & ",getdate()) "
        End If

        '申請月數
        If SCHSMonth.Text <> "" Then
            sql += " and a.ApplyMonth >= " & SCHSMonth.Text & vbCrLf
        End If

        If SCHEMonth.Text <> "" Then
            sql += " and a.ApplyMonth <= " & SCHEMonth.Text & vbCrLf
        End If

        '申請金額
        If SCHSMoney.Text <> "" Then
            sql += " and a.ApplyMoney >= " & SCHSMoney.Text & vbCrLf
        End If

        If SCHEMoney.Text <> "" Then
            sql += " and a.ApplyMoney <= " & SCHEMoney.Text & vbCrLf
        End If
        dt = DbAccess.GetDataTable(sql, objconn)

        TablePage.Visible = False
        DataGrid1.Visible = False
        mesg.Text = "查無符合條件之資料"
        If dt.Rows.Count > 0 Then
            TablePage.Visible = True
            DataGrid1.Visible = True
            mesg.Text = ""

            PageControler1.PageDataTable = dt ' PageControler1.SqlString = sql
            PageControler1.ControlerLoad()
        End If
        '設定 DataGrid 滑鼠 scroll 顏色
        CommonUtil.set_row_color(DataGrid1)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Select Case e.CommandName
            Case "View"
                Dim IDNO As String = TIMS.GetMyValue(sCmdArg, "IDNO")
                'Dim idno1 As LinkButton = DataGrid1.Items(e.Item.ItemIndex).FindControl("txtidno")
                'idno1.Text = TIMS.ChangeIDNO(idno1.Text)
                Dim dt1 As DataTable
                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " select a.*" & vbCrLf
                sql &= " ,b.name" & vbCrLf
                sql &= " ,b.birthday " & vbCrLf
                sql &= " from sub_subsidyapply a " & vbCrLf
                sql &= " join stud_studentinfo b on a.SID = b.SID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and a.IDNO = '" & IDNO & "'" & vbCrLf
                dt1 = DbAccess.GetDataTable(sql, objconn)

                'strsql = "select a.*, b.name, b.birthday from sub_subsidyapply a, stud_studentinfo b "
                'strsql += "where a.SID = b.SID and a.IDNO = '" & idno1.Text & "'"
                'dt1 = DbAccess.GetDataTable(strsql)
                If dt1.Rows.Count = 0 Then
                    Datagrid2.Visible = False
                    TablePage2.Visible = False
                    searchPanel.Visible = True
                    historyPanel.Visible = False
                    Exit Sub
                End If

                Datagrid2.Visible = True
                TablePage2.Visible = True
                searchPanel.Visible = False
                historyPanel.Visible = True
                Datagrid2.DataSource = dt1
                Datagrid2.DataBind()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Not Me.ViewState("sort") Is Nothing Then
                    Dim img As New UI.WebControls.Image
                    Dim i As Integer

                    Select Case Me.ViewState("sort")
                        Case "OrgName", "OrgName desc"
                            i = 0
                        Case "IDNO", "IDNO desc"
                            i = 1
                        Case "stdname", "stdname desc"
                            i = 2
                        Case "stdBirthday", "stdBirthday desc"
                            i = 3
                        Case "ApplyDate", "ApplyDate desc"
                            i = 4
                    End Select

                    If Me.ViewState("sort").ToString.IndexOf("desc") = -1 Then
                        img.ImageUrl = "../../images/SortUp.gif"
                    Else
                        img.ImageUrl = "../../images/SortDown.gif"
                    End If

                    e.Item.Cells(i).Controls.Add(img)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim txtidno As LinkButton = e.Item.FindControl("txtidno")
                Dim objname As Label = e.Item.FindControl("txtname")
                Dim objbirth As Label = e.Item.FindControl("birth")
                Dim objapply As Label = e.Item.FindControl("apply")
                Dim objasf As Label = e.Item.FindControl("asf")
                Dim objasfin As Label = e.Item.FindControl("asfin")
                Dim objisdl As Label = e.Item.FindControl("isdl")

                objname.Text = Convert.ToString(drv("stdName"))
                txtidno.Text = Convert.ToString(drv("IDNO"))
                Dim sCmdArag As String = ""
                TIMS.SetMyValue(sCmdArag, "IDNO", Convert.ToString(drv("IDNO")))
                txtidno.CommandArgument = sCmdArag

                '生日
                If Not Convert.IsDBNull(drv("stdBirthday")) Then
                    objbirth.Text = Convert.ToDateTime(drv("stdBirthday")).ToString("yyyy/MM/dd")
                End If

                '申請日期
                If Not Convert.IsDBNull(drv("ApplyDate")) Then
                    objapply.Text = Convert.ToDateTime(drv("ApplyDate")).ToString("yyyy/MM/dd")
                End If

                '受訓起訖
                e.Item.Cells(3).Text = Convert.ToDateTime(drv("TSDate")).ToString("yyyy/MM/dd") & "<br>" & Convert.ToDateTime(drv("TEDate")).ToString("yyyy/MM/dd")

                '申請月數、金額
                e.Item.Cells(4).Text = Convert.ToString(drv("ApplyMonth")) & "<br>" & Convert.ToString(drv("ApplyMoney"))

                If Convert.ToString(drv("AppliedStatusF")) = "Y" Then
                    objasf.Text = "通過"
                ElseIf Convert.ToString(drv("AppliedStatusF")) = "N" Then
                    objasf.Text = "未過"
                End If

                Select Case drv("AppliedStatusFin").ToString
                    Case "Y"
                        objasfin.Text = "通過"
                    Case ""
                        objasfin.Text = ""
                    Case Else
                        objasfin.Text = drv("AppliedStatusFin").ToString
                End Select

                If Convert.ToString(drv("isDownload")) = "1" Then
                    objisdl.Text = "已送"
                End If
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If e.SortExpression = Me.ViewState("sort") Then
            Me.ViewState("sort") = e.SortExpression & " desc"
        Else
            Me.ViewState("sort") = e.SortExpression
        End If

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()

        '設定 DataGrid 滑鼠 scroll 顏色
        CommonUtil.set_row_color(DataGrid1)
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim objasf1 As Label = e.Item.FindControl("asf1")
                Dim objasfin1 As Label = e.Item.FindControl("asfin1")
                Dim objisdl1 As Label = e.Item.FindControl("isdl1")

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                If Convert.ToString(drv("AppliedStatusF")) = "Y" Then
                    objasf1.Text = "通過"
                ElseIf Convert.ToString(drv("AppliedStatusF")) = "N" Then
                    objasf1.Text = "未過"
                End If

                Select Case drv("AppliedStatusFin").ToString
                    Case "Y"
                        objasfin1.Text = "通過"
                    Case ""
                        objasfin1.Text = ""
                    Case Else
                        objasfin1.Text = drv("AppliedStatusFin").ToString
                End Select

                If Convert.ToString(drv("isDownload")) = "1" Then
                    objisdl1.Text = "已送"
                End If

                e.Item.Cells(1).Text = Convert.ToDateTime(drv("applydate")).ToString("yyyy/MM/dd")
                e.Item.Cells(2).Text = Convert.ToDateTime(drv("tsdate")).ToString("yyyy/MM/dd") & "<br>" & Convert.ToDateTime(drv("tedate")).ToString("yyyy/MM/dd")
                e.Item.Cells(3).Text = Convert.ToString(drv("ApplyMonth")) & "<br>" & Convert.ToString(drv("ApplyMoney"))
                e.Item.Cells(4).Text = Convert.ToString(drv("PayMonth")) & "<br>" & Convert.ToString(drv("PayMoney"))

                headidno.Text = Convert.ToString(drv("IDNO"))
                headbrith.Text = Convert.ToDateTime(drv("Birthday")).ToString("yyyy/MM/dd")
                headname.Text = Convert.ToString(drv("Name"))
        End Select
    End Sub

    Private Sub backButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackButton.Click
        searchPanel.Visible = True
        historyPanel.Visible = False
    End Sub

End Class
