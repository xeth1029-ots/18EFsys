Partial Class CM_01_001
    Inherits AuthBasePage

    'SELECT * FROM PLAN_BUDGETCAN
    'SELECT * FROM SYS_BUDGETCLOSE
    'Dim FunDr As DataRow
    Dim Auth_Relship As DataTable
    Dim CancelID As Integer
    Dim PlanKind As Integer
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '分頁設定---------------Start
        PageControler1.PageDataGrid = DG_Budget
        '分頁設定---------------End

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        bt_search.Enabled = False
        '        If FunDr("Sech") = "1" Then
        '            bt_search.Enabled = True
        '        End If
        '    End If
        'End If

        Dim sql As String = ""
        Dim dr As DataRow

        '檢查核銷方式--------------------Start
        sql = "SELECT CancelID FROM Plan_BudgetCan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, "尚未設定計價方式，此功能無效")
            bt_search.Enabled = False
        Else
            CancelID = dr("CancelID")
        End If
        '檢查核銷方式--------------------End

        '經費關帳
        Call getBudgetClose()

        '依sm.UserInfo.PlanID取得PlanKind
        PlanKind = TIMS.Get_PlanKind(Me, objconn)

        sql = ""
        sql += " SELECT a.RID,b.OrgName "
        sql += " FROM Auth_Relship a "
        sql += " JOIN Org_OrgInfo b ON a.OrgID=b.OrgID"
        Auth_Relship = DbAccess.GetDataTable(sql, objconn)

        If Not IsPostBack Then
            DataGridTable.Visible = False

            If Session("BudgetSearchStr") Is Nothing Then
                center.Text = sm.UserInfo.OrgName
                RIDValue.Value = sm.UserInfo.RID
            End If

            '取得查詢條件
            If Not Session("BudgetSearchStr") Is Nothing Then
                Dim MyValue As String = ""
                Dim sSession1 As String = Convert.ToString(Session("BudgetSearchStr"))
                Session("BudgetSearchStr") = Nothing
                center.Text = TIMS.GetMyValue(sSession1, "center")
                RIDValue.Value = TIMS.GetMyValue(sSession1, "RIDValue")
                OCID1.Text = TIMS.GetMyValue(sSession1, "OCID1")
                TMID1.Text = TIMS.GetMyValue(sSession1, "TMID1")
                OCIDValue1.Value = TIMS.GetMyValue(sSession1, "OCIDValue1")
                TMIDValue1.Value = TIMS.GetMyValue(sSession1, "TMIDValue1")
                start_date.Text = TIMS.GetMyValue(sSession1, "start_date")
                end_date.Text = TIMS.GetMyValue(sSession1, "end_date")

                MyValue = TIMS.GetMyValue(sSession1, "DropDownList1")
                Common.SetListItem(DropDownList1, MyValue)
                TB_cycltype.Text = TIMS.GetMyValue(sSession1, "TB_cycltype")

                MyValue = TIMS.GetMyValue(sSession1, "PageIndex")
                Me.ViewState("PageIndex") = MyValue

                MyValue = TIMS.GetMyValue(sSession1, "Button1")
                If MyValue = "True" Then
                    bt_search_Click(sender, e)
                    If IsNumeric(Me.ViewState("PageIndex")) Then
                        PageControler1.PageIndex = Me.ViewState("PageIndex")
                        PageControler1.CreateData()
                    End If
                End If
            End If


            If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
                Button1.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
            Else
                Button1.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
            End If
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", False)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    '取得關帳設定資料
    Private Sub getBudgetClose()
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim dr As DataRow = Nothing

        Dim sql As String = ""

        Try
            conn.Open()

            sql = ""
            sql &= " select close1,close2,close3,close4,close5,close6 from Sys_BudgetClose "
            sql += " where ((tplanid=' ') or (distid=@distid and tplanid=@tplanid)) "
            sql += " order by tplanid desc"

            With sda
                .SelectCommand = New SqlCommand(sql, conn)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("distid", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                .SelectCommand.Parameters.Add("tplanid", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
                .Fill(ds)
            End With

            If ds.Tables(0).Rows.Count > 0 Then
                dr = ds.Tables(0).Rows(0)

                hidClose1.Value = Convert.ToString(dr("close1"))
                hidClose2.Value = Convert.ToString(dr("close2"))
                hidClose3.Value = Convert.ToString(dr("close3"))
                hidClose4.Value = Convert.ToString(dr("close4"))
                hidClose5.Value = Convert.ToString(dr("close5"))
                hidClose6.Value = Convert.ToString(dr("close6"))
            End If

        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
        Finally
            conn.Close()
            If Not sda Is Nothing Then sda.Dispose()
            If Not ds Is Nothing Then ds.Dispose()
        End Try
    End Sub

    'classTag 依 class_classinfo 別名不同而送入
    Function GetWhereSql(ByVal classTag As String) As String
        Dim Rst As String
        Rst = ""

        If OCIDValue1.Value <> "" Then
            Rst += " and " & classTag & ".OCID =" & OCIDValue1.Value & "" & vbCrLf
        End If

        If start_date.Text <> "" Then
            Rst += " and " & classTag & ".STDate>= " & TIMS.to_date(start_date.Text) & vbCrLf
        End If

        If end_date.Text <> "" Then
            Rst += " and " & classTag & ".FTDate<= " & TIMS.to_date(start_date.Text) & vbCrLf '" & end_date.Text & "' " & vbCrLf
        End If

        If TB_cycltype.Text <> "" Then
            If IsNumeric(TB_cycltype.Text) Then
                TB_cycltype.Text = CInt(TB_cycltype.Text)

                If CInt(TB_cycltype.Text) < 10 Then
                    Rst += " and " & classTag & ".CyclType='0" & TB_cycltype.Text & "'" & vbCrLf
                Else
                    Rst += " and " & classTag & ".CyclType='" & TB_cycltype.Text & "'" & vbCrLf
                End If
            Else
                TB_cycltype.Text = ""
            End If
        End If

        Return Rst
    End Function

    Sub search1()
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
        If Me.TxtPageSize.Text <> Me.DG_Budget.PageSize Then Me.DG_Budget.PageSize = Me.TxtPageSize.Text

        Dim sql As String = ""
        'Dim Relship As String = ""
        If RIDValue.Value = "" Then
            RIDValue.Value = sm.UserInfo.RID
        End If
        'sql = "SELECT Relship FROM Auth_Relship WHERE RID='" & RIDValue.Value & "'"
        'Relship = DbAccess.ExecuteScalar(sql, objconn)
        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        sql = "" & vbCrLf
        sql += " select  " & vbCrLf
        sql += "  a.OCID,aa.ClassCName,a.CyclType,b.AdmPercent,a.RID,a.STDate,a.FTDate,a.THours,a.Tnum" & vbCrLf
        sql += " ,rr.orgname,rr.relship,a.planid,a.comidno,a.seqno,c.classid" & vbCrLf
        sql += " ,dbo.NVL(aa.no1,0) no1" & vbCrLf
        sql += " ,dbo.NVL(aa2.x01,0) x01" & vbCrLf
        sql += " ,dbo.NVL(aa2.x02,0) x02" & vbCrLf
        sql += " ,dbo.NVL(aa2.x03,0) x03" & vbCrLf
        sql += " ,dbo.NVL(cc.TotalCancelCost,0) TotalCancelCost" & vbCrLf
        sql += " ,((dbo.NVL(aa.no1,0)* a.Tnum) -(dbo.NVL(cc.TotalCancelCost,0))) balance" & vbCrLf

        'CostMode
        sql += " ,dd.CostMode" & vbCrLf
        sql += " ,dbo.NVL(dd.TotalCost,0) TotalCost" & vbCrLf
        sql += " ,dbo.NVL(b.AdmPercent,0) * dbo.NVL(dd.AdmCost,0) AdmCost" & vbCrLf
        '
        sql += " ,dbo.NVL(dd.TotalCost,0)+(dbo.NVL(b.AdmPercent,0) * dbo.NVL(dd.AdmCost,0)) TotalAdmCost" & vbCrLf
        sql += " ,dbo.NVL(dd.TotalCost,0)+(dbo.NVL(b.AdmPercent,0) * dbo.NVL(dd.AdmCost,0))-dbo.NVL(cc.TotalCancelCost,0) TotalAdmCancelCost" & vbCrLf

        sql += " from class_classinfo a" & vbCrLf
        sql += " join plan_planinfo b on a.planid=b.planid and a.comidno=b.comidno and a.seqno=b.seqno" & vbCrLf
        sql += " join id_plan ip on ip.planid =a.planid " & vbCrLf
        sql += " AND a.IsSuccess='Y' " & vbCrLf
        sql += " AND a.STDate<= dbo.TRUNC_DATETIME(getdate())" & vbCrLf
        sql += " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql += " AND ip.PlanID=" & sm.UserInfo.PlanID & " " & vbCrLf
        End If
        sql += " join view_ridname rr on rr.rid =a.rid" & vbCrLf
        sql += " AND rr.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Relship & "%')"

        sql += " join id_class c on a.clsid=c.clsid " & vbCrLf
        sql += " join (" & vbCrLf
        sql += "  	select a.ocid " & vbCrLf
        sql &= "    ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql += "    ,dbo.NVL(b.DefGovCost,0) DefGovCost" & vbCrLf
        sql += "    ,dbo.NVL(b.DefUnitCost,0) DefUnitCost" & vbCrLf
        sql += "    ,dbo.NVL(b.DefStdCost,0) DefStdCost" & vbCrLf
        sql += "    ,dbo.NVL(b.DefGovCost,0)+dbo.NVL(b.DefUnitCost,0)+dbo.NVL(b.DefStdCost,0) as no1" & vbCrLf
        sql += "  	from class_classinfo a" & vbCrLf
        sql += " 	join plan_planinfo b on a.planid=b.planid and a.comidno=b.comidno and a.seqno=b.seqno" & vbCrLf
        sql += "  	where 1=1 " & vbCrLf
        sql += GetWhereSql("a")

        sql += " ) aa on aa.ocid =a.ocid " & vbCrLf
        sql += " left join  (" & vbCrLf
        sql += " 	select a.ocid" & vbCrLf
        sql += " 	,sum(case when b.budgetid='01' then 1 else 0 end) x01" & vbCrLf
        sql += " 	,sum(case when b.budgetid='02' then 1 else 0 end) x02" & vbCrLf
        sql += " 	,sum(case when b.budgetid='03' then 1 else 0 end) x03" & vbCrLf
        sql += "  from class_classinfo a " & vbCrLf
        sql += " 	join class_studentsofclass b on a.ocid=b.ocid " & vbCrLf
        sql += " 	join plan_planinfo c on a.planid=c.planid and a.comidno=c.comidno and a.seqno=c.seqno  " & vbCrLf
        sql += "  	where 1=1 " & vbCrLf
        sql += GetWhereSql("a")

        sql += "   group by a.ocid " & vbCrLf
        sql += " ) aa2 on aa2.ocid = a.ocid " & vbCrLf
        sql += " left join (" & vbCrLf
        sql += " 	select a.ocid" & vbCrLf
        sql += " 	,sum(CancelCost) TotalCancelCost " & vbCrLf
        sql += " 	from class_classinfo  a " & vbCrLf
        sql += " 	join Budget_ClassCancel b on a.OCID=b.OCID " & vbCrLf
        sql += " 	join plan_planinfo c on a.planid=c.planid and a.comidno=c.comidno and a.seqno=c.seqno  " & vbCrLf
        sql += "  	where 1=1 " & vbCrLf
        sql += GetWhereSql("a")

        sql += "   group by a.OCID" & vbCrLf
        sql += " ) cc on cc.ocid=a.ocid " & vbCrLf

        sql += " left join (" & vbCrLf
        sql += "    select " & vbCrLf
        sql += " 	a.ocid" & vbCrLf
        sql += "    , min(c.CostMode) CostMode" & vbCrLf
        sql += "    , sum ( c.OPrice* c.Itemage*dbo.NVL(c.ItemCost,1))  TotalCost" & vbCrLf
        sql += "    , sum ( case when c.AdmFlag='Y' then c.OPrice* c.Itemage*dbo.NVL(c.ItemCost,1) else 0 end )  AdmCost" & vbCrLf
        sql += "    from Plan_CostItem c " & vbCrLf
        sql += "    join class_classinfo a on a.planid=c.planid and a.comidno=c.comidno and a.seqno=c.seqno"
        sql += "    where 1=1" & vbCrLf
        sql += GetWhereSql("a")
        sql += "    group by a.ocid" & vbCrLf
        sql += " ) dd on dd.ocid=a.ocid " & vbCrLf

        sql += " WHERE 1=1" & vbCrLf
        sql += GetWhereSql("a")

        If DropDownList1.SelectedValue <> "0" Then
            Dim chkVal As String = ""
            Select Case Me.DropDownList1.SelectedValue
                Case "1" '未結清
                    chkVal = ">0"
                Case "2" '已結清
                    chkVal = "=0"
                Case "3" '超支
                    chkVal = "<0"
            End Select
            If chkVal <> "" Then
                sql += " and ((dbo.NVL(aa.no1,0)* a.Tnum) -( dbo.NVL(cc.TotalCancelCost,0))) " & chkVal & vbCrLf
            End If
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Visible = False
        Me.DG_Budget.Visible = False
        msg.Visible = True
        msg.Text = "查無資料!!"

        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            Me.DG_Budget.Visible = True
            msg.Visible = False
            msg.Text = ""

            DG_Budget.DataKeyField = "OCID"
            DG_Budget.Columns(6).Visible = True '公務
            DG_Budget.Columns(7).Visible = True '就安
            DG_Budget.Columns(8).Visible = True '就保

            'PageControler1.SqlPrimaryKeyDataCreate(Sql, "OCID", "ClassID,CyclType")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "ClassID,CyclType"
            PageControler1.ControlerLoad()
        End If
    End Sub

    '查詢
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call search1()
    End Sub

    Private Sub DG_Budget_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Budget.ItemCommand
        If e.CommandName = "edit" Then
            GetSearchStr()
            TIMS.Utl_Redirect1(Me, "CM_01_001_add.aspx?ID=" & Request("ID") & " &" & e.CommandArgument & "")
        End If
    End Sub

    Private Sub DG_Budget_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Budget.ItemDataBound
        'Dim cost As Double
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Not Me.ViewState("sort") Is Nothing Then
                    Dim img As New UI.WebControls.Image
                    Dim i As Integer
                    Select Case Me.ViewState("sort")
                        Case "ClassCName", "ClassCName desc"
                            i = 2
                        Case "STDate", "STDate desc"
                            i = 3
                        Case "TotalAdmCancelCost", "TotalAdmCancelCost desc"
                            i = 11
                    End Select

                    If Me.ViewState("sort").ToString.IndexOf("desc") = -1 Then
                        img.ImageUrl = "../../images/SortUp.gif"
                    Else
                        img.ImageUrl = "../../images/SortDown.gif"
                    End If
                    e.Item.Cells(i).Controls.Add(img)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim but1 As Button = e.Item.FindControl("edit_but")
                Dim dr As DataRowView = e.Item.DataItem 'drv
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DG_Budget.CurrentPageIndex * DG_Budget.PageSize
                e.Item.Cells(3).Text = "" & dr("STDate") & "<br>" & dr("FTDate") & ""
                '計畫人數/每人費用
                '每人費用
                'cost = Convert.ToDouble(dr("no1") / dr("Tnum")) '(no1/tnum) 
                e.Item.Cells(5).Text = "" & dr("Tnum") & "/" & Format(CDbl(dr("no1")), "#,##0.00") & ""
                '就安人數/金額
                e.Item.Cells(6).Text = "" & dr("x02") & "/" & Format(CDbl(dr("no1") * dr("x02")), "#,##0.00") & " "
                ''就保人數/金額
                e.Item.Cells(7).Text = "" & dr("x03") & "/" & Format(CDbl(dr("no1") * dr("x03")), "#,##0.00") & ""
                '公務人數/金額
                e.Item.Cells(8).Text = "" & dr("x01") & "/" & Format(CDbl(dr("no1") * dr("x01")), "#,##0.00") & ""
                '計畫總經費
                'Dim TotalCost As Double
                'Dim AdmCost As Double
                'For Each dr1 As DataRow In CostTable.Select("PlanID='" & dr("PlanID") & "' and ComIDNO='" & dr("ComIDNO") & "' and SeqNo='" & dr("SeqNo") & "'")
                '    TotalCost += IIf(IsDBNull(dr1("OPrice")), 1, dr1("OPrice")) * IIf(IsDBNull(dr1("Itemage")), 1, dr1("Itemage")) * IIf(IsDBNull(dr1("ItemCost")), 1, dr1("ItemCost"))
                'Next
                'For Each dr1 As DataRow In CostTable.Select("PlanID='" & dr("PlanID") & "' and ComIDNO='" & dr("ComIDNO") & "' and SeqNo='" & dr("SeqNo") & "' and AdmFlag='Y'")
                '    AdmCost += IIf(IsDBNull(dr1("OPrice")), 1, dr1("OPrice")) * IIf(IsDBNull(dr1("Itemage")), 1, dr1("Itemage")) * IIf(IsDBNull(dr1("ItemCost")), 1, dr1("ItemCost"))
                'Next
                'AdmCost = IIf(IsDBNull(dr("AdmPercent")), 0, dr("AdmPercent")) * AdmCost

                '計畫總經費
                'e.Item.Cells(9).Text = Format(CDbl(TotalCost + AdmCost), "#,##0.00")
                e.Item.Cells(9).Text = Format(CDbl(dr("TotalAdmCost")), "#,##0.00")
                '已核銷總金額
                e.Item.Cells(10).Text = Format(CDbl(dr("TotalCancelCost")), "#,##0.00")
                '結餘總金額
                'e.Item.Cells(11).Text = Format(CDbl(CDbl(TotalCost + AdmCost) - dr("TotalCancelCost")), "#,##0.00")
                e.Item.Cells(11).Text = Format(CDbl(dr("TotalAdmCancelCost")), "#,##0.00")

                If CDbl(dr("TotalAdmCancelCost")) < 0 Then '負數則顯示紅色
                    e.Item.Cells(11).ForeColor = Color.Red
                End If

                Dim Parent As String
                If Split(dr("Relship"), "/").Length > 2 Then
                    Parent = Split(dr("Relship"), "/")(Split(dr("Relship"), "/").Length - 3)
                    e.Item.Cells(1).Text = "<font color='Blue'>" & Auth_Relship.Select("RID='" & Parent & "'")(0)("OrgName") & "</font>-" & e.Item.Cells(1).Text
                End If

                'Sys_BudgetClose-關帳設定
                If PlanKind = 1 Then
                    '1.自辦(內訓)
                    If Year(dr("STDate")) = Year(dr("FTDate")) Then
                        If Now > CDate(Year(dr("FTDate")) + 1 & "/" & hidClose2.Value) Then
                            but1.Enabled = False
                        Else
                            but1.Enabled = True
                            'If FunDr("Adds") = "1" Then
                            '    but1.Enabled = True
                            'Else
                            '    but1.Enabled = False
                            'End If
                        End If
                    Else
                        If Now > CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose1.Value)) Then
                            but1.Enabled = False
                        Else
                            but1.Enabled = True
                            'If FunDr("Adds") = "1" Then
                            '    but1.Enabled = True
                            'Else
                            '    but1.Enabled = False
                            'End If
                        End If
                    End If
                Else
                    '2.委外(PlanKind!=1)
                    '結訓日開訓日同年
                    If Year(dr("STDate")) = Year(dr("FTDate")) Then
                        If dr("RID").ToString.Length = 1 Then '有分署(中心)以上權限者
                            Me.ViewState("closedate") = Common.FormatDate(CDate(Year(dr("FTDate")) + 1 & "/" & hidClose6.Value))
                            If Now > CDate(Year(dr("FTDate")) + 1 & "/" & hidClose6.Value) Then
                                but1.Enabled = False
                                'TIMS.Tooltip(but1, "結訓日開訓日同年，中心委外課程停止核銷日為" & Me.ViewState("closedate"))
                                TIMS.Tooltip(but1, "結訓日開訓日同年，分署委外課程停止核銷日為" & Me.ViewState("closedate"))
                                'Else
                                '    If FunDr("Adds") = "1" Then
                                '        but1.Enabled = True
                                '    Else
                                '        but1.Enabled = False
                                '        TIMS.Tooltip(but1, "結訓日開訓日同年，中心委外課程 無核銷權限[Adds]")
                                '    End If
                            End If
                        Else
                            Me.ViewState("closedate") = Common.FormatDate(CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose5.Value)))
                            If Now > CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose5.Value)) Then
                                but1.Enabled = False
                                TIMS.Tooltip(but1, "結訓日開訓日同年，委訓委外課程停止核銷日為" & Me.ViewState("closedate"))
                            Else
                                but1.Enabled = True
                                'If FunDr("Adds") = "1" Then
                                '    but1.Enabled = True
                                'Else
                                '    but1.Enabled = False
                                '    TIMS.Tooltip(but1, "結訓日開訓日同年，委訓委外課程 無核銷權限[Adds]")
                                'End If
                            End If
                        End If
                    Else
                        If dr("RID").ToString.Length = 1 Then
                            Me.ViewState("closedate") = Common.FormatDate(CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose4.Value)))
                            If Now > CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose4.Value)) Then
                                but1.Enabled = False
                                'TIMS.Tooltip(but1, "結訓日開訓日跨年，中心委外課程停止核銷日為" & Me.ViewState("closedate"))
                                TIMS.Tooltip(but1, "結訓日開訓日跨年，分署委外課程停止核銷日為" & Me.ViewState("closedate"))
                            Else
                                but1.Enabled = True
                                'If FunDr("Adds") = "1" Then
                                '    but1.Enabled = True
                                'Else
                                '    but1.Enabled = False
                                '    TIMS.Tooltip(but1, "結訓日開訓日跨年，中心委外課程 無核銷權限[Adds]")
                                'End If
                            End If
                        Else
                            Me.ViewState("closedate") = Common.FormatDate(CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose3.Value)))
                            If Now > CDate(dr("FTDate")).AddMonths(Convert.ToInt16(hidClose3.Value)) Then
                                but1.Enabled = False
                                TIMS.Tooltip(but1, "結訓日開訓日跨年，委訓委外課程停止核銷日為" & Me.ViewState("closedate"))
                            Else
                                but1.Enabled = True
                                'If FunDr("Adds") = "1" Then
                                '    but1.Enabled = True
                                'Else
                                '    but1.Enabled = False
                                '    TIMS.Tooltip(but1, "結訓日開訓日跨年，委訓委外課程 無核銷權限[Adds]")
                                'End If
                            End If
                        End If
                    End If
                End If

                but1.CommandArgument = "PlanID=" & dr("PlanID") & "&ComIDNO=" & dr("ComIDNO") & "&SeqNO=" & dr("SeqNO") & "&OCID=" & dr("OCID") & "&toalprice=" & e.Item.Cells(9).Text
                but1.CommandArgument += "&CostMode=" & Convert.ToString(dr("CostMode"))
        End Select
    End Sub

    Private Sub DG_Budget_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_Budget.SortCommand
        If e.SortExpression = Me.ViewState("sort") Then
            Me.ViewState("sort") = e.SortExpression & " desc"
        Else
            Me.ViewState("sort") = e.SortExpression
        End If

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    Sub GetSearchStr()
        Session("BudgetSearchStr") = "center=" & center.Text
        Session("BudgetSearchStr") += "&RIDValue=" & RIDValue.Value
        Session("BudgetSearchStr") += "&OCID1=" & OCID1.Text
        Session("BudgetSearchStr") += "&TMID1=" & TMID1.Text
        Session("BudgetSearchStr") += "&OCIDValue1=" & OCIDValue1.Value
        Session("BudgetSearchStr") += "&TMIDValue1=" & TMIDValue1.Value
        Session("BudgetSearchStr") += "&start_date=" & start_date.Text
        Session("BudgetSearchStr") += "&end_date=" & end_date.Text
        Session("BudgetSearchStr") += "&DropDownList1=" & DropDownList1.SelectedValue
        Session("BudgetSearchStr") += "&TB_cycltype=" & TB_cycltype.Text
        Session("BudgetSearchStr") += "&PageIndex=" & DG_Budget.CurrentPageIndex + 1
        Session("BudgetSearchStr") += "&Button1=" & DG_Budget.Visible
    End Sub
End Class
