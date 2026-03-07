Partial Class LevPlan
    Inherits AuthBasePage

    Dim dtOrgBlack As DataTable  '取出系統現有黑名單
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        dtOrgBlack = TIMS.Get_OrgBlackList(Me, objconn)  '取出系統現有黑名單

        If Not Page.IsPostBack Then
            cCreate1()
            planlist.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        End If
    End Sub

    Sub cCreate1()
        hid_ORGBLACKLIST.Value = ""
        If TIMS.Check_OrgBlackList(Me, "", objconn) Then hid_ORGBLACKLIST.Value = "Y"
        cbplanlistAll.Visible = False
        'Select Case sm.UserInfo.TPlanID
        '    Case "17"
        '        If sm.UserInfo.RoleID <= 5 And sm.UserInfo.LID <= 1 Then cbplanlistAll.Visible = True
        '        If sm.UserInfo.RoleID = "0" Then cbplanlistAll.Visible = False
        'End Select
        yearlist = TIMS.GetSyear(yearlist)
        planlist.Items.Clear()
        planlist.Items.Insert(0, New ListItem("==請選擇==", ""))

        '帶預設值
        Dim parms As New Hashtable
        parms.Add("PLANID", Val(sm.UserInfo.PlanID))
        Dim sql As String
        sql = " SELECT YEARS FROM ID_PLAN WHERE PLANID=@PLANID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr IsNot Nothing Then
            Common.SetListItem(yearlist, dr("YEARS"))
            'yearlist_SelectedIndexChanged(sender, e)
            Call SHOW_PLANLIST()

            Common.SetListItem(planlist, sm.UserInfo.PlanID)
            'planlist_SelectedIndexChanged(sender, e)
            Call sSearch1()
            If sm.UserInfo.LID = 2 Then Table1.Visible = False
        End If

    End Sub

    Sub AddTreeNodes(ByVal dr As DataRow, ByVal objTable As DataTable, ByVal objTreeView As TreeView, ByVal ParentNode As TreeNode)
        If dr Is Nothing Then Exit Sub

        Dim NewNode As New TreeNode
        Dim drChild As DataRow
        Dim strFilter As String

        NewNode.Text = Convert.ToString(dr("OrgName"))
        Dim v_planlist As String = TIMS.GetListValue(planlist) 'Me.planlist.SelectedValue
        Dim v_planlist_txt As String = TIMS.GetListText(planlist) 'Me.planlist.SelectedItem.Text

        Dim sb_wsNavigateUrl As New StringBuilder
        'wsNavigateUrl = ""
        sb_wsNavigateUrl.Append("javascript:returnValue('").Append(Convert.ToString(dr("RID")))
        sb_wsNavigateUrl.Append("','").Append(v_planlist)
        sb_wsNavigateUrl.Append("','").Append(v_planlist_txt).Append(" _ ").Append(Convert.ToString(dr("OrgName")))
        sb_wsNavigateUrl.Append("','").Append(Convert.ToString(dr("orgid")))
        sb_wsNavigateUrl.Append("','").Append(Convert.ToString(dr("isBlack"))) 'isBlack
        sb_wsNavigateUrl.Append("');")
        NewNode.NavigateUrl = sb_wsNavigateUrl.ToString()

        If ParentNode Is Nothing Then
            objTreeView.Nodes.Add(NewNode)
            NewNode.ToolTip = GET_ToolTip(dr)
        Else
            'ParentNode.Nodes.Add(NewNode)
            ParentNode.ChildNodes.Add(NewNode)
            NewNode.ToolTip = GET_ToolTip(dr)
        End If

        '加入子節點
        Dim strRid As String = Convert.ToString(dr("RID")) & "/"
        strFilter = "Relship like '%" & strRid & "%'"        '先找出符合父節點 xxx\ 開頭的關係
        For Each drChild In objTable.Select(strFilter, "OrgName")
            Dim strRelship As String = drChild("Relship")
            Dim pos As Integer = strRelship.IndexOf(strRid)
            '若出現格式為「%父節點/子節點/」的Relship值，視為子節點
            If pos <> -1 And (pos + strRid.Length) < strRelship.Length Then
                'If strRelship.IndexOf("/", pos + strRid.Length) = strRelship.Length - 1 Then AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                If strRelship.IndexOf("/", pos + strRid.Length) = strRelship.Length - 1 Then
                    Dim s_OrgName As String = Convert.ToString(drChild("OrgName"))
                    Dim s_lk_OrgName As String = txtSearch1.Text ' TIMS.ClearSQM(txtSearch1.Text)
                    If s_lk_OrgName <> "" AndAlso s_OrgName.Contains(s_lk_OrgName) Then
                        AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                    ElseIf s_lk_OrgName = "" Then
                        AddTreeNodes(drChild, objTable, objTreeView, NewNode)
                    End If
                End If

            End If
        Next
    End Sub

    Function GET_ToolTip(ByRef dr As DataRow) As String
        Dim sb_rst As New StringBuilder
        sb_rst.Append(String.Concat("ComIDNO=", dr("ComIDNO")))
        sb_rst.Append(String.Concat(",OrgID=", dr("OrgID")))
        sb_rst.Append(String.Concat(",RID=", dr("RID")))
        sb_rst.Append(String.Concat(",Relship=", dr("Relship")))
        Return sb_rst.ToString()
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim CookieTable_CanSave As Boolean = False

        OrgRID.Value = TIMS.ClearSQM(OrgRID.Value)
        CookieTable_CanSave = TIMS.CheckRIDsPLAN(Me, OrgRID.Value, objconn)  '確認RID 業務權限 與PLANID 計畫 為登入權限 sm.UserInfo.PlanID

        Dim ReqGetOther As String = TIMS.sUtl_GetRqValue(Me, "GetOther")

        If CookieTable_CanSave Then
            If ReqGetOther <> "1" Then
                dt = TIMS.GetCookieTable(Me, da, objconn)
                For i As Integer = 1 To 5
                    If dt.Select("ItemName='ActOrgRID" & i & "'").Length = 0 Then
                        Dim InsertFlag As Boolean = True
                        For j As Integer = 1 To 5
                            If dt.Select(String.Format("ItemName='ActOrgRID{0}' and ItemValue='{1}'", j, OrgRID.Value)).Length <> 0 Then
                                InsertFlag = False
                                Exit For
                            End If
                        Next
                        If InsertFlag = True Then
                            TIMS.InsertCookieTable(Me, dt, da, "ActOrgName" & i, OrgName.Value, False, objconn)
                            TIMS.InsertCookieTable(Me, dt, da, "ActOrgRID" & i, OrgRID.Value, True, objconn)
                        End If
                        Exit For
                    Else
                        If i = 5 Then
                            For j As Integer = 1 To 4
                                Dim s_filt1a As String = String.Format("ItemName='ActOrgRID{0}'", (j + 1))
                                Dim s_filt1b As String = String.Format("ItemName='ActOrgRID{0}'", j)
                                Dim NewDr As DataRow = dt.Select(s_filt1a)(0)
                                Dim OldDr As DataRow = dt.Select(s_filt1b)(0)
                                OldDr("ItemValue") = NewDr("ItemValue")

                                Dim s_filt2a As String = String.Format("ItemName='ActOrgName{0}'", (j + 1))
                                Dim s_filt2b As String = String.Format("ItemName='ActOrgName{0}'", j)
                                NewDr = dt.Select(s_filt2a)(0)
                                OldDr = dt.Select(s_filt2b)(0)
                                OldDr("ItemValue") = NewDr("ItemValue")
                            Next
                            Dim InsertFlag As Boolean = True
                            For j As Integer = 1 To 5
                                Dim s_filt1 As String = String.Format("ItemName='ActOrgRID{0}' and ItemValue='{1}'", j, OrgRID.Value)
                                If dt.Select(s_filt1).Length <> 0 Then
                                    InsertFlag = False
                                    Exit For
                                End If
                            Next
                            If InsertFlag = True Then
                                TIMS.InsertCookieTable(Me, dt, da, "ActOrgName" & i, OrgName.Value, False, objconn)
                                TIMS.InsertCookieTable(Me, dt, da, "ActOrgRID" & i, OrgRID.Value, True, objconn)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    ''' <summary>查詢計畫</summary>
    Sub SHOW_PLANLIST()
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If

        'Dim objreader As SqlDataReader
        'Dim sqlstr1 As String

        TreeView1.Visible = False

        Dim s_parms As New Hashtable
        s_parms.Clear()
        Dim s_sqlstr1 As String = ""
        s_sqlstr1 = "" & vbCrLf
        s_sqlstr1 &= " SELECT data1.* FROM (" & vbCrLf
        s_sqlstr1 &= " SELECT DISTINCT concat(a.Years,b.Name,c.PlanName,a.seq) PlanName ,a.PlanID ,a.DistID" & vbCrLf
        s_sqlstr1 &= " FROM ID_Plan a" & vbCrLf
        s_sqlstr1 &= " JOIN ID_District b ON a.DistID = b.DistID" & vbCrLf
        s_sqlstr1 &= " JOIN Key_Plan c ON a.TPlanID = c.TPlanID" & vbCrLf
        s_sqlstr1 &= " JOIN Auth_AccRWPlan d ON a.PlanID = d.PlanID" & vbCrLf
        s_sqlstr1 &= " WHERE 1=1" & vbCrLf

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist = "" Then
            '一定要選年度
            Me.planlist.Items.Clear()
            Me.planlist.Items.Insert(0, New ListItem("==請選擇==", ""))
            'Me.showcenter.Visible = False
            TreeView1.Visible = False
            Exit Sub
        End If

        Dim rqSAH As String = TIMS.sUtl_GetRqValue(Me, "SAH")
        Dim flag_SAH3 As Boolean = TIMS.IsSuperUser(sm, 3)
        Dim s_SAH_YN As String = If(flag_SAH3, rqSAH, "")

        If sm.UserInfo.RoleID = "0" Then '超級管理者不卡自己擁有什麼計畫
            If s_SAH_YN = "Y" Then
                '(跨轄區)
                s_parms.Add("YEARS", v_yearlist)
                s_sqlstr1 &= " AND a.YEARS=@YEARS" & vbCrLf
            Else
                '(限制轄區)
                s_parms.Add("YEARS", v_yearlist)
                s_parms.Add("DistID", sm.UserInfo.DistID)
                s_sqlstr1 &= " AND a.YEARS=@YEARS  AND a.DistID=@DistID" & vbCrLf
            End If
        Else
            '(限制轄區)
            s_parms.Add("YEARS", v_yearlist)
            s_parms.Add("DistID", sm.UserInfo.DistID)
            s_sqlstr1 &= " AND a.YEARS=@YEARS AND a.DistID=@DistID" & vbCrLf
            If cbplanlistAll.Visible AndAlso Not cbplanlistAll.Checked Then
                s_parms.Add("Account", sm.UserInfo.UserID)
                s_sqlstr1 &= " AND d.Account=@Account" & vbCrLf
            ElseIf Not cbplanlistAll.Visible Then
                s_parms.Add("Account", sm.UserInfo.UserID)
                s_sqlstr1 &= " AND d.Account=@Account" & vbCrLf
            End If
        End If
        'sqlstr1 += " order by 1 NULLS First" & vbCrLf
        s_sqlstr1 &= " ) data1 "
        s_sqlstr1 &= " ORDER BY (CASE WHEN data1.PlanName IS NULL THEN 0 ELSE 1 END), data1.PlanName "

        Dim dt1 As DataTable = DbAccess.GetDataTable(s_sqlstr1, objconn, s_parms)
        'objreader = DbAccess.GetReader(sqlstr1, objconn)
        With planlist
            .Items.Clear()
            .DataSource = dt1 'objreader
            .DataTextField = "PlanName"
            .DataValueField = "PlanID"
            .DataBind()

            .Items.Insert(0, New ListItem("==請選擇==", ""))
        End With
        'objreader.Close()

    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        Call SHOW_PLANLIST()
    End Sub

    Private Sub cbplanlistAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbplanlistAll.CheckedChanged
        Call SHOW_PLANLIST()
    End Sub

    '查詢機構
    Private Sub planlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles planlist.SelectedIndexChanged
        Call sSearch1()
    End Sub

    ''' <summary> 業務機構查詢 </summary>
    Sub sSearch1()
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If

        txtSearch1.Text = TIMS.ClearSQM(txtSearch1.Text)
        'Dim s_lk_OrgName As String = txtSearch1.Text

        Dim rqSAH As String = TIMS.sUtl_GetRqValue(Me, "SAH")
        Dim flag_SAH3 As Boolean = TIMS.IsSuperUser(sm, 3)
        Dim s_SAH_YN As String = If(flag_SAH3, rqSAH, "")

        Dim v_planlist As String = TIMS.GetListValue(planlist)
        If v_planlist = "" Then Exit Sub

        TreeView1.Visible = True
        Me.TreeView1.Nodes.Clear()

        Dim parms As Hashtable = New Hashtable()
        parms.Clear()
        parms.Add("PlanID", Val(v_planlist))

        Dim objstr As String = ""
        objstr = "" & vbCrLf
        objstr &= " SELECT a.OrgID, a.OrgKind, a.OrgName, a.ComIDNO, a.ComCIDNO" & vbCrLf
        objstr &= " ,a.IsConUnit, a.TradeID, a.EmpNum, a.OrgUrl, a.OrgKind2, a.LastYearExeRate" & vbCrLf
        objstr &= " ,a.IsConTTQS, a.BankName, a.ExBankName, a.AccNo, a.AccName" & vbCrLf
        objstr &= " ,b.RSID, b.PlanID, b.RID, b.Relship, b.OrgLevel, b.DistID ,'N' isBlack" & vbCrLf
        objstr &= " FROM ORG_ORGINFO a" & vbCrLf
        objstr &= " JOIN AUTH_RELSHIP b ON a.OrgID = b.OrgID and a.ORGID not in (200,201,164)" & vbCrLf
        objstr &= " WHERE (b.PlanID = 0 OR (b.PlanID=@PlanID AND b.PlanID <> 0))" & vbCrLf
        'If s_lk_OrgName <> "" Then objstr &= " AND a.OrgName LIKE '%'+@OrgName+'%' "
        'If s_lk_OrgName <> "" Then parms.Add("OrgName", s_lk_OrgName)

        Dim objdt As DataTable = DbAccess.GetDataTable(objstr, objconn, parms)

        If hid_ORGBLACKLIST.Value.Equals("Y") Then
            For Each odr As DataRow In objdt.Rows
                If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS<>'38'").Length > 0 Then
                    odr("isBlack") = "Y"
                Else
                    If dtOrgBlack.Select("isBlack='Y' AND ComIDNO='" & odr("ComIDNO") & "' AND OBTERMS='38' AND DistID='" & odr("DistID") & "'").Length > 0 Then
                        odr("isBlack") = "Y"
                    End If
                End If
            Next
            objdt.AcceptChanges()
        End If

        'objAdapter = New SqlDataAdapter(objstr, objconn)
        'objAdapter.Fill(objtable)

        Dim strFilter As String = If(s_SAH_YN = "Y", "1=1", String.Concat("RID='", sm.UserInfo.RID, "'"))

        For Each dr As DataRow In objdt.Select(strFilter)
            AddTreeNodes(dr, objdt, Me.TreeView1, Nothing)
        Next
        'treeview end 
        Call TreeView1.ExpandAll() '以程式設計方式展開節點
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Call sSearch1()
    End Sub
End Class