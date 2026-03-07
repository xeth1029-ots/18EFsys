Partial Class TC_03_002
    Inherits AuthBasePage

    '班級複製
    'Dim Auth_Relship As DataTable

    Const Cst_序號 As Integer = 0
    Const Cst_計畫年度 As Integer = 1
    Const Cst_申請日期 As Integer = 2
    Const Cst_訓練起日 As Integer = 3
    Const Cst_訓練迄日 As Integer = 4
    Const Cst_機構名稱 As Integer = 5
    Const Cst_班別名稱 As Integer = 6

    'Const Cst_訓練性質ID = 7
    'Const Cst_訓練性質Name = 8
    Const Cst_學分班 As Integer = 7 '9

    Const cst_copy_tb_TV As String = "1:課程表,2:材料明細"

    Dim vsMsg2 As String = "" '確認機構是否為黑名單

    Const cst_ccopy As String = "ccopy" 'Request(cst_ccopy)
    Const cst_ccopy1 As String = "&ccopy=1" 'Request(cst_ccopy)

    'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' Common.RespWrite(Me, Session("search"))
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        'iPYNum = TIMS.sUtl_GetPYNum(Me)

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Me.LabTMID.Text = "訓練業別"

        trddlDIST.Visible = False
        Select Case sm.UserInfo.LID
            Case 2 '委訓單位，可跨轄區
                trddlDIST.Visible = True
                Org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
            Case Else
                Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
                Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), TIMS.GetListValue(PlanYear))
        End Select

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            cCreate1()
        End If


        'Dim sql As String = ""
        'sql = "SELECT a.RID,b.OrgName FROM Auth_Relship a,Org_OrgInfo b WHERE a.OrgID=b.OrgID"
        'Auth_Relship = DbAccess.GetDataTable(sql, objconn)

        '確認機構是否為黑名單
        'Dim vsMsg2 As String = ""'確認機構是否為黑名單
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            'Button2.Enabled = False 'TIMS.Tooltip(Button2, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If

    End Sub

    Sub cCreate1()
        trCOPYSUB.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trCOPYSUB.Visible = True
            Dim copy_tb_ARY As String() = cst_copy_tb_TV.Split(",")

            '1:課程表 /2:材料明細
            With CBL_COPYSUB
                .Items.Clear()
                For Each s_TV As String In copy_tb_ARY
                    .Items.Add(New ListItem(s_TV.Split(":")(1), s_TV.Split(":")(0)))
                Next
                '.Items.Add(New ListItem("課程表", "1"))
                '.Items.Add(New ListItem("材料明細", "2"))
            End With
            '1:課程表 /2:材料明細
            TIMS.SetCblValue(CBL_COPYSUB, "1,2")
        End If

        msg.Text = ""
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        PlanYear = TIMS.GetSyear(PlanYear)
        'Common.SetListItem(PlanYear, Now.Year)
        Common.SetListItem(PlanYear, sm.UserInfo.Years)
        ddlDIST = TIMS.Get_DistID(ddlDIST, TIMS.dtNothing(), objconn)
        Common.SetListItem(ddlDIST, sm.UserInfo.DistID)

        If Session("search") IsNot Nothing Then
            Dim MyValue As String = ""
            Dim strSession As String = ""
            strSession = Session("search")

            MyValue = TIMS.GetMyValue(strSession, "PlanYear")
            If MyValue <> "" Then Common.SetListItem(PlanYear, MyValue)
            center.Text = TIMS.GetMyValue(strSession, "center")
            RIDValue.Value = TIMS.GetMyValue(strSession, "RIDValue")
            TB_career_id.Text = TIMS.GetMyValue(strSession, "TB_career_id")
            trainValue.Value = TIMS.GetMyValue(strSession, "trainValue")
            txtCJOB_NAME.Text = TIMS.GetMyValue(strSession, "txtCJOB_NAME")
            cjobValue.Value = TIMS.GetMyValue(strSession, "cjobValue")
            ClassName.Text = TIMS.GetMyValue(strSession, "ClassName")
            CyclType.Text = TIMS.GetMyValue(strSession, "CyclType")
            CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
            MyValue = TIMS.GetMyValue(strSession, "PageIndex")
            If MyValue <> "" Then PageControler1.PageIndex = MyValue
            If MyValue <> "" Then
                'Button2_Click(sender, e)
                Call sSearch1()
            End If
            Session("search") = Nothing
        End If
    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            Me.isBlack.Value = "Y"
            Me.Blackorgname.Value = sm.UserInfo.OrgName
        End If
        Return rst
    End Function

    '查詢
    Sub sSearch1()
        'Dim sRelship As String = ""
        'Dim sOrgid As String = ""
        'Dim sSearchStr As String = ""
        'Dim str2 As String = ""
        'Dim sql As String = ""
        'hidTC03002PlanID.Value = TIMS.ClearSQM(hidTC03002PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "訓練機構資料異常，請重新選擇訓練機構!!")
            Exit Sub
        End If
        Dim sqlR As String = "SELECT Relship,OrgID FROM AUTH_RELSHIP WHERE RID='" & RIDValue.Value & "'"
        Dim dr99 As DataRow = DbAccess.GetOneRow(sqlR, objconn)
        If dr99 Is Nothing Then
            Common.MessageBox(Me, "訓練機構資料異常，請重新選擇訓練機構!!")
            Exit Sub
        End If
        Dim sRelship As String = Convert.ToString(dr99("Relship"))
        Dim sOrgid As String = Convert.ToString(dr99("OrgID"))
        If sRelship = "" OrElse sOrgid = "" Then
            Common.MessageBox(Me, "訓練機構資料異常，請重新選擇訓練機構!!")
            Exit Sub
        End If

        'Dim flagTC03002PlanID As Boolean = False
        'If hidTC03002PlanID.Value <> "" Then
        '    '自辦計畫 2013年後 非署(局) 【自辦應該是分署(中心)】
        '    If sm.UserInfo.RID <> "A" AndAlso sm.UserInfo.TPlanID = "02" _
        '        AndAlso sm.UserInfo.Years >= "2013" Then
        '        flagTC03002PlanID = True
        '    End If
        'End If

        'SQL
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT a.PlanID" & vbCrLf
        sql &= " ,a.ComIDNO" & vbCrLf
        sql &= " ,a.SeqNo" & vbCrLf
        sql &= " ,a.PlanYear" & vbCrLf
        sql &= " ,a.STDate" & vbCrLf
        sql &= " ,a.FDDate" & vbCrLf
        sql &= " ,a.ClassName" & vbCrLf
        sql &= " ,a.CyclType" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,a.AppliedDate" & vbCrLf
        sql &= " ,c.OrgName" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,b.Relship" & vbCrLf
        sql &= " ,a.PointYN" & vbCrLf
        'sql += " ,a.ProcID, isnull(kcp.ProcName,'未選擇') ProcName " & vbcrlf
        sql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP b ON a.RID=b.RID " & vbCrLf
        sql &= " JOIN dbo.ID_PLAN ip on ip.PlanID=a.PlanID " & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO c ON b.OrgID=c.OrgID " & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_TRAINTYPE tt ON tt.TMID=a.TMID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'If flagTC03002PlanID Then
        '    '自辦計畫 2013年後 非署(局) 【自辦應該是分署(中心)】
        '    sql &= " AND ip.PlanID='" & hidTC03002PlanID.Value & "'"
        '    sql &= " AND a.IsApprPaper='Y'" & vbCrLf
        'Else
        '    sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        '    sql &= " AND a.IsApprPaper='Y'" & vbCrLf
        'End If

        sql &= " AND a.IsApprPaper='Y'" & vbCrLf
        Dim sel_DISTID As String = ""

        '階層代碼0:署(局) 1:分署(中心) 2:委訓 【SELECT LID ,COUNT(1) CNT FROM AUTH_ACCOUNT GROUP BY LID ORDER BY 1】
        Select Case sm.UserInfo.LID
            Case 0
                '跨轄區、同計畫、同機構的複製 
                sel_DISTID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                If RIDValue.Value <> sm.UserInfo.RID Then
                    sel_DISTID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    sql &= " AND ip.DistID ='" & sel_DISTID & "'" & vbCrLf
                Else
                    sql &= " AND ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                End If
                sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                If Len(RIDValue.Value) > 1 Then
                    '委訓單位
                    sql &= " and c.ORGID = '" & sOrgid & "'" & vbCrLf
                Else
                    sql &= " and b.Relship like '" & sRelship & "%'" & vbCrLf
                End If

            Case 1
                '跨轄區、同計畫、同機構的複製 
                If RIDValue.Value <> sm.UserInfo.RID Then
                    sel_DISTID = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    sql &= " AND ip.DistID ='" & sel_DISTID & "'" & vbCrLf
                Else
                    sql &= " AND ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf
                End If
                sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                'sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                If Len(RIDValue.Value) > 1 Then
                    '委訓單位
                    sql &= " and c.ORGID = '" & sOrgid & "'" & vbCrLf
                Else
                    sql &= " and b.Relship like '" & sRelship & "%'" & vbCrLf
                End If

            Case Else
                '跨轄區
                Dim v_DISTID As String = TIMS.ClearSQM(ddlDIST.SelectedValue)
                If v_DISTID = "" Then v_DISTID = sm.UserInfo.DistID
                '同轄區、同計畫、同機構的複製 
                sql &= " AND ip.DistID='" & v_DISTID & "'" & vbCrLf
                sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                'sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                '委訓單位
                sql &= " AND c.ORGID = '" & sm.UserInfo.OrgID & "'" & vbCrLf
        End Select
        If PlanYear.SelectedIndex <> 0 Then
            sql &= " AND ip.Years='" & PlanYear.SelectedValue & "'" & vbCrLf
        Else
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        End If

        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            sql &= " AND a.CJOB_UNKEY='" & cjobValue.Value & "'" & vbCrLf
        End If

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Me.LabTMID.Text = "訓練業別"
            'If iPYNum >= 3 Then
            '    If trainValue.Value <> "" Then
            '        sql &= " and a.TMID = " & Me.trainValue.Value & vbCrLf
            '    End If
            'Else
            '    If jobValue.Value <> "" Then
            '        sql &= " and (1!=1"
            '        sql &= " OR tt.JOBTMID = " & jobValue.Value & vbCrLf
            '        sql &= " OR tt.TMID = " & jobValue.Value & vbCrLf
            '        sql &= " )"
            '    End If
            'End If
            If trainValue.Value <> "" Then
                sql &= " and a.TMID = " & Me.trainValue.Value & vbCrLf
            End If
        Else
            If trainValue.Value <> "" Then
                sql &= " and a.TMID = " & Me.trainValue.Value & vbCrLf
            End If
        End If
        'If trainValue.Value <> "" Then
        '    SearchStr += " and TMID='" & trainValue.Value & "'"
        'End If
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then
            sql &= " and a.ClassName like '%" & ClassName.Text & "%'" & vbCrLf
        End If
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If CInt(Val(CyclType.Text)) < 10 Then
                    sql &= " and a.CyclType='0" & CInt(Val(CyclType.Text)) & "'" & vbCrLf
                Else
                    sql &= " and a.CyclType='" & CyclType.Text & "'" & vbCrLf
                End If
            End If
        End If

        'If str2 <> "" Then sql += str2
        'If TIMS.sUtl_ChkTest() Then sql = cls_test.test2_Sql4() '測試用

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        msg.Text = "查無資料"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.Sort = "STDate,ClassName"
            PageControler1.ControlerLoad()
        End If
    End Sub

    '查詢
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call sSearch1()
    End Sub

    Sub KEEP_SEARCH(ByRef sCOPYSUB As String)
        Dim ssSearch1 As String = ""
        ssSearch1 = "PlanYear=" & TIMS.ClearSQM(PlanYear.SelectedValue)
        ssSearch1 &= "&center=" & TIMS.ClearSQM(center.Text)
        ssSearch1 &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        ssSearch1 &= "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        ssSearch1 &= "&trainValue=" & TIMS.ClearSQM(trainValue.Value)
        ssSearch1 &= "&txtCJOB_NAME=" & TIMS.ClearSQM(txtCJOB_NAME.Text)
        ssSearch1 &= "&cjobValue=" & TIMS.ClearSQM(cjobValue.Value)
        ssSearch1 &= "&ClassName=" & TIMS.ClearSQM(TB_career_id.Text)
        ssSearch1 &= "&CyclType=" & TIMS.ClearSQM(CyclType.Text)
        If (sCOPYSUB <> "") Then ssSearch1 &= sCOPYSUB
        ssSearch1 &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1

        Session("search") = ssSearch1
    End Sub

    Sub COPY_DATA1(ByRef sCmdArg As String)
        Dim sCOPYSUB As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'COPYSUB 1:課程表 /2:材料明細
            Dim fsCOPYSUB As String() = TIMS.GetCblValue(CBL_COPYSUB).Split(",")
            Dim fsCOPYSUBtxt As String = TIMS.GetCblText(CBL_COPYSUB)
            For Each sV As String In fsCOPYSUB
                If sV <> "" Then sCOPYSUB &= String.Format("&COPYSUB{0}=Y", sV)
            Next
        End If

        KEEP_SEARCH(sCOPYSUB)

        'Common.RespWrite(Me, Session("search"))
        Dim sPlanYear28 As String = "&PlanYear=" & PlanYear.SelectedValue
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Dim sUrl1 As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用
            sUrl1 = "TC_03_006.aspx?ID=" & rqMID
            sUrl1 &= sPlanYear28 'PlanYear
            sUrl1 &= sCOPYSUB
            sUrl1 &= cst_ccopy1
            sUrl1 &= sCmdArg 'e.CommandArgument
        Else
            sUrl1 = "TC_03_001.aspx?ID=" & rqMID 'Request("ID")
            sUrl1 &= cst_ccopy1
            sUrl1 &= sCmdArg 'e.CommandArgument
        End If
        TIMS.Utl_Redirect(Me, objconn, sUrl1)
    End Sub
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Dim PlanYear28 As String
        'PCS PYN
        If e.CommandArgument = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Select Case e.CommandName
            Case "copy" '複製
                If isBlack.Value = "Y" Then
                    Common.MessageBox(Me, vsMsg2)
                    Return 'Exit Sub
                End If
                If e.CommandArgument = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Return
                End If
                Dim sCmdArg As String = e.CommandArgument
                COPY_DATA1(sCmdArg)
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Me.ViewState("sort") <> "" Then
                    'Dim mylabel As String
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "PlanYear", "PlanYear DESC"
                            i = Cst_計畫年度
                            mysort.ImageUrl = If(ViewState("sort") = "PlanYear", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "AppliedDate", "AppliedDate DESC"
                            'mylabel = "ComName"
                            i = Cst_申請日期
                            mysort.ImageUrl = If(ViewState("sort") = "AppliedDate", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "STDate", "STDate DESC"
                            'mylabel = "ComName"
                            i = Cst_訓練起日
                            mysort.ImageUrl = If(ViewState("sort") = "STDate", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "FDDate", "FDDate DESC"
                            'mylabel = "ComName"
                            i = Cst_訓練迄日
                            mysort.ImageUrl = If(ViewState("sort") = "FDDate", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "OrgName", "OrgName DESC"
                            'mylabel = "ComName"
                            i = Cst_機構名稱
                            mysort.ImageUrl = If(ViewState("sort") = "OrgName", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "ClassCName", "ClassCName DESC"
                            'mylabel = "ComName"
                            i = Cst_班別名稱
                            mysort.ImageUrl = If(ViewState("sort") = "ClassCName", "../../images/SortUp.gif", "../../images/SortDown.gif")
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex

                Dim ParentName As String = TIMS.Get_ParentRID(drv("Relship"), objconn)
                If ParentName <> "" Then
                    Dim sParentOrg As String = "<font color='Blue'>" & ParentName & "</font>"
                    e.Item.Cells(Cst_機構名稱).Text = sParentOrg & "-" & Convert.ToString(drv("OrgName"))
                End If

                'If IsNumeric(drv("CyclType")) Then
                '    If Int(drv("CyclType")) <> 0 Then
                '        e.Item.Cells(Cst_班別名稱).Text += "第" & Int(drv("CyclType")) & "期"
                '    End If
                'End If
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "ComIDNO", Convert.ToString(drv("ComIDNO")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
                TIMS.SetMyValue(sCmdArg, "PointYN", Convert.ToString(drv("PointYN")))

                Dim Button3 As LinkButton = e.Item.FindControl("Button3")
                Button3.CommandArgument = sCmdArg
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If
        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    Protected Sub PlanYear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlanYear.SelectedIndexChanged

    End Sub

    Protected Sub ddlDIST_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlDIST.SelectedIndexChanged

    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
