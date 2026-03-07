Partial Class SD_02_004_R
    Inherits AuthBasePage

    'Maintest_list.jrxml '(SD_02_004_R1 /單筆)
    'Maintest_list_org.jrxml 'Stud_EnterType
    'Maintest_list*.jrxml
    'http://tims2.turbotech.com.tw:8080/ReportServer/report.do?RptID=Maintest_list_org&OCID=57678&planid=2004&start_date=&end_date=&Mailtype1=0&Mailtype2=0&Mailtype3=0&Mailtype4=0&Mailtype5=0&chk1=1&chk2=1&chk3=1&chk4=1&chk5=1&UserID=rinc

    Const cst_reportFN1 As String = "Maintest_list_org2" '列印全部

    '報名管道
    Const cst_rw2不區分 As String = "A"
    Const cst_rw2一般推介單 As String = "CH4"
    Const cst_rw2免試推介單 As String = "EPW"
    Const cst_rw2專案核定報名 As String = "EP2P"

    '<asp:LinkButton ID = "Button2" runat="server" Text="列印全部" CommandName="all" CssClass="asp_Export_M"></asp:LinkButton>
    '<asp:LinkButton ID = "Button3" runat="server" Text="個別列印" CommandName="only" CssClass="asp_Export_M"></asp:LinkButton>
    '<asp:LinkButton ID = "btnExport1" runat="server" Text="匯出全部" CommandName="exp1" CssClass="asp_Export_M"></asp:LinkButton>
    Const cst_print_all As String = "all" '列印全部
    Const cst_print_only As String = "only" '個別列印
    Const cst_print_exp1 As String = "exp1" '匯出全部

    Const cst_SessSch As String = "SD_02_004_R_session"
    Const cst_errMsg1 As String = "准考證號碼有誤或重複，無法列印資料，請進入個別列印修正"

    Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
        blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        If blnP0 Then Trwork2013a.Visible = True

        ''就服單位協助報名
        'Trwork2013a.Visible = False
        'If sm.UserInfo.Years >= 2013 AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then Trwork2013a.Visible = True
        'End If

        '帶入查詢參數
        If Not IsPostBack Then
            'start_date.ReadOnly = True
            'end_date.ReadOnly = True
            PageControler1.Visible = False
            If Not Session(cst_SessSch) Is Nothing Then
                TMID1.Text = TIMS.GetMyValue(Session(cst_SessSch), "TMID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session(cst_SessSch), "TMIDValue1")
                OCID1.Text = TIMS.GetMyValue(Session(cst_SessSch), "OCID1")
                OCIDValue1.Value = TIMS.GetMyValue(Session(cst_SessSch), "OCIDValue1")
                start_date.Text = TIMS.GetMyValue(Session(cst_SessSch), "start_date")
                end_date.Text = TIMS.GetMyValue(Session(cst_SessSch), "end_date")
                If TIMS.GetMyValue(Session(cst_SessSch), "submit") = "1" Then
                    'Button1_Click(sender, e)
                    Call Search1()
                End If
                Session(cst_SessSch) = Nothing
            End If
        End If
        msg.Text = ""
        'Button1.Attributes("onclick") = "javascript:return search()"

        Me.table11.Style("display") = "none"

        If Not IsPostBack Then
            Me.start_date.Text = ""
            Me.end_date.Text = "" 'Common.FormatDate(Now())
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button7_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    '查詢
    Sub Search1()
        'Dim finddate = ""
        Call KeepSearch()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Me.table11.Style("display") = "none"
        Me.Item_Note.Visible = False
        Me.Button5.Visible = False

        Dim parms As Hashtable = New Hashtable()
        'Class_ClassInfo a
        Dim strWhere2 As String = "" '同時有2個where條件。請注意。
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0"
                strWhere2 &= " AND ip.TPlanID = @TPlanID " & vbCrLf '20090908 andy edit
                strWhere2 &= " AND ip.Years = @Years " & vbCrLf '20090908 andy edit
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
                parms.Add("Years", sm.UserInfo.Years)
            Case Else
                strWhere2 &= " AND a.PlanID = @PlanID " & vbCrLf '20090908 andy edit
                parms.Add("PlanID", sm.UserInfo.PlanID)
        End Select

        If OCIDValue1.Value <> "" Then
            strWhere2 &= " AND a.RID = @RID " & vbCrLf
            strWhere2 &= " AND a.OCID = @OCID " & vbCrLf
            parms.Add("RID", RIDValue.Value)
            parms.Add("OCID", OCIDValue1.Value)
        Else
            strWhere2 &= " AND a.RID LIKE @RID+'%' " & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If

        If start_date.Text <> "" Then
            'strWhere2 &= " AND a.STDate >= " & TIMS.to_date(start_date.Text) & vbCrLf
            strWhere2 &= " AND a.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", start_date.Text)
        End If
        If end_date.Text <> "" Then
            'strWhere2 &= " AND a.STDate <= " & TIMS.to_date(end_date.Text) & vbCrLf
            strWhere2 &= " AND a.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", end_date.Text)
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.PlanID " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassCName,a.cycltype) ClassCName" & vbCrLf
        sql &= " ,a.CyclType " & vbCrLf
        sql &= " ,a.LevelType " & vbCrLf
        sql &= " ,a.OCID " & vbCrLf
        sql &= " ,c.ClassID " & vbCrLf
        sql &= " ,g.TOTAL " & vbCrLf
        sql &= " ,ip.DISTID " & vbCrLf '2018-09-07 add 查轄區代碼
        'PSMEMO1 sql &= " ,CONVERT(varchar, NULL) PSMEMO1" & vbCrLf
        sql &= " FROM Class_ClassInfo a " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = a.planid " & vbCrLf
        sql &= " JOIN ID_Class c ON a.CLSID = c.CLSID " & vbCrLf
        'SELECT ENTERPATH FROM STUD_ENTERTYPE WHERE ROWNUM <=10

        'GROUP 1
        sql &= " JOIN ( " & vbCrLf
        sql &= "  SELECT a.OCID ,COUNT(1) TOTAL " & vbCrLf
        sql &= "  FROM Class_ClassInfo a " & vbCrLf
        sql &= "  JOIN ID_PLAN ip ON ip.planid = a.planid " & vbCrLf
        sql &= "  JOIN STUD_ENTERTYPE b ON b.OCID1 = a.OCID " & vbCrLf
        sql &= "  JOIN STUD_ENTERTEMP b2 ON b2.SETID = b.SETID " & vbCrLf
        sql &= "  WHERE 1=1 " & vbCrLf

        'Select Case rblEnterPathW2.SelectedValue
        '    Case cst_rw2不區分
        '    Case cst_rw2一般推介單
        '        sql &= " AND dbo.NVL(b.ENTERCHANNEL,0)=4" & vbCrLf
        '        sql &= " AND dbo.NVL(b.ENTERPATH,' ')!='W'" & vbCrLf '排除免試
        '        sql &= " AND dbo.NVL(b.ENTERPATH2,' ')!='P'" '專案核定
        '    Case cst_rw2免試推介單
        '        sql &= " AND dbo.NVL(b.ENTERPATH,' ')='W'" & vbCrLf
        '    Case cst_rw2專案核定報名
        '        sql &= " AND dbo.NVL(b.ENTERPATH2,' ')='P'" & vbCrLf
        'End Select

        Select Case rblEnterPathW.SelectedValue
            Case "Y"
                sql &= " AND b.EnterPath = 'W' " & vbCrLf
            Case "N"
                sql &= " AND ISNULL(b.EnterPath,' ') != 'W' " & vbCrLf
        End Select
        sql &= strWhere2
        sql &= " GROUP BY a.OCID " & vbCrLf
        sql &= " ) g ON g.OCID = a.OCID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= strWhere2
        sql &= " ORDER BY c.ClassID, a.CyclType " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGrid1.Visible = True
            PageControler1.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
        'Else
        '    Common.MessageBox(Me, "請先設定甄試通知單內容!!")
        'End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1()
    End Sub

    '(search條件) 組合chk字串
    Function Get_CBListInfo() As String
        Dim sStr As String = ""
        Const cst_Title As String = "chk"
        For i As Integer = 0 To Me.cblist_info.Items.Count - 1
            If sStr <> "" Then sStr += "&"
            If cblist_info.Items(i).Selected Then
                sStr &= cst_Title & (i + 1) & "=1"
            Else
                sStr &= cst_Title & (i + 1) & "=0"
            End If
        Next
        Return sStr
    End Function

    '(DataGrid1) 組合Mailtype字串
    Function Get_DG1CheckList1(ByVal ocid As String) As String
        Dim sRst As String = ""
        Const cst_Title As String = "Mailtype"
        For Each eItem As DataGridItem In DataGrid1.Items
            'Dim drv As DataRowView = Item.DataItem
            Dim Mailtype1 As CheckBoxList = eItem.FindControl("Mailtype1")
            Dim hidOCID As HtmlInputHidden = eItem.FindControl("hidOCID")
            If hidOCID.Value = ocid Then
                Dim sStr As String = ""
                'sStr = "OCID=" & hidOCID.Value
                sStr = ""
                For i As Integer = 0 To Mailtype1.Items.Count - 1
                    If sStr <> "" Then sStr += "&"
                    If Mailtype1.Items(i).Selected Then
                        sStr &= cst_Title & (i + 1) & "=1"
                    Else
                        sStr &= cst_Title & (i + 1) & "=0"
                    End If
                Next
                sRst = sStr
                Exit For
            End If
        Next
        Return sRst
    End Function

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Dim SMpath As String = ReportQuery.GetSmartQueryPath
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim sOCID1 As String = TIMS.GetMyValue(sCmdArg, "OCID")
        If sOCID1 = "" Then Exit Sub
        Dim strCBListInfo As String = Get_CBListInfo()
        Dim strMailType As String = Get_DG1CheckList1(sOCID1)
        'Dim strScript As String = ""
        Select Case e.CommandName
            Case cst_print_exp1 '"exp1"
                Dim sMyValue As String = ""
                sMyValue = ""
                sMyValue += sCmdArg
                sMyValue += "&" & strMailType
                sMyValue += "&" & strCBListInfo

                ''排除就服單位協助報名
                'If rblEnterPathNW.SelectedValue = "Y" Then
                '    ''Y:排除 N:不排除 才須加入此一條件
                '    MyValue += "&EnterPathNW=W" 'Y:排除 N:不排除
                'End If

                '就服單位協助報名
                Select Case rblEnterPathW.SelectedValue
                    Case "A"
                    Case "Y", "N"
                        sMyValue &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
                End Select
                'Select Case rblEnterPathW2.SelectedValue
                '    Case cst_rw2不區分
                '    Case cst_rw2一般推介單
                '        sMyValue &= "&ENTERPATH1=Y"
                '    Case cst_rw2免試推介單
                '        sMyValue &= "&ENTERPATH2=Y"
                '    Case cst_rw2專案核定報名
                '        sMyValue &= "&ENTERPATH3=Y"
                'End Select

                Dim dt As DataTable = Nothing
                dt = TIMS.LoadDatab39(1, sMyValue, objconn)
                If dt.Rows.Count = 0 Then
                    Common.MessageBox(Me, "查無資料!!")
                    Exit Sub
                End If
                Call TIMS.ExpRptb39(Me, objconn, dt)

            Case cst_print_all '"all" 'Button2 'member 'Maintest_list_org
                'openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=Maintest_list_org&path=' + SMpath + CommandArgument + '&DistID=' + DistID + '&Mailtype1=' + M1 + '&Mailtype2=' + M2 + '&Mailtype3=' + M3 + '&Mailtype4=' + M4  + '&Mailtype5=' + M5  + document.getElementById('chkvalue').value ); 
                Dim MyValue As String = ""
                MyValue = ""
                MyValue += sCmdArg
                MyValue += "&" & strMailType
                MyValue += "&" & strCBListInfo

                ''排除就服單位協助報名
                'If rblEnterPathNW.SelectedValue = "Y" Then
                '    ''Y:排除 N:不排除 才須加入此一條件
                '    MyValue += "&EnterPathNW=W" 'Y:排除 N:不排除
                'End If

                '就服單位協助報名
                Select Case rblEnterPathW.SelectedValue
                    Case "A"
                    Case "Y", "N"
                        MyValue &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
                End Select
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN1, MyValue)

                'strScript = "<script language=""javascript"">"
                'strScript += "print_rpt('" & SMpath & "','" & e.CommandArgument & "','" & Me.sm.UserInfo.DistID & "');"
                'strScript += "</script>"
                'Page.RegisterStartupScript("PrintRepost", strScript)

            Case cst_print_only '"only" 'Button3
                'Dim strCBListInfo As String = Get_CBListInfo()
                'Dim strMailType As String = Get_DG1CheckList1(TIMS.GetMyValue(e.CommandArgument, "ocid"))
                Dim sUrl As String = ""
                sUrl = "SD_02_004_R1.aspx?ID=" & Request("ID")
                sUrl += e.CommandArgument
                sUrl += "&" & strMailType
                sUrl += "&" & strCBListInfo

                ''排除就服單位協助報名
                'sUrl += "&EnterPathNW=" & rblEnterPathNW.SelectedValue 'Y:排除 N:不排除
                '就服單位協助報名
                Select Case rblEnterPathW.SelectedValue
                    Case "A"
                    Case "Y", "N"
                        sUrl &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
                End Select
                'Response.Redirect(sUrl)
                'Dim url1 As String = ""
                Call TIMS.Utl_Redirect(Me, objconn, sUrl)

                ''='SD_02_004_R1.aspx?' + CommandArgument + '&ID='+ID+ '&Mailtype1=' + M1+ '&Mailtype2=' + M2 + '&Mailtype3=' + M3+ '&Mailtype4=' + M4 + '&Mailtype5=' + M5+ document.getElementById('chkvalue').value ;
                'strScript = "<script language=""javascript"">" + vbCrLf
                'strScript += "RedictPage('" & e.CommandArgument & "','" & Request("ID") & "');"
                'strScript += "</script>"
                'Page.RegisterStartupScript("PrintRepost", strScript)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Button2 As LinkButton = e.Item.FindControl("Button2") '列印全部
                Dim Button3 As LinkButton = e.Item.FindControl("Button3") '個別列印
                Dim btnExport1 As LinkButton = e.Item.FindControl("btnExport1") '匯出全部
                Dim Mailtype1 As CheckBoxList = e.Item.FindControl("Mailtype1") '增加郵寄型態
                Dim hidOCID As HtmlInputHidden = e.Item.FindControl("hidOCID")
                hidOCID.Value = Convert.ToString(drv("OCID"))

                'https://jira.turbotech.com.tw/browse/TIMSC-219
                '推介資料註銷處理
                'Dim LPSMEMO1 As Label = e.Item.FindControl("LPSMEMO1") '備註
                'LPSMEMO1.Text = TIMS.Get_TICKETPS1(objconn, Convert.ToString(drv("OCID")))
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                'If CInt(Val(e.Item.Cells(6).Text)) <> 0 Then e.Item.Cells(1).Text += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(6).Text)) & "期"
                'Mailtype1.Attributes("onclick") = "return checkselect('" & Mailtype1.ClientID & "');"
                'mybut1.Attributes("onclick") =     ReportQuery.ReportScript(Me, "list", "Maintest_list", "OCID1=" & DataGrid1.DataKeys(e.Item.ItemIndex) & "&DistID=" & sm.UserInfo.DistID & "&Mailtype1=" & M1.Value & "&Mailtype2=" & M2.Value & "&Mailtype3=" & M3.Value & "&Mailtype4=" & M4.Value & "&Mailtype5=" & M5.Value)
                'If TIMS.CheckDblExamNo(Convert.ToString(drv("OCID")), "", objconn) Then
                '    Button2.Enabled = False
                '    TIMS.Tooltip(Button2, "准考證號碼有誤或重複，無法列印資料，請進入個別列印修正")
                'End If
                'Dim dateParment As String = ""
                'If start_date.Text <> "" Then dateParment += "&start_date=" & (Convert.ToDateTime(start_date.Text)).ToString("yyyy-MM-dd")
                'If end_date.Text <> "" Then dateParment += "&end_date=" & (Convert.ToDateTime(end_date.Text)).ToString("yyyy-MM-dd")
                'Dim XXX As String = GetParentOrgID(Convert.ToString(drv("OCID")))
                'Common.RespWrite(Me, XXX)
                'Response.End()

                If TIMS.CheckDblExamNo(Convert.ToString(drv("OCID")), "", objconn) Then
                    Button2.Enabled = False
                    TIMS.Tooltip(Button2, cst_errMsg1)
                    btnExport1.Enabled = False
                    TIMS.Tooltip(btnExport1, cst_errMsg1)
                End If

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "planid", Convert.ToString(drv("planid")))
                TIMS.SetMyValue(sCmdArg, "distid", Convert.ToString(drv("distid")))
                TIMS.SetMyValue(sCmdArg, "start_date", TIMS.Cdate3(start_date.Text))
                TIMS.SetMyValue(sCmdArg, "end_date", TIMS.Cdate3(end_date.Text))

                If Button2.Enabled Then Button2.CommandArgument = sCmdArg
                If Button3.Enabled Then Button3.CommandArgument = sCmdArg
                If btnExport1.Enabled Then btnExport1.CommandArgument = sCmdArg
                'Button2.CommandArgument = "&OCID=" & Convert.ToString(drv("OCID")) & "&planid=" & Convert.ToString(drv("planID")) & dateParment '& "&ParentOrgID=" & GetParentOrgID(Convert.ToString(drv("OCID"))) & "" '20090908 andy edit 
                'Button3.CommandArgument = "&OCID=" & Convert.ToString(drv("OCID")) & "&planid=" & Convert.ToString(drv("planID")) & dateParment '& "&ParentOrgID=" & GetParentOrgID(Convert.ToString(drv("OCID"))) & "" '20090908 andy edit 
        End Select
    End Sub

    '設定通知單內容
    Sub sUtl_SetNoticeContent()
        Me.table11.Style("display") = "inline"
        Me.Item_Note.Visible = True
        Me.DataGrid1.Visible = False
        Me.Button5.Visible = True
        Dim sql As String = ""

        '將參數設定-甄試成績內容代入----start
        sql = " SELECT * FROM SYS_ORGVAR WHERE RID = '" & sm.UserInfo.RID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow = dt.Rows(0)
            If Not IsDBNull(dr("ItemVar")) Then Me.Item_Note.Text = Convert.ToString(dr("ItemVar"))
        Else
            sql = " SELECT * FROM SYS_ORGVAR WHERE RID = '" & sm.UserInfo.RID & "'  AND TPlanID IS NULL "
            dt = DbAccess.GetDataTable(sql, objconn)

            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(0)
                If Not IsDBNull(dr("ItemVar")) Then Me.Item_Note.Text = Convert.ToString(dr("ItemVar"))
            Else
                sql = " SELECT * FROM SYS_GLOBALVAR WHERE GVID = '6' AND DistID = '" & sm.UserInfo.DistID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count <> 0 Then
                    Dim dr As DataRow = dt.Rows(0)
                    If Not IsDBNull(dr("ItemVar1")) Then Me.Item_Note.Text = Convert.ToString(dr("ItemVar1"))
                End If
            End If
        End If
        '---end
    End Sub

    '設定通知單內容
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call sUtl_SetNoticeContent()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing

        sql = " SELECT * FROM SYS_ORGVAR WHERE RID = '" & sm.UserInfo.RID & "' AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0)
            dr("ItemVar") = Me.Item_Note.Text.ToString
            dr("TPlanID") = sm.UserInfo.TPlanID
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now()
        Else
            dr = dt.NewRow()
            dt.Rows.Add(dr)
            dr("RID") = sm.UserInfo.RID
            dr("ItemVar") = Me.Item_Note.Text.ToString
            dr("TPlanID") = sm.UserInfo.TPlanID
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now()
        End If
        DbAccess.UpdateDataTable(dt, da)

        Me.Button5.Visible = False
        Common.MessageBox(Me, "儲存成功!!")
    End Sub

    Sub KeepSearch()
        Session(cst_SessSch) = ""
        Session(cst_SessSch) += "&TMID1=" & TMID1.Text
        Session(cst_SessSch) += "&TMIDValue1=" & TMIDValue1.Value
        Session(cst_SessSch) += "&OCID1=" & OCID1.Text
        Session(cst_SessSch) += "&OCIDValue1=" & OCIDValue1.Value
        Session(cst_SessSch) += "&start_date=" & start_date.Text
        Session(cst_SessSch) += "&end_date=" & end_date.Text
        If DataGrid1.Visible Then
            Session(cst_SessSch) += "&submit=1"
        Else
            Session(cst_SessSch) += "&submit=0"
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGrid1.Visible = False
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGrid1.Visible = False
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub
End Class