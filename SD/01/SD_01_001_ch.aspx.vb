Partial Class SD_01_001_ch
    Inherits AuthBasePage

#Region "參數/變數 設定"

    'Const cst_IJC As String = "民眾-XXX的參訓資格，因與委外實施基準條款有抵觸，請確認是否要同意此民眾的報名?"
    Const Cst_session1 As String = "SD_01_001_ch_ClassSort"
    Const Cst_rqwish As String = "wish" 'Request(Cst_rqwish)
    Const Cst_rqStudIDNO As String = "StudIDNO" 'Request(Cst_rqStudIDNO)
    Const cst_TMsg1 As String = "課程為多選，已協助查詢。不可再使用。"
    Const cst_AlertMsg1 As String = "依計畫機構及查詢條件，查無開班資料!!"
    Const cst_AlertMsg2 As String = "請選擇志願班級!"
    Const cst_AlertMsg3 As String = "只能勾選兩個班級!"
    Const cst_AlertMsg4 As String = "只能勾選一個志願班級!"
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    Dim objconn As SqlConnection

#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
#Region "在這裡放置使用者程式碼以初始化網頁"

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '是否為超級使用者
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1)
        Dim flag_ChkTest As Boolean = TIMS.sUtl_ChkTest() '測試

        trCenter.Visible = False
        If flgROLEIDx0xLIDx0 Then trCenter.Visible = True

        StudIDNO.Value = Request(Cst_rqStudIDNO)
        StudIDNO.Value = TIMS.ClearSQM(StudIDNO.Value)
        Dim rqWish As String = Request(Cst_rqwish) 'wish 1:單選／2:多選
        rqWish = TIMS.ClearSQM(rqWish)
        'StudBirth.Value = Request("StudBirth")

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DataGridTable1.Visible = False
            DataGridTable2.Visible = False
            msg.Text = ""
            search_but.Attributes("onclick") = "javascript:return chkdata();"

            'wish 1:單選／2:多選
            If rqWish <> 1 Then
                search_but.Enabled = False
                TIMS.Tooltip(search_but, cst_TMsg1)
                Call sSearch1()
                'search_but_Click(sender, e)
            End If

            send.Attributes("onclick") = "return CheckData(1);"
            Button1.Attributes("onclick") = "return CheckData(2);"

            Select Case Convert.ToString(Session(Cst_session1))
                Case "1", "2", "3", "4", "5"
                    Common.SetListItem(ClassSort, Session(Cst_session1))
                    Session(Cst_session1) = Nothing
            End Select

            Call Chk_OrgBlackList()  '確認登入帳號之機構是否在黑名單中
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

#End Region
    End Sub

    Function Chk_OrgBlackList() As Boolean
#Region "確認登入帳號之機構是否在黑名單中"

        Dim rst As Boolean = False '若為黑名單為True 不是為False;
        'Chk_OrgBlackList = False
        '若有session消失離開此搜尋
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Return True
        'If Convert.ToString(sm.UserInfo.OrgID) = "" Then Return True
        'If Convert.ToString(sm.UserInfo.OrgName) = "" Then Return True
        Me.ViewState("ComIDNO") = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)

        If TIMS.Check_OrgBlackList(Me, Me.ViewState("ComIDNO"), objconn) Then
            'Chk_OrgBlackList = True
            rst = True
            'sm.UserInfo.OrgName
            Me.ViewState("msg2") = sm.UserInfo.OrgName & "，已列入處分名單!!"
            send.Attributes.Remove("onclick")
            Button1.Attributes.Remove("onclick")
            send.Enabled = False
            Button1.Enabled = False
            TIMS.Tooltip(send, "")
            TIMS.Tooltip(Button1, "")
            TIMS.Tooltip(DataGrid1, "")
            TIMS.Tooltip(send, Me.ViewState("msg2"))
            TIMS.Tooltip(Button1, Me.ViewState("msg2"))
            TIMS.Tooltip(DataGrid1, Me.ViewState("msg2"))
        End If
        Return rst

#End Region
    End Function

    Function CheckData1(ByRef Errmsg As String) As Boolean
#Region "CheckData1"

        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then Errmsg += "開訓日期 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then Errmsg += "開訓日期 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then Errmsg += "【開訓日期】的起日不得大於【開訓日期】的迄日!!" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst

#End Region
    End Function

    'LevelFlag 階段班級 0:是 1:否
    Sub SchLev0()
#Region "SchLev0"

        Dim rqWish As String = Request(Cst_rqwish) 'wish 1:單選／2:多選
        rqWish = TIMS.ClearSQM(rqWish)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID " & vbCrLf
        sql &= " ,cc.PlanID " & vbCrLf
        sql &= " ,cc.ComIDNO " & vbCrLf
        sql &= " ,cc.SeqNO " & vbCrLf
        sql &= " ,ip.TPlanID " & vbCrLf
        sql &= " ,dbo.DECODE6(cc.IsApplic,'Y','可挑選志願','y','可挑選志願','不可挑選志願') IsApplic " & vbCrLf
        sql &= " ,cc.LevelType " & vbCrLf
        'sql += " ,cc.IsApplic " & vbCrLf
        'sql += " ,cc.ClassCName " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME " & vbCrLf
        sql &= " ,cc.ClassCName ClassCName1 " & vbCrLf
        sql &= " ,cc.CyclType " & vbCrLf
        sql &= " ,cc.STDate " & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDate, 111) STDate1 " & vbCrLf
        sql &= " ,CONVERT(varchar, cc.SENTERDATE, 111) SENTERDATE " & vbCrLf '報名開始日
        sql &= " ,CONVERT(varchar, cc.FENTERDATE, 111) FENTERDATE " & vbCrLf '報名截止日
        sql &= " ,cc.Thours " & vbCrLf
        sql &= " ,cc.CLSID " & vbCrLf
        sql &= " ,cc.TMID " & vbCrLf
        sql &= " ,ktt.TrainID " & vbCrLf
        sql &= " ,'[' + ktt.TrainID + ']' + ktt.TrainName TrainName " & vbCrLf
        sql &= " FROM Class_ClassInfo cc " & vbCrLf
        sql &= " JOIN ID_Class ic ON cc.CLSID = ic.CLSID " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid = cc.planid " & vbCrLf
        sql &= " LEFT JOIN Key_TrainType ktt ON ktt.TMID = cc.TMID " & vbCrLf
        sql &= " LEFT JOIN MVIEW_RELSHIP23 R3 ON R3.RID3 = cc.RID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y' " & vbCrLf
        sql &= " AND cc.NotOpen != 'Y' " & vbCrLf
        'sql &= " and rownum <=100" & vbCrLf
        If ClassCName.Text <> "" Then
            sql &= " AND cc.ClassCName LIKE '%" & ClassCName.Text & "%' " & vbCrLf
        End If
        If CyclType.Text <> "" Then
            sql &= " AND cc.CyclType = '" & CyclType.Text & "' " & vbCrLf
        End If
        If start_date.Text <> "" Then
            sql &= " AND cc.STDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
        End If
        If end_date.Text <> "" Then
            sql &= " AND cc.STDate <= " & TIMS.To_date(end_date.Text) & vbCrLf
        End If
        If trainValue.Value <> "" Then
            sql &= " AND cc.TMID = '" & trainValue.Value & "' " & vbCrLf
        End If

        '<asp@ListItem Value="1" Selected="True">報名尚未結束班級</asp@ListItem>
        '<asp@ListItem Value="2">報名結束班級</asp@ListItem>
        '<asp@ListItem Value="4">尚未甄試班級</asp@ListItem>
        '<asp@ListItem Value="5">未結訓班級</asp@ListItem>
        '<asp@ListItem Value="3">所有的班級</asp@ListItem>

        Session(Cst_session1) = ""
        If ClassSort.SelectedValue <> "" Then
            '1.2.4.5.3
            Session(Cst_session1) = ClassSort.SelectedValue
        End If
        sql &= " AND cc.SEnterDate <= getdate() " & vbCrLf '報名已開始
        Select Case ClassSort.SelectedValue
            Case "1" '報名尚未結束班級
                sql &= " AND cc.FEnterDate > getdate() " & vbCrLf '報名尚未結束 (時間可能為00:00)
            Case "2" '報名結束班級
                sql &= " AND cc.FEnterDate <= getdate() " & vbCrLf '報名尚未結束 (時間可能為00:00)
            Case "4" '尚未甄試班級
                sql &= " AND cc.IsCalculate != 'Y' " & vbCrLf
            Case "5" '未結訓班級
                sql &= " AND cc.FTDate > dbo.TRUNC_DATETIME(getdate()) " & vbCrLf '未結訓班級 (時間為00:00)
            Case "3" '所有的班級
            Case Else
                Session(Cst_session1) = ""
        End Select

        If Not rqWish = "1" Then
            If Convert.ToString(Session("wish1_date")) <> "" Then
                sql &= " AND cc.STDate = " & TIMS.To_date(Session("wish1_date")) & vbCrLf
            End If
            sql &= " AND cc.IsApplic = 'Y' " & vbCrLf
        End If

#Region "(No Use)"

        'If TestStr = "AmuTest" Then '測試用
        '    sql &= " and cc.RID like '" & sm.UserInfo.RID & "%' " & vbCrLf '測試用
        'Else '測試用
        '    sql &= " and cc.RID='" & sm.UserInfo.RID & "' " & vbCrLf '測試用
        'End If '測試用
        'If TestStr <> "AmuTest" Then
        '    sql &= " and cc.RID='" & sm.UserInfo.RID & "' " & vbCrLf
        'End If

#End Region

        'Dim flag_useRID As Boolean = True
        Select Case sm.UserInfo.TPlanID
            Case "17"
                If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
                '補助地方政府
                'LID: 階層代碼【0:署(局) 1:分署(中心) 2:委訓(縣市政府)】 NUMBER
                Select Case sm.UserInfo.LID
                    Case "0", "1"
                        sql &= " AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
                    Case "2"
                        'flag_useRID = False
                        '【0:署(職訓局) 1:分署(中心) 2:委訓(補助單位) 3:(委訓)】
                        Select Case sm.UserInfo.OrgLevel
                            Case "3" '委訓單位
                                sql &= " AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
                            Case "2" '縣市政府(補助單位)。
                                sql &= " AND R3.RID2='" & RIDValue.Value & "' " & vbCrLf
                            Case Else
                                sql &= " AND cc.RID='" & RIDValue.Value & "' " & vbCrLf
                        End Select
                    Case Else
                        sql &= " AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
                End Select
            Case Else
#Region "(No Use)"
                'LID: 階層代碼【0:署(局) 1:分署(中心) 2:委訓】 NUMBER
                'Select Case sm.UserInfo.LID
                '    Case "0", "1"
                '        flag_useRID = False
                '    Case Else
                '        flag_useRID = True
                'End Select
                'sql &= " and cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf

#End Region
        End Select

        'sql += " and cc.RID='" & sm.UserInfo.RID & "' " & vbCrLf
        'sql &= " and cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf

        If trCenter.Visible Then
            If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
            sql &= " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
            If Len(sm.UserInfo.RID) = 1 Then
                Select Case sm.UserInfo.RID
                    Case "A" '署(局)權限
                        sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
                        sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                    Case Else
                        sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                        sql &= " AND ip.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf
                        sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
                End Select
            Else
                '依登入年度計畫@PlanID
                sql &= " AND cc.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
            End If
        Else
            sql &= " AND cc.RID = '" & sm.UserInfo.RID & "' " & vbCrLf
            sql &= " AND cc.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If

        sql &= " ORDER BY ic.ClassID, cc.CyclType " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        Me.ViewState("dv") = dt
        msg.Text = cst_AlertMsg1
        DataGridTable1.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable1.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "OCID"
            DataGrid1.DataBind()
        End If

        '改變顯示模式(1:單選／2:多選)
        If rqWish = "1" Then
            DataGrid1.Columns(0).Visible = True
            DataGrid1.Columns(1).Visible = False
        Else
            DataGrid1.Columns(0).Visible = False
            DataGrid1.Columns(1).Visible = True
        End If

        '確認登入帳號之機構是否在黑名單中
        If Me.Chk_OrgBlackList() Then
            DataGrid1.Columns(0).Visible = False '無法選擇該機構。
            DataGrid1.Columns(1).Visible = False '無法選擇該機構。
        End If

#End Region
    End Sub

    'LevelFlag 階段班級 0:是 1:否
    Sub SchLev1()
#Region "SchLev1"

        'Dim datestr As String = ""
        'Dim TMIDStr As String = ""
        'Dim sort As String = ""
        'Session(Cst_session1) = ""
        'If ClassSort.SelectedValue <> "" Then Session(Cst_session1) = ClassSort.SelectedValue

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME " & vbCrLf
        sql &= " ,a.CyclType " & vbCrLf
        sql &= " ,b.LSDate " & vbCrLf
        sql &= " ,b.LevelName " & vbCrLf
        sql &= " ,b.LevelSDate " & vbCrLf
        sql &= " ,b.CCLID " & vbCrLf
        sql &= " ,b.Num " & vbCrLf
        sql &= " FROM Class_ClassInfo a " & vbCrLf
        sql &= " JOIN Class_ClassLevel b ON a.OCID = b.OCID " & vbCrLf
        sql &= " JOIN ID_Class c ON a.CLSID = c.CLSID " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PlanID = a.PlanID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " and a.SEnterDate <= getdate() " & vbCrLf '報名已開始
        sql &= " AND a.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        sql &= " AND a.RID = '" & sm.UserInfo.RID & "' " & vbCrLf
        sql &= " AND b.LevelName != '01' " & vbCrLf
        'sql += " AND b.Num IS Not NULL " & vbCrLf

        'start_date.Text = TIMS.ClearSQM(start_date.Text)
        'end_date.Text = TIMS.ClearSQM(end_date.Text)
        'trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If start_date.Text <> "" Then
            sql &= " AND a.STDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
        End If
        If end_date.Text <> "" Then
            sql &= " AND a.STDate <= " & TIMS.To_date(end_date.Text) & vbCrLf
        End If
        If trainValue.Value <> "" Then
            sql &= " AND a.TMID = '" & trainValue.Value & "' " & vbCrLf
        End If
        If trCenter.Visible Then
            If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
            sql &= " AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
            If Len(sm.UserInfo.RID) = 1 Then
                Select Case sm.UserInfo.RID
                    Case "A" '署(局)權限
                        sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
                        sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                    Case Else
                        sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                        sql &= " AND ip.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf
                        sql &= " AND ip.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
                End Select
            Else
                '依登入年度計畫@PlanID
                sql &= " AND a.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
            End If
        Else
            sql &= " AND a.RID = '" & sm.UserInfo.RID & "' " & vbCrLf
            sql &= " AND a.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        sql &= " ORDER BY c.ClassID, a.CyclType " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        'msg.Text = "查無資料!!"
        msg.Text = cst_AlertMsg1
        DataGridTable2.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable2.Visible = True
            DataGrid2.DataSource = dt
            DataGrid2.DataBind()
        End If

#End Region
    End Sub

    '查詢 [SQL]
    Sub sSearch1()
#Region "sSearch1"

        '若有session消失離開此搜尋 (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        '整理
        ClassCName.Text = TIMS.ClearSQM(ClassCName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        If CyclType.Text <> "" Then
            CyclType.Text = CInt(Val(CyclType.Text))
            If CyclType.Text.Length > 2 Then CyclType.Text = Left(CyclType.Text, 2)
            If CyclType.Text.Length < 2 Then CyclType.Text = "0" & CyclType.Text
            'sql += " and cc.CyclType='" & CyclType.Text & "'" & vbCrLf
        End If
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)

        Session(Cst_session1) = ""
        If ClassSort.SelectedValue <> "" Then
            Session(Cst_session1) = ClassSort.SelectedValue  '1.2.4.5.3
        End If
        'sql += " and cc.SEnterDate <= getdate()" & vbCrLf '報名已開始
        Select Case ClassSort.SelectedValue
            Case "1" '報名尚未結束班級
                'sql += " and cc.FEnterDate > getdate()" & vbCrLf '報名尚未結束 (時間可能為00:00)
            Case "2" '報名結束班級
                'sql += " and cc.FEnterDate <= getdate()" & vbCrLf '報名尚未結束 (時間可能為00:00)
            Case "4" '尚未甄試班級
                'sql += " and cc.IsCalculate != 'Y' " & vbCrLf
            Case "5" '未結訓班級
                'sql += " and cc.FTDate > dbo.TRUNC_DATETIME(getdate()) " & vbCrLf '未結訓班級 (時間為00:00)
            Case "3" '所有的班級
            Case Else
                Session(Cst_session1) = ""
        End Select

        'Dim dt As DataTable
        'Dim sql As String = ""
        'LevelFlag 階段班級 0:是 1:否
        Select Case LevelFlag.SelectedIndex
            Case 0 'LevelFlag 階段班級 0:是 1:否
                Call SchLev0()
            Case 1 'LevelFlag 階段班級 0:是 1:否
                Call SchLev1()
        End Select

#End Region
    End Sub

    '查詢鈕
    Private Sub search_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search_but.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call sSearch1() '查詢 [SQL]
    End Sub

    '送出
    Private Sub send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click
#Region "送出"

        Me.ViewState("hidstar3") = ""
        Dim rqClass1 As String = Request("class1")
        rqClass1 = TIMS.ClearSQM(rqClass1)
        'Request("class2")
        Dim rqClass2 As String = Request("class2")
        rqClass2 = TIMS.ClearSQM(rqClass2)
        'Request(Cst_rqwish) = 1 'wish 1:單選／2:多選
        Dim rqWish As String = Request(Cst_rqwish) 'wish 1:單選／2:多選
        rqWish = TIMS.ClearSQM(rqWish)

        Dim row() As DataRow
        Dim data_search As DataTable = Me.ViewState("dv")

#Region "(No Use)"

        'If rqWish = 1 Then
        '    If rqClass1 <> "" Then
        '        row = data_search.Select("OCID='" & rqClass1 & "'")
        '        If row.Length <> 0 AndAlso StudIDNO.Value <> "" Then
        '            'TIMS 非產投計畫(在職)。
        '            Dim flagCanUseChk1 As Boolean = False '是否可執行檢核(true:可 false:不可)
        '            'If TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID) = -1 Then flagCanUseChk1 = True
        '            If TIMS.Cst_TPlanID_PreUseLimited17e.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagCanUseChk1 = True
        '            If flagCanUseChk1 Then
        '                'Chk_IsJobsCounseling
        '                Dim ss As String = "" 'ss = ""
        '                Call TIMS.SetMyValue(ss, "IDNO", StudIDNO.Value)
        '                Call TIMS.SetMyValue(ss, "STDate1", row(0).Item("STDate1"))
        '                Call TIMS.SetMyValue(ss, "FENTERDATE", row(0).Item("FENTERDATE"))
        '                Call TIMS.SetMyValue(ss, "ClassCName", row(0).Item("ClassCName1"))
        '                Call TIMS.SetMyValue(ss, "Thours", row(0).Item("Thours"))
        '                Call TIMS.SetMyValue(ss, "OCID", row(0).Item("OCID"))
        '                Dim iNum As Integer = 0
        '                If TIMS.Chk_IsJobsCounseling(ss, iNum, objconn) Then
        '                    Hid_IJC.Value = TIMS.cst_IJC2
        '                End If
        '                'If TestStr = "AmuTest" Then Hid_IJC.Value = cst_IJC '測試用
        '            End If
        '        End If
        '    End If
        'End If

#End Region

        Select Case rqWish 'wish 1:單選／2:多選
            Case "1" 'wish 1:單選／2:多選
                If rqClass1 = "" Then
                    Common.MessageBox(Me, cst_AlertMsg2)
                    Exit Sub
                End If
                If StudIDNO.Value <> "" Then
                    If TIMS.Chk_StudStatus(StudIDNO.Value, rqClass1, objconn) Then
                        Me.ViewState("hidstar3") = "1"
                    End If
                End If
                row = data_search.Select("OCID='" & rqClass1 & "'")
                If row.Length = 0 Then
                    '資料連結錯誤
                    Common.MessageBox(Me, cst_AlertMsg2)
                    Exit Sub
                End If
                Session("wish1_date") = row(0).Item("STDate")
                Common.RespWrite(Me, "<script language=javascript>")
                Common.RespWrite(Me, "function returnNum(){")
                'Common.RespWrite(Me, "window.opener.document.form1.TMID1.value='[" & row(0).Item("TrainID") & "]" & row(0).Item("TrainName") & "';")
                Common.RespWrite(Me, "window.opener.document.form1.ComIDNO1.value='" & row(0).Item("ComIDNO") & "';")      ''第一志願的廠商統一編號
                Common.RespWrite(Me, "window.opener.document.form1.SeqNO1.value='" & row(0).Item("SeqNO") & "';")      ''第一志願的計畫主檔序號
                Common.RespWrite(Me, "window.opener.document.form1.CCLID.value='';")

                Common.RespWrite(Me, "window.opener.document.form1.TMID1.value='" & row(0).Item("TrainName") & "';")
                Common.RespWrite(Me, "window.opener.document.form1.TMIDValue1.value='" & row(0).Item("TMID") & "';")
                Common.RespWrite(Me, "window.opener.document.form1.OCID1.value='" & row(0).Item("ClassCName") & "';")
                Common.RespWrite(Me, "window.opener.document.form1.OCIDValue1.value='" & row(0).Item("OCID") & "';")
                'If Convert.ToString(row(0).Item("IsApplic")).ToUpper = "N" Then          '如果第一志願的志願選項是N的話，停止選擇第二、三志願
                '    Common.RespWrite(Me, "window.opener.document.form1.Button2.disabled=true;")
                '    Common.RespWrite(Me, "window.opener.document.form1.Button3.disabled=true;")
                'Else
                '    Common.RespWrite(Me, "window.opener.document.form1.Button2.disabled=false;")
                '    Common.RespWrite(Me, "window.opener.document.form1.Button3.disabled=false;")
                'End If
                If Me.ViewState("hidstar3") = "1" Then
                    '仍在訓中
                    Common.RespWrite(Me, "if(window.opener.document.form1.hidstar3!=null) window.opener.document.form1.hidstar3.value='" & Me.ViewState("hidstar3") & "';")
                End If
                If Hid_IJC.Value <> "" Then
                    '與委外實施基準條款有抵觸
                    Common.RespWrite(Me, "if(window.opener.document.form1.HidIJCMsg!=null) window.opener.document.form1.HidIJCMsg.value='" & Hid_IJC.Value & "';")
                End If
                Common.RespWrite(Me, "window.close();")
                Common.RespWrite(Me, "}")
                Common.RespWrite(Me, "returnNum();")
                Common.RespWrite(Me, "</script>")

            Case "2" 'wish 1:單選／2:多選
                If rqClass2 = "" Then
                    Common.MessageBox(Me, cst_AlertMsg2)
                    Exit Sub
                End If
                Dim all() As String = Split(rqClass2, ",", , CompareMethod.Text)
                If all.Length <> 2 Then
                    Common.MessageBox(Me, cst_AlertMsg3)
                    Exit Sub
                End If
                row = data_search.Select("OCID='" & all(0) & "'")
                Common.RespWrite(Me, "<script language=javascript>")
                Common.RespWrite(Me, "function returnNum(){")
                'Common.RespWrite(Me, "window.opener.document.form1.TMID2.value='" & row(0).Item("TrainName") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.TMIDValue2.value='" & row(0).Item("TMID") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.OCID2.value='" & row(0).Item("ClassCName") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.OCIDValue2.value='" & row(0).Item("OCID") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.ComIDNO2.value='" & row(0).Item("ComIDNO") & "';")      ''第二志願的廠商統一編號
                'Common.RespWrite(Me, "window.opener.document.form1.SeqNO2.value='" & row(0).Item("SeqNO") & "';")      ''第二志願的計畫主檔序號
                'If all.Length = 2 Then
                '    row = data_search.Select("OCID='" & all(1) & "'")
                '    Common.RespWrite(Me, "window.opener.document.form1.TMID3.value='" & row(0).Item("TrainName") & "';")
                '    Common.RespWrite(Me, "window.opener.document.form1.TMIDValue3.value='" & row(0).Item("TMID") & "';")
                '    Common.RespWrite(Me, "window.opener.document.form1.OCID3.value='" & row(0).Item("ClassCName") & "';")
                '    Common.RespWrite(Me, "window.opener.document.form1.OCIDValue3.value='" & row(0).Item("OCID") & "';")
                '    Common.RespWrite(Me, "window.opener.document.form1.ComIDNO3.value='" & row(0).Item("ComIDNO") & "';")      ''第三志願的廠商統一編號
                '    Common.RespWrite(Me, "window.opener.document.form1.SeqNO3.value='" & row(0).Item("SeqNO") & "';")      ''第三志願的計畫主檔序號
                'End If
                Common.RespWrite(Me, "window.close();")
                Common.RespWrite(Me, "}")
                Common.RespWrite(Me, "returnNum();")
                Common.RespWrite(Me, "</script>")
            Case Else '沒有 'wish 1:單選／2:多選
                If rqClass2 = "" Then
                    Common.MessageBox(Me, cst_AlertMsg2)
                    Exit Sub
                End If
                Dim all() As String = Split(rqClass2, ",", , CompareMethod.Text)
                If all.Length <> 1 Then
                    Common.MessageBox(Me, cst_AlertMsg4)
                    Exit Sub
                End If
                row = data_search.Select("OCID='" & all(0) & "'")
                Common.RespWrite(Me, "<script language=javascript>")
                Common.RespWrite(Me, "function returnNum(){")
                'Common.RespWrite(Me, "window.opener.document.form1.TMID3.value='" & row(0).Item("TrainName") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.TMIDValue3.value='" & row(0).Item("TMID") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.OCID3.value='" & row(0).Item("ClassCName") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.OCIDValue3.value='" & row(0).Item("OCID") & "';")
                'Common.RespWrite(Me, "window.opener.document.form1.ComIDNO3.value='" & row(0).Item("ComIDNO") & "';")      ''第三志願的廠商統一編號
                'Common.RespWrite(Me, "window.opener.document.form1.SeqNO3.value='" & row(0).Item("SeqNO") & "';")      ''第三志願的計畫主檔序號
                Common.RespWrite(Me, "window.close();")
                Common.RespWrite(Me, "}")
                Common.RespWrite(Me, "returnNum();")
                Common.RespWrite(Me, "</script>")
        End Select
    End Sub

    '插班的DataGrid列表
    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If

                Dim drv As DataRowView = e.Item.DataItem
                Dim CCLID As HtmlInputRadioButton = e.Item.FindControl("CCLID")
                CCLID.Value = drv("CCLID").ToString
                CCLID.Attributes("onclick") = "SelectItem(" & e.Item.ItemIndex + 1 & ")"

                e.Item.Cells(2).Text = TIMS.ChangeNum(Val(drv("LevelName"))) '換中文。
        End Select

#End Region
    End Sub

    '階段送出按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
#Region "階段送出按鈕"

        Dim sql As String
        'Dim dt As DataTable
        Dim dr As DataRow
        Dim CCLIDValue As Integer
        Dim ClassCName As String

        For Each item As DataGridItem In DataGrid2.Items
            Dim CCLID As HtmlInputRadioButton = item.FindControl("CCLID")
            If CCLID.Checked Then
                CCLIDValue = CCLID.Value
            End If
        Next

        CCLIDValue = TIMS.ClearSQM(CCLIDValue)
        sql = ""
        sql &= " SELECT a.CCLID " & vbCrLf
        sql &= " ,a.LevelName " & vbCrLf
        sql &= " ,b.TMID " & vbCrLf
        sql &= " ,'[' + c.TrainID + ']' + c.TrainName TrainName " & vbCrLf
        sql &= " ,b.OCID " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME " & vbCrLf
        sql &= " ,b.CyclType " & vbCrLf
        sql &= " ,b.PlanID " & vbCrLf
        sql &= " ,b.ComIDNO " & vbCrLf
        sql &= " ,b.SeqNo " & vbCrLf
        sql &= " FROM Class_ClassLevel a " & vbCrLf
        sql &= " JOIN Class_ClassInfo b ON a.OCID = b.OCID " & vbCrLf
        sql &= " JOIN Key_TrainType c ON b.TMID = c.TMID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.CCLID = '" & CCLIDValue & "' " & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        ClassCName = dr("ClassCName").ToString
        ClassCName &= "(第" & TIMS.ChangeNum(Val(dr("LevelName"))) & "階段)"

        Common.RespWrite(Me, "<script>")
        Common.RespWrite(Me, "window.opener.document.form1.TMID1.value='" & dr("TrainName") & "';")
        Common.RespWrite(Me, "window.opener.document.form1.TMIDValue1.value='" & dr("TMID") & "';")
        Common.RespWrite(Me, "window.opener.document.form1.OCID1.value='" & ClassCName & "';")
        Common.RespWrite(Me, "window.opener.document.form1.OCIDValue1.value='" & dr("OCID") & "';")
        Common.RespWrite(Me, "window.opener.document.form1.ComIDNO1.value='" & dr("ComIDNO") & "';")
        Common.RespWrite(Me, "window.opener.document.form1.SeqNO1.value='" & dr("SeqNo") & "';")
        Common.RespWrite(Me, "window.opener.document.form1.CCLID.value='" & dr("CCLID") & "';")
        'Common.RespWrite(Me, "window.opener.document.form1.Button2.disabled=true;")
        'Common.RespWrite(Me, "window.opener.document.form1.Button3.disabled=true;")
        Common.RespWrite(Me, "window.close();")
        Common.RespWrite(Me, "</script>")

#End Region
    End Sub
End Class