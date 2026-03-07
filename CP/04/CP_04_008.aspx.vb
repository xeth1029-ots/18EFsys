Partial Class CP_04_008
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            Call Create1()
        End If
    End Sub

    Sub Create1()
        '檢查日期格式
        Me.SSTDate.Attributes("onchange") = "check_date();"
        Me.ESTDate.Attributes("onchange") = "check_date();"
        Me.SFTDate.Attributes("onchange") = "check_date();"
        Me.EFTDate.Attributes("onchange") = "check_date();"

        'Dim dt As DataTable
        ''Dim dr As DataRow
        'Dim sqlstr As String

        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)

        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        Sql = "SELECT NAME,DISTID FROM ID_DISTRICT ORDER BY DISTID"
        dt = DbAccess.GetDataTable(Sql, objconn)
        Me.DistrictList.DataSource = dt
        Me.DistrictList.DataTextField = "Name"
        Me.DistrictList.DataValueField = "DistID"
        Me.DistrictList.DataBind()
        Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

        PlanList = TIMS.Get_TPlan(PlanList, , 1, "Y")

        '縣市
        CityList = TIMS.Get_CityName(CityList, TIMS.dtNothing)

        sql = "SELECT TMID,BUSID+'.'+BUSNAME BUSNAME FROM KEY_TRAINTYPE WHERE Levels=0 AND BUSID!='G' ORDER BY TMID"
        dt = DbAccess.GetDataTable(sql, objconn)
        With TMID
            .DataSource = dt
            .DataTextField = "BusName"
            .DataValueField = "TMID"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", ""))
        End With

        ''預算來源

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        '選擇全部縣市
        Me.CityList.Attributes("onclick") = "SelectAll('CityList','CityHidden');"

        '選擇全部訓練計畫
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"

        '當分署(中心)使用者使用時,轄區應該都要鎖死該轄區,不可選擇其它轄區
        Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "0"
            Case "1"
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
            Case Else
                Common.SetListItem(DistrictList, sm.UserInfo.DistID)
                DistrictList.Enabled = False
                'DistrictList.Style.Item("display") = "none"
        End Select
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        '選擇轄區
        'Dim objitem As ListItem
        'Dim DistID, DistName, ICity As String
        'Dim newDistID, newICity As String
        'Dim TPlanID As String
        'Dim newTPlanID As String
        'Dim newDistName, ICityName, newICityName, TPlanName, newTPlanIDName As String
        'Dim NotOpenStausStr, TMIDName As String
        'Dim i As Integer

        Dim itemstr As String = ""
        For Each objitem As ListItem In Me.DistrictList.Items
            If objitem.Selected = True AndAlso objitem.Value <> "" Then
                If itemstr <> "" Then itemstr &= ","
                itemstr &= "'" & objitem.Value & "'"
            End If
        Next

        '報表要用的轄區參數
        Dim DistID As String = ""
        Dim DistName As String = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected AndAlso DistrictList.Items(i).Value <> "" Then
                If DistID <> "" Then DistID &= ","
                DistID &= "\'" & Me.DistrictList.Items(i).Value & "\'"
                If DistName <> "" Then DistName &= ","
                DistName &= Me.DistrictList.Items(i).Text
            End If
        Next
        If DistID <> "" AndAlso Me.DistrictList.Items(0).Selected Then
            DistName = "全部"
        End If
        'If DistID <> "" Then
        '    If Me.DistrictList.Items(0).Selected Then
        '        newDistID = Mid(DistID, 1, DistID.Length - 1)
        '        newDistName = "全部"
        '    Else
        '        newDistID = Mid(DistID, 1, DistID.Length - 1)
        '        newDistName = Mid(DistName, 1, DistName.Length - 1)
        '    End If
        'End If

        '選擇縣市
        Dim itemcity As String = ""
        For Each objitem As ListItem In Me.CityList.Items
            If objitem.Selected = True AndAlso objitem.Value <> "" Then
                If itemcity <> "" Then itemcity &= ","
                itemcity &= "'" & objitem.Value & "'"
            End If
        Next

        '報表要用的縣市參數
        Dim ICity As String = ""
        Dim ICityName As String = ""
        For i As Integer = 1 To Me.CityList.Items.Count - 1
            If Me.CityList.Items(i).Selected AndAlso Me.CityList.Items(i).Value <> "" Then
                If ICity <> "" Then ICity &= ","
                ICity &= "\'" & Me.CityList.Items(i).Value & "\'"
                If ICityName <> "" Then ICityName &= ","
                ICityName &= Me.CityList.Items(i).Text
            End If
        Next
        'If ICity <> "" Then
        '    newICity = Mid(ICity, 1, ICity.Length - 1)
        'End If
        If ICity <> "" AndAlso Me.CityList.Items(0).Selected Then
            ICityName = "全部"
        End If

        '選擇訓練計畫
        Dim itemplan As String = ""
        For Each objitem As ListItem In Me.PlanList.Items
            If objitem.Selected = True AndAlso objitem.Value <> "" Then
                If itemplan <> "" Then itemplan &= ","
                itemplan &= "\'" & objitem.Value & "\'"
            End If
        Next

        '報表要用的訓練計畫參數
        Dim TPlanID As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To Me.PlanList.Items.Count - 1
            If Me.PlanList.Items(i).Selected AndAlso Me.PlanList.Items(i).Value <> "" Then
                If TPlanID <> "" Then TPlanID &= ","
                TPlanID &= "\'" & Me.PlanList.Items(i).Value & "\'"
                If TPlanName <> "" Then TPlanName &= ","
                TPlanName &= Me.PlanList.Items(i).Text
            End If
        Next

        'If TPlanID <> "" Then
        '    newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        'End If
        If TPlanID <> "" AndAlso Me.PlanList.Items(0).Selected Then
            TPlanName = "全部"
        End If

        '選擇開班狀態
        Dim NotOpenStaus As String = ""
        Dim NotOpenStausStr As String = ""
        If Me.NotOpenStaus.Items(0).Selected AndAlso Me.NotOpenStaus.Items(1).Selected Then
            NotOpenStausStr = "開班,不開班"
        Else
            Select Case Me.NotOpenStaus.SelectedIndex
                Case 0 '0 開班
                    NotOpenStaus = "N"
                    NotOpenStausStr = "開班"
                Case Else '1 不開班
                    NotOpenStaus = "Y"
                    NotOpenStausStr = "不開班"
            End Select
        End If

        Session("_search") = "prog=CP_04_008"
        Session("_search") += "&itemstr=" & itemstr
        Session("_search") += "&itemplan=" & itemplan
        Session("_search") += "&SSTDate=" & Me.SSTDate.Text
        Session("_search") += "&ESTDate=" & Me.ESTDate.Text
        Session("_search") += "&SFTDate=" & Me.SFTDate.Text
        Session("_search") += "&EFTDate=" & Me.EFTDate.Text
        Session("_search") += "&NotOpenStaus=" & NotOpenStaus
        Session("_search") += "&newDistID=" & DistID 'newDistID
        Session("_search") += "&newTPlanID=" & TPlanID 'newTPlanID
        Session("_search") += "&TMID=" & TMID.SelectedValue
        'Session("_search") += "&itembudget=" & itembudget
        'Session("_search") += "&newBudgetID=" & newBudgetID
        Session("_search") += "&OrgName=" & OrgName.Text
        Session("_search") += "&ClassCName=" & ClassCName.Text
        Session("itemcity") = itemcity
        Session("newICity") = ICityName 'newICity
        Session("_search") += "&newICityName=" & ICityName 'newICityName
        Session("_search") += "&newTPlanIDName=" & TPlanName 'newTPlanIDName
        Session("_search") += "&NotOpenStausStr=" & NotOpenStausStr
        Session("_search") += "&TMIDName=" & TMID.SelectedItem.Text
        'Session("_search") += "&newBudgetName=" & newBudgetName
        'Response.Redirect("CP_04_003_add.aspx?yearlist=" & Me.yearlist.SelectedValue)
        Dim url1 As String = "CP_04_003_add.aspx?ID=" & Request("ID") & "&yearlist=" & Me.yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click
        'Reset
        Me.yearlist.SelectedIndex = 1

        Dim i As Short
        For i = 0 To Me.DistrictList.Items.Count - 1
            Me.DistrictList.Items(i).Selected = False
        Next

        For i = 0 To Me.CityList.Items.Count - 1
            Me.CityList.Items(i).Selected = False
        Next

        For i = 0 To Me.PlanList.Items.Count - 1
            Me.PlanList.Items(i).Selected = False
        Next

        Me.SSTDate.Text = ""
        Me.ESTDate.Text = ""
        Me.SFTDate.Text = ""
        Me.EFTDate.Text = ""
        Me.OrgName.Text = ""
        Me.ClassCName.Text = ""

        Me.NotOpenStaus.SelectedIndex = 0
        Me.TMID.SelectedIndex = 0

        'For i = 0 To Me.BudgetList.Items.Count - 1
        '    Me.BudgetList.Items(i).Selected = False
        'Next
    End Sub

End Class
