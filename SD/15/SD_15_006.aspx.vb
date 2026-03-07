Partial Class SD_15_006
    Inherits AuthBasePage

    'SD_15_006_R & RadioItem.SelectedValue & FuncID.SelectedValue
    'SD_15_006_R0*.jrxml '(0.1.2.) '(0~12)
    'SD_15_006_R1*.jrxml '(0.1.2.) '(0~12)
    'SD_15_006_R2*.jrxml '(0.1.2.) '(0~12)

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        '訓練機構
        If Not IsPostBack Then
            SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
            Common.SetListItem(SearchPlan, "G")

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, Year(Now))
            Get_FuncList(FuncID)
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            RadioItem.SelectedIndex = 0

            trPlanKind.Style("display") = "none"
            trPackageType.Style("display") = "none"
            '54:充電起飛計畫（在職）判斷方式
            If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                trPackageType.Style("display") = TIMS.cst_inline1 '"inline"
            Else
                '28:產業人才投資方案
                '計畫範圍 產投
                If sm.UserInfo.Years >= 2008 Then
                    trPlanKind.Style("display") = TIMS.cst_inline1 '"inline"
                End If
            End If

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button5_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "javascript:return print();"
    End Sub

    '0~12
    Public Shared Sub Get_FuncList(ByRef obj As ListControl)
        With obj
            If TypeOf obj Is DropDownList Then
                .Items.Insert(0, New ListItem("性別", "0"))
                .Items.Insert(1, New ListItem("年齡", "1"))
                .Items.Insert(2, New ListItem("教育程度", "2"))
                .Items.Insert(3, New ListItem("身分別", "3"))
                .Items.Insert(4, New ListItem("工作年資", "4"))
                .Items.Insert(5, New ListItem("地理分佈", "5"))
                .Items.Insert(6, New ListItem("公司行業別", "6"))
                .Items.Insert(7, New ListItem("公司規模", "7"))
                .Items.Insert(8, New ListItem("參訓動機", "8"))
                .Items.Insert(9, New ListItem("訓後動向", "9"))
                .Items.Insert(10, New ListItem("參訓單位類別", "10"))
                '.Items.Insert(11, New ListItem("參加課程業別", "11"))
                '.Items.Insert(12, New ListItem("參加課程職類別", "12"))
                .Items.Insert(11, New ListItem("參加課程職能別", "11"))
                .Items.Insert(12, New ListItem("參加課程型態別", "12"))
            End If
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Const Cst_RptName As String = "SD_15_006_R"

        Dim ReportName As String
        Dim ReportRadio1 As String
        Dim ReportRadio2 As String

        Years.Value = Me.yearlist.SelectedValue

        '統計項目 0:已用補助費統計
        '統計項目 1:已得學分統計
        '統計項目 2:已上課程數量統計
        ReportRadio1 = RadioItem.SelectedValue '(0.1.2.)
        '交叉查詢選項 參考Get_FuncList
        ReportRadio2 = FuncID.SelectedValue

        Dim SearchPlan1 As String = ""
        If SearchPlan.SelectedIndex <> 0 Then
            If SearchPlan.SelectedValue <> "" Then
                SearchPlan1 = SearchPlan.SelectedValue
            End If
        End If

        ReportName = Cst_RptName & ReportRadio1 & ReportRadio2
        If RIDValue.Value = "A" Then RIDValue.Value = ""

        Dim sPackType As String = ""
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            SearchPlan1 = "" '清空
            If PackageType.SelectedValue <> "A" Then
                sPackType = PackageType.SelectedValue
            End If
        End If

        Dim MyValue As String = ""
        MyValue &= "&Years=" & Years.Value
        MyValue &= "&RID=" & RIDValue.Value
        MyValue &= "&OCID=" & OCIDValue1.Value
        MyValue &= "&SearchPlan=" & SearchPlan1
        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue &= "&PackageType=" & sPackType
        ReportQuery.PrintReport(Me, "Report", ReportName, MyValue) '"Years=" & Years.Value & "&RID=" & RIDValue.Value & "&OCID=" & OCIDValue1.Value & "&SearchPlan=" & SearchPlan1)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        If dr Is Nothing OrElse Convert.ToString(dr("total")) <> "1" Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

End Class