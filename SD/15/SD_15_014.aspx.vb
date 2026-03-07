Partial Class SD_15_014
    Inherits AuthBasePage

#Region "NO USE"
    'SELECT  ip.distname
    ',ip.years
    ',oo.orgname 
    ',cc.ocid
    ',convert(varchar,cc.STDate,111) STDate
    ',convert(varchar,cc.FTDate,111) FTDate
    ',isnull(cs.closePeo,0) closePeo
    ',isnull(cs.quesPeo,0) quesPeo
    'from class_classinfo cc
    ' join plan_planinfo pp on pp.planid =cc.planid and pp.comidno =cc.comidno and pp.seqno =cc.seqno 
    ' join view_plan ip on ip.planid=cc.planid
    ' join org_orginfo oo on oo.comidno=cc.comidno
    ' join id_class icc on icc.CLSID=cc.CLSID
    ' left join (
    '	select c.ocid 
    '	,sum(case when c.studstatus not in (2,3) and c.studstatus=5 then 1 end) closePeo 
    '	,sum(case when c.studstatus not in (2,3) and c.studstatus=5 and q1.socid is not null then 1 end) quesPeo 
    '	from class_studentsofclass c
    '	left join Stud_QuestionFin q1 on q1.socid =c.socid 
    '	group by c.ocid 
    ' ) cs on cs.ocid =cc.ocid 
    'WHERE 1=1
    ' and cc.IsSuccess	='Y'
    ' and cc.NotOpen='N'
    'and ip.years >=2011
    'and ip.years <=2011
    'and isnull(cs.quesPeo,0)>0
    'and isnull(cs.quesPeo,0) !=isnull(cs.closePeo,0)
    'order by 
    'ip.distname,oo.orgname ,cc.classcname,cc.cycltype
#End Region

    Const cst_printFN1 As String = "SD_15_014_R"
    'SD_15_014_R (Stud_QuestionFin)

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        '訓練機構
        If Not IsPostBack Then
            '含有不區分
            SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
            Common.SetListItem(SearchPlan, "G")

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, Year(Now))

            '計畫範圍 產投
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

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

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

        Years.Value = sm.UserInfo.Years

        Button1.Attributes("onclick") = "javascript:return print();"
    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim SearchPlan1 As String = ""
        If RIDValue.Value.ToString = "A" Then RIDValue.Value = ""
        '28:產業人才投資方案
        If SearchPlan.SelectedIndex <> 0 Then
            SearchPlan1 = SearchPlan.SelectedValue.ToString
        End If

        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            SearchPlan1 = ""
        End If
        Dim sPackType As String = ""
        If PackageType.SelectedValue <> "A" Then
            sPackType = PackageType.SelectedValue
        End If

        Dim MyValue As String = ""
        MyValue += "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue += "&Years=" & yearlist.SelectedValue
        MyValue += "&OCID=" & OCIDValue1.Value
        MyValue += "&RID=" & RIDValue.Value
        MyValue += "&SearchPlan=" & SearchPlan1 '"",G:產業人才投資計畫,W:提升勞工自主學習計畫
        MyValue += "&PackageType=" & sPackType '"",2:企業包班,3:企業聯合包班
        MyValue += "&FTDate1=" & FTDate1.Text
        MyValue += "&FTDate2=" & FTDate2.Text
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_15_014_R", MyValue)
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
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
