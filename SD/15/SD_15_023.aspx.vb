Partial Class SD_15_023
    Inherits AuthBasePage

    '【綜合查詢統計表】
    '實際開訓人次： 班級開訓，學員開訓後14天實際錄訓人數,且有選擇預算別(公務/就安/就保或公務(ECFA))
    '結訓人次：班級開訓，學員補助符合補助者沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))
    '撥款人次：班級開訓，學員補助符合補助者 沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))-學員經費撥款狀態：已撥款之人數
    '【交叉分析統計表】： -判斷邏輯調整成跟【綜合查詢統計表】
    '參訓人數： 班級開訓，學員沒有離退訓，只剩開結訓人數-判斷邏輯調整成跟【綜合查詢統計表】
    '結訓人數：班級開訓，學員已結訓人數-判斷邏輯調整成跟【綜合查詢統計表】
    '撥款人數：班級開訓，學員經費撥款狀態：已撥款之人數-判斷邏輯調整成跟【綜合查詢統計表】

    '產投使用
    '#Region "參數/變數 設定"

    Const Cst_性別 As String = "1"
    Const Cst_年齡 As String = "2"
    Const Cst_教育程度 As String = "3"
    Const Cst_特定對象 As String = "4"
    Const Cst_結訓後動向 As String = "5"
    Const Cst_工作年資 As String = "6"
    Const Cst_受訓學員地理分布 As String = "7" 'CTID (受訓學員) 通訊地址 STUD_SUBDATA.ZipCode1
    Const Cst_所屬公司行業別 As String = "8"
    Const Cst_所屬公司規模 As String = "9"
    Const Cst_參訓動機 As String = "10"

    Const Cst_報名人數 As String = "11"
    Const Cst_參訓人數 As String = "12"
    Const Cst_結訓人數 As String = "13"
    Const Cst_撥款人數 As String = "14"

    Const Cst_參訓單位類別 As String = "15"
    Const Cst_職能課程分類 As String = "16" '參加課職能別 
    Const Cst_參加課程型態 As String = "17"
    Const Cst_外籍配偶類別 As String = "18"

    Const Cst_外配類別1 As String = "本國"
    Const Cst_外配類別2 As String = "外籍(大陸人士)"
    Const Cst_外配類別3 As String = "外籍(非大陸人士)"
    Const Cst_sch_本國 As String = " and PassPortNO='1'"
    Const Cst_sch_外籍_大陸人士 As String = " and PassPortNO='2' and ChinaOrNot='1'" 'ChinaOrNot 1:大陸人士 /2:非大陸人士
    Const Cst_sch_外籍_非大陸人士 As String = " and PassPortNO='2' and ChinaOrNot='2'"
    Const Cst_sch_本國外籍範圍 As String = " and PassPortNO in ('1','2')"

    Const Cst_sch_本國x As String = "PassPortNO='1' "
    Const Cst_sch_外籍_大陸人士x As String = "PassPortNO='2' and ChinaOrNot='1' "
    Const Cst_sch_外籍_非大陸人士x As String = "PassPortNO='2' and ChinaOrNot='2' "
    Const Cst_sch_本國外籍範圍x As String = "PassPortNO in ('1','2') "

    Const Cst_ageIN As String = "(AGE IS NOT NULL)"
    '調整為15-19、20-24、25-29、30-34、35-39、40-44、45-49、50-54、55-59、60-64、65歲以上，共n-1個級距
    Const cst_ARE_TXT As String = ",15-17,18-19,20-24,25-29,30-34,35-39,40-44,45-49,50-54,55-59,60-64,65歲以上"

    Dim objconn As SqlConnection

    '#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '#Region "在這裡放置使用者程式碼以初始化網頁"
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.TestDbConn(Me, objConn)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            cCreate1()
        End If
        '#End Region
    End Sub

    Sub cCreate1()
        cblAPPSTAGE = TIMS.Get_APPSTAGE2(cblAPPSTAGE)
        cblAPPSTAGE.Attributes("onclick") = "SelectAll('cblAPPSTAGE','cblAPPSTAGE_Hidden');"

        BudgetList = TIMS.Get_Budget(BudgetList, 23)
        '選擇全部預算
        BudgetList.Attributes("onclick") = "SelectAll('BudgetList','BudgetHidden');"
        spBudgetList2.Style.Add("display", "none")

        'If (BudgetList.Items.Count > 0) Then
        '    For Each bd1 As ListItem In BudgetList.Items
        '        bd1.Attributes.Add("class", "GBudgetList")
        '    Next
        'End If
        'StudStatus.Attributes("onclick")="chk_StudStatus();"
        'StudStatus.Attributes("onchang")="chk_StudStatus();"
        Dim js_auto1 As String = "chk_StudStatus();"
        StudStatus.Attributes("onclick") = js_auto1 '"javascript:autorecsubtotal();"
        StudStatus.Attributes("onblur") = js_auto1 '"javascript:autorecsubtotal();"
        StudStatus.Attributes("onchange") = js_auto1 '"javascript:autorecsubtotal();"

        Dim s_jsc As String = "<script>" & js_auto1 & "</script>"
        TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), s_jsc)

        StudStatus.AppendDataBoundItems = True
        Common.SetListItem(StudStatus, "12")

        SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
        Common.SetListItem(SearchPlan, "G")

        STDate1.Text = TIMS.Cdate3(Now.Year & "/01/01")
        STDate2.Text = TIMS.Cdate3(Now.ToString("yyyy/MM/dd"))

        DataGroupTable.Visible = False
        Dim Sql As String = "SELECT DISTID,NAME FROM ID_DISTRICT WHERE DISTID!='000' ORDER BY DISTID"
        Dim dtDIST As DataTable = DbAccess.GetDataTable(Sql, objconn)
        DistID = TIMS.Get_DistID(DistID, dtDIST)

        trPlanKind.Style("display") = "none"
        trPackageType.Style("display") = "none"
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trPackageType.Style("display") = TIMS.cst_inline1 '"inline"
        Else
            '28:產業人才投資方案
            '計畫範圍 產投
            If sm.UserInfo.Years >= 2008 Then trPlanKind.Style("display") = TIMS.cst_inline1 '"inline"
        End If

        Button1.Attributes("onclick") = "return CheckSearch();"
        Button3.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"
    End Sub

    '與 SD_15_003_R.aspx 共用
    Public Shared Sub CreateData(ByVal dt As DataTable, ByVal XRollValue As String, ByVal YRollValue As String, ByVal YRollText As String,
                                 ByVal DataTable1 As Table, ByVal tConn As SqlConnection)
        '#Region "CreateData"

        Dim MyRow As TableRow
        Dim MyCell As TableCell
        Dim Key_Degree As DataTable = TIMS.Get_KeyTable("dbo.KEY_DEGREE", "DegreeType='1'", tConn)
        'Dim Key_Identity As DataTable
        Dim ID_City As DataTable = TIMS.Get_KeyTable("dbo.ID_CITY", tConn)
        Dim Key_Trade As DataTable = TIMS.Get_KeyTable("dbo.KEY_TRADE", tConn)
        Dim Key_ClassCatelog As DataTable = TIMS.Get_KeyTable("dbo.KEY_CLASSCATELOG", tConn)
        Dim Key_OrgType As DataTable = TIMS.Get_KeyTable("dbo.KEY_ORGTYPE", tConn)

        Dim iTotal As Integer = 0
        'Dim sql As String=""
        'sql=""
        'sql &= " SELECT IDENTITYID"
        'sql &= " ,CASE WHEN IDENTITYID='02' THEN CONVERT(NVARCHAR(30),'非自願離職者')"
        'sql &= "  WHEN IDENTITYID='10' THEN CONVERT(NVARCHAR(30),'其他(就服法24條)')"
        'sql &= "  ELSE NAME END NAME"
        'sql &= " FROM dbo.KEY_IDENTITY WITH(NOLOCK)"
        'sql &= " WHERE 1=1 AND IDENTITYID IN ( '01','04','05','06','07','09','10','11','13','14','26','28','40','30','33','34','37','43','31','32','36')"
        'sql &= " ORDER BY IDENTITYID "
        'Key_Identity=DbAccess.GetDataTable(sql, tConn)

        Dim sm As SessionModel = SessionModel.Instance()
        Dim Key_Identity As DataTable = TIMS.Get_dtIdentity(11, tConn, sm)
        Dim s_MIdentityID As String = ""
        s_MIdentityID = TIMS.Cst_Identity28_2019_11
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then s_MIdentityID = TIMS.Cst_Identity06_2019_11

        '29:天然災害受災民眾
        '33:中低收入戶
        'Cst_特定對象 TITLE
        ' TIMS.cst_Identity28b 該資料不能包含 (28 / 03) (28:獨力負擔家計者 / 03:負擔家計婦女)
        Dim sIdentityID_title As String = " IdentityID IN (" & s_MIdentityID & ")"
        'Dim sIdentityID_title2 As String=" IdentityID='28'" '28	獨力負擔家計者

        'Cst_特定對象 VALUE
        ' TIMS.cst_Identity28b 該資料不能包含 (28 / 03) (28:獨力負擔家計者 / 03:負擔家計婦女)
        Dim sIdentityID As String = " IdentityID IN (" & s_MIdentityID & ")" '090923 andy edit 加入28獨力負擔家計者29:天然災害受災民眾；
        'Dim sIdentityID_2 As String=" (MIdentityID='03' OR MIdentityID='28') " ''090923 andy edit 03負擔家計婦女統計結果併入--> 28獨力負擔家計者(拆開讓迴圈跑兩次)

        Select Case XRollValue
            Case Cst_性別
                MyRow = CreateRow(DataTable1)
                MyCell = CreateCell(MyRow, YRollText)
                MyCell.Width = Unit.Pixel(150)

                Select Case YRollValue
                    Case Cst_性別
                    Case Else
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                End Select

                Select Case YRollValue
                    Case Cst_性別
                        CreateCell(MyRow, "人數")
                        CreateCell(MyRow, "比率")
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, New DataView(dt, "Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F')", Nothing, DataViewRowState.CurrentRows).Count, 1)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, New DataView(dt, "Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F')", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Case Cst_年齡
                        Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                        For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, A_ARE_TXT(i_age))
                            CreateCell(MyRow, New DataView(dt, $"Sex='M' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, $"Sex='F' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"Sex IN ('M','F') and {Cst_ageIN}", Nothing, DataViewRowState.CurrentRows).Count)
                        Next

                    Case Cst_教育程度
                        For Each dr As DataRow In Key_Degree.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_特定對象
                        For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and MIdentityID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next

                    Case Cst_結訓後動向
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "轉換工作")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q3 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "留任")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q3 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "其他")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q3 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                    Case Cst_工作年資
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "5年以下")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "5~10年")
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "10~15年")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "15~20年")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "20~25年")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "25~30年")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "30年以上")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                    Case Cst_受訓學員地理分布
                        For Each dr As DataRow In ID_City.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("CTName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and CTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_所屬公司行業別
                        'MyCell.Width=Nothing
                        For Each dr As DataRow In Key_Trade.Rows
                            MyRow = CreateRow(DataTable1)
                            MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                            MyCell.HorizontalAlign = HorizontalAlign.Left
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q4 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_所屬公司規模
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "中小企業")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q5 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "非中小企業")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and Q5 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                    Case Cst_參訓動機
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "為補充與原專長相關之技能")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)

                        iTotal = New DataView(dt, "Sex IN ('M','F') and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Sex IN ('M','F') and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Sex IN ('M','F') and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Sex IN ('M','F') and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                        SubPercent(MyRow, iTotal)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "轉換其他行職業所需技能")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, iTotal)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "拓展工作領域及視野")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, iTotal)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "其他")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, iTotal)
                    Case Cst_參訓單位類別
                        For Each dr As DataRow In Key_OrgType.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_職能課程分類
                        For Each dr As DataRow In Key_ClassCatelog.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("CCName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and ClassCate IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_參加課程型態
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "學分班")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and PointYN IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "非學分班")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and PointYN IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "企業包班")
                        CreateCell(MyRow, New DataView(dt, "Sex='M' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, "Sex='F' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and IsBusiness IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                    Case Cst_外籍配偶類別
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, Cst_外配類別1)
                        CreateCell(MyRow, New DataView(dt, $"Sex='M'{Cst_sch_本國}", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, $"Sex='F'{Cst_sch_本國}", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, $"Sex IN ('M','F'){Cst_sch_本國外籍範圍}", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, Cst_外配類別2)
                        CreateCell(MyRow, New DataView(dt, $"Sex='M'{Cst_sch_外籍_大陸人士}", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, $"Sex='F'{Cst_sch_外籍_大陸人士}", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, $"Sex IN ('M','F'){Cst_sch_本國外籍範圍}", Nothing, DataViewRowState.CurrentRows).Count)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, Cst_外配類別3)
                        CreateCell(MyRow, New DataView(dt, $"Sex='M'{Cst_sch_外籍_非大陸人士}", Nothing, DataViewRowState.CurrentRows).Count)
                        CreateCell(MyRow, New DataView(dt, $"Sex='F'{Cst_sch_外籍_非大陸人士}", Nothing, DataViewRowState.CurrentRows).Count)
                        Subtotal(MyRow)
                        SubPercent(MyRow, New DataView(dt, $"Sex IN ('M','F'){Cst_sch_本國外籍範圍}", Nothing, DataViewRowState.CurrentRows).Count)
                End Select
            Case Cst_年齡
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                    For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, A_ARE_TXT(i_age))
                        CreateCell(MyRow, New DataView(dt, $"AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, Cst_ageIN, Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next

                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                    For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                        CreateCell(MyRow, A_ARE_TXT(i_age))
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"Sex='M' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and Sex IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"Sex='F' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and Sex IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                    CreateCell(MyRow, New DataView(dt, $"DegreeID='{dr("DegreeID")}' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} And MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and MIdentityID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            iTotal = New DataView(dt, $"{Cst_ageIN} and Q3 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                        Case Cst_工作年資
                            iTotal = New DataView(dt, $"{Cst_ageIN} and Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND CTID='{dr("CTID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and CTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            iTotal = New DataView(dt, $"{Cst_ageIN} and Q4 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q4='{dr("TradeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and Q5 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and Q5 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, $"{Cst_ageIN} and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, $"{Cst_ageIN} and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, $"{Cst_ageIN} and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, $"{Cst_ageIN} and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND ORGKIND='{dr("OrgTypeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND ClassCate='{dr("ccid")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and ClassCate IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and PointYN IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and PointYN IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"{Cst_ageIN} and IsBusiness IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age}{Cst_sch_本國}", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"AGE IS NOT NULL{Cst_sch_本國外籍範圍}", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age}{Cst_sch_外籍_大陸人士}", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"AGE IS NOT NULL{Cst_sch_本國外籍範圍}", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                CreateCell(MyRow, New DataView(dt, $"AGE={i_age}{Cst_sch_外籍_非大陸人士}", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, $"AGE IS NOT NULL{Cst_sch_本國外籍範圍}", Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If

            Case Cst_教育程度
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_Degree.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("Name").ToString)
                        CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_Degree.Rows
                        CreateCell(MyRow, dr("Name").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                For Each dr As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} and DegreeID='{dr("DegreeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_教育程度
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, $"DegreeID='{dr1("DegreeID")}' and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            iTotal = New DataView(dt, "DegreeID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)

                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "DegreeID IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "DegreeID IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "DegreeID IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "DegreeID IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_特定對象
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("Name").ToString)
                        CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next

                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_Identity.Select(sIdentityID_title) '20090923 andy edit
                        CreateCell(MyRow, dr("Name").ToString)
                    Next
                    'CreateCell(MyRow, Key_Identity.Select(sIdentityID_title2)(0)("Name").ToString) '負擔家計婦女併至獨立負擔家計者
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, $"Sex='M' and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            '20090923 andy edit'負擔家計婦女併至獨立負擔家計者
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, $"Sex='F' and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            '20090923 andy edit '負擔家計婦女併至獨立負擔家計者
                            'CreateCell(MyRow, New DataView(dt, "Sex='F'  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                '20090923 andy edit'負擔家計婦女併至獨立負擔家計者
                                'CreateCell(MyRow, New DataView(dt, $"AGE={i_age} and " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Identity.Select(sIdentityID)
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                'CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "'  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q3=1" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q3=2" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q3=3" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=1" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=2" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=3" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=4" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=5" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=6" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q61=7" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Select(sIdentityID)
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                'CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "'" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)   '20090923 andy edit
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                For Each dr1 As DataRow In Key_Identity.Select(sIdentityID)
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                'CreateCell(MyRow, New DataView(dt, sIdentityID_2 & " and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q5=1" & "  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)   '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q5=0  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "MIdentityID IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "MIdentityID IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "MIdentityID IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "MIdentityID IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q21=1  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q22=2  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q23=3  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Q24=4  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Identity.Select(sIdentityID)
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                'CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and  orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Select(sIdentityID)
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                'CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "'  and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and  " & sIdentityID_2, Nothing, DataViewRowState.CurrentRows).Count)  '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, sIdentityID_2 & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)   '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, sIdentityID_2 & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)   '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, sIdentityID_2 & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)    '20090923 andy edit
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_結訓後動向
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    'debug by nick 060301
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "轉換工作")
                    CreateCell(MyRow, New DataView(dt, "Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q3 is not null", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "留任")
                    CreateCell(MyRow, New DataView(dt, "Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "其他")
                    CreateCell(MyRow, New DataView(dt, "Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "轉換工作")
                    CreateCell(MyRow, "留任")
                    CreateCell(MyRow, "其他")
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            iTotal = New DataView(dt, "Q3 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                CreateCell(MyRow, New DataView(dt, $"Q3=1 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q3=2 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q3=3 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next

                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, $"Q3=1 and DegreeID='{dr("DegreeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q3=2 and DegreeID='{dr("DegreeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q3=3 and DegreeID='{dr("DegreeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, $"Q3=1 and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q3=2 and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q3=3 and MIdentityID='{dr("IdentityID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q3=1 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=2 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=3 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                CreateCell(MyRow, New DataView(dt, "Q3=1 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=2 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=3 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "Q3 IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q3 IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q3 IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q3 IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q3=1 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=2 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=3 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q3=1 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=2 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q3=3 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            CreateCell(MyRow, New DataView(dt, "Q3=1 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            CreateCell(MyRow, New DataView(dt, "Q3=1" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            CreateCell(MyRow, New DataView(dt, "Q3=1" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            CreateCell(MyRow, New DataView(dt, "Q3=1" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=2" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q3=3" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_工作年資
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "5年以下")
                    CreateCell(MyRow, New DataView(dt, "Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "5~10年")
                    CreateCell(MyRow, New DataView(dt, "Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "10~15年")
                    CreateCell(MyRow, New DataView(dt, "Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "15~20年")
                    CreateCell(MyRow, New DataView(dt, "Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "20~25年")
                    CreateCell(MyRow, New DataView(dt, "Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "25~30年")
                    CreateCell(MyRow, New DataView(dt, "Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "30年以上")
                    CreateCell(MyRow, New DataView(dt, "Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "5年以下")
                    CreateCell(MyRow, "5~10年")
                    CreateCell(MyRow, "10~15年")
                    CreateCell(MyRow, "15~20年")
                    CreateCell(MyRow, "20~25年")
                    CreateCell(MyRow, "25~30年")
                    CreateCell(MyRow, "30年以上")
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                CreateCell(MyRow, New DataView(dt, $"Q61=1 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q61=2 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q61=3 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q61=4 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q61=5 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q61=6 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q61=7 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and AGE IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q61=1 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=2 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=3 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=4 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=5 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=6 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=7 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q61=1 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=2 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=3 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=4 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=5 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=6 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=7 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q61=1 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=2 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=3 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=4 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=5 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=6 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=7 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                CreateCell(MyRow, New DataView(dt, "Q61=1 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=2 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=3 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=4 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=5 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=6 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=7 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "Q61 IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q61 IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q61 IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q61 IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q61=1 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=2 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=3 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=4 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=5 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=6 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=7 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q61=1 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=2 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=3 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=4 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=5 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=6 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q61=7 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            CreateCell(MyRow, New DataView(dt, "Q61=1 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            CreateCell(MyRow, New DataView(dt, "Q61=1" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            CreateCell(MyRow, New DataView(dt, "Q61=1" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            CreateCell(MyRow, New DataView(dt, "Q61=1" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=2" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=3" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=4" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=5" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=6" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q61=7" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q61 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_受訓學員地理分布
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In ID_City.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("CTName").ToString)
                        CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID").ToString & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In ID_City.Rows
                        CreateCell(MyRow, dr("CTName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                For Each dr As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} and CTID='{dr("CTID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15歲以下")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=1 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=2 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=3 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=4 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=5 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=6 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'For Each dr As DataRow In ID_City.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=7 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "CTID IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "CTID IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "CTID IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "CTID IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID") & "'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_所屬公司行業別
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(300)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_Trade.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("TradeName").ToString)
                        MyCell.Width = Unit.Pixel(300)
                        MyCell.HorizontalAlign = HorizontalAlign.Left
                        CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    For Each dr As DataRow In Key_Trade.Rows
                        CreateCell(MyRow, dr("TradeName").ToString)
                        MyCell.Width = Unit.Pixel(300)
                        MyCell.HorizontalAlign = HorizontalAlign.Left
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                For Each dr As DataRow In Key_Trade.Rows
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} and Q4='{dr("TradeID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15歲以下")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=1 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=2 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=3 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=4 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=5 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=6 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'For Each dr As DataRow In Key_Trade.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=7 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Trade.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and Q4='" & dr1("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Trade.Rows
                                    CreateCell(MyRow, New DataView(dt, "Q4='" & dr1("TradeID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Trade.Rows
                                    CreateCell(MyRow, New DataView(dt, "Q4='" & dr1("TradeID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "Q4 IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q4 IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q4 IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q4 IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Trade.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q4='" & dr1("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                For Each dr1 As DataRow In Key_Trade.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q4='" & dr1("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For Each dr As DataRow In Key_Trade.Rows
                                CreateCell(MyRow, New DataView(dt, "Q4='" & dr("TradeID") & "'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q4 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_所屬公司規模
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "中小企業")
                    CreateCell(MyRow, New DataView(dt, "Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "非中小企業")
                    CreateCell(MyRow, New DataView(dt, "Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "中小企業")
                    CreateCell(MyRow, "非中小企業")
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                CreateCell(MyRow, New DataView(dt, $"Q5=1 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q5=0 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q5=1 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q5=0 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q5=1 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q5=0 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q5=1 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q5=0 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.Width = Unit.Pixel(300)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                CreateCell(MyRow, New DataView(dt, "Q5=1 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q5=0 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "Q5 IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q5 IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q5 IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q5 IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q5=1 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q5=0 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q5=1 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q5=0 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            CreateCell(MyRow, New DataView(dt, "Q5=1 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            CreateCell(MyRow, New DataView(dt, "Q5=1" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            CreateCell(MyRow, New DataView(dt, "Q5=1" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            CreateCell(MyRow, New DataView(dt, "Q5=1" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q5=0" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Q5 IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_參訓動機
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    iTotal = New DataView(dt, "Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "為補充與原專長相關之技能")
                    CreateCell(MyRow, New DataView(dt, "Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, iTotal, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "轉換其他行職業所需技能")
                    CreateCell(MyRow, New DataView(dt, "Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, iTotal, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "拓展工作領域及視野")
                    CreateCell(MyRow, New DataView(dt, "Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, iTotal, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "其他")
                    CreateCell(MyRow, New DataView(dt, "Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, iTotal, 1)
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    CreateCell(MyRow, "為補充與原專長相關之技能")
                    CreateCell(MyRow, "轉換其他行職業所需技能")
                    CreateCell(MyRow, "拓展工作領域及視野")
                    CreateCell(MyRow, "其他")
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            iTotal = New DataView(dt, "Sex IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Sex IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Sex IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Sex IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_年齡
                            iTotal = New DataView(dt, $"{Cst_ageIN} AND Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, $"{Cst_ageIN} and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, $"{Cst_ageIN} and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, $"{Cst_ageIN} and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                CreateCell(MyRow, New DataView(dt, $"Q21=1 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q22=2 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q23=3 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"Q24=4 and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15歲以下")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'CreateCell(MyRow, New DataView(dt, "Q21=1 and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q22=2 and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q23=3 and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Q24=4 and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, iTotal)
                        Case Cst_教育程度
                            iTotal = New DataView(dt, "DegreeID IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "DegreeID IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "DegreeID IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "DegreeID IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q21=1 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q22=2 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q23=3 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q24=4 and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                        Case Cst_特定對象
                            iTotal = New DataView(dt, "MIdentityID IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "MIdentityID IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "MIdentityID IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "MIdentityID IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q21=1 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q22=2 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q23=3 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q24=4 and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next

                        Case Cst_結訓後動向
                            iTotal = New DataView(dt, "Q3 IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q3 IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q3 IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q3 IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_工作年資
                            iTotal = New DataView(dt, "Q61 IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q61 IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q61 IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q61 IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_受訓學員地理分布
                            iTotal = New DataView(dt, "CTID IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "CTID IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "CTID IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "CTID IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q21=1 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q22=2 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q23=3 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q24=4 and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                        Case Cst_所屬公司行業別
                            iTotal = New DataView(dt, "Q4 IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q4 IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q4 IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q4 IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                CreateCell(MyRow, New DataView(dt, "Q21=1 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q22=2 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q23=3 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q24=4 and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                        Case Cst_所屬公司規模
                            iTotal = New DataView(dt, "Q5 IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q5 IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q5 IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "Q5 IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓動機
                        Case Cst_參訓單位類別
                            iTotal = New DataView(dt, "orgkind IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "orgkind IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "orgkind IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "orgkind IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q21=1 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q22=2 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q23=3 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q24=4 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                        Case Cst_職能課程分類
                            iTotal = New DataView(dt, "ClassCate IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "ClassCate IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "ClassCate IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "ClassCate IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                CreateCell(MyRow, New DataView(dt, "Q21=1 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q22=2 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q23=3 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "Q24=4 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next
                        Case Cst_參加課程型態
                            iTotal = New DataView(dt, "PointYN IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "PointYN IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "PointYN IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "PointYN IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            iTotal = New DataView(dt, "IsBusiness IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "IsBusiness IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "IsBusiness IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "IsBusiness IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            CreateCell(MyRow, New DataView(dt, "Q21=1 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4 and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_外籍配偶類別
                            iTotal = 0
                            iTotal += New DataView(dt, "Q21 IS Not NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count
                            iTotal += New DataView(dt, "Q22 IS Not NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count
                            iTotal += New DataView(dt, "Q23 IS Not NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count
                            iTotal += New DataView(dt, "Q24 IS Not NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            CreateCell(MyRow, New DataView(dt, "Q21=1" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            CreateCell(MyRow, New DataView(dt, "Q21=1" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            iTotal = New DataView(dt, "IsBusiness IS Not NULL and Q21 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "IsBusiness IS Not NULL and Q22 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "IsBusiness IS Not NULL and Q23 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "IsBusiness IS Not NULL and Q24 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            CreateCell(MyRow, New DataView(dt, "Q21=1" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q22=2" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q23=3" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Q24=4" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                    End Select
                End If
            Case Cst_參訓單位類別
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_OrgType.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("Name").ToString)
                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_OrgType.Rows
                        CreateCell(MyRow, dr("Name").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, " orgkind IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            iTotal = New DataView(dt, "ORGKIND IS NOT NULL AND AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                For Each dr As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, $"ORGKIND='{dr("OrgTypeID")}' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next

                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'For Each dr As DataRow In Key_OrgType.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=2 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'For Each dr As DataRow In Key_OrgType.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=3 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'For Each dr As DataRow In Key_OrgType.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=4 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'For Each dr As DataRow In Key_OrgType.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=5 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'For Each dr As DataRow In Key_OrgType.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=6 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'For Each dr As DataRow In Key_OrgType.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=7 and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "orgkind IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "orgkind IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "orgkind IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "orgkind IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_職能課程分類
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_ClassCatelog.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("CCName").ToString)
                        CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_ClassCatelog.Rows
                        CreateCell(MyRow, dr("CCName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            iTotal = New DataView(dt, "CLASSCATE IS NOT NULL AND AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                For Each dr As DataRow In Key_ClassCatelog.Rows
                                    CreateCell(MyRow, New DataView(dt, $"AGE={i_age} AND CLASSCATE='{dr("CCID")}'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next

                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15歲以下")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=1 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=2 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=3 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=4 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=5 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=6 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'For Each dr As DataRow In Key_ClassCatelog.Rows
                            '    CreateCell(MyRow, New DataView(dt, "AGE=7 and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_ClassCatelog.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr1("ccid") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_ClassCatelog.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr1("ccid") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_ClassCatelog.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr1("ccid") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                For Each dr1 As DataRow In Key_ClassCatelog.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr1("ccid") & "' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "ClassCate IS NOT NULL and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "ClassCate IS NOT NULL and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "ClassCate IS NOT NULL and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "ClassCate IS NOT NULL and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_ClassCatelog.Rows
                                    CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr1("ccid") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                        Case Cst_參加課程型態
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and PointYN IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL and IsBusiness IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_外籍配偶類別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                CreateCell(MyRow, New DataView(dt, "ClassCate='" & dr("ccid") & "'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "ClassCate IS NOT NULL" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count)
                    End Select
                End If
            Case Cst_參加課程型態
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "學分班")
                    CreateCell(MyRow, New DataView(dt, "PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "PointYN is not null", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "非學分班")
                    CreateCell(MyRow, New DataView(dt, "PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "PointYN IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, "企業包班")
                    CreateCell(MyRow, New DataView(dt, "IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "IsBusiness IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "學分班")
                    CreateCell(MyRow, "非學分班")
                    CreateCell(MyRow, "企業包班")
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            iTotal = New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                CreateCell(MyRow, New DataView(dt, $"PointYN='Y' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"PointYN='N' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"IsBusiness='Y' and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, iTotal)
                            Next

                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15歲以下")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'CreateCell(MyRow, New DataView(dt, "PointYN='Y' and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "PointYN='N' and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "PointYN='Y' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "PointYN='N' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "PointYN='Y' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "PointYN='N' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                CreateCell(MyRow, New DataView(dt, "PointYN='Y' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "PointYN='N' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, "PointYN='Y' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "PointYN='N' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                CreateCell(MyRow, New DataView(dt, "PointYN='Y' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "PointYN='N' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, "IsBusiness='Y' and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL) and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                        Case Cst_外籍配偶類別
                            iTotal = 0
                            iTotal += New DataView(dt, "(PointYN IS NOT NULL or IsBusiness IS NOT NULL)" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別1)
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y'" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別2)
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y'" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, Cst_外配類別3)
                            CreateCell(MyRow, New DataView(dt, "PointYN='Y'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "PointYN='N'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "IsBusiness='Y'" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                    End Select
                End If
            Case Cst_外籍配偶類別
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, Cst_外配類別1)
                    CreateCell(MyRow, New DataView(dt, "1=1" & Cst_sch_本國, Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "1=1" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, Cst_外配類別2)
                    CreateCell(MyRow, New DataView(dt, "1=1" & Cst_sch_外籍_大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "1=1" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count, 1)
                    MyRow = CreateRow(DataTable1)
                    CreateCell(MyRow, Cst_外配類別3)
                    CreateCell(MyRow, New DataView(dt, "1=1" & Cst_sch_外籍_非大陸人士, Nothing, DataViewRowState.CurrentRows).Count)
                    SubPercent(MyRow, New DataView(dt, "1=1" & Cst_sch_本國外籍範圍, Nothing, DataViewRowState.CurrentRows).Count, 1)
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, Cst_外配類別1)
                    CreateCell(MyRow, Cst_外配類別2)
                    CreateCell(MyRow, Cst_外配類別3)
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            Dim A_ARE_TXT As String() = cst_ARE_TXT.Split(",")
                            For i_age As Integer = 1 To A_ARE_TXT.Length - 1
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, A_ARE_TXT(i_age))
                                CreateCell(MyRow, New DataView(dt, $"{Cst_sch_本國x}and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"{Cst_sch_外籍_大陸人士x}and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, $"{Cst_sch_外籍_非大陸人士x}and AGE={i_age}", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, $"{Cst_sch_本國外籍範圍x}and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15歲以下")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "15~24")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "25~34")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "35~44")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Q3 IS NOT NULL and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "45~54")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "55~64")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=6", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'MyRow=CreateRow(DataTable1)
                            'CreateCell(MyRow, "65歲以上")
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and AGE=7", Nothing, DataViewRowState.CurrentRows).Count)
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and AGE IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_特定對象
                            For Each dr As DataRow In Key_Identity.Select(sIdentityID)
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_結訓後動向
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換工作")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q3=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "留任")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q3=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q3=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q3 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_工作年資
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5年以下")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "5~10年")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "10~15年")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "15~20年")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "20~25年")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=5", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "25~30年")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "30年以上")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q61=7", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q61 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司行業別
                            MyCell.Width = Nothing
                            For Each dr As DataRow In Key_Trade.Rows
                                MyRow = CreateRow(DataTable1)
                                MyCell = CreateCell(MyRow, dr("TradeName").ToString)
                                MyCell.HorizontalAlign = HorizontalAlign.Left
                                CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q4='" & dr("TradeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q4 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_所屬公司規模
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "中小企業")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q5=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非中小企業")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q5=0", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and Q5 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_參訓動機
                            iTotal = New DataView(dt, Cst_sch_本國外籍範圍x & "and Q21 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, Cst_sch_本國外籍範圍x & "and Q22 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, Cst_sch_本國外籍範圍x & "and Q23 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count + New DataView(dt, Cst_sch_本國外籍範圍x & "and Q24 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "為補充與原專長相關之技能")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q21=1", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "轉換其他行職業所需技能")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q22=2", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "拓展工作領域及視野")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q23=3", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "其他")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and Q24=4", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_職能課程分類
                            For Each dr As DataRow In Key_ClassCatelog.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CCName").ToString)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and ClassCate='" & dr("ccid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, Cst_sch_本國外籍範圍x & "and ClassCate IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_參加課程型態
                            iTotal = 0
                            iTotal += New DataView(dt, Cst_sch_本國外籍範圍x & "and (PointYN IS NOT NULL or IsBusiness IS NOT NULL)", Nothing, DataViewRowState.CurrentRows).Count
                            MyCell.Width = Nothing
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "學分班")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and PointYN='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "非學分班")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and PointYN='N'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "企業包班")
                            CreateCell(MyRow, New DataView(dt, Cst_sch_本國x & "and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_大陸人士x & "and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, Cst_sch_外籍_非大陸人士x & "and IsBusiness='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Case Cst_外籍配偶類別
                    End Select
                End If
        End Select

        '小計
        MyRow = CreateRow(DataTable1)
        CreateCell(MyRow, "小計")
        For i As Integer = 1 To DataTable1.Rows(0).Cells.Count - 1
            iTotal = 0
            For j As Integer = 1 To DataTable1.Rows.Count - 2
                If IsNumeric(DataTable1.Rows(j).Cells(i).Text) Then
                    iTotal += DataTable1.Rows(j).Cells(i).Text
                End If
            Next

            If i = DataTable1.Rows(0).Cells.Count - 1 Then
                If DataTable1.Rows(DataTable1.Rows.Count - 1).Cells(DataTable1.Rows(0).Cells.Count - 2).Text = 0 Then
                    CreateCell(MyRow, "0%")
                Else
                    CreateCell(MyRow, "100%")
                End If
            Else
                CreateCell(MyRow, iTotal)
            End If
        Next
    End Sub

    Public Shared Sub Subtotal(ByVal MyRow As TableRow)
        Dim iTotal As Integer = 0
        For i As Integer = 1 To MyRow.Cells.Count - 1
            'iTotal += Int(TIMS.GetValue2(MyRow.Cells(i).Text))
            iTotal += Int(MyRow.Cells(i).Text)
        Next
        CreateCell(MyRow, iTotal)
    End Sub

    Public Shared Sub SubPercent(ByVal MyRow As TableRow, ByVal iRecordCount As Integer, Optional ByVal iEndNum As Integer = 2)
        Dim iTotal As Integer = 0
        For i As Integer = 1 To MyRow.Cells.Count - iEndNum
            Dim i_VPNUM As Integer = 0
            If TIMS.IsNumeric1(MyRow.Cells(i).Text) Then i_VPNUM = Val(MyRow.Cells(i).Text)
            iTotal += i_VPNUM
        Next

        If iRecordCount = 0 Then
            CreateCell(MyRow, 0)
        Else
            CreateCell(MyRow, Math.Round(iTotal * 100 / iRecordCount, 2) & "%")
        End If
    End Sub

    Public Shared Function CreateRow(ByVal DataTable1 As Table) As TableRow
        Dim MyRow As New TableRow
        DataTable1.Rows.Add(MyRow)
        Return MyRow
    End Function

    Public Shared Function CreateCell(ByRef MyRow As TableRow, ByVal MyText As String) As TableCell
        Dim MyCell As New TableCell
        MyRow.Cells.Add(MyCell)
        MyCell.Text = MyText
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.HorizontalAlign = HorizontalAlign.Center
        Return MyCell
    End Function

    '查詢檢核
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)

        'If Trim(STDate1.Text) <> "" Then STDate1.Text=Trim(STDate1.Text) Else STDate1.Text=""
        'If Trim(STDate2.Text) <> "" Then STDate2.Text=Trim(STDate2.Text) Else STDate2.Text=""
        'If Trim(FTDate1.Text) <> "" Then FTDate1.Text=Trim(FTDate1.Text) Else FTDate1.Text=""
        'If Trim(FTDate2.Text) <> "" Then FTDate2.Text=Trim(FTDate2.Text) Else FTDate2.Text=""

        If STDate1.Text <> "" Then
            If Not TIMS.IsDate1(STDate1.Text) Then Errmsg += "開訓期間 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
        Else
            'Errmsg += "開訓期間 起始日期 為必填" & vbCrLf
        End If

        If STDate2.Text <> "" Then
            If Not TIMS.IsDate1(STDate2.Text) Then Errmsg += "開訓期間 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
        Else
            'Errmsg += "開訓期間 迄止日期 為必填" & vbCrLf
        End If
        If FTDate1.Text <> "" Then
            If Not TIMS.IsDate1(FTDate1.Text) Then Errmsg += "結訓期間 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
        Else
            'Errmsg += "結訓期間 起始日期 為必填" & vbCrLf
        End If

        If FTDate2.Text <> "" Then
            If Not TIMS.IsDate1(FTDate2.Text) Then Errmsg += "結訓期間 迄止日期格式有誤" & vbCrLf
            If Errmsg = "" Then FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
        Else
            'Errmsg += "結訓期間 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    ''' <summary> sql語法 </summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function Search1_sql(ByRef parms As Hashtable) As String
        'Dim dt As DataTable=Nothing
        'Dim parms As Hashtable=New Hashtable()
        Dim V_StudStatus As String = TIMS.GetListValue(StudStatus)

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.STDATE,cc.FTDATE,oo.ORGKIND,pp.CLASSCATE,pp.PointYN,pp.IsBusiness,cc.OCID,cc.NOTOPEN" & vbCrLf
        sql &= "  FROM dbo.ID_PLAN ip" & vbCrLf
        sql &= "  JOIN dbo.PLAN_PLANINFO pp on ip.planid=pp.planid" & vbCrLf
        sql &= "  JOIN dbo.ORG_ORGINFO oo on oo.comidno=pp.comidno" & vbCrLf
        sql &= "  JOIN dbo.CLASS_CLASSINFO cc on cc.planid=pp.planid and cc.COMIDNO=pp.COMIDNO and cc.seqno=pp.seqno" & vbCrLf
        sql &= "  WHERE pp.IsApprPaper='Y'" & vbCrLf
        'And ip.DISTID='001'and ip.YEARS='2018'and ip.TPLANID='28'
        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'If SearchPlan.SelectedIndex <> 0 Then sql &= " and oo.OrgKind2 in ('" & SearchPlan.SelectedValue & "')" & vbCrLf
            Select Case TIMS.GetListValue(SearchPlan)'.SelectedValue
                Case "A"
                    sql &= " and oo.OrgKind2 IN ('G','W')" & vbCrLf
                Case "G"
                    sql &= " and oo.OrgKind2='G'" & vbCrLf
                Case "W"
                    sql &= " and oo.OrgKind2='W'" & vbCrLf
            End Select
        End If

        Dim v_cblAPPSTAGE As String = TIMS.GetCblValue(cblAPPSTAGE)
        Dim v_cblAPPSTAGE_in As String = TIMS.CombiSQLINM3(v_cblAPPSTAGE)
        If v_cblAPPSTAGE <> "" AndAlso v_cblAPPSTAGE_in <> "" Then
            sql &= $" AND pp.APPSTAGE IN ({v_cblAPPSTAGE_in})" & vbCrLf
        End If

        '54:充電起飛計畫（在職）判斷方式
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If v_PackageType <> "A" Then
                sql &= " and pp.PackageType=@PackageType" & vbCrLf
                parms.Add("PackageType", v_PackageType)
            End If
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
        Else
            sql &= " AND ip.TPlanID In (" & TIMS.Cst_TPlanID28AppPlan & ")" & vbCrLf
        End If
        If DistID.SelectedIndex <> 0 Then
            sql &= " and ip.DistID=@DistID" & vbCrLf
            parms.Add("DistID", DistID.SelectedValue)
        End If
        If RIDValue.Value <> "" Then
            sql &= " and pp.RID=@RID" & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If
        If PlanID.Value <> "" Then
            sql &= " and ip.PlanID=@PlanID" & vbCrLf
            parms.Add("PlanID", PlanID.Value)
        End If
        If STDate1.Text <> "" Then
            'SearchStr &= " and a.STDate>=convert(datetime, '" & STDate1.Text & "', 111)" & vbCrLf
            sql &= " and cc.STDate>=CONVERT(date, @STDate1)" & vbCrLf
            parms.Add("STDate1", STDate1.Text)
        End If
        If STDate2.Text <> "" Then
            'SearchStr &= " and a.STDate<=convert(datetime, '" & STDate2.Text & "', 111)" & vbCrLf
            sql &= " and cc.STDate<=CONVERT(date, @STDate2)" & vbCrLf
            parms.Add("STDate2", STDate2.Text)
        End If
        If FTDate1.Text <> "" Then
            'SearchStr &= " and a.FTDate>=convert(datetime, '" & FTDate1.Text & "', 111)" & vbCrLf
            sql &= " and cc.FTDate>=CONVERT(date, @FTDate1)" & vbCrLf
            parms.Add("FTDate1", FTDate1.Text)
        End If
        If FTDate2.Text <> "" Then
            'SearchStr &= " and a.FTDate<=convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
            sql &= " and cc.FTDate<=CONVERT(date, @FTDate2)" & vbCrLf
            parms.Add("FTDate2", FTDate2.Text)
        End If
        sql &= " )" & vbCrLf

        'Dim sql As String=""
        Select Case V_StudStatus' StudStatus.SelectedValue
            Case Cst_報名人數
                sql &= " SELECT st.SEX ,st.DegreeID" & vbCrLf
            Case Else 'Cst_參訓人數/Cst_結訓人數/Cst_撥款人數
                sql &= " SELECT ss.SEX ,ss.DegreeID" & vbCrLf
        End Select
        '年齡
        sql &= " ,dbo.FN_YEARSOLDID2D2(DATEDIFF(YEAR,st.BIRTHDAY,CC.STDATE)) AGE" & vbCrLf
        sql &= " ,cc.ORGKIND,cc.CLASSCATE,cc.PointYN,cc.IsBusiness,b.MIdentityID,b.IsApprPaper" & vbCrLf
        sql &= " ,b.STUDSTATUS,b.REJECTTDATE1,b.REJECTTDATE2,b.BudgetID,ssc.AppliedStatus" & vbCrLf
        '結訓後動向
        sql &= " ,e.Q3" & vbCrLf
        '工作年資
        sql &= " ,case when e.Q61<5 then 1 when e.Q61<10 then 2 when e.Q61<15 then 3 when e.Q61<20 then 4" & vbCrLf
        sql &= " when e.Q61<25 then 5 when e.Q61<30 then 6 when e.Q61<35 then 7 end Q61" & vbCrLf
        '受訓學員地理分布
        sql &= " ,g.CTID" & vbCrLf
        '所屬公司行業別
        sql &= " ,e.Q4" & vbCrLf
        '所屬公司規模
        sql &= " ,e.Q5" & vbCrLf
        '參訓動機
        sql &= " ,h.Q21" & vbCrLf
        '參訓單位類別
        sql &= " ,h.Q22" & vbCrLf
        sql &= " ,h.Q23" & vbCrLf
        sql &= " ,h.Q24" & vbCrLf
        sql &= " ,case when ss.PassPortNO is not null and isnull(ss.ChinaOrNot,0) in (1,2) then 2" & vbCrLf
        sql &= "  else case when ss.PassPortNO is not null then 1 else 0 end end PassPortNO" & vbCrLf
        'ChinaOrNot 1:大陸人士 /2:非大陸人士
        sql &= " ,ss.ChinaOrNot" & vbCrLf

        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN ( SELECT a.IDNO,b.OCID1 ,MAX(a.Birthday) Birthday ,MAX(a.SEX) SEX ,MAX(a.DegreeID) DegreeID" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.OCID1=cc.OCID" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTEMP2 a WITH(NOLOCK) ON b.eSETID=a.eSETID" & vbCrLf
        sql &= " GROUP BY a.IDNO,b.OCID1" & vbCrLf
        sql &= " ) st ON st.OCID1=cc.OCID" & vbCrLf
        Select Case V_StudStatus' StudStatus.SelectedValue
            Case Cst_報名人數
                sql &= " LEFT JOIN dbo.STUD_STUDENTINFO ss WITH(NOLOCK) ON ss.IDNO=st.IDNO" & vbCrLf
            Case Else 'Cst_參訓人數/Cst_結訓人數/Cst_撥款人數
                sql &= " JOIN dbo.STUD_STUDENTINFO ss WITH(NOLOCK) ON ss.IDNO=st.IDNO" & vbCrLf
        End Select
        sql &= " LEFT JOIN dbo.CLASS_STUDENTSOFCLASS b WITH(NOLOCK) on b.OCID=cc.OCID AND b.SID=ss.SID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SUBSIDYCOST ssc WITH(NOLOCK) ON ssc.SOCID=b.SOCID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_TRAINBG e WITH(NOLOCK) ON b.SOCID=e.SOCID" & vbCrLf 'ENTERTRAIN2
        sql &= " LEFT JOIN dbo.STUD_SUBDATA f WITH(NOLOCK) ON f.SID=ss.SID" & vbCrLf

        sql &= " LEFT JOIN dbo.ID_ZIP g WITH(NOLOCK) ON f.ZipCode1=g.ZipCode" & vbCrLf
        sql &= " LEFT JOIN dbo.V_TRAINBGQ2 h WITH(NOLOCK) ON b.SOCID=h.SOCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        Dim sql_budval As String = TIMS.CombiSQM2IN(TIMS.GetCblValue(BudgetList))
        'Dim sql_budwhere As String=" and b.BudgetID IN ('03','02','01','97')" & vbCrLf
        Dim sql_budwhere As String = "" '註：如未勾選則統計全部預算別,含未選擇預算別或不補助
        If sql_budval <> "" Then sql_budwhere = String.Format(" and b.BudgetID IN ({0})", sql_budval) & vbCrLf

        Select Case V_StudStatus' StudStatus.SelectedValue
            Case Cst_報名人數
                sql &= " And cc.NOTOPEN='N'" & vbCrLf
            Case Cst_參訓人數 '12
                '/*開訓人次-合計*/ Cst_實際開訓人次加總
                '班級開訓，學員開訓後14天實際錄訓人數,且有選擇預算別(公務/就安/就保或公務(ECFA))
                sql &= " and cc.NOTOPEN='N' and b.IsApprPaper='Y'" & vbCrLf
                sql &= " and dbo.FN_GET_STUDCNT14B(b.STUDSTATUS,b.REJECTTDATE1,b.REJECTTDATE2,cc.STDATE)=1" & vbCrLf
                If sql_budwhere <> "" Then sql &= sql_budwhere & vbCrLf
                'sql &= " and cc.NOTOPEN='N' AND b.STUDSTATUS IN (1,5)" & vbCrLf'班級開訓，學員沒有離退訓，只剩開結訓人數
            Case Cst_結訓人數 '13
                '/*結訓-合計人次*/Cst_結訓人次
                '班級開訓，學員補助符合補助者 沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))
                sql &= " and cc.NOTOPEN='N' and b.IsApprPaper='Y'" & vbCrLf
                sql &= " and b.CreditPoints IS NOT NULL and b.STUDSTATUS NOT IN (2,3)" & vbCrLf
                sql &= " and cc.FTDate < GETDATE()" & vbCrLf
                If sql_budwhere <> "" Then sql &= sql_budwhere & vbCrLf
                'sql &= " and cc.NOTOPEN='N' AND b.STUDSTATUS IN (5) AND cc.FTDate < GETDATE()" & vbCrLf '班級開訓，學員已結訓人數
            Case Cst_撥款人數 '14
                '/*協助合計撥款人次*/Cst_撥款人次
                '班級開訓，學員補助符合補助者 沒有離退訓，只剩開結訓人數，且結訓日期已過今天,且有選擇預算別(公務/就安/就保或公務(ECFA))-學員經費撥款狀態：已撥款之人數 
                sql &= " and cc.NOTOPEN='N' and b.IsApprPaper='Y'" & vbCrLf
                sql &= " and b.CreditPoints IS NOT NULL and b.STUDSTATUS NOT IN (2,3)" & vbCrLf
                sql &= " and cc.FTDate < GETDATE()" & vbCrLf
                If sql_budwhere <> "" Then sql &= sql_budwhere & vbCrLf
                sql &= " and ssc.AppliedStatus='1'" & vbCrLf
                'sql &= " and cc.NOTOPEN='N' and ssc.APPLIEDSTATUS=1" & vbCrLf '撥款人數->學員經費撥款狀態 班級開訓，學員經費撥款狀態-已撥款人數
        End Select
        Return sql
    End Function

    ''' <summary> 查詢 (SQL) </summary>
    Sub Search1()
        '#Region "search1"
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim dt As DataTable = Nothing
        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = Search1_sql(parms)

        Try
            dt = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢!!")
            Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg &= TIMS.GetErrorMsg(Page) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try
        If dt Is Nothing Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        ViewState("SD_15_023_SqlStr") = sql
        ViewState("SD_15_023_parms") = parms
        DataGroupTable.Visible = True
        Call CreateData(dt, XRoll.SelectedValue, YRoll.SelectedValue, YRoll.SelectedItem.Text, DataTable1, objconn)
    End Sub

    ''' <summary>
    ''' 匯出xls/ods
    ''' </summary>
    Sub ExportDiv1()
        Call Search1()

        'Dim sTitle1 As String=""
        'sTitle1=CStr(sm.UserInfo.Years - 1911) & "年度－交叉分析統計表" '標題抬頭
        'Response.Clear()
        'Response.ClearHeaders()
        'Response.Buffer=True
        'Response.Charset="UTF-8" '"BIG5"
        ''Response.ContentType="Application/octet-stream"
        ''Response.ContentType="application/vnd.ms-excel"
        'Response.ContentType="application/ms-excel;charset=utf-8" '內容型態設為Excel
        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(sFileName1, System.Text.Encoding.UTF8) & ".xls")
        ''Response.ContentEncoding=System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding=System.Text.Encoding.GetEncoding("UTF-8")

        Const cst_TitleS1 As String = "交叉分析統計表" '檔案名
        Dim sFileName1 As String = cst_TitleS1 & TIMS.GetToday(objconn)

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType)
        Dim strHTML As String = ""
        Select Case v_ExpType
            Case "EXCEL"
                strHTML &= ("<html>")
                strHTML &= ("<head>")
                'strHTML &=("<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
                strHTML &= ("<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
                '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
                '套CSS值
                'mso-number-format:"0" 
                strHTML &= ("<style>")
                strHTML &= ("td{mso-number-format:""\@"";}")
                strHTML &= (".noDecFormat{mso-number-format:""0"";}")
                strHTML &= ("</style>")
                strHTML &= ("</head>")
                strHTML &= ("<body>")

                'strHTML &= ("<div>")
                strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
                strHTML &= ("<tr><td>")
                '標題抬頭
                'Dim ExportStr As String="" '建立輸出文字
                'ExportStr="<tr>"
                'ExportStr &= "<td colspan='" & iColSpanCount & "' align='center'>" & sTitle1 & "</td>" '& vbTab
                'ExportStr &= "</tr>" & vbCrLf
                'strHTML &=(ExportStr)
                Dim objStringWriter As New System.IO.StringWriter
                Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
                Div1.RenderControl(objHtmlTextWriter)
                strHTML &= (Convert.ToString(objStringWriter))
                strHTML &= ("</td></tr>")
                strHTML &= ("</table>")
        'strHTML &= ("</div>")
        'strHTML &= ("</body>")
            Case "PDF"

            Case "ODS"
                Dim objStringWriter As New System.IO.StringWriter
                Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
                Div1.RenderControl(objHtmlTextWriter)
                strHTML &= (Convert.ToString(objStringWriter))
            Case Else
        End Select

        'TIMS.writeLog(Me, "v_ExpType: " & v_ExpType)
        'TIMS.writeLog(Me, "strHTML: " & strHTML)

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1()
    End Sub

    ''' <summary>
    ''' 列印
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnPrint2_Click(sender As Object, e As EventArgs) Handles BtnPrint2.Click
        'Session("SqlString")=Me.ViewState("SqlString")
        Session("SD_15_023_SqlStr") = ViewState("SD_15_023_SqlStr") '= Sql
        Session("SD_15_023_parms") = ViewState("SD_15_023_parms") '= parms
        Dim v_XRoll As String = TIMS.GetListValue(XRoll)
        Dim v_YRoll As String = TIMS.GetListValue(YRoll)
        Dim t_YRoll As String = TIMS.GetListText(YRoll)
        Dim v_SearchPlan As String = TIMS.GetListText(SearchPlan)
        Dim s_jsc As String = "<script>wopen('SD_15_023_R.aspx?X=" & v_XRoll & "&Y=" & v_YRoll & "&YText=" & Server.UrlEncode(t_YRoll) & "&SearchPlan=" & v_SearchPlan & "');</script>"
        Page.RegisterStartupScript("open", s_jsc)
        Call Search1()
    End Sub

    ''' <summary>
    ''' 匯出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp2_Click(sender As Object, e As EventArgs) Handles BtnExp2.Click
        Call ExportDiv1()
    End Sub
End Class