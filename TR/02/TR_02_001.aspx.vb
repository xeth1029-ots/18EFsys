Partial Class TR_02_001
    Inherits AuthBasePage

    Dim Key_AgeRange As DataTable
    Dim Key_WorkYear As DataTable
    Dim Key_ProSkill As DataTable
    Dim Key_Degree As DataTable
    Dim Key_Military As DataTable

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Key_AgeRange = TIMS.Get_KeyTable("Key_AgeRange", objconn)
        Key_WorkYear = TIMS.Get_KeyTable("Key_WorkYear", objconn)
        Key_ProSkill = TIMS.Get_KeyTable("Key_ProSkill", objconn)
        Key_Degree = TIMS.Get_KeyTable("Key_Degree", objconn)
        Key_Military = TIMS.Get_KeyTable("Key_Military", objconn)

        If Not IsPostBack Then
            msg1.Text = ""
            Session("Bus_VisitTR") = Nothing
            CreateItem()

            SearchTable.Visible = True
            ResultTable.Visible = False
            DetailTable.Visible = False
            ListTable.Visible = False

            If sm.UserInfo.RID = "A" Then
                DistTr.Visible = True
            Else
                DistTr.Visible = False
            End If
            Common.SetListItem(SDistID, sm.UserInfo.DistID)

            Me.ViewState("BVID") = ""
            Me.ViewState("BDID") = ""
            Me.ViewState("ProType") = ""

            SVisitDate2.Text = Now.Date
        End If

        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        City.Attributes("onblur") = "getzipname(this.value,'City','Zip');"
        Button4.Attributes("onclick") = "return CheckData();"
        Button12.Attributes("onclick") = "wopen('../../Common/ProSkill.aspx?Skill_Field=KPID&Skill_Name_Field=ProName','skill',500,500,1);"
        Button9.Attributes("onclick") = "wopen('TR_02_001_Finder.aspx','Finder',550,500,1);"
        Button10.Attributes("onclick") = "return CheckAdd();"

        PageControler1.PageDataGrid = DataGrid1
    End Sub

    Sub CreateItem()
        SCTID = TIMS.Get_CityName(SCTID, TIMS.dtNothing)
        SCTID.Items.Insert(0, New ListItem("全選", ""))
        SCTID.Attributes("onclick") = "SelectAll();"

        SDistID = TIMS.Get_DistID(SDistID)
        TradeID = TIMS.Get_KeyControl(TradeID, "Key_Trade", "TRADENAME", "TRADEID", objconn)
        KEID = TIMS.Get_KeyControl(KEID, "Key_Emp", "KENAME", "KEID", objconn)
        DegreeID = TIMS.Get_Degree(DegreeID, 1, objconn)
        ARID = TIMS.Get_KeyControl(ARID, "Key_AgeRange", "ARNAME", "ARID", objconn)
        WYID = TIMS.Get_KeyControl(WYID, "Key_WorkYear", "WYNAME", "WYID", objconn)
        MilitaryID = TIMS.Get_Military(MilitaryID, 1, objconn)
        BVCID = TIMS.Get_KeyControl(BVCID, "Key_BusVisitCase", "BVCNAME", "BVCID", objconn)
    End Sub

    '查詢按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim sql As String
        'Dim NewSql As String
        'Dim dt As DataTable
        'Dim CTRound As String
        'Dim DateRound As String
        Dim DistStr As String = ""
        Dim CTRound As String = ""
        Dim DateRound As String = ""

        ResultTable.Visible = True

        For Each item As ListItem In SCTID.Items
            If item.Selected = True AndAlso item.Value <> "" Then
                If CTRound <> "" Then CTRound += ","
                CTRound += item.Value
            End If
        Next
        If CTRound <> "" Then
            CTRound = " and Zip IN (SELECT ZipCode FROM ID_ZIP WHERE CTID IN (" & CTRound & "))"
        End If
        If SVisitDate1.Text <> "" Then
            DateRound += " and VisitDate>= " & TIMS.To_date(SVisitDate1.Text) & vbCrLf '"'"
        End If
        If SVisitDate2.Text <> "" Then
            DateRound += " and VisitDate<= " & TIMS.To_date(SVisitDate2.Text) & vbCrLf '" & SVisitDate2.Text & "'"
        End If
        DistStr = " and DistID='" & SDistID.SelectedValue & "'"


        SDistName.Text = SDistID.SelectedItem.Text
        Dim sql As String = ""
        sql = ""
        sql += "SELECT a.BDID,c.Name as DistName,a.Uname,b.Num,a.Intaxno,d.TradeName,e.KEName FROM "
        sql += "(SELECT * FROM Bus_BasicData WHERE Uname like '%" & SUname.Text & "%' and Intaxno like '%" & SIntaxno.Text & "%'" & CTRound & ") a "
        sql += "JOIN (SELECT BDID,DistID,count(*) as Num FROM Bus_VisitInfo WHERE VisitKind='" & SVisitKind.SelectedValue & "'" & DateRound & DistStr & " Group By BDID,DistID) b ON a.BDID=b.BDID "
        sql += "LEFT JOIN ID_District c ON b.DistID=c.DistID "
        sql += "LEFT JOIN Key_Trade d ON a.TradeID=d.TradeID "
        sql += "LEFT JOIN Key_Emp e ON a.KEID=e.KEID "

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        msg1.Text = "查無資料"
        DataGridTable1.Visible = False
        If dt.Rows.Count > 0 Then
            msg1.Text = ""
            DataGridTable1.Visible = True

            'PageControler1.SqlPrimaryKeyDataCreate(sql, "BDID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "BDID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim btn1 As LinkButton = e.Item.FindControl("LinkButton1")
                Dim btn2 As Button = e.Item.FindControl("Button2")
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                btn1.Text = drv("Uname") & "(" & drv("Num") & ")"
                btn1.CommandArgument = drv("BDID")
                btn2.CommandArgument = drv("BDID")

        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "view"
                SearchTable.Visible = False
                ListTable.Visible = True
                Me.ViewState("BDID") = e.CommandArgument

                LDistName.Text = SDistID.SelectedItem.Text
                ShowVisitList(e.CommandArgument)
            Case "add"
                SearchTable.Visible = False
                DetailTable.Visible = True

                ProType.Text = "-新增"
                Me.ViewState("ProType") = "add"
                CleanData()

                CreateVisiterData(e.CommandArgument)

                Dim dr As DataRow
                dr = DbAccess.GetOneRow("SELECT * FROM Bus_BasicData WHERE BDID='" & e.CommandArgument & "'", objconn)
                Me.ViewState("BDID") = e.CommandArgument
                Uname.Text = dr("Uname").ToString
                BDID.Value = dr("BDID").ToString
                Intaxno.Text = dr("Intaxno").ToString
                Button9.Visible = False
                If dr("Zip").ToString <> "" Then
                    City.Text = "(" & dr("Zip").ToString & ")" & TIMS.Get_ZipName(dr("Zip").ToString)
                    Zip.Value = dr("Zip").ToString
                End If
                Addr.Text = dr("Addr").ToString
                Common.SetListItem(TradeID, dr("TradeID").ToString)
                Common.SetListItem(KEID, dr("KEID").ToString)
                Common.SetListItem(Labor, If(dr("Labor"), "1", "0"))
        End Select
    End Sub

    Sub CleanData()
        OrgName.Text = ""
        RID.Value = ""
        Visiter.Text = ""
        VisitDate.Text = ""
        VisitNum.Text = ""
        VisitKind.SelectedIndex = -1
        Uname.Text = ""
        BDID.Value = ""
        Intaxno.Text = ""
        City.Text = ""
        Zip.Value = ""
        Addr.Text = ""
        TradeID.SelectedIndex = -1
        KEID.SelectedIndex = -1
        Labor.SelectedIndex = -1
        VisitedName.Text = ""
        VisitedTitle.Text = ""
        VisitedTel.Text = ""
        VistiedFax.Text = ""
        VisitedMob.Text = ""
        VisitedMail.Text = ""
        BVCID.SelectedIndex = -1
        VisitOther.Text = ""
        BPKind1.Checked = False
        BPKind2.Checked = False
        BPKind3.Checked = False
        BPKind_Year.Text = ""
        BPKind_Mon.Text = ""
        BPKind_Day.Text = ""
        ProName.Text = ""
        KPID.Value = ""
        DegreeID.SelectedIndex = -1
        ARID.SelectedIndex = -1
        WYID.SelectedIndex = -1
        MilitaryID.SelectedIndex = -1
        License.Text = ""
        RPNum.Text = ""
        ProYear.Text = ""
        ProMonth.Text = ""

        DataGrid2.Visible = False
    End Sub

    Sub CreateVisiterData(ByVal BDID As String)
        Dim dr As DataRow
        OrgName.Text = sm.UserInfo.OrgName
        RID.Value = sm.UserInfo.RID
        Visiter.Text = DbAccess.ExecuteScalar("SELECT Name FROM Auth_Account WHERE Account='" & sm.UserInfo.UserID & "'", objconn)
        VisitDate.Text = Now.Date
        dr = DbAccess.GetOneRow("SELECT MAX(VisitNum)+1 as Num FROM Bus_VisitInfo WHERE BDID='" & BDID & "' and RID='" & sm.UserInfo.RID & "'", objconn)
        If IsDBNull(dr("Num")) Then
            VisitNum.Text = 1
        Else
            VisitNum.Text = dr("Num")
        End If
    End Sub

    Sub ShowVisitList(ByVal BDID As String)
        Dim sql As String = ""
        Dim dt As DataTable = Nothing

        sql = "SELECT b.BVID,b.RID,a.Uname,b.VisitDate,c.Name,b.VisitKind,d.BVCID,d.BVCName,b.VisitOther,b.BPKind,b.BPKind_Year,b.BPKind_Mon,b.BPKind_Day,e.TradeName,f.KEName FROM "
        sql += "(SELECT * FROM Bus_BasicData WHERE BDID='" & BDID & "') a "
        sql += "JOIN (SELECT * FROM Bus_VisitInfo WHERE BDID='" & BDID & "' and DistID='" & SDistID.SelectedValue & "') b ON a.BDID=b.BDID "
        sql += "LEFT JOIN Auth_Account c ON b.Visiter=c.Account "
        sql += "LEFT JOIN Key_BusVisitCase d ON b.BVCID=d.BVCID "
        sql += "LEFT JOIN Key_Trade e ON a.TradeID=e.TradeID "
        sql += "LEFT JOIN Key_Emp f ON a.KEID=f.KEID "

        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            SearchTable.Visible = True
            ListTable.Visible = False
            Button1_Click(Button1, Nothing)
        Else
            LUname.Text = dt.Rows(0)("Uname").ToString
            LTradeID.Text = dt.Rows(0)("TradeName").ToString
            LKEID.Text = dt.Rows(0)("KEName").ToString

            DataGrid3.DataSource = dt
            DataGrid3.DataBind()
        End If
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button7")
                Dim btn2 As Button = e.Item.FindControl("Button8")
                Dim btn3 As Button = e.Item.FindControl("Button13")

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1

                Select Case drv("VisitKind").ToString
                    Case "1"
                        e.Item.Cells(3).Text = "一般計畫"
                    Case "2"
                        e.Item.Cells(3).Text = "訓用合一"
                    Case Else
                End Select

                If drv("BVCID").ToString = "99" Then
                    e.Item.Cells(4).Text = "(" & drv("VisitOther").ToString & ")"
                End If

                Select Case drv("BPKind").ToString
                    Case "1"
                        e.Item.Cells(5).Text = "存檔備查"
                    Case "2"
                        e.Item.Cells(5).Text = "持續聯繫，並e-mail或傳真相關資料提供參考。"
                    Case "3"
                        e.Item.Cells(5).Text = "再次前往訪視(預計訪視日期西元" & drv("BPKind_Year").ToString & "年" & drv("BPKind_Mon").ToString & "月" & drv("BPKind_Day").ToString & "日)"
                End Select

                If sm.UserInfo.RID <> drv("RID") Then
                    btn1.Visible = False
                    btn2.Visible = False
                    btn3.Visible = True
                Else
                    btn1.Visible = True
                    btn2.Visible = True
                    btn3.Visible = False
                End If

                btn1.CommandArgument = drv("BVID").ToString
                btn2.CommandArgument = drv("BVID").ToString
                btn3.CommandArgument = drv("BVID").ToString

                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End Select
    End Sub

    Private Sub DataGrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid3.ItemCommand
        Select Case e.CommandName
            Case "edit", "view"
                ListTable.Visible = False
                DetailTable.Visible = True

                ProType.Text = "-修改"
                CleanData()

                Me.ViewState("BVID") = e.CommandArgument
                Me.ViewState("ProType") = e.CommandName
                create(e.CommandArgument)

                '檢視狀態隱藏儲存按鈕
                If Me.ViewState("ProType") = "view" Then
                    Button4.Visible = False
                    Button10.Visible = False
                Else
                    Button4.Visible = True
                    Button10.Visible = True
                End If
            Case "del"
                Dim sql As String
                sql = "DELETE Bus_VisitInfo WHERE BVID='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                sql = "DELETE Bus_VisitTR WHERE BVID='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功")
                ShowVisitList(Me.ViewState("BDID"))
        End Select
    End Sub

    Sub create(ByVal BVID As String)
        Dim sql As String
        Dim dr As DataRow
        'Dim dt As DataTable
        sql = ""
        sql &= " SELECT a.*,b.*,d.OrgName "
        sql += " FROM (SELECT * FROM Bus_VisitInfo WHERE BVID='" & BVID & "') a "
        sql += " JOIN Bus_BasicData b ON a.BDID=b.BDID "
        sql += " JOIN Auth_Relship c ON a.RID=c.RID "
        sql += " JOIN Org_OrgInfo d ON c.OrgID=d.OrgID "
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr Is Nothing Then
            Common.MessageBox(Me, "查無資料")
        Else
            OrgName.Text = dr("OrgName").ToString
            RID.Value = dr("RID").ToString
            Visiter.Text = dr("Visiter").ToString
            VisitDate.Text = FormatDateTime(dr("VisitDate"), 2)
            VisitNum.Text = dr("VisitNum").ToString
            Common.SetListItem(VisitKind, dr("VisitKind").ToString)
            Uname.Text = dr("Uname").ToString
            BDID.Value = dr("BDID").ToString
            Intaxno.Text = dr("Intaxno").ToString
            Button9.Visible = False
            If dr("Zip").ToString <> "" Then
                City.Text = "(" & dr("Zip").ToString & ")" & TIMS.Get_ZipName(dr("Zip").ToString)
                Zip.Value = dr("Zip").ToString
            End If
            Addr.Text = dr("Addr").ToString
            Common.SetListItem(TradeID, dr("TradeID").ToString)
            Common.SetListItem(KEID, dr("KEID").ToString)
            Common.SetListItem(Labor, If(dr("Labor"), "1", "0"))

            VisitedName.Text = dr("VisitedName").ToString
            VisitedTitle.Text = dr("VisitedTitle").ToString
            VisitedTel.Text = dr("VisitedTel").ToString
            VistiedFax.Text = dr("VistiedFax").ToString
            VisitedMob.Text = dr("VisitedMob").ToString
            VisitedMail.Text = dr("VisitedMail").ToString
            Common.SetListItem(BVCID, dr("BVCID").ToString)
            VisitOther.Text = dr("VisitOther").ToString
            Select Case dr("BPKind").ToString
                Case "1"
                    BPKind1.Checked = True
                Case "2"
                    BPKind2.Checked = True
                Case "3"
                    BPKind3.Checked = True
            End Select
            BPKind_Year.Text = dr("BPKind_Year").ToString
            BPKind_Mon.Text = dr("BPKind_Mon").ToString
            BPKind_Day.Text = dr("BPKind_Day").ToString
        End If

        CreateBus_VisitTR(BVID)

    End Sub

    Sub CreateBus_VisitTR(ByVal BVID As String)
        Dim sql As String
        Dim dt As DataTable
        sql = "SELECT * FROM Bus_VisitTR WHERE BVID='" & BVID & "'"
        dt = DbAccess.GetDataTable(sql, objconn)
        dt.Columns("BVTID").AutoIncrement = True
        dt.Columns("BVTID").AutoIncrementSeed = -1
        dt.Columns("BVTID").AutoIncrementStep = -1
        Session("Bus_VisitTR") = dt

        If dt.Rows.Count = 0 Then
            DataGrid2.Visible = False
        Else
            DataGrid2.Visible = True

            DataGrid2.DataSource = dt
            DataGrid2.DataBind()
        End If
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As Button = e.Item.FindControl("Button5")

                e.Item.Cells(0).Text = "【" & Key_ProSkill.Select("KPID='" & drv("KPID") & "'")(0)("ProID") & "】" & Key_ProSkill.Select("KPID='" & drv("KPID") & "'")(0)("ProName")
                e.Item.Cells(1).Text = Key_Degree.Select("DegreeID='" & drv("DegreeID") & "'")(0)("Name")
                e.Item.Cells(2).Text = Key_AgeRange.Select("ARID='" & drv("ARID") & "'")(0)("ARName")
                e.Item.Cells(3).Text = Key_WorkYear.Select("WYID='" & drv("WYID") & "'")(0)("WYName")
                e.Item.Cells(4).Text = Key_Military.Select("MilitaryID='" & drv("MilitaryID") & "'")(0)("Name")

                If IsNumeric(drv("ProYear")) Then
                    If Int(drv("ProYear")) <> 0 Then
                        e.Item.Cells(7).Text = drv("ProYear") & "年"
                    End If
                End If
                If IsNumeric(drv("ProMonth")) Then
                    If Int(drv("ProMonth")) <> 0 Then
                        e.Item.Cells(7).Text += drv("ProMonth") & "月"
                    End If
                End If
                btn.CommandArgument = drv("BVTID").ToString
                btn.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Select Case e.CommandName
            Case "del"
                Dim dt As DataTable = Session("Bus_VisitTR")
                Dim dr As DataRow
                If dt.Select("BVTID='" & e.CommandArgument & "'").Length <> 0 Then
                    dr = dt.Select("BVTID='" & e.CommandArgument & "'")(0)
                    dr.Delete()
                End If
                Session("Bus_VisitTR") = dt

                If dt.Rows.Count = 0 Then
                    DataGrid2.Visible = False
                Else
                    DataGrid2.Visible = True
                    DataGrid2.DataSource = dt
                    DataGrid2.DataBind()
                End If
        End Select
    End Sub

    '新增訪視表
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        CleanData()
        ProType.Text = "-新增"
        Me.ViewState("ProType") = "add"
        CreateVisiterData("")
        Button9.Visible = True
        SearchTable.Visible = False
        DetailTable.Visible = True
    End Sub

    '新增職類需求
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim dt As DataTable
        Dim dr As DataRow
        Dim sql As String

        If Session("Bus_VisitTR") Is Nothing Then
            sql = "SELECT * FROM Bus_VisitTR WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, objconn)

            dt.Columns("BVTID").AutoIncrement = True
            dt.Columns("BVTID").AutoIncrementSeed = -1
            dt.Columns("BVTID").AutoIncrementStep = -1
        Else
            dt = Session("Bus_VisitTR")
        End If

        dr = dt.NewRow
        dt.Rows.Add(dr)
        If Me.ViewState("BVID") <> "" Then
            dr("BVID") = Me.ViewState("BVID")
        End If
        dr("KPID") = KPID.Value
        dr("DegreeID") = DegreeID.SelectedValue
        dr("ARID") = ARID.SelectedValue
        dr("WYID") = WYID.SelectedValue
        dr("MilitaryID") = MilitaryID.SelectedValue
        dr("License") = License.Text
        dr("RPNum") = RPNum.Text
        dr("ProYear") = If(ProYear.Text = "", Convert.DBNull, ProYear.Text)
        dr("ProMonth") = If(ProMonth.Text = "", Convert.DBNull, ProMonth.Text)

        Session("Bus_VisitTR") = dt

        DataGrid2.Visible = True
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()

        ProName.Text = ""
        KPID.Value = ""
        DegreeID.SelectedIndex = -1
        ARID.SelectedIndex = -1
        WYID.SelectedIndex = -1
        MilitaryID.SelectedIndex = -1
        License.Text = ""
        RPNum.Text = ""
        ProYear.Text = ""
        ProMonth.Text = ""
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        SearchTable.Visible = True
        ListTable.Visible = False
        Me.ViewState("BDID") = ""
    End Sub

    '儲存訪視表
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim conn As SqlConnection = DbAccess.GetConnection
        Dim trans As SqlTransaction = Nothing
        Dim dr As DataRow = Nothing

        Try
            trans = DbAccess.BeginTrans(conn)
            sql = "SELECT * FROM Bus_VisitInfo WHERE BVID='" & Me.ViewState("BVID") & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
                dt.Rows.Add(dr)
            Else
                dr = dt.Rows(0)
            End If

            dr("RID") = RID.Value
            If Me.ViewState("BVID") = "" Then
                dr("DistID") = sm.UserInfo.DistID
            End If
            dr("VisitDate") = VisitDate.Text
            dr("VisitKind") = VisitKind.SelectedValue
            dr("BDID") = BDID.Value
            If Me.ViewState("BVID") = "" Then
                Dim drTemp As DataRow
                drTemp = DbAccess.GetOneRow("SELECT MAX(VisitNum)+1 as Num FROM Bus_VisitInfo WHERE BDID='" & BDID.Value & "' and RID='" & sm.UserInfo.RID & "'", trans)
                If IsDBNull(drTemp("Num")) Then
                    dr("VisitNum") = 1
                Else
                    dr("VisitNum") = drTemp("Num")
                End If
            End If
            dr("Visiter") = Visiter.Text
            dr("VisitedName") = VisitedName.Text
            dr("VisitedTel") = VisitedTel.Text
            dr("VisitedMob") = If(VisitedMob.Text = "", Convert.DBNull, VisitedMob.Text)
            dr("VisitedTitle") = If(VisitedTitle.Text = "", Convert.DBNull, VisitedTitle.Text)
            dr("VistiedFax") = If(VistiedFax.Text = "", Convert.DBNull, VistiedFax.Text)
            dr("VisitedMail") = If(VisitedMail.Text = "", Convert.DBNull, VisitedMail.Text)
            dr("BVCID") = BVCID.SelectedValue
            If BVCID.SelectedValue = "99" Then
                dr("VisitOther") = If(VisitOther.Text = "", Convert.DBNull, VisitOther.Text)
            Else
                dr("VisitOther") = Convert.DBNull
            End If

            If BPKind1.Checked Then
                dr("BPKind") = 1
            ElseIf BPKind2.Checked Then
                dr("BPKind") = 2
            ElseIf BPKind3.Checked Then
                dr("BPKind") = 3
            End If
            If BPKind3.Checked Then
                dr("BPKind_Year") = If(BPKind_Year.Text = "", Convert.DBNull, BPKind_Year.Text)
                dr("BPKind_Mon") = If(BPKind_Mon.Text = "", Convert.DBNull, BPKind_Mon.Text)
                dr("BPKind_Day") = If(BPKind_Day.Text = "", Convert.DBNull, BPKind_Day.Text)
            Else
                dr("BPKind_Year") = Convert.DBNull
                dr("BPKind_Mon") = Convert.DBNull
                dr("BPKind_Day") = Convert.DBNull
            End If

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now

            DbAccess.UpdateDataTable(dt, da, trans)
            If Me.ViewState("BVID") = "" Then
                Me.ViewState("BVID") = DbAccess.GetNewId(trans.Connection, "BUS_VISITINFO_BVID_SEQ")
            End If

            sql = "SELECT * FROM Bus_BasicData WHERE BDID='" & BDID.Value & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                If TradeID.SelectedIndex <> 0 Then
                    dr("TradeID") = TradeID.SelectedValue
                    DbAccess.UpdateDataTable(dt, da, trans)
                End If
            End If

            If Not Session("Bus_VisitTR") Is Nothing Then
                sql = "SELECT * FROM Bus_VisitTR WHERE BVID='" & Me.ViewState("BVID") & "'"
                dt = DbAccess.GetDataTable(sql, da, trans)
                Dim dtTemp As DataTable

                dtTemp = Session("Bus_VisitTR")
                For Each dr In dtTemp.Rows
                    If dr.RowState <> DataRowState.Deleted Then
                        dr("BVID") = Me.ViewState("BVID")
                    End If
                Next
                dt = dtTemp.Copy

                DbAccess.UpdateDataTable(dt, da, trans)

                Session("Bus_VisitTR") = Nothing
            End If

            DbAccess.CommitTrans(trans)
            Call TIMS.CloseDbConn(conn)
            Common.MessageBox(Me, "儲存成功")

            Me.ViewState("BVID") = ""
            DetailTable.Visible = False
            Select Case Me.ViewState("ProType")
                Case "add"
                    Button1_Click(Button1, Nothing)
                    SearchTable.Visible = True
                Case "edit", "view"
                    ShowVisitList(Me.ViewState("BDID"))
                    ListTable.Visible = True
            End Select
            ProType.Text = ""
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(conn)
            Throw ex
        End Try

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Me.ViewState("BVID") = ""
        Session("Bus_VisitTR") = Nothing
        DetailTable.Visible = False
        ProType.Text = ""
        Select Case Me.ViewState("ProType")
            Case "add"
                SearchTable.Visible = True
            Case "edit", "view"
                ListTable.Visible = True
        End Select
    End Sub
End Class
