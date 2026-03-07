Partial Class CP_04_002_add_01
    Inherits AuthBasePage

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
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Me.ViewState("itemstr") = Session("itemstr")
            Me.ViewState("itemplan") = Session("itemplan")
            Me.ViewState("itemcity") = Session("itemcity")
            Me.ViewState("SSTDate") = Session("SSTDate")
            Me.ViewState("ESTDate") = Session("ESTDate")
            Me.ViewState("ConRID") = Session("ConRID")
            Me.ViewState("newDistName") = Session("newDistName")
            Me.ViewState("newICityName") = Session("newICityName")
            Me.ViewState("newTPlanIDName") = Session("newTPlanIDName")
            Me.ViewState("newConRIDName") = Session("newConRIDName")
            Session("itemstr") = Nothing
            Session("itemplan") = Nothing
            Session("itemcity") = Nothing
            Session("SSTDate") = Nothing
            Session("ESTDate") = Nothing
            Session("ConRID") = Nothing
            Session("newDistName") = Nothing
            Session("newICityName") = Nothing
            Session("newTPlanIDName") = Nothing
            Session("newConRIDName") = Nothing
            create()
        End If

        '回上一頁
        Me.Button2.Attributes.Add("onclick", "location.href='CP_04_002.aspx';return false;")
    End Sub

    Sub create()
        'Dim dt As DataTable
        'Dim dr As DataRow
        Dim sqlstr, str1, str2 As String
        Dim yearlist As String = Request("yearlist")
        Dim itemstr As String = Me.ViewState("itemstr")
        Dim itemplan As String = Me.ViewState("itemplan")
        Dim itemcity As String = Me.ViewState("itemcity")
        Dim SSTDate As String = Me.ViewState("SSTDate")
        Dim ESTDate As String = Me.ViewState("ESTDate")

        str1 = "" & vbCrLf
        str2 = "" & vbCrLf
        '選擇訓練計畫
        If itemplan <> "" Then
            str2 = str2 & " and kp.TPlanID IN (" & itemplan & ") " & vbCrLf
        End If

        '選擇轄區
        If itemstr <> "" Then
            str1 = str1 & " and ip.DistID IN (" & itemstr & ") " & vbCrLf
        End If

        '選擇年度
        If yearlist <> "" Then
            str1 = str1 & " and ip.Years='" & Trim(yearlist) & "' " & vbCrLf
        End If

        '選擇縣市
        If itemcity <> "" Then

        End If

        '選擇縣市
        If itemcity <> "" Then
            str1 += " and (1!=1" & vbCrLf
            str1 += "  OR iz.CTID IN (" & itemcity & ") " & vbCrLf
            str1 += "  OR iz1.CTID IN (" & itemcity & ") " & vbCrLf
            str1 += "  OR iz2.CTID IN (" & itemcity & ") " & vbCrLf
            str1 += " )" & vbCrLf
        End If

        '開訓日期起
        If SSTDate <> "" Then
            str1 += " and cc.STDate >= " & TIMS.To_date(SSTDate) & vbCrLf
        End If

        '開訓日期迄
        If ESTDate <> "" Then
            str1 += " and cc.STDate <= " & TIMS.To_date(ESTDate) & vbCrLf
        End If

        '管控單位
        If Me.ViewState("ConRID") <> "" Then
            Dim Relship As String
            Dim RelshipStr As String '多筆含逗號
            RelshipStr = ""
            For i As Integer = 0 To Split(Me.ViewState("ConRID"), ",").Length - 1
                Relship = DbAccess.ExecuteScalar("SELECT Relship FROM Auth_Relship WHERE RID ='" & Split(Me.ViewState("ConRID"), ",")(i) & "'", objconn)

                If RelshipStr <> "" Then RelshipStr &= ","
                RelshipStr &= Relship
            Next

            If Split(RelshipStr, ",").Length > 1 Then
                '多筆 Split(RelshipStr, ",")(i) 
                str1 += " and (1!=1"
                For i As Integer = 0 To Split(RelshipStr, ",").Length - 1
                    str1 += " or cc.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Split(RelshipStr, ",")(i) & "%')" & vbCrLf
                Next
                str1 += " )"
            Else
                '單1筆 (RelshipStr)
                str1 += " and cc.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & RelshipStr & "%')"
            End If
        End If


        'If Me.ViewState("ConRID") <> "" Then
        '    Dim Relship As String
        '    Dim RelshipStr As String
        '    For i As Integer = 0 To Split(Me.ViewState("ConRID"), ",").Length - 1
        '        Relship = DbAccess.ExecuteScalar("SELECT Relship FROM Auth_Relship WHERE RID ='" & Split(Me.ViewState("ConRID"), ",")(i) & "'", objconn)

        '        If RelshipStr = "" Then
        '            If Split(Me.ViewState("ConRID"), ",").Length = 1 Then
        '                RelshipStr = " and cc.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Relship & "%')" & vbCrLf
        '            Else
        '                RelshipStr = " and (cc.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Relship & "%')" & vbCrLf
        '            End If
        '        Else
        '            If i <> Split(Me.ViewState("ConRID"), ",").Length - 1 Then
        '                RelshipStr += " or cc.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Relship & "%')" & vbCrLf
        '            Else
        '                RelshipStr += " or cc.RID IN (SELECT RID FROM Auth_Relship WHERE Relship like '" & Relship & "%'))" & vbCrLf
        '            End If
        '        End If
        '    Next
        '    str1 += RelshipStr & vbCrLf
        'End If

        sqlstr = "SELECT kp.TPlanID,kp.PlanName " & vbCrLf

        sqlstr += " ,(select count(1) " & vbCrLf
        sqlstr += " from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf
        sqlstr += " where cc.NotOpen = 'N' and cc.IsSuccess = 'Y' " & vbCrLf
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and cc.stdate <= getdate() " & vbCrLf '已開班
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += "" & str1 & ") AS Class1 " & vbCrLf

        sqlstr += " ,(select count(1) " & vbCrLf
        sqlstr += " from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf

        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf

        sqlstr += " where cc.NotOpen = 'N' and cc.IsSuccess = 'Y' " & vbCrLf
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and cc.stdate > getdate() " & vbCrLf '未開班
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += "" & str1 & ") AS Class2 " & vbCrLf

        sqlstr += " ,(select count(1) " & vbCrLf
        sqlstr += " from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf
        sqlstr += " where cc.NotOpen = 'Y'  and cc.IsSuccess = 'Y'" & vbCrLf '不開班
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += "" & str1 & ") AS Class3 " & vbCrLf

        sqlstr += " ,(select count(1) " & vbCrLf
        sqlstr += " from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf

        sqlstr += " where cc.IsSuccess = 'Y'" & vbCrLf '已轉入 開班數
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += "" & str1 & ") AS Class4 " & vbCrLf

        '計畫總人數
        sqlstr += " ,(select count(1) from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf

        'sqlstr += " where cc.NotOpen = 'N' and cc.IsSuccess = 'Y'" & vbCrLf '已轉入 開班數
        sqlstr += " where cc.IsSuccess = 'Y'" & vbCrLf '已轉入 開班數
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += "" & str1 & ") AS student1 " & vbCrLf

        '在訓總人數
        sqlstr += " ,(select count(1) from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf

        sqlstr += " where cc.NotOpen = 'N' and cc.IsSuccess = 'Y'" & vbCrLf '已轉入 開班數
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += " and cs.StudStatus = '1' " & vbCrLf
        sqlstr += "" & str1 & ") AS student2 " & vbCrLf

        '結訓總人數
        sqlstr += " ,(select count(1) from Class_ClassInfo cc " & vbCrLf
        sqlstr += " JOIN plan_planinfo pp on pp.planid = cc.planid and pp.comidno = cc.comidno and pp.seqno = cc.seqno " & vbCrLf
        sqlstr += " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr += " JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID " & vbCrLf
        sqlstr += " JOIN Auth_Relship ar ON ar.RID = cc.RID " & vbCrLf
        sqlstr += " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip " & vbCrLf
        ' /* 產投上課地址學科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace sp   on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz1   on iz1.zipCode=sp.ZipCode" & vbCrLf
        ' /* 產投上課地址術科場地代碼 */
        sqlstr += " LEFT JOIN Plan_TrainPlace tp   on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sqlstr += " LEFT JOIN ID_ZIP iz2   on iz2.zipCode=tp.ZipCode" & vbCrLf

        sqlstr += " where cc.NotOpen = 'N' and cc.IsSuccess = 'Y'" & vbCrLf '已轉入 開班數
        sqlstr += " AND pp.isapprpaper = 'Y' AND pp.appliedresult = 'Y' AND pp.transflag = 'Y' " & vbCrLf
        sqlstr += " and ip.TPlanID = kp.TPlanID " & vbCrLf
        sqlstr += " and cs.StudStatus = '5' " & vbCrLf
        sqlstr += "" & str1 & ") AS student3 " & vbCrLf

        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where cc.NotOpen = 'N' and cc.IsSuccess = 'Y' and ip.TPlanID = kp.TPlanID " & str1 & ") AS Class1," & vbCrLf
        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where cc.NotOpen = 'N' and cc.IsSuccess = 'N' and ip.TPlanID = kp.TPlanID " & str1 & ") AS Class2," & vbCrLf
        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where cc.NotOpen = 'Y'  and ip.TPlanID = kp.TPlanID " & str1 & ") AS Class3," & vbCrLf
        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where ip.TPlanID = kp.TPlanID " & str1 & ") AS Class4," & vbCrLf
        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where ip.TPlanID = kp.TPlanID " & str1 & ") AS student1," & vbCrLf
        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where ip.TPlanID = kp.TPlanID and cs.StudStatus = '1' " & str1 & ") AS student2," & vbCrLf
        'sqlstr += "(select count(1) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = cc.TaddressZip where ip.TPlanID = kp.TPlanID and cs.StudStatus = '5' " & str1 & ") AS student3" & vbCrLf
        sqlstr += " FROM  Key_Plan kp where 1 = 1 " & vbCrLf

        sqlstr += str2 & vbCrLf
        '以轄區、訓練計畫做排序

        Me.ViewState("cp_sql") = sqlstr

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Me.NoData.Text = "<font color=red>查無資料</font>"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.NoData.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True
            'PageControler1.SqlPrimaryKeyDataCreate(sqlstr, "TPlanID", "TPlanID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "TPlanID"
            PageControler1.Sort = "TPlanID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            '序號
            e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cp_sql As String
        Dim table As DataTable
        Dim dr As DataRow

        cp_sql = Me.ViewState("cp_sql")
        table = DbAccess.GetDataTable(cp_sql, objconn)

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("CP_Data", System.Text.Encoding.UTF8) & ".xls")
        Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim ExportStr As String             '建立輸出文字
        ExportStr = "計畫名稱" & vbTab & "已開班" & vbTab & "未開班" & vbTab & "不開班" & vbTab & "總開班" & vbTab & "計畫總人數" & vbTab & "在訓總人數" & vbTab & "結訓總人數" & vbTab
        ExportStr += vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        '建立資料面
        For Each dr In table.Rows
            ExportStr = ""
            ExportStr = ExportStr & dr("PlanName") & vbTab
            ExportStr = ExportStr & dr("Class1") & vbTab
            ExportStr = ExportStr & dr("Class2") & vbTab
            ExportStr = ExportStr & dr("Class3") & vbTab
            ExportStr = ExportStr & dr("Class4") & vbTab
            ExportStr = ExportStr & dr("student1") & vbTab
            ExportStr = ExportStr & dr("student2") & vbTab
            ExportStr = ExportStr & dr("student3") & vbTab
            ExportStr += vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        Response.End()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '報表列印
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        Dim newyear As String = Request("yearlist")
        Dim SnewDate As String = Me.ViewState("SSTDate")
        Dim EnewDate As String = Me.ViewState("ESTDate")
        Dim newstr As String = Session("newDistID")
        Dim newplan As String = Session("newTPlanID")
        Dim newcity As String = Session("newICity")
        Dim newRID As String = Session("newConRID")
        Dim newDistName As String = Me.ViewState("newDistName")
        Dim newICityName As String = Me.ViewState("newICityName")
        Dim newTPlanIDName As String = Me.ViewState("newTPlanIDName")
        Dim newConRIDName As String = Me.ViewState("newConRIDName")

        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=Report&path=TIMS&filename=CP_04_002_01_Rpt"
        'strScript += "&Years=" & newyear & "&DistID=" & newstr & "&TPlanID=" & newplan & "&SSTDate=" & SnewDate & "&ESTDate=" & EnewDate & "&itemcity=" & newcity & ""
        'strScript += "&ConRID=" & newRID & ""
        'strScript += "&newDistName='+escape('" & newDistName & "')+'&newICityName='+escape('" & newICityName & "')+'&newTPlanIDName='+escape('" & newTPlanIDName & "')+'&newConRIDName='+escape('" & newConRIDName & "'));" + vbCrLf
        'strScript += "</script>"

        'Page.RegisterStartupScript("window_onload", strScript)

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "Years=" & newyear
        MyValue += "&DistID=" & newstr
        MyValue += "&TPlanID=" & newplan
        MyValue += "&SSTDate=" & SnewDate
        MyValue += "&ESTDate=" & EnewDate
        MyValue += "&itemcity=" & newcity
        MyValue += "&ConRID=" & newRID
        MyValue += "&newDistName=" & newDistName
        MyValue += "&newICityName=" & newICityName
        MyValue += "&newTPlanIDName=" & newTPlanIDName
        MyValue += "&newConRIDName=" & newConRIDName

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "CP_04_002_01_Rpt", MyValue)
    End Sub
End Class
