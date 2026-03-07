Partial Class CP_04_001_02
    Inherits AuthBasePage

    Const cst_printFN1 As String = "CP_04_001_02_Rpt"

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
        PageControler1.PageDataGrid = DataGrid1
        '檢查Session是否存在 End

        If Not IsPostBack Then
            If Not Session("itemstr") Is Nothing Then Me.ViewState("itemstr") = Session("itemstr")
            If Not Session("itemplan") Is Nothing Then Me.ViewState("itemplan") = Session("itemplan")
            If Not Session("itemcity") Is Nothing Then Me.ViewState("itemcity") = Session("itemcity")
            Session("itemstr") = Nothing
            Session("itemplan") = Nothing
            Session("itemcity") = Nothing
            create()
        End If

        '回上一頁
        Me.Button2.Attributes.Add("onclick", "location.href='CP_04_001.aspx';return false;")
    End Sub

    Sub create()

        'Dim dr As DataRow
        'Dim sqlstr, str1 As String
        Dim yearlist As String = Request("yearlist")
        Dim itemstr As String = Me.ViewState("itemstr")
        Dim itemplan As String = Me.ViewState("itemplan")
        Dim itemcity As String = Me.ViewState("itemcity")
        'Dim SelectPlanTimes As Integer = Request("SelectPlanTimes")
        'Dim i As Integer
        'Dim str1 As String = ""

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select cc.OCID,oo.OrgID, oo.OrgName,cc.NotOpen,cc.IsSuccess" & vbCrLf
        sql &= " FROM Class_ClassInfo cc" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID = cc.PlanID" & vbCrLf
        sql &= " JOIN Auth_Relship ar ON ar.RID = cc.RID" & vbCrLf
        sql &= " JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID" & vbCrLf
        sql &= " JOIN Org_OrgInfo oo ON oo.OrgID=ar.OrgID" & vbCrLf
        sql &= " JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql &= " and rownum <=1000" & vbCrLf
        '選擇訓練計畫
        If itemplan <> "" Then
            sql &= " and ip.TPlanID IN (" & itemplan & ") "
        End If
        '選擇轄區
        If itemstr <> "" Then
            sql &= " and ip.DistID IN (" & itemstr & ") "
        End If
        '選擇年度
        If yearlist <> "" Then
            sql &= " and ip.Years='" & Trim(yearlist) & "' "
        End If
        '選擇縣市
        If itemcity <> "" Then
            sql &= " and iz.CTID IN (" & itemcity & ") "
        End If
        sql &= " )" & vbCrLf
        sql &= " ,WC2 AS (" & vbCrLf
        sql &= " select cc.OrgID, cc.OrgName" & vbCrLf
        sql &= " ,count(case when cc.NotOpen = 'N' and cc.IsSuccess = 'Y' then 1 end) Class1" & vbCrLf
        sql &= " ,count(case when cc.NotOpen = 'N' and cc.IsSuccess = 'N' then 1 end ) Class2" & vbCrLf
        sql &= " ,count(case when cc.NotOpen = 'Y' then 1 end) Class3" & vbCrLf
        sql &= " ,count(1) Class4" & vbCrLf
        sql &= " from WC1 cc" & vbCrLf
        sql &= " group by cc.OrgID, cc.OrgName" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " select cc.orgid,cs.StudStatus" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs on cs.ocid =cc.ocid" & vbCrLf
        sql &= " join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        sql &= " join stud_subdata ss2 on ss2.sid =cs.sid" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WS2 AS (" & vbCrLf
        sql &= " select ss.OrgID" & vbCrLf
        sql &= " ,count(1) student1" & vbCrLf
        sql &= " ,count(case when ss.StudStatus = '1' then 1 end) student2" & vbCrLf
        sql &= " ,count(case when ss.StudStatus = '5' then 1 end) student3" & vbCrLf
        sql &= " from WS1 ss" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " group by ss.OrgID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " select wc2.OrgID, wc2.OrgName" & vbCrLf
        sql &= " ,WC2.class1" & vbCrLf
        sql &= " ,WC2.class2" & vbCrLf
        sql &= " ,WC2.class3" & vbCrLf
        sql &= " ,WC2.class4" & vbCrLf
        sql &= " ,Ws2.Student1" & vbCrLf
        sql &= " ,Ws2.Student2" & vbCrLf
        sql &= " ,Ws2.Student3" & vbCrLf
        sql &= " FROM WC2" & vbCrLf
        sql &= " JOIN WS2 on ws2.orgid =wc2.orgid" & vbCrLf

        'sqlstr = "Select * from (SELECT oo.OrgID, oo.OrgName," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.NotOpen = 'N' and cc.IsSuccess = 'Y' and cc.ComIDNO = oo.ComIDNO " & str1 & ") AS Class1," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.NotOpen = 'N' and cc.IsSuccess = 'N' and cc.ComIDNO = oo.ComIDNO " & str1 & ") AS Class2," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.NotOpen = 'Y'  and cc.ComIDNO = oo.ComIDNO " & str1 & ") AS Class3," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.ComIDNO = oo.ComIDNO " & str1 & ") AS Class4," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.ComIDNO = oo.ComIDNO " & str1 & ") AS student1," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.ComIDNO = oo.ComIDNO and cs.StudStatus = '1' " & str1 & ") AS student2," & vbCrLf
        'sqlstr += "(select count(*) from Class_ClassInfo cc JOIN ID_Plan ip ON ip.PlanID = cc.PlanID JOIN Class_StudentsOfClass cs ON cs.OCID = cc.OCID JOIN Auth_Relship ar ON ar.RID = cc.RID JOIN Org_OrgPlanInfo op ON op.RSID = ar.RSID JOIN ID_ZIP iz ON iz.ZipCode = op.ZipCode where cc.ComIDNO = oo.ComIDNO and cs.StudStatus = '5' " & str1 & ") AS student3" & vbCrLf
        'sqlstr += " FROM Org_OrgInfo oo) aa  where 1 = 1 and (aa.Class1 <> 0 or Class2 <> 0 or Class3 <> 0 or Class4 <> 0 or student1 <> 0 or student2 <> 0 or student3 <> 0) order by aa.OrgID " & vbCrLf

        '以轄區、訓練計畫做排序
        Me.ViewState("cp_sql") = sql
        'Dim dt As DataTable
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        Me.NoData.Text = "<font color=red>查無資料</font>"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.NoData.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True

            'PageControler1.SqlPrimaryKeyDataCreate(sqlstr, "OrgID", "OrgID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OrgID"
            PageControler1.Sort = "OrgID"
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
        ExportStr = "訓練機構名稱" & vbTab & "已開班" & vbTab & "未開班" & vbTab & "不開班" & vbTab & "總開班" & vbTab & "訓練總人數" & vbTab & "在訓總人數" & vbTab & "結訓總人數" & vbTab
        ExportStr += vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        '建立資料面
        For Each dr In table.Rows
            ExportStr = ""
            ExportStr = ExportStr & dr("OrgName") & vbTab
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
        Dim newstr As String = Session("newDistID")
        Dim newplan As String = Session("newTPlanID")
        Dim newcity As String = Session("newICity")

        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=Report&path=TIMS&filename=CP_04_001_02_Rpt"
        'strScript += "&Years=" & newyear & "&DistID=" & newstr & "&TPlanID=" & newplan & ""
        'strScript += "&itemcity=" & newcity & "');" + vbCrLf
        'strScript += "</script>"

        'Page.RegisterStartupScript("window_onload", strScript)

        newstr = Replace(newstr, "\", "")
        newstr = Replace(newstr, "'", "")
        newplan = Replace(newplan, "\", "")
        newplan = Replace(newplan, "'", "")
        newcity = Replace(newcity, "\", "")
        newcity = Replace(newcity, "'", "")
        Dim MyValue As String = ""
        MyValue = "Years=" & newyear
        If newstr <> "" Then
            MyValue += "&DistID=" & newstr
        End If
        If newplan <> "" Then
            MyValue += "&TPlanID=" & newplan
        End If
        If newcity <> "" Then
            MyValue += "&itemcity=" & newcity
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

    End Sub
End Class
