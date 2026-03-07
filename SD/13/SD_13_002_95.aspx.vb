Public Class SD_13_002_95
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    End Sub

    Protected WithEvents center As System.Web.UI.WebControls.TextBox
    Protected WithEvents HistoryRID As System.Web.UI.WebControls.Table
    Protected WithEvents TMID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents OCID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents HistoryTable As System.Web.UI.WebControls.Table
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents msg As System.Web.UI.WebControls.Label
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents RIDValue As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Button2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents TMIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents OCIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents DataGridTable As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents TitleLab1 As System.Web.UI.WebControls.Label
    Protected WithEvents TitleLab2 As System.Web.UI.WebControls.Label
    Protected WithEvents Button4 As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在--------------------------End

        msg.Text = ""

        If Not IsPostBack Then
            DataGridTable.Style("display") = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

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
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button2.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg.aspx');"
        Else
            Button2.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg1.aspx');"
        End If
        Button1.Attributes("onclick") = "return CheckSearch();"
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        sql = ""
        sql += " SELECT d.SOCID,Right(d.StudentID,2) as StudentID,e.Name,e.IDNO " & vbCrLf
        '-- by nick 20060410     by amu 後來改正不用除人數 2008-01-08
        'sql += "       ,ISNULL(c.Total,0)/b.TNum as Total,d.CreditPoints, " & vbCrLf
        sql += "        ,ISNULL(c.Total,0) as Total,d.CreditPoints, " & vbCrLf
        sql += "        a.THours,ISNULL(f.CountHours,0) as CountHours,e.DegreeID, " & vbCrLf
        sql += "        d.StudStatus,d.MIdentityID,a.STDate, " & vbCrLf
        sql += "        g.SOCID as Exist,g.SumOfMoney,g.PayMoney,g.AppliedStatus,g.AppliedNote " & vbCrLf
        sql += " FROM Class_ClassInfo a " & vbCrLf
        sql += " JOIN Plan_PlanInfo b ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo " & vbCrLf
        'sql += " LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,convert(int,Sum(ISNULL(OPrice,1)*ISNULL(Itemage,1)*ISNULL(ItemCost,1))) as Total FROM Plan_CostItem Group By PlanID,ComIDNO,SeqNo) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo "
        sql += " LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,convert(int,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)),0) as Total " & vbCrLf
        sql += " FROM Plan_CostItem " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql += " WHERE COSTMODE = 5 " & vbCrLf
        Else
            sql += " WHERE COSTMODE <> 5 " & vbCrLf
        End If
        sql += " Group By PlanID,ComIDNO,SeqNo) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo " & vbCrLf
        sql += " JOIN Class_StudentsOfClass d ON a.OCID=d.OCID " & vbCrLf
        sql += " JOIN Stud_StudentInfo e ON d.SID=e.SID " & vbCrLf
        sql += " LEFT JOIN (SELECT SOCID,Sum(Hours) as CountHours FROM Stud_Turnout2 Group By SOCID) f ON d.SOCID=f.SOCID " & vbCrLf
        sql += " JOIN Stud_SubsidyCost g ON d.SOCID=g.SOCID " & vbCrLf
        sql += " WHERE a.OCID='" & OCIDValue1.Value & "' order by StudentID ASC " & vbCrLf
        dt = DbAccess.GetDataTable(sql)
        If dt.Rows.Count = 0 Then
            DataGridTable.Style("display") = "none"
            msg.Text = "查無資料"
        Else
            DataGridTable.Style("display") = "inline"
            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"
                Dim DropDownList1 As DropDownList = e.Item.FindControl("DropDownList1")
                DropDownList1.Attributes("onchange") = "SelectAll(this.selectedIndex)"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim Flag As Integer = 0
                Dim drv As DataRowView = e.Item.DataItem
                Dim AppliedStatus As DropDownList = e.Item.FindControl("AppliedStatus")
                Dim AppliedNote As TextBox = e.Item.FindControl("AppliedNote")
                If IsDBNull(drv("CreditPoints")) Then
                    e.Item.Cells(3).Text = "否"
                Else
                    If drv("CreditPoints") Then
                        e.Item.Cells(3).Text = "是"
                        Flag += 1
                    Else
                        e.Item.Cells(3).Text = "否"
                    End If
                End If

                If (drv("THours") - drv("CountHours")) / drv("THours") > 2 / 3 Then
                    e.Item.Cells(4).Text = "是"
                Else
                    e.Item.Cells(4).Text = "否"
                End If
                Dim sql As String
                Dim dr As DataRow
                Dim Total As Integer = 20000            '可用餘額(最早是提供2萬)
                sql = "SELECT ISNULL(Sum(SumOfMoney),0) as SumOfMoney FROM view_SubsidyCost WHERE IDNO='" & drv("IDNO") & "' and STDate>=DateAdd(Year,-5,'" & drv("STDate") & "') and STDate<='" & drv("STDate") & "' and SOCID<>'" & drv("SOCID") & "'"
                dr = DbAccess.GetOneRow(sql)
                If Not dr Is Nothing Then
                    Total -= dr("SumOfMoney")
                    If Total < 0 Then Total = 0
                End If
                If Total - drv("SumOfMoney") >= 0 Then
                    e.Item.Cells(9).Text = Total - drv("SumOfMoney")
                Else
                    e.Item.Cells(9).Text = "<font color=Red>" & Total - drv("SumOfMoney") & "</font>"
                End If

                If Flag = 1 Then
                    e.Item.Cells(5).Text = "是"
                Else
                    e.Item.Cells(5).Text = "否"
                End If
                If IsDBNull(drv("AppliedStatus")) Then
                    AppliedStatus.SelectedIndex = 0
                Else
                    If drv("AppliedStatus") = 1 Then
                        AppliedStatus.SelectedIndex = 1
                    Else
                        AppliedStatus.SelectedIndex = 2
                    End If
                End If
                AppliedNote.Text = drv("AppliedNote").ToString
        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim conn As SqlConnection = DbAccess.GetConnection
        Dim sql As String
        Dim dr As DataRow
        Dim dt As DataTable
        Dim da As SqlDataAdapter
        Dim i As Integer
        sql = "SELECT * FROM Stud_SubsidyCost WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "')"
        dt = DbAccess.GetDataTable(sql, da, conn)
        For Each item As DataGridItem In DataGrid1.Items
            Dim AppliedStatus As DropDownList = item.FindControl("AppliedStatus")
            Dim AppliedNote As TextBox = item.FindControl("AppliedNote")
            dr = dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'")(0)
            Select Case AppliedStatus.SelectedIndex
                Case 0
                    dr("AppliedStatus") = Convert.DBNull
                Case 1
                    dr("AppliedStatus") = 1
                Case 2
                    dr("AppliedStatus") = 0
            End Select
            dr("AppliedNote") = IIf(AppliedNote.Text = "", Convert.DBNull, AppliedNote.Text)
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            i += 1
        Next
        DbAccess.UpdateDataTable(dt, da)
        Turbo.Common.MessageBox(Me, "儲存成功")
        Button1_Click(sender, e)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow
        dr = TIMS.GET_OnlyOne_OCID(RIDValue.Value) '判斷機構是否只有一個班級
        If Not dr Is Nothing Then
            If dr("total") = "1" Then '如果只有一個班級
                TMID1.Text = dr("trainname")
                OCID1.Text = dr("classname")
                TMIDValue1.Value = dr("trainid")
                OCIDValue1.Value = dr("ocid")
                DataGridTable.Style("display") = "none"
            Else '不只一個班級
                TMID1.Text = ""
                OCID1.Text = ""
                TMIDValue1.Value = ""
                OCIDValue1.Value = ""
                DataGridTable.Style("display") = "none"
            End If
        Else
            TMID1.Text = ""
            OCID1.Text = ""
            TMIDValue1.Value = ""
            OCIDValue1.Value = ""
            DataGridTable.Style("display") = "none"
        End If
    End Sub
End Class