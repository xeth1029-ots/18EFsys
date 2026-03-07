Public Class SD_13_001_95
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
    Protected WithEvents RIDValue As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Button2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents TMIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents OCIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents DataGridTable As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents msg As System.Web.UI.WebControls.Label
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
        '在這裡放置使用者程式碼以初始化網頁
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
        Button3.Attributes("onclick") = "return CheckData();"
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        'change by nick 20060410
        ' sql = "SELECT d.SOCID,Right(d.StudentID,2) as StudentID,e.Name,e.IDNO,ISNULL(c.Total,0)/a.TNum as Total,d.CreditPoints, "
        sql = " SELECT d.SOCID ,RIGHT(d.StudentID,2) AS StudentID ,e.Name ,e.IDNO ,c.Total ,d.CreditPoints, a.THours ,ISNULL(f.CountHours,0) as CountHours ,e.DegreeID, "
        sql += "       d.StudStatus ,d.MIdentityID ,a.STDate ,g.SOCID as Exist ,g.SumOfMoney ,g.PayMoney ,g.AppliedStatus ,g.AppliedNote "
        sql += " FROM Class_ClassInfo a "
        sql += " JOIN Plan_PlanInfo b ON a.PlanID=b.PlanID and a.ComIDNO=b.ComIDNO and a.SeqNo=b.SeqNo "
        'sql += " LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,convert(int,Sum(ISNULL(OPrice,1)*ISNULL(Itemage,1)*ISNULL(ItemCost,1)),0) as Total FROM Plan_CostItem Group By PlanID,ComIDNO,SeqNo) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo "
        'sql += " LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,convert(int,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)),0) as Total FROM Plan_CostItem Group By PlanID,ComIDNO,SeqNo) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo "
        sql += " LEFT JOIN (SELECT PlanID,ComIDNO,SeqNo,convert(int,Sum(ISNULL(OPrice,1)*ISNULL(ItemCost,1)),0) as Total " & vbCrLf
        sql += "            FROM Plan_CostItem " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql += " WHERE COSTMODE = 5 " & vbCrLf
        Else
            sql += " WHERE COSTMODE <> 5 " & vbCrLf
        End If
        sql += " GROUP BY PlanID,ComIDNO,SeqNo) c ON a.PlanID=c.PlanID and a.ComIDNO=c.ComIDNO and a.SeqNo=c.SeqNo " & vbCrLf
        sql += "JOIN Class_StudentsOfClass d ON a.OCID=d.OCID "
        sql += "JOIN Stud_StudentInfo e ON d.SID=e.SID "
        sql += "LEFT JOIN (SELECT SOCID,Sum(Hours) as CountHours FROM Stud_Turnout2 Group By SOCID) f ON d.SOCID=f.SOCID "
        sql += "LEFT JOIN Stud_SubsidyCost g ON d.SOCID=g.SOCID "
        sql += "WHERE a.OCID='" & OCIDValue1.Value & "' order by StudentID ASC"
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
                e.Item.Cells(5).ToolTip = "是否有請領補助津貼的資格"
                e.Item.Cells(7).ToolTip = "預定要補助的金額(可自行變動，未申請前系統會根據可用餘額推算)"
                e.Item.Cells(8).ToolTip = "學員自行要支付的金額(會根據補助費用所輸入的值來調動)"
                e.Item.Cells(9).ToolTip = "學員目前可用餘額-這次預定補助費用的剩餘金額(成為負數時會以紅字表示)"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim Flag As Integer = 0
                Dim drv As DataRowView = e.Item.DataItem
                Dim DataGrid2 As DataGrid = e.Item.FindControl("DataGrid2")
                Dim CreditPoints As Label = e.Item.FindControl("CreditPoints")
                Dim SumOfMoney As TextBox = e.Item.FindControl("SumOfMoney")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim RemainSub As HtmlInputHidden = e.Item.FindControl("RemainSub")
                Dim MaxSub As HtmlInputHidden = e.Item.FindControl("MaxSub")
                Dim PayMoney As HtmlInputHidden = e.Item.FindControl("PayMoney")
                If IsDBNull(drv("CreditPoints")) Then
                    CreditPoints.Text = "<font color='RED'>否</font>"
                Else
                    If drv("CreditPoints") Then
                        CreditPoints.Text = "是"
                        Flag = 1
                    Else
                        CreditPoints.Text = "<font color='RED'>否</font>"
                    End If
                End If
                If (drv("THours") - drv("CountHours")) / drv("THours") > 2 / 3 Then
                    e.Item.Cells(4).Text = "是"
                Else
                    e.Item.Cells(4).Text = "否"
                End If
                If Int(drv("DegreeID")) <= 2 Then
                End If
                Dim sql As String
                Dim dr As DataRow
                Dim dt As DataTable
                Dim Total As Integer = 20000            '可用餘額
                sql = "SELECT b.ClassCName,a.SumOfMoney,case a.AppliedStatus When 1 then '審核通過' else '申請中' end as AppliedStatus FROM "
                sql += "(SELECT * FROM view_SubsidyCost WHERE IDNO='" & drv("IDNO") & "' and STDate>=DateAdd(Year,-5,'" & drv("STDate") & "') and STDate<='" & drv("STDate") & "' and SOCID<>'" & drv("SOCID") & "') a "
                sql += "JOIN view_StudentBasicData b ON a.SOCID=b.SOCID "
                dt = DbAccess.GetDataTable(sql)
                If dt.Rows.Count = 0 Then
                    DataGrid2.Visible = False
                Else
                    DataGrid2.Style("display") = "none"
                    DataGrid2.Style("POSITION") = "absolute"
                    DataGrid2.Visible = True
                    DataGrid2.DataSource = dt
                    DataGrid2.DataBind()
                    For Each dr In dt.Rows
                        Total -= dr("SumOfMoney")
                    Next
                End If
                If Total < 0 Then Total = 0
                RemainSub.Value = Total
                If IsDBNull(drv("Exist")) Then      '表示沒資料,以新增的型態顯示
                    If Flag = 1 Then
                        e.Item.Cells(5).Text = "是"
                        If drv("MIdentityID").ToString <> "" Then
                            If drv("MIdentityID").ToString = "01" Then
                                If Total >= Int(Math.Round(drv("Total") * 0.8)) Then
                                    SumOfMoney.Text = Int(Math.Round(drv("Total") * 0.8))
                                Else
                                    SumOfMoney.Text = Total
                                End If
                            Else
                                If Total >= drv("Total") Then
                                    SumOfMoney.Text = drv("Total")
                                Else
                                    SumOfMoney.Text = Total
                                End If
                            End If
                            MaxSub.Value = SumOfMoney.Text
                            e.Item.Cells(8).Text = drv("Total") - SumOfMoney.Text
                            PayMoney.Value = drv("Total") - SumOfMoney.Text
                            e.Item.Cells(9).Text = Total - SumOfMoney.Text
                        Else
                            SumOfMoney.Enabled = False
                            Checkbox1.Disabled = True
                            e.Item.Cells(9).Text = Total
                        End If
                    Else
                        e.Item.Cells(5).Text = "<font color='RED'>否</font>"
                        SumOfMoney.Enabled = False
                        Checkbox1.Disabled = True
                        e.Item.Cells(9).Text = Total
                    End If
                Else
                    If Flag = 1 Then
                        e.Item.Cells(5).Text = "是"
                    Else
                        e.Item.Cells(5).Text = "<font color='RED'>否</font>"
                    End If
                    If drv("MIdentityID").ToString = "01" Then
                        If Total >= Int(Math.Round(drv("Total") * 0.8)) Then
                            MaxSub.Value = Int(Math.Round(drv("Total") * 0.8))
                        Else
                            MaxSub.Value = Total
                        End If
                    Else
                        If Total >= drv("Total") Then
                            MaxSub.Value = drv("Total")
                        Else
                            MaxSub.Value = Total
                        End If
                    End If
                    SumOfMoney.Text = drv("SumOfMoney").ToString
                    PayMoney.Value = drv("PayMoney").ToString
                    e.Item.Cells(8).Text = drv("PayMoney").ToString
                    Checkbox1.Checked = True
                    If IsDBNull(drv("AppliedStatus")) Then
                        Checkbox1.Disabled = False
                        e.Item.Cells(11).Text = "審核中"
                    Else
                        If drv("AppliedStatus") Then
                            Checkbox1.Disabled = True
                            SumOfMoney.ReadOnly = True
                            e.Item.Cells(11).Text = "審核通過"
                        Else
                            Checkbox1.Disabled = True
                            e.Item.Cells(11).Text = "審核失敗"
                        End If
                    End If
                    If Total - SumOfMoney.Text >= 0 Then
                        e.Item.Cells(9).Text = Total - SumOfMoney.Text
                    Else
                        e.Item.Cells(9).Text = "<font color=Red>" & Total - SumOfMoney.Text & "</font>"
                    End If
                End If
                For i As Integer = 0 To 2
                    e.Item.Cells(i).Attributes("onmouseover") = "if(document.getElementById('" & DataGrid2.ClientID & "')){document.getElementById('" & DataGrid2.ClientID & "').style.display='inline';}"
                    e.Item.Cells(i).Attributes("onmouseout") = "if(document.getElementById('" & DataGrid2.ClientID & "')){document.getElementById('" & DataGrid2.ClientID & "').style.display='none';}"
                    e.Item.Cells(i).Style("CURSOR") = "hand"
                Next
                SumOfMoney.Attributes("onchange") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
                SumOfMoney.Attributes("onblur") = "ChangeMoney(" & e.Item.ItemIndex + 1 & ",'" & SumOfMoney.ClientID & "','" & RemainSub.ClientID & "','" & MaxSub.ClientID & "','" & PayMoney.ClientID & "');"
        End Select
    End Sub

    '儲存按鈕
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
            Dim SumOfMoney As TextBox = item.FindControl("SumOfMoney")
            Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
            Dim RemainSub As HtmlInputHidden = item.FindControl("RemainSub")
            Dim PayMoney As HtmlInputHidden = item.FindControl("PayMoney")
            'If Checkbox1.Disabled = False Then
            If SumOfMoney.ReadOnly = False Then
                If Checkbox1.Checked = True Then
                    If dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'").Length = 0 Then
                        dr = dt.NewRow()
                        dt.Rows.Add(dr)
                        dr("SOCID") = DataGrid1.DataKeys(i)
                    Else
                        dr = dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'")(0)
                    End If
                    dr("SumOfMoney") = SumOfMoney.Text
                    dr("PayMoney") = PayMoney.Value
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                Else
                    If dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'").Length <> 0 Then dt.Select("SOCID='" & DataGrid1.DataKeys(i) & "'")(0).Delete()
                End If
            End If
            i += 1
        Next
        DbAccess.UpdateDataTable(dt, da)
        Turbo.Common.MessageBox(Me, "儲存成功")
        Button1_Click(sender, e)
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow
        dr = TIMS.GET_ONLYONE_OCID(RIDValue.Value) '判斷機構是否只有一個班級
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