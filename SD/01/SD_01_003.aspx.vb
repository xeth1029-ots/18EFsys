Imports System.Data.SqlClient
Imports System.Data
Imports Turbo
Public Class SD_01_003
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents msg As System.Web.UI.WebControls.Label
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents msg2 As System.Web.UI.WebControls.Label
    Protected WithEvents DG_stud As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel
    Protected WithEvents center As System.Web.UI.WebControls.TextBox
    Protected WithEvents TMID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents OCID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents search As System.Web.UI.WebControls.Button
    Protected WithEvents StudentTable As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents RIDValue As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents TMIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents OCIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents check_search As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents check_add As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents check_mod As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents check_del As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents HistoryTable As System.Web.UI.WebControls.Table
    Protected WithEvents HistoryRID As System.Web.UI.WebControls.Table
    Protected WithEvents Button2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents TitleLab1 As System.Web.UI.WebControls.Label
    Protected WithEvents TitleLab2 As System.Web.UI.WebControls.Label

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region
    'Dim sqlAdapter As SqlClient.SqlDataAdapter
    'Dim objconn As SqlConnection
    Dim stud_table, class_table As DataTable
    Dim FunDr As DataRow

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        msg.Text = ""
        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button2.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg.aspx');"
        Else
            Button2.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg1.aspx');"
        End If
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If sm.UserInfo.RoleID <> 0 Then
            If sm.UserInfo.FunDt Is Nothing Then
                Common.RespWrite(Me, "<script>alert('Session過期');</script>")
                Common.RespWrite(Me, "<script>Top.location.href='../../logout.aspx';</script>")
            Else
                Dim FunDt As DataTable = sm.UserInfo.FunDt
                Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

                If FunDrArray.Length = 0 Then
                    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
                    Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
                Else
                    FunDr = FunDrArray(0)
                    If FunDr("Adds") = "1" Then
                        check_add.Value = "1"
                    Else
                        check_add.Value = "0"
                    End If
                    If FunDr("Sech") = "1" Then
                        search.Enabled = True
                        check_search.Value = "1"
                    Else
                        search.Enabled = False
                        check_search.Value = "0"
                    End If
                    If FunDr("Del") = "1" Then
                        check_del.Value = "1"
                    Else
                        check_del.Value = "0"
                    End If
                    If FunDr("Mod") = "1" Then
                        check_mod.Value = "1"
                    Else
                        check_mod.Value = "0"
                    End If
                End If
            End If
        End If

        If Not Page.IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            StudentTable.Style.Item("display") = "none"

            Dim re_ocid, ClassName, ProcessType As String
            Dim dr As DataRow
            re_ocid = Request("ocid")
            ProcessType = Request("ProcessType")

            If ProcessType = "Back" Then

                If Not Session("QuestionaryStud_EnterQStr") Is Nothing Then
                    Dim MyArray As Array = Split(Session("QuestionaryStud_EnterQStr"), "&")
                    Dim MyItem As String
                    Dim MyValue As String
                    Dim StudentTableState As String

                    For i As Integer = 0 To MyArray.Length - 1
                        MyItem = Split(MyArray(i), "=")(0)
                        MyValue = Split(MyArray(i), "=")(1)
                        Select Case MyItem
                            Case "center"
                                center.Text = MyValue
                            Case "RIDValue"
                                RIDValue.Value = MyValue
                            Case "TMID1"
                                TMID1.Text = MyValue
                            Case "OCID1"
                                OCID1.Text = MyValue
                            Case "TMIDValue1"
                                TMIDValue1.Value = MyValue
                            Case "OCIDValue1"
                                OCIDValue1.Value = MyValue
                            Case "Button1"
                                If MyValue = "True" Then
                                    search_Click(sender, e)
                                End If
                            Case "StudentTable"
                                StudentTableState = MyValue
                        End Select
                    Next
                    Session("QuestionaryStud_EnterQStr") = Nothing
                End If

            End If
        End If

    End Sub

    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        Me.Panel1.Visible = True
        StudentTable.Style.Item("display") = "none"
        msg2.Visible = False
        Label1.Visible = False

        Dim sqlstr_class As String = "select  a.OCID,a.CyclType, a.LevelType, a.ClassCName, a.FTDate,b.total,isnull(c.num,0) as num  from class_classinfo a join  "
        sqlstr_class += " (select OCID,count(*) as total  from class_studentsofclass  WHERE  StudStatus IN (1,4,5) GROUP BY OCID ) b "
        sqlstr_class += "   on a.ocid=b.ocid left  join (select OCID,count(*) as num  from Stud_EnterQMain   group by OCID) c on c.OCID=a.OCID"
        sqlstr_class += " where 1=1 "
        If RIDValue.Value <> "" Then
            sqlstr_class += " and   a.RID='" & RIDValue.Value & "'"
        End If
        If OCIDValue1.Value <> "" Then
            sqlstr_class += "  and a.OCID='" & OCIDValue1.Value & "' "
        End If
        sqlstr_class += " ORDER BY a.OCID"

        'class_table = DbAccess.GetDataTable(sqlstr_class, sqlAdapter, objconn)
        class_table = DbAccess.GetDataTable(sqlstr_class)

        DataGrid1.Visible = False
        DataGrid1.Style.Item("display") = "none"
        msg.Visible = True
        msg.Text = "查無資料!!"

        If class_table.Rows.Count > 0 Then
            DataGrid1.Visible = True
            DataGrid1.Style.Item("display") = "inline"
            msg.Visible = False
            msg.Text = ""

            DataGrid1.DataKeyField = "OCID"
            DataGrid1.DataSource = class_table
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1


            Dim drv As DataRowView = e.Item.DataItem
            Dim OCID_Namestr As String
            OCID_Namestr = drv("ClassCName").ToString
            If CInt(e.Item.Cells(6).Text) <> 0 Then
                OCID_Namestr += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(6).Text)) & "期"
            End If
            e.Item.Cells(1).Text = OCID_Namestr

            Dim mybut1 As Button = e.Item.Cells(4).FindControl("Button1")
            If check_search.Value = "1" Then
                mybut1.Enabled = True
            Else
                mybut1.Enabled = False
            End If

            mybut1.CommandArgument = DataGrid1.DataKeys(e.Item.ItemIndex)

        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "view" Then
            Dim sqlstr_stud As String
            sqlstr_stud = "SELECT a.*,ISNULL(b.StudentCount,0) as StudentCount,ISNULL(c.TrainCount,0) as TrainCount,ISNULL(d.LeaveCount,0) as LeaveCount FROM "
            sqlstr_stud += "(SELECT * FROM Class_ClassInfo WHERE OCID='" & e.CommandArgument & "') a "
            sqlstr_stud += "LEFT JOIN (SELECT OCID,Count(*) AS StudentCount FROM Class_StudentsOfClass WHERE OCID='" & e.CommandArgument & "' Group By OCID) b ON a.OCID=b.OCID "
            sqlstr_stud += "LEFT JOIN (SELECT OCID,Count(*) AS TrainCount FROM Class_StudentsOfClass WHERE OCID='" & e.CommandArgument & "' and StudStatus IN (1,4,5) Group By OCID) c ON a.OCID=c.OCID "
            sqlstr_stud += "LEFT JOIN (SELECT OCID,Count(*) AS LeaveCount FROM Class_StudentsOfClass WHERE OCID='" & e.CommandArgument & "' and StudStatus IN (2,3) Group By OCID) d ON a.OCID=d.OCID "

            Dim dr As DataRow = DbAccess.GetOneRow(sqlstr_stud)
            If Not dr Is Nothing Then
                Label1.Text = "班別：" & dr("ClassCName")
                If CInt(dr("CyclType")) <> 0 Then
                    Label1.Text += "第" & TIMS.GetChtNum(CInt(dr("CyclType"))) & "期"
                End If

                If Not IsDBNull(dr("LevelType")) Then
                    If CInt(dr("LevelType")) <> 0 Then
                        Label1.Text += "第" & TIMS.GetChtNum(CInt(dr("LevelType"))) & "階段"
                    End If
                End If

                Label1.Text += "(開訓人數:" & dr("StudentCount").ToString & "&nbsp;&nbsp;在結訓人數:" & dr("TrainCount").ToString & "&nbsp;&nbsp;離退訓人數:" & dr("LeaveCount").ToString & ")"
            End If
            sqlstr_stud = "  select b.studentid,b.StudStatus,c.name,b.OCID,b.RejectTDate1,b.RejectTDate2,b.SOCID  from class_classinfo a join class_studentsofclass b on a.ocid=b.ocid join stud_studentinfo c on b.sid=c.sid WHERE  a.ocid='" & e.CommandArgument & "' order by b.studentid"
            Dim dt As DataTable = DbAccess.GetDataTable(sqlstr_stud)
            If dt.Rows.Count = 0 Then
                msg2.Text = "查無此班學生資料!"
                StudentTable.Style.Item("display") = "none"
                Label1.Visible = False
            Else
                Label1.Visible = True
                Session("QTable_Stuednt") = dt
                StudentTable.Style.Item("display") = "inline"
                DG_stud.DataSource = dt
                'DG_stud.DataKeyField = "SOCID"
                DG_stud.DataBind()
            End If
        End If
    End Sub

    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Dim drv As DataRowView
        drv = e.Item.DataItem
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = Right(drv("StudentID"), 2)
            Dim but5 As Button
            Dim but4, but6 As Button
            but4 = e.Item.Cells(3).FindControl("Button4") '新增
            but5 = e.Item.Cells(3).FindControl("Button5") '查看
            but6 = e.Item.Cells(3).FindControl("Button6") '清除重填

            If drv("RejectTDate1").ToString <> "" Then
                e.Item.Cells(1).Text += "(" & FormatDateTime(drv("RejectTDate1"), 2) & ")"
            End If
            If drv("RejectTDate2").ToString <> "" Then
                e.Item.Cells(1).Text += "(" & FormatDateTime(drv("RejectTDate2"), 2) & ")"
            End If

            Dim sqlstr_list As String = "select * from Stud_EnterQMain where OCID='" & e.Item.Cells(4).Text & "' and SOCID= '" & e.Item.Cells(5).Text & "'"
            If DbAccess.GetCount(sqlstr_list) > 0 Then '已有資料
                e.Item.Cells(2).Text = "是"
                but4.Enabled = False '不可新增
                but5.Enabled = True '可以查看
                but6.Enabled = True '可清除重填
                If check_mod.Value = "0" And check_del.Value = "0" Then '兩者功能皆沒有時,不能使用
                    but6.Enabled = False
                Else
                    but6.Enabled = True
                End If
                If check_search.Value = "1" Then
                    but5.Enabled = True
                Else
                    but5.Enabled = False
                End If
            Else
                e.Item.Cells(2).Text = "否"
                but4.Enabled = True '可新增
                but5.Enabled = False '不可查看
                but6.Enabled = False '不可清除重填
                If check_add.Value = "1" Then
                    but4.Enabled = True
                Else
                    but4.Enabled = False
                End If

            End If

        End If
    End Sub

    Private Sub DG_stud_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_stud.ItemCommand
        If e.CommandName = "insert" Then
            GetSearchStr()
            TIMS.Utl_Redirect1(Me, "SD_01_003_add.aspx?ProcessType=Insert&SOCID=" & e.Item.Cells(5).Text & "&ocid=" & e.Item.Cells(4).Text & "&ID=" & Request("ID") & "&studentid=" & e.Item.Cells(6).Text)
        ElseIf e.CommandName = "clear" Then
            GetSearchStr()
            Dim strScript As String
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "if (window.confirm('此動作會刪除期末學員滿意度調查表資料，是否確定刪除?')){" + vbCrLf
            strScript += "location.href ='SD_01_003_add.aspx?ProcessType=del&SOCID=" & e.Item.Cells(5).Text & "&ocid=" & e.Item.Cells(4).Text & "&ID=" & Request("ID") & "&studentid=" & e.Item.Cells(6).Text & "';}" + vbCrLf
            strScript += "</script>"
            Page.RegisterStartupScript("", strScript)
        Else
            GetSearchStr()
            TIMS.Utl_Redirect1(Me, "SD_01_003_add.aspx?ProcessType=check&SOCID=" & e.Item.Cells(5).Text & "&ocid=" & e.Item.Cells(4).Text & "&ID=" & Request("ID") & "&studentid=" & e.Item.Cells(6).Text)
        End If
    End Sub

    Function GetSearchStr()
        '取新的名稱
        Session("QuestionaryStud_EnterQStr") += "&center=" & center.Text
        Session("QuestionaryStud_EnterQStr") += "&RIDValue=" & RIDValue.Value
        Session("QuestionaryStud_EnterQStr") = "TMID1=" & TMID1.Text
        Session("QuestionaryStud_EnterQStr") += "&OCID1=" & OCID1.Text
        Session("QuestionaryStud_EnterQStr") += "&TMIDValue1=" & TMIDValue1.Value
        Session("QuestionaryStud_EnterQStr") += "&OCIDValue1=" & OCIDValue1.Value
        Session("QuestionaryStud_EnterQStr") += "&Button1=" & DG_stud.Visible
        Session("QuestionaryStud_EnterQStr") += "&StudentTable=" & StudentTable.Style.Item("display")
    End Function
End Class
