Imports System.Data.SqlClient
Imports Turbo
Imports System.Data
Public Class SV_01_005
    Inherits System.Web.UI.Page
    Protected WithEvents PageControler1 As PageControler

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents IDNO As System.Web.UI.WebControls.TextBox
    Protected WithEvents birth_date As System.Web.UI.WebControls.TextBox
    Protected WithEvents labPageSize As System.Web.UI.WebControls.Label
    Protected WithEvents TxtPageSize As System.Web.UI.WebControls.TextBox
    Protected WithEvents bt_search As System.Web.UI.WebControls.Button
    Protected WithEvents msg As System.Web.UI.WebControls.Label
    Protected WithEvents DG_ClassInfo As System.Web.UI.WebControls.DataGrid
    Protected WithEvents Panel As System.Web.UI.WebControls.Panel
    Protected WithEvents td6 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents td5 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents NY As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents search_tbl As System.Web.UI.HtmlControls.HtmlTable

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region
    Dim objconn As SqlConnection
    Dim objreader As SqlDataReader
    Dim ProcessType As String
    Dim space As String
    Dim FunDr As DataRow
    Dim RelshipTable As DataTable

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) '☆
        '檢查Session是否存在--------------------------End

        objconn = DbAccess.GetConnection()
        '分頁設定---------------Start
        PageControler1.PageDataGrid = DG_ClassInfo
        '分頁設定---------------End

        If Not IsPostBack Then
            bt_search.Attributes("onclick") = "return CheckData();"
            msg.Text = "在問卷調查起迄期間內且起迄日期均有值才可新增或修改"
            msg.Visible = False
        End If

        Dim sql As String
        sql = "SELECT a.RID,a.Relship,b.OrgName FROM "
        sql += "Auth_Relship a "
        sql += "JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "
        RelshipTable = DbAccess.GetDataTable(sql)
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('" & "顯示列數不正確，以10 帶入" & "');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('" & "顯示列數不正確，以10 帶入" & "');</script>")
            Me.TxtPageSize.Text = 10
        End If
        Me.DG_ClassInfo.PageSize = Me.TxtPageSize.Text

        Me.Panel.Visible = True

        If CheckIDNO(Me.IDNO.Text) = False Then
            Panel.Visible = False
            DG_ClassInfo.Visible = False
            Turbo.Common.MessageBox(Me, "身分證號碼格式錯誤!")
            Exit Sub
        End If

        Dim sqlAdapter As SqlClient.SqlDataAdapter
        Dim class_info As DataTable
        Dim sqlstr As String

        sqlstr = "select a.OCID,a.QaySDate,a.QayFDate,a.RID,b.StudentID,c.Name,e.OrgName,b.socid, "
        sqlstr += "(case when a.CyclType <> '00' then a.ClassCName+'(第'+a.CyclType+'期)' else a.ClassCName end) as ClassCName, "
        sqlstr += "kp.QuesID "
        sqlstr += "from Class_ClassInfo a "
        sqlstr += "join Class_StudentsOfClass b on a.OCID=b.OCID "
        sqlstr += "join Stud_StudentInfo c on b.SID=c.SID "
        sqlstr += "join Auth_Relship f  on a.RID=f.RID "
        sqlstr += "join Org_OrgInfo e on f.OrgID=e.Orgid "
        sqlstr += "join ID_Plan ip on ip.Planid = a.Planid "
        sqlstr += "join Key_Plan kp on kp.TPlanid = ip.TPlanid "
        sqlstr += "where c.IDNO='" & IDNO.Text & "' and c.Birthday='" & birth_date.Text & "' "
        sqlstr += "and a.QaySDate <= getdate() and a.QayFDate >= getdate() and kp.QuesID is not NULL "
        sqlstr += "order by a.OCID "

        If TIMS.Get_SQLRecordCount(sqlstr) = 0 Then
            Panel.Visible = False
            DG_ClassInfo.Visible = False
            Turbo.Common.MessageBox(Me, "查無資料!!")
        Else
            Panel.Visible = True
            DG_ClassInfo.Visible = True

            PageControler1.SqlString = sqlstr
            PageControler1.PrimaryKey = "OCID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('" & "顯示列數不正確，以10 帶入" & "');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('" & "顯示列數不正確，以10 帶入" & "');</script>")
            Me.TxtPageSize.Text = 10
        End If
        Me.DG_ClassInfo.PageSize = Me.TxtPageSize.Text

        Dim dr As DataRowView
        Dim is_parent As String
        dr = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim btn_Edit As Button = e.Item.Cells(8).FindControl("Edit")

                If dr("RID").ToString.Length <> 1 Then
                    If RelshipTable.Select("RID='" & dr("RID") & "'").Length <> 0 Then
                        Dim Relship As String
                        Dim Parent As String
                        Relship = RelshipTable.Select("RID='" & dr("RID") & "'")(0)("Relship")
                        Parent = Split(Relship, "/")(Split(Relship, "/").Length - 3)
                        If RelshipTable.Select("RID='" & Parent & "'").Length <> 0 Then
                            e.Item.Cells(2).Text = RelshipTable.Select("RID='" & Parent & "'")(0)("OrgName")
                        End If
                    End If
                End If

                '填寫狀況
                Dim sqlstr_list As String = "select * from stud_survey where socid ='" & e.Item.Cells(11).Text & "'"
                If DbAccess.GetCount(sqlstr_list) > 0 Then    '已有資料
                    'If (DbAccess.GetCount(sqlstr_list) > 0 Or DbAccess.GetCount(sqlstr) > 0) Then   '已有資料
                    e.Item.Cells(7).Text = "是"
                    btn_Edit.Text = "修改"
                Else
                    e.Item.Cells(7).Text = "否"
                    btn_Edit.Text = "新增"
                End If

                '問卷期間起日
                If dr("QaySDate").ToString <> "" Then e.Item.Cells(5).Text = Turbo.Common.FormatDate(dr("QaySDate").ToString)

                '問卷期間迄日
                If dr("QayFDate").ToString <> "" Then e.Item.Cells(6).Text = Turbo.Common.FormatDate(dr("QayFDate").ToString)

                If (dr("QaySDate").ToString <> "") And (dr("QayFDate").ToString <> "") Then
                    If (CDate(dr("QaySDate")) <= CDate(Turbo.Common.FormatDate(Now()))) And (CDate(dr("QayFDate")) >= CDate(Turbo.Common.FormatDate(Now()))) Then
                        btn_Edit.Enabled = True
                        If e.Item.Cells(7).Text = "是" Then

                            btn_Edit.Attributes("onclick") = "window.open('SV_08_004_Insert.aspx?Type=E&inline=1&OCID=" & e.Item.Cells(9).Text & "&Stuedntid=" & e.Item.Cells(10).Text & "&IDNO=" & IDNO.Text & "&socid=" & e.Item.Cells(11).Text & "&SVID=" & e.Item.Cells(12).Text & "'); return false;"
                        Else
                            Dim sql As String
                            Dim dr5 As DataRow
                            sql = "SELECT DLID from Stud_DataLid where OCID = " & e.Item.Cells(9).Text & ""
                            dr5 = DbAccess.GetOneRow(sql)
                            If dr5 Is Nothing Then
                                btn_Edit.Enabled = False
                                btn_Edit.ToolTip = "結訓資料卡封面檔沒有填,請聯絡承辦人員"
                            End If

                            btn_Edit.Attributes("onclick") = "window.open('SV_08_004_Insert.aspx?Type=I&inline=1&BtnName=bt_search&OCID=" & e.Item.Cells(9).Text & "&Stuedntid=" & e.Item.Cells(10).Text & "&IDNO=" & IDNO.Text & "&socid=" & e.Item.Cells(11).Text & "&SVID=" & e.Item.Cells(12).Text & "'); return false;"
                        End If
                    End If
                Else
                    btn_Edit.Enabled = False
                End If
            Case ListItemType.Header
                If Not Me.ViewState("sort") Is Nothing Then
                    Dim img As New UI.WebControls.Image
                    Dim i As Integer
                    Select Case Me.ViewState("sort")
                        Case "OrgName", "OrgName desc"
                            i = 3
                    End Select

                    If Me.ViewState("sort").ToString.IndexOf("desc") = -1 Then
                        img.ImageUrl = "../../images/SortUp.gif"
                    Else
                        img.ImageUrl = "../../images/SortDown.gif"
                    End If
                    e.Item.Cells(i).Controls.Add(img)
                End If
        End Select

    End Sub

    Private Sub DG_ClassInfo_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_ClassInfo.SortCommand
        If e.SortExpression = Me.ViewState("sort") Then
            Me.ViewState("sort") = e.SortExpression & " desc"
        Else
            Me.ViewState("sort") = e.SortExpression
        End If

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    '檢查身分證號碼
    Function CheckIDNO(ByVal IDNO As String) As Boolean
        Dim ErrString As String = ""
        Dim IDString As String
        IDNO = IDNO.ToUpper             '將字串換成大寫
        If IDNO.Length <> 10 Then       '非10個字元的身分號則退出檢驗
            Return False
        End If
        Dim IDdigit As New ArrayList
        For i As Integer = 0 To 9
            IDdigit.Add(IDNO.Chars(i))
        Next

        Dim CharEng As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        IDdigit(0) = CharEng.IndexOf(IDNO.Chars(0))
        If CharEng.IndexOf(IDNO.Chars(0)) = -1 Then         '檢查第一個字是否為英文字
            Return False
        Else
            If (IDNO.Chars(1) <> "1" And IDNO.Chars(1) <> "2") Then         '檢查第二個字元是否為1或2
                Return False
            End If
        End If
        Dim Array1 As New ArrayList
        Array1.Add(1)
        Array1.Add(10)
        Array1.Add(19)
        Array1.Add(28)
        Array1.Add(37)
        Array1.Add(46)
        Array1.Add(55)
        Array1.Add(64)
        Array1.Add(39)
        Array1.Add(73)
        Array1.Add(82)
        Array1.Add(2)
        Array1.Add(11)
        Array1.Add(20)
        Array1.Add(48)
        Array1.Add(29)
        Array1.Add(38)
        Array1.Add(47)
        Array1.Add(56)
        Array1.Add(65)
        Array1.Add(74)
        Array1.Add(83)
        Array1.Add(21)
        Array1.Add(3)
        Array1.Add(12)
        Array1.Add(30)
        Dim result As String = Array1(IDdigit(0))
        For i As Integer = 1 To 9
            Dim Number As String = "0123456789"
            IDdigit(i) = Number.IndexOf(IDdigit(i))

            If IDdigit(i) = -1 Then
                Return False
            Else
                result += IDdigit(i) * (9 - i)
            End If
        Next

        result += 1 * IDdigit(9)
        If result Mod 10 <> 0 Then
            Return False
        Else
            Return True
        End If
    End Function

End Class
