Imports System.Data.SqlClient
Imports System.Data
Imports Turbo

Public Class SD_11_001_add
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    End Sub

    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents TMID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents OCID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents SOCID As System.Web.UI.WebControls.DropDownList
    Protected WithEvents StudStatus As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RejectTDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents RTReasonID As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents NeedPay As System.Web.UI.WebControls.DropDownList
    Protected WithEvents SumOfPay As System.Web.UI.WebControls.TextBox
    Protected WithEvents HadPay As System.Web.UI.WebControls.TextBox
    Protected WithEvents TMIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents OCIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents SLTID As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents IMG1 As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents SumOfPay1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents RadioButtonList1_1 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R1_1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList1_2 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R1_2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents Re_R1_3 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList1_3 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R2_1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList2_1 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R2_2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList2_2 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R2_3 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList2_3 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R2_4 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList2_4 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R2_5 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList2_5 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R3_1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList3_1 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R3_2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList3_2 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R3_3 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList3_3 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R3_4 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList3_4 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RadioButtonList3_5 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RadioButtonList3_6 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RadioButtonList3_7 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R4_1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList4_1 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R4_2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList4_2 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R4_3 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList4_3 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R4_4 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList4_4 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R4_5 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList4_5 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R4_6 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList4_6 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R5_1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList5_1 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RadioButtonList5_2 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RadioButtonList5_4 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R6_1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList6_1 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R6_2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList6_2 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R6_3 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList6_3 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R6_4 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList6_4 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Re_R6_5 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RadioButtonList6_5 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Summary As System.Web.UI.WebControls.ValidationSummary
    Protected WithEvents Re_OCID As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Re_Studentid As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents clientscript As System.Web.UI.WebControls.Literal
    Protected WithEvents CustomValidator1 As System.Web.UI.WebControls.CustomValidator
    Protected WithEvents Label_Stud As System.Web.UI.WebControls.Label
    Protected WithEvents Label_Name As System.Web.UI.WebControls.Label
    Protected WithEvents ProcessType As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Button2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents Re_ID As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents next_but As System.Web.UI.WebControls.Button
    Protected WithEvents Label_Status As System.Web.UI.WebControls.Label
    Protected WithEvents StdTr As System.Web.UI.HtmlControls.HtmlTableRow
    Protected WithEvents Label_R1_1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R1_2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R1_3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R2_1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R2_2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R6_1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R6_2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R6_3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R6_4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R6_5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R2_3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R2_4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R2_5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_6 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R3_7 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R4_1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R4_2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R4_3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R4_4 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R4_5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R4_6 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R5_2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R5_3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label_R5_4 As System.Web.UI.WebControls.Label
    Protected WithEvents TD_R3_6 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R3_7 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents RadioButtonList5_3 As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents TD_R1_1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R1_2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R1_3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R2_1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R2_2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R2_3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R2_4 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R2_5 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R3_1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R3_2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R3_3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R3_4 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R3_5 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R4_1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R4_2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R4_3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R4_4 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R4_5 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R4_6 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R5_1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R5_2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R5_3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R5_4 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R6_1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R6_2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R6_3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R6_4 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents TD_R6_5 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents Label_R5_1 As System.Web.UI.WebControls.Label
    Protected WithEvents TD_R6 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents Re_R5_2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents Re_R5_3 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents Re_R5_4 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents Qtype_Value As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Table3 As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents BtnBak As System.Web.UI.WebControls.Button
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
    ''Dim stud_table As DataTable
    'Dim QuestType As String  '問卷類型
    'Stud_Questionary
    Dim FunDr As DataRow

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在--------------------------End

        'Button2.Attributes("onclick") = "history.go(-1);"
        Dim objconn As SqlConnection = DbAccess.GetConnection()
        TIMS.OpenDbConn(Me, objconn)
        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        ProcessType.Value = Request("ProcessType")

        '20080723 andy ------------------新增(問卷類型) start 
        If GetQType() = False Then
            '計畫未設定問卷類型，請先設定後，再進入問卷填寫
            Dim TD_Stud As HtmlTableCell = Me.FindControl("TD_Stud")
            StdTr.Visible = False
            Table3.Style("display") = "none"
            Label1.Visible = True
            Button1.Visible = False
            Button2.Visible = False
            next_but.Visible = False
            BtnBak.Visible = False
            Turbo.Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再進入問卷填寫")
            '離開此功能
            Exit Sub
        End If

        QuestionType() '問卷種類

        Select Case ProcessType.Value
            Case "Insert", "Next"
                Qtype_Value.Value = viewstate("QName")
                If viewState("QName") = "B" Then
                    CustomValidator1.Enabled = False
                    RadioButtonList5_4.Enabled = True
                Else
                    RadioButtonList3_4.Attributes("onclick") = "disable_radio3();"
                    RadioButtonList5_3.Attributes("onclick") = "disable_radio5();"
                End If
        End Select

        'Button2.Attributes("onclick") = "history.go(-1);"

        '20080723 andy ------------------新增(問卷類型)end
        'If sm.UserInfo.RoleID <> 0 Then
        'End If
        If sm.UserInfo.FunDt Is Nothing Then
            Common.RespWrite(Me, "<script>alert('Session過期');</script>")
            Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        Else
            Dim FunDt As DataTable = sm.UserInfo.FunDt
            Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
            Re_ID.Value = Request("ID")

            If FunDrArray.Length = 0 Then
                Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
                Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
            Else
                FunDr = FunDrArray(0)
                If ProcessType.Value = "Update" Then
                    If FunDr("Mod") = "0" AndAlso FunDr("Del") = "0" Then
                        Button1.Enabled = False
                    Else
                        Button1.Enabled = True
                    End If
                ElseIf ProcessType.Value = "Insert" Or ProcessType.Value = "Next" Then
                    If FunDr("Adds") = "1" Then
                        Button1.Enabled = True
                    Else
                        Button1.Enabled = False
                    End If
                End If
            End If
        End If

        If Not IsPostBack Then
            Re_OCID.Value = Request("ocid")
            Re_Studentid.Value = Request("Stuedntid")
            If ProcessType.Value <> "Print" Then
                Dim sql As String
                Dim dr As DataRow
                Dim dt As DataTable
                sql = "SELECT StudentID, CASE WHEN LEN(a.StudentID) = 12 THEN b.Name + '(' + RIGHT(a.StudentID,3) + ')' ELSE b.Name + '(' + RIGHT(a.StudentID,2)+')' END AS Name, a.SOCID "
                sql += "FROM (SELECT * FROM Class_StudentsOfClass WHERE OCID = '" & Request("OCID") & "') a "
                sql += "JOIN (SELECT * FROM Stud_StudentInfo) b ON a.SID = b.SID "
                dt = DbAccess.GetDataTable(sql)
                dt.DefaultView.Sort = "StudentID"
                With SOCID
                    .DataSource = dt
                    .DataTextField = "Name"
                    .DataValueField = "StudentID"
                    .DataBind()
                    .Items.Insert(0, New ListItem("===請選擇===", ""))
                End With
                Turbo.Common.SetListItem(SOCID, Request("Stuedntid"))
                Me.ViewState("QuestionarySearchStr") = Session("QuestionarySearchStr")
                Session("QuestionarySearchStr") = Nothing
                'Re_OCID.Value = Request("ocid")
                'Re_Studentid.Value = Request("Stuedntid")
                Dim sqlstr As String = "select b.studentid,c.name,b.StudStatus,b.RejectTDate1,b.RejectTDate2  from class_classinfo a join class_studentsofclass b on a.ocid=b.ocid join stud_studentinfo c on b.sid=c.sid where b.OCID='" & Re_OCID.Value & "' and b.studentid='" & Re_Studentid.Value & "'"
                Dim row As DataRow = DbAccess.GetOneRow(sqlstr)
                Me.Label_Name.Text = row("name")
                Me.Label_Stud.Text = row("studentid")
                Dim str As String
                Select Case row("StudStatus").ToString
                    Case "1"
                        Me.Label_Status.Text = "在訓"
                    Case "2"
                        str += "離訓"
                        str += "(" + row("RejectTDate1") + ")"
                        Me.Label_Status.Text = str
                    Case "3"
                        str += "退訓"
                        str += "(" + row("RejectTDate2") + ")"
                        Me.Label_Status.Text = str
                    Case "4"
                        Me.Label_Status.Text = "續訓"
                    Case "5"
                        Me.Label_Status.Text = "結訓"
                End Select
            End If

            '刪除學員問卷答案
            Select Case ProcessType.Value
                Case "clear"
                    Button2.Visible = False
                    BtnBak.Visible = True '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = True
                    next_but.Visible = True
                Case "del"
                    Dim sqlstrdel As String = "delete from  Stud_Questionary where  OCID='" & Re_OCID.Value & "' and StudID= '" & Re_Studentid.Value & "'"
                    DbAccess.ExecuteNonQuery(sqlstrdel, objconn)
                    Button2.Visible = False
                    BtnBak.Visible = True '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = True
                    next_but.Visible = True
                Case "check"              '檢視學員問卷答案
                    SOCID.Enabled = False
                    create(Re_Studentid.Value)
                    Button2.Visible = True
                    BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = False
                    next_but.Visible = False
                Case "Edit"                '修改
                    SOCID.Enabled = False
                    create(Re_Studentid.Value)
                    Button2.Visible = True
                    BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = True
                    next_but.Visible = False
                Case "Next"                '下一個
                    check_next()
                Case "Print"               '列印空白 
                    next_but.Visible = False
                    Button2.Visible = False
                    BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = False
                    'Me.RegisterStartupScript("scripprint", "<script>printDoc();history.back(1);</script>")
                    If Session("QuestionarySearchStr") = Nothing Then Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr")
                    Me.RegisterStartupScript("scripprint", "<script>printDoc();self.location.href='SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2&ocid=" & Me.Re_OCID.Value & "';</script>")
                Case "Print2"              '列印
                    create(Re_Studentid.Value)
                    next_but.Visible = False
                    Button2.Visible = False
                    BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = False
                    'Me.RegisterStartupScript("scripprint", "<script>printDoc();history.back(1);</script>")
                    If Session("QuestionarySearchStr") = Nothing Then Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr")
                    Me.RegisterStartupScript("scripprint", "<script>printDoc();self.location.href='SD_11_001.aspx?ID=" & Me.Re_ID.Value & "&ProcessType=Back2&ocid=" & Me.Re_OCID.Value & "';</script>")
                Case Else
                    Button2.Visible = True
                    BtnBak.Visible = False '清除專用因為若使用原上一頁btn會出現重覆的確認訊息
                    Button1.Visible = True
                    next_but.Visible = True
            End Select
        End If
    End Sub

    Private Sub create(ByVal StrStudID As String)
        '取得學員問卷答案
        Dim sqlstr As String
        sqlstr = "select * from Stud_Questionary where OCID='" & Re_OCID.Value & "' and StudID= '" & StrStudID & "'"
        Dim row_list As DataRow = DbAccess.GetOneRow(sqlstr)

        '答案清除 -- Start
        RadioButtonList1_1.SelectedIndex = -1
        RadioButtonList1_2.SelectedIndex = -1
        RadioButtonList1_3.SelectedIndex = -1
        RadioButtonList2_1.SelectedIndex = -1
        RadioButtonList2_2.SelectedIndex = -1
        RadioButtonList2_3.SelectedIndex = -1
        RadioButtonList2_4.SelectedIndex = -1
        RadioButtonList2_5.SelectedIndex = -1
        RadioButtonList3_1.SelectedIndex = -1
        RadioButtonList3_2.SelectedIndex = -1
        RadioButtonList3_3.SelectedIndex = -1
        RadioButtonList3_4.SelectedIndex = -1
        RadioButtonList3_5.SelectedIndex = -1
        RadioButtonList3_6.SelectedIndex = -1
        RadioButtonList3_7.SelectedIndex = -1
        RadioButtonList4_1.SelectedIndex = -1
        RadioButtonList4_2.SelectedIndex = -1
        RadioButtonList4_3.SelectedIndex = -1
        RadioButtonList4_4.SelectedIndex = -1
        RadioButtonList4_5.SelectedIndex = -1
        RadioButtonList4_6.SelectedIndex = -1
        RadioButtonList5_1.SelectedIndex = -1
        RadioButtonList5_2.SelectedIndex = -1
        RadioButtonList5_3.SelectedIndex = -1
        RadioButtonList5_4.SelectedIndex = -1
        RadioButtonList6_1.SelectedIndex = -1
        RadioButtonList6_2.SelectedIndex = -1
        RadioButtonList6_3.SelectedIndex = -1
        RadioButtonList6_4.SelectedIndex = -1
        RadioButtonList6_5.SelectedIndex = -1
        '答案清除 -- End

        '有資料
        If Not row_list Is Nothing Then
            If row_list("Q1_1").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList1_1, row_list("Q1_1").ToString)
            If row_list("Q1_2").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList1_2, row_list("Q1_2").ToString)
            If row_list("Q1_3").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList1_3, row_list("Q1_3").ToString)
            If row_list("Q2_1").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList2_1, row_list("Q2_1").ToString)
            If row_list("Q2_2").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList2_2, row_list("Q2_2").ToString)
            If row_list("Q2_3").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList2_3, row_list("Q2_3").ToString)
            If row_list("Q2_4").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList2_4, row_list("Q2_4").ToString)
            If row_list("Q2_5").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList2_5, row_list("Q2_5").ToString)
            If row_list("Q3_1").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_1, row_list("Q3_1").ToString)
            If row_list("Q3_2").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_2, row_list("Q3_2").ToString)
            If row_list("Q3_3").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_3, row_list("Q3_3").ToString)
            If row_list("Q3_4").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_4, row_list("Q3_4").ToString)
            If row_list("Q3_5").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_5, row_list("Q3_5").ToString)
            If row_list("Q3_6").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_6, row_list("Q3_6").ToString)
            If row_list("Q3_7").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList3_7, row_list("Q3_7").ToString)
            If row_list("Q4_1").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList4_1, row_list("Q4_1").ToString)
            If row_list("Q4_2").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList4_2, row_list("Q4_2").ToString)
            If row_list("Q4_3").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList4_3, row_list("Q4_3").ToString)
            If row_list("Q4_4").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList4_4, row_list("Q4_4").ToString)
            If row_list("Q4_5").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList4_5, row_list("Q4_5").ToString)
            If row_list("Q4_6").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList4_6, row_list("Q4_6").ToString)
            If row_list("Q5_1").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList5_1, row_list("Q5_1").ToString)
            If row_list("Q5_2").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList5_2, row_list("Q5_2").ToString)
            If row_list("Q5_3").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList5_3, row_list("Q5_3").ToString)
            If row_list("Q5_4").ToString <> "" Then Turbo.Common.SetListItem(RadioButtonList5_4, row_list("Q5_4").ToString)
            If Not IsDBNull(row_list("Q6_1")) Then Turbo.Common.SetListItem(RadioButtonList6_1, row_list("Q6_1"))
            If Not IsDBNull(row_list("Q6_2")) Then Turbo.Common.SetListItem(RadioButtonList6_2, row_list("Q6_2"))
            If Not IsDBNull(row_list("Q6_3")) Then Turbo.Common.SetListItem(RadioButtonList6_3, row_list("Q6_3"))
            If Not IsDBNull(row_list("Q6_4")) Then Turbo.Common.SetListItem(RadioButtonList6_4, row_list("Q6_4"))
            If Not IsDBNull(row_list("Q6_5")) Then Turbo.Common.SetListItem(RadioButtonList6_5, row_list("Q6_5"))

            Select Case ProcessType.Value
                Case "check"
                Case Else '"Edit", "Print2"
                    '顯示提示字眼
                    Dim strScript As String = ""
                    strScript = "<script language=""javascript"">" + vbCrLf
                    strScript += "alert('此學員，已經填寫!!');" + vbCrLf
                    strScript += "</script>"
                    Page.RegisterStartupScript("", strScript)
            End Select
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Fill_Questionary()
    End Sub

    Private Sub CustomValidator1_ServerValidate(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles CustomValidator1.ServerValidate
        If RadioButtonList3_4.SelectedValue = "2" Then
            args.IsValid = True '通過驗證
        Else
            args.IsValid = False
            source.errormessage = ""
            If RadioButtonList3_5.SelectedValue = "" Then source.errormessage &= "請選擇第三部分的問題五" & vbCrLf
            If RadioButtonList3_6.SelectedValue = "" Then source.errormessage &= "請選擇第三部分的問題六" & vbCrLf
            If RadioButtonList3_7.SelectedValue = "" Then source.errormessage &= "請選擇第三部分的問題七" & vbCrLf
            If source.errormessage = "" Then
                args.IsValid = True
            Else
                args.IsValid = False
            End If
        End If
        If RadioButtonList5_3.SelectedValue = "2" Then
            args.IsValid = True '通過驗證
        ElseIf RadioButtonList5_3.SelectedValue = "" Then
            args.IsValid = True '通過驗證
        ElseIf RadioButtonList5_3.SelectedValue = "1" Then
            args.IsValid = False
            source.errormessage = ""
            If RadioButtonList5_4.SelectedValue = "" Then source.errormessage &= "請選擇第五部分的問題四" & vbCrLf
            If source.errormessage = "" Then
                args.IsValid = True
            Else
                args.IsValid = False
            End If
        End If
    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        Dim strMessage As String = ""
        For Each obj As WebControls.BaseValidator In Page.Validators
            If obj.IsValid = False Then strMessage &= obj.ErrorMessage & vbCrLf
        Next
        If strMessage <> "" Then Turbo.Common.MessageBox(Page, strMessage)
    End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    'Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr")
    'Response.Redirect("SD_11_001.aspx?ProcessType=Back&ID=" & Re_ID.Value & "")
    'End Sub

    '已為此班級中最後一筆學員!!(顯示訊息)
    Private Sub check_last()
        Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr")
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "alert('已為此班級中最後一筆學員!!');" + vbCrLf
        strScript += "location.href ='SD_11_001.aspx?ProcessType=Back&ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub next_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles next_but.Click
        If SOCID.Items.Count > 0 Then
            Dim NowIndex As Integer
            Dim MaxIndex As Integer
            MaxIndex = SOCID.Items.Count - 1
            NowIndex = SOCID.SelectedIndex
            If NowIndex = MaxIndex Then
                check_last() '已為此班級中最後一筆學員!!
            Else
                SOCID.SelectedIndex = NowIndex + 1
                Re_Studentid.Value = SOCID.SelectedValue
                create(SOCID.SelectedValue)
            End If
        End If
    End Sub

    '取得下一位新增學員(若沒有才可新增，若直到最後一位，則顯示訊息)
    Private Sub check_next()
        Dim Student_Table As DataTable
        Dim rows() As DataRow
        Try
            If Session("DTable_Stuednt") Is Nothing Then
                Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');top.location.href='../../MOICA_Login.aspx';</script>")
                Me.Response.End()
                Exit Sub
            End If
            Student_Table = Session("DTable_Stuednt")
            If Student_Table.Select("studentid > '" & Re_Studentid.Value & "'").Length > 0 Then
                rows = Student_Table.Select("studentid > '" & Re_Studentid.Value & "'")
                If Not rows.Length = 0 Then
                    For i As Integer = 0 To rows.Length - 1
                        Dim dr As DataRow = rows(i) 'Stud_Questionary
                        Dim sqlstr_list As String = "select * from Stud_Questionary where OCID='" & dr("OCID") & "' and StudID= '" & dr("studentid") & "'"
                        If DbAccess.GetCount(sqlstr_list) = 0 Then '沒有資料
                            '取得下1位學員基礎資料 (學員號與班級號資料)
                            Re_Studentid.Value = dr("studentid")
                            Re_OCID.Value = dr("OCID")
                            Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr") '存取搜尋頁面條件
                            '重新呼叫頁面(重新載入)
                            TIMS.Utl_Redirect1(Me, "SD_11_001_add.aspx?ocid=" & Me.Re_OCID.Value & "&Stuedntid=" & Re_Studentid.Value & "&ID=" & Re_ID.Value & "")
                            '執行不到這一句，但還是寫離開此迴圈
                            Exit For
                        ElseIf rows.Length = 0 Or rows.Length = 1 Then
                            check_last() '已為此班級中最後一筆學員!!
                        ElseIf i = rows.Length - 1 Then
                            check_last() '已為此班級中最後一筆學員!!
                        End If
                    Next
                Else
                    check_last() '已為此班級中最後一筆學員!!
                End If
            Else
                check_last() '已為此班級中最後一筆學員!!
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        Re_Studentid.Value = SOCID.SelectedValue
        create(SOCID.SelectedValue)
    End Sub

    Function QuestionType()
        Dim TD_R3_4 As HtmlTableCell = Me.FindControl("TD_R3_4")
        Dim TD_R3_5 As HtmlTableCell = Me.FindControl("TD_R3_5")
        Dim TD_R3_6 As HtmlTableCell = Me.FindControl("TD_R3_6")
        Dim TD_R3_7 As HtmlTableCell = Me.FindControl("TD_R3_7")
        Dim TD_R6_1 As HtmlTableCell = Me.FindControl("TD_R6_1")
        Dim TD_R6_2 As HtmlTableCell = Me.FindControl("TD_R6_2")
        Dim TD_R6_3 As HtmlTableCell = Me.FindControl("TD_R6_3")
        Dim TD_R6_4 As HtmlTableCell = Me.FindControl("TD_R6_4")
        Dim TD_R6_5 As HtmlTableCell = Me.FindControl("TD_R6_5")
        Dim TD_R6 As HtmlTableCell = Me.FindControl("TD_R6")
        Dim sqlstr As String
        Dim qtype As String

        qtype = viewstate("QName")
        Select Case qtype '問卷類型  ID_Questionary
            Case "B"
                '重新設定問卷題目 (在職)
                CustomValidator1.Enabled = False
                TD_R3_4.Style("display") = "none"
                TD_R3_5.Style("display") = "none"
                TD_R3_6.Style("display") = "none"
                TD_R3_7.Style("display") = "none"
                TD_R6_1.Style("display") = "none"
                TD_R6_2.Style("display") = "none"
                TD_R6_3.Style("display") = "none"
                TD_R6_4.Style("display") = "none"
                TD_R6_5.Style("display") = "none"
                TD_R6.Style("display") = "none"
                Re_R3_4.Enabled = False
                Re_R6_1.Enabled = False
                Re_R6_2.Enabled = False
                Re_R6_3.Enabled = False
                Re_R6_4.Enabled = False
                Re_R6_5.Enabled = False
                '第一部份
                Label_R1_1.Text = "1.請問您對這次的職訓課程內容安排是否滿意？"
                Label_R1_2.Text = "2.請問您對這次的職訓課程時數安排是否滿意？"
                Label_R1_3.Text = "3.請問您對使用的上課教材與訓練設施（如工具 /材料）是否滿意？"
                '第二部份
                Label_R2_1.Text = "1.請問您是否滿意老師之專業知識？"
                Label_R2_2.Text = "2.請問您是否滿意老師之教學態度？"
                Label_R2_3.Text = "3.請問您是否滿意老師之教學方法？"
                Label_R2_4.Text = "4.請問您是否滿意老師之教材內容？"
                Label_R2_5.Text = "5.請問您是否滿意老師與學員間之互動？"
                '第三部份
                Label_R3_1.Text = "1.請問您是否滿意上課環境？"
                Label_R3_2.Text = "2.請問您對這次職訓上課地點公共設施(如消防安全及無障礙設施）是否滿意？"
                Label_R3_3.Text = "3.請問您對訓練單位行政支援（如問題解決及申訴管道）是否滿意？"
                '第四部份
                Label_R4_1.Text = "1.請問您對上課內容吸收程度如何？"
                Label_R4_2.Text = "2.您對於自己參加職訓這段期間的總體學習表現打幾分？"
                Label_R4_3.Text = "3.若用考試評估您的學習效果，您的學習效果如何？"
                Label_R4_4.Text = "4.若用交作業評估您的學習效果，您的學習效果如何？ "
                Label_R4_5.Text = "5.若用實習評估您的學習效果，您的學習效果如何？"
                Label_R4_6.Text = "6.相較於您參訓前對課程的期待，您對本次參訓結果是否滿意？"
                '第五部份
                Turbo.Common.AddClientScript(Page, "ChgFont('B');")
                Label_R5_1.Text = "1.受訓之課程內容與目前的工作內容是否相關？ "
                Label_R5_2.Text = "2.受訓所學知識技能，對目前工作或轉業有無幫助？"
                Label_R5_3.Text = "3.您對本次訓練認為最需要改進的地方為何？（單選）"
                Label_R5_4.Text = "4.您是否有意願繼續參加與工作有關之進修訓練活動？"
                '第一次
                Select Case viewstate("IsReLoaded")
                    Case ""
                        viewstate("IsReLoaded") = "N" '第一次
                    Case "N"
                        viewstate("IsReLoaded") = "Y" '第2次以後
                End Select

                If viewstate("IsReLoaded") = "N" Then '第一次設定
                    RadioButtonList5_1.Items.Clear()
                    RadioButtonList5_1.Items.Insert(0, New ListItem("是", 1))
                    RadioButtonList5_1.Items.Insert(1, New ListItem("否", 2))
                    RadioButtonList5_3.Items.Clear()
                    RadioButtonList5_3.Items.Insert(0, New ListItem("很滿意，無需改進", 1))
                    RadioButtonList5_3.Items.Insert(1, New ListItem("參訓職類不符就業市場需求", 2))
                    RadioButtonList5_3.Items.Insert(2, New ListItem("訓練設備不符產業需求", 3))
                    RadioButtonList5_3.Items.Insert(3, New ListItem("訓練時數不足", 4))
                    RadioButtonList5_3.Items.Insert(4, New ListItem("教學課程安排不當", 5))
                    RadioButtonList5_3.Items.Insert(5, New ListItem("訓練師專業及熱忱不足", 6))
                    RadioButtonList5_3.RepeatColumns = 1
                    RadioButtonList5_3.Enabled = True
                    RadioButtonList5_4.Items.Clear()
                    RadioButtonList5_4.Items.Insert(0, New ListItem("政府無補助也願意", 1))
                    RadioButtonList5_4.Items.Insert(1, New ListItem("政府提供50%以上之補助才願意", 2))
                    RadioButtonList5_4.Items.Insert(2, New ListItem("政府有補助才願意，無補助就不願意", 3))
                    RadioButtonList5_4.Items.Insert(3, New ListItem("政府提供補助也不願意", 4))
                    RadioButtonList5_4.RepeatColumns = 1
                    RadioButtonList5_4.Enabled = True
                End If
            Case Else
                '預設問卷為「A」
                Re_R5_2.Enabled = False
                Re_R5_3.Enabled = False
                Re_R5_4.Enabled = False
                TD_R3_4.Style("display") = "inline"
                TD_R3_5.Style("display") = "inline"
                TD_R3_6.Style("display") = "inline"
                TD_R3_7.Style("display") = "inline"
                TD_R6_1.Style("display") = "inline"
                TD_R6_2.Style("display") = "inline"
                TD_R6_3.Style("display") = "inline"
                TD_R6_4.Style("display") = "inline"
                TD_R6_5.Style("display") = "inline"
                TD_R6.Style("display") = "inline"
        End Select
    End Function

    Function Fill_Questionary()
        Dim objconn As SqlConnection = DbAccess.GetConnection()
        TIMS.OpenDbConn(Me, objconn)
        Dim sqlAdapter As SqlClient.SqlDataAdapter
        Dim sqldr As DataRow
        Dim dr_row As DataRow
        Dim dtTable As DataTable
        Dim State As String = ""
        'Dim Qtyp As String
        Const cst_update = "update"
        Const cst_add = "add"

        If Not Page.IsValid Then Exit Function
        Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr")
        Dim sqlstr_update = "select * from Stud_Questionary where OCID='" & Re_OCID.Value & "' AND StudID='" & Re_Studentid.Value & "' "
        sqldr = DbAccess.GetOneRow(sqlstr_update)
        If Not sqldr Is Nothing Then
            State = cst_update
            dr_row = DbAccess.GetUpdateRow(sqlstr_update, dtTable, sqlAdapter, objconn)
        Else
            State = cst_add
            dr_row = DbAccess.GetInsertRow("Stud_Questionary", dtTable, sqlAdapter, objconn)
            dr_row = dtTable.NewRow
            dr_row("OCID") = Re_OCID.Value
            dr_row("StudID") = Re_Studentid.Value
        End If
        dr_row("FillFormDate") = Now()

        Select Case viewState("QName")
            Case "A"  '問卷A
                dr_row("Q1_1") = RadioButtonList1_1.SelectedValue
                dr_row("Q1_2") = RadioButtonList1_2.SelectedValue
                dr_row("Q1_3") = RadioButtonList1_3.SelectedValue
                dr_row("Q2_1") = RadioButtonList2_1.SelectedValue
                dr_row("Q2_2") = RadioButtonList2_2.SelectedValue
                dr_row("Q2_3") = RadioButtonList2_3.SelectedValue
                dr_row("Q2_4") = RadioButtonList2_4.SelectedValue
                dr_row("Q2_5") = RadioButtonList2_5.SelectedValue
                dr_row("Q3_1") = RadioButtonList3_1.SelectedValue
                dr_row("Q3_2") = RadioButtonList3_2.SelectedValue
                dr_row("Q3_3") = RadioButtonList3_3.SelectedValue
                dr_row("Q3_4") = RadioButtonList3_4.SelectedValue
                If RadioButtonList3_5.SelectedValue <> "" And RadioButtonList3_5.Enabled = True Then
                    dr_row("Q3_5") = RadioButtonList3_5.SelectedValue
                Else
                    dr_row("Q3_5") = Convert.DBNull
                End If
                If RadioButtonList3_6.SelectedValue <> "" And RadioButtonList3_6.Enabled = True Then
                    dr_row("Q3_6") = RadioButtonList3_6.SelectedValue
                Else
                    dr_row("Q3_6") = Convert.DBNull
                End If
                If RadioButtonList3_7.SelectedValue <> "" And RadioButtonList3_7.Enabled = True Then
                    dr_row("Q3_7") = RadioButtonList3_7.SelectedValue
                Else
                    dr_row("Q3_7") = Convert.DBNull
                End If
                dr_row("Q4_1") = RadioButtonList4_1.SelectedValue
                dr_row("Q4_2") = RadioButtonList4_2.SelectedValue
                dr_row("Q4_3") = RadioButtonList4_3.SelectedValue
                dr_row("Q4_4") = RadioButtonList4_4.SelectedValue
                dr_row("Q4_5") = RadioButtonList4_5.SelectedValue
                dr_row("Q4_6") = RadioButtonList4_6.SelectedValue
                dr_row("Q5_1") = RadioButtonList5_1.SelectedValue
                If RadioButtonList5_2.SelectedValue <> "" Then
                    dr_row("Q5_2") = RadioButtonList5_2.SelectedValue
                Else
                    dr_row("Q5_2") = Convert.DBNull
                End If
                If RadioButtonList5_3.SelectedValue <> "" Then
                    dr_row("Q5_3") = RadioButtonList5_3.SelectedValue
                Else
                    dr_row("Q5_3") = Convert.DBNull
                End If
                If RadioButtonList5_4.SelectedValue <> "" And RadioButtonList5_4.Enabled = True Then
                    dr_row("Q5_4") = RadioButtonList5_4.SelectedValue
                Else
                    dr_row("Q5_4") = Convert.DBNull
                End If
                dr_row("Q6_1") = RadioButtonList6_1.SelectedValue
                dr_row("Q6_2") = RadioButtonList6_2.SelectedValue
                dr_row("Q6_3") = RadioButtonList6_3.SelectedValue
                dr_row("Q6_4") = RadioButtonList6_4.SelectedValue
                dr_row("Q6_5") = RadioButtonList6_5.SelectedValue
            Case "B"   '問卷B
                '第一部份
                dr_row("Q1_1") = RadioButtonList1_1.SelectedValue
                dr_row("Q1_2") = RadioButtonList1_2.SelectedValue
                dr_row("Q1_3") = RadioButtonList1_3.SelectedValue
                '第二部份
                dr_row("Q2_1") = RadioButtonList2_1.SelectedValue
                dr_row("Q2_2") = RadioButtonList2_2.SelectedValue
                dr_row("Q2_3") = RadioButtonList2_3.SelectedValue
                dr_row("Q2_4") = RadioButtonList2_4.SelectedValue
                dr_row("Q2_5") = RadioButtonList2_5.SelectedValue
                '第三部份
                dr_row("Q3_1") = RadioButtonList3_1.SelectedValue
                dr_row("Q3_2") = RadioButtonList3_2.SelectedValue
                dr_row("Q3_3") = RadioButtonList3_3.SelectedValue
                '第四部份
                dr_row("Q4_1") = RadioButtonList4_1.SelectedValue
                dr_row("Q4_2") = RadioButtonList4_2.SelectedValue
                dr_row("Q4_3") = RadioButtonList4_3.SelectedValue
                dr_row("Q4_4") = RadioButtonList4_4.SelectedValue
                dr_row("Q4_5") = RadioButtonList4_5.SelectedValue
                dr_row("Q4_6") = RadioButtonList4_6.SelectedValue
                '第五部份
                dr_row("Q5_1") = RadioButtonList5_1.SelectedValue
                dr_row("Q5_2") = RadioButtonList5_2.SelectedValue
                dr_row("Q5_3") = RadioButtonList5_3.SelectedValue
                dr_row("Q5_4") = RadioButtonList5_4.SelectedValue
        End Select
        dr_row("QID") = CInt(viewState("QID"))
        dr_row("ModifyAcct") = sm.UserInfo.UserID
        dr_row("ModifyDate") = Now()
        Select Case State
            Case cst_add
                dtTable.Rows.Add(dr_row)
        End Select
        sqlAdapter.Update(dtTable)
        If ProcessType.Value <> "Edit" Then '是新增或是填下一個才跑這一個
            Turbo.Common.AddClientScript(Page, "insert_next();")
        Else
            Turbo.Common.AddClientScript(Page, "BAK();")
            ''Session("QuestionarySearchStr") = Me.ViewState("QuestionarySearchStr")
            'Response.Redirect("SD_11_001.aspx?ProcessType=Back&ID=" & Re_ID.Value & "&ocid=" & Re_OCID.Value & "")
        End If
    End Function

    Function GetQType() As Boolean
        '判斷是否設定問卷類別
        '搜尋計畫問卷類型是否設定
        '若有設定 viewstate("QName")=QName ; viewstate("QID")=QID

        Dim sqlstr As String
        Dim dt As DataTable
        Dim dr As DataRow
        sqlstr = " SELECT b.QName, b.QID " & vbCrLf
        sqlstr += " FROM Plan_Questionary a " & vbCrLf
        sqlstr += " LEFT JOIN ID_Questionary b ON a.QID = b.QID " & vbCrLf
        sqlstr += " WHERE TPlanID = '" & sm.UserInfo.TPlanID & "' "
        dt = DbAccess.GetDataTable(sqlstr)

        If dt.Rows.Count = 0 Then
            '計畫未設定問卷類型
            'Dim TD_Stud As HtmlTableCell = Me.FindControl("TD_Stud")
            'StdTr.Visible = False
            'Table3.Style("display") = "none"
            'Label1.Visible = True
            'Button1.Visible = False
            'Button2.Visible = False
            'next_but.Visible = False
            Return False
        Else
            '計畫已設定問卷類型
            dr = dt.Rows(0)
            viewstate("QName") = dr("QName").ToString
            viewstate("QID") = dr("QID").ToString
            Return True
        End If
    End Function

    Private Sub BtnBak_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBak.Click
        Turbo.Common.AddClientScript(Page, "BAK();")
    End Sub
End Class