Public Class QuestionSearch
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents DataGrid1 As System.Web.UI.WebControls.DataGrid
    Protected WithEvents msg As System.Web.UI.WebControls.Label
    Protected WithEvents table_F As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents Ipt_Name As System.Web.UI.HtmlControls.HtmlInputText
    Protected WithEvents search As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents Table2 As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents PageControler1 As PageControler
    Protected WithEvents send As System.Web.UI.WebControls.Button

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region
    Dim FunDr As DataRow
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在--------------------------Start
        '' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9) '☆
        '檢查Session是否存在--------------------------End
        '分頁設定
        PageControler1.PageDataGrid = DataGrid1

        '分頁設定
        If sm.UserInfo.RoleID <> 0 Then
            If sm.UserInfo.FunDt Is Nothing Then
                Common.RespWrite(Me, "<script>alert('Session過期');</script>")
                Common.RespWrite(Me, "<script>Top.location.href='../../logout.aspx';</script>")
                'Else
                '    Dim FunDt As DataTable = sm.UserInfo.FunDt
                'Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

                'If FunDrArray.Length = 0 Then
                'Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
                'Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
                'Else
                '    'FunDr = FunDrArray(0)

                '    'If FunDr("Adds") = "1" Then
                '    '    Save.Disabled = False
                '    'Else
                '    '    Save.Disabled = True
                '    'End If

                '    'If FunDr("Sech") = "1" Then
                '    '    search.Disabled = False
                '    'Else
                '    '    search.Disabled = True
                '    '    Save.Disabled = True
                '    'End If
                'End If
            End If
        End If

        If Not IsPostBack Then

            table_F.Visible = True
            PageControler1.Visible = False
            send.Visible = False

        End If

     
    End Sub

    Private Sub search_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.ServerClick
        dt_search()
    End Sub
    Function dt_search()

        Dim str As String
        Dim sql As String
        Dim dt As DataTable

        If Ipt_Name.Value <> "" Then   '搜尋條件

            str = " and Name like '%" & Ipt_Name.Value & "%' "

        End If

        sql = "select SVID, Name, case Avail when 'Y' then '啟用' else '不啟用' end as Avail from ID_Survey where 1=1 and Avail <> 'N' " & str & ""

        If TIMS.Get_SQLRecordCount(sql) = 0 Then

            msg.Text = "查無資料"
            msg.Visible = True
            table_F.Visible = True
            Table2.Visible = False
            DataGrid1.Visible = False
            PageControler1.Visible = False
        Else
            msg.Visible = False
            table_F.Visible = True
            Table2.Visible = True
            send.Visible = True
            DataGrid1.Visible = True
            PageControler1.Visible = True
            PageControler1.SqlString = sql
            PageControler1.ControlerLoad()
        End If


    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

        If e.Item.ItemType <> ListItemType.Footer And e.Item.ItemType <> ListItemType.Header Then

            e.Item.Cells(1).Text = e.Item.ItemIndex + 1 + DataGrid1.PageSize * DataGrid1.CurrentPageIndex
        End If

    End Sub

    Private Sub send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click

        If Request("Check1") <> "" Then

            Dim sql As String = ""
            Dim dr As DataRow
            Dim dt As DataTable

            sql += "select SVID, Name from ID_Survey where 1=1 and Avail <> 'N' and SVID ='" & Request("Check1") & "' "
            dr = DbAccess.GetOneRow(sql)

            Common.RespWrite(Me, "<script language=javascript>")
            Common.RespWrite(Me, "function returnNum(){")
            Common.RespWrite(Me, "window.opener.document.FDUpdate.QuesType.value = '" & dr("Name") & "' ;")
            Common.RespWrite(Me, "window.opener.document.FDUpdate.SVID.value = '" & dr("SVID") & "';")
            Common.RespWrite(Me, "window.close();")
            Common.RespWrite(Me, "}")
            Common.RespWrite(Me, "returnNum();")
            Common.RespWrite(Me, "</script>")
            Common.RespWrite(Me, "<script>window.close();</script>")
        Else
            Turbo.Common.MessageBox(Me, "請先勾選班級!")
        End If
    End Sub

End Class
