Partial Class DataGridPage
    Inherits System.Web.UI.UserControl

    Dim PageCount As Integer                 '指定總頁數
    'Dim Me.ViewState(cstSql) As String = Nothing
    'Dim Me.ViewState(cstSort) As String = Nothing
    'Dim Me.ViewState(cstPrimaryKey) As String = Nothing
    'Dim Me.ViewState(cstMyRecord) As String = Nothing
    Const cstSql As String = "Sql"
    Const cstSort As String = "Sort"
    Const cstPrimaryKey As String = "MyPk"
    Const cstMyRecord As String = "MyRecord"

    Public MyRecord As Integer                  '指定查詢的筆數                     設定
    Public MyDataGrid As DataGrid               '指定分頁的DataGrid                 設定
    Public MyDataTable As DataTable
    Public MySort As String                     'DataGrid的排序方式                 可選擇
    Public MySqlStr As String                   '查詢的SQL語法                      設定
    Public MyPrimaryKey As String               '查詢的SQL的主鍵                    設定

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁        
        If Not IsPostBack Then
            If PageCount <= 1 Then
                Me.Visible = False
            End If
        End If
    End Sub

    '第一次載入需要的參數
    Sub FirstTime()
        If Not MyDataTable Is Nothing Then
            Session("MyDataTable") = MyDataTable
        End If
        Me.ViewState(cstSql) = MySqlStr
        Me.ViewState(cstSort) = MySort
        Me.ViewState(cstPrimaryKey) = MyPrimaryKey
        Me.ViewState(cstMyRecord) = MyRecord

        If Not Me.ViewState(cstSql) Is Nothing Then
            PageCount = MyRecord \ MyDataGrid.PageSize
            If MyRecord Mod MyDataGrid.PageSize Then
                PageCount += 1
            End If
        Else
            PageCount = MyDataTable.Rows.Count \ MyDataGrid.PageSize
            If MyDataTable.Rows.Count Mod MyDataGrid.PageSize Then
                PageCount += 1
            End If
        End If

        TextBox1.Text = 1
        TextBox2.Text = PageCount

        If PageCount > 1 Then
            Me.Visible = True
        Else
            Me.Visible = False
        End If
    End Sub

    Sub GetSort()
        Me.ViewState(cstSort) = MySort
        Call create()
    End Sub

    '建立資料集
    Sub create()
        Dim dt As DataTable
        If Not Me.ViewState(cstSql) Is Nothing Then
            Dim sql As String = Me.ViewState(cstSql)

            If sql Is Nothing Then
                Common.MessageBox(Me.Page, "沒有SQL字串")
            Else
                If Me.ViewState(cstPrimaryKey) Is Nothing Then
                    dt = DbAccess.GetDataTable(sql)
                    MyDataGrid.AllowCustomPaging = False
                Else
                    sql = TIMS.Get_SQLPAGE(sql, TextBox1.Text, MyDataGrid.PageSize, Me.ViewState(cstPrimaryKey), Me.ViewState(cstSort))
                    dt = DbAccess.GetDataTable(sql)
                    MyDataGrid.AllowCustomPaging = True
                    MyDataGrid.VirtualItemCount = Me.ViewState(cstMyRecord)
                End If

                MyDataGrid.CurrentPageIndex = TextBox1.Text - 1
                MyDataGrid.DataSource = dt
                MyDataGrid.DataBind()
            End If
        Else
            dt = Session("MyDataTable")
            'MyDataGrid.VirtualItemCount = Me.ViewState(cstMyRecord)
            MyDataGrid.AllowCustomPaging = False
            MyDataGrid.CurrentPageIndex = TextBox1.Text - 1
            If Not Me.ViewState(cstSort) Is Nothing Then
                dt.DefaultView.Sort = Me.ViewState(cstSort)
            End If
            MyDataGrid.DataSource = dt
            MyDataGrid.DataBind()
        End If
    End Sub

    Private Sub ImageButton1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click, ImageButton2.Click, ImageButton3.Click, ImageButton4.Click
        If sender Is ImageButton1 Then
            TextBox1.Text = 1
        ElseIf sender Is ImageButton2 Then
            If Int(TextBox1.Text) > 1 Then
                TextBox1.Text -= 1
            End If
        ElseIf sender Is ImageButton3 Then
            If Int(TextBox1.Text) < Int(TextBox2.Text) Then
                TextBox1.Text += 1
            End If
        ElseIf sender Is ImageButton4 Then
            TextBox1.Text = TextBox2.Text
        End If

        If MyDataGrid.EditItemIndex <> -1 Then
            MyDataGrid.EditItemIndex = -1
        End If
        Call create()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Call create()
    End Sub
End Class
