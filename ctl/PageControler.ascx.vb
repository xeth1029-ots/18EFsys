Partial Class PageControler
    Inherits System.Web.UI.UserControl

    'Session(cst_ssDataTable)
    Private _PageDataGrid As DataGrid           '控制項的DataGrid
    Private _SqlString As String                'SQL字串
    Private _PageIndex As Integer               '目前指定頁數
    Private _PrimaryKey As String               'SQL字串主鍵
    Private _Sort As String                     '排序
    Private _RecordCount As Integer             '資料筆數
    Private _PageDataTable As DataTable         'DataTable
    Private _PageDataTable2 As DataTable        'DataTable(非必要勿使用)
    Private _PageCount As Integer
    Private _StartPage As Integer
    Private _EndPage As Integer

    Const cst_ssDataTable As String = "_DataTable"
    Const cst_ssDataTable2 As String = "_DataTable2"
    Const cstSql As String = "Sql"
    Const cstSort As String = "Sort"
    Const cstPrimaryKey As String = "PrimaryKey"

    'PageControler1.PageDataTable = dt
    'PageControler1.PrimaryKey = "SeqNo"
    'PageControler1.Sort = "StudentID"
    'PageControler1.ControlerLoad() 'PageDataGrid.DataBind()

    Public Property PageDataGrid() As DataGrid
        Get
            Return _PageDataGrid
        End Get
        Set(ByVal Value As DataGrid)
            _PageDataGrid = Value
        End Set
    End Property

    Public Property SqlString() As String
        Get
            Return _SqlString
        End Get
        Set(ByVal Value As String)
            _SqlString = Value
        End Set
    End Property

    Public Property SSSDTRID() As String
        Get
            Return HidSSSDTRID.Value
        End Get
        Set(ByVal Value As String)
            '設定一次
            If HidSSSDTRID.Value <> "" AndAlso Session(HidSSSDTRID.Value) IsNot Nothing Then Session(HidSSSDTRID.Value) = Nothing
            HidSSSDTRID.Value = Value
        End Set
    End Property

    Public Property PageIndex() As Integer
        Get
            Return If(_PageIndex <= 0, 1, _PageIndex)
        End Get
        Set(ByVal Value As Integer)
            If Value <= 0 Then
                _PageIndex = 1
                NowPage.Value = "1"
            Else
                _PageIndex = Value
                NowPage.Value = Convert.ToString(Value)
            End If
        End Set
    End Property

    Public Property PrimaryKey() As String
        Get
            Return If(_PrimaryKey Is Nothing, "", _PrimaryKey)
        End Get
        Set(ByVal Value As String)
            _PrimaryKey = Value
        End Set
    End Property

    Public Property Sort() As String
        Get
            Return If(_Sort Is Nothing, "", _Sort)
        End Get
        Set(ByVal Value As String)
            _Sort = Value
        End Set
    End Property

    Public Property RecordCount() As Integer
        Get
            Return _RecordCount
        End Get
        Set(ByVal Value As Integer)
            _RecordCount = Value
        End Set
    End Property

    Public Property PageDataTable() As DataTable
        Get
            Return _PageDataTable
        End Get
        Set(ByVal Value As DataTable)
            _PageDataTable = Value
        End Set
    End Property

    Public Property PageDataTable2() As DataTable
        Get
            Return _PageDataTable2
        End Get
        Set(ByVal Value As DataTable)
            _PageDataTable2 = Value
        End Set
    End Property

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        Button1.Style("display") = "none"
    End Sub

#Region "NO USE"
    'Public Sub SqlDataCreate(ByVal sql As String, Optional ByVal DataSort As String = "", Optional ByVal Index As String = Nothing)
    '    Me.ViewState(cstSql) = sql
    '    Me.ViewState(cstSort) = DataSort
    '    If Index Is Nothing Then
    '        NowPage.Value = Convert.ToString(PageIndex)
    '    Else
    '        NowPage.Value = Index
    '    End If
    '    CreateData()
    'End Sub

    'Public Sub SqlPrimaryKeyDataCreate(ByVal sql As String, ByVal pk As String, Optional ByVal DataSort As String = "", Optional ByVal Index As String = Nothing)
    '    Me.ViewState(cstSql) = sql
    '    Me.ViewState(cstPrimaryKey) = pk
    '    Me.ViewState(cstSort) = DataSort
    '    If Index Is Nothing Then
    '        NowPage.Value = Convert.ToString(PageIndex)
    '    Else
    '        NowPage.Value = Index
    '    End If
    '    CreateData()
    'End Sub

    'Sub Utl_ErrorX1()
    '    Dim sHTTPHOST As String = TIMS.GetHTTP_HOST(Page)
    '    Dim strErrmsg As String = ""
    '    strErrmsg = ""
    '    strErrmsg &= TIMS.GetErrorMsg(Page) '取得錯誤資訊寫入
    '    strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
    '    strErrmsg &= "GetHTTP_HOST:" & vbCrLf & sHTTPHOST & vbCrLf
    '    strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
    '    Call TIMS.SendMailTest(strErrmsg)
    '    With Page
    '        .Response.Status = TIMS.cst_404NotFound
    '        .Response.StatusCode = 404
    '        .Response.End()
    '    End With
    'End Sub
#End Region

    Public Sub DataTableCreate(ByVal dt As DataTable, Optional ByVal DataSort As String = "", Optional ByVal Index As String = Nothing)
        If HidSSSDTRID.Value <> "" Then
            Session(HidSSSDTRID.Value) = dt
        Else
            Session(cst_ssDataTable) = dt
        End If
        Me.ViewState(cstSort) = DataSort
        NowPage.Value = If(Index Is Nothing, Convert.ToString(PageIndex), Index)
        Call CreateData()
    End Sub

    'PageDataGrid.DataBind()
    Public Sub ControlerLoad()
        Me.ViewState(cstSql) = SqlString
        Me.ViewState(cstPrimaryKey) = PrimaryKey
        Me.ViewState(cstSort) = Sort
        If HidSSSDTRID.Value <> "" Then
            If PageDataTable IsNot Nothing Then Session(HidSSSDTRID.Value) = PageDataTable
        Else
            If PageDataTable IsNot Nothing Then Session(cst_ssDataTable) = PageDataTable
        End If
        If PageDataTable2 IsNot Nothing Then Session(cst_ssDataTable2) = PageDataTable2

        NowPage.Value = Convert.ToString(PageIndex)

        Call CreateData()
    End Sub

    Public Sub ChangeSort()
        Me.ViewState(cstSort) = Sort
        CreateData()
    End Sub

    Public Sub ChangeSort(ByVal MySort As String)
        Me.ViewState(cstSort) = MySort
        CreateData()
    End Sub

    Public Sub CreateData()
        '備用 DataTable2
        If Session(cst_ssDataTable2) IsNot Nothing Then _PageDataTable2 = CType(Session(cst_ssDataTable2), DataTable)

        If Me.ViewState(cstSql) IsNot Nothing Then
            If Convert.ToString(Me.ViewState(cstPrimaryKey)) = "" Then
                _SqlString = Convert.ToString(Me.ViewState(cstSql))
                PageDataGrid.AllowCustomPaging = False

                _PageDataTable = DbAccess.GetDataTable(_SqlString)
                If Convert.ToString(Me.ViewState(cstSort)) <> "" Then
                    _PageDataTable.DefaultView.Sort = Convert.ToString(Me.ViewState(cstSort))
                End If
                RecordCount = TIMS.Get_SQLRecordCount(Convert.ToString(Me.ViewState(cstSql)))
            Else
                _SqlString = TIMS.Get_SQLPAGE(Convert.ToString(Me.ViewState(cstSql)), CInt(NowPage.Value), PageDataGrid.PageSize, Me.ViewState(cstPrimaryKey).ToString, Me.ViewState(cstSort).ToString)
                PageDataGrid.AllowCustomPaging = True

                _PageDataTable = DbAccess.GetDataTable(_SqlString)

                RecordCount = TIMS.Get_SQLRecordCount(Me.ViewState(cstSql).ToString)
                PageDataGrid.VirtualItemCount = RecordCount
            End If
        Else
            'Try
            '    dt = CType(Session(cst_ssDataTable), DataTable)
            '    If Convert.ToString(Me.ViewState(cstSort)) <> "" Then
            '        dt.DefaultView.Sort = Convert.ToString(Me.ViewState(cstSort))
            '    End If
            'Catch ex As Exception
            '    Exit Sub
            'End Try
            _PageDataTable = If(HidSSSDTRID.Value <> "", CType(Session(HidSSSDTRID.Value), DataTable), CType(Session(cst_ssDataTable), DataTable))

            If _PageDataTable IsNot Nothing Then
                If Convert.ToString(Me.ViewState(cstSort)) <> "" Then
                    _PageDataTable.DefaultView.Sort = Convert.ToString(Me.ViewState(cstSort))
                End If
            End If
            RecordCount = 0
            If _PageDataTable IsNot Nothing Then RecordCount = _PageDataTable.Rows.Count

            'PageDataGrid.DataGrid is Nothing 
            '分頁設定 Start
            'PageControler1.PageDataGrid = DataGrid1
            '分頁設定 End
            PageDataGrid.AllowCustomPaging = False
        End If

        Me.Visible = If(RecordCount = 0, False, True)
        If RecordCount <> 0 Then
            'Me.Visible = True
            '無效的 CurrentPageIndex 值。必須是 >= 0 且 < PageCount。
            If NowPage.Value <> "" Then NowPage.Value = Val(NowPage.Value)
            If (CInt(NowPage.Value) - 1) >= 0 AndAlso (CInt(NowPage.Value) - 1) < PageDataGrid.PageCount Then
                PageDataGrid.CurrentPageIndex = (CInt(NowPage.Value) - 1)
            Else
                NowPage.Value = "1"
                PageDataGrid.CurrentPageIndex = 0
                If (CInt(NowPage.Value) - 1) >= 0 Then
                    If (CInt(NowPage.Value) - 1) < PageDataGrid.PageCount Then
                        '合理
                        PageDataGrid.CurrentPageIndex = (CInt(NowPage.Value) - 1)
                        NowPage.Value = CStr(PageDataGrid.CurrentPageIndex + 1)
                    Else
                        '不合理
                        If Val(PageDataGrid.PageCount) <> 0 Then
                            PageDataGrid.CurrentPageIndex = Val(PageDataGrid.PageCount)
                            NowPage.Value = CStr(PageDataGrid.CurrentPageIndex + 1)
                        End If
                    End If
                End If
            End If

            Dim TryNPV1 As Boolean = False
            'PageDataGrid.DataSource = dt
            PageDataGrid.DataSource = _PageDataTable.DefaultView
            Try
                'CurrentPageIndex 值。必須是 >= 0 且 < PageCount。
                PageDataGrid.DataBind()
            Catch ex As Exception
                Dim str_err_msg1 As String = ex.ToString
                Dim flag_write_1 As Boolean = CHK_ERR_MSG(str_err_msg1)
                If flag_write_1 Then TIMS.WriteTraceLog(Page, ex, str_err_msg1)
                '錯誤動作呼叫執行。
                'Call TIMS.sUtl_ErrorAction(Page, Me, _PageDataTable, Convert.ToString(ViewState(cstSql)), ex)
                TryNPV1 = True
                'Throw ex
            End Try

            If TryNPV1 Then
                Try
                    NowPage.Value = "1"
                    PageDataGrid.CurrentPageIndex = 0
                    PageDataGrid.DataBind()
                Catch ex As Exception
                    Dim str_err_msg1 As String = ex.ToString
                    Dim flag_write_1 As Boolean = CHK_ERR_MSG(str_err_msg1)
                    If flag_write_1 Then TIMS.WriteTraceLog(Page, ex, str_err_msg1)
                    '錯誤動作呼叫執行。
                    'Call TIMS.sUtl_ErrorAction(Page, Me, _PageDataTable, Convert.ToString(ViewState(cstSql)), ex)
                    Me.ViewState(cstSql) = Nothing '資訊雙重異常，清除資訊。
                    Session(cst_ssDataTable) = Nothing '資訊雙重異常，清除資訊。
                    If HidSSSDTRID.Value <> "" Then Session(HidSSSDTRID.Value) = Nothing '資訊雙重異常，清除資訊。
                    'Me.Visible = False
                    'Common.RespWrite(Me, "跳頁功能產生問題，請重新查詢有效資料!")
                    'Common.RespWrite(Me, Common.GetJsString(ex.ToString))

                    If TIMS.sUtl_ChkTest() Then
                        Dim exStr1 As String = TIMS.GetResponseWrite(ex.ToString)
                        Common.RespWrite(Page, TIMS.sUtl_AntiXss(exStr1))
                        Response.End()
                    End If
                    'Call Utl_ErrorX1()

                    Dim altMsg As String = ""
                    altMsg = "alert('跳頁功能產生問題，請重新查詢有效資料');return false;"
                    PageIndexFlag.Text = ""
                    FirstButton.Attributes("onclick") = altMsg
                    PreButton.Attributes("onclick") = altMsg
                    NextButton.Attributes("onclick") = altMsg
                    LastButton.Attributes("onclick") = altMsg
                    PrePreButton.Attributes("onclick") = altMsg
                    NextNextButton.Attributes("onclick") = altMsg

                    'Common.RespWrite(Me, "<script>alert('" & Common.GetJsString(ex.ToString) & "');</script>")
                    'Common.RespWrite(Me, "<script>" & altMsg & "</script>")
                    'Response.End()
                    Exit Sub
                    'Throw ex
                End Try

            End If

            _PageCount = RecordCount \ PageDataGrid.PageSize
            If RecordCount Mod PageDataGrid.PageSize <> 0 Then _PageCount += 1
            PageCountLabel.Text = _PageCount.ToString

            PageIndexFlag.Text = ""
            '處裡頁數範圍
            _StartPage = 1
            If CInt(NowPage.Value) - 5 > 1 Then
                _StartPage = CInt(NowPage.Value) - 5
                If CInt(NowPage.Value) >= _PageCount - 5 Then
                    '表示接近最後五頁
                    _StartPage = If(_PageCount - 9 > 1, _PageCount - 9, 1)
                End If
            End If

            _EndPage = _PageCount
            If CInt(NowPage.Value) + 5 < _PageCount Then
                _EndPage = CInt(NowPage.Value) + 5
                If CInt(NowPage.Value) <= 5 Then
                    _EndPage = If(_PageCount >= 10, 10, _PageCount)
                End If
            End If

            If CInt(NowPage.Value) = 1 Then
                FirstButton.Attributes("onclick") = "alert('已經到第一頁');return false;"
                PreButton.Attributes("onclick") = "alert('已經到第一頁');return false;"
            Else
                FirstButton.Attributes("onclick") = String.Concat("ChangePage('" & NowPage.ClientID & "',1,'" & Button1.ClientID & "');return false;")
                PreButton.Attributes("onclick") = String.Concat("AddPage(-1,'" & NowPage.ClientID & "','" & Button1.ClientID & "');return false;")
            End If

            If CInt(NowPage.Value) = _PageCount Then
                NextButton.Attributes("onclick") = "alert('已經到最後一頁');return false;"
                LastButton.Attributes("onclick") = "alert('已經到最後一頁');return false;"
            Else
                LastButton.Attributes("onclick") = String.Concat("ChangePage('", NowPage.ClientID, "','", _PageCount, "','", Button1.ClientID, "');return false;")
                NextButton.Attributes("onclick") = String.Concat("AddPage(1,'", NowPage.ClientID, "','", Button1.ClientID, "');return false;") '
            End If

            If CInt(NowPage.Value) - 10 > 1 Then
                PrePreButton.Attributes("onclick") = String.Concat("LastPage(", CInt(NowPage.Value) - 10, ",'", NowPage.ClientID, "','", Button1.ClientID, "');return false;")
            Else
                If CInt(NowPage.Value) = 1 Then
                    PrePreButton.Attributes("onclick") = "alert('已經到第一頁');return false;"
                Else
                    PrePreButton.Attributes("onclick") = String.Concat("LastPage(1,'", NowPage.ClientID, "','", Button1.ClientID, "');return false;")
                End If
            End If

            If CInt(NowPage.Value) + 10 < _PageCount Then
                NextNextButton.Attributes("onclick") = String.Concat("LastPage(", CInt(NowPage.Value) + 10, ",'", NowPage.ClientID, "','", Button1.ClientID, "');return false;")
            Else
                If CInt(NowPage.Value) = _PageCount Then
                    NextNextButton.Attributes("onclick") = "alert('已經到最後一頁');return false;"
                Else
                    NextNextButton.Attributes("onclick") = String.Concat("LastPage(", _PageCount, ",'", NowPage.ClientID, "','", Button1.ClientID, "');return false;")
                End If
            End If

            For i As Integer = _StartPage To _EndPage
                Dim SpanStyle1 As String
                Dim SpanStyle2 As String
                SpanStyle1 = String.Concat("<span class=""PageLinkSpan"" title=""", i, """ onclick=""ChangePage('", NowPage.ClientID, "',", i, ",'", Button1.ClientID, "')"">", i, "</span>")
                SpanStyle2 = String.Concat("<span class=""OverLinkSpan"" title=""", i, """>", i, "</span>")
                If PageIndexFlag.Text = "" Then
                    If CInt(NowPage.Value) = i Then
                        PageIndexFlag.Text = SpanStyle2
                    Else
                        PageIndexFlag.Text = SpanStyle1
                    End If
                Else
                    If CInt(NowPage.Value) = i Then
                        PageIndexFlag.Text += String.Concat("&nbsp;&nbsp;", SpanStyle2)
                    Else
                        PageIndexFlag.Text += String.Concat("&nbsp;&nbsp;", SpanStyle1)
                    End If
                End If
            Next
        End If
    End Sub

    Function CHK_ERR_MSG(ByVal str_err_msg As String) As Boolean
        Dim rst As Boolean = True '若為true:要寫log /false:不寫log 
        Const cst_err_msg1a As String = "System.IndexOutOfRangeException"
        Const cst_err_msg1b As String = "找不到資料行"
        If str_err_msg.IndexOf(cst_err_msg1a) > -1 AndAlso str_err_msg.IndexOf(cst_err_msg1b) > -1 Then
            rst = False '常見錯誤，不寫log
        End If
        Return rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            CreateData()
        Catch ex As Exception
            Dim str_err_msg1 As String = ex.ToString
            Dim flag_write_1 As Boolean = CHK_ERR_MSG(str_err_msg1)
            If flag_write_1 Then TIMS.WriteTraceLog(Page, ex, str_err_msg1)
            '資料有誤
            'Me.Visible = False
            'Common.RespWrite(Me, "跳頁功能產生問題，請重新查詢有效資料!")
            Response.End()
            ''錯誤動作呼叫執行。
            'Call TIMS.sUtl_ErrorAction(Page, Me, _PageDataTable, Convert.ToString(ViewState(cstSql)), ex)
        End Try
    End Sub

End Class
