Partial Class SYS_03_007
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

        If Not Me.IsPostBack Then
            'Hid_guid1.Value = TIMS.GetGUID()
            Me.SDate.Text = ""
            Me.EDate.Text = ""
        End If
        If Hid_guid1.Value = "" Then Hid_guid1.Value = TIMS.GetGUID()

        Button3.Attributes("onclick") = "javascript:return search()"
    End Sub

    'Function checkjob(ByVal idn As String, ByRef iCountex As Integer) As String
    '    Dim rutstr As String = ""
    '    Dim dr As DataTable
    '    Dim dr2 As DataRow
    '    Try
    '        Dim sql1 As String = ""
    '        sql1 = "SELECT * FROM temp93_2 where upper(IDNO) = '" & UCase(idn) & "' and MDate >= " & TIMS.to_date(Me.SDate.Text) & " and MDate <= " & TIMS.to_date(Me.SDate.Text) & " order by MDate desc"
    '        dr = DbAccess.GetDataTable(sql1, objconn)

    '        sql1 = "SELECT * FROM temp93_2 a left join Bus_BasicData b on a.ACTNO=b.ubno where upper(a.IDNO) = '" & UCase(idn) & "' and MDate >= " & TIMS.to_date(Me.SDate.Text) & " and MDate <= " & TIMS.to_date(Me.SDate.Text) & " and a.ACTNO='" & dr.DefaultView(0)(4).ToString & "' and a.ChangeMode = '2'"
    '        dr2 = DbAccess.GetOneRow(sql1, objconn)
    '        If dr.DefaultView(0)(5) = "4" Then
    '            If dr2 Is Nothing Then
    '                rutstr = "加保"
    '            Else
    '                rutstr = "退保"
    '            End If
    '        Else
    '            rutstr = "退保"
    '        End If
    '    Catch ex As Exception
    '        iCountex += 1
    '    End Try
    '    Return rutstr
    'End Function

    Sub createdata()
        Session(Hid_guid1.Value) = Nothing
        Dim dt As DataTable = Nothing
        Dim sql As String
        sql = "select Distinct a.Name,a.FM,a.IDNO from temp93 a,temp93_2 b where a.IDNO = b.IDNO and a.Name is not null"
        sql = "SELECT DISTINCT A.NAME,A.IDNO FROM STUD_BLIGATEDATA_TEMP A WHERE 1<>1"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim changemode As New DataColumn("changeMode", System.Type.GetType("System.String"))
        dt.Columns.Add(changemode)

        Dim iCountex As Integer = 0
        '想辦法更新dt並把changemode這個cloumns加入資料
        For i As Integer = 0 To dt.Rows.Count - 1
            'dt.Rows(i).BeginEdit()
            'dt.Rows(i).Item("changeMode") = checkjob(dt.Rows(i).Item("IDNO").ToString, iCountex)
            'dt.Rows(i).EndEdit()
        Next
        Common.RespWrite(Me, "錯誤資料筆數:" & CStr(iCountex)) '.ToString
        Common.RespWrite(Me, "共處理" & dt.Rows.Count.ToString & "筆資料")
        DataGrid1.PagerStyle.Mode = PagerMode.NumericPages
        DataGrid1.DataSource = dt
        Session(Hid_guid1.Value) = dt
        DataGridtable.Style.Item("display") = "none"
        DataGrid1.Enabled = False
        DataGrid1.DataBind()
    End Sub

    Sub binddata()
        If Session(Hid_guid1.Value) Is Nothing Then Exit Sub
        Dim dt As DataTable = Session(Hid_guid1.Value)

        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Sub DataGrid1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As DataGridPageChangedEventArgs)
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        binddata()
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    'oConn = DbAccess.GetConnection()
    '    createdata()
    'End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call EXP_XLS1()
    End Sub

    Sub EXP_XLS1()
        'oConn = DbAccess.GetConnection()
        createdata()
        If Session(Hid_guid1.Value) Is Nothing Then Exit Sub
        Dim dt As DataTable = Session(Hid_guid1.Value)
        If dt Is Nothing Then Exit Sub
        If dt.Rows.Count = 0 Then Exit Sub

        Dim MyFilePath As String
        Dim sr As System.IO.FileStream
        Dim srw As System.IO.StreamWriter
        Dim sql As String = Me.ViewState("ExportSqlStr")
        MyFilePath = Server.MapPath("temp.xls")
        If IO.File.Exists(MyFilePath) = False Then
            Try
                sr = IO.File.Create(Server.MapPath("temp.xls"))
                sr.Close()
            Catch ex As Exception
                Common.MessageBox(Me, "資料寫入錯誤，可能的原因是因為Temp的資料夾尚未設定ASP.NET的權限")
                Exit Sub
            End Try
        End If
        '建立輸出文字
        Dim ExportStr As String = ""
        '建立表頭
        For Each col As Data.DataColumn In dt.Columns
            ExportStr += Replace(col.ColumnName, vbTab, "") & vbTab
        Next
        ExportStr += vbCrLf
        '建立資料面
        For Each dr As DataRow In dt.Rows
            For i As Integer = 0 To dt.Columns.Count - 1
                Dim OneRow As String = ""
                OneRow = Replace(Replace(dr(i).ToString, vbTab, ""), vbCrLf, "")
                ExportStr += OneRow & IIf(dt.Columns(i).DataType.Name = "String" And IsNumeric(dr(i).ToString), Chr(128), "") & vbTab
            Next
            ExportStr += vbCrLf
        Next
        'srw = New IO.StreamWriter(MyFilePath, False, System.Text.Encoding.Default)
        srw = New IO.StreamWriter(MyFilePath, False, System.Text.Encoding.Unicode)
        srw.WriteLine(ExportStr)
        srw.Close()

        '將新建立的excel存入記憶體下載-----   Start
        Dim strErrmsg As String = ""
        strErrmsg = ""
        Try
            Dim fr As New System.IO.FileStream(MyFilePath, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte
            fr.Read(buf, 0, fr.Length)
            fr.Close()
            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=" & Now.ToString)
            'Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
            Response.ContentType = "Application/vnd.ms-Excel"
            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
            Response.BinaryWrite(buf)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "無法存取該檔案!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
        Finally
            '刪除Temp中的資料
            Call TIMS.MyFileDelete(MyFilePath)
            'If MyFile.Exists(MyFilePath) Then MyFile.Delete(MyFilePath)
            If strErrmsg = "" Then Response.End()
        End Try
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
        End If
        '將新建立的excel存入記憶體下載-----   End

        'Dim fr As New System.IO.FileStream(MyFilePath, IO.FileMode.Open)
        'Dim br As New System.IO.BinaryReader(fr)
        'Dim buf(fr.Length) As Byte
        'fr.Read(buf, 0, fr.Length)
        'fr.Close()
        'If MyFile.Exists(MyFilePath) Then
        '    MyFile.Delete(MyFilePath)
        'End If
        'Response.Clear()
        'Response.ClearHeaders()
        'Response.Buffer = True
        'Response.AddHeader("content-disposition", "attachment;filename=" & Now.ToString)
        ''HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8)
        ''將編碼改成UTF8，可以下載中文檔名
        'Response.ContentType = "Application/vnd.ms-Excel"
        'Response.BinaryWrite(buf)
        'Response.End()

    End Sub
End Class
