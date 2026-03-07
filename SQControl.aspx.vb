Partial Class SQControl
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        Call ReportQuery.Redirect(Me)
    End Sub

#Region "NO USE"
    'Sub z()
    ''Dim cGuid As String =   ReportQuery.GetGuid(Me)
    ''Dim cGuid As String = ""
    ''Dim strScript As String
    '    Dim sUrl As String =   ReportQuery.GetUrl(Me)
    '    Dim NewStr As String = ""
    '    Dim MyUrl As String = Request.Url.ToString
    '    Dim MyValue As String = Right(MyUrl, MyUrl.Length - MyUrl.IndexOf("?") - 1)

    '    Dim PrintNum As Integer
    '    PrintNum = 100000 * TIMS.Rnd1X() + 1

    '    If MyValue.Chars(0) <> "&" Then
    '        MyValue = "&" & MyValue
    '    End If

    '    For i As Integer = 0 To MyValue.Length - 1
    '        If AscW(MyValue.Chars(i)) > 127 Then
    '            NewStr += Me.Server.UrlEncode(MyValue.Chars(i))
    '        Else
    '            NewStr += MyValue.Chars(i)
    '        End If
    '    Next

    '    Const cst_filename As String = "filename="
    '    Const cst_RptID As String = "RptID=" '匯出Excel參數 Export=xls
    '    If MyValue.IndexOf(cst_filename) > -1 Then
    '        MyValue = Replace(MyValue, cst_filename, cst_RptID)
    '    End If

    '    Dim sRed As String = ""
    '    sRed = sUrl & "GUID=" & PrintNum & MyValue
    '    TIMS.Utl_Redirect1(Me, sRed) '正式
    'End Sub
#End Region

End Class
