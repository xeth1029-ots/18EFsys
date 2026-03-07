Partial Class CheckZip
    Inherits AuthBasePage

    Dim strCityName As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Function Check_ZipCodeNumber(ByVal ZipCode As String) As Boolean
        Dim rst As Boolean = False
        hidValue.Value = TIMS.ClearSQM(hidValue.Value)
        If hidValue.Value.Length > 10 Then Return rst
        If hidValue.Value = "" Then Return rst
        If Not TIMS.IsNumeric1(hidValue.Value) Then Return rst
        'hidValue.Value = TIMS.ClearSQM(hidValue.Value)
        'Dim dr As DataRow
        'Dim dt As DataTable
        Dim sql As String = ""
        sql &= " SELECT concat('(',a.zipcode,')', CASE WHEN b.ctname=a.zipname THEN a.zipname ELSE concat(b.ctname,a.zipname) END) cityname,a.zipcode" & vbCrLf
        sql &= " FROM id_zip a JOIN id_city b ON b.ctid = a.ctid WHERE a.zipcode != '999'" & vbCrLf
        sql &= String.Format(" AND a.zipcode ={0} ", hidValue.Value) & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Return rst
        rst = True
        strCityName = $"{dt.Rows(0)("cityname")}"
        Return rst
    End Function

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        strCityName = ""
        If Page.IsPostBack Then
            hidCityID.Value = TIMS.ClearSQM(hidCityID.Value)
            hidZipID.Value = TIMS.ClearSQM(hidZipID.Value)
            hidValue.Value = TIMS.ClearSQM(hidValue.Value)
            If hidCityID.Value <> "" AndAlso hidZipID.Value <> "" AndAlso hidValue.Value <> "" Then
                Dim flag_CheckOK As Boolean = Check_ZipCodeNumber(hidValue.Value)
                If flag_CheckOK Then
                    Dim script1 As String = $"window.parent.document.getElementById('{hidCityID.Value}').value='{strCityName}';"
                    Common.RespWrite(Me, String.Format("<script>{0}</script>", script1))
                Else
                    Dim script1 As String = $"window.parent.document.getElementById('{hidCityID.Value}').value='';"
                    Dim msg1 As String = $"alert('查無 {hidValue.Value} 郵遞區號!');"
                    Common.RespWrite(Me, $"<script>{script1}{msg1}</script>")
                End If
            End If
        End If
    End Sub
End Class