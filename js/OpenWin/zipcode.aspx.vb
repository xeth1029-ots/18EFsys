Imports System.Text
'Imports System.Data.SqlClient
'Imports Turbo

Partial Class zipcode
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents PlaceHolder1 As System.Web.UI.WebControls.PlaceHolder
    Protected WithEvents Literal1 As System.Web.UI.WebControls.Literal

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region
    Private dtCity As DataTable
    Private dtZip As DataTable
    Private objConnection As SqlConnection

    Private Sub GetCityTable()
        '請修改為從你的縣市代碼表，能讀出CITY_ID, CITY_NAME這兩個固定欄位名稱的SQL查詢
        '-----------------------------------------------------------------------
        Dim strSql As String = "SELECT CTID CITY_ID,CTName CITY_NAME FROM ID_CITY"
        '-----------------------------------------------------------------------
        Dim daCity As New SqlDataAdapter(strSql, objConnection)

        dtCity = New DataTable
        daCity.Fill(dtCity)
    End Sub

    Private Sub GetZipTable(ByVal strCityId As String)
        '請修改為從你的郵遞區碼代碼表，能讀出ZIP_ID, ZIP_NAME這兩個固定欄位名稱的SQL查詢
        Dim strSql As String = ""

        '-----------------------------------------------------------------------
        If Request("local") = "Y" Then
            strSql = ""
            strSql += " SELECT a.ZipCode ZIP_ID,a.ZipName + '[' + b.Name + ']' ZIP_NAME "
            strSql += " FROM ID_ZIP a"
            strSql += " JOIN Key_location b on a.LCID=b.LCID "
            strSql += " WHERE a.CTID='" & strCityId & "'"
        Else
            strSql = "SELECT ZipCode ZIP_ID,ZipName ZIP_NAME FROM ID_ZIP  WHERE CTID='" & strCityId & "'"
        End If

        Dim daZip As New SqlDataAdapter(strSql, objConnection)

        dtZip = New DataTable
        daZip.Fill(dtZip)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objConnection = DbAccess.GetConnection

        If Not Me.IsPostBack Then
            GetCityTable()
            GenerateCityList()
        End If

        If Me.QueryCityId.Value <> "" Then
            GetZipTable(Me.QueryCityId.Value)
            GenerateZipList()
        End If
    End Sub

    Private Sub GenerateCityList()
        Dim strCode As New StringBuilder

        strCode.Append("<script language=""javascript"">" & vbCrLf)
        For Each row As DataRow In dtCity.Rows
            strCode.AppendFormat("zip_list['{0}'] = new Array();" & vbCrLf & _
                                 "zip_list['{0}']['city_name'] = '{1}';" & vbCrLf, row("CITY_ID"), row("CITY_NAME"))
        Next
        strCode.Append("</script>" & vbCrLf)

        Me.CityList.Text = strCode.ToString
    End Sub

    Private Sub GenerateZipList()
        Dim strCode As New StringBuilder

        strCode.Append("<script language=""javascript"">" & vbCrLf)
        For Each row As DataRow In dtZip.Rows
            strCode.AppendFormat("zip_list['{0}']['{1}'] = '{2}';" & vbCrLf, Me.QueryCityId.Value, row("ZIP_ID"), row("ZIP_NAME"))
        Next
        strCode.Append("</script>" & vbCrLf)

        Me.ZipList.Text = strCode.ToString
    End Sub
End Class
