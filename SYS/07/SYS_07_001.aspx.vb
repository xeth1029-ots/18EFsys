'Imports turbo
'Imports System.Data.SqlClient
'Imports System.IO
'Imports System.Xml
'Imports System.Data
'Imports System.Web.HttpServerUtility

Partial Class SYS_07_001
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End
        Dim ErrorMsg As String = ""
        ErrorMsg = ETRAIN.Main(Me)
        If ErrorMsg <> "" Then
            Common.MessageBox(Me, ErrorMsg)
        Else
            Common.MessageBox(Me, "XML檔案已建立，請到指定位置取得")
        End If
    End Sub

End Class
