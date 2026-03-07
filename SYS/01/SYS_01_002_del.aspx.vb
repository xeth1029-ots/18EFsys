Imports System.Data.SqlClient
Imports System.Data
Imports Turbo
Public Class SYS_01_002_del
    Inherits System.Web.UI.Page
    Dim objconn As SqlConnection

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        Dim PlanID = Request("PID")
        Dim RID = Request("RID")
        Dim OrgName = Request("ON")
        Dim AccountName = Request("AN")
        Dim delstr As String
        delstr = "Delete from Auth_AccRWPlan where Account='" & AccountName & "'" & _
                 " and PlanID=" & PlanID & " and RID='" & RID & "'"

        'DbAccess.ExecuteNonQuery(delstr, objconn)
        objconn.Close()

        'Response.Redirect("SYS_01_002.aspx?RIDValue=" & RID & "&OrgName=" & OrgName & "&accountname=" & AccountName)
    End Sub

End Class
