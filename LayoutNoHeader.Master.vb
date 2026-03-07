Public Class LayoutNoHeader
    Inherits System.Web.UI.MasterPage

    Public ReqAppPath As String
    Public Property sm As SessionModel

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        sm = SessionModel.Instance()

    End Sub

    ''' <summary>
    ''' 這個 Method 會在頁面內容產生完畢, 輸出實際 HTML 前被呼叫,
    ''' 可用來對最後輸出的HTML內容進行前置處理
    ''' </summary>
    Protected Sub Page_PreRender() Handles Me.PreRender
        ReqAppPath = Request.ApplicationPath
        ' 將 Session Model 中的 Message 輸出到頁面中
        Me.LastErrorMessage.Text = sm.LastErrorMessage
        Me.LastResultMessage.Text = sm.LastResultMessage
        Me.RedirectUrlAfterBlock.Text = sm.RedirectUrlAfterBlock
        Me.Lithelppdf1.Text = sm.HelpPdf1
    End Sub
End Class
