Imports System.IO

''' <summary>
''' 產生圖型驗證碼
''' </summary>
Partial Class ValidateCode
    Inherits System.Web.UI.Page

    ''' <summary>
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        'Dim sm As SessionModel = SessionModel.Instance()
        'Dim flag_NG3 As Boolean = False
        'If Not flag_NG3 AndAlso Session(TIMS.cst_MOICA_Login) Is Nothing Then flag_NG3 = True
        'If Not flag_NG3 AndAlso Not Session(TIMS.cst_MOICA_Login).ToString().Equals("xxx") Then flag_NG3 = True
        'If flag_NG3 Then
        '    Response.StatusCode = 404
        '    Response.End()
        'End If
        Dim vc As Turbo.Commons.ValidateCode = New Turbo.Commons.ValidateCode()
        Dim vCode As String = $"{Now.ToString("fff") Mod 10}{Now.ToString("fff")}"
        If Request("Audio") = "Y" Then
            '輸出驗證碼的語音內容
            vCode = SessionModel.Instance().LoginValidateCode  '從 SessionModel 中取得已儲存的 vCode
            If IsNothing(vCode) Then
                'Session 中找不到 vCode, 回應 404 Not Found
                If (Response.StatusCode <> 404) Then Response.StatusCode = 404
                Response.End()
                Return
            End If

            Dim audioPath As String = Server.MapPath("~/Content/audio/")
            Using stream As MemoryStream = vc.CreateValidateAudio(vCode, audioPath)
                '需要輸出合成的音檔 要修改HTTP頭 
                Response.ClearContent()
                Response.ContentType = "audio/wav"

                '告訴 browser 停用 Cache
                Response.Headers.Add("Cache-Control", "no-cache, no-store")
                Response.Cache.SetExpires(DateTime.UtcNow.AddYears(-1))
                Response.BinaryWrite(stream.ToArray())
                If (Response.IsClientConnected) Then Response.Flush()
            End Using

        Else
            Dim flag_NG3 As Boolean = ($"{Request("rand")}".Length = 0)
            If flag_NG3 Then
                If (Response.StatusCode <> 404) Then Response.StatusCode = 404
                Response.End()
                Return
            End If

            '輸出驗證碼圖型
            vCode = vc.CreateValidateCode(4)
            Using stream As MemoryStream = vc.CreateValidateGraphic(vCode)
                SessionModel.Instance().LoginValidateCode = vCode  '將 ValidateCode 保存在 Session 中
                '需要輸出圖型信息 要修改HTTP頭 
                Response.ClearContent()
                Response.ContentType = "image/jpeg"
                'Response.Headers.Add("X-Content-Type-Options", "nosniff")

                '告訴 browser 停用 Cache
                Response.Headers.Add("Cache-Control", "no-cache, no-store")
                Response.Cache.SetExpires(DateTime.UtcNow.AddYears(-1))
                Response.BinaryWrite(stream.ToArray())
                If (Response.IsClientConnected) Then Response.Flush()
            End Using
        End If
    End Sub
End Class