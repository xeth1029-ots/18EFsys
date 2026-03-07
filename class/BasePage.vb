Imports System.IO

''' <summary>
''' 所有頁面的基底類
''' </summary>
Public Class BasePage
    Inherits System.Web.UI.Page

    Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    ''' <summary>
    ''' 當前的 SessionModel
    ''' </summary>
    Protected sm As SessionModel = Nothing

    Protected Overrides Sub OnPreInit(e As EventArgs)
        sm = SessionModel.Instance()

    End Sub

    ''' <summary>
    ''' Msg_helppdf1
    ''' </summary>
    ''' <param name="s_ReqID"></param>
    ''' <param name="sm"></param>
    Sub Utl_SETHELPPDF(ByRef s_ReqID As String, ByRef sm As SessionModel)
        If s_ReqID Is Nothing Then Return
        '非NOTHING 有值時才執行
        Dim auditInfo As AuditLogInfo = sm.AuditInfo
        sm.HelpPdf1 = auditInfo.GetHelpPdf1(s_ReqID)
        'localhost:12986/Doc/HELP/TC_01_001_28.pdf?r=3064781
        If sm.HelpPdf1 IsNot Nothing AndAlso sm.HelpPdf1.Length > 1 Then
            Dim MyPathFiles1 As String = Server.MapPath($"~/Doc/HELP/{sm.HelpPdf1}")
            If Not IO.File.Exists(MyPathFiles1) Then sm.HelpPdf1 = ""
        End If

        'Const cst_Hid_helppdf1 As String = "Hid_helppdf1"
        'If Me.FindControl(cst_Hid_helppdf1) IsNot Nothing Then
        '    Dim Hid_helppdf1 As HiddenField = CType(Me.Form.FindControl(cst_Hid_helppdf1), HiddenField)
        '    Hid_helppdf1.Value = If(sm.HelpPdf1 <> "", sm.HelpPdf1, "")
        'End If
        'Const cst_Hid_helppdf1 As String = "Hid_helppdf1"
        'Dim hf1 As HiddenField = New HiddenField()
        'hf1.ID = cst_Hid_helppdf1
        ''hf1.Value = If(sm.HelpPdf1 <> "", sm.HelpPdf1, "")
        'If Me.Form.FindControl(cst_Hid_helppdf1) Is Nothing Then Me.Form.Controls.Add(hf1)
        'hf1.Value = If(sm.HelpPdf1 <> "", sm.HelpPdf1, "")
    End Sub

    Protected Overrides Sub OnInit(e As EventArgs)
        Dim funcPath As String = Request.Path
        Dim s_ReqID As String = Request("ID")
        Dim userId As String = "ANOYMOUS"
        If Not IsNothing(sm.UserInfo) Then userId = sm.UserInfo.UserID

        Utl_SETHELPPDF(s_ReqID, sm)

        logger.Info($"Page Request({userId}): {funcPath}")

        ' 功能頁面的存取記錄寫入DB ' 注意: ' 要更新 auditInfo 的 Property, 須先將整個 AuditInfo 取出 ' 更新完所有 Property 之後再整個設回去, 才會正確
        Dim auditInfo As AuditLogInfo = sm.AuditInfo

        auditInfo.LastFuncPath = funcPath
        auditInfo.LastAccessTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")

        sm.AuditInfo = auditInfo
        sm.AuditInfo.WriteAccessLog()

    End Sub

    ''' <summary>PreRender 事件-在 Control 物件載入之後但在呈現之前發生。</summary>
    ''' <param name="e"></param>
    Protected Overrides Sub OnPreRender(ByVal e As EventArgs)
        MyBase.OnPreRender(e)

        ' ~/index.aspx 不注入動態內容
        If Request.Path.IndexOf("/index.aspx", StringComparison.OrdinalIgnoreCase) > -1 Then Return ' Exit Sub
        ' ~/index.aspx 不注入動態內容
        If Request.Path.IndexOf("/index", StringComparison.OrdinalIgnoreCase) > -1 Then Return ' Exit Sub
        ' 動態加入 LastResultMessage,LastErrorMessage 輸出用的 div
        Dim lastResult As New HtmlGenericControl("div") With {.ID = "Msg_LastResultMessage", .InnerHtml = sm.LastResultMessage}
        lastResult.Attributes.CssStyle.Add("display", "none")
        Dim lastError As New HtmlGenericControl("div") With {.ID = "Msg_LastErrorMessage", .InnerHtml = sm.LastErrorMessage}
        lastError.Attributes.CssStyle.Add("display", "none")
        Dim redirUrl As New HtmlGenericControl("div") With {.ID = "Msg_RedirectUrlAfterBlock", .InnerHtml = sm.RedirectUrlAfterBlock}
        redirUrl.Attributes.CssStyle.Add("display", "none")
        Dim helppdf1 As New HtmlGenericControl("div") With {.ID = "Msg_helppdf1", .InnerHtml = sm.HelpPdf1}
        helppdf1.Attributes.CssStyle.Add("display", "none")

        If IsNothing(Form) Then Return ' Exit Sub

        '登入計畫為28 才顯示，其它一律不顯示 TRPlanPoint28 Visible  = False '排除54 搜尋 
        Const cst_TRPlanPoint28 As String = "TRPlanPoint28" '計畫別 (產業人才投資計畫/提升勞工自主學習計畫)
        '登入計畫為28 才顯示，其它一律不顯示 tr_AppStage_TP28 Visible  = False '排除54 搜尋 
        Const cst_tr_AppStage_TP28 As String = "tr_AppStage_TP28" '(申請階段)
        If Form.FindControl(cst_TRPlanPoint28) IsNot Nothing Then
            CType(Me.Form.FindControl(cst_TRPlanPoint28), HtmlTableRow).Visible = If(sm.UserInfo.TPlanID = TIMS.Cst_TPlanID28, True, False)
        End If
        If Form.FindControl(cst_tr_AppStage_TP28) IsNot Nothing Then
            CType(Me.Form.FindControl(cst_tr_AppStage_TP28), HtmlTableRow).Visible = If(sm.UserInfo.TPlanID = TIMS.Cst_TPlanID28, True, False)
        End If
        Const cst_tr_ddl_INQUIRY_S As String = "tr_ddl_INQUIRY_S" '(查詢原因)
        If Form.FindControl(cst_tr_ddl_INQUIRY_S) IsNot Nothing AndAlso Not TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then
            CType(Me.Form.FindControl(cst_tr_ddl_INQUIRY_S), HtmlTableRow).Visible = False
        End If
        Form.Controls.Add(lastResult)
        Form.Controls.Add(lastError)
        Form.Controls.Add(redirUrl)
        Form.Controls.Add(helppdf1)

        ' 因為 global-iframe.js 中會去 override alert() 
        ' 若在前端 page.load() 再動態載入, 
        ' 對 RegisterStratScripts 而言, 會失去效用
        ' 要將 /Scripts/global-iframe.js 內容直接注入頁面中
        Dim jsFile As String = "~/Scripts/global-iframe.js"
        Dim objReader As New StreamReader(Server.MapPath(jsFile))
        Dim scriptTag As New HtmlGenericControl("script")
        scriptTag.Attributes.Add("type", "text/javascript")
        scriptTag.InnerHtml = $"{vbCrLf}/* {jsFile}*/{vbCrLf}{objReader.ReadToEnd()}"
        objReader.Close()
        objReader = Nothing

        Me.Form.Controls.Add(scriptTag)
    End Sub

    ''' <summary>「HiQPdf」的SerialNumber</summary>
    ''' <returns></returns>
    Public ReadOnly Property HiQPdf_SerialNumber As String
        Get
            Return "/7eWrq+b-mbOWnY2e-jYbOz9HP-387fy9/G-yM7fzM7R-zs3RxsbG-xg==" '「HiQPdf」的SerialNumber
        End Get
    End Property

End Class
