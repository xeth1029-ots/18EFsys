Imports System.Web.Routing

''' <summary>
''' 系統路由組態
''' </summary>
Public Class RouteConfig

    Private Shared logger As ILog = LogManager.GetLogger(GetType(RouteConfig))

    ''' <summary>
    ''' 路由組態設定
    ''' <para>應在 Global.asax 中被呼叫以啟用</para>
    ''' </summary>
    ''' <param name="routes"></param>
    Public Shared Sub RegisterRoutes(ByVal routes As RouteCollection)

        routes.MapPageRoute("Common", "Common/{func}", "~/Common/{func}.aspx")

        Const cst_routeUrl1 As String = "{func}"
        Const cst_physicalFile1 As String = "{func}.aspx"
        Const cst_routeUrl2 As String = "{dir}/{func}"
        Const cst_physicalFile2 As String = "{dir}/{func}.aspx"

        'Dim routeName As String = "" '路徑的名稱
        'Dim routeUrl As String = "" ' 路徑的 URL 模式
        'Dim physicalFile As String = "" '路由的實體 URL
        Dim sFUNSORT As String() = TIMS.c_FUNSORT.Split(",")
        For i As Integer = 0 To sFUNSORT.Length - 1
            Dim sFUNID As String = sFUNSORT(i)
            routes.MapPageRoute($"{sFUNID}1", $"{sFUNID}/{cst_routeUrl1}", $"~/{sFUNID}/{cst_physicalFile1}") '加入至路由集合的路由
            routes.MapPageRoute($"{sFUNID}2", $"{sFUNID}/{cst_routeUrl2}", $"~/{sFUNID}/{cst_physicalFile2}") '加入至路由集合的路由
        Next

        routes.MapPageRoute("VisualChart1", "VisualChart/{func}", "~/VisualChart/{func}.aspx")
        routes.MapPageRoute("VisualChart2", "VisualChart/{dir}/{func}", "~/VisualChart/{dir}/{func}.aspx")

        ' 根路徑的路由要另外處理    
        Dim urls As IList = MyRouteHandler.GetRouteDef()
        For Each url As String In urls
            logger.Debug($"RegisterRoutes: add Route '{url}'")
            routes.Add(url, New Route(url, New MyRouteHandler()))
        Next

    End Sub


End Class
