Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Web.Compilation
Imports System.Web.Routing
Imports System.Xml.Linq

Public Class MyRouteHandler
    Implements IRouteHandler

    Private Shared logger As ILog = LogManager.GetLogger(GetType(MyRouteHandler))

    Public Sub New()

    End Sub

    ''' <summary>
    ''' 資安調整(沒有作用)
    ''' </summary>
    ''' <param name="requestContext"></param>
    ''' <returns></returns>
    Public Function IRouteHandler_GetHttpHandler(requestContext As RequestContext) As IHttpHandler Implements IRouteHandler.GetHttpHandler
        Dim reqPath As String = requestContext.HttpContext.Request.Path
        Dim routePath As String = String.Empty
        Dim idx As Integer = reqPath.IndexOf(".aspx", StringComparison.OrdinalIgnoreCase)

        If reqPath.EndsWith("/") Then
            routePath = reqPath
        ElseIf idx > 0 Then
            routePath = reqPath
        Else
            idx = reqPath.IndexOf("?", StringComparison.OrdinalIgnoreCase)
            routePath = If(idx > 0, String.Concat(reqPath.Substring(0, idx), ".aspx?", reqPath.Substring(idx + 1)), String.Concat(reqPath, ".aspx"))
        End If

        routePath = If(Not routePath.StartsWith("/"), String.Concat("~/", routePath), String.Concat("~", routePath))

        logger.Debug(String.Concat(reqPath, " => ", routePath))

        Dim handlePage As Page = BuildManager.CreateInstanceFromVirtualPath(routePath, GetType(Page))

        Return handlePage
    End Function

    ''' <summary>
    ''' 根目錄的修改
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function GetRouteDef() As IList
        Dim urls As IList = New List(Of String)()

        Dim assembly As Assembly = Assembly.GetExecutingAssembly()
        Dim binPath As String = Path.GetDirectoryName(GetType(MyRouteHandler).Assembly.CodeBase)
        Dim config As XDocument = XDocument.Load(binPath + "//class//RouteConfig.xml")
        Dim q As IEnumerable(Of XElement) =
            From el In config.Elements("RouteList").Elements("RouteDef")
            Select el

        If q.Count > 0 Then
            Dim item As XElement
            For Each item In q
                urls.Add(item.Attribute("url").Value)
            Next
        End If

        Return urls
    End Function

#Region "NO-USE"
    'Public Shared Function RouteRootFunc() As String

    '    Dim appPath As String = HttpContext.Current.Request.ApplicationPath
    '    Dim reqPath As String = TIMS.Get_RequestPath()
    '    Dim routePath As String = String.Empty

    '    Dim p As Integer = reqPath.IndexOf("/")
    '    If p > 0 Then
    '        ' Not Root Path Function, Ignored
    '    Else
    '        If Not String.IsNullOrEmpty(reqPath) Then
    '            If Not reqPath.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then
    '                routePath = reqPath & ".aspx"
    '                logger.Debug("RouteRootFunc: (appPath: " & appPath & ", reqPath: " & reqPath & ") routePath: [" & routePath & "]")
    '            End If
    '        Else
    '            ' Default Home Page, Ignored
    '        End If
    '    End If

    '    Return routePath
    'End Function
#End Region

End Class
