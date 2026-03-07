Public Class GetAddressPTID1
    Implements System.Web.IHttpHandler

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        context.Response.ContentType = "application/json"

        Dim UTF8bytes As Byte()

        Dim result As New AjaxResultStruct

        Dim check_parms_ok As Boolean = True
        Dim COMIDNO As String = TIMS.ClearSQM(context.Request("comidno"))
        Dim PLACEID As String = TIMS.ClearSQM(context.Request("placeid"))
        If COMIDNO = "" Then check_parms_ok = False
        If PLACEID = "" Then check_parms_ok = False
        If Not check_parms_ok Then
            TIMS.LOG.Warn(String.Format("#GetAddressPTID1 check parms is Empty ,{0},{1} ", COMIDNO, PLACEID))
            result.status = False
            result.message = "fail"
            result.data = "check parms is Empty"
            UTF8bytes = Encoding.UTF8.GetBytes(result.Serialize())
            context.Response.AddHeader("Content-Length", UTF8bytes.Length.ToString())
            context.Response.BinaryWrite(UTF8bytes)
            context.Response.Flush()
            'context.Response.StatusCode = 401
            context.Response.End()
            Return
        End If

        Dim objconn As SqlConnection = DbAccess.GetConnection()

        If Not TIMS.OpenDbConn(objconn) Then
            DbAccess.CloseDbConn(objconn)
            TIMS.LOG.Error(String.Concat("#Not DbAccess.Open()!"))
            context.Response.StatusCode = 401
            context.Response.End()
            Return
        End If

        Dim s_PTID As String = TIMS.Get_PTID(PLACEID, COMIDNO, objconn)
        DbAccess.CloseDbConn(objconn)

        result.status = True
        result.message = "ok"
        result.data = If(s_PTID <> "", s_PTID, "-1")
        UTF8bytes = Encoding.UTF8.GetBytes(result.Serialize())
        context.Response.AddHeader("Content-Length", UTF8bytes.Length.ToString())
        context.Response.BinaryWrite(UTF8bytes)
        context.Response.Flush()
        context.Response.End()
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return True
        End Get
    End Property

End Class