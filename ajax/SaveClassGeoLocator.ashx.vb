Public Class SaveClassGeoLocator
    Implements System.Web.IHttpHandler

    Const cst_s_ip As String = "210.68.37.161"

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        'context.Response.ContentType = "text/plain"
        'context.Response.Write("Hello World!")
        context.Response.ContentType = "application/json"

        Const cst_message_ok As String = "ok"

        Dim message As String = cst_message_ok

        Dim str As String = ""

        Dim UTF8bytes As Byte()

        Dim check_parms_ok As Boolean = True
        Dim OCID As String = context.Request("ocid")
        Dim WGS84_Y As String = context.Request("wgs84_y")
        Dim WGS84_X As String = context.Request("wgs84_x")
        Dim AREANET As String = context.Request("areanet")
        If OCID = "" Then check_parms_ok = False 'Return
        If Not TIMS.IsNumeric1(OCID) Then check_parms_ok = False 'Return
        If WGS84_Y = "" Then check_parms_ok = False 'Return
        If WGS84_X = "" Then check_parms_ok = False 'Return
        If check_parms_ok AndAlso Not TIMS.IsNumeric1(WGS84_Y) Then check_parms_ok = False 'Return
        If check_parms_ok AndAlso WGS84_Y = 0 Then check_parms_ok = False 'Return
        If check_parms_ok AndAlso Not TIMS.IsNumeric1(WGS84_X) Then check_parms_ok = False 'Return
        If check_parms_ok AndAlso WGS84_X = 0 Then check_parms_ok = False 'Return
        TIMS.LOG.Debug(String.Format("#SaveClassGeoLocator AREANET: {0}", AREANET))
        If Not check_parms_ok Then
            'DbAccess.CloseDbConn(objconn)
            TIMS.LOG.Warn(String.Format("#SaveClassGeoLocator check parms is Empty ,OCID: {0},{1},{2}", OCID, WGS84_Y, WGS84_X))

            message = String.Format("fail, {0}", "check parms is Empty")

            str = String.Format("{0}""status"":""{1}""{2}", "{", message, "}")

            UTF8bytes = Encoding.UTF8.GetBytes(str)

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

        Dim fromIp As String = Common.GetIpAddress()
        Dim s_SaveGeoIP As String = TIMS.Utl_GetConfigVAL0(objconn, "SaveGeoIP", 0)
        TIMS.LOG.Debug(String.Concat("#ProcessRequest: ", " ,context.Request.IsLocal: ", context.Request.IsLocal, " ,fromIp: ", fromIp, " ,SaveGeoIP: ", s_SaveGeoIP))
        Dim flag_NG As Boolean = (Not context.Request.IsLocal AndAlso fromIp <> cst_s_ip AndAlso fromIp <> s_SaveGeoIP)
        Dim flag_NG2 As Boolean = ((AREANET = "") OrElse (AREANET <> "" AndAlso fromIp <> AREANET))
        If flag_NG AndAlso flag_NG2 Then
            DbAccess.CloseDbConn(objconn)
            TIMS.LOG.Warn(String.Format("Try to access from {0}", fromIp))
            context.Response.StatusCode = 401
            context.Response.End()
            Return
        End If

        '更新機構座標
        'param.Clear()
        Dim param As New Hashtable From {
            {"TWD97_X", WGS84_X},
            {"TWD97_Y", WGS84_Y},
            {"OCID", OCID}
        }

        Savedate1(objconn, message, param)

        If (message = cst_message_ok) Then Savedate2(objconn, message, param)

        DbAccess.CloseDbConn(objconn)

        'message = cst_message_ok

        str = String.Format("{0}""status"":""{1}""{2}", "{", message, "}")

        'Dim UTF8bytes As Byte() = Encoding.UTF8.GetBytes(str)
        UTF8bytes = Encoding.UTF8.GetBytes(str)

        context.Response.AddHeader("Content-Length", UTF8bytes.Length.ToString())
        context.Response.BinaryWrite(UTF8bytes)
        context.Response.Flush()
        context.Response.End()
    End Sub

    Sub Savedate1(ByRef oConn As SqlConnection, ByRef message As String, ByRef param As Hashtable)
        '更新機構座標
        'Dim param As New Hashtable
        'param.Clear()
        'param.Add("TWD97_X", WGS84_X)
        'param.Add("TWD97_Y", WGS84_Y)
        'param.Add("OCID", OCID)
        Dim sql As String = ""
        sql = "UPDATE CLASS_CLASSINFO SET TWD97_X=@TWD97_X, TWD97_Y=@TWD97_Y WHERE OCID=@OCID"

        Dim sCmd As New SqlCommand(sql, oConn)

        DbAccess.HashParmsChange(sCmd, param)

        Dim i_rst As Integer = 0
        Try

            i_rst = sCmd.ExecuteNonQuery()

        Catch ex As Exception
            Dim OCID As String = TIMS.GetMyValue2(param, "OCID")
            Dim sErrmsg As String = ""
            sErrmsg = String.Format("#SaveClassGeoLocator Savedate1 WGS84 fail,OCID: {0} ", OCID) & vbCrLf
            TIMS.LOG.Error(sErrmsg + ex.Message, ex)
            message = String.Format("fail, {0}", ex.Message)
        End Try

        TIMS.LOG.Debug(String.Format("#SaveClassGeoLocator Savedate1: {0},cnt: {1}", TIMS.GetMyValue3(param), i_rst))

    End Sub

    Sub Savedate2(ByRef oConn As SqlConnection, ByRef message As String, ByRef param As Hashtable)
        Dim OCID As String = TIMS.GetMyValue2(param, "OCID")
        Dim TWD97_X As String = TIMS.GetMyValue2(param, "TWD97_X")
        Dim TWD97_Y As String = TIMS.GetMyValue2(param, "TWD97_Y")

        'sql = "" & vbCrLf
        'sql &= " SELECT a.TWD97_X ,a.TWD97_Y" & vbCrLf
        'sql &= " ,A.taddresszip,A.taddress" & vbCrLf
        'sql &= " ,B.TWD97_X ,B.TWD97_Y" & vbCrLf
        'sql &= " ,b.taddresszip,b.taddress" & vbCrLf
        Dim sql As String = ""
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO b on b.taddresszip=a.taddresszip and b.taddress=a.taddress AND a.ocid!=b.ocid" & vbCrLf
        sql &= " WHERE (a.TWD97_X IS NULL or a.TWD97_Y IS NULL )" & vbCrLf
        sql &= " AND (b.TWD97_X IS NOT NULL or b.TWD97_Y IS NOT NULL )" & vbCrLf
        sql &= " and b.OCID=@OCID" & vbCrLf

        'param2.Clear()
        Dim param2 As New Hashtable From {
            {"OCID", OCID}
        }

        Dim sCmd2 As New SqlCommand(sql, oConn)

        DbAccess.HashParmsChange(sCmd2, param2)

        Dim dt2 As New DataTable

        dt2.Load(sCmd2.ExecuteReader())

        If dt2.Rows.Count = 0 Then Return

        TIMS.LOG.Debug(String.Format("#SaveClassGeoLocator Savedate2: {0},dt2.Rows: {1}", TIMS.GetMyValue3(param), dt2.Rows.Count))

        Try

            For Each dr2 As DataRow In dt2.Rows
                'param3.Clear()
                Dim param3 As New Hashtable From {
                    {"TWD97_X", TWD97_X},
                    {"TWD97_Y", TWD97_Y},
                    {"OCID", dr2("OCID")}
                }

                Savedate1(oConn, message, param3)
            Next

        Catch ex As Exception
            Dim sErrmsg As String = ""
            sErrmsg = String.Format("#SaveClassGeoLocator Savedate2 WGS84 fail,OCID: {0} ", OCID) & vbCrLf
            TIMS.LOG.Error(sErrmsg + ex.Message, ex)
            message = String.Format("fail, {0}", ex.Message)
        End Try

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return True
        End Get
    End Property

End Class
