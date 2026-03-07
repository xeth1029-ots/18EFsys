Partial Class TC_01_004_del
    Inherits AuthBasePage

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Dim ReqID As String = TIMS.ClearSQM(Request("ID"))
        Dim Re_ocid As String = TIMS.ClearSQM(Request("ocid"))
        Dim Re_planid As String = TIMS.ClearSQM(Request("PlanID"))
        Dim Re_ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim Re_SeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        Dim Re_Years As String = TIMS.ClearSQM(Request("Years"))

        If Convert.ToString(Re_ocid) = "" Then Errmsg += "傳入班級代號有誤" & vbCrLf

        Dim drC As DataRow = TIMS.GetOCIDDate(Re_ocid, objconn)
        If drC Is Nothing Then Errmsg += "傳入班級代號有誤" & vbCrLf

        If Convert.ToString(Re_planid) = "" _
            OrElse Convert.ToString(Re_ComIDNO) = "" _
            OrElse Convert.ToString(Re_SeqNO) = "" Then
            Errmsg += "傳入開班序號有誤" & vbCrLf
        End If

        Dim drP As DataRow = TIMS.GetPCSDate(Re_planid, Re_ComIDNO, Re_SeqNO, objconn)
        If drP Is Nothing Then Errmsg += "傳入開班序號有誤" & vbCrLf

        If (drC("OCID") <> drP("OCID")) Then Errmsg += "傳入開班序號與代號有誤" & vbCrLf

        If Convert.ToString(Re_Years) = "" Then
            Errmsg += "傳入計畫年度有誤" & vbCrLf
        End If

        If Errmsg <> "" Then Return False '離開。

        Dim i_a As Integer = 0
        Dim sqlstr As String = ""
        sqlstr = "SELECT COUNT(1) X FROM Stud_EnterType where ocid1='" & Re_ocid & "'"
        i_a += Convert.ToInt16(DbAccess.ExecuteScalar(sqlstr, objconn))
        If i_a = 0 Then
            sqlstr = "SELECT COUNT(1) X FROM Stud_EnterType where ocid2='" & Re_ocid & "'"
            i_a += Convert.ToInt16(DbAccess.ExecuteScalar(sqlstr, objconn))
        End If
        If i_a = 0 Then
            sqlstr = "SELECT COUNT(1) X FROM Stud_EnterType where ocid3='" & Re_ocid & "'"
            i_a += Convert.ToInt16(DbAccess.ExecuteScalar(sqlstr, objconn))
        End If
        If i_a = 0 Then
            sqlstr = "SELECT COUNT(1) X FROM Stud_EnterType2 where ocid1='" & Re_ocid & "'"
            i_a += Convert.ToInt16(DbAccess.ExecuteScalar(sqlstr, objconn))
        End If

        Dim sqlstr_B As String = "SELECT COUNT(1) FROM Class_StudentsOfClass where ocid='" & Re_ocid & "'"
        Dim i_b As Integer = Convert.ToInt16(DbAccess.ExecuteScalar(sqlstr_B, objconn))

        Dim sqlstr_C As String = "SELECT COUNT(1) FROM Class_Schedule where ocid='" & Re_ocid & "' "
        Dim i_c As Integer = Convert.ToInt16(DbAccess.ExecuteScalar(sqlstr_C, objconn))

        Dim is_parent As String = If((i_a + i_b + i_c) >= 1, TIMS.c_true, TIMS.c_false)

        If is_parent = TIMS.c_true Then
            Dim strMsg As String = ""
            If i_a > 0 Then strMsg += "此班級檔 尚有報名資料(Stud_EnterType:" & Re_ocid & "),已有資料參照,不可刪除!!!\n"

            If i_b > 0 Then strMsg += "此班級檔 尚有班級學員(Class_StudentsOfClass:" & Re_ocid & "),已有資料參照,不可刪除!!!\n"

            If i_c > 0 Then strMsg += "此班級檔 尚有排課(Class_Schedule:" & Re_ocid & "),已有資料參照,不可刪除!!!\n"

            If strMsg <> "" Then Errmsg += strMsg
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            If Session("ClassSearchStr") IsNot Nothing Then ViewState("ClassSearchStr") = Session("ClassSearchStr")
        End If

        '檢查若有異常離開。
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call DEL_CLASSINFO()

    End Sub

    Sub DEL_CLASSINFO()

        Dim ReqID As String = TIMS.ClearSQM(Request("ID"))
        Dim Re_ocid As String = TIMS.ClearSQM(Request("ocid"))
        Dim Re_planid As String = TIMS.ClearSQM(Request("PlanID"))
        Dim Re_ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim Re_SeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        Dim Re_Years As String = TIMS.ClearSQM(Request("Years"))

        Dim strClassname As String = ""
        Dim strCycle As String = ""
        Dim strRID As String = ""
        Dim strClsid As String = ""
        Dim strTPlanID As String = ""
        Dim strTPlanName As String = ""
        Dim strOrgname As String = ""
        Dim strOrgID As String = ""
        Dim strClassID As String = ""

        Dim drC As DataRow = TIMS.GetOCIDDate(Re_ocid, objconn)
        If drC Is Nothing Then Return

        Dim Sql As String = ""
        Dim dr As DataRow
        Sql = "SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & Re_ocid & "'"
        dr = DbAccess.GetOneRow(Sql, objconn)

        strClassname = Convert.ToString(dr("ClassCName"))
        strCycle = Convert.ToString(dr("CyclType"))
        strRID = Convert.ToString(dr("RID"))
        strClsid = Convert.ToString(dr("CLSID"))

        Sql = "SELECT TPLANID FROM ID_PLAN WHERE PlanID='" & Re_planid & "'"
        dr = DbAccess.GetOneRow(Sql, objconn)
        strTPlanID = Convert.ToString(dr("TPlanID"))

        Sql = "SELECT PLANNAME FROM KEY_PLAN WHERE TPlanID='" & strTPlanID & "'"
        dr = DbAccess.GetOneRow(Sql, objconn)
        strTPlanName = Convert.ToString(dr("PlanName"))

        Sql = "SELECT B.ORGNAME,B.ORGID FROM AUTH_RELSHIP A JOIN ORG_ORGINFO B ON A.ORGID=B.ORGID WHERE a.RID='" & strRID & "'"
        dr = DbAccess.GetOneRow(Sql, objconn)
        strOrgname = Convert.ToString(dr("OrgName"))
        strOrgID = Convert.ToString(dr("orgid"))

        Sql = "SELECT CLASSID FROM ID_CLASS WHERE CLSID='" & strClsid & "'"
        dr = DbAccess.GetOneRow(Sql, objconn)
        strClassID = Convert.ToString(dr("ClassID"))

        Dim Check_Sql_1 As String = "" 'check 課程階段資料
        Check_Sql_1 = "SELECT * FROM CLASS_CLASSLEVEL WHERE OCID='" & Re_ocid & "'"
        '有課程階段資料才刪除
        Dim dtCLASSLEVEL As DataTable = DbAccess.GetDataTable(Check_Sql_1, objconn)

        '已無此計畫轉入之班級資料,則將此筆計畫TransFlag='N'
        Dim Check_Sql2 As String = ""
        Check_Sql2 = "" & vbCrLf
        Check_Sql2 += " SELECT OCID FROM CLASS_CLASSINFO" & vbCrLf
        Check_Sql2 += " WHERE 1=1" & vbCrLf
        Check_Sql2 += " and OCID<>'" & Re_ocid & "'" & vbCrLf
        Check_Sql2 += " and Years='" & Re_Years & "'" & vbCrLf
        Check_Sql2 += " and PlanID=" & Re_planid & vbCrLf
        Check_Sql2 += " and ComIDNO='" & Re_ComIDNO & "'" & vbCrLf
        Check_Sql2 += " and SeqNO=" & Re_SeqNO & vbCrLf
        Check_Sql2 += " and isnull(CyclType,'')='" & strCycle & "'" & vbCrLf
        Check_Sql2 += " and RID='" & strRID & "'" & vbCrLf
        Check_Sql2 += " and CLSID='" & strClsid & "'" & vbCrLf
        'Check_Sql2 += " and ClassNum is not null " & vbCrLf
        Dim dtCLASSINFO As DataTable = DbAccess.GetDataTable(Check_Sql2, objconn)

        Dim str As String = ""
        '刪除[訓練計畫名稱]-[機構名稱]-[(班別代碼)班別名稱]-[期別]
        str = "刪除[" & strTPlanName & "]-[" & strOrgname & "]-[(" & strClassID & ")" & strClassname & "]-[" & strCycle & "]"

        Dim delFlagOk As Boolean = False '刪除情況ok嗎？

        TIMS.InsertDelLog(sm.UserInfo.UserID, ReqID, sm.UserInfo.DistID, str, strOrgID, strRID, Re_planid, Re_ComIDNO, Re_SeqNO, Re_ocid)

        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'BeginTrans 'objTrans = DbAccess.BeginTrans(objconn)

            '刪除開班資料
            Sql = " DELETE CLASS_CLASSINFO WHERE OCID='" & Re_ocid & "'"
            DbAccess.ExecuteNonQuery(Sql, objTrans)

            '刪除課程階段資料
            If dtCLASSLEVEL.Rows.Count > 0 Then  '有課程階段資料才刪除
                Sql = " DELETE CLASS_CLASSLEVEL WHERE OCID='" & Re_ocid & "'"
                DbAccess.ExecuteNonQuery(Sql, objTrans)
            End If
            If dtCLASSINFO.Rows.Count = 0 Then '已無此計畫轉入之班級資料,則將此筆計畫TransFlag='N'
                Sql = " UPDATE PLAN_PLANINFO SET TransFlag='N' WHERE PlanID=" & Re_planid & " and ComIDNO='" & Re_ComIDNO & "' and SeqNO=" & Re_SeqNO
                DbAccess.ExecuteNonQuery(Sql, objTrans)
            End If
            DbAccess.CommitTrans(objTrans)

            If Session("ClassSearchStr") Is Nothing AndAlso ViewState("ClassSearchStr") IsNot Nothing Then Session("ClassSearchStr") = ViewState("ClassSearchStr")

            delFlagOk = True '刪除情況ok嗎？ok

        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            TIMS.CloseDbConn(objconn)
            Common.MessageBox(Me, "!!刪除失敗!!")
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
        End Try

        If delFlagOk Then
            '刪除情況ok嗎？ok
            'Response.Redirect("TC_01_004.aspx?ID=" & Request("ID") & "")
            'Response.Redirect("TC_01_004.aspx?ProcessType=del&ID=" & Request("ID") & "")
            Dim url1 As String = "TC_01_004.aspx?ProcessType=del&ID=" & ReqID & ""
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        TIMS.CloseDbConn(objconn)
    End Sub
End Class
