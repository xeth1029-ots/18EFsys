Imports System.Threading.Tasks

Partial Class SYS_03_009
    Inherits AuthBasePage

    'Dim chkEnableAll As CheckBox
    'Dim chkAddsAll As CheckBox
    'Dim chkModAll As CheckBox
    'Dim chkDelAll As CheckBox
    'Dim chkSechAll As CheckBox
    'Dim chkPrntAll As CheckBox
    'Dim trID As Integer = 0

#Region "Function"
    Private Function Get_StudStudentInfo(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        'Call TIMS.OpenDbConn(objConn)
        Dim rst As DataTable = Nothing
        Dim sqlStr As String = "SELECT SID,IDNO,Name,format(Birthday,'yyyy/MM/dd') Birthday FROM STUD_STUDENTINFO WHERE 1=1 "
        If tmpIDNO <> "" Then sqlStr &= " and IDNO like @IDNO "
        If tmpName <> "" Then sqlStr &= " and Name like @Name "

        Dim oDt As New DataTable
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            If tmpIDNO <> "" Then .Parameters.Add("IDNO", SqlDbType.VarChar).Value = String.Concat(UCase(tmpIDNO), "%")
            If tmpName <> "" Then .Parameters.Add("Name", SqlDbType.NVarChar).Value = String.Concat("%", tmpName, "%")
            oDt.Load(.ExecuteReader())
        End With

        If oDt.Rows.Count > 0 Then rst = oDt
        Return rst
    End Function

    Private Function Get_StudEnterTemp(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        'Call TIMS.OpenDbConn(objConn)
        Dim rst As DataTable = Nothing
        Dim sqlStr As String = "SELECT eSETID,SETID,IDNO,Name,format(Birthday,'yyyy/MM/dd') Birthday FROM STUD_ENTERTEMP WHERE 1=1 "
        If tmpIDNO <> "" Then sqlStr &= " and IDNO like @IDNO "
        If tmpName <> "" Then sqlStr &= " and Name like @Name "

        Dim oDt As New DataTable
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            If tmpIDNO <> "" Then .Parameters.Add("IDNO", SqlDbType.VarChar).Value = String.Concat(UCase(tmpIDNO), "%")
            If tmpName <> "" Then .Parameters.Add("Name", SqlDbType.NVarChar).Value = String.Concat("%", tmpName, "%")
            oDt.Load(.ExecuteReader())
        End With
        If oDt.Rows.Count > 0 Then rst = oDt
        Return rst
    End Function

    Private Function Get_StudEnterTemp2(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        'Call TIMS.OpenDbConn(objConn)
        Dim rst As DataTable = Nothing
        Dim sqlStr As String = "SELECT eSETID,SETID,IDNO,Name,format(Birthday,'yyyy/MM/dd') Birthday FROM STUD_ENTERTEMP2 WHERE 1=1 "
        If tmpIDNO <> "" Then sqlStr &= " and IDNO like @IDNO "
        If tmpName <> "" Then sqlStr &= " and Name like @Name "

        Dim oDt As New DataTable
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            If tmpIDNO <> "" Then .Parameters.Add("IDNO", SqlDbType.VarChar).Value = String.Concat(UCase(tmpIDNO), "%")
            If tmpName <> "" Then .Parameters.Add("Name", SqlDbType.NVarChar).Value = String.Concat("%", tmpName, "%")
            oDt.Load(.ExecuteReader())
        End With
        If oDt.Rows.Count > 0 Then rst = oDt
        Return rst
    End Function

    Function Get_StudEnterTempDelData(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        'Call TIMS.OpenDbConn(objConn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select '2' stype,a.eSETID,a.SETID,a.IDNO,a.Name,format(a.Birthday,'yyyy/MM/dd') Birthday FROM STUD_ENTERTEMPDELDATA a" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If tmpIDNO <> "" Then
            sql &= " AND (1!=1 " & vbCrLf
            sql &= " OR a.IDNO like @IDNO " & vbCrLf
            sql &= " OR a.SETID IN (SELECT SETID FROM STUD_ENTERTEMP WHERE IDNO like @IDNO)" & vbCrLf
            sql &= " )" & vbCrLf
        End If
        If tmpName <> "" Then sql &= " and a.Name like @Name "

        sql &= " union " & vbCrLf
        sql &= " select '1' stype,eSETID,SETID,IDNO,Name,format(Birthday,'yyyy/MM/dd') Birthday FROM STUD_ENTERTEMP" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If tmpIDNO <> "" Then sql &= " and IDNO like @IDNO "
        If tmpName <> "" Then sql &= " and Name like @Name "

        Dim Rst As New DataTable '= Nothing
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            If tmpIDNO <> "" Then .Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO) & "%"
            If tmpName <> "" Then .Parameters.Add("Name", SqlDbType.NVarChar).Value = "%" & tmpName & "%"
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Function Get_StudStudentInfoDelData(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        'Call TIMS.OpenDbConn(objConn)
        'Dim Rst As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select 'D' TT,a.SID,b.SOCID, a.IDNO,a.Name,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " FROM STUD_STUDENTINFODELDATA a" & vbCrLf
        sql &= " LEFT JOIN CLASS_STUDENTSOFCLASSDELDATA b on b.SID =a.SID" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If tmpIDNO <> "" Then
            sql &= " AND (1!=1 " & vbCrLf
            sql &= " OR a.IDNO like @IDNO " & vbCrLf
            sql &= " OR a.SID IN (SELECT SID FROM STUD_STUDENTINFO WHERE IDNO like @IDNO)" & vbCrLf
            sql &= " )" & vbCrLf
        End If
        If tmpName <> "" Then sql &= " and a.Name like @Name " & vbCrLf

        sql &= " UNION " & vbCrLf
        sql &= " select 'N' TT,a.SID,b.SOCID, a.IDNO,a.Name,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASSDELDATA b on b.SID =a.SID" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If tmpIDNO <> "" Then sql &= " and a.IDNO like @IDNO " & vbCrLf
        If tmpName <> "" Then sql &= " and a.Name like @Name " & vbCrLf

        Dim Rst As New DataTable '= Nothing
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            If tmpIDNO <> "" Then .Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO) & "%"
            If tmpName <> "" Then .Parameters.Add("Name", SqlDbType.NVarChar).Value = "%" & tmpName & "%"
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Function Get_StudEnterTemp2DelData(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        'Call TIMS.OpenDbConn(objConn)
        'Dim dt2 As DataTable = Get_StudEnterTemp2(tmpIDNO, tmpName)
        'Dim flag_dt2 As Boolean = If(dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0, True, False)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select '2' stype,a.eSETID,a.SETID,a.IDNO,a.Name,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2DELDATA a" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If tmpIDNO <> "" Then
            sql &= " AND (1!=1 " & vbCrLf
            sql &= " OR a.IDNO like @IDNO " & vbCrLf
            sql &= " OR a.eSETID IN (SELECT eSETID FROM STUD_ENTERTEMP2 WHERE IDNO like @IDNO )" & vbCrLf
            sql &= " )" & vbCrLf
        End If
        If tmpName <> "" Then sql &= " and a.Name like @Name "

        sql &= " union " & vbCrLf
        sql &= " select '1' stype,eSETID,SETID,IDNO,Name,format(Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If tmpIDNO <> "" Then sql &= " and IDNO like @IDNO "
        If tmpName <> "" Then sql &= " and Name like @Name "

        Dim Rst As New DataTable '= Nothing
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            If tmpIDNO <> "" Then .Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO) & "%"
            If tmpName <> "" Then .Parameters.Add("Name", SqlDbType.NVarChar).Value = "%" & tmpName & "%"
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Private Function Get_ClassStudentsOfClass(ByVal columnName As String, ByVal tmpID As String) As DataTable
        Dim oDt As New DataTable
        If tmpID = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)

        Dim sqlStr As String
        Dim rst As DataTable = Nothing
        Dim columns() As String = {"SID", "SETID"}
        Dim fg_OK As Boolean = Array.IndexOf(columns, columnName) <> -1
        If Not fg_OK Then Return oDt 'TIMS.dtNothing()

        sqlStr = "" & vbCrLf
        sqlStr &= " select a.SID,a.SOCID,a.OCID" & vbCrLf
        sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlStr &= " FROM CLASS_STUDENTSOFCLASS a WITH(NOLOCK)" & vbCrLf
        sqlStr &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID" & vbCrLf
        sqlStr &= String.Concat(" WHERE a.", columnName, "= @tmpID") & vbCrLf
        sqlStr &= " ORDER BY a.SID,a.SOCID,a.OCID"
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            Select Case UCase(columnName)
                Case "SID"
                    .Parameters.Add("tmpID", SqlDbType.VarChar).Value = tmpID
                Case "SETID"
                    .Parameters.Add("tmpID", SqlDbType.Int).Value = tmpID
            End Select
            oDt.Load(.ExecuteReader())
        End With
        If oDt.Rows.Count > 0 Then rst = oDt

        Return rst
    End Function

    'Async Function Async_StudEnterType(ByVal columnName As String, ByVal tmpID As String) As Task(Of DataTable)
    Private Function Get_StudEnterType(ByVal columnName As String, ByVal tmpID As String) As DataTable
        Dim oDt As New DataTable
        If tmpID = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)

        Dim sqlStr As String
        Dim rst As DataTable = Nothing
        Dim columns() As String = {"SETID", "eSETID"}
        Dim fg_OK As Boolean = Array.IndexOf(columns, columnName) <> -1
        If Not fg_OK Then Return oDt 'TIMS.dtNothing()

        sqlStr = "" & vbCrLf
        sqlStr &= " select a.SETID,a.SerNum,a.OCID1" & vbCrLf
        sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlStr &= " FROM STUD_ENTERTYPE a WITH(NOLOCK)" & vbCrLf
        sqlStr &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) ON b.OCID=a.OCID1" & vbCrLf
        sqlStr &= String.Concat(" WHERE a.", columnName, "= @tmpID") & vbCrLf
        sqlStr &= " ORDER BY a.SETID,a.SerNum,a.OCID1"
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("tmpID", SqlDbType.Int).Value = Val(tmpID)
            oDt.Load(.ExecuteReader())
        End With
        If oDt.Rows.Count > 0 Then rst = oDt
        Return rst
    End Function

    Private Function Get_StudEnterType2(ByVal columnName As String, ByVal tmpID As String) As DataTable
        Dim oDt As New DataTable
        If tmpID = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)

        Dim sqlStr As String
        Dim rst As DataTable = Nothing
        Dim columns() As String = {"SETID", "eSETID"}
        Dim fg_OK As Boolean = Array.IndexOf(columns, columnName) <> -1
        If Not fg_OK Then Return oDt 'TIMS.dtNothing()

        sqlStr = "" & vbCrLf
        sqlStr &= " select a.SETID,a.eSETID,a.eSerNum,a.OCID1" & vbCrLf
        sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlStr &= " FROM STUD_ENTERTYPE2 a WITH(NOLOCK)" & vbCrLf
        sqlStr &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID1" & vbCrLf
        sqlStr &= String.Concat(" WHERE a.", columnName, "= @tmpID") & vbCrLf
        sqlStr &= " ORDER BY a.SETID,a.eSETID,a.eSerNum,a.OCID1"
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("tmpID", SqlDbType.Int).Value = Val(tmpID)
            oDt.Load(.ExecuteReader())
        End With
        If oDt.Rows.Count > 0 Then rst = oDt

        Return rst
    End Function

    Function GET_STUDENTSOFCLASSDELDATA(ByVal columnName As String, ByVal tmpValue As String) As DataTable
        'If tmpValue = "" Then Return TIMS.dtNothing()
        Dim oDt As New DataTable '= Nothing
        If tmpValue = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " select dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlstr &= " ,b.STDATE,b.FTDATE" & vbCrLf
        sqlstr &= " ,a.SOCID,a.SID" & vbCrLf
        sqlstr &= " ,a.OCID" & vbCrLf
        sqlstr &= " ,a.Modifydate,a.Modifyacct" & vbCrLf
        sqlstr &= " FROM CLASS_STUDENTSOFCLASSDELDATA a WITH(NOLOCK)" & vbCrLf
        sqlstr &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID" & vbCrLf
        sqlstr &= " where 1=1" & vbCrLf
        sqlstr &= String.Format(" and a.{0}={1}", columnName, tmpValue)
        'SOCID
        Dim sCmd As New SqlCommand(sqlstr, objConn)
        With sCmd
            .Parameters.Clear()
            oDt.Load(.ExecuteReader())
        End With
        Return oDt
    End Function

    Function Get_StudEnterTypeDelData(ByVal columnName As String, ByVal tmpValue As String) As DataTable
        Dim oDt As New DataTable '= Nothing
        If tmpValue = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " select dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlstr &= " ,b.STDATE,b.FTDATE" & vbCrLf
        sqlstr &= " ,a.eSerNum,a.eSETID" & vbCrLf
        sqlstr &= " ,a.SETID,a.EnterDate,a.SerNum,a.OCID1" & vbCrLf
        sqlstr &= " ,a.RelEnterDate" & vbCrLf
        sqlstr &= " ,a.Modifydate,a.Modifyacct" & vbCrLf
        sqlstr &= " FROM STUD_ENTERTYPEDELDATA a WITH(NOLOCK)" & vbCrLf
        sqlstr &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID1" & vbCrLf
        sqlstr &= " where 1=1" & vbCrLf
        sqlstr &= String.Format(" and a.{0}={1}", columnName, tmpValue)
        'SETID
        Dim sCmd As New SqlCommand(sqlstr, objConn)
        With sCmd
            .Parameters.Clear()
            oDt.Load(.ExecuteReader())
        End With
        Return oDt
    End Function

    Function Get_STUD_DELENTERTYPE(ByVal columnName As String, ByVal tmpValue As String) As DataTable
        Dim oDt As New DataTable '= Nothing
        If tmpValue = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)
        'If tmpValue = "" Then Return TIMS.dtNothing()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,b.STDATE,b.FTDATE" & vbCrLf
        sql &= " ,a.eSerNum,a.eSETID" & vbCrLf
        sql &= " ,a.SETID,a.EnterDate,a.SerNum,a.OCID1" & vbCrLf
        sql &= " ,a.RelEnterDate" & vbCrLf
        sql &= " ,a.Modifydate,a.Modifyacct" & vbCrLf
        sql &= " FROM STUD_DELENTERTYPE a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID1" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= String.Format(" and a.{0}={1}", columnName, tmpValue)
        'SETID
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            oDt.Load(.ExecuteReader())
        End With
        Return oDt
        'Dim Rdt As DataTable = DbAccess.GetDataTable(sql, objConn)
        'Return Rdt
    End Function

    Function Get_StudEnterType2DelData(ByVal columnName As String, ByVal tmpValue As String) As DataTable
        Dim oDt As New DataTable '= Nothing
        If tmpValue = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)
        'If tmpValue = "" Then Return TIMS.dtNothing()
        Dim Rst As DataTable = Nothing
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " select dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlstr &= " ,b.STDATE,b.FTDATE" & vbCrLf
        sqlstr &= " ,a.eSerNum,a.eSETID" & vbCrLf
        sqlstr &= " ,a.SETID,a.EnterDate,a.SerNum,a.OCID1" & vbCrLf
        sqlstr &= " ,a.RelEnterDate" & vbCrLf
        sqlstr &= " ,a.Modifydate,a.Modifyacct" & vbCrLf
        sqlstr &= " FROM STUD_ENTERTYPE2DELDATA a WITH(NOLOCK)" & vbCrLf
        sqlstr &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID1" & vbCrLf
        sqlstr &= " where 1=1" & vbCrLf
        sqlstr &= String.Format(" and a.{0}={1}", columnName, tmpValue)
        'eSETID
        Dim sCmd As New SqlCommand(sqlstr, objConn)
        With sCmd
            .Parameters.Clear()
            oDt.Load(.ExecuteReader())
        End With
        Return oDt
        'Rst = DbAccess.GetDataTable(sqlstr, objConn)
        'Return Rst
    End Function

    Function Get_STUD_DELENTERTYPE2(ByVal columnName As String, ByVal tmpValue As String) As DataTable
        Dim oDt As New DataTable '= Nothing
        If tmpValue = "" Then Return oDt 'TIMS.dtNothing()
        'Call TIMS.OpenDbConn(objConn)
        'If tmpValue = "" Then Return TIMS.dtNothing()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,b.STDATE,b.FTDATE" & vbCrLf
        sql &= " ,a.eSerNum,a.eSETID" & vbCrLf
        sql &= " ,a.SETID,a.EnterDate,a.SerNum,a.OCID1" & vbCrLf
        sql &= " ,a.RelEnterDate" & vbCrLf
        sql &= " ,a.Modifydate,a.Modifyacct" & vbCrLf
        sql &= " FROM STUD_DELENTERTYPE2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) on b.OCID=a.OCID1" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= String.Format(" and a.{0}={1}", columnName, tmpValue)
        'SETID
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            oDt.Load(.ExecuteReader())
        End With
        Return oDt
        'Dim Rdt As DataTable = DbAccess.GetDataTable(sql, objConn)
        'Return Rdt
    End Function

#Region "NO USE"
    'Private Function Get_SubSubSidyApply(ByVal tmpSID As String) As DataTable
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim objDS As New DataSet
    '    Dim oDt As New DataTable
    '    Dim sqlStr As String
    '    sqlStr = "" & vbCrLf
    '    sqlStr &= " select a.SID" & vbCrLf
    '    sqlStr &= " ,a.OCID" & vbCrLf
    '    sqlStr &= " ,a.SUBID" & vbCrLf
    '    sqlStr &= " ,a.AppliedStatusF" & vbCrLf
    '    sqlStr &= " ,a.AppliedStatusFin" & vbCrLf
    '    sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
    '    sqlStr &= " FROM SUB_SUBSIDYAPPLY a " & vbCrLf
    '    sqlStr &= " JOIN CLASS_CLASSINFO b on b.OCID=a.OCID" & vbCrLf
    '    sqlStr &= " where SID= @SID "
    '    sqlStr &= " order by SUBID asc "
    '    With sqlAdp
    '        .SelectCommand = New SqlCommand(sqlStr, objConn)
    '        .SelectCommand.Parameters.Clear()
    '        .SelectCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '        '.Fill(objDS, "Data")
    '        .Fill(oDt)
    '    End With
    '    If oDt.Rows.Count > 0 Then rst = oDt
    '    Dim rst As DataTable = Nothing
    '    Return rst
    'End Function
#End Region

    Private Sub Show_DataGrid2(ByVal tmpIDNO As String, ByVal tmpName As String, Optional ByVal tmpPage As Integer = 0)
        Dim dt_Student As DataTable = Nothing
        'DataGrid2.DataSource = Nothing
        'DataGrid2.DataBind()

        DataGrid2.DataKeyField = ""
        Select Case Convert.ToString(ViewState(cst_vs_Option))
            Case "STUDENTINFO"
                dt_Student = Get_StudStudentInfo(tmpIDNO, tmpName)
                DataGrid2.DataKeyField = "SID"
            Case "ENTERTEMP"
                dt_Student = Get_StudEnterTemp(tmpIDNO, tmpName)
                DataGrid2.DataKeyField = "SETID"
            Case "ENTERTEMP2"
                dt_Student = Get_StudEnterTemp2(tmpIDNO, tmpName)
                DataGrid2.DataKeyField = "eSETID"

            Case "3" 'StudentInfo(DEL LOG)
                dt_Student = Get_StudStudentInfoDelData(tmpIDNO, tmpName)
                DataGrid2.DataKeyField = "SID"
            Case "4" 'EnterTemp(DEL LOG)
                dt_Student = Get_StudEnterTempDelData(tmpIDNO, tmpName)
                DataGrid2.DataKeyField = "SETID"
            Case "5" 'EnterTemp2(DEL LOG)
                dt_Student = Get_StudEnterTemp2DelData(tmpIDNO, tmpName)
                DataGrid2.DataKeyField = "eSETID"

        End Select

        If dt_Student Is Nothing Then
            lab_Msg.Visible = True
            DataGrid2.Visible = False
            Return
        End If

        DataGrid2.DataSource = dt_Student
        DataGrid2.CurrentPageIndex = tmpPage
        DataGrid2.DataBind()
        DataGrid2.Visible = True
        lab_Msg.Visible = False

    End Sub

    Private Sub Del_StudStudentInfo(ByVal tmpSID As String, ByVal tmpTrans As SqlTransaction)
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String

        Try
            sqlStr = "UPDATE STUD_STUDENTINFO SET ModifyAcct= @ModifyAcct,ModifyDate=GETDATE() where SID=@SID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
                .UpdateCommand.ExecuteNonQuery()
            End With
            sqlStr = "INSERT INTO STUD_STUDENTINFODELDATA SELECT * FROM STUD_STUDENTINFO WHERE SID=@SID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.ExecuteNonQuery()
            End With
            sqlStr = "DELETE STUD_STUDENTINFO WHERE SID=@SID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
                .DeleteCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .DeleteCommand.ExecuteNonQuery()
            End With
        Catch ex As Exception
            tmpTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            'objConn.Close()
            sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub Del_StudSubData(ByVal tmpSID As String, ByVal tmpTrans As SqlTransaction)
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String

        Try
            sqlStr = "update Stud_SubData set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SID= @SID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
                .UpdateCommand.ExecuteNonQuery()
            End With
            sqlStr = "insert into Stud_SubDataDelData select * from Stud_SubData where SID= @SID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.ExecuteNonQuery()
            End With
            sqlStr = "delete Stud_SubData where SID= @SID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
                .DeleteCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .DeleteCommand.ExecuteNonQuery()
            End With
        Catch ex As Exception
            tmpTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            'objConn.Close()
            sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub Del_StudEnterTemp(ByVal tmpSETID As String, ByVal tmpTrans As SqlTransaction)
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String

        Try
            sqlStr = "update STUD_ENTERTEMP set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
                .UpdateCommand.ExecuteNonQuery()
            End With
            sqlStr = "insert into STUD_ENTERTEMPDelData select * from STUD_ENTERTEMP where SETID= @SETID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.ExecuteNonQuery()
            End With
            sqlStr = "delete STUD_ENTERTEMP where SETID= @SETID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
                .DeleteCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .DeleteCommand.ExecuteNonQuery()
            End With
        Catch ex As Exception
            tmpTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            'objConn.Close()
            sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub Del_StudSelResult(ByVal tmpSETID As String, ByVal tmpTrans As SqlTransaction)
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String

        Try
            sqlStr = "update Stud_SelResult set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
                .UpdateCommand.ExecuteNonQuery()
            End With
            sqlStr = "insert into Stud_SelResultDelData select * from Stud_SelResult where SETID= @SETID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.ExecuteNonQuery()
            End With
            sqlStr = "delete Stud_SelResult where SETID= @SETID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
                .DeleteCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .DeleteCommand.ExecuteNonQuery()
            End With
        Catch ex As Exception
            tmpTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            'objConn.Close()
            sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub Del_StudEnterTemp2(ByVal tmpeSETID As String, ByVal tmpTrans As SqlTransaction)
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String

        Try
            sqlStr = "update STUD_ENTERTEMP2 set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where eSETID= @eSETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("eSETID", SqlDbType.VarChar).Value = tmpeSETID
                .UpdateCommand.ExecuteNonQuery()
            End With
            sqlStr = "insert into STUD_ENTERTEMP2DELDATA select * from STUD_ENTERTEMP2 where eSETID= @eSETID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("eSETID", SqlDbType.VarChar).Value = tmpeSETID
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.ExecuteNonQuery()
            End With
            sqlStr = "delete STUD_ENTERTEMP2 where eSETID= @eSETID AND ModifyAcct= @ModifyAcct"
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("eSETID", SqlDbType.VarChar).Value = tmpeSETID
                .DeleteCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .DeleteCommand.ExecuteNonQuery()
            End With
        Catch ex As Exception
            tmpTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            'objConn.Close()
            sqlAdp.Dispose()
        End Try
    End Sub
#End Region
    Const cst_vs_IDNO As String = "IDNO"
    Const cst_vs_Name As String = "Name"
    Const cst_vs_Option As String = "Option"

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn)

        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
            str_superuser1 = CStr(sm.UserInfo.UserID)
        End If
        If Not flgROLEIDx0xLIDx0 Then
            Common.MessageBox(Me, "權限錯誤，無法使用此功能。")
            Exit Sub
        End If

        'If Convert.ToString(sm.UserInfo.UserID) <> str_superuser1 Then
        '    Common.MessageBox(Me, "權限錯誤，無法使用此功能。")
        '    Exit Sub
        'End If
        'If objConn.State = ConnectionState.Closed Then objConn.Open()
    End Sub

    '查詢
    Private Sub btn_Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Query.Click
        If str_superuser1 <> CStr(sm.UserInfo.UserID) Then Exit Sub

        txt_IDNO.Text = UCase(txt_IDNO.Text) '換大寫，若有小寫問題，直接修改資料庫
        txt_IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txt_IDNO.Text))
        txt_Name.Text = TIMS.ClearSQM(txt_Name.Text)

        ViewState(cst_vs_IDNO) = txt_IDNO.Text '.Trim(" ")
        ViewState(cst_vs_Name) = txt_Name.Text '.Trim(" ")
        ViewState(cst_vs_Option) = UCase(TIMS.GetListValue(rdo_Option))

        If Convert.ToString(ViewState(cst_vs_IDNO)) = "" AndAlso Convert.ToString(ViewState(cst_vs_Name)) = "" Then Return

        Show_DataGrid2(Convert.ToString(ViewState(cst_vs_IDNO)), Convert.ToString(ViewState(cst_vs_Name)))
    End Sub

    '刪除
    Private Sub DataGrid2_ItemCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim sCmdName As String = If(e IsNot Nothing, e.CommandName, "")
        Dim sCmdArg As String = If(e IsNot Nothing, e.CommandArgument, "")
        If sCmdName = "" Then Return
        If sCmdArg = "" Then Return
        'ViewState(cst_vs_Option) = UCase(TIMS.GetListValue(rdo_Option))
        If Convert.ToString(ViewState(cst_vs_Option)) <> UCase(rdo_Option.SelectedValue) Then
            Common.MessageBox(Me, "功能與刪除內容不符請重新操作該功能!!")
            Return
        End If

        Select Case UCase(e.CommandName)
            Case "DEL"
                Dim labSID As Label = e.Item.FindControl("lab_SID")
                If labSID Is Nothing OrElse labSID.Text = "" Then Return
                If labSID.Text <> "" AndAlso ViewState(cst_vs_Option) = UCase(TIMS.GetListValue(rdo_Option)) Then
                    Dim sqlTrans As SqlTransaction = objConn.BeginTransaction()
                    Try
                        Select Case Convert.ToString(ViewState(cst_vs_Option))
                            Case "STUDENTINFO"
                                Dim s_SID As String = TIMS.GetMyValue(sCmdArg, "SID")
                                If s_SID = "" Then Return
                                Del_StudStudentInfo(s_SID, sqlTrans)
                                Del_StudSubData(s_SID, sqlTrans)
                            Case "ENTERTEMP"
                                Dim s_SETID As String = TIMS.GetMyValue(sCmdArg, "SETID")
                                If s_SETID = "" Then Return
                                Del_StudEnterTemp(s_SETID, sqlTrans)
                                Del_StudSelResult(s_SETID, sqlTrans)
                            Case "ENTERTEMP2"
                                Dim s_eSETID As String = TIMS.GetMyValue(sCmdArg, "eSETID")
                                If s_eSETID = "" Then Return
                                Del_StudEnterTemp2(s_eSETID, sqlTrans)
                            Case Else
                                Common.MessageBox(Me, "暫不提供刪除功能()!!")
                        End Select
                        sqlTrans.Commit()

                        Dim i_tmpPage As Integer = If(DataGrid2.CurrentPageIndex >= DataGrid2.PageSize, DataGrid2.PageSize - 1, DataGrid2.CurrentPageIndex)
                        Show_DataGrid2(Convert.ToString(ViewState(cst_vs_IDNO)), Convert.ToString(ViewState(cst_vs_Name)), i_tmpPage)
                    Catch ex As Exception
                        sqlTrans.Rollback()

                        'sqlTrans.Dispose()
                        Common.MessageBox(Me, ex.ToString)
                    End Try

                End If
        End Select

        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
    End Sub

    Sub SHOW_DG2_IDB_HEADER(ByRef s_Option As String, ByRef e As System.Web.UI.WebControls.DataGridItemEventArgs)
        If e Is Nothing Then Return
        If e.Item.ItemType <> ListItemType.Header Then Return

        Dim labTitle1 As Label = e.Item.FindControl("lab_Title1")
        Dim labTitle2 As Label = e.Item.FindControl("lab_Title2")
        'Dim labSNo As Label = e.Item.FindControl("lab_SNo")
        'Dim labIDNO As Label = e.Item.FindControl("lab_IDNO")
        'Dim labSID As Label = e.Item.FindControl("lab_SID")
        'Dim labName As Label = e.Item.FindControl("lab_Name")
        'Dim labBirthday As Label = e.Item.FindControl("lab_Birthday")
        'Dim labClass As Label = e.Item.FindControl("lab_Class")
        'Dim lab_Lapm As Label = e.Item.FindControl("lab_Lapm")
        'Dim btnDel As LinkButton = e.Item.FindControl("btn_Del")

        Select Case s_Option'Convert.ToString(ViewState(cst_0vs_Option))
            Case "STUDENTINFO"
                labTitle1.Text = "現有參訓班級"
                labTitle2.Text = "現有津貼申請"
            Case "ENTERTEMP"
                labTitle1.Text = "現有報名班級"
                labTitle2.Text = "其他班級資訊"
            Case "ENTERTEMP2"
                labTitle1.Text = "現有報名班級"
                labTitle2.Text = "其他班級資訊"
            Case "3"
                labTitle1.Text = "刪除學員上課資料"
                labTitle2.Text = "學員上課資料"
            Case "4"
                labTitle1.Text = "刪除報名班級LOG"
                labTitle2.Text = ""
            Case "5"
                labTitle1.Text = "刪除報名班級LOG"
                labTitle2.Text = ""
        End Select
    End Sub

    Sub SHOW_DG2_IDB_ITEM(ByRef s_Option As String, ByRef sender As System.Object, ByRef e As System.Web.UI.WebControls.DataGridItemEventArgs)
        If e Is Nothing Then Return
        If e.Item.ItemType <> ListItemType.AlternatingItem AndAlso e.Item.ItemType <> ListItemType.Item Then Return

        Dim dr_Data As DataRowView = e.Item.DataItem

        Dim labSNo As Label = e.Item.FindControl("lab_SNo")
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO")
        Dim labSID As Label = e.Item.FindControl("lab_SID")
        Dim labName As Label = e.Item.FindControl("lab_Name")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday")

        Dim labClass As Label = e.Item.FindControl("lab_Class")
        Dim labLapm As Label = e.Item.FindControl("lab_Lapm")
        Dim btnDel As LinkButton = e.Item.FindControl("btn_Del")

        'labSNo.Text = Convert.ToString(DataGrid2.CurrentPageIndex * DataGrid2.PageSize + e.Item.ItemIndex + 1)
        labSNo.Text = TIMS.Get_DGSeqNo(sender, e)
        labIDNO.Text = Convert.ToString(dr_Data("IDNO"))
        labName.Text = Convert.ToString(dr_Data("Name"))
        labBirthday.Text = Convert.ToString(dr_Data("Birthday"))

        If btnDel IsNot Nothing Then btnDel.Visible = False

        Select Case s_Option' Convert.ToString(ViewState(cst_vs_Option))
            Case "STUDENTINFO"
                labSID.Text = Convert.ToString(dr_Data("SID"))
                Dim dt_Class As DataTable = Get_ClassStudentsOfClass("SID", Convert.ToString(dr_Data("SID")))
                Dim v_labClass As String = ""
                If dt_Class IsNot Nothing Then
                    For Each dr As DataRow In dt_Class.Rows
                        v_labClass &= String.Concat("(", dr("OCID"), ")", dr("ClassName"), "<br />")
                    Next
                    labClass.Text = v_labClass
                End If
                If btnDel IsNot Nothing AndAlso v_labClass = "" Then
                    btnDel.Visible = True
                    Dim sCmdArg As String = ""
                    TIMS.SetMyValue(sCmdArg, "SID", Convert.ToString(dr_Data("SID")))
                    btnDel.CommandArgument = sCmdArg
                    btnDel.Attributes.Add("style", "cursor@hand")
                    btnDel.Attributes.Add("onclick", String.Concat("return confirm('確認要將【", labIDNO.Text, "】", labName.Text, "刪除嗎?');"))
                End If

            Case "ENTERTEMP"
                labSID.Text = Convert.ToString(dr_Data("SETID"))
                Dim dt_Class As DataTable = Get_StudEnterType("SETID", Convert.ToString(dr_Data("SETID")))
                Dim v_labClass As String = ""
                If dt_Class IsNot Nothing Then
                    For Each dr As DataRow In dt_Class.Rows
                        v_labClass &= String.Concat("(", dr("OCID1"), ")", dr("ClassName"), "<br />")
                    Next
                    labClass.Text = v_labClass
                End If

                Dim dt_Lapm As DataTable = Get_ClassStudentsOfClass("SETID", Convert.ToString(dr_Data("SETID")))
                Dim v_labLapm As String = ""
                If dt_Lapm IsNot Nothing Then
                    For Each dr As DataRow In dt_Lapm.Rows
                        v_labLapm &= String.Concat("(", dr("OCID"), ")", dr("ClassName"), "<br />")
                    Next
                    labLapm.Text &= String.Concat("Class_StudentsOfClass：<br />", v_labLapm)
                End If

                Dim dt_Lapm2 As DataTable = Get_StudEnterType2("SETID", Convert.ToString(dr_Data("SETID")))
                Dim v_labLapm2 As String = ""
                If dt_Lapm2 IsNot Nothing Then
                    For Each dr As DataRow In dt_Lapm2.Rows
                        v_labLapm2 &= String.Concat("(", dr("OCID1"), ")", dr("ClassName"), "<br />")
                    Next
                    labLapm.Text &= String.Concat("Stud_EnterType2：<br />", v_labLapm2)
                End If

                If btnDel IsNot Nothing AndAlso v_labClass = "" AndAlso v_labLapm = "" AndAlso v_labLapm2 = "" Then
                    btnDel.Visible = True
                    Dim sCmdArg As String = ""
                    TIMS.SetMyValue(sCmdArg, "SETID", Convert.ToString(dr_Data("SETID")))
                    btnDel.CommandArgument = sCmdArg
                    btnDel.Attributes.Add("style", "cursor@hand")
                    btnDel.Attributes.Add("onclick", String.Concat("return confirm('確認要將【", labIDNO.Text, "】", labName.Text, "刪除嗎?');"))
                End If

                Dim vsTip As String = String.Concat("SETID:", dr_Data("SETID"), vbCrLf, "eSETID:", dr_Data("eSETID"))
                TIMS.Tooltip(labClass, vsTip)
                TIMS.Tooltip(labIDNO, vsTip)

            Case "ENTERTEMP2"
                labSID.Text = Convert.ToString(dr_Data("eSETID"))
                Dim dt_Class As DataTable = Get_StudEnterType2("eSETID", Convert.ToString(dr_Data("eSETID")))
                Dim v_labClass As String = ""
                If dt_Class IsNot Nothing Then
                    For Each dr As DataRow In dt_Class.Rows
                        v_labClass &= String.Concat("(", dr("OCID1"), ")", dr("ClassName"), "<br />")
                    Next
                    labClass.Text = v_labClass
                End If
                'labLapm.Text = ""
                Dim dt_Lapm As DataTable = Get_StudEnterType("eSETID", Convert.ToString(dr_Data("eSETID")))
                Dim v_labLapm As String = ""
                If dt_Lapm IsNot Nothing Then
                    For Each dr As DataRow In dt_Lapm.Rows
                        v_labLapm &= String.Concat("(", dr("OCID1"), ")", dr("ClassName"), "<br />")
                    Next
                    labLapm.Text = String.Concat("Stud_EnterType：<br />", v_labLapm)
                End If

                If btnDel IsNot Nothing AndAlso v_labClass = "" AndAlso v_labLapm = "" Then
                    btnDel.Visible = True
                    Dim sCmdArg As String = ""
                    TIMS.SetMyValue(sCmdArg, "eSETID", Convert.ToString(dr_Data("eSETID")))
                    'TIMS.SetMyValue(sCmdArg, "IDNO", Convert.ToString(dr_Data("IDNO")))
                    btnDel.CommandArgument = sCmdArg
                    btnDel.Attributes.Add("style", "cursor@hand")
                    btnDel.Attributes.Add("onclick", String.Concat("return confirm('確認要將【", labIDNO.Text, "】", labName.Text, "刪除嗎?');"))
                End If

                Dim vsTip As String = String.Concat("SETID:", dr_Data("SETID"), vbCrLf, "eSETID:", dr_Data("eSETID"))
                TIMS.Tooltip(labClass, vsTip)
                TIMS.Tooltip(labIDNO, vsTip)

            Case "3" 'StudentInfo(DEL LOG)
                labSID.Text = Convert.ToString(dr_Data("SID"))
                btnDel.Visible = False
                hidSID.Value = Convert.ToString(dr_Data("SID"))
                hidSOCID.Value = Convert.ToString(dr_Data("SOCID"))
                Dim ss_sid As String = "&nbsp;(now)SID："
                If Convert.ToString(dr_Data("TT")) = "D" Then ss_sid = "&nbsp;(Del)SID："

                If hidSOCID.Value = "" Then
                    labClass.Text &= String.Concat(ss_sid, hidSID.Value)
                    Return
                End If

                Dim dt_Class As DataTable = GET_STUDENTSOFCLASSDELDATA("SOCID", Convert.ToString(hidSOCID.Value))
                If dt_Class IsNot Nothing AndAlso dt_Class.Rows.Count > 0 Then
                    Dim v_labClass As String = ""
                    v_labClass &= String.Concat(ss_sid, hidSID.Value, "<br />")
                    'labClass.Text = "CLASS_STUDENTSOFCLASSDELDATA：<br />"
                    For Each dr As DataRow In dt_Class.Rows
                        v_labClass &= String.Concat("&nbsp;(", dr("OCID"), ")", dr("ClassName"))
                        v_labClass &= String.Concat("<br />&nbsp;開結訓日期：", CDate(dr("STDATE")).ToString("yyyy/MM/dd"), "-", CDate(dr("FTDATE")).ToString("yyyy/MM/dd"))
                        v_labClass &= String.Concat("<br />&nbsp;SOCID：", hidSOCID.Value)
                        v_labClass &= String.Concat("<br />&nbsp;異動(刪除日期)：", CDate(dr("Modifydate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_labClass &= String.Concat("<br />&nbsp;異動者：", dr("Modifyacct"))
                        v_labClass &= "<br />"
                    Next
                    labClass.Text &= v_labClass
                End If

            Case "4" 'EnterTemp(DEL LOG)
                labSID.Text = Convert.ToString(dr_Data("SETID"))
                btnDel.Visible = False
                hsType.Value = Convert.ToString(dr_Data("sType"))
                hSETID.Value = Convert.ToString(dr_Data("SETID"))
                heSETID.Value = Convert.ToString(dr_Data("eSETID"))

                Dim vsTip As String = ""
                Select Case hsType.Value
                    Case "2"
                        vsTip += "sType@STUD_ENTERTEMPDelData" & vbCrLf
                    Case "1"
                        vsTip += "sType@STUD_ENTERTEMP" & vbCrLf
                End Select
                vsTip += "SETID:" & hSETID.Value & vbCrLf
                vsTip += "eSETID:" & heSETID.Value
                TIMS.Tooltip(labClass, vsTip)
                TIMS.Tooltip(labIDNO, vsTip)

                Dim dt_Class As DataTable = Get_StudEnterTypeDelData("SETID", Convert.ToString(hSETID.Value))
                Dim v_labClass As String = ""
                If dt_Class IsNot Nothing AndAlso dt_Class.Rows.Count > 0 Then
                    v_labClass = "Stud_EnterTypeDelData：<br />"
                    For Each dr As DataRow In dt_Class.Rows
                        v_labClass &= String.Concat("&nbsp;(", dr("OCID1"), ")", dr("ClassName"))
                        v_labClass &= String.Concat("<br />&nbsp;開結訓日期：", CDate(dr("STDATE")).ToString("yyyy/MM/dd"), "-", CDate(dr("FTDATE")).ToString("yyyy/MM/dd"))
                        v_labClass &= String.Concat("<br />&nbsp;實際報名日期：", CDate(dr("RelEnterDate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_labClass &= String.Concat("<br />&nbsp;異動(刪除日期)：", CDate(dr("Modifydate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_labClass &= String.Concat("<br />&nbsp;異動者：", Convert.ToString(dr("Modifyacct")))
                        v_labClass &= "<br />"
                    Next
                    labClass.Text = v_labClass
                End If

                'STUD_DELENTERTYPE
                Dim dt_C2 As DataTable = Get_STUD_DELENTERTYPE("SETID", Convert.ToString(hSETID.Value))
                Dim v_lab2 As String = ""
                If dt_C2 IsNot Nothing AndAlso dt_C2.Rows.Count > 0 Then
                    v_lab2 = "STUD_DELENTERTYPE：<br />"
                    For Each dr As DataRow In dt_C2.Rows
                        v_lab2 &= String.Concat("&nbsp;(", dr("OCID1"), ")", dr("ClassName"))
                        v_lab2 &= String.Concat("<br />&nbsp;開結訓日期：", CDate(dr("STDATE")).ToString("yyyy/MM/dd"), "-", CDate(dr("FTDATE")).ToString("yyyy/MM/dd"))
                        v_lab2 &= String.Concat("<br />&nbsp;實際報名日期：", CDate(dr("RelEnterDate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_lab2 &= String.Concat("<br />&nbsp;異動(刪除日期)：", CDate(dr("Modifydate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_lab2 &= String.Concat("<br />&nbsp;異動者：", Convert.ToString(dr("Modifyacct")))
                        v_lab2 &= "<br />"
                    Next
                    labClass.Text &= v_lab2
                End If

            Case "5" 'EnterTemp2(DEL LOG)
                labSID.Text = Convert.ToString(dr_Data("eSETID"))
                btnDel.Visible = False
                hsType.Value = Convert.ToString(dr_Data("sType"))
                hSETID.Value = Convert.ToString(dr_Data("SETID"))
                heSETID.Value = Convert.ToString(dr_Data("eSETID"))

                Dim vsTip As String = ""
                Select Case hsType.Value
                    Case "2"
                        vsTip &= "sType@STUD_ENTERTEMP2DelData" & vbCrLf
                    Case "1"
                        vsTip &= "sType@STUD_ENTERTEMP2" & vbCrLf
                End Select
                vsTip &= "SETID:" & hSETID.Value & vbCrLf
                vsTip &= "eSETID:" & heSETID.Value
                TIMS.Tooltip(labClass, vsTip)
                TIMS.Tooltip(labIDNO, vsTip)

                Dim dt_Class As DataTable = Get_StudEnterType2DelData("eSETID", Convert.ToString(heSETID.Value))
                Dim v_labClass As String = ""
                If dt_Class IsNot Nothing AndAlso dt_Class.Rows.Count > 0 Then
                    v_labClass = "Stud_EnterType2DelData：<br />"
                    For Each dr As DataRow In dt_Class.Rows
                        v_labClass &= String.Concat("&nbsp;(", dr("OCID1"), ")", dr("ClassName"))
                        v_labClass &= String.Concat("<br />&nbsp;開結訓日期：", CDate(dr("STDATE")).ToString("yyyy/MM/dd"), "-", CDate(dr("FTDATE")).ToString("yyyy/MM/dd"))
                        v_labClass &= String.Concat("<br />&nbsp;實際報名日期：", CDate(dr("RelEnterDate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_labClass &= String.Concat("<br />&nbsp;異動(刪除日期)：", CDate(dr("Modifydate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_labClass &= String.Concat("<br />&nbsp;異動者：", Convert.ToString(dr("Modifyacct")))
                        v_labClass &= "<br />"
                    Next
                    labClass.Text = v_labClass
                End If

                'STUD_DELENTERTYPE2
                Dim dt_C2 As DataTable = Get_STUD_DELENTERTYPE2("SETID", Convert.ToString(hSETID.Value))
                Dim v_lab2 As String = ""
                If dt_C2 IsNot Nothing AndAlso dt_C2.Rows.Count > 0 Then
                    v_lab2 = "STUD_DELENTERTYPE2：<br />"
                    For Each dr As DataRow In dt_C2.Rows
                        v_lab2 &= String.Concat("&nbsp;(", dr("OCID1"), ")", dr("ClassName"))
                        v_lab2 &= String.Concat("<br />&nbsp;開結訓日期：", CDate(dr("STDATE")).ToString("yyyy/MM/dd"), "-", CDate(dr("FTDATE")).ToString("yyyy/MM/dd"))
                        v_lab2 &= String.Concat("<br />&nbsp;實際報名日期：", CDate(dr("RelEnterDate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_lab2 &= String.Concat("<br />&nbsp;異動(刪除日期)：", CDate(dr("Modifydate")).ToString("yyyy-MM-dd HH:mm:ss"))
                        v_lab2 &= String.Concat("<br />&nbsp;異動者：", Convert.ToString(dr("Modifyacct")))
                        v_lab2 &= "<br />"
                    Next
                    labClass.Text &= v_lab2
                End If

        End Select

    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        SHOW_DG2_IDB_HEADER(Convert.ToString(ViewState(cst_vs_Option)), e)
        SHOW_DG2_IDB_ITEM(Convert.ToString(ViewState(cst_vs_Option)), sender, e)
    End Sub

    Private Sub DataGrid2_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid2.PageIndexChanged
        Show_DataGrid2(Convert.ToString(ViewState(cst_vs_IDNO)), Convert.ToString(ViewState(cst_vs_Name)), e.NewPageIndex)
    End Sub

    Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged

    End Sub
End Class
