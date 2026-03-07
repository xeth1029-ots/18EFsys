Partial Class SYS_03_010
    Inherits AuthBasePage

    Public Shared PGErrMsg1 As String = ""
    Public Shared iRst As Integer = 0
    Public Shared aNow1 As DateTime
    Public Shared sTimeMsg As String = ""

    Dim vs_IDNO As String = ""

    Const Cst_DEL1 As String = "DEL1" 'itemname
    Const Cst_DelVal1 As String = "DelVal1" 'Cst_SYS_03_009 '依系統參數判斷是否提供刪除
    Const Cst_SYS_03_009 As String = "SYS_03_009" 'spage
    'S * FROM SYS_VAR WHERE SPAGE ='SYS_03_009' AND ITEMNAME='DEL1' AND ITEMVALUE='N'
    'U SYS_VAR SET ITEMVALUE='Y' FROM SYS_VAR WHERE SPAGE ='SYS_03_009' AND ITEMNAME='DEL1' AND ITEMVALUE='N'
    'U SYS_VAR SET ITEMVALUE='N' FROM SYS_VAR WHERE SPAGE ='SYS_03_009' AND ITEMNAME='DEL1' AND ITEMVALUE='Y'

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objConn As SqlConnection

#Region "REM 1"
    'select * from e_member 
    'where mem_idno='A122778615'
    'select * from stud_studentinfo 
    'where idno='A122778615'
    'select * from stud_subdata
    'where sid ='2005102511360304'
    'select * from class_studentsofclass 
    'where sid ='2005102511360304'
    'SELECT * FROM Stud_Turnout
    'where socid in (12365
    ',66570
    ',636847
    ')
    'select * from class_classinfo 
    'where ocid in (1228
    ',3885
    ',29578
    ')
    'select * from Stud_entertemp
    'where idno='A122778615'
    'select * from Stud_entertype
    'where SETID='1874'
    'select * from Stud_entertemp2
    'where idno='A122778615'
    'select * from Stud_entertype2
    'where eSETID='81152'
#End Region

#Region "SHOW 1"
    Const cst_REC1_Class1 As Integer = 0
    Const cst_REC1_SID1 As Integer = 2
    Const cst_REC1_SETID1 As Integer = 3
    Const cst_REC2_Class2 As Integer = 0
    Const cst_REC2_SETID2 As Integer = 1
    Const cst_REC2_eSETID2 As Integer = 5
    Const cst_REC3_Class3 As Integer = 0
    Const cst_REC3_SETID3 As Integer = 2
    Const cst_REC3_eSETID3 As Integer = 3

    Private Sub Show_RdoEditClass1(ByVal tmpDT As DataTable)
        If rdo_EditClass1.Items.Count > 0 Then rdo_EditClass1.Items.Clear()
        If tmpDT IsNot Nothing Then
            For Each dr As DataRow In tmpDT.Rows
                rdo_EditClass1.Items.Add(New ListItem("OCID：" & Convert.ToString(dr("OCID")) &
                                                      "--" & Convert.ToString(dr("ClassName")) &
                                                      "；SOCID：" & Convert.ToString(dr("SOCID")) &
                                                      "；SID：" & Convert.ToString(dr("SID")) &
                                                      "；SETID：" & Convert.ToString(dr("SETID")), Convert.ToString(dr("SOCID"))))
            Next
        End If
    End Sub

    Private Sub Show_RdoEditClass2(ByVal tmpDT As DataTable)
        If rdo_EditClass2.Items.Count > 0 Then rdo_EditClass2.Items.Clear()
        If tmpDT IsNot Nothing Then
            For Each dr As DataRow In tmpDT.Rows
                Dim txCN0 As String = String.Format("OCID：{0}--{1}", dr("OCID1"), dr("ClassName"))
                Dim txCN1 As String = String.Format("；SETID：{0}；EnterDate：{1}；SerNum：{2}", dr("SETID"), dr("EnterDate"), dr("SerNum"))
                Dim txCN2 As String = String.Format("；EXAMNO：{0}；eSETID：{1}；eSerNum：{2}", dr("EXAMNO"), dr("eSETID"), dr("eSErNum"))
                Dim txCNALL As String = String.Format("{0}{1}{2}", txCN0, txCN1, txCN2)
                Dim valALL As String = String.Format("{0};{1};{2};{3}", dr("SETID"), dr("EnterDate"), dr("SerNum"), dr("CanUseYears28"))
                rdo_EditClass2.Items.Add(New ListItem(txCNALL, valALL))
            Next
        End If
    End Sub

    Private Sub Show_RdoEditClass3(ByVal tmpDT As DataTable)
        If rdo_EditClass3.Items.Count > 0 Then rdo_EditClass3.Items.Clear()
        If tmpDT IsNot Nothing Then
            For Each dr As DataRow In tmpDT.Rows
                Dim txCN0 As String = String.Format("OCID：{0}--{1}", dr("OCID1"), dr("ClassName"))
                Dim txCN1 As String = String.Format("；SerNum：{0}；SETID：{1}；eSETID：{2}", dr("eSerNum"), dr("SETID"), dr("eSETID"))
                Dim txCNALL As String = String.Format("{0}{1}", txCN0, txCN1)
                Dim valALL As String = String.Format("{0}", dr("eSerNum"))
                rdo_EditClass3.Items.Add(New ListItem(txCNALL, valALL))
            Next
        End If
    End Sub

#End Region

#Region "FUNC 1"

    Function GET_E_MEMBER(ByVal Smem_idno As String) As DataTable
        Dim Rst As New DataTable
        If Smem_idno = "" Then Return Rst

        Smem_idno = UCase(Smem_idno)
        Smem_idno = TIMS.ClearSQM(Smem_idno)
        Dim saIDNO() As String = Split(Smem_idno, ",")
        Dim sqlStr As String = ""
        sqlStr &= " SELECT convert(varchar, a.ELOGIN, 120) eloginT" & vbCrLf
        sqlStr &= " ,a.*" & vbCrLf
        sqlStr &= " FROM E_MEMBER a" & vbCrLf
        sqlStr &= " WHERE 1=1" & vbCrLf
        If saIDNO.Length > 1 Then
            sqlStr &= " AND (1!=1" & vbCrLf
            For i As Integer = 0 To saIDNO.Length - 1
                sqlStr &= " OR a.MEM_IDNO= @MEM_IDNO" & CStr(i) & vbCrLf
                'sqlStr += " or UPPER(a.MEM_IDNO)= @MEM_IDNO" & CStr(i) & vbCrLf
                'da.SelectCommand.Parameters.Add("mem_idno" & CStr(i), SqlDbType.VarChar).Value = UCase(Trim(saIDNO(i)))
            Next
            sqlStr &= " )" & vbCrLf
        Else
            If Smem_idno <> "" Then
                sqlStr &= " AND a.MEM_IDNO= @MEM_IDNO" & vbCrLf
                'sqlStr += " and UPPER(a.mem_idno)= @mem_idno"
                'da.SelectCommand.Parameters.Add("mem_idno", SqlDbType.VarChar).Value = UCase(Smem_idno)
            End If
        End If

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try

        'TIMS.OpenDbConn(objConn)
        Dim oCmd As New SqlCommand(sqlStr, objConn)
        With oCmd
            .Parameters.Clear()
            If saIDNO.Length > 1 Then
                For i As Integer = 0 To saIDNO.Length - 1
                    'sqlStr += " or UPPER(a.mem_idno)= @mem_idno" & CStr(i) & vbCrLf
                    .Parameters.Add("MEM_IDNO" & CStr(i), SqlDbType.VarChar).Value = TIMS.ClearSQM(saIDNO(i))
                Next
            Else
                If Smem_idno <> "" Then
                    'sqlStr += " and UPPER(a.mem_idno)= @mem_idno"
                    .Parameters.Add("MEM_IDNO", SqlDbType.VarChar).Value = TIMS.ClearSQM(Smem_idno)
                End If
            End If
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Function Get_StudTurnout(ByVal sIDNO As String) As DataTable
        Dim Rst As New DataTable
        If sIDNO = "" Then Return Rst

        sIDNO = UCase(sIDNO)
        Dim saIDNO() As String = Split(sIDNO, ",")
        Dim sql As String = ""
        sql &= " SELECT  CONVERT(varchar, st.socid)" & vbCrLf
        sql &= " +','+CONVERT(varchar, st.LeaveDate, 111)" & vbCrLf
        sql &= " +','+CONVERT(varchar, st.SeqNo) stkey" & vbCrLf
        sql &= " ,ss.IDNO" & vbCrLf
        sql &= " ,CONVERT(varchar, st.LeaveDate, 111) LeaveDate" & vbCrLf
        sql &= " ,cc.ocid,st.socid,st.LeaveID,st.SeqNo,c1.Name c1Name" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111)+'~'+CONVERT(varchar, cc.ftdate, 111) sftdate" & vbCrLf
        sql &= " ,st.Hours" & vbCrLf
        sql &= " ,cc.THours" & vbCrLf
        sql &= " ,isnull(st.c1,'x')+isnull(st.c2,'x')+isnull(st.c3,'x')+isnull(st.c4,'x')" & vbCrLf
        sql &= " +isnull(st.c5,'x')+isnull(st.c6,'x')+isnull(st.c7,'x')+isnull(st.c8,'x')" & vbCrLf
        sql &= " +isnull(st.c9,'x')+isnull(st.c10,'x')+isnull(st.c11,'x')+isnull(st.c12,'x') C12" & vbCrLf
        sql &= " ,isnull(st.TurnoutIgnore,'0') TurnoutIgnore" & vbCrLf
        sql &= " FROM Stud_Turnout st" & vbCrLf
        sql &= " join class_studentsofclass cs on st.socid =cs.socid" & vbCrLf
        sql &= " join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        sql &= " join class_classinfo cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " left join Key_Leave c1 on c1.LeaveID=st.LeaveID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If saIDNO.Length > 1 Then
            sql &= " and (1!=1" & vbCrLf
            For i As Integer = 0 To saIDNO.Length - 1
                sql &= String.Concat(" or ss.idno= @sIDNO", i, vbCrLf)
                'da.SelectCommand.Parameters.Add("sIDNO" & CStr(i), SqlDbType.VarChar).Value = UCase(Trim(saIDNO(i)))
            Next
            sql &= " )" & vbCrLf
        Else
            If sIDNO <> "" Then
                sql &= " and ss.idno= @sIDNO"
                'da.SelectCommand.Parameters.Add("sIDNO", SqlDbType.VarChar).Value = UCase(sIDNO)
            End If
        End If

        'TIMS.OpenDbConn(objConn)
        Dim oCmd As New SqlCommand(sql, objConn)
        With oCmd
            .Parameters.Clear()
            If saIDNO.Length > 1 Then
                For i As Integer = 0 To saIDNO.Length - 1
                    'sql += " or UPPER(ss.idno)= @sIDNO" & CStr(i) & vbCrLf
                    .Parameters.Add("sIDNO" & CStr(i), SqlDbType.VarChar).Value = TIMS.ClearSQM(saIDNO(i))
                Next
            Else
                If sIDNO <> "" Then
                    'sql += " and UPPER(ss.idno)= @sIDNO"
                    .Parameters.Add("sIDNO", SqlDbType.VarChar).Value = TIMS.ClearSQM(sIDNO)
                End If
            End If
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Function Get_StudTurnout2(ByVal sIDNO As String) As DataTable
        Dim Rst As New DataTable
        If sIDNO = "" Then Return Rst

        'sIDNO = UCase(sIDNO)
        Dim saIDNO() As String = Split(sIDNO, ",")

        Dim sql As String = ""
        sql &= " select st.STOID" & vbCrLf
        sql &= " ,ss.IDNO" & vbCrLf
        sql &= " ,CONVERT(varchar, st.LeaveDate, 111) LeaveDate" & vbCrLf
        sql &= " ,cc.ocid,st.socid" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111)+'~'+CONVERT(varchar, cc.ftdate, 111) sftdate" & vbCrLf
        sql &= " ,st.Hours" & vbCrLf
        sql &= " ,cc.THours" & vbCrLf
        sql &= " FROM Stud_Turnout2 st" & vbCrLf
        sql &= " join class_studentsofclass cs on st.socid =cs.socid" & vbCrLf
        sql &= " join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        sql &= " join class_classinfo cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If saIDNO.Length > 1 Then
            sql &= " and (1!=1" & vbCrLf
            For i As Integer = 0 To saIDNO.Length - 1
                saIDNO(i) = TIMS.ClearSQM(saIDNO(i))
                If saIDNO(i) <> "" Then
                    sql &= String.Concat(" or ss.idno= @sIDNO", i, vbCrLf)
                End If
                'da.SelectCommand.Parameters.Add("sIDNO" & CStr(i), SqlDbType.VarChar).Value = UCase(Trim(saIDNO(i)))
            Next
            sql &= " )" & vbCrLf
        Else
            If sIDNO <> "" Then
                sql &= " and ss.idno= @sIDNO"
                'da.SelectCommand.Parameters.Add("sIDNO", SqlDbType.VarChar).Value = UCase(sIDNO)
            End If
        End If

        'TIMS.OpenDbConn(objConn)
        Dim oCmd As New SqlCommand(sql, objConn)
        With oCmd
            .Parameters.Clear()
            If saIDNO.Length > 1 Then
                For i As Integer = 0 To saIDNO.Length - 1
                    saIDNO(i) = TIMS.ClearSQM(saIDNO(i))
                    If saIDNO(i) <> "" Then
                        .Parameters.Add("sIDNO" & CStr(i), SqlDbType.VarChar).Value = TIMS.ClearSQM(saIDNO(i))
                        'sql += " or ss.idno= @sIDNO" & CStr(i) & vbCrLf
                    End If
                    'sql += " or UPPER(ss.idno)= @sIDNO" & CStr(i) & vbCrLf
                Next
            Else
                If sIDNO <> "" Then
                    'sql += " and UPPER(ss.idno)= @sIDNO"
                    .Parameters.Add("sIDNO", SqlDbType.VarChar).Value = TIMS.ClearSQM(sIDNO)
                End If
            End If
            Rst.Load(.ExecuteReader())
        End With

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
        'If da.SelectCommand.Connection.State = ConnectionState.Open Then da.SelectCommand.Connection.Close()

        Return Rst
    End Function

    Function Get_BLIGATEDATA28(ByRef sIDNO As String) As DataTable
        Dim sql As String = ""
        sql &= " SELECT a.SB2ID" & vbCrLf '  /*PK*/" & vbCrLf
        sql &= " ,a.IDNO,a.NAME" & vbCrLf
        sql &= " ,format(a.BIRTHDAY,'yyyy/MM/dd') BIRTHDAY" & vbCrLf
        sql &= " ,a.UTYPE,a.ACTNO,a.COMNAME,a.CHANGEMODE" & vbCrLf
        sql &= " ,format(a.MDATE,'yyyy/MM/dd') MDATE" & vbCrLf
        sql &= " ,a.SALARY" & vbCrLf
        sql &= " ,a.DEPARTMENT,a.BIEF" & vbCrLf
        sql &= " ,format(a.MODIFYDATE,'yyyy/MM/dd') MODIFYDATE" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " FROM STUD_BLIGATEDATA28 a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.SOCID=a.SOCID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.OCID=cs.OCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sIDNO.IndexOf(",") > -1 Then
            '有逗號 用IN
            sql &= String.Concat(" AND a.IDNO IN (", TIMS.CombiSQM2IN(sIDNO), ") ")
        Else
            sql &= String.Concat(" AND a.IDNO like '", TIMS.ClearSQM(sIDNO), "'+'%' ")
        End If
        Dim Rst As New DataTable
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Function Get_STUDBLACKLIST(ByRef sIDNO As String) As DataTable
        Dim Rst As New DataTable
        If sIDNO = "" Then Return Rst

        Dim sql As String = ""
        'sql &= " SELECT TOP 500" & vbCrLf
        sql &= " SELECT a.SBSN,a.IDNO" & vbCrLf
        sql &= " ,a.SBSDATE" & vbCrLf ' 處分起日" & vbCrLf
        sql &= " ,a.SBYEARS" & vbCrLf ' 年限" & vbCrLf
        sql &= " ,a.SBCOMMENT" & vbCrLf ' 事由" & vbCrLf
        sql &= " ,a.SBTERMS" & vbCrLf ' 處分緣由" & vbCrLf
        'sql &= " ,a.AVAIL,a.MODIFYACCT,a.MODIFYDATE,a.SBNUM,a.OCID,a.RID,a.DISTID,a.TPLANID,a.NAME" & vbCrLf
        sql &= " FROM STUD_BLACKLIST a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sIDNO.IndexOf(",") > -1 Then
            '有逗號 用IN
            sql &= String.Concat(" AND a.IDNO IN (", TIMS.CombiSQM2IN(sIDNO), ") ")
        Else
            sql &= String.Concat(" AND a.IDNO like '", TIMS.ClearSQM(sIDNO), "'+'%' ")
        End If
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function

    Function Get_STUDSELRESULTBLI(ByVal sIDNO As String) As DataTable
        Dim Rst As New DataTable
        If sIDNO = "" Then Return Rst

        Dim sql As String = ""
        sql &= " select st.SB3ID" & vbCrLf
        sql &= " ,st.IDNO" & vbCrLf
        sql &= " ,st.NAME" & vbCrLf
        sql &= " ,st.COMNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, st.ENTERDATE, 111) ENTERDATE" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111)+'~'+CONVERT(varchar, cc.ftdate, 111) sftdate" & vbCrLf
        sql &= " FROM STUD_SELRESULTBLI st" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.ocid =st.ocid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If sIDNO.IndexOf(",") > -1 Then
            '有逗號 用IN
            sql &= " AND IDNO IN (" & TIMS.CombiSQM2IN(sIDNO) & ") "
        Else
            sql &= " AND IDNO like '" & TIMS.ClearSQM(sIDNO) & "'+'%' "
        End If
        Dim sCmd As New SqlCommand(sql, objConn)
        'TIMS.OpenDbConn(objConn)
        With sCmd
            .Parameters.Clear()
            Rst.Load(.ExecuteReader())
        End With
        Return Rst
    End Function


    Function ChkSubDATA(ByVal SID As String) As Boolean
        Dim Rst As Boolean = False
        SID = TIMS.ClearSQM(SID)
        If SID = "" Then Return Rst

        Dim hPMS As New Hashtable From {{"SID", SID}}
        Dim sql As String = "SELECT 'x' FROM STUD_SUBDATA WHERE SID=@SID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objConn, hPMS)
        If dr IsNot Nothing Then Rst = True
        Return Rst
    End Function

    Private Function Get_StudStudentInfo(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        Dim oDt As New DataTable 'SqlDataAdapter
        tmpIDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(tmpIDNO))
        tmpName = TIMS.ClearSQM(tmpName)
        If tmpIDNO = "" AndAlso tmpName = "" Then Return oDt

        Dim sqlStr As String = ""
        sqlStr = " select SID,IDNO,Name,CONVERT(varchar, Birthday, 111) Birthday from Stud_StudentInfo" & vbCrLf
        sqlStr &= " where 1=1 "
        If tmpIDNO.IndexOf(",") > -1 Then
            '有逗號 用IN
            sqlStr &= " AND IDNO IN (" & TIMS.CombiSQM2IN(tmpIDNO) & ") "
        Else
            sqlStr &= " AND IDNO like '" & tmpIDNO & "'+'%' "
        End If
        If tmpName <> "" Then
            sqlStr &= " AND Name like '%'+@Name+'%' "
            '.SelectCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = tmpName
        End If
        sqlStr &= " ORDER BY SID ASC  "
        Dim sCmd As New SqlCommand(sqlStr, objConn)

        'Call TIMS.OpenDbConn(objConn)
        With sCmd
            .Parameters.Clear()
            If tmpName <> "" Then
                'sqlStr += " and Name like '%'+@Name+'%' "
                .Parameters.Add("Name", SqlDbType.NVarChar).Value = tmpName
            End If
            oDt.Load(.ExecuteReader())
        End With
        If oDt.Rows.Count = 0 Then oDt = Nothing
        Return oDt
    End Function

    Private Function Get_StudEnterTemp(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        Dim sqlAdp As New SqlDataAdapter
        'Dim objDS As New DataSet

        Dim rst As DataTable = Nothing
        Dim sIDNO() As String = Split(tmpIDNO, ",")

        Dim oDt As New DataTable
        If tmpIDNO = "" AndAlso tmpName = "" Then Return oDt

        Dim sqlStr As String = ""
        sqlStr &= " SELECT SETID,IDNO,Name,eSETID" & vbCrLf
        sqlStr &= " ,CONVERT(varchar, Birthday, 111) Birthday" & vbCrLf
        sqlStr &= " FROM STUD_ENTERTEMP " & vbCrLf
        sqlStr &= " where 1=1" & vbCrLf
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objConn)
            .SelectCommand.Parameters.Clear()
            If tmpIDNO <> "" Then
                If sIDNO.Length > 1 Then
                    Dim cntIDNO As Integer = 1
                    .SelectCommand.CommandText += " and IDNO in ("
                    For Each itm As String In sIDNO
                        .SelectCommand.CommandText += If(cntIDNO > 1, ",", "") & "@IDNO" & cntIDNO
                        .SelectCommand.Parameters.Add("IDNO" & cntIDNO, SqlDbType.VarChar).Value = UCase(itm)
                        cntIDNO += 1
                    Next
                    .SelectCommand.CommandText += ") "
                Else
                    .SelectCommand.CommandText += " and IDNO like @IDNO+'%' "
                    .SelectCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO)
                End If
            End If
            If tmpName <> "" Then
                .SelectCommand.CommandText += " and Name like '%'+@Name+'%' "
                .SelectCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = tmpName
            End If
            .SelectCommand.CommandText += " ORDER BY SETID asc "
            .Fill(oDt) '.Fill(objDS, "Data")
        End With
        If oDt.Rows.Count > 0 Then rst = oDt
        Return rst
    End Function

    Private Function Get_StudEnterTemp2(ByVal tmpIDNO As String, ByVal tmpName As String) As DataTable
        Dim Rst As DataTable = Nothing
        If tmpIDNO = "" AndAlso tmpName = "" Then Return Rst

        Dim sqlAdp As SqlDataAdapter = TIMS.GetOneDA(objConn)
        Dim sIDNO() As String = Split(tmpIDNO, ",")
        Dim sqlStr As String = ""
        sqlStr &= " select SETID,eSETID,IDNO,Name"
        sqlStr &= " ,CONVERT(varchar, Birthday, 111) Birthday"
        sqlStr &= " from Stud_EnterTemp2 " & vbCrLf
        sqlStr &= " where 1=1 "
        With sqlAdp
            .SelectCommand.CommandText = sqlStr ' = New SqlCommand(sqlStr, objConn)
            .SelectCommand.Parameters.Clear()
            If tmpIDNO <> "" Then
                If sIDNO.Length > 1 Then
                    Dim cntIDNO As Integer = 1
                    .SelectCommand.CommandText += " and IDNO in ("
                    For Each itm As String In sIDNO
                        .SelectCommand.CommandText += If(cntIDNO > 1, ",", "") & "@IDNO" & cntIDNO
                        .SelectCommand.Parameters.Add("IDNO" & cntIDNO, SqlDbType.VarChar).Value = UCase(itm)
                        cntIDNO += 1
                    Next
                    .SelectCommand.CommandText += ") "
                Else
                    .SelectCommand.CommandText += " and IDNO like @IDNO+'%' "
                    .SelectCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO)
                End If
            End If
            If tmpName <> "" Then
                .SelectCommand.CommandText += " and Name like '%'+@Name+'%' "
                .SelectCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = tmpName
            End If
            .SelectCommand.CommandText += "ORDER BY eSETID asc "

            Rst = New DataTable
            .Fill(Rst)
        End With

        If Rst IsNot Nothing AndAlso Rst.Rows.Count = 0 Then
            Rst = Nothing
        End If

        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    'objConn.Close()
        '    'If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    'If Not Rst Is Nothing Then Rst.Dispose()
        '    'If Not objTrans Is Nothing Then objTrans.Dispose()
        'End Try
        Return Rst
    End Function

    Private Function Get_ClassStudentsOfClass(ByVal columnName As String, ByVal tmpID As String) As DataTable
        Dim Rst As DataTable = Nothing
        Dim sqlAdp As SqlDataAdapter = TIMS.GetOneDA(objConn)
        Dim saTMPID() As String = Split(tmpID, ",")
        Dim columns() As String = {"SID", "SETID", "IDNO"} '檢查columns
        sqlAdp.SelectCommand.Parameters.Clear()
        If Array.IndexOf(columns, columnName) <> -1 Then
            Dim sqlStr As String = ""
            sqlStr &= " select a.SID,a.SOCID,a.OCID,a.SETID" & vbCrLf
            sqlStr &= " ,a.StudStatus,a.CreditPoints ,ss.IDNO ,ss.Name" & vbCrLf
            sqlStr &= " ,CONVERT(varchar, a.EnterDate, 111) as EnterDate" & vbCrLf
            sqlStr &= " ,CONVERT(varchar, b.STDate, 111) as STDate" & vbCrLf
            sqlStr &= " ,CONVERT(varchar, b.FTDate, 111) as FTDate" & vbCrLf
            sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
            sqlStr &= " ,b.PlanID,vp.TPlanID,vp.PlanName,vp.YEARS" & vbCrLf
            sqlStr &= " ,vo.OrgName" & vbCrLf
            sqlStr &= " ,CONVERT(varchar, a.CloseDate, 111) CloseDate" & vbCrLf
            sqlStr &= " ,CONVERT(varchar, a.RejectTDate1, 111) RejectTDate1" & vbCrLf
            sqlStr &= " ,CONVERT(varchar, a.RejectTDate2, 111) RejectTDate2" & vbCrLf
            'Datagrid6 : CanUseYears28
            sqlStr &= " ,case when vp.Years>='2021' AND vp.TPlanID='28' AND a.Modifydate >=getdate()-400 then 'Y' END CanUseYears28" & vbCrLf
            sqlStr &= " from Class_StudentsOfClass a" & vbCrLf
            sqlStr &= " join Class_ClassInfo b on b.OCID=a.OCID" & vbCrLf
            sqlStr &= " join View_LoginPlan vp on vp.PlanID=b.PlanID" & vbCrLf
            sqlStr &= " join view_orgplaninfo vo on vo.RID=b.RID" & vbCrLf
            sqlStr &= " LEFT JOIN stud_studentinfo ss on ss.SID=a.SID" & vbCrLf
            sqlStr &= " WHERE 1=1" & vbCrLf
            If saTMPID.Length > 1 Then
                Select Case UCase(columnName)
                    Case "SID", "SETID"
                        sqlStr &= " and a." & columnName & " in (" & vbCrLf
                    Case "IDNO"
                        sqlStr &= " and ss." & columnName & " in (" & vbCrLf
                End Select
                For i As Integer = 0 To saTMPID.Length - 1
                    sqlStr += If(i = 0, "", ",")
                    sqlStr += "@tmpID" & CStr(i)
                    Select Case UCase(columnName)
                        Case "SID", "IDNO"
                            sqlAdp.SelectCommand.Parameters.Add("tmpID" & CStr(i), SqlDbType.VarChar).Value = UCase(saTMPID(i))
                        Case Else '"SETID"
                            sqlAdp.SelectCommand.Parameters.Add("tmpID" & CStr(i), SqlDbType.VarChar).Value = saTMPID(i)
                    End Select
                Next
                sqlStr += ")" & vbCrLf
            Else
                Select Case UCase(columnName)
                    Case "SID", "SETID"
                        sqlStr &= " and a." & columnName & "= @tmpID " & vbCrLf
                    Case "IDNO"
                        sqlStr &= " and ss." & columnName & "= @tmpID " & vbCrLf
                End Select
                Select Case UCase(columnName)
                    Case "SID", "IDNO"
                        sqlAdp.SelectCommand.Parameters.Add("tmpID", SqlDbType.VarChar).Value = UCase(tmpID)
                    Case Else '"SETID"
                        sqlAdp.SelectCommand.Parameters.Add("tmpID", SqlDbType.Int).Value = tmpID
                End Select
            End If
            sqlStr &= " ORDER BY a.SOCID asc" & vbCrLf

            sqlAdp.SelectCommand.CommandText = sqlStr

            Rst = New DataTable
            sqlAdp.Fill(Rst)

            If Rst IsNot Nothing AndAlso Rst.Rows.Count = 0 Then
                Rst = Nothing
            End If

        End If

        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    'objConn.Close()
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    If Not Rst Is Nothing Then Rst.Dispose()
        '    'If Not objTrans Is Nothing Then objTrans.Dispose()
        'End Try
        Return Rst
    End Function

    Private Function Get_StudEnterType(ByVal columnName As String, ByVal tmpID As String) As DataTable
        Dim sqlAdp As New SqlDataAdapter
        Dim objDS As New DataSet

        Dim rst As DataTable = Nothing
        Dim columns() As String = {"SETID", "eSETID"}
        If Array.IndexOf(columns, columnName) <> -1 Then
            Dim sqlStr As String = ""
            sqlStr &= " select a.SETID,CONVERT(varchar, a.EnterDate, 111) EnterDate" & vbCrLf
            sqlStr &= " ,a.SerNum,a.OCID1,a.eSETID,a.eSerNum,a.EXAMNO" & vbCrLf
            sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
            sqlStr &= " ,b.ExamPeriod,k1.EPName ,f.Admission,f.SelResultID" & vbCrLf
            'Datagrid6 : CanUseYears28
            sqlStr &= " ,case when vp.Years>='2021' AND vp.TPlanID='28' AND a.Modifydate >=getdate()-400 then 'Y' END CanUseYears28" & vbCrLf
            sqlStr &= " FROM STUD_ENTERTYPE a" & vbCrLf
            sqlStr &= " JOIN CLASS_CLASSINFO b on b.OCID=a.OCID1" & vbCrLf
            sqlStr &= " LEFT JOIN VIEW_PLAN vp on vp.PlanID=b.PlanID" & vbCrLf
            sqlStr &= " LEFT JOIN KEY_EXAMPERIOD k1 on k1.EPID=b.ExamPeriod" & vbCrLf
            sqlStr &= " LEFT JOIN STUD_SELRESULT f ON a.SETID=f.SETID and a.EnterDate=f.EnterDate and a.SerNum=f.SerNum" & vbCrLf
            sqlStr &= " WHERE a." & columnName & "= @tmpID" & vbCrLf
            sqlStr &= " ORDER BY a.OCID1,a.SETID,a.EnterDate asc,a.SerNum,a.eSETID,a.eSerNum"
            With sqlAdp
                .SelectCommand = New SqlCommand(sqlStr, objConn)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("tmpID", SqlDbType.Int).Value = tmpID
                .Fill(objDS, "Data")
            End With
            If objDS.Tables("Data") IsNot Nothing AndAlso objDS.Tables("Data").Rows.Count > 0 Then
                rst = objDS.Tables("Data")
            End If
        End If
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    If Not objDS Is Nothing Then objDS.Dispose()
        'End Try
        Return rst
    End Function

    Private Function Get_StudEnterType2(ByVal columnName As String, ByVal tmpID As String) As DataTable
        Dim sqlAdp As New SqlDataAdapter
        Dim objDS As New DataSet

        Dim rst As DataTable = Nothing
        Dim columns() As String = {"SETID", "eSETID"}
        If Array.IndexOf(columns, columnName) <> -1 Then
            Dim sqlStr As String = ""
            sqlStr &= " select a.SETID,a.eSETID,a.eSerNum,a.OCID1,a.signUpStatus" & vbCrLf
            sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
            sqlStr &= " ,b.ExamPeriod,k1.EPName" & vbCrLf
            'Datagrid6 : CanUseYears28
            sqlStr &= " ,case when vp.Years>='2021' AND vp.TPlanID='28' AND a.Modifydate >=getdate()-400 then 'Y' END CanUseYears28" & vbCrLf
            sqlStr &= " FROM STUD_ENTERTYPE2 a" & vbCrLf
            sqlStr &= " JOIN CLASS_CLASSINFO b on b.OCID=a.OCID1" & vbCrLf
            sqlStr &= " LEFT JOIN VIEW_PLAN vp on vp.PlanID=b.PlanID" & vbCrLf
            sqlStr &= " LEFT JOIN KEY_EXAMPERIOD k1 on k1.EPID=b.ExamPeriod" & vbCrLf
            sqlStr &= " where a." & columnName & "= @tmpID" & vbCrLf
            sqlStr &= " ORDER BY a.OCID1,a.eSETID,a.eSerNum,a.SETID,a.EnterDate asc "
            With sqlAdp
                .SelectCommand = New SqlCommand(sqlStr, objConn)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("tmpID", SqlDbType.Int).Value = tmpID
                .Fill(objDS, "Data")
            End With
            If objDS.Tables("Data") IsNot Nothing AndAlso objDS.Tables("Data").Rows.Count > 0 Then
                rst = objDS.Tables("Data")
            End If
        End If
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    If Not objDS Is Nothing Then objDS.Dispose()
        'End Try
        Return rst
    End Function

    Private Function Get_SubSubSidyApply(ByVal columnName As String, ByVal tmpID As String) As DataTable
        'Dim sqlAdp As New SqlDataAdapter
        'Dim objDS As New DataSet
        TIMS.OpenDbConn(objConn)
        Dim rst As New DataTable '= Nothing
        Dim sqlStr As String = ""
        sqlStr &= " select a.SID,a.OCID,a.SUBID,a.AppliedStatusF,a.AppliedStatusFin" & vbCrLf
        sqlStr &= " ,dbo.FN_GET_CLASSCNAME(b.CLASSCNAME,b.CYCLTYPE) CLASSNAME" & vbCrLf
        sqlStr &= " from Sub_SubSidyApply a" & vbCrLf
        sqlStr &= " left join Class_ClassInfo b on b.OCID=a.OCID" & vbCrLf
        sqlStr &= " where " & columnName & "= @IDX" & vbCrLf
        sqlStr &= " ORDER BY a.SUBID ASC"
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("IDX", SqlDbType.VarChar).Value = tmpID
            rst.Load(.ExecuteReader())
        End With
        If rst IsNot Nothing And rst.Rows.Count = 0 Then rst = Nothing
        Return rst
    End Function

    Private Function Get_StudSubSidyCost(ByVal tmpID As Integer) As DataTable
        'Dim sqlAdp As New SqlDataAdapter
        'Dim objDS As New DataSet
        TIMS.OpenDbConn(objConn)
        Dim rst As New DataTable '= Nothing
        Dim sqlStr As String = ""
        sqlStr &= " select SOCID,AppliedStatus,AppliedStatusM,SumOfMoney" & vbCrLf
        sqlStr &= " ,CONVERT(varchar, AllotDate, 111) AllotDate" & vbCrLf
        sqlStr &= " FROM STUD_SUBSIDYCOST" & vbCrLf
        sqlStr &= " WHERE SOCID=@SOCID" & vbCrLf
        Dim sCmd As New SqlCommand(sqlStr, objConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = tmpID
            rst.Load(.ExecuteReader())
        End With
        If rst IsNot Nothing And rst.Rows.Count = 0 Then rst = Nothing
        Return rst
    End Function

    Private Function Get_SID(ByVal tmpIDNO As String) As DataTable
        Dim sqlAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim rst As DataTable = Nothing
        tmpIDNO = TIMS.ClearSQM(UCase(tmpIDNO))
        If tmpIDNO = "" Then Return rst

        Dim sIDNO() As String = Split(tmpIDNO, ",")
        Dim sqlStr As String = ""
        sqlStr = "select distinct SID from Stud_StudentInfo where 1=1 "
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objConn)
            .SelectCommand.Parameters.Clear()
            If sIDNO.Length > 1 Then
                Dim cntIDNO As Integer = 1
                .SelectCommand.CommandText += " and IDNO in ("
                For Each itm As String In sIDNO
                    .SelectCommand.CommandText += If(cntIDNO > 1, ",", "") & "@IDNO" & cntIDNO
                    .SelectCommand.Parameters.Add("IDNO" & cntIDNO, SqlDbType.VarChar).Value = UCase(itm)
                    cntIDNO += 1
                Next
                .SelectCommand.CommandText += ") "
            Else
                .SelectCommand.CommandText += " and IDNO like @IDNO+'%' "
                .SelectCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO)
            End If
            .SelectCommand.CommandText += " ORDER BY SID asc"
            .Fill(objDS, "Data")
        End With
        If objDS.Tables("Data") IsNot Nothing AndAlso objDS.Tables("Data").Rows.Count > 0 Then
            rst = objDS.Tables("Data")
        End If
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    If Not objDS Is Nothing Then objDS.Dispose()
        'End Try
        Return rst
    End Function

    Private Function Get_SETID(ByVal tmpIDNO As String, ByVal source As String) As DataTable
        Dim sqlAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim Rst As DataTable = Nothing
        If tmpIDNO = "" Then Return Rst

        tmpIDNO = UCase(tmpIDNO)
        Dim sIDNO() As String = Split(tmpIDNO, ",")
        Dim sqlStr As String = ""
        Select Case UCase(source)
            Case "STUD_STUDENTINFO"
                sqlStr = "select distinct b.SETID from Stud_StudentInfo a join Class_StudentsOfClass b on b.SID=a.SID where 1=1 "
            Case "STUD_ENTERTEMP"
                sqlStr = "select distinct b.SETID from Stud_EnterTemp a join Stud_EnterType b on b.SETID=a.SETID where 1=1 "
            Case "STUD_ENTERTEMP2"
                sqlStr = "select distinct b.SETID from Stud_EnterTemp2 a join Stud_EnterType2 b on b.eSETID=a.eSETID where 1=1 "
        End Select
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objConn)
            .SelectCommand.Parameters.Clear()
            If sIDNO.Length > 1 Then
                Dim i_cntIDNO As Integer = 1
                .SelectCommand.CommandText += " and a.IDNO in ("
                For Each itm As String In sIDNO
                    .SelectCommand.CommandText += If(i_cntIDNO > 1, ",", "") & "@IDNO" & i_cntIDNO
                    .SelectCommand.Parameters.Add("IDNO" & i_cntIDNO, SqlDbType.VarChar).Value = UCase(itm)
                    i_cntIDNO += 1
                Next
                .SelectCommand.CommandText += ") "
            Else
                .SelectCommand.CommandText += " and a.IDNO like @IDNO+'%' "
                .SelectCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO)
            End If

            .SelectCommand.CommandText += "ORDER BY b.SETID asc"
            .Fill(objDS, "Data")
        End With
        If objDS.Tables("Data") IsNot Nothing AndAlso objDS.Tables("Data").Rows.Count > 0 Then
            Rst = objDS.Tables("Data")
        End If
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    If Not objDS Is Nothing Then objDS.Dispose()
        '    If Not Rst Is Nothing Then Rst.Dispose()
        'End Try
        Return Rst
    End Function

    Private Function Get_eSETID(ByVal tmpIDNO As String, ByVal source As String) As DataTable
        Dim sqlAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim rst As DataTable = Nothing
        If tmpIDNO = "" Then Return rst

        Dim sIDNO() As String = Split(tmpIDNO, ",")
        Dim sqlStr As String = ""
        Select Case UCase(source)
            Case "STUD_ENTERTEMP"
                sqlStr = "select distinct b.eSETID from Stud_EnterTemp a join Stud_EnterType b on b.SETID=a.SETID where 1=1 "
            Case "STUD_ENTERTEMP2"
                sqlStr = "select distinct b.eSETID from Stud_EnterTemp2 a join Stud_EnterType2 b on b.eSETID=a.eSETID where 1=1 "
        End Select
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objConn)
            .SelectCommand.Parameters.Clear()
            If sIDNO.Length > 1 Then
                Dim cntIDNO As Integer = 1
                .SelectCommand.CommandText += " and a.IDNO in ("
                For Each itm As String In sIDNO
                    .SelectCommand.CommandText += If(cntIDNO > 1, ",", "") & "@IDNO" & cntIDNO
                    .SelectCommand.Parameters.Add("IDNO" & cntIDNO, SqlDbType.VarChar).Value = UCase(itm)
                    cntIDNO += 1
                Next
                .SelectCommand.CommandText += ") "
            Else
                .SelectCommand.CommandText += " and a.IDNO like @IDNO+'%' "
                .SelectCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(tmpIDNO)
            End If
            .SelectCommand.CommandText += "ORDER BY b.eSETID asc"
            .Fill(objDS, "Data")
        End With
        If objDS.Tables("Data") IsNot Nothing AndAlso objDS.Tables("Data").Rows.Count > 0 Then
            rst = objDS.Tables("Data")
        End If
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        '    If Not objDS Is Nothing Then objDS.Dispose()
        '    If Not rst Is Nothing Then rst.Dispose()
        'End Try
        Return rst
    End Function

#Region "NO USE"
    'Private Function Get_SUBID(ByVal columnName As String, ByVal tmpID As String) As DataTable
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim objDS As New DataSet
    '    Dim sqlStr As String
    '    Dim rst As DataTable = Nothing

    '    Try
    '        sqlStr = "select SUBID,SID,SOCID,AppliedStatusF,AppliedStatusFin from Sub_SubSidyApply where " & columnName & "= @ID "
    '        With sqlAdp
    '            .SelectCommand = New SqlCommand(sqlStr, objConn)
    '            .SelectCommand.Parameters.Clear()
    '            .SelectCommand.Parameters.Add("ID", SqlDbType.VarChar).Value = tmpID
    '            .Fill(objDS, "Data")
    '        End With
    '        If Not objDS.Tables("Data") Is Nothing Then
    '            If objDS.Tables("Data").Rows.Count > 0 Then
    '                rst = objDS.Tables("Data")
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        objConn.Close()
    '        sqlAdp.Dispose()
    '        objDS.Dispose()
    '    End Try
    '    Return rst
    'End Function
#End Region

    '共用顯示DG
    Public Shared Sub sUtl_ShowDGData1(ByRef labMsg As Label, ByRef objDG As DataGrid,
                                       ByRef dtS As DataTable, ByVal tmpPage As Integer, ByVal sDKField As String)
        sTimeMsg = ":" & DateDiff(DateInterval.Second, aNow1, Now) & "秒"
        labMsg.Visible = True
        labMsg.Text = "查無資料" & sTimeMsg

        With objDG
            .Visible = False
            If dtS Is Nothing Then Exit Sub
            If dtS.Rows.Count = 0 Then Exit Sub
            .Visible = True
            If sDKField <> "" Then .DataKeyField = sDKField
            .CurrentPageIndex = tmpPage
            .DataSource = dtS
            .DataBind()
        End With
        labMsg.Text = sTimeMsg
        'lab_Msg1.Visible = False
    End Sub

    Private Sub Show_DataGrid1(ByVal tmpIDNO As String, ByVal tmpName As String, ByVal tmpPage As Integer)
        If tmpIDNO = "" AndAlso tmpName = "" Then Exit Sub

        Dim dt_Student As DataTable = Nothing
        dt_Student = Get_StudStudentInfo(tmpIDNO, tmpName)
        Call sUtl_ShowDGData1(lab_Msg1, DataGrid1, dt_Student, tmpPage, "SID")
    End Sub

    Private Sub Show_DataGrid2(ByVal tmpIDNO As String, ByVal tmpName As String, ByVal tmpPage As Integer)
        If tmpIDNO = "" AndAlso tmpName = "" Then Exit Sub

        Dim dt_Student As DataTable = Nothing
        dt_Student = Get_StudEnterTemp(tmpIDNO, tmpName)
        Call sUtl_ShowDGData1(lab_Msg2, Datagrid2, dt_Student, tmpPage, "SETID")
    End Sub

    Private Sub Show_DataGrid3(ByVal tmpIDNO As String, ByVal tmpName As String, ByVal tmpPage As Integer)
        If tmpIDNO = "" AndAlso tmpName = "" Then Exit Sub

        Dim dt_Student As DataTable = Nothing
        dt_Student = Get_StudEnterTemp2(tmpIDNO, tmpName)
        Call sUtl_ShowDGData1(lab_Msg3, Datagrid3, dt_Student, tmpPage, "eSETID")
    End Sub

    Private Sub Show_DataGrid4(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql &= " select cc.ocid, cs.SOCID, rr.courid, ss.name, ss.idno, cc.classcname, rr.results " & vbCrLf
        sql &= " ,cr.CourseName" & vbCrLf
        sql &= " from stud_studentinfo ss" & vbCrLf
        sql &= " join class_studentsofclass cs on cs.SID =ss.SID" & vbCrLf
        sql &= " join class_classinfo cc on cc.OCID =cs.OCID" & vbCrLf
        sql &= " join Stud_TrainingResults rr on rr.socid =cs.socid" & vbCrLf
        sql &= " JOIN Course_CourseInfo cr on cr.courid=rr.courid" & vbCrLf
        sql &= " where ss.idno='" & TIMS.ClearSQM(tmpIDNO) & "'" & vbCrLf
        sql &= " ORDER BY cc.ocid , cs.SOCID,rr.courid" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objConn)
        Call sUtl_ShowDGData1(lab_Msg4, Datagrid4, dt, tmpPage, "")
    End Sub

    Private Sub Show_DataGrid5(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dt As DataTable = Nothing
        Dim sqlstr As String = ""
        sqlstr &= " SELECT B.name ,A.idno ,A.TRN_CLASS ,A.TransToTIMS ,A.TIMSModifyDate" & vbCrLf
        sqlstr &= " ,A.SOCID SOCID1" & vbCrLf
        sqlstr &= " ,cs.SOCID SOCID2" & vbCrLf
        sqlstr &= " ,cc.OCID" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sqlstr &= " FROM Adp_GOVTRNData A" & vbCrLf
        sqlstr &= " LEFT JOIN Adp_StdData B ON A.IDNO = B.IDNO" & vbCrLf
        sqlstr &= " LEFT JOIN class_classinfo cc ON cc.ocid=a.TRN_CLASS" & vbCrLf
        sqlstr &= " LEFT JOIN stud_studentinfo ss on ss.idno=a.idno" & vbCrLf
        sqlstr &= " LEFT JOIN class_studentsofclass cs on cs.sid=ss.sid and cs.ocid =cc.ocid" & vbCrLf
        sqlstr &= " WHERE A.IDNO ='" & TIMS.ClearSQM(tmpIDNO) & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sqlstr, objConn)
        Call sUtl_ShowDGData1(lab_Msg5, Datagrid5, dt, tmpPage, "")
    End Sub

    Private Sub Show_DataGrid6(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dt As DataTable = Nothing
        dt = Get_ClassStudentsOfClass("IDNO", tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg6, Datagrid6, dt, tmpPage, "")
    End Sub

    Private Sub Show_DataGrid7(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dt As DataTable = Nothing
        Dim sqlstr As String = ""
        sqlstr &= " SELECT s.IDNO,s.Name" & vbCrLf
        sqlstr &= " ,c.DLID, d.SubNo" & vbCrLf
        sqlstr &= " ,CONVERT(varchar, b.RejectTDate1, 111) RejectTDate1" & vbCrLf
        sqlstr &= " ,CONVERT(varchar, b.RejectTDate2, 111) RejectTDate2" & vbCrLf
        sqlstr &= " ,b.SOCID,b.SID" & vbCrLf
        sqlstr &= " ,b.OCID" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sqlstr &= " ,b.StudentID, d.StudentID as StudentID2" & vbCrLf
        sqlstr &= " FROM Class_StudentsOfClass b" & vbCrLf
        sqlstr &= " JOIN Stud_StudentInfo s on s.SID =b.SID" & vbCrLf
        sqlstr &= " left join Stud_DataLid c on c.OCID =b.OCID" & vbCrLf
        sqlstr &= " left join Class_ClassInfo cc on cc.OCID =b.OCID" & vbCrLf
        sqlstr &= " left join Stud_ResultStudData d on d.socid =b.SOCID" & vbCrLf
        sqlstr &= " WHERE s.IDNO ='" & TIMS.ClearSQM(tmpIDNO) & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sqlstr, objConn)
        Call sUtl_ShowDGData1(lab_Msg7, Datagrid7, dt, tmpPage, "")
    End Sub

    Private Sub Show_DataGrid8(ByVal tmpIDNO As String, ByVal tmpPage As Integer)

        Dim dt As DataTable = Nothing
        Dim sqlstr As String = ""
        sqlstr &= " select  cs.SOCID,ss.name,  cc.ocid, cc.classcname" & vbCrLf
        sqlstr &= " ,rr.TechPoint,rr.RemedPoint,rr.MinusLeave,rr.MinusSanction" & vbCrLf
        sqlstr &= " from stud_studentinfo ss" & vbCrLf
        sqlstr &= " join class_studentsofclass cs on cs.SID =ss.SID" & vbCrLf
        sqlstr &= " join class_classinfo cc on cc.OCID =cs.OCID" & vbCrLf
        sqlstr &= " join Stud_Conduct rr on rr.socid =cs.socid" & vbCrLf
        sqlstr &= " WHERE ss.idno ='" & TIMS.ClearSQM(tmpIDNO) & "'" & vbCrLf
        sqlstr &= " ORDER BY 1,3,5,6,7,8" & vbCrLf
        dt = DbAccess.GetDataTable(sqlstr, objConn)
        Call sUtl_ShowDGData1(lab_Msg8, Datagrid8, dt, tmpPage, "")

    End Sub

    Private Sub Show_DataGrid9(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dt As DataTable = Nothing
        Dim sqlstr As String = ""
        sqlstr &= " select cs.SOCID,ss.name" & vbCrLf
        sqlstr &= " ,rr.OrigClassID" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(cc1.CLASSCNAME,cc1.CYCLTYPE) CLASSCNAME1" & vbCrLf
        sqlstr &= " ,rr.NewClassID" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(cc2.CLASSCNAME,cc2.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sqlstr &= " ,rr.ApplyDate,rr.Reason" & vbCrLf
        sqlstr &= " from stud_studentinfo ss" & vbCrLf
        sqlstr &= " join class_studentsofclass cs on cs.SID =ss.SID" & vbCrLf
        sqlstr &= " join Stud_TranClassRecord rr on rr.socid =cs.socid" & vbCrLf
        sqlstr &= " join class_classinfo cc1 on cc1.OCID =rr.OrigClassID" & vbCrLf
        sqlstr &= " join class_classinfo cc2 on cc2.OCID =rr.NewClassID" & vbCrLf
        sqlstr &= " WHERE ss.idno ='" & TIMS.ClearSQM(tmpIDNO) & "'" & vbCrLf
        sqlstr &= " ORDER BY 1,3,5 " & vbCrLf
        dt = DbAccess.GetDataTable(sqlstr, objConn)
        Call sUtl_ShowDGData1(lab_Msg9, Datagrid9, dt, tmpPage, "")

    End Sub

    Private Sub Show_DataGrid10(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        If tmpIDNO = "" Then Exit Sub

        Dim dtD As DataTable = Nothing
        dtD = GET_E_MEMBER(tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg10, Datagrid10, dtD, tmpPage, "mem_sn")
    End Sub

    Private Sub Show_DataGrid11(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dtD As DataTable = Nothing
        dtD = Get_StudTurnout(tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg11, Datagrid11, dtD, tmpPage, "stkey")
    End Sub

    Private Sub Show_DataGrid12(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dtD As DataTable = Nothing
        dtD = Get_StudTurnout2(tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg12, Datagrid12, dtD, tmpPage, "STOID")
    End Sub

    Private Sub Show_DataGrid13(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dtD As DataTable = Nothing
        dtD = Get_STUDSELRESULTBLI(tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg13, Datagrid13, dtD, tmpPage, "SB3ID")
    End Sub

    Private Sub Show_DataGrid14(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dtD As DataTable = Nothing
        dtD = Get_STUDBLACKLIST(tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg14, DataGrid14, dtD, tmpPage, "SBSN")
    End Sub

    Private Sub Show_DataGrid15(ByVal tmpIDNO As String, ByVal tmpPage As Integer)
        Dim dtD As DataTable = Nothing
        dtD = Get_BLIGATEDATA28(tmpIDNO)
        Call sUtl_ShowDGData1(lab_Msg15, DataGrid15, dtD, tmpPage, "SB2ID")
    End Sub

    Public Shared Sub Del_STUD_SELRESULTBLI(ByVal MyPage As Page,
                                            ByVal SB3ID As String,
                                            ByVal oConn As SqlConnection)
        'Dim sqlAdp As New SqlDataAdapter
        Const cst_COLUMN_1 As String = "SB3ID,IDNO,NAME,BIRTHDAY,UTYPE,ACTNO,COMNAME,CHANGEMODE,MDATE,SALARY,DEPARTMENT,MODIFYDATE,SETID,ENTERDATE,SERNUM,OCID,CREATEDATE,BIEF,BIEFDESC"
        Dim sqlStr As String = ""
        'Call TIMS.OpenDbConn(oConn)
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn) 'Dim oTrans As SqlTransaction = oConn.BeginTransaction()
        Try
            sqlStr = "UPDATE STUD_SELRESULTBLI SET MODIFYDATE=GETDATE() WHERE SB3ID=@SB3ID "
            Dim uCmd As New SqlCommand(sqlStr, oConn, oTrans)
            sqlStr = String.Concat("INSERT INTO STUD_SELRESULTBLIDELDATA (", cst_COLUMN_1, ") SELECT ", cst_COLUMN_1, " FROM STUD_SELRESULTBLI WHERE SB3ID=@SB3ID")
            Dim iCmd As New SqlCommand(sqlStr, oConn, oTrans)
            sqlStr = "delete STUD_SELRESULTBLI where SB3ID=@SB3ID "
            Dim dCmd As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("SB3ID", SqlDbType.VarChar).Value = SB3ID
                .ExecuteNonQuery()
            End With
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("SB3ID", SqlDbType.VarChar).Value = SB3ID
                .ExecuteNonQuery()
            End With
            With dCmd
                .Parameters.Clear()
                .Parameters.Add("SB3ID", SqlDbType.VarChar).Value = SB3ID
                .ExecuteNonQuery()
            End With
            oTrans.Commit()
        Catch ex As Exception
            oTrans.Rollback()
            Common.MessageBox(MyPage, ex.ToString)
            'objConn.Close()
            'sqlAdp.Dispose()
        End Try
    End Sub

    Sub Del_E_MEMBER(ByVal mem_sn As String, ByRef oConn As SqlConnection)
        'Dim conn As SqlConnection
        'conn = DbAccess.GetConnection()
        Const cst_COLUMN_1 As String = "MEM_SN,MEM_IDNO,MEM_PWD,MEM_NAME,MEM_FOREIGN,MEM_EDU,MEM_BIRTH,MEM_SEX,MEM_MILITARY,MEM_MARRY,MEM_GRADUATE,MEM_SCHOOL,MEM_DEPART,MEM_ZIP,MEM_ADDR,MEM_TEL,MEM_TELN,MEM_MOBILE,MEM_EMAIL,MEM_OPENSEC,MEM_MEMO,MEM_TIMS,EPAPER,ELOGIN,MEM_USR_ID,MEM_REGTIME,MEM_OPUSER,MEM_UDATE,HANDTYPEID,HANDLEVELID,MEM_ZIP2W,STOPMEM,MEM_LOGINCNT,HANDTYPEID2,HANDLEVELID2,MFLAG,MEM_ZIP6W"
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
        Try
            sqlStr = "update E_MEMBER set mem_opuser= @mem_opuser,mem_udate=getdate() where mem_sn= @mem_sn "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, oConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("mem_opuser", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("mem_sn", SqlDbType.VarChar).Value = mem_sn
                .UpdateCommand.ExecuteNonQuery()
            End With
            sqlStr = String.Concat("INSERT INTO E_MEMBERDELDATA(", cst_COLUMN_1, ") select ", cst_COLUMN_1, " from E_Member where mem_sn= @mem_sn")
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, oConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("mem_sn", SqlDbType.VarChar).Value = mem_sn
                .InsertCommand.ExecuteNonQuery()
            End With
            sqlStr = "delete E_MEMBER where mem_sn= @mem_sn "
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, oConn, oTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("mem_sn", SqlDbType.VarChar).Value = mem_sn
                .DeleteCommand.ExecuteNonQuery()
            End With
            DbAccess.CommitTrans(oTrans)
            'Call TIMS.CloseDbConn(conn)
            'If conn.State = ConnectionState.Open Then conn.Close()
            Common.MessageBox(Me, "刪除完成。")
        Catch ex As Exception
            DbAccess.RollbackTrans(oTrans)
            Common.MessageBox(Me, "刪除失敗。" & vbCrLf & ex.ToString & vbCrLf)
        End Try
        If oTrans IsNot Nothing Then oTrans = Nothing
    End Sub

    Sub Stop_E_MEMBER(ByVal mem_sn As String, i_TYPE As Integer)
        'i_TYPE : 1:stop , 2.unstop
        Try
            Dim pms_S1 As New Hashtable From {{"mem_sn", mem_sn}}
            Dim sqlStr_S1 As String = "SELECT mem_memo FROM E_MEMBER where mem_sn= @mem_sn"
            Dim drM1 As DataRow = DbAccess.GetOneRow(sqlStr_S1, objConn, pms_S1)
            Dim s_mem_memo As String = If(drM1 IsNot Nothing, Convert.ToString(drM1("mem_memo")), "")
            If s_mem_memo.Length > 0 Then s_mem_memo &= ","

            If i_TYPE = 1 Then
                Dim pms_u1 As New Hashtable From {
                {"mem_pwd", Left(TIMS.GetGUID(), 30)},
                {"mem_memo", String.Concat(s_mem_memo, "該帳號停用於", Common.FormatNow())},
                {"mem_opuser", sm.UserInfo.UserID},
                {"mem_sn", mem_sn}}
                Dim sqlStr_U1 As String = "update E_MEMBER set stopmem='Y',mem_memo= @mem_memo,mem_pwd= @mem_pwd, mem_opuser= @mem_opuser,mem_udate=GETDATE() where mem_sn= @mem_sn"
                DbAccess.ExecuteNonQuery(sqlStr_U1, objConn, pms_u1)
                Common.MessageBox(Me, "停用完成。")
                Return
            ElseIf i_TYPE = 2 Then
                Dim pms_u2 As New Hashtable From {
                {"mem_memo", String.Concat(s_mem_memo, "帳號啟用", Common.FormatNow())},
                {"mem_opuser", sm.UserInfo.UserID},
                {"mem_sn", mem_sn}}
                Dim sqlStr_U2 As String = "update E_MEMBER set stopmem=null ,mem_memo= @mem_memo ,mem_opuser= @mem_opuser,mem_udate=GETDATE() where mem_sn= @mem_sn"
                DbAccess.ExecuteNonQuery(sqlStr_U2, objConn, pms_u2)
                Common.MessageBox(Me, "已啟用。")
                Return
            End If
        Catch ex As Exception
            Common.MessageBox(Me, "停用失敗。" & vbCrLf & ex.ToString & vbCrLf)
        End Try
    End Sub

    Private Sub Clear_Edit()
        txt_IDNO.Text = TIMS.ClearSQM(txt_IDNO.Text)
        If vs_IDNO = "" Then vs_IDNO = txt_IDNO.Text

        tr_Info.Visible = True
        tr_Edit1.Visible = False
        tr_Edit2.Visible = False
        tr_Edit3.Visible = False
        tr_Edit10.Visible = False

        If vs_IDNO = "" Then Exit Sub
        Show_DataGrid1(vs_IDNO, "", 0)
        Show_DataGrid2(vs_IDNO, "", 0)
        Show_DataGrid3(vs_IDNO, "", 0)
        Show_DataGrid10(vs_IDNO, 0)
        Call QuerySearch1()
    End Sub

    'Public Shared Sub UPDATE_GOVTRNDATA(ByRef oConn As SqlConnection, ByVal OCID As String, ByVal IDNO As String, ByVal SOCIDValue As String)
    '    'Dim sql As String = ""
    '    'Dim dt As DataTable = Nothing
    '    'Dim da As SqlDataAdapter = Nothing
    '    'Dim dr As DataRow = Nothing

    '    '學習券新增，更新三合一資料
    '    Dim OrgName As String = ""
    '    Dim ComIDNO As String = ""
    '    Dim ContactName As String = ""
    '    Dim ContactPhone As String = ""
    '    Dim THours As String = ""
    '    Dim TaddressZip As String = ""
    '    Dim TAddress As String = ""
    '    Dim EnterDate As String = ""
    '    Dim OpenDate As String = ""
    '    Dim CloseDate As String = ""
    '    Dim ClassName As String = ""

    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT a.THours,a.TaddressZip,a.TAddress" & vbCrLf
    '    sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
    '    sql &= " ,b.OrgName,b.ComIDNO,d.ContactName,d.Phone" & vbCrLf
    '    sql &= " FROM Class_ClassInfo a" & vbCrLf
    '    sql &= " JOIN Org_OrgInfo b ON a.ComIDNO=b.ComIDNO and a.OCID='" & OCID & "'" & vbCrLf
    '    sql &= " JOIN Auth_Relship c ON a.RID=c.RID" & vbCrLf
    '    sql &= " JOIN Org_OrgPlanInfo d ON c.RSID=d.RSID" & vbCrLf
    '    Dim dr As DataRow = DbAccess.GetOneRow(sql, oConn)
    '    If dr IsNot Nothing Then
    '        THours = dr("THours").ToString
    '        TaddressZip = dr("TaddressZip").ToString
    '        TAddress = dr("TAddress").ToString
    '        OrgName = dr("OrgName").ToString
    '        ComIDNO = dr("ComIDNO").ToString
    '        ContactName = dr("ContactName").ToString
    '        ContactPhone = dr("Phone").ToString
    '        ClassName = dr("ClassCName").ToString
    '    End If

    '    sql = "" & vbCrLf
    '    sql &= " SELECT c.OpenDate" & vbCrLf
    '    sql &= " ,c.CloseDate" & vbCrLf
    '    sql &= " ,c.EnterDate" & vbCrLf
    '    sql &= " ,c.IdentityID IdentityIDEX" & vbCrLf
    '    sql &= " ,c.SubsidyID SubsidyIDEX " & vbCrLf
    '    sql &= " FROM Stud_StudentInfo a " & vbCrLf
    '    sql &= " join Stud_SubData b on b.SID=a.SID " & vbCrLf
    '    sql &= " join Class_StudentsOfClass c on c.SID=a.SID " & vbCrLf
    '    sql &= " WHERE a.SID=b.SID " & vbCrLf
    '    sql &= " AND c.SID=b.SID " & vbCrLf
    '    sql &= " and c.SOCID='" & SOCIDValue & "' " & vbCrLf
    '    dr = DbAccess.GetOneRow(sql, oConn)
    '    If dr IsNot Nothing Then
    '        If Convert.ToString(dr("OpenDate")) <> "" Then
    '            OpenDate = FormatDateTime(Convert.ToString(dr("OpenDate")), DateFormat.ShortDate)
    '        End If
    '        If Convert.ToString(dr("CloseDate")) <> "" Then
    '            CloseDate = FormatDateTime(Convert.ToString(dr("CloseDate")), DateFormat.ShortDate)
    '        End If
    '        If Convert.ToString(dr("EnterDate")) <> "" Then
    '            EnterDate = FormatDateTime(Convert.ToString(dr("EnterDate")), DateFormat.ShortDate)
    '        End If
    '    End If

    '    '下列請先參考SD_03_002_add 再做修改
    '    Dim tmpTicketNO As String = ""
    '    '為了避免身分證號又被修改而導致更新錯誤，所以重新取得一次推介單號
    '    tmpTicketNO = TIMS.Get_GOVTRNData(OCID, IDNO, oConn)
    '    If tmpTicketNO <> "" Then
    '        If SOCIDValue <> "" Then

    '            Dim dt As New DataTable
    '            sql = "SELECT * FROM Adp_GOVTRNData WHERE TICKET_NO='" & tmpTicketNO & "'"
    '            Dim sCmd As New SqlCommand(sql, oConn)
    '            With sCmd
    '                dt.Load(.ExecuteReader())
    '            End With
    '            If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
    '            '若有此單號，更新下列狀況

    '            Dim UsSql As String = ""
    '            UsSql = "" & vbCrLf
    '            UsSql &= " UPDATE ADP_GOVTRNDATA" & vbCrLf
    '            UsSql &= " SET SOCID=@SOCID" & vbCrLf
    '            UsSql &= " ,ARVL_STATE=@ARVL_STATE" & vbCrLf
    '            UsSql &= " ,ARVL_DATE=@ARVL_DATE" & vbCrLf
    '            UsSql &= " ,ARVL_SDATE=@ARVL_SDATE" & vbCrLf
    '            UsSql &= " ,ARVL_EDATE=@ARVL_EDATE" & vbCrLf

    '            UsSql &= " ,ARVL_HOURS=@ARVL_HOURS" & vbCrLf
    '            UsSql &= " ,ARVL_UNIT_ZIP=@ARVL_UNIT_ZIP" & vbCrLf
    '            UsSql &= " ,ARVL_UNIT_ADDR=@ARVL_UNIT_ADDR" & vbCrLf
    '            UsSql &= " ,ARVL_UNIT_NAME=@ARVL_UNIT_NAME" & vbCrLf
    '            UsSql &= " ,ARVL_CLASS_NAME=@ARVL_CLASS_NAME" & vbCrLf
    '            UsSql &= " ,ARVL_UNIT_USER=@ARVL_UNIT_USER" & vbCrLf
    '            UsSql &= " ,ARVL_UNIT_TEL=@ARVL_UNIT_TEL" & vbCrLf
    '            UsSql &= " ,ARVL_FSH_DATE=@ARVL_FSH_DATE" & vbCrLf
    '            UsSql &= " ,TIMSModifyDate=GETDATE()" & vbCrLf
    '            UsSql &= " ,TransToTIMS='Y'" & vbCrLf
    '            UsSql &= " WHERE TICKET_NO=@TICKET_NO" & vbCrLf
    '            '若有此單號，更新下列狀況
    '            Dim UsCmd As New SqlCommand(UsSql, oConn)
    '            With UsCmd
    '                .Parameters.Clear()
    '                .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SOCIDValue)
    '                .Parameters.Add("ARVL_STATE", SqlDbType.NVarChar).Value = "1"
    '                .Parameters.Add("ARVL_DATE", SqlDbType.DateTime).Value = TIMS.cdate2(EnterDate)
    '                .Parameters.Add("ARVL_SDATE", SqlDbType.DateTime).Value = TIMS.cdate2(OpenDate)
    '                .Parameters.Add("ARVL_EDATE", SqlDbType.DateTime).Value = TIMS.cdate2(CloseDate)

    '                .Parameters.Add("ARVL_HOURS", SqlDbType.BigInt).Value = Val(THours)
    '                .Parameters.Add("ARVL_UNIT_ZIP", SqlDbType.NVarChar).Value = TaddressZip
    '                .Parameters.Add("ARVL_UNIT_ADDR", SqlDbType.NVarChar).Value = TAddress
    '                .Parameters.Add("ARVL_UNIT_NAME", SqlDbType.NVarChar).Value = OrgName
    '                .Parameters.Add("ARVL_CLASS_NAME", SqlDbType.NVarChar).Value = ClassName
    '                .Parameters.Add("ARVL_UNIT_USER", SqlDbType.NVarChar).Value = ContactName
    '                .Parameters.Add("ARVL_UNIT_TEL", SqlDbType.NVarChar).Value = ContactPhone
    '                .Parameters.Add("ARVL_FSH_DATE", SqlDbType.NVarChar).Value = TIMS.cdate2(CloseDate)
    '                .Parameters.Add("TICKET_NO", SqlDbType.NVarChar).Value = tmpTicketNO
    '                .ExecuteNonQuery()
    '            End With
    '        Else
    '            'TIMS.Update_GOVTRNData(tmpTicketNO)
    '            Dim sqlStr As String = ""
    '            sqlStr = "" & vbCrLf
    '            sqlStr &= " UPDATE Adp_GOVTRNData" & vbCrLf
    '            sqlStr &= " set TransToTIMS='Y'" & vbCrLf
    '            sqlStr &= " ,TIMSModifyDate=GETDATE()" & vbCrLf
    '            sqlStr &= " where TICKET_NO='" & tmpTicketNO & "'" & vbCrLf
    '            'sqlStr += "  and (ARVL_STATE not in ('2','9') OR ARVL_STATE IS NULL) and TransToTIMS='N' " & vbCrLf
    '            DbAccess.ExecuteNonQuery(sqlStr, oConn)
    '        End If
    '    End If

    'End Sub

#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.TestDbConn(Me, objConn)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn)
        '檢查Session是否存在 End

        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
            str_superuser1 = CStr(sm.UserInfo.UserID)
        End If

        If Not flgROLEIDx0xLIDx0 Then
            Common.MessageBox(Me, "權限錯誤，無法使用此功能。")
            tr_Info.Visible = True
            tr_Edit1.Visible = False
            tr_Edit2.Visible = False
            tr_Edit3.Visible = False
            tr_Edit10.Visible = False
            Exit Sub
        End If

        ViewState(Cst_DelVal1) = Get_SYSTEM_VS_DELVAL1()

        If Not Me.IsPostBack Then
            txt_IDNO.Attributes.Add("onBlur", "this.value=this.value.toUpperCase();")
            txt_EditIDNO1.Attributes.Add("onBlur", "this.value=this.value.toUpperCase();")
            txt_EditIDNO2.Attributes.Add("onBlur", "this.value=this.value.toUpperCase();")
            txt_EditIDNO3.Attributes.Add("onBlur", "this.value=this.value.toUpperCase();")
            txt_EditIdno10.Attributes.Add("onBlur", "this.value=this.value.toUpperCase();")

            btn_EditdelClass1.Attributes.Add("onClick", "return confirm('確認要刪除該筆資料??');")
            btn_EditdelClass2.Attributes.Add("onClick", "return confirm('確認要刪除該筆資料??');")
            '取消課程報到資料
            btn_EditUpdateCls2.Attributes.Add("onClick", "return confirm('確認要刪除學員資料，取消課程報到??');")
            btn_EditdelClass3.Attributes.Add("onClick", "return confirm('確認要刪除該筆資料??');")

            btn_EditSave1.Attributes.Add("onClick", "return confirm('請先確認修改資料正確無誤，確定要修改?');")
            btn_EditSave2.Attributes.Add("onClick", "return confirm('請先確認修改資料正確無誤，確定要修改?');")
            btn_EditSave3.Attributes.Add("onClick", "return confirm('請先確認修改資料正確無誤，確定要修改?');")
            btn_EditSave10.Attributes.Add("onClick", "return confirm('請先確認修改資料正確無誤，確定要修改?');")

            tr_Info.Visible = True
            tr_Edit1.Visible = False
            tr_Edit2.Visible = False
            tr_Edit3.Visible = False
            tr_Edit10.Visible = False
        End If

        'btn_EditSave1.Attributes("onclick") = "return confirm('請先確認修改資料正確無誤，確定要修改?');"
        'btn_EditSave2.Attributes("onclick") = "return confirm('請先確認修改資料正確無誤，確定要修改?');"
        'btn_EditSave3.Attributes("onclick") = "return confirm('請先確認修改資料正確無誤，確定要修改?');"
        'btn_EditSave10.Attributes("onclick") = "return confirm('請先確認修改資料正確無誤，確定要修改?');"
    End Sub

    Function Get_SYSTEM_VS_DELVAL1() As String
        If ViewState(Cst_DelVal1) IsNot Nothing Then
            If Convert.ToString(ViewState(Cst_DelVal1)).Length > 0 Then
                Return ViewState(Cst_DelVal1)
            End If
        End If
        ViewState(Cst_DelVal1) = TIMS.GetSystemValue(Cst_SYS_03_009, "", Cst_DEL1, objConn)
        Return ViewState(Cst_DelVal1)
    End Function

    ''' <summary>'查詢</summary>
    Sub QuerySearch1()
        If str_superuser1 <> CStr(sm.UserInfo.UserID) Then
            Common.MessageBox(Me, "權限錯誤，無法使用此功能。")
            Exit Sub
        End If
        'txt_IDNO.Text = UCase(txt_IDNO.Text)
        txt_IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txt_IDNO.Text))
        If txt_IDNO.Text = "" Then
            Common.MessageBox(Me, "請輸入身分證號")
            Exit Sub
        End If

        vs_IDNO = txt_IDNO.Text
        If vs_IDNO <> "" Then
            aNow1 = Now
            Show_DataGrid1(vs_IDNO, "", 0)
            aNow1 = Now
            Show_DataGrid2(vs_IDNO, "", 0)
            aNow1 = Now
            Show_DataGrid3(vs_IDNO, "", 0)
            aNow1 = Now
            Show_DataGrid4(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid5(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid6(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid7(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid8(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid9(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid10(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid11(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid12(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid13(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid14(vs_IDNO, 0)
            aNow1 = Now
            Show_DataGrid15(vs_IDNO, 0)
        End If
    End Sub

    '查詢鈕
    Private Sub btn_Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Query.Click
        Call QuerySearch1()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO1")
        Dim labSID As Label = e.Item.FindControl("lab_SID1")
        Dim labName As Label = e.Item.FindControl("lab_Name1")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday1")
        Dim labClass As Label = e.Item.FindControl("lab_Class1")
        Dim labSUBID As Label = e.Item.FindControl("lab_SUBID1")

        Select Case e.CommandName
            Case "EDIT1"
                Dim dt_Class As DataTable = Nothing

                txt_EditIDNO1.Text = labIDNO.Text
                txt_EditName1.Text = labName.Text
                txt_EditBirthday1.Text = labBirthday.Text
                lab_EditSID1.Text = labSID.Text
                lab_EditSUBID1.Text = labSUBID.Text.Replace("<br />", ";")
                'lab_EditSUBID1.Text = lab_EditSUBID1.Text
                lab_msg_stud.Text = ""
                Show_RdoEditClass1(Get_ClassStudentsOfClass("SID", labSID.Text))
                tr_Info.Visible = False
                tr_Edit1.Visible = True
                tr_Edit2.Visible = False
                tr_Edit3.Visible = False
                tb_EditClass1.Visible = False

                vs_IDNO = TIMS.ClearSQM(labIDNO.Text)
                If vs_IDNO = "" Then Exit Sub

                'SID下拉選單
                list_EditSID1.DataSource = Get_SID(vs_IDNO)
                list_EditSID1.DataTextField = "SID"
                list_EditSID1.DataValueField = "SID"
                list_EditSID1.DataBind()
                list_EditSID1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                list_EditSID1.SelectedIndex = -1
                list_EditSID1.Attributes.Add("onChange", "document.getElementById('" & txt_EditClassSID1.ClientID & "').value=this.value;")
                'SETID下拉選單
                list_EditSETID1.DataSource = Get_SETID(vs_IDNO, "Stud_StudentInfo")
                list_EditSETID1.DataTextField = "SETID"
                list_EditSETID1.DataValueField = "SETID"
                list_EditSETID1.DataBind()
                list_EditSETID1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                list_EditSETID1.SelectedIndex = -1
                list_EditSETID1.Attributes.Add("onChange", "document.getElementById('" & txt_EditClassSETID1.ClientID & "').value=this.value;")

            Case "DELE1"
                Dim sCmdArg As String = e.CommandArgument
                Dim vs_IDNO As String = TIMS.GetMyValue(sCmdArg, "IDNO")
                Dim vs_SID As String = TIMS.GetMyValue(sCmdArg, "SID")
                If sCmdArg = "" OrElse vs_IDNO = "" OrElse vs_SID = "" Then Exit Sub

                Dim sqlTrans As SqlTransaction = objConn.BeginTransaction()
                Try
                    TIMS.DEL_STUDSTUDENTINFO(sm, sqlTrans, objConn, vs_SID)
                    TIMS.DEL_STUDSUBDATA(sm, sqlTrans, objConn, vs_SID)
                    sqlTrans.Commit()
                Catch ex As Exception
                    sqlTrans.Rollback()
                    'sqlTrans.Dispose()
                    Common.MessageBox(Me, ex.ToString)
                    Return
                End Try

                Call QuerySearch1()
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim labSNo As Label = e.Item.FindControl("lab_SNo1")
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO1")
        Dim labSID As Label = e.Item.FindControl("lab_SID1")
        Dim labName As Label = e.Item.FindControl("lab_Name1")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday1")
        Dim labClass As Label = e.Item.FindControl("lab_Class1")
        Dim labSUBID As Label = e.Item.FindControl("lab_SUBID1")

        'Case ListItemType.Header
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim btnEdit As LinkButton = e.Item.FindControl("btn_Edit1")
                btnEdit.CommandName = "EDIT1"

                Dim btnDele1 As LinkButton = e.Item.FindControl("btn_Dele1")
                btnDele1.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDele1.CommandName = "DELE1"

                Dim dr_Data As DataRowView = e.Item.DataItem

                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
                labIDNO.Text = Convert.ToString(dr_Data("IDNO"))
                labName.Text = Convert.ToString(dr_Data("Name"))
                labBirthday.Text = Convert.ToString(dr_Data("Birthday"))
                labSID.Text = Convert.ToString(dr_Data("SID")) 'SID
                If Not ChkSubDATA(Convert.ToString(dr_Data("SID"))) Then
                    labSID.Text += "<br><font color='red'>(副檔資料為空)</font> "
                End If

                '取得班及資料
                Dim dt_Class As DataTable = Get_ClassStudentsOfClass("SID", Convert.ToString(dr_Data("SID")))

                'btnEdit.Visible = (dt_Class IsNot Nothing AndAlso dt_Class.Rows.Count > 0)
                btnDele1.Visible = (dt_Class Is Nothing OrElse dt_Class.Rows.Count = 0)

                Call SHOW_STUDENTINFO(dt_Class, labClass)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "IDNO", Convert.ToString(dr_Data("IDNO")))
                TIMS.SetMyValue(sCmdArg, "SID", Convert.ToString(dr_Data("SID")))
                btnEdit.CommandArgument = sCmdArg
                btnDele1.CommandArgument = sCmdArg
        End Select
    End Sub

    Sub SHOW_STUDENTINFO(ByRef dt_Class As DataTable, ByRef labClass As Label)
        If dt_Class Is Nothing OrElse dt_Class.Rows.Count = 0 Then Return
        Dim TotalSumOfMoney As Integer = 0 '補助金額總合
        'TotalSumOfMoney = 0
        For Each dr As DataRow In dt_Class.Rows
            Dim dt_SOCID As DataTable = Nothing
            Dim dt_SUBID As DataTable = Nothing
            'document.getElementById('plan" & Convert.ToString(dr("OCID")) & "').style.display

            'labClass.Text += "<div onmouseover=""javascript:if (document.getElementById('plan" & Convert.ToString(dr("OCID")) & "')){"
            'labClass.Text += "document.getElementById('plan" & Convert.ToString(dr("OCID")) & "').style.display='inline';"
            'labClass.Text += "}"" onmouseout=""javascript:if (document.getElementById('plan" & Convert.ToString(dr("OCID")) & "')){"
            'labClass.Text += "document.getElementById('plan" & Convert.ToString(dr("OCID")) & "').style.display='none';"
            'labClass.Text += "}"""
            'labClass.Text += ">"

            labClass.Text += "<span onclick=""javascript:if (document.getElementById('plan" & Convert.ToString(dr("OCID")) & "')){"
            labClass.Text += "document.getElementById('plan" & Convert.ToString(dr("OCID")) & "').style.display='inline';"
            labClass.Text += "}"" ondblclick=""javascript:if (document.getElementById('plan" & Convert.ToString(dr("OCID")) & "')){"
            labClass.Text += "document.getElementById('plan" & Convert.ToString(dr("OCID")) & "').style.display='none';"
            labClass.Text += "}"""
            labClass.Text += "><b>"
            labClass.Text += "OCID：" & Convert.ToString(dr("OCID"))
            labClass.Text += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Convert.ToString(dr("ClassName")) & "<br />"
            labClass.Text += "</b></span>"

            labClass.Text += "<span id=""plan" & Convert.ToString(dr("OCID")) & """ style=""display:none"" "
            labClass.Text += "onclick=""javascript:if (document.getElementById('plan" & Convert.ToString(dr("OCID")) & "')){"
            labClass.Text += "document.getElementById('plan" & Convert.ToString(dr("OCID")) & "').style.display='none';"
            labClass.Text += "}"""
            labClass.Text += ">"
            labClass.Text += "PlanID：" & Convert.ToString(dr("PlanID"))
            labClass.Text += "&nbsp;&nbsp;TPlanID：" & Convert.ToString(dr("TPlanID"))
            labClass.Text += "&nbsp;&nbsp;" & Convert.ToString(dr("PlanName"))
            labClass.Text += "&nbsp;&nbsp;" & Convert.ToString(dr("OrgName")) & "<br />"
            labClass.Text += "</span>"

            labClass.Text += "SOCID：" & Convert.ToString(dr("SOCID"))
            labClass.Text += "&nbsp;&nbsp;SETID：" & Convert.ToString(dr("SETID"))
            labClass.Text += "&nbsp;&nbsp;開訓日期：" & Convert.ToString(dr("STDate"))
            labClass.Text += "&nbsp;&nbsp;學員狀態：" & Convert.ToString(dr("StudStatus"))
            If Convert.ToString(dr("CreditPoints")) <> "" Then
                Select Case Convert.ToString(dr("CreditPoints"))
                    Case "0"
                        labClass.Text += "&nbsp;&nbsp;結訓資格：否"
                    Case "1"
                        labClass.Text += "&nbsp;&nbsp;結訓資格：是"
                End Select
            Else
                labClass.Text += "&nbsp;&nbsp;結訓資格：空"
            End If
            Select Case Convert.ToString(dr("StudStatus"))
                Case "1"
                    labClass.Text += "&nbsp;StudStatus:在訓"
                Case "2"
                    labClass.Text += "&nbsp;StudStatus:離訓" & Convert.ToString(dr("RejectTDate1"))
                Case "3"
                    labClass.Text += "&nbsp;StudStatus:退訓" & Convert.ToString(dr("RejectTDate2"))
                Case "4"
                    labClass.Text += "&nbsp;StudStatus:續訓"
                Case "5"
                    labClass.Text += "&nbsp;StudStatus:結訓" & Convert.ToString(dr("CloseDate"))
                Case Else
                    labClass.Text += "&nbsp;StudStatus:(異常)" & Convert.ToString(dr("StudStatus"))
            End Select

            labClass.Text += "" & "<br />"
            '取得補助金資料
            dt_SOCID = Get_StudSubSidyCost(dr("SOCID"))

            If dt_SOCID IsNot Nothing Then
                For Each dr_SOCID As DataRow In dt_SOCID.Rows
                    Dim SumOfMoney As String = "" '補助金額
                    Dim AllotDate As String = "" '撥款日期
                    SumOfMoney = Convert.ToString(dr_SOCID("SumOfMoney"))
                    AllotDate = Convert.ToString(dr_SOCID("AllotDate"))
                    Select Case Convert.ToString(dr_SOCID("AppliedStatusM"))
                        Case "Y"
                            TotalSumOfMoney += SumOfMoney
                            labClass.Text += "<font color='red'>(有一筆審核成功的產投補助金。$:" & SumOfMoney & ")</font><br />"
                        Case "N"
                            labClass.Text += "<font color='blue'>(有一筆審核失敗的產投補助金。$:" & SumOfMoney & ")</font><br />"
                        Case "R"
                            labClass.Text += "<font color='green'>(有一筆退件修正的產投補助金。$:" & SumOfMoney & ")</font><br />"
                        Case Else
                            labClass.Text += "<font color='orange'>(有一筆未審核的產投補助金。$:" & SumOfMoney & ")</font><br />"
                    End Select
                Next

            End If

            '取得生活津貼
            dt_SUBID = Get_SubSubSidyApply("SOCID", Convert.ToString(dr("SOCID")))

            If dt_SUBID IsNot Nothing Then
                For Each dr_SUBID As DataRow In dt_SUBID.Rows
                    labClass.Text += "<font color='blue'>(有一筆初審"
                    labClass.Text += If(Convert.ToString(dr_SUBID("AppliedStatusF")) = "Y", "通過", "未審或不通過")
                    labClass.Text += "與勾稽" & If(Convert.ToString(dr_SUBID("AppliedStatusFin")) = "Y", "通過", "未審或不通過")
                    labClass.Text += "之生活津貼 FROM Sub_SubSidyApply WHERE SUBID=" & dr_SUBID("SUBID") & ")</font><br />"
                Next
            End If

        Next
        If TotalSumOfMoney > 0 Then
            labClass.Text += String.Concat("<font color='red'>合計產投補助金:", TotalSumOfMoney, "</font><br />")
        End If

    End Sub

    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO2")
        Dim labSETID As Label = e.Item.FindControl("lab_SETID2")
        Dim labeSETID As Label = e.Item.FindControl("lab_eSETID2")
        Dim labName As Label = e.Item.FindControl("lab_Name2")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday2")
        Dim labClass As Label = e.Item.FindControl("lab_Class2")
        Dim dt_Class As DataTable = Nothing

        txt_EditIDNO2.Text = labIDNO.Text
        txt_EditName2.Text = labName.Text
        txt_EditBirthday2.Text = labBirthday.Text
        lab_EditSETID2.Text = labSETID.Text
        txt_EditeSETID2.Text = labeSETID.Text
        lab_SelResult_msg.Text = ""
        Call Show_RdoEditClass2(Get_StudEnterType("SETID", labSETID.Text))
        tr_Info.Visible = False
        tr_Edit1.Visible = False
        tr_Edit2.Visible = True
        tr_Edit3.Visible = False
        tb_EditClass2.Visible = False

        vs_IDNO = labIDNO.Text
        vs_IDNO = TIMS.ClearSQM(vs_IDNO)
        If vs_IDNO = "" Then Exit Sub

        'SETID下拉選單
        list_EditSETID2.DataSource = Get_SETID(vs_IDNO, "Stud_EnterTemp")
        list_EditSETID2.DataTextField = "SETID"
        list_EditSETID2.DataValueField = "SETID"
        list_EditSETID2.DataBind()

        list_EditSETID2.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        list_EditSETID2.SelectedIndex = -1
        list_EditSETID2.Attributes.Add("onChange", "document.getElementById('" & txt_EditClassSETID2.ClientID & "').value=this.value;")
        'eSETID下拉選單
        list_EditeSETID2.DataSource = Get_eSETID(vs_IDNO, "Stud_EnterTemp")
        list_EditeSETID2.DataTextField = "eSETID"
        list_EditeSETID2.DataValueField = "eSETID"
        list_EditeSETID2.DataBind()

        list_EditeSETID2.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        list_EditeSETID2.SelectedIndex = -1
        list_EditeSETID2.Attributes.Add("onChange", "document.getElementById('" & txt_EditClasseSETID2.ClientID & "').value=this.value;")
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Dim labSNo As Label = e.Item.FindControl("lab_SNo2")
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO2")
        Dim labSETID As Label = e.Item.FindControl("lab_SETID2")
        Dim labeSETID As Label = e.Item.FindControl("lab_eSETID2")
        Dim labName As Label = e.Item.FindControl("lab_Name2")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday2")
        Dim labClass As Label = e.Item.FindControl("lab_Class2")
        Dim btnEdit As LinkButton = e.Item.FindControl("btn_Edit2")
        Dim dr_Data As DataRowView = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dt_Class As DataTable = Nothing

                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
                labIDNO.Text = Convert.ToString(dr_Data("IDNO"))
                labName.Text = Convert.ToString(dr_Data("Name"))
                labBirthday.Text = Convert.ToString(dr_Data("Birthday"))
                labSETID.Text = Convert.ToString(dr_Data("SETID"))
                labeSETID.Text = Convert.ToString(dr_Data("eSETID"))
                dt_Class = Get_StudEnterType("SETID", Convert.ToString(dr_Data("SETID")))
                If dt_Class IsNot Nothing Then
                    For Each dr As DataRow In dt_Class.Rows
                        labClass.Text += "<b>"
                        labClass.Text += "OCID：" & Convert.ToString(dr("OCID1"))
                        labClass.Text += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Convert.ToString(dr("ClassName"))
                        labClass.Text += "</b>"

                        labClass.Text += "&nbsp;&nbsp;是否錄取："
                        Select Case Convert.ToString(dr("Admission"))
                            Case "Y"
                                labClass.Text += "Y:錄取"
                            Case "N"
                                labClass.Text += "N:未錄取"
                            Case Else
                                labClass.Text += Convert.ToString(dr("Admission")) & ":未審核"
                        End Select
                        labClass.Text += "&nbsp;&nbsp;甄試結果代碼："
                        Select Case Convert.ToString(dr("SelResultID"))
                            Case "01"
                                labClass.Text += "01:正取"
                            Case "02"
                                labClass.Text += "02:備取"
                            Case "03"
                                labClass.Text += "03:不錄取(未錄取)"
                            Case Else
                                labClass.Text += Convert.ToString(dr("SelResultID")) & ":未審核"
                        End Select
                        labClass.Text += "&nbsp;&nbsp;甄試時段：" & Convert.ToString(dr("EPName"))
                        labClass.Text += "<br />"

                        labClass.Text += "SETID：" & Convert.ToString(dr("SETID"))
                        labClass.Text += "&nbsp;&nbsp;EnterDate：" & Convert.ToString(dr("EnterDate"))
                        labClass.Text += "&nbsp;&nbsp;SerNum：" & Convert.ToString(dr("SerNum"))
                        labClass.Text += "&nbsp;&nbsp;EXAMNO：" & Convert.ToString(dr("EXAMNO"))
                        labClass.Text += "&nbsp;&nbsp;eSETID：" & Convert.ToString(dr("eSETID"))
                        labClass.Text += "&nbsp;&nbsp;eSerNum：" & Convert.ToString(dr("eSerNum")) & "<br />"
                    Next
                End If
        End Select
    End Sub

    Private Sub Datagrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid3.ItemCommand
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO3")
        Dim labSETID As Label = e.Item.FindControl("lab_SETID3")
        Dim labeSETID As Label = e.Item.FindControl("lab_eSETID3")
        Dim labName As Label = e.Item.FindControl("lab_Name3")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday3")
        Dim labClass As Label = e.Item.FindControl("lab_Class3")

        txt_EditIDNO3.Text = labIDNO.Text
        txt_EditName3.Text = labName.Text
        txt_EditBirthday3.Text = labBirthday.Text
        txt_EditSETID3.Text = labSETID.Text
        lab_EditeSETID3.Text = labeSETID.Text
        Lab_ENTERTYPE2_mag.Text = ""
        Show_RdoEditClass3(Get_StudEnterType2("eSETID", labeSETID.Text))
        tr_Info.Visible = False
        tr_Edit1.Visible = False
        tr_Edit2.Visible = False
        tr_Edit3.Visible = True
        tb_EditClass3.Visible = False

        vs_IDNO = labIDNO.Text
        vs_IDNO = TIMS.ClearSQM(vs_IDNO)
        If vs_IDNO = "" Then Exit Sub
        'eSETID下拉選單
        list_EditeSETID3.DataSource = Get_eSETID(vs_IDNO, "Stud_EnterTemp2")
        list_EditeSETID3.DataTextField = "eSETID"
        list_EditeSETID3.DataValueField = "eSETID"
        list_EditeSETID3.DataBind()

        list_EditeSETID3.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        list_EditeSETID3.SelectedIndex = -1
        list_EditeSETID3.Attributes.Add("onChange", "document.getElementById('" & txt_EditClasseSETID3.ClientID & "').value=this.value;")
        'SETID下拉選單
        list_EditSETID3.DataSource = Get_SETID(vs_IDNO, "Stud_EnterTemp2")
        list_EditSETID3.DataTextField = "SETID"
        list_EditSETID3.DataValueField = "SETID"
        list_EditSETID3.DataBind()

        list_EditSETID3.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        list_EditSETID3.SelectedIndex = -1
        list_EditSETID3.Attributes.Add("onChange", "document.getElementById('" & txt_EditClassSETID3.ClientID & "').value=this.value;")
    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Dim labSNo As Label = e.Item.FindControl("lab_SNo3")
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO3")
        Dim labSETID As Label = e.Item.FindControl("lab_SETID3")
        Dim labeSETID As Label = e.Item.FindControl("lab_eSETID3")
        Dim labName As Label = e.Item.FindControl("lab_Name3")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday3")
        Dim labClass As Label = e.Item.FindControl("lab_Class3")
        Dim btnEdit As LinkButton = e.Item.FindControl("btn_Edit3")
        Dim dr_Data As DataRowView = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dt_Class As DataTable = Nothing

                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
                labIDNO.Text = Convert.ToString(dr_Data("IDNO"))
                labName.Text = Convert.ToString(dr_Data("Name"))
                labBirthday.Text = Convert.ToString(dr_Data("Birthday"))
                labSETID.Text = Convert.ToString(dr_Data("SETID"))
                labeSETID.Text = Convert.ToString(dr_Data("eSETID"))
                dt_Class = Get_StudEnterType2("eSETID", Convert.ToString(dr_Data("eSETID")))
                If dt_Class IsNot Nothing Then
                    For Each dr As DataRow In dt_Class.Rows
                        labClass.Text += "<b>"
                        labClass.Text += "OCID：" & Convert.ToString(dr("OCID1"))
                        labClass.Text += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Convert.ToString(dr("ClassName"))
                        labClass.Text += "</b>"

                        labClass.Text += "&nbsp;&nbsp;報名狀態："
                        Select Case Convert.ToString(dr("signUpStatus"))
                            Case "0"
                                labClass.Text += "0:收件完成"
                            Case "1"
                                labClass.Text += "1:報名成功"
                            Case "2"
                                labClass.Text += "2:報名失敗"
                            Case "3"
                                labClass.Text += "3:正取"
                            Case "4"
                                labClass.Text += "4:備取"
                            Case "5"
                                labClass.Text += "5:未錄取"
                            Case Else
                                labClass.Text += Convert.ToString(dr("signUpStatus")) & ":異常資料"
                        End Select
                        labClass.Text += "&nbsp;&nbsp;甄試時段：" & Convert.ToString(dr("EPName"))

                        labClass.Text += "<br />"
                        labClass.Text += "eSerNum：" & Convert.ToString(dr("eSerNum"))
                        labClass.Text += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;eSETID：" & Convert.ToString(dr("eSETID"))
                        labClass.Text += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;SETID：" & Convert.ToString(dr("SETID")) & "<br />"
                    Next
                End If
        End Select
    End Sub

    Private Sub rdo_EditClass1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo_EditClass1.SelectedIndexChanged
        Dim strInfo() As String = Split(rdo_EditClass1.SelectedItem.Text, "；")
        If strInfo(cst_REC1_Class1) IsNot Nothing Then lab_EditClass1.Text = strInfo(cst_REC1_Class1) 'const cst_rec1_OCID as integer =0
        If strInfo(cst_REC1_SID1) IsNot Nothing Then txt_EditClassSID1.Text = strInfo(cst_REC1_SID1).Split("：")(1)
        If strInfo(cst_REC1_SETID1) IsNot Nothing Then txt_EditClassSETID1.Text = strInfo(cst_REC1_SETID1).Split("：")(1)
        If txt_EditClassSID1.Text <> list_EditSID1.SelectedValue Then
            list_EditSID1.SelectedIndex = -1
        End If
        If txt_EditClassSETID1.Text <> list_EditSETID1.SelectedValue Then
            list_EditSETID1.SelectedIndex = -1
        End If
        tb_EditClass1.Visible = True
    End Sub

    Private Sub rdo_EditClass2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo_EditClass2.SelectedIndexChanged
        tb_EditClass2.Visible = False
        Dim v_rdo_EditClass2 As String = TIMS.GetListValue(rdo_EditClass2)
        Dim txt_rdo_EditClass2 As String = TIMS.GetListText(rdo_EditClass2)
        If txt_rdo_EditClass2 <> "" Then
            Dim strInfo() As String = Split(txt_rdo_EditClass2, "；")
            If strInfo(cst_REC2_Class2) IsNot Nothing Then lab_EditClass2.Text = strInfo(cst_REC2_Class2)
            If strInfo(cst_REC2_SETID2) IsNot Nothing Then txt_EditClassSETID2.Text = strInfo(cst_REC2_SETID2).Split("：")(1)
            If strInfo(cst_REC2_eSETID2) IsNot Nothing Then txt_EditClasseSETID2.Text = strInfo(cst_REC2_eSETID2).Split("：")(1)
        End If
        If txt_EditClassSETID2.Text <> list_EditSETID2.SelectedValue Then list_EditSETID2.SelectedIndex = -1
        If txt_EditClasseSETID2.Text <> list_EditeSETID2.SelectedValue Then list_EditeSETID2.SelectedIndex = -1
        tb_EditClass2.Visible = True

        BTNSCH_T1_A.Visible = False
        'Common.MessageBox(Me, "請選取要處理的資料。")
        If v_rdo_EditClass2 = "" Then Exit Sub
        Dim strPK() As String = Split(v_rdo_EditClass2, ";")
        Dim dr1 As DataRow = Nothing 'STUD_ENTERTYPE
        Dim dr2 As DataRow = Nothing 'STUD_ENTERTEMP
        Dim dr3 As DataRow = Nothing 'CLASS_STUDENTSOFCLASS
        Call GET_STUDDATA3(strPK, dr1, dr2, dr3)
        'If dr1 Is Nothing Then Exit Sub
        'If dr2 Is Nothing Then Exit Sub
        Dim s_CanUseYears28 As String = ""
        If v_rdo_EditClass2 <> "" Then s_CanUseYears28 = If(v_rdo_EditClass2.Split(";").Length > 3, v_rdo_EditClass2.Split(";")(3), "")
        If dr3 Is Nothing AndAlso s_CanUseYears28 = "Y" Then BTNSCH_T1_A.Visible = True '可執行 (產投)查詢重複  

    End Sub

    Private Sub rdo_EditClass3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo_EditClass3.SelectedIndexChanged
        Dim strInfo() As String = Split(rdo_EditClass3.SelectedItem.Text, "；")
        If strInfo(cst_REC3_Class3) IsNot Nothing Then lab_EditClass3.Text = strInfo(cst_REC3_Class3)
        If strInfo(cst_REC3_SETID3) IsNot Nothing Then txt_EditClassSETID3.Text = strInfo(cst_REC3_SETID3).Split("：")(1)
        If strInfo(cst_REC3_eSETID3) IsNot Nothing Then txt_EditClasseSETID3.Text = strInfo(cst_REC3_eSETID3).Split("：")(1)

        If txt_EditClasseSETID3.Text <> list_EditeSETID3.SelectedValue Then list_EditeSETID3.SelectedIndex = -1

        If txt_EditClassSETID3.Text <> list_EditSETID3.SelectedValue Then list_EditSETID3.SelectedIndex = -1

        tb_EditClass3.Visible = True
    End Sub

    Private Sub btn_EditCancel1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditCancel1.Click
        Clear_Edit()
    End Sub

    Private Sub btn_EditCancel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditCancel2.Click
        Clear_Edit()
    End Sub

    Private Sub btn_EditCancel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditCancel3.Click
        Clear_Edit()
    End Sub

    '儲存課程變更
    Private Sub btn_EditSaveClass1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSaveClass1.Click
        'Dim sqlAdp As New SqlDataAdapter
        'Dim sqlStr As String = ""
        'Dim objTrans As SqlTransaction = Nothing

        If rdo_EditClass1.SelectedIndex = -1 AndAlso rdo_EditClass1.SelectedValue = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        If txt_EditClassSID1.Text = "" Then
            Common.MessageBox(Me, "SID不能沒填。")
            Exit Sub
        End If

        iRst = 0
        Dim hParms As New Hashtable
        hParms.Add("SID", txt_EditClassSID1.Text)
        hParms.Add("SETID", txt_EditClassSETID1.Text)
        hParms.Add("SOCID", rdo_EditClass1.SelectedValue)
        Call UPDATE_CLASS_STUDENTSOFCLASS(objConn, hParms)
        iRst = 1

        'Common.MessageBox(Me, "修改完成。")
        lab_msg_stud.Text = String.Concat(Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), "，更新筆數：", iRst)
        Show_RdoEditClass1(Get_ClassStudentsOfClass("SID", lab_EditSID1.Text))
        tb_EditClass1.Visible = False
    End Sub

    Public Shared Sub UPDATE_CLASS_STUDENTSOFCLASS(ByRef oConn As SqlConnection, ByRef hParms As Hashtable)
        'DbAccess.Open(objConn)
        Dim v_SID As String = TIMS.GetMyValue2(hParms, "SID")
        Dim v_SETID As String = TIMS.GetMyValue2(hParms, "SETID")
        Dim v_SOCID As String = TIMS.GetMyValue2(hParms, "SOCID")
        Dim sqlStr As String = ""
        '更新學員課程資料Class_StudentsOfClass的SID、SETID
        sqlStr = "UPDATE CLASS_STUDENTSOFCLASS SET SID=@SID,SETID=@SETID WHERE SOCID=@SOCID"
        Dim u_cmd As New SqlCommand(sqlStr, oConn)
        With u_cmd
            .Parameters.Clear()
            .Parameters.Add("SID", SqlDbType.VarChar).Value = v_SID 'txt_EditClassSID1.Text
            .Parameters.Add("SETID", SqlDbType.Int).Value = If(v_SETID = "", Convert.DBNull, Val(v_SETID))
            .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(v_SOCID)
            .ExecuteNonQuery()
        End With

        '更新津貼資料Sub_SubSidyApply的SID
        sqlStr = "UPDATE SUB_SUBSIDYAPPLY SET SID=@SID WHERE SOCID=@SOCID"
        Dim u_cmd2 As New SqlCommand(sqlStr, oConn)
        With u_cmd2
            .Parameters.Clear()
            .Parameters.Add("SID", SqlDbType.VarChar).Value = v_SID 'txt_EditClassSID1.Text
            .Parameters.Add("SOCID", SqlDbType.Int).Value = Val(v_SOCID) 'rdo_EditClass1.SelectedValue
            .ExecuteNonQuery()
        End With
    End Sub

    '儲存課程變更 '含有修正SQL 語法。
    Private Sub btn_EditSaveClass2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSaveClass2.Click
        Dim v_rdo_EditClass2 As String = TIMS.GetListValue(rdo_EditClass2)
        If rdo_EditClass2.SelectedIndex = -1 OrElse v_rdo_EditClass2 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        Dim strPK() As String = Split(v_rdo_EditClass2, ";")
        If txt_EditClassSETID2.Text = "" Then
            Common.MessageBox(Me, "SETID不能沒填。")
            Exit Sub
        End If

        'Dim flagDblRows As Boolean = False '更新資料有重複 False:沒有
        '新的與舊的資料不同，做判斷
        If txt_EditClassSETID2.Text <> Convert.ToString(strPK(0)) Then
            Dim dt As New DataTable
            Dim sqlStr As String = "SELECT * FROM Stud_SelResult WHERE SETID= @newSETID and EnterDate= @EnterDate and SerNum= @SerNum"
            Dim sCmd As New SqlCommand(sqlStr, objConn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = txt_EditClassSETID2.Text
                '.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                dt.Load(.ExecuteReader())
            End With
            '更新資料有重複 False:沒有
            Dim flagDblRows As Boolean = (dt.Rows.Count > 0) '更新資料有重複 False:沒有 True:有
            If flagDblRows Then
                Common.MessageBox(Me, "修改資料異常，請先確認資料正確性!!")

                Dim pParms As New Hashtable
                pParms.Add("SETID_1", txt_EditClassSETID2.Text)
                pParms.Add("SETID_2", strPK(0))
                pParms.Add("ENTERDATE", strPK(1))
                pParms.Add("SERNUM", strPK(2))
                Call SHOW_ERROR_HANDLING_TIPS(pParms)
                'Show_RdoEditClass2(Get_StudEnterType("SETID", lab_EditSETID2.Text))
                'tb_EditClass2.Visible = False
                Return
            End If
        End If

        Dim UParms As New Hashtable
        UParms.Add("SETID_1", txt_EditClassSETID2.Text)
        UParms.Add("SETID_2", strPK(0))
        UParms.Add("ENTERDATE", strPK(1))
        UParms.Add("SERNUM", strPK(2))
        UParms.Add("eSETID", txt_EditClasseSETID2.Text)
        PGErrMsg1 = UPDATE_STUD_ENTERTYPE(objConn, UParms)
        If PGErrMsg1 <> "" Then
            Common.MessageBox(Me, PGErrMsg1)
            Return
        End If
        iRst = 1

        'Common.MessageBox(Me, "修改完成。")
        'lab_SelResult_msg.Text = Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
        lab_SelResult_msg.Text = String.Concat(Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), "，更新筆數：", iRst)
        Call Show_RdoEditClass2(Get_StudEnterType("SETID", lab_EditSETID2.Text))
        tb_EditClass2.Visible = False
    End Sub

    Public Shared Function UPDATE_STUD_ENTERTYPE(ByRef oConn As SqlConnection, ByRef pParms As Hashtable) As String
        Dim rst As String = ""
        Dim v_SETID_1 As String = TIMS.GetMyValue2(pParms, "SETID_1") 'txt_EditClassSETID2.Text

        Dim v_SETID_2 As String = TIMS.GetMyValue2(pParms, "SETID_2") 'strPK(0)
        Dim v_ENTERDATE As String = TIMS.GetMyValue2(pParms, "ENTERDATE") 'strPK(1)
        Dim v_SERNUM As String = TIMS.GetMyValue2(pParms, "SERNUM") 'strPK(2)

        Dim v_eSETID As String = TIMS.GetMyValue2(pParms, "eSETID") 'txt_EditClasseSETID2.Text
        If v_SETID_1 = "" OrElse v_SETID_2 = "" OrElse v_ENTERDATE = "" OrElse v_SERNUM = "" Then Return rst

        'Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn) 'objTrans = oConn.BeginTransaction()
        Try
            '更新報名資料Stud_EnterType的SETID、eSETID
            sqlStr = "UPDATE STUD_ENTERTYPE SET SETID=@newSETID,eSETID=@eSETID WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum"
            Dim uCmd As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
                .Parameters.Add("eSETID", SqlDbType.Int).Value = If(v_eSETID = "", Convert.DBNull, v_eSETID)

                .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
                .ExecuteNonQuery()
            End With

            '更新 STUD_ENTERTRAIN
            sqlStr = "UPDATE STUD_ENTERTRAIN SET SETID=@newSETID WHERE SETID=@SETID AND EnterDate=@EnterDate AND SERNUM=@SerNum"
            Dim uCmd7 As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd7
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
                .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
                .ExecuteNonQuery()
            End With

            '更新錄取結果Stud_SelResult的SETID
            sqlStr = "UPDATE STUD_SELRESULT SET SETID=@newSETID WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum"
            Dim uCmd2 As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd2
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
                .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
                .ExecuteNonQuery()
            End With

            '更新 STUD_SELRESULTBLI
            sqlStr = "UPDATE STUD_SELRESULTBLI SET SETID=@newSETID WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum"
            Dim uCmd3 As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd3
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
                .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
                .ExecuteNonQuery()
            End With

            '更新 STUD_BLIGATEDATA28E
            'sqlStr = "UPDATE STUD_BLIGATEDATA28E SET SETID=@newSETID WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum"
            'Dim uCmd4 As New SqlCommand(sqlStr, oConn, oTrans)
            'With uCmd4
            '    .Parameters.Clear()
            '    .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
            '    .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
            '    .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
            '    .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
            '    .ExecuteNonQuery()
            'End With

            '更新 CLASS_STUDENTSOFCLASS
            sqlStr = "UPDATE CLASS_STUDENTSOFCLASS SET SETID=@newSETID WHERE SETID=@SETID AND ETENTERDATE=@ETENTERDATE AND SERNUM=@SerNum"
            Dim uCmd5 As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd5
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
                .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
                .Parameters.Add("ETENTERDATE", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
                .ExecuteNonQuery()
            End With

            '更新 STUD_CONFIRM
            sqlStr = "UPDATE STUD_CONFIRM SET SETID=@newSETID WHERE SETID=@SETID AND EnterDate=@EnterDate AND SERNUM=@SerNum"
            Dim uCmd6 As New SqlCommand(sqlStr, oConn, oTrans)
            With uCmd6
                .Parameters.Clear()
                .Parameters.Add("newSETID", SqlDbType.Int).Value = Val(v_SETID_1)
                .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(v_SETID_2)
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(v_ENTERDATE)
                .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(v_SERNUM)
                .ExecuteNonQuery()
            End With

            DbAccess.CommitTrans(oTrans) 'objTrans.Commit()
        Catch ex As Exception
            rst = String.Concat("#ERROR:UPDATE_STUD_ENTERTYPE:", vbCrLf, ex.ToString())
            DbAccess.RollbackTrans(oTrans) 'objTrans.Rollback() '.Rollback("SETIDTrans")
            'Common.MessageBox(Me, ex.ToString)
            'If sqlAdp IsNot Nothing Then sqlAdp.Dispose()
            If oTrans IsNot Nothing Then oTrans.Dispose()
            'If Not objDS Is Nothing Then objDS.Dispose()
        End Try
        Return rst
    End Function

    ''' <summary>錯誤處置提示</summary>
    ''' <param name="pParms"></param>
    Private Sub SHOW_ERROR_HANDLING_TIPS(ByRef pParms As Hashtable)
        Dim v_SETID_1 As String = TIMS.GetMyValue2(pParms, "SETID_1") 'txt_EditClassSETID2.Text
        Dim v_SETID_2 As String = TIMS.GetMyValue2(pParms, "SETID_2") 'strPK(0)
        Dim v_ENTERDATE As String = TIMS.GetMyValue2(pParms, "ENTERDATE") 'strPK(1)
        Dim v_SERNUM As String = TIMS.GetMyValue2(pParms, "SERNUM") 'strPK(2)

        Dim sValue As String = ""
        sValue = String.Concat("/*", txt_EditIDNO2.Text, "*/", "<BR>")
        sValue &= String.Concat(" SELECT * FROM STUD_SELRESULT WHERE SETID='", v_SETID_1, "'")
        sValue &= String.Concat(" AND ENTERDATE=", TIMS.To_date(v_ENTERDATE))
        sValue &= String.Concat(" AND SERNUM='", v_SERNUM, "'", "<BR>")
        sValue &= String.Concat(" UNION ", "<BR>")
        sValue &= String.Concat(" SELECT * FROM STUD_SELRESULT WHERE SETID='", v_SETID_2, "'")
        sValue &= String.Concat(" AND ENTERDATE=", TIMS.To_date(v_ENTERDATE))
        sValue &= String.Concat(" AND SERNUM='", v_SERNUM, "'", "<BR>")
        sValue &= "<BR>"
        Common.RespWrite(Me, sValue)

        sValue = String.Concat("/*", txt_EditIDNO2.Text, "*/")
        sValue &= String.Concat("<BR>", " SELECT * FROM STUD_SELRESULT WHERE SETID IN ('", v_SETID_1, "','", v_SETID_2, "')")
        sValue &= String.Concat(" AND EnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>")
        sValue &= String.Concat("<BR>", " SELECT * FROM STUD_ENTERTYPE WHERE SETID IN ('", v_SETID_1, "','", v_SETID_2, "')")
        sValue &= String.Concat(" AND EnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>")
        sValue &= String.Concat("<BR>", " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SETID IN ('", v_SETID_1, "','", v_SETID_2, "')")
        sValue &= String.Concat(" AND ETEnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>")
        Common.RespWrite(Me, sValue)

        sValue = String.Concat("/*", txt_EditIDNO2.Text, "*/", "<BR>")
        sValue &= " /* <BR>"
        sValue &= " BEGIN <BR>"
        sValue &= String.Concat(" DECLARE @SerNum INT=", Val(v_SERNUM) + 1, ";<BR>")
        sValue &= String.Concat(" UPDATE Stud_SelResult ", "<BR>", " SET SerNum=@SerNum ", "<BR>")
        sValue &= String.Concat(" WHERE SETID='", v_SETID_2, "'<BR>", " and EnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>", " and SerNum='", v_SERNUM, "';<BR>", "<BR>")
        sValue &= String.Concat(" UPDATE Stud_entertype ", "<BR>", " SET SerNum=@SerNum ", "<BR>")
        sValue &= String.Concat(" WHERE SETID='", v_SETID_2, "'<BR>", " and EnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>", " and SerNum='", v_SERNUM, "';<BR>", "<BR>")
        sValue &= " END; <BR>"
        sValue &= " */ <BR>"
        sValue &= " /* OR */ <BR>"
        sValue &= " /* <BR>"
        sValue &= " BEGIN <BR>"
        sValue &= String.Concat(" DECLARE @SerNum INT=", Val(v_SERNUM) + 1, ";<BR>")
        sValue &= String.Concat(" UPDATE Stud_SelResult ", "<BR>", " SET SerNum=@SerNum ", "<BR>")
        sValue &= String.Concat(" WHERE SETID='", v_SETID_1, "'<BR>", " and EnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>", " and SerNum='", v_SERNUM, "';<BR>", "<BR>")
        sValue &= String.Concat(" UPDATE Stud_entertype ", "<BR>", " SET SerNum=@SerNum ", "<BR>")
        sValue &= String.Concat(" WHERE SETID='", v_SETID_1, "'<BR>", " and EnterDate=", TIMS.To_date(v_ENTERDATE), "<BR>", " and SerNum='", v_SERNUM, "';<BR>", "<BR>")
        sValue &= " END; <BR>"
        sValue &= " */ <BR>"
        Common.RespWrite(Me, sValue)

        sValue = "<br>" & vbCrLf
        sValue &= " select cc.classcname,b.setid,b.enterdate,b.sernum,b.ocid1<br>" & vbCrLf
        sValue &= " from stud_entertemp a <br>" & vbCrLf
        sValue &= " join Stud_entertype b on a.setid =b.setid<br>" & vbCrLf
        sValue &= " left join class_classinfo cc on cc.ocid=b.ocid1<br>" & vbCrLf
        sValue &= " where a.idno ='" & txt_EditIDNO2.Text & "'<br>" & vbCrLf
        sValue &= " and exists (<br>" & vbCrLf
        sValue &= " 	select bx.enterdate,bx.sernum<br>" & vbCrLf
        sValue &= " 	,count(1) cnt <br>" & vbCrLf
        'sValue &= " 	--,max(b.ocid1) maxocid1,min(b.ocid1) minocid1<br>" & vbCrLf
        sValue &= " 	from stud_entertemp ax <br>" & vbCrLf
        sValue &= " 	join Stud_entertype bx on ax.setid =bx.setid<br>" & vbCrLf
        sValue &= " 	where ax.idno ='" & txt_EditIDNO2.Text & "'<br>" & vbCrLf
        sValue &= " 	group by  bx.enterdate,bx.sernum<br>" & vbCrLf
        sValue &= " 	having count(1) >1<br>" & vbCrLf
        sValue &= " 	and bx.enterdate=b.enterdate <br>" & vbCrLf
        sValue &= " 	and bx.sernum=b.sernum<br>" & vbCrLf
        sValue &= " )<br>" & vbCrLf
        sValue &= "<BR>"
        Common.RespWrite(Me, sValue)
    End Sub

    ''' <summary>
    ''' '儲存課程變更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btn_EditSaveClass3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSaveClass3.Click
        'Dim sqlAdp As New SqlDataAdapter
        Dim v_rdo_EditClass3 As String = TIMS.GetListValue(rdo_EditClass3)
        If rdo_EditClass3.SelectedIndex = -1 OrElse v_rdo_EditClass3 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        If txt_EditClasseSETID3.Text = "" Then
            Common.MessageBox(Me, "eSETID不能沒填。")
            Exit Sub
        End If

        Dim pParms As New Hashtable
        pParms.Add("SETID", txt_EditClassSETID3.Text)
        pParms.Add("eSETID", txt_EditClasseSETID3.Text)
        pParms.Add("eSERNUM", v_rdo_EditClass3) 'TIMS.GetListValue(rdo_EditClass3)
        Call UPDATE_STUD_ENTERTYPE2B(objConn, pParms)
        iRst = 1

        'Common.MessageBox(Me, "修改完成。")
        Lab_ENTERTYPE2_mag.Text = String.Concat(Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), "，更新筆數：", iRst)
        Show_RdoEditClass3(Get_StudEnterType2("eSETID", lab_EditeSETID3.Text))
        tb_EditClass3.Visible = False

        'Try
        '    If rdo_EditClass3.SelectedIndex <> -1 Then
        '        If txt_EditClasseSETID3.Text = "" Then
        '            Common.MessageBox(Me, "eSETID不能沒填。")
        '        Else

        '        End If
        '    Else
        '        Common.MessageBox(Me, "請選取要處理的資料。")
        '    End If
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        'End Try
    End Sub

    Public Shared Sub UPDATE_STUD_ENTERTYPE2B(ByRef oConn As SqlConnection, ByRef pParms As Hashtable)
        Dim vSETID As String = TIMS.GetMyValue2(pParms, "SETID") 'txt_EditClassSETID3.Text
        Dim veSETID As String = TIMS.GetMyValue2(pParms, "eSETID") 'txt_EditClasseSETID3.Text
        Dim veSERNUM As String = TIMS.GetMyValue2(pParms, "eSERNUM") 'rdo_EditClass3.selectvalue
        If veSERNUM = "" Then Return
        'Dim v_rdo_EditClass3 As String = TIMS.GetListValue(rdo_EditClass3)

        'STUD_ENTERTYPE2 ,STUD_ENTERSUBDATA2 ,STUD_ENTERTRAIN2
        Dim sqlStr As String = "UPDATE STUD_ENTERTYPE2 set SETID= @SETID,eSETID= @eSETID where eSerNum= @eSerNum"
        Dim uCmd As New SqlCommand(sqlStr, oConn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("SETID", SqlDbType.Int).Value = If(vSETID = "", Convert.DBNull, vSETID)
            .Parameters.Add("eSETID", SqlDbType.Int).Value = veSETID
            .Parameters.Add("eSerNum", SqlDbType.Int).Value = veSERNUM
            .ExecuteNonQuery()
        End With

        'UPDATE STUD_BLIGATEDATA28E 
        Dim sqlStr2 As String = "UPDATE STUD_BLIGATEDATA28E SET eSETID=@eSETID WHERE eSerNum=@eSerNum AND eSETID IS NOT NULL"
        Dim uCmd2 As New SqlCommand(sqlStr2, oConn)
        With uCmd2
            .Parameters.Clear()
            .Parameters.Add("eSETID", SqlDbType.Int).Value = veSETID
            .Parameters.Add("eSerNum", SqlDbType.Int).Value = veSERNUM
            .ExecuteNonQuery()
        End With
    End Sub

    '確認資料的正確性囉  身分證號、生日、姓名不能沒填
    Function CheckData3(ByRef sIDNO As String, ByRef sBirthday As String, ByRef sName As String) As String
        Dim tmpMsg As String = "" '回傳錯誤訊息
        tmpMsg = ""

        sName = Trim(sName)
        If sIDNO = "" OrElse sBirthday = "" OrElse sName = "" Then tmpMsg += "身分證號、生日、姓名不能沒填。" & vbCrLf

        Dim sName_org As String = sName
        sName = TIMS.ClearSQM(sName)
        If sName <> sName_org Then tmpMsg += "姓名不能沒填。" & vbCrLf


        '1:國民身分證 -檢查
        Dim flagIdno1 As Boolean = TIMS.CheckIDNO(sIDNO) '檢查身分證號
        '2:居留證 4:居留證2021 -檢查
        Dim flagPermit2 As Boolean = TIMS.CheckIDNO2(sIDNO, 2) '可檢查居留證號1
        Dim flagPermit4 As Boolean = TIMS.CheckIDNO2(sIDNO, 4) '可檢查居留證號2
        If Not flagIdno1 AndAlso Not flagPermit2 AndAlso Not flagPermit4 Then tmpMsg += "身分證號 或 居留證號錯誤，請確認。" & vbCrLf

        If IsDate(sBirthday) = False Then tmpMsg += "輸入的生日格式錯誤，請確認。" & vbCrLf

        If tmpMsg = "" Then
            '0~119歲，應該很夠了吧!
            If Convert.ToDateTime(sBirthday) > Now() OrElse DateDiff(DateInterval.Year, Convert.ToDateTime(sBirthday), Now()) > 120 Then
                tmpMsg += "輸入的生日 範圍有誤，請確認。" & vbCrLf
            End If
        End If

        If vs_IDNO = "" Then vs_IDNO = sIDNO

        Return tmpMsg
    End Function

    Private Sub btn_EditSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSave1.Click
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""
        Dim tmpMsg As String = ""

        'Dim Trans As SqlTransaction = Nothing
        Dim flagTransBegin As Boolean = False
        Dim flagCommit As Boolean = False
        flagTransBegin = False
        flagCommit = False

        tmpMsg = CheckData3(txt_EditIDNO1.Text, txt_EditBirthday1.Text, txt_EditName1.Text)
        If tmpMsg <> "" Then
            Common.MessageBox(Me, tmpMsg)
            Exit Sub
        End If

        Const cst_COLUMN1 As String = "SID,IDNO,NAME,ENGNAME,PASSPORTNO,SEX,BIRTHDAY,MARITALSTATUS,DEGREEID,GRADUATESTATUS,MILITARYID,IDENTITYID,SUBSIDYID,JOBLESSID,REALJOBLESS,GETCERTIFICATE,GETSUBSIDY,ISAGREE,LAINFLAG,MODIFYACCT,MODIFYDATE,CHINAORNOT,NATIONALITY,PPNO,JOBSTATE,FTYPE,ACTNO,MDATE,SALID,FIXID,JOBLESSID_99,GRADUATEY,nationid,RMPNAME"
        'DbAccess.Open(objConn) 'If objConn.State = ConnectionState.Closed Then objConn.Open()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'Trans = objConn.BeginTransaction()
        Try
            flagTransBegin = True
            '修改 學員資料 人時
            sqlStr = "update Stud_StudentInfo set ModifyAcct=@ModifyAcct,ModifyDate=getdate() where SID=@SID"
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = lab_EditSID1.Text
                .UpdateCommand.ExecuteNonQuery()
            End With
            '備份 學員資料 人時
            sqlStr = String.Concat("insert into Stud_StudentInfoDelData (", cst_COLUMN1, ") select ", cst_COLUMN1, " from Stud_StudentInfo where ModifyAcct=@ModifyAcct AND SID=@SID")
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = lab_EditSID1.Text
                .InsertCommand.ExecuteNonQuery()
            End With

            '修改 學員資料 
            sqlStr = "update Stud_StudentInfo set IDNO=@IDNO,Name=@Name,Birthday=@Birthday where SID=@SID"
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = txt_EditIDNO1.Text
                .UpdateCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = txt_EditName1.Text
                .UpdateCommand.Parameters.Add("Birthday", SqlDbType.DateTime).Value = Convert.ToDateTime(txt_EditBirthday1.Text)

                .UpdateCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = lab_EditSID1.Text
                .UpdateCommand.ExecuteNonQuery()
            End With

            sqlStr = ""
            sqlStr &= " select 'X' " & vbCrLf
            sqlStr &= " from Stud_ResultStudData sp" & vbCrLf
            sqlStr &= " join stud_studentinfo ss on sp.stdpid=ss.idno" & vbCrLf
            sqlStr &= " where sp.stdName!=ss.Name" & vbCrLf
            sqlStr &= " and sp.stdpid=@IDNO2 " & vbCrLf
            Dim Parms As New Hashtable
            Parms.Add("IDNO2", txt_EditIDNO1.Text)
            Dim dt3 As DataTable = DbAccess.GetDataTable(sqlStr, oTrans, Parms)
            If dt3.Rows.Count > 0 Then
                '修改 結訓學員資料卡
                sqlStr = ""
                sqlStr &= " update Stud_ResultStudData" & vbCrLf
                sqlStr &= " set stdName = ( select max(name) name from stud_studentinfo where idno=@IDNO1 )" & vbCrLf
                sqlStr &= " where stdpid= @IDNO2 " & vbCrLf
                With sqlAdp
                    .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                    .UpdateCommand.Parameters.Clear()
                    .UpdateCommand.Parameters.Add("IDNO1", SqlDbType.VarChar).Value = txt_EditIDNO1.Text
                    .UpdateCommand.Parameters.Add("IDNO2", SqlDbType.VarChar).Value = txt_EditIDNO1.Text
                    .UpdateCommand.ExecuteNonQuery()
                End With
            End If

            oTrans.Commit()
            flagCommit = True

            Clear_Edit()
        Catch ex As Exception
            If flagCommit = False AndAlso flagTransBegin = True Then oTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            If sqlAdp IsNot Nothing Then sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub btn_EditSave2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSave2.Click
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""
        Dim tmpMsg As String = ""

        Dim flagTransBegin As Boolean = False
        Dim flagCommit As Boolean = False
        flagTransBegin = False
        flagCommit = False

        tmpMsg = CheckData3(txt_EditIDNO2.Text, txt_EditBirthday2.Text, txt_EditName2.Text)
        If tmpMsg <> "" Then
            Common.MessageBox(Me, tmpMsg)
            Exit Sub
        End If

        Const cst_COLUMN1 As String = "SETID,IDNO,NAME,SEX,BIRTHDAY,PASSPORTNO,MARITALSTATUS,DEGREEID,GRADID,SCHOOL,DEPARTMENT,MILITARYID,ZIPCODE,ADDRESS,PHONE1,PHONE2,CELLPHONE,EMAIL,PASSWORD,NOTES,ISAGREE,LAINFLAG,MODIFYACCT,MODIFYDATE,ESETID,ZIPCODE2W,ZIPCODE_N,ZIPCODE6W"

        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'Trans = objConn.BeginTransaction()
        Try
            flagTransBegin = True
            '修改 學員報名資料 人時
            sqlStr = "update Stud_EnterTemp set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.Int).Value = lab_EditSETID2.Text
                .UpdateCommand.ExecuteNonQuery()
            End With
            '備份 學員報名資料 人時
            sqlStr = String.Concat("insert into Stud_EnterTempDelData(", cst_COLUMN1, ")  select ", cst_COLUMN1, " from Stud_EnterTemp where ModifyAcct= @ModifyAcct AND SETID= @SETID ")
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.Parameters.Add("SETID", SqlDbType.Int).Value = lab_EditSETID2.Text
                .InsertCommand.ExecuteNonQuery()
            End With
            '修改 學員報名資料
            sqlStr = "update Stud_EnterTemp set IDNO= @IDNO,Name= @Name,Birthday= @Birthday,eSETID= @eSETID where SETID= @SETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = txt_EditIDNO2.Text
                .UpdateCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = txt_EditName2.Text
                .UpdateCommand.Parameters.Add("Birthday", SqlDbType.DateTime).Value = Convert.ToDateTime(txt_EditBirthday2.Text)
                .UpdateCommand.Parameters.Add("eSETID", SqlDbType.Int).Value = If(txt_EditeSETID2.Text = "", Convert.DBNull, txt_EditeSETID2.Text)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.Int).Value = lab_EditSETID2.Text
                .UpdateCommand.ExecuteNonQuery()
            End With
            oTrans.Commit()
            flagCommit = True

            Clear_Edit()
        Catch ex As Exception
            If flagCommit = False AndAlso flagTransBegin = True Then oTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub btn_EditSave3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSave3.Click
        Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""
        Dim tmpMsg As String = ""

        'Dim Trans As SqlTransaction = Nothing
        Dim flagTransBegin As Boolean = False
        Dim flagCommit As Boolean = False
        flagTransBegin = False
        flagCommit = False

        tmpMsg = CheckData3(txt_EditIDNO3.Text, txt_EditBirthday3.Text, txt_EditName3.Text)
        If tmpMsg <> "" Then
            Common.MessageBox(Me, tmpMsg)
            Exit Sub
        End If

        Const cst_COLUMN1 As String = "ESETID,SETID,IDNO,NAME,SEX,BIRTHDAY,PASSPORTNO,MARITALSTATUS,DEGREEID,GRADID,SCHOOL,DEPARTMENT,MILITARYID,ZIPCODE,ADDRESS,PHONE1,PHONE2,CELLPHONE,EMAIL,ISAGREE,MODIFYACCT,MODIFYDATE,ZIPCODE2W,LAINFLAG,ZIPCODE_N,ZIPCODE6W"

        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'Trans = objConn.BeginTransaction()
        Try
            flagTransBegin = True
            '修改 學員報名資料2 人時
            sqlStr = "update Stud_EnterTemp2 set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where eSETID= @eSETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("eSETID", SqlDbType.Int).Value = lab_EditeSETID3.Text
                .UpdateCommand.ExecuteNonQuery()
            End With
            '備份 學員報名資料2 人時
            sqlStr = String.Concat("insert into Stud_EnterTemp2DelData(", cst_COLUMN1, ") select ", cst_COLUMN1, " from Stud_EnterTemp2 where ModifyAcct= @ModifyAcct AND eSETID= @eSETID ")
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .InsertCommand.Parameters.Add("eSETID", SqlDbType.Int).Value = lab_EditeSETID3.Text
                .InsertCommand.ExecuteNonQuery()
            End With
            '修改 學員報名資料2
            sqlStr = "update Stud_EnterTemp2 set IDNO= @IDNO,Name= @Name,Birthday= @Birthday,SETID= @SETID where eSETID= @eSETID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("IDNO", SqlDbType.VarChar).Value = txt_EditIDNO3.Text
                .UpdateCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = txt_EditName3.Text
                .UpdateCommand.Parameters.Add("Birthday", SqlDbType.DateTime).Value = txt_EditBirthday3.Text
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.Int).Value = If(txt_EditSETID3.Text = "", Convert.DBNull, txt_EditSETID3.Text)
                .UpdateCommand.Parameters.Add("eSETID", SqlDbType.Int).Value = lab_EditeSETID3.Text
                .UpdateCommand.ExecuteNonQuery()
            End With
            oTrans.Commit()
            flagCommit = True

            Clear_Edit()
        Catch ex As Exception
            If flagCommit = False AndAlso flagTransBegin = True Then oTrans.Rollback()
            Common.MessageBox(Me, ex.ToString)
            If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        End Try
    End Sub

    Private Sub btn_EditSave10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditSave10.Click
        'Dim sqlAdp As SqlDataAdapter = TIMS.GetOneDA("SU", objConn)
        Dim sqlStr As String = ""
        Dim tmpMsg As String = ""

        tmpMsg = CheckData3(txt_EditIdno10.Text, txt_BirthDay10.Text, txt_EditName10.Text)
        If tmpMsg <> "" Then
            Common.MessageBox(Me, tmpMsg)
            Exit Sub
        End If

        Try
            If tmpMsg = "" Then
                'Dim dt As DataTable = Nothing
                sqlStr = "select mem_IDNO,mem_sn from e_Member WHERE mem_IDNO= @mem_IDNO AND mem_sn!= @mem_sn "
                Dim sCmd As New SqlCommand(sqlStr, objConn)
                'Call TIMS.OpenDbConn(objConn)
                Dim dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("mem_IDNO", SqlDbType.VarChar).Value = txt_EditIdno10.Text
                    .Parameters.Add("mem_sn", SqlDbType.VarChar).Value = lab_mem_sn.Text
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count > 0 Then
                    tmpMsg += "該身分證號有重複，請查明後再修改!!" & vbCrLf
                End If
            End If

            If tmpMsg = "" Then
                'Dim dt As DataTable = Nothing
                sqlStr = "select mem_IDNO,mem_sn from e_Member WHERE mem_IDNO= @mem_IDNO AND mem_sn= @mem_sn AND STOPMEM='Y' "
                Dim sCmd As New SqlCommand(sqlStr, objConn)
                'Call TIMS.OpenDbConn(objConn)
                Dim dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("mem_IDNO", SqlDbType.VarChar).Value = txt_EditIdno10.Text
                    .Parameters.Add("mem_sn", SqlDbType.VarChar).Value = lab_mem_sn.Text
                    dt.Load(.ExecuteReader())
                End With
                'With sqlAdp
                '    .SelectCommand.CommandText = sqlStr
                '    .SelectCommand.Parameters.Clear()
                '    .SelectCommand.Parameters.Add("mem_IDNO", SqlDbType.VarChar).Value = txt_EditIdno10.Text
                '    .SelectCommand.Parameters.Add("mem_sn", SqlDbType.VarChar).Value = lab_mem_sn.Text
                '    dt = New DataTable
                '    .Fill(dt)
                'End With
                If dt.Rows.Count > 0 Then
                    tmpMsg += "該身分證號已停用，請查明後再修改!!" & vbCrLf
                End If
            End If

            If tmpMsg = "" Then
                sqlStr = "update e_Member set mem_IDNO= @mem_IDNO,mem_Name= @mem_Name,mem_birth= @mem_birth where mem_sn= @mem_sn "
                Dim uCmd As New SqlCommand(sqlStr, objConn)
                'Call TIMS.OpenDbConn(objConn)
                'Dim dt As New DataTable
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("mem_IDNO", SqlDbType.VarChar).Value = txt_EditIdno10.Text
                    .Parameters.Add("mem_Name", SqlDbType.NVarChar).Value = txt_EditName10.Text
                    .Parameters.Add("mem_birth", SqlDbType.DateTime).Value = Convert.ToDateTime(txt_BirthDay10.Text)
                    .Parameters.Add("mem_sn", SqlDbType.VarChar).Value = lab_mem_sn.Text
                    .ExecuteNonQuery()
                End With

                'With sqlAdp
                '    .UpdateCommand.CommandText = sqlStr
                '    .UpdateCommand.Parameters.Clear()
                '    .UpdateCommand.Parameters.Add("mem_IDNO", SqlDbType.VarChar).Value = txt_EditIdno10.Text
                '    .UpdateCommand.Parameters.Add("mem_Name", SqlDbType.NVarChar).Value = txt_EditName10.Text
                '    .UpdateCommand.Parameters.Add("mem_birth", SqlDbType.DateTime).Value = Convert.ToDateTime(txt_BirthDay10.Text)
                '    .UpdateCommand.Parameters.Add("mem_sn", SqlDbType.VarChar).Value = lab_mem_sn.Text
                '    If .UpdateCommand.Connection.State = ConnectionState.Closed Then .UpdateCommand.Connection.Open()
                '    .UpdateCommand.ExecuteNonQuery()
                '    If .UpdateCommand.Connection.State = ConnectionState.Open Then .UpdateCommand.Connection.Close()
                'End With
                Clear_Edit()
            End If

        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            'If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
        End Try

        If tmpMsg <> "" Then
            Common.MessageBox(Me, tmpMsg)
            Exit Sub
        End If

        Common.MessageBox(Me, tmpMsg)
    End Sub

    Private Sub btn_EditCanedl10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditCanedl10.Click
        Clear_Edit()
    End Sub

    '刪Class_StudentsOfClass
    Private Sub btn_EditDelClass1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditdelClass1.Click
        'Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = ""
        'Dim objTrans As SqlTransaction = Nothing
        Dim v_rdo_EditClass1 As String = TIMS.GetListValue(rdo_EditClass1)

        If rdo_EditClass1.SelectedIndex = -1 OrElse v_rdo_EditClass1 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Return
        End If

        Dim sqlAdp As New SqlDataAdapter
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'objTrans = objConn.BeginTransaction()
        Try

            '變更Modify人時
            sqlStr = "update Class_StudentsOfClass set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SOCID= @SOCID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rdo_EditClass1.SelectedValue
                .UpdateCommand.ExecuteNonQuery()
            End With
            '搬移資料
            sqlStr = "insert into Class_StudentsOfClassDelData select * from Class_StudentsOfClass where SOCID= @SOCID "
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rdo_EditClass1.SelectedValue
                .InsertCommand.ExecuteNonQuery()
            End With
            '刪除Class_StudentsOfClass
            sqlStr = "delete Class_StudentsOfClass where SOCID= @SOCID "
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rdo_EditClass1.SelectedValue
                .DeleteCommand.ExecuteNonQuery()
            End With
            '刪除Sub_SubSidyApply
            'sqlStr = "delete Sub_SubSidyApply where SOCID= @SOCID "
            'With sqlAdp
            '    .DeleteCommand = New SqlCommand(sqlStr, objConn, objTrans)
            '    .DeleteCommand.Parameters.Clear()
            '    .DeleteCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rdo_EditClass1.SelectedValue
            '    .DeleteCommand.ExecuteNonQuery()
            'End With
            oTrans.Commit()
        Catch ex As Exception
            oTrans.Rollback() '.Rollback("SIDTrans")
            Common.MessageBox(Me, ex.ToString)
            If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
            If Not oTrans Is Nothing Then oTrans.Dispose()
        End Try
    End Sub

    '刪 Stud_EnterType .Stud_SelResult
    Private Sub btn_EditDelClass2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditdelClass2.Click
        Dim v_rdo_EditClass2 As String = TIMS.GetListValue(rdo_EditClass2)
        If v_rdo_EditClass2 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        Dim strPK() As String = Split(v_rdo_EditClass2, ";")

        Dim sqlStr As String = ""
        'Dim objTrans As SqlTransaction = Nothing

        Dim sqlAdp As New SqlDataAdapter
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'objTrans = objConn.BeginTransaction()
        Try

            '變更Modify人時
            sqlStr = "update Stud_EnterType set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .UpdateCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .UpdateCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .UpdateCommand.ExecuteNonQuery()
            End With
            '搬移資料
            sqlStr = "insert into Stud_EnterTypeDelData select * from Stud_EnterType where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum "
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .InsertCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .InsertCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .InsertCommand.ExecuteNonQuery()
            End With
            '刪除Stud_EnterType
            sqlStr = "delete Stud_EnterType where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum "
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .DeleteCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .DeleteCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .DeleteCommand.ExecuteNonQuery()
            End With
            '變更Modify人時
            sqlStr = "update Stud_SelResult set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .UpdateCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .UpdateCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .UpdateCommand.ExecuteNonQuery()
            End With
            '搬移資料
            sqlStr = "insert into Stud_SelResultDelData select * from Stud_SelResult where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum "
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .InsertCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .InsertCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .InsertCommand.ExecuteNonQuery()
            End With
            '刪除Stud_SelResult
            sqlStr = "delete Stud_SelResult where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum "
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .DeleteCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .DeleteCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .DeleteCommand.ExecuteNonQuery()
            End With
            oTrans.Commit()
            Common.MessageBox(Me, "刪除完成。")
            Call Show_RdoEditClass2(Get_StudEnterType("SETID", lab_EditSETID2.Text))

            tb_EditClass2.Visible = False
        Catch ex As Exception
            oTrans.Rollback() '.Rollback("SETIDTrans")
            Common.MessageBox(Me, ex.ToString)
            If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
            If Not oTrans Is Nothing Then oTrans.Dispose()
        End Try
    End Sub

    '取消課程報到資料(刪除學員資料，取消課程報到)
    Protected Sub btn_EditUpdateCls2_Click(sender As Object, e As EventArgs) Handles btn_EditUpdateCls2.Click
        Dim v_rdo_EditClass2 As String = TIMS.GetListValue(rdo_EditClass2)
        If v_rdo_EditClass2 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        Dim strPK() As String = Split(v_rdo_EditClass2, ";")

        Dim sql As String = ""
        sql &= " SELECT r.SETID" & vbCrLf '/*PK*/ 
        sql &= " ,r.ENTERDATE" & vbCrLf '/*PK*/ 
        sql &= " ,r.SERNUM" & vbCrLf '/*PK*/ 
        sql &= " ,r.OCID" & vbCrLf '/*PK*/ 
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " FROM STUD_SELRESULT r" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b on b.SETID=r.SETID AND b.ENTERDATE=r.ENTERDATE AND b.SERNUM=r.SERNUM AND r.OCID =b.OCID1" & vbCrLf
        sql &= " JOIN STUD_ENTERTemp a on a.SETID = b.SETID " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND r.SETID= @SETID AND r.EnterDate= @EnterDate AND r.SerNum= @SerNum "
        Dim dt1 As New DataTable
        Dim sCmd As New SqlCommand(sql, objConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
            .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
            .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count <> 1 Then
            Common.MessageBox(Me, "查無有效資料取消動作1。")
            Exit Sub
        End If
        Dim dr1 As DataRow = dt1.Rows(0)

        sql = "" & vbCrLf
        sql &= " select cs.SOCID " & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs " & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid " & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.ocid =cs.ocid " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND ss.IDNO=@IDNO" & vbCrLf
        sql &= " AND cs.OCID=@OCID"
        Dim dt2 As New DataTable
        Dim sCmd2 As New SqlCommand(sql, objConn)
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = Convert.ToString(dr1("IDNO"))
            .Parameters.Add("OCID", SqlDbType.Int).Value = Convert.ToInt32(dr1("OCID"))
            dt2.Load(.ExecuteReader())
        End With
        If dt2.Rows.Count <> 1 Then
            Common.MessageBox(Me, "查無有效資料取消動作2。")
            Exit Sub
        End If
        Dim dr2 As DataRow = dt2.Rows(0)

        Dim sqlStr As String = ""
        Dim sqlAdp As New SqlDataAdapter
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'objTrans = objConn.BeginTransaction()
        Try
            '變更Modify人時
            sqlStr = "update Class_StudentsOfClass set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SOCID= @SOCID "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = Convert.ToInt32(dr2("SOCID")) 'rdo_EditClass1.SelectedValue
                .UpdateCommand.ExecuteNonQuery()
            End With
            '搬移資料
            sqlStr = "insert into Class_StudentsOfClassDelData select * from Class_StudentsOfClass where SOCID= @SOCID "
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = Convert.ToInt32(dr2("SOCID")) 'rdo_EditClass1.SelectedValue
                .InsertCommand.ExecuteNonQuery()
            End With
            '刪除Class_StudentsOfClass
            sqlStr = "delete Class_StudentsOfClass where SOCID= @SOCID "
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = Convert.ToInt32(dr2("SOCID")) 'rdo_EditClass1.SelectedValue
                .DeleteCommand.ExecuteNonQuery()
            End With
            '變更Modify人時
            sqlStr = ""
            sqlStr &= " UPDATE STUD_SELRESULT"
            sqlStr &= " SET APPLIEDSTATUS='N',ModifyAcct= @ModifyAcct,ModifyDate=getdate()"
            sqlStr &= " WHERE SETID= @SETID"
            sqlStr &= " AND EnterDate= @EnterDate"
            sqlStr &= " AND SerNum= @SerNum "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                .UpdateCommand.Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                .UpdateCommand.Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                .UpdateCommand.ExecuteNonQuery()
            End With
            oTrans.Commit()
            Common.MessageBox(Me, "刪除學員資料，取消課程報到 完成。")
            Call Show_RdoEditClass2(Get_StudEnterType("SETID", lab_EditSETID2.Text))

            tb_EditClass2.Visible = False
        Catch ex As Exception
            oTrans.Rollback() '.Rollback("SETIDTrans")
            Common.MessageBox(Me, ex.ToString)
            If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
            If Not oTrans Is Nothing Then oTrans.Dispose()
        End Try
    End Sub

    '刪Stud_EnterType2
    Private Sub btn_EditDelClass3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_EditdelClass3.Click
        Dim sqlAdp As New SqlDataAdapter

        If rdo_EditClass3.SelectedIndex = -1 Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If

        Dim sqlStr As String = ""
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objConn) 'objTrans = objConn.BeginTransaction()
        Try
            '變更Modify人時
            sqlStr = "update Stud_EnterType2 set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where eSerNum= @eSerNum "
            With sqlAdp
                .UpdateCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .UpdateCommand.Parameters.Clear()
                .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                .UpdateCommand.Parameters.Add("eSerNum", SqlDbType.Int).Value = rdo_EditClass3.SelectedValue
                .UpdateCommand.ExecuteNonQuery()
            End With
            '搬移資料
            sqlStr = "insert into Stud_EnterType2DelData select * from Stud_EnterType2 where eSerNum= @eSerNum "
            With sqlAdp
                .InsertCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .InsertCommand.Parameters.Clear()
                .InsertCommand.Parameters.Add("eSerNum", SqlDbType.Int).Value = rdo_EditClass3.SelectedValue
                .InsertCommand.ExecuteNonQuery()
            End With
            '刪除 Stud_EnterType2
            sqlStr = "delete Stud_EnterType2 where eSerNum= @eSerNum "
            With sqlAdp
                .DeleteCommand = New SqlCommand(sqlStr, objConn, oTrans)
                .DeleteCommand.Parameters.Clear()
                .DeleteCommand.Parameters.Add("eSerNum", SqlDbType.Int).Value = rdo_EditClass3.SelectedValue
                .DeleteCommand.ExecuteNonQuery()
            End With
            oTrans.Commit()
            Common.MessageBox(Me, "刪除完成。")
            Show_RdoEditClass3(Get_StudEnterType2("eSETID", lab_EditeSETID3.Text))

            tb_EditClass3.Visible = False
        Catch ex As Exception
            oTrans.Rollback() '.Rollback("eSETIDTrans")
            Common.MessageBox(Me, ex.ToString)
            If Not sqlAdp Is Nothing Then sqlAdp.Dispose()
            If Not oTrans Is Nothing Then oTrans.Dispose()
        End Try
    End Sub

    Private Sub Datagrid4_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid4.ItemCommand
        Select Case e.CommandName
            Case "btnDelete"
                Dim sql As String = ""
                Dim vs_SOCID As String = ""
                Dim vs_courid As String = ""
                vs_SOCID = TIMS.GetMyValue(e.CommandArgument, "SOCID")
                vs_courid = TIMS.GetMyValue(e.CommandArgument, "courid")
                vs_IDNO = TIMS.GetMyValue(e.CommandArgument, "idno")
                vs_IDNO = UCase(vs_IDNO)

                vs_SOCID = TIMS.ClearSQM(vs_SOCID)
                vs_courid = TIMS.ClearSQM(vs_courid)
                sql = "delete Stud_TrainingResults where SOCID='" & vs_SOCID & "' AND courid='" & vs_courid & "' "
                DbAccess.ExecuteNonQuery(sql, objConn)
                'Common.MessageBox(Me, "清除成功！")

                Call Show_DataGrid4(vs_IDNO, 0)

        End Select
    End Sub

    Private Sub Datagrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo4")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                Dim btnDelete As LinkButton = e.Item.FindControl("btnDelete")
                btnDelete.Enabled = False
                If Me.ViewState(Cst_DelVal1) = "Y" Then
                    btnDelete.Enabled = True

                    btnDelete.CommandArgument = "btnDelete=1"
                    btnDelete.CommandArgument += "&SOCID=" & drv("SOCID")
                    btnDelete.CommandArgument += "&courid=" & drv("courid")
                    btnDelete.CommandArgument += "&idno=" & drv("idno")
                    btnDelete.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                Else
                    TIMS.Tooltip(btnDelete, "刪除功能請系統管理者處理!!")
                End If

        End Select
    End Sub

    Private Sub Datagrid5_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid5.ItemCommand
        Select Case e.CommandName
            Case "btnUpdate"
                Dim vs_OCID As String = ""
                Dim vs_SOCID2 As String = ""
                vs_OCID = TIMS.GetMyValue(e.CommandArgument, "OCID")
                vs_SOCID2 = TIMS.GetMyValue(e.CommandArgument, "SOCID2")
                vs_IDNO = TIMS.GetMyValue(e.CommandArgument, "idno")
                vs_IDNO = UCase(vs_IDNO)
                If vs_SOCID2 = "" Then
                    Common.MessageBox(Me, "無班級學員序號!!")
                    Exit Sub
                End If
                If vs_SOCID2 <> "" Then
                    '送3合1 'Call Update_GOVTRNData(vs_OCID, vs_IDNO, vs_SOCID2)
                    Show_DataGrid5(vs_IDNO, 0)
                End If
        End Select
    End Sub

    Private Sub Datagrid5_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid5.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo5")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                Dim btnUpdate As LinkButton = e.Item.FindControl("btnUpdate")
                Dim sCmdArg As String = ""
                sCmdArg = "btnUpdate=1"
                sCmdArg += "&idno=" & TIMS.ClearSQM(drv("idno"))
                sCmdArg += "&OCID=" & TIMS.ClearSQM(drv("OCID"))
                sCmdArg += "&SOCID2=" & TIMS.ClearSQM(drv("SOCID2"))
                btnUpdate.CommandArgument = sCmdArg
                btnUpdate.Attributes("onclick") = "return confirm('您確定要送出這一筆資料?');"

                btnUpdate.Enabled = True '可送出
                If Convert.ToString(drv("SOCID2")) = "" Then
                    btnUpdate.CommandArgument = ""
                    btnUpdate.Enabled = False
                    TIMS.Tooltip(btnUpdate, "無班級學員序號", True)
                End If
                If Convert.ToString(drv("SOCID1")) <> "" Then
                    btnUpdate.CommandArgument = ""
                    btnUpdate.Enabled = False
                    TIMS.Tooltip(btnUpdate, "已送出", True)
                End If


        End Select
    End Sub

    Private Sub Datagrid6_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles Datagrid6.ItemCommand
        'Dim btnDG6UPD1 As LinkButton = e.Item.FindControl("btnDG6UPD1")
        If e.CommandName = "" Then Return
        If e.CommandArgument = "" Then Return
        '<asp:LinkButton ID = "btnDG6SCH1A" runat="server" Text="查詢重複" CommandName="SCH1A" CssClass="linkbutton"></asp: LinkButton<> br />
        '                                   <asp:LinkButton ID="btnDG6UPD1A" runat="server" Text="(產投)<br/>暫時離訓" CommandName="UPD1A" CssClass="linkbutton"></asp:LinkButton> < br />
        '                                   <asp:LinkButton ID="btnDG6UPD1B" runat="server" Text="還原<br/>(暫時離訓)" CommandName="UPD1B" CssClass="linkbutton"></asp:LinkButton>
        '                               <
        Select Case e.CommandName
            Case "SCH1A" '查詢重複
                'WebRequest物件如何忽略憑證問題
                System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
                'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
                System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
                '檢核學員重複參訓。
                'http://163.29.199.211/TIMSWS/timsService1.asmx
                'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
                Dim timsSer1 As New timsService1.timsService1

                '檢核學員重複參訓。
                'Dim aIDNO1 As String = CStr(drET2("IDNO"))
                Dim aOCID1 As String = TIMS.GetMyValue(e.CommandArgument, "ocid")
                Dim aIDNO1 As String = TIMS.GetMyValue(e.CommandArgument, "idno")
                aOCID1 = TIMS.ClearSQM(aOCID1)
                aIDNO1 = TIMS.ClearSQM(aIDNO1)
                Dim ERRMSG As String = ""
                '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
                Dim xStudInfo As String = String.Format("&IDNO={0}&OCID1={1}&STEST=Y", aIDNO1, aOCID1)
                Call TIMS.ChkStudDouble(timsSer1, ERRMSG, "", xStudInfo)
                If ERRMSG = "" Then ERRMSG = "(OK)該班未發生重複"
                Common.MessageBox(Me, ERRMSG)
                Return

            Case "UPD1A" '(產投)暫時離訓
                'Dim sql As String = ""
                Dim vs_idno As String = TIMS.GetMyValue(e.CommandArgument, "idno")
                Dim vs_ocid As String = TIMS.GetMyValue(e.CommandArgument, "ocid")
                Dim vs_socid As String = TIMS.GetMyValue(e.CommandArgument, "socid")
                Dim vs_studstatus As String = TIMS.GetMyValue(e.CommandArgument, "studstatus")
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("STUDSTATUS", vs_studstatus)
                parms.Add("socid", vs_socid)
                parms.Add("ocid", vs_ocid)
                Dim sql As String = ""
                sql &= " update CLASS_STUDENTSOFCLASS "
                sql &= " set STUDSTATUS=2 "
                sql &= " FROM CLASS_STUDENTSOFCLASS "
                sql &= " where STUDSTATUS=@STUDSTATUS "
                sql &= " and socid =@socid"
                sql &= " and ocid =@ocid "
                DbAccess.ExecuteNonQuery(sql, objConn, parms)

                Call Show_DataGrid6(vs_idno, 0)
            Case "UPD1B" '還原(暫時離訓)
                'Dim sql As String = ""
                Dim vs_idno As String = TIMS.GetMyValue(e.CommandArgument, "idno")
                Dim vs_ocid As String = TIMS.GetMyValue(e.CommandArgument, "ocid")
                Dim vs_socid As String = TIMS.GetMyValue(e.CommandArgument, "socid")
                Dim vs_studstatus As String = TIMS.GetMyValue(e.CommandArgument, "studstatus")
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("STUDSTATUS", vs_studstatus)
                parms.Add("socid", vs_socid)
                parms.Add("ocid", vs_ocid)
                Dim sql As String = ""
                sql &= " update CLASS_STUDENTSOFCLASS "
                sql &= " set STUDSTATUS=1 "
                sql &= " FROM CLASS_STUDENTSOFCLASS "
                sql &= " where STUDSTATUS=@STUDSTATUS "
                sql &= " and socid =@socid"
                sql &= " and ocid =@ocid "
                DbAccess.ExecuteNonQuery(sql, objConn, parms)

                Call Show_DataGrid6(vs_idno, 0)
        End Select
    End Sub

    ''' <summary>
    ''' '檢核按鈕是否顯示 與 加入參數
    ''' </summary>
    ''' <param name="eItem"></param>
    Sub CHECK_DG6(ByRef eItem As DataGridItem)
        Dim drv As DataRowView = eItem.DataItem
        Dim sCmdArg As String = ""
        'Dim labSNo As Label = eItem.FindControl("lab_SNo6")
        Dim btnDG6UPD1A As LinkButton = eItem.FindControl("btnDG6UPD1A")
        Dim btnDG6UPD1B As LinkButton = eItem.FindControl("btnDG6UPD1B")
        Dim btnDG6SCH1A As LinkButton = eItem.FindControl("btnDG6SCH1A")

        btnDG6UPD1A.Visible = False
        btnDG6UPD1B.Visible = False
        btnDG6SCH1A.Visible = False
        '年度2021，且為產投，且為在訓
        If Convert.ToString(drv("TPLANID")) = "28" Then
            If Convert.ToString(drv("CanUseYears28")) = "Y" Then btnDG6SCH1A.Visible = True '可查詢是否重複

            If Convert.ToString(drv("CanUseYears28")) = "Y" AndAlso Convert.ToString(drv("studstatus")) = "1" Then btnDG6UPD1A.Visible = True '可執行暫時離訓
            '產投/離訓但沒有離訓日期，異常資料
            If Convert.ToString(drv("studstatus")) = "2" AndAlso Convert.ToString(drv("RejectTDate1")) = "" Then btnDG6UPD1B.Visible = True '可執行暫時離訓-還原

            sCmdArg = ""
            TIMS.SetMyValue(sCmdArg, "TPLANID", Convert.ToString(drv("TPLANID")))
            TIMS.SetMyValue(sCmdArg, "idno", Convert.ToString(drv("idno")))
            TIMS.SetMyValue(sCmdArg, "socid", Convert.ToString(drv("socid")))
            TIMS.SetMyValue(sCmdArg, "ocid", Convert.ToString(drv("ocid")))
            TIMS.SetMyValue(sCmdArg, "studstatus", Convert.ToString(drv("studstatus")))
            If btnDG6UPD1A.Visible Then btnDG6UPD1A.CommandArgument = sCmdArg
            If btnDG6UPD1B.Visible Then btnDG6UPD1B.CommandArgument = sCmdArg
            If btnDG6SCH1A.Visible Then btnDG6SCH1A.CommandArgument = sCmdArg
        End If

        Const Cst_StudStatus As Integer = 8
        Dim StudStatus As String = ""
        Select Case Convert.ToString(drv("StudStatus"))
            Case "1"
                StudStatus = "1:在訓"
            Case "2"
                StudStatus = "2:離訓"
            Case "3"
                StudStatus = "3:退訓"
            Case "4"
                StudStatus = "4:續訓"
            Case "5"
                StudStatus = "5:結訓"
            Case Else
                StudStatus = Convert.ToString(drv("StudStatus"))
        End Select
        eItem.Cells(Cst_StudStatus).Text = StudStatus
    End Sub

    Private Sub Datagrid6_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid6.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                'Dim drv As DataRowView = e.Item.DataItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo6")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
                '檢核按鈕是否顯示 與 加入參數
                Call CHECK_DG6(e.Item)
        End Select
    End Sub

    Private Sub Datagrid7_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid7.ItemCommand
        Select Case e.CommandName
            Case "btnDelete"
                Dim vs_SOCID As String = ""
                Dim vs_DLID As String = ""
                Dim vs_SubNo As String = ""
                vs_SOCID = TIMS.GetMyValue(e.CommandArgument, "SOCID")
                vs_DLID = TIMS.GetMyValue(e.CommandArgument, "DLID")
                vs_SubNo = TIMS.GetMyValue(e.CommandArgument, "SubNo")
                vs_IDNO = TIMS.GetMyValue(e.CommandArgument, "IDNO")
                vs_IDNO = TIMS.ClearSQM(UCase(vs_IDNO))

                vs_DLID = TIMS.ClearSQM(vs_DLID)
                vs_SubNo = TIMS.ClearSQM(vs_SubNo)
                Dim sql As String = ""
                sql = "Delete Stud_ResultStudData  WHERE DLID='" & vs_DLID & "' AND SubNo='" & vs_SubNo & "'"
                DbAccess.ExecuteNonQuery(sql, objConn)
                'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
                'BY AMU 2009-07-30
                '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
                sql = "Delete Stud_ResultIdentData  WHERE DLID='" & vs_DLID & "' AND SubNo='" & vs_SubNo & "'"
                DbAccess.ExecuteNonQuery(sql, objConn)
                sql = "Delete Stud_ResultTwelveData  WHERE DLID='" & vs_DLID & "' AND SubNo='" & vs_SubNo & "'"
                DbAccess.ExecuteNonQuery(sql, objConn)
                'Common.MessageBox(Me, "清除成功！")
                Show_DataGrid7(vs_IDNO, 0)
        End Select
    End Sub

    Private Sub Datagrid7_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid7.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo7")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                Dim btnDelete7 As LinkButton = e.Item.FindControl("btnDelete7")
                btnDelete7.Enabled = False

                Dim xdr As DataRow = Nothing
                Dim xSql As String = ""
                xSql = "SELECT '結訓成績資料' as Reason FROM Stud_TrainingResults WHERE SOCID='" & drv("SOCID") & "' and Results>0"
                xdr = DbAccess.GetOneRow(xSql, objConn)

                If Me.ViewState(Cst_DelVal1) = "Y" OrElse xdr Is Nothing Then
                    btnDelete7.Enabled = True

                    btnDelete7.CommandArgument = "btnDelete=1"
                    btnDelete7.CommandArgument += "&SOCID=" & drv("SOCID")
                    btnDelete7.CommandArgument += "&DLID=" & drv("DLID")
                    btnDelete7.CommandArgument += "&SubNo=" & drv("SubNo")
                    btnDelete7.CommandArgument += "&IDNO=" & drv("IDNO")

                    btnDelete7.Enabled = False
                    'SELECT '結訓成績資料' as Reason FROM Stud_TrainingResults WHERE SOCID='" & drv("SOCID") & "' and Results>0
                    If Convert.ToString(drv("DLID")) <> "" AndAlso Convert.ToString(drv("SubNo")) <> "" Then
                        btnDelete7.Enabled = True
                        btnDelete7.Attributes("onclick") = "return confirm('您確定要清除這一筆資料?');"
                    Else
                        TIMS.Tooltip(btnDelete7, "已清除", True)
                    End If
                Else
                    TIMS.Tooltip(btnDelete7, "刪除功能請系統管理者處理!!")
                End If
        End Select
    End Sub

    Private Sub Datagrid10_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid10.ItemCommand
        If e Is Nothing Then Return
        If e.CommandArgument Is Nothing Then Return
        If Convert.ToString(e.CommandArgument) = "" Then Return

        lab_mem_sn.Text = TIMS.GetMyValue(e.CommandArgument, "mem_sn")
        txt_EditIdno10.Text = TIMS.GetMyValue(e.CommandArgument, "mem_idno")
        txt_EditName10.Text = TIMS.GetMyValue(e.CommandArgument, "mem_name")
        txt_BirthDay10.Text = TIMS.GetMyValue(e.CommandArgument, "mem_birth")
        vs_IDNO = txt_EditIdno10.Text.Trim(" ")

        Select Case e.CommandName
            Case "btnUpdata"
                tr_Info.Visible = False
                tr_Edit10.Visible = True
            Case "btnDelete"
                Del_E_MEMBER(lab_mem_sn.Text, objConn)
                Show_DataGrid10(vs_IDNO, 0)
            Case "btnStop"
                Stop_E_MEMBER(lab_mem_sn.Text, 1)
                Show_DataGrid10(vs_IDNO, 0)
            Case "btnUnStop"
                Stop_E_MEMBER(lab_mem_sn.Text, 2)
                Show_DataGrid10(vs_IDNO, 0)
        End Select
    End Sub

    Private Sub Datagrid10_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid10.ItemDataBound
        Dim labSNo As Label = e.Item.FindControl("lab_SNo10")
        Dim lmem_birth As Label = e.Item.FindControl("lmem_birth")
        Dim btnEdit As LinkButton = e.Item.FindControl("btn_Edit10")
        Dim btn_Delete10 As LinkButton = e.Item.FindControl("btn_Delete10")
        Dim btn_Stop10 As LinkButton = e.Item.FindControl("btn_Stop10")
        Dim btn_UnStop10 As LinkButton = e.Item.FindControl("btn_UnStop10")

        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
                lmem_birth.Text = If(Convert.ToString(drv("mem_birth")) <> "", Common.FormatDate(Convert.ToString(drv("mem_birth"))), "")
                lmem_birth.Text = TIMS.ClearSQM(lmem_birth.Text)
                Dim sCmdArg As String = ""
                sCmdArg &= "&mem_sn=" & Convert.ToString(drv("mem_sn"))
                sCmdArg &= "&mem_idno=" & Convert.ToString(drv("mem_idno"))
                sCmdArg &= "&mem_name=" & Convert.ToString(drv("mem_name"))
                sCmdArg &= "&mem_birth=" & lmem_birth.Text 'Convert.ToString(drv("mem_birth")) 

                btnEdit.CommandArgument = String.Concat("Edit=Update", sCmdArg)

                btn_Delete10.Enabled = False
                If Me.ViewState(Cst_DelVal1) = "Y" Then
                    btn_Delete10.Enabled = True
                    btn_Delete10.Attributes.Add("onClick", "return confirm('確認要刪除該筆資料??');")
                    btn_Delete10.CommandArgument = String.Concat("Edit=Delete", sCmdArg)
                Else
                    TIMS.Tooltip(btn_Delete10, "刪除功能請系統管理者處理!!")
                End If

                btn_UnStop10.Visible = (Convert.ToString(drv("stopmem")) = "Y")
                If Convert.ToString(drv("stopmem")) = "Y" Then
                    '已停用該筆資料
                    btnEdit.Enabled = False
                    TIMS.Tooltip(btnEdit, "(帳號已停用不開放修改)!!")

                    btn_Stop10.Enabled = False
                    TIMS.Tooltip(btn_Stop10, "(已停用)!!")
                    e.Item.Cells(1).Text = "<font color=red>" & Convert.ToString(drv("mem_name")) & "(已停用)</font>"

                    btn_UnStop10.Enabled = True
                    btn_UnStop10.Attributes.Add("onClick", "return confirm('確認要啟用該筆資料??');")
                    btn_UnStop10.CommandArgument = String.Concat("Edit=UnStop", sCmdArg)
                Else
                    '停用該筆資料
                    btn_Stop10.Enabled = True
                    btn_Stop10.Attributes.Add("onClick", "return confirm('確認要停用該筆資料??');")
                    btn_Stop10.CommandArgument = String.Concat("Edit=Stop", sCmdArg)
                End If

        End Select
    End Sub

    'Private Sub LinkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
    '    'Response.Redirect("SYS_03_010_a.aspx")
    '    Dim url1 As String = "SYS_03_010_a.aspx?ID=" & Request("ID")
    '    Call TIMS.Utl_Redirect(Me, objConn, url1)
    'End Sub

    Private Sub Datagrid11_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid11.ItemCommand
        Select Case e.CommandName
            Case "btnDelete"
                Dim sql As String = ""
                Dim vs_SOCID As String = ""
                Dim vs_LeaveDate As String = ""
                Dim vs_SeqNo As String = ""
                Dim vs_LeaveID As String = ""
                vs_SOCID = TIMS.GetMyValue(e.CommandArgument, "SOCID")
                vs_LeaveDate = TIMS.GetMyValue(e.CommandArgument, "LeaveDate")
                vs_SeqNo = TIMS.GetMyValue(e.CommandArgument, "SeqNo")
                vs_LeaveID = TIMS.GetMyValue(e.CommandArgument, "LeaveID")
                vs_IDNO = TIMS.GetMyValue(e.CommandArgument, "IDNO")

                'Exit Sub
                If vs_SOCID <> "" AndAlso vs_LeaveDate <> "" AndAlso vs_SeqNo <> "" Then
                    sql = ""
                    sql &= " Delete Stud_Turnout" & vbCrLf
                    sql &= " WHERE 1=1" & vbCrLf
                    sql &= " AND SOCID='" & vs_SOCID & "'" & vbCrLf
                    sql &= " AND LeaveDate =" & TIMS.To_date(vs_LeaveDate) & vbCrLf
                    sql &= " AND SeqNo='" & vs_SeqNo & "'" & vbCrLf
                    DbAccess.ExecuteNonQuery(sql, objConn)
                End If

                Show_DataGrid11(vs_IDNO, 0)
        End Select
    End Sub

    Private Sub Datagrid11_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid11.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo11")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                Dim btnDelete As LinkButton = e.Item.FindControl("btnDelete11")
                btnDelete.Enabled = False
                If Me.ViewState(Cst_DelVal1) = "Y" Then
                    btnDelete.Enabled = True

                    btnDelete.CommandArgument = "btnDelete=1"
                    btnDelete.CommandArgument += "&SOCID=" & drv("SOCID")
                    btnDelete.CommandArgument += "&LeaveDate=" & drv("LeaveDate")
                    btnDelete.CommandArgument += "&SeqNo=" & drv("SeqNo")

                    btnDelete.CommandArgument += "&IDNO=" & drv("IDNO")
                    btnDelete.CommandArgument += "&LeaveID=" & drv("LeaveID")

                    btnDelete.Enabled = False
                    If Convert.ToString(drv("SOCID")) <> "" AndAlso Convert.ToString(drv("LeaveDate")) <> "" AndAlso Convert.ToString(drv("SeqNo")) <> "" Then
                        btnDelete.Enabled = True
                        btnDelete.Attributes("onclick") = "return confirm('您確定要清除這一筆資料?');"
                    End If
                Else
                    TIMS.Tooltip(btnDelete, "刪除功能請系統管理者處理!!")
                End If
        End Select
    End Sub

    Private Sub Datagrid12_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid12.ItemCommand
        Select Case e.CommandName
            Case "btnDelete"
                Dim sql As String = ""
                Dim vs_SOCID As String = ""
                Dim vs_LeaveDate As String = ""
                Dim vs_STOID As String = ""
                vs_SOCID = TIMS.GetMyValue(e.CommandArgument, "SOCID")
                vs_LeaveDate = TIMS.GetMyValue(e.CommandArgument, "LeaveDate")
                vs_STOID = TIMS.GetMyValue(e.CommandArgument, "STOID")
                vs_IDNO = TIMS.GetMyValue(e.CommandArgument, "IDNO")

                'Exit Sub
                If vs_SOCID <> "" AndAlso vs_LeaveDate <> "" AndAlso vs_STOID <> "" Then
                    sql = ""
                    sql &= " DELETE STUD_TURNOUT2" & vbCrLf
                    sql &= " WHERE SOCID='" & vs_SOCID & "'" & vbCrLf
                    'sql += " AND convert(varchar,LeaveDate,111) ='" & vs_LeaveDate & "'" & vbCrLf
                    sql &= " AND LeaveDate =" & TIMS.To_date(vs_LeaveDate) & vbCrLf
                    sql &= " AND STOID='" & vs_STOID & "'" & vbCrLf
                    DbAccess.ExecuteNonQuery(sql, objConn)
                End If

                Show_DataGrid12(vs_IDNO, 0)
        End Select
    End Sub

    Private Sub Datagrid12_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid12.ItemDataBound
        Const Cst_labSNo As String = "lab_SNo12"
        Const Cst_btnDelete As String = "btnDelete12"

        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl(Cst_labSNo)
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                Dim btnDelete As LinkButton = e.Item.FindControl(Cst_btnDelete)
                btnDelete.Enabled = False
                If Me.ViewState(Cst_DelVal1) = "Y" Then
                    btnDelete.Enabled = True

                    btnDelete.CommandArgument = "btnDelete=1"
                    btnDelete.CommandArgument += "&SOCID=" & drv("SOCID")
                    btnDelete.CommandArgument += "&LeaveDate=" & drv("LeaveDate")
                    btnDelete.CommandArgument += "&STOID=" & drv("STOID")

                    btnDelete.CommandArgument += "&IDNO=" & drv("IDNO")

                    btnDelete.Enabled = False
                    If Convert.ToString(drv("SOCID")) <> "" AndAlso Convert.ToString(drv("LeaveDate")) <> "" AndAlso Convert.ToString(drv("STOID")) <> "" Then
                        btnDelete.Enabled = True
                        btnDelete.Attributes("onclick") = "return confirm('您確定要清除這一筆資料?');"
                    End If
                Else
                    TIMS.Tooltip(btnDelete, "刪除功能請系統管理者處理!!")
                End If
        End Select
    End Sub

    Private Sub btn_Stud_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Stud.Click
        'GetSearchStr()'
        'edit
        Dim vsOCID As String = TIMS.Get_OCIDforSOCID(rdo_EditClass1.SelectedValue, objConn)
        Session("SearchSOCID") = rdo_EditClass1.SelectedValue '傳送的socid
        Session("RetrunUrl") = "../../SYS/03/SYS_03_010.aspx" '目前程式名稱
        'Response.Redirect("../../SD/03/SD_03_002_add.aspx?ID=" & Request("ID") & "&OCID=" & vsOCID & "&SOCID=" & rdo_EditClass1.SelectedValue)
        Dim url1 As String = "../../SD/03/SD_03_002_add.aspx?ID=" & Request("ID") & "&OCID=" & vsOCID & "&SOCID=" & rdo_EditClass1.SelectedValue
        Call TIMS.Utl_Redirect(Me, objConn, url1)
    End Sub

#Region "no use"
    'Private Sub Del_StudStudentInfo(ByVal tmpSID As String, ByVal tmpTrans As SqlTransaction)
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String

    '    Try
    '        sqlStr = "update Stud_StudentInfo set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SID= @SID "
    '        With sqlAdp
    '            .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .UpdateCommand.Parameters.Clear()
    '            .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
    '            .UpdateCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '            .UpdateCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "insert into Stud_StudentInfoDelData select * from Stud_StudentInfo where SID= @SID "
    '        With sqlAdp
    '            .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .InsertCommand.Parameters.Clear()
    '            .InsertCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '            .InsertCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "delete Stud_StudentInfo where SID= @SID "
    '        With sqlAdp
    '            .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .DeleteCommand.Parameters.Clear()
    '            .DeleteCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '            .DeleteCommand.ExecuteNonQuery()
    '        End With
    '    Catch ex As Exception
    '        tmpTrans.Rollback()
    '        Common.MessageBox(Me, ex.ToString)
    '        objConn.Close()
    '        sqlAdp.Dispose()
    '    End Try
    'End Sub

    'Private Sub Del_StudSubData(ByVal tmpSID As String, ByVal tmpTrans As SqlTransaction)
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String

    '    Try
    '        sqlStr = "update Stud_SubData set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SID= @SID "
    '        With sqlAdp
    '            .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .UpdateCommand.Parameters.Clear()
    '            .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
    '            .UpdateCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '            .UpdateCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "insert into Stud_SubDataDelData select * from Stud_SubData where SID= @SID "
    '        With sqlAdp
    '            .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .InsertCommand.Parameters.Clear()
    '            .InsertCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '            .InsertCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "delete Stud_SubData where SID= @SID "
    '        With sqlAdp
    '            .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .DeleteCommand.Parameters.Clear()
    '            .DeleteCommand.Parameters.Add("SID", SqlDbType.VarChar).Value = tmpSID
    '            .DeleteCommand.ExecuteNonQuery()
    '        End With
    '    Catch ex As Exception
    '        tmpTrans.Rollback()
    '        Common.MessageBox(Me, ex.ToString)
    '        objConn.Close()
    '        sqlAdp.Dispose()
    '    End Try
    'End Sub

    'Private Sub Del_StudEnterTemp(ByVal tmpSETID As String, ByVal tmpTrans As SqlTransaction)
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String

    '    Try
    '        sqlStr = "update Stud_EnterTemp set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID "
    '        With sqlAdp
    '            .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .UpdateCommand.Parameters.Clear()
    '            .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
    '            .UpdateCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
    '            .UpdateCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "insert into Stud_EnterTempDelData select * from Stud_EnterTemp where SETID= @SETID "
    '        With sqlAdp
    '            .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .InsertCommand.Parameters.Clear()
    '            .InsertCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
    '            .InsertCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "delete Stud_EnterTemp where SETID= @SETID "
    '        With sqlAdp
    '            .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .DeleteCommand.Parameters.Clear()
    '            .DeleteCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
    '            .DeleteCommand.ExecuteNonQuery()
    '        End With
    '    Catch ex As Exception
    '        tmpTrans.Rollback()
    '        Common.MessageBox(Me, ex.ToString)
    '        objConn.Close()
    '        sqlAdp.Dispose()
    '    End Try
    'End Sub

    'Private Sub Del_StudSelResult(ByVal tmpSETID As String, ByVal tmpTrans As SqlTransaction)
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String

    '    Try
    '        sqlStr = "update Stud_SelResult set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where SETID= @SETID "
    '        With sqlAdp
    '            .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .UpdateCommand.Parameters.Clear()
    '            .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
    '            .UpdateCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
    '            .UpdateCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "insert into Stud_SelResultDelData select * from Stud_SelResult where SETID= @SETID "
    '        With sqlAdp
    '            .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .InsertCommand.Parameters.Clear()
    '            .InsertCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
    '            .InsertCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "delete Stud_SelResult where SETID= @SETID "
    '        With sqlAdp
    '            .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .DeleteCommand.Parameters.Clear()
    '            .DeleteCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpSETID
    '            .DeleteCommand.ExecuteNonQuery()
    '        End With
    '    Catch ex As Exception
    '        tmpTrans.Rollback()
    '        Common.MessageBox(Me, ex.ToString)
    '        objConn.Close()
    '        sqlAdp.Dispose()
    '    End Try
    'End Sub

    'Private Sub Del_StudEnterTemp2(ByVal tmpeSETID As String, ByVal tmpTrans As SqlTransaction)
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String

    '    Try
    '        sqlStr = "update Stud_EnterTemp2 set ModifyAcct= @ModifyAcct,ModifyDate=getdate() where eSETID= @eSETID "
    '        With sqlAdp
    '            .UpdateCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .UpdateCommand.Parameters.Clear()
    '            .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
    '            .UpdateCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpeSETID
    '            .UpdateCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "insert into Stud_EnterTemp2DelData select * from Stud_EnterTemp2 where eSETID= @eSETID "
    '        With sqlAdp
    '            .InsertCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .InsertCommand.Parameters.Clear()
    '            .InsertCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpeSETID
    '            .InsertCommand.ExecuteNonQuery()
    '        End With
    '        sqlStr = "delete Stud_EnterTemp2 where eSETID= @eSETID "
    '        With sqlAdp
    '            .DeleteCommand = New SqlCommand(sqlStr, objConn, tmpTrans)
    '            .DeleteCommand.Parameters.Clear()
    '            .DeleteCommand.Parameters.Add("SETID", SqlDbType.VarChar).Value = tmpeSETID
    '            .DeleteCommand.ExecuteNonQuery()
    '        End With
    '    Catch ex As Exception
    '        tmpTrans.Rollback()
    '        Common.MessageBox(Me, ex.ToString)
    '        objConn.Close()
    '        sqlAdp.Dispose()
    '    End Try
    'End Sub
#End Region

    Private Sub Datagrid9_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid9.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo9")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                Dim btnDelete As LinkButton = e.Item.FindControl("btnDelete9")
                btnDelete.Enabled = False
                If Me.ViewState(Cst_DelVal1) = "Y" Then
                    btnDelete.Enabled = True

                    btnDelete.CommandArgument = "btnDelete=1"
                    'btnDelete.CommandArgument += "&SOCID=" & drv("SOCID")
                    'btnDelete.CommandArgument += "&courid=" & drv("courid")
                    'btnDelete.CommandArgument += "&idno=" & drv("idno")
                    btnDelete.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                Else
                    TIMS.Tooltip(btnDelete, "暫無提供刪除功能!!")
                End If

        End Select
    End Sub

    Private Sub Datagrid13_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid13.ItemCommand
        Const CommandName_btnDelete As String = "btnDelete"
        'sCmdArg += "&SB3ID=" & drv("SB3ID")
        'sCmdArg += "&ENTERDATE=" & drv("ENTERDATE")
        'sCmdArg += "&IDNO=" & drv("IDNO")
        'btnDelete.CommandArgument = sCmdArg
        Select Case e.CommandName
            Case CommandName_btnDelete
                Dim sql As String = ""
                Dim vs_SB3ID As String = TIMS.GetMyValue(e.CommandArgument, "SB3ID")
                Dim vs_ENTERDATE As String = TIMS.GetMyValue(e.CommandArgument, "ENTERDATE")
                vs_IDNO = TIMS.GetMyValue(e.CommandArgument, "IDNO")

                'Exit Sub
                If vs_SB3ID <> "" AndAlso vs_ENTERDATE <> "" AndAlso vs_IDNO <> "" Then
                    Call Del_STUD_SELRESULTBLI(Me, vs_SB3ID, objConn)
                End If

                Show_DataGrid13(vs_IDNO, 0)
        End Select
    End Sub

    Private Sub Datagrid13_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid13.ItemDataBound
        Const CommandName_btnDelete As String = "btnDelete"
        Const Cst_labSNo As String = "lab_SNo13"
        Const Cst_btnDelete As String = "btnDelete13"
        Const cst_lab_IDNO As String = "lab_IDNO13"
        Const cst_lab_Name As String = "lab_Name13"
        Const cst_lab_Class As String = "lab_Class13"
        Const cst_lab_SB3ID As String = "lab_SB3ID"

        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl(Cst_labSNo)
                Dim lab_IDNO As Label = e.Item.FindControl(cst_lab_IDNO)
                Dim lab_Name As Label = e.Item.FindControl(cst_lab_Name)
                Dim lab_SB3ID As Label = e.Item.FindControl(cst_lab_SB3ID)
                Dim lab_Class As Label = e.Item.FindControl(cst_lab_Class)
                Dim btnDelete As LinkButton = e.Item.FindControl(Cst_btnDelete)
                btnDelete.CommandName = CommandName_btnDelete

                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
                lab_IDNO.Text = Convert.ToString(drv("IDNO"))
                lab_Name.Text = Convert.ToString(drv("Name"))
                lab_SB3ID.Text = Convert.ToString(drv("SB3ID"))
                lab_Class.Text = Convert.ToString(drv("ClassName"))

                btnDelete.Enabled = False
                If Me.ViewState(Cst_DelVal1) = "Y" Then
                    btnDelete.Enabled = True

                    Dim sCmdArg As String = ""
                    sCmdArg = "btnDelete=1"
                    sCmdArg += "&SB3ID=" & drv("SB3ID")
                    sCmdArg += "&ENTERDATE=" & drv("ENTERDATE")
                    sCmdArg += "&IDNO=" & drv("IDNO")
                    btnDelete.CommandArgument = sCmdArg

                    btnDelete.Enabled = False
                    If Convert.ToString(drv("SB3ID")) <> "" AndAlso Convert.ToString(drv("ENTERDATE")) <> "" AndAlso Convert.ToString(drv("IDNO")) <> "" Then

                        btnDelete.Enabled = True
                        btnDelete.Attributes("onclick") = "return confirm('您確定要清除這一筆資料?');"
                    End If
                Else
                    TIMS.Tooltip(btnDelete, "刪除功能請系統管理者處理!!")
                End If
        End Select
    End Sub

    Public Shared Sub UPDATE_STUD_ENTERTYPE2(ByVal oConn As SqlConnection, ByRef htSS As Hashtable)
        'If drOC Is Nothing Then Return
        'UPDATE STUD_SELRESULT
        'UPDATE STUD_ENTERTYPE2
        Dim tOCID1 As String = TIMS.GetMyValue2(htSS, "tOCID1")
        Dim tSETID As String = TIMS.GetMyValue2(htSS, "tSETID")
        Dim tSERNUM As String = TIMS.GetMyValue2(htSS, "tSERNUM")
        Dim tENTERDATE As String = TIMS.GetMyValue2(htSS, "tENTERDATE")

        'E網報名結果
        Dim sql As String = ""
        sql &= " SELECT 'X' FROM STUD_ENTERTYPE2" & vbCrLf
        sql &= " WHERE OCID1=@OCID1 AND SETID=@SETID AND SerNum=@SerNum AND EnterDate=@EnterDate"
        Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("OCID1", Val(tOCID1))
        s_parms.Add("SETID", Val(tSETID))
        s_parms.Add("SerNum", Val(tSERNUM))
        s_parms.Add("EnterDate", TIMS.Cdate2(tENTERDATE))
        Dim dt1 As DataTable = Nothing
        dt1 = DbAccess.GetDataTable(sql, oConn, s_parms)
        If dt1.Rows.Count = 0 Then Return '查無資料(異常)
        If dt1.Rows.Count > 1 Then Return '多筆資料(異常)

        'Dim dr1 As DataRow = dt1.Rows(0)
        sql = ""
        sql &= " UPDATE STUD_ENTERTYPE2" & vbCrLf
        sql &= " SET signUpStatus=@signUpStatus" & vbCrLf
        sql &= " WHERE OCID1=@OCID1 AND SETID=@SETID AND SerNum=@SerNum AND EnterDate=@EnterDate"
        Dim u_parms As New Hashtable
        u_parms.Clear()
        u_parms.Add("signUpStatus", 5) '未錄取
        u_parms.Add("OCID1", Val(tOCID1))
        u_parms.Add("SETID", Val(tSETID))
        u_parms.Add("SerNum", Val(tSERNUM))
        u_parms.Add("EnterDate", TIMS.Cdate2(tENTERDATE))
        DbAccess.ExecuteNonQuery(sql, oConn, u_parms)
    End Sub



    Sub GET_STUDDATA3(ByRef strPK() As String, ByRef dr1 As DataRow, ByRef dr2 As DataRow, ByRef dr3 As DataRow)
        'Dim dr1 As DataRow = Nothing 'STUD_ENTERTYPE
        'Dim dr2 As DataRow = Nothing 'STUD_ENTERTEMP
        'Dim dr3 As DataRow = Nothing 'CLASS_STUDENTSOFCLASS

        Dim sqlStr As String = ""
        '變更Modify人時
        sqlStr = "select * from STUD_ENTERTYPE where SETID= @SETID and EnterDate= @EnterDate and SerNum= @SerNum"
        Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("SETID", strPK(0))
        s_parms.Add("EnterDate", CDate(strPK(1)))
        s_parms.Add("SerNum", strPK(2))
        'Dim dr1 As DataRow = DbAccess.GetOneRow(sqlStr, objConn, s_parms)
        dr1 = DbAccess.GetOneRow(sqlStr, objConn, s_parms)
        If dr1 Is Nothing Then Return

        'If dr1 Is Nothing Then
        '    Common.MessageBox(Me, "查詢資料有誤!")
        '    Exit Sub
        'End If
        sqlStr = "select * from STUD_ENTERTEMP where SETID= @SETID"
        'Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("SETID", strPK(0))
        'Dim dr2 As DataRow = DbAccess.GetOneRow(sqlStr, objConn, s_parms)
        dr2 = DbAccess.GetOneRow(sqlStr, objConn, s_parms)
        If dr2 Is Nothing Then Return

        Dim sql As String = ""
        sql &= " select cs.SID,cs.SOCID,cs.OCID" & vbCrLf
        sql &= " from STUD_STUDENTINFO ss" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.sid=ss.sid" & vbCrLf
        sql &= " where ss.IDNO=@IDNO" & vbCrLf 'E220267151'" & vbCrLf
        sql &= " and cs.OCID=@OCID" & vbCrLf  '133538'" & vbCrLf
        'Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("IDNO", dr2("IDNO"))
        s_parms.Add("OCID", dr1("OCID1"))
        'Dim dr3 As DataRow = DbAccess.GetOneRow(sql, objConn, s_parms)
        dr3 = DbAccess.GetOneRow(sql, objConn, s_parms)

    End Sub


    Protected Sub BTNSCH_T1_A_Click(sender As Object, e As EventArgs) Handles BTNSCH_T1_A.Click
        Dim v_rdo_EditClass2 As String = TIMS.GetListValue(rdo_EditClass2)
        If v_rdo_EditClass2 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        Dim sqlStr As String = ""
        Dim strPK() As String = Split(v_rdo_EditClass2, ";")
        Dim dr1 As DataRow = Nothing 'STUD_ENTERTYPE
        Dim dr2 As DataRow = Nothing 'STUD_ENTERTEMP
        Dim dr3 As DataRow = Nothing 'CLASS_STUDENTSOFCLASS
        Call GET_STUDDATA3(strPK, dr1, dr2, dr3)
        If dr1 Is Nothing Then
            Common.MessageBox(Me, "查詢資料有誤!")
            Exit Sub
        End If
        If dr2 Is Nothing Then
            Common.MessageBox(Me, "查詢資料有誤!!")
            Exit Sub
        End If
        If dr3 IsNot Nothing Then
            Common.MessageBox(Me, "已為學員，不提供查詢狀況!!")
            Exit Sub
        End If

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        Dim aOCID1 As String = TIMS.ClearSQM(dr1("OCID1"))
        Dim aIDNO1 As String = TIMS.ClearSQM(dr2("IDNO"))
        Dim ERRMSG As String = ""
        '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
        Dim xStudInfo As String = String.Format("&IDNO={0}&OCID1={1}&STEST=Y", aIDNO1, aOCID1)
        Call TIMS.ChkStudDouble(timsSer1, ERRMSG, "", xStudInfo)
        If ERRMSG = "" Then ERRMSG = "(OK)該班未發生重複"
        Common.MessageBox(Me, ERRMSG)
        Return
    End Sub

    Private Sub DataGrid14_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid14.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo14")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
        End Select
    End Sub

    Private Sub DataGrid15_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid15.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo15")
                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)
        End Select
    End Sub

    '(批次)儲存課程變更
    Protected Sub btn_EditSaveClass1B_Click(sender As Object, e As EventArgs) Handles btn_EditSaveClass1B.Click
        If rdo_EditClass1.SelectedIndex = -1 AndAlso rdo_EditClass1.SelectedValue = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。") 'SOCID
            Exit Sub
        End If
        If txt_EditClassSID1.Text = "" Then
            Common.MessageBox(Me, "SID不能沒填。")
            Exit Sub
        End If

        iRst = 0
        For Each objItem As ListItem In rdo_EditClass1.Items
            Dim hParms As New Hashtable
            hParms.Add("SID", txt_EditClassSID1.Text)
            hParms.Add("SETID", txt_EditClassSETID1.Text)
            hParms.Add("SOCID", objItem.Value)
            Call UPDATE_CLASS_STUDENTSOFCLASS(objConn, hParms)
            iRst += 1
        Next

        'Common.MessageBox(Me, "修改完成。")
        lab_msg_stud.Text = String.Concat(Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), "，更新筆數：", iRst)
        Show_RdoEditClass1(Get_ClassStudentsOfClass("SID", lab_EditSID1.Text))
        tb_EditClass1.Visible = False
    End Sub

    '(批次)儲存課程變更
    Protected Sub btn_EditSaveClass2B_Click(sender As Object, e As EventArgs) Handles btn_EditSaveClass2B.Click
        Dim v_rdo_EditClass2 As String = TIMS.GetListValue(rdo_EditClass2)
        If rdo_EditClass2.SelectedIndex = -1 OrElse v_rdo_EditClass2 = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        'Dim strPK() As String = Split(v_rdo_EditClass2, ";")
        If txt_EditClassSETID2.Text = "" Then
            Common.MessageBox(Me, "SETID不能沒填。")
            Exit Sub
        End If

        iRst = 0
        For Each objItem As ListItem In rdo_EditClass2.Items
            Dim strPK() As String = Split(objItem.Value, ";")
            'Dim flagDblRows As Boolean = False '更新資料有重複 False:沒有
            '新的與舊的資料不同，做判斷
            If txt_EditClassSETID2.Text <> Convert.ToString(strPK(0)) Then
                Dim dt As New DataTable
                Dim sqlStr As String = "SELECT * FROM STUD_SELRESULT WHERE SETID=@newSETID and EnterDate=@EnterDate and SerNum=@SerNum"
                Dim sCmd As New SqlCommand(sqlStr, objConn)
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("newSETID", SqlDbType.Int).Value = txt_EditClassSETID2.Text
                    '.Parameters.Add("SETID", SqlDbType.Int).Value = Convert.ToInt32(strPK(0))
                    .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = Convert.ToDateTime(strPK(1))
                    .Parameters.Add("SerNum", SqlDbType.Int).Value = Convert.ToInt32(strPK(2))
                    dt.Load(.ExecuteReader())
                End With
                '更新資料有重複 False:沒有
                Dim flagDblRows As Boolean = (dt.Rows.Count > 0) '更新資料有重複 False:沒有 True:有
                If flagDblRows Then
                    Common.MessageBox(Me, "修改資料異常，請先確認資料正確性!!") '資料有重複

                    Dim pParms As New Hashtable
                    pParms.Add("SETID_1", txt_EditClassSETID2.Text)
                    pParms.Add("SETID_2", strPK(0))
                    pParms.Add("ENTERDATE", strPK(1))
                    pParms.Add("SERNUM", strPK(2))
                    Call SHOW_ERROR_HANDLING_TIPS(pParms)
                    'Show_RdoEditClass2(Get_StudEnterType("SETID", lab_EditSETID2.Text))
                    'tb_EditClass2.Visible = False
                    Return
                End If
            End If

            Dim UParms As New Hashtable
            UParms.Add("SETID_1", txt_EditClassSETID2.Text)
            UParms.Add("SETID_2", strPK(0))
            UParms.Add("ENTERDATE", strPK(1))
            UParms.Add("SERNUM", strPK(2))
            UParms.Add("eSETID", txt_EditClasseSETID2.Text)
            PGErrMsg1 = UPDATE_STUD_ENTERTYPE(objConn, UParms)
            If PGErrMsg1 <> "" Then
                Common.MessageBox(Me, PGErrMsg1)
                Return
            End If

            iRst += 1
        Next

        'Common.MessageBox(Me, "修改完成。")
        'lab_SelResult_msg.Text = Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
        lab_SelResult_msg.Text = String.Concat(Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), "，更新筆數：", iRst)
        Call Show_RdoEditClass2(Get_StudEnterType("SETID", lab_EditSETID2.Text))
        tb_EditClass2.Visible = False
    End Sub

    '(批次)儲存課程變更
    Protected Sub btn_EditSaveClass3B_Click(sender As Object, e As EventArgs) Handles btn_EditSaveClass3B.Click
        If rdo_EditClass3.SelectedIndex = -1 OrElse rdo_EditClass3.SelectedValue = "" Then
            Common.MessageBox(Me, "請選取要處理的資料。")
            Exit Sub
        End If
        If txt_EditClasseSETID3.Text = "" Then
            Common.MessageBox(Me, "eSETID不能沒填。")
            Exit Sub
        End If

        iRst = 0
        For Each objItem As ListItem In rdo_EditClass3.Items
            Dim pParms As New Hashtable
            pParms.Add("SETID", txt_EditClassSETID3.Text)
            pParms.Add("eSETID", txt_EditClasseSETID3.Text)
            pParms.Add("eSERNUM", objItem.Value) 'TIMS.GetListValue(rdo_EditClass3)
            Call UPDATE_STUD_ENTERTYPE2B(objConn, pParms)
            iRst += 1
        Next

        'Common.MessageBox(Me, "修改完成。")
        Lab_ENTERTYPE2_mag.Text = String.Concat(Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), "，更新筆數：", iRst)
        Show_RdoEditClass3(Get_StudEnterType2("eSETID", lab_EditeSETID3.Text))
        tb_EditClass3.Visible = False
    End Sub

End Class
