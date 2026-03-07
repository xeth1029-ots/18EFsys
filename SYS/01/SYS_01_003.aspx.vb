Partial Class SYS_01_003
    Inherits AuthBasePage

    Dim selectAllPlans As String
    Dim selectAllAccounts As String
    Dim auditAllPlan As DropDownList
    Dim auditAllAccount As DropDownList
    'Dim FunDr As DataRow

#Region "Function"
    '依條件取得帳號申請資料
    Private Function Get_AccountApplyList() As DataTable
        Dim dt As New DataTable

        Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        Dim sqlStr As String = ""
        sqlStr = ""
        sqlStr += "select a.AcctID,a.Account,a.Name,a.IDNO,b.Name as RoleName,c.OrgName " & vbCrLf
        sqlStr += ",dbo.NVL(a.AuditStatus,'X') as AuditStatus,dbo.NVL(d.CntAcctPID,0) as CntAcctPID " & vbCrLf
        sqlStr += "FROM AUTH_ACCOUNTTEMP a " & vbCrLf
        sqlStr += "join ID_Role b on a.RoleID=b.RoleID " & vbCrLf
        sqlStr += "join Org_OrgInfo c on c.OrgID=a.OrgID " & vbCrLf

        sqlStr += " left join (" & vbCrLf
        sqlStr += " 	select Account,count(AcctPID) CntAcctPID,AcctID " & vbCrLf
        sqlStr += " 	from Auth_AccRWPlanTemp " & vbCrLf
        sqlStr += " 	where ActMode is null " & vbCrLf
        sqlStr += " 	and AuditStatus is null " & vbCrLf
        If sm.UserInfo.RoleID <> "0" Then
            sqlStr += " 	and DistID= @DistID " & vbCrLf
        End If
        sqlStr += " 	group by Account,AcctID" & vbCrLf
        sqlStr += " ) d on d.Account=a.Account and d.AcctID=a.AcctID " & vbCrLf
        sqlStr += " where 1=1 " & vbCrLf

        If sm.UserInfo.RoleID <> "0" Then
            sqlStr += " and a.DistID= @DistID " & vbCrLf
            da.SelectCommand.Parameters.Add("DistID", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.DistID)
        End If
        If Me.ViewState("nameid") <> "" Then
            sqlStr += "and a.Account like @Account " & vbCrLf
            da.SelectCommand.Parameters.Add("Account", SqlDbType.NVarChar).Value = Convert.ToString(Me.ViewState("nameid")).Trim
        End If
        If Me.ViewState("namefield") <> "" Then
            sqlStr += "and a.Name like @Name " & vbCrLf
            da.SelectCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = "%" & Convert.ToString(Me.ViewState("namefield")).Trim & "%"
        End If
        If Me.ViewState("OrgID") <> "" Then
            sqlStr += "and a.OrgID= @OrgID " & vbCrLf
            da.SelectCommand.Parameters.Add("OrgID", SqlDbType.Int).Value = Convert.ToInt32(Me.ViewState("OrgID"))
        End If
        If Me.ViewState("Resultsrh") <> "" Then
            sqlStr += "and dbo.NVL(a.AuditStatus,'X')= @AuditStatus " & vbCrLf
            da.SelectCommand.Parameters.Add("AuditStatus", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("Resultsrh"))
        End If
        sqlStr += " order by a.ApplyDate " & vbCrLf
        TIMS.Fill(sqlStr, da, dt)
        Return dt
    End Function

    Private Function Get_PlanApplyList(ByVal tmpAccount As String) As DataTable
        Dim rst As DataTable
        Dim sqlStr As String = "SELECT * FROM AUTH_ACCRWPLANTEMP where Account= @Account and AuditStatus is null and ActMode is null "
        Dim cmd As New SqlCommand(sqlStr, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("Account", SqlDbType.VarChar).Value = tmpAccount
            rst = New DataTable
            rst.Load(.ExecuteReader())
        End With
        Return rst
    End Function

    '依條件取得計劃申請資料
    Function Get_PlanApplyList() As DataTable
        'Dim dt As New DataTable
        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        Dim sqlStr As String = ""
        sqlStr = "" & vbCrLf
        sqlStr += " select a.AcctPID" & vbCrLf
        sqlStr += " ,a.Account" & vbCrLf
        sqlStr += " ,dbo.NVL(b.Name,c.Name) Name" & vbCrLf
        sqlStr += " ,(case when b.Name is null then e.OrgName else d.OrgName end) OrgName" & vbCrLf
        sqlStr += " ,f.Years+g.Name+h.PlanName+f.Seq+(case when f.TPlanID in (17,22) then '_'+dbo.NVL(CONVERT(varchar, l.OrgName),'機構名稱異常') end) PlanName" & vbCrLf
        sqlStr += " ,dbo.NVL(a.AuditStatus,'X') AuditStatus" & vbCrLf
        sqlStr += " ,dbo.NVL(i.AcctID,0) AcctID" & vbCrLf
        sqlStr += " ,a.PlanID,a.DistID" & vbCrLf
        sqlStr += " ,case when j.AcctPID is null then 'N' else 'Y' end Shared" & vbCrLf
        sqlStr += " from Auth_AccRWPlanTemp a " & vbCrLf
        sqlStr += " join ID_Plan f on f.PlanID=a.PlanID " & vbCrLf
        sqlStr += " join ID_District g on g.DistID=a.DistID " & vbCrLf
        sqlStr += " join Key_Plan h on h.TPlanID=f.TPlanID " & vbCrLf
        sqlStr += " left join Auth_Account b on a.Account=b.Account " & vbCrLf
        '帳號已經審核通過者 isnull(c.AuditStatus,'X')<>'N'
        sqlStr += " left join Auth_AccountTemp c on c.Account=a.Account and dbo.NVL(c.AuditStatus,'X')<>'N' " & vbCrLf
        sqlStr += " left join Org_OrgInfo d on d.OrgID=b.OrgID " & vbCrLf
        sqlStr += " left join Org_OrgInfo e on e.OrgID=c.OrgID " & vbCrLf
        sqlStr += " left join (" & vbCrLf
        sqlStr += " 	select Account,AcctID " & vbCrLf
        sqlStr += " 	from Auth_AccountTemp " & vbCrLf
        '判斷是否有帳號申請待審核，有的話不能進行計劃審核
        sqlStr += " 	where AuditStatus is null  " & vbCrLf
        If sm.UserInfo.RoleID <> "0" Then
            sqlStr += " and DistID= @DistID " & vbCrLf
        End If
        sqlStr += " ) i on i.Account=a.Account " & vbCrLf
        sqlStr += " left join Auth_RelshipTemp j on j.AcctPID=a.AcctPID " & vbCrLf
        sqlStr += " left join Auth_Relship k on (case when j.Relship is null then k.RID else k.Relship end)=(case when j.Relship is null then a.RID else j.Relship end) " & vbCrLf
        sqlStr += " left join view_RIDName l on l.RID=(case when j.Relship is not null and k.OrgLevel>1 and dbo.NVL(a.AuditStatus,'')<>'Y' then k.RID else replace(replace( dbo.SUBSTR(k.Relship,5,Len(k.Relship)-4),k.RID,''),'/','') end) " & vbCrLf
        sqlStr += " where a.ActMode is null  " & vbCrLf
        'sqlStr += " and a.Account like Account" & vbCrLf
        'sqlStr += " and isnull(a.AuditStatus,'X')= @AuditStatus " & vbCrLf
        'sqlStr += " order by a.Account,a.ApplyDate" & vbCrLf

        Dim parms As New Hashtable
        parms.Clear()
        If sm.UserInfo.RoleID <> "0" Then
            sqlStr += "and a.DistID= @DistID " & vbCrLf
            parms.Add("DistID", sm.UserInfo.DistID)
            'da.SelectCommand.Parameters.Add("DistID", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.DistID)
        End If

        If Me.ViewState("nameid") <> "" Then
            sqlStr += "and a.Account like @Account" & vbCrLf
            parms.Add("Account", Convert.ToString(Me.ViewState("nameid")) & "%")
        End If

        If Me.ViewState("namefield") <> "" Then
            sqlStr += "and (b.Name like @Name or c.Name like @Name) " & vbCrLf
            parms.Add("Name", "%" & Convert.ToString(Me.ViewState("namefield")) & "%")
            'da.SelectCommand.Parameters.Add("Name", SqlDbType.NVarChar).Value = "%" & Convert.ToString(Me.ViewState("namefield")) & "%"
        End If
        If Me.ViewState("OrgID") <> "" Then
            sqlStr += "and (b.OrgID= @OrgID or c.OrgID= @OrgID) " & vbCrLf
            parms.Add("OrgID", Convert.ToInt32(Me.ViewState("OrgID")))
            'da.SelectCommand.Parameters.Add("OrgID", SqlDbType.Int).Value = Convert.ToInt32(Me.ViewState("OrgID"))
        End If
        If Me.ViewState("Resultsrh") <> "" Then
            sqlStr += "and dbo.NVL(a.AuditStatus,'X')= @AuditStatus " & vbCrLf
            parms.Add("AuditStatus", Me.ViewState("Resultsrh"))
            'da.SelectCommand.Parameters.Add("AuditStatus", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("Resultsrh"))
        End If
        If sm.UserInfo.RoleID <> "0" And sm.UserInfo.RoleID <> "1" Then
            sqlStr += "and a.PlanID= @PlanID " & vbCrLf
            parms.Add("PlanID", Convert.ToInt32(sm.UserInfo.PlanID))
            'da.SelectCommand.Parameters.Add("PlanID", SqlDbType.Int).Value = Convert.ToInt32(sm.UserInfo.PlanID)
        End If
        sqlStr += "ORDER BY a.Account,a.ApplyDate " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlStr, objconn, parms)
        'TIMS.Fill(sqlStr, da, dt)

        Return dt
    End Function

    '更新帳號審核狀態
    Sub Update_AccountApplyList(ByVal acctID As Integer, ByVal auditStatus As String)
        Dim sqlStr As String = "UPDATE AUTH_ACCOUNTTEMP SET AuditStatus= @AuditStatus,AuditAcct= @AuditAcct,AuditDate=getdate() WHERE AcctID=@AcctID"
        Call TIMS.OpenDbConn(objconn)
        Dim objSqlCmd As New SqlCommand(sqlStr, objconn)
        With objSqlCmd
            .Parameters.Clear()
            .Parameters.Add("AuditStatus", SqlDbType.Char).Value = auditStatus
            .Parameters.Add("AuditAcct", SqlDbType.NVarChar).Value = Convert.ToString(sm.UserInfo.UserID).Trim(" ")
            .Parameters.Add("AcctID", SqlDbType.Int).Value = acctID
            .ExecuteNonQuery()
        End With
    End Sub

    '將帳號資料從TEMP轉存到正式
    Sub Save_Account(ByVal acctID As Integer)
        Dim sqlStr As String = ""
        sqlStr &= " INSERT INTO AUTH_ACCOUNT(Account,RoleID,LID,Name,Phone,Email,OrgID,IDNO,Serialno,ModifyAcct,ModifyDate) " & vbCrLf
        sqlStr &= " SELECT Account,RoleID,LID,Name,Phone,Email,OrgID,IDNO,Serialno,@ModifyAcct,GETDATE() FROM AUTH_ACCOUNTTEMP WHERE AcctID= @AcctID"
        Call TIMS.OpenDbConn(objconn)
        Dim objSqlCmd As New SqlCommand(sqlStr, objconn)
        With objSqlCmd
            .Parameters.Clear()
            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID 'Convert.ToString().Trim(" ")
            .Parameters.Add("AcctID", SqlDbType.Int).Value = acctID
            .ExecuteNonQuery()
        End With
    End Sub

    Sub Update_AuthAccount(ByVal tmpAcct As String, ByVal tmpUsed As String)
        Dim sqlStr As String = "UPDATE AUTH_ACCOUNT SET ISUSED= @ISUSED WHERE ACCOUNT= @ACCOUNT "
        Call TIMS.OpenDbConn(objconn)
        Dim cmd As New SqlCommand(sqlStr, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("ISUSED", SqlDbType.Char).Value = tmpUsed
            .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = tmpAcct
            .ExecuteNonQuery()
        End With
    End Sub

    '檢查Auth_Account是否存在有Auth_AccountTemp審核通過的資料
    Private Function Chk_Account(ByVal tmpAcctID As Integer) As Boolean
        Dim dt As New DataTable
        Dim sqlStr As String = "SELECT 1 FROM AUTH_ACCOUNT WHERE ACCOUNT IN (select Account FROM AUTH_ACCOUNTTEMP WHERE AcctID=@AcctID)"
        Call TIMS.OpenDbConn(objconn)
        Dim cmd As New SqlCommand(sqlStr, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("AcctID", SqlDbType.Int).Value = tmpAcctID
            dt.Load(.ExecuteReader())
        End With
        Return TIMS.dtHaveDATA(dt)
    End Function

    '更新計劃審核狀態
    Sub Update_PlanApplyList(ByVal acctPID As Integer, ByVal auditStatus As String)
        Dim sqlStr As String = "UPDATE AUTH_ACCRWPLANTEMP SET AuditStatus= @AuditStatus,AuditAcct= @AuditAcct,AuditDate=GETDATE() WHERE AcctPID= @AcctPID"
        Call TIMS.OpenDbConn(objconn)
        Dim cmd As New SqlCommand(sqlStr, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("AuditStatus", SqlDbType.Char).Value = auditStatus
            .Parameters.Add("AuditAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID).Trim(" ")
            .Parameters.Add("AcctPID", SqlDbType.Int).Value = acctPID
            .ExecuteNonQuery()
        End With
    End Sub

    'ACTMODE更新 'C:如果有申請計畫，自動將計劃失效 Y:更新TEMP的審核狀態為通過
    Sub Update_PlanApplyList(ByVal tmpAccount As String, ByVal tmpActStatus As String, Optional ByVal tmpNote As String = "")
        Dim sqlStr As String = "UPDATE AUTH_ACCRWPLANTEMP SET ActMode= @ActMode,AuditNote= @AuditNote WHERE Account=@Account and AuditStatus is null and ActMode is null"
        Call TIMS.OpenDbConn(objconn)
        Dim sqlCmd As New SqlCommand(sqlStr, objconn)
        With sqlCmd
            .Parameters.Clear()
            .Parameters.Add("ActMode", SqlDbType.Char).Value = tmpActStatus
            .Parameters.Add("AuditNote", SqlDbType.NVarChar).Value = If(tmpNote <> "", tmpNote, Convert.DBNull)
            .Parameters.Add("Account", SqlDbType.VarChar).Value = tmpAccount
            .ExecuteNonQuery()
        End With
    End Sub

    '將計劃資料從TEMP轉存到正式
    Sub Save_Plan(ByVal acctPID As Integer, ByVal tmpShared As String)
        'Dim objAdp As SqlCommand
        'Dim sqlStr As String
        Dim newRID As String = ""
        Dim RSID As Integer = 0

        If tmpShared = "Y" Then
            newRID = Get_NewRID(acctPID)    '取得新的RID
            If newRID <> "" Then    '新的共用
                Me.ViewState("newRID") = newRID
                Call Save_AccRWPlanTemp(acctPID) '更新Auth_AccRWPlanTemp的RID
                Call Save_RelshipTemp(acctPID, 2) '更新Auth_RelshipTemp的RID
            End If
        End If
        Call Save_AccRWPlan(acctPID) '將資料從 Auth_AccRWPlanTemp 轉到Auth_AccRWPlan
        If tmpShared = "Y" Then
            Dim dt_RelshipTemp As DataTable = Nothing
            RSID = Save_Relship(acctPID)    '將資料從Auth_RelshipTemp轉到Auth_Relship
            If RSID <> 0 Then Save_OrgPlanInfo(acctPID, RSID) '將資料從Org_OrgPlanInfoTemp轉到Org_OrgPlanInfo

            '檢查是否有相同的共用資料待審，有的情況下讓這些待審的資料失效
            dt_RelshipTemp = Chk_RelshipTemp(acctPID)
            If Not dt_RelshipTemp Is Nothing Then
                For i As Integer = 0 To dt_RelshipTemp.Rows.Count - 1
                    Dim othAcctPID As Integer = Convert.ToInt32(dt_RelshipTemp.Rows(i).Item("AcctPID"))
                    Save_AccRWPlanTemp(othAcctPID)  '更新這些待審資料在Auth_AccRWPlanTemp的RID
                    Save_RelshipTemp(othAcctPID, 3) '失效待審的Auth_RelshipTemp
                    Save_OrgPlanInfoTemp(othAcctPID) '失效待審的Org_OrgPlanInfoTemp
                Next
            End If
        End If
    End Sub

    '檢查Auth_Relship是否存在有Auth_RelshipTemp審核通過的資料
    Private Function Chk_Relship(ByVal OrgID As Integer, ByVal planID As Integer, ByVal DistID As String) As Boolean
        Dim dt As New DataTable
        Dim sqlStr As String = "SELECT RID FROM AUTH_RELSHIP WHERE PlanID= @PlanID and OrgID= @OrgID and DistID= @DistID"
        Call TIMS.OpenDbConn(objconn)
        Dim sqlCmd As New SqlCommand(sqlStr, objconn)
        With sqlCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.Int).Value = planID
            .Parameters.Add("OrgID", SqlDbType.Int).Value = OrgID
            .Parameters.Add("DistID", SqlDbType.VarChar).Value = DistID
            dt.Load(.ExecuteReader())
        End With
        Return TIMS.dtHaveDATA(dt)
    End Function

    '儲存Auth_AccRWPlanTemp(tmpAct=0--刪除,tmpAct=1--新增,tmpAct=2--更新RID)
    Private Sub Save_AccRWPlanTemp(ByVal tmpAcctPID As Integer, Optional ByVal tmpAct As Integer = 2)
        Select Case tmpAct
            Case 0
            Case 1
            Case 2  '更新RID
                Dim sqlStr As String
                sqlStr = "update Auth_AccRWPlanTemp set RID= @RID where AcctPID= @AcctPID "
                Call TIMS.OpenDbConn(objconn)
                Dim sqlCmd As New SqlCommand(sqlStr, objconn)
                With sqlCmd
                    .Parameters.Clear()
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("newRID"))
                    .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                    .ExecuteNonQuery()
                End With
        End Select
    End Sub

    '儲存Auth_ReshipTemp(tmpAct=0--刪除,tmpAct=1--新增,tmpAct=2--更新RID)
    Sub Save_RelshipTemp(ByVal tmpAcctPID As Integer, Optional ByVal tmpAct As Integer = 2)
        'Dim sqlCmd As SqlCommand
        Dim sqlStr As String = ""
        Select Case tmpAct
            Case 0
            Case 1
            Case 2  '更新RID
                sqlStr = "update Auth_RelshipTemp set RID= @RID,Relship=Relship+@RID+'/' where AcctPID= @AcctPID "
            Case 3  '更新ActMode為失效
                sqlStr = "update Auth_RelshipTemp set ActMode='C' where AcctPID= @AcctPID"
        End Select
        Call TIMS.OpenDbConn(objconn)
        Dim sqlCmd As New SqlCommand(sqlStr, objconn)
        With sqlCmd
            Select Case tmpAct
                Case 0
                Case 1
                Case 2  '更新RID
                    .Parameters.Clear()
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("newRID"))
                    .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                    .ExecuteNonQuery()
                Case 3  '更新ActMode為失效
                    .Parameters.Clear()
                    .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                    .ExecuteNonQuery()
            End Select
        End With
    End Sub

    '儲存Org_OrgPlanInfoTemp(tmpAct=0--刪除,tmpAct=1--新增,tmpAct=2--更新ActMode為失效)
    Sub Save_OrgPlanInfoTemp(ByVal tmpAcctPID As Integer, Optional ByVal tmpAct As Integer = 2)
        'Dim sqlCmd As SqlCommand
        Dim sqlStr As String = ""
        Select Case tmpAct
            Case 0
            Case 1
            Case 2  '更新ActMode為失效
                sqlStr = "update Org_OrgPlanInfoTemp set ActMode='C' where AcctPID= @AcctPID"
        End Select
        Dim sqlCmd As New SqlCommand(sqlStr, objconn)
        With sqlCmd
            Select Case tmpAct
                Case 0
                Case 1
                Case 2  '更新ActMode為失效
                    .Parameters.Clear()
                    .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                    If objconn.State = ConnectionState.Closed Then objconn.Open()
                    .ExecuteNonQuery()
            End Select
        End With
    End Sub

    '儲存Auth_AccRWPlan(tmpAct=0--刪除,tmpAct=1--新增)
    Sub Save_AccRWPlan(ByVal tmpAcctPID As Integer, Optional ByVal tmpAct As Integer = 1)
        'Dim sqlCmd As SqlCommand
        Dim sqlStr As String = ""
        Select Case tmpAct
            Case 0
            Case 1
                sqlStr &= " INSERT INTO AUTH_ACCRWPLAN(Account,PlanID,RID,CreateByAcc,ModifyAcct,ModifyDate) "
                sqlStr &= " select Account,PlanID,RID,@CreatebyAcc CreateByAcc,@ModifyAcct ModifyAcct,getdate() ModifyDate from Auth_AccRWPlanTemp where AcctPID= @AcctPID "
        End Select
        Select Case tmpAct
            Case 0
            Case 1
                Dim createByAcc As String = ""
                If Chk_AccRWPlan(tmpAcctPID, createByAcc) = False Then
                    Dim sqlCmd As New SqlCommand(sqlStr, objconn)
                    With sqlCmd
                        .Parameters.Clear()
                        .Parameters.Add("CreateByAcc", SqlDbType.NVarChar).Value = createByAcc
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID).Trim
                        .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                        If objconn.State = ConnectionState.Closed Then objconn.Open()
                        .ExecuteNonQuery()
                    End With

                    '查詢主要權限是否存在 CreateByAcc='Y'
                    Dim objstr As String = $" SELECT * FROM AUTH_ACCRWPLANTEMP WHERE AcctPID={tmpAcctPID}"
                    Dim dr As DataRow = DbAccess.GetOneRow(objstr, objconn)
                    If Not dr Is Nothing Then
                        Dim objrow As DataRow = Nothing
                        Dim objadapter As SqlDataAdapter = Nothing
                        Dim objtable As DataTable = Nothing
                        objstr = "Select count(1) cnt from Auth_AccRWPlan where CreateByAcc='Y' and Account='" & dr("Account") & "' and RID ='" & dr("RID") & "'"
                        If DbAccess.ExecuteScalar(objstr) = 0 Then
                            '增加1筆主要權限 CreateByAcc='Y'
                            objstr = "Select * from Auth_AccRWPlan where Account='" & dr("Account") & "' and RID='" & dr("RID") & "' AND ROWNUM<=1"
                            objtable = DbAccess.GetDataTable(objstr, objadapter, objconn)
                            If objtable.Rows.Count > 0 Then
                                objrow = objtable.Rows(0)
                                objrow("CreateByAcc") = "Y"
                                DbAccess.UpdateDataTable(objtable, objadapter)
                            End If
                        End If
                    End If

                End If
        End Select
    End Sub

    '檢查Auth_AccRWPlan是否存在有Auth_AccRWPlanTemp審核通過的資料
    Private Function Chk_AccRWPlan(ByVal tmpAcctPID As Integer, Optional ByRef tmpCBA As String = "N") As Boolean
        Dim objAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim sqlStr As String = ""
        Dim rst As Boolean = False

        '取得Auth_AccRWPlanTemp的帳號、計畫
        sqlStr = ""
        sqlStr &= " select a.Account,a.PlanID,dbo.NVL(b.OrgLevel,(select OrgLevel from Auth_Relship where RID=a.RID)) as OrgLevel,a.RID "
        sqlStr += " from Auth_AccRWPlanTemp a left join Auth_RelshipTemp b on b.AcctPID=a.AcctPID "
        sqlStr += " where a.AuditStatus is null and a.ActMode is null and a.AcctPID= @AcctPID"
        Try
            With objAdp
                .SelectCommand = New SqlCommand(sqlStr, objconn)
                .SelectCommand.Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                .Fill(objDS, "Temp")
            End With
            If objDS.Tables("Temp").Rows.Count > 0 Then
                '以帳號、計畫檢查Auth_AccRWPlan是否有相同資料
                sqlStr = "select RID from Auth_AccRWPlan where Account= @Account and PlanID= @PlanID "
                With objAdp
                    .SelectCommand = New SqlCommand(sqlStr, objconn)
                    .SelectCommand.Parameters.Add("Account", SqlDbType.VarChar).Value = Convert.ToString(objDS.Tables("Temp").Rows(0).Item("Account"))
                    .SelectCommand.Parameters.Add("PlanID", SqlDbType.Int).Value = Convert.ToInt32(objDS.Tables("Temp").Rows(0).Item("PlanID"))
                    If objDS.Tables("Temp").Rows(0).Item("OrgLevel") = 3 Then
                        .SelectCommand.CommandText += "and RID like @RID "
                        .SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = Left(Convert.ToString(objDS.Tables("Temp").Rows(0).Item("RID")), Len(Convert.ToString(objDS.Tables("Temp").Rows(0).Item("RID"))) - 3) & "%"
                    End If
                    .Fill(objDS, "Auth")
                End With
                If objDS.Tables("Auth").Rows.Count > 0 Then rst = True Else rst = False
                '以帳號檢查是否初次新增計畫，並返回檢驗值
                sqlStr = "select RID from Auth_AccRWPlan where Account= @Account "
                With objAdp
                    .SelectCommand = New SqlCommand(sqlStr, objconn)
                    .SelectCommand.Parameters.Add("Account", SqlDbType.VarChar).Value = Convert.ToString(objDS.Tables("Temp").Rows(0).Item("Account"))
                    .Fill(objDS, "CBA")
                End With
                If objDS.Tables("CBA").Rows.Count > 0 Then tmpCBA = "N" Else tmpCBA = "Y"
            Else
                rst = True
            End If
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Throw ex
        End Try
        Return rst
    End Function

    '儲存Auth_Relship(tmpAct=0--刪除,tmpAct=1--新增)
    Private Function Save_Relship(ByVal tmpAcctPID As Integer, Optional ByVal tmpAct As Integer = 1) As Integer
        Dim iRSID As Integer = 0
        'Dim sqlCmd As SqlCommand
        Dim sqlStr As String = ""
        Select Case tmpAct
            Case 0
            Case 1
                sqlStr = ""
                sqlStr += " insert into Auth_Relship(RSID,PlanID,RID,OrgID,Relship,OrgLevel,DistID,ModifyAcct,ModifyDate) " & vbCrLf
                sqlStr += " select @RSID,PlanID,RID,OrgID,Relship,OrgLevel,DistID,@ModifyAcct as ModifyAcct,getdate() as ModifyDate from Auth_RelshipTemp where AcctPID= @AcctPID "
        End Select
        Dim sqlCmd As New SqlCommand(sqlStr, objconn)
        With sqlCmd
            Select Case tmpAct
                Case 0
                Case 1
                    If Chk_Relship(tmpAcctPID) = False Then
                        iRSID = DbAccess.GetNewId(objconn, "AUTH_RELSHIP_RSID_SEQ,AUTH_RELSHIP,RSID")
                        .Parameters.Clear()
                        .Parameters.Add("RSID", SqlDbType.Int).Value = iRSID
                        .Parameters.Add("AcctPID", SqlDbType.NVarChar).Value = tmpAcctPID
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID).Trim
                        If objconn.State = ConnectionState.Closed Then objconn.Open()
                        .ExecuteNonQuery()
                        '.ExecuteScalar()
                        'sqlStr = "select :@identity as RSID" 'AUTH_RELSHIP_RSID_SEQ
                    End If
            End Select
        End With
        Return iRSID
    End Function

    '檢查Auth_Relship中是否存在有Auth_RelshipTemp審核通過的資料
    Private Function Chk_Relship(ByVal tmpAcctPID As Integer) As Boolean
        Dim objAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim sqlStr As String = ""

        Dim rst As Boolean = True

        sqlStr = "select * from Auth_RelshipTemp where ActMode is null and AcctPID= @AcctPID"
        objAdp.SelectCommand = New SqlCommand(sqlStr, objconn)
        With objAdp.SelectCommand
            .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
        End With
        objAdp.Fill(objDS, "Temp")

        If objDS.Tables("Temp").Rows.Count > 0 Then
            sqlStr = "select RSID from Auth_Relship where PlanID= @PlanID and OrgID= @OrgID and DistID= @DistID and Relship like @Relship and OrgLevel= @OrgLevel "
            objAdp.SelectCommand = New SqlCommand(sqlStr, objconn)
            With objAdp.SelectCommand
                .Parameters.Add("PlanID", SqlDbType.Int).Value = Convert.ToInt32(objDS.Tables("Temp").Rows(0).Item("PlanID"))
                .Parameters.Add("OrgID", SqlDbType.Int).Value = Convert.ToInt32(objDS.Tables("Temp").Rows(0).Item("OrgID"))
                .Parameters.Add("DistID", SqlDbType.NVarChar).Value = Convert.ToString(objDS.Tables("Temp").Rows(0).Item("DistID"))
                .Parameters.Add("Relship", SqlDbType.VarChar).Value = Convert.ToString(objDS.Tables("Temp").Rows(0).Item("Relship")) & "%"
                .Parameters.Add("OrgLevel", SqlDbType.Int).Value = Convert.ToInt32(objDS.Tables("Temp").Rows(0).Item("OrgLevel"))
            End With
            objAdp.Fill(objDS, "Auth")
            If objDS.Tables("Auth").Rows.Count > 0 Then rst = True Else rst = False
        End If

        Return rst
    End Function

    '儲存Auth_Relship(tmpAct=0--刪除,tmpAct=1--新增)
    Sub Save_OrgPlanInfo(ByVal tmpAcctPID As Integer, ByVal tmpRSID As Integer, Optional ByVal tmpAct As Integer = 1)
        'Dim sqlCmd As SqlCommand
        Dim sqlStr As String = ""
        Select Case tmpAct
            Case 0
            Case 1
                sqlStr = "insert into Org_OrgPlanInfo(RSID,OrgPName,ZipCode,Address,Phone,MasterName,ContactName,ContactEmail " & vbCrLf
                sqlStr += ",ContactCellPhone,TrainCap,ProTrainKind,FireControlState,ComSumm,ActNo,ModifyAcct,ModifyDate,PlanMaster,PlanMasterPhone,ContactFax,ContactSex " & vbCrLf
                sqlStr += ",ContactTitle,PayTax,AssistUnit,AssistUnit01,AssistUnit02,AssistUnit03,AssistUnitOther,ZipCODE6W) " & vbCrLf

                sqlStr += "select @RSID as RSID,OrgPName,ZipCode,Address,Phone,MasterName,ContactName,ContactEmail,ContactCellPhone,TrainCap,ProTrainKind " & vbCrLf
                sqlStr += ",FireControlState,ComSumm,ActNo,@ModifyAcct as ModifyAcct,getdate() as ModifyDate,PlanMaster,PlanMasterPhone,ContactFax,ContactSex,ContactTitle " & vbCrLf
                sqlStr += ",PayTax,AssistUnit,AssistUnit01,AssistUnit02,AssistUnit03,AssistUnitOther,ZipCODE6W from Org_OrgPlanInfoTemp where AcctPID= @AcctPID "
        End Select
        If sqlStr = "" Then Exit Sub
        Dim sqlCmd As New SqlCommand(sqlStr, objconn)
        With sqlCmd
            Select Case tmpAct
                Case 0
                Case 1
                    .Parameters.Clear()
                    .Parameters.Add("RSID", SqlDbType.Int).Value = tmpRSID
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID).Trim
                    .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                    If objconn.State = ConnectionState.Closed Then objconn.Open()
                    .ExecuteNonQuery()
            End Select
        End With
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    Throw ex
        'End Try
    End Sub

    '檢查是否有其他相同的共用資料待審
    Private Function Chk_RelshipTemp(ByVal tmpAcctPID As Integer) As DataTable
        Dim objAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim sqlStr As String
        Dim rst As DataTable = Nothing

        '取得審核通過的資料
        sqlStr = "select * from Auth_RelshipTemp where AcctPID= @AcctPID and ActMode is null"
        Try
            With objAdp
                .SelectCommand = New SqlCommand(sqlStr, objconn)
                .SelectCommand.Parameters.Add("AcctPID", tmpAcctPID).Value = tmpAcctPID
                .Fill(objDS, "Temp")
            End With
            If objDS.Tables("Temp").Rows.Count > 0 Then
                '用計畫、訓練機構、轄區去找出是否還有其他待審核的共用資料
                Dim tmpPlanID As Integer = Convert.ToInt32(objDS.Tables("Temp").Rows(0).Item("PlanID"))
                Dim tmpOrgID As Integer = Convert.ToInt32(objDS.Tables("Temp").Rows(0).Item("OrgID"))
                Dim tmpDistID As String = objDS.Tables("Temp").Rows(0).Item("DistID")
                Dim tmpRelship As String = Left(objDS.Tables("Temp").Rows(0).Item("Relship"), Len(objDS.Tables("Temp").Rows(0).Item("Relship")) - Len(objDS.Tables("Temp").Rows(0).Item("RID") & "/"))

                sqlStr = "select * from Auth_RelshipTemp where PlanID= @PlanID and OrgID= @OrgID and DistID= @DistID and Relship= @Relship and AcctPID<>@AcctPID and ActMode is null "
                With objAdp
                    .SelectCommand = New SqlCommand(sqlStr, objconn)
                    .SelectCommand.Parameters.Add("PlanID", SqlDbType.Int).Value = tmpPlanID
                    .SelectCommand.Parameters.Add("OrgID", SqlDbType.Int).Value = tmpOrgID
                    .SelectCommand.Parameters.Add("DistID", SqlDbType.VarChar).Value = tmpDistID
                    .SelectCommand.Parameters.Add("Relship", SqlDbType.VarChar).Value = tmpRelship
                    .SelectCommand.Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                    .Fill(objDS, "Relship")
                End With
                If objDS.Tables("Relship").Rows.Count > 0 Then
                    rst = objDS.Tables("Relship")
                End If
            End If
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Throw ex
        End Try
        Return rst
    End Function

    '取得新的RID 依 Auth_RelshipTemp@AcctPID
    Private Function Get_NewRID(ByVal tmpAcctPID As Integer) As String
        Dim rst As String = ""
        Dim sqlStr As String
        'Dim objAdp As New SqlDataAdapter
        'Dim objDS As New DataSet

        Call TIMS.OpenDbConn(objconn)

        '從業務關係暫存檔取出暫存的共用關係
        sqlStr = "select * from Auth_RelshipTemp where AcctPID= @AcctPID and ActMode is null "
        Dim dt1 As New DataTable 'Auth_RelshipTemp
        Dim oCmd1 As New SqlCommand(sqlStr, objconn)
        With oCmd1
            .Parameters.Clear()
            .Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count = 0 Then Return rst '異常

        '有暫存的共用關係時，產生新的RID，沒有的話RID設空值
        If dt1.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt1.Rows(0) 'Auth_RelshipTemp
            sqlStr = ""
            sqlStr &= " select RID from Auth_Relship "
            sqlStr &= " where 1=1"
            sqlStr &= " and PlanID= @PlanID "
            sqlStr &= " and OrgID= @OrgID "
            sqlStr &= " and DistID= @DistID "
            sqlStr &= " and OrgLevel= @OrgLevel "
            If dr1("OrgLevel") = 3 Then
                sqlStr &= " and Relship= @Relship+RID+'/' "
            End If

            Dim dt2 As New DataTable 'Auth_Relship
            Dim oCmd2 As New SqlCommand(sqlStr, objconn)
            With oCmd2
                .Parameters.Clear()
                .Parameters.Add("PlanID", SqlDbType.Int).Value = dr1("PlanID")
                .Parameters.Add("OrgID", SqlDbType.Int).Value = dr1("OrgID")
                .Parameters.Add("DistID", SqlDbType.NVarChar).Value = dr1("DistID")
                .Parameters.Add("OrgLevel", SqlDbType.Int).Value = dr1("OrgLevel")
                If dr1("OrgLevel") = 3 Then
                    .Parameters.Add("Relship", SqlDbType.VarChar).Value = dr1("Relship")
                End If
                dt2.Load(.ExecuteReader())
            End With

            If dt2.Rows.Count > 0 Then
                '同業務使用同一種RID
                rst = dt2.Rows(0).Item("RID") 'Auth_Relship
            Else
                Select Case CInt(dr1("OrgLevel"))
                    Case 2
                        Dim RIDKey As String
                        '取得RID的KeyWord
                        RIDKey = Split(Convert.ToString(dr1("Relship")), "/")(Convert.ToInt32(dr1("OrgLevel")) - 1)
                        '從正式業務關係中，依Relship、OrgLevel取得RID的最大數字+1
                        sqlStr = "select max(CONVERT(numeric, replace(RID,@RIDKey,'')))+1 NewRID from Auth_Relship where Relship like @Relship and OrgLevel= @OrgLevel "
                        Dim dt3 As New DataTable
                        Dim oCmd3 As New SqlCommand(sqlStr, objconn)
                        With oCmd3
                            .Parameters.Clear()
                            .Parameters.Add("RIDKey", SqlDbType.VarChar).Value = RIDKey
                            .Parameters.Add("Relship", SqlDbType.VarChar).Value = Convert.ToString(dr1("Relship")) & "%"
                            .Parameters.Add("OrgLevel", SqlDbType.Int).Value = Convert.ToInt32(dr1("OrgLevel"))
                            dt3.Load(.ExecuteReader())
                        End With
                        '組合出新的RID
                        rst = RIDKey & Convert.ToString(dt3.Rows(0)("NewRID"))
                    Case 3
                        Dim RIDKey As String
                        '取得RID的KeyWord
                        RIDKey = Split(Convert.ToString(dr1("Relship")), "/")(Convert.ToInt32(dr1("OrgLevel")) - 1)
                        '從正式業務關係中，依Relship、OrgLevel取得RID的最大數字+1
                        sqlStr = "select max(CONVERT(numeric, replace(RID,@RIDKey,'')))+1  NewRID from Auth_Relship where Relship like @Relship and OrgLevel= @OrgLevel "
                        Dim dt3 As New DataTable
                        Dim oCmd3 As New SqlCommand(sqlStr, objconn)
                        With oCmd3
                            .Parameters.Clear()
                            .Parameters.Add("RIDKey", SqlDbType.VarChar).Value = RIDKey
                            .Parameters.Add("Relship", SqlDbType.VarChar).Value = Convert.ToString(dr1("Relship")) & "%"
                            .Parameters.Add("OrgLevel", SqlDbType.Int).Value = Convert.ToInt32(dr1("OrgLevel"))
                            dt3.Load(.ExecuteReader())
                        End With
                        '組合出新的RID
                        rst = RIDKey & Right("000" & Convert.ToString(dt3.Rows(0)("NewRID")), 3) '目前上限為001~999 
                        If rst = "000" Then rst = "001" '若為000表示無資料，重新建立資料。
                End Select
            End If
            '依機構階層不同，用不同的方法產生RID
        End If

        Return rst
    End Function

#Region "NO USE"
    ''取得轄區職訓局的單位代號
    'Private Function Get_DistRID(ByVal tmpDistID As String) As String
    '    Dim rst As String = ""
    '    Dim sqlStr As String
    '    Dim objAdp As SqlDataAdapter
    '    Dim objDS As DataSet

    '    sqlStr = "select RID from Auth_Relship where OrgLevel=2 and DistID= @DistID"
    '    Try
    '        With objAdp
    '            .SelectCommand = New SqlCommand(sqlStr, objconn)
    '            .SelectCommand.Parameters.Add("DistID", SqlDbType.VarChar).Value = tmpDistID
    '            .Fill(objDS, "RID")
    '        End With
    '        rst = objDS.Tables("RID").Rows(0).Item("RID")
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '        Throw ex
    '    End Try
    '    Return rst
    'End Function
#End Region


    '存儲群組權限☆
    Sub Save_AccRWFun(ByVal tmpAcct As String, Optional ByVal tmpLID As Integer = 2)
        Dim dt As New DataTable
        Dim sqlStr As String = ""
        sqlStr = "select gid from auth_groupcontra where gtype='2'"
        Call TIMS.OpenDbConn(objconn)
        Dim cmdSelect As New SqlCommand(sqlStr, objconn)
        With cmdSelect
            'dt = New DataTable
            dt.Load(.ExecuteReader())
        End With
        sqlStr = "insert into auth_groupacct(gid,account,modifyacct,modifydate) "
        sqlStr += "values(@gid, @account, @modifyacct,getdate())"
        Call TIMS.OpenDbConn(objconn)
        Dim cmdInsert As New SqlCommand(sqlStr, objconn)
        With cmdInsert
            .Parameters.Clear()
            .Parameters.Add("gid", SqlDbType.VarChar).Value = Convert.ToString(dt.Rows(0)("gid"))
            .Parameters.Add("account", SqlDbType.VarChar).Value = tmpAcct
            .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
            .ExecuteNonQuery()
        End With
    End Sub

    Private Function Check_AccRWPlanTemp(ByVal tmpAcctPID As Integer) As Boolean
        Dim sqlAdp As New SqlDataAdapter
        Dim objDS As New DataSet
        Dim sqlStr As String
        Dim rst As Boolean = False

        sqlStr = "select AuditStatus from Auth_AccRWPlanTemp where AcctPID= @AcctPID and ActMode is null "
        Try
            With sqlAdp
                .SelectCommand = New SqlCommand(sqlStr, objconn)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("AcctPID", SqlDbType.Int).Value = tmpAcctPID
                .Fill(objDS, "Plan")
            End With
            If objDS.Tables("Plan").Rows.Count > 0 Then
                If IsDBNull(objDS.Tables("Plan").Rows(0).Item("AuditStatus")) = False Then rst = True
            Else
                rst = True
            End If
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Throw ex
        End Try
        Return rst
    End Function

    Private Function Check_AccountTemp(ByVal tmpAcctID As Integer) As Boolean
        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Try
            Dim sqlStr As String = "SELECT AUDITSTATUS FROM AUTH_ACCOUNTTEMP WHERE AcctID= @AcctID and ActMode is null "
            Dim SCommand As New SqlCommand(sqlStr, objconn)
            With SCommand
                .Parameters.Clear()
                .Parameters.Add("AcctID", SqlDbType.Int).Value = tmpAcctID
                dt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Throw ex
        End Try
        If TIMS.dtNODATA(dt) Then Return True
        If Not IsDBNull(dt.Rows(0).Item("AuditStatus")) Then Return True
        Return False
    End Function
#End Region

    'SELECT RID,COUNT(1) FROM AUTH_RELSHIP GROUP BY RID HAVING COUNT(1) >1
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            TB_Condition.Visible = True
            TR_Account.Visible = False
            TR_AuditAccount.Visible = False
            TR_Plan.Visible = False
            TR_AuditPlan.Visible = False
            TR_Acc1.Visible = False
            msg.Text = ""
            Me.ViewState("sort") = "RoleID"
        End If

        choice_button.Attributes("onclick") = "wopen('../../Common/LevPlan2.aspx?" &
                                                "YearsField=" & Me.YearsValue.ClientID &
                                                "&DistField=" & Me.DistValue.ClientID &
                                                "&PlanIDField=" & Me.PlanIDValue.ClientID &
                                                "&RIDField=" & Me.RIDValue.ClientID &
                                                "&OrgIDField=" & Me.OrgIDValue.ClientID &
                                                "&TextField=" & Me.TBplan.ClientID &
                                                "','計畫階段',1100,600,1);"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)

        '            If FunDr("Sech") = "1" Then
        '                but_search.Enabled = True
        '            Else
        '                but_search.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    'Button--查詢
    Private Sub but_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_search.Click
        Dim dt_Account As DataTable = Nothing
        Dim dt_Plan As DataTable = Nothing

        Me.ViewState("nameid") = TIMS.ClearSQM(nameid.Text)
        Me.ViewState("namefield") = TIMS.ClearSQM(namefield.Text)
        Me.ViewState("OrgID") = TIMS.ClearSQM(OrgIDValue.Value)
        Me.ViewState("ApplyType") = TIMS.GetListValue(ApplyType) '.SelectedValue
        Me.ViewState("Resultsrh") = TIMS.GetListValue(Resultsrh) '.SelectedValue.Trim(" ")
        Me.ViewState("planToAcct") = False
        Me.ViewState("acctToPlan") = False

        If ApplyType.SelectedValue = "Account" Then '帳號審核
            Dim chk As Boolean = False

            dt_Account = Get_AccountApplyList()
            If Not dt_Account Is Nothing Then
                If dt_Account.Rows.Count > 0 Then chk = True
            End If
            If chk = True Then
                DataGrid1.DataSource = dt_Account
                DataGrid1.DataKeyField = "AcctID"
                DataGrid1.CurrentPageIndex = 0
                DataGrid1.DataBind()

                TR_Account.Visible = True
                TR_AuditAccount.Visible = True
                TR_Acc1.Visible = True
                TR_Plan.Visible = False
                TR_AuditPlan.Visible = False
                msg.Text = ""
            Else
                TR_Account.Visible = False
                TR_AuditAccount.Visible = False
                TR_Acc1.Visible = False
                TR_Plan.Visible = False
                TR_AuditPlan.Visible = False
                msg.Text = "查無帳號申請資料"
            End If
        ElseIf ApplyType.SelectedValue = "Plan" Then    '計劃審核
            Dim chk As Boolean = False

            dt_Plan = Get_PlanApplyList()
            If Not dt_Plan Is Nothing Then
                If dt_Plan.Rows.Count > 0 Then chk = True
            End If
            If chk = True Then
                Datagrid2.DataSource = dt_Plan
                Datagrid2.DataKeyField = "AcctPID"
                Datagrid2.CurrentPageIndex = 0
                Datagrid2.DataBind()

                TR_Plan.Visible = True
                TR_AuditPlan.Visible = True
                TR_Account.Visible = False
                TR_AuditAccount.Visible = False
                TR_Acc1.Visible = False
                msg.Text = ""
            Else
                TR_Plan.Visible = False
                TR_AuditPlan.Visible = False
                TR_Account.Visible = False
                TR_AuditAccount.Visible = False
                TR_Acc1.Visible = False
                msg.Text = "查無計畫申請資料"
            End If
        End If
    End Sub

    '帳號審核
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim dt_Plan As DataTable
        Dim accountNote As LinkButton = e.Item.FindControl("AccountNote")

        Select Case e.CommandName
            Case "AuditPlan"    '點選帳號所附屬的待審核計畫，進行計劃審核
                '獲取帳號條件
                Me.ViewState("nameid") = e.Item.Cells(0).Text
                Me.ViewState("namefield") = ""
                Me.ViewState("OrgID") = ""
                Me.ViewState("Resultsrh") = "X" '透過帳號帶出的，都是預設未審核
                '獲取計劃資料
                dt_Plan = Get_PlanApplyList()
                If dt_Plan.Rows.Count > 0 Then
                    Datagrid2.DataSource = dt_Plan
                    Datagrid2.DataKeyField = "AcctPID"
                    Me.ViewState("acctToPlan") = True
                    Datagrid2.CurrentPageIndex = 0
                    Datagrid2.DataBind()

                    TR_Account.Visible = False
                    TR_AuditAccount.Visible = False
                    TR_Acc1.Visible = False
                    TR_Plan.Visible = True
                    TR_AuditPlan.Visible = True
                End If
                Me.ViewState("nameid") = nameid.Text
                Me.ViewState("namefield") = namefield.Text
                Me.ViewState("OrgID") = OrgIDValue.Value
                Me.ViewState("Resultsrh") = Resultsrh.SelectedValue.Trim
        End Select
    End Sub

    '帳號審核
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType = ListItemType.Header Then auditAllAccount = e.Item.FindControl("AuditAllAccount")
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim dr_Account As DataRowView = e.Item.DataItem
            Dim auditList As DropDownList = e.Item.FindControl("AuditListAccount")
            Dim auditStatus As Label = e.Item.FindControl("AuditAccountStatus")
            Dim accountNote As LinkButton = e.Item.FindControl("AccountNote")

            Select Case dr_Account("AuditStatus")   '審核狀態
                Case "Y"    '審核通過
                    auditList.Items.Item(1).Selected = True
                    auditList.Enabled = False
                    auditStatus.Visible = False
                Case "N"    '審核不通過
                    auditList.Items.Item(2).Selected = True
                    auditList.Enabled = False
                    auditStatus.Visible = False
                Case "X"    '未審核
                    auditList.Items.Item(0).Selected = True
                    auditList.Enabled = True
                    auditStatus.Visible = False
                Case Else   '無法判定
                    auditList.Visible = False
                    auditStatus.Visible = False
            End Select

            If dr_Account("CntAcctPID") <> 0 Then   '顯示是否有順帶申請計畫未審
                accountNote.Text = "有" & Convert.ToString(dr_Account("CntAcctPID")) & "個計畫申請待審核"
                If Me.ViewState("planToAcct") = True Then accountNote.Enabled = False Else accountNote.Enabled = True
                '如果承辦人沒有權限可以審核相關計畫時，就將連結Disable。
                If sm.UserInfo.RoleID <> "0" And sm.UserInfo.RoleID <> "1" Then
                    Dim dt_Plan As DataTable

                    Me.ViewState("nameid") = dr_Account("Account")
                    Me.ViewState("namefield") = ""
                    Me.ViewState("OrgID") = ""
                    Me.ViewState("Resultsrh") = "X"
                    dt_Plan = Get_PlanApplyList()
                    If dt_Plan Is Nothing Then
                        accountNote.Enabled = False
                    Else
                        If dt_Plan.Rows.Count <= 0 Then accountNote.Enabled = False
                    End If
                    Me.ViewState("nameid") = nameid.Text
                    Me.ViewState("namefield") = namefield.Text
                    Me.ViewState("OrgID") = OrgIDValue.Value
                    Me.ViewState("Resultsrh") = Resultsrh.SelectedValue.Trim
                End If
            End If

            selectAllAccounts += " if(document.getElementById('" & auditList.ClientID & "').disabled==false){ document.getElementById('" & auditList.ClientID & "').value=document.getElementById('" & auditAllAccount.ClientID & "').value; }"
        End If
        If e.Item.ItemType = ListItemType.Footer Then
            auditAllAccount.Attributes("onChange") = selectAllAccounts
        End If
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        Dim dt_Account As DataTable

        dt_Account = Get_AccountApplyList()
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataSource = dt_Account
        DataGrid1.DataKeyField = "AcctID"
        DataGrid1.DataBind()
    End Sub

    '計劃審核
    Private Sub Datagrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid2.ItemCommand
        Dim dt_Account As DataTable
        Dim planNote As LinkButton = e.Item.FindControl("PlanNote")

        Select Case e.CommandName
            Case "AuditAccount" '點選計劃審核所對應的帳號申請，進行帳號審核
                '獲取條件
                Me.ViewState("nameid") = e.Item.Cells(0).Text
                Me.ViewState("namefield") = ""
                Me.ViewState("OrgID") = ""
                Me.ViewState("Resultsrh") = "X" '透過計畫帶出的，都是預設未審核
                '獲取帳號資料
                dt_Account = Get_AccountApplyList()
                If dt_Account.Rows.Count > 0 Then
                    DataGrid1.DataSource = dt_Account
                    DataGrid1.DataKeyField = "AcctID"
                    Me.ViewState("planToAcct") = True
                    DataGrid1.DataBind()
                    TR_Account.Visible = True
                    TR_AuditAccount.Visible = True
                    TR_Acc1.Visible = True
                    TR_Plan.Visible = False
                    TR_AuditPlan.Visible = False
                End If
                Me.ViewState("nameid") = nameid.Text
                Me.ViewState("namefield") = namefield.Text
                Me.ViewState("OrgID") = OrgIDValue.Value
                Me.ViewState("Resultsrh") = Resultsrh.SelectedValue.Trim
        End Select
    End Sub

    '計劃審核
    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        If e.Item.ItemType = ListItemType.Header Then auditAllPlan = e.Item.FindControl("AuditAllPlan")
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim dr_Plan As DataRowView = e.Item.DataItem
            Dim auditList As DropDownList = e.Item.FindControl("AuditListPlan")
            Dim auditStatus As Label = e.Item.FindControl("AuditPlanStatus")
            Dim planNote As LinkButton = e.Item.FindControl("PlanNote")
            Dim othPlanNote As Label = e.Item.FindControl("OthPlanNote")

            Select Case dr_Plan("AuditStatus")  '審核狀態
                Case "Y"    '審核通過
                    auditList.Items.Item(1).Selected = True
                    auditList.Enabled = False
                    auditStatus.Visible = False
                Case "N"    '審核不通過
                    auditList.Items.Item(2).Selected = True
                    auditList.Enabled = False
                    auditStatus.Visible = False
                Case "X"    '未審核
                    auditList.Items.Item(0).Selected = True
                    auditList.Enabled = True
                    auditStatus.Visible = False
                Case Else   '無法判定
                    auditList.Visible = False
                    auditStatus.Visible = False
            End Select
            '判斷是否有帳號申請待審核，有的話不能進行計劃審核
            If dr_Plan("AcctID") <> 0 Then
                planNote.Text = "帳號申請待審核。"
                auditList.Visible = False
                auditStatus.Text = "有帳號申請未審核，所以計畫無法審核。"
                auditStatus.Visible = True
                If Me.ViewState("acctToPlan") = True Then planNote.Enabled = False Else planNote.Enabled = True
            End If

            Dim msgStr As String = ""
            '計畫需被共用時，需要顯示的告警訊息
            'selectAllPlans += " if(document.getElementById('" & auditList.ClientID & "').disabled==false){ document.getElementById('" & auditList.ClientID & "').value=document.getElementById('" & auditAllPlan.ClientID & "').value; }"
            'If dr_Plan("Shared") = "Y" Then
            '    If dr_Plan("AcctID") <> 0 Then othPlanNote.Text += "<br>"
            '    othPlanNote.Text += "此計劃跨區申請需被共用。"
            '    msgStr += e.Item.Cells(1).Text & "所申請的計畫：" & e.Item.Cells(3).Text & "，\n"
            '    msgStr += e.Item.Cells(2).Text & " 不存在 " & TIMS.GET_DistName(e.Item.Cells(7).Text) & "，\n"
            '    msgStr += "確認是否要通過申請並完成共用？"
            '    auditList.Attributes("onChange") = "if(document.getElementById('" & auditList.ClientID & "').disabled==false){ if(document.getElementById('" & auditList.ClientID & "').value=='Y'){ if(confirm('" & msgStr & "')){ document.getElementById('" & auditList.ClientID & "').value='Y'; }else{ document.getElementById('" & auditList.ClientID & "').value='N'; }}}"
            '    selectAllPlans += " if(document.getElementById('" & auditList.ClientID & "').disabled==false){ if(document.getElementById('" & auditList.ClientID & "').value=='Y'){ if(confirm('" & msgStr & "')){ document.getElementById('" & auditList.ClientID & "').value='Y'; }else{ document.getElementById('" & auditList.ClientID & "').value='N'; }}}"
            'End If
        End If
        If e.Item.ItemType = ListItemType.Footer Then
            auditAllPlan.Attributes("onChange") = selectAllPlans
        End If
    End Sub

    Private Sub Datagrid2_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles Datagrid2.PageIndexChanged
        Dim dt_plan As DataTable

        dt_plan = Get_PlanApplyList()
        Datagrid2.CurrentPageIndex = e.NewPageIndex
        Datagrid2.DataSource = dt_plan
        Datagrid2.DataKeyField = "AcctPID"
        Datagrid2.DataBind()
    End Sub

    'Button帳號審核確認
    Private Sub AuditAccont_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AuditAccont.Click
        Dim msgStr As String = ""

        For i As Integer = 0 To DataGrid1.Items.Count - 1
            Dim auditList As DropDownList = DataGrid1.Items(i).FindControl("AuditListAccount")

            '當帳號審核狀態是可以被選取改辦的狀態下，才進行處理
            If auditList.Enabled = True And auditList.Visible = True And auditList.SelectedValue <> "" Then
                '審核通過且帳號不存在正式資料庫中
                If auditList.SelectedValue = "Y" Then
                    If Chk_Account(CInt(DataGrid1.DataKeys.Item(i))) = True Then
                        Update_AccountApplyList(CInt(DataGrid1.DataKeys.Item(i)), "N")
                        msgStr += DataGrid1.Items(i).Cells(0).Text & "已經存在，所以無法正確審核通過，將自動設為不通過。\n"
                    Else
                        If Check_AccountTemp(CInt(DataGrid1.DataKeys.Item(i))) = False Then
                            Save_Account(CInt(DataGrid1.DataKeys.Item(i)))    '將資料從Temp轉到正式
                            Save_AccRWFun(DataGrid1.Items(i).Cells(0).Text) '產生使用者權限功能
                            Update_AccountApplyList(CInt(DataGrid1.DataKeys.Item(i)), "Y")    '更新TEMP的審核狀態為通過
                        Else
                            msgStr += DataGrid1.Items(i).Cells(0).Text & "已經審核過，所以無法再次審核。"
                        End If
                    End If
                Else   '審核不通過，或帳號已經存在正式資料庫時
                    Update_AccountApplyList(CInt(DataGrid1.DataKeys.Item(i)), "N")    '更新TEMP的審核狀態為不通過
                    '如果有申請計畫，自動將計劃失效
                    Update_PlanApplyList(DataGrid1.Items(i).Cells(0).Text, "C", "帳號申請不通過，計畫申請自動失效")
                End If
            End If
        Next

        '判斷是否由計劃審核轉到帳號審核，"是"則轉回計劃審核，"否"則依條件重新篩選
        If Me.ViewState("planToAcct") = True Then

            Try
                Dim planNote As LinkButton = Nothing
                If Convert.ToString(Me.ViewState("acctPID")) <> "" Then
                    planNote = Datagrid2.Items(Datagrid2.AccessKey.IndexOf(Me.ViewState("acctPID"))).FindControl("PlanNote")
                End If
                If Not planNote Is Nothing Then
                    planNote.Text = ""
                End If
            Catch ex As Exception
                Throw ex
            End Try

            Me.ViewState("planToAcct") = False
            TR_Account.Visible = False
            TR_AuditAccount.Visible = False
            TR_Acc1.Visible = False
            TR_Plan.Visible = True
            TR_AuditPlan.Visible = True
        Else
            nameid.Text = Me.ViewState("nameid")
            namefield.Text = Me.ViewState("namefield")
            OrgIDValue.Value = Me.ViewState("OrgID")
            ApplyType.SelectedValue = Me.ViewState("ApplyType")
            Resultsrh.SelectedValue = Me.ViewState("Resultsrh")
            but_search_Click(sender, e) '重新查詢

        End If
    End Sub

    'Button計劃審核確認
    Private Sub AuditPlan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AuditPlan.Click
        'Dim msgStr As String

        For i As Integer = 0 To Datagrid2.Items.Count - 1
            Dim auditList As DropDownList = Datagrid2.Items(i).FindControl("AuditListPlan")
            Dim othPlanNote As Label = Datagrid2.Items(i).FindControl("OthPlanNote")

            '當帳號審核狀態是可以被選取改辦的狀態下，才進行處理
            If auditList.Enabled = True And auditList.Visible = True And auditList.SelectedValue <> "" Then
                If auditList.SelectedValue = "Y" Then   '審核通過
                    If Check_AccRWPlanTemp(CInt(Datagrid2.DataKeys.Item(i))) = False Then
                        Save_Plan(CInt(Datagrid2.DataKeys.Item(i)), Datagrid2.Items(i).Cells(8).Text)     '將計畫從TEMP轉到正式
                        Update_PlanApplyList(CInt(Datagrid2.DataKeys.Item(i)), "Y")   '更新TEMP的審核狀態為通過
                        Update_AuthAccount(Datagrid2.Items(i).Cells(0).Text, "Y")
                        'If Me.ViewState("acctID") Is Nothing Then Me.ViewState("acctID") = Datagrid2.Items(i).Cells(0).Text
                    End If
                ElseIf auditList.SelectedValue = "N" Then   '審核不通過
                    Update_PlanApplyList(CInt(Datagrid2.DataKeys.Item(i)), "N")   '更新TEMP的審核狀態為不通過
                End If
            End If

            ''判斷是否由帳號審核轉到計畫審核，"是"則轉回帳號審核，"否"則依條件重新篩選
            'If Me.ViewState("acctToPlan") = True Then
            '    If Not Me.ViewState("acctID") Is Nothing Then
            '        Dim accountNote As LinkButton = DataGrid1.Items(DataGrid1.AccessKey.IndexOf(Me.ViewState("acctID"))).FindControl("AccountNote")
            '        accountNote.Text = ""
            '        Me.ViewState("acctID") = Nothing
            '    End If
            'End If
        Next

        '判斷是否由帳號審核轉到計畫審核，"是"則轉回帳號審核，"否"則依條件重新篩選
        If Me.ViewState("acctToPlan") = True Then
            'Dim accountNote As LinkButton = DataGrid1.Items(DataGrid1.AccessKey.IndexOf(Me.ViewState("acctID"))).FindControl("AccountNote")
            'accountNote.Text = ""

            Me.ViewState("acctToPlan") = False
            TR_Account.Visible = True
            TR_AuditAccount.Visible = True
            TR_Acc1.Visible = True
            TR_Plan.Visible = False
            TR_AuditPlan.Visible = False
        Else
            nameid.Text = Me.ViewState("nameid")
            namefield.Text = Me.ViewState("namefield")
            OrgIDValue.Value = Me.ViewState("OrgID")
            ApplyType.SelectedValue = Me.ViewState("ApplyType")
            Resultsrh.SelectedValue = Me.ViewState("Resultsrh")
            but_search_Click(sender, e)
        End If

    End Sub

    Private Sub Btn_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Cancel.Click
        '判斷是否由計劃審核轉到帳號審核，"是"則轉回計劃審核，"否"則依條件重新篩選
        If Me.ViewState("planToAcct") = True Then
            nameid.Text = Me.ViewState("nameid")
            namefield.Text = Me.ViewState("namefield")
            OrgIDValue.Value = Me.ViewState("OrgID")
            ApplyType.SelectedValue = Me.ViewState("ApplyType")
            Resultsrh.SelectedValue = Me.ViewState("Resultsrh")
            but_search_Click(sender, e)
        Else
            TR_Account.Visible = False
            TR_AuditAccount.Visible = False
            TR_Acc1.Visible = False
        End If
    End Sub

    Private Sub Btn_Cancel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Cancel2.Click
        '判斷是否由帳號審核轉到計畫審核，"是"則轉回帳號審核，"否"則依條件重新篩選
        If Me.ViewState("acctToPlan") = True Then
            nameid.Text = Me.ViewState("nameid")
            namefield.Text = Me.ViewState("namefield")
            OrgIDValue.Value = Me.ViewState("OrgID")
            ApplyType.SelectedValue = Me.ViewState("ApplyType")
            Resultsrh.SelectedValue = Me.ViewState("Resultsrh")
            but_search_Click(sender, e)
        Else
            TR_Plan.Visible = False
            TR_AuditPlan.Visible = False
        End If
    End Sub
End Class
