Partial Class SYS_01_002
    Inherits AuthBasePage

    Dim dtG1 As DataTable '取得所選取帳號目前已賦予之計畫
    Dim dtG2 As DataTable '取得登入帳號可賦予之計畫
    '更動的TABLE: Auth_AccRWPlan pk@Account, PlanID , RID
    '一切都是為了 RID
    '暫定 目前跨轄區只有分署(中心)單位可使用
    Const Cst_DefInput As String = "請輸入姓名關鍵字"

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    Dim objconn As SqlConnection

    'SELECT * FROM AUTH_ACCRWPLAN WHERE RID ='B1000' 'SELECT * FROM AUTH_RELSHIP WHERE RID ='B1000'

#Region "Sub"
    '顯示此登入者的可用計畫年度dropdownlist
    Sub Show_Years()
        '異常資料。 SELECT RID,COUNT(1) FROM Auth_Relship GROUP BY RID HAVING COUNT(1)>1
        'SELECT * FROM Auth_Relship where RID='B1000'
        'Dim sql As String 'Dim objtable As DataTable

        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)

        Lis_Plan.Items.Clear()

        Dim hDIC As New Hashtable From {{"account", v_Lis_acc}}
        Dim objstr As String = "select MAX(RoleID) from auth_account where account=@account"
        Dim oTmpRoleID As Object = DbAccess.ExecuteScalar(objstr, objconn, hDIC) '角色(SELECT * FROM ID_ROLE)
        Dim iTmpRoleID As Integer = If(Convert.ToString(oTmpRoleID) <> "" AndAlso TIMS.IsNumeric1(oTmpRoleID), Val(oTmpRoleID), -1)

        Dim objtable As DataTable = Nothing
        If sm.UserInfo.RoleID <= 1 Then
            '假如登入者為系統管理者
            Dim pms1 As New Hashtable From {{"RID", RIDValue.Value}}
            objstr = "SELECT DISTINCT YEARS FROM VIEW_LOGINPLAN WHERE DistID=(SELECT DistID FROM Auth_Relship WHERE RID=@RID) ORDER BY YEARS DESC"
            objtable = DbAccess.GetDataTable(objstr, objconn, pms1)
        Else
            If iTmpRoleID = 1 Then
                Dim pms1 As New Hashtable From {{"RID", RIDValue.Value}}
                objstr = "SELECT DISTINCT YEARS FROM VIEW_LOGINPLAN WHERE DistID=(SELECT DistID FROM Auth_Relship WHERE RID=@RID) ORDER BY YEARS DESC"
                objtable = DbAccess.GetDataTable(objstr, objconn, pms1)
            Else
                Dim pms1 As New Hashtable From {{"PlanID", sm.UserInfo.PlanID}}
                objstr = "SELECT DISTINCT YEARS FROM VIEW_LOGINPLAN WHERE PlanID=@PlanID ORDER BY YEARS DESC"
                objtable = DbAccess.GetDataTable(objstr, objconn, pms1)
            End If
        End If
        If objtable Is Nothing AndAlso Years.Items.Count > 0 Then Years.Items.Clear()
        'Dim objtable As DataTable = DbAccess.GetDataTable(objstr, objconn)
        With Years
            .DataSource = objtable
            .DataTextField = "YEARS"
            .DataValueField = "YEARS"
            .DataBind()
        End With
        Years.Visible = If(Years.Items.Count = 1, False, True)
        Common.SetListItem(Years, sm.UserInfo.Years)
    End Sub

    '*** 取出目前可賦予之權限 sm.UserInfo.UserID
    Private Sub getRelishTB()
        'Dim sda As New SqlDataAdapter 'Dim ds As DataSet
        Dim sql As String = ""
        Dim sql_authuse As String = ""

        'ds = New DataSet
        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc) 'Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)
        Dim YearsSVal As String = TIMS.GetListValue(Years)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        If v_Lis_acc <> "" And Years.SelectedValue <> "" Then
            lsbShow.Items.Clear()
            Try
                Dim xFlag1 As Boolean = True '分署(中心)、署(局) @True '非分署(中心)、署(局)  @False
                Dim xFlag2 As Boolean = False '不只能賦予本身的給別人@True '只能賦予本身的給別人@False 

                sql = "SELECT ISNULL(AUTHUSE,'0') AUTHUSE FROM AUTH_ACCOUNT WHERE ACCOUNT= @account"
                Dim sCmd2 As New SqlCommand(sql, objconn)
                Call TIMS.OpenDbConn(objconn)
                Dim dtAccount As New DataTable
                With sCmd2
                    .Parameters.Clear()
                    .Parameters.Add("account", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    dtAccount.Load(.ExecuteReader())
                End With
                If dtAccount.Rows.Count > 0 Then
                    'Auth_Account.AuthUse 若是1, 不只能賦予本身的給別人而以還可以加注外來權限
                    'Auth_Account.AuthUse 若不是1, 則只能賦予本身的給別人而以
                    If Convert.ToString(dtAccount.Rows(0)("authuse")) = "1" Then
                        xFlag2 = True '有權限
                    End If
                End If

                'With sda
                '    .SelectCommand = New SqlCommand(sql, objconn)
                '    .SelectCommand.Parameters.Clear()
                '    .SelectCommand.Parameters.Add("account", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '    .Fill(ds, "authuse")
                'End With

                'Auth_Account.AuthUse 若不是1, 則只能賦予本身的給別人而以
                'Auth_Account.AuthUse 若是1, 不只能賦予本身的給別人而以還可以加注外來權限
                sql_authuse = "" '(不限定)
                If Not xFlag2 Then
                    '該帳號本身計畫(限定)
                    sql_authuse = " and a.planid in (select planid from auth_accrwplan where account='" & sm.UserInfo.UserID & "')" & vbCrLf
                End If

                If RIDValue.Value.Length <= 1 Then
                    '分署(中心)、署(局)  
                    sql = ""
                    sql &= " select a.planid,b.distid,b.rid,o.orgid,o.orgname,a.planname,a.years" & vbCrLf
                    sql &= " ,null OrgName3" & vbCrLf
                    sql &= " ,null OrgName2" & vbCrLf
                    sql &= " ,a.tplanid tplanid" & vbCrLf
                    sql &= " from VIEW_LOGINPLAN a" & vbCrLf
                    sql &= " join AUTH_RELSHIP b on b.distid=a.distid" & vbCrLf
                    sql &= " join ORG_ORGINFO o on b.orgid=o.orgid" & vbCrLf
                    sql &= " where a.years='" & Years.SelectedValue & "' and b.rid='" & RIDValue.Value & "'" & vbCrLf
                    '同轄區
                    If Convert.ToString(sm.UserInfo.RID) = RIDValue.Value Then
                        If Not xFlag2 Then sql += sql_authuse '本身
                    End If

                    '不同轄區 可將計畫 賦予 單位
                    If Convert.ToString(sm.UserInfo.RID) <> RIDValue.Value Then
                        sql &= " union" & vbCrLf
                        sql &= " select a.planid,b.distid,b.rid,o.orgid,o.orgname,a.planname,a.years" & vbCrLf
                        sql &= " ,null OrgName3" & vbCrLf
                        sql &= " ,null OrgName2" & vbCrLf
                        sql &= " ,a.tplanid tplanid" & vbCrLf
                        sql &= " from VIEW_LOGINPLAN a" & vbCrLf
                        sql &= " join AUTH_RELSHIP b on b.distid=a.distid" & vbCrLf
                        sql &= " join ORG_ORGINFO o on b.orgid=o.orgid" & vbCrLf
                        sql &= " where a.years='" & YearsSVal & "' and b.RID='" & Convert.ToString(sm.UserInfo.RID) & "'" & vbCrLf
                        If Not xFlag2 Then sql += sql_authuse '本身

                        'ORDER BY 
                        If RIDValue.Value > Convert.ToString(sm.UserInfo.RID) Then
                            sql += "order by rid desc, planid"
                        Else
                            sql += "order by rid, planid"
                        End If
                    End If

                Else
                    '非分署(中心)、署(局) 
                    xFlag1 = False

                    Dim rPlanID As String = ""

                    If ChkRIDOrgLevel2(RIDValue.Value, rPlanID) Then
                        '為輔助地方政府
                        '查看該年度 %計畫% 的 機構層級為 2 且底下有機構承辦業務
                        sql = ""
                        sql &= " SELECT o.orgid,o.orgname" & vbCrLf
                        sql &= " ,a.PlanName+'　('+o.OrgName+')' PlanName ,a.years" & vbCrLf
                        sql &= " ,AR.planid, AR.RID, AR.DistID" & vbCrLf
                        sql &= " ,null OrgName3" & vbCrLf
                        sql &= " ,o.orgname OrgName2" & vbCrLf
                        sql &= " ,a.tplanid tplanid" & vbCrLf

                        sql &= " FROM AUTH_RELSHIP ar" & vbCrLf
                        sql &= " join VIEW_LOGINPLAN a on a.planid=ar.planid" & vbCrLf
                        sql &= " JOIN ORG_ORGINFO o on o.orgid=ar.orgid" & vbCrLf
                        sql &= " where ar.orglevel='2'" & vbCrLf
                        sql &= " and a.years='" & YearsSVal & "'" & vbCrLf
                        sql &= " and a.planid ='" & rPlanID & "'" & vbCrLf

                        '同轄區
                        If Convert.ToString(sm.UserInfo.RID) = RIDValue.Value Then
                            'sql_authuse = Replace(sql_authuse, " a.", " ip.") '換字
                            If Not xFlag2 Then sql += sql_authuse '本身
                        End If
                    Else
                        'select rid ,count(1) cnt from Auth_Relship group by rid having count(1)>1
                        '委訓單位(OrgLevel可能為2.3)
                        sql = ""
                        sql &= " SELECT o.orgid,o.orgname" & vbCrLf
                        sql &= " ,a.PlanName ,a.years" & vbCrLf
                        sql &= " ,AR.planid, AR.RID, AR.DistID" & vbCrLf
                        sql &= " ,o3.OrgName OrgName3" & vbCrLf
                        sql &= " ,o2.OrgName OrgName2" & vbCrLf
                        sql &= " ,a.tplanid tplanid" & vbCrLf

                        sql &= " FROM AUTH_RELSHIP ar" & vbCrLf
                        sql &= " join VIEW_LOGINPLAN a on a.planid=ar.planid" & vbCrLf
                        sql &= " JOIN ORG_ORGINFO o on ar.orgid=o.orgid" & vbCrLf
                        sql &= " left join MVIEW_RELSHIP23 r3 on r3.RID3=ar.RID" & vbCrLf
                        sql &= " left join org_orginfo o3 on o3.orgid=r3.orgid3" & vbCrLf
                        sql &= " left join org_orginfo o2 on o2.orgid=r3.orgid2" & vbCrLf
                        sql &= " where a.years='" & YearsSVal & "'" & vbCrLf
                        sql &= " and ar.orgid=(select orgid from Auth_Relship where rid='" & RIDValue.Value & "')" & vbCrLf
                        '同轄區
                        If Convert.ToString(sm.UserInfo.RID) = RIDValue.Value Then
                            'sql_authuse = Replace(sql_authuse, " a.", " ip.") '換字
                            If Not xFlag2 Then sql += sql_authuse '本身
                        End If
                    End If
                End If

                Dim sCmd As New SqlCommand(sql, objconn)
                Call TIMS.OpenDbConn(objconn)
                Dim dtData As New DataTable
                With sCmd
                    .Parameters.Clear()
                    dtData.Load(.ExecuteReader())
                End With

                'With sda
                '    .SelectCommand = New SqlCommand(sql, objconn)
                '    .SelectCommand.Parameters.Clear()
                '    .Fill(ds, "data")
                'End With

                trView.Visible = False

                Call ClearDataGrid2()
                Call ClearDataGrid1()
                Datagrid2.Visible = False
                DataGrid1.Visible = False

                If dtData.Rows.Count >= 0 Then
                    trView.Visible = True

                    dtG1 = getDT(Years.SelectedValue, v_Lis_acc)
                    dtG2 = getLis_Plan(Me, Years.SelectedValue, v_Lis_acc)

                    Datagrid2.Visible = True
                    Datagrid2.DataSource = dtData.DefaultView
                    Datagrid2.DataBind()
                    Data_Bind()
                End If

                'If Not sda Is Nothing Then sda.Dispose()
                'If Not ds Is Nothing Then ds.Dispose()
            Catch ex As Exception
                Me.Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
            End Try

        End If
    End Sub

    '建置訓練計畫清單
    Private Sub Data_Bind()
        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
        Dim YearsSVal As String = TIMS.GetListValue(Years)
        'Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)

        Dim pms1 As New Hashtable From {{"ACCOUNT", v_Lis_acc}, {"YEARS", YearsSVal}}
        Dim Sql As String = ""
        Sql &= " select a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
        Sql &= " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
        Sql &= "  ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')' END" & vbCrLf
        Sql &= "  +case when f.Clsyear is null or f.Clsyear > b.Years then ' ' else '…已停用'+CONVERT(varchar, f.Clsyear) end AS PlanName" & vbCrLf
        Sql &= " ,d.OrgID, d.OrgName, ac.isused, c.DistID, c.OrgLevel" & vbCrLf
        Sql &= " ,r3.RID2 CRID,o2.OrgID COrgID, o2.OrgName COrgName" & vbCrLf
        Sql &= " FROM Auth_AccRWPlan a" & vbCrLf
        Sql &= " join ID_Plan b on a.PlanID=b.PlanID" & vbCrLf
        Sql &= " join Auth_Relship c on c.RID=a.RID" & vbCrLf
        Sql &= " join Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
        Sql &= " join ID_District e on b.DistID=e.DistID" & vbCrLf
        Sql &= " join Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
        Sql &= " join auth_account ac on a.account=ac.account" & vbCrLf
        Sql &= " LEFT JOIN VIEW_RELSHIP23X r3 on r3.distid =e.distid and r3.planid=b.planid and r3.rid3=c.rid" & vbCrLf
        Sql &= " left join Org_OrgInfo o2 on o2.OrgID=r3.OrgID2" & vbCrLf
        Sql &= " WHERE a.ACCOUNT=@ACCOUNT AND b.YEARS=@YEARS" & vbCrLf
        Dim objtable As DataTable = DbAccess.GetDataTable(Sql, objconn, pms1)

        msg.Text = "查無資料"
        DataGrid1.Visible = False

        Me.hidRoleId.Value = TIMS.Get_RoleID(v_Lis_acc, objconn)

        If objtable.Rows.Count > 0 Then
            msg.Text = ""
            DataGrid1.Visible = True

            DataGrid1.DataSource = objtable
            DataGrid1.DataBind()
        End If
    End Sub

    '儲存訓練計畫
    Private Sub SAVE_Auth_AccRWPlan(ByVal count As Integer, ByRef objtable As DataTable, ByRef objAdapter As SqlDataAdapter, ByRef Errmsg As String)
        Dim objrow As DataRow
        Dim sql As String

        Errmsg = ""
        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
        Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        objrow = objtable.NewRow()
        objrow("Account") = v_Lis_acc '賦予的帳號
        objrow("PlanID") = v_Lis_Plan '賦予的計畫

        '跨轄區 【此功能應開放為只有分署(中心)有權限使用，可賦予本系統委訓、中心的使用者】 ，依RID頭一個字，可判斷轄區是否相同
        If sm.UserInfo.RID <> RIDValue.Value.Chars(0) Then
            objrow("RID") = OrgRID.Value '隸屬機構的RID
            If count = 0 Then '若權限數為0
                objrow("createByAcc") = "Y" '創由帳號建立
            Else
                objrow("createByAcc") = "N"
            End If
        Else
            '同屬轄區
            '但必須要檢查機構是否共用於此計劃
            If RIDValue.Value.Length > 1 Then '分署(中心)以下權限賦予
                Dim dr As DataRow '補助地方政府之計畫直接賦予

                sql = " SELECT * FROM Auth_Relship" & vbCrLf
                sql &= " WHERE PlanID='" & v_Lis_Plan & "'" & vbCrLf
                sql &= " AND OrgID=(SELECT OrgID FROM Auth_Relship WHERE RID='" & RIDValue.Value & "')" & vbCrLf
                sql &= " AND RID='" & RIDValue.Value & "'" & vbCrLf

                dr = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then
                    'Common.MessageBox(Page, "要先共用機構到此計劃才能進行計畫賦予")
                    Errmsg += "要先共用機構到此計劃才能進行計畫賦予" & vbCrLf
                    Exit Sub
                Else
                    objrow("RID") = dr("RID")
                    objrow("createByAcc") = "N"
                End If
            Else
                objrow("RID") = RIDValue.Value '分署(中心)權限賦予
                objrow("createByAcc") = "N"
            End If
        End If
        objrow("ModifyAcct") = sm.UserInfo.UserID
        objrow("ModifyDate") = Now()

        objtable.Rows.Add(objrow)
        DbAccess.UpdateDataTable(objtable, objAdapter)
    End Sub

    ''' <summary>
    ''' 修改訓練計畫
    ''' </summary>
    ''' <param name="AccountVal"></param>
    ''' <param name="i_opt">1:  '取任一筆RID 設定 CreateByAcc='Y' /2: 指定  CreateByAcc='Y'</param>
    ''' <param name="RIDVal"></param>
    ''' <param name="planid"></param>
    Sub UPDATE_Auth_AccRWPlan_CreateByAcc(ByVal AccountVal As String, ByVal i_opt As Int16, ByVal RIDVal As String, ByVal planid As String)
        'Optional ByVal RIDVal As String = "", Optional ByVal planid As String = ""
        Select Case i_opt
            Case 1
                '取任一筆RID 設定 CreateByAcc='Y' 依 AccountVal
                Dim pms1 As New Hashtable From {{"Account", AccountVal}}
                Dim sql As String = "select * from Auth_AccRWPlan where Account=@Account"
                Dim objrow As DataRow = DbAccess.GetOneRow(sql, objconn, pms1) 'Account PlanID RID
                If objrow Is Nothing Then Return

                Dim pms_u As New Hashtable From {{"Account", objrow("Account")}, {"PlanID", objrow("PlanID")}, {"RID", objrow("RID")}}
                Dim u_sql As String = "update Auth_AccRWPlan set CreateByAcc='Y' WHERE Account=@Account AND PlanID=@PlanID AND RID=@RID"
                DbAccess.ExecuteNonQuery(u_sql, objconn, pms_u)
            Case 2
                '指定  CreateByAcc='Y' 依 AccountVal,RIDVal,planid
                Dim pms_u As New Hashtable From {{"Account", AccountVal}, {"PlanID", planid}, {"RID", RIDVal}}
                Dim u_sql As String = "update Auth_AccRWPlan set CreateByAcc='Y' WHERE Account=@Account AND PlanID=@PlanID AND RID=@RID"
                DbAccess.ExecuteNonQuery(u_sql, objconn, pms_u)
        End Select
    End Sub

    '刪除訓練計畫
    Private Sub DeleteCmd(ByVal AccountVal As String, ByVal PlanIDval As String, ByVal RID As String)
        Dim Createbyacc As String
        Dim objstr As String
        Dim objtable As DataTable
        Dim objrow As DataRow
        'Dim objadapter As SqlDataAdapter
        AccountVal = TIMS.ClearSQM(AccountVal)
        PlanIDval = TIMS.ClearSQM(PlanIDval)
        RID = TIMS.ClearSQM(RID)

        '刪除目標 SELECT 
        objstr = "select CreateByAcc from Auth_AccRWPlan where Account='" & AccountVal & "' and PlanID = " + PlanIDval & " and RID='" & RID & "'"
        Createbyacc = DbAccess.ExecuteScalar(objstr, objconn)

        '刪除指令啟動
        objstr = "Delete Auth_AccRWPlan where Account='" & AccountVal & "' and PlanID = " & PlanIDval & " and RID='" & RID & "'"
        DbAccess.ExecuteNonQuery(objstr, objconn)

        If Createbyacc = "N" Then
            '查詢主要權限是否存在 CreateByAcc='Y'
            objstr = "Select count(1) cnt from Auth_AccRWPlan where CreateByAcc='Y' and Account='" & AccountVal & "' and RID like '" & RID.Chars(0) & "%'"

            If DbAccess.ExecuteScalar(objstr, objconn) = 0 Then
                Dim objadapter As SqlDataAdapter = Nothing
                '增加1筆主要權限 CreateByAcc='Y'
                'objstr = "SELECT * FROM Auth_AccRWPlan where Account='" & AccountVal & "' and RID like '" & RID.Chars(0) & "%' AND ROWNUM<=1"
                objstr = "SELECT TOP 1 * FROM Auth_AccRWPlan where Account='" & AccountVal & "' and RID like '" & RID.Chars(0) & "%'"
                objtable = DbAccess.GetDataTable(objstr, objadapter, objconn)

                If objtable.Rows.Count <> 0 Then
                    objrow = objtable.Rows(0)
                    objrow("CreateByAcc") = "Y"
                    DbAccess.UpdateDataTable(objtable, objadapter)
                End If
            End If

        ElseIf Createbyacc = "Y" Then
            Dim objadapter As SqlDataAdapter = Nothing
            '增加1筆主要權限 CreateByAcc='Y'
            'objstr = "SELECT * FROM Auth_AccRWPlan where Account='" & AccountVal & "' and RID like '" & RID.Chars(0) & "%' and ROWNUM<=1"
            objstr = "SELECT TOP 1 * FROM Auth_AccRWPlan where Account='" & AccountVal & "' and RID like '" & RID.Chars(0) & "%'"
            objtable = DbAccess.GetDataTable(objstr, objadapter, objconn)

            If objtable.Rows.Count <> 0 Then
                objrow = objtable.Rows(0)
                objrow("CreateByAcc") = "Y"
                DbAccess.UpdateDataTable(objtable, objadapter)
            End If
        End If
    End Sub

    'DataGrid1 刪除命令使用用 (可否刪除後增加1筆主要權限)
    Public Sub CL_Delete(ByVal sender As System.Object, ByVal e As DataGridCommandEventArgs)
        Dim Createbyacc As String = ""
        Dim objstr As String = ""
        Dim objtable As DataTable = Nothing
        Dim objrow As DataRow = Nothing
        Dim objadapter As SqlDataAdapter = Nothing

        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
        'Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)

        Const Cst_CellsRID As Integer = 3
        '刪除目標 SELECT 
        objstr = ""
        objstr += " select CreateByAcc from Auth_AccRWPlan"
        objstr += " where Account='" & v_Lis_acc & "'"
        objstr += " and PlanID = '" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'"
        objstr += " and RID='" & e.Item.Cells(Cst_CellsRID).Text & "'"
        Createbyacc = DbAccess.ExecuteScalar(objstr, objconn)

        '刪除指令啟動
        objstr = ""
        objstr += " Delete Auth_AccRWPlan "
        objstr += " where Account='" & v_Lis_acc & "'"
        objstr += " and PlanID = '" & DataGrid1.DataKeys(e.Item.ItemIndex) & "'"
        objstr += " and RID='" & e.Item.Cells(Cst_CellsRID).Text & "'"
        DbAccess.ExecuteNonQuery(objstr, objconn)

        If Createbyacc = "N" Then
            '查詢主要權限是否存在 CreateByAcc='Y'
            objstr = "Select count(1) cnt from Auth_AccRWPlan where CreateByAcc='Y' and Account='" & v_Lis_acc & "' and RID like '" & e.Item.Cells(3).Text.Chars(0) & "%'"
            If DbAccess.ExecuteScalar(objstr, objconn) = 0 Then
                '增加1筆主要權限 CreateByAcc='Y'
                'top 1 AND ROWNUM<=1
                'objstr = "Select * from Auth_AccRWPlan where Account='" & Me.v_Lis_acc & "' and RID like '" & e.Item.Cells(3).Text.Chars(0) & "%' AND ROWNUM<=1"
                objstr = "Select TOP 1 * from Auth_AccRWPlan where Account='" & v_Lis_acc & "' and RID like '" & e.Item.Cells(3).Text.Chars(0) & "%'"
                objtable = DbAccess.GetDataTable(objstr, objadapter, objconn)
                If objtable.Rows.Count <> 0 Then
                    objrow = objtable.Rows(0)
                    objrow("CreateByAcc") = "Y"
                    DbAccess.UpdateDataTable(objtable, objadapter)
                End If
            End If
        ElseIf Createbyacc = "Y" Then
            '增加1筆主要權限 CreateByAcc='Y'
            'objstr = "Select * from Auth_AccRWPlan where Account='" & Me.v_Lis_acc & "' and RID like '" & e.Item.Cells(3).Text.Chars(0) & "%' AND ROWNUM<=1"
            objstr = "Select TOP 1 * from Auth_AccRWPlan where Account='" & v_Lis_acc & "' and RID like '" & e.Item.Cells(3).Text.Chars(0) & "%'"
            objtable = DbAccess.GetDataTable(objstr, objadapter, objconn)
            If objtable.Rows.Count <> 0 Then
                objrow = objtable.Rows(0)
                objrow("CreateByAcc") = "Y"
                DbAccess.UpdateDataTable(objtable, objadapter)
            End If
        End If
    End Sub

    '清除資料
    Private Sub ClearShow()
        Datagrid2.Visible = False
        DataGrid1.Visible = False
        bt_save.Visible = False
    End Sub
#End Region

#Region "Function"
    '取得計畫dropdownlist    '取得登入帳號可賦予之計畫
    Private Function getLis_Plan(ByVal MyPage As Page, ByVal YearsVal As String, ByVal Account As String) As DataTable
        Dim dt As DataTable
        Dim sql As String
        Dim dr As DataRow
        Dim RoleID As Integer

        Account = TIMS.ClearSQM(Account)
        YearsVal = TIMS.ClearSQM(YearsVal)

        '選擇賦予的使用者 
        sql = "select * from auth_account where account = '" & Account & "'"
        dr = DbAccess.GetOneRow(sql, objconn)

        RoleID = dr("RoleID") '角色代碼 0：超級使用者 1：系統管理者 2：一級以上 3：一級 4：二級 5：承辦人 99：一般使用者
        'Me.ViewState("LID") = dr("LID") '階層代碼 0：署(局) 1：分署(中心) 2：委訓(縣市政府、一般培訓單位)

        If sm.UserInfo.RoleID <= 1 Then  '假如登入者角色代碼為系統管理者(或超級使用者)
            sql = "SELECT * FROM VIEW_LOGINPLAN WHERE DistID='" & sm.UserInfo.DistID & "' and Years='" & YearsVal & "'"
        Else
            Select Case RoleID '角色代碼 1：系統管理者
                Case 1 '選擇賦予的使用者 為系統管理者 ，依登入者轄區顯示計劃
                    sql = "SELECT * FROM VIEW_LOGINPLAN WHERE DistID='" & sm.UserInfo.DistID & "' and Years='" & YearsVal & "'"

                Case Else '選擇賦予的使用者 不為系統管理者，依登入者計畫顯示計畫
                    sql = "SELECT * FROM VIEW_LOGINPLAN WHERE PlanID='" & sm.UserInfo.PlanID & "' and Years='" & YearsVal & "'"
            End Select
        End If
        dt = DbAccess.GetDataTable(sql, objconn)
        Return dt
    End Function

    '取得所選取帳號目前已賦予之計畫
    Private Function getDT(ByVal YearsVal As String, ByVal AccountVal As String) As DataTable
        'Me.Years.SelectedValue
        'Me.v_Lis_acc
        'Dim objAdapter As SqlDataAdapter
        Dim da As New SqlDataAdapter
        Dim dt As New DataTable
        Dim sql As String = ""

        Try
            da.SelectCommand = New SqlCommand
            da.SelectCommand.Connection = objconn

            'sql += " SELECT a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
            'sql += " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
            'sql += " ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')'" & vbCrLf
            'sql += " END" & vbCrLf
            'sql += " +case when f.Clsyear is null or f.Clsyear > b.Years then '' else '…已停用'+ convert(varchar,Clsyear) end AS PlanName" & vbCrLf
            'sql += " ,d.OrgID, d.OrgName, ac.isused, c.DistID, c.OrgLevel" & vbCrLf
            'sql += " ,a2.RID CRID,o2.OrgID COrgID, o2.OrgName COrgName" & vbCrLf
            'sql += " FROM" & vbCrLf
            'sql += " (SELECT * FROM Auth_AccRWPlan WHERE account='" & AccountVal & "') a" & vbCrLf
            'sql += "   JOIN ID_Plan b on a.PlanID=b.PlanID" & vbCrLf
            'sql += "   JOIN ( select * , CASE" & vbCrLf
            'sql += "        WHEN len(Relship)-4 >0" & vbCrLf
            'sql += "        THEN replace(replace(substring(Relship,5,len(Relship)-4),RID,''),'/','')" & vbCrLf
            'sql += "        END AS CRID from Auth_Relship) c on a.RID=c.RID " & vbCrLf
            'sql += "   JOIN Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
            'sql += "   JOIN ID_District e on b.DistID=e.DistID" & vbCrLf
            'sql += "   JOIN Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
            'sql += "   JOIN auth_account ac on a.account=ac.account" & vbCrLf
            'sql += "   left join Auth_Relship a2 on a2.RID=c.CRID" & vbCrLf
            'sql += "   left join Org_OrgInfo o2 on a2.OrgID=o2.OrgID" & vbCrLf
            'sql += " WHERE b.Years='" & YearsVal & "'" & vbCrLf


            sql = ""
            sql &= " SELECT a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
            sql &= " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
            sql &= " ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')'" & vbCrLf
            sql &= " END" & vbCrLf
            sql &= " +case when f.Clsyear is null or f.Clsyear > b.Years then ' ' else '…已停用'+CONVERT(varchar, f.Clsyear) end AS PlanName" & vbCrLf
            sql &= " ,d.OrgID, d.OrgName, ac.isused, c.DistID, c.OrgLevel" & vbCrLf
            sql &= " ,r3.RID2 CRID,o2.OrgID COrgID, o2.OrgName COrgName" & vbCrLf
            sql &= " FROM Auth_AccRWPlan a" & vbCrLf
            sql &= " join ID_Plan b on a.PlanID=b.PlanID and a.account='" & AccountVal & "'" & vbCrLf
            sql &= " join auth_relship c on a.RID=c.RID" & vbCrLf
            sql &= " join Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
            sql &= " join ID_District e on b.DistID=e.DistID" & vbCrLf
            sql &= " join Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
            sql &= " join auth_account ac on a.account=ac.account" & vbCrLf
            sql &= " left join view_Relship23x r3 on r3.distid =e.distid and r3.planid=b.planid and r3.rid3=c.rid" & vbCrLf
            sql &= " left join Org_OrgInfo o2 on r3.OrgID2=o2.OrgID" & vbCrLf
            sql &= " WHERE b.Years='" & YearsVal & "'" & vbCrLf
            da.SelectCommand.CommandText = sql
            da.Fill(dt)

            If Not da Is Nothing Then da.Dispose()
            'If Not dt Is Nothing Then dt.Dispose()
        Catch ex As Exception
            Throw ex
            Me.Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' 取得計畫筆數 '計算該帳號目前所有已賦予權限的計畫
    ''' </summary>
    ''' <param name="Account"></param>
    ''' <param name="i_opt"></param>
    ''' <param name="years"></param>
    ''' <returns></returns>
    Function get_PlanCount(ByVal Account As String, ByVal i_opt As Int16, ByVal years As String) As Integer
        'Optional ByVal i_opt As Int16 = 0, Optional ByVal years As String = ""
        Dim iRst As Integer = 0
        Account = TIMS.ClearSQM(Account)
        years = TIMS.ClearSQM(years)

        Dim sql As String = ""
        Select Case i_opt
            Case 0 '確認計畫數
                sql = "SELECT COUNT(1) CNT FROM Auth_AccRWPlan WHERE Account='" & Account & "'"
            Case 3 '是否有 CreateByAcc='Y' 
                sql = "SELECT COUNT(1) CNT FROM Auth_AccRWPlan WHERE Account='" & Account & "' AND CreateByAcc='Y' "
            Case 1 '確認計畫數 依年度
                sql = ""
                sql &= " select count(1) cnt" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= " 	select a.account ,b.years" & vbCrLf
                sql &= " 	from Auth_AccRWPlan a" & vbCrLf
                sql &= " 	join id_plan b on a.planid=b.planid" & vbCrLf
                sql &= " 	where a.account='" & Account & "'" & vbCrLf
                sql &= " 	and b.years in ('" & years & "')" & vbCrLf
                sql &= " ) a" & vbCrLf
            Case 2 '確認計畫數 依(其它)年度
                sql = ""
                sql &= " select count(1) cnt" & vbCrLf
                sql &= " from (" & vbCrLf
                sql &= " 	select a.account ,b.years" & vbCrLf
                sql &= " 	FROM AUTH_ACCRWPLAN a" & vbCrLf
                sql &= " 	JOIN ID_PLAN b on a.planid=b.planid" & vbCrLf
                sql &= " 	where a.account='" & Account & "'" & vbCrLf
                sql &= " 	and b.years not in ('" & years & "')" & vbCrLf
                sql &= " ) a" & vbCrLf
        End Select
        iRst = DbAccess.ExecuteScalar(sql, objconn)
        Return iRst 'DbAccess.ExecuteScalar(sql, objconn)
    End Function

#End Region

#Region "NO USE"
    'Datagrid2 = Nothing
    'DataGrid1 = Nothing
    'Datagrid2.DataSource = New DataTable
    'Datagrid2.DataBind()
    'DataGrid1.DataSource = New DataTable
    'DataGrid1.DataBind()

    'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "but_search")
    'If Not HistoryRID Is Nothing Then
    '    If HistoryRID.Rows.Count <> 0 Then
    '        center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
    '        HistoryRID.Attributes("onclick") = "ShowFrame();"
    '        center.Style("CURSOR") = "hand"
    '    End If
    'End If

    'If TIMS.Check_PlanCount(Me.v_Lis_acc, dr("RID")) > 1 Then
    '    'chk1.ToolTip = "刪除計劃時請謹慎"
    '    vtitle = "該業務機構尚有其他計畫代碼 擁有該控制權限,刪除該計劃時請謹慎"
    '    TIMS.Tooltip(chk1, vtitle)
    'End If

    'If dr("isused") = "N" Then
    '    'chk1.Attributes.Add("onclick", "return confirm('目前此帳號為【停用帳號】強制刪除計劃時請謹慎!');")
    '    'bt_save.ToolTip = "目前此帳號為【停用帳號】強制刪除計劃時請謹慎!"
    '    vtitle = "目前此帳號為【停用帳號】強制刪除計劃時請謹慎!"
    '    chk1.Attributes.Add("onclick", "return confirm('" & vtitle & "');")
    '    TIMS.Tooltip(bt_save, vtitle)
    'End If


#End Region

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
        'TIMS.TestDbConn(Me, objconn, True)
        '檢查Session是否存在 End
        'conn = DbAccess.GetConnection()
        'If conn.State = ConnectionState.Closed Then conn.Open()
        Me.but_search.Style("display") = "none"
        HidDefInput.Value = Cst_DefInput

        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
            str_superuser1 = CStr(sm.UserInfo.UserID)
        End If

        Hid_RoleID.Value = sm.UserInfo.RoleID
        'submit=='true' 將觸動Click做but_search
        If sm.UserInfo.RID <> "A" AndAlso sm.UserInfo.RoleID = "1" Then
            '分署(中心)？
            btu_org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?submit=true&GetOther=1'+'&OrgField=center&fisBlack=isBlack')"
        ElseIf sm.UserInfo.RID = "A" Then
            '署(局)
            btu_org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?submit=true'+'&OrgField=center&fisBlack=isBlack')"
        Else
            '使用者
            btu_org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx?btnName=but_search'+'&OrgField=center&fisBlack=isBlack')"
        End If

        '隸屬機構選擇
        Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?PlanID='+document.form1.Lis_Plan.value+'&OrgField=OrgName&fisBlack=isBlack2')"
        bt_save.Attributes("onclick") = "return confirm('確定儲存?');"

        If Not IsPostBack Then
            trView.Visible = False
            ShowTable.Style.Item("display") = "none"

            If Request("RID") <> "" AndAlso Request("AN") <> "" Then
                Dim sql As String
                Dim dr As DataRow

                sql = "SELECT OrgName, ComIDNO FROM Org_OrgInfo WHERE OrgID=(SELECT OrgID FROM Auth_Relship WHERE RID='" & Request("RID") & "')"
                dr = DbAccess.GetOneRow(sql, objconn)

                If Not dr Is Nothing Then
                    center.Text = dr("OrgName").ToString
                    '確認登入帳號之機構是否在黑名單中
                    If TIMS.Check_OrgBlackList(Me, Convert.ToString(dr("ComIDNO")), objconn) Then
                        Me.isBlack.Value = "Y"
                    End If
                End If

                RIDValue.Value = Request("RID")
                Call sSearch1() 'but_search_Click(sender, e)

                Common.SetListItem(Lis_acc, Request("AN"))
                Lis_acc_SelectedIndexChanged(sender, e)
                Common.SetListItem(Years, Request("Years"))
                Years_SelectedIndexChanged(sender, e)
            Else
                Me.ViewState("ComIDNO") = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)

                '確認登入帳號之機構是否在黑名單中
                If TIMS.Check_OrgBlackList(Me, Me.ViewState("ComIDNO"), objconn) Then
                    Me.isBlack.Value = "Y"
                End If

                center.Text = sm.UserInfo.OrgName
                RIDValue.Value = sm.UserInfo.RID
                Call sSearch1() 'but_search_Click(sender, e)
            End If

            txtSchAccount.Text = Cst_DefInput ' "請輸入姓名關鍵字"
            txtSchAccount.Style.Add("color", "#858585")
            txtSchAccount.Attributes.Add("onclick", "chkName();")
            txtSchAccount.Attributes.Add("onblur", "chkName();")
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    If sm.UserInfo.RoleID <> 0 Then
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '           'Dim FunDr As DataRow = FunDrArray(0)
        '            If FunDr("Adds") = "1" Then
        '                btu_add.Enabled = True
        '            Else
        '                btu_add.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

        '因隸屬機構計畫賦予會有問題,暫改為不顯示隸屬機構 
        OrgTable.Style.Item("visibility") = "hidden"
        btu_add.Style.Item("visibility") = "hidden"
        Lis_Plan.Style.Item("visibility") = "hidden"
        butRefer.Attributes("onclick") = "return openrefer();"
    End Sub

    '使用者 查詢 Lis_acc
    Sub sSearch1()
        'Dim objtable As DataTable
        'Dim sql As String = ""
        Call ClearShow()
        ShowTable.Style.Item("display") = ""
        Years.Visible = False
        Years.Items.Clear()
        Lis_Plan.Items.Clear()
        Lis_Plan.Items.Add("請選擇帳號")
        OrgName.Text = ""
        OrgRID.Value = ""

        If Convert.ToString(Me.Request("accountname")) <> "" Then
            Common.SetListItem(Me.Lis_acc, Convert.ToString(Me.Request("accountname")))
            'Me.v_Lis_acc = Convert.ToString(Me.Request("accountname"))
        End If

        '造成原因為使用 ShowHistoryRID ，但卻沒有值，而產生此問題
        If RIDValue.Value = "" Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        '修改可直接賦予計畫
        Dim sql As String = ""
        sql &= " select a.Account" & vbCrLf
        sql &= " ,a.RoleID" & vbCrLf
        sql &= " ,a.Name+'('+b.Name+') ['+a.Account+']' Name" & vbCrLf
        sql &= " ,'1' flag" & vbCrLf
        sql &= " FROM AUTH_ACCOUNT a" & vbCrLf
        sql &= " join ID_Role b on a.RoleID=b.RoleID" & vbCrLf
        sql &= " join Auth_Relship c on c.OrgID=a.OrgID" & vbCrLf

        '跨轄區選擇時 (含委訓單位Left(,1))
        If Left(Me.RIDValue.Value, 1) <> sm.UserInfo.RID Then
            If Convert.ToString(sm.UserInfo.DistID) = "000" AndAlso sm.UserInfo.UserID = str_superuser1 Then
                '(署(局)" & cst_supermaster & "登入)  
                'objstr += "join Auth_AccRwDist rd on rd.Account=a.Account" & vbCrLf
            Else
                '依登入者使用轄區權限賦予計畫
                sql &= " join Auth_AccRwDist rd on rd.Account=a.Account AND rd.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            End If
        End If
        sql &= " WHERE c.RID='" & Me.RIDValue.Value & "'" & vbCrLf
        sql &= " AND a.RoleID >=" & sm.UserInfo.RoleID & vbCrLf

        Dim v_rdoIsUsed As String = TIMS.GetListValue(rdoIsUsed)
        If sm.UserInfo.RID <> RIDValue.Value.Chars(0) Then
            If sm.UserInfo.RoleID = "0" Then '超級使用者啦
                Select Case v_rdoIsUsed 'rdoIsUsed.SelectedValue
                    Case "A"
                    Case "Y"
                        sql &= " and a.IsUsed='Y'" & vbCrLf
                    Case Else
                        sql &= " and a.IsUsed='N'" & vbCrLf
                End Select
            Else
                Select Case v_rdoIsUsed 'rdoIsUsed.SelectedValue '非超級使用者(跨區了)
                    Case "A"
                        sql &= " and a.RoleID not in (0,1)" & vbCrLf
                    Case "Y"
                        sql &= " and a.RoleID not in (0,1) and a.IsUsed='Y'" & vbCrLf
                    Case Else
                        sql &= " and a.RoleID not in (0,1) and a.IsUsed='N'" & vbCrLf
                End Select
            End If

            OrgTable.Style.Item("display") = "inline" '顯示隸屬機構 
            btu_add.Attributes("onclick") = "return chkaccount(2);"
        Else
            '限定不可為超級使用者
            Select Case v_rdoIsUsed 'rdoIsUsed.SelectedValue
                Case "A"
                    sql &= " and a.RoleID<>0" & vbCrLf
                Case "Y"
                    sql &= " and a.RoleID<>0 and a.IsUsed='Y'" & vbCrLf
                Case Else
                    sql &= " and a.RoleID<>0 and a.IsUsed='N'" & vbCrLf
            End Select
            OrgTable.Style.Item("display") = "none" '不顯示隸屬機構
            btu_add.Attributes("onclick") = "return chkaccount(1);"
        End If

        'super user snoopy " & cst_supermaster & " 帳號清單中要顯示snoopy,除此之外,任何一個帳號登入至此功能,都不得顯示snoopy帳號
        If sm.UserInfo.UserID = str_superuser1 Then
            sql &= " or a.account='" & str_superuser1 & "'" & vbCrLf
        Else
            sql &= " and a.account<>'" & str_superuser1 & "'" & vbCrLf
        End If

        '跨轄區選擇時 (含委訓單位Left(,1))
        If Left(Me.RIDValue.Value, 1) = sm.UserInfo.RID Then
            sql &= " union" & vbCrLf

            sql &= " select a.Account,a.RoleID" & vbCrLf
            sql &= " ,a.Name+'('+b.Name+') ['+a.Account+']'" & vbCrLf
            sql &= " + '[' + CASE (c.RID)" & vbCrLf
            sql &= " when 'A' then '署'" & vbCrLf
            sql &= " when 'B' then '北基宜花金馬分署'" & vbCrLf
            sql &= " when 'C' then '泰山訓練場'" & vbCrLf
            sql &= " when 'D' then '桃竹苗分署'" & vbCrLf
            sql &= " when 'E' then '中彰投分署'" & vbCrLf
            sql &= " when 'F' then '雲嘉南分署'" & vbCrLf
            sql &= " when 'G' then '高屏澎東分署'" & vbCrLf
            sql &= " end + ']' Name" & vbCrLf
            'sql += " when 'A' then '局'" & vbCrLf
            'sql += " when 'B' then '北區'" & vbCrLf
            'sql += " when 'C' then '泰山'" & vbCrLf
            'sql += " when 'D' then '桃園'" & vbCrLf
            'sql += " when 'E' then '中區'" & vbCrLf
            'sql += " when 'F' then '台南'" & vbCrLf
            'sql += " when 'G' then '南區'" & vbCrLf
            sql &= " ,'2' flag" & vbCrLf
            sql &= " FROM AUTH_ACCOUNT a" & vbCrLf
            sql &= " join ID_Role b on a.RoleID=b.RoleID" & vbCrLf
            sql &= " join Auth_Relship c on c.OrgID=a.OrgID" & vbCrLf
            sql &= " join Auth_AccRwDist rd on rd.Account=a.Account AND rd.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            sql &= " where a.RoleID >=" & sm.UserInfo.RoleID & vbCrLf

            Select Case v_rdoIsUsed 'rdoIsUsed.SelectedValue '非超級使用者(跨區了)
                Case "A"
                    sql &= " and a.RoleID not in (0,1)" & vbCrLf
                Case "Y"
                    sql &= " and a.RoleID not in (0,1) and a.IsUsed='Y'" & vbCrLf
                Case Else
                    sql &= " and a.RoleID not in (0,1) and a.IsUsed='N'" & vbCrLf
            End Select

        End If
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        'objtable = DbAccess.GetDataTable(sql, objconn)
        'objtable.DefaultView.Sort = "flag,RoleID,Name"
        dt.DefaultView.Sort = "flag,RoleID,Name"
        With Me.Lis_acc
            .DataSource = dt.DefaultView ' objtable.DefaultView
            .DataTextField = "name"
            .DataValueField = "account"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        DataGrid1.Visible = False
    End Sub

    '搜尋 帳號 姓名關鍵字
    Sub KeywordSch1()
        Dim strName As String = ""
        Dim bolCheck As Boolean = False
        If txtSchAccount.Text <> "" Then txtSchAccount.Text = Trim(txtSchAccount.Text)
        txtSchAccount.Text = TIMS.ClearSQM(txtSchAccount.Text)
        '不為空值&預設定時,看DropDownList是否有查到關鍵字
        If txtSchAccount.Text <> Cst_DefInput Then
            For i As Integer = 0 To Lis_acc.Items.Count - 1
                If i > 0 Then
                    '取得姓名/'取得帳號
                    strName = Mid(Lis_acc.Items(i).Text, 1, Lis_acc.Items(i).Text.IndexOf("("))
                    strName &= Mid(Lis_acc.Items(i).Text, Lis_acc.Items(i).Text.IndexOf("[") + 2, Lis_acc.Items(i).Text.IndexOf("]") - Lis_acc.Items(i).Text.IndexOf("[") - 1)
                    If strName.IndexOf(txtSchAccount.Text) <> -1 Then
                        bolCheck = True
                        Exit For
                    End If
                End If
            Next

            '如有關鍵字時,開始刪非關鍵字item(若無,則不動)
            If bolCheck = True Then
                For i As Integer = Lis_acc.Items.Count - 1 To 0 Step -1
                    If i > 0 Then
                        strName = Mid(Lis_acc.Items(i).Text, 1, Lis_acc.Items(i).Text.IndexOf("("))
                        strName &= Mid(Lis_acc.Items(i).Text, Lis_acc.Items(i).Text.IndexOf("[") + 2, Lis_acc.Items(i).Text.IndexOf("]") - Lis_acc.Items(i).Text.IndexOf("[") - 1)
                        If strName.IndexOf(txtSchAccount.Text) = -1 Then
                            Lis_acc.Items.RemoveAt(i)
                        End If
                    End If
                Next
            End If
        End If
    End Sub

    '使用者 查詢 Lis_acc
    Private Sub but_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_search.Click
        Call sSearch1()
    End Sub

    '搜尋 帳號 姓名關鍵字
    Private Sub btnSchAccount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSchAccount.Click
        Call sSearch1() ' but_search_Click(sender, e)
        Call KeywordSch1()

        'txtSchAccount.Style.Remove("color")
        txtSchAccount.Style.Add("color", "#858585") '提示style
        If txtSchAccount.Text <> Cst_DefInput Then
            'txtSchAccount.Style.Remove("color")
            txtSchAccount.Style.Add("color", "black") '實體style
        End If
    End Sub

    Private Sub btu_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btu_add.Click
        'Dim Errmsg As String = ""
        Dim objAdapter As SqlDataAdapter = Nothing
        Dim objtable As DataTable = Nothing
        Dim objstr As String = ""
        Dim sql As String = ""
        Dim count As Integer = 0

        'v_Lis_acc
        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
        Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)

        Dim Errmsg As String = ""
        If sm.UserInfo.RID <> RIDValue.Value.Chars(0) Then '新增時 確認登入者的RID與賦予者的RID 轄區是否相同?
            '不相同 '依 隸屬機構RID確認是否有此計畫
            sql = "select count(1) from Auth_AccRWPlan where account='" & v_Lis_acc & "' and RID like '" & Me.OrgRID.Value.Chars(0) & "%'"
            objstr = "select * from Auth_AccRWPlan where account='" & v_Lis_acc & "' and PlanID=" & v_Lis_Plan & " and RID='" & Me.OrgRID.Value & "'"
        Else
            '相同 '依 賦予者機構RID確認是否有此計畫
            sql = "select count(1) from Auth_AccRWPlan where account='" & v_Lis_acc & "' and RID='" & Me.RIDValue.Value & "'"
            objstr = "select * from Auth_AccRWPlan where account='" & v_Lis_acc & "' and PlanID=" & v_Lis_Plan & " and RID='" & Me.RIDValue.Value & "'"
        End If

        count = DbAccess.ExecuteScalar(sql, objconn) '依轄區RID計算出使用者目前的賦予的權限數

        objtable = DbAccess.GetDataTable(objstr, objAdapter, objconn)

        If objtable.Rows.Count < 1 Then
            '權限儲存
            SAVE_Auth_AccRWPlan(count, objtable, objAdapter, Errmsg)

            If Errmsg <> "" Then
                Common.MessageBox(Page, Errmsg)
            Else
                Common.MessageBox(Page, "新增成功!!!")
            End If
        Else
            Common.MessageBox(Page, "已經新增，不可重複!!!")
        End If

        Call getRelishTB()
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem '
                Dim chk1 As CheckBox = e.Item.FindControl("chk1")
                Dim Hid_RID As HiddenField = e.Item.FindControl("Hid_RID")
                Dim Hid_planid As HiddenField = e.Item.FindControl("Hid_planid")
                Dim Hid_tplanid As HiddenField = e.Item.FindControl("Hid_tplanid")
                Hid_RID.Value = Convert.ToString(drv("RID"))
                Hid_planid.Value = Convert.ToString(drv("planid"))
                Hid_tplanid.Value = Convert.ToString(drv("tplanid"))

                Dim vtitle As String = ""
                '分署(中心)登入的話所有計畫都可以自己加,【分署(中心)以下要卡】
                If Convert.ToString(sm.UserInfo.RID).Length > 1 Then
                    chk1.Enabled = False

                    If Not dtG2 Is Nothing Then
                        For Each dr2 As DataRow In dtG2.Rows
                            '有在登入帳號可賦予之計畫才開放使用
                            If dr2("planid") = drv("planid") Then
                                chk1.Enabled = True
                                Exit For '迴圈目的達成離開
                            End If
                        Next
                    End If
                Else
                    If Not dtG1 Is Nothing Then
                        For Each dr As DataRow In dtG1.Rows

                            If Convert.ToString(dr("RID")) = Convert.ToString(drv("RID")) _
                                AndAlso Convert.ToString(dr("planid")) = Convert.ToString(drv("planid")) Then
                                chk1.Checked = True
                                vtitle = "含有該年度計畫 業務機構之權限!" & vbCrLf
                                TIMS.Tooltip(chk1, vtitle)
                            End If

                            If chk1.Checked Then
                                Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
                                'Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)
                                If get_PlanCount(v_Lis_acc, 0, "") <= 1 Then
                                    chk1.Enabled = False
                                    vtitle = "此帳號至少須保留一個計劃賦予之計劃,若有需要刪除,請先新增其他計劃即可刪除此計劃 !" & vbCrLf
                                    TIMS.Tooltip(chk1, vtitle)
                                End If
                                Exit For '迴圈目的達成離開
                            End If
                        Next
                    End If
                End If

                '補助地方政府
                If Convert.ToString(drv("OrgName2")) <> "" Then
                    vtitle = "補助地方政府單位：" & Convert.ToString(drv("OrgName2")) & vbCrLf
                    TIMS.Tooltip(e.Item.Cells(1), vtitle)
                End If
                '補助地方政府的委訓單位
                If Convert.ToString(drv("OrgName3")) <> "" Then
                    vtitle = "補助地方政府的委訓單位：" & Convert.ToString(drv("OrgName3")) & vbCrLf
                    TIMS.Tooltip(e.Item.Cells(1), vtitle)
                End If

                '署(局)帳號登入有最大權限
                If sm.UserInfo.RID <> "A" Then
                    '不同轄區 不可互相執行計畫賦予
                    If sm.UserInfo.RID <> Convert.ToString(drv("RID")).Chars(0) Then
                        chk1.Enabled = False
                        vtitle = "登入者與賦予權限者的業務機構不相同!" & Convert.ToString(drv("RID")) & vbCrLf
                        TIMS.Tooltip(e.Item.Cells(1), vtitle)
                        TIMS.Tooltip(chk1, vtitle)
                    Else
                        vtitle = "登入者與賦予權限者的業務機構相同" & Convert.ToString(drv("RID")) & vbCrLf
                        TIMS.Tooltip(e.Item.Cells(1), vtitle)
                        TIMS.Tooltip(chk1, vtitle)
                    End If
                    If drv("DistID") <> sm.UserInfo.DistID Then
                        chk1.Enabled = False
                        vtitle = "登入者與賦予權限者的轄區不相同!" & vbCrLf
                        TIMS.Tooltip(e.Item.Cells(1), vtitle)
                        TIMS.Tooltip(chk1, vtitle)
                    Else
                        vtitle = "登入者與賦予權限者的轄區相同" & vbCrLf
                        TIMS.Tooltip(e.Item.Cells(1), vtitle)
                        TIMS.Tooltip(chk1, vtitle)
                    End If
                Else
                    'vtitle = "局登入者目前擁有該控制權限!" & vbCrLf
                    vtitle = "署登入者目前擁有該控制權限!" & vbCrLf
                    TIMS.Tooltip(chk1, vtitle)
                End If

                '顯示清單控制
                If chk1.Checked Then lsbShow.Items.Add(Convert.ToString(drv("PlanName")))
                chk1.Attributes.Add("onclick", "showList();")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "Att" '委訓單位歸屬
                TIMS.Utl_Redirect1(Me, e.CommandArgument)

            Case "AttB" '委訓單位歸屬全選
                ViewState("RoleID") = TIMS.GetMyValue(e.CommandArgument, "RoleID")
                ViewState("PlanID") = TIMS.GetMyValue(e.CommandArgument, "PlanID")
                ViewState("Account") = TIMS.GetMyValue(e.CommandArgument, "Account")
                ViewState("RIDValue") = TIMS.GetMyValue(e.CommandArgument, "RIDValue")
                'TIMS.Update_AUTH_ACCTORG(Me, ViewState("RoleID"), ViewState("PlanID"), ViewState("Account"), ViewState("RIDValue"))
                Dim iRoleID As Integer = Val(ViewState("RoleID"))
                Dim iPlanID As Integer = Val(ViewState("PlanID"))
                Dim sAccount As String = CStr(ViewState("Account"))
                Dim sRIDValue As String = CStr(ViewState("RIDValue"))
                '委訓單位歸屬全選
                Call TIMS.Update_AUTH_ACCTORG(iRoleID, iPlanID, sAccount, sRIDValue, objconn)
                Common.MessageBox(Me, "歸屬完成!!!")
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const cst_計畫代碼 As Integer = 0
        Const cst_訓練機構 As Integer = 1
        Const cst_功能 As Integer = 2

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim strRoleId As String = "" 'TIMS.Get_RoleID(Me.v_Lis_acc)
                strRoleId = Me.hidRoleId.Value
                Dim delflag As Integer = 0
                delflag = 0
                'Dim Label1 As Label
                Dim But_Att As Button = e.Item.Cells(cst_功能).FindControl("But_Att") '委訓單位歸屬
                Dim But_Del As Button = e.Item.Cells(cst_功能).FindControl("But_Del") '刪除
                Dim But_AttB As Button = e.Item.Cells(cst_功能).FindControl("But_AttB") '委訓單位歸屬全選

                But_Del = e.Item.Cells(cst_功能).FindControl("But_Del")
                But_Del.Attributes("onclick") = "return confirm('確定要刪除這一筆計畫賦予?');"
                But_Del.Style("display") = "none" '2010010203  andy   刪除鈕隱藏
                But_Del.Visible = False '停用
                But_Del.Enabled = False '停用

                TIMS.Tooltip(e.Item.Cells(cst_計畫代碼), drv("TPlanID").ToString)
                TIMS.Tooltip(e.Item.Cells(cst_計畫代碼), drv("PlanID").ToString)

                TIMS.Tooltip(e.Item.Cells(cst_訓練機構), "[" & drv("RID").ToString & "] " & drv("OrgID").ToString)
                TIMS.Tooltip(e.Item.Cells(cst_訓練機構), " CreateByAcc:" & drv("CreateByAcc").ToString)

                If But_Att IsNot Nothing Then
                    Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
                    'Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)

                    But_Att.CommandName = "Att"
                    But_Att.CommandArgument = "SYS_01_002_att.aspx?ID=" & Request("ID") & "&PID=" & Convert.ToString(drv("PlanID")) & "&RID=" & Me.RIDValue.Value & "&ON=" & Me.center.Text & "&Years=" & Years.SelectedValue & "&AN=" & v_Lis_acc & "&RoleID=" & strRoleId

                    But_AttB.CommandName = "AttB"
                    But_AttB.CommandArgument = "BCmd=AttB"
                    But_AttB.CommandArgument += "&RoleID=" & strRoleId
                    But_AttB.CommandArgument += "&PlanID=" & Convert.ToString(drv("PlanID"))
                    But_AttB.CommandArgument += "&Account=" & v_Lis_acc
                    But_AttB.CommandArgument += "&RIDValue=" & Me.RIDValue.Value

                    If sm.UserInfo.RID <> RIDValue.Value.Chars(0) Then
                        If sm.UserInfo.RoleID <> 0 Then
                            But_Att.Visible = False
                        Else
                            Select Case strRoleId
                                Case 4, 5 ' 二級 與 承辦人 可做委訓單位歸屬
                                    But_Att.Visible = True
                                    TIMS.Tooltip(But_Att, "二級 與 承辦人 可做委訓單位歸屬(管理者跨轄區功能開放)")
                                Case Else
                                    But_Att.Visible = False
                            End Select
                        End If
                    Else
                        Select Case strRoleId
                            Case 4, 5 ' 二級 與 承辦人 可做委訓單位歸屬
                                But_Att.Visible = True
                                TIMS.Tooltip(But_Att, "二級 與 承辦人 可做委訓單位歸屬")
                            Case Else
                                But_Att.Visible = False
                        End Select
                    End If
                End If

                But_AttB.Visible = But_Att.Visible 'But_AttB顯示同But_Att
        End Select
    End Sub

    '再檢查一次 Auth_AccRWPlan 裡面是否帳號存在至少一筆 CreateByAcc='Y'的資料
    Sub save_data_c2()
        Dim i_allchk_count2 As Int16 = 0
        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc) '.SelectedValue
        '除了所選年度之外都沒有賦予該帳號計畫權限
        If get_PlanCount(v_Lis_acc, 2, Years.SelectedValue) = 0 Then
            Dim s_RIDval As String = ""
            Dim si_Planid2 As String = ""
            For Each eItem As DataGridItem In Datagrid2.Items
                Dim chk1 As CheckBox = eItem.FindControl("chk1") ' Datagrid2.Items(i).Cells(1).FindControl("chk1")
                Dim Hid_RID As HiddenField = eItem.FindControl("Hid_RID")
                Dim Hid_planid As HiddenField = eItem.FindControl("Hid_planid")
                Dim Hid_tplanid As HiddenField = eItem.FindControl("Hid_tplanid")
                If chk1.Checked Then
                    s_RIDval = Hid_RID.Value
                    si_Planid2 = Hid_planid.Value
                    i_allchk_count2 += 1
                End If
            Next
            'For i As Integer = 0 To Datagrid2.Items.Count - 1
            '    Dim chk1 As CheckBox = Datagrid2.Items(i).Cells(1).FindControl("chk1")
            '    If chk1.Checked = True Then
            '        allchk_count2 = allchk_count2 + 1
            '        RIDval = Datagrid2.Items(i).Cells(2).Text
            '        planid2 = Datagrid2.Items(i).Cells(3).Text
            '    End If
            'Next
            If i_allchk_count2 = 1 Then
                UPDATE_Auth_AccRWPlan_CreateByAcc(v_Lis_acc, 2, s_RIDval, si_Planid2)
            End If
        End If
        ''不管怎樣至少保留一筆CreateByAcc='Y'的資料
        'If get_PlanCount(Me.v_Lis_acc) = 1 Then
        '    UPDATE_Auth_AccRWPlan_CreateByAcc(Me.v_Lis_acc, 1)
        'End If
        '不管怎樣至少保留一筆CreateByAcc='Y'的資料
        If get_PlanCount(v_Lis_acc, 3, "") = 0 Then
            UPDATE_Auth_AccRWPlan_CreateByAcc(v_Lis_acc, 1, "", "")
        End If

    End Sub

    Sub save_data_c1(ByRef dt2 As DataTable)
        Dim CreateByAcc As String = "N" '"N":'由計畫賦予 "Y":'由計畫新增(由帳號建立)

        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
        hidRoleId.Value = TIMS.Get_RoleID(v_Lis_acc, objconn)
        Dim i_ROLEID As Integer = Val(hidRoleId.Value)

        Dim i_parms As New Hashtable
        Dim sqladd As String = ""
        sqladd = ""
        sqladd += " insert into Auth_AccRWPlan (Account,PlanID,RID,CreateByAcc,ModifyAcct,ModifyDate)"
        sqladd += " values (@Account,@PlanID,@RID,@CreateByAcc,@ModifyAcct,getdate())"

        For Each eItem As DataGridItem In Datagrid2.Items
            Dim drv As DataRowView = eItem.DataItem '
            Dim chk1 As CheckBox = eItem.FindControl("chk1")
            Dim Hid_RID As HiddenField = eItem.FindControl("Hid_RID")
            Dim Hid_planid As HiddenField = eItem.FindControl("Hid_planid")
            Dim Hid_tplanid As HiddenField = eItem.FindControl("Hid_tplanid")
            Dim s_RID As String = Hid_RID.Value
            Dim iPlanid As Integer = Val(Hid_planid.Value)

            If chk1.Checked = True Then
                Dim flag_can_save As Boolean = False
                If dt2.Rows.Count > 0 Then
                    If dt2.Select("rid='" & s_RID & "' and PlanID='" & iPlanid.ToString() & "'").Length = 0 Then
                        flag_can_save = True
                        'sqladd += " insert into Auth_AccRWPlan (Account,PlanID,RID,CreateByAcc,ModifyAcct,ModifyDate)"
                        'sqladd += "  values('" & v_Lis_acc & "', '" & planid & "' " 'Me.v_Lis_acc '賦予的帳號  'PlanList.Items(i).Value  '賦予的計畫
                        'sqladd += ",'" & RID & "', '" & CreateByAcc & "' ,"
                        'sqladd += "'" & sm.UserInfo.UserID & "', getdate() )  "
                    End If
                Else
                    flag_can_save = True
                    'sqladd += " insert into Auth_AccRWPlan (Account,PlanID,RID,CreateByAcc,ModifyAcct,ModifyDate)"
                    'sqladd += "  values('" & v_Lis_acc & "', '" & planid & "' "        'Me.v_Lis_acc '賦予的帳號  'PlanList.Items(i).Value  '賦予的計畫
                    'sqladd += ",'" & RID & "', '" & CreateByAcc & "' ,"
                    'sqladd += "'" & sm.UserInfo.UserID & "', getdate() )  "
                End If

                If flag_can_save Then
                    i_parms.Clear()
                    i_parms.Add("Account", v_Lis_acc)
                    i_parms.Add("PlanID", iPlanid)
                    i_parms.Add("RID", s_RID)
                    i_parms.Add("CreateByAcc", CreateByAcc)
                    i_parms.Add("ModifyAcct", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(sqladd, objconn, i_parms)

                    Call TIMS.Update_AUTH_ACCTORG(i_ROLEID, iPlanid, v_Lis_acc, s_RID, objconn)
                End If
            ElseIf chk1.Checked = False Then
                If get_PlanCount(v_Lis_acc, 0, "") > 1 Then
                    '此帳號至少須保留一個計劃賦予之計劃,若有需要刪除<br>請先新增其他計劃即可刪除此計劃 !
                    If dt2.Rows.Count > 0 Then
                        If dt2.Select("rid='" & s_RID & "'  and PlanID='" & iPlanid.ToString() & "'").Length > 0 Then
                            DeleteCmd(v_Lis_acc, iPlanid.ToString(), s_RID)
                        End If
                    End If
                End If
            End If
            'If sqladd <> "" Then
            '    DbAccess.ExecuteNonQuery(sqladd, objconn)
            '    da.SelectCommand = New SqlCommand(sqladd, objconn)
            '    da.SelectCommand.ExecuteNonQuery()
            'End If
        Next
        'For i As Integer = 0 To Datagrid2.Items.Count - 1
        '    Dim RID As String = ""
        '    Dim planid As String = ""
        '    Dim chk1 As CheckBox = Datagrid2.Items(i).Cells(0).FindControl("chk1")
        '    RID = Datagrid2.Items(i).Cells(2).Text
        '    planid = Datagrid2.Items(i).Cells(3).Text
        'Next
    End Sub

    '儲存鈕
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        '*登入帳號權限為--署(局)    則只能賦予局的權限，不可賦予局以下的權限 
        '*登入帳號權限為--分署(中心)以下【含】 則不做限制(所可賦予計畫之帳號必不含局只會有中心以下之權限)
        Dim v_Lis_acc As String = TIMS.GetListValue(Lis_acc)
        'Dim v_Lis_Plan As String = TIMS.GetListValue(Lis_Plan)

        Dim dt2 As DataTable = getDT(Me.Years.SelectedValue, v_Lis_acc)
        'Dim CreateByAcc As String = "N" '"N":'由計畫賦予 "Y":'由計畫新增(由帳號建立)
        'Dim sqladd As String = ""
        'Dim da As New SqlDataAdapter
        'Dim chk_count As Int16 = 0
        Dim i_allchk_count As Int16 = 0
        If get_PlanCount(v_Lis_acc, 2, Years.SelectedValue) = 0 Then     '除了所選年度之外都沒有賦予該帳號計畫權限
            For i As Integer = 0 To Datagrid2.Items.Count - 1
                Dim chk1 As CheckBox = Datagrid2.Items(i).Cells(1).FindControl("chk1")
                If chk1.Checked = True Then i_allchk_count += 1
            Next
            If i_allchk_count < 1 Then
                Me.Page.RegisterStartupScript("Errmsg", "<script>alert('此帳號至少需保留一個計畫賦予！');</script>")
                lsbShow.Items.Clear()
                Exit Sub
            End If
        End If

        If Lis_acc.SelectedIndex = 0 Then
            Me.Page.RegisterStartupScript("Errmsg", "<script>alert('" & "請選取帳號！".Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
            Exit Sub
        End If

        Try

            save_data_c1(dt2)
            'UPDATE_Auth_AccRWPlan_CreateByAcc(Me.v_Lis_acc, Me.RIDValue.Value)
            '再檢查一次 Auth_AccRWPlan 裡面是否帳號存在至少一筆 CreateByAcc='Y'的資料
            save_data_c2()

        Catch ex As Exception
            Me.Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try

        'dsdsfdsdsf
        'Dim iRoleID As Integer = Val(ViewState("RoleID"))
        'Dim iPlanID As Integer = Val(ViewState("PlanID"))
        'Dim sAccount As String = CStr(ViewState("Account"))
        'Dim sRIDValue As String = CStr(ViewState("RIDValue"))
        'TIMS.Update_AUTH_ACCTORG(iRoleID, iPlanID, sAccount, sRIDValue, objconn)

        Me.Page.RegisterStartupScript("Errmsg", "<script>alert('儲存成功！');</script>")
        getRelishTB()

    End Sub

    Private Sub rdoIsUsed_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoIsUsed.SelectedIndexChanged
        Call sSearch1() 'but_search_Click(sender, e)
    End Sub

    Private Sub center_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles center.TextChanged
        'getRelishTB()
        Call ClearDataGrid2()
        Call ClearDataGrid1()
        Datagrid2.Visible = False
        DataGrid1.Visible = False
    End Sub

    'SQL 查詢該帳號擁有權限
    Private Sub Lis_acc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Lis_acc.SelectedIndexChanged
        If Lis_acc.SelectedIndex <> 0 Then
            '顯示此登入者的可用計畫年度dropdownlist
            Show_Years() '篩選可用年度
            'Years_SelectedIndexChanged(sender, e) '啟動
            '取出目前可賦予之權限 sm.UserInfo.UserID
            Call getRelishTB()

            Datagrid2.Visible = True
            DataGrid1.Visible = True
            bt_save.Visible = True
        Else
            'DataGrid1.DataSource = Nothing 'DataGrid1.DataBind()
            Call ClearDataGrid1()
            ClearShow()
        End If
    End Sub

    '年度擁有計畫？
    Private Sub Years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Years.SelectedIndexChanged
        '取出目前可賦予之權限 sm.UserInfo.UserID
        Call getRelishTB()
    End Sub

    Sub ClearDataGrid1()
        '依據SQL清除
        'Dim dt As New DataTable '= Nothing
        ''Dim da As SqlDataAdapter = TIMS.GetOneDA()
        'Dim Sql As String = ""
        'Sql += " select a.RID,b.TPlanID, b.PlanID,a.CreateByAcc" & vbCrLf
        'Sql += " ,CASE WHEN o2.OrgName is null then b.Years+e.Name+f.PlanName+b.Seq" & vbCrLf
        'Sql += " 	ELSE b.Years+e.Name+f.PlanName+b.Seq +'　('+o2.OrgName+')' END" & vbCrLf
        'Sql += "  +case when f.Clsyear is null or f.Clsyear > b.Years then '' else '…已停用'+ CONVERT(varchar, f.Clsyear) end AS PlanName" & vbCrLf
        'Sql += " ,d.OrgID, d.OrgName, ac.isused, c.DistID, c.OrgLevel" & vbCrLf
        'Sql += " ,r3.RID2 CRID,o2.OrgID COrgID, o2.OrgName COrgName" & vbCrLf
        'Sql += " FROM Auth_AccRWPlan a" & vbCrLf
        'Sql += " join ID_Plan b on a.PlanID=b.PlanID" & vbCrLf
        'Sql += " join Auth_Relship c on c.RID=a.RID" & vbCrLf
        'Sql += " join Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
        'Sql += " join ID_District e on b.DistID=e.DistID" & vbCrLf
        'Sql += " join Key_Plan f on b.TPlanID=f.TPlanID" & vbCrLf
        'Sql += " join auth_account ac on a.account=ac.account" & vbCrLf
        'Sql += " left join view_Relship23x r3 on r3.distid =e.distid and r3.planid=b.planid and r3.rid3=c.rid" & vbCrLf
        'Sql += " left join Org_OrgInfo o2 on o2.OrgID=r3.OrgID2" & vbCrLf
        'Sql += " WHERE 1<>1" & vbCrLf
        'da.SelectCommand.Parameters.Clear()
        'TIMS.Fill(Sql, da, dt)
        Dim dt As New DataTable '= Nothing
        DataGrid1.DataSource = dt 'dt.DefaultView
        DataGrid1.DataBind()
    End Sub

    Sub ClearDataGrid2()
        '依據SQL清除
        'Dim dt As DataTable = Nothing
        'Dim da As SqlDataAdapter = TIMS.GetOneDA()
        'Dim Sql As String = ""
        'Sql += " SELECT o.orgid,o.orgname,ip.PlanName,ip.years" & vbCrLf
        'Sql += " ,AR.planid, AR.RID, AR.DistID" & vbCrLf
        'Sql += " ,null OrgName3" & vbCrLf
        'Sql += " ,null OrgName2" & vbCrLf
        'Sql += " from auth_relship ar" & vbCrLf
        'Sql += " join VIEW_LOGINPLAN ip on ip.planid=ar.planid" & vbCrLf
        'Sql += " left join org_orginfo o on ar.orgid=o.orgid" & vbCrLf
        'Sql += " where 1<>1" & vbCrLf
        'da.SelectCommand.Parameters.Clear()
        'TIMS.Fill(Sql, da, dt)
        Dim dt As New DataTable '= Nothing
        Datagrid2.DataSource = dt 'dt.DefaultView
        Datagrid2.DataBind()
    End Sub

    '為輔助地方政府 '查看該 %計畫% 的 機構層級為 2 且底下有機構承辦業務
    Function ChkRIDOrgLevel2(ByVal sRID As String, ByRef oPlanID As String) As Boolean
        oPlanID = ""
        Dim Rst As Boolean = False
        Dim sql As String = ""
        sql &= " select b.PlanID" & vbCrLf
        sql &= " from Auth_Relship b" & vbCrLf
        sql &= " join org_orginfo o2 on o2.orgid=b.orgid" & vbCrLf
        sql &= " where b.OrgLevel='2'" & vbCrLf
        sql &= " and exists ( select 'x' from view_Relship23X x where x.RID2=b.RID )" & vbCrLf
        sql &= " and b.RID=@RID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sRID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            oPlanID = dt.Rows(0)("PlanID")
            Rst = True
        End If
        Return Rst
    End Function

    Protected Sub btn_QUERY2_Click(sender As Object, e As EventArgs) Handles btn_QUERY2.Click
        'Response.Redirect()
        TIMS.Utl_Redirect1(Me, String.Concat("SYS_01_002_mq.aspx?ID=", TIMS.Get_MRqID(Me)))
    End Sub
End Class
