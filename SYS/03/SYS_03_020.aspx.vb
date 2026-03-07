Partial Class SYS_03_020
    Inherits AuthBasePage

    'SqlDataAdapter DataSet TIMS.GetOneDA(objconn) TIMS.GetOneDA()
    'Catch ex As Exception
    'Dim objConn As SqlConnection=DbAccess.GetConnection()
    'arrFun=FunSort.Split(",")
    'Dim arrFun As String()  '= {"TC", "SD", "CP", "TR", "CM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"} 'fun排列順序
    'Dim FunSort As String=System.Configuration.ConfigurationSettings.AppSettings("FunSort")

    Const cst_DG1cmd_EDIT1 As String = "EDIT1"  '修改
    Const cst_DG1cmd_DEL1 As String = "DEL1"   '刪除
    Const cst_DG1cmd_PRINT1 As String = "PRINT1" '報表

    Dim arrFun As String()

#Region "Update_IDFunction"

    Sub UPDATE_IDFUNCTION(ByVal tmpID As Integer)
        Dim sqlStr As String = ""
        sqlStr = ""
        sqlStr &= " UPDATE ID_FUNCTION "
        sqlStr &= " SET Name=@Name ,SPage=@SPage ,Kind=@Kind ,Levels=@Levels ,Parent=@Parent ,Valid=@Valid "
        sqlStr &= "  ,Sort=@Sort ,ModifyAcct=@ModifyAcct ,ModifyDate=current_timestamp ,Memo= @Memo ,Adds= @Adds ,Mod= @Mod ,Del= @Del "
        sqlStr &= "  ,Sech=@Sech ,Prnt=@Prnt ,ISREPORT=@ISREPORT "
        sqlStr &= " WHERE FunID= @FunID "
        Dim uCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("Name", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtfunname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtfunname")))
            .Parameters.Add("SPage", SqlDbType.VarChar).Value = If(Convert.ToString(Me.ViewState("txtfunroot")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtfunroot")))
            .Parameters.Add("Kind", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("listmainmenu"))
            .Parameters.Add("Levels", SqlDbType.VarChar).Value = If(Convert.ToString(Me.ViewState("listparentfun")) = "0", "0", "1")
            .Parameters.Add("Parent", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("listparentfun"))
            .Parameters.Add("Valid", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("chkvalid"))
            .Parameters.Add("Sort", SqlDbType.Int).Value = CInt(Val(ViewState("listsort")))
            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
            .Parameters.Add("Memo", SqlDbType.VarChar).Value = If(Convert.ToString(Me.ViewState("txtnote")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtnote")))
            .Parameters.Add("Adds", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("adds"))
            .Parameters.Add("Mod", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("mod"))
            .Parameters.Add("Del", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("del"))
            .Parameters.Add("Sech", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("sech"))
            .Parameters.Add("Prnt", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("prnt"))
            .Parameters.Add("ISREPORT", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("ISREPORT"))
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    ''' <summary>
    ''' [刪除] 但只是功能停用
    ''' </summary>
    ''' <param name="tmpID"></param>
    ''' <param name="tmpState"></param>
    Sub UPDATE_IDFUNCTION(ByVal tmpID As Integer, ByVal tmpState As String)
        Dim sqlStr As String = ""
        sqlStr = ""
        sqlStr &= " UPDATE ID_FUNCTION"
        sqlStr &= " Set FSTATE=@FState "
        If UCase(tmpState) = "D" Then sqlStr &= " ,VALID='N' "
        sqlStr &= " WHERE FUNID=@FunID "
        Dim uCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("FState", SqlDbType.VarChar).Value = tmpState
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    ''' <summary>
    ''' 修改-功能類型-KIND
    ''' </summary>
    ''' <param name="tmpID"></param>
    ''' <param name="tmpKind"></param>
    Sub UPDATE_IDFUNCTION(ByVal tmpID As String, ByVal tmpKind As String)
        Dim sqlStr As String = ""
        sqlStr = " UPDATE ID_FUNCTION SET KIND=@Kind WHERE Parent=@Parent "
        Dim uCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("Kind", SqlDbType.VarChar).Value = tmpKind
            .Parameters.Add("Parent", SqlDbType.VarChar).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    Sub UPDATE_IDFUNCTION(ByVal tmpID As Integer, ByVal tmpSort As Integer)
        Dim sqlStr As String = ""
        sqlStr = " UPDATE ID_FUNCTION SET SORT=@Sort WHERE FunID=@FunID "
        Dim uCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("Sort", SqlDbType.Int).Value = tmpSort
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    Sub UPDATE_AUTHGROUPFUN(ByVal tmpID As Integer, ByVal tmpItem As String)
        Dim tmpItem_up As String = TIMS.ClearSQM(tmpItem).ToUpper()
        If tmpItem_up = "" Then Return
        Select Case tmpItem_up
            Case "ADDS"
            Case "MOD"
            Case "DEL"
            Case "SECH"
            Case "PRNT"
            Case Else
                Return
        End Select
        TIMS.OpenDbConn(objconn)
        Dim sqlStr As String = ""
        sqlStr = String.Concat(" UPDATE AUTH_GROUPFUN SET ", tmpItem_up, "='0' WHERE FunID=@FunID")
        Dim uCmd As New SqlCommand(sqlStr, objconn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    ''' <summary>刪除群組中的功能項</summary>
    ''' <param name="tmpID"></param>
    Sub DEL_AUTHGROUPFUN(ByVal tmpID As Integer)
        Dim sqlStr As String = ""
        sqlStr = " DELETE AUTH_GROUPFUN WHERE FunID=@FunID "
        Dim dCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    ''' <summary>
    ''' 刪除計畫中的功能項
    ''' </summary>
    ''' <param name="tmpID"></param>
    Sub REMOVE_AUTHPLANFUNCTION(ByVal tmpID As Integer)
        Dim sqlStr As String = ""
        sqlStr = " DELETE AUTH_PLANFUNCTION WHERE FunID=@FunID "
        Dim dCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

    ''' <summary>
    ''' 刪除使用者的功能項
    ''' </summary>
    ''' <param name="tmpID"></param>
    Sub DEL_AUTHACCRWFUN(ByVal tmpID As Integer)
        Dim sqlStr As String = ""
        sqlStr = " DELETE AUTH_ACCRWFUN WHERE FunID=@FunID "
        Dim dCmd As New SqlCommand(sqlStr, objconn)
        TIMS.OpenDbConn(objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            .ExecuteNonQuery()
        End With
    End Sub

#End Region

#Region "Function"

    '查詢 SQL
    Sub SSearch1(ByVal sMenu3Val As String)
        'Optional ByVal sMenu3Val As String=""
        'Dim arrFun()={"TC", "SD", "CP", "TR", "CM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"} 'fun排列順序
        'arrFun=FunSort.Split(",")
        'sch_txSPAGE.Text=TIMS.ClearSQM(sch_txSPAGE.Text)
        arrFun = TIMS.c_FUNSORT.Split(",")
        Dim sql As String = ""

        Dim dt_fun As New DataTable
        Dim dt_kind As DataTable = Nothing
        Dim dr As DataRow = Nothing
        '組清單用DataTable
        dt_fun.Columns.Add(New DataColumn("funid"))
        dt_fun.Columns.Add(New DataColumn("name"))
        dt_fun.Columns.Add(New DataColumn("spage"))
        dt_fun.Columns.Add(New DataColumn("kind"))
        dt_fun.Columns.Add(New DataColumn("levels"))
        dt_fun.Columns.Add(New DataColumn("parent"))
        dt_fun.Columns.Add(New DataColumn("sort"))
        dt_fun.Columns.Add(New DataColumn("memo"))
        dt_fun.Columns.Add(New DataColumn("valid"))
        dt_fun.Columns.Add(New DataColumn("subs"))

        Dim ds As New DataSet
        Dim sdaSelect As New SqlDataAdapter
        For i As Integer = 0 To arrFun.Length - 1
            '取得功能清單
            sql = "" & vbCrLf
            sql &= " SELECT a.funid,a.name ,a.spage ,a.kind ,a.levels,a.parent" & vbCrLf
            sql &= " ,ISNULL(p.sort,a.sort) psort ,a.sort ,a.memo ,a.valid" & vbCrLf
            sql &= " ,(CASE a.levels WHEN '0' THEN (SELECT COUNT(1) FROM ID_FUNCTION x WHERE x.PARENT=a.funid) ELSE 0 END) subs" & vbCrLf
            sql &= " ,concat(vk.KINDNAME,'>',p.Name,'>',a.Name) COLLATE Chinese_Taiwan_Stroke_CI_AS funPath" & vbCrLf
            sql &= " FROM ID_FUNCTION a" & vbCrLf
            sql &= " LEFT JOIN V_FUNCKIND vk ON vk.kind=a.Kind COLLATE Chinese_Taiwan_Stroke_CI_AS" & vbCrLf
            sql &= " LEFT JOIN ID_FUNCTION p ON p.funid=a.parent" & vbCrLf
            sql &= " WHERE ISNULL(a.FState,' ') NOT IN ('D')" & vbCrLf
            sql &= " AND a.kind=@kind " & vbCrLf
            'If sch_txSPAGE.Text <> "" Then
            '    Dim sch_txSPAGE_lk As String=String.Concat("%", sch_txSPAGE, "%").ToUpper()
            '    sql &= " AND (a.spage is null" & vbCrLf
            '    sql &= String.Concat(" OR UPPER(a.spage) Like '", sch_txSPAGE_lk, "'") & vbCrLf
            '    sql &= String.Concat(" OR concat(vk.KINDNAME,'>',p.Name,'>',a.Name) LIKE '", sch_txSPAGE_lk, "' COLLATE Chinese_Taiwan_Stroke_CI_AS") & vbCrLf
            '    sql &= " )" & vbCrLf
            'End If
            If sMenu3Val <> "" Then
                Select Case sMenu3Val
                    Case "ALL"
                    Case "P"
                        sql &= " AND a.spage IS NOT NULL AND a.parent='0' AND a.levels='0' " & vbCrLf
                    Case Else
                        sql &= " AND (1!=1 " & vbCrLf
                        sql &= " OR (a.spage IS NOT NULL AND a.parent='" & sMenu3Val & "' AND a.levels='1') " & vbCrLf
                        sql &= " OR (a.funid='" & sMenu3Val & "')" & vbCrLf
                        sql &= " )" & vbCrLf
                End Select
            End If
            sql &= " ORDER BY psort ,a.kind ,a.levels ,a.sort " & vbCrLf

            Dim v_list_MainMenu2 As String = TIMS.GetListValue(list_MainMenu2)
            With sdaSelect
                .SelectCommand = New SqlCommand(sql, objconn)
                .SelectCommand.Parameters.Clear()
                If list_MainMenu2.SelectedIndex <> 0 AndAlso v_list_MainMenu2 <> "" Then
                    .SelectCommand.Parameters.Add("kind", SqlDbType.VarChar).Value = v_list_MainMenu2
                Else
                    .SelectCommand.Parameters.Add("kind", SqlDbType.VarChar).Value = arrFun(i)
                End If
                .Fill(ds, "select_fun" & i)
            End With

            For j As Integer = 0 To ds.Tables("select_fun" & i).Rows.Count - 1
                dr = dt_fun.NewRow
                dt_fun.Rows.Add(dr)
                dr("funid") = ds.Tables("select_fun" & i).Rows(j)("funid")
                dr("name") = ds.Tables("select_fun" & i).Rows(j)("name")
                dr("spage") = ds.Tables("select_fun" & i).Rows(j)("spage")
                dr("kind") = ds.Tables("select_fun" & i).Rows(j)("kind")
                dr("levels") = ds.Tables("select_fun" & i).Rows(j)("levels")
                dr("parent") = ds.Tables("select_fun" & i).Rows(j)("parent")
                dr("sort") = ds.Tables("select_fun" & i).Rows(j)("sort")
                dr("memo") = ds.Tables("select_fun" & i).Rows(j)("memo")
                dr("valid") = ds.Tables("select_fun" & i).Rows(j)("valid")
                dr("subs") = ds.Tables("select_fun" & i).Rows(j)("subs")
            Next

            'list_MainMenu3 設定
            If sMenu3Val = "" AndAlso list_MainMenu2.SelectedIndex <> 0 AndAlso v_list_MainMenu2 <> "" Then
                list_MainMenu3.Items.Clear()
                Dim dt2 As New DataTable '= Nothing
                sql = "" & vbCrLf
                sql &= " SELECT a.funid " & vbCrLf
                sql &= "  ,a.name ,a.spage ,a.kind ,a.levels " & vbCrLf
                sql &= "  ,a.parent " & vbCrLf
                sql &= "  ,ISNULL(p.sort,a.sort) psort ,a.sort ,a.memo " & vbCrLf
                'sql += " ,REPLACE(a.valid,' ','') valid " & vbCrLf
                sql &= "  ,a.valid " & vbCrLf
                sql &= "  ,(CASE a.levels WHEN '0' THEN (SELECT COUNT(funid) cnt FROM id_function WHERE parent=a.funid) ELSE 0 END) AS subs " & vbCrLf
                sql &= "  FROM id_function a " & vbCrLf
                sql &= "  LEFT JOIN id_function p ON p.funid=a.parent " & vbCrLf
                sql &= "  WHERE ISNULL(a.FState,' ') NOT IN ('D') " & vbCrLf
                sql &= "   AND a.spage IS NULL AND a.levels='0' " & vbCrLf
                sql &= "   AND a.kind=@kind " & vbCrLf
                sql &= "  ORDER BY a.kind ,psort ,a.levels ,a.sort " & vbCrLf
                Dim sCmd As New SqlCommand(sql, objconn)
                TIMS.OpenDbConn(objconn)
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("kind", SqlDbType.VarChar).Value = v_list_MainMenu2 'list_MainMenu2.SelectedValue
                    dt2.Load(.ExecuteReader())
                End With
                'da.SelectCommand.Parameters.Clear()
                'da.SelectCommand.Parameters.Add("kind", SqlDbType.VarChar).Value=list_MainMenu2.SelectedValue
                'TIMS.Fill(sql, da, dt2)
                'If da.SelectCommand.Connection.State=ConnectionState.Open Then da.SelectCommand.Connection.Close()
                With list_MainMenu3
                    .DataSource = dt2
                    .DataValueField = "funid"
                    .DataTextField = "name"
                    .DataBind()
                End With
                list_MainMenu3.Items.Insert(0, New ListItem("無", "P"))
                list_MainMenu3.Items.Insert(0, New ListItem("全部", "ALL"))
            End If

            If list_MainMenu2.SelectedIndex <> 0 Then Exit For
        Next

        '取得功能種類
        sql = "SELECT DISTINCT KIND FROM ID_FUNCTION WITH(NOLOCK) ORDER BY KIND"
        With sdaSelect
            .SelectCommand = New SqlCommand(sql, objconn)
            .Fill(ds, "select_kind")
        End With
        If ds.Tables("select_kind") IsNot Nothing Then
            If ds.Tables("select_kind").Rows.Count > 0 Then dt_kind = ds.Tables("select_kind")
        End If
        For i As Integer = 0 To dt_kind.Rows.Count - 1
            dr = dt_kind.Rows(i)
            Me.ViewState(dr("kind").ToString) = dt_fun.Select("kind='" & dr("kind").ToString & "'").Length
        Next

        'intCntSch=0
        DataGrid1.DataSource = dt_fun
        DataGrid1.DataKeyField = "FunID"
        DataGrid1.DataBind()

        '    Try
        '        'If objConn.State=ConnectionState.Open Then objConn.Close()
        '        'Call TIMS.CloseDbConn(objConn)
        '        If sdaSelect IsNot Nothing Then sdaSelect.Dispose()
        '        If ds IsNot Nothing Then ds.Dispose()
        '    Catch ex As Exception
        '        Common.MessageBox(Me, ex.ToString)
        '    End Try
    End Sub

    '取得功能清單
    Function GET_IDFUNCTION(ByVal tmpKind As String, ByVal tmpLV As String, ByVal tmpParent As String) As DataTable
        Dim rst As New DataTable

        Dim sqlStr As String = ""
        sqlStr &= " SELECT a.FunID,a.Name " & vbCrLf
        sqlStr &= " ,a.Name+CASE WHEN a.SPage IS NOT NULL THEN ' (SP)' ELSE '' END Name2 " & vbCrLf
        sqlStr &= " ,a.SPage,a.Kind,a.Levels" & vbCrLf
        sqlStr &= " ,(CASE CONVERT(VARCHAR, a.Levels) WHEN '0' THEN CONVERT(VARCHAR, a.FunID) ELSE CONVERT(VARCHAR, a.Parent) END) Parent " & vbCrLf
        sqlStr &= " ,a.Sort,a.Memo,a.Adds,a.Mod,a.Del,a.Sech,a.Prnt,a.ISREPORT,a.Valid" & vbCrLf
        sqlStr &= " ,(CASE CONVERT(VARCHAR, a.levels) WHEN '0' THEN (SELECT COUNT(x.FunID) FROM ID_Function x WHERE x.Parent=a.FunID) ELSE 0 END) Subs " & vbCrLf
        sqlStr &= " ,a.PSORT " & vbCrLf
        sqlStr &= " FROM dbo.V_FUNCTION a " & vbCrLf
        sqlStr &= " WHERE ISNULL(a.FState,' ') NOT IN ('D') " & vbCrLf
        If tmpKind <> "" Then sqlStr &= " AND a.Kind=@Kind " & vbCrLf
        If tmpLV <> "" Then sqlStr &= " AND a.Levels=@Levels " & vbCrLf
        If tmpParent <> "" Then sqlStr &= " AND a.Parent=@Parent " & vbCrLf
        sqlStr &= " ORDER BY a.Kind ,a.PSORT ,a.Levels ,a.Sort " & vbCrLf
        Dim sCmd As New SqlCommand(sqlStr, objconn)

        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            If tmpKind <> "" Then .Parameters.Add("Kind", tmpKind)
            If tmpLV <> "" Then .Parameters.Add("Levels", tmpLV)
            If tmpParent <> "" Then .Parameters.Add("Parent", tmpParent)
            rst.Load(.ExecuteReader())
        End With
        If rst.Rows.Count = 0 Then rst = Nothing '若為無資料設定為nothing

        Return rst
    End Function

    Function GET_IDFUNCTION(ByVal tmpID As Integer) As DataRow
        Dim rst As DataRow = Nothing
        Dim sqlStr As String
        sqlStr = " SELECT * FROM ID_FUNCTION WHERE ISNULL(FState,' ') NOT IN ('D') AND FunID=@FunID "
        Dim sCmd As New SqlCommand(sqlStr, objconn)
        Dim dt As New DataTable
        TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = dt.Rows(0)
        Return rst
    End Function

    Function INSERT_IDFUNCTION() As Integer
        Dim iFUNID As Integer = 0
        Dim sqlStr As String = ""
        sqlStr &= " INSERT INTO ID_FUNCTION(FUNID,Name,SPage,Kind,Levels,Parent,Valid,Sort,ModifyAcct,ModifyDate,Memo,Adds,Mod,Del,Sech,Prnt,ISREPORT) "
        sqlStr &= " VALUES(@FUNID,@Name,@SPage,@Kind,@Levels,@Parent,@Valid,@Sort,@ModifyAcct,GETDATE(),@Memo,@Adds,@Mod,@Del,@Sech,@Prnt,@ISREPORT) "
        Dim iCmd As New SqlCommand(sqlStr, objconn)
        iFUNID = DbAccess.GetNewId(objconn, "ID_FUNCTION_FUNID_SEQ,ID_FUNCTION,FUNID")
        TIMS.OpenDbConn(objconn)
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("FUNID", SqlDbType.Int).Value = iFUNID
            .Parameters.Add("Name", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtfunname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtfunname")))
            .Parameters.Add("SPage", SqlDbType.VarChar).Value = If(Convert.ToString(Me.ViewState("txtfunroot")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtfunroot")))
            .Parameters.Add("Kind", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("listmainmenu"))
            .Parameters.Add("Levels", SqlDbType.VarChar).Value = If(Convert.ToString(Me.ViewState("listparentfun")) = "0", "0", "1")
            .Parameters.Add("Parent", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("listparentfun"))
            .Parameters.Add("Valid", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("chkvalid"))
            .Parameters.Add("Sort", SqlDbType.Int).Value = CInt(Val(ViewState("listsort")))
            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
            .Parameters.Add("Memo", SqlDbType.VarChar).Value = If(Convert.ToString(Me.ViewState("txtnote")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtnote")))
            .Parameters.Add("Adds", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("adds"))
            .Parameters.Add("Mod", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("mod"))
            .Parameters.Add("Del", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("del"))
            .Parameters.Add("Sech", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("sech"))
            .Parameters.Add("Prnt", SqlDbType.Char).Value = Convert.ToString(Me.ViewState("prnt"))
            .Parameters.Add("ISREPORT", SqlDbType.VarChar).Value = Convert.ToString(Me.ViewState("ISREPORT"))
            .ExecuteNonQuery()
        End With

        Return iFUNID
    End Function

    Sub Show_DataGrid1(ByVal tmpDT As DataTable, ByVal tmpKind As String)
        ' Optional ByVal tmpKind As String=""
        If tmpKind = "" Then
            DataGrid1.DataSource = tmpDT
        Else
            DataGrid1.DataSource = tmpDT.Select("Kind='" & tmpKind & "'")
        End If
        DataGrid1.DataKeyField = "FunID"
        DataGrid1.DataBind()
    End Sub

    Sub Show_ListParentFun(ByVal tmpDT As DataTable, ByVal tmpSelect As String)
        'Optional ByVal tmpSelect As String=""
        'dv.RowFilter="Subs<>'0'"
        Dim dv As DataView = tmpDT.DefaultView
        dv.RowFilter = "SPage Is NULL"
        dv.Sort = "PSORT"

        list_ParentFun.Items.Clear()
        If tmpDT IsNot Nothing Then
            list_ParentFun.DataSource = dv 'tmpDT
            list_ParentFun.DataTextField = "Name"
            list_ParentFun.DataValueField = "FunID"
            list_ParentFun.DataBind()

            list_ParentFun.Items.Insert(0, New ListItem(cst_請選擇3, ""))
            If tmpSelect <> "" And tmpSelect <> "0" Then
                Common.SetListItem(list_ParentFun, tmpSelect)
            Else
                list_ParentFun.SelectedIndex = 0
            End If
            'If tmpSelect <> "" And tmpSelect <> "0" Then list_ParentFun.SelectedValue=tmpSelect Else list_ParentFun.SelectedIndex=0
        End If
        list_ParentFun.Enabled = (tmpDT IsNot Nothing)
    End Sub

    '存在(有選擇)為True 不存在(沒有選擇)為false
    Public Shared Function Check_cblTPlanIDed(ByVal str_NOTPlanID As String, ByVal TPlanIDval As String) As Boolean
        Dim Rst As Boolean = False
        Dim ary_NOTPlanID As String() = Split(str_NOTPlanID, ",")
        'Check_cblTPlanID2=False
        For i As Integer = 0 To ary_NOTPlanID.Length - 1
            If ary_NOTPlanID(i).ToString = TPlanIDval Then
                Rst = True
                Exit For
            End If
        Next
        Return Rst
    End Function

    '搜尋排除名單
    Public Shared Function Get_cblTPlanIDSelected(ByVal chkobj As CheckBoxList) As String
        Dim str As String = ""
        For i As Integer = 1 To chkobj.Items.Count - 1
            If chkobj.Items(i).Selected Then
                str &= $"{If(str <> "", ",", "")}{chkobj.Items(i).Value}"
            End If
        Next
        Return str
    End Function

    '新增 做排除
    Sub Insert_AuthPlanFunction(ByVal tmpID As Integer)
        'tmpID: FunID 
        Dim flagCanSave As Boolean = False '可以儲存(false不可以)
        If tmpID > 0 Then flagCanSave = True '可以儲存
        If Not flagCanSave Then Return 'Exit Sub 'false不可以:離開

        Dim sqlAdp As New SqlDataAdapter
        Dim objDt As DataTable
        Dim sqlStr As String
        Dim str_NOTPlanID As String = Get_cblTPlanIDSelected(cblTPlanID2)

        Dim sTPlanID As String = ""

        '新增 動作
        sqlStr = " INSERT INTO AUTH_PLANFUNCTION(TPlanID,FunID,ModifyAcct,ModifyDate) values(@TPlanID,@FunID,@ModifyAcct,GETDATE()) "
        sqlAdp.InsertCommand = New SqlCommand(sqlStr, objconn)
        '刪除 動作
        sqlStr = " DELETE Auth_PlanFunction WHERE TPlanID= @TPlanID  AND FunID= @FunID "
        sqlAdp.DeleteCommand = New SqlCommand(sqlStr, objconn)

        'sqlStr=" SELECT TPlanID FROM Key_Plan WHERE TPlanID NOT IN (SELECT TPlanID FROM Auth_PlanFunction WHERE FunID=@FunID) "
        sqlStr = " SELECT TPlanID FROM Key_Plan p WHERE EXISTS (SELECT 'x' FROM Auth_PlanFunction x WHERE x.FunID=@FunID AND x.TPlanID=p.TPlanID) "
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objconn)
            .SelectCommand.Parameters.Clear()
            .SelectCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            objDt = New DataTable
            .Fill(objDt)
        End With
        If objDt.Rows.Count > 0 Then
            For i As Integer = 0 To objDt.Rows.Count - 1
                sTPlanID = Convert.ToString(objDt.Rows(i).Item("TPlanID"))
                With sqlAdp
                    If Check_cblTPlanIDed(str_NOTPlanID, sTPlanID) Then
                        '若是排除名單做刪除
                        .DeleteCommand.Parameters.Clear()
                        .DeleteCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sTPlanID
                        .DeleteCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
                        .DeleteCommand.ExecuteNonQuery()
                    End If
                End With
            Next
        End If

        sqlStr = " SELECT TPLANID FROM KEY_PLAN P WHERE NOT EXISTS (SELECT 'x' FROM Auth_PlanFunction x WHERE x.FunID=@FunID AND x.TPlanID=p.TPlanID) "
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objconn)
            .SelectCommand.Parameters.Clear()
            .SelectCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            objDt = New DataTable
            .Fill(objDt)
        End With
        If objDt.Rows.Count > 0 Then
            For i As Integer = 0 To objDt.Rows.Count - 1
                sTPlanID = Convert.ToString(objDt.Rows(i).Item("TPlanID"))
                With sqlAdp
                    If Not Check_cblTPlanIDed(str_NOTPlanID, sTPlanID) Then
                        '若不是排除名單可新增
                        .InsertCommand.Parameters.Clear()
                        .InsertCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sTPlanID
                        .InsertCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
                        .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        .InsertCommand.ExecuteNonQuery()
                    End If
                End With
            Next
        End If

        'Try

        '    If sqlAdp IsNot Nothing Then sqlAdp.Dispose()
        '    If objDt IsNot Nothing Then objDt.Dispose()
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
    End Sub

    '刪除 做新增
    Sub Insert_AuthPlanFunction2(ByVal tmpID As Integer)
        'tmpID: FunID 
        Dim flagCanSave As Boolean = False '可以儲存(false不可以)
        If tmpID > 0 Then flagCanSave = True '可以儲存
        If Not flagCanSave Then Return 'Exit Sub 'false不可以:離開

        Dim sqlAdp As New SqlDataAdapter
        Dim objDt As DataTable
        Dim sqlStr As String
        Dim str_AddPlanID As String = Get_cblTPlanIDSelected(cblTPlanID3)

        Dim sTPlanID As String = ""

        '新增 動作
        sqlStr = " INSERT INTO Auth_PlanFunction(TPlanID,FunID,ModifyAcct,ModifyDate) values(@TPlanID,@FunID,@ModifyAcct,current_timestamp) "
        sqlAdp.InsertCommand = New SqlCommand(sqlStr, objconn)
        '刪除 動作
        sqlStr = " DELETE Auth_PlanFunction WHERE TPlanID=@TPlanID AND FunID= @FunID "
        sqlAdp.DeleteCommand = New SqlCommand(sqlStr, objconn)

        'sqlStr=" SELECT TPlanID FROM Key_Plan WHERE TPlanID NOT IN (SELECT TPlanID FROM Auth_PlanFunction WHERE FunID=@FunID) "
        sqlStr = " SELECT TPlanID FROM Key_Plan p WHERE NOT EXISTS (SELECT 'x' FROM Auth_PlanFunction x WHERE x.FunID=@FunID AND x.TPlanID=p.TPlanID) "
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objconn)
            .SelectCommand.Parameters.Clear()
            .SelectCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            objDt = New DataTable
            .Fill(objDt)
        End With

        If objDt.Rows.Count > 0 Then
            For i As Integer = 0 To objDt.Rows.Count - 1
                sTPlanID = Convert.ToString(objDt.Rows(i).Item("TPlanID"))
                With sqlAdp
                    If Check_cblTPlanIDed(str_AddPlanID, sTPlanID) Then
                        '若是新增名單做 新增
                        '.DeleteCommand.Parameters.Clear()
                        '.DeleteCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value=sTPlanID
                        '.DeleteCommand.Parameters.Add("FunID", SqlDbType.Int).Value=tmpID
                        '.DeleteCommand.ExecuteNonQuery()
                        .InsertCommand.Parameters.Clear()
                        .InsertCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sTPlanID
                        .InsertCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
                        .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        .InsertCommand.ExecuteNonQuery()
                    End If
                End With
            Next
        End If

        sqlStr = " SELECT TPlanID FROM Key_Plan p WHERE EXISTS (SELECT 'x' FROM Auth_PlanFunction x WHERE x.FunID=@FunID AND x.TPlanID=p.TPlanID) "
        With sqlAdp
            .SelectCommand = New SqlCommand(sqlStr, objconn)
            .SelectCommand.Parameters.Clear()
            .SelectCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
            objDt = New DataTable
            .Fill(objDt)
        End With
        If objDt.Rows.Count > 0 Then
            For i As Integer = 0 To objDt.Rows.Count - 1
                sTPlanID = Convert.ToString(objDt.Rows(i).Item("TPlanID"))

                With sqlAdp
                    If Not Check_cblTPlanIDed(str_AddPlanID, sTPlanID) Then
                        '若不是 新增名單做 刪除
                        .DeleteCommand.Parameters.Clear()
                        .DeleteCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value = sTPlanID
                        .DeleteCommand.Parameters.Add("FunID", SqlDbType.Int).Value = tmpID
                        .DeleteCommand.ExecuteNonQuery()
                    End If
                End With
            Next
        End If

        'Try

        '    If sqlAdp IsNot Nothing Then sqlAdp.Dispose()
        '    If objDt IsNot Nothing Then objDt.Dispose()
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try
    End Sub

    Sub Show_ListSort(ByVal tmpDT As DataTable, ByVal tmpSelect As String)
        'Optional ByVal tmpSelect As String=""
        list_Sort.Items.Clear()
        'If tmpDT Is Nothing Then Return
        If tmpDT IsNot Nothing Then
            Dim dv As DataView = tmpDT.DefaultView
            dv.RowFilter = "1=1"
            dv.Sort = "SORT"
            list_Sort.DataSource = dv 'tmpDT
            list_Sort.DataTextField = "Name2"
            list_Sort.DataValueField = "FunID"
            list_Sort.DataBind()
        End If
        If tmpSelect <> "" Then
            list_Sort.SelectedValue = tmpSelect
            list_Sort.Attributes.Add("onChange", "document.getElementById('" & list_Sort.ClientID & "').value=document.getElementById('hide_FunID').value;")
        Else
            Dim v_txt_FunName As String = If(txt_FunName.Text <> "", txt_FunName.Text, "新項目")
            Dim v_hide_FunID As String = If(hide_FunID.Value <> "", hide_FunID.Value, "new")
            list_Sort.Items.Add(New ListItem(v_txt_FunName, v_hide_FunID))
            list_Sort.SelectedValue = v_hide_FunID 'If(hide_FunID.Value="", "new", hide_FunID.Value)
            list_Sort.Attributes.Add("onChange", String.Concat("document.getElementById('", list_Sort.ClientID, "').value='new';"))
        End If
        If list_Sort IsNot Nothing Then hide_Sort.Value = list_Sort.SelectedIndex
    End Sub

    Sub Clear_Items()
        list_MainMenu.SelectedIndex = 0
        txt_FunName.Text = ""
        TIMS.Tooltip(txt_FunName, "")
        hide_FunID.Value = ""
        hide_Subs.Value = ""
        txt_FunRoot.Text = ""
        hide_Sort.Value = -1 '""

        list_ParentFun.Items.Clear()
        list_ParentFun.Items.Add(New ListItem(cst_請選擇3, ""))
        list_Sort.Items.Clear()
        list_Sort.Items.Add(New ListItem("新增項目", "new"))
        CHK_VALID.Checked = True
        CHK_DEL_REL_1.Checked = False
        txt_Note.Text = ""
        For i As Integer = 0 To chk_Option.Items.Count - 1
            chk_Option.Items.Item(i).Selected = True
        Next

        Me.ViewState("listmainmenu") = Nothing
        Me.ViewState("txtfunname") = Nothing
        Me.ViewState("txtfunroot") = Nothing
        Me.ViewState("listparentfun") = Nothing
        Me.ViewState("chkvalid") = Nothing
        Me.ViewState("listsort") = Nothing
        Me.ViewState("txtnote") = Nothing
        Me.ViewState("adds") = Nothing
        Me.ViewState("del") = Nothing
        Me.ViewState("mod") = Nothing
        Me.ViewState("sech") = Nothing
        Me.ViewState("prnt") = Nothing
        Me.ViewState("ISREPORT") = Nothing
    End Sub

#End Region

    Const cst_tab2 As String = "&nbsp;&nbsp;&nbsp;　"
    Const cst_t請選擇 As String = "==請選擇=="
    Const cst_v請選擇 As String = "==請選擇=="
    Const cst_請選擇3 As String = TIMS.cst_ddl_PleaseChoose3

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'objConn.Open()

        If Not IsPostBack Then
            list_MainMenu2 = TIMS.Get_ddlFunction(list_MainMenu2, 2)
            list_MainMenu = TIMS.Get_ddlFunction(list_MainMenu, cst_t請選擇, cst_v請選擇)

            btn_Up.Attributes.Add("style", "cursor@hand;")
            btn_Up.Attributes.Add("onClick", "return Show_ListSort(document.getElementById('" & list_Sort.ClientID & "'),'up');")

            btn_Down.Attributes.Add("style", "cursor@hand;")
            btn_Down.Attributes.Add("onClick", "return Show_ListSort(document.getElementById('" & list_Sort.ClientID & "'),'down');")

            btn_Save.Attributes.Add("onClick", "return Check_Data();")

            list_Sort.Attributes.Add("readonly", "readonly")

            tb_View.Visible = True
            tb_Edit.Visible = False

            '查詢 SQL
            Call SSearch1("")

            '建立資料
            Call SUtl_create1()

            'Me.chk_RemoveAll.Attributes("onclick")="if(document.getElementById('" & chk_RemoveAll.ClientID & "')) {if(document.getElementById('" & chk_RemoveAll.ClientID & "').checked){ return confirm('將會把該功能從所有計畫移除，確定要移除嗎?');}}"
            TPlanIDFun.Attributes("onclick") = "TPlanIDFunSelected(this,'SPTPlanID2');"
            TPlanIDFun.Attributes("onkeypress") = "TPlanIDFunSelected(this,'SPTPlanID2');"
            SPTPlanID2.Style.Item("display") = "none"
            cblTPlanID2.Attributes("onclick") = "SelectAll('cblTPlanID2','HidTPlanID2','OthTPlanID2');"

            TPlanIDFunDEL.Attributes("onclick") = "TPlanIDFunSelected(this,'SPTPlanID3');"
            TPlanIDFunDEL.Attributes("onkeypress") = "TPlanIDFunSelected(this,'SPTPlanID3');"
            SPTPlanID3.Style.Item("display") = "none"
            cblTPlanID3.Attributes("onclick") = "SelectAll('cblTPlanID3','HidTPlanID3','OthTPlanID3');"

            SPTPlanIDhave.Style.Item("display") = "none"
            SpAuthGroup.Style.Item("display") = "none"
            cblAuthGroup.Attributes("onclick") = "SelectAll2('cblAuthGroup','HidAuthGroup');"
        End If
    End Sub

    '建立資料
    Sub SUtl_create1()
        '排除預設---start
        cblTPlanID2 = TIMS.Get_TPlan(cblTPlanID2, , 2)
        cblTPlanID2.Items.Insert(0, New ListItem("全選", ""))
        OthTPlanID2.Value = ","
        For i As Integer = 1 To cblTPlanID2.Items.Count - 1
            'Cst_NotTPlanID3 
            If TIMS.Cst_NotTPlanID3.IndexOf(cblTPlanID2.Items(i).Value) > -1 Then
                cblTPlanID2.Items(i).Selected = True
                '最後補逗點 (符合javascript)
                OthTPlanID2.Value += "cblTPlanID2_" & CStr(i) & ","
            End If
        Next
        '排除預設---end

        '加入預設---start
        cblTPlanID3 = TIMS.Get_TPlan(cblTPlanID3, , 2)
        cblTPlanID3.Items.Insert(0, New ListItem("全選", ""))
        OthTPlanID3.Value = ","
        For i As Integer = 1 To cblTPlanID3.Items.Count - 1
            'Cst_NotTPlanID3 
            If TIMS.Cst_AddTPlanID3.IndexOf(cblTPlanID3.Items(i).Value) > -1 Then
                cblTPlanID3.Items(i).Selected = True
                '最後補逗點 (符合javascript)
                OthTPlanID3.Value += "cblTPlanID3_" & CStr(i) & ","
            End If
        Next
        '加入預設---end

        '加入預設---start
        cblTPlanIDhave = TIMS.Get_TPlan(cblTPlanIDhave, , 2, , , objconn)
        cblAuthGroup = TIMS.Get_AuthGroup(cblAuthGroup, objconn)
        cblAuthGroup.Items.Insert(0, New ListItem("全選", ""))
        '加入預設---end
    End Sub

    '已存在的群組
    Sub Show_Funid_Group(ByVal sFunID As String)
        sFunID = TIMS.ClearSQM(sFunID) : If sFunID = "" Then sFunID = "0"
        Dim sql As String = ""
        sql = " SELECT * FROM Auth_GroupFun WHERE FunID=@FunID "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.VarChar).Value = sFunID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            For Each xLI1 As ListItem In cblAuthGroup.Items
                If Convert.ToString(xLI1.Value) <> "" Then
                    If dt.Select("GID ='" & xLI1.Value & "'").Length > 0 Then
                        xLI1.Selected = True
                    End If
                End If
            Next
            'For i As Integer=0 To cblAuthGroup.Items.Count - 1
            '    'cblAuthGroup.Items(i).Selected=False
            '    If Convert.ToString(cblAuthGroup.Items(i).Value) <> "" Then
            '        If dt.Select("GID ='" & cblAuthGroup.Items(i).Value & "'").Length > 0 Then
            '            cblAuthGroup.Items(i).Selected=True
            '        End If
            '    End If
            'Next
        End If
    End Sub

    '已存在的計畫
    Sub Show_Funid_TPlan(ByVal sFunID As String)
        sFunID = TIMS.ClearSQM(sFunID) : If sFunID = "" Then sFunID = "0"

        Dim PMS1 As New Hashtable From {{"FunID", sFunID}}
        Dim sql As String = " SELECT * FROM AUTH_PLANFUNCTION WHERE FunID=@FunID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)
        If TIMS.dtNODATA(dt) Then Return

        For Each xLI1 As ListItem In cblTPlanIDhave.Items
            If xLI1.Value <> "" Then
                If dt.Select("TPlanID ='" & xLI1.Value & "'").Length > 0 Then
                    xLI1.Selected = True
                End If
            End If
        Next
    End Sub

    '顯示資料
    Sub Show_Edit(ByRef e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        Dim labFunName As Label = e.Item.FindControl("lab_FunName")
        Dim hideSubs As HtmlInputHidden = e.Item.FindControl("hide_Subs2")

        Dim dr_Function As DataRow = Nothing
        Dim dt_ParentFun As DataTable = Nothing
        Dim dt_Sort As DataTable = Nothing

        tb_View.Visible = False
        tb_Edit.Visible = True
        dr_Function = GET_IDFUNCTION(CInt(DataGrid1.DataKeys.Item(e.Item.ItemIndex)))

        If dr_Function IsNot Nothing Then
            hide_FunID.Value = Convert.ToString(dr_Function("FunID"))
            list_MainMenu.SelectedValue = Convert.ToString(dr_Function("Kind"))
            hide_Kind.Value = Convert.ToString(dr_Function("Kind"))
            txt_FunName.Text = Convert.ToString(dr_Function("Name"))
            TIMS.Tooltip(txt_FunName, hide_FunID.Value, True)
            hide_Subs.Value = hideSubs.Value
            txt_FunRoot.Text = Convert.ToString(dr_Function("SPage"))
            dt_ParentFun = GET_IDFUNCTION(Convert.ToString(dr_Function("Kind")), "0", "")

            If dt_ParentFun IsNot Nothing Then
                Call Show_ListParentFun(dt_ParentFun, Convert.ToString(dr_Function("Parent")))
                If Convert.ToString(dr_Function("Levels")) = "0" Then list_ParentFun.SelectedIndex = 0
            End If

            CHK_VALID.Checked = If(Convert.ToString(dr_Function("Valid")).Replace(" ", "") = "Y", True, False)
            dt_Sort = GET_IDFUNCTION(Convert.ToString(dr_Function("Kind")), Convert.ToString(dr_Function("Levels")), Convert.ToString(dr_Function("Parent")))

            If dt_Sort IsNot Nothing Then Show_ListSort(dt_Sort, Convert.ToString(dr_Function("FunID")))

            txt_Note.Text = Convert.ToString(dr_Function("Memo"))
            chk_Option.Items(0).Selected = If(Convert.ToString(dr_Function("Adds")) = "1", True, False)
            chk_Option.Items(1).Selected = If(Convert.ToString(dr_Function("Del")) = "1", True, False)
            chk_Option.Items(2).Selected = If(Convert.ToString(dr_Function("Mod")) = "1", True, False)
            chk_Option.Items(3).Selected = If(Convert.ToString(dr_Function("Sech")) = "1", True, False)
            chk_Option.Items(4).Selected = If(Convert.ToString(dr_Function("Prnt")) = "1", True, False)
            chk_Option.Items(5).Selected = If(Convert.ToString(dr_Function("ISREPORT")) = "Y", True, False)

            If CInt(hideSubs.Value) > 0 Then list_ParentFun.Enabled = False

            hide_Sort.Value = Convert.ToInt16(dr_Function("Sort")) - 1
        Else
            Common.MessageBox(Me, labFunName.Text & "資料無法調閱。")
        End If

        'If dt_ParentFun IsNot Nothing Then dt_ParentFun.Dispose()
        'If dt_Sort IsNot Nothing Then dt_Sort.Dispose()
    End Sub

    '修正
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Dim labFunName As Label=e.Item.FindControl("lab_FunName")
        'Dim hideSubs As HtmlInputHidden=e.Item.FindControl("hide_Subs2")
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Select Case e.CommandName'UCase()
                    Case cst_DG1cmd_EDIT1 '"EDIT"
                        Call Show_Edit(e)

                    Case cst_DG1cmd_DEL1 '"DEL"
                        UPDATE_IDFUNCTION(CInt(DataGrid1.DataKeys.Item(e.Item.ItemIndex)), "D")
                        If CHK_DEL_REL_2.Checked Then
                            DEL_AUTHGROUPFUN(CInt(DataGrid1.DataKeys.Item(e.Item.ItemIndex)))
                            REMOVE_AUTHPLANFUNCTION(CInt(DataGrid1.DataKeys.Item(e.Item.ItemIndex)))
                            DEL_AUTHACCRWFUN(CInt(DataGrid1.DataKeys.Item(e.Item.ItemIndex)))
                        End If
                        Call SSearch1("")
                        'list_MainMenu2_SelectedIndexChanged(Nothing, Nothing)

                    Case cst_DG1cmd_PRINT1
                        Dim s_CMDARG As String = e.CommandArgument
                        Dim s_FunID As String = TIMS.GetMyValue(s_CMDARG, "FUNID")
                        If s_FunID = "" Then Return 'Exit Sub

                        Dim sUrl As String = "SYS_03_020_P.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                        sUrl &= "&FUNID=" & s_FunID
                        TIMS.Utl_Redirect1(Me, sUrl)
                End Select
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound

        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim labMainMenu As Label = e.Item.FindControl("lab_MainMenu")
                Dim labValid As Label = e.Item.FindControl("lab_Valid")
                Dim labFunName As Label = e.Item.FindControl("lab_FunName")
                Dim hideSubs As HtmlInputHidden = e.Item.FindControl("hide_Subs2")
                Dim btnEdit As LinkButton = e.Item.FindControl("btn_Edit")
                Dim btnDel As LinkButton = e.Item.FindControl("btn_Del")
                Dim BtnPrint1 As LinkButton = e.Item.FindControl("BtnPrint1")
                Dim dr_Data As DataRowView = e.Item.DataItem

                'labMainMenu.Text=Get_MainMenuName(Convert.ToString(dr_Data("Kind")))
                Dim strkind As String = Convert.ToString(dr_Data("Kind"))
                labMainMenu.Text = TIMS.Get_MainMenuName(strkind)
                labFunName.Text = If(Convert.ToString(dr_Data("levels")) = "1", cst_tab2 & Convert.ToString(dr_Data("Name")), Convert.ToString(dr_Data("Name")))
                labValid.Text = If(Convert.ToString(dr_Data("Valid")).Replace(" ", "") = "Y", "是", "否")

                If Convert.ToString(Me.ViewState("mmname")) <> labMainMenu.Text Then
                    Me.ViewState("mmname") = labMainMenu.Text
                    e.Item.Cells(0).RowSpan = Me.ViewState(Convert.ToString(dr_Data("Kind")))
                    e.Item.Cells(0).BackColor = Color.FromArgb(241, 249, 252)
                Else
                    e.Item.Cells(0).Visible = False
                End If
                e.Item.Cells(1).ToolTip = dr_Data("FunID")

                Dim s_page As String = Convert.ToString(dr_Data("SPage")).ToUpper().Replace(".ASPX", "")  '(清除".aspx"關鍵字，by:20181031)
                Dim s_CMDARG As String = ""
                TIMS.SetMyValue(s_CMDARG, "FUNID", Convert.ToString(dr_Data("FUNID")))
                BtnPrint1.CommandArgument = s_CMDARG
                BtnPrint1.Visible = If(s_page <> "", True, False)

                If Convert.ToString(dr_Data("Subs")) <> "0" Then
                    e.Item.BackColor = Color.FromArgb(235, 243, 254)
                    If CInt(dr_Data("Subs")) > 0 Then
                        btnDel.Enabled = False
                        btnDel.ToolTip = "有子項目無法刪除!"
                    End If
                End If

                hideSubs.Value = If(Convert.ToString(dr_Data("Subs")) = "", "0", Convert.ToString(dr_Data("Subs")))
                btnDel.Attributes.Add("onClick", "return confirm('確認要刪除該項目?');")
                e.Item.Cells(3).Text = s_page 'Convert.ToString(dr_Data("SPage")).ToUpper().Replace(".ASPX", "")  '(清除".aspx"關鍵字，by:20181031)
            Case ListItemType.Footer
                Me.ViewState("mmname") = Nothing
        End Select
    End Sub

    Private Sub Btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        Clear_Items()
        tb_View.Visible = False
        tb_Edit.Visible = True
    End Sub

    '回上一頁
    Private Sub Btn_LoadBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_LoadBack.Click
        Clear_Items()
        tb_View.Visible = True
        tb_Edit.Visible = False
    End Sub

    Sub SaveData1()
        Call TIMS.OpenDbConn(objconn)

        Dim tmpFunID As Integer = 0
        Dim dt_Function As DataTable = Nothing
        Dim tmpSort As Integer = list_Sort.SelectedIndex '原排序選取位置
        Dim intCntSave As Integer = 1 '判斷存儲成功(0=>否,1=>是)

        Dim v_txt_FunRoot As String = txt_FunRoot.Text
        If v_txt_FunRoot <> "" Then v_txt_FunRoot = Trim(v_txt_FunRoot)
        Me.ViewState("listmainmenu") = list_MainMenu.SelectedValue
        Me.ViewState("txtfunname") = txt_FunName.Text
        If v_txt_FunRoot.IndexOf("\") > -1 Then v_txt_FunRoot = v_txt_FunRoot.Replace("\", "/")
        Me.ViewState("txtfunroot") = v_txt_FunRoot
        Me.ViewState("listparentfun") = If(list_ParentFun.SelectedValue = "", "0", list_ParentFun.SelectedValue)
        Me.ViewState("chkvalid") = If(CHK_VALID.Checked, "Y", "N")
        Me.ViewState("listsort") = If(hide_Sort.Value = "", tmpSort, CInt(hide_Sort.Value)) + 1 '新排序選取位置
        Me.ViewState("txtnote") = txt_Note.Text
        Me.ViewState("adds") = If(chk_Option.Items.Item(0).Selected, "1", "0")
        Me.ViewState("del") = If(chk_Option.Items.Item(1).Selected, "1", "0")
        Me.ViewState("mod") = If(chk_Option.Items.Item(2).Selected, "1", "0")
        Me.ViewState("sech") = If(chk_Option.Items.Item(3).Selected, "1", "0")
        Me.ViewState("prnt") = If(chk_Option.Items.Item(4).Selected, "1", "0")
        Me.ViewState("ISREPORT") = If(chk_Option.Items.Item(5).Selected, "Y", Convert.DBNull)
        txt_FunName.Text = v_txt_FunRoot

        If hide_FunID.Value = "" Then   '新增
            tmpFunID = INSERT_IDFUNCTION() '新增
            list_Sort.Items.FindByValue("new").Value = Convert.ToString(tmpFunID)
            hide_FunID.Value = Convert.ToString(tmpFunID)
        Else
            '修改
            Call UPDATE_IDFUNCTION(CInt(hide_FunID.Value))

            '更新群組功能
            If CHK_VALID.Checked Then
                If chk_Option.Items.Item(0).Selected = False Then UPDATE_AUTHGROUPFUN(CInt(hide_FunID.Value), "Adds")
                If chk_Option.Items.Item(1).Selected = False Then UPDATE_AUTHGROUPFUN(CInt(hide_FunID.Value), "Del")
                If chk_Option.Items.Item(2).Selected = False Then UPDATE_AUTHGROUPFUN(CInt(hide_FunID.Value), "Mod")
                If chk_Option.Items.Item(3).Selected = False Then UPDATE_AUTHGROUPFUN(CInt(hide_FunID.Value), "Sech")
                If chk_Option.Items.Item(4).Selected = False Then UPDATE_AUTHGROUPFUN(CInt(hide_FunID.Value), "Prnt")
                'If chk_Option.Items.Item(5).Selected=False Then Update_AuthGroupFun(CInt(hide_FunID.Value), "ISREPORT")
            End If

            '停用，且要刪除相關連結
            If Not CHK_VALID.Checked AndAlso CHK_DEL_REL_1.Checked Then
                DEL_AUTHGROUPFUN(CInt(hide_FunID.Value))            '刪除群組中的功能項
                REMOVE_AUTHPLANFUNCTION(CInt(hide_FunID.Value))     '刪除計畫中的功能項
                DEL_AUTHACCRWFUN(CInt(hide_FunID.Value))            '刪除使用者的功能項
            End If
        End If

        '當父項改變類型時，子項一併改變
        Dim i_hide_Subs As Integer = CInt(If(hide_Subs.Value <> "", hide_Subs.Value, "0"))
        If i_hide_Subs > 0 AndAlso hide_Kind.Value <> list_MainMenu.SelectedValue Then
            UPDATE_IDFUNCTION(hide_FunID.Value, list_MainMenu.SelectedValue)
        End If

        '更新排序
        Dim newSort As Integer = Val(ViewState("listsort")) - 1
        Dim newSort2 As Integer = list_Sort.Items.Count - 1
        If tmpSort <> -1 Then
            If tmpSort <> newSort Then
                Dim tmpText As String = list_Sort.Items(tmpSort).Text
                Dim tmpValue As String = list_Sort.Items(tmpSort).Value
                list_Sort.Items.RemoveAt(tmpSort)
                If newSort > newSort2 Then newSort = newSort2
                If (newSort < 0) Then newSort = -1
                list_Sort.Items.Insert(newSort, New ListItem(tmpText, tmpValue))
                list_Sort.Items(newSort).Selected = True
            End If
            For i As Integer = 0 To list_Sort.Items.Count - 1
                If Not list_Sort.Items.Item(i).Selected Then
                    UPDATE_IDFUNCTION(CInt(list_Sort.Items.Item(i).Value), i + 1)
                End If
            Next

            ''判斷是否加入所有計畫
            'If chk_AutoAdd.Checked=True Then
            '    Insert_AuthPlanFunction(CInt(hide_FunID.Value))
            'End If

            ''判斷是否加入所有計畫
            'If Me.chk_RemoveAll.Checked=True Then
            '    Remove_AuthPlanFunction(CInt(hide_FunID.Value))
            'End If

            '判斷是否 加入 所有計畫
            If TPlanIDFun.Checked Then
                Call Insert_AuthPlanFunction(CInt(hide_FunID.Value))
            End If

            '判斷是否 移除 所有計畫
            If TPlanIDFunDEL.Checked Then Call Insert_AuthPlanFunction2(CInt(hide_FunID.Value))
        Else
            intCntSave = 0
            Common.MessageBox(Me, "請重新點選功能排序")
            Return
        End If

        'Try

        '    'Call TIMS.CloseDbConn(objConn)
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    Return
        'End Try

        If intCntSave = 1 Then
            Call Clear_Items() '清除資料
            Call SSearch1("")

            tb_View.Visible = True
            tb_Edit.Visible = False
            Common.MessageBox(Me, "儲存成功!")
        End If
    End Sub

    '儲存
    Private Sub Btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click
        Call SaveData1()
    End Sub

    Private Sub List_MainMenu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_MainMenu.SelectedIndexChanged
        Dim dt_Function As DataTable = Nothing
        Dim dt_Sort As DataTable = Nothing

        Dim v_listMainMenu As String = TIMS.GetListValue(list_MainMenu)
        dt_Function = GET_IDFUNCTION(v_listMainMenu, "0", "")
        If dt_Function IsNot Nothing Then
            Call Show_ListParentFun(dt_Function, "")
            Call Show_ListSort(dt_Function, "")
        End If

        Dim i_hide_Subs As Integer = CInt(If(hide_Subs.Value <> "", hide_Subs.Value, "0"))
        If i_hide_Subs > 0 Then list_ParentFun.Enabled = False

        'If dt_Function IsNot Nothing Then dt_Function.Dispose()
        'If dt_Sort IsNot Nothing Then dt_Sort.Dispose()
    End Sub

    Private Sub List_ParentFun_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_ParentFun.SelectedIndexChanged
        Dim dt_Function As DataTable = Nothing
        Dim v_listMainMenu As String = TIMS.GetListValue(list_MainMenu)
        Dim v_listParentFun As String = TIMS.GetListValue(list_ParentFun)
        If v_listParentFun = "" Then
            dt_Function = GET_IDFUNCTION(v_listMainMenu, "0", "")
        Else
            dt_Function = GET_IDFUNCTION(v_listMainMenu, "1", v_listParentFun)
        End If
        Call Show_ListSort(dt_Function, "")
        'If dt_Function IsNot Nothing Then dt_Function.Dispose()
    End Sub

    Private Sub List_MainMenu2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_MainMenu2.SelectedIndexChanged
        Call SSearch1("")
    End Sub

    Private Sub List_MainMenu3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_MainMenu3.SelectedIndexChanged
        'list_MainMenu3.SelectedValue
        Dim v_list_MainMenu3 As String = TIMS.GetListValue(list_MainMenu3)
        Call SSearch1(v_list_MainMenu3)
    End Sub

    Private Sub BtnTest1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest1.Click
        Dim sPath1 As String = ""
        '==請選擇==
        If list_MainMenu.SelectedValue <> cst_v請選擇 _
            AndAlso list_MainMenu.SelectedValue <> "" _
            AndAlso txt_FunRoot.Text = "" Then

            sPath1 &= list_MainMenu.SelectedValue & "/"
            If list_ParentFun.SelectedValue <> "" Then
                sPath1 &= "00/" & list_MainMenu.SelectedValue & "_00_001.aspx"
            Else
                sPath1 &= list_MainMenu.SelectedValue & "_00_001.aspx"
            End If
            txt_FunRoot.Text = sPath1
        End If

        'hide_FunID
        '顯示 '已存在的計畫
        SPTPlanIDhave.Style.Item("display") = "inline"
        '已存在的計畫
        Call Show_Funid_TPlan(hide_FunID.Value)

        '顯示 '已存在的群組
        SpAuthGroup.Style.Item("display") = "inline"
        '已存在的群組
        Call Show_Funid_Group(hide_FunID.Value)
    End Sub

    '使用功能移除
    Private Sub Btn_use1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_use1.Click
        If TPlanIDFunDEL.Checked AndAlso SPTPlanIDhave.Style.Item("display") = "inline" Then
            For i As Integer = 0 To cblTPlanIDhave.Items.Count - 1
                If cblTPlanIDhave.Items(i).Selected Then
                    For j As Integer = 1 To cblTPlanID3.Items.Count - 1
                        If cblTPlanID3.Items(j).Value = cblTPlanIDhave.Items(i).Value Then
                            cblTPlanID3.Items(j).Selected = True
                            Exit For
                        End If
                    Next
                Else
                    For j As Integer = 1 To cblTPlanID3.Items.Count - 1
                        If cblTPlanID3.Items(j).Value = cblTPlanIDhave.Items(i).Value Then
                            cblTPlanID3.Items(j).Selected = False
                            Exit For
                        End If
                    Next
                End If
            Next
            SPTPlanID3.Style.Item("display") = "inline"
        Else
            Common.MessageBox(Me, "未勾選移除功能!!或未使用測試功能!!")
            Return 'Exit Sub
        End If
    End Sub

    '測試群組
    Protected Sub Btn_use3_Click(sender As Object, e As EventArgs) Handles Btn_use3.Click
        SpAuthGroup.Style.Item("display") = "inline"

        Dim sql As String = "SELECT * FROM Auth_GroupFun WHERE FunID=@FunID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("FunID", SqlDbType.VarChar).Value = hide_FunID.Value
            dt.Load(.ExecuteReader())
        End With
        With cblAuthGroup
            For i As Integer = 0 To .Items.Count - 1
                .Items(i).Selected = False
            Next
        End With
        If dt.Rows.Count > 0 Then
            With cblAuthGroup
                For i As Integer = 0 To .Items.Count - 1
                    If .Items(i).Value <> "" Then
                        If dt.Select("GID='" & .Items(i).Value & "'").Length > 0 Then
                            .Items(i).Selected = True
                        End If
                    End If
                Next
            End With
        End If
    End Sub

    '使用群組增減
    Protected Sub Btn_use2_Click(sender As Object, e As EventArgs) Handles Btn_use2.Click
        Dim sGroup As String = ""

        Dim sql As String = ""
        sql = " SELECT * FROM Auth_GroupFun WHERE FunID=@FunID AND GID=@GID "
        Dim sCmd As New SqlCommand(sql, objconn)
        sql = " DELETE Auth_GroupFun WHERE FunID=@FunID AND GID=@GID "
        Dim dCmd As New SqlCommand(sql, objconn)
        sql = ""
        sql &= " INSERT INTO Auth_GroupFun(GID,FUNID,ADDS,MOD,DEL,SECH,PRNT,MODIFYACCT,MODIFYDATE) "
        sql &= " VALUES (@GID,@FUNID,0,0,0,0,0,@MODIFYACCT,GETDATE()) "
        Dim iCmd As New SqlCommand(sql, objconn)

        If SpAuthGroup.Style.Item("display") = "inline" Then
            With cblAuthGroup
                For i As Integer = 0 To .Items.Count - 1
                    If .Items(i).Value <> "" Then
                        Dim GIDValue As String = .Items(i).Value
                        If .Items(i).Selected Then
                            Dim dt As New DataTable
                            With sCmd
                                .Parameters.Clear()
                                .Parameters.Add("FunID", SqlDbType.VarChar).Value = hide_FunID.Value
                                .Parameters.Add("GID", SqlDbType.VarChar).Value = GIDValue
                                dt.Load(.ExecuteReader())
                            End With
                            If dt.Rows.Count = 0 Then
                                With iCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("FunID", SqlDbType.VarChar).Value = hide_FunID.Value
                                    .Parameters.Add("GID", SqlDbType.VarChar).Value = GIDValue
                                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                    .ExecuteNonQuery()
                                End With
                            End If
                        Else
                            With dCmd
                                .Parameters.Clear()
                                .Parameters.Add("FunID", SqlDbType.VarChar).Value = hide_FunID.Value
                                .Parameters.Add("GID", SqlDbType.VarChar).Value = GIDValue
                                .ExecuteNonQuery()
                            End With
                        End If
                        If sGroup <> "" Then sGroup &= ","
                        sGroup &= .Items(i).Value
                    End If
                Next
            End With
        End If
        If sGroup = "" Then
            Common.MessageBox(Me, "未勾選群組!!或未使用測試功能!!")
            Return 'Exit Sub
        Else
            Common.MessageBox(Me, "已儲存!")
        End If
    End Sub

    'Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
    '    Dim v_list_MainMenu3 As String=TIMS.GetListValue(list_MainMenu3)
    '    '查詢 SQL
    '    Call sSearch1(v_list_MainMenu3)
    'End Sub

#Region "(No Use)"

    '功能類別對照 取得中文名稱
    'Function Get_MainMenuName(ByVal tmpCode As String) As String
    '    Dim rst As String=""

    '    Select Case UCase(tmpCode)
    '        Case "TC"
    '            rst="訓練機構管理"
    '        Case "SD"
    '            rst="學員動態管理"
    '        Case "CP"
    '            rst="查核/績效管理"
    '        Case "TR"
    '            rst="訓練需求管理"
    '        Case "CM"
    '            rst="訓練經費控管"
    '        Case "SYS"
    '            rst="系統管理"
    '        Case "FAQ"
    '            rst="問答集"
    '        Case "OB"
    '            rst="委外訓練管理"
    '        Case "SE"
    '            rst="技能檢定管理"
    '        Case "EXAM"
    '            rst="甄試管理"
    '        Case "SV"
    '            rst="問卷管理"
    '        Case "OO"
    '            rst="其他系統"
    '        Case Else
    '            rst=tmpCode
    '    End Select

    '    Return rst
    'End Function

#End Region

End Class