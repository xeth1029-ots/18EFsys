Partial Class SD_05_026
    Inherits AuthBasePage

    'Stud_EnterTemp 'Stud_EnterType
    'Stud_EnterTemp2 'Stud_EnterType2
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        PageControler1.PageDataGrid = DataGrid2
        'Button2.Style("display")="none"  'Button1.Attributes.Add("onclick")="return chkdata();"
        Button1.Attributes("onclick") = "return chkdata();"

        If Not IsPostBack Then
            'DistID=TIMS.Get_DistID(DistID)
            'TPlanID=TIMS.Get_TPlan(TPlanID)
            ShowDatatable.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Call sSearch1()
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Call TIMS.WriteTraceLog(Me, ex)
            Dim ExcepMsg1 As String = String.Concat(TIMS.cst_ErrorMsg9, ",", ex.Message)
            Common.MessageBox(Me, ExcepMsg1) '"資料庫效能異常，請重新查詢")
        End Try
    End Sub

    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid2)

        'Dim sql As String=""
        'Dim dt As DataTable
        Dim str As String = "" '總集合
        Dim str_1 As String = ""
        Dim str_2 As String = ""

        '消除單引號。
        If Name.Text <> "" Then Name.Text = TIMS.ClearSQM(Name.Text)
        'If Name.Text <> "" Then Name.Text=Trim(Name.Text)
        'If Name.Text.IndexOf("'") > -1 Then Name.Text=Replace(Name.Text, "'", "")
        If IDNO.Text <> "" Then IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(UCase(IDNO.Text)))
        RelEnterDate1.Text = TIMS.Cdate3(RelEnterDate1.Text)
        RelEnterDate2.Text = TIMS.Cdate3(RelEnterDate2.Text)
        Dim NoRdate12 As Boolean = Not (RelEnterDate1.Text <> "" AndAlso RelEnterDate2.Text <> "")
        If IDNO.Text = "" AndAlso Name.Text = "" AndAlso NoRdate12 Then
            Common.MessageBox(Me, "至少輸入一查詢條件或日期區間!!")
            Exit Sub
        End If

        If IDNO.Text <> "" Then
            str &= " AND a.IDNO='" & IDNO.Text & "'" & vbCrLf
            str_1 &= " AND se.IDNO='" & IDNO.Text & "'" & vbCrLf
            str_2 &= " AND se2.IDNO='" & IDNO.Text & "'" & vbCrLf
        End If
        If Name.Text <> "" Then
            str &= " AND a.Name LIKE N'" & Name.Text & "'" & vbCrLf
            str_1 &= " AND se.Name LIKE  N'" & Name.Text & "'" & vbCrLf
            str_2 &= " AND se2.Name LIKE  N'" & Name.Text & "'" & vbCrLf
        End If
        If RelEnterDate1.Text <> "" Then
            str &= " AND a.RelEnterDate >= " & TIMS.To_date(RelEnterDate1.Text) & vbCrLf '★
            str_1 &= " AND sy.RelEnterDate >= " & TIMS.To_date(RelEnterDate1.Text) & vbCrLf '★
            str_2 &= " AND sy2.RelEnterDate >= " & TIMS.To_date(RelEnterDate1.Text) & vbCrLf '★
        End If
        If RelEnterDate2.Text <> "" Then
            str &= " AND a.RelEnterDate <= " & TIMS.To_date(RelEnterDate2.Text) & vbCrLf '★
            str_1 &= " AND sy.RelEnterDate <= " & TIMS.To_date(RelEnterDate2.Text) & vbCrLf '★
            str_2 &= " AND sy2.RelEnterDate <=  " & TIMS.To_date(RelEnterDate2.Text) & vbCrLf '★
        End If

        Dim sql As String = ""
        sql &= " SELECT DISTINCT a.Name,a.IDNO,a.Birthday,CONVERT(varchar, a.RelEnterDate, 111) Eterdate,a.Admission,a.EnterChannel" & vbCrLf
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql &= " ,a.signUpStatus,a.ENTERPATH2,cc.DistName,cc.Years,cc.Tplanid,cc.PlanName,cc.orgName,ke.JobName,ke.TrainName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDate, 111) + '~' + CONVERT(varchar, cc.FTDate, 111) AS SFdate" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " JOIN VIEW_TRAINTYPE ke ON ke.TMID=cc.TMID" & vbCrLf
        sql &= " JOIN ( SELECT 'n' EnterType ,sy.SETID ,sy.EnterDate ,sy.OCID1 ,sy.RelEnterDate ,se.Name" & vbCrLf
        sql &= "   ,se.IDNO,se.Birthday ,ss.Admission ,sy.EnterChannel ,NULL signUpStatus" & vbCrLf
        sql &= "   ,sy.ENTERPATH2 "
        sql &= "   FROM STUD_ENTERTEMP se WITH(NOLOCK)" & vbCrLf
        sql &= "   JOIN STUD_ENTERTYPE sy WITH(NOLOCK) ON se.SETID=sy.SETID" & vbCrLf
        sql &= "   LEFT JOIN Stud_SelResult ss ON ss.SETID=sy.SETID AND ss.EnterDate=sy.EnterDate AND ss.SerNum=sy.SerNum" & vbCrLf
        sql &= "   WHERE 1=1" & vbCrLf
        sql &= str_1

        If IDNO.Text <> "" OrElse Name.Text <> "" Then
            'E網優先，排除內部報名。
            sql &= "   AND NOT EXISTS ( SELECT 1" & vbCrLf
            sql &= "  FROM STUD_ENTERTYPE2 c WITH(NOLOCK)" & vbCrLf
            sql &= "  JOIN STUD_ENTERTEMP2 c2 WITH(NOLOCK) ON c2.eSETID=c.eSETID" & vbCrLf
            If IDNO.Text <> "" Then sql &= " AND c2.IDNO ='" & IDNO.Text & "'" & vbCrLf
            If Name.Text <> "" Then sql &= " AND c2.Name =N'" & Name.Text & "'" & vbCrLf
            sql &= "  WHERE c.SETID=sy.SETID AND c.EnterDate=sy.EnterDate AND c.SerNum=sy.SerNum" & " )" & vbCrLf
        End If

        sql &= "   UNION" & vbCrLf
        sql &= "   SELECT 'e' EnterType ,sy2.eSETID ,sy2.EnterDate ,sy2.OCID1 ,sy2.RelEnterDate ,se2.Name" & vbCrLf
        sql &= "   ,se2.IDNO,se2.Birthday ,null Admission,null EnterChannel,sy2.signUpStatus" & vbCrLf
        sql &= "   ,null enterpath2"
        sql &= "   FROM STUD_ENTERTEMP2 se2 WITH(NOLOCK)" & vbCrLf
        sql &= "   JOIN STUD_ENTERTYPE2 sy2 WITH(NOLOCK) on se2.eSETID=sy2.eSETID" & vbCrLf
        sql &= "   WHERE 1=1" & vbCrLf
        sql &= str_2 & " ) a on a.OCID1=cc.OCID" & vbCrLf
        sql &= " WHERE cc.NotOpen='N'" & vbCrLf '排除未開訓班級
        'sql &= " and ROWNUM <= 1000" & vbCrLf

        '排除已結訓班級
        Select Case rblFTDate.SelectedValue
            Case "Y"
                '已結訓班級
                sql &= " AND cc.FTDate <= GETDATE()" & vbCrLf
            Case "N"
                '排除已結訓班級
                sql &= " AND cc.FTDate >= GETDATE()" & vbCrLf
                'Case Else'A
        End Select

        sql &= str
        'sql &= " AND upper(a.IDNO)=upper('k122412880')" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        msg.Visible = True
        ShowDatatable.Visible = False
        Searchtable.Visible = True

        If TIMS.dtNODATA(dt) Then Return

        msg.Text = ""
        'msg.Visible=True
        ShowDatatable.Visible = True
        Searchtable.Visible = False

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub Button5_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.ServerClick, Button4.ServerClick
        ShowDatatable.Visible = False
        Searchtable.Visible = True
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LTMID As Label = e.Item.FindControl("LTMID")
                Dim LEnterType As Label = e.Item.FindControl("LEnterType")
                Dim LAdmission As Label = e.Item.FindControl("LAdmission")
                Dim LEnterChannel As Label = e.Item.FindControl("LEnterChannel")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                LTMID.Text = Convert.ToString(drv("TrainName"))
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(Convert.ToString(drv("TPlanID"))) > -1 Then LTMID.Text = Convert.ToString(drv("JobName"))

                If Convert.ToString(drv("signUpStatus")) <> "" Then
                    'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                    Select Case Convert.ToString(drv("signUpStatus"))
                        Case "0"
                            LEnterType.Text = "收件完成"
                        Case "1"
                            LEnterType.Text = "報名成功"
                        Case "2"
                            LEnterType.Text = "報名失敗"
                        Case "3"
                            LEnterType.Text = "正取"
                        Case "4"
                            LEnterType.Text = "備取"
                        Case "5"
                            LEnterType.Text = "未錄取"
                        Case Else
                            LEnterType.Text = "-"
                    End Select
                Else    '如果報名狀態是空值
                    Select Case Convert.ToString(drv("EnterChannel"))
                        Case "2", "3"
                            LEnterType.Text = "報名成功"
                        Case Else
                            LEnterType.Text = "-"
                    End Select
                End If
                'LAdmission 錄取狀態
                Select Case Convert.ToString(drv("Admission"))
                    Case "Y"
                        LAdmission.Text = "錄取"
                    Case "N"
                        LAdmission.Text = "不錄取"
                    Case Else
                        LAdmission.Text = "-"
                End Select

                '報名管道
                Dim strEnterP2 As String = ""
                Select Case Convert.ToString(drv("ENTERPATH2"))
                    Case "P"
                        strEnterP2 = "(專案核定報名)"
                End Select
                Dim strEnter As String = ""
                Select Case Convert.ToString(drv("EnterChannel"))
                    Case "1"
                        strEnter = "網路"
                    Case "2"
                        strEnter = "現場"
                    Case "3"
                        strEnter = "通訊"
                    Case "4"
                        strEnter = "推介"
                    Case Else
                        strEnter = "-"
                End Select
                LEnterChannel.Text = strEnterP2 & strEnter
        End Select
    End Sub
End Class