Partial Class SD_02_003_other
    Inherits AuthBasePage

    '挑選其他志願

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

        If Not IsPostBack Then create1()
    End Sub

    Sub create1()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        If rqOCID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WA1 AS ( " & vbCrLf
        sql &= "   SELECT 2 AS Wish, SETID, SerNum, EnterDate, WriteResult, OralResult, TotalResult, ExamNo " & vbCrLf
        sql &= "   FROM Stud_EnterType WHERE OCID2 = '" & rqOCID & "' " & vbCrLf
        sql &= "   UNION " & vbCrLf
        sql &= "   SELECT 3 AS Wish, SETID, SerNum, EnterDate, WriteResult, OralResult, TotalResult, ExamNo " & vbCrLf
        sql &= "   FROM Stud_EnterType WHERE OCID3 = '" & rqOCID & "' " & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WA2 AS ( " & vbCrLf
        sql &= "   SELECT a.Wish, a.SETID, a.SerNum, a.EnterDate, a.WriteResult, a.OralResult, a.TotalResult, a.ExamNo, b.NAME " & vbCrLf
        sql &= "   FROM WA1 a " & vbCrLf
        sql &= "   JOIN Stud_EnterTemp b ON b.SETID = a.SETID " & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT c.SETID ,c.ENTERDATE ,c.SERNUM ,c.OCID ,c.SUMOFGRAD ,c.APPLIEDSTATUS ,c.ADMISSION ,c.SELRESULTID ,c.TRNDTYPE " & vbCrLf
        sql &= " ,c.RID ,c.PLANID ,c.MODIFYACCT ,c.MODIFYDATE ,c.SELSORT ,c.NOTES2 ,d.Wish ,d.WriteResult ,d.OralResult ,d.TotalResult ,d.ExamNo ,d.NAME " & vbCrLf
        sql &= " ,dbo.FN_SELRESULTID(c.SELRESULTID,1) SELRESULTID1" & vbCrLf
        'sql &= " ,dbo.FN_SELRESULTID(c.SELRESULTID,2) SELRESULTID2" & vbCrLf
        sql &= " ,CASE WHEN d.Wish=2 THEN '二' WHEN d.Wish=3 THEN '三' END WISH_TXT" & vbCrLf
        sql &= " FROM STUD_SELRESULT c " & vbCrLf
        sql &= " JOIN WA2 d ON c.SETID = d.SETID AND c.SerNum = d.SerNum AND c.EnterDate = d.EnterDate " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND c.SelResultID <> '01' " & vbCrLf
        sql &= " AND ISNULL(c.Admission,'N') = 'N' " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            DataGrid1.Visible = False
            button1.Visible = False
            msg.Text = "查無資料!"
            Exit Sub
        End If
        DataGrid1.Visible = True
        button1.Visible = True
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                'Dim stSELRESULTID As String = ""
                'Select Case Convert.ToString(drv("SELRESULTID"))
                '    Case "01"
                '        stSELRESULTID = "正取"
                '    Case "02"
                '        stSELRESULTID = "備取"
                '    Case "03"
                '        stSELRESULTID = "未錄取"
                'End Select
                'e.Item.Cells(6).Text = stSELRESULTID
                'Dim stWish As String = ""
                'Select Case Convert.ToString(drv("Wish"))
                '    Case "2"
                '        stWish = "二"
                '    Case "3"
                '        stWish = "三"
                'End Select
                'e.Item.Cells(7).Text = stWish
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim rqSETID As String = TIMS.ClearSQM(Request("SETID"))
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        If rqSETID = "" Then
            Common.MessageBox(Me, "請勾選學員")
            Exit Sub
        End If
        If rqOCID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim all() As String = Split(rqSETID, ",", , CompareMethod.Text)
        For i As Integer = 0 To all.Length - 1
            Dim AllMember() As String = Split(all(i), "@#", , CompareMethod.Text)
            Dim sql As String = ""
            sql = ""
            sql &= " UPDATE Stud_SelResult"
            sql &= " SET OCID='" & rqOCID & "'"
            sql &= " ,ModifyAcct='" & sm.UserInfo.UserID & "'"
            sql &= " ,ModifyDate=getdate()"
            sql &= " WHERE SETID='" & AllMember(0) & "'"
            sql &= " and EnterDate=convert(datetime, '" & AllMember(1) & "', 111)"
            sql &= " and SerNum='" & AllMember(2) & "'"
            DbAccess.ExecuteNonQuery(sql, objconn)
        Next
        'Common.RespWrite(Me, "<script language=javascript>window.close();</script>")
        Common.RespWrite(Me, "<script language=javascript>opener.form1.submit(); window.close();</script>")
    End Sub
End Class