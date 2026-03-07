Partial Class SD_05_036c
    Inherits System.Web.UI.Page

    Const cst_SD_05_036c_NEW1 As String = "NEW1"
    Const cst_SD_05_036c_ZIPCODE As String = "SD_05_036c_ZIPCODE"
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        Button2.Attributes("onclick") = "javascript:return result();"
        Button3.Attributes("onclick") = "MoveItem(1);"
        Button4.Attributes("onclick") = "MoveItem(2);"
        Button5.Attributes("onclick") = "MoveItem(3);"
        Button6.Attributes("onclick") = "MoveItem(4);"

        If Not IsPostBack Then
            Table3.Visible = False
            Button2.Visible = False '送出
            Call SSearch1()
        End If
    End Sub

    '取得該序號目前的災害地區
    Function Get_rqADID_ZIPCODE(ByVal ADID As String) As String
        Dim rst As String = ""
        Select Case ADID
            Case cst_SD_05_036c_NEW1 '新增一筆資料
            Case Else
                rst = TIMS.GetDISASTER2(ADID, objconn)
        End Select
        Return rst
    End Function

    Sub SSearch1()
        Dim flagCanUseSess As Boolean = False
        If Not Session(cst_SD_05_036c_ZIPCODE) Is Nothing Then
            If Convert.ToString(Session(cst_SD_05_036c_ZIPCODE)) <> "" Then
                flagCanUseSess = True
            End If
        End If

        Dim rqADID As String = TIMS.sUtl_GetRqValue(Me, "ADID")
        Dim rqZIPCODE As String = ""
        If rqADID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        If flagCanUseSess Then
            rqZIPCODE = Convert.ToString(Session(cst_SD_05_036c_ZIPCODE))
        Else
            rqZIPCODE = TIMS.sUtl_GetRqValue(Me, "ZIPCODE")
        End If
        If rqZIPCODE <> "" Then
            HID_ZIPCODES1.Value = rqZIPCODE
        Else
            HID_ZIPCODES1.Value = Get_rqADID_ZIPCODE(rqADID)
        End If

        HidADID.Value = ""
        Select Case rqADID
            Case cst_SD_05_036c_NEW1 '新增一筆資料
            Case Else
                HidADID.Value = rqADID
        End Select

        Call SHOWListBox1()
        Call SHOWListBox2(HidADID.Value)
    End Sub

    Sub SHOWListBox1()
        Dim sql As String = ""
        sql &= " SELECT ZIPCODE, concat(CTNAME,'-',ZNAME) ZIPNAME2" & vbCrLf
        sql &= " FROM VIEW_ZIPNAME" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        sql &= " ORDER BY ZIPCODE" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt2.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!")
            Table3.Visible = False
            Exit Sub
        End If
        Table3.Visible = True
        Button2.Visible = True '送出
        With ListBox1
            .DataSource = dt2
            .DataTextField = "ZIPNAME2"
            .DataValueField = "ZIPCODE"
            .DataBind()
        End With
    End Sub

    Sub SHOWListBox2(ByVal byADID As String)
        byADID = TIMS.ClearSQM(byADID)
        If byADID = "" Then Return

        HID_ZIPCODES1.Value = TIMS.CombiSQLIN(HID_ZIPCODES1.Value)
        Dim pms1 As New Hashtable
        Dim sql As String = ""
        sql &= " WITH WC1 AS (" & vbCrLf
        If HID_ZIPCODES1.Value = "" Then
            pms1.Add("ADID", byADID)
            sql &= " SELECT ZIPCODE FROM ADP_DISASTER2 WHERE ADID=@ADID" & vbCrLf
        Else
            sql &= " SELECT ZIPCODE FROM ID_ZIP WHERE ZIPCODE IN (" & HID_ZIPCODES1.Value & ")" & vbCrLf
        End If
        sql &= " )" & vbCrLf

        sql &= " SELECT a.ZIPCODE, concat(a.CTNAME,'-',a.ZNAME) ZIPNAME2" & vbCrLf
        sql &= " FROM VIEW_ZIPNAME a" & vbCrLf
        sql &= " JOIN WC1 b on b.ZIPCODE=a.ZIPCODE" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        sql &= " ORDER BY a.ZIPCODE" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
        If dt2.Rows.Count = 0 Then Exit Sub

        With ListBox2
            .DataSource = dt2
            .DataTextField = "ZIPNAME2"
            .DataValueField = "ZIPCODE"
            .DataBind()
        End With

        Dim ZIPCODEstr As String = ""
        If dt2.Rows.Count > 0 Then
            For Each dr As DataRow In dt2.Rows
                If Convert.ToString(dr("ZIPCODE")) <> "" Then
                    Dim tmpValue1 As String = String.Concat("'", dr("ZIPCODE"), "'")
                    If tmpValue1 <> "" AndAlso ZIPCODEstr.IndexOf(tmpValue1) = -1 Then '重複字過濾
                        ZIPCODEstr &= String.Concat(If(ZIPCODEstr <> "", ",", ""), tmpValue1)
                    End If
                End If
            Next
        End If
        HID_ZIPCODES1.Value = ZIPCODEstr

        For Each dr2 As DataRow In dt2.Rows
            If Not ListBox1.Items.FindByValue(dr2("ZIPCODE")) Is Nothing Then
                ListBox1.Items.Remove(ListBox1.Items.FindByValue(dr2("ZIPCODE")))
            End If
        Next

    End Sub

    '送出
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sessName1 As String = cst_SD_05_036c_ZIPCODE & TIMS.GetRnd6Eng()
        Session(sessName1) = HID_ZIPCODES1.Value

        '上層欄位要注意
        Dim coZIPCODES1 As String = Replace(HID_ZIPCODES1.Value, "'", "\'")
        Dim ssScript1 As String = ""
        ssScript1 &= "<script language=javascript>" & vbCrLf
        ssScript1 &= "   var mylabel=window.opener.document.getElementById('Label1');" & vbCrLf
        ssScript1 &= "   mylabel.innerHTML='選取完畢!';" & vbCrLf
        'ssScript1 &= "   window.opener.document.form1.hid_ZIPCODE.value='" & coZIPCODES1 & "';" & vbCrLf
        ssScript1 &= "   window.opener.document.form1.hid_sessName1.value='" & sessName1 & "';" & vbCrLf
        ssScript1 &= "   window.close();" & vbCrLf
        ssScript1 &= "</script>" & vbCrLf
        Common.RespWrite(Me, ssScript1)
    End Sub
End Class
