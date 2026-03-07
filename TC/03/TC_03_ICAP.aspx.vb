Public Class TC_03_ICAP
    Inherits AuthBasePage

    Dim fg_nodata1 As Boolean = True '預設查無資料:true

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lblMsg.Text = ""
        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    Sub CCreate1()
        Dim rqICAPNUM1 As String = TIMS.sUtl_GetRqValue(Me, "ICAPNUM1")

        Call Clear_DATA()

        fg_nodata1 = True '預設查無資料:true
        lblMsg.Text = TIMS.cst_NODATAMsg1

        Dim icap_data As tw.gov.wda.wltims1.iCAPData = Nothing
        Try
            icap_data = TIMS.GET_iCAPData(rqICAPNUM1)
        Catch ex As Exception
            lblMsg.Text = ex.Message
            Return
        End Try
        If icap_data Is Nothing Then Return

        If Convert.ToString(icap_data.ErrMsg) <> "" Then
            lblMsg.Text = icap_data.ErrMsg
            Return
        End If

        Call SHOW_ICAP_DATA(icap_data)

        If fg_nodata1 Then
            '預設查無資料:true Common.MessageBox2(Me, TIMS.cst_NODATAMsg1)
            '查無資料:true
            Return
        End If

        '跑完程序，表示有資料 fg_nodata1 = False
        lblMsg.Text = ""
    End Sub

    Sub Clear_DATA()
        Dim dtcp2 As DataTable = CreateC2dt()
        ListView1.DataSource = dtcp2
        ListView1.DataBind()

        labCLASS_ID.Text = "" ' ic1.CLASS_ID
        labCASE_ID.Text = "" 'ic1.CASE_ID
        labCOMPANY.Text = "" 'ic1.COMPANY
        labC_ID.Text = "" 'ic1.C_ID
        labCCNAME.Text = "" 'ic1.NAME
        labTRAIN_COURSE_HOURS.Text = "" 'ic1.TRAIN_COURSE_HOURS
        labCLASS_LEVEL.Text = "" 'ic1.CLASS_LEVEL
        labTARGET.Text = "" 'ic1.TARGET
        labESSENTIAL.Text = "" 'ic1.ESSENTIAL
    End Sub

    Function CreateC2dt() As DataTable
        Dim dtcp2 As New DataTable
        dtcp2.Columns.Add(New DataColumn("UNIT_SEQ"))
        dtcp2.Columns.Add(New DataColumn("UNIT_NAME"))
        dtcp2.Columns.Add(New DataColumn("TEA_EXP"))
        dtcp2.Columns.Add(New DataColumn("SUP_EXP"))
        dtcp2.Columns.Add(New DataColumn("TEACH_MATERIALS"))
        dtcp2.Columns.Add(New DataColumn("TEACH_EQUIPMENT"))
        dtcp2.Columns.Add(New DataColumn("ACTUAL_TEACHER"))
        dtcp2.Columns.Add(New DataColumn("TEACH_METHOD"))
        dtcp2.Columns.Add(New DataColumn("EVALUATION_METHOD"))
        dtcp2.Columns.Add(New DataColumn("OUTLINE"))
        Return dtcp2
    End Function

    Function Get_CP2DETAILdt(ic1 As tw.gov.wda.wltims1.Course) As DataTable
        Dim dtcp2 As DataTable = Nothing

        If ic1.DETAIL Is Nothing Then Return dtcp2

        If ic1.DETAIL.Length = 0 Then Return dtcp2

        'Dim dtcp2 As DataTable = CreateC2dt()
        dtcp2 = CreateC2dt()

        Dim iRow As Integer = 0
        For Each ic2 As tw.gov.wda.wltims1.Unit In ic1.DETAIL
            iRow += 1
            Dim cp2DR As DataRow = dtcp2.NewRow
            dtcp2.Rows.Add(cp2DR)
            cp2DR("UNIT_SEQ") = String.Concat("課程單元 ", iRow)
            cp2DR("UNIT_NAME") = ic2.UNIT_NAME
            cp2DR("TEA_EXP") = ic2.TEA_EXP
            cp2DR("SUP_EXP") = ic2.SUP_EXP
            cp2DR("TEACH_MATERIALS") = ic2.TEACH_MATERIALS
            cp2DR("TEACH_EQUIPMENT") = ic2.TEACH_EQUIPMENT
            cp2DR("ACTUAL_TEACHER") = ic2.ACTUAL_TEACHER
            cp2DR("TEACH_METHOD") = ic2.TEACH_METHOD
            cp2DR("EVALUATION_METHOD") = ic2.EVALUATION_METHOD
            cp2DR("OUTLINE") = ic2.OUTLINE
        Next

        Return dtcp2
    End Function

    Private Sub SHOW_ICAP_DATA(icap_data As tw.gov.wda.wltims1.iCAPData)
        If icap_data Is Nothing Then Return

        If icap_data.iCAPData1 Is Nothing Then Return

        If icap_data.iCAPData1.Length = 0 Then Return

        Dim ic1 As tw.gov.wda.wltims1.Course = icap_data.iCAPData1(0)
        labCLASS_ID.Text = ic1.CLASS_ID
        labCASE_ID.Text = ic1.CASE_ID
        labCOMPANY.Text = ic1.COMPANY
        labC_ID.Text = ic1.C_ID
        labCCNAME.Text = ic1.NAME
        labTRAIN_COURSE_HOURS.Text = ic1.TRAIN_COURSE_HOURS
        labCLASS_LEVEL.Text = ic1.CLASS_LEVEL
        labTARGET.Text = ic1.TARGET
        labESSENTIAL.Text = ic1.ESSENTIAL

        If ic1.DETAIL Is Nothing Then Return

        If ic1.DETAIL.Length = 0 Then Return

        Dim dtcp2 As DataTable = Get_CP2DETAILdt(ic1)
        'Dim dtcp2 As DataTable = TIMS.ToDataTableUnit(ic1.DETAIL)

        ListView1.DataSource = dtcp2
        ListView1.DataBind()

        '跑完程序，表示有資料 fg_nodata1 = False
        fg_nodata1 = False
    End Sub

    'Public Shared Function GET_iCAPData(s_ICAPNUM1 As String) As tw.gov.wda.wltims1.iCAPData
    '    Dim rst As tw.gov.wda.wltims1.iCAPData = Nothing
    '    'WebRequest物件如何忽略憑證問題
    '    System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
    '    'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
    '    System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12 '3072
    '    'Dim rqICAPNUM1 As String = TIMS.sUtl_GetRqValue(Me, "ICAPNUM1")
    '    Const cst_s_AUTH As String = "ce130f098d90e52491a8097fcd119b4d"
    '    Dim icap_data As tw.gov.wda.wltims1.iCAPData = Nothing
    '    Using iCAPws1 As New tw.gov.wda.wltims1.iCapWebS1
    '        icap_data = iCAPws1.Ws_onlineClass2(cst_s_AUTH, s_ICAPNUM1)
    '        If icap_data Is Nothing Then Return rst
    '    End Using
    '    'If Convert.ToString(icap_data.ErrMsg) <> "" Then
    '    '    lblMsg.Text = icap_data.ErrMsg
    '    '    Return rst
    '    'End If
    '    Return rst
    'End Function

End Class