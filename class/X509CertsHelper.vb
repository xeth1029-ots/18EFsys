Imports System.Security.Cryptography.X509Certificates
Public Class X509CertsHelper

    Private logger As ILog = LogManager.GetLogger(GetType(X509CertsHelper))

    Private x509 As X509Certificate = Nothing

    ''' <summary>
    ''' 傳入 Base64 編碼的 X509 憑證資料建構 X509CertsHelper
    ''' </summary>
    ''' <param name="sCert"></param>
    Public Sub New(ByVal sCert As String)

        Dim certData As Byte()
        Try
            certData = Convert.FromBase64String(sCert)
        Catch ex As Exception
            Throw New ArgumentException("參數 sCert 不是合法的 Base64 格式字串")
        End Try

        Try
            x509 = New X509Certificate(certData)
        Catch ex As Exception
            Throw New ArgumentException($"以參數 sCert 起始 X509Certificate 失敗, {ex.Message}", ex)
        End Try

    End Sub

    ''' <summary>
    ''' 憑證的主旨辨別名稱
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Subject As String
        Get
            If IsNothing(x509) Then Return Nothing

            Return x509.Subject
        End Get
    End Property

    ''' <summary>
    ''' 憑證內部序號
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property SerialNumber As String
        Get
            If IsNothing(x509) Then Return Nothing

            Return x509.GetSerialNumberString
        End Get
    End Property

    ''' <summary>
    ''' 憑證授權(簽發)單位名稱
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Issuer As String
        Get
            If IsNothing(x509) Then Return Nothing

            Return x509.Issuer
        End Get
    End Property

    ''' <summary>
    ''' 憑證有效日期(起)
    ''' <para>格式: 2020/12/16 下午 11:59:59</para>
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property NotBefore As String
        Get
            If IsNothing(x509) Then Return Nothing

            Return x509.GetEffectiveDateString
        End Get
    End Property

    ''' <summary>
    ''' 憑證有效日期(迄)
    ''' <para>格式: 2015/12/16 下午 02:38:47</para>
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property NotAfter As String
        Get
            If IsNothing(x509) Then Return Nothing

            Return x509.GetExpirationDateString
        End Get
    End Property

    ''' <summary>
    ''' 憑證核發對象名稱
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Name As String
        Get
            If IsNothing(x509) Then Return Nothing

            ' C=TW, CN=ｏｏｏ, SERIALNUMBER=000000000000000
            Dim str As String = x509.GetName
            Dim tokens As String() = Split(str, ", ")
            If IsNothing(tokens) OrElse tokens.Length < 3 Then Return Nothing

            ' tokens(1): CN=ｏｏｏ
            Dim parts As String() = Split(tokens(1), "=")
            If IsNothing(parts) OrElse parts.Length < 2 Then Return Nothing

            Return parts(1)
        End Get
    End Property

    ''' <summary>
    ''' 憑證簽發卡號
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CardID As String
        Get
            If IsNothing(x509) Then Return Nothing

            ' C=TW, CN=ｏｏｏ, SERIALNUMBER=000000000000000
            Dim str As String = x509.GetName
            Dim tokens As String() = Split(str, ", ")
            If IsNothing(tokens) OrElse tokens.Length < 3 Then Return Nothing

            ' tokens(2): SERIALNUMBER=000000000000000
            Dim parts As String() = Split(tokens(2), "=")
            If IsNothing(parts) OrElse parts.Length < 2 Then Return Nothing

            Return parts(1)
        End Get
    End Property

    ''' <summary>
    ''' 此憑證是否有效
    ''' <para>當前系統日期是否在憑證效期內</para>
    ''' </summary>
    ''' <returns></returns>
    Public Function IsExpired() As Boolean
        Dim expired As Boolean = True
        Try
            Dim dStart As DateTime = Convert.ToDateTime(Me.NotBefore) 'DateTime.ParseExact(Me.NotBefore, "yyyy/MM/dd tt HH:mm:ss", Nothing)

            Dim dEnd As DateTime = Convert.ToDateTime(Me.NotAfter) 'DateTime.ParseExact(Me.NotAfter, "yyyy/MM/dd tt HH:mm:ss", Nothing)

            Dim now As DateTime = DateTime.Now

            If now.CompareTo(dStart) >= 0 AndAlso now.CompareTo(dEnd) <= 0 Then expired = False

        Catch ex As Exception
            logger.Warn("IsExpired: " & ex.Message, ex)
        End Try
        Return expired
    End Function


    ''' <summary>
    ''' 是否為廢止憑證(CRL)
    ''' </summary>
    ''' <returns></returns>
    Public Function IsCRL() As Boolean

        Return False
    End Function

    Public Overrides Function ToString() As String
        Return $"[Subject] {Subject}{vbCrLf}[Issuer] {Issuer}{vbCrLf}[NotBefore] {NotBefore}{vbCrLf}[NotAfter] {NotAfter}{vbCrLf}[SerialNumber] {SerialNumber}{vbCrLf}[Name] {Name}{vbCrLf}[CardID] {CardID}{vbCrLf}[Expired] {IsExpired()}{vbCrLf}"
    End Function

End Class
