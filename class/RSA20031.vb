Imports System.Security.Cryptography
'Imports System.Security.Cryptography.ECDiffieHellmanCng
'Imports System.Web.Security
'Imports System.Security
'Imports System.Security.Cryptography
'https://dotblogs.com.tw/yc421206/archive/2012/06/25/73041.aspx
'https://dotblogs.com.tw/yc421206/2012/06/27/73097
'產生加密和解密金鑰
'https://msdn.microsoft.com/zh-tw/library/5e9ft273(v=vs.100).aspx
'http://big5.webasp.net/article/5/4098_print.htm
'http://www.blueshop.com.tw/board/FUM200410061525290EW/BRD20121005093915Q2Q.html
'http://blog.xuite.net/yan.kee/CSharp/19047933-%E7%94%A8VB.NET%E7%B7%A8%E5%AF%ABDES%E5%8A%A0%E5%AF%86
'http://note.artchiu.org/2009/02/06/%E7%94%A8-vbnet-%E7%B7%A8%E5%AF%AB-des-%E5%8A%A0%E5%AF%86%E7%A8%8B%E5%BA%8F/

Public Class RSA20031

    '取得8個亂數文字 (大小寫)
    Public Shared Function GetRnd8Eng() As String
        '(ASCII 字元碼)
        'https://msdn.microsoft.com/zh-tw/library/60ecse8t(v=vs.80).aspx
        Dim Code As String = ""
        For i As Integer = 0 To 7
            Dim iChr As Integer = CInt((25 + 5) * TIMS.Rnd1X()) + 65 '大寫字元 增加5個亂數
            If iChr >= 91 Then
                iChr = CInt(25 * TIMS.Rnd1X()) + 97 '小寫字元
            End If
            Code &= Chr(iChr)
        Next
        Return Code
    End Function

#Region "DES"

    '加密
    'Public Shared Function Encrypt(ByVal pToEncrypt As String, ByVal sKey As String) As String
    '    Dim rst As String = ""
    '    If pToEncrypt.Length = 0 Then Return rst
    '    If sKey.Length <> 8 Then Return rst
    '    Try
    '        '於System.Security.Cryptography庫， 建議使用AesCryptoServiceProvider
    '        'Dim desAES As New AesCryptoServiceProvider() ' Compliant
    '        Dim des As New DESCryptoServiceProvider()
    '        Dim inputByteArray() As Byte = Encoding.Default.GetBytes(pToEncrypt)
    '        '建立加密對象的密鑰和偏移量
    '        '原文使用ASCIIEncoding.ASCII方法的GetBytes方法
    '        '使得輸入密碼必須輸入英文文本
    '        des.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
    '        des.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
    '        '寫二進制數組到加密流
    '        '(把內存流中的內容全部寫入)
    '        Dim ms As New System.IO.MemoryStream()
    '        Dim cs As New CryptoStream(ms, des.CreateEncryptor, CryptoStreamMode.Write)
    '        '寫二進制數組到加密流
    '        '(把內存流中的內容全部寫入)
    '        cs.Write(inputByteArray, 0, inputByteArray.Length)
    '        cs.FlushFinalBlock()

    '        '建立輸出字符串     
    '        Dim ret As New StringBuilder()
    '        Dim b As Byte
    '        For Each b In ms.ToArray()
    '            ret.AppendFormat("{0:X2}", b)
    '        Next
    '        rst = ret.ToString()
    '    Catch ex As Exception
    '        TIMS.LOG.Warn(ex.Message, ex)
    '    End Try
    '    Return rst
    'End Function

    ''加密2
    'Public Shared Function Encrypt2(ByVal pToEncrypt As String) As String
    '    Dim sKEY As String = GetRnd8Eng()
    '    Return sKEY & "," & Encrypt(pToEncrypt, sKEY)
    'End Function

    ''解密方法 (可能會ERROR)
    'Public Shared Function Decrypt(ByVal pToDecrypt As String, ByVal sKey As String) As String
    '    Dim rst As String = ""
    '    If pToDecrypt.Length = 0 Then Return rst
    '    If sKey.Length <> 8 Then Return rst
    '    '於System.Security.Cryptography庫， 建議使用AesCryptoServiceProvider
    '    'Dim desAES As New AesCryptoServiceProvider() ' Compliant
    '    Dim des As New DESCryptoServiceProvider()
    '    '把字符串放入byte數組
    '    Dim iLen As Integer = pToDecrypt.Length / 2 - 1
    '    'len = pToDecrypt.Length / 2 - 1
    '    Dim inputByteArray(iLen) As Byte
    '    'Dim i As Integer=0
    '    For x As Integer = 0 To iLen
    '        Dim i As Integer = Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16)
    '        inputByteArray(x) = CType(i, Byte)
    '    Next
    '    '建立加密對象的密鑰和偏移量，此值重要，不能修改
    '    des.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
    '    des.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
    '    Dim ms As New System.IO.MemoryStream()
    '    Dim cs As New CryptoStream(ms, des.CreateDecryptor, CryptoStreamMode.Write)
    '    cs.Write(inputByteArray, 0, inputByteArray.Length)
    '    cs.FlushFinalBlock()
    '    rst = Encoding.Default.GetString(ms.ToArray)
    '    Return rst
    'End Function

    ''解密方法2 (解Encrypt2)
    'Public Shared Function Decrypt2(ByVal vStr As String) As String
    '    Dim rst As String = ""
    '    If vStr Is Nothing Then Return rst
    '    If vStr.IndexOf(",") = -1 Then Return rst
    '    If vStr.Split(",").Length < 2 Then Return rst
    '    Dim sKey As String = vStr.Split(",")(0)
    '    Dim pToDecrypt As String = vStr.Split(",")(1)
    '    If pToDecrypt.Length = 0 Then Return rst
    '    If sKey.Length <> 8 Then Return rst

    '    '於System.Security.Cryptography庫， 建議使用AesCryptoServiceProvider
    '    'Dim desAES As New AesCryptoServiceProvider() ' Compliant
    '    'Dim des As New DESCryptoServiceProvider()
    '    Using RsaDES As New DESCryptoServiceProvider
    '        '把字符串放入byte數組
    '        Dim iLen As Integer = pToDecrypt.Length / 2 - 1
    '        'len = pToDecrypt.Length / 2 - 1
    '        Dim inputByteArray(iLen) As Byte
    '        'Dim i As Integer=0
    '        For x As Integer = 0 To iLen
    '            Dim i As Integer = Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16)
    '            inputByteArray(x) = CType(i, Byte)
    '        Next
    '        '建立加密對象的密鑰和偏移量，此值重要，不能修改
    '        RsaDES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
    '        RsaDES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
    '        Try
    '            Dim ms As New System.IO.MemoryStream()
    '            Dim cs As New CryptoStream(ms, RsaDES.CreateDecryptor, CryptoStreamMode.Write)
    '            cs.Write(inputByteArray, 0, inputByteArray.Length)
    '            cs.FlushFinalBlock()
    '            rst = Encoding.Default.GetString(ms.ToArray)
    '        Catch ex As Exception
    '            TIMS.LOG.Warn(ex.Message, ex)
    '        End Try
    '    End Using
    '    Return rst
    'End Function

#End Region


#Region "AesCrypto"

    '加密
    Public Shared Function AesEncrypt(ByVal pToEncrypt As String, ByVal sKey As String) As String
        Dim rst As String = ""
        If pToEncrypt.Length = 0 Then Return rst
        If sKey.Length <> 8 Then Return rst
        Try
            '於System.Security.Cryptography庫， 建議使用AesCryptoServiceProvider
            Dim desAES As New AesCryptoServiceProvider() ' Compliant 'Dim des As New DESCryptoServiceProvider()
            Dim inputByteArray() As Byte = Encoding.Default.GetBytes(pToEncrypt)
            '建立加密對象的密鑰和偏移量
            'Byte[] key = sha256.ComputeHash(Encoding.UTF8.GetBytes(CryptoKey))
            'Byte[] iv = md5.ComputeHash(Encoding.UTF8.GetBytes(CryptoKey));
            Dim sha256 As New Security.Cryptography.SHA256CryptoServiceProvider
            Dim MD5cp As MD5CryptoServiceProvider = New MD5CryptoServiceProvider() 'Dim md5cn As MD5 = MD5.Create()
            desAES.Key = sha256.ComputeHash(Encoding.UTF8.GetBytes(sKey)) 'ASCIIEncoding.ASCII.GetBytes(sKey)
            desAES.IV = MD5cp.ComputeHash(Encoding.UTF8.GetBytes(sKey)) 'ASCIIEncoding.ASCII.GetBytes(sKey)

            '寫二進制數組到加密流 '(把內存流中的內容全部寫入)
            Dim ms As New System.IO.MemoryStream()
            Dim cs As New CryptoStream(ms, desAES.CreateEncryptor, CryptoStreamMode.Write)
            cs.Write(inputByteArray, 0, inputByteArray.Length)
            cs.FlushFinalBlock()

            '建立輸出字符串     
            Dim ret As New StringBuilder()
            Dim b As Byte
            For Each b In ms.ToArray()
                ret.AppendFormat("{0:X2}", b)
            Next
            rst = ret.ToString()
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
        End Try
        Return rst
    End Function

    '加密2
    Public Shared Function AesEncrypt2(ByVal pToEncrypt As String) As String
        Dim sKEY As String = GetRnd8Eng()
        Return $"{sKEY},{AesEncrypt(pToEncrypt, sKEY)}"
    End Function

    '解密方法 (可能會ERROR)
    Public Shared Function AesDecrypt(ByVal pToDecrypt As String, ByVal sKey As String) As String
        Dim rst As String = ""
        If pToDecrypt.Length = 0 Then Return rst
        If sKey.Length <> 8 Then Return rst
        '於System.Security.Cryptography庫， 建議使用AesCryptoServiceProvider
        Dim desAES As New AesCryptoServiceProvider() ' Compliant 'Dim des As New DESCryptoServiceProvider()
        '把字符串放入byte數組
        Dim iLen As Integer = pToDecrypt.Length / 2 - 1
        'len = pToDecrypt.Length / 2 - 1
        Dim inputByteArray(iLen) As Byte
        'Dim i As Integer=0
        For x As Integer = 0 To iLen
            Dim i As Integer = Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16)
            inputByteArray(x) = CType(i, Byte)
        Next
        '建立加密對象的密鑰和偏移量，此值重要，不能修改
        'Byte[] key = sha256.ComputeHash(Encoding.UTF8.GetBytes(CryptoKey))
        'Byte[] iv = md5.ComputeHash(Encoding.UTF8.GetBytes(CryptoKey));
        Dim sha256 As New Security.Cryptography.SHA256CryptoServiceProvider
        Dim MD5cp As MD5CryptoServiceProvider = New MD5CryptoServiceProvider() 'Dim md5cn As MD5 = MD5.Create()
        desAES.Key = sha256.ComputeHash(Encoding.UTF8.GetBytes(sKey)) 'ASCIIEncoding.ASCII.GetBytes(sKey)
        desAES.IV = MD5cp.ComputeHash(Encoding.UTF8.GetBytes(sKey)) 'ASCIIEncoding.ASCII.GetBytes(sKey)

        Dim ms As New System.IO.MemoryStream()
        Dim cs As New CryptoStream(ms, desAES.CreateDecryptor, CryptoStreamMode.Write)
        cs.Write(inputByteArray, 0, inputByteArray.Length)
        cs.FlushFinalBlock()
        rst = Encoding.Default.GetString(ms.ToArray)
        Return rst
    End Function

    '解密方法2 (解Encrypt2)
    Public Shared Function AesDecrypt2(ByVal vStr As String) As String
        Dim rst As String = ""
        If vStr Is Nothing Then Return rst
        If vStr.IndexOf(",") = -1 Then Return rst
        If vStr.Split(",").Length < 2 Then Return rst
        Dim sKey As String = vStr.Split(",")(0)
        Dim pToDecrypt As String = vStr.Split(",")(1)
        If pToDecrypt.Length = 0 Then Return rst
        If sKey.Length <> 8 Then Return rst

        '於System.Security.Cryptography庫， 建議使用AesCryptoServiceProvider
        'Dim desAES As New AesCryptoServiceProvider() ' Compliant 'Dim des As New DESCryptoServiceProvider()
        Using RsaAES As New AesCryptoServiceProvider 'DESCryptoServiceProvider
            '把字符串放入byte數組
            Dim iLen As Integer = pToDecrypt.Length / 2 - 1
            'len = pToDecrypt.Length / 2 - 1
            Dim inputByteArray(iLen) As Byte
            'Dim i As Integer=0
            For x As Integer = 0 To iLen
                Dim i As Integer = Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16)
                inputByteArray(x) = CType(i, Byte)
            Next
            '建立加密對象的密鑰和偏移量，此值重要，不能修改
            'Byte[] key = sha256.ComputeHash(Encoding.UTF8.GetBytes(CryptoKey))
            'Byte[] iv = md5.ComputeHash(Encoding.UTF8.GetBytes(CryptoKey));
            Dim sha256 As New Security.Cryptography.SHA256CryptoServiceProvider
            Dim MD5cp As MD5CryptoServiceProvider = New MD5CryptoServiceProvider() 'Dim md5cn As MD5 = MD5.Create()
            RsaAES.Key = sha256.ComputeHash(Encoding.UTF8.GetBytes(sKey)) 'ASCIIEncoding.ASCII.GetBytes(sKey)
            RsaAES.IV = MD5cp.ComputeHash(Encoding.UTF8.GetBytes(sKey)) 'ASCIIEncoding.ASCII.GetBytes(sKey)

            Try
                Dim ms As New System.IO.MemoryStream()
                Dim cs As New CryptoStream(ms, RsaAES.CreateDecryptor, CryptoStreamMode.Write)
                cs.Write(inputByteArray, 0, inputByteArray.Length)
                cs.FlushFinalBlock()
                rst = Encoding.Default.GetString(ms.ToArray)
            Catch ex As Exception
                TIMS.LOG.Warn(ex.Message, ex)
            End Try
        End Using
        Return rst
    End Function

#End Region

End Class
