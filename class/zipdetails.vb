Public Class zipdetails
    Private m_zipcode1 As String
    Private m_zipcode3 As String
    Private m_zipcode6 As String
    Private m_zipcodeN As String
    Private m_city1 As String
    Private m_address1 As String

    Public Sub New()
        m_zipcode1 = ""
        m_zipcode3 = ""
        m_zipcode6 = ""
        m_city1 = ""
        m_address1 = ""
    End Sub

    Public Property ZipCode1() As String
        Get
            Return m_zipcode1
        End Get
        Set
            m_zipcode1 = Value
        End Set
    End Property
    Public Property ZipCode3() As String
        Get
            Return m_zipcode3
        End Get
        Set
            m_zipcode3 = Value
        End Set
    End Property

    Public Property ZipCode6() As String
        Get
            Return m_zipcode6
        End Get
        Set
            m_zipcode6 = Value
        End Set
    End Property

    Public Property City1() As String
        Get
            Return m_city1
        End Get
        Set
            m_city1 = Value
        End Set
    End Property

    Public Property Address1() As String
        Get
            Return m_address1
        End Get
        Set
            m_address1 = Value
        End Set
    End Property

End Class
