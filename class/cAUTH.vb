Public Class cAUTH
    'blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt
    Private m_blnCanAdds As Boolean
    Private m_blnCanMod As Boolean
    Private m_blnCanDel As Boolean
    Private m_blnCanSech As Boolean
    Private m_blnCanPrnt As Boolean

    Public Sub New()
        m_blnCanAdds = False '新增
        m_blnCanMod = False '修改
        m_blnCanDel = False '刪除
        m_blnCanSech = False '查詢
        m_blnCanPrnt = False '列印
    End Sub

    Property blnCanAdds() As Boolean
        Get
            Return m_blnCanAdds
        End Get
        Set(ByVal value As Boolean)
            m_blnCanAdds = value
        End Set
    End Property

    Property blnCanMod() As Boolean
        Get
            Return m_blnCanMod
        End Get
        Set(ByVal value As Boolean)
            m_blnCanMod = value
        End Set
    End Property

    Property blnCanDel() As Boolean
        Get
            Return m_blnCanDel
        End Get
        Set(ByVal value As Boolean)
            m_blnCanDel = value
        End Set
    End Property

    Property blnCanSech() As Boolean
        Get
            Return m_blnCanSech
        End Get
        Set(ByVal value As Boolean)
            m_blnCanSech = value
        End Set
    End Property

    Property blnCanPrnt() As Boolean
        Get
            Return m_blnCanPrnt
        End Get
        Set(ByVal value As Boolean)
            m_blnCanPrnt = value
        End Set
    End Property

End Class
