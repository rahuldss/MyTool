Namespace NDS.LIB

    Public Class LIBAllDBTypesSQLTable

        Private _ID As Int32
        Public Property ID() As Int32
            Get
                Return _ID
            End Get
            Set(ByVal value As Int32)
                _ID = value
            End Set
        End Property

        Private _A As Int64
        Public Property A() As Int64
            Get
                Return _A
            End Get
            Set(ByVal value As Int64)
                _A = value
            End Set
        End Property

        Private _B As Byte()
        Public Property B() As Byte()
            Get
                Return _B
            End Get
            Set(ByVal value As Byte())
                _B = value
            End Set
        End Property

        Private _C As Boolean
        Public Property C() As Boolean
            Get
                Return _C
            End Get
            Set(ByVal value As Boolean)
                _C = value
            End Set
        End Property

        Private _D As String
        Public Property D() As String
            Get
                Return _D
            End Get
            Set(ByVal value As String)
                _D = value
            End Set
        End Property

        Private _E As DateTime
        Public Property E() As DateTime
            Get
                Return _E
            End Get
            Set(ByVal value As DateTime)
                _E = value
            End Set
        End Property

        Private _F As Decimal
        Public Property F() As Decimal
            Get
                Return _F
            End Get
            Set(ByVal value As Decimal)
                _F = value
            End Set
        End Property

        Private _G As Decimal
        Public Property G() As Decimal
            Get
                Return _G
            End Get
            Set(ByVal value As Decimal)
                _G = value
            End Set
        End Property

        Private _H As Double
        Public Property H() As Double
            Get
                Return _H
            End Get
            Set(ByVal value As Double)
                _H = value
            End Set
        End Property

        Private _I As Byte()
        Public Property I() As Byte()
            Get
                Return _I
            End Get
            Set(ByVal value As Byte())
                _I = value
            End Set
        End Property

        Private _J As Int32
        Public Property J() As Int32
            Get
                Return _J
            End Get
            Set(ByVal value As Int32)
                _J = value
            End Set
        End Property

        Private _K As Decimal
        Public Property K() As Decimal
            Get
                Return _K
            End Get
            Set(ByVal value As Decimal)
                _K = value
            End Set
        End Property

        Private _L As String
        Public Property L() As String
            Get
                Return _L
            End Get
            Set(ByVal value As String)
                _L = value
            End Set
        End Property

        Private _M As String
        Public Property M() As String
            Get
                Return _M
            End Get
            Set(ByVal value As String)
                _M = value
            End Set
        End Property

        Private _N As Decimal
        Public Property N() As Decimal
            Get
                Return _N
            End Get
            Set(ByVal value As Decimal)
                _N = value
            End Set
        End Property

        Private _O As Double
        Public Property O() As Double
            Get
                Return _O
            End Get
            Set(ByVal value As Double)
                _O = value
            End Set
        End Property

        Private _P As String
        Public Property P() As String
            Get
                Return _P
            End Get
            Set(ByVal value As String)
                _P = value
            End Set
        End Property

        Private _Q As String
        Public Property Q() As String
            Get
                Return _Q
            End Get
            Set(ByVal value As String)
                _Q = value
            End Set
        End Property

        Private _R As Single
        Public Property R() As Single
            Get
                Return _R
            End Get
            Set(ByVal value As Single)
                _R = value
            End Set
        End Property

        Private _S As DateTime
        Public Property S() As DateTime
            Get
                Return _S
            End Get
            Set(ByVal value As DateTime)
                _S = value
            End Set
        End Property

        Private _T As int16
        Public Property T() As int16
            Get
                Return _T
            End Get
            Set(ByVal value As int16)
                _T = value
            End Set
        End Property

        Private _U As Decimal
        Public Property U() As Decimal
            Get
                Return _U
            End Get
            Set(ByVal value As Decimal)
                _U = value
            End Set
        End Property

        Private _V As Object
        Public Property V() As Object
            Get
                Return _V
            End Get
            Set(ByVal value As Object)
                _V = value
            End Set
        End Property

        Private _W As String
        Public Property W() As String
            Get
                Return _W
            End Get
            Set(ByVal value As String)
                _W = value
            End Set
        End Property

        Private _X As Byte()
        Public Property X() As Byte()
            Get
                Return _X
            End Get
            Set(ByVal value As Byte())
                _X = value
            End Set
        End Property

        Private _Y As Byte
        Public Property Y() As Byte
            Get
                Return _Y
            End Get
            Set(ByVal value As Byte)
                _Y = value
            End Set
        End Property

        Private _Z As Guid
        Public Property Z() As Guid
            Get
                Return _Z
            End Get
            Set(ByVal value As Guid)
                _Z = value
            End Set
        End Property

        Private _A1 As Byte()
        Public Property A1() As Byte()
            Get
                Return _A1
            End Get
            Set(ByVal value As Byte())
                _A1 = value
            End Set
        End Property

        Private _B1 As Byte()
        Public Property B1() As Byte()
            Get
                Return _B1
            End Get
            Set(ByVal value As Byte())
                _B1 = value
            End Set
        End Property

        Private _C1 As String
        Public Property C1() As String
            Get
                Return _C1
            End Get
            Set(ByVal value As String)
                _C1 = value
            End Set
        End Property

        Private _D1 As String
        Public Property D1() As String
            Get
                Return _D1
            End Get
            Set(ByVal value As String)
                _D1 = value
            End Set
        End Property

        Private _E1 As String
        Public Property E1() As String
            Get
                Return _E1
            End Get
            Set(ByVal value As String)
                _E1 = value
            End Set
        End Property

    End Class


    <Serializable()> _
    Public Class LIBAllDBTypesSQLTableListing
        Inherits List(Of LIBAllDBTypesSQLTable)

    End Class
End Namespace