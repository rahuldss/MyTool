Namespace NDS.LIB

    Public Class LIBTable_2

        Private _c As String
        Public Property c() As String
            Get
                Return _c
            End Get
            Set(ByVal value As String)
                _c = value
            End Set
        End Property

        Private _d As String
        Public Property d() As String
            Get
                Return _d
            End Get
            Set(ByVal value As String)
                _d = value
            End Set
        End Property

        Private _ImgVarBinary As Byte()
        Public Property ImgVarBinary() As Byte()
            Get
                Return _ImgVarBinary
            End Get
            Set(ByVal value As Byte())
                _ImgVarBinary = value
            End Set
        End Property
    End Class


    <Serializable()> _
    Public Class LIBTable_2Listing
        Inherits List(Of LIBTable_2)

    End Class
End Namespace
