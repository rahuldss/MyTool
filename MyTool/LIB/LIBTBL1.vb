Namespace NDS.LIB

    Public Class LIBTBL1

        Private _ID As Int32
        Public Property ID() As Int32
            Get
                Return _ID
            End Get
            Set(ByVal value As Int32)
                _ID = value
            End Set
        End Property

        Private _Img As Byte()
        Public Property Img() As Byte()
            Get
                Return _Img
            End Get
            Set(ByVal value As Byte())
                _Img = value
            End Set
        End Property

    End Class


    <Serializable()> _
    Public Class LIBTBL1Listing
        Inherits List(Of LIBTBL1)

    End Class
End Namespace