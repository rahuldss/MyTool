Namespace NDS.LIB

    Public Class LIBTable_1

        Private _a As String
        Public Property a() As String
            Get
                Return _a
            End Get
            Set(ByVal value As String)
                _a = value
            End Set
        End Property

        Private _b As String
        Public Property b() As String
            Get
                Return _b
            End Get
            Set(ByVal value As String)
                _b = value
            End Set
        End Property

    End Class


    <Serializable()> _
    Public Class LIBTable_1Listing
        Inherits List(Of LIBTable_1)

    End Class
End Namespace
