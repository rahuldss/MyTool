Imports MyTool.NDS.LIB
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace NDS.DAL

    Public Class DALTable_1

        
        ''' <summary>
        ''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTable_1Details() As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objLIBTable_1Listing As New LIBTable_1Listing
            Dim Packet As New MyCLS.TransportationPacket

            Try
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromTable_1")
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBTable_1 As New LIBTable_1
                                    objLIBTable_1Listing.Add(oLIBTable_1)
                                    objLIBTable_1Listing(i).a = ds.Tables(0).Rows(i)("a").ToString
                                    objLIBTable_1Listing(i).b = ds.Tables(0).Rows(i)("b").ToString
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If
                Packet.MessageResultset = objLIBTable_1Listing
                Packet.MessageResultsetDS = ds

            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


       
        ''' <summary>
        ''' Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTable_1Details(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objParamList As New List(Of SqlParameter)()
            Dim objLIBTable_1Listing As New LIBTable_1Listing

            Try
                objParamList.Add(New SqlParameter("@Id", Packet.MessagePacket))
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromTable_1ById", objParamList)
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBTable_1 As New LIBTable_1
                                    objLIBTable_1Listing.Add(oLIBTable_1)
                                    objLIBTable_1Listing(i).a = ds.Tables(0).Rows(i)("a").ToString
                                    objLIBTable_1Listing(i).b = ds.Tables(0).Rows(i)("b").ToString
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If

                Packet.MessageResultset = objLIBTable_1Listing

            Catch ex As Exception
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


       
        ''' <summary>
        ''' Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <param name="Packet"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function InsertTable_1(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim strOutParamValues As String()
            Dim objParamList As New List(Of SqlParameter)()
            Dim objParamListOut As New List(Of SqlParameter)()
            Try
                Dim objLIBTable_1 As New LIBTable_1
                objLIBTable_1 = Packet.MessagePacket

                objParamList.Add(New SqlParameter("@a", objLIBTable_1.a))
                objParamList.Add(New SqlParameter("@b", objLIBTable_1.b))
                objParamListOut.Add(New SqlParameter("@@a", SqlDbType.VarChar, 100))
                'MyCLS.clsCOMMON.ConOpen(true)
                strOutParamValues = MyCLS.clsExecuteStoredProcSql.ExecuteSPNonQueryOutPut("SP_InsertTable_1", objParamList, objParamListOut, Packet.MessageId)
                'Result = 1
                'MyCLS.clsCOMMON.ConClose()

                Packet.MessageResultset = strOutParamValues

            Catch ex As Exception
                'Result = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function

    End Class
End Namespace
