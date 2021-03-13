Imports MyTool.NDS.LIB
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace NDS.DAL

    Public Class DALTable_2

        'PUT IT IN LOAD EVENTS

        'MyCLS.strConnStringOLEDB = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
        'MyCLS.strConnStringSQLCLIENT = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"

        '*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************
        'Try
        '    Dim objLIBTable_2Listing As New LIBTable_2Listing
        '    Dim objDALTable_2 As New DALTable_2
        '    Dim tp As New MyCLS.TransportationPacket

        ''    objLIBTable_2Listing(0).c = ""
        ''    objLIBTable_2Listing(1).d = ""
        ''    objLIBTable_2Listing(2).ImgVarBinary = ""
        ''    objLIBTable_2Listing(3).ImgImage = ""
        '    tp = objDALTable_2.GetTable_2Details()
        '    If tp.MessageId = 1 Then
        '        objLIBTable_2Listing = tp.MessageResultset
        '        MsgBox(objLIBTable_2Listing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTable_2Details() As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objLIBTable_2Listing As New LIBTable_2Listing
            Dim Packet As New MyCLS.TransportationPacket

            Try
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromTable_2")
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBTable_2 As New LIBTable_2
                                    objLIBTable_2Listing.Add(oLIBTable_2)
                                    objLIBTable_2Listing(i).c = ds.Tables(0).Rows(i)("c").ToString
                                    objLIBTable_2Listing(i).d = ds.Tables(0).Rows(i)("d").ToString
                                    objLIBTable_2Listing(i).ImgVarBinary = ds.Tables(0).Rows(i)("ImgVarBinary")                                    
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If
                Packet.MessageResultset = objLIBTable_2Listing

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
        Public Function GetTable_2Details(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objParamList As New List(Of SqlParameter)()
            Dim objLIBTable_2Listing As New LIBTable_2Listing

            Try
                objParamList.Add(New SqlParameter("@c", Packet.MessagePacket))
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromTable_2ById", objParamList)
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBTable_2 As New LIBTable_2
                                    objLIBTable_2Listing.Add(oLIBTable_2)
                                    objLIBTable_2Listing(i).c = ds.Tables(0).Rows(i)("c").ToString
                                    objLIBTable_2Listing(i).d = ds.Tables(0).Rows(i)("d").ToString
                                    objLIBTable_2Listing(i).ImgVarBinary = ds.Tables(0).Rows(i)("ImgVarBinary")
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If

                Packet.MessageResultset = objLIBTable_2Listing

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
        Public Function InsertTable_2(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim strOutParamValues As String()
            Dim objParamList As New List(Of SqlParameter)()
            Dim objParamListOut As New List(Of SqlParameter)()
            Dim Result As Int16 = 0
            Try
                Dim objLIBTable_2 As New LIBTable_2
                objLIBTable_2 = Packet.MessagePacket

                objParamList.Add(New SqlParameter("@c", objLIBTable_2.c))
                objParamList.Add(New SqlParameter("@d", objLIBTable_2.d))
                'objParamList.Add(New SqlParameter("@ImgVarBinary", SqlDbType.VarBinary, objLIBTable_2.ImgVarBinary.Length, ParameterDirection.Input, True, 0, 0, "", DataRowVersion.Current, objLIBTable_2.ImgVarBinary)) 
                objParamList.Add(New SqlParameter("@ImgVarBinary", objLIBTable_2.ImgVarBinary))
                objParamListOut.Add(New SqlParameter("@@c", SqlDbType.VarChar, 100))
                strOutParamValues = MyCLS.clsExecuteStoredProcSql.ExecuteSPNonQueryOutPut("SP_InsertTable_2", objParamList, objParamListOut, Packet.MessageId)
                Packet.MessageResultset = strOutParamValues

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function

    End Class
End Namespace
