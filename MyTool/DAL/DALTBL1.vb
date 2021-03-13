Imports MyTool.NDS.LIB
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace NDS.DAL

    Public Class DALTBL1

        'PUT IT IN LOAD EVENTS

        'MyCLS.strConnStringOLEDB = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
        'MyCLS.strConnStringSQLCLIENT = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"

        '*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************
        'Try
        '    Dim objLIBTBL1Listing As New LIBTBL1Listing
        '    Dim objDALTBL1 As New DALTBL1
        '    Dim tp As New MyCLS.TransportationPacket
        '    Dim ds As New Dataset

        '' txt.Text = objLIBTBL1Listing(0).ID
        '' txt.Text = objLIBTBL1Listing(0).Img
        '    tp = objDALTBL1.GetTBL1Details()
        '    If tp.MessageId = 1 Then
        '        objLIBTBL1Listing = tp.MessageResultset
        '        ds = tp.MessageResultsetDS
        '        MyCLS.clsImaging.ByteArray2Image(,)
        '        MsgBox(objLIBTBL1Listing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTBL1Details() As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objLIBTBL1Listing As New LIBTBL1Listing
            Dim Packet As New MyCLS.TransportationPacket

            Try
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromTBL1")
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBTBL1 As New LIBTBL1
                                    objLIBTBL1Listing.Add(oLIBTBL1)
                                    objLIBTBL1Listing(i).ID = ds.Tables(0).Rows(i)("ID").ToString
                                    objLIBTBL1Listing(i).Img = ds.Tables(0).Rows(i)("Img")
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If
                Packet.MessageResultsetDS = ds
                Packet.MessageResultset = objLIBTBL1Listing

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


        '*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************
        'Try
        '    Dim objLIBTBL1Listing As New LIBTBL1Listing
        '    Dim objDALTBL1 As New DALTBL1
        '    Dim tp As New MyCLS.TransportationPacket
        '    Dim ds As New Dataset
        '    tp.MessagePacket = 1    'ID to be Passed

        '' txt.Text = objLIBTBL1Listing(0).ID
        '' txt.Text = objLIBTBL1Listing(0).Img
        '    tp = objDALTBL1.GetTBL1Details(tp)
        '    If tp.MessageId = 1 Then
        '        objLIBTBL1Listing = tp.MessageResultset
        '        ds = tp.MessageResultsetDS
        '        MyCLS.clsImaging.ByteArray2Image(,)
        '        MsgBox(objLIBTBL1Listing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTBL1Details(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objParamList As New List(Of SqlParameter)()
            Dim objLIBTBL1Listing As New LIBTBL1Listing

            Try
                objParamList.Add(New SqlParameter("@Id", Packet.MessagePacket))
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromTBL1ById", objParamList)
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBTBL1 As New LIBTBL1
                                    objLIBTBL1Listing.Add(oLIBTBL1)
                                    objLIBTBL1Listing(i).ID = ds.Tables(0).Rows(i)("ID").ToString
                                    objLIBTBL1Listing(i).Img = ds.Tables(0).Rows(i)("Img")
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If

                Packet.MessageResultsetDS = ds
                Packet.MessageResultset = objLIBTBL1Listing

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


        '*******COPY IT TO USE BELOW FUNCTION - INSERT************
        'Try
        '    Dim objLIBTBL1 As New LIBTBL1
        '    Dim objDALTBL1 As New DALTBL1
        '    Dim tp As New MyCLS.TransportationPacket

        '    objLIBTBL1.ID = txt.Text
        '    objLIBTBL1.Img = MyCLS.clsImaging.PictureBoxToByteArray()
        '    tp.MessagePacket = objLIBTBL1
        '    tp = objDALTBL1.InsertTBL1(tp)

        '    If tp.MessageId > -1 Then
        '        Dim strOutParamValues As String() = tp.MessageResultset
        '        MsgBox(strOutParamValues(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=Packet, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <param name="Packet"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function InsertTBL1(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim strOutParamValues As String()
            Dim objParamList As New List(Of SqlParameter)()
            Dim objParamListOut As New List(Of SqlParameter)()
            Dim Result As Int16 = 0
            Try
                Dim objLIBTBL1 As New LIBTBL1
                objLIBTBL1 = Packet.MessagePacket

                objParamList.Add(New SqlParameter("@ID", objLIBTBL1.ID))
                objParamList.Add(New SqlParameter("@Img", objLIBTBL1.Img))
                objParamListOut.Add(New SqlParameter("@@ID", SqlDbType.Int))
                strOutParamValues = MyCLS.clsExecuteStoredProcSql.ExecuteSPNonQueryOutPut("SP_InsertTBL1", objParamList, objParamListOut, Packet.MessageId)
                Packet.MessageResultset = strOutParamValues

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function

    End Class
End Namespace