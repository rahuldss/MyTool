Imports MyTool.NDS.LIB
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace NDS.DAL

    Public Class DALAllDBTypesSQLTable

        'PUT IT IN LOAD EVENTS

        'MyCLS.strConnStringOLEDB = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
        'MyCLS.strConnStringSQLCLIENT = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"

        '*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************
        'Try
        '    Dim objLIBAllDBTypesSQLTableListing As New LIBAllDBTypesSQLTableListing
        '    Dim objDALAllDBTypesSQLTable As New DALAllDBTypesSQLTable
        '    Dim tp As New MyCLS.TransportationPacket
        '    Dim ds As New Dataset

        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).ID
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).A
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).B
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).C
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).D
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).E
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).F
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).G
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).H
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).I
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).J
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).K
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).L
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).M
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).N
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).O
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).P
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).Q
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).R
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).S
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).T
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).U
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).V
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).W
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).X
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).Y
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).Z
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).A1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).B1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).C1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).D1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).E1
        '    tp = objDALAllDBTypesSQLTable.GetAllDBTypesSQLTableDetails()
        '    If tp.MessageId = 1 Then
        '        objLIBAllDBTypesSQLTableListing = tp.MessageResultset
        '        ds = tp.MessageResultsetDS
        '        MyCLS.clsImaging.ByteArray2Image(,)
        '        MsgBox(objLIBAllDBTypesSQLTableListing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllDBTypesSQLTableDetails() As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objLIBAllDBTypesSQLTableListing As New LIBAllDBTypesSQLTableListing
            Dim Packet As New MyCLS.TransportationPacket

            Try
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromAllDBTypesSQLTable")
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBAllDBTypesSQLTable As New LIBAllDBTypesSQLTable
                                    objLIBAllDBTypesSQLTableListing.Add(oLIBAllDBTypesSQLTable)
                                    objLIBAllDBTypesSQLTableListing(i).ID = ds.Tables(0).Rows(i)("ID").ToString
                                    objLIBAllDBTypesSQLTableListing(i).A = ds.Tables(0).Rows(i)("A").ToString
                                    objLIBAllDBTypesSQLTableListing(i).B = ds.Tables(0).Rows(i)("B")
                                    objLIBAllDBTypesSQLTableListing(i).C = ds.Tables(0).Rows(i)("C").ToString
                                    objLIBAllDBTypesSQLTableListing(i).D = ds.Tables(0).Rows(i)("D").ToString
                                    objLIBAllDBTypesSQLTableListing(i).E = ds.Tables(0).Rows(i)("E").ToString
                                    objLIBAllDBTypesSQLTableListing(i).F = ds.Tables(0).Rows(i)("F").ToString
                                    objLIBAllDBTypesSQLTableListing(i).G = ds.Tables(0).Rows(i)("G").ToString
                                    objLIBAllDBTypesSQLTableListing(i).H = ds.Tables(0).Rows(i)("H").ToString
                                    objLIBAllDBTypesSQLTableListing(i).I = ds.Tables(0).Rows(i)("I")
                                    objLIBAllDBTypesSQLTableListing(i).J = ds.Tables(0).Rows(i)("J").ToString
                                    objLIBAllDBTypesSQLTableListing(i).K = ds.Tables(0).Rows(i)("K").ToString
                                    objLIBAllDBTypesSQLTableListing(i).L = ds.Tables(0).Rows(i)("L").ToString
                                    objLIBAllDBTypesSQLTableListing(i).M = ds.Tables(0).Rows(i)("M").ToString
                                    objLIBAllDBTypesSQLTableListing(i).N = ds.Tables(0).Rows(i)("N").ToString
                                    objLIBAllDBTypesSQLTableListing(i).O = ds.Tables(0).Rows(i)("O").ToString
                                    objLIBAllDBTypesSQLTableListing(i).P = ds.Tables(0).Rows(i)("P").ToString
                                    objLIBAllDBTypesSQLTableListing(i).Q = ds.Tables(0).Rows(i)("Q").ToString
                                    objLIBAllDBTypesSQLTableListing(i).R = ds.Tables(0).Rows(i)("R").ToString
                                    objLIBAllDBTypesSQLTableListing(i).S = ds.Tables(0).Rows(i)("S").ToString
                                    objLIBAllDBTypesSQLTableListing(i).T = ds.Tables(0).Rows(i)("T").ToString
                                    objLIBAllDBTypesSQLTableListing(i).U = ds.Tables(0).Rows(i)("U").ToString
                                    objLIBAllDBTypesSQLTableListing(i).V = ds.Tables(0).Rows(i)("V").ToString
                                    objLIBAllDBTypesSQLTableListing(i).W = ds.Tables(0).Rows(i)("W").ToString
                                    objLIBAllDBTypesSQLTableListing(i).X = ds.Tables(0).Rows(i)("X")
                                    objLIBAllDBTypesSQLTableListing(i).Y = ds.Tables(0).Rows(i)("Y").ToString
                                    objLIBAllDBTypesSQLTableListing(i).Z = ds.Tables(0).Rows(i)("Z")
                                    objLIBAllDBTypesSQLTableListing(i).A1 = ds.Tables(0).Rows(i)("A1")
                                    objLIBAllDBTypesSQLTableListing(i).B1 = ds.Tables(0).Rows(i)("B1")
                                    objLIBAllDBTypesSQLTableListing(i).C1 = ds.Tables(0).Rows(i)("C1").ToString
                                    objLIBAllDBTypesSQLTableListing(i).D1 = ds.Tables(0).Rows(i)("D1").ToString
                                    objLIBAllDBTypesSQLTableListing(i).E1 = ds.Tables(0).Rows(i)("E1").ToString
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If
                Packet.MessageResultsetDS = ds
                Packet.MessageResultset = objLIBAllDBTypesSQLTableListing

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


        '*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************
        'Try
        '    Dim objLIBAllDBTypesSQLTableListing As New LIBAllDBTypesSQLTableListing
        '    Dim objDALAllDBTypesSQLTable As New DALAllDBTypesSQLTable
        '    Dim tp As New MyCLS.TransportationPacket
        '    Dim ds As New Dataset
        '    tp.MessagePacket = 1    'ID to be Passed

        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).ID
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).A
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).B
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).C
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).D
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).E
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).F
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).G
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).H
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).I
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).J
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).K
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).L
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).M
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).N
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).O
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).P
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).Q
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).R
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).S
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).T
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).U
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).V
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).W
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).X
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).Y
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).Z
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).A1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).B1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).C1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).D1
        '' txt.Text = objLIBAllDBTypesSQLTableListing(0).E1
        '    tp = objDALAllDBTypesSQLTable.GetAllDBTypesSQLTableDetails(tp)
        '    If tp.MessageId = 1 Then
        '        objLIBAllDBTypesSQLTableListing = tp.MessageResultset
        '        ds = tp.MessageResultsetDS
        '        MyCLS.clsImaging.ByteArray2Image(,)
        '        MsgBox(objLIBAllDBTypesSQLTableListing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllDBTypesSQLTableDetails(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objParamList As New List(Of SqlParameter)()
            Dim objLIBAllDBTypesSQLTableListing As New LIBAllDBTypesSQLTableListing

            Try
                objParamList.Add(New SqlParameter("@Id", Packet.MessagePacket))
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromAllDBTypesSQLTableById", objParamList)
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBAllDBTypesSQLTable As New LIBAllDBTypesSQLTable
                                    objLIBAllDBTypesSQLTableListing.Add(oLIBAllDBTypesSQLTable)
                                    objLIBAllDBTypesSQLTableListing(i).ID = ds.Tables(0).Rows(i)("ID").ToString
                                    objLIBAllDBTypesSQLTableListing(i).A = ds.Tables(0).Rows(i)("A").ToString
                                    objLIBAllDBTypesSQLTableListing(i).B = ds.Tables(0).Rows(i)("B")
                                    objLIBAllDBTypesSQLTableListing(i).C = ds.Tables(0).Rows(i)("C").ToString
                                    objLIBAllDBTypesSQLTableListing(i).D = ds.Tables(0).Rows(i)("D").ToString
                                    objLIBAllDBTypesSQLTableListing(i).E = ds.Tables(0).Rows(i)("E").ToString
                                    objLIBAllDBTypesSQLTableListing(i).F = ds.Tables(0).Rows(i)("F").ToString
                                    objLIBAllDBTypesSQLTableListing(i).G = ds.Tables(0).Rows(i)("G").ToString
                                    objLIBAllDBTypesSQLTableListing(i).H = ds.Tables(0).Rows(i)("H").ToString
                                    objLIBAllDBTypesSQLTableListing(i).I = ds.Tables(0).Rows(i)("I")
                                    objLIBAllDBTypesSQLTableListing(i).J = ds.Tables(0).Rows(i)("J").ToString
                                    objLIBAllDBTypesSQLTableListing(i).K = ds.Tables(0).Rows(i)("K").ToString
                                    objLIBAllDBTypesSQLTableListing(i).L = ds.Tables(0).Rows(i)("L").ToString
                                    objLIBAllDBTypesSQLTableListing(i).M = ds.Tables(0).Rows(i)("M").ToString
                                    objLIBAllDBTypesSQLTableListing(i).N = ds.Tables(0).Rows(i)("N").ToString
                                    objLIBAllDBTypesSQLTableListing(i).O = ds.Tables(0).Rows(i)("O").ToString
                                    objLIBAllDBTypesSQLTableListing(i).P = ds.Tables(0).Rows(i)("P").ToString
                                    objLIBAllDBTypesSQLTableListing(i).Q = ds.Tables(0).Rows(i)("Q").ToString
                                    objLIBAllDBTypesSQLTableListing(i).R = ds.Tables(0).Rows(i)("R").ToString
                                    objLIBAllDBTypesSQLTableListing(i).S = ds.Tables(0).Rows(i)("S").ToString
                                    objLIBAllDBTypesSQLTableListing(i).T = ds.Tables(0).Rows(i)("T").ToString
                                    objLIBAllDBTypesSQLTableListing(i).U = ds.Tables(0).Rows(i)("U").ToString
                                    objLIBAllDBTypesSQLTableListing(i).V = ds.Tables(0).Rows(i)("V").ToString
                                    objLIBAllDBTypesSQLTableListing(i).W = ds.Tables(0).Rows(i)("W").ToString
                                    objLIBAllDBTypesSQLTableListing(i).X = ds.Tables(0).Rows(i)("X")
                                    objLIBAllDBTypesSQLTableListing(i).Y = ds.Tables(0).Rows(i)("Y").ToString
                                    objLIBAllDBTypesSQLTableListing(i).Z = ds.Tables(0).Rows(i)("Z")
                                    objLIBAllDBTypesSQLTableListing(i).A1 = ds.Tables(0).Rows(i)("A1")
                                    objLIBAllDBTypesSQLTableListing(i).B1 = ds.Tables(0).Rows(i)("B1")
                                    objLIBAllDBTypesSQLTableListing(i).C1 = ds.Tables(0).Rows(i)("C1").ToString
                                    objLIBAllDBTypesSQLTableListing(i).D1 = ds.Tables(0).Rows(i)("D1").ToString
                                    objLIBAllDBTypesSQLTableListing(i).E1 = ds.Tables(0).Rows(i)("E1").ToString
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If

                Packet.MessageResultsetDS = ds
                Packet.MessageResultset = objLIBAllDBTypesSQLTableListing

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


        '*******COPY IT TO USE BELOW FUNCTION - INSERT************
        'Try
        '    Dim objLIBAllDBTypesSQLTable As New LIBAllDBTypesSQLTable
        '    Dim objDALAllDBTypesSQLTable As New DALAllDBTypesSQLTable
        '    Dim tp As New MyCLS.TransportationPacket

        '    objLIBAllDBTypesSQLTable.ID = txt.Text
        '    objLIBAllDBTypesSQLTable.A = txt.Text
        '    objLIBAllDBTypesSQLTable.B = txt.Text
        '    objLIBAllDBTypesSQLTable.C = txt.Text
        '    objLIBAllDBTypesSQLTable.D = txt.Text
        '    objLIBAllDBTypesSQLTable.E = txt.Text
        '    objLIBAllDBTypesSQLTable.F = txt.Text
        '    objLIBAllDBTypesSQLTable.G = txt.Text
        '    objLIBAllDBTypesSQLTable.H = txt.Text
        '    objLIBAllDBTypesSQLTable.I = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable.J = txt.Text
        '    objLIBAllDBTypesSQLTable.K = txt.Text
        '    objLIBAllDBTypesSQLTable.L = txt.Text
        '    objLIBAllDBTypesSQLTable.M = txt.Text
        '    objLIBAllDBTypesSQLTable.N = txt.Text
        '    objLIBAllDBTypesSQLTable.O = txt.Text
        '    objLIBAllDBTypesSQLTable.P = txt.Text
        '    objLIBAllDBTypesSQLTable.Q = txt.Text
        '    objLIBAllDBTypesSQLTable.R = txt.Text
        '    objLIBAllDBTypesSQLTable.S = txt.Text
        '    objLIBAllDBTypesSQLTable.T = txt.Text
        '    objLIBAllDBTypesSQLTable.U = txt.Text
        '    objLIBAllDBTypesSQLTable.V = txt.Text
        '    objLIBAllDBTypesSQLTable.W = txt.Text
        '    objLIBAllDBTypesSQLTable.X = txt.Text
        '    objLIBAllDBTypesSQLTable.Y = txt.Text
        '    objLIBAllDBTypesSQLTable.Z = txt.Text
        '    objLIBAllDBTypesSQLTable.A1 = txt.Text
        '    objLIBAllDBTypesSQLTable.B1 = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable.C1 = txt.Text
        '    objLIBAllDBTypesSQLTable.D1 = txt.Text
        '    objLIBAllDBTypesSQLTable.E1 = txt.Text
        '    tp.MessagePacket = objLIBAllDBTypesSQLTable
        '    tp = objDALAllDBTypesSQLTable.InsertAllDBTypesSQLTable(tp)

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
        Public Function InsertAllDBTypesSQLTable(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim strOutParamValues As String()
            Dim objParamList As New List(Of SqlParameter)()
            Dim objParamListOut As New List(Of SqlParameter)()
            Dim Result As Int16 = 0
            Try
                Dim objLIBAllDBTypesSQLTable As New LIBAllDBTypesSQLTable
                objLIBAllDBTypesSQLTable = Packet.MessagePacket

                objParamList.Add(New SqlParameter("@ID", objLIBAllDBTypesSQLTable.ID))
                objParamList.Add(New SqlParameter("@A", objLIBAllDBTypesSQLTable.A))
                objParamList.Add(New SqlParameter("@B", objLIBAllDBTypesSQLTable.B))
                objParamList.Add(New SqlParameter("@C", objLIBAllDBTypesSQLTable.C))
                objParamList.Add(New SqlParameter("@D", objLIBAllDBTypesSQLTable.D))
                objParamList.Add(New SqlParameter("@E", objLIBAllDBTypesSQLTable.E))
                objParamList.Add(New SqlParameter("@F", objLIBAllDBTypesSQLTable.F))
                objParamList.Add(New SqlParameter("@G", objLIBAllDBTypesSQLTable.G))
                objParamList.Add(New SqlParameter("@H", objLIBAllDBTypesSQLTable.H))
                objParamList.Add(New SqlParameter("@I", objLIBAllDBTypesSQLTable.I))
                objParamList.Add(New SqlParameter("@J", objLIBAllDBTypesSQLTable.J))
                objParamList.Add(New SqlParameter("@K", objLIBAllDBTypesSQLTable.K))
                objParamList.Add(New SqlParameter("@L", objLIBAllDBTypesSQLTable.L))
                objParamList.Add(New SqlParameter("@M", objLIBAllDBTypesSQLTable.M))
                objParamList.Add(New SqlParameter("@N", objLIBAllDBTypesSQLTable.N))
                objParamList.Add(New SqlParameter("@O", objLIBAllDBTypesSQLTable.O))
                objParamList.Add(New SqlParameter("@P", objLIBAllDBTypesSQLTable.P))
                objParamList.Add(New SqlParameter("@Q", objLIBAllDBTypesSQLTable.Q))
                objParamList.Add(New SqlParameter("@R", objLIBAllDBTypesSQLTable.R))
                objParamList.Add(New SqlParameter("@S", objLIBAllDBTypesSQLTable.S))
                objParamList.Add(New SqlParameter("@T", objLIBAllDBTypesSQLTable.T))
                objParamList.Add(New SqlParameter("@U", objLIBAllDBTypesSQLTable.U))
                objParamList.Add(New SqlParameter("@V", objLIBAllDBTypesSQLTable.V))
                objParamList.Add(New SqlParameter("@W", objLIBAllDBTypesSQLTable.W))
                objParamList.Add(New SqlParameter("@X", objLIBAllDBTypesSQLTable.X))
                objParamList.Add(New SqlParameter("@Y", objLIBAllDBTypesSQLTable.Y))
                objParamList.Add(New SqlParameter("@Z", objLIBAllDBTypesSQLTable.Z))
                objParamList.Add(New SqlParameter("@A1", objLIBAllDBTypesSQLTable.A1))
                objParamList.Add(New SqlParameter("@B1", objLIBAllDBTypesSQLTable.B1))
                objParamList.Add(New SqlParameter("@C1", objLIBAllDBTypesSQLTable.C1))
                objParamList.Add(New SqlParameter("@D1", objLIBAllDBTypesSQLTable.D1))
                objParamList.Add(New SqlParameter("@E1", objLIBAllDBTypesSQLTable.E1))
                objParamListOut.Add(New SqlParameter("@@ID", SqlDbType.Int))
                strOutParamValues = MyCLS.clsExecuteStoredProcSql.ExecuteSPNonQueryOutPut("SP_InsertAllDBTypesSQLTable", objParamList, objParamListOut, Packet.MessageId)
                Packet.MessageResultset = strOutParamValues

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function

    End Class
End Namespace