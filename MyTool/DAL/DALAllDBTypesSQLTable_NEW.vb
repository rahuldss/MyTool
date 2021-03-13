Imports MyTool.NDS.LIB
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient

Namespace NDS.DAL

    Public Class DALAllDBTypesSQLTable_NEW

        'PUT IT IN LOAD EVENTS

        'MyCLS.strConnStringOLEDB = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;Provider=SQLOLEDB.1"
        'MyCLS.strConnStringSQLCLIENT = "Initial Catalog=AB;Data Source=127.0.0.1;UID=sa;PWD=sa123;"

        '*******COPY IT TO USE BELOW FUNCTION - SELECT ALL************
        'Try
        '    Dim objLIBAllDBTypesSQLTable_NEWListing As New LIBAllDBTypesSQLTable_NEWListing
        '    Dim objDALAllDBTypesSQLTable_NEW As New DALAllDBTypesSQLTable_NEW
        '    Dim tp As New MyCLS.TransportationPacket
        '    Dim ds As New Dataset

        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).ID
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).bigint_A
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).binary_B
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).bit_C
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).char_D
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).datetime_E
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).decimal_F
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).decimal_G
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).float_H
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).image_I
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).int_J
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).money_K
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).nchar_L
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).ntext_M
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).numeric_N
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).numeric_O
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).nvarchar_P
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).nvarchar_Max_Q
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).real_R
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).smalldatetime_S
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).smallint_T
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).smallmoney_U
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).sql_variant_V
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).text_W
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).timestamp_X
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).tinyint_Y
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).uniqueidentifier_Z
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varbinary_A1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varbinary_Max_B1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varchar_C1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varchar_Max_D1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).xml_E1
        '    tp = objDALAllDBTypesSQLTable_NEW.GetAllDBTypesSQLTable_NEWDetails()
        '    If tp.MessageId = 1 Then
        '        objLIBAllDBTypesSQLTable_NEWListing = tp.MessageResultset
        '        ds = tp.MessageResultsetDS
        '        MyCLS.clsImaging.ByteArray2Image(,)
        '        MsgBox(objLIBAllDBTypesSQLTable_NEWListing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=Nothing, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllDBTypesSQLTable_NEWDetails() As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objLIBAllDBTypesSQLTable_NEWListing As New LIBAllDBTypesSQLTable_NEWListing
            Dim Packet As New MyCLS.TransportationPacket

            Try
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromAllDBTypesSQLTable_NEW")
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBAllDBTypesSQLTable_NEW As New LIBAllDBTypesSQLTable_NEW
                                    objLIBAllDBTypesSQLTable_NEWListing.Add(oLIBAllDBTypesSQLTable_NEW)
                                    objLIBAllDBTypesSQLTable_NEWListing(i).ID = ds.Tables(0).Rows(i)("ID").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).bigint_A = ds.Tables(0).Rows(i)("bigint_A").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).binary_B = ds.Tables(0).Rows(i)("binary_B")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).bit_C = ds.Tables(0).Rows(i)("bit_C").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).char_D = ds.Tables(0).Rows(i)("char_D").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).datetime_E = ds.Tables(0).Rows(i)("datetime_E").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).decimal_F = ds.Tables(0).Rows(i)("decimal_F").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).decimal_G = ds.Tables(0).Rows(i)("decimal_G").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).float_H = ds.Tables(0).Rows(i)("float_H").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).image_I = ds.Tables(0).Rows(i)("image_I")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).int_J = ds.Tables(0).Rows(i)("int_J").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).money_K = ds.Tables(0).Rows(i)("money_K").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).nchar_L = ds.Tables(0).Rows(i)("nchar_L").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).ntext_M = ds.Tables(0).Rows(i)("ntext_M").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).numeric_N = ds.Tables(0).Rows(i)("numeric_N").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).numeric_O = ds.Tables(0).Rows(i)("numeric_O").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).nvarchar_P = ds.Tables(0).Rows(i)("nvarchar_P").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).nvarchar_Max_Q = ds.Tables(0).Rows(i)("nvarchar_Max_Q").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).real_R = ds.Tables(0).Rows(i)("real_R").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).smalldatetime_S = ds.Tables(0).Rows(i)("smalldatetime_S").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).smallint_T = ds.Tables(0).Rows(i)("smallint_T").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).smallmoney_U = ds.Tables(0).Rows(i)("smallmoney_U").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).sql_variant_V = ds.Tables(0).Rows(i)("sql_variant_V").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).text_W = ds.Tables(0).Rows(i)("text_W").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).timestamp_X = ds.Tables(0).Rows(i)("timestamp_X")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).tinyint_Y = ds.Tables(0).Rows(i)("tinyint_Y").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).uniqueidentifier_Z = ds.Tables(0).Rows(i)("uniqueidentifier_Z")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varbinary_A1 = ds.Tables(0).Rows(i)("varbinary_A1")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varbinary_Max_B1 = ds.Tables(0).Rows(i)("varbinary_Max_B1")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varchar_C1 = ds.Tables(0).Rows(i)("varchar_C1").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varchar_Max_D1 = ds.Tables(0).Rows(i)("varchar_Max_D1").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).xml_E1 = ds.Tables(0).Rows(i)("xml_E1").ToString
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If
                Packet.MessageResultsetDS = ds
                Packet.MessageResultset = objLIBAllDBTypesSQLTable_NEWListing

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


        '*******COPY IT TO USE BELOW FUNCTION - SELECT BY ID************
        'Try
        '    Dim objLIBAllDBTypesSQLTable_NEWListing As New LIBAllDBTypesSQLTable_NEWListing
        '    Dim objDALAllDBTypesSQLTable_NEW As New DALAllDBTypesSQLTable_NEW
        '    Dim tp As New MyCLS.TransportationPacket
        '    Dim ds As New Dataset
        '    tp.MessagePacket = 1    'ID to be Passed

        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).ID
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).bigint_A
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).binary_B
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).bit_C
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).char_D
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).datetime_E
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).decimal_F
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).decimal_G
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).float_H
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).image_I
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).int_J
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).money_K
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).nchar_L
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).ntext_M
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).numeric_N
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).numeric_O
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).nvarchar_P
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).nvarchar_Max_Q
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).real_R
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).smalldatetime_S
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).smallint_T
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).smallmoney_U
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).sql_variant_V
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).text_W
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).timestamp_X
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).tinyint_Y
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).uniqueidentifier_Z
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varbinary_A1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varbinary_Max_B1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varchar_C1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).varchar_Max_D1
        '' txt.Text = objLIBAllDBTypesSQLTable_NEWListing(0).xml_E1
        '    tp = objDALAllDBTypesSQLTable_NEW.GetAllDBTypesSQLTable_NEWDetails(tp)
        '    If tp.MessageId = 1 Then
        '        objLIBAllDBTypesSQLTable_NEWListing = tp.MessageResultset
        '        ds = tp.MessageResultsetDS
        '        MyCLS.clsImaging.ByteArray2Image(,)
        '        MsgBox(objLIBAllDBTypesSQLTable_NEWListing(0))
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try

        ''' <summary>
        ''' Accepts=TransportationPacket, Return=Packet, Result=Packet.MessageId, Return Values=Packet.MessageResultset
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllDBTypesSQLTable_NEWDetails(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim ds As New DataSet
            Dim objParamList As New List(Of SqlParameter)()
            Dim objLIBAllDBTypesSQLTable_NEWListing As New LIBAllDBTypesSQLTable_NEWListing

            Try
                objParamList.Add(New SqlParameter("@Id", Packet.MessagePacket))
                ds = MyCLS.clsExecuteStoredProcSql.ExecuteSPDataSet("SP_GetDetailsFromAllDBTypesSQLTable_NEWById", objParamList)
                If ds IsNot Nothing Then
                    If ds.Tables IsNot Nothing Then
                        If ds.Tables(0).Rows IsNot Nothing Then
                            If ds.Tables(0).Rows.Count > 0 Then
                                For i As Int16 = 0 To ds.Tables(0).Rows.Count - 1
                                    Dim oLIBAllDBTypesSQLTable_NEW As New LIBAllDBTypesSQLTable_NEW
                                    objLIBAllDBTypesSQLTable_NEWListing.Add(oLIBAllDBTypesSQLTable_NEW)
                                    objLIBAllDBTypesSQLTable_NEWListing(i).ID = ds.Tables(0).Rows(i)("ID").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).bigint_A = ds.Tables(0).Rows(i)("bigint_A").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).binary_B = ds.Tables(0).Rows(i)("binary_B")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).bit_C = ds.Tables(0).Rows(i)("bit_C").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).char_D = ds.Tables(0).Rows(i)("char_D").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).datetime_E = ds.Tables(0).Rows(i)("datetime_E").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).decimal_F = ds.Tables(0).Rows(i)("decimal_F").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).decimal_G = ds.Tables(0).Rows(i)("decimal_G").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).float_H = ds.Tables(0).Rows(i)("float_H").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).image_I = ds.Tables(0).Rows(i)("image_I")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).int_J = ds.Tables(0).Rows(i)("int_J").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).money_K = ds.Tables(0).Rows(i)("money_K").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).nchar_L = ds.Tables(0).Rows(i)("nchar_L").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).ntext_M = ds.Tables(0).Rows(i)("ntext_M").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).numeric_N = ds.Tables(0).Rows(i)("numeric_N").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).numeric_O = ds.Tables(0).Rows(i)("numeric_O").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).nvarchar_P = ds.Tables(0).Rows(i)("nvarchar_P").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).nvarchar_Max_Q = ds.Tables(0).Rows(i)("nvarchar_Max_Q").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).real_R = ds.Tables(0).Rows(i)("real_R").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).smalldatetime_S = ds.Tables(0).Rows(i)("smalldatetime_S").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).smallint_T = ds.Tables(0).Rows(i)("smallint_T").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).smallmoney_U = ds.Tables(0).Rows(i)("smallmoney_U").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).sql_variant_V = ds.Tables(0).Rows(i)("sql_variant_V").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).text_W = ds.Tables(0).Rows(i)("text_W").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).timestamp_X = ds.Tables(0).Rows(i)("timestamp_X")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).tinyint_Y = ds.Tables(0).Rows(i)("tinyint_Y").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).uniqueidentifier_Z = ds.Tables(0).Rows(i)("uniqueidentifier_Z")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varbinary_A1 = ds.Tables(0).Rows(i)("varbinary_A1")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varbinary_Max_B1 = ds.Tables(0).Rows(i)("varbinary_Max_B1")
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varchar_C1 = ds.Tables(0).Rows(i)("varchar_C1").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).varchar_Max_D1 = ds.Tables(0).Rows(i)("varchar_Max_D1").ToString
                                    objLIBAllDBTypesSQLTable_NEWListing(i).xml_E1 = ds.Tables(0).Rows(i)("xml_E1").ToString
                                Next
                                Packet.MessageId = 1
                            Else
                                Packet.MessageId = -1
                            End If
                        End If
                    End If
                End If

                Packet.MessageResultsetDS = ds
                Packet.MessageResultset = objLIBAllDBTypesSQLTable_NEWListing

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function


        '*******COPY IT TO USE BELOW FUNCTION - INSERT************
        'Try
        '    Dim objLIBAllDBTypesSQLTable_NEW As New LIBAllDBTypesSQLTable_NEW
        '    Dim objDALAllDBTypesSQLTable_NEW As New DALAllDBTypesSQLTable_NEW
        '    Dim tp As New MyCLS.TransportationPacket

        '    objLIBAllDBTypesSQLTable_NEW.ID = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.bigint_A = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.binary_B = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable_NEW.bit_C = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.char_D = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.datetime_E = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.decimal_F = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.decimal_G = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.float_H = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.image_I = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable_NEW.int_J = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.money_K = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.nchar_L = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.ntext_M = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.numeric_N = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.numeric_O = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.nvarchar_P = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.nvarchar_Max_Q = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.real_R = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.smalldatetime_S = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.smallint_T = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.smallmoney_U = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.sql_variant_V = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.text_W = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.timestamp_X = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable_NEW.tinyint_Y = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.uniqueidentifier_Z = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.varbinary_A1 = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable_NEW.varbinary_Max_B1 = MyCLS.clsImaging.PictureBoxToByteArray()
        '    objLIBAllDBTypesSQLTable_NEW.varchar_C1 = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.varchar_Max_D1 = txt.Text
        '    objLIBAllDBTypesSQLTable_NEW.xml_E1 = txt.Text
        '    tp.MessagePacket = objLIBAllDBTypesSQLTable_NEW
        '    tp = objDALAllDBTypesSQLTable_NEW.InsertAllDBTypesSQLTable_NEW(tp)

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
        Public Function InsertAllDBTypesSQLTable_NEW(ByVal Packet As MyCLS.TransportationPacket) As MyCLS.TransportationPacket
            Dim strOutParamValues As String()
            Dim objParamList As New List(Of SqlParameter)()
            Dim objParamListOut As New List(Of SqlParameter)()
            Dim Result As Int16 = 0
            Try
                Dim objLIBAllDBTypesSQLTable_NEW As New LIBAllDBTypesSQLTable_NEW
                objLIBAllDBTypesSQLTable_NEW = Packet.MessagePacket

                objParamList.Add(New SqlParameter("@ID", objLIBAllDBTypesSQLTable_NEW.ID))
                objParamList.Add(New SqlParameter("@bigint_A", objLIBAllDBTypesSQLTable_NEW.bigint_A))
                objParamList.Add(New SqlParameter("@binary_B", objLIBAllDBTypesSQLTable_NEW.binary_B))
                objParamList.Add(New SqlParameter("@bit_C", objLIBAllDBTypesSQLTable_NEW.bit_C))
                objParamList.Add(New SqlParameter("@char_D", objLIBAllDBTypesSQLTable_NEW.char_D))
                objParamList.Add(New SqlParameter("@datetime_E", objLIBAllDBTypesSQLTable_NEW.datetime_E))
                objParamList.Add(New SqlParameter("@decimal_F", objLIBAllDBTypesSQLTable_NEW.decimal_F))
                objParamList.Add(New SqlParameter("@decimal_G", objLIBAllDBTypesSQLTable_NEW.decimal_G))
                objParamList.Add(New SqlParameter("@float_H", objLIBAllDBTypesSQLTable_NEW.float_H))
                objParamList.Add(New SqlParameter("@image_I", objLIBAllDBTypesSQLTable_NEW.image_I))
                objParamList.Add(New SqlParameter("@int_J", objLIBAllDBTypesSQLTable_NEW.int_J))
                objParamList.Add(New SqlParameter("@money_K", objLIBAllDBTypesSQLTable_NEW.money_K))
                objParamList.Add(New SqlParameter("@nchar_L", objLIBAllDBTypesSQLTable_NEW.nchar_L))
                objParamList.Add(New SqlParameter("@ntext_M", objLIBAllDBTypesSQLTable_NEW.ntext_M))
                objParamList.Add(New SqlParameter("@numeric_N", objLIBAllDBTypesSQLTable_NEW.numeric_N))
                objParamList.Add(New SqlParameter("@numeric_O", objLIBAllDBTypesSQLTable_NEW.numeric_O))
                objParamList.Add(New SqlParameter("@nvarchar_P", objLIBAllDBTypesSQLTable_NEW.nvarchar_P))
                objParamList.Add(New SqlParameter("@nvarchar_Max_Q", objLIBAllDBTypesSQLTable_NEW.nvarchar_Max_Q))
                objParamList.Add(New SqlParameter("@real_R", objLIBAllDBTypesSQLTable_NEW.real_R))
                objParamList.Add(New SqlParameter("@smalldatetime_S", objLIBAllDBTypesSQLTable_NEW.smalldatetime_S))
                objParamList.Add(New SqlParameter("@smallint_T", objLIBAllDBTypesSQLTable_NEW.smallint_T))
                objParamList.Add(New SqlParameter("@smallmoney_U", objLIBAllDBTypesSQLTable_NEW.smallmoney_U))
                objParamList.Add(New SqlParameter("@sql_variant_V", objLIBAllDBTypesSQLTable_NEW.sql_variant_V))
                objParamList.Add(New SqlParameter("@text_W", objLIBAllDBTypesSQLTable_NEW.text_W))
                objParamList.Add(New SqlParameter("@timestamp_X", objLIBAllDBTypesSQLTable_NEW.timestamp_X))
                objParamList.Add(New SqlParameter("@tinyint_Y", objLIBAllDBTypesSQLTable_NEW.tinyint_Y))
                objParamList.Add(New SqlParameter("@uniqueidentifier_Z", objLIBAllDBTypesSQLTable_NEW.uniqueidentifier_Z))
                objParamList.Add(New SqlParameter("@varbinary_A1", objLIBAllDBTypesSQLTable_NEW.varbinary_A1))
                objParamList.Add(New SqlParameter("@varbinary_Max_B1", objLIBAllDBTypesSQLTable_NEW.varbinary_Max_B1))
                objParamList.Add(New SqlParameter("@varchar_C1", objLIBAllDBTypesSQLTable_NEW.varchar_C1))
                objParamList.Add(New SqlParameter("@varchar_Max_D1", objLIBAllDBTypesSQLTable_NEW.varchar_Max_D1))
                objParamList.Add(New SqlParameter("@xml_E1", objLIBAllDBTypesSQLTable_NEW.xml_E1))
                objParamListOut.Add(New SqlParameter("@@ID", SqlDbType.Int))
                strOutParamValues = MyCLS.clsExecuteStoredProcSql.ExecuteSPNonQueryOutPut("SP_InsertAllDBTypesSQLTable_NEW", objParamList, objParamListOut, Packet.MessageId)
                Packet.MessageResultset = strOutParamValues

            Catch ex As Exception
                Packet.MessageId = -1
                MyCLS.clsHandleException.HandleEx(ex, System.Reflection.MethodBase.GetCurrentMethod.ToString())
            End Try
            Return Packet
        End Function

    End Class
End Namespace