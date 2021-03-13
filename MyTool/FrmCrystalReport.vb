'************USE TO CALL THIS FORM*********
'Dim objForm As New FrmCrystalReport
''objForm.ViewReport("D:\Narender\Projects\VB.NET\2008\MyTool\MyTool\CrystalReport\rptAddressCodeList.rpt", , "@parameter1=IN000001&parameter2=IN000001")
'objForm.ViewReport("D:\Narender\Projects\VB.NET\2008\MyTool\MyTool\CrystalReport\rptAssignDays.rpt", , )
'objForm.show()
'************USE TO CALL THIS FORM*********

Public Class FrmCrystalReport


    ''Friend Function ViewReport1(ByVal sReportName As String, Optional ByVal sSelectionFormula As String = "", Optional ByVal param As String = "") As Boolean

    ''    Dim intCounter As Integer
    ''    Dim intCounter1 As Integer
    ''    Dim strTableName As String
    ''    Dim objReportsParameters As frmReportsParameters
    ''    Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
    ''    Dim mySection As CrystalDecisions.CrystalReports.Engine.Section
    ''    Dim mySections As CrystalDecisions.CrystalReports.Engine.Sections


    ''    Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo

    ''    Dim paraValue As New CrystalDecisions.Shared.ParameterDiscreteValue
    ''    Dim currValue As CrystalDecisions.Shared.ParameterValues
    ''    Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject
    ''    Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    ''    Dim strParamenters As String
    ''    Dim strParValPair() As String
    ''    Dim strVal() As String
    ''    Dim sFileName As String
    ''    Dim index As Integer

    ''    Try


    ''        sFileName = DownloadReport(sReportName, m_strReportDir)

    ''        objReport.Load(sFileName)

    ''        intCounter = objReport.DataDefinition.ParameterFields.Count
    ''        If intCounter = 1 Then
    ''            If InStr(objReport.DataDefinition.ParameterFields(0).ParameterFieldName, ".", CompareMethod.Text) > 0 Then
    ''                intCounter = 0
    ''            End If
    ''        End If


    ''        If intCounter > 0 And Trim(param) <> "" Then

    ''            strParValPair = strParamenters.Split("&")
    ''            For index = 0 To UBound(strParValPair)
    ''                If InStr(strParValPair(index), "=") > 0 Then
    ''                    strVal = strParValPair(index).Split("=")
    ''                    paraValue.Value = strVal(1)
    ''                    currValue = objReport.DataDefinition.ParameterFields(strVal(0)).CurrentValues
    ''                    currValue.Add(paraValue)
    ''                    objReport.DataDefinition.ParameterFields(strVal(0)).ApplyCurrentValues(currValue)
    ''                End If
    ''            Next
    ''        End If



    ''        ConInfo.ConnectionInfo.UserID = objDataBase.UserName
    ''        ConInfo.ConnectionInfo.Password = objDataBase.Password
    ''        ConInfo.ConnectionInfo.ServerName = objDataBase.Server
    ''        ConInfo.ConnectionInfo.DatabaseName = objDataBase.Database

    ''        For intCounter = 0 To objReport.Database.Tables.Count - 1
    ''            objReport.Database.Tables(intCounter).ApplyLogOnInfo(ConInfo)
    ''        Next



    ''        For index = 0 To objReport.ReportDefinition.Sections.Count - 1
    ''            For intCounter = 0 To objReport.ReportDefinition.Sections(index).ReportObjects.Count - 1
    ''                With objReport.ReportDefinition.Sections(index)
    ''                    If .ReportObjects(intCounter).Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
    ''                        mySubReportObject = CType(.ReportObjects(intCounter), CrystalDecisions.CrystalReports.Engine.SubreportObject)
    ''                        mySubRepDoc = mySubReportObject.OpenSubreport(mySubReportObject.SubreportName)
    ''                        For intCounter1 = 0 To mySubRepDoc.Database.Tables.Count - 1
    ''                            mySubRepDoc.Database.Tables(intCounter1).ApplyLogOnInfo(ConInfo)
    ''                        Next
    ''                    End If
    ''                End With
    ''            Next
    ''        Next





    ''        If sSelectionFormula.Length > 0 Then
    ''            objReport.RecordSelectionFormula = sSelectionFormula
    ''        End If


    ''        rptViewer.ReportSource = Nothing
    ''        rptViewer.ReportSource = objReport
    ''        rptViewer.Show()

    ''        Application.DoEvents()

    ''        Me.Text = sReportName
    ''        MyBase.Visible = True
    ''        Me.BringToFront()

    ''        Return True

    ''    Catch ex As System.Exception
    ''        MsgBox(ex.Message)
    ''    End Try
    ''End Function

    Friend Function ViewReport(ByVal sReportName As String, Optional ByVal sSelectionFormula As String = "", Optional ByVal param As String = "") As Boolean
        'Declaring variablesables
        Dim intCounter As Integer
        Dim intCounter1 As Integer

        'Crystal Report's report document object
        Dim objReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        'object of table Log on info of Crystal report
        Dim ConInfo As New CrystalDecisions.Shared.TableLogOnInfo

        'Parameter value object of crystal report 
        ' parameters used for adding the value to parameter.
        Dim paraValue As New CrystalDecisions.Shared.ParameterDiscreteValue

        'Current parameter value object(collection) of crystal report parameters.
        Dim currValue As CrystalDecisions.Shared.ParameterValues

        'Sub report object of crystal report.
        Dim mySubReportObject As CrystalDecisions.CrystalReports.Engine.SubreportObject

        'Sub report document of crystal report.
        Dim mySubRepDoc As New CrystalDecisions.CrystalReports.Engine.ReportDocument

        Dim strParValPair() As String
        Dim strVal() As String
        Dim index As Integer

        Try

            'Load the report
            objReport.Load(sReportName)

            'Check if there are parameters or not in report.
            intCounter = objReport.DataDefinition.ParameterFields.Count

            'As parameter fields collection also picks the selection 
            ' formula which is not the parameter
            ' so if total parameter count is 1 then we check whether 
            ' its a parameter or selection formula.

            If intCounter = 1 Then
                If InStr(objReport.DataDefinition.ParameterFields(0).ParameterFieldName, ".", CompareMethod.Text) > 0 Then
                    intCounter = 0
                End If
            End If

            'If there are parameters in report and 
            'user has passed them then split the 
            'parameter string and Apply the values 
            'to their concurrent parameters.
            Try
                If intCounter > 0 And Trim(param) <> "" Then
                    strParValPair = param.Split("&")

                    For index = 0 To UBound(strParValPair)
                        If InStr(strParValPair(index), "=") > 0 Then
                            strVal = strParValPair(index).Split("=")
                            paraValue.Value = strVal(1)
                            currValue = objReport.DataDefinition.ParameterFields(strVal(0)).CurrentValues
                            currValue.Add(paraValue)
                            objReport.DataDefinition.ParameterFields(strVal(0)).ApplyCurrentValues(currValue)
                        End If
                    Next
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
            'Set the connection information to ConInfo 
            'object so that we can apply the 
            'connection information on each table in the report
            '*****Added By Me***
            Dim objConnectionInfo As New MyCLS.ConnectionInfo
            objConnectionInfo = MyCLS.clsCOMMON.ConOpenFromXMLFile(True)
            '*****Added By Me***
            ConInfo.ConnectionInfo.UserID = objConnectionInfo.UserID
            ConInfo.ConnectionInfo.Password = objConnectionInfo.Password
            ConInfo.ConnectionInfo.ServerName = objConnectionInfo.ServerName
            ConInfo.ConnectionInfo.DatabaseName = objConnectionInfo.Database

            For intCounter = 0 To objReport.Database.Tables.Count - 1
                objReport.Database.Tables(intCounter).ApplyLogOnInfo(ConInfo)
            Next

            ' Loop through each section on the report then look 
            ' through each object in the section
            ' if the object is a subreport, then apply logon info 
            ' on each table of that sub report

            For index = 0 To objReport.ReportDefinition.Sections.Count - 1
                For intCounter = 0 To objReport.ReportDefinition.Sections(index).ReportObjects.Count - 1
                    With objReport.ReportDefinition.Sections(index)
                        If .ReportObjects(intCounter).Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
                            mySubReportObject = CType(.ReportObjects(intCounter), CrystalDecisions.CrystalReports.Engine.SubreportObject)
                            mySubRepDoc = mySubReportObject.OpenSubreport(mySubReportObject.SubreportName)
                            For intCounter1 = 0 To mySubRepDoc.Database.Tables.Count - 1
                                mySubRepDoc.Database.Tables(intCounter1).ApplyLogOnInfo(ConInfo)
                                'sp;
                                'mySubRepDoc.Database.Tables(intCounter1).ApplyLogOnInfo(ConInfo)
                            Next
                        End If
                    End With
                Next
            Next
            'If there is a selection formula passed to this function then use that
            If sSelectionFormula.Length > 0 Then
                objReport.RecordSelectionFormula = sSelectionFormula
            End If
            'Re setting control 
            rptViewer.ReportSource = Nothing

            'Set the current report object to report.
            rptViewer.ReportSource = objReport

            'Show the report
            rptViewer.Show()
            Return True
        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
    End Function
End Class