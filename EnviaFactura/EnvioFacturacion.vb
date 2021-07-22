Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net.Mail

Public Class EnvioFacturacion

    Friend WithEvents SBOApplication As SAPbouiCOM.Application
    Public SBOCompany As SAPbobsCOM.Company
    Dim pdf, xml, pdfSAP, xmlSAP As String

    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Function Facturacion(ByVal DocNum As String, ByVal Tipo As String, ByVal psDirectory As String)

        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String
        Dim DocEntry, CardCode, CardName, DocDate, EmailC, ReportId As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Call EnvioCorreo_SemiAutomatico('" & DocNum & "','" & Tipo & "')"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                DocEntry = oRecSettxb.Fields.Item("DocEntry").Value
                CardCode = oRecSettxb.Fields.Item("CardCode").Value
                CardName = oRecSettxb.Fields.Item("CardName").Value
                DocDate = oRecSettxb.Fields.Item("DocDate").Value
                EmailC = oRecSettxb.Fields.Item("E_Mail").Value
                ReportId = oRecSettxb.Fields.Item("ReportID").Value

                ValidarDoc(DocEntry, ReportId, Tipo, DocDate, CardCode, DocNum, EmailC)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en la funcion Cotizacion. " & ex.Message)

        End Try

    End Function


    Public Function ValidarDoc(ByVal DocEntry As String, ByVal ReportID As String, ByVal Tipo As String, ByVal DocDate As Date, ByVal CardCode As String, ByVal DocNum As String, ByVal EmailC As String)

        'MsgBox("Exportar Documento Exitoso")
        Dim Ruta, RutaSAP As String
        pdf = Nothing
        xml = Nothing
        pdfSAP = Nothing
        xmlSAP = Nothing

        Try

            Ruta = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\IN"
            RutaSAP = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\IN"

            Dim dir As New System.IO.DirectoryInfo(Ruta)

            Dim fileList = dir.GetFiles("*.pdf", System.IO.SearchOption.TopDirectoryOnly)

            Dim FileQuery = From file In fileList
                            Where file.Extension = ".pdf" And file.Name.Trim.ToString.EndsWith(ReportID & ".pdf") And file.Name.Trim.ToString.StartsWith(ReportID & ".pdf")
                            Order By file.CreationTime
                            Select file

            pdf = Ruta & "\" & ReportID & ".pdf"
            pdfSAP = RutaSAP & "\" & ReportID & ".pdf"

            Dim fileList1 = dir.GetFiles("*.xml", System.IO.SearchOption.TopDirectoryOnly)

            Dim fileQuery1 = From file In fileList1
                             Where file.Extension = ".xml" And file.Name.Trim.ToString.EndsWith(ReportID & ".xml") And file.Name.Trim.ToString.StartsWith(ReportID & ".xml")
                             Order By file.CreationTime
                             Select file

            xml = Ruta & "\" & ReportID & ".xml"
            xmlSAP = RutaSAP & "\" & ReportID & ".xml"

            If FileQuery.Count > 0 And fileQuery1.Count > 0 Then

                If EmailC <> "" Then

                    UpdatePDFXML(DocNum, pdfSAP, xmlSAP)
                    EnviarCorreo(DocNum, EmailC, pdf, xml, Tipo, pdfSAP, xmlSAP, CardCode)

                Else

                    SBOApplication.MessageBox("El socio de negocios no tiene asignado un correo electronico")

                End If

            ElseIf FileQuery.Count = 0 And fileQuery1.Count > 0 Then

                ExportarPDF(DocEntry, ReportID, Tipo, DocDate, CardCode, DocNum, EmailC)

            ElseIf fileQuery1.Count = 0 Then

                SBOApplication.MessageBox("A la factura no se le ha creado un xml")

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en ValidarDoc. " & ex.Message)

        End Try

    End Function


    Public Function ExportarPDF(ByVal DocEntry As String, ByVal ReportId As String, ByVal Tipo As String, ByVal DocDate As Date, ByVal CardCode As String, ByVal DocNum As String, ByVal EmailC As String)

        'MsgBox("Consulta de Documentos exitosa")
        Dim reportDocument As ReportDocument
        Dim diskFileDestinationOption As DiskFileDestinationOptions

        Try

            reportDocument = New ReportDocument

            reportDocument.Load("C:\TareasProgramadas\EnvioCorreos\FC2.rpt")

            Dim count As Integer = reportDocument.DataSourceConnections.Count
            reportDocument.DataSourceConnections(0).SetLogon(My.Settings.DbUserName, My.Settings.DbPassword)

            reportDocument.SetParameterValue(0, DocEntry)

            reportDocument.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            reportDocument.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
            diskFileDestinationOption = New DiskFileDestinationOptions

            diskFileDestinationOption.DiskFileName = My.Settings.Ruta & "\" & DocDate.ToString("yyyy-MM") & "\" & CardCode & "\IN\" & ReportId & ".pdf"

            reportDocument.ExportOptions.ExportDestinationOptions = diskFileDestinationOption
            reportDocument.ExportOptions.ExportFormatOptions = New PdfRtfWordFormatOptions

            reportDocument.Export()
            'MsgBox("Exportacion de Documento Exitosa")
            reportDocument.Close()
            reportDocument.Dispose()
            GC.SuppressFinalize(reportDocument)

            UpdatePDFXML(DocNum, pdfSAP, xmlSAP)

            If EmailC <> "" Then

                EnviarCorreo(DocNum, EmailC, pdf, xml, Tipo, pdfSAP, xmlSAP, CardCode)

            Else

                SBOApplication.MessageBox("El socio de negocios no tiene asignado un correo electronico")

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en ExportarPDF. " & ex.Message)

        End Try

    End Function


    Public Function UpdatePDFXML(ByVal DocNum As String, ByVal pdfSAP As String, ByVal xmlSAP As String)

        Dim oRecSettxb1, oRecSettxb2 As SAPbobsCOM.Recordset
        Dim stQuerytxb1, stQuerytxb2 As String

        Try

            oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb1 = "Update OINV set ""U_XML""='" & xmlSAP & "' where ""DocNum""=" & DocNum
            oRecSettxb1.DoQuery(stQuerytxb1)

            oRecSettxb2 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb2 = "Update OINV set ""U_PDF""='" & pdfSAP & "' where ""DocNum""=" & DocNum
            oRecSettxb2.DoQuery(stQuerytxb2)

        Catch ex As Exception

            SBOApplication.MessageBox("Error en UpdatePDFXML. " & ex.Message)

        End Try

    End Function


    Public Function EnviarCorreo(ByVal DocNum As String, ByVal EmailC As String, ByVal pdf As String, ByVal xml As String, ByVal Tipo As String, ByVal pdfSAP As String, ByVal xmlSAP As String, ByVal CardCode As String)

        'MsgBox("Validacion de Documentos exitosa")
        Dim message As New MailMessage
        Dim smtp As New SmtpClient
        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String
        Dim EmailU, Pass, EmailCC, Subject, Body, smtpService, Puerto, SegSSL As String

        Try

            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Select ""U_Email"",""U_Password"",""U_EmailCC"",""U_Subject"",""U_Body"",""U_SMTP"",""U_Puerto"",""U_SeguridadSSL"" from ""@CORREOTEKNO"" where ""Name""='Automático'"
            oRecSettxb.DoQuery(stQuerytxb)

            If oRecSettxb.RecordCount > 0 Then

                oRecSettxb.MoveFirst()

                EmailU = oRecSettxb.Fields.Item("U_Email").Value
                Pass = oRecSettxb.Fields.Item("U_Password").Value
                EmailCC = oRecSettxb.Fields.Item("U_EmailCC").Value

                Subject = oRecSettxb.Fields.Item("U_Subject").Value
                Body = oRecSettxb.Fields.Item("U_Body").Value
                smtpService = oRecSettxb.Fields.Item("U_SMTP").Value
                Puerto = oRecSettxb.Fields.Item("U_Puerto").Value
                SegSSL = oRecSettxb.Fields.Item("U_SeguridadSSL").Value

                'Limpiamos correo destinatario, correo copia y archivos adjuntos
                message.To.Clear()
                message.CC.Clear()
                message.Attachments.Clear()

                'Llenamos encabezado de correo
                message.From = New MailAddress(EmailU)
                EmailC = ArreglarTexto(EmailC, ";", ",")
                message.To.Add(EmailC)

                If EmailCC.Count > 0 Then
                    message.CC.Add(EmailCC)
                End If

                message.Subject = Subject & " Factura " & DocNum

                'Llenamos el cuerpo del correo y prioridad
                message.Body = Body
                message.Priority = MailPriority.Normal

                'Adjuntamos archivos xml y pdf
                Dim attxml As New Net.Mail.Attachment(xml)
                message.Attachments.Add(attxml)

                Dim attpdf As New Net.Mail.Attachment(pdf)
                message.Attachments.Add(attpdf)

                'Llenamos datos de smtp
                smtp.Host = smtpService
                smtp.Credentials = New Net.NetworkCredential(EmailU, Pass)
                smtp.Port = Puerto
                smtp.EnableSsl = SegSSL

                'Enviamos Correo
                smtp.Send(message)

                SBOApplication.MessageBox("El correo se envio correctamente.")
                UpdateCorreoEnviado(DocNum)

            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error en EnviarCorreo. " & ex.Message)

        End Try

    End Function


    Public Function UpdateCorreoEnviado(ByVal DocNum As String)

        Dim oRecSettxb1 As SAPbobsCOM.Recordset
        Dim stQuerytxb1 As String

        Try

            oRecSettxb1 = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb1 = "Update OINV set ""U_TekEnviado""='Y' where ""DocNum""=" & DocNum
            oRecSettxb1.DoQuery(stQuerytxb1)

        Catch ex As Exception

            SBOApplication.MessageBox("Error en UpdateCorreoEnviado. " & ex.Message)

        End Try

    End Function


    Public Function ArreglarTexto(ByVal TextoOriginal As String, ByVal QuitarCaracter As String, ByVal PonerCaracter As String)

        TextoOriginal = TextoOriginal.Replace(QuitarCaracter, PonerCaracter)
        Return TextoOriginal

    End Function


End Class