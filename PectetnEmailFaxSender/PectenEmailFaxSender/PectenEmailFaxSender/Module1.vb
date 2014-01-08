Imports System.Windows.Forms
Imports PectenEmailFaxSender.EmailFax
Imports system.Net.Mail
Imports System
Imports System.Net
Imports SubSystems.RP
Imports System.IO
Module Module1
    Public Sub writetoconsole(ByVal text As String)
        Console.WriteLine(text)
    End Sub
    Sub Main()

        Dim s As EmailFaxTableAdapters.tblEmailFaxTableAdapter = Nothing
        Dim Dt As EmailFax.tblEmailFaxDataTable
        Dim Dr As EmailFax.tblEmailFaxRow

        Try

            writetoconsole("starting the application")

            s = New EmailFaxTableAdapters.tblEmailFaxTableAdapter
            Dt = s.SelectEmailFax()
            Dim PathOfDocumentToConvert As String = ""
            For Each Dr In Dt.Rows
                Try
                    PathOfDocumentToConvert = ""
                    Dim PathofPDFConvertedDocument As String = ""

                    'MM 05/31/10
                    '//if there is not attachment then dont bother filling those fields
                    If Not String.IsNullOrEmpty(Dr.AttachmentFile.ToString.Trim) Then

                        PathOfDocumentToConvert = Dr.AttachmentFile.ToString.Trim
                        'PathofPDFConvertedDocument = ConverToPdf(PathOfDocumentToConvert) 'Have to do this
                        PathofPDFConvertedDocument = PathOfDocumentToConvert 'Have to do this
                        Dr.PDFFile = PathofPDFConvertedDocument
                    Else
                        PathofPDFConvertedDocument = ""
                    End If

                    SendEmail(Dr)
                    s.UpdateEmailFaxSendFlag(Dr.rowid, DateTime.Now, PathofPDFConvertedDocument, "")
                Catch ex As Exception
                    s.UpdateEmailFaxSendFlag(Dr.rowid, Nothing, Nothing, ex.ToString)
                End Try

            Next

        Catch ex As Exception
            writetoconsole(ex.ToString)
            logerror(ex)
        Finally
            s = Nothing
            Dt = Nothing
            Dr = Nothing
            writetoconsole("Done processing!")
            'Console.ReadLine()
        End Try

    End Sub


    Public Sub logerror(ByVal ex As Exception)

        Dim s As EmailFaxTableAdapters.EmailFaxConfigErrorsTableAdapter = Nothing
        Try

            s = New EmailFaxTableAdapters.EmailFaxConfigErrorsTableAdapter()
            s.InsertEmailFaxConfigErrors("Errors Occured while trying to run." + ex.ToString)
        Catch
        Finally
            s = Nothing
        End Try


    End Sub

    Public Sub SendEmail(ByVal DrEmailfax As EmailFax.tblEmailFaxRow)

        Dim s As EmailFaxTableAdapters.TblEmailFaxConfigTableAdapter
        Dim dtemailfaxconfig As TblEmailFaxConfigDataTable
        Dim DrEmailFaxConfig As EmailFax.TblEmailFaxConfigRow
        Dim recepientsstr As String = "'"
        Try

            s = New EmailFaxTableAdapters.TblEmailFaxConfigTableAdapter()

            dtemailfaxconfig = s.SelectEmailFaxConfig()
            DrEmailFaxConfig = dtemailfaxconfig.Item(0)

            recepientsstr = DrEmailfax.RecepientStr.ToString.Trim
            recepientsstr = recepientsstr.Replace("  ", ",")
            recepientsstr = recepientsstr.Replace(";", ",")

            If recepientsstr.LastIndexOf(",") = recepientsstr.Length - 1 Then
                recepientsstr = recepientsstr.Substring(0, recepientsstr.Length - 1)
            End If

            Dim fmailaddress As MailAddress = New MailAddress(DrEmailFaxConfig.FromEmail.ToString.Trim, DrEmailFaxConfig.FromName.Trim)

            Dim mail As MailMessage = New MailMessage(DrEmailFaxConfig.FromEmail.ToString.Trim(), recepientsstr, DrEmailfax.Subject.Trim, DrEmailfax.Body.Trim)
            mail.From = fmailaddress
            mail.IsBodyHtml = True
            'MM 05/31/10
            '//if there is an attachment then add it.
            If Not String.IsNullOrEmpty(DrEmailfax.PDFFile.ToString.Trim) Then

                Dim a As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(DrEmailfax.PDFFile.ToString.Trim)
                mail.Attachments.Add(a)

            End If

            Dim c As SmtpClient = New SmtpClient(DrEmailFaxConfig.SMTP, DrEmailFaxConfig.SendingPort)
            c.DeliveryMethod = SmtpDeliveryMethod.Network
            c.EnableSsl = False
            c.UseDefaultCredentials = False

            Dim oCredential As NetworkCredential = New NetworkCredential(DrEmailFaxConfig.FromEmail.ToString.Trim(), DrEmailFaxConfig.password.ToString.Trim())
            c.Credentials = oCredential
            c.Send(mail)
            c = Nothing

        Catch ex As Exception

            Throw ex
        Finally

        End Try
    End Sub

    Public Function ConverToPdf(ByVal PathOfDocumentToConvert As String)

        Dim PDFDocumentPath As String = ""
        Dim rp As Rpn
        Dim result As Boolean = False

        Try

            If PathOfDocumentToConvert.ToUpper.Trim.EndsWith("PDF") Then
                Return PathOfDocumentToConvert
            Else

                rp = New Rpn()
                PDFDocumentPath = PathOfDocumentToConvert.Replace(".", "") + ".pdf"

                Try
                    If File.Exists(PDFDocumentPath) Then File.Delete(PDFDocumentPath)
                Catch
                End Try


                result = rp.RpsConvertFile(PathOfDocumentToConvert, PDFDocumentPath)
                Threading.Thread.Sleep(New TimeSpan(0, 0, 1))
                If result = False Then
                    Throw New Exception("Not able to convert the file " + PathOfDocumentToConvert + " to pdf.")
                Else
                    Return PDFDocumentPath

                End If

            End If

        Catch ex As Exception
            'Return ""
            Throw ex


        Finally
            rp = Nothing
        End Try



    End Function


End Module
