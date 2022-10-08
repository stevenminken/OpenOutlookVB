Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Stimulsoft.Report
Imports Stimulsoft.Report.Export
Imports Outlook = Microsoft.Office.Interop.Outlook

' Programma om Stimulsoft . mrt file te exporteren naar .rtf bestand en deze te openen in Outlook mail 
Module Module1

    Sub Main()

        Dim application = GetApplicationObject()
        CreateSendItem(application)

    End Sub

    ' Exporteert mrt bestand en opent mail 
    Private Function CreateSendItem(ByVal oApp As Outlook.Application)

        Dim mailItem As Outlook.MailItem = Nothing
        Dim mailRecipients As Outlook.Recipients = Nothing
        Dim mailRecipient As Outlook.Recipient = Nothing
        Dim report As StiReport

        Dim fullPath As String = "c:\Users\Steven\Documents\Reports\KleinBevestigingDatumTijd.mrt"

        report = New StiReport()
        report.Load(fullPath)
        report.Render(False)

        Dim rtfSettings As StiRtfExportSettings = New StiRtfExportSettings()
        rtfSettings.ImageQuality = 1.0F
        report.ExportDocument(StiExportFormat.Rtf, "c:\Users\Steven\Documents\Reports\KleinBevestigingDatumTijd.rtf", rtfSettings)
        Console.WriteLine("The export action is complete.", "Export Report")

        mailItem = CType(oApp.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        mailItem.[To] = "stevenminken@hotmail.com"
        mailItem.Subject = "A programatically generated e-mail"
        mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText
        mailItem.RTFBody = System.Text.Encoding.ASCII.GetBytes(My.Computer.FileSystem.ReadAllText("c:\Users\Steven\Documents\Reports\KleinBevestigingDatumTijd.rtf"))
        mailItem.Display()

        Return Nothing

    End Function

    ' Retourneert het actieve Outlook object
    Function GetApplicationObject() As Outlook.Application

        Dim application As Outlook.Application

        ' Check whether there is an Outlook process running.
        If Process.GetProcessesByName("OUTLOOK").Count() > 0 Then

            ' If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
            application = DirectCast(Marshal.GetActiveObject("Outlook.Application"), Outlook.Application)
        Else

            ' If not, create a new instance of Outlook and sign in to the default profile.
            application = New Outlook.Application()
            Dim ns As Outlook.NameSpace = application.GetNamespace("MAPI")
            ns.Logon("", "", Missing.Value, Missing.Value)
            ns = Nothing
        End If

        ' Return the Outlook Application object.
        Return application
    End Function

End Module
