Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            Dim fPDF As New clsPDFMaker()
            Dim FormType As String = ""
            Dim strYA As String = ""

            FormType = fPDF.GetFormType()
            strYA = fPDF.GetYA()

            fPDF.CloseStamper()

            'if FormType is Empty then exit application
            If String.IsNullOrEmpty(FormType) Then Exit Sub

            'if Year of Assessment is Empty then exit application
            If String.IsNullOrEmpty(strYA) Then Exit Sub

            'else select the specific form
            Dim objBorang As New clsBorangHandler(strYA, FormType)
            objBorang.CreateBorang()

            'Close application
            e.Cancel = True

        End Sub

    End Class

End Namespace

