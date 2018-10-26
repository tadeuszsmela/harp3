Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Private Sub MyApplication_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
            Dim a As Integer

            '            MessageBox.Show("KOniec pracy - " & aktualny_uzytkownik.id)
            If aktualny_uzytkownik.id > 0 Then
                a = zapisz_log_wejsc_wyjsc(aktualny_uzytkownik.id, 0)
            End If



            Try
                polaczenie_sql.Close()
            Catch ex As Exception

            End Try
        End Sub
        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup


            HARPApp.Initialize()
            MainForm = HARPApp.MainForm



        End Sub

        Private Sub MyApplication_StartupNextInstance(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupNextInstanceEventArgs) Handles Me.StartupNextInstance

            Dim msg As String
            Dim a As Integer

            msg = "HARP jest ju¿ uruchomiony na tym komputerze."
            msg = msg & vbCrLf & "Czy uruchomiæ program w oddzielnym oknie ?"

            a = MessageBox.Show(msg, "HARP", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If a = vbNo Then
                End
            End If

        End Sub
    End Class

End Namespace
