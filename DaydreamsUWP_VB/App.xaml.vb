''' <summary>
''' Fornisci un comportamento specifico dell'applicazione in supplemento alla classe Application predefinita.
''' </summary>
NotInheritable Class App
    Inherits Application

    ''' <summary>
    ''' Richiamato quando l'applicazione viene avviata normalmente dall'utente. All'avvio dell'applicazione
    ''' verranno usati altri punti di ingresso per aprire un file specifico, per visualizzare
    ''' risultati della ricerca e così via.
    ''' </summary>
    ''' <param name="e">Dettagli sulla richiesta e sul processo di avvio.</param>
    Protected Overrides Sub OnLaunched(e As Windows.ApplicationModel.Activation.LaunchActivatedEventArgs)
        Dim rootFrame As Frame = TryCast(Window.Current.Content, Frame)

        ' Non ripetere l'inizializzazione dell'applicazione se la finestra già dispone di contenuto,
        ' assicurarsi solo che la finestra sia attiva

        If rootFrame Is Nothing Then
            ' Creare un frame che agisca da contesto di navigazione e passare alla prima pagina
            rootFrame = New Frame()

            AddHandler rootFrame.NavigationFailed, AddressOf OnNavigationFailed

            If e.PreviousExecutionState = ApplicationExecutionState.Terminated Then
                ' TODO: caricare lo stato dall'applicazione sospesa in precedenza
            End If
            ' Posizionare il frame nella finestra corrente
            Window.Current.Content = rootFrame
        End If

        If e.PrelaunchActivated = False Then
            If rootFrame.Content Is Nothing Then
                ' Quando lo stack di esplorazione non viene ripristinato, passare alla prima pagina
                ' e configurare la nuova pagina passando le informazioni richieste come parametro
                ' parametro
                rootFrame.Navigate(GetType(MainPage), e.Arguments)
            End If

            ' Assicurarsi che la finestra corrente sia attiva
            Window.Current.Activate()
        End If
    End Sub

    ''' <summary>
    ''' Chiamato quando la navigazione a una determinata pagina ha esito negativo
    ''' </summary>
    ''' <param name="sender">Frame la cui navigazione non è riuscita</param>
    ''' <param name="e">Dettagli sull'errore di navigazione.</param>
    Private Sub OnNavigationFailed(sender As Object, e As NavigationFailedEventArgs)
        Throw New Exception("Failed to load Page " + e.SourcePageType.FullName)
    End Sub

    ''' <summary>
    ''' Richiamato quando l'esecuzione dell'applicazione viene sospesa. Lo stato dell'applicazione viene salvato
    ''' senza che sia noto se l'applicazione verrà terminata o ripresa con il contenuto
    ''' della memoria ancora integro.
    ''' </summary>
    ''' <param name="sender">Origine della richiesta di sospensione.</param>
    ''' <param name="e">Dettagli relativi alla richiesta di sospensione.</param>
    Private Sub OnSuspending(sender As Object, e As SuspendingEventArgs) Handles Me.Suspending
        Dim deferral As SuspendingDeferral = e.SuspendingOperation.GetDeferral()
        ' TODO: Salvare lo stato dell'applicazione e interrompere qualsiasi attività in background
        deferral.Complete()
    End Sub

End Class
