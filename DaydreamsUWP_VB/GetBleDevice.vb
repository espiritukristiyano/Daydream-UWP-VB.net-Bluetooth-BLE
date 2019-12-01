Imports Windows.Devices.Enumeration
Imports Windows.UI.Core

Public Class GetBleDevice

    Private KnownDevices As ObservableCollection(Of DeviceInformation) = New ObservableCollection(Of DeviceInformation)()
    Private deviceWatcher As DeviceWatcher

    Public Event GetBleDevices(DevicesBle As ObservableCollection(Of DeviceInformation))

    Private dispatcher As CoreDispatcher

    '#######################################################################################################################################
    '################################################################################## ROUTINE NEW WITH DISPATCHER
    Public Sub New(iDispatcher As CoreDispatcher)
        dispatcher = iDispatcher
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE START BLE DIVICE WATCH
    Public Sub StartBleDeviceWatcher()

        Dim requestedProperties As String() = {"System.Devices.Aep.DeviceAddress", "System.Devices.Aep.IsConnected", "System.Devices.Aep.Bluetooth.Le.IsConnectable"}
        Dim aqsAllBluetoothLEDevices As String = "(System.Devices.Aep.ProtocolId:=""{bb7bb05e-5972-42b5-94fc-76eaa7084d49}"")"

        deviceWatcher = DeviceInformation.CreateWatcher(aqsAllBluetoothLEDevices, requestedProperties, DeviceInformationKind.AssociationEndpoint)

        AddHandler deviceWatcher.Added, AddressOf DeviceWatcher_Added
        AddHandler deviceWatcher.Updated, AddressOf DeviceWatcher_Updated
        AddHandler deviceWatcher.Removed, AddressOf DeviceWatcher_Removed
        AddHandler deviceWatcher.EnumerationCompleted, AddressOf DeviceWatcher_EnumerationCompleted
        AddHandler deviceWatcher.Stopped, AddressOf DeviceWatcher_Stopped

        KnownDevices.Clear()
        deviceWatcher.Start()
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE STOP BLE DIVICE WATCH
    Public Sub StopBleDeviceWatcher()
        If deviceWatcher IsNot Nothing Then
            RemoveHandler deviceWatcher.Added, AddressOf DeviceWatcher_Added
            RemoveHandler deviceWatcher.Updated, AddressOf DeviceWatcher_Updated
            RemoveHandler deviceWatcher.Removed, AddressOf DeviceWatcher_Removed
            RemoveHandler deviceWatcher.EnumerationCompleted, AddressOf DeviceWatcher_EnumerationCompleted
            RemoveHandler deviceWatcher.Stopped, AddressOf DeviceWatcher_Stopped

            deviceWatcher.Stop()
            deviceWatcher = Nothing
        End If
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE FIND BLE DEVICE DISPLAY
    Private Function FindBluetoothLEDeviceDisplay(ByVal id As String) As DeviceInformation
        For Each dev As DeviceInformation In KnownDevices
            If id = dev.Id Then '###### IF YOU FOUND IT
                Return dev
            End If
        Next
        Return Nothing
    End Function

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT ADDED
    Private Async Sub DeviceWatcher_Added(ByVal sender As DeviceWatcher, ByVal deviceInfo As DeviceInformation)
        Await dispatcher.RunAsync(CoreDispatcherPriority.Normal, Sub()
                                                                     SyncLock Me
                                                                         If deviceWatcher.Equals(sender) Then
                                                                             KnownDevices.Add(deviceInfo)
                                                                             '###### RAISE EVENT GET BLE DEVICES
                                                                             RaiseEvent GetBleDevices(KnownDevices)
                                                                         End If
                                                                     End SyncLock
                                                                 End Sub)
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT UPDATED
    Private Async Sub DeviceWatcher_Updated(ByVal sender As DeviceWatcher, ByVal deviceInfoUpdate As DeviceInformationUpdate)
        Await dispatcher.RunAsync(CoreDispatcherPriority.Normal, Sub()
                                                                     SyncLock Me
                                                                         If deviceWatcher.Equals(sender) Then
                                                                             Dim dev As DeviceInformation = FindBluetoothLEDeviceDisplay(deviceInfoUpdate.Id)
                                                                             If dev IsNot Nothing Then
                                                                                 '###### 
                                                                                 KnownDevices.Remove(dev)
                                                                                 dev.Update(deviceInfoUpdate)
                                                                                 KnownDevices.Add(dev)
                                                                                 '###### RAISE EVENT GET BLE DEVICES
                                                                                 RaiseEvent GetBleDevices(KnownDevices)
                                                                                 Return
                                                                             End If
                                                                         End If
                                                                     End SyncLock
                                                                 End Sub)
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT REMOVED
    Private Async Sub DeviceWatcher_Removed(ByVal sender As DeviceWatcher, ByVal deviceInfoUpdate As DeviceInformationUpdate)
        Await dispatcher.RunAsync(CoreDispatcherPriority.Normal, Sub()
                                                                     SyncLock Me
                                                                         If deviceWatcher.Equals(sender) Then
                                                                             Dim dev As DeviceInformation = FindBluetoothLEDeviceDisplay(deviceInfoUpdate.Id)
                                                                             If dev IsNot Nothing Then
                                                                                 KnownDevices.Remove(dev)
                                                                                 '###### RAISE EVENT GET BLE DEVICES
                                                                                 RaiseEvent GetBleDevices(KnownDevices)
                                                                             End If
                                                                         End If
                                                                     End SyncLock
                                                                 End Sub)
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT ENUMERATION COMPLETED
    Private Async Sub DeviceWatcher_EnumerationCompleted(ByVal sender As DeviceWatcher, ByVal e As Object)
        Await dispatcher.RunAsync(CoreDispatcherPriority.Normal, Sub()
                                                                     If deviceWatcher.Equals(sender) Then
                                                                         '###### RAISE EVENT GET BLE DEVICES
                                                                         RaiseEvent GetBleDevices(KnownDevices)
                                                                     End If
                                                                 End Sub)
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT STOPPED
    Private Async Sub DeviceWatcher_Stopped(ByVal sender As DeviceWatcher, ByVal e As Object)
        Await dispatcher.RunAsync(CoreDispatcherPriority.Normal, Sub()
                                                                     If deviceWatcher.Equals(sender) Then
                                                                         If sender.Status = DeviceWatcherStatus.Aborted Then
                                                                             Throw New Exception("No longer watching for devices.")
                                                                         End If

                                                                     End If
                                                                 End Sub)
    End Sub

End Class
