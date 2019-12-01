Imports Windows.Devices.Bluetooth
Imports Windows.Devices.Bluetooth.GenericAttributeProfile
Imports Windows.Devices.Enumeration
Imports Windows.UI.Core

Public Class BleManagement

    Private BleServices As ObservableCollection(Of GattDeviceService) = New ObservableCollection(Of GattDeviceService)()
    Private BleCharacteristic As ObservableCollection(Of GattCharacteristic) = New ObservableCollection(Of GattCharacteristic)()
    Private bluetoothLeDevice As BluetoothLEDevice = Nothing

    Private selectedCharacteristic As GattCharacteristic
    Private registeredCharacteristic As GattCharacteristic
    Private presentationFormat As GattPresentationFormat

    Private dispatcher As CoreDispatcher

    Public Event GetServicesCompleted(ServicesBle As ObservableCollection(Of GattDeviceService))
    Public Event GetCharacteristicCompleted(CharacteristicBle As ObservableCollection(Of GattCharacteristic))
    Public Event PrepareCharacteristicCommpleted()
    Public Event BleNotifyChanged(data As Byte(), presentationFormat As GattPresentationFormat, selectedCharacteristic As GattCharacteristic)
    Public Event BleReadCharacteristic(data As Byte(), presentationFormat As GattPresentationFormat, selectedCharacteristic As GattCharacteristic)


    '#######################################################################################################################################
    '################################################################################## ROUTINE NEW WITH DISPATCHER
    Public Sub New(iDispatcher As CoreDispatcher)
        dispatcher = iDispatcher
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE RESET BLE 
    Public Sub ResetBle()

        RemoveValueChangedHandler()

        BleServices.Clear()
        BleCharacteristic.Clear()

        bluetoothLeDevice?.Dispose()
        bluetoothLeDevice = Nothing

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE GET SERVICES
    Public Async Function GetServices(SelectedBleDeviceId As String) As Task

        Try
            '###### CONNETTITI AL DISPOSITIVO BLUETOOTH
            bluetoothLeDevice = Await BluetoothLEDevice.FromIdAsync(SelectedBleDeviceId)
            '###### NO CONNECT
            If bluetoothLeDevice Is Nothing Then
                Throw New Exception("Failed to connect to device.")
                Exit Function
            End If
        Catch ex As Exception When ex.HResult = &H800710DF 'E_DEVICE_NOT_AVAILABLE
            Throw New Exception("Bluetooth radio is not on.")
            Exit Function
        End Try

        '###### IS DEVICE BLUETOOTH CONNECTED
        If bluetoothLeDevice IsNot Nothing Then
            Dim result As GattDeviceServicesResult = Await bluetoothLeDevice.GetGattServicesAsync(BluetoothCacheMode.Uncached) '###### RESTITUISCE I SERVIZI DISPONIBILI

            '###### SE LA CONNESSIONE E ANDATA A BUON FINE
            If result.Status = GattCommunicationStatus.Success Then
                Dim services = result.Services

                '###### GET ALL SERVICES
                For Each service In services
                    '###### PREPARE NEW COLLECTION SERVICE
                    BleServices.Add(service)
                Next

                '###### RAISE EVENT GET SERVICE COMPLETED
                RaiseEvent GetServicesCompleted(BleServices)

            Else
                Throw New Exception("Device unreachable")
            End If
        End If

    End Function

    '#######################################################################################################################################
    '################################################################################## ROUTINE GET CHARACTERISTIC
    Public Async Function GetCharacteristic(GDS As GattDeviceService) As Task

        '###### 
        BleCharacteristic.Clear()
        RemoveValueChangedHandler() '###### REMOVE CONNECTION EVENT CHANGE CHARACTERISTIC

        Dim characteristics As IReadOnlyList(Of GattCharacteristic) = Nothing

        Try
            Dim accessStatus = Await GDS.RequestAccessAsync

            If accessStatus = DeviceAccessStatus.Allowed Then
                Dim result = Await GDS.GetCharacteristicsAsync(BluetoothCacheMode.Uncached)

                If result.Status = GattCommunicationStatus.Success Then
                    characteristics = result.Characteristics
                Else
                    Throw New Exception("Error accessing service.")
                    characteristics = New List(Of GattCharacteristic)
                End If
            Else
                Throw New Exception("Error accessing service.")
                characteristics = New List(Of GattCharacteristic)
            End If

        Catch ex As Exception
            Throw New Exception("Restricted service. Can't read characteristics: " & ex.Message)
            characteristics = New List(Of GattCharacteristic)
        End Try

        '###### GET ALL CHARACTERISTIC
        For Each c In characteristics
            '###### PREPARE NEW COLLECTION CHARACTERISTIC
            BleCharacteristic.Add(c)
        Next

        '###### RAISE EVENT GET CHARACTERISTIC COMPLETED
        RaiseEvent GetCharacteristicCompleted(BleCharacteristic)

    End Function

    '#######################################################################################################################################
    '################################################################################## ROUTINE PREPARE CHARACTERISTIC
    Public Async Sub PrepareCharacteristic(GCH As GattCharacteristic)

        If GCH Is Nothing Then
            Throw New Exception("Attribute Info Disp not set.")
            Return
        End If

        '###### FOR EVENT CHANGECHARACTERISTIC
        selectedCharacteristic = GCH

        If GCH Is Nothing Then
            Throw New Exception("No characteristic selected")
            Return
        End If

        Dim result = Await GCH.GetDescriptorsAsync(BluetoothCacheMode.Uncached)

        If result.Status <> GattCommunicationStatus.Success Then
            Throw New Exception("Descriptor read failure: " & result.Status.ToString)
        End If

        presentationFormat = Nothing

        If GCH.PresentationFormats.Count > 0 Then
            If GCH.PresentationFormats.Count.Equals(1) Then
                presentationFormat = GCH.PresentationFormats(0)
            Else
            End If
        End If

        '###### RAISE EVENT PREPARE CHARACTERISTIC COMPLETED
        RaiseEvent PrepareCharacteristicCommpleted()

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE ACTIVE NOTIFY
    Public Async Sub ActiveNotify(GCH As GattCharacteristic)

        Dim status = GattCommunicationStatus.Unreachable
        Dim cccdValue = GattClientCharacteristicConfigurationDescriptorValue.None

        If GCH.CharacteristicProperties.HasFlag(GattCharacteristicProperties.Indicate) Then
            cccdValue = GattClientCharacteristicConfigurationDescriptorValue.Indicate
        ElseIf GCH.CharacteristicProperties.HasFlag(GattCharacteristicProperties.Notify) Then
            cccdValue = GattClientCharacteristicConfigurationDescriptorValue.Notify
        End If

        Try
            status = Await GCH.WriteClientCharacteristicConfigurationDescriptorAsync(cccdValue)

            If status = GattCommunicationStatus.Success Then
                registeredCharacteristic = GCH
                AddHandler registeredCharacteristic.ValueChanged, AddressOf Characteristic_ValueChanged
            Else
                Throw New Exception("Error registering for value changes: " & status)
            End If

        Catch ex As UnauthorizedAccessException
            Throw New Exception(ex.Message)
        End Try

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE READ CHARACTERISTIC
    Public Async Sub ReadCharacteristic(GCH As GattCharacteristic)
        Dim result = Await GCH.ReadValueAsync(BluetoothCacheMode.Uncached)
        Dim data As Byte() = result.Value.ToArray()

        If result.Status = GattCommunicationStatus.Success Then
            '###### RAISE EVENT BLE READ CHARACTERISTIC 
            RaiseEvent BleReadCharacteristic(data, presentationFormat, GCH)
        Else
            '###### RAISE EVENT BLE READ CHARACTERISTIC ERROR
            Throw New Exception("Read failed: " & result.Status)
        End If

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT CHARACTERISTIC VALUE CHANGED
    Private Async Sub Characteristic_ValueChanged(ByVal sender As GattCharacteristic, ByVal args As GattValueChangedEventArgs)
        Await dispatcher.RunAsync(CoreDispatcherPriority.Normal, Sub()
                                                                     '###### RAISE EVENT BLE NOTIFY CHANGED
                                                                     Dim data As Byte() = args.CharacteristicValue.ToArray()
                                                                     RaiseEvent BleNotifyChanged(data, presentationFormat, selectedCharacteristic)
                                                                 End Sub)
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE REMOVE CONNECTION EVENT CHANGE CHARACTERISTIC
    Private Sub RemoveValueChangedHandler()
        If registeredCharacteristic IsNot Nothing Then
            RemoveHandler registeredCharacteristic.ValueChanged, AddressOf Characteristic_ValueChanged
            registeredCharacteristic = Nothing
        End If
    End Sub


End Class
