
'DEVICE INFORMATION
'https://docs.microsoft.com/it-it/uwp/api/Windows.Devices.Enumeration.DeviceInformation
'CLIENT GATT
'https://docs.microsoft.com/it-it/windows/uwp/devices-sensors/gatt-client

Imports Windows.Devices.Enumeration
Imports Windows.Devices.Bluetooth.GenericAttributeProfile

''' <summary>
''' Pagina vuota che può essere usata autonomamente oppure per l'esplorazione all'interno di un frame.
''' </summary>
Public NotInheritable Class MainPage
    Inherits Page

    Dim BleDevices As ObservableCollection(Of DeviceInformation) = New ObservableCollection(Of DeviceInformation)()
    Dim BleServices As ObservableCollection(Of GattDeviceService) = New ObservableCollection(Of GattDeviceService)()
    Dim BleCharacteristic As ObservableCollection(Of GattCharacteristic) = New ObservableCollection(Of GattCharacteristic)()

    Dim SelectedBleDeviceId As String

    WithEvents GetBleDev As New GetBleDevice(Dispatcher)
    WithEvents BleMan As New BleManagement(Dispatcher)

    '#######################################################################################################################################
    '################################################################################## ROUTINE MAIN PAGE
    Private Sub MainPage_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

    End Sub

#Region "     ################################################################################## GET DEVICES BLE AVAILABLE"


    '#######################################################################################################################################
    '################################################################################## ROUTINE BUTTON WATCH DEVICES BLE
    Private Sub BtnWatch_Click(sender As Object, e As RoutedEventArgs) Handles btnWatch.Click
        GetBleDev.StartBleDeviceWatcher()
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT GET BLE DEVICES AVAILABLE
    Private Sub GetBleDev_GetBleDevices(DevicesBle As ObservableCollection(Of DeviceInformation)) Handles GetBleDev.GetBleDevices

        BleDevices = DevicesBle

        lstBleDevices.Items.Clear()
        For Each dev As DeviceInformation In BleDevices
            lstBleDevices.Items.Add(dev.Name)
        Next

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE SELECTION CHANGED LIST OF DEVICES
    Private Sub DeviceInterfacesOutputList_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstBleDevices.SelectionChanged
        '###### 
        GetBleDev.StopBleDeviceWatcher()
        '###### Save the selected device's ID for use in other scenarios.
        For Each dev As DeviceInformation In BleDevices
            If dev.Name = lstBleDevices.SelectedItem Then
                SelectedBleDeviceId = dev.Id
                Exit For
            End If
        Next
    End Sub

#End Region

    '#######################################################################################################################################
    '#######################################################################################################################################
    '#######################################################################################################################################
#Region "     ################################################################################## GET BLE SERVICES/CHARACTERISTIC AVAILABLE"


    '#######################################################################################################################################
    '################################################################################## ROUTINE BUTTON CONNECT SERVICES->CHARACTERISTIC->(NOTIFY/READ)
    Private Async Sub BtnConnect_Click(sender As Object, e As RoutedEventArgs) Handles btnConnect.Click

        BleDevices.Clear()

        '###### RESET BLE
        BleMan.ResetBle()

        '###### GET ALL SERVICES FROM SELECTED BLE DEVICE ID
        Await BleMan.GetServices(SelectedBleDeviceId)

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE BUTTON RESET
    Private Sub BtnReset_Click(sender As Object, e As RoutedEventArgs) Handles btnReset.Click
        BleMan.ResetBle()
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT GET SERVICES COMPLETED
    Private Async Sub BleMan_GetServicesCompleted(ServicesBle As ObservableCollection(Of GattDeviceService)) Handles BleMan.GetServicesCompleted

        '###### 
        BleServices = ServicesBle

        '###### GET ALL SERVICES
        lstBleServices.Items.Clear()
        For Each service As GattDeviceService In BleServices
            '###### INSERT INTO LISTBOX
            lstBleServices.Items.Add(service.Uuid)
        Next

        '###### GET CHARACTERISTIC FROM BLE SERVICE
        Await BleMan.GetCharacteristic(BleServices(3)) '###### DAYDREAM
        'Await BleMan.GetCharacteristic(BleServices(5)) '###### BATTERY


    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT GET CHARACTERISITC COMPLETED
    Private Sub BleMan_GetCharacteristicCompleted(CharacteristicBle As ObservableCollection(Of GattCharacteristic)) Handles BleMan.GetCharacteristicCompleted

        '###### 
        BleCharacteristic = CharacteristicBle

        '###### GET ALL CHARACTERISTIC
        lstBleCharacteristic.Items.Clear()
        For Each characteristic As GattCharacteristic In BleCharacteristic
            '###### INSERT INTO LISTBOX
            lstBleCharacteristic.Items.Add(characteristic.Uuid)
        Next

        '###### PREPARE CHARACTERISTIC (FOR READ/NOTIFY) FROM BLE CHARACTERISTIC
        BleMan.PrepareCharacteristic(BleCharacteristic(0))

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT PREPARE CHARACTERISITC COMPLETED
    Private Sub BleMan_PrepareCharacteristicCommpleted() Handles BleMan.PrepareCharacteristicCommpleted
        '###### ACTIVE NOTIFY CHARACTERISTIC
        BleMan.ActiveNotify(BleCharacteristic(0))
        '###### READ VALUE FROM CHARACTERISTIC
        'BleMan.ReadCharacteristic(BleCharacteristic(0))

    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT BLE NOTIFY CHANGED
    Private Sub BleMan_BleNotifyChanged(data() As Byte, presentationFormat As GattPresentationFormat, selectedCharacteristic As GattCharacteristic) Handles BleMan.BleNotifyChanged
        '###### FOR OTHER USE
        'OutputText.Text = FormatValueByPresentation(data, presentationFormat, selectedCharacteristic)
        '###### DAYDREAMS
        GetDaydreamProperty(DDControl, data, selectedCharacteristic)
    End Sub

    '#######################################################################################################################################
    '################################################################################## ROUTINE EVENT BLE READ VALUE
    Private Sub BleMan_BleReadCharacteristic(data() As Byte, presentationFormat As GattPresentationFormat, selectedCharacteristic As GattCharacteristic) Handles BleMan.BleReadCharacteristic
        '###### FOR OTHER USE
        'OutputText.Text = FormatValueByPresentation(data, presentationFormat, selectedCharacteristic)
        '###### DAYDREAMS
        GetDaydreamProperty(DDControl, data, selectedCharacteristic)
    End Sub

#End Region

#Region "     ################################################################################## GET PROPERTY DAYDREAM"

    Structure DEF_DAYDREAMS
        Dim isClickDown As Boolean
        Dim isAppDown As Boolean
        Dim isHomeDown As Boolean
        Dim isVolPlusDown As Boolean
        Dim isVolMinusDown As Boolean

        Dim Time As Integer
        Dim PPS As Byte

        Dim xOri As Integer
        Dim yOri As Integer
        Dim zOri As Integer
        Dim xAcc As Integer
        Dim yAcc As Integer
        Dim zAcc As Integer

        Dim xGyro As Integer
        Dim yGyro As Integer
        Dim zGyro As Integer

        Dim xTouch As Integer
        Dim yTouch As Integer

        Dim BatteryLevel As Byte
    End Structure

    Dim DDControl As DEF_DAYDREAMS

    '#######################################################################################################################################
    '################################################################################## ROUTINE GET DAYDREAM PROPERTY
    Sub GetDaydreamProperty(DDControl As DEF_DAYDREAMS, ByVal data As Byte(), GCH As GattCharacteristic)

        If data IsNot Nothing Then

            '###### DAYDREAM SERVICE
            If GCH.Uuid.Equals(New Guid("00000001-1000-1000-8000-00805F9B34FB")) Then

                Try

                    With DDControl

                        .isClickDown = (data(18) And &H1) > 0
                        .isAppDown = (data(18) And &H4) > 0
                        .isHomeDown = (data(18) And &H2) > 0
                        .isVolPlusDown = (data(18) And &H10) > 0
                        .isVolMinusDown = (data(18) And &H8) > 0

                        .Time = (data(0) And &HFF) << 1 Or data(0) And &H80
                        .PPS = (data(1) And &H7C) >> 2

                        .xOri = ((data(1) And &H3) << 11 Or (data(2) And &HFF) << 3 Or (data(3) And &H80) >> 5) << 19 >> 19
                        .yOri = (data(3) And &H1F) << 8 Or data(4) And &HFF << 19 >> 19
                        .zOri = ((data(5) And &HFF) << 5 Or (data(6) And &HF8) >> 3) << 19 >> 19
                        .xAcc = ((data(6) And &H7) << 10 Or (data(7) And &HFF) << 2 Or (data(8) And &HC0) >> 6) << 19 >> 19
                        .yAcc = ((data(8) And &H3F) << 7 Or (data(9) And &HFE) >> 1) << 19 >> 19
                        .zAcc = ((data(9) And &H1) << 12 Or (data(10) And &HFF) << 4 Or (data(11) And &HF0) >> 4) << 19 >> 19

                        .xGyro = ((data(11) And &HF) << 9 Or (data(12) And &HFF) << 1 Or (data(13) And &H80) >> 7) << 19 >> 19
                        .yGyro = ((data(13) And &H7F) << 6 Or (data(14) And &HFC) >> 2) << 19 >> 19
                        .zGyro = ((data(14) And &H3) << 11 Or (data(15) And &HFF) << 3 Or (data(16) And &HE0) >> 5) << 19 >> 19

                        .xTouch = (data(16) And &H1F) << 3 Or (data(17) And &HE0) >> 5
                        .yTouch = (data(17) And &H1F) << 3 Or (data(18) And &HE0) >> 5

                        '###### PRINTED ONLY FOR TEST
                        Dim v = "isClickDown:" & Convert.ToString(.isClickDown) & vbLf
                        v += " isAppDown:" & Convert.ToString(.isAppDown) & vbLf
                        v += " isHomeDown:" & Convert.ToString(.isHomeDown) & vbLf
                        v += " isVolPlusDown:" & Convert.ToString(.isVolPlusDown) & vbLf
                        v += " isVolMinusDown:" & Convert.ToString(.isVolMinusDown) & vbLf

                        v += " Time:" & Convert.ToString(.Time) & vbLf
                        v += " Seq:" & Convert.ToString(.PPS) & vbLf

                        v += "xOri:" & Convert.ToString(.xOri) & vbLf
                        v += "yOri:" & Convert.ToString(.yOri) & vbLf
                        v += "zOri:" & Convert.ToString(.zOri) & vbLf

                        v += "xAcc:" & Convert.ToString(.xAcc) & vbLf
                        v += "yAcc:" & Convert.ToString(.yAcc) & vbLf
                        v += "zAcc:" & Convert.ToString(.zAcc) & vbLf

                        v += "xGyro:" & Convert.ToString(.xGyro) & vbLf
                        v += "yGyro:" & Convert.ToString(.yGyro) & vbLf
                        v += "zGyro:" & Convert.ToString(.zGyro) & vbLf

                        v += "xTouch:" & Convert.ToString(.xTouch) & vbLf
                        v += "yTouch:" & Convert.ToString(.yTouch) & vbLf
                        OutputText.Text = v
                        '######

                    End With

                    Return
                Catch ExArg As ArgumentException
                    Throw New Exception("(unable to parse)")
                    Return
                End Try

            ElseIf GCH.Uuid.Equals(GattCharacteristicUuids.BatteryLevel) Then '###### BATTERY

                Try
                    '###### GET BATTERY LEVEL
                    DDControl.BatteryLevel = data(0)

                    '###### PRINTED ONLY FOR TEST
                    OutputText.Text = "Battery Level: " & DDControl.BatteryLevel & "%"
                    '######
                    Return
                Catch ExArg As ArgumentException
                    Throw New Exception("Battery Level: (unable to parse)")
                    Return
                End Try
            End If

        Else
            Throw New Exception("Empty data received")
            Return
        End If

        Throw New Exception("Unknown format ()")

    End Sub

#End Region

    '#######################################################################################################################################
    '################################################################################## ROUTINE FOR OTHER USE
    Private Function FormatValueByPresentation(ByVal data As Byte(), ByVal format As GattPresentationFormat, GCH As GattCharacteristic) As String
        Dim v = ""

        If format IsNot Nothing Then
            Stop
            'If format.FormatType = GattPresentationFormatTypes.UInt32 AndAlso data.Length >= 4 Then
            '    Return BitConverter.ToInt32(data, 0).ToString
            'ElseIf format.FormatType = GattPresentationFormatTypes.Utf8 Then
            '    Try
            '        Return Encoding.UTF8.GetString(data)
            '    Catch ExArg As ArgumentException
            '        Return "(error: Invalid UTF-8 string)"
            '    End Try
            'Else
            '    Return "Unsupported format: " & CryptographicBuffer.EncodeToHexString(buffer)
            'End If
        ElseIf data IsNot Nothing Then

            '###### MORE CONTROL
            If GCH.Uuid.Equals(GattCharacteristicUuids.HeartRateMeasurement) Then

                Try

                    '###### CODE CONTROL

                    Return ""
                Catch ExArg As ArgumentException
                    Return "(unable to parse)"
                End Try

            ElseIf GCH.Uuid.Equals(GattCharacteristicUuids.BatteryLevel) Then '###### BATTERY

                Try
                    Return "Battery Level: " & data(0).ToString & "%"
                Catch ExArg As ArgumentException
                    Return "Battery Level: (unable to parse)"
                End Try
            End If

        Else
            Return "Empty data received"
        End If

        Return "Unknown format ()"
    End Function


End Class
