# Daydream UWP VB.net Bluetooth BLE

This code written in VB.Net is used to read data from a Bluetooth Ble device.
Born primarily to read data from Google's Daydream you can also use it for other purposes, that is, reading data from other Bluetooth Ble devices.

The code consists mainly of two classes, one to find Bluetooth Ble devices "GetBleDivices()" and another "BleManagement()" to connect with the device and then take its services and features, with the ability to activate notifications where possible or to read more simply the data of a feature.
The part that code that has been taken most into account and then developed is the one concerning the Daydream, in fact you can connect a Daydream and read all the data coming from it that are:

All 5 Buttons
Orientation
Acceleration
Gyro
Touch x, y

When you launch the project we will have a series of buttons and listboxes so divided:
"1) Ble Watch button"
This button starts searching for Bluetooth Ble devices available on Windows and will be reflected in the listbox below the button.
In the Listbox containing the names of the devices found, click Daydream by selecting it.
After this click on the "3) Connect" button to connect the selected device in the previous listbox.
When and if the connection is successful, the two listboxes below will display the services in the first and the features in the second. The latter will be connected to service number 4 (0->3). In fact, in the code it is already all set up in order to connect to the service and then to its characteristic example:

Await BleMan.GetCharacteristic(BleServices(3)) '# DAYDREAM

This line looks for the characteristics of service number 3

Await BleMan.GetCharacteristic(BleServices(5)) '# BATTERY

The one looks for the features of service number 5 which in this case is the Daydream battery.

The code is very simple and you can deduce how it works in all its parts, just have some knowledge of VB and Bluetooth Ble devices.
