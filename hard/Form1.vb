Imports System

Imports System.Management

Public Class Form1
    Structure INFO_HARDDISK
        Dim diskId As String          '0
        Dim diskSerialNumber As String    '1
        Dim diskModel As String           '1
        Dim Availability As UInt16
        Dim Description As String
        Dim DeviceID As String
        Dim Manufacturer As String
        Dim MaxMediaSize As UInt64
        Dim Name As String
        Dim size As UInt64
        Dim TotalCylinders As UInt64
        Dim TotalHeads As UInt32
        Dim TotalSectors As UInt64
        Dim totalTracks As UInt64
        Dim TracksPerCylinder As UInt32

    End Structure
    Public hd(1) As INFO_HARDDISK
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim strTemp As String = ""
        Dim cmicWmi As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDrive")
   
        Dim diskId As String = "" '数字ID
        Dim diskSerialNumber As String = "" '这个我们暂且称其为序列号码
        Dim diskModel As String = "" '序列号
        Dim i As Int32
        i = 0
        For Each cmicWmiObj As ManagementObject In cmicWmi.Get
            If i < 2 Then
                hd(i).Description = cmicWmiObj("Description")
                hd(i).diskId = cmicWmiObj("signature")
                hd(i).diskSerialNumber = cmicWmiObj("serialnumber")
                hd(i).diskModel = cmicWmiObj("Model")
                hd(i).Availability = cmicWmiObj("Availability")
                hd(i).DeviceID = cmicWmiObj("DeviceID")
                hd(i).Manufacturer = cmicWmiObj("Manufacturer")
                hd(i).MaxMediaSize = cmicWmiObj("MaxMediaSize")
                hd(i).Name = cmicWmiObj("Name")
                hd(i).size = cmicWmiObj("size")
                hd(i).TotalCylinders = cmicWmiObj("TotalCylinders")
                hd(i).TotalHeads = cmicWmiObj("TotalHeads")
                hd(i).TotalSectors = cmicWmiObj("TotalSectors")
                hd(i).totalTracks = cmicWmiObj("totalTracks")
                hd(i).TracksPerCylinder = cmicWmiObj("TracksPerCylinder")
            End If

            'diskId = cmicWmiObj("signature")
            'diskSerialNumber = cmicWmiObj("serialnumber")
            'diskModel = cmicWmiObj("Model")
            i += 1
        Next

       

        '        class Win32_DiskDrive : CIM_DiskDrive
        '{
        '  uint16   Availability;
        '  uint32   BytesPerSector;
        '  uint16   Capabilities[];
        '  string   CapabilityDescriptions[];
        '  string   Caption;
        '  string   CompressionMethod;
        '  uint32   ConfigManagerErrorCode;
        '  boolean  ConfigManagerUserConfig;
        '  string   CreationClassName;
        '  uint64   DefaultBlockSize;
        '  string   Description;
        '  string   DeviceID;
        '  boolean  ErrorCleared;
        '  string   ErrorDescription;
        '  string   ErrorMethodology;
        '  string   FirmwareRevision;
        '  uint32   Index;
        '  datetime InstallDate;
        '  string   InterfaceType;
        '  uint32   LastErrorCode;
        '  string   Manufacturer;
        '  uint64   MaxBlockSize;
        '  uint64   MaxMediaSize;
        '  boolean  MediaLoaded;
        '  string   MediaType;
        '  uint64   MinBlockSize;
        '  string   Model;
        '  string   Name;
        '  boolean  NeedsCleaning;
        '  uint32   NumberOfMediaSupported;
        '  uint32   Partitions;
        '  string   PNPDeviceID;
        '  uint16   PowerManagementCapabilities[];
        '  boolean  PowerManagementSupported;
        '  uint32   SCSIBus;
        '  uint16   SCSILogicalUnit;
        '  uint16   SCSIPort;
        '  uint16   SCSITargetId;
        '  uint32   SectorsPerTrack;
        '  string   SerialNumber;
        '  uint32   Signature;
        '  uint64   Size;
        '  string   Status;
        '  uint16   StatusInfo;
        '  string   SystemCreationClassName;
        '  string   SystemName;
        '  uint64   TotalCylinders;
        '  uint32   TotalHeads;
        '  uint64   TotalSectors;
        '  uint64   TotalTracks;
        '  uint32   TracksPerCylinder;
        '};

        'Button1.Text = InfoHarddisk.diskSerialNumber


    End Sub
    Public Function 硬盘序列号() As String
        Try
            Dim myInfo As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("HARDWARE\DEVICEMAP\Scsi\Scsi Port 0\Scsi Bus 0\Target Id 0\Logical Unit Id 0")
            硬盘序列号 = Trim(myInfo.GetValue("Identifier")) '这是型号
        Catch
            Try
                Dim myInfo As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("HARDWARE\DEVICEMAP\Scsi\Scsi Port 1\Scsi Bus 1\Target Id 0\Logical Unit Id 0")
                硬盘序列号 = Trim(myInfo.GetValue("Identifier"))
            Catch
                硬盘序列号 = ""
            End Try
        End Try
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myInfo As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey("HARDWARE\DESCRIPTION\System\MultifunctionAdapter\0\DiskController\0\DiskPeripheral\0")
        Button2.Text = Trim(myInfo.GetValue("Identifier")) '这个的后半部分是硬盘序列号
    End Sub
End Class
