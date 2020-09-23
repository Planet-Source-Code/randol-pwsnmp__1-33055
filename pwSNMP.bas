Attribute VB_Name = "Module1"
Global snmpPacketSize As Integer
Global snmpBinary(255) As Byte
Global Const pduGet = &HA0
Global Const pduGetNext = &HA1

Public Type lenSNMP
   packetLen As Byte
   requestLen As Byte
   structLen As Byte
End Type

Public Type verSNMP
   byteType As Byte
   packetLen As Byte
   packetData() As Byte
End Type

Public Type commSNMP
   byteType As Byte
   packetLen As Byte
   packetData() As Byte
End Type

Public Type contextSNMP
   byteType As Byte
   packetLen As Byte
End Type

Public Type requestSNMP
   byteType As Byte
   packetLen As Byte
   packetData() As Byte
End Type

Public Type errorSNMP
   byteType As Byte
   packetLen As Byte
   packetData() As Byte
End Type

Public Type indexSNMP
   byteType As Byte
   packetLen As Byte
   packetData() As Byte
End Type

Public Type struct1SNMP
   byteType As Byte
   packetLen As Byte
End Type

Public Type struct2SNMP
   byteType As Byte
   packetLen As Byte
End Type

Public Type objectSNMP
   byteType As Byte
   packetLen As Byte
   packetData() As Byte
End Type

Public Type endSNMP
   byte01 As Byte
   byte02 As Byte
End Type


Public Type snmpPacket
   UniSeq As Byte
   packetLenth As Byte   'Length of Packet minus 2
   verSNMP As verSNMP
   commSNMP As commSNMP
   contextSNMP As contextSNMP
   requestSNMP As requestSNMP
   errorSNMP As errorSNMP
   indexSNMP As indexSNMP
   struct1SNMP As struct1SNMP
   struct2SNMP As struct2SNMP
   objectSNMP As objectSNMP
   endSNMP As endSNMP
End Type

Public Type snmpData
   UniSeq01 As Byte
   UniSeq02 As Byte
   packetLenth As Byte   'Length of Packet minus 2
   verSNMP As verSNMP
   commSNMP As commSNMP
   contextSNMP As contextSNMP
   requestSNMP As requestSNMP
   errorSNMP As errorSNMP
   indexSNMP As indexSNMP
   struct1SNMP As struct1SNMP
   struct2SNMP As struct2SNMP
   objectSNMP As objectSNMP
   endSNMP As endSNMP
End Type

Sub convertBinArray(xsnmp As snmpPacket)
   Dim iINT As Integer
   Dim pos As Integer
   Dim structSize As Byte

   snmpBinary(0) = xsnmp.UniSeq
   snmpPacketSize = xsnmp.packetLenth

   snmpBinary(2) = xsnmp.verSNMP.byteType
   snmpBinary(3) = xsnmp.verSNMP.packetLen
   snmpBinary(4) = xsnmp.verSNMP.packetData(0)

   snmpBinary(5) = xsnmp.commSNMP.byteType
   snmpBinary(6) = xsnmp.commSNMP.packetLen

   For iINT = 0 To xsnmp.commSNMP.packetLen - 1
      snmpBinary(7 + iINT) = xsnmp.commSNMP.packetData(iINT)
   Next iINT
   pos = 7 + iINT

   snmpBinary(pos) = xsnmp.contextSNMP.byteType
   snmpBinary(pos + 1) = xsnmp.contextSNMP.packetLen
   pos = pos + 2

   snmpBinary(pos) = xsnmp.requestSNMP.byteType
   snmpBinary(pos + 1) = xsnmp.requestSNMP.packetLen
   snmpBinary(pos + 2) = xsnmp.requestSNMP.packetData(0)
   pos = pos + 3

   snmpBinary(pos) = xsnmp.errorSNMP.byteType
   snmpBinary(pos + 1) = xsnmp.errorSNMP.packetLen
   snmpBinary(pos + 2) = xsnmp.errorSNMP.packetData(0)
   pos = pos + 3

   snmpBinary(pos) = xsnmp.indexSNMP.byteType
   snmpBinary(pos + 1) = xsnmp.indexSNMP.packetLen
   snmpBinary(pos + 2) = xsnmp.indexSNMP.packetData(0)
   pos = pos + 3

   snmpBinary(pos) = xsnmp.struct1SNMP.byteType
   snmpBinary(pos + 1) = xsnmp.struct1SNMP.packetLen
   pos = pos + 2

   snmpBinary(pos) = xsnmp.struct2SNMP.byteType
   snmpBinary(pos + 1) = xsnmp.struct2SNMP.packetLen
   pos = pos + 2

   snmpBinary(pos) = xsnmp.objectSNMP.byteType
   snmpBinary(pos + 1) = xsnmp.objectSNMP.packetLen
   pos = pos + 2

   For iINT = 0 To xsnmp.objectSNMP.packetLen - 1
      snmpBinary(pos + iINT) = xsnmp.objectSNMP.packetData(iINT)
   Next iINT
   pos = pos + iINT

   snmpBinary(pos) = xsnmp.endSNMP.byte01
   snmpBinary(pos + 1) = xsnmp.endSNMP.byte02
   snmpBinary(1) = CByte(pos)

End Sub

Sub convertSnmp(snmpData As String)

   Dim packetLen As Integer
   Dim pos As Integer
   Dim skipLen As Integer
   Dim xbinary() As Byte

   ReDim xbinary(Len(snmpData))

   For iINT = 0 To Len(snmpData) - 1
      xbinary(iINT) = Asc(Mid(snmpData, iINT + 1, 1))
   Next iINT

   If xbinary(0) = 48 And ((xbinary(1) And &HF0) = 128) And xbinary(2) <> 2 Then
      Call Version02(xbinary)
   Else
      Call Version01(xbinary)
   End If

End Sub


Sub Version01(xbinary() As Byte)
   Dim dataLen As Integer
   Dim pos As Integer
   Dim DisplayString As String

   dataLen = xbinary(1)
   pos = 1
   dataLen = xbinary(pos + 1)
   pos = pos + 1
   pos = pos + dataLen
   pos = pos + 1
   dataLen = xbinary(pos + 1)
   pos = dataLen + pos + 3
   dataLen = pos + 1
   pos = 1 + dataLen
   dataLen = xbinary(pos)
   pos = dataLen + pos + 2
   pos = pos + 3
   dataLen = xbinary(pos)
   pos = dataLen + pos + 1

   If (xbinary(pos) = &H30) Then
      pos = pos + 1
      dataLen = 1
      pos = pos + dataLen
   End If

   If (xbinary(pos) = &H30) Then
      pos = pos + 1
      dataLen = 1
      pos = pos + dataLen
   End If

   DisplayObject pos, xbinary


End Sub

Sub Version02(xbinary() As Byte)
   Dim dataLen As Integer
   Dim pos As Integer
   Dim DisplayString As String

   dataLen = xbinary(1) And &HF
   pos = 1 + dataLen
   dataLen = xbinary(pos + 1)
   pos = pos + 1
   pos = pos + dataLen
   pos = pos + 1
   dataLen = xbinary(pos + 1)
   pos = dataLen + pos + 3
   dataLen = pos + (xbinary(pos) And &HF)
   pos = 2 + dataLen
   dataLen = xbinary(pos)
   pos = dataLen + pos + 2
   pos = pos + 3
   dataLen = xbinary(pos)
   pos = dataLen + pos + 1

   If (xbinary(pos) = &H30) Then
      pos = pos + 1
      dataLen = xbinary(pos) And &HF
      pos = pos + dataLen + 1
   End If

   If (xbinary(pos) = &H30) Then
      pos = pos + 1
      dataLen = xbinary(pos) And &HF
      pos = pos + dataLen + 1
   End If

   DisplayObject pos, xbinary

End Sub


Sub DisplayObject(pos As Integer, xbinary() As Byte)
   Dim DisplayString As String
   Dim OIDArray As Variant
   Dim OIDString As String

   DisplayString = ""
   Select Case xbinary(pos)   ' Evaluate Byte
   Case &H43   ' Display Timer Ticks
      pos = pos + 1
      For i = 1 To xbinary(pos)
         DisplayString = DisplayString & Hex(xbinary(pos + i))
      Next i
      Form1.Text4.Text = Val("&H" & DisplayString) & vbCrLf
   Case &H41   ' Display Timer Ticks
      pos = pos + 1
      For i = 1 To xbinary(pos)
         DisplayString = DisplayString & Hex(xbinary(pos + i))
      Next i
   Form1.Text4.Text = Val("&H" & DisplayString) & vbCrLf
   Case &H42   ' Display Timer Ticks
      pos = pos + 1
      For i = 1 To xbinary(pos)
         DisplayString = DisplayString & Hex(xbinary(pos + i))
      Next i
      Form1.Text4.Text = Val("&H" & DisplayString) & vbCrLf
   Case &H3   ' Display Bits
      pos = pos + 1
      For i = 1 To xbinary(pos)
         DisplayString = DisplayString & Hex(xbinary(pos + i))
      Next i
      Form1.Text4.Text = Val("&H" & DisplayString) & vbCrLf
   Case &H40   ' Display IP Addr
      pos = pos + 1
      For i = 1 To xbinary(pos)
         DisplayString = DisplayString & Hex(xbinary(pos + i))
      Next i
      Form1.Text4.Text = Val("&H" & DisplayString) & vbCrLf
   Case &H2   ' Display Timer Ticks
      pos = pos + 1
      For i = 1 To xbinary(pos)
         DisplayString = DisplayString & Hex(xbinary(pos + i))
      Next i
      Form1.Text4.Text = Val("&H" & DisplayString) & vbCrLf
   Case &H4   ' Display String
      pos = pos + 1
      For i = 0 To xbinary(pos)
         DisplayString = DisplayString & Chr(xbinary(pos + 1 + i))
      Next i
      Form1.Text4.Text = DisplayString
   Case &H5   ' Display NULL
      pos = pos + 1
      For i = 0 To xbinary(pos)
         DisplayString = DisplayString & "NULL "
      Next i
      Form1.Text4.Text = DisplayString
   Case &H6   ' Process Object
      Do While (pos <= (xbinary(1) - 2))
         DoEvents
         If xbinary(pos) = 6 Then
            pos = pos + 1
            dataLen = xbinary(pos)
            For i = 1 To dataLen
               OIDString = OIDString & Hex(xbinary(pos + i)) & "#"
            Next i
            OIDArray = Split(OIDString, "#", , vbBinaryCompare)
            Form1.Text3.Text = "1.3"
            For i = 1 To UBound(OIDArray) - 1
               Form1.Text3.Text = Form1.Text3.Text & "." & Val("&H" & OIDArray(i))
            Next i
            pos = pos + dataLen + 1
            If xbinary(pos) = 6 Then pos = pos + 2
            DisplayObject pos, xbinary
            pos = pos + xbinary(pos) + 1
         End If
         pos = 1 + pos
      Loop
   Case Else   ' Other values.
         Form1.Text4.Text = "Unknown &H" & Val("&H" & xbinary(pos))
   End Select
End Sub
