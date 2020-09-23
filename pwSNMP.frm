VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6600
      TabIndex        =   9
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get Next SNMP"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   9375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "192.168.136.161"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get SNMP"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Text            =   "1.3.6.1.2.1."
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "public"
      Top             =   360
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock pwSocket 
      Left            =   9600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   161
   End
   Begin VB.Label Label4 
      Caption         =   "Start Walking at MIB Value"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "OID"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Community Name"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP Address"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DataArrivalCount As Integer
Public DataReturned As String

Public Sub snmpSend(destination As String, Community As String, Oid As String, getType As Byte)
   Dim iINT As Integer
   Dim convertType(255) As Byte
   Dim convertString As String
   Dim xsnmp As snmpPacket
   Dim OIDArray As Variant
   Dim OIDsize As Integer

   'Text4.Text = ""

   xsnmp.UniSeq = &H30
   xsnmp.verSNMP.byteType = 2
   xsnmp.verSNMP.packetLen = 1
   ReDim xsnmp.verSNMP.packetData(0)
   xsnmp.verSNMP.packetData(0) = 0
   
   xsnmp.commSNMP.byteType = 4
   ReDim xsnmp.commSNMP.packetData(CByte(Len(Trim$(Community)) - 1))
   For iINT = 0 To UBound(xsnmp.commSNMP.packetData)
      xsnmp.commSNMP.packetData(iINT) = Asc(Mid(Trim$(Community), iINT + 1, 1))
   Next iINT
   xsnmp.commSNMP.packetLen = CByte(UBound(xsnmp.commSNMP.packetData) + 1)
      
   xsnmp.contextSNMP.byteType = getType
   
   xsnmp.requestSNMP.byteType = 2
   ReDim xsnmp.requestSNMP.packetData(0)
   xsnmp.requestSNMP.packetData(0) = 1
   xsnmp.requestSNMP.packetLen = 1
   
   xsnmp.errorSNMP.byteType = 2
   ReDim xsnmp.errorSNMP.packetData(0)
   xsnmp.errorSNMP.packetData(0) = 0
   xsnmp.errorSNMP.packetLen = 1
   
   xsnmp.indexSNMP.byteType = 2
   ReDim xsnmp.indexSNMP.packetData(0)
   xsnmp.indexSNMP.packetData(0) = 0
   xsnmp.indexSNMP.packetLen = 1
   
   xsnmp.struct1SNMP.byteType = &H30
   
   xsnmp.struct2SNMP.byteType = &H30
   
   xsnmp.objectSNMP.byteType = 6
   
   OIDArray = Split(Trim(Oid), ".", , vbBinaryCompare)
   OIDsize = UBound(OIDArray)
   If OIDArray(UBound(OIDArray)) = "" Then OIDsize = OIDsize - 1
   
   
   ReDim xsnmp.objectSNMP.packetData(OIDsize - 1)
   xsnmp.objectSNMP.packetData(0) = &H2B
   
   j = 2
   
   For i = 2 To OIDsize
      If OIDArray(i) > 255 Then
         remainder = OIDArray(i) Mod 128
         quoient = OIDArray(i) / 128
         OIDsize = OIDsize + 1
         ReDim Preserve xsnmp.objectSNMP.packetData(OIDsize - 1)
         xsnmp.objectSNMP.packetData(j - 1) = 128 + quoient
         xsnmp.objectSNMP.packetData(j) = remainder
         j = j + 1
      Else
         xsnmp.objectSNMP.packetData(j - 1) = OIDArray(i)
      End If
      j = j + 1
   Next i
   
   xsnmp.endSNMP.byte01 = 5
   xsnmp.endSNMP.byte02 = 0
   xsnmp.objectSNMP.packetLen = OIDsize
   xsnmp.struct2SNMP.packetLen = xsnmp.objectSNMP.packetLen + 4
   xsnmp.struct1SNMP.packetLen = xsnmp.struct2SNMP.packetLen + 2
   xsnmp.contextSNMP.packetLen = xsnmp.struct1SNMP.packetLen + 11
   
   Call convertBinArray(xsnmp)
   Debug.Print snmpBinary
   pwSocket.RemotePort = 161
   pwSocket.RemoteHost = Trim(destination)
   pwSocket.SendData snmpBinary
End Sub

Private Sub Combo1_Click()
   Dim OIDStart As String
   Dim OIDSplit As Variant
   OIDStart = Combo1.Text
   OIDSplit = Split(OIDStart, "-", , vbBinaryCompare)
   Text3.Text = Trim(OIDSplit(1))
End Sub

Private Sub Command1_Click()
   snmpSend Text1.Text, Text2.Text, Text3.Text, pduGet
End Sub

Private Sub Command2_Click()
   snmpSend Text1.Text, Text2.Text, Text3.Text, pduGetNext
End Sub

Private Sub Form_Load()
   DataArrivalCount = 0
   'Open "packetData.bin" For Binary Access Write As #1
   
   Combo1.AddItem "MIB2 - 1.3.6.1.2.1."
   Combo1.AddItem "Enterprise - 1.3.6.1.4.1"
   Combo1.AddItem "SNMP Modules - 1.3.6.1.6.3"
   
End Sub

Private Sub pwSocket_Close()
   MsgBox "Closed"
End Sub

Private Sub pwSocket_Connect()
   MsgBox "Connected"
End Sub

Private Sub pwSocket_ConnectionRequest(ByVal requestID As Long)
MsgBox "Connection Request"
End Sub

Private Sub pwSocket_DataArrival(ByVal bytesTotal As Long)
  Dim OIDData As String
  pwSocket.GetData OIDData
  'Put #1, , OIDData

  'TestCvt (OIDData)

  convertSnmp (OIDData)
  DataArrivalCount = DataArrivalCount + 1
End Sub

Private Sub TestCvt(x)
 For i = 1 To Len(x)
   Debug.Print Asc(Mid(x, i, 1)) & vbTab & Mid(x, i, 1)
 Next
End Sub

Private Sub pwSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   MsgBox "Socket Error", vbInformation, Description & " " & Source
End Sub

