VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3840
      Top             =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "read"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3720
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OFF"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ON"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   1320
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EXIT_Click()

Form1.Hide

End Sub

Private Sub Command1_Click()
  MSComm1.Output = "a"

Text1.Text = "LED IS ON"

Shape1.FillColor = vbGreen
End Sub

Private Sub Command2_Click()
MSComm1.Output = "b"

Text1.Text = "LED IS OFF"

Shape1.FillColor = vbRed
End Sub

Private Sub Form_Load()

MSComm1.RThreshold = 3

MSComm1.Settings = "9600,n,8,1"

MSComm1.CommPort = 12 '(Write Port Number on which your Arduino Board is available).

MSComm1.PortOpen = True

MSComm1.DTREnable = False

Text1.Text = ""

Shape1.FillColor = vbRed

End Sub


Private Sub Timer1_Timer()
     ' Text1.Text = ""
 ' MSComm1.Output = "c"

  ' Text1.Text = MSComm1.Input
  ''Text2.Text = ""
  cc = MSComm1.Input
  Me.Caption = cc
   If cc <> "" Then
              Text2.Text = cc
              FL = "c:\restsys\printouts\printout.txt"
              
              Open FL For Output As #17
              Print #17, cc
              Close
              On Error GoTo xxxx
              If Not Dir("c:\restsys\xxx.txt") = "" Then
                Kill "c:\restsys\xxx.txt"
              End If
                FileCopy FL, "c:\restsys\xxx.txt"
            '    rc = Shell("c:\restsys\cprn.bat")
                rc = Shell("c:\restsys\cprn.bat", vbHide)
    
    End If
    Exit Sub
    
xxxx:
    Text2.Text = Err.Description
    Text2.Text = MSComm1.Input
    
End Sub
