VERSION 5.00
Begin VB.Form frmellipses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Ellipses Program"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   975
   End
   Begin VB.HScrollBar hsbRadius 
      Height          =   255
      Left            =   840
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   480
      Value           =   1
      Width           =   5535
   End
   Begin VB.HScrollBar hsbaspect 
      Height          =   255
      Left            =   840
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   120
      Value           =   1
      Width           =   5535
   End
   Begin VB.Label lblAspect 
      Caption         =   "Aspect:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblRadius 
      Caption         =   "Radius:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblInfo 
      Caption         =   "Aspect :"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmellipses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Created by Deon van Zyl
End
End Sub

Private Sub cmdExit_Click()
End

End Sub

Private Sub Form_Load()

    ' Initialize the radius and aspect scroll bars.
    hsbRadius.Value = 10
    hsbaspect.Value = 10
    
    'Intialize the info Label
    lblInfo.Caption = "Aspect: 1"
    
    'Set the Drawwidth property of the form
    frmellipses.DrawWidth = 1
    
    
End Sub

Private Sub hsbaspect_Change()

    Dim x, y
    Dim Info
    'Calculatethe center of the form
    x = frmellipses.ScaleWidth / 2
    y = frmellipses.ScaleHeight / 2
    
    'Clear the form
    frmellipses.Cls
    
    'Draw the ellipse.
    frmellipses.Circle (x, y), hsbRadius.Value * 10, _
        RGB(255, 0, 0), , , hsbaspect.Value / 10
        
    'Prepare the Info String
    Info = "Aspect : " + Str(hsbaspect.Value / 10)
    
    'Display the value of the aspect
    frmellipses.lblInfo.Caption = Info
    
        
End Sub

Private Sub hsbaspect_Scroll()

    hsbaspect_Change
    
End Sub

Private Sub hsbRadius_Change()

Dim x, y
    Dim Info
    'Calculatethe center of the form
    x = frmellipses.ScaleWidth / 2
    y = frmellipses.ScaleHeight / 2
    
    'Clear the form
    frmellipses.Cls
    
    'Draw the ellipse.
    frmellipses.Circle (x, y), hsbRadius.Value * 10, _
        RGB(255, 0, 0), , , hsbaspect.Value / 10
        
    'Prepare the Info String
    Info = "Aspect : " + Str(hsbaspect.Value / 10)
    
    'Display the value of the aspect
    frmellipses.lblInfo.Caption = Info



End Sub

Private Sub hsbRadius_Scroll()
    
    hsbRadius_Change
    
End Sub
