VERSION 5.00
Begin VB.Form Temps 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TNumberOfDays 
      Height          =   390
      Left            =   3525
      TabIndex        =   0
      Top             =   1095
      Width           =   1965
   End
   Begin VB.Label LFixedRate 
      Caption         =   "Label5"
      Height          =   480
      Left            =   3630
      TabIndex        =   5
      Top             =   2685
      Width           =   1830
   End
   Begin VB.Label LDoubleRate 
      Caption         =   "LDoubleRate"
      Height          =   510
      Left            =   3615
      TabIndex        =   4
      Top             =   2055
      Width           =   1770
   End
   Begin VB.Label Label3 
      Caption         =   "Fixed Rate"
      Height          =   420
      Left            =   1515
      TabIndex        =   3
      Top             =   2820
      Width           =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Double Rate"
      Height          =   495
      Left            =   1515
      TabIndex        =   2
      Top             =   2010
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Number Of Days"
      Height          =   435
      Left            =   1500
      TabIndex        =   1
      Top             =   1155
      Width           =   1620
   End
End
Attribute VB_Name = "Temps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    LDoubleRate.Caption = ""
    LFixedRate.Caption = ""
End Sub

Private Sub calculate()
Dim dRate As Double
Dim fRate As Double
Dim noOfDays As Long
Dim i As Long
    
    noOfDays = Val(TNumberOfDays.Text)
    i = 0
        
    dRate = 1
    If noOfDays <= 0 Then
        dRate = 0
    End If
    While i < noOfDays
        dRate = dRate + dRate * i
        i = i + 1
    Wend
    
    i = 1
    While i <= noOfDays
        fRate = fRate + 100
        i = i + 1
    Wend

    LDoubleRate.Caption = "$" & Format(dRate, "0.00")
    LFixedRate.Caption = "$" & Format(fRate, "0.00")
End Sub

Private Sub TNumberOfDays_Change()
    calculate
End Sub
