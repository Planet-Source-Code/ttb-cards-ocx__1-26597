VERSION 5.00
Object = "*\ACards.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project2.Deck Deck1 
      Height          =   1545
      Left            =   480
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2725
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next Card"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Suit"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Random Card"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Deck1.CardNumber = Int(Rnd * 13)
Deck1.CardSuit = Int(Rnd * 5)
Deck1.ShowCard
End Sub

Private Sub Command2_Click()
Deck1.CardSuit = Deck1.CardSuit + 1
Deck1.ShowCard
If Deck1.CardSuit = 4 Then Deck1.CardSuit = 1
End Sub

Private Sub Command3_Click()
Deck1.CardNumber = Deck1.CardNumber + 1
Deck1.ShowCard
If Deck1.CardNumber = 13 Then Deck1.CardNumber = 1
End Sub

Private Sub Form_Load()
Deck1.CardNumber = 1
Deck1.CardSuit = Diamonds
Deck1.ShowCard
End Sub
