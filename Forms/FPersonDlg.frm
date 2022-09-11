VERSION 5.00
Begin VB.Form FPersonDlg 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Person"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label LblBirthDay 
      AutoSize        =   -1  'True
      Caption         =   "Birthday:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "FPersonDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Person As Person
Private m_DlgRet As VbMsgBoxResult

Public Function ShowDialog(aPerson As Person, aOwnerForm As Form) As VbMsgBoxResult
    Set m_Person = aPerson
    Me.Show vbModal, aOwnerForm
    ShowDialog = m_DlgRet
End Function
