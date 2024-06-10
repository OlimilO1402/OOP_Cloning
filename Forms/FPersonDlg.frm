VERSION 5.00
Begin VB.Form FPerson 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Person"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtCity 
      Height          =   435
      Left            =   1440
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox TxtBirthDay 
      Height          =   435
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton BtnCancel 
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
   Begin VB.TextBox TxtName 
      Height          =   435
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label LblBirthDay 
      AutoSize        =   -1  'True
      Caption         =   "Birthday:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   750
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "FPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Person As Person
Private m_Return As VbMsgBoxResult

Public Function ShowDialog(aPerson As Person, aOwnerForm As Form) As VbMsgBoxResult
    Set m_Person = aPerson.Clone
    UpdateView
    Me.Show vbModal, aOwnerForm
    aPerson.NewC m_Person
    ShowDialog = m_Return
End Function

Sub UpdateView()
    TxtName.Text = m_Person.Name
    TxtBirthDay.Text = m_Person.BirthDay
    TxtCity.Text = m_Person.City.Name
End Sub

Sub UpdateData()
    Set m_Person = Mnew.Person(TxtBirthDay.Text, m_Person.Brain.Clone, MData.Cities_Add(TxtCity.Text), m_Person.Index, TxtName.Text)
End Sub

Private Sub BtnOK_Click()
    UpdateData
    m_Return = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Return = VbMsgBoxResult.vbCancel
    Unload Me
End Sub
