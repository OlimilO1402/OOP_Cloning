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
   Begin VB.ComboBox CmbCity 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox TxtBirthDay 
      Height          =   435
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox TxtName 
      Height          =   435
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   600
      Width           =   750
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
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

Private Sub Form_Load()
    MData.Cities_ToListCtrl CmbCity
End Sub

Public Function ShowDialog(Person As Person, aOwnerForm As Form) As VbMsgBoxResult
    Set m_Person = Person.Clone
    UpdateView
    
    'important_1:
    'from this point on, the dialog will stuck in this proecdure,
    'all Events are processed and when the Form gets closed, the
    'dialog jumps back into this procedure after show vbModal
    Me.Show vbModal, aOwnerForm
    'important_2:
    'Now in order to "clone back" all data, so the original object
    'will be updated, we use the same function "NewC" as we did for
    'cloning the object, so in every case only private write access
    'to all data is needed
    Person.NewC m_Person
    ShowDialog = m_Return
End Function

Sub UpdateView()
    If m_Person Is Nothing Then MsgBox "The Person does not exist": Exit Sub
    TxtName.Text = m_Person.Name
    TxtBirthDay.Text = m_Person.BirthDay
    If m_Person.City Is Nothing Then MsgBox "The City does not exist": Exit Sub
    CmbCity.Text = m_Person.City.Name
End Sub

Function UpdateData() As Boolean
    Dim bd As Date
    UpdateData = Date_TryParse(TxtBirthDay.Text, bd)
    If Not UpdateData Then Exit Function
    Set m_Person = MNew.Person(bd, m_Person.Brain.Clone, MData.Cities_Add(MNew.City(CmbCity.Text)), m_Person.Index, TxtName.Text)
    UpdateData = True
End Function

Private Sub BtnOK_Click()
    If Not UpdateData Then Exit Sub
    m_Return = VbMsgBoxResult.vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Return = VbMsgBoxResult.vbCancel
    Unload Me
End Sub
