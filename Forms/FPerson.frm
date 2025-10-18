VERSION 5.00
Begin VB.Form FPerson 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Person"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4575
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
   ScaleHeight     =   2175
   ScaleWidth      =   4575
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
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1680
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
Private m_Result As VbMsgBoxResult
Private m_Object As Person

Private Sub Form_Load()
    m_Result = VbMsgBoxResult.vbCancel
    MData.Cities_ToListCtrl CmbCity
End Sub

Public Function ShowDialog(obj As Person, Owner As Form) As VbMsgBoxResult
    Set m_Object = obj.Clone
    UpdateView
    
    'important_1:
    'from this point on, the dialog will stuck in this proecdure,
    'all Events are processed and when the Form gets closed, the
    'dialog jumps back into this procedure after show vbModal
    Me.Show vbModal, Owner
    'important_2:
    'Now in order to "clone back" all data, so the original object
    'will be updated, we use the same function "NewC" as we did for
    'cloning the object, so in every case only private write access
    'to all data is needed
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    obj.NewC m_Object
End Function

Sub UpdateView()
    If m_Object Is Nothing Then MsgBox "The Person does not exist": Exit Sub
    TxtName.Text = m_Object.Name
    TxtBirthDay.Text = m_Object.BirthDay
    If m_Object.City Is Nothing Then MsgBox "The City does not exist": Exit Sub
    CmbCity.Text = m_Object.City.Name
End Sub

Function UpdateData() As Boolean
    Dim s As String: s = TxtBirthDay.Text
    Dim bd As Date
    UpdateData = Date_TryParse(s, bd)
    If Not UpdateData Then
        Dim mr As VbMsgBoxResult: mr = MsgBox("Please give a valid Date value: " & vbCrLf & s, vbOKCancel)
        If mr = vbOK Then
            TxtBirthDay.SetFocus
            Exit Function
        End If
        UpdateView
        Exit Function
    End If
    Set m_Object = MNew.Person(bd, m_Object.Brain.Clone, MData.Cities_Add(MNew.City(CmbCity.Text)), m_Object.Index, TxtName.Text)
    UpdateData = True
End Function

Private Sub BtnOK_Click()
    If Not UpdateData Then Exit Sub
    m_Result = VbMsgBoxResult.vbOK
    Unload Me
End Sub
Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub TxtBirthDay_LostFocus()
    UpdateData
End Sub
