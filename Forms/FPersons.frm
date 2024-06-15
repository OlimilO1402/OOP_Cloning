VERSION 5.00
Begin VB.Form FPersons 
   Caption         =   "Persons"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Delete [ - ]"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnEdit 
      Caption         =   "Edit [ / ]"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton BtnAdd 
      Caption         =   "Add [ + ]"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "FPersons.frx":0000
      Left            =   0
      List            =   "FPersons.frx":0002
      TabIndex        =   0
      ToolTipText     =   "DoubleClick to Edit"
      Top             =   360
      Width           =   13815
   End
End
Attribute VB_Name = "FPersons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    UpdateView
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single: T = List1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
End Sub

Private Sub BtnAdd_Click()
    Dim p As New Person
    If FPerson.ShowDialog(p, Me) = vbCancel Then Exit Sub
    MData.Persons_Add p
    UpdateView
End Sub

Private Sub BtnEdit_Click()
    List1_DblClick
End Sub

Private Sub BtnDelete_Click()
    Dim i As Long, Obj As Person: Set Obj = MData.Persons_ObjectFromListCtrl(List1, i)
    If Obj Is Nothing Then Exit Sub
    If MsgBox("Do you really want to delete this person?" & vbCrLf & Obj.ToStr, vbOKCancel) = vbCancel Then Exit Sub
    MData.Persons_Remove Obj
    UpdateView
End Sub

Private Sub List1_DblClick()
    Dim i As Long, Obj As Person: Set Obj = MData.Persons_ObjectFromListCtrl(List1, i)
    If Obj Is Nothing Then Exit Sub
    If FPerson.ShowDialog(Obj, Me) = vbCancel Then Exit Sub
    UpdateView1 i, Obj
End Sub

Private Sub UpdateView()
    MData.Persons_ToListCtrl List1
End Sub

Private Sub UpdateView1(ByVal Index As Long, ByVal Obj As Person)
    List1.List(Index) = Obj.ToStr
End Sub

