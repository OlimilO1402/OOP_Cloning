VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModalDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_Object
Private m_Dialog
Private WithEvents mDialog As Form
Attribute mDialog.VB_VarHelpID = -1
Private WithEvents mBtnOK  As CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCcl As CommandButton
Attribute mBtnCcl.VB_VarHelpID = -1

Private Sub Class_Initialize()
    m_Result = VbMsgBoxResult.vbCancel
End Sub

Friend Sub New_(aDialog As Form, BtnOK As CommandButton, BtnCancel As CommandButton)
    Set m_Dialog = aDialog: Set mDialog = aDialog: Set mBtnOK = BtnOK: Set mBtnCcl = BtnCancel
End Sub

Public Function ShowDialog(Obj, Owner As Form) As VbMsgBoxResult
    Set m_Object = Obj.Clone
    m_Dialog.UpdateView m_Object
    mDialog.Show vbModal, Owner
    ShowDialog = m_Result
    If ShowDialog = vbCancel Then Exit Function
    Obj.NewV m_Object
End Function

Private Sub mBtnOK_Click()
    If Not m_Dialog.UpdateData(m_Object) Then Exit Sub 'm_Dialog.UpdateView
    m_Result = VbMsgBoxResult.vbOK
    Unload mDialog
End Sub

Private Sub mBtnCancel_Click()
    Unload mDialog
End Sub
