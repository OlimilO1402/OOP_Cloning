VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BrainSmart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Brain
Private m_Smartness As Double

Private Sub Class_Initialize()
    m_Smartness = 100
End Sub

Friend Sub NewC(other As BrainSmart)
    m_Smartness = other.Smartness
End Sub

Public Function Clone() As BrainSmart
    Set Clone = New BrainSmart: Clone.NewC Me
End Function

Public Function ToStr() As String
    ToStr = TypeName(Me) & " {smartness: " & TypeName(m_Smartness) & "; Value: " & m_Smartness & "}"
End Function

Public Property Get Smartness() As Double
    Smartness = m_Smartness
End Property

Public Function IsEqual(other As BrainSmart) As Boolean
    IsEqual = m_Smartness = other.Smartness
End Function

Public Function IsSame(other As BrainSmart) As Boolean
    If Me Is other Then IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

' #################### ' Implements Brain ' #################### '
Private Function Brain_ToStr() As String
    Brain_ToStr = Me.ToStr
End Function

Private Property Get Brain_Smartness() As Double
    Brain_Smartness = Me.Smartness
End Property

Public Property Let Brain_Smartness(ByVal Value As Double)
    m_Smartness = Value
End Property

Private Function Brain_Clone() As Brain
    Set Brain_Clone = Me.Clone
End Function

Private Function Brain_IsEqual(other As Brain) As Boolean
    Brain_IsEqual = Me.IsEqual(other)
End Function

Private Function Brain_IsSame(other As Brain) As Boolean
    Brain_IsSame = Me.IsSame(other)
End Function

