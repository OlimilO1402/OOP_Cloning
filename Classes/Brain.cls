VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Brain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Smartness As Single

Friend Sub New_()
    m_Smartness = 50
End Sub
Friend Sub NewC(other As Brain)
    m_Smartness = other.Smartness
End Sub
Public Function Clone() As Brain
    'Set Clone = Mnew.BrainC(Me)
    Set Clone = New Brain: Clone.NewC Me
End Function

Public Function ToStr() As String
    ToStr = TypeName(Me) & " {smartness: " & TypeName(m_Smartness) & "; Value: " & m_Smartness & "}"
End Function

Public Property Get Smartness() As Double
    Smartness = m_Smartness
End Property
Public Property Let Smartness(ByVal Value As Double)
    m_Smartness = Value
End Property

Public Function IsEqual(other As Brain) As Boolean
    IsEqual = m_Smartness = other.Smartness
End Function

Public Function IsSame(other As Brain) As Boolean
    'If Me Is other??? wozu soll das gut sein?
    
    If Me Is other Then IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

