VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "City"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name As String

Friend Sub New_(Name As String)
    m_Name = Name
End Sub

Friend Function Clone() As City
    Set Clone = Me 'there is only one of it's kind
End Function

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Ptr() As LongPtr
    Ptr = ObjPtr(Me)
End Property

Public Property Get Key() As String
    Key = m_Name
End Property

Public Function IsEqual(other As City) As Boolean
    IsEqual = m_Name = other.Name
End Function

Public Function IsSame(other As City) As Boolean
    If Me Is other Then IsSame = Me.Ptr = other.Ptr
End Function

Public Function ToStr() As String
    ToStr = m_Name
End Function

