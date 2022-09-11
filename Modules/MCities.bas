Attribute VB_Name = "MCities"
Option Explicit

Private m_Cities As New Collection

Public Property Get List() As Collection
    Set List = m_Cities
End Property

Public Function Contains(key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_Cities(key)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Sub ToListBox(aLB)
    aLB.Clear
    Dim c As City
    For i = 1 To m_Cities.Count
        Set c = m_Cities.Item(i)
        aLB.AddItem c.Name
    Next
End Sub

