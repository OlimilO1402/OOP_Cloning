Attribute VB_Name = "MData"
Option Explicit

Public Persons As Collection
Public Cities  As Collection

Public Sub Init()
    Set Persons = New Collection
    Set Cities = New Collection
End Sub

' v ############################## v '    Cities     ' v ############################## v '
Public Function Cities_Contains(key As String) As Boolean
    On Error Resume Next
    If IsEmpty(Cities(key)) Then: 'DoNothing
    Cities_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Cities_Add(ByVal Name As String) As City
    If Cities_Contains(Name) Then
        Set Cities_Add = Cities.Item(Name)
    Else
        Set Cities_Add = New City: Cities_Add.New_ Name
        Cities.Add Cities_Add, Name
    End If
End Function
' ^ ############################## ^ '    Cities     ' ^ ############################## ^ '

' v ############################## v '    Persons    ' v ############################## v '
Public Sub Persons_Add(Person As Person)
    Persons.Add Person, Person.key
End Sub

Public Function Persons_Contains(ByVal key As String) As Boolean
    On Error Resume Next
    If IsEmpty(Persons(key)) Then: 'DoNothing
    Persons_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Sub Persons_Remove(Person As Person)
    Dim p As Person
    For Each p In Persons
        If p.IsSame(Person) Then
            If Persons_Contains(Person.key) Then Persons.Remove Person.key
        End If
    Next
End Sub

Public Sub Persons_ToListBox(aLB As ListBox)
    Dim i As Long, p As Person
    With aLB
        .Clear
        For i = 1 To Persons.Count
            Set p = Persons.Item(i)
            .AddItem p.ToStr
            .ItemData(i - 1) = p.key
        Next
    End With
End Sub

Public Property Get Persons_Item(Index) As Person
    'Debug.Print m_Persons.Count
    Set Persons_Item = Persons.Item(Index)
End Property
' ^ ############################## ^ '    Persons     ' ^ ############################## ^ '
