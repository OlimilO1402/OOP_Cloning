Attribute VB_Name = "MData"
Option Explicit

Public Persons As Collection
Public Cities  As Collection

Public Sub Init()
    Set Persons = New Collection
    Set Cities = New Collection
End Sub

' v ############################## v '    Cities     ' v ############################## v '
Public Function Cities_Add(City As City) As City
    Set Cities_Add = Col_AddOrGet(Cities, City)
End Function

Public Function Cities_Contains(Key As String) As Boolean
    Cities_Contains = Col_Contains(Cities, Key)
End Function

Public Sub Cities_ToListCtrl(ComboBoxOrListBox)
    Col_ToListCtrl Cities, ComboBoxOrListBox, False, True
End Sub
' ^ ############################## ^ '    Cities     ' ^ ############################## ^ '

' v ############################## v '    Persons    ' v ############################## v '
Public Function Persons_Add(Person As Person) As Person
    Set Persons_Add = MPtr.Col_AddOrGet(Persons, Person) 'Persons.Add( Person)', CStr(Person.Key)
End Function

Public Function Persons_Contains(ByVal Key As String) As Boolean
    Persons_Contains = Col_Contains(Persons, Key)
End Function

Public Sub Persons_Remove(Person As Person)
    Dim p As Person
    For Each p In Persons
        If p.IsSame(Person) Then
            If Persons_Contains(Person.Key) Then Persons.Remove Person.Key
        End If
    Next
End Sub

Public Property Get Persons_ObjectFromListCtrl(ComboBoxOrListBox, i_out As Long) As Person
    Set Persons_ObjectFromListCtrl = Col_ObjectFromListCtrl(Persons, ComboBoxOrListBox, i_out)
End Property

Public Sub Persons_ToListCtrl(ComboBoxOrListBox)
    Col_ToListCtrl Persons, ComboBoxOrListBox, False, True
End Sub

Public Property Get Persons_Item(Index) As Person
    Set Persons_Item = Persons.Item(Index)
End Property
' ^ ############################## ^ '    Persons     ' ^ ############################## ^ '

Public Function Date_TryParse(ByVal s As String, d_out As Date) As Boolean
Try: On Error GoTo Catch
    d_out = CDate(s)
    Date_TryParse = True
Catch:
End Function
