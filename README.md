# OOP_Cloning
## Correct creating and cloning of objects with a constructor and a clone function  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/OOP_Cloning?style=plastic)](https://github.com/OlimilO1402/OOP_Cloning/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/OOP_Cloning?style=plastic)](https://github.com/OlimilO1402/OOP_Cloning/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/OOP_Cloning/total.svg)](https://github.com/OlimilO1402/OOP_Cloning/releases/download/v2024.06.15/CtorCloning_v2024.06.15.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/OOP_Cloning/watchers)  

and asking  
 * if one object is  equal   to another object (IsEqual) and  
 * if one object is the same as another object (IsSame)  
  
Project started around may 2005.  
  
In most OOP-languages nowadays you have an object-constructor (aka ctor) for creating objects and for first initialization of necessary properties of the object in just one line of code.  
And also a copy-constructor for cloning resp copying the complete state of all properties of one object into another new object also in just one line of code.  
In Visual Basic Classic (VBC & VBA) we do not have the convenience of a language inbuilt constructor, but this does not mean we have to disclaim about it. Of course we can write our own object initialization constructor and object cloning functions for every class.  
  
1. Simple example of using an **Object Constructor**:  
```vb6
Private Sub BntOpenFile_Click()
    Dim file As PathFileName
    Set file = MNew.PathFileName("C:\MyPath\MyFile.ext", FileAccess.BinaryWrite, FileMode.OpenOrCreate)  
    If file.Open Then
        '
    End If
End Sub
```  
  
1.1. To achieve this in VB do the following:  
1.1a) In every Class we need one procedure with a name that makes clear it is meant for creating a New Object e.g. "New_" (or maybe "Init" or whatever)
Making it "Friend", has the benefit it has not to be implemented in derived classes and every derived class can again have it's own ctor-function:  
```vb6  
Friend Sub New_(PathFileName As String, FileAccess As FileAccess, FileMode As FileMode)  
    'inside the class we have access to all private variables  
    '. . . init and assign all private variables  
End Sub  
```  
1.1b) We have a standard-module as a factory for convenient creating objects. The name of this module should mirror the purpose of object creation, should be short and should be the same in every project like e.g. "MNew".
      with Public functions with the name of every class and every function parameter is the same as in the corresponding Friend Sub New_:  
```vb6  
Public Function PathFileName(PathFileName As String, FileAccess As FileAccess, FileMode As FileMode)  
    Set PathFileName = New PathFileName: PathFileName.New_ PathFileName, FileAccess, FileMode  
End Function  
```  
But pay attention to the following, if things begin to change inf your project over time, you have to synchronize all function parameters between 
"Friend Sub New_(<all function parameters>)" and "Public Function MyClass(<all function parameters>) As MyClass, but the benefit of more clear and concise code will soon pay off the little overhead
  
2. Example of using **Cloning** Of Objects  
```vb6  
Private Sub BtnPerson_Click()  
    Dim simon As Person: Set simon = garfunkel.Clone  
    Dim dolly As Sheep:  Set dolly = shaun.Clone  
    '. . .  
End Sub  
```  
  
2.1 To achieve this do the following:  
2.1a) In every cloneable class we need a Function Clone that creates and returns a new object of the type as the class itself:  
```vb6  
Friend Function Clone() As Person
    Set Clone = MNew.Person(Me.Name, Me.EyeColor, Me.HairColor)
End Function
```  
  
2.1.b) And because we need direct private write access to class members in the new object, we can also do this via a procedure Friend Sub NewC where we just hand over the old object to the new object  
```vb6  
Friend Sub NewC(other As Person) As Person
    With other
        m_Name = .Name
        m_BirthDate = .BirthDate
        m_EyeColor = .EyeColor
        m_HairColor = .HairColor
    End With
End Function
```  
Then in the Public Function Clone you just use it like this:  
```vb6  
Public Function Clone() As Person  
    Set Clone = New Person: Clone.NewC Me
End Function  
```  
Maybe in the first place, you somewhat have to wrap your brain around it before you dig how it works. So I would advise: just write it down yourself in a new empty project, and step through your code, so you will get used to it quickly.  
  
3. Modal dialogs  
  
In a modal dialog the user is allowed to edit all object properties and either saving all edits with the OK-button or maybe throw away all edits with the Cancel-button just in case the user is not sure anymore whether the edits are correct or not.  
  
3.1 To achive this, again we could use the Cloning functions. The modal dialog needs a Function ShowDialog where we give the object to edit and the parentwindow and return which button was pressed (OK or Cancel).  
```vb6  
Public Function ShowDialog(Obj As Person, Owner As Form) As vbMsgBoxResult  
```  
In the dialog we hold a clone of the object as a private member. The clone is created from the object right at the beginning of the Function ShowDialog
```vb6  
Option Explicit
Private m_Person As Person
Public Function ShowDialog(Obj As Person, Owner As Form) As vbMsgBoxResult  
    Set m_Person = Obj.Clone
```  
then we update the view and we actually show the modal-dialog with Show vbModal
```vb6  
Public Function ShowDialog(Obj As Person, Owner As Form) As vbMsgBoxResult  
    Set m_Person = Obj.Clone
    UpdateView
    Me.Show vbModal, Owner
```  
The proecdure stops in this line, but no problem all events will be done. At the moment the dialog will be closed, the next line in this procedure gets executed. Now is the time to write all edits to the original object.
We just have to use the same procedure that was used when the Clone was created namely "NewC"
```vb6  
Public Function ShowDialog(Obj As Person, Owner As Form) As vbMsgBoxResult  
    Set m_Person = Obj.Clone
    UpdateView
    Me.Show vbModal, Owner
    Obj.NewC m_Person
End Function
```  

3.2 the Function ShowDialog will be used like this  
```vb6  
Private Sub  BtnEdit_Click() 
    Dim i As Long, Obj As Person: Set Obj = Col_ObjectFromListCtrl(MData.Persons, List1, i)
    If Obj Is Nothing Then Exit Sub
    If FPerson.ShowDialog(Obj, Me) = vbCancel Then Exit Sub
    UpdateView1 i, Obj
End Function
```  
That's all there is to it folks, pretty easy stuff, of course you could use this knowledge in all other languages as well.  

![OOP_Cloning Image](Resources/PCloningIsEqualOrSame.png "OOP-Cloning Image")