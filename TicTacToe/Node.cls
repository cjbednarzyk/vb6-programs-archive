VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarParent As Node 'local copy
Private mvarChildren As Collection 'local copy
Public Property Set Children(ByVal vData As Collection)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Children = Form1
    Set mvarChildren = vData
End Property


Public Property Get Children() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Children
    Set Children = mvarChildren
End Property



Public Property Set Parent(ByVal vData As Node)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Parent = Form1
    Set mvarParent = vData
End Property


Public Property Get Parent() As Node
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Set Parent = mvarParent
End Property



