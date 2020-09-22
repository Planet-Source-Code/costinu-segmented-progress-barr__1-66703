VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Segment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarMinValue As Long 'local copy
Private mvarMaxValue As Long 'local copy
Private mvarValue As Long 'local copy
Public Property Let Value(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
End Property


Public Property Get Value() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    Value = mvarValue
End Property



Public Property Let MaxValue(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MaxValue = 5
    mvarMaxValue = vData
End Property


Public Property Get MaxValue() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaxValue
    MaxValue = mvarMaxValue
End Property



Public Property Let MinValue(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MinValue = 5
    mvarMinValue = vData
End Property


Public Property Get MinValue() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinValue
    MinValue = mvarMinValue
End Property



