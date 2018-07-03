VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropertyIs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge. All Rights Reserved.
'
'    This file is part of VBMock.
'
'    VBMock is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    VBMock is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

'
'   PropertyIs
'
Option Explicit
Implements IConstraint

Private mCalls() As String
Private mPropertyCallIndex As Long
Private mConstraint As IConstraint
Private mMock As Object
Private mMessage As String
Private mArgs() As Variant



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal Mock As Object, ByVal PropertyString As String, ByRef Value As Variant, ByRef Args() As Variant)
    Set mMock = Mock
    SetPropertyNavigation PropertyString
    SetConstraint Value
    SetArguments Args
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetConstraint(ByRef Value As Variant)
    If TypeOf Value Is IConstraint Then
        Set mConstraint = Value
    Else
        Set mConstraint = Test.IsEqual(Value)
    End If
End Sub



Private Sub SetPropertyNavigation(ByVal Path As String)
    mCalls = Split(Path, ".", , vbTextCompare)
    mPropertyCallIndex = UBound(mCalls)
End Sub



' The arguments must be in reverse order for the
' late-binding invoke to work correctly.
Private Sub SetArguments(ByRef Args() As Variant)
    Dim i As Long
    Dim l As Long
    
    l = GetLength(Args)
    If l > 0 Then
        ReDim mArgs(0 To l - 1)
        For i = 0 To l - 1
            SetVariant Args(i), mArgs(l - i - 1)
        Next i
    End If
End Sub



Private Function GetPropertyValue() As Variant
    If GetLength(mArgs) = 0 Then
        SetVariant InvokeHook(GetLastObjectInChain, mCalls(mPropertyCallIndex), INVOKE_PROPERTYGET), GetPropertyValue
    Else
        SetVariant InvokeHookArray(GetLastObjectInChain, mCalls(mPropertyCallIndex), INVOKE_PROPERTYGET, mArgs), GetPropertyValue
    End If
End Function



Private Function GetLastObjectInChain() As Object
    Dim i As Long
    Dim obj As Object
    
    Set obj = mMock
    For i = 0 To mPropertyCallIndex - 1
        Set obj = InvokeHook(obj, mCalls(i), INVOKE_PROPERTYGET)
    Next i
    Set GetLastObjectInChain = obj
End Function



Private Sub SetVariant(ByRef src As Variant, ByRef dst As Variant)
    If IsObject(src) Then
        Set dst = src
    Else
        dst = src
    End If
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Eval(Value As Variant) As Boolean
    If mConstraint.Eval(GetPropertyValue) Then
        mMessage = ""
        IConstraint_Eval = True
        Exit Function
    Else
        mMessage = mConstraint.Message
    End If
End Function

Private Property Get IConstraint_Message() As String
    IConstraint_Message = mMessage
End Property
