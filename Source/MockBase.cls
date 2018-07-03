VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'   MockBase
'
Option Explicit
Implements IMockCallback

Private mMethods As Collection
Private mStrict As Boolean




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function SetVoidCall(ByVal Name As String, ByRef ExpectedValues() As Variant) As VoidMethod
    Set SetVoidCall = GetVoidMethod(Name)
    SetVoidCall.Newsignature ExpectedValues
End Function



Public Function SetReturnCall(ByVal Name As String, ByRef ExpectedValues() As Variant) As ReturnMethod
    Set SetReturnCall = GetReturnMethod(Name)
    SetReturnCall.Newsignature ExpectedValues
End Function



Public Sub Verify()
    Dim Method As MethodBase
    For Each Method In mMethods
        Method.Verify
    Next Method
End Sub



Public Property Get Strict() As Boolean
    Strict = mStrict
End Property
Public Property Let Strict(ByVal RHS As Boolean)
    mStrict = RHS
End Property



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetMethodBase(ByVal Name As String) As MethodBase
    Dim ret As MethodBase
    
    On Error Resume Next
    Set ret = mMethods(Name)
    If ret Is Nothing Then
        Set ret = NewMethodBase(Name)
        mMethods.Add ret, Name
    End If
    Set GetMethodBase = ret
End Function

Private Function GetVoidMethod(ByVal Name As String) As VoidMethod
    Set GetVoidMethod = New VoidMethod
    GetVoidMethod.Init GetMethodBase(Name)
End Function



Private Function GetReturnMethod(ByVal Name As String) As ReturnMethod
    Set GetReturnMethod = New ReturnMethod
    GetReturnMethod.Init GetMethodBase(Name)
End Function



Private Function NewMethodBase(ByVal Name As String) As MethodBase
    Set NewMethodBase = New MethodBase
    NewMethodBase.Init Name
End Function



Private Function Contains(ByVal Name As String) As Boolean
    On Error Resume Next
    Dim o As Object
    Set o = mMethods(Name)
    Contains = (Not o Is Nothing)
End Function



Private Sub Class_Initialize()
    Set mMethods = New Collection
End Sub



Private Sub IMockCallback_CallReturn(ByVal Name As String, retval As Variant, ParamArray Args() As Variant)
    Dim Method As ReturnMethod
    Dim vArgs() As Variant
    If (Not Contains(Name)) And (mStrict = False) Then Exit Sub
    vArgs = Args
    Set Method = GetReturnMethod(Name)
    Method.Invoke vArgs, retval
End Sub

Private Sub IMockCallback_CallVoid(ByVal Name As String, ParamArray Args() As Variant)
    Dim Method As VoidMethod
    Dim vArgs() As Variant
    If (Not Contains(Name)) And (mStrict = False) Then Exit Sub
    vArgs = Args
    Set Method = GetVoidMethod(Name)
    Method.Invoke vArgs
End Sub
