VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2004 Kelly Ethridge
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
'    Module: MockControl
'
Option Explicit
Private mBase As MockBase


' Used to tailor a series of call behaviors for a single parameter set.
'
Public Function SetVoidCall(ByVal Name As String, ParamArray ExpectedValues() As Variant) As VoidMethod
    Dim vExpectedValues() As Variant
    vExpectedValues = ExpectedValues
    Set SetVoidCall = mBase.SetVoidCall(Name, vExpectedValues)
End Function



' Used to tailor a series of call behaviors for a single parameter set.
'
Public Function SetReturnCall(ByVal Name As String, ParamArray ExpectedValues() As Variant) As ReturnMethod
    Dim vExpectedValues() As Variant
    vExpectedValues = ExpectedValues
    Set SetReturnCall = mBase.SetReturnCall(Name, vExpectedValues)
End Function



' Used for a quick single line setup of a sub call.
'
Public Sub ExpectVoid(ByVal Name As String, ParamArray ExpectedValues() As Variant)
    Dim vExpectedValues() As Variant
    vExpectedValues = ExpectedValues
    mBase.SetVoidCall(Name, vExpectedValues).AddCall 1
End Sub



' Used for a quick single line setup of a function call with a return value.
'
Public Sub ExpectReturn(ByVal Name As String, ByRef retval As Variant, ParamArray ExpectedValues() As Variant)
    Dim vExpectedValues() As Variant
    vExpectedValues = ExpectedValues
    mBase.SetReturnCall(Name, vExpectedValues).AddCall retval, 1
End Sub



Public Sub Verify()
    mBase.Verify
End Sub



Public Function GetMock(ByVal obj As IMockInitialize) As Object
    obj.SetMockBase mBase
    Set GetMock = obj
End Function



Public Property Get Strict() As Boolean
    Strict = mBase.Strict
End Property
Public Property Let Strict(ByVal RHS As Boolean)
    mBase.Strict = RHS
End Property



Private Sub Class_Initialize()
    Set mBase = New MockBase
End Sub
