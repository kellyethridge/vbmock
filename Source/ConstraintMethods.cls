VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstraintMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
'   ConstraintMethods
'
Option Explicit



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Public Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' These will appear as constructors when access using the Test.* method.
'
Public Function IsEqual(ByRef ExpectedValue As Variant, Optional ByVal Delta As Double) As IsEqual
    Set IsEqual = New IsEqual
    IsEqual.Init ExpectedValue, Delta
End Function

Public Function IsAnything() As IsAnything
    Set IsAnything = New IsAnything
End Function

Public Function IsNothing() As IsNothing
    Set IsNothing = New IsNothing
End Function

Public Function IsTypeOf(ByVal TypeName As String) As IsTypeOf
    Set IsTypeOf = New IsTypeOf
    IsTypeOf.Init TypeName
End Function

Public Function IsIn(ParamArray Values() As Variant) As IsIn
    Set IsIn = New IsIn
    IsIn.Init ToArray(Values)
End Function

Public Function IsNot(ByRef Value As Variant) As IsNot
    Set IsNot = New IsNot
    IsNot.Init Value
End Function

Public Function IsVarType(ByVal VarType As VbVarType) As IsVarType
    Set IsVarType = New IsVarType
    IsVarType.Init VarType
End Function

Public Function bAnd(ByRef a As Variant, ByRef b As Variant) As bAnd
    Set bAnd = New bAnd
    bAnd.Init a, b
End Function

Public Function bOr(ByRef a As Variant, ByRef b As Variant) As bOr
    Set bOr = New bOr
    bOr.Init a, b
End Function

Public Function NotEqual(ByRef Value As Variant) As NotEqual
    Set NotEqual = New NotEqual
    NotEqual.Init Value
End Function

Public Function NotNothing() As NotNothing
    Set NotNothing = New NotNothing
End Function

Public Function NotIn(ParamArray Values() As Variant) As NotIn
    Set NotIn = New NotIn
    NotIn.Init ToArray(Values)
End Function

Public Function IsEqualIgnoreCase(ByRef Value As Variant) As IsEqualIgnoreCase
    Set IsEqualIgnoreCase = New IsEqualIgnoreCase
    IsEqualIgnoreCase.Init Value
End Function

Public Function StartsWith(ByVal Value As String) As StartsWith
    Set StartsWith = New StartsWith
    StartsWith.Init Value
End Function

Public Function IsLikeArray(ByRef Arr As Variant) As IsLikeArray
    Set IsLikeArray = New IsLikeArray
    IsLikeArray.Init Arr
End Function

Public Function AsParamArray(ParamArray Params() As Variant) As AsParamArray
    Set AsParamArray = New AsParamArray
    AsParamArray.Init ToArray(Params)
End Function

Public Function PropertyIs(ByVal Mock As Object, ByVal PropertyName As String, ByRef Value As Variant, ParamArray Args() As Variant) As PropertyIs
    Set PropertyIs = New PropertyIs
    PropertyIs.Init Mock, PropertyName, Value, ToArray(Args)
End Function

Public Function IsLike(ByRef Expression As Variant) As IsLike
    Set IsLike = New IsLike
    IsLike.Init Expression
End Function

Public Function IsInContainer(ByVal Container As IUnknown) As IsInContainer
    Set IsInContainer = New IsInContainer
    IsInContainer.Init Container
End Function

Public Function IsInArray(ByRef Arr As Variant) As IsInArray
    Set IsInArray = New IsInArray
    IsInArray.Init Arr
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToArray(ParamArray Arr() As Variant) As Variant()
    ToArray = Arr(0)
End Function

