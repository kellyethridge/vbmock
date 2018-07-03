Attribute VB_Name = "modStaticFunctions"
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
'   modStaticFunctions
'
Option Explicit

Public Const RES_WrongArgCount As Long = 101
Public Const RES_IncorrectParameter As Long = 102
Public Const RES_TypeMismatch As Long = 103
Public Const RES_UnexpectedCall As Long = 104
Public Const RES_MissingCall As Long = 105
Public Const RES_ObjectTypeMismatch As Long = 106
Public Const RES_ObjectShouldBeNothing As Long = 107
Public Const RES_NotSameObjectReference As Long = 108
Public Const RES_NoValueReceived As Long = 109
Public Const RES_ExpectedValue As Long = 110
Public Const RES_WrongVarType As Long = 111
Public Const RES_StartsWithError As Long = 112
Public Const RES_WrongArraySize As Long = 113
Public Const RES_WrongArrayElement As Long = 114



Private mConstraintMethods As ConstraintMethods


Public Function Test() As ConstraintMethods
    If mConstraintMethods Is Nothing Then
        Set mConstraintMethods = New ConstraintMethods
    End If
    Set Test = mConstraintMethods
End Function



Public Function GetResourceString(ByVal Index As Long, ParamArray Args() As Variant) As String
    Dim ret As String
    Dim i As Long
    
    ret = LoadResString(Index)
    For i = 0 To UBound(Args)
        ret = Replace$(ret, "{" & i & "}", Args(i))
    Next i
    GetResourceString = ret
End Function



Public Function GetLength(ByRef arr As Variant) As Long
    On Error Resume Next
    GetLength = UBound(arr) - LBound(arr) + 1
    Err.Clear
End Function

