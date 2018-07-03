VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsLikeArray"
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
'   IsLikeArray
'
Option Explicit
Implements IConstraint

Private mConstraints As ConstraintList
Private mMessage As String



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByRef Arr As Variant)
    Set mConstraints = New ConstraintList
    mConstraints.Init Arr
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Eval(Value As Variant) As Boolean
    Dim l As Long
    Dim i As Long
    Dim lb As Long
    
    l = GetLength(Value)
    If l = mConstraints.Count Then
        lb = LBound(Value)
        For i = 0 To l - 1
            If mConstraints(i).Eval(Value(lb)) = False Then
                mMessage = GetResourceString(RES_WrongArrayElement, lb)
                Exit Function
            End If
            lb = lb + 1
        Next i
        IConstraint_Eval = True
        mMessage = ""
    Else
        mMessage = GetResourceString(RES_WrongArraySize, GetLength(mConstraints), l)
    End If
End Function

Private Property Get IConstraint_Message() As String
    IConstraint_Message = mMessage
End Property
