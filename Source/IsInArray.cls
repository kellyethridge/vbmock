VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsInArray"
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
'   IsInArray
'
Option Explicit
Implements IConstraint

Private mConstraints As ConstraintList
Private mMessage As String



Friend Sub Init(ByRef Arr As Variant)
    Set mConstraints = New ConstraintList
    mConstraints.Init Arr
End Sub



Private Function IConstraint_Eval(Value As Variant) As Boolean
    Dim i As Long
    For i = 0 To mConstraints.Count - 1
        If mConstraints(i).Eval(Value) Then
            mMessage = ""
            IConstraint_Eval = True
            Exit Function
        End If
    Next i
    mMessage = "Value was not found in the array."
End Function

Private Property Get IConstraint_Message() As String
    IConstraint_Message = mMessage
End Property
