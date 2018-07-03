VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockCall"
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
'   MockCall
'
Option Explicit
Implements IMockCall

Private mBase As CallBase
Private mRetVal As Variant
Private mSignature As CallSignature



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub GetRetVal(ByRef retval As Variant)
    SetVariant mRetVal, retval
End Sub



Friend Sub Init(ByVal Min As Long, ByVal Max As Long, ByRef retval As Variant, ByRef Signature As CallSignature)
    Set mBase = New CallBase
    mBase.Init Min, Max
    SetVariant retval, mRetVal
    Set mSignature = Signature
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetVariant(ByRef src As Variant, ByRef dst As Variant)
    If IsObject(src) Then
        Set dst = src
    Else
        dst = src
    End If
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IMockCall Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get IMockCall_Count() As Long
    IMockCall_Count = mBase.Count
End Property

Private Function IMockCall_Invoke(Args() As Variant) As Boolean
    If mBase.Invoke(Args) Then
        mSignature.Invoke Args
        IMockCall_Invoke = True
    End If
End Function

Private Property Get IMockCall_Min() As Long
    IMockCall_Min = mBase.Min
End Property
