VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IsVarType"
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
'   IsVarType
'
Option Explicit
Implements IConstraint

Private mVarType As VbVarType
Private mMessage As String



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Friend Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub Init(ByVal VarType As VbVarType)
    mVarType = VarType
End Sub



Private Function GetTypeName(ByVal VarType As VbVarType) As String
    Dim ret As String
    
    Select Case VarType
        Case vbEmpty
            ret = "Empty"
        Case vbNull
            ret = "Null"
        Case vbInteger
            ret = "Integer"
        Case vbLong
            ret = "Long"
        Case vbSingle
            ret = "Single"
        Case vbDouble
            ret = "Double"
        Case vbCurrency
            ret = "Currency"
        Case vbDate
            ret = "Date"
        Case vbString
            ret = "String"
        Case vbObject
            ret = "Object"
        Case vbError
            ret = "Error"
        Case vbBoolean
            ret = "Boolean"
        Case vbVariant
            ret = "Variant"
        Case vbDataObject
            ret = "DataObject"
        Case vbDecimal
            ret = "Decimal"
        Case vbByte
            ret = "Byte"
        Case vbUserDefinedType
            ret = "UserDefinedType"
        Case vbArray
            ret = "Array"
    End Select
    GetTypeName = ret
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IConstraint Interface
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IConstraint_Eval(Value As Variant) As Boolean
    If VarType(Value) = mVarType Then
        IConstraint_Eval = True
        mMessage = ""
    Else
        mMessage = GetResourceString(RES_WrongVarType, GetTypeName(mVarType), GetTypeName(VarType(Value)))
    End If
End Function

Private Property Get IConstraint_Message() As String
    IConstraint_Message = mMessage
End Property
