VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------
'
' [MFunctionCreater] v1
'
' Copyright (c) 2019 Yasuhiro Watanabe
' https://github.com/RelaxTools/MFunctionCreater
' author:relaxtools@opensquare.net
'
' The MIT License (MIT)
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'
'-----------------------------------------------------------------------------------------------------
' Power Query M functions Table Class
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor

Public Name As String
Private mValue As String
Private mParent As MTable
'---------------------------------------
' �R���X�g���N�^
'---------------------------------------
Private Function IConstructor_Instancing(Args As Collection) As Object
    Select Case Args.Count
        Case 1
            mValue = Args(1)
            Set mParent = Nothing
        Case 2
            mValue = Args(1)
            Set mParent = Args(2)
        Case Else
            Exit Function
    End Select
    
    Set IConstructor_Instancing = Me

End Function
'---------------------------------------
' PromoteHeaders
'---------------------------------------
Public Function PromoteHeaders(t As MTable, ParamArray param() As Variant) As MTable

    Dim strBuf As String
    Dim sb As StringBuilder
    Dim v As Variant
    
    Set sb = New StringBuilder

    For Each v In param
        sb.Append v
    Next

    If t.Name = "" Then
        strBuf = t.ToString
    Else
        strBuf = t.Name
    End If
    
    Set PromoteHeaders = Constructor(New MTable, GetClassName & "PromoteHeaders(" & strBuf & ", [" & sb.ToJoin(", ") & "])", t)

End Function
'---------------------------------------
' Skip
'---------------------------------------
Public Function Skip(t As MTable, countOrCondition As Long) As MTable
    
    Dim strBuf As String
    
    If t.Name = "" Then
        strBuf = t.ToString
    Else
        strBuf = t.Name
    End If
    
    Set Skip = Constructor(New MTable, GetClassName & "Skip(" & strBuf & ", " & countOrCondition & ")", t)
    
End Function
'---------------------------------------
' �l���ꐶ��
'---------------------------------------
Public Function ToString() As String
    ToString = mValue
End Function
'---------------------------------------
' �N���X���擾
'---------------------------------------
Private Function GetClassName() As String
    GetClassName = Mid$(TypeName(Me), 2) & "."
End Function
