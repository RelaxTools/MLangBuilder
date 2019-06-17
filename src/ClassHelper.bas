Attribute VB_Name = "ClassHelper"
'-----------------------------------------------------------------------------------------------------
'
' [MLangBuilder] v1
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
' コンストラクタ生成
'-----------------------------------------------------------------------------------------------------
Option Explicit

Public Function Constructor(ByRef obj As Object, ParamArray Args() As Variant) As Object

    Dim c As IConstructor
    Dim v As Variant
    
    '引数をCollectionに詰め替える
    Dim col As Collection
    Set col = New Collection
    
    For Each v In Args
        col.Add v
    Next
    
    'IConstructor Interfaceを呼び出す。
    Set c = obj
    Set Constructor = c.Instancing(col)
    
    'オブジェクトが返却されなかった場合エラー
    If Constructor Is Nothing Then
        Err.Raise vbObjectError + 512 + 1, "Argument Error"
    End If

End Function



