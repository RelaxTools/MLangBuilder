VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Power Query M Language List Class(Collection Wrapper)
'-----------------------------------------------------------------------------------------------------
Option Explicit
Implements IConstructor
Implements ICursor

Private mCol As Collection
Private v() As Variant
Private mIndex As Long
'--------------------------------------------------------------
'  Initialize
'--------------------------------------------------------------
Private Sub Class_Initialize()
    Set mCol = New Collection
    mIndex = 0
End Sub
'--------------------------------------------------------------
'  Terminate
'--------------------------------------------------------------
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'--------------------------------------------------------------
'  コンストラクタ
'--------------------------------------------------------------
Private Function IConstructor_Instancing(ByRef Args As Collection) As Object

    Dim v As Variant
    Dim i As Long
    
    Select Case Args.Count
        
        '引数なし
        Case 0
        
        '引数１個
        Case 1
            Select Case True
                
                'コレクションからの変換
                Case TypeOf Args(1) Is Collection
                    Me.Clear
                    Set mCol = Args(1)
                
                '１次元配列からの変換
                Case IsArray(Args(1))
                    Me.Clear
                    For i = LBound(Args(1)) To UBound(Args(1))
                        mCol.Add Args(1)(i)
                    Next
                Case Else
                    'エラー
                    Exit Function
            End Select
        
        '引数からの追加
        Case Else
            Me.Clear
            For Each v In Args
                mCol.Add v
            Next
    End Select
    
    Set IConstructor_Instancing = Me
    
End Function
'--------------------------------------------------------------
'  MList から全ての項目を削除する。
'--------------------------------------------------------------
Public Sub Clear()

    Set mCol = New Collection
    mIndex = 0
    
End Sub
'--------------------------------------------------------------
'  MList に項目を追加する。
'--------------------------------------------------------------
Public Sub Add(obj As Variant)
    
    mCol.Add obj

End Sub
'--------------------------------------------------------------
'  MList の指定位置に項目を追加する。
'--------------------------------------------------------------
Public Sub Insert(Idx As Long, v As Variant)

    If Idx < 0 Or Idx >= mCol.Count Then
        Err.Raise vbObjectError + 512 + 1, , "IndexOutOfRangeException"
    End If
    
    mCol.Add v, , Idx + 1

End Sub
'--------------------------------------------------------------
'  For Each 用
'--------------------------------------------------------------
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCol.[_NewEnum]
End Function
'--------------------------------------------------------------
'  MList 要素数を返却
'--------------------------------------------------------------
Public Property Get Count() As Long
    Count = mCol.Count
End Property
'--------------------------------------------------------------
'  MList の指定位置の項目を取得
'--------------------------------------------------------------
Public Property Get Item(Idx As Long) As Variant
Attribute Item.VB_UserMemId = 0
    
    If Idx < 0 Or Idx >= mCol.Count Then
        Err.Raise vbObjectError + 512 + 1, , "IndexOutOfRangeException"
    End If
        
    If VBA.IsObject(mCol.Item(Idx + 1)) Then
        Set Item = mCol.Item(Idx + 1)
    Else
        Item = mCol.Item(Idx + 1)
    End If
    
End Property
'--------------------------------------------------------------
'  MList の指定位置の項目を削除
'--------------------------------------------------------------
Public Sub RemoveAt(Idx As Long)
    
    If Idx < 0 Or Idx >= mCol.Count Then
        Err.Raise vbObjectError + 512 + 1, , "IndexOutOfRangeException"
    End If
    
    mCol.Remove Idx + 1
End Sub
'--------------------------------------------------------------
'  MList のソート
'--------------------------------------------------------------
Public Sub Sort()

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim col2 As Collection
    Dim var() As Variant
    Dim wk As String

    'Collectionが空ならなにもしない
    If mCol Is Nothing Then
        Exit Sub
    End If

    'Collectionの要素数が０または１の場合ソート不要
    If mCol.Count <= 1 Then
        Exit Sub
    End If

    n = mCol.Count
    ReDim var(1 To n)

    For i = 1 To n
        var(i) = mCol.Item(i)
    Next

    '挿入ソート
    For i = 2 To n

        wk = var(i)

        If var(i - 1) > wk Then

            j = i

            Do

                var(j) = var(j - 1)

                j = j - 1

                If j = 1 Then
                    Exit Do
                End If

            Loop While var(j - 1) > wk
            var(j) = wk

        End If
    Next

    Set col2 = New Collection

    For i = 1 To n
        col2.Add var(i)
    Next

    Set mCol = col2
    Set col2 = Nothing

End Sub
'--------------------------------------------------------------
'  MList を１次元配列に変換
'--------------------------------------------------------------
Public Function ToArray() As Variant

    Dim c As Variant
    Dim i As Long
    
    ReDim v(0 To mCol.Count)
    
    For i = 0 To mCol.Count - 1

        If IsObject(mCol(i + 1)) Then
            Set v(i) = mCol(i + 1)
        Else
            Let v(i) = mCol(i + 1)
        End If
        
    Next

    ToArray = v
    
End Function
'--------------------------------------------------------------
'  MList をCollectionに変換
'--------------------------------------------------------------
Public Function ToCollection() As Collection
    
    Set ToCollection = mCol

End Function
'--------------------------------------------------------------
'  MList を文字列に変換
'--------------------------------------------------------------
Public Function ToString() As String

    Dim v As Variant
    Dim sb As StringBuilder
    Set sb = New StringBuilder
    
    For Each v In mCol
    
        Select Case True
            Case IsObject(v)

                sb.Append v.ToString

'            Case IsEmpty(v)
'
'                sb.Append "null"
'
'            Case VarType(v) = vbBoolean
'
'                sb.Append LCase(CStr(v))
'
'            Case VarType(v) = vbString
'
'                sb.Append """" & v & """"

            Case Else

                sb.Append v

        End Select
    
    Next

    ToString = "{" & sb.ToJoin(", ") & "}"

End Function
'--------------------------------------------------------------
'  ICursor インターフェースを取得
'--------------------------------------------------------------
Public Property Get GetCursor() As ICursor
    Set GetCursor = Me
End Property
'--------------------------------------------------------------
'  ICursor の Eof プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Eof() As Boolean
    ICursor_Eof = mIndex >= mCol.Count
End Property
'--------------------------------------------------------------
'  ICursor の Item プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_Item(Optional ByVal opt As Variant) As Variant
    
    If mIndex >= mCol.Count Then
        ICursor_Item = mCol(mCol.Count)
    Else
        ICursor_Item = mCol(mIndex + 1)
    End If
    
End Property
'--------------------------------------------------------------
'  ICursor の PreviousItem プロパティを取得
'--------------------------------------------------------------
Private Property Get ICursor_PreviousItem(Optional ByVal opt As Variant) As Variant
    
    If mIndex <= 0 Then
        Err.Raise vbObjectError + 512 + 1, , "IndexOutOfRangeException"
    Else
        ICursor_PreviousItem = mCol(mIndex)
    End If

End Property
'--------------------------------------------------------------
'  ICursor の カーソル位置を初期化
'--------------------------------------------------------------
Private Sub ICursor_MoveFirst()
    mIndex = 0
End Sub
'--------------------------------------------------------------
'  ICursor の カーソル位置を次に移動
'--------------------------------------------------------------
Private Sub ICursor_MoveNext()

    mIndex = mIndex + 1
    If mIndex > mCol.Count Then
        mIndex = mCol.Count
    End If

End Sub

