# MLangBuilder
Power Querry M Language Builder For VBA  
VBA で Powr Query の M言語 を組み立てるときに便利なクラス群

** →Hidennotareに吸収されました **

## クラス一覧

| クラス | 説明 |
----|---- 
| MCsv | Csv の関数の入るクラス。 |
| MFile | File の関数の入るクラス。 |
| MTable | Table の関数の入るクラス。 |
| MRecord | Record の組み立てクラス(Dictionaryラッパークラス)。 |
| MList | List の組み立てクラス(Collectionラッパークラス)。 |
| MCommand | 各クラスからM言語を作り出すクラス |

## 呼び出しサンプル
```
'------------------------------------------------
' MCommandをVBAで作成する場合のヘルパークラス
'------------------------------------------------
Sub Sample()

    '-----------------------------------
    ' MCommandを代入せずに作成する場合
    '-----------------------------------
    Dim t1 As MTable
    Dim t2 As MTable
    Dim t3 As MTable
    
    Set t1 = MCsv.Document(MFile.Contents("C:\Test.csv"), "[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]")
    Set t2 = MTable.Skip(t1, 2)
    Set t3 = MTable.PromoteHeaders(t2, "[PromoteAllScalars=true]")

    Dim m1 As MCommand
    Set m1 = New MCommand
    
    m1.Append t3
    Debug.Print m1.ToString
    
    '結果
    'let Source1 = Table.PromoteHeaders(Table.Skip(Csv.Document(File.Contents("C:\Test.csv"),
    '              [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]), 2), [PromoteAllScalars=true]) in Source1

    
    '-----------------------------------
    ' MCommandに代入して作成する場合
    '-----------------------------------
    Dim m2 As MCommand
    Set m2 = New MCommand
    
    m2.Append MCsv.Document(MFile.Contents("C:\Test.csv"), "[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]")
    m2.Append MTable.Skip(m2.Table, 2)
    m2.Append MTable.PromoteHeaders(m2.Table, "[PromoteAllScalars=true]")

    Debug.Print m2.ToString

    '結果
    'let Source1 = Csv.Document(File.Contents("C:\Test.csv"), [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    '    Source2 = Table.Skip(Source1, 2),
    '    Source3 = Table.PromoteHeaders(Source2, [PromoteAllScalars=true]) in Source3


    '-----------------------------------
    ' MRecord/MListを用いたサンプル
    '-----------------------------------
    Dim m3 As MCommand
    
    'MRecord(M言語のRecord) は DictionaryをWrapしたもの。使用方法はDictionary同等。
    Dim rec As MRecord
    Set rec = New MRecord
            
    rec.Add "Column1", """No."""
    rec.Add "Column2", """NAME"""
    rec.Add "Column3", """AGE"""
    rec.Add "Column4", """ADDRESS"""
    rec.Add "Column5", """TEL"""

    'MList(M言語のList) は CollectionをWrapしたもの。使用方法はCollectionと同等。
    Dim lst As MList
    Set lst = New MList
    lst.Add rec
    
    Set m3 = New MCommand

    m3.Append MCsv.Document(MFile.Contents("C:\Test.csv"), "[Delimiter="","", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]")
    m3.Append MTable.Skip(m3.Table, 2)
    m3.Append MTable.InsertRows(m3.Table, 0, lst)
    m3.Append MTable.PromoteHeaders(m3.Table, "[PromoteAllScalars=true]")

    Debug.Print m3.ToString

    '結果
    'let Source1 = Csv.Document(File.Contents("C:\Test.csv"), [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    '    Source2 = Table.Skip(Source1, 2),
    '    Source3 = Table.InsertRows(Source2, 0, {[Column1="No.", Column2="NAME", Column3="AGE", Column4="ADDRESS", Column5="TEL"]}),
    '    Source4 = Table.PromoteHeaders(Source3, [PromoteAllScalars=true]) in Source4

End Sub

```
