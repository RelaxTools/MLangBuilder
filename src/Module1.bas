Attribute VB_Name = "Module1"
Option Explicit

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
    
    Set t1 = MCsv.Document(MFile.Contents("C:\Test.csv"), "Delimiter="",""", "Columns=5", "Encoding=65001", "QuoteStyle=QuoteStyle.Csv")
    Set t2 = MTable.Skip(t1, 2)
    Set t3 = MTable.PromoteHeaders(t2, "PromoteAllScalars=true")

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
    
    m2.Append MCsv.Document(MFile.Contents("C:\Test.csv"), "Delimiter="",""", "Columns=5", "Encoding=65001", "QuoteStyle=QuoteStyle.Csv")
    m2.Append MTable.Skip(m2.LastTable, 2)
    m2.Append MTable.PromoteHeaders(m2.LastTable, "PromoteAllScalars=true")

    Debug.Print m2.ToString

    '結果
    'let Source1 = Csv.Document(File.Contents("C:\Test.csv"), [Delimiter=",", Columns=5, Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    '    Source2 = Table.Skip(Source1, 2),
    '    Source3 = Table.PromoteHeaders(Source2, [PromoteAllScalars=true]) in Source3

End Sub

