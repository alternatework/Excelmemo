# Excelmemo
#=IF(WEEKDAY(B4)=1,IF(COUNTIF(B:B,B4+1)=1,B4+2,IF(COUNTIF(B:B,B4+2)=1,B4+3,IF(COUNTIF(B:B,B4+3)=1,B4+4,IF(COUNTIF(B:B,B4+4)=1,B4+5,IF(COUNTIF(B:B,B4+5)=1,B4+6,B4+1))))),B4)

=IF(SUM(IF((A21>=Sheet3!A1:A100)*(A21<=Sheet3!B1:B100)*(Sheet3!C1:C100="3-Mtr")+(Sheet3!D1:D100="3-Mtr"), 1, 0))>0, "×", "")

=IF(SUM(IF(((A21>=Sheet3!A1:A100)*(A21<=Sheet3!B1:B100)*((ISNUMBER(SEARCH(3-Mtr", Sheet3!C1:C100)))+(ISNUMBER(SEARCH("3-Mtr", Sheet3!D1:D100)))), 1, 0))>0, "×", "")

=QUERY({Sheet1!A2:Z; Sheet2!A2:Z; Sheet3!A2:Z; Sheet4!A2:Z}, "SELECT * WHERE Col1 IS NOT NULL", 0)

Excel 2016では、QUERY関数が直接使用できないため、代わりにPower Query（またはGet & Transformデータ）を使用してデータをまとめることができます。以下は手順です：

Power Query Editorを開く:

データの範囲を含む各シートを選択します。
データタブからクエリの編集をクリックします。
クエリエディターでの結合:

各シートのデータがPower Query Editorに表示されます。
各テーブルを選択して、ホームタブからマージ クエリをクリックします。
適切な結合キー（共通の列）を選択して、マージを行います。
フィルター:

必要に応じて、不要な列や行をフィルタリングします。
データを結合:

クエリエディターの左上にある閉じると適用ボタンをクリックして変更を保存し、元のExcelファイルにデータを結合します。
Power QueryはExcel 2016以降で利用可能で、上記の手順は基本的なものです。手順は実際のデータの形式によって変わる可能性があるため、具体的なデータの構造に合わせて調整が必要かもしれません。
