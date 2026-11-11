# M社 月次資料AI再現 スクリプト一式

LibreOffice（UNO）をheadlessで起動してから実行する：
  soffice --headless --invisible --norestore "--accept=socket,host=localhost,port=2002;urp;" &
（pkill -f soffice は使わない。シェルごと止まる）

## 流れ
1. 棚卸：python3 run_kai.py  （tanaoroshi_builder.build → AI要確認シート・黄色ハイライト）
   - kai_engine.py：在庫・売上明細・まとめ値の計算ロジック
   - build_materials.py：売上明細シート用の表（sales_tables）
2. 月次報告：python3 run_all.py 08  → n08_3.xlsx → python3 fix_kana.py n08_3.xlsx n08_4.xlsx
             python3 run_all.py 09 n08_4.xlsx → n09_3.xlsx → fix_kana.py
   - build_month.py：貼付・新科目追加・手入力値シート・参照の張替え（mode=roll/refresh）
   - manual.py：手入力値（派遣・不動産取得税・MA・人員）の月別値とMAコメント。毎月ここを更新
   - patch_pl.py：新しいPL科目の行追加（今回は減損損失）
   - apply_comments.py + comments.py：数字から書けるコメント生成 / finalize.py：コメント書込とAI作業ログ
3. 取締役会資料：python3 run_deck2.py（board_deck.build + board_bs.bs_params）→ python3 highlight.py in.pptx out.pptx
4. 照合：python3 compare.py AI版.xlsx 実物.xlsx 40
5. 案件別消費税：python3 journal_to_long.py 仕訳帳.csv long.pkl → tax_alloc.py（long2.pkl を読む。実物値JSONと締め日を引数）

## 注意
- 入力ファイルのパス（/mnt/user-data/uploads/Downloads/ など）はスクリプト内に直書き。環境に合わせて変更
- UNOのsetFormulaは関数の区切りが「;」、シート参照は $'シート名'.A1 形式
- 日本語ファイル名で保存するとLOが失敗することがある。ASCII名で保存してからコピー
- 予算（budget8/budget9）は run_deck2.py に実物資料から転記した値を直書き

## 2026/10/27 追加
- cf_loan.py：取締役会資料2枚目の「借入の内訳」「現預金の増減内訳」「在庫粗利」を計算（run_deck2.pyから呼ぶ）
  - 材料：借入金返済明細.xlsx（17期シート：PJ・運転・設備・賞与・納税の月別借入/返済）、月次報告BS増減分析（現預金・在庫・社債）、kintone買取再販台帳CSV
  - 新店投資等＝運転の返済、税金・賞与融資の返済＝設備＋賞与＋納税の返済、在庫増加ーPJ借入＝PJ純増−在庫増加、運転資金・社債＝運転等の借入＋社債、営業活動＝差額
  - 億円は各項目を四捨五入、現預金は切り捨て、営業活動は表示値の差額。8月実物と全一致、9月は実物側の「運転資金+705」「新店投資等△2.6」（端数）のみ差
  - 在庫粗利：仕入決済実績日≦月末かつ売却決済実績日が空白または月末より後の案件の「事業利益(成約予定)」合計。台帳は月末時点の出力が必要
- tax_alloc_v2.py：案件別消費税。メモに複数値（「…|案件0004866」）がある行は案件番号だけを取り出してから按分（8月25件中23件一致）
