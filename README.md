# mdash-full

MDASH（数理・データサイエンス・AI教育プログラム認定制度）リテラシーレベルの申請書を整理したリポジトリです。

## フォルダ構成

```text
.
├─ literacy_application/        # 認定申請書
│  ├─ r3/                       # 令和3年度認定
│  ├─ r4/                       # 令和4年度認定
│  ├─ r5/                       # 令和5年度認定
│  ├─ r6/                       # 令和6年度認定
│  ├─ r7/                       # 令和7年度認定
│  └─ r7_pre/                   # 令和7年度先行認定（R7先行）
└─ literacy_change_notification/ # 変更届（提供後に追加予定）
```

現在、変更届は未収録です。変更届は7月中を目途に提供される予定で、提供され次第追加します。今後件数が増えることを想定し、申請書とは分けて管理します。

申請書のファイル名は、`機関名.xlsx` の形式です。

## 認定年度と教育プログラムの実績

原則として、各年度の認定は前年度までに実施された教育プログラムの実績を対象とします。たとえば、令和3年度認定では、令和2年度までに実施された教育プログラムが対象です。

ただし「R7先行」は例外です。令和7年度に開設した教育プログラムを、試行的に同じ令和7年度中に先行認定したものです。

## 修了要件・構成科目の JSON 抽出

年度ごとの様式差に対応した抽出スクリプトを `scripts/` に収録しています。

```powershell
python -m pip install -r requirements.txt

# 全年度を抽出し、結合版・不完全校一覧も生成
python scripts/extract_all_years.py --progress

# 年度別 JSON から結合版・集計だけを再生成
python scripts/extract_all_years.py --combine-only

# 年度単位で実行する場合
python scripts/extract_r3.py
python scripts/extract_r4.py
python scripts/extract_r5.py
python scripts/extract_r6.py
python scripts/extract_r7.py
python scripts/extract_r7_pre.py

# 生成した JSON の件数・出典セル・重複を検証
python scripts/validate_extraction.py

# 全年度を横断検索できる単一 HTML を生成
python scripts/build_curriculum_html.py
```

出力先は `outputs/curriculum/` です。`r3.json` から `r7_pre.json` までの年度別ファイルに加え、`all_years.json`、`summary.json`、`incomplete_universities.json`、`validation_report.json` を生成します。HTML 閲覧用の `curriculum_browser.html` はデータを内包し、全大学・全プログラムを展開した状態で単体表示できます。

各プログラムには修了要件、最低科目数・単位数、対象学部・学科、構成科目を収録します。構成科目は年度様式内で重複して記載される場合があるため科目名で統合し、単位数、必須区分、モデルカリキュラム対応番号、元シート・セルを保持します。
