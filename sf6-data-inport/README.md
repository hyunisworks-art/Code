# sf6-data-inport

Street Fighter 6 の Buckler's Boot Camp 実績ページのテキストを CSV に変換・追記するツール。

## ファイル構成

| ファイル                       | 説明                                                       |
| ------------------------------ | ---------------------------------------------------------- |
| `playlog.py`                   | メインスクリプト                                           |
| `analyze_playlog.py`           | CSVの傾向分析スクリプト                                    |
| `collect_playlog.py`           | ランキング＋プロフィール実績を一括でplaylogへ反映          |
| `scrape_rankings.py`           | 公開ランキング収集スクリプト                               |
| `scrape_profiles.py`           | short_id一覧からプロフィール実績CSVを作るスクリプト        |
| `import_ranking_to_playlog.py` | ランキングCSVを `sf6-playlog-out.csv` へ取り込むスクリプト |
| `LP_master.json`               | LP → ランク変換マスター                                    |
| `MR_master.json`               | MR → マスター帯ランク変換マスター                          |
| `sf6_playlog_in.txt`           | 最後に貼り付けた実績テキスト（自動保存）                   |
| `sf6-playlog-out.csv`          | 出力 CSV（追記形式）                                       |

## 使い方

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" playlog.py
```

実行後の操作手順：

1. Buckler's Boot Camp の実績ページ全体をコピー
2. ターミナルに貼り付けて Enter
3. リーグポイントを入力して Enter（不明なら空でEnter）
4. LP が 25000 以上の場合は MR を入力して Enter

## 一括実行の操作手順（推奨）

ランキング取得からプロフィール実績取得、`sf6-playlog-out.csv` への反映までを1回で実行します。

### 1. 実行条件

1. Python 3.13+ が利用できること
2. `sf6-data-inport` 配下に次のファイルがあること
   - `collect_playlog.py`
   - `sf6-playlog-out.csv`（既存CSV。存在しない場合はエラー終了）
   - `LP_master.json`
   - `MR_master.json`
3. Bucklerアクセス用Cookieが使えること（`403`回避用）

Cookieは次のどちらかで渡します。

1. 推奨: `.buckler_cookie.txt` に1行で保存
2. 任意: `--cookie` オプションで直接指定

### 2. 実行前チェック

1. 作業ディレクトリを `sf6-data-inport` にする
2. `.buckler_cookie.txt` の内容が最新であることを確認する（期限切れ時は再取得）
   1. bucklerサイトを開いて、ログインする
   2. F12を実行してコンソールを開く
   3. 「Disable cache」のチェックボックスをチェックする
   4. 「Doc」タブを押す
   5. Nameで「master」を選択し、「Headers」タブを選択する
   6. 「Request Headers」内の「Cookie」の「CookieConsent=」を除いた値がキャッシュになるのでコピー
3. まず `--dry-run` で対象件数を確認する

```powershellのコマンド
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" collect_playlog.py --end-page 3 --dry-run
```

実行時の内部動作:

1. ランキングページを取得して候補プレイヤーを収集
2. `sf6-playlog-out.csv` を読み込み、次の3種類に振り分け
   - 実績列が空の既存行: 更新対象
   - 未登録行: 新規追加対象
   - 実績列が埋まっている既存行: スキップ
3. 各プレイヤーのプロフィール `battle_stats` を取得して即時保存
   - 更新時: CSV全体を上書き保存（1件ごとにチェックポイント）
   - 追加時: 行を末尾に追記（1件ごとに保存）

### 3. 本実行

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" collect_playlog.py --end-page 3
```

#### 今回実行コマンドセット

```powershell
# dry-run で確認
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" collect_playlog.py `
   --ranking-type league `
   --start-page 1 `
   --end-page 466913 `
   --page-step 46691 `
   --random-start-offset `
   --dry-run

# 本実行（約10ページ × 20人 = 200人をランダム層化サンプリング）
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" collect_playlog.py `
--ranking-type league --start-page 245000 --end-page 294000 --page-step 5000 --random-start-offset --random-seed 303 --limit 50 --delay 2.0
```

### 4. 主なオプション

| オプション              | 既定値                        | 説明                                                           |
| ----------------------- | ----------------------------- | -------------------------------------------------------------- |
| `--ranking-type`        | `master`                      | 取得するランキング種別（`master` / `league`）                  |
| `--start-page`          | `1`                           | 開始ページ                                                     |
| `--end-page`            | `3`                           | 終了ページ                                                     |
| `--page-step`           | なし                          | Nページ飛ばし取得（`start, start+N...`）。`10000` 未満はエラー |
| `--random-start-offset` | なし                          | `--page-step` 時に開始ページをランダム化（偏り分散）           |
| `--random-seed`         | なし                          | `--random-start-offset` の乱数シード（再現実行用）             |
| `--locale`              | `en`                          | 取得ロケール                                                   |
| `--delay`               | `1.5`                         | リクエスト間の待機秒数                                         |
| `--timeout`             | `scrape_rankings.py` の既定値 | 通信タイムアウト秒                                             |
| `--cookie`              | 空                            | Cookie文字列を直接指定                                         |
| `--cookie-file`         | `.buckler_cookie.txt`         | Cookie文字列ファイル                                           |
| `--output`              | `sf6-playlog-out.csv`         | 書き込み先CSV                                                  |
| `--limit`               | 省略時は全件                  | 処理件数上限                                                   |
| `--dry-run`             | なし                          | 書き込みせず対象件数だけ確認                                   |

### 5. 実行後の確認

1. 末尾に `完了: 更新=... 新規追加=... エラー=... スキップ=...` が出ること
2. `エラー=0` であること
3. `sf6-playlog-out.csv` の実績列（列6以降）が更新されていること
4. 実行後に同日重複行・欠損行・LP 9000未満の行が自動で削除されること（ログに `整形:` が表示されます）

途中で止まっても、保存済み分はCSVに反映済みです。再実行すると残りだけ処理されます。

### 6. よくあるエラーと対処

1. `403` が出る
   - Cookieの期限切れが多いです。Bucklerで再ログインして `.buckler_cookie.txt` を更新してください。
2. `出力CSVが見つかりません` が出る
   - `--output` で正しいCSVを指定するか、`sf6-playlog-out.csv` の配置場所を確認してください。
3. `ページ範囲が不正です` が出る
   - `--start-page` と `--end-page` を見直してください（`start <= end` かつ `start >= 1`）。

## 出力 CSV の列構成

| インデックス | 列名                        | 内容                                         |
| ------------ | --------------------------- | -------------------------------------------- |
| 0            | No                          | 連番                                         |
| 1            | データ取得日                | 実行日（YYYY/M/D）                           |
| 2            | プレイヤー名                | テキストから自動取得                         |
| 3            | リーグポイント              | LP（手動入力）                               |
| 4            | ランク                      | LP_master / MR_master から自動解決           |
| 5            | MR                          | マスター帯のみ入力（非マスターは 0）         |
| 6–12         | ドライブゲージ各種割合      | パリィ・インパクト・オーバードライブ等       |
| 13–17        | ドライブリバーサル / パリィ | 使用回数・ジャストパリィ等                   |
| 18–23        | ドライブインパクト          | 決めた・受けた・パニカン等                   |
| 24–27        | SAゲージ使用割合            | Lv1/2/3/CA                                   |
| 28–29        | スタン                      | させた / させられた                          |
| 30–32        | 投げ                        | 決めた・受けた・投げ抜け                     |
| 33–34        | 壁際                        | 追い詰めた / 追い詰められた（秒）            |
| 35–38        | プレイ回数                  | ランクマッチ・カジュアル・ルーム・バトルハブ |
| 39           | 累計プレイポイント          | PT                                           |

## ランク解決ロジック

1. LP を `LP_master.json` に照合して基本ランク（例: GOLD3, PLAT1, MASTER）を決定
2. LP が 25000 以上（MASTER 帯）の場合のみ MR を追加入力
3. MR を `MR_master.json` に照合して詳細ランク（MASTER / HIGH / GRAND / ULTIMATE）に上書き

| MR範囲     | ランク          | abbr     |
| ---------- | --------------- | -------- |
| ～1599     | MASTER          | MASTER   |
| 1600～1699 | HIGH MASTER     | HIGH     |
| 1700～1799 | GRAND MASTER    | GRAND    |
| 1800～     | ULTIMATE MASTER | ULTIMATE |

## コマンドラインオプション

| オプション        | 既定値                | 説明                         |
| ----------------- | --------------------- | ---------------------------- |
| `--input`         | `sf6_playlog_in.txt`  | 貼り付けテキスト保存先       |
| `--output`        | `sf6-playlog-out.csv` | 出力 CSV                     |
| `--lp-master`     | `LP_master.json`      | LP マスターファイルパス      |
| `--mr-master`     | `MR_master.json`      | MR マスターファイルパス      |
| `--player`        | （自動取得）          | プレイヤー名を手動上書き     |
| `--league-points` | （対話入力）          | LP を引数で渡す              |
| `--dry-run`       | —                     | CSV に追記せず生成行だけ表示 |

## 傾向分析

現在の `sf6-playlog-out.csv` を使って、

1. LP 25000 未満は `リーグポイント`
2. LP 25000 以上は `MR`

を目的変数として探索的な傾向分析を行えます。

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" analyze_playlog.py
```

実行すると次の3種類を出力します。

1. LP傾向分析（LP 25000 未満）
2. MR傾向分析（LP 25000 以上）
3. 進捗スコア傾向分析（LP/MR を段階別に標準化した統合指標）

詳細結果は `analysis-output` フォルダに CSV で保存されます。

## 公開ランキング収集

Buckler の公開ランキングページから、ランキング一覧を JSONL / CSV で保存できます。

実行時に Buckler のWEB画面を開き続ける必要はありません。
必要になるのは「403が出たときに Cookie を渡すこと」で、これは1回取得して渡せば実行できます。

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" scrape_rankings.py --ranking-type master --start-page 1 --end-page 3
```

出力先は既定で `ranking-output` フォルダです。

生成されるファイル：

1. `master_p1-p3.csv`
2. `master_p1-p3.jsonl`
3. `master_p1-p3.meta.json`

主なオプション：

| オプション       | 既定値                | 説明                           |
| ---------------- | --------------------- | ------------------------------ |
| `--ranking-type` | `master`              | `master` または `league`       |
| `--start-page`   | `1`                   | 取得開始ページ                 |
| `--end-page`     | `3`                   | 取得終了ページ                 |
| `--locale`       | `en`                  | 取得に使うロケール             |
| `--delay`        | `1.2`                 | ページ間の待機秒数             |
| `--cookie`       | 空                    | ブラウザの Cookie 文字列       |
| `--cookie-file`  | `.buckler_cookie.txt` | Cookie文字列を保存したファイル |
| `--output-dir`   | `ranking-output`      | 出力先フォルダ                 |

CloudFront 側の制御でランキングJSONが 403 になる場合があります。その場合はブラウザで Buckler を開いた状態の Cookie を `--cookie` か `BUCKLER_COOKIE` 環境変数で渡してください。

### Cookie が必要になるケース

1. そのまま実行して成功する場合: Cookie 不要
2. 403 で失敗する場合: Cookie 必要

### 実行イメージ

1. まず Cookie なしで実行
2. 403 が出たらブラウザで Buckler を開いて Cookie を取得
3. `BUCKLER_COOKIE` に設定して再実行

※ 実行中ずっとブラウザを開いておく必要はありません。

### 毎回入力しない方法（推奨）

1. ブラウザで Buckler にログイン
2. `Cookie:` ヘッダーの値を丸ごとコピー
3. `sf6-data-inport/.buckler_cookie.txt` に保存（1行）
4. 以後は通常実行だけでOK（自動で読み込み）

### Edge での Cookie 取得手順（5ステップ）

1. Edge で Buckler のランキングページを開き、ログイン状態にする
2. `F12` で開発者ツールを開き、`Network` タブを選ぶ
3. ページを再読み込みして、`master.json?page=1` か `ranking/master?page=1` の通信をクリック
4. 右ペインの `Headers` → `Request Headers` から `cookie:` の値を丸ごとコピー
5. `sf6-data-inport/.buckler_cookie.txt` に1行で貼り付け保存して、`scrape_rankings.py` を再実行

### それでも 403 のときの確認

1. `.buckler_cookie.txt` が1行で、先頭に `Cookie:` を付けていない
2. 値の中に `;` 区切りで複数Cookieが入っている
3. ログインし直して最新Cookieを取り直している（期限切れ対策）
4. `{stamp:...necessary:true...}` のような同意Cookie単体ではない（認証Cookie全体が必要）

### ランキングCSVを playlog CSV に取り込む

`master_p1-p3.csv` の内容を `sf6-playlog-out.csv` に追記できます。
（取り込み時は `No, データ取得日, プレイヤー名, リーグポイント, ランク, MR` のみ埋まり、実績詳細列は空欄です）

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" import_ranking_to_playlog.py --ranking-csv ranking-output/master_p1-p3.csv --output sf6-playlog-out.csv
```

取り込み前に件数だけ確認したい場合：

```powershell
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" import_ranking_to_playlog.py --ranking-csv ranking-output/master_p1-p3.csv --output sf6-playlog-out.csv --dry-run
```

例：

```powershell
# ※ cookie1=value1; cookie2=value2 は説明用のダミー値です
$env:BUCKLER_COOKIE = "cookie1=value1; cookie2=value2"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" scrape_rankings.py --ranking-type master --start-page 1 --end-page 1
```

## 分析ダッシュボード（Streamlit）

サンプル数、分析根拠（相関上位）、総括文ドラフトを1画面で確認できます。

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" -m streamlit run dashboard.py
```

初回のみ、次の追加パッケージをインストールしてください。

```powershell
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/pip.exe" install streamlit plotly pandas
```

### 個別トレーニング提案（ユーザーコード診断）

ダッシュボード下部の「個別トレーニング提案」で、次を入力すると個別課題を表示します。

1. ユーザーコード（short_id）
2. 目標ランク
3. 必要なら Cookie（未入力時は `.buckler_cookie.txt` を利用）

特徴:

1. 傾向分析用CSVに本人データがなくても診断可能（ユーザーコードから最新データを直接取得）
2. 良い部分（強み）と不足部分（優先課題）を分けて表示
3. プレイ時間系は別軸で評価し、実戦量不足かトレモ優先かを提案

## 動作環境

- Python 3.13+
- 基本機能（収集・整形・傾向分析）は標準ライブラリのみ
- ダッシュボード機能は `streamlit`, `plotly`, `pandas` が必要

---

## #04記事用 ゲージ推移分析スクリプト（2026-04-18 追加）

note記事「スト６を数字で見てみた #04｜ゲージの使い方は、LPとMRでどう変わるか」用の分析スクリプト群。

### 新規スクリプト

| ファイル | 用途 |
|---|---|
| `_plot_gauge_continuous.py` | LP/MR連続軸でDゲージ・SAゲージ使用率を散布＋ガウシアン平滑化で可視化（グラフ4枚出力） |
| `_collect_mr_band.py` | MR 1500-1800 帯のサンプルをBucklerランキングから追加取得してSupabase投入 |
| `_probe_mr_pages.py` | マスターランキングの飛び飛びページからMR分布を確認（取得戦略の立案用） |
| `_count_sample_split.py` | 全サンプルを最高到達LP・最高到達MR基準で分類カウント |
| `_inspect_sa_boundary.py` | 特定ランクのSAゲージ記述統計（集計検証用） |
| `_aggregate_gauge_by_rank.py` | ランク順序×ゲージ使用率のSpearman順位相関を出力 |

### Bucklerマスターランキングの構造（重要）

- マスターランキングは **「キャラ単位のエントリ」**。1プレイヤーが複数キャラでマスター到達していたら複数回ランキングに載る
- favorite_character_league_info（現在選択中キャラ）のMR順で並ぶ
- そのため **前半ページ（1〜500あたり）にはサブキャラでMR 1500-1800のエントリが混じる**（本人はハイマス以上）→ 一次絞込を通っても最高MRが1840+で全件破棄になりがち
- **page 3000 周辺**: MR 1655近辺に同値が集中（Bucklerランキングの区切り仕様と推測）
- **page 15000 周辺**: MR 1344-1831（広めに分布）
- **page 25000 周辺**: MR 1382-1587（マスター到達直後の下限層）

### サンプリング戦略（MR 1500-1800 帯を集める手順）

表示MR（ランキング上）と最高到達MR（プレイヤー全キャラの最大MR）は乖離するため、二段階フィルタが必須。

1. `scrape_rankings.py` または `_collect_mr_band.py` でランキングページからshort_id取得
2. **一次絞込**: 表示MR で 1500-1800 フィルタ
3. Buckler profile API で play データ取得
4. **二次検証**: 全is_playedキャラから最高到達MRを計算
5. 最高MRが 1500-1800 範囲内のサンプルのみSupabaseへ投入

`_collect_mr_band.py` は上記の手順を一括で実行。MAX_REQUESTS_PER_SESSION=50 の制約があり、1セッションで約30件投入できる。

### 実行例

```powershell
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"

# 1. マスターランキングの飛び飛びページからMR分布を確認（取得戦略立案）
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _probe_mr_pages.py

# 2. MR 1500-1800 帯を追加取得（dry-run）
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _collect_mr_band.py --page-start 2995 --page-end 3005 --dry-run

# 3. 本取得（page 3000周辺 / page 15000周辺 / page 25000周辺 を組み合わせて均等分布を狙う）
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _collect_mr_band.py --page-start 2995 --page-end 3005
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _collect_mr_band.py --page-start 25000 --page-end 25010
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _collect_mr_band.py --page-start 15000 --page-end 15010

# 4. グラフ生成（analysis-output/ に4枚PNG）
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _plot_gauge_continuous.py

# 5. 分類カウント（対象・除外の内訳を確認）
& "c:/Users/hyuni/OneDrive/ドキュメント/GitHub/Code/.venv/Scripts/python.exe" _count_sample_split.py
```

### 記事用データ（2026-04-18時点）

- Supabase 全734件（初期581件 + 今回取得153件）
- 対象内訳:
  - MR帯（最高MR 1500-1800）: 132件（縦帯MR 1655-1656 を各12件上限で間引き）
  - LP帯（最高LP 3000-24999）: 114件
  - 合計対象: 246件
- 除外: ハイマス以上（最高MR>1800）281件・累積LP>24999 が27件・LP/MR取れず153件

### 分析の基準値

- LP範囲: 3000 〜 24999（ブロンズ〜ダイヤ相当）
- MR範囲: 1500 〜 1800（マスター到達 〜 ハイマス到達ライン）
- グラフX軸: LP値・MR値そのまま（連続軸）
- 平滑化: ガウシアンカーネル（bandwidth_frac=0.12・200点リサンプリング）
- Y軸上限: 60%で固定（D/SA両方）
- 同一MR値の縦帯対策: 各MR値の件数を上限12件で間引き
