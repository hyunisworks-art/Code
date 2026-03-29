# SF6成長アプリ Phase 0 - 使い方ガイド

## 準備

### Python環境
Python 3.10以上が必要です。追加ライブラリのインストールは不要（標準ライブラリのみ使用）。

### 作業フォルダ
すべてのコマンドは `sf6-data-inport/` フォルダで実行してください。

```
cd "c:\Users\hyuni\OneDrive\ドキュメント\GitHub\Code\sf6-data-inport"
```

---

## ステップ1: 同ランク帯との比較（毎日使うメイン機能）

`analyze_step1.py` を使います。`sf6-playlog-out.csv` にデータが必要です。

```bash
python analyze_step1.py --player "あなたのプレイヤー名"
```

### オプション

| オプション | 説明 | 例 |
|-----------|------|----|
| `--player` | プレイヤー名（必須） | `--player "YourName"` |
| `--input` | 分析対象CSV（省略時: sf6-playlog-out.csv） | `--input sf6-playlog-out.csv` |
| `--rank` | 比較ランク帯（省略時: あなたのランクを自動使用） | `--rank PLATINUM` |

### 出力の見方

```
======================================================================
  Step1分析: YourName vs PLATINUM帯 (45名)
  * p<0.05（有意差あり）  ** p<0.01（強い有意差）
======================================================================
  指標                           あなた   同帯平均     差分  sig  評価
  -----------------------------------------------------------------------
  ドライブパリィ%                  8.2%     14.5%    -6.3%   *   [課題]
  DI被弾回数                      18.3     12.1      +6.2        [課題]
  端追い詰め時間                   31.2     38.7      -7.5        [課題]
======================================================================
```

- `[課題]` : 同ランク帯より劣っている指標 → 練習優先
- `[良]  ` : 同ランク帯より優れている指標
- `[--]  ` : ほぼ同じ
- `*` / `**` : 統計的に有意な差あり（サンプル数30以上で信頼性が上がります）

---

## ステップ2: サンプルデータ収集（週1〜2回推奨）

`collect_samples.py` を使います。同ランク帯のサンプルを収集します。

### 基本的な実行（全ランク帯・各50件）
```bash
python collect_samples.py
```

### 特定ランク帯だけ収集
```bash
python collect_samples.py --rank platinum
python collect_samples.py --rank platinum --count 30
```

### 動作確認（dry-run）
```bash
python collect_samples.py --rank platinum --count 5 --dry-run
```

### Cookieが必要な場合
ブラウザからCookieを取得して `.buckler_cookie.txt` に貼り付けてください。

```bash
python collect_samples.py --rank platinum --cookie-file .buckler_cookie.txt
```

### 注意事項
- 収集データは `data/samples/YYYY-MM-DD_rank_[ランク].csv` に保存されます
- 30日以上古いファイルは自動削除されます
- リクエスト間隔: 3秒（Bucklerへの負荷軽減のため変更しないでください）
- 1セッションの上限: 50リクエスト

---

## ステップ3: 個人データ取得（週1回推奨）

`fetch_my_data.py` を使います。ひゅーさん自身のデータを取得します。

### short_id の確認方法
BucklerのプロフィールURL を確認してください:
```
https://www.streetfighter.com/6/buckler/profile/XXXXXXXX
                                                 ↑これがshort_id
```

### 実行方法
```bash
python fetch_my_data.py --short-id XXXXXXXX
```

### 動作確認（dry-run）
```bash
python fetch_my_data.py --short-id XXXXXXXX --dry-run
```

### 保存先
`data/my/YYYY-MM-DD_my_data.csv` に日付付きで保存されます。
個人データは削除されません（履歴として保持）。

---

## 毎日の使い方（推奨フロー）

```bash
# 1. 個人データを取得（週1回）
python fetch_my_data.py --short-id XXXXXXXX

# 2. 同ランク帯サンプルを収集（週1〜2回）
python collect_samples.py --rank platinum --count 30

# 3. Step1分析で比較（毎日）
python analyze_step1.py --player "YourName"
```

---

## よくあるエラー

### `エラー: ファイルが見つかりません: sf6-playlog-out.csv`
Step1分析には `sf6-playlog-out.csv` が必要です。
`collect_playlog.py` でデータを取得してから実行してください。

### `警告: PLATINUM帯のデータが少ないです（5件）`
サンプル数が30件未満の場合は精度が低くなります。
`collect_samples.py` でサンプルを追加収集してください。

### `403 エラー`
Bucklerの認証が必要な場合は `.buckler_cookie.txt` にCookieを貼り付けてください。

---

## ファイル構成

```
sf6-data-inport/
├── analyze_step1.py      # Step1分析（同ランク帯との比較）
├── analyze_playlog.py    # Step2分析（ランク間相関）
├── collect_samples.py    # サンプルデータ収集
├── fetch_my_data.py      # 個人データ取得
├── collect_playlog.py    # プレイログ収集
├── sf6-playlog-out.csv   # 分析データ（Step1/Step2の入力）
├── data/
│   ├── samples/          # ランク帯別サンプル（自動管理・30日で削除）
│   └── my/               # 個人データ履歴（削除しない）
└── HOWTO.md              # このファイル
```
