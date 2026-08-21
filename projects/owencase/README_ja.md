<p align="center">
  <img src="assets/ppt-mcp-logo-letter.png" alt="IBM Bob PowerPoint MCP" width="480">
</p>

<p align="center">
  <a href="README.md">English</a>
</p>

<p align="center">
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python 3.10+"></a>
  <img src="https://img.shields.io/badge/Platform-Windows-0078d4.svg" alt="Windows">
  <img src="https://img.shields.io/badge/License-MIT-green.svg" alt="MIT License">
  <img src="https://img.shields.io/badge/IBM%20Bob-ready-052FAD.svg" alt="IBM Bob ready">
</p>

<h1 align="center">IBM Bob PowerPoint MCP</h1>

<p align="center">
  <strong>Windows COM 自動化による、IBM Bob 向けの安全なリアルタイム PowerPoint 編集。</strong>
</p>

IBM Bob から、起動中の Microsoft PowerPoint を確認・編集するための Model Context Protocol（MCP）サーバーです。Windows COM を介してデスクトップ版 PowerPoint を直接操作するため、変更は即座に画面へ反映され、PowerPoint ネイティブのオブジェクト、レイアウト、テーマ、アニメーション、メディアを維持できます。

`python-pptx` のようなファイルベースのライブラリとは異なり、PowerPoint で開いているプレゼンテーションを直接操作します。対象ファイルのロック、図形状態の事前条件、差分、検証、再試行可能性を含むエラーによって、エージェントによる編集をより安全にします。

## 主な特徴

- **IBM Bob 向け設計** — 確認、編集、差分、検証の順序をツール指示とガードで徹底します。
- **リアルタイム COM 自動化** — ファイルモデルを再構築せず、起動中の PowerPoint を直接制御します。
- **27 カテゴリ・165 ツール** — プレゼンテーション、スライド、図形、テキスト、表、グラフ、テーマ、アニメーション、メディア、エクスポートなどを扱えます。
- **フェイルクローズな対象管理** — 対象が閉じられた場合や見つからない場合に、別のプレゼンテーションへ勝手に切り替わりません。
- **事前条件付き編集** — 移動、削除、ビジュアル置換では、変更前に確認した図形状態を検証します。
- **再試行を考慮したエラー** — MCP エラーに `retryable` と修正用の `hint` を含め、危険な無条件再試行を防ぎます。
- **出力先の制限** — 名前を付けて保存、およびエクスポートは信頼済みディレクトリ内に限定されます。
- **視覚的な確認** — スライドプレビュー、図形スナップショットの差分、保存前検証を利用できます。

## 動作環境

- Windows 10 または Windows 11
- デスクトップ版 Microsoft PowerPoint
- Python 3.10 以降
- [`uv`](https://docs.astral.sh/uv/getting-started/installation/)

PowerPoint COM 自動化は Windows 専用です。Python テストの多くは PowerPoint を起動せず実行できますが、実際のプレゼンテーション編集には Windows 上の PowerPoint が必要です。

## このリポジトリからインストール

`projects/owencase` で実行します。

```powershell
uv sync --frozen
uv run ppt-mcp-bob
```

サーバーは MCP の stdio 通信を使用します。通常は直接起動するのではなく、IBM Bob が MCP 設定を通じて起動します。

## IBM Bob への登録

Bob がこのローカルソースを実行するように設定します。次のパスを Windows マシン上の絶対パスへ置き換えてください。

```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "C:\\Users\\YOUR_NAME\\.local\\bin\\uv.exe",
      "args": [
        "--directory",
        "C:\\ABSOLUTE\\PATH\\TO\\ppt-mcp-bob\\projects\\owencase",
        "run",
        "ppt-mcp-bob"
      ],
      "env": {
        "PPT_TEMPLATES_DIR": "C:\\ABSOLUTE\\PATH\\TO\\templates",
        "PPT_MCP_OUTPUT_DIR": "C:\\ABSOLUTE\\PATH\\TO\\output"
      }
    }
  }
}
```

開発・配置時は `--directory` でこのローカルディレクトリを指定してください。`uvx ppt-mcp` は別途公開されているアップストリームパッケージを解決するため、このディレクトリのコードは実行されません。

従来の設定との互換性のため `ppt-mcp` コマンドも残していますが、このプロジェクトの正式なコマンドは `ppt-mcp-bob` です。

## Bob の安全な編集手順

既存のプレゼンテーションを編集する場合、Bob は次の順序で操作します。

1. `ppt_activate_presentation` を呼び出し、`ppt_get_presentation_info`、`ppt_list_shapes`、スライドプレビューで対象を確認します。
2. 確認したフルパスとスライド数を指定して `ppt_set_work_mode` を呼び出します。
3. 移動、サイズ変更、編集、削除、置換では `allow_create=false` を維持します。
4. 図形スナップショットを取得し、安定した `shape_id` を保持します。
5. 確認済み状態を事前条件として、`ppt_transform_shapes`、`ppt_delete_shapes`、または `ppt_replace_visual` を使います。
6. `ppt_proofread_text` を呼び出し、返されたすべてのテキストを文脈に沿って確認し、綴りや分かち書きの問題を修正して再実行します。
7. 保存前に `ppt_diff_shape_snapshot` と `ppt_validate_presentation` を実行します。
8. 校正結果が空で、差分と検証結果が依頼内容に一致した場合のみ保存します。

`ppt_proofread_text` は読み取り専用です。信頼度の高い韓国語・英語の誤記、
カスタム置換、単語の重複、句読点、括弧、制御文字、文字化けを検査します。
また、図形、グループ、表、SmartArt、グラフから位置情報付きテキストを返し、
Bob が文脈に沿って確認できるようにします。製品名には `allowed_terms`、組織固有
の用語には `custom_replacements` を使用できます。発表者ノートは任意です。

MCP サーバーは Bob に静かな完了報告ポリシーも渡します。プレゼンテーションの
作業中は、計画、ツール呼び出しとその結果、進捗、プレビューなどの中間情報を
ユーザーへ表示しません。保存と検証が完了した後、結果、出力先と作業範囲、検証
結果を、ユーザーの言語で正確に 3 行だけ報告します。ただし、ツール呼び出しカード
の表示は MCP ではなくクライアント UI が制御するため、クライアントによっては表示
される場合があります。

既定のワークモードは意図的に厳格です。

| 設定 | 既定値 | 効果 |
|---|---:|---|
| `allow_create` | `false` | 作成が明示的に許可されるまで、追加、コピー、複製ツールをブロックします。 |
| `require_preconditions` | `true` | 変更時に対象プレゼンテーションのフルパスと図形状態の事前条件を要求します。 |

## 環境変数

| 変数 | 必須 | 説明 |
|---|---:|---|
| `PPT_TEMPLATES_DIR` | いいえ | ディレクトリ引数がない場合に `ppt_list_templates` が検索する場所です。 |
| `PPT_MCP_OUTPUT_DIR` | 推奨 | 名前を付けて保存、およびエクスポート先の信頼境界です。未指定の場合、利用可能であればロック済みプレゼンテーションのディレクトリを使用します。 |
| `PPT_DOWNLOAD_TIMEOUT_SECONDS` | いいえ | リモート画像とアイコンメタデータのタイムアウトです。既定値は `15` 秒です。 |
| `PPT_MAX_DOWNLOAD_BYTES` | いいえ | リモートダウンロードの最大サイズです。既定値は `20971520` バイト（20 MiB）です。 |
| `PPT_AUTO_DISMISS_DIALOG` | いいえ | PowerPoint がビジー状態で COM 呼び出しを拒否した場合に Escape を送信します。既定では無効です。 |

`PPT_AUTO_DISMISS_DIALOG=true` は無人実行に便利ですが、ユーザーが開いているダイアログをキャンセルする可能性があります。対話操作では、意図して自動キャンセルを使う場合を除き無効のままにしてください。

## ツールカテゴリ

| カテゴリ | ツール数 | 対応範囲 |
|---|---:|---|
| アプリ | 5 | 接続、アプリ状態、アクティブウィンドウ、開いているプレゼンテーション |
| プレゼンテーション | 8 | 作成、開く、保存、閉じる、有効化、情報取得、テンプレート一覧 |
| スライド | 10 | 追加、削除、複製、移動、コピー、情報取得、ノート、移動 |
| 図形 | 10 | 安定した ID による図形、テキストボックス、画像、線の追加と確認 |
| 安全な編集 | 8 | ワークモード、安全な変形・削除・置換、スナップショット、差分、検証 |
| テキスト | 11 | 内容、範囲、段落、箇条書き、検索、抽出、タイポグラフィ、校正 |
| プレースホルダー | 6 | プレースホルダー内容の確認と更新 |
| 書式 | 3 | 塗りつぶし、線、影 |
| 表 | 13 | データ、セル、行、列、結合・分割、スタイル、配置、罫線 |
| エクスポート | 4 | PDF、画像、スライドプレビュー、クリップボード |
| スライドショー | 6 | 開始、終了、移動、状態確認 |
| グラフ | 7 | 作成、確認、データ設定、書式、種類変更 |
| アニメーション | 6 | 画面切り替えとアニメーションのライフサイクル |
| テーマ | 4 | テーマ適用、テーマカラー、ヘッダー・フッター |
| グループ | 3 | グループ化、解除、項目確認 |
| コネクタ | 2 | コネクタの追加と書式 |
| ハイパーリンク | 3 | 追加、確認、削除 |
| セクション | 3 | 追加、一覧、管理 |
| プロパティ | 2 | プレゼンテーションメタデータの取得と更新 |
| メディア | 3 | 動画、音声、メディア設定 |
| SmartArt | 3 | 追加、変更、レイアウト一覧 |
| 編集操作 | 6 | 元に戻す、やり直し、図形・書式のコピー |
| レイアウト | 7 | 整列、均等配置、サイズ、背景、反転、結合 |
| 効果 | 3 | 光彩、反射、ぼかし |
| コメント | 3 | 追加、一覧、削除 |
| 高度な操作 | 19 | タグ、フォント、トリミング、画像、選択、アイコン、URL、バッチ適用 |
| フリーフォーム | 7 | パス作成とノードの確認・編集 |
| **合計** | **165** | |

## 開発

`projects/owencase` で実行します。

```powershell
uv sync --group dev
uv run pytest
```

テストでは、スキーマの厳格性、対象ロック、再試行動作、パス制約、安定した図形 ID、スライド操作、検証ロジックを確認します。

## ライセンスとクレジット

MIT License で公開しています。

この IBM Bob 統合は [owencase](https://github.com/owencase) が保守し、[ykuwai/ppt-mcp](https://github.com/ykuwai/ppt-mcp) の PowerPoint MCP を基盤としています。[FastMCP](https://github.com/jlowin/fastmcp)、[pywin32](https://github.com/mhammond/pywin32)、[Model Context Protocol](https://modelcontextprotocol.io/) を使用しています。
