# translate-pptx

PPTX ファイル（PowerPoint プレゼンテーション）を英語から日本語へ翻訳する GitHub Copilot エージェント。元のレイアウト・書式・画像・図形・色をそのまま保ち、自然な日本語に翻訳して `_JA` サフィックス付きの新しい PPTX ファイルを出力

## 前提条件

- 対応 OS: **Windows / PowerShell**、**Linux**、**macOS**、**WSL2** （いずれも検証済み）
- Python 3.10 以上
- VS Code + GitHub Copilot
- 推奨モデル: `Claude Opus 4.7`

## セットアップ

依存パッケージをインストール:

```bash
pip install -r requirements.txt
```

> 必要なパッケージは `defusedxml` のみです。

## 使い方

1. 翻訳対象 PPTX を [pptx/](pptx/) フォルダに配置
2. GitHub Copilot の Agent モード `Autopilot` で `/PPTX翻訳依頼` プロンプトを起動

```text
/PPTX翻訳依頼
```

## 動作イメージ

プロンプトが `pptx/` 配下の対象ファイルを列挙し、1 ファイルずつ順番にエージェント翻訳を実行。翻訳が完了すると逐次 `pptx/<元ファイル名>_JA.pptx` が生成され、最後に出力ファイル一覧が報告される

![動作イメージ](snapshot.png)

## 特長

- **レイアウト保持**: スライドの配置、書式、画像、図形、文字色は変更しない
- **本文 + ノート翻訳**: スライド本文 (`ppt/slides/`) に加え、発表者ノート (`ppt/notesSlides/`) も翻訳
- **常用漢字の厳守**: 文化庁告示の常用漢字表（2,136 字）のみを使用
- **中国語字形の混入防止**: 簡体字・繁体字を排除し、日本語の正字を使用
- **フォント統一**: slides / notesSlides / slideLayouts / slideMasters / notesMasters / theme の全ファイルで東アジアフォント (`<a:ea>`) を **Yu Gothic UI** に統一
- **ロケール更新**: `lang="en-*"` を `lang="ja-JP"` に自動更新
- **シャード分割翻訳辞書**: 翻訳辞書を `temp/translations/chunk_NNN.json` に分割し、LLM の出力トークン上限による途中切れを回避
- **シリアル実行**: ファイル間は並列処理せず、1 ファイルずつ確実に処理

## 主要フォルダ構成

```
translate-pptx/
├── .github/
│   ├── agents/
│   │   ├── translate-pptx.agent.md           # エージェント定義（翻訳手順・ルール）
│   │   └── scripts/
│   │       ├── 1_unpack_pptx.py              # PPTX を展開
│   │       ├── 2_extract_texts.py            # ユニーク原文を抽出（正規化済み）
│   │       ├── 3_apply_translations.py       # 翻訳適用
│   │       ├── 4_pack_pptx.py                # 展開フォルダを PPTX に再パック
│   │       └── lib/
│   │           ├── __init__.py
│   │           └── _text_normalize.py        # 共通の文字正規化処理
│   └── prompts/
│       └── PPTX翻訳依頼.prompt.md             # エージェント起動プロンプト
├── pptx/                                     # 翻訳対象 PPTX を配置（ここに翻訳結果も出力）
└── README.md
```

## 翻訳ルール

詳細は [.github/agents/translate-pptx.agent.md](.github/agents/translate-pptx.agent.md) を参照

### 漢字の選択

- 常用漢字表（2,136 字）に含まれない漢字はひらがなで表記
- 補助動詞（「ください」「いたします」など）はひらがな
- 中国語字形（簡体字・繁体字）禁止 — `机` / `认证` / `应用` ではなく日本語正字 `機` / `認証` / `応用`

### 表記の統一

| 項目 | ルール | 例 |
|------|--------|-----|
| カタカナ語の長音 | 長音符号「ー」を付ける | サーバー、ユーザー、メモリー |
| 英数字 | 半角 | API、100、v2 |
| 括弧・句読点 | 全角 | （注意）、「定義」、、。 |
| 日本語と英数字の間 | 半角スペース | Azure の API |

### 技術用語の訳語統一（抜粋）

| 英語 | 日本語 |
|------|--------|
| Authentication | 認証 |
| Authorization | 認可（「承認」ではない） |
| Deploy / Deployment | デプロイ / デプロイメント（「展開」ではない） |
| Governance | ガバナンス |
| Observability | オブザーバビリティ |
| Credential | 資格情報 |
| On-premises | オンプレミス |
| Rate limiting | レート制限 |
| Load balancing | 負荷分散 |

### 翻訳対象外

以下は翻訳せず原文のまま保持:

- 製品名・サービス名（Azure、Microsoft Foundry、MCP、A2A など）
- 人名・会社名
- URL、メールアドレス
- 技術略語（REST、GraphQL、gRPC、RBAC、SDK など）

## 制約事項

- スライドのレイアウト・書式・画像・図形・文字色は変更しない
- XML の構造（タグ、属性、名前空間）は維持
- 翻訳対象は `<a:t>` タグ内のテキストのみ（本文 + ノート）
- `<a:latin>` / `<a:cs>` は変更しません（英数字部分の見た目を保持）
- `<a:ea>` 挿入は OOXML スキーマ順序を厳守（誤ると `<a:solidFill>` が無視され文字が白色化する不具合あり）
- 元ファイルは上書きせず、`_JA` サフィックスを付けた新規ファイルとして保存
- 作業用中間ファイルはすべて `temp/` 配下で扱い、終了時に削除