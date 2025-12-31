

# Repository Guidelines for Humans and Agents (AGENTS.md)

本ファイルは、このリポジトリで作業する **人間および自動化エージェント（Codex等）** が、
環境・手順・運用ルールを誤解なく共有するための **唯一の正本** です。
エージェントは本ファイルの内容を最優先で遵守してください。

---

## 0. Purpose（目的）

Receipt Parser GAS は、レシート解析結果を Google Apps Script（GAS）で処理し、
CSV などの出力に整形するためのツール群です。  
本リポジトリでは、**Issue と確実に紐づく Pull Request 運用**と、  
**Dev Container 前提の安全な自動化作業**を重視します。

---

## 1. Language Preference（言語ルール）

- Codex を含むすべてのエージェントは、**日本語で応答してください**。
- 説明・確認・エラーメッセージも日本語とします。
- ユーザーが明示的に要求しない限り、英語に切り替えてはいけません。
- コード、コマンド、ファイル名は英語のままで構いません。

---

## 2. Development Environment（開発環境）

- 本リポジトリは **VSCode Dev Container（Ubuntu）前提**で開発します。
- エージェントによるファイル操作・コマンド実行は、**コンテナ内に限定**されます。
- ローカル Mac / Windows 環境への直接操作を前提にしてはいけません。

---

## 3. Project Structure & Module Organization

GAS プロジェクトのため、現時点のソースはリポジトリ直下に配置します。

- `appsscript.json` : GAS マニフェスト
- `*.js` / `*.gs` : GAS ソース（例: `コード.js`）
- `.clasp.json` : GAS プロジェクト連携情報
- `skills/` : Codex 用スキル定義
- `docs/` : 設計・仕様ドキュメント（必要になったら追加）
- `tests/` : 自動テスト（必要になったら追加）

構成が固まったら、本ファイルを必ず更新してください。

---

## 4. Build / Test / Run Commands

現在、標準的な build / test / run コマンドは定義されていません。

将来的に導入した場合は、以下のように **実行可能な例** を必ず明記してください。

- `./scripts/dev`
- `./scripts/test`
- `./scripts/build`

GAS 同期は `npx clasp push` を使います（詳細は 9 章参照）。

---

## 5. Coding Style & Naming Conventions

- 現時点では formatter / linter の公式ルールはありません。
- 新しい言語やツールを導入する場合は、
  `.editorconfig`、`pyproject.toml`、`.prettierrc` 等を追加し、
  主要ルール（インデント、行長、命名規則）を本ファイルに記載してください。
- 既存の GAS ファイル名（例: `コード.js`）は、明示的な指示がない限り変更しないこと。

---

## 6. Git / Issue / Pull Request Rules（最重要）

### 6.1 Issue と PR の紐づけ（必須）

- **PR 本文に必ず Closing keyword を含めること**
  - 例: `Closes #123` / `Fixes #123` / `Resolves #123`
- これが無い PR は、Issue 管理不備として不合格とします。

### 6.2 ブランチ命名規則

- 機能追加: `feature/<issue番号>-<short-slug>`
- バグ修正: `fix/<issue番号>-<short-slug>`
- リファクタ: `refactor/<issue番号>-<short-slug>`
- ドキュメント: `docs/<issue番号>-<short-slug>`

例:
- `feature/123-add-login`

### 6.3 コミットメッセージ（推奨）

- Issue番号を含めることを推奨します。

例:
- `feat: add login endpoint (#123)`
- `fix: handle null config (#123)`

---

## 7. Pull Request 作成ルール

### PRタイトル（推奨）
- `#123 <Issueの要約>`

### PR本文（必須構成）

- Closing keyword 行（必須）  
  `Closes #123`
- 何をしたか（要点）
- 動作確認内容

テンプレ例:

```
## Summary
- ...

## Related Issue
Closes #123

## Changes
- ...

## Testing
- [ ] Unit tests
- [ ] Manual check: ...
```

---

## 8. gh CLI コメント作成ルール（重要）

### 基本方針
`gh issue comment` / `gh pr comment` では、  
`\n` を含む1行文字列を `--body` に渡してはいけません。

### 必須ルール
- **`--body-file -` を使用する**
- 実際の改行を含む本文を標準入力で渡す

NG例（禁止）:
```bash
gh issue comment 123 --body "背景\n要望\n..."
```

---

## 9. GAS（clasp）運用ルール

### 方針
GAS 関連ファイルを更新した場合、条件を満たせば
**エージェントはユーザー確認なしで `npx clasp push` を実行してよい**。

### 実行条件
- `appsscript.json` が存在する
- `@google/clasp` が利用可能（基本は `npx clasp` を使用）
- 変更が GAS ファイルに限定されている

### 禁止事項
- 明示的な指示が無い限り `clasp pull` を実行してはいけません。
- `.clasprc.json`（認証情報）はコミットしないこと。

---

## 10. Agent Checklist（PR作成前）

- [ ] Issue番号が確定している
- [ ] ブランチ名に Issue番号が含まれている
- [ ] PR本文に `Closes #<番号>` が含まれている
- [ ] 変更内容とテスト結果が記載されている

---

## 11. Agent-Specific Notes

- エージェントは本ファイルを **憲法レベルのルール**として扱ってください。
- 不明点・曖昧点がある場合は、**勝手に判断せず必ずユーザーに確認**してください。
- 本ファイルはプロジェクト進化に合わせて随時更新されます。

---

## 12. Issue-driven Workflow（必須）

エージェントは、ユーザーから Issue 番号が渡されたら必ず次の順で進める。

### 12.1 Issueの読取り
- `gh issue view <N> --json number,title,body,labels,assignees -q ".number,.title,.body"`
- 不明点があれば作業前にユーザーへ確認する。

### 12.2 方針コメント（実装前に必須）
- まず Issue にコメントする。内容は以下を含める：
  - 実装方針（設計/手順）
  - 変更ファイル見込み
  - 影響範囲（破壊的変更の有無）
  - テスト方針（実行コマンド or 手動確認）
  - 追加で必要な確認事項
- コメント投稿は `gh issue comment <N> --body-file -` を必須とする（`\n` を含む `--body` 禁止）。

### 12.3 ユーザー承認待ち
- ユーザーが「進めてOK」と明示するまで、実装（コード変更）を開始してはいけない。

### 12.4 実装開始
- ブランチ名は規則に従い Issue番号を含める。
- コミットメッセージにも Issue番号を含める（推奨）。

### 12.5 PR作成（Issue紐づけ必須）
- PR本文に必ず `Closes #<N>` を入れる。
- PRには概要、変更点、テスト結果を含める。
