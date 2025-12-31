---
name: gh-pg
description: GitHub Issue の実装フェーズを支援するスキル。指定 Issue の「## Implementation Plan (Codex)」コメントを実装仕様として解釈し、必要なコード変更、コミット作成、Pull Request 作成までを行う。すでに当該 Issue に紐づく PR が存在する場合は新規 PR を作らず、PR のレビュー指摘（コメント/レビュー/Review Threads）を収集して修正方針を提示し、同一ブランチへ追加コミットして PR を更新する。計画が存在しない/不十分な場合は実装せずユーザーに確認する。
---

# GH PG

## Overview

Issue の計画コメントを前提に、実装・コミット・PR 作成までを実行する。すでに PR が出ている場合は、PR の指摘コメントを読み取り、修正方針→修正コミット→PR 更新までを行う（新規 PR は作らない）。計画が無い/不足している場合は停止してユーザーに確認する。

## Workflow

1. リポジトリの `AGENTS.md` を読み、言語/運用ルールに従う。
2. 先に当該 Issue に紐づく既存 PR を探索する（存在する場合は **同一 PR を更新** する）。
   例:
   `gh issue view <N> --json "linkedPullRequests(first: 10, states: OPEN) { number,title,headRefName,url }"`
   - 1件だけ見つかった場合: その PR を対象にする。
   - 0件の場合: 新規実装フローへ進む（以降の手順でブランチ作成→PR 作成）。
   - 複数件見つかった場合: **実装を止め**、対象 PR の選択をユーザーに求める（番号とタイトルとURLを提示）。
3. Issue 本文を取得する。
   `gh issue view <N> --json number,title,body,labels,assignees -q ".number,.title,.body"`
4. Issue コメントから `## Implementation Plan (Codex)` を探し、最新の計画を取得する。
   例: `gh issue view <N> --json comments -q ".comments[].body"` を使い、該当見出しを抽出する。
5. 計画が存在しない/不十分な場合は実装を止め、ユーザーに確認する。

A. 新規 PR 作成フロー（Step 2 で PR が見つからなかった場合）
6. 計画内容を実装仕様として解釈し、必要なコード変更を行う。
7. 変更が GAS ファイルに限定される場合は、`npx clasp push` を実行する（条件は `AGENTS.md` に従う）。
8. ブランチ命名規則に従ってブランチを作成する。
9. Issue 番号を含むコミットメッセージでコミットを作成する。
10. PR を作成する。本文は `Closes #<N>` を含め、`references/pr-template.md` を使用する。
   コメント/PR 作成は必ず `--body-file -` を使う。

B. 既存 PR 更新フロー（Step 2 で PR が見つかった場合）
11. 対象 PR の詳細を取得し、レビュー指摘を収集する。
    例:
    - PR 本文/状態/ブランチ: `gh pr view <PR> --json number,title,state,headRefName,body,url`
    - PR コメント: `gh pr view <PR> --json comments -q ".comments[].body"`
    - Review Threads（インライン指摘）: `gh pr view <PR> --json reviewThreads`
12. 収集した指摘を「指摘 → 対応方針 → 変更予定箇所」の形に整理し、PR に `## Fix Plan (Codex)` 見出しでコメント投稿する。
    - 指摘が曖昧/再現条件不足の場合は、このコメント内に `ユーザー確認事項` として質問を列挙し、**回答が来るまで修正コミットは作らない**。
13. 指摘が解消できるだけの情報が揃っている場合、PR の head ブランチをチェックアウトして修正する。
    - 既存ブランチに追加コミットする（新規 PR は作らない）。
14. Issue 番号と PR 番号を含むコミットメッセージで追加コミットを作成し、push する。
15. 修正内容の要約と、各指摘がどう解消されたか（もしくは未対応理由）を PR にコメントする（`--body-file -`）。

## Notes

- 返答は必ず日本語にする。
- 計画コメントは実装仕様として最優先する。
- `--body` に `\n` を含める投稿は禁止。必ず `--body-file -` を使う。
- 既存 PR がある場合、**新しい PR を作らない**。同一 PR の head ブランチに追加コミットして更新する。
- PR 指摘の取得は「PR コメント」だけでなく「Review Threads（インライン）」も対象にする。
- ユーザーが「PR にコメントしました」「指摘入れました」等と言ったら、まず Step 11 のコマンドを再実行して最新指摘を再収集する。
- `clasp` を使う場合は `npx clasp` を優先する（`@google/clasp` は devDependencies）。
- 明示的な指示が無い限り `clasp pull` は実行しない。
