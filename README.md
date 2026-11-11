# marks-getsuji-shiryo

マークスライフ株式会社（M社）の月次資料（棚卸・月次報告・取締役会資料・売上構成分析）を作るためのスキル。
中身の手順は `SKILL.md`、見せ方のルールは `rules.md`、スクリプトは `scripts/`。

## 使い方

### Claude Code
```bash
git clone https://github.com/prime-consulting-inc/marks-getsuji-shiryo.git ~/.claude/skills/marks-getsuji-shiryo
```
「M社の11月分を作って」のように頼むと使われる。更新は `git pull`。

### claude.ai
このフォルダをzipにして、設定 → 機能 → スキル からアップロード。
組織で配る場合は管理者が組織のスキルとして登録する。

## 注意
- 先方の実データ（仕訳帳、棚卸、月次報告など）はリポジトリに入れない（.gitignore 済み）。毎月チャットに添付する
- `scripts/monthly/2026-10/` は10月の実行例。パス・月・手入力値が直書きなので、当月用にコピーして書き換える
- ルールが増えたら `rules.md` に追記してコミットする
