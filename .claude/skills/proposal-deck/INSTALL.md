# インストール

正しく置けたかの判定はひとつだけ：**`<スキル置き場>\proposal-deck\SKILL.md` が存在すること。**
`proposal-deck` フォルダが二重（`…\skills\proposal-deck-codex\proposal-deck\SKILL.md` など）だと
Codex は見つけられない。スキルは置き場の**ちょうど1階層下**に置く。

| エージェント | スキル置き場（Windows） | （Mac / Linux） |
|---|---|---|
| Codex（現行） | `%USERPROFILE%\.agents\skills\` | `~/.agents/skills/` |
| Codex（旧版） | `%USERPROFILE%\.codex\skills\` | `~/.codex/skills/` |
| Claude Code | `.skill` ファイルの「Save skill」 | 同左 |

## Windows：Codex に入れてもらう（いちばん確実）

`proposal-deck-codex.zip` をダウンロードフォルダに保存してから、Codex に次をそのまま貼る:

```
ダウンロードフォルダの proposal-deck-codex.zip を展開して、
%USERPROFILE%\.agents\skills\proposal-deck\SKILL.md と
%USERPROFILE%\.codex\skills\proposal-deck\SKILL.md が存在する配置にしてください
（proposal-deck フォルダが二重にならないように）。
そのあと次の2つを実行して、最後の行を見せてください。
python "$env:USERPROFILE\.agents\skills\proposal-deck\scripts\setup_env.py"
python "$env:USERPROFILE\.agents\skills\proposal-deck\tests\run_e2e.py"
```

`ALL PASSED` が出たら、**Codex で新しいスレッドを開いて**から `$proposal-deck` で依頼する
（スキル一覧はスレッド開始時に読まれるため）。

## Windows：自分で PowerShell から入れる

```powershell
$zip = "$env:USERPROFILE\Downloads\proposal-deck-codex.zip"
foreach ($d in "$env:USERPROFILE\.agents\skills", "$env:USERPROFILE\.codex\skills") {
  New-Item -ItemType Directory -Force $d | Out-Null
  Expand-Archive -Path $zip -DestinationPath $d -Force
}
Test-Path "$env:USERPROFILE\.agents\skills\proposal-deck\SKILL.md"   # True ならOK

python "$env:USERPROFILE\.agents\skills\proposal-deck\scripts\setup_env.py"
python "$env:USERPROFILE\.agents\skills\proposal-deck\tests\run_e2e.py"
```

注意: Windows の PowerShell では `python ~/.agents/...` の `~` が展開されないことがある。
パスは必ず `$env:USERPROFILE\...` で書く。

## Mac / Linux

```bash
mkdir -p ~/.agents/skills ~/.codex/skills
unzip -o ~/Downloads/proposal-deck-codex.zip -d ~/.agents/skills
unzip -o ~/Downloads/proposal-deck-codex.zip -d ~/.codex/skills
python3 ~/.agents/skills/proposal-deck/scripts/setup_env.py
python3 ~/.agents/skills/proposal-deck/tests/run_e2e.py
```

## うまくいかないとき

| 症状 | 原因と対処 |
|---|---|
| Codex が「proposal-deck が見つからない」と言い、内蔵の Presentations スキルで作り始める | 置き場所が違う。上の `Test-Path` が True になる配置に直し、新しいスレッドで依頼する |
| `python` が見つからない／Microsoft Store が開く | Python 未導入。`winget install Python.Python.3.12` の後、ターミナルを開き直す |
| `setup_env.py` が LibreOffice を入れられない | `winget install TheDocumentFoundation.LibreOffice` を手動で実行 |
| `run_e2e.py` の rendered で失敗 | LibreOffice の導入直後はPCの再起動かターミナルの開き直しで直ることが多い |
