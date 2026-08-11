# Gotchas（踩坑记录）

> 可追加。新增经验直接在对应分区下加一条即可，不用改其它内容。

## GitHub · 挂代理 / 大文件推送（2026-08-11）

**GHW 环境推 GitHub 卡死 · 大流量走 443**
- 症状：`gh repo create --push` 生成的 HTTPS 远程推送时 `git send-pack`/`pack-objects` 进程存活但 0% CPU，卡到天荒地老；改成 SSH:22 也一样。
- 判断：控制通道（`ssh -T git@github.com`、`gh api`、`git ls-remote`）全通，只有 >10MB 的上传断流 ≈ 墙内大流量通道被掐，不是客户端问题。
- 解法：`~/.ssh/config` 加 `Host github.com → HostName ssh.github.com / Port 443`，并对 known_hosts 补 `ssh-keyscan -p 443 ssh.github.com`。切换后 133MB 一次推成。
- 换 Host 后首次连接报 `Host key verification failed` 是正常现象（22 与 443 的 known_hosts 条目不同），补 key 即可，数据通道本身走 443 更稳。

**gh 登录是会话级的**
- `gh auth login` 凭据在 `~/.config/gh/hosts.yml`，不在 git 钥匙串；SSH key 能 `ssh -T` ≠ gh API 已登录。
- `gh auth login` 交互式，SSH 不能建仓库，必须走 gh API；靠 device code（`! gh auth login` → web browser）完成。

**离线后 flash 提交要防误收录**
- `git add -A` 会把工作区里所有新文件带进来——本会话曾把插件的 skill zip（`*skill-local.zip`，含 127MB）一起 commit 了。`.gitignore` 至少要有 `*.zip`、`.DS_Store`。
- 大文件误 commit、还没 push 时：`git commit --amend` + `git gc --prune=now` 让大 blob 变 dangling，push 就不会传（`git fsck --unreachable | grep -c blob` 验证清零）。
- push 前 `git rev-list --objects --all` 没有直接量体积——用 `git ls-tree -rl HEAD | awk '{s+=$4} END {printf "%.0fMiB", s/1048576}'` 估算实际要传的字节。

**macOS 没有 GNU `timeout`**
- `timeout 300 cmd` 不存在（那是 coreutils）。要超时/盯结果用 `run_in_background` + Monitor，不要 `sleep` 轮询（会被 block）。

## Claude Code skills（2026-08-11）

**skill 目录即包**
- 本项目 8 个技能（docx / article-to-short-video-script / wechat-high-energy-commentary / modern-qimin-jimeng-video / pdfs / slides / spreadsheets / geopolitical-deep-analysis-wechat）均位于 `.claude/skills/<name>/`（`geopolitical-deep-analysis-wechat` 由 zip 解压而来）。
- `SKILL.md` 的 frontmatter `name` 必须与目录名一致，才能被 Claude 正确索引/触发。
- 新加的 skill 需要**重启会话**（或 `/reload-skills`）才会出现在可用 skills 列表里——`/reload-skills` 成功会打印 "Reloaded skills: N available (M added)"，否则确认技能目录在 `.claude/skills/` 下且有 `SKILL.md`。
- skill 包以 zip 分发时（如 `xxx-skill-local.zip`）：先 `unzip -l` 确认包含 `SKILL.md`，再解压后按 frontmatter name 移入 `.claude/skills/<name>/`，随后删 zip（`.gitignore` 已挡 `*.zip`）。

**目录命名要较真**
- 曾把工作区目录误拼成 `geoplitical`（漏了 e），文档里也跟着错。修正方法是 `mv geoplitical geopolitical`，再 grep 全仓把残留引用改干净。中文拼音/英文混排的目录名尤其容易手滑，改名后务必 `grep -rn` 检查 CLAUDE.md / README 里的引用。

**DOCX 交付必过渲染门**
- 任何产出 `.docx` 的任务必须 `cd .claude/skills/docx && python render_docx.py in.docx --output_dir out` 逐页看图过关后才交付，文本抽取会漏版式缺陷。

## 环境速查（2026-08-11）

- `gh` 安装在 Homebrew（`/opt/homebrew/bin/gh`），认证对应当前 GitHub 账号 `1998x-stack`。
- 本项目私有仓库：`content-prod-lab`（已推 main，224 文件）。默认走 SSH 443。
- 设备流授权：让用户在输入框输入 `! gh auth login` → 选择 GitHub.com / HTTPS / web browser，把 device code 粘贴到浏览器即可，无需手动灌 token。