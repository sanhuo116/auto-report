# 政企部项目自动化巡检

这是政企部项目自动化巡检平台的 `auto-report` 独立栏目。

## 功能

- 需求提交：选择业务场景、租户名称、报告周期、巡检模板并上传巡检材料。
- 报告生成：调用对应 Codex skill 生成巡检报告。
- 报告提取：查看制作中、失败、成功状态并下载报告。
- 飞书归档：将生成的巡检报告上传至配置的飞书多维表格。

模板与 skill：

- 语音 SaaS：通用模板，`saas-xunjian-tongyong`
- 文本：文本项目通用模板，`chatbot-xjbg`
- 语音私有化：余杭社保，`yuhangshebao-report`

## 启动

在仓库根目录执行：

```bash
cd auto-report
python3 server.py
```

打开 `http://127.0.0.1:4174/`。

## 自动上传代码

网页服务启动后会在后台每 60 秒检查当前仓库的代码变更，并自动提交推送到 GitHub `main` 分支。运行数据、上传文件和生成报告保存在本地 `runtime/`，不会提交。

同步脚本为 `scripts/auto-sync-github.sh`；`launchd/` 中保留了后台同步配置。

远程仓库：

`https://github.com/sanhuo116/auto-report`
