from __future__ import annotations

import cgi
import json
import os
import shutil
import subprocess
import threading
import time
import uuid
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode, unquote, urlparse
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parent
RUNTIME = ROOT / "runtime"
UPLOADS = RUNTIME / "uploads"
REPORTS = RUNTIME / "reports"
TASKS_FILE = RUNTIME / "tasks.json"
CODEX_BIN = shutil.which("codex") or "/Applications/ChatGPT.app/Contents/Resources/codex"
FEISHU_API_BASE = os.environ.get("FEISHU_API_BASE", "https://open.feishu.cn").rstrip("/")
DEFAULT_FEISHU_WIKI_TOKEN = "VMOGwc4cAivY0hkPtIzcfP7jnHh"
DEFAULT_FEISHU_BASE_TOKEN = "PAzTbBX7EamTfzs9aXzcWBaanZe"
DEFAULT_FEISHU_TABLE_ID = "tblfCjtJ4uImwNyi"
DEFAULT_FEISHU_VIEW_ID = "vew4bWEOlm"
DEFAULT_FEISHU_REPORT_FIELD_ID = "fld4ab4qf8"
LARK_CLI_BIN = os.environ.get("FEISHU_CLI_BIN", "").strip() or shutil.which("lark-cli") or next(
    (
        str(path)
        for path in sorted(
            Path.home().glob(".nvm/versions/node/*/lib/node_modules/@larksuite/cli/bin/lark-cli")
        )
        if path.exists()
    ),
    "",
)

TEMPLATE_CATALOG = {
    "saas-general": {
        "scenario": "语音 - SaaS",
        "template": "通用模板",
        "taskName": "语音saas巡检",
        "skill": "saas-xunjian-tongyong",
    },
    "hangzhou-shebao": {
        "scenario": "文本",
        "template": "杭州社保",
        "taskName": "杭州社保巡检",
        "skill": "hangzhoushebao-report",
    },
    "text-project": {
        "scenario": "文本",
        "template": "文本项目通用模板",
        "taskName": "文本项目巡检报告",
        "skill": "chatbot-xjbg",
    },
    "yuhang-shebao": {
        "scenario": "语音 - 私有化",
        "template": "余杭社保",
        "taskName": "余杭社保巡检",
        "skill": "yuhangshebao-report",
    },
}


def ensure_runtime() -> None:
    UPLOADS.mkdir(parents=True, exist_ok=True)
    REPORTS.mkdir(parents=True, exist_ok=True)
    RUNTIME.mkdir(parents=True, exist_ok=True)
    if not TASKS_FILE.exists():
        TASKS_FILE.write_text("[]", encoding="utf-8")


def auto_sync_loop() -> None:
    """Keep the GitHub mirror updated while the inspection service is running."""
    sync_script = ROOT / "scripts" / "auto-sync-github.sh"
    repo_root = ROOT
    if not sync_script.exists():
        repo_root = ROOT.parent
        sync_script = repo_root / "scripts" / "auto-sync-github.sh"
    if not sync_script.exists() or not (repo_root / ".git").exists():
        return

    while True:
        try:
            completed = subprocess.run(
                ["/bin/zsh", str(sync_script)],
                cwd=repo_root,
                capture_output=True,
                text=True,
                timeout=120,
            )
            if completed.returncode != 0:
                detail = (completed.stderr or completed.stdout).strip()
                print(f"GitHub 自动同步失败: {detail[-800:]}", flush=True)
        except (OSError, subprocess.SubprocessError) as error:
            print(f"GitHub 自动同步异常: {error}", flush=True)
        time.sleep(60)


def read_tasks() -> list[dict]:
    ensure_runtime()
    try:
        return json.loads(TASKS_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return []


def write_tasks(tasks: list[dict]) -> None:
    ensure_runtime()
    temp_file = TASKS_FILE.with_suffix(".tmp")
    temp_file.write_text(json.dumps(tasks, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_file.replace(TASKS_FILE)


def update_task(task_id: str, **updates) -> dict | None:
    tasks = read_tasks()
    for task in tasks:
        if task["id"] == task_id:
            task.update(updates)
            task["updatedAt"] = datetime.now().isoformat(timespec="seconds")
            write_tasks(tasks)
            return task
    return None


def safe_filename(name: str) -> str:
    return Path(name).name.replace("/", "_").replace("\\", "_")


def parse_object(value) -> dict:
    if isinstance(value, dict):
        return value
    if not value:
        return {}
    try:
        parsed = json.loads(value)
    except (TypeError, json.JSONDecodeError):
        return {}
    return parsed if isinstance(parsed, dict) else {}


def feishu_request(
    method: str,
    path: str,
    access_token: str = "",
    payload: dict | None = None,
    body: bytes | None = None,
    content_type: str = "application/json",
) -> dict:
    request_body = body
    headers = {"Accept": "application/json"}
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"
    if payload is not None:
        request_body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        headers["Content-Type"] = content_type
    elif body is not None:
        headers["Content-Type"] = content_type

    request = Request(f"{FEISHU_API_BASE}{path}", data=request_body, headers=headers, method=method)
    try:
        with urlopen(request, timeout=45) as response:
            response_body = response.read()
    except HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")[:800]
        raise RuntimeError(f"飞书接口 HTTP {error.code}: {detail}") from error
    except URLError as error:
        raise RuntimeError(f"无法连接飞书接口: {error.reason}") from error

    try:
        result = json.loads(response_body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RuntimeError("飞书接口返回了无法解析的响应") from error
    if result.get("code", 0) != 0:
        raise RuntimeError(f'飞书接口失败（{result.get("code")}）：{result.get("msg", "未知错误")}')
    return result


def feishu_destination(task: dict) -> dict:
    destination = parse_object(task.get("feishuDestination"))
    fields = destination.get("fields") if isinstance(destination.get("fields"), dict) else {}
    return {
        "base_token": os.environ.get("FEISHU_BASE_TOKEN", "").strip()
        or destination.get("baseToken")
        or DEFAULT_FEISHU_BASE_TOKEN,
        "table_id": os.environ.get("FEISHU_TABLE_ID", "").strip()
        or destination.get("tableId")
        or DEFAULT_FEISHU_TABLE_ID,
        "view_id": os.environ.get("FEISHU_VIEW_ID", "").strip()
        or destination.get("viewId")
        or DEFAULT_FEISHU_VIEW_ID,
        "report_field_id": os.environ.get("FEISHU_REPORT_FIELD_ID", "").strip()
        or destination.get("reportFieldId")
        or DEFAULT_FEISHU_REPORT_FIELD_ID,
        "cli_bin": LARK_CLI_BIN,
        "cli_profile": os.environ.get("FEISHU_CLI_PROFILE", "").strip(),
        "fields": {
            "tenant": os.environ.get("FEISHU_FIELD_TENANT", "").strip()
            or fields.get("tenant")
            or "客户名称",
            "period": os.environ.get("FEISHU_FIELD_PERIOD", "").strip()
            or fields.get("period")
            or "巡检报告周期",
            "report": os.environ.get("FEISHU_FIELD_REPORT", "").strip()
            or fields.get("report")
            or "巡检报告",
        },
    }


def lark_cli_json(config: dict, args: list[str]) -> dict:
    if not config["cli_bin"]:
        raise RuntimeError("未找到 lark-cli，请先安装并登录飞书 CLI")
    command = [config["cli_bin"], *args]
    if config["cli_profile"]:
        command.extend(["--profile", config["cli_profile"]])
    command.extend(["--format", "json"])
    try:
        completed = subprocess.run(
            command,
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=120,
        )
    except OSError as error:
        raise RuntimeError(f"无法启动飞书 CLI：{error}") from error
    raw_output = (completed.stdout or "").strip()
    if completed.returncode != 0:
        details = (completed.stderr or raw_output or "").strip()
        raise RuntimeError(f"飞书 CLI 执行失败：{details[-1200:]}")
    try:
        result = json.loads(raw_output[raw_output.find("{") :])
    except (ValueError, json.JSONDecodeError) as error:
        raise RuntimeError(f"飞书 CLI 返回无法解析：{raw_output[-800:]}") from error
    if result.get("ok") is False:
        error = result.get("error", {})
        raise RuntimeError(error.get("message") or "飞书 CLI 返回失败")
    return result


def extract_record_id(response: dict) -> str:
    candidates = [
        response.get("data", {}).get("record", {}).get("record_id"),
        *(response.get("data", {}).get("record", {}).get("record_id_list") or []),
        response.get("data", {}).get("record_id"),
        response.get("record", {}).get("record_id"),
    ]
    return next((str(candidate) for candidate in candidates if candidate), "")


def archive_report_to_feishu(task: dict, report_path: Path) -> str:
    config = feishu_destination(task)
    record_id = task.get("feishuRecordId", "")
    if not record_id:
        period = f'{task["periodStart"]} 至 {task["periodEnd"]}'
        response = lark_cli_json(
            config,
            [
                "base",
                "+record-upsert",
                "--base-token",
                config["base_token"],
                "--table-id",
                config["table_id"],
                "--json",
                json.dumps(
                    {
                        config["fields"]["tenant"]: task["tenant"],
                        config["fields"]["period"]: period,
                    },
                    ensure_ascii=False,
                ),
            ],
        )
        record_id = extract_record_id(response)
        if not record_id:
            raise RuntimeError("飞书 CLI 创建记录成功但未返回记录 ID")
        update_task(task["id"], feishuRecordId=record_id)

    lark_cli_json(
        config,
        [
            "base",
            "+record-upload-attachment",
            "--base-token",
            config["base_token"],
            "--table-id",
            config["table_id"],
            "--record-id",
            record_id,
            "--field-id",
            config["report_field_id"],
            "--file",
            str(report_path),
        ],
    )
    return record_id


def archive_task(task_id: str) -> None:
    task = next((item for item in read_tasks() if item["id"] == task_id), None)
    if not task:
        return
    report_path = Path(task["reportPath"]) if task.get("reportPath") else None
    if not report_path or not report_path.exists():
        update_task(task_id, feishuSyncStatus="归档失败", feishuError="报告文件不存在")
        return
    update_task(task_id, feishuSyncStatus="归档中", feishuError="")
    try:
        record_id = archive_report_to_feishu(task, report_path)
        update_task(
            task_id,
            feishuSyncStatus="已归档",
            feishuRecordId=record_id,
            feishuError="",
            feishuArchivedAt=datetime.now().isoformat(timespec="seconds"),
        )
    except Exception as error:
        update_task(task_id, feishuSyncStatus="归档失败", feishuError=str(error))


def resolve_template_id(task: dict) -> str:
    template_id = task.get("templateId", "")
    if template_id in TEMPLATE_CATALOG:
        return template_id
    for candidate_id, catalog in TEMPLATE_CATALOG.items():
        if (
            catalog["scenario"] == task.get("scenario")
            and catalog["skill"] == task.get("skill")
        ):
            return candidate_id
    return ""


def task_response(task: dict) -> dict:
    template_id = resolve_template_id(task)
    result = {
        "id": task["id"],
        "taskId": task["id"],
        "tenant": task["tenant"],
        "name": task["name"],
        "reportName": task["name"],
        "scenario": task["scenario"],
        "templateId": template_id,
        "template": task["template"],
        "skill": task["skill"],
        "executor": "codex",
        "taskName": task["taskName"],
        "periodStart": task["periodStart"],
        "periodEnd": task["periodEnd"],
        "period": f'{task["periodStart"].replace("-", ".")} - {task["periodEnd"].replace("-", ".")}',
        "generatedAt": task.get("generatedAt") or task["createdAt"],
        "createdAt": task["createdAt"],
        "status": task["status"],
        "statusText": {"processing": "制作中", "success": "成功", "failed": "失败"}[task["status"]],
        "feishuSyncStatus": task.get("feishuSyncStatus", "待归档"),
        "feishuRecordId": task.get("feishuRecordId", ""),
        "feishuError": task.get("feishuError", ""),
    }
    if task.get("reportPath"):
        result["downloadUrl"] = f'/api/inspection-reports/{task["id"]}/download'
    if task.get("error"):
        result["error"] = task["error"]
    return result


def run_codex_task(task_id: str) -> None:
    task = next((item for item in read_tasks() if item["id"] == task_id), None)
    if not task:
        return

    template_id = resolve_template_id(task)
    template_catalog = TEMPLATE_CATALOG.get(template_id, {})
    task_dir = UPLOADS / task_id
    output_path = REPORTS / f'{task_id}.docx'
    last_message = task_dir / "codex-last-message.txt"
    skill_dir = Path.home() / ".codex" / "skills" / task["skill"]
    call_files = [
        str(path)
        for path in sorted(task_dir.iterdir())
        if path.is_file() and path.name != "request.json"
    ]
    calls_path = "\n".join(f"- {path}" for path in call_files)
    prompt = f"""
执行 Codex 任务：{task.get("taskName") or template_catalog.get("taskName", "巡检报告")}。

请严格使用 skill `{task["skill"]}`，根据上传的通话记录生成客户可交付的巡检报告。
当前网页选择的巡检模板：{task.get("template") or template_catalog.get("template", "未命名模板")}（templateId: {template_id}）。
请将本次任务视为真实生产任务，直接读取并检查上传目录中的全部材料，不要忽略图片、Word、PDF、压缩包或文本文件。
客户名称：{task["tenant"]}
业务场景：{task["scenario"]}
报告周期：{task["periodStart"]} 至 {task["periodEnd"]}
上传材料：
{calls_path}
报告输出路径：{output_path}
备注：{task.get("remarks") or "无"}

必须实际生成 DOCX 文件并完成 skill 要求的校验，不要只返回文字说明。
如果有多个通话记录文件，请按 skill 要求合并或逐一处理后再汇总。
使用客户名称和报告周期生成正式报告，报告完成后保存在指定输出路径。
如本任务使用 `chatbot-xjbg`，请执行“文本项目巡检报告”流程，结合聊天记录、知识库统计、案例截图和参考模板生成文本场景巡检报告；材料不足时基于现有材料完成可交付版本，并在报告中保留真实客户名称和报告周期。
    """
    command = [
        CODEX_BIN,
        "exec",
        "--ephemeral",
        "--cd",
        str(ROOT),
        "--add-dir",
        str(skill_dir),
        "--skip-git-repo-check",
        "--dangerously-bypass-approvals-and-sandbox",
        "--output-last-message",
        str(last_message),
        prompt,
    ]

    try:
        completed = subprocess.run(
            command,
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=30 * 60,
        )
        if completed.returncode != 0 or not output_path.exists():
            details = (completed.stderr or completed.stdout or "").strip()
            if len(details) > 1200:
                details = details[-1200:]
            raise RuntimeError(details or "Codex 任务未生成报告文件")
        update_task(
            task_id,
            status="success",
            generatedAt=datetime.now().isoformat(timespec="seconds"),
            reportPath=str(output_path),
            feishuSyncStatus="待归档",
            feishuError="",
        )
        archive_task(task_id)
    except Exception as error:
        update_task(task_id, status="failed", error=str(error))


class InspectionHandler(SimpleHTTPRequestHandler):
    server_version = "InspectionReportServer/1.0"

    def send_json(self, payload: dict | list, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        path = unquote(urlparse(self.path).path)
        if path == "/api/inspection-reports":
            tasks = sorted(read_tasks(), key=lambda item: item["createdAt"], reverse=True)
            self.send_json([task_response(task) for task in tasks])
            return

        prefix = "/api/inspection-reports/"
        if path.startswith(prefix) and path.endswith("/download"):
            task_id = path[len(prefix) : -len("/download")].strip("/")
            task = next((item for item in read_tasks() if item["id"] == task_id), None)
            report_path = Path(task["reportPath"]) if task and task.get("reportPath") else None
            if not report_path or not report_path.exists():
                self.send_json({"message": "报告尚未生成"}, status=404)
                return
            body = report_path.read_bytes()
            download_name = safe_filename(task["name"]) + ".docx"
            encoded_name = quote(download_name)
            self.send_response(200)
            self.send_header(
                "Content-Disposition",
                f'attachment; filename="inspection-report-{task_id}.docx"; filename*=UTF-8\'\'{encoded_name}',
            )
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        super().do_GET()

    def do_POST(self) -> None:
        path = unquote(urlparse(self.path).path)
        if path == "/api/inspection-requests":
            self.create_task()
            return

        prefix = "/api/inspection-requests/"
        if path.startswith(prefix) and path.endswith("/retry"):
            task_id = path[len(prefix) : -len("/retry")].strip("/")
            task = next((item for item in read_tasks() if item["id"] == task_id), None)
            if not task:
                self.send_json({"message": "任务不存在"}, status=404)
                return
            update_task(task_id, status="processing", error="", reportPath="")
            threading.Thread(target=run_codex_task, args=(task_id,), daemon=True).start()
            refreshed = next(item for item in read_tasks() if item["id"] == task_id)
            self.send_json(task_response(refreshed), status=202)
            return

        report_prefix = "/api/inspection-reports/"
        if path.startswith(report_prefix) and path.endswith("/archive"):
            task_id = path[len(report_prefix) : -len("/archive")].strip("/")
            task = next((item for item in read_tasks() if item["id"] == task_id), None)
            if not task:
                self.send_json({"message": "任务不存在"}, status=404)
                return
            if not task.get("reportPath") or not Path(task["reportPath"]).exists():
                self.send_json({"message": "报告文件尚未生成"}, status=409)
                return
            update_task(task_id, feishuSyncStatus="归档中", feishuError="")
            threading.Thread(target=archive_task, args=(task_id,), daemon=True).start()
            refreshed = next(item for item in read_tasks() if item["id"] == task_id)
            self.send_json(task_response(refreshed), status=202)
            return

        self.send_json({"message": "接口不存在"}, status=404)

    def create_task(self) -> None:
        content_type = self.headers.get("content-type", "")
        if "multipart/form-data" not in content_type:
            self.send_json({"message": "请使用 multipart/form-data 提交"}, status=415)
            return

        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type},
        )

        def value(name: str) -> str:
            field = form[name] if name in form else None
            if isinstance(field, list):
                field = field[0]
            return str(field.value).strip() if field is not None else ""

        tenant = value("tenant")
        scenario = value("scenario")
        template_id = value("templateId")
        catalog = TEMPLATE_CATALOG.get(template_id)
        if not tenant or not catalog or catalog["scenario"] != scenario:
            self.send_json({"message": "租户名称、业务场景或巡检模板无效"}, status=400)
            return

        file_fields = form["files"] if "files" in form else []
        if not isinstance(file_fields, list):
            file_fields = [file_fields]
        file_fields = [field for field in file_fields if getattr(field, "filename", None)]
        if not file_fields:
            self.send_json({"message": "至少上传一份通话记录"}, status=400)
            return

        task_id = uuid.uuid4().hex
        created_at = datetime.now().isoformat(timespec="seconds")
        task_dir = UPLOADS / task_id
        task_dir.mkdir(parents=True, exist_ok=True)
        saved_files = []
        for field in file_fields:
            filename = safe_filename(field.filename)
            file_path = task_dir / filename
            with file_path.open("wb") as output:
                shutil.copyfileobj(field.file, output)
            saved_files.append(filename)

        request_data = {
            "tenant": tenant,
            "scenario": scenario,
            "templateId": template_id,
            "template": catalog["template"],
            "skill": catalog["skill"],
            "taskName": catalog["taskName"],
            "periodStart": value("periodStart"),
            "periodEnd": value("periodEnd"),
            "remarks": value("remarks"),
            "destination": value("destination"),
            "files": saved_files,
        }
        (task_dir / "request.json").write_text(json.dumps(request_data, ensure_ascii=False, indent=2), encoding="utf-8")

        task = {
            "id": task_id,
            "tenant": tenant,
            "name": f'{tenant}巡检报告',
            "scenario": scenario,
            "templateId": template_id,
            "template": request_data["template"],
            "skill": request_data["skill"],
            "taskName": catalog["taskName"],
            "periodStart": request_data["periodStart"],
            "periodEnd": request_data["periodEnd"],
            "remarks": request_data["remarks"],
            "createdAt": created_at,
            "status": "processing",
            "feishuSyncStatus": "待归档",
            "feishuDestination": request_data["destination"],
            "files": saved_files,
        }
        tasks = read_tasks()
        tasks.append(task)
        write_tasks(tasks)
        threading.Thread(target=run_codex_task, args=(task_id,), daemon=True).start()
        self.send_json(task_response(task), status=202)


def main() -> None:
    ensure_runtime()
    threading.Thread(target=auto_sync_loop, name="github-auto-sync", daemon=True).start()
    os.chdir(ROOT)
    server = ThreadingHTTPServer(("127.0.0.1", 4174), InspectionHandler)
    print("Inspection report server running at http://127.0.0.1:4174/")
    server.serve_forever()


if __name__ == "__main__":
    main()
