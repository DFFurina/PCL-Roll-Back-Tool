from __future__ import annotations

import hashlib
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "PCL Roll Back Tool"
TAG = "2.14.0.0"
REPOSITORIES = ("Meloong-Git/PCL", "DFFurina/PCLBackup")
TARGET_DIR_NAME = f"PCL-{TAG}"
SOURCE_ROOT_NAME = f"PCL-{TAG}"
EXE_NAME = "Plain Craft Launcher 2.exe"
RELEASE_ZIP_NAME = "最新正式版.zip"
MAX_DOWNLOAD_BYTES = 250 * 1024 * 1024
MAX_EXTRACT_BYTES = 600 * 1024 * 1024
MAX_ARCHIVE_FILES = 30_000
USER_AGENT = "PCL-Standalone-Rollback/1.0"
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

LogCallback = Callable[[str], None]
ProgressCallback = Callable[[float], None]
StatusCallback = Callable[[str], None]


class FetchError(Exception):
    pass


class NetworkFailure(FetchError):
    pass


class RemoteNotFound(FetchError):
    pass


class RollbackError(Exception):
    pass


@dataclass(frozen=True)
class ExeInfo:
    path: Path
    file_version: str | None
    product_version: str | None
    product_name: str | None
    sha256: str
    size: int


@dataclass(frozen=True)
class FetchResult:
    repository: str
    source_url: str
    source_sha256: str
    exe_path: Path
    exe_version: str | None
    exe_sha256: str


@dataclass(frozen=True)
class RunningProcess:
    pid: int
    executable_path: Path | None


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def run_powershell_json(command: str, timeout: int = 25) -> object:
    completed = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            command,
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        creationflags=CREATE_NO_WINDOW,
        check=False,
    )
    if completed.returncode != 0:
        raise RollbackError(
            f"PowerShell 查询失败：{completed.stderr.strip() or completed.stdout.strip()}"
        )
    text = completed.stdout.strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError as error:
        raise RollbackError(f"PowerShell 返回无法解析：{text}") from error


def inspect_exe(path: Path) -> ExeInfo:
    path = path.resolve()
    if not path.is_file():
        raise RollbackError(f"文件不存在：{path}")
    with path.open("rb") as stream:
        if stream.read(2) != b"MZ":
            raise RollbackError(f"不是有效的 Windows EXE：{path}")
    escaped = str(path).replace("'", "''")
    command = (
        f"$item = Get-Item -LiteralPath '{escaped}' -ErrorAction Stop; "
        "[pscustomobject]@{"
        "FileVersion=$item.VersionInfo.FileVersion;"
        "ProductVersion=$item.VersionInfo.ProductVersion;"
        "ProductName=$item.VersionInfo.ProductName"
        "} | ConvertTo-Json -Compress"
    )
    raw = run_powershell_json(command)
    if not isinstance(raw, dict):
        raise RollbackError(f"读取 EXE 版本信息失败：{path}")
    return ExeInfo(
        path=path,
        file_version=str(raw.get("FileVersion") or "") or None,
        product_version=str(raw.get("ProductVersion") or "") or None,
        product_name=str(raw.get("ProductName") or "") or None,
        sha256=sha256_file(path),
        size=path.stat().st_size,
    )


def request_json(url: str, timeout: int = 25) -> dict:
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": USER_AGENT,
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            payload = response.read()
    except urllib.error.HTTPError as error:
        if error.code == 404:
            raise RemoteNotFound(f"HTTP 404：{url}") from error
        if error.code in (403, 429):
            raise NetworkFailure(
                f"HTTP {error.code}，可能是 GitHub API 限流或网络策略拦截：{url}"
            ) from error
        if 500 <= error.code < 600:
            raise NetworkFailure(f"GitHub 服务端 HTTP {error.code}：{url}") from error
        raise FetchError(f"HTTP {error.code}：{url}") from error
    except (urllib.error.URLError, TimeoutError, OSError) as error:
        raise NetworkFailure(f"网络连接失败：{url}（{error}）") from error
    try:
        value = json.loads(payload.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise FetchError(f"服务器返回的不是有效 JSON：{url}") from error
    if not isinstance(value, dict):
        raise FetchError(f"服务器返回的 JSON 结构异常：{url}")
    return value


def check_repository(repo: str, log: LogCallback) -> dict[str, bool]:
    result = {"repository": False, "tag": False, "release": False}
    try:
        data = request_json(f"https://api.github.com/repos/{repo}")
        result["repository"] = True
        log(
            f"仓库存在：{repo}（默认分支 {data.get('default_branch', '?')}，"
            f"archived={data.get('archived', '?')}）"
        )
    except RemoteNotFound:
        log(f"仓库不存在或不可见：{repo}")
        return result
    except FetchError as error:
        log(f"检查仓库失败：{repo}：{error}")
        return result

    try:
        request_json(
            f"https://api.github.com/repos/{repo}/git/ref/tags/{urllib.parse.quote(TAG)}"
        )
        result["tag"] = True
        log(f"标签存在：{repo}@{TAG}")
    except RemoteNotFound:
        log(f"标签不存在：{repo}@{TAG}")
    except FetchError as error:
        log(f"检查标签失败：{repo}@{TAG}：{error}")

    try:
        request_json(
            f"https://api.github.com/repos/{repo}/releases/tags/{urllib.parse.quote(TAG)}"
        )
        result["release"] = True
        log(f"Release 存在：{repo}@{TAG}")
    except RemoteNotFound:
        log(f"Release 不存在：{repo}@{TAG}；仍可尝试 Git 标签源码包")
    except FetchError as error:
        log(f"检查 Release 失败：{repo}@{TAG}：{error}")
    return result


def download_file(
    url: str,
    destination: Path,
    log: LogCallback,
    progress: ProgressCallback,
) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    partial = destination.with_name(destination.name + ".part")
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    log(f"下载源码包：{url}")
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            length_text = response.headers.get("Content-Length")
            total = int(length_text) if length_text else 0
            if total > MAX_DOWNLOAD_BYTES:
                raise FetchError(f"源码包超过允许大小：{total} 字节")
            downloaded = 0
            with partial.open("wb") as output:
                while True:
                    chunk = response.read(256 * 1024)
                    if not chunk:
                        break
                    downloaded += len(chunk)
                    if downloaded > MAX_DOWNLOAD_BYTES:
                        raise FetchError("源码包超过允许大小")
                    output.write(chunk)
                    if total:
                        progress(min(downloaded / total, 1.0))
        if downloaded < 1024:
            raise FetchError("下载内容过小，可能是错误页面")
        os.replace(partial, destination)
        log(f"源码包下载完成：{destination.stat().st_size} 字节")
    except urllib.error.HTTPError as error:
        partial.unlink(missing_ok=True)
        if error.code == 404:
            raise RemoteNotFound(f"源码包 HTTP 404：{url}") from error
        raise NetworkFailure(f"源码包 HTTP {error.code}：{url}") from error
    except (urllib.error.URLError, TimeoutError, OSError) as error:
        partial.unlink(missing_ok=True)
        if isinstance(error, FetchError):
            raise
        raise NetworkFailure(f"源码包下载失败：{url}（{error}）") from error
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def safe_extract_zip(
    archive: Path,
    destination: Path,
    log: LogCallback,
) -> None:
    destination = destination.resolve()
    destination.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(archive) as zip_file:
            bad = zip_file.testzip()
            if bad:
                raise FetchError(f"ZIP 校验失败：{bad}")
            members = zip_file.infolist()
            if len(members) > MAX_ARCHIVE_FILES:
                raise FetchError(f"ZIP 文件数量异常：{len(members)}")
            total = sum(member.file_size for member in members)
            if total > MAX_EXTRACT_BYTES:
                raise FetchError(f"ZIP 解压后过大：{total} 字节")
            for member in members:
                normalized = member.filename.replace("\\", "/")
                parts = Path(normalized).parts
                if (
                    not normalized
                    or normalized.startswith("/")
                    or ".." in parts
                    or (parts and ":" in parts[0])
                ):
                    raise FetchError(f"ZIP 包含不安全路径：{member.filename}")
                target = (destination / normalized).resolve()
                if target != destination and destination not in target.parents:
                    raise FetchError(f"ZIP 路径越界：{member.filename}")
                if member.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with zip_file.open(member) as source, target.open("wb") as output:
                    shutil.copyfileobj(source, output, length=1024 * 1024)
    except zipfile.BadZipFile as error:
        raise FetchError(f"不是有效 ZIP：{archive}") from error
    log(f"ZIP 解压完成：{archive.name} → {destination}")


def archive_urls(repo: str, allow_mirror: bool) -> list[str]:
    official = f"https://github.com/{repo}/archive/refs/tags/{TAG}.zip"
    urls = [
        official,
        f"https://codeload.github.com/{repo}/zip/refs/tags/{TAG}",
    ]
    if allow_mirror:
        urls.extend(
            [
                f"https://ghproxy.net/{official}",
                f"https://ghfast.top/{official}",
            ]
        )
    return urls


def find_source_root(extracted: Path) -> Path:
    """Find the repository root without relying on the GitHub ZIP root name."""
    preferred = extracted / SOURCE_ROOT_NAME
    if preferred.is_dir() and any(preferred.iterdir()):
        return preferred

    children = [child for child in extracted.iterdir() if child.is_dir()]
    if not children:
        if (extracted / RELEASE_ZIP_NAME).is_file():
            return extracted
        raise FetchError("源码包解压后找不到源码根目录")

    def score(candidate: Path) -> int:
        value = 0
        if (candidate / RELEASE_ZIP_NAME).is_file():
            value += 100
        if (candidate / "Plain Craft Launcher 2.sln").is_file():
            value += 40
        if (candidate / "Plain Craft Launcher 2").is_dir():
            value += 30
        if (candidate / ".gitmodules").is_file():
            value += 20
        for name in ("README.md", "LICENCE", "LICENSE"):
            if (candidate / name).is_file():
                value += 10
                break
        lowered = candidate.name.lower()
        if lowered in {
            SOURCE_ROOT_NAME.lower(),
            "pclbackup",
            f"pclbackup-{TAG}".lower(),
        }:
            value += 20
        elif "pcl" in lowered or "plaincraft" in lowered:
            value += 5
        return value

    ranked = sorted(children, key=lambda item: (score(item), item.name), reverse=True)
    if ranked and score(ranked[0]) > 0:
        return ranked[0]
    if len(children) == 1:
        return children[0]
    raise FetchError(
        "源码包解压后存在多个候选目录，且无法通过内容识别源码根目录"
    )


def fetch_and_prepare(
    base_dir: Path,
    allow_mirror: bool,
    log: LogCallback,
    progress: ProgressCallback,
    status: StatusCallback,
) -> FetchResult:
    base_dir = base_dir.resolve()
    base_dir.mkdir(parents=True, exist_ok=True)
    status("检查仓库、标签和 Release……")
    progress(0.02)
    states = {repo: check_repository(repo, log) for repo in REPOSITORIES}
    ordered = [
        repo
        for repo in REPOSITORIES
        if all(states[repo].values())
    ]
    ordered.extend(repo for repo in REPOSITORIES if repo not in ordered)

    target = (base_dir / TARGET_DIR_NAME).resolve()
    if target.parent != base_dir or target.name != TARGET_DIR_NAME:
        raise FetchError("目标目录路径未通过安全校验")
    work_dir = Path(tempfile.mkdtemp(prefix=".pcl-fetch-", dir=base_dir))
    source_zip = work_dir / f"PCL-{TAG}.zip"
    extracted = work_dir / "extracted"
    downloaded = False
    selected_repo = ""
    selected_url = ""
    target_prepared = False
    completed = False
    network_failed = False
    failures: list[str] = []
    try:
        for repo in ordered:
            for url in archive_urls(repo, allow_mirror):
                try:
                    status(f"下载 {repo}@{TAG} 源码包……")
                    download_file(
                        url,
                        source_zip,
                        log,
                        lambda value: progress(0.1 + value * 0.55),
                    )
                    selected_repo = repo
                    selected_url = url
                    downloaded = True
                    break
                except NetworkFailure as error:
                    network_failed = True
                    failures.append(str(error))
                    log(f"下载源失败，继续回退：{error}")
                except RemoteNotFound as error:
                    failures.append(str(error))
                    log(f"下载源不存在，继续回退：{error}")
            if downloaded:
                break
        if not downloaded:
            if all(not state["repository"] for state in states.values()):
                raise FetchError("主仓库和备份仓库均不存在或不可访问")
            if network_failed:
                raise NetworkFailure(
                    "所有下载源均失败，更可能是网络、TLS、代理或 GitHub 限流问题"
                )
            raise FetchError("所有官方源和已启用镜像均不可用")

        if target.exists():
            log(f"目标目录已存在，自动覆盖：{target}")
            shutil.rmtree(target)
        target.mkdir(parents=True)
        target_prepared = True

        status("校验并解压源码包……")
        source_hash = sha256_file(source_zip)
        log(f"源码包 SHA-256：{source_hash}")
        extracted.mkdir(parents=True)
        safe_extract_zip(source_zip, extracted, log)
        source_root = find_source_root(extracted)
        staged_source = target / SOURCE_ROOT_NAME
        shutil.move(str(source_root), str(staged_source))

        nested = staged_source / RELEASE_ZIP_NAME
        if not nested.is_file():
            matches = list(staged_source.rglob(RELEASE_ZIP_NAME))
            if not matches:
                raise FetchError(f"源码包中未找到 {RELEASE_ZIP_NAME}")
            nested = matches[0]
        log(f"找到内置启动器包：{nested}")
        progress(0.75)

        status("提取启动器……")
        release_dir = target / "Release"
        safe_extract_zip(nested, release_dir, log)
        exe_path = release_dir / EXE_NAME
        if not exe_path.is_file():
            matches = list(release_dir.rglob(EXE_NAME))
            if not matches:
                raise FetchError(f"内置 ZIP 中未找到 {EXE_NAME}")
            exe_path = matches[0]
        info = inspect_exe(exe_path)
        log(f"启动器路径：{info.path}")
        log(f"启动器实际版本：{info.file_version or '无法读取'}")
        log(f"启动器 SHA-256：{info.sha256}")

        nested.unlink(missing_ok=True)
        source_zip.unlink(missing_ok=True)
        manifest = {
            "fetched_at_utc": time.strftime(
                "%Y-%m-%dT%H:%M:%SZ", time.gmtime()
            ),
            "requested_tag": TAG,
            "repository": selected_repo,
            "source_url": selected_url,
            "source_sha256": source_hash,
            "exe_path": str(info.path),
            "exe_file_version": info.file_version,
            "exe_sha256": info.sha256,
        }
        (target / "fetch-manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        completed = True
        progress(1.0)
        status("下载与准备完成")
        return FetchResult(
            repository=selected_repo,
            source_url=selected_url,
            source_sha256=source_hash,
            exe_path=info.path,
            exe_version=info.file_version,
            exe_sha256=info.sha256,
        )
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
        if target_prepared and not completed:
            shutil.rmtree(target, ignore_errors=True)


def find_running_pcl() -> list[RunningProcess]:
    command = (
        "Get-CimInstance Win32_Process -Filter \"Name='Plain Craft Launcher 2.exe'\" "
        "| Select-Object ProcessId,ExecutablePath | ConvertTo-Json -Compress"
    )
    raw = run_powershell_json(command)
    if raw is None:
        return []
    items = raw if isinstance(raw, list) else [raw]
    result: list[RunningProcess] = []
    for item in items:
        if not isinstance(item, dict):
            continue
        try:
            pid = int(item.get("ProcessId"))
        except (TypeError, ValueError):
            continue
        path = item.get("ExecutablePath")
        result.append(
            RunningProcess(
                pid=pid,
                executable_path=Path(path).resolve() if path else None,
            )
        )
    return result


def same_path(left: Path, right: Path) -> bool:
    try:
        return os.path.normcase(str(left.resolve())) == os.path.normcase(
            str(right.resolve())
        )
    except OSError:
        return False


def ensure_not_running(target: Path) -> None:
    matches = [
        process
        for process in find_running_pcl()
        if process.executable_path and same_path(process.executable_path, target)
    ]
    if matches:
        details = "\n".join(
            f"PID {process.pid}：{process.executable_path}"
            for process in matches
        )
        raise RollbackError(
            "目标启动器仍在运行，请先手动退出后重试。\n\n" + details
        )


def backup_exe(current: ExeInfo) -> Path:
    backup_dir = current.path.parent / "PCL-Rollback-Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    version = re.sub(r"[^0-9A-Za-z._-]+", "_", current.file_version or "unknown")
    backup = (
        backup_dir
        / f"{current.path.stem}.{version}.{stamp}.{time.time_ns()}.exe.bak"
    )
    shutil.copy2(current.path, backup)
    if sha256_file(backup) != current.sha256:
        raise RollbackError("当前 EXE 备份校验失败")
    return backup


def atomic_copy(source: Path, destination: Path) -> None:
    temporary = destination.with_name(
        f".{destination.name}.rollback-{os.getpid()}-{time.time_ns()}.tmp"
    )
    try:
        shutil.copy2(source, temporary)
        with temporary.open("rb+") as stream:
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def update_setup_ini(setup_path: Path) -> tuple[Path | None, bool]:
    setup_path = setup_path.resolve()
    setup_path.parent.mkdir(parents=True, exist_ok=True)
    original = setup_path.read_bytes() if setup_path.exists() else b""
    if original.startswith(b"\xef\xbb\xbf"):
        text, encoding = original.decode("utf-8-sig"), "utf-8-sig"
    elif original.startswith((b"\xff\xfe", b"\xfe\xff")):
        text, encoding = original.decode("utf-16"), "utf-16"
    else:
        text, encoding = "", "utf-8"
        for candidate in ("utf-8", "gb18030"):
            try:
                text = original.decode(candidate)
                encoding = candidate
                break
            except UnicodeDecodeError:
                continue
        if original and not text:
            raise RollbackError("无法识别 Setup.ini 文本编码")

    lines = text.splitlines(keepends=True)
    pattern = re.compile(r"^(\s*)SystemSystemUpdate\s*:(.*?)(\r?\n)?$")
    replaced = 0
    changed = False
    for index, line in enumerate(lines):
        match = pattern.match(line)
        if not match:
            continue
        if match.group(2).strip() != "3":
            changed = True
        lines[index] = f"SystemSystemUpdate:3{match.group(3) or ''}"
        replaced += 1
    if replaced == 0:
        changed = True
        if lines and not lines[-1].endswith(("\n", "\r")):
            lines[-1] += "\r\n"
        lines.append("SystemSystemUpdate:3\r\n")
    if not changed:
        return None, False

    backup = None
    if setup_path.exists():
        backup = setup_path.with_name(
            f"Setup.ini.{time.strftime('%Y%m%d-%H%M%S')}.{time.time_ns()}.bak"
        )
        shutil.copy2(setup_path, backup)
    temporary = setup_path.with_name(
        f".Setup.ini.rollback-{os.getpid()}-{time.time_ns()}.tmp"
    )
    try:
        temporary.write_text("".join(lines), encoding=encoding, newline="")
        os.replace(temporary, setup_path)
    finally:
        temporary.unlink(missing_ok=True)
    return backup, True


def perform_rollback(
    current: ExeInfo,
    target: ExeInfo,
    disable_update: bool,
    log: LogCallback,
) -> dict:
    ensure_not_running(current.path)
    if target.product_name and "Plain Craft Launcher" not in target.product_name:
        raise RollbackError(f"目标产品名不是 PCL：{target.product_name}")
    log(f"当前版本：{current.file_version or '未知'}")
    log(f"目标版本：{target.file_version or '未知'}")
    backup = backup_exe(current)
    log(f"当前 EXE 已备份：{backup}")
    replaced = False
    try:
        atomic_copy(target.path, current.path)
        replaced = True
        after = inspect_exe(current.path)
        if after.sha256 != target.sha256:
            raise RollbackError("替换后的 SHA-256 与目标不一致")
    except Exception:
        if replaced or current.path.exists():
            log("替换失败，正在恢复原 EXE……")
            try:
                atomic_copy(backup, current.path)
                restored = inspect_exe(current.path)
                if restored.sha256 != current.sha256:
                    raise RollbackError("恢复后的 SHA-256 与备份不一致")
            except Exception as restore_error:
                raise RollbackError(
                    "替换失败且自动恢复失败。\n"
                    f"手动备份：{backup}\n恢复错误：{restore_error}"
                ) from restore_error
        raise

    setup_backup = None
    setup_modified = False
    if disable_update:
        setup_path = current.path.parent / "PCL" / "Setup.ini"
        log("将 SystemSystemUpdate 设为 3，关闭自动更新/更新提示……")
        setup_backup, setup_modified = update_setup_ini(setup_path)
        log(
            f"Setup.ini 已修改：{setup_path}"
            if setup_modified
            else f"Setup.ini 已确保为 SystemSystemUpdate:3：{setup_path}"
        )
    final = inspect_exe(current.path)
    return {
        "before": {"version": current.file_version, "sha256": current.sha256},
        "after": {"version": final.file_version, "sha256": final.sha256},
        "exe_backup": str(backup),
        "setup_backup": str(setup_backup) if setup_backup else None,
        "setup_modified": setup_modified,
    }


def assess_target(current: ExeInfo, target: ExeInfo) -> str | None:
    if same_path(current.path, target.path):
        return "目标与当前文件是同一个文件，无需回退。"
    if current.sha256 == target.sha256:
        return "目标 EXE 与当前 EXE 完全相同，替换不会改变版本。"
    if target.product_name and "Plain Craft Launcher" not in target.product_name:
        return f"目标产品名为 {target.product_name}，不是 PCL。"
    return None


class AutoRollbackApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x720")
        self.minsize(820, 620)
        self.events: queue.Queue[tuple] = queue.Queue()
        self.fetch_result: FetchResult | None = None
        self.current_info: ExeInfo | None = None
        self.target_info: ExeInfo | None = None
        self.audit_result: dict | None = None
        self._build_ui()
        self.after(100, self._poll_events)
        self.after(300, self._autodetect)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text="PCL 自动下载与版本回退",
            font=("Microsoft YaHei UI", 17, "bold"),
        ).pack(anchor=tk.W)
        ttk.Label(
            outer,
            text=(
                "单文件工具：检查仓库与标签、下载源码、提取并校验 EXE、"
                "展示完整计划，经确认后备份并原子替换。"
            ),
            wraplength=920,
        ).pack(fill=tk.X, pady=(8, 14))

        form = ttk.Frame(outer)
        form.pack(fill=tk.X)
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="当前 PCL").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5
        )
        self.current_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.current_var).grid(
            row=0, column=1, sticky=tk.EW, pady=5
        )
        ttk.Button(form, text="选择…", command=self._browse_current).grid(
            row=0, column=2, padx=(8, 0), pady=5
        )
        ttk.Label(form, text="下载目录").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5
        )
        self.output_var = tk.StringVar(value=str(Path(__file__).resolve().parent))
        ttk.Entry(form, textvariable=self.output_var).grid(
            row=1, column=1, sticky=tk.EW, pady=5
        )
        ttk.Button(form, text="选择…", command=self._browse_output).grid(
            row=1, column=2, padx=(8, 0), pady=5
        )

        options = ttk.Frame(outer)
        options.pack(fill=tk.X, pady=(10, 12))
        self.mirror_var = tk.BooleanVar(value=False)
        self.disable_update_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options,
            text="允许第三方 GitHub 镜像（不建议）",
            variable=self.mirror_var,
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Checkbutton(
            options,
            text="回退后关闭自动更新（SystemSystemUpdate=3）",
            variable=self.disable_update_var,
        ).grid(row=0, column=1, sticky=tk.W)

        controls = ttk.Frame(outer)
        controls.pack(fill=tk.X)
        self.start_button = ttk.Button(
            controls, text="开始：下载并准备回退", command=self._start_download
        )
        self.start_button.pack(side=tk.LEFT)
        self.execute_button = ttk.Button(
            controls,
            text="确认执行回退",
            command=self._execute,
            state=tk.DISABLED,
        )
        self.execute_button.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(
            controls, text="检查运行中的 PCL", command=self._check_running
        ).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(
            controls, text="导出审计 JSON", command=self._export_audit
        ).pack(side=tk.LEFT, padx=(8, 0))
        self.status_var = tk.StringVar(value="等待开始")
        ttk.Label(controls, textvariable=self.status_var).pack(
            side=tk.LEFT, padx=(16, 0)
        )

        self.progress = ttk.Progressbar(outer, maximum=100)
        self.progress.pack(fill=tk.X, pady=(10, 10))
        plan_box = ttk.LabelFrame(outer, text="计划与校验结果", padding=10)
        plan_box.pack(fill=tk.X, pady=(0, 10))
        self.plan_var = tk.StringVar(
            value="尚未下载目标。下载目录存在时会自动覆盖。"
        )
        ttk.Label(
            plan_box,
            textvariable=self.plan_var,
            justify=tk.LEFT,
            wraplength=900,
        ).pack(fill=tk.X)

        log_frame = ttk.Frame(outer)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Consolas", 10),
            background="#111820",
            foreground="#d7e1ea",
        )
        scrollbar = ttk.Scrollbar(
            log_frame, orient=tk.VERTICAL, command=self.log_text.yview
        )
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        ttk.Label(
            outer,
            text=(
                "不会结束正在运行的 PCL，不会自动启动新版本，"
                "也不会模拟官方更新服务器。"
            ),
            foreground="#8a5a00",
        ).pack(anchor=tk.W, pady=(8, 0))

    def _post(self, kind: str, payload=None) -> None:
        self.events.put((kind, payload))

    def _log(self, message: str) -> None:
        self._post("log", f"[{time.strftime('%H:%M:%S')}] {message}")

    def _progress(self, value: float) -> None:
        self._post("progress", value)

    def _status(self, message: str) -> None:
        self._post("status", message)

    def _autodetect(self) -> None:
        try:
            paths = [
                process.executable_path
                for process in find_running_pcl()
                if process.executable_path
            ]
            if len(paths) == 1 and not self.current_var.get().strip():
                self.current_var.set(str(paths[0]))
                self._log(f"自动识别当前 PCL：{paths[0]}")
        except Exception as error:
            self._log(f"自动识别失败，可手动选择：{error}")

    def _browse_current(self) -> None:
        path = filedialog.askopenfilename(
            title="选择当前 PCL",
            filetypes=[("EXE 文件", "*.exe"), ("所有文件", "*.*")],
        )
        if path:
            self.current_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.askdirectory(title="选择源码下载目录")
        if path:
            self.output_var.set(path)

    def _current_path(self) -> Path:
        text = self.current_var.get().strip()
        if not text:
            raise RollbackError("请选择当前 PCL 启动器。")
        return Path(text)

    def _start_download(self) -> None:
        try:
            current = inspect_exe(self._current_path())
            output = Path(self.output_var.get().strip())
            output.mkdir(parents=True, exist_ok=True)
        except Exception as error:
            messagebox.showerror(APP_TITLE, str(error), parent=self)
            return
        self.current_info = current
        self.target_info = None
        self.fetch_result = None
        self.audit_result = None
        self.execute_button.configure(state=tk.DISABLED)
        self.start_button.configure(state=tk.DISABLED)
        self.progress.configure(value=0)
        self._log(
            f"当前 PCL：{current.path}\n"
            f"实际版本：{current.file_version or '未知'}\n"
            f"SHA-256：{current.sha256}"
        )
        threading.Thread(
            target=self._download_worker,
            args=(output, self.mirror_var.get()),
            daemon=True,
        ).start()

    def _download_worker(self, output: Path, allow_mirror: bool) -> None:
        try:
            result = fetch_and_prepare(
                output,
                allow_mirror,
                self._log,
                self._progress,
                self._status,
            )
            self._post("download_done", result)
        except NetworkFailure as error:
            self._post("download_error", f"网络问题：{error}")
        except (FetchError, RollbackError) as error:
            self._post("download_error", str(error))
        except Exception as error:
            self._post("download_error", f"{type(error).__name__}: {error}")

    def _on_download_done(self, result: FetchResult) -> None:
        self.start_button.configure(state=tk.NORMAL)
        self.fetch_result = result
        try:
            self.current_info = inspect_exe(self._current_path())
            self.target_info = inspect_exe(result.exe_path)
        except Exception as error:
            messagebox.showerror(APP_TITLE, str(error), parent=self)
            return
        blocking = assess_target(self.current_info, self.target_info)
        plan = [
            f"当前版本：{self.current_info.file_version or '未知'}",
            f"当前路径：{self.current_info.path}",
            f"当前 SHA-256：{self.current_info.sha256}",
            "",
            f"目标版本：{self.target_info.file_version or '未知'}",
            f"目标路径：{self.target_info.path}",
            f"目标 SHA-256：{self.target_info.sha256}",
            "",
            f"下载来源：{result.repository}",
            f"源码 URL：{result.source_url}",
            "",
            "自动更新设置："
            + (
                "关闭（SystemSystemUpdate=3）"
                if self.disable_update_var.get()
                else "不修改"
            ),
        ]
        if blocking:
            plan.extend(["", f"阻止执行：{blocking}"])
        self.plan_var.set("\n".join(plan))
        self._log("完整计划：\n" + "\n".join(plan))
        if blocking:
            self.status_var.set("目标校验未通过")
            messagebox.showwarning(APP_TITLE, blocking, parent=self)
            return
        self.status_var.set("下载和校验完成，等待确认")
        self.execute_button.configure(state=tk.NORMAL)
        messagebox.showinfo(
            APP_TITLE,
            "下载和校验完成。\n\n"
            f"当前版本：{self.current_info.file_version or '未知'}\n"
            f"目标版本：{self.target_info.file_version or '未知'}\n\n"
            "完整路径、哈希和配置计划已显示，请检查后点击“确认执行回退”。",
            parent=self,
        )

    def _execute(self) -> None:
        if not self.current_info or not self.target_info:
            return
        blocking = assess_target(self.current_info, self.target_info)
        if blocking:
            messagebox.showerror(APP_TITLE, blocking, parent=self)
            return
        confirm = messagebox.askyesno(
            "最终确认",
            "即将执行：\n\n"
            f"1. 备份当前 EXE：{self.current_info.path}\n"
            f"2. 用目标 EXE 替换：{self.target_info.path}\n"
            "3. 替换后校验 SHA-256\n"
            "4. 自动更新设置："
            f"{'关闭（SystemSystemUpdate:3）' if self.disable_update_var.get() else '不修改'}\n\n"
            "PCL 必须已关闭，且不会自动启动新版本。\n\n是否继续？",
            parent=self,
        )
        if not confirm:
            return
        self.start_button.configure(state=tk.DISABLED)
        self.execute_button.configure(state=tk.DISABLED)
        threading.Thread(target=self._rollback_worker, daemon=True).start()

    def _rollback_worker(self) -> None:
        try:
            result = perform_rollback(
                self.current_info,
                self.target_info,
                self.disable_update_var.get(),
                self._log,
            )
            if self.fetch_result:
                result["download"] = {
                    "repository": self.fetch_result.repository,
                    "source_url": self.fetch_result.source_url,
                    "source_sha256": self.fetch_result.source_sha256,
                    "requested_tag": TAG,
                }
            self.audit_result = result
            self._post("rollback_done", result)
        except Exception as error:
            self._post("rollback_error", str(error))

    def _check_running(self) -> None:
        try:
            running = find_running_pcl()
            text = (
                "没有检测到正在运行的 PCL。"
                if not running
                else "\n".join(
                    f"PID {item.pid}：{item.executable_path or '路径不可读'}"
                    for item in running
                )
            )
        except Exception as error:
            text = f"检查失败：{error}"
        self._log(text)
        messagebox.showinfo(APP_TITLE, text, parent=self)

    def _export_audit(self) -> None:
        payload = {
            "generated_at_utc": time.strftime(
                "%Y-%m-%dT%H:%M:%SZ", time.gmtime()
            ),
            "tool": APP_TITLE,
            "requested_tag": TAG,
            "current": (
                {
                    "path": str(self.current_info.path),
                    "version": self.current_info.file_version,
                    "sha256": self.current_info.sha256,
                }
                if self.current_info
                else None
            ),
            "target": (
                {
                    "path": str(self.target_info.path),
                    "version": self.target_info.file_version,
                    "sha256": self.target_info.sha256,
                }
                if self.target_info
                else None
            ),
            "download": (
                {
                    "repository": self.fetch_result.repository,
                    "source_url": self.fetch_result.source_url,
                    "source_sha256": self.fetch_result.source_sha256,
                }
                if self.fetch_result
                else None
            ),
            "disable_auto_update_requested": self.disable_update_var.get(),
            "rollback_result": self.audit_result,
            "log": self.log_text.get("1.0", tk.END).strip(),
        }
        path = filedialog.asksaveasfilename(
            title="导出审计 JSON",
            defaultextension=".json",
            initialfile=time.strftime("pcl-rollback-audit-%Y%m%d-%H%M%S.json"),
            filetypes=[("JSON 文件", "*.json")],
        )
        if not path:
            return
        Path(path).write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        messagebox.showinfo(APP_TITLE, f"审计记录已导出：\n{path}", parent=self)

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log":
                    self.log_text.configure(state=tk.NORMAL)
                    self.log_text.insert(tk.END, f"{payload}\n")
                    self.log_text.see(tk.END)
                    self.log_text.configure(state=tk.DISABLED)
                elif kind == "status":
                    self.status_var.set(payload)
                elif kind == "progress":
                    self.progress.configure(value=payload * 100)
                elif kind == "download_done":
                    self._on_download_done(payload)
                elif kind == "download_error":
                    self.start_button.configure(state=tk.NORMAL)
                    self.execute_button.configure(state=tk.DISABLED)
                    self.status_var.set("下载失败")
                    self._log(f"下载失败：{payload}")
                    messagebox.showerror(APP_TITLE, str(payload), parent=self)
                elif kind == "rollback_done":
                    self.start_button.configure(state=tk.NORMAL)
                    self.execute_button.configure(state=tk.NORMAL)
                    self.status_var.set("回退完成")
                    messagebox.showinfo(
                        APP_TITLE,
                        "版本回退完成。\n\n"
                        f"当前版本：{payload['after']['version'] or '未知'}\n"
                        f"EXE 备份：{payload['exe_backup']}\n"
                        f"Setup.ini 备份：{payload['setup_backup'] or '未修改'}",
                        parent=self,
                    )
                elif kind == "rollback_error":
                    self.start_button.configure(state=tk.NORMAL)
                    self.execute_button.configure(state=tk.NORMAL)
                    self.status_var.set("回退失败")
                    self._log(f"回退失败：{payload}")
                    messagebox.showerror(APP_TITLE, str(payload), parent=self)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)


def main() -> int:
    app = AutoRollbackApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
