from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Iterable


class UploadPort(ABC):
    @abstractmethod
    def upload_file(self, file_path: Path, metadata: dict[str, Any]) -> str:
        raise NotImplementedError

    @abstractmethod
    def upload_records(self, records: Iterable[dict[str, Any]], metadata: dict[str, Any]) -> str:
        raise NotImplementedError

    @abstractmethod
    def check_upload_status(self, task_id: str) -> str:
        raise NotImplementedError


class DisabledUploadPort(UploadPort):
    def upload_file(self, file_path: Path, metadata: dict[str, Any]) -> str:
        raise RuntimeError("MVP阶段上传功能未启用")

    def upload_records(self, records: Iterable[dict[str, Any]], metadata: dict[str, Any]) -> str:
        raise RuntimeError("MVP阶段上传功能未启用")

    def check_upload_status(self, task_id: str) -> str:
        raise RuntimeError("MVP阶段上传功能未启用")

