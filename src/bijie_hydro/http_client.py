from __future__ import annotations

import json
import random
import re
import time
from dataclasses import dataclass
from urllib.error import HTTPError, URLError
from urllib.parse import urljoin, urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

from .errors import AccessRestrictedError, SourceError


@dataclass(slots=True, frozen=True)
class HttpResponse:
    url: str
    status: int
    body: bytes
    attempts: int
    content_type: str


def validate_source_url(url: str, page_url: str) -> str:
    source, page = urlsplit(url), urlsplit(page_url)
    if source.scheme not in {"http", "https"} or source.username or source.password or (source.hostname, source.port) != (page.hostname, page.port):
        raise SourceError(f"拒绝跨来源请求: {url}")
    return url


class SameOriginRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        validate_source_url(newurl, req.full_url)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


class HttpClient:
    def __init__(self, timeout: int, max_attempts: int, retry_base_seconds: float, user_agent: str):
        self.timeout = timeout
        self.max_attempts = max_attempts
        self.retry_base_seconds = retry_base_seconds
        self.user_agent = user_agent
        self.total_attempts = 0
        self.opener = build_opener(SameOriginRedirect())

    def request(self, url: str, method: str = "GET", body: bytes | None = None, referer: str | None = None) -> HttpResponse:
        last_error: Exception | None = None
        for attempt in range(1, self.max_attempts + 1):
            self.total_attempts += 1
            headers = {
                "User-Agent": self.user_agent,
                "Accept": "application/json,text/html;q=0.9,*/*;q=0.8",
                "Connection": "close",
            }
            if referer:
                headers["Referer"] = referer
            if method == "POST":
                headers["Content-Type"] = "application/json"
                headers["Origin"] = urljoin(url, "/").rstrip("/")
            request = Request(url=url, data=body, headers=headers, method=method)
            try:
                with self.opener.open(request, timeout=self.timeout) as response:
                    response_body = response.read(20 * 1024 * 1024 + 1)
                    if len(response_body) > 20 * 1024 * 1024:
                        raise SourceError("响应超过20MB上限，停止采集")
                    return HttpResponse(
                        url=response.geturl(),
                        status=response.status,
                        body=response_body,
                        attempts=attempt,
                        content_type=response.headers.get("Content-Type", ""),
                    )
            except (HTTPError, URLError, TimeoutError, ConnectionError) as exc:
                last_error = exc
                if isinstance(exc, HTTPError) and exc.code == 429:
                    raise AccessRestrictedError(f"官网返回429，停止本批次；Retry-After={exc.headers.get('Retry-After', '未提供')}") from exc
                if isinstance(exc, HTTPError) and exc.code in {401, 403}:
                    raise AccessRestrictedError(f"官网拒绝访问({exc.code})，停止采集并人工检查授权/连通情况") from exc
                if isinstance(exc, HTTPError) and 400 <= exc.code < 500 and exc.code not in (408, 429):
                    break
                if attempt < self.max_attempts:
                    delay = self.retry_base_seconds * (2 ** (attempt - 1)) + random.uniform(0, 0.5)
                    time.sleep(delay)
        raise SourceError(f"请求失败: {url}; {last_error}") from last_error


_ENDPOINT_PATTERN = re.compile(r"url\s*:\s*['\"](?P<url>/[^'\"]+)['\"]", re.IGNORECASE)


def discover_endpoint(page_url: str, html: str) -> str:
    candidates = [match.group("url") for match in _ENDPOINT_PATTERN.finditer(html)]
    candidates = [item for item in candidates if not item.lower().endswith((".js", ".css", ".html"))]
    candidates = sorted(set(candidates))
    if len(candidates) != 1:
        raise SourceError(f"无法唯一确定大型水库数据接口，候选数={len(candidates)}")
    return validate_source_url(urljoin(page_url, candidates[0]), page_url)


def parse_api_payload(body: bytes) -> tuple[list[dict], dict]:
    try:
        payload = json.loads(body.decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise SourceError(f"接口响应不是有效UTF-8 JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise SourceError("接口响应顶层结构不是JSON对象")
    if payload.get("returncode") != 0:
        raise SourceError(f"接口业务状态异常: returncode={payload.get('returncode')}, message={payload.get('message')}")
    result = payload.get("result")
    if not isinstance(result, list):
        raise SourceError("接口响应缺少result数组")
    if any(not isinstance(item, dict) for item in result):
        raise SourceError("接口result包含非对象记录")
    return result, payload

