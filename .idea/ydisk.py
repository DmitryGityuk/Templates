# -*- coding: utf-8 -*-
"""
ydisk.py — минимальный клиент REST API Яндекс Диска.
Документация API: https://yandex.ru/dev/disk/api/concepts/about.html
Авторизация: OAuth-токен (как получить — см. README, раздел «Облачный режим»).
"""
import os
import requests

API = "https://cloud-api.yandex.net/v1/disk"


class YDiskError(RuntimeError):
    pass


class YDisk:
    def __init__(self, token, base="disk:/Генератор договоров"):
        self.headers = {"Authorization": "OAuth " + token.strip()}
        self.base = base.rstrip("/")

    # ------------------------------------------------------------- служебное
    def _req(self, method, url, ok404=False, **kw):
        try:
            r = requests.request(method, url, headers=self.headers, timeout=60, **kw)
        except requests.RequestException as e:
            raise YDiskError(f"нет связи с Яндекс Диском: {e}") from e
        if r.status_code == 404 and ok404:
            return None
        if r.status_code >= 400:
            try:
                j = r.json()
                msg = j.get("message") or j.get("description") or r.text[:200]
            except Exception:
                msg = r.text[:200]
            raise YDiskError(f"Диск ответил {r.status_code}: {msg}")
        return r

    # ------------------------------------------------------------- операции
    def check(self):
        """Проверка токена. Возвращает строку «владелец, свободно N ГБ»."""
        j = self._req("GET", API).json()
        free = (j.get("total_space", 0) - j.get("used_space", 0)) / 2 ** 30
        user = (j.get("user") or {}).get("display_name", "")
        return f"{user}, свободно {free:.1f} ГБ".strip(", ")

    def exists(self, path):
        return self._req("GET", API + "/resources",
                         params={"path": path, "fields": "name"}, ok404=True) is not None

    def mkdir(self, path):
        """Создает папку; если уже есть — молча продолжает."""
        try:
            self._req("PUT", API + "/resources", params={"path": path})
        except YDiskError as e:
            if "409" not in str(e):
                raise

    def ensure_path(self, path):
        """Создает всю цепочку папок: disk:/A/B/C или app:/A/B."""
        if ":/" not in path:
            raise YDiskError(f"путь должен начинаться с disk:/ или app:/ — получено {path!r}")
        scheme, rest = path.split(":/", 1)
        cur = scheme + ":"
        for seg in [s for s in rest.split("/") if s]:
            cur += "/" + seg
            self.mkdir(cur)

    def upload_bytes(self, data, path):
        href = self._req("GET", API + "/resources/upload",
                         params={"path": path, "overwrite": "true"}).json()["href"]
        r = requests.put(href, data=data, timeout=300)
        if r.status_code >= 400:
            raise YDiskError(f"не удалось загрузить {path}: HTTP {r.status_code}")

    def upload_file(self, local_path, path):
        with open(local_path, "rb") as f:
            self.upload_bytes(f.read(), path)

    def download(self, path):
        """Возвращает содержимое файла (bytes) или None, если файла нет."""
        r = self._req("GET", API + "/resources/download",
                      params={"path": path}, ok404=True)
        if r is None:
            return None
        href = r.json()["href"]
        r2 = requests.get(href, timeout=300)
        if r2.status_code >= 400:
            raise YDiskError(f"не удалось скачать {path}: HTTP {r2.status_code}")
        return r2.content

    def publish(self, path):
        """Делает ресурс публичным и возвращает ссылку «поделиться»."""
        self._req("PUT", API + "/resources/publish", params={"path": path})
        j = self._req("GET", API + "/resources",
                      params={"path": path, "fields": "public_url"}).json()
        return j.get("public_url", "")

    # ------------------------------------------------------------- сценарии
    def bootstrap(self, local_tpl_dir):
        """Создает структуру папок и докладывает на Диск недостающие шаблоны
        (существующие на Диске НЕ перезаписывает — правки сохраняются)."""
        self.ensure_path(self.base + "/Шаблоны")
        self.ensure_path(self.base + "/Документы")
        for name in sorted(os.listdir(local_tpl_dir)):
            if name.endswith(".docx") and not self.exists(self.base + "/Шаблоны/" + name):
                self.upload_file(os.path.join(local_tpl_dir, name),
                                 self.base + "/Шаблоны/" + name)

    def fetch_template(self, filename, fallback_dir):
        """Шаблон с Диска, а если его там нет — локальный из комплекта."""
        data = self.download(self.base + "/Шаблоны/" + filename)
        if data is None:
            with open(os.path.join(fallback_dir, filename), "rb") as f:
                data = f.read()
        return data
