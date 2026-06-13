# -*- coding: utf-8 -*-
"""
ydisk.py — минимальный клиент REST API Яндекс Диска.
Документация API: https://yandex.ru/dev/disk/api/concepts/about.html
Авторизация: OAuth-токен (как получить — см. README, раздел «Облачный режим»).
"""
import os
import requests

API = "https://cloud-api.yandex.net/v1/disk"

# Имя папки внутри templates/, которая считается «единственным комплектом»
# для варианта без выбора (app_single). Должно совпадать с core.SINGLE_SET.
SINGLE_SET = "Самозанятые"


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

    def listdir(self, path):
        """Содержимое папки: список (имя, тип) где тип 'dir' или 'file'.
        None, если папки нет."""
        r = self._req("GET", API + "/resources",
                      params={"path": path, "limit": 500,
                              "fields": "_embedded.items.name,_embedded.items.type"},
                      ok404=True)
        if r is None:
            return None
        items = (r.json().get("_embedded") or {}).get("items", [])
        return [(i["name"], i["type"]) for i in items]

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
    def bootstrap(self, local_tpl_dir, multi=True):
        """Создает структуру папок и докладывает на Диск недостающие шаблоны
        (существующее на Диске НЕ перезаписывается — правки целы).

        multi=True  — комплекты раскладываются по подпапкам Шаблоны/<Комплект>/.
        multi=False — один комплект кладётся плоско в Шаблоны/ (без выбора).
        Источник для single — папка SINGLE_SET внутри local_tpl_dir.
        Папки и файлы, начинающиеся с «_», на Диск не выгружаются.

        Если локальной папки-источника нет (шаблоны не залиты в репозиторий) —
        создаём только структуру папок и выходим: шаблоны в этом случае ведутся
        прямо на Диске, локальный резерв не обязателен."""
        self.ensure_path(self.base + "/Шаблоны")
        self.ensure_path(self.base + "/Документы")
        if not local_tpl_dir or not os.path.isdir(local_tpl_dir):
            return  # нет локального резерва — работаем только с Диском
        if multi:
            for set_name in sorted(os.listdir(local_tpl_dir)):
                local_set = os.path.join(local_tpl_dir, set_name)
                if set_name.startswith("_") or not os.path.isdir(local_set):
                    continue
                remote_set = f"{self.base}/Шаблоны/{set_name}"
                self.mkdir(remote_set)
                have = {n for n, t in (self.listdir(remote_set) or [])}
                for fname in sorted(os.listdir(local_set)):
                    if fname not in have:
                        self.upload_file(os.path.join(local_set, fname),
                                         remote_set + "/" + fname)
        else:
            src = os.path.join(local_tpl_dir, SINGLE_SET)
            if not os.path.isdir(src):
                return
            have = {n for n, t in (self.listdir(self.base + "/Шаблоны") or [])}
            for fname in sorted(os.listdir(src)):
                if fname not in have:
                    self.upload_file(os.path.join(src, fname),
                                     self.base + "/Шаблоны/" + fname)

    def list_flat_files(self):
        """docx прямо в Шаблоны/ (для варианта с одним комплектом)."""
        items = self.listdir(self.base + "/Шаблоны") or []
        return sorted(n for n, t in items if t == "file" and n.endswith(".docx"))

    def fetch_flat(self, filename, fallback_dir):
        data = self.download(f"{self.base}/Шаблоны/{filename}")
        if data is None:
            with open(os.path.join(fallback_dir, SINGLE_SET, filename), "rb") as f:
                data = f.read()
        return data

    def list_sets(self):
        """Имена папок-комплектов на Диске."""
        items = self.listdir(self.base + "/Шаблоны") or []
        return sorted(n for n, t in items if t == "dir" and not n.startswith("_"))

    def list_set_files(self, set_name):
        items = self.listdir(f"{self.base}/Шаблоны/{set_name}") or []
        return sorted(n for n, t in items if t == "file")

    def fetch(self, set_name, filename, fallback_dir):
        """Файл комплекта с Диска, а если его там нет — локальный."""
        data = self.download(f"{self.base}/Шаблоны/{set_name}/{filename}")
        if data is None:
            with open(os.path.join(fallback_dir, set_name, filename), "rb") as f:
                data = f.read()
        return data
