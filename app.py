# -*- coding: utf-8 -*-
"""Форма генерации договоров. Запуск:  streamlit run app.py
Облачный режим: секреты YANDEX_DISK_TOKEN и APP_PASSWORD (см. README)."""
import datetime, io, os, shutil, tempfile, zipfile
import streamlit as st
import core, ydisk

st.set_page_config(page_title="Генератор договоров", page_icon="📄", layout="wide")

KEEP_KEYS = {"auth_ok", "yd_token", "yd_status", "yd_error", "user_name",
             "sets", "set_pick", "результат", "db_synced"}


def get_secret(name, default=""):
    try:
        v = st.secrets.get(name)
        if v:
            return str(v)
    except Exception:
        pass
    return os.environ.get(name, default)


def clear_form():
    """Очистка всех полей формы после генерации."""
    for k in list(st.session_state.keys()):
        if k not in KEEP_KEYS:
            st.session_state.pop(k, None)


# ---------------------------------------------------------------- пароль
_pw = get_secret("APP_PASSWORD")
if _pw and not st.session_state.get("auth_ok"):
    st.title("🔒 Генератор договоров")
    if st.text_input("Пароль", type="password") == _pw:
        st.session_state["auth_ok"] = True
        st.rerun()
    st.stop()

st.title("📄 Генератор договоров")

# ---------------------------------------------------------------- Яндекс Диск
def get_disk():
    token = get_secret("YANDEX_DISK_TOKEN") or st.session_state.get("yd_token", "")
    if not token:
        return None
    return ydisk.YDisk(token, get_secret("YANDEX_DISK_FOLDER",
                                         "disk:/Генератор договоров"))


disk = get_disk()
if disk and "yd_status" not in st.session_state:
    try:
        st.session_state["yd_status"] = disk.check()
        disk.bootstrap(core.TPL_DIR)
        data = disk.download(disk.base + "/контрагенты.xlsx")
        if data:
            with open(core.DB_PATH, "wb") as f:
                f.write(data)
        st.session_state["db_synced"] = True
    except Exception as e:
        st.session_state["yd_status"] = None
        st.session_state["yd_error"] = str(e)

yd_ok = bool(disk and st.session_state.get("yd_status"))


def pull_db():
    """Свежая база с Диска перед записью — сводит конфликты команды к секундам."""
    if yd_ok:
        try:
            data = disk.download(disk.base + "/контрагенты.xlsx")
            if data:
                with open(core.DB_PATH, "wb") as f:
                    f.write(data)
        except Exception:
            pass


def push_db():
    if yd_ok and os.path.exists(core.DB_PATH):
        try:
            disk.upload_file(core.DB_PATH, disk.base + "/контрагенты.xlsx")
        except Exception as e:
            st.warning(f"База сохранена локально, но не загрузилась на Диск: {e}")


# ---------------------------------------------------------------- комплекты
def load_sets(force=False):
    if not force and st.session_state.get("sets"):
        return st.session_state["sets"]
    sets = {}
    if yd_ok:
        try:
            for name in disk.list_sets():
                files = [f for f in disk.list_set_files(name) if f.endswith(".docx")]
                if not files:
                    continue
                raw = disk.download(f"{disk.base}/Шаблоны/{name}/{core.MANIFEST}")
                m = core.read_manifest(raw, name)
                m["files"] = files
                sets[name] = m
        except Exception as e:
            st.warning(f"Не удалось прочитать комплекты с Диска ({e}) — "
                       f"использую встроенные.")
    if not sets:
        for name, files in core.local_sets().items():
            raw = None
            mp = os.path.join(core.TPL_DIR, name, core.MANIFEST)
            if os.path.exists(mp):
                raw = open(mp, "rb").read()
            m = core.read_manifest(raw, name)
            m["files"] = [f for f in files if f.endswith(".docx")]
            sets[name] = m
    st.session_state["sets"] = sets
    return sets


def fetch_templates(set_name, files):
    out = []
    for f in files:
        if yd_ok:
            data = disk.fetch(set_name, f, core.TPL_DIR)
        else:
            with open(os.path.join(core.TPL_DIR, set_name, f), "rb") as fh:
                data = fh.read()
        out.append((f, data))
    return out


sets = load_sets()

# ---------------------------------------------------------------- сайдбар
with st.sidebar:
    st.text_input("👤 Ваше имя (для журнала и истории)", key="user_name",
                  placeholder="например, Олег")
    st.divider()
    st.header("☁️ Яндекс Диск")
    if disk is None:
        st.caption("Без Диска всё работает локально — пакет скачивается zip-архивом.")
        t = st.text_input("OAuth-токен Диска", type="password")
        if t:
            st.session_state["yd_token"] = t
            st.rerun()
    elif yd_ok:
        st.success(f"Подключено: {st.session_state['yd_status']}")
        st.caption(f"Папка: {disk.base.replace('disk:', '') or '/'}")
    else:
        st.error(f"Диск недоступен: {st.session_state.get('yd_error', '')}")
        if st.button("Повторить подключение"):
            st.session_state.pop("yd_status", None)
            st.rerun()
    if st.button("🔄 Обновить список шаблонов"):
        st.session_state.pop("sets", None)
        st.rerun()

# ---------------------------------------------------------------- результат
res = st.session_state.get("результат")
if res:
    st.success(res["text"])
    for w in res.get("warnings", []):
        st.warning(w)
    if res.get("url"):
        st.markdown(f"📂 Пакет на Яндекс Диске: [{res['folder']}]({res['url']})")
    elif res.get("folder"):
        st.markdown(f"📂 Пакет на Яндекс Диске: **{res['folder']}**")
    st.download_button("⬇️ Скачать пакет (zip)", res["zip"],
                       file_name=res["zipname"], mime="application/zip",
                       use_container_width=True, key="dl_result")
    if st.button("✖ Скрыть", key="hide_result"):
        st.session_state.pop("результат", None)
        st.rerun()
    st.divider()

if not sets:
    st.error("Не найдено ни одного комплекта шаблонов (папки в «Шаблоны/»).")
    st.stop()

# ---------------------------------------------------------------- комплект
c_set, c_test = st.columns([4, 1])
set_name = c_set.selectbox("📁 Комплект шаблонов", list(sets.keys()), key="set_pick")
m = sets[set_name]
исп_тип = m["тип_исполнителя"]
префикс = m["префикс"]
c_set.caption(f"Исполнитель: {исп_тип} · нумерация «{префикс}-N»"
              + (" · с НДС" if m["ндс"] else ""))
if c_test.button("🧪 Проверить\nкомплект", use_container_width=True):
    td = core.test_data(исп_тип)
    td["договор_номер"] = префикс + "-ТЕСТ"
    rows = []
    for fname, data in fetch_templates(set_name, m["files"]):
        try:
            core.render_docx_bytes(data, core.build_context(td))
            rows.append(f"✅ {fname}")
        except Exception as e:
            rows.append(f"❌ {fname} — {e}")
    (st.success if all(r.startswith("✅") for r in rows) else st.error)(
        "\n\n".join(rows))

# ---------------------------------------------------------------- стороны
parties = core.load_parties()
labels = [f'{p["наименование"]} · ИНН {p["инн"] or "—"} · {p["тип"]}' for p in parties]

PARTY_FIELDS = [
    ("наименование", "Наименование / ФИО", "ООО «Ромашка» или Иванов Иван Иванович"),
    ("инн", "ИНН", ""), ("кпп", "КПП", ""), ("огрн", "ОГРН / ОГРНИП", ""),
    ("адрес", "Адрес", ""),
    ("должность_род", "Должность подписанта (род. падеж)", "Генерального директора"),
    ("подписант_род", "ФИО подписанта (род. падеж)", "Петрова Петра Петровича"),
    ("подписант", "ФИО подписанта (им. падеж)", "Петров Петр Петрович"),
    ("основание_род", "Действует на основании (род. падеж)", "Устава / доверенности №…"),
    ("паспорт", "Паспорт", "серия, номер, кем и когда выдан"),
    ("счет", "Расчетный счет", ""), ("банк", "Банк", ""),
    ("бик", "БИК", ""), ("корсчет", "Корр. счет", ""),
    ("телефон", "Телефон", ""), ("email", "E-mail", ""),
]
FIELDS_BY_KIND = {  # какие поля показывать
    "ЮЛ": [k for k, _, _ in PARTY_FIELDS if k != "паспорт"],
    "ИП": ["наименование", "инн", "огрн", "адрес", "счет", "банк", "бик",
           "корсчет", "телефон", "email"],
    "Самозанятый": ["наименование", "инн", "паспорт", "адрес", "счет", "банк",
                    "бик", "корсчет", "телефон", "email"],
}


def _fill_from_db(role):
    pick = st.session_state.get(f"{role}_pick")
    if pick in labels:
        p = parties[labels.index(pick)]
        for k, v in p.items():
            st.session_state[f"{role}_{k}"] = v
        st.session_state[f"{role}_snapshot"] = dict(p)


def party_form(role, title, kind):
    st.subheader(title)
    st.selectbox("📇 Подставить сохранённого контрагента",
                 ["— ввести вручную —"] + labels, key=f"{role}_pick",
                 on_change=_fill_from_db, args=(role,))
    snap = st.session_state.get(f"{role}_snapshot")
    if snap and role == "i" and snap.get("тип") and snap["тип"] != kind:
        st.warning(f"Выбранный контрагент сохранён как «{snap['тип']}», "
                   f"а комплект — «{kind}». Проверьте, тот ли это исполнитель.")
    data = {"тип": kind if role == "i" else
            st.session_state.get(f"{role}_тип") or "ЮЛ"}
    show = FIELDS_BY_KIND[kind] if role == "i" else FIELDS_BY_KIND["ЮЛ"]
    for k, label, hint in PARTY_FIELDS:
        if k in show:
            data[k] = st.text_input(label, key=f"{role}_{k}", placeholder=hint)
    cur = core.find_party(data)
    if cur:
        diff = core.party_diff(cur, data)
        if diff:
            st.caption("✏️ Отличается от записи в базе: " +
                       ", ".join(k for k, _, _ in diff))
    return data


col_z, col_i = st.columns(2)
with col_z:
    заказчик = party_form("z", "Заказчик", "ЮЛ")
with col_i:
    исполнитель = party_form("i", f"Исполнитель ({исп_тип})", исп_тип)

# ---------------------------------------------------------------- документ
st.subheader("Документ")
suggest = core.next_contract_number(префикс)
c1, c2, c3, c4, c5 = st.columns(5)
номер = c1.text_input("Номер", value=suggest, key=f"номер_{префикс}")
дата = c2.date_input("Дата", value=datetime.date.today(), format="DD.MM.YYYY",
                     key="дата")
город = c3.text_input("Город", value="Екатеринбург", key="город")
начало = c4.date_input("Работы: начало", value=datetime.date.today(),
                       format="DD.MM.YYYY", key="начало")
окончание = c5.date_input("Работы: окончание",
                          value=datetime.date.today() + datetime.timedelta(days=30),
                          format="DD.MM.YYYY", key="окончание")

st.subheader("Услуги / работы")
st.caption("Сумма строки = количество × цена, итог считается сам. Даты заполнять "
           "не обязательно.")
услуги_df = st.data_editor(
    [{"наименование": "", "колво": 1, "цена": 0.0, "начало": None,
      "окончание": None}],
    num_rows="dynamic", use_container_width=True, key="услуги",
    column_config={
        "наименование": st.column_config.TextColumn("Наименование", width="large"),
        "колво": st.column_config.NumberColumn("Кол-во", min_value=0, format="%g"),
        "цена": st.column_config.NumberColumn("Цена, руб.", min_value=0.0,
                                              format="%.2f"),
        "начало": st.column_config.DateColumn("Начало", format="DD.MM.YYYY"),
        "окончание": st.column_config.DateColumn("Окончание", format="DD.MM.YYYY"),
    })

# ---------------------------------------------------------------- условия
st.subheader("Условия")
u1, u2 = st.columns(2)
блок_ис = u1.checkbox("Пункт об интеллектуальной собственности", True,
                      key="блок_ис",
                      help="Снимите, если объектов ИС в работах не возникает")
ндс_ставка = 0
if m["ндс"]:
    ндс_ставка = u2.number_input("Ставка НДС, %", 0, 30, 20, key="ндс_ставка")

with st.expander("Дополнительные поля (приложение, счёт-оферта)"):
    t1, t2 = st.columns(2)
    прил_номер = t1.text_input("№ приложения к договору", value="1", key="прил_номер")
    прил_дата = t2.date_input("Дата приложения", value=datetime.date.today(),
                              format="DD.MM.YYYY", key="прил_дата")
    оферта_оплата = st.text_input(
        "Срок оплаты аванса", key="оферта_оплата",
        value="в течение 5 (пяти) рабочих дней с даты выставления Счета")
    оферта_срок = st.text_input(
        "Срок выполнения работ", key="оферта_срок",
        value="в течение 10 (десяти) рабочих дней с даты внесения аванса")
    оферта_результат = st.text_input("Результат работ", key="оферта_результат",
                                     value="результат работ, указанных в Счете")
    оферта_формат = st.text_input("Формат передачи результата", key="оферта_формат",
                                  value="ссылкой на облачное хранилище")
    ндс_строка = st.text_input("Строка НДС (пусто = автоматически)", value="",
                               key="ндс_строка")

st.subheader("Какие документы сгенерировать")
dcols = st.columns(max(len(m["files"]), 1))
выбранные = []
for i, fname in enumerate(m["files"]):
    if dcols[i].checkbox(os.path.splitext(fname)[0], True,
                         key=f"doc_{set_name}_{fname}"):
        выбранные.append(fname)

сохранить = st.checkbox("Сохранить контрагентов в базу и записать в журнал", True,
                        key="сохранить")


# ---------------------------------------------------------------- генерация
def collect_data():
    услуги = []
    for u in услуги_df:
        if not str(u.get("наименование") or "").strip():
            continue
        услуги.append({k: (None if v is None or str(v) in ("NaT", "nan") else v)
                       for k, v in u.items()})
    data = {
        "договор_номер": номер.strip(), "договор_дата": дата, "город": город,
        "услуги_начало": начало, "услуги_окончание": окончание,
        "заказчик": заказчик, "исполнитель": исполнитель, "услуги": услуги,
        "блок_ис": блок_ис, "ндс_ставка": ндс_ставка,
        "прил_номер": прил_номер, "прил_дата": прил_дата,
        "оферта_срок_оплаты": оферта_оплата, "оферта_срок_работ": оферта_срок,
        "оферта_результат": оферта_результат, "оферта_формат": оферта_формат,
        "ндс_строка": ндс_строка or None,
    }
    return data, услуги


def run_generation(choices):
    """choices: {'z'|'i': 'update'|'skip'} для конфликтных контрагентов."""
    user = st.session_state.get("user_name", "")
    data, услуги = collect_data()
    pull_db()
    if data["договор_номер"] == suggest:  # номер не правили — пересчитать свежий
        data["договор_номер"] = core.next_contract_number(префикс)

    warnings = core.validate_party(заказчик, "Заказчик") + \
        core.validate_party(исполнитель, "Исполнитель")
    if data["договор_номер"] in core.journal_numbers():
        warnings.append(f"Номер {data['договор_номер']} уже встречался в журнале.")

    tmp = tempfile.mkdtemp()
    try:
        paths = core.generate_files(data, tmp,
                                    fetch_templates(set_name, выбранные))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for p in paths:
                z.write(p, os.path.basename(p))

        folder_name = url = ""
        if yd_ok:
            folder_name = core.safe_name(
                f"{data['договор_номер']} — {исполнитель['наименование']}")
            folder = f"{disk.base}/Документы/{folder_name}"
            try:
                disk.ensure_path(folder)
                for p in paths:
                    disk.upload_file(p, folder + "/" + os.path.basename(p))
                url = disk.publish(folder)
            except Exception as e:
                warnings.append(f"На Диск загрузить не удалось: {e}. Скачайте zip.")
                folder_name = url = ""

        if сохранить:
            for role, p in [("z", заказчик), ("i", исполнитель)]:
                if choices.get(role) != "skip":
                    core.save_party(p, user=user)
            data["стоимость"] = sum(float(u.get("цена") or 0) *
                                    float(u.get("колво") or 1) for u in услуги)
            core.append_journal(data, [os.path.splitext(f)[0] for f in выбранные],
                                комплект=set_name, user=user)
            push_db()

        st.session_state["результат"] = {
            "text": f"Готово! {data['договор_номер']}: {len(paths)} док., "
                    f"сумма {core.money_fmt(data.get('стоимость') or 0)} руб. "
                    f"Форма очищена для следующего пакета.",
            "warnings": warnings, "folder": folder_name, "url": url,
            "zip": buf.getvalue(),
            "zipname": core.safe_name(
                f"{data['договор_номер']} {исполнитель['наименование']}") + ".zip",
        }
        clear_form()
        st.rerun()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def find_conflicts():
    out = []
    for role, p, who in [("z", заказчик, "Заказчик"), ("i", исполнитель, "Исполнитель")]:
        cur = core.find_party(p)
        if cur:
            diff = core.party_diff(cur, p)
            if diff:
                out.append({"role": role, "who": who,
                            "name": p.get("наименование"), "diff": diff})
    return out


pending = st.session_state.get("pending_conflicts")
if pending:
    st.warning("Данные контрагентов отличаются от сохранённых в базе. "
               "Что сделать с базой?")
    for c in pending:
        with st.container(border=True):
            st.markdown(f"**{c['who']}: {c['name']}**")
            for k, a, b in c["diff"]:
                st.markdown(f"- {k}: ~~{a}~~ → **{b}**")
            st.radio("Действие", ["Обновить запись (старая версия — в Архив)",
                                  "Не сохранять (только в этот документ)"],
                     key=f"conf_{c['role']}", horizontal=True)
    cc1, cc2 = st.columns(2)
    if cc1.button("✅ Подтвердить и сгенерировать", type="primary",
                  use_container_width=True):
        choices = {c["role"]: ("skip" if "Не сохранять" in
                               st.session_state.get(f"conf_{c['role']}", "")
                               else "update") for c in pending}
        st.session_state.pop("pending_conflicts", None)
        try:
            run_generation(choices)
        except Exception as e:
            st.error(str(e))
    if cc2.button("↩️ Отмена", use_container_width=True):
        st.session_state.pop("pending_conflicts", None)
        st.rerun()
else:
    if st.button("🚀 Сгенерировать пакет документов", type="primary",
                 use_container_width=True):
        data, услуги = collect_data()
        if not услуги:
            st.error("Добавьте хотя бы одну услугу/работу с наименованием.")
        elif not заказчик["наименование"] or not исполнитель["наименование"]:
            st.error("Заполните наименования Заказчика и Исполнителя.")
        elif not выбранные:
            st.error("Отметьте хотя бы один документ.")
        else:
            conflicts = find_conflicts() if сохранить else []
            if conflicts:
                st.session_state["pending_conflicts"] = conflicts
                st.rerun()
            else:
                try:
                    run_generation({})
                except Exception as e:
                    st.error(str(e))

# ---------------------------------------------------------------- история
st.divider()
with st.expander("📜 База контрагентов: история изменений и восстановление"):
    if not parties:
        st.caption("База пока пуста.")
    else:
        pick = st.selectbox("Контрагент", labels, key="hist_pick")
        p = parties[labels.index(pick)]
        hist = core.party_history(p)
        if not hist:
            st.caption("Архивных версий нет — запись ни разу не перезаписывалась.")
        else:
            опции = [f'{v["заменено"]} — {v["кем"] or "имя не указано"}'
                     for v in hist]
            vi = st.selectbox("Версия (новые сверху)", опции, key="hist_ver")
            v = hist[опции.index(vi)]
            st.table([{"поле": k, "в этой версии": v.get(k, "") or "—",
                       "сейчас в базе": p.get(k, "") or "—"}
                      for k in core.PARTY_COLS
                      if (v.get(k) or p.get(k)) and v.get(k) != p.get(k)])
            if st.button("↩️ Вернуть эту версию", key="hist_restore"):
                pull_db()
                core.restore_party(v, user=st.session_state.get("user_name", ""))
                push_db()
                st.success("Версия восстановлена; предыдущая ушла в Архив.")
                st.rerun()
