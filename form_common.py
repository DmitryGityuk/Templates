# -*- coding: utf-8 -*-
"""Общая логика формы для обоих вариантов (один комплект / выбор комплекта).
Используется из app_single.py и app_multi.py."""
import datetime, io, os, shutil, tempfile, zipfile
import streamlit as st
import core, ydisk

KEEP_KEYS = {"auth_ok", "yd_token", "yd_status", "yd_error", "user_name",
             "sets", "результат", "db_synced", "_do_clear"}


def get_secret(name, default=""):
    try:
        v = st.secrets.get(name)
        if v:
            return str(v)
    except Exception:
        pass
    return os.environ.get(name, default)


def _apply_pending_clear():
    """Очистка выполняется В НАЧАЛЕ прогона (до создания виджетов), иначе
    Streamlit восстановит значения из value=. Ставим флаг -> rerun -> чистим."""
    if st.session_state.pop("_do_clear", False):
        for k in list(st.session_state.keys()):
            if k not in KEEP_KEYS:
                st.session_state.pop(k, None)


def request_clear():
    st.session_state["_do_clear"] = True


def password_gate():
    pw = get_secret("APP_PASSWORD")
    if pw and not st.session_state.get("auth_ok"):
        st.title("🔒 Генератор договоров")
        if st.text_input("Пароль", type="password") == pw:
            st.session_state["auth_ok"] = True
            st.rerun()
        st.stop()


# ----------------------------------------------------------------- Диск
def get_disk():
    token = get_secret("YANDEX_DISK_TOKEN") or st.session_state.get("yd_token", "")
    if not token:
        return None
    return ydisk.YDisk(token, get_secret("YANDEX_DISK_FOLDER",
                                         "disk:/Генератор договоров"))


def connect_disk(disk, multi):
    """Первое подключение: проверка токена, структура папок, скачивание базы.
    Сбой bootstrap (например, нет локального резерва) не должен ронять связь."""
    if disk and "yd_status" not in st.session_state:
        try:
            st.session_state["yd_status"] = disk.check()
        except Exception as e:
            st.session_state["yd_status"] = None
            st.session_state["yd_error"] = str(e)
            return False
        try:
            disk.bootstrap(core.TPL_DIR, multi=multi)
        except Exception as e:
            st.session_state["yd_error"] = f"шаблоны: {e}"  # не критично
        try:
            data = disk.download(disk.base + "/контрагенты.xlsx")
            if data:
                with open(core.DB_PATH, "wb") as f:
                    f.write(data)
            st.session_state["db_synced"] = True
        except Exception as e:
            st.session_state["yd_error"] = f"база: {e}"
    return bool(disk and st.session_state.get("yd_status"))


def pull_db(disk, yd_ok):
    if yd_ok:
        try:
            data = disk.download(disk.base + "/контрагенты.xlsx")
            if data:
                with open(core.DB_PATH, "wb") as f:
                    f.write(data)
        except Exception:
            pass


def push_db(disk, yd_ok):
    if yd_ok and os.path.exists(core.DB_PATH):
        try:
            disk.upload_file(core.DB_PATH, disk.base + "/контрагенты.xlsx")
        except Exception as e:
            st.warning(f"База сохранена локально, но не загрузилась на Диск: {e}")


def sidebar(disk, yd_ok):
    with st.sidebar:
        st.text_input("👤 Ваше имя (для журнала и истории)", key="user_name",
                      placeholder="например, Олег")
        st.divider()
        st.header("☁️ Яндекс Диск")
        if disk is None:
            st.caption("Без Диска всё работает локально — пакет скачивается zip.")
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


def show_result_panel():
    res = st.session_state.get("результат")
    if not res:
        return
    st.success(res["text"])
    for w in res.get("warnings", []):
        st.warning(w)
    if res.get("url"):
        st.markdown(f"📂 Пакет на Яндекс Диске: [{res['folder']}]({res['url']})")
    elif res.get("folder"):
        st.markdown(f"📂 Пакет на Яндекс Диске: **{res['folder']}**")
    st.download_button("⬇️ Скачать пакет (zip)", res["zip"], file_name=res["zipname"],
                       mime="application/zip", use_container_width=True,
                       key="dl_result")
    if st.button("✖ Скрыть", key="hide_result"):
        st.session_state.pop("результат", None)
        st.rerun()
    st.divider()


# ----------------------------------------------------------------- стороны
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
FIELDS_BY_KIND = {
    "ЮЛ": [k for k, _, _ in PARTY_FIELDS if k != "паспорт"],
    "ИП": ["наименование", "инн", "огрн", "адрес", "счет", "банк", "бик",
           "корсчет", "телефон", "email"],
    "Самозанятый": ["наименование", "инн", "паспорт", "адрес", "счет", "банк",
                    "бик", "корсчет", "телефон", "email"],
}


def make_filler(parties, labels):
    def _fill(role):
        pick = st.session_state.get(f"{role}_pick")
        if pick in labels:
            p = parties[labels.index(pick)]
            for k, v in p.items():
                st.session_state[f"{role}_{k}"] = v
            st.session_state[f"{role}_snapshot"] = dict(p)   # что подставили из базы
        else:
            st.session_state.pop(f"{role}_snapshot", None)
    return _fill


def party_form(role, title, kind, parties, labels, filler):
    st.subheader(title)
    st.selectbox("📇 Подставить сохранённого контрагента",
                 ["— ввести вручную —"] + labels, key=f"{role}_pick",
                 on_change=filler, args=(role,))
    snap = st.session_state.get(f"{role}_snapshot")
    if snap and role == "i" and snap.get("тип") and snap["тип"] != kind \
            and kind in core._STATUS_LABEL:
        st.warning(f"Контрагент сохранён как «{snap['тип']}», а комплект — "
                   f"«{kind}». Проверьте, тот ли это исполнитель.")
    data = {"тип": kind if role == "i" else
            (snap.get("тип") if snap else "ЮЛ")}
    show = FIELDS_BY_KIND.get(kind, FIELDS_BY_KIND["ЮЛ"]) if role == "i" \
        else FIELDS_BY_KIND["ЮЛ"]
    for k, label, hint in PARTY_FIELDS:
        if k in show:
            data[k] = st.text_input(label, key=f"{role}_{k}", placeholder=hint)
    # подсветка отличий от ПОДСТАВЛЕННОЙ записи (а не от любой по ИНН)
    base = snap if (snap and core._same_party(snap, data)) else None
    if base:
        diff = core.party_diff(base, data)
        if diff:
            st.caption("✏️ Отличается от выбранной записи: " +
                       ", ".join(k for k, _, _ in diff))
    return data


# ----------------------------------------------------------------- общие поля
def document_fields(with_vat):
    st.subheader("Документ")
    c1, c2, c3, c4, c5 = st.columns(5)
    номер = c1.text_input("Номер", key="номер", placeholder="например, 12 или 7/2026")
    дата = c2.date_input("Дата", value=datetime.date.today(), format="DD.MM.YYYY",
                         key="дата")
    город = c3.text_input("Город", value="Екатеринбург", key="город")
    начало = c4.date_input("Работы: начало", value=datetime.date.today(),
                           format="DD.MM.YYYY", key="начало")
    окончание = c5.date_input("Работы: окончание",
                              value=datetime.date.today() + datetime.timedelta(days=30),
                              format="DD.MM.YYYY", key="окончание")

    st.subheader("Услуги / работы")
    st.caption("Сумма строки = количество × цена, итог считается сам. "
               "Даты заполнять не обязательно.")
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

    st.subheader("Условия")
    u1, u2 = st.columns(2)
    блок_ис = u1.checkbox("Пункт об интеллектуальной собственности", True,
                          key="блок_ис",
                          help="Снимите, если объектов ИС в работах не возникает")
    ндс_ставка = 0
    if with_vat:
        ндс_ставка = u2.number_input("Ставка НДС, %", 0, 30, 20, key="ндс_ставка")

    with st.expander("Дополнительные поля (приложение, счёт-оферта)"):
        t1, t2 = st.columns(2)
        прил_номер = t1.text_input("№ приложения к договору", value="1",
                                   key="прил_номер")
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
    return {
        "номер": номер, "дата": дата, "город": город, "начало": начало,
        "окончание": окончание, "услуги_df": услуги_df, "блок_ис": блок_ис,
        "ндс_ставка": ндс_ставка, "прил_номер": прил_номер, "прил_дата": прил_дата,
        "оферта_оплата": оферта_оплата, "оферта_срок": оферта_срок,
        "оферта_результат": оферта_результат, "оферта_формат": оферта_формат,
        "ндс_строка": ндс_строка,
    }


def assemble(f, заказчик, исполнитель):
    услуги = []
    for u in f["услуги_df"]:
        if not str(u.get("наименование") or "").strip():
            continue
        услуги.append({k: (None if v is None or str(v) in ("NaT", "nan") else v)
                       for k, v in u.items()})
    data = {
        "договор_номер": (f["номер"] or "").strip(), "договор_дата": f["дата"],
        "город": f["город"], "услуги_начало": f["начало"],
        "услуги_окончание": f["окончание"], "заказчик": заказчик,
        "исполнитель": исполнитель, "услуги": услуги, "блок_ис": f["блок_ис"],
        "ндс_ставка": f["ндс_ставка"], "прил_номер": f["прил_номер"],
        "прил_дата": f["прил_дата"], "оферта_срок_оплаты": f["оферта_оплата"],
        "оферта_срок_работ": f["оферта_срок"], "оферта_результат": f["оферта_результат"],
        "оферта_формат": f["оферта_формат"], "ндс_строка": f["ндс_строка"] or None,
    }
    return data, услуги


# ----------------------------------------------------------------- генерация
def generate_and_store(data, услуги, templates, set_name, disk, yd_ok, choices,
                       заказчик, исполнитель):
    """Единая процедура: рендер -> zip -> (Диск) -> база/журнал -> результат."""
    user = st.session_state.get("user_name", "")
    pull_db(disk, yd_ok)

    warnings = core.validate_party(заказчик, "Заказчик") + \
        core.validate_party(исполнитель, "Исполнитель")
    if data["договор_номер"] and data["договор_номер"] in core.journal_numbers():
        warnings.append(f"Номер {data['договор_номер']} уже встречался в журнале.")

    tmp = tempfile.mkdtemp()
    try:
        paths = core.generate_files(data, tmp, templates)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for p in paths:
                z.write(p, os.path.basename(p))

        folder_name = url = ""
        if yd_ok:
            ном = data["договор_номер"] or datetime.date.today().strftime("%Y-%m-%d")
            folder_name = core.safe_name(f"{ном} — {исполнитель['наименование']}")
            folder = f"{disk.base}/Документы/{folder_name}"
            try:
                disk.ensure_path(folder)
                for p in paths:
                    disk.upload_file(p, folder + "/" + os.path.basename(p))
                url = disk.publish(folder)
            except Exception as e:
                warnings.append(f"На Диск загрузить не удалось: {e}. Скачайте zip.")
                folder_name = url = ""

        if st.session_state.get("сохранить", True):
            for role, p in [("z", заказчик), ("i", исполнитель)]:
                if choices.get(role) != "skip":
                    core.save_party(p, user=user)
            data["стоимость"] = sum(float(u.get("цена") or 0) *
                                    float(u.get("колво") or 1) for u in услуги)
            core.append_journal(data, [os.path.splitext(f)[0] for f, _ in templates],
                                комплект=set_name, user=user)
            push_db(disk, yd_ok)

        ном = data["договор_номер"] or "без номера"
        st.session_state["результат"] = {
            "text": f"Готово! {ном}: {len(paths)} док., сумма "
                    f"{core.money_fmt(data.get('стоимость') or 0)} руб. "
                    f"Форма очищена для следующего пакета.",
            "warnings": warnings, "folder": folder_name, "url": url,
            "zip": buf.getvalue(),
            "zipname": core.safe_name(f"{ном} {исполнитель['наименование']}") + ".zip",
        }
        request_clear()
        st.rerun()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def find_conflicts(заказчик, исполнитель):
    """Конфликт = подставленная из базы запись изменена в форме."""
    out = []
    for role, p, who in [("z", заказчик, "Заказчик"),
                         ("i", исполнитель, "Исполнитель")]:
        snap = st.session_state.get(f"{role}_snapshot")
        if snap and core._same_party(snap, p):
            diff = core.party_diff(snap, p)
            if diff:
                out.append({"role": role, "who": who,
                            "name": p.get("наименование"), "diff": diff})
    return out


def conflict_or_generate(заказчик, исполнитель, gen_callback):
    """Рисует диалог конфликта или кнопку генерации. gen_callback(choices)."""
    pending = st.session_state.get("pending_conflicts")
    if pending:
        st.warning("Данные контрагентов отличаются от выбранных из базы. "
                   "Что сделать с базой?")
        for c in pending:
            with st.container(border=True):
                st.markdown(f"**{c['who']}: {c['name']}**")
                for k, a, b in c["diff"]:
                    st.markdown(f"- {k}: ~~{a}~~ → **{b}**")
                st.radio("Действие",
                         ["Обновить запись (старая версия — в Архив)",
                          "Не сохранять (только в этот документ)"],
                         key=f"conf_{c['role']}", horizontal=True)
        cc1, cc2 = st.columns(2)
        if cc1.button("✅ Подтвердить и сгенерировать", type="primary",
                      use_container_width=True):
            choices = {c["role"]: ("skip" if "Не сохранять" in
                                   st.session_state.get(f"conf_{c['role']}", "")
                                   else "update") for c in pending}
            st.session_state.pop("pending_conflicts", None)
            gen_callback(choices)
        if cc2.button("↩️ Отмена", use_container_width=True):
            st.session_state.pop("pending_conflicts", None)
            st.rerun()
        return

    if st.button("🚀 Сгенерировать пакет документов", type="primary",
                 use_container_width=True):
        save_on = st.session_state.get("сохранить", True)
        conflicts = find_conflicts(заказчик, исполнитель) if save_on else []
        if conflicts:
            st.session_state["pending_conflicts"] = conflicts
            st.rerun()
        else:
            gen_callback({})


def history_expander(disk, yd_ok):
    parties = core.load_parties()
    labels = [f'{p["наименование"]} · ИНН {p["инн"] or "—"} · {p["тип"]}'
              for p in parties]
    st.divider()
    with st.expander("📜 База контрагентов: история изменений и восстановление"):
        if not parties:
            st.caption("База пока пуста.")
            return
        pick = st.selectbox("Контрагент", labels, key="hist_pick")
        p = parties[labels.index(pick)]
        hist = core.party_history(p)
        if not hist:
            st.caption("Архивных версий нет — запись ни разу не перезаписывалась.")
            return
        опции = [f'{v["заменено"]} — {v["кем"] or "имя не указано"}' for v in hist]
        vi = st.selectbox("Версия (новые сверху)", опции, key="hist_ver")
        v = hist[опции.index(vi)]
        st.table([{"поле": k, "в этой версии": v.get(k, "") or "—",
                   "сейчас в базе": p.get(k, "") or "—"}
                  for k in core.PARTY_COLS
                  if (v.get(k) or p.get(k)) and v.get(k) != p.get(k)])
        if st.button("↩️ Вернуть эту версию", key="hist_restore"):
            pull_db(disk, yd_ok)
            core.restore_party(v, user=st.session_state.get("user_name", ""))
            push_db(disk, yd_ok)
            st.success("Версия восстановлена; предыдущая ушла в Архив.")
            st.rerun()


def check_set_button(container, set_name, files, kind, fetch_fn):
    if container.button("🧪 Проверить\nкомплект", use_container_width=True,
                        key="check_set"):
        td = core.test_data(kind)
        rows = []
        for fname, data in fetch_fn(set_name, files):
            try:
                core.render_docx_bytes(data, core.build_context(td))
                rows.append(f"✅ {fname}")
            except Exception as e:
                rows.append(f"❌ {fname} — {e}")
        (st.success if all(r.startswith("✅") for r in rows) else st.error)(
            "\n\n".join(rows))
