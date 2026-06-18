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


def form_gen():
    return st.session_state.get("form_gen", 0)


def wk(name):
    """Ключ виджета с «поколением» формы. Меняя поколение (request_clear),
    мы заставляем Streamlit создать свежие виджеты с дефолтами — это надёжно
    очищает форму на любой версии Streamlit, в т.ч. в облаке."""
    return f"{name}__g{form_gen()}"


def _apply_pending_clear():
    """Очистка = увеличить поколение формы и стереть старые ключи полей.
    Делать в начале прогона, до создания виджетов."""
    if st.session_state.pop("_do_clear", False):
        g = st.session_state.get("form_gen", 0)
        # удаляем все ключи прошлого поколения и снапшоты/конфликты
        for key in list(st.session_state.keys()):
            if key.endswith(f"__g{g}") or key in (
                    "z_snapshot", "i_snapshot", "pending_conflicts", "set_pick"):
                if key not in KEEP_KEYS:
                    st.session_state.pop(key, None)
        st.session_state["form_gen"] = g + 1


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
            # подсев заказчиков Точки при первом запуске (если их нет)
            if core.seed_customers():
                push_db(disk, True)
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
        st.markdown(f"📂 Папка на Яндекс Диске: [{res['folder']}]({res['url']})")
        st.markdown("🔗 **Ссылка для отправки** (по ней документы откроет любой, "
                    "кому вы её перешлёте):")
        st.code(res["url"], language=None)   # поле с кнопкой «копировать» справа
        if hasattr(st, "link_button"):
            st.link_button("🔗 Открыть ссылку для отправки", res["url"],
                           use_container_width=True)
    elif res.get("folder"):
        st.markdown(f"📂 Пакет на Яндекс Диске: **{res['folder']}** "
                    "(ссылка не создалась — скачайте архивом ниже)")
    st.download_button("⬇️ Скачать пакет (zip)", res["zip"], file_name=res["zipname"],
                       mime="application/zip", use_container_width=True,
                       key="dl_result")
    if st.button("✖ Скрыть", key="hide_result"):
        st.session_state.pop("результат", None)
        st.rerun()
    st.divider()


# ----------------------------------------------------------------- стороны
PARTY_FIELDS = [
    ("наименование", "Наименование / ФИО", "ООО «ОНОВАМНАДО» или Иванов Иван Иванович"),
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
        pick = st.session_state.get(wk(f"{role}_pick"))
        if pick in labels:
            p = parties[labels.index(pick)]
            for fld, v in p.items():
                st.session_state[wk(f"{role}_{fld}")] = v
            st.session_state[f"{role}_snapshot"] = dict(p)   # что подставили из базы
        else:
            st.session_state.pop(f"{role}_snapshot", None)
    return _fill


def party_form(role, title, kind, parties, labels, filler):
    st.subheader(title)
    st.selectbox("📇 Подставить сохранённого контрагента",
                 ["— ввести вручную —"] + labels, key=wk(f"{role}_pick"),
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
    for fld, label, hint in PARTY_FIELDS:
        if fld in show:
            data[fld] = st.text_input(label, key=wk(f"{role}_{fld}"), placeholder=hint)
    # подсветка отличий от ПОДСТАВЛЕННОЙ записи (а не от любой по ИНН)
    base = snap if (snap and core._same_party(snap, data)) else None
    if base:
        diff = core.party_diff(base, data)
        if diff:
            st.caption("✏️ Отличается от выбранной записи: " +
                       ", ".join(fld for fld, _, _ in diff))
    return data


# ----------------------------------------------------------------- общие поля
_TYPE_LABEL = {"прил": "Приложение к договору", "счет": "Счёт-оферта",
               "акт": "Акт", "дог": "Соглашение / договор"}


def _doc_keys(files, doc_types=None):
    """Группирует выбранные документы по типам (из манифеста) и возвращает
    по одной паре полей на тип: [(тип, подпись, [файлы этого типа])].
    Номер и дата типа применяются ко ВСЕМ документам этого типа."""
    doc_types = doc_types or {}
    order = ["прил", "счет", "акт", "дог"]
    groups = {}
    for fn in files:
        тип = doc_types.get(fn)
        if not тип:  # запасной разбор по имени, если в манифесте нет
            low = os.path.splitext(fn)[0].lower()
            тип = ("прил" if "приложен" in low else
                   "счет" if ("счёт" in low or "счет" in low or "оферт" in low) else
                   "акт" if "акт" in low else "дог")
        groups.setdefault(тип, []).append(fn)
    return [(т, _TYPE_LABEL.get(т, т), groups[т]) for т in order if т in groups]


def document_fields(with_vat, files=None, вид="обычный", doc_types=None):
    files = files or []

    # ---- Режим НДА: своя короткая форма (без услуг/НДС/номеров счетов) ----
    if вид == "нда":
        st.subheader("Соглашение о конфиденциальности (НДА)")
        g1, g2, g3 = st.columns(3)
        город = g1.text_input("Город", value="Москва", key=wk("город"))
        дата = g2.date_input("Дата соглашения", value=datetime.date.today(),
                             format="DD.MM.YYYY", key=wk("дог_дата"))
        отношения = g3.date_input("Сила с (дата отношений)",
                                  value=datetime.date.today(), format="DD.MM.YYYY",
                                  key=wk("нда_отношения"))
        st.caption("Реквизиты и подписанты обеих сторон берутся из выбранных "
                   "контрагентов (Заказчик/Компания и Исполнитель).")
        return {"город": город, "дог_номер": "", "дог_дата": дата,
                "оферта_отношения_с": отношения, "услуги_df": [],
                "блок_ис": False, "ндс_ставка": 0, "ндс_строка": None,
                "оферта_оплата": "", "оферта_срок": "", "оферта_результат": "",
                "оферта_формат": "", "вид": "нда"}

    docmap = _doc_keys(files, doc_types)

    st.subheader("Общие сведения")
    g1, g2, g3 = st.columns(3)
    город = g1.text_input("Город", value="Екатеринбург", key=wk("город"))
    начало = g2.date_input("Работы: начало", value=datetime.date.today(),
                           format="DD.MM.YYYY", key=wk("начало"))
    окончание = g3.date_input("Работы: окончание",
                              value=datetime.date.today() + datetime.timedelta(days=30),
                              format="DD.MM.YYYY", key=wk("окончание"))

    st.subheader("Номера и даты документов")
    st.caption("Один номер и дата на каждый ТИП документа. Если отмечено несколько "
               "счетов — у всех будет номер и дата из строки «Счёт».")
    номера = {}
    for тип, подпись, файлы in docmap:
        cc1, cc2, _ = st.columns([2, 2, 1])
        примечание = (f" (для всех: {len(файлы)} шт.)" if len(файлы) > 1 else "")
        номера[f"{тип}_номер"] = cc1.text_input(
            f"№ — {подпись}{примечание}", key=wk(f"num_{тип}"),
            placeholder="например, 12")
        номера[f"{тип}_дата"] = cc2.date_input(
            f"Дата — {подпись}", value=datetime.date.today(),
            format="DD.MM.YYYY", key=wk(f"date_{тип}"))

    st.subheader("Услуги / работы")
    st.caption("Сумма строки = количество × цена, итог считается сам.")
    услуги_df = st.data_editor(
        [{"наименование": "", "колво": 1, "цена": 0.0}],
        num_rows="dynamic", use_container_width=True, key=wk("услуги"),
        column_config={
            "наименование": st.column_config.TextColumn("Наименование", width="large"),
            "колво": st.column_config.NumberColumn("Кол-во", min_value=0, format="%g"),
            "цена": st.column_config.NumberColumn("Цена, руб.", min_value=0.0,
                                                  format="%.2f"),
        })

    st.subheader("Условия")
    u1, u2, u3 = st.columns(3)
    блок_ис = u1.checkbox("Пункт об интеллектуальной собственности", True,
                          key=wk("блок_ис"),
                          help="Снимите, если объектов ИС в работах не возникает")
    ндс_ставка = 0
    if with_vat:
        ндс_ставка = u2.number_input("Ставка НДС, %", 0, 30, 20, key=wk("ндс_ставка"))
    аванс_процент = u3.number_input("Аванс, % (для «аванс+доплата»)", 0, 100, 50,
                                    key=wk("аванс_процент"))

    with st.expander("Дополнительные поля документов (сроки, место, основание)"):
        оферта_оплата = st.text_input(
            "Срок оплаты", key=wk("оферта_оплата"), value="не позднее ______")
        оферта_срок = st.text_input(
            "Срок выполнения работ / поставки / период услуг", key=wk("оферта_срок"),
            value="в течение 10 (десяти) рабочих дней с даты внесения аванса")
        оферта_результат = st.text_input("Результат работ", key=wk("оферта_результат"),
                                         value="результат работ, указанных в Счете")
        оферта_формат = st.text_input("Формат передачи результата",
                                      key=wk("оферта_формат"),
                                      value="ссылкой на облачное хранилище")
        c1, c2 = st.columns(2)
        оферта_размещения = c1.text_input("Срок размещения (реклама)",
                                          value="______", key=wk("оферта_размещения"))
        оферта_отношения = c2.date_input("Отношения сторон с (дата)",
                                          value=datetime.date.today(),
                                          format="DD.MM.YYYY",
                                          key=wk("оферта_отношения_с"),
                                          help="Для постоплаты и НДА: дата, с которой "
                                               "договор распространяет силу")
        оферта_место = st.text_input(
            "Место оказания услуг",
            value="г. Екатеринбург, ул. Сакко и Ванцетти, д. 61",
            key=wk("оферта_место"))
        c3, c4 = st.columns(2)
        осн_номер = c3.text_input("Основание — № договора (акт услуг)",
                                  value="", key=wk("осн_номер"))
        осн_дата = c4.date_input("Основание — дата", value=datetime.date.today(),
                                 format="DD.MM.YYYY", key=wk("осн_дата"))
        прил_оплата = st.text_area(
            "Способ оплаты (для приложения, п. 3)", key=wk("прил_оплата"),
            value="Заказчик обязуется оплатить Услуги в размере 100% их стоимости до "
                  "начала оказания Услуг путем перечисления денежных средств на "
                  "банковский счет Исполнителя на основании счета на оплату, "
                  "выставленного Исполнителем. Акт оказанных услуг направляется "
                  "Исполнителем Заказчику по факту оказания Услуг.",
            help="Например: постоплата в течение 5 рабочих дней после подписания "
                 "акта; или аванс 50% + доплата")
        ндс_строка = st.text_input("Строка НДС (пусто = автоматически)", value="",
                                   key=wk("ндс_строка"))
    return {
        "город": город, "начало": начало, "окончание": окончание,
        "услуги_df": услуги_df, "блок_ис": блок_ис, "ндс_ставка": ндс_ставка,
        "оферта_оплата": оферта_оплата, "оферта_срок": оферта_срок,
        "оферта_результат": оферта_результат, "оферта_формат": оферта_формат,
        "оферта_размещения": оферта_размещения, "оферта_отношения_с": оферта_отношения,
        "оферта_место": оферта_место, "осн_номер": осн_номер, "осн_дата": осн_дата,
        "прил_оплата": прил_оплата,
        "аванс_процент": аванс_процент, "ндс_строка": ндс_строка, **номера,
    }


def assemble(f, заказчик, исполнитель):
    услуги = []
    for u in f.get("услуги_df", []):
        if not str(u.get("наименование") or "").strip():
            continue
        услуги.append({k: (None if v is None or str(v) in ("NaT", "nan") else v)
                       for k, v in u.items()})
    данные_док = {}
    for тип in ("прил", "счет", "акт", "дог"):
        if f"{тип}_номер" in f:
            данные_док[f"{тип}_номер"] = (f.get(f"{тип}_номер") or "").strip()
            данные_док[f"{тип}_дата"] = f.get(f"{тип}_дата")
    data = {
        "город": f.get("город", ""), "услуги_начало": f.get("начало"),
        "услуги_окончание": f.get("окончание"), "заказчик": заказчик,
        "исполнитель": исполнитель, "услуги": услуги,
        "блок_ис": f.get("блок_ис", True), "ндс_ставка": f.get("ндс_ставка", 0),
        "оферта_срок_оплаты": f.get("оферта_оплата") or None,
        "оферта_срок_работ": f.get("оферта_срок") or None,
        "оферта_результат": f.get("оферта_результат") or None,
        "оферта_формат": f.get("оферта_формат") or None,
        "оферта_срок_размещения": f.get("оферта_размещения") or None,
        "оферта_отношения_с": f.get("оферта_отношения_с") or None,
        "оферта_место": f.get("оферта_место") or None,
        "оферта_основание_номер": f.get("осн_номер") or None,
        "оферта_основание_дата": f.get("осн_дата"),
        "прил_оплата": f.get("прил_оплата") or None,
        "аванс_процент": f.get("аванс_процент", 50),
        "ндс_строка": f.get("ндс_строка") or None,
        **данные_док,
    }
    return data, услуги


# ----------------------------------------------------------------- генерация
def generate_and_store(data, услуги, templates, set_name, disk, yd_ok, choices,
                       заказчик, исполнитель, doc_types=None):
    """Единая процедура: рендер -> zip -> (Диск) -> база/журнал -> результат."""
    user = st.session_state.get("user_name", "")
    pull_db(disk, yd_ok)

    warnings = core.validate_party(заказчик, "Заказчик") + \
        core.validate_party(исполнитель, "Исполнитель")
    seen_nums = set(core.journal_numbers())
    for тип in ("прил", "счет", "акт", "дог"):
        n = (data.get(f"{тип}_номер") or "").strip()
        if n and n in seen_nums:
            warnings.append(f"Номер {n} уже встречался в журнале.")

    главный = core._primary_number(data) or datetime.date.today().strftime("%Y-%m-%d")

    tmp = tempfile.mkdtemp()
    try:
        paths = core.generate_files(data, tmp, templates, doc_types=doc_types)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for p in paths:
                z.write(p, os.path.basename(p))

        folder_name = url = ""
        if yd_ok:
            folder_name = core.safe_name(f"{главный} — {исполнитель['наименование']}")
            folder = f"{disk.base}/Документы/{folder_name}"
            try:
                disk.ensure_path(folder)
                for p in paths:
                    disk.upload_file(p, folder + "/" + os.path.basename(p))
                url = disk.publish(folder)
            except Exception as e:
                warnings.append(f"На Диск загрузить не удалось: {e}. Скачайте zip.")
                folder_name = url = ""

        if st.session_state.get(wk("сохранить"), True):
            for role, p in [("z", заказчик), ("i", исполнитель)]:
                if choices.get(role) != "skip":
                    core.save_party(p, user=user)
            data["стоимость"] = sum(float(u.get("цена") or 0) *
                                    float(u.get("колво") or 1) for u in услуги)
            core.append_journal(data, [os.path.splitext(f)[0] for f, _ in templates],
                                комплект=set_name, user=user)
            push_db(disk, yd_ok)

        st.session_state["результат"] = {
            "text": f"Готово! {главный}: {len(paths)} док., сумма "
                    f"{core.money_fmt(data.get('стоимость') or 0)} руб. "
                    f"Форма очищена для следующего пакета.",
            "warnings": warnings, "folder": folder_name, "url": url,
            "zip": buf.getvalue(),
            "zipname": core.safe_name(f"{главный} {исполнитель['наименование']}") + ".zip",
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
        save_on = st.session_state.get(wk("сохранить"), True)
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
