# -*- coding: utf-8 -*-
"""Форма генерации договоров. Запуск:  streamlit run app.py
Облачный режим: задайте секреты YANDEX_DISK_TOKEN и APP_PASSWORD (см. README)."""
import datetime, io, os, shutil, tempfile, zipfile
import streamlit as st
import core, ydisk

st.set_page_config(page_title="Генератор договоров", page_icon="📄", layout="wide")


def get_secret(name, default=""):
    try:
        v = st.secrets.get(name)
        if v:
            return str(v)
    except Exception:
        pass
    return os.environ.get(name, default)


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
    # первое подключение в сессии: проверяем токен, создаем папки,
    # докладываем недостающие шаблоны и забираем базу контрагентов с Диска
    try:
        st.session_state["yd_status"] = disk.check()
        disk.bootstrap(core.TPL_DIR)
        data = disk.download(disk.base + "/контрагенты.xlsx")
        if data:
            with open(core.DB_PATH, "wb") as f:
                f.write(data)
    except Exception as e:
        st.session_state["yd_status"] = None
        st.session_state["yd_error"] = str(e)

yd_ok = bool(disk and st.session_state.get("yd_status"))


def push_db():
    """Отправляет базу контрагентов на Диск после изменений."""
    if yd_ok and os.path.exists(core.DB_PATH):
        try:
            disk.upload_file(core.DB_PATH, disk.base + "/контрагенты.xlsx")
        except Exception as e:
            st.warning(f"База сохранена локально, но не загрузилась на Диск: {e}")


with st.sidebar:
    st.header("☁️ Яндекс Диск")
    if disk is None:
        st.caption("Без Диска всё работает — пакет можно скачать zip-архивом. "
                   "С Диском документы и база контрагентов хранятся в облаке.")
        t = st.text_input("OAuth-токен Диска", type="password",
                          help="Как получить токен — раздел «Облачный режим» в README")
        if t:
            st.session_state["yd_token"] = t
            st.rerun()
    elif yd_ok:
        st.success(f"Подключено: {st.session_state['yd_status']}")
        st.caption(f"Папка: {disk.base.replace('disk:', '') or '/'}\n\n"
                   "Шаблоны можно править прямо на Диске — изменения "
                   "подхватываются при следующей генерации.")
    else:
        st.error(f"Диск недоступен: {st.session_state.get('yd_error', '')}")
        if st.button("Повторить подключение"):
            st.session_state.pop("yd_status", None)
            st.rerun()

# ---------------------------------------------------------------- база
parties = core.load_parties()
labels = [f'{p["наименование"]} (ИНН {p["инн"]})' if p["инн"] else p["наименование"]
          for p in parties]

PARTY_FIELDS = [  # (ключ, подпись, подсказка)
    ("наименование", "Наименование / ФИО", "ООО «Ромашка» или Иванов Иван Иванович"),
    ("инн", "ИНН", ""), ("кпп", "КПП (для ООО)", ""), ("огрн", "ОГРН / ОГРНИП", ""),
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
ORG_ONLY = {"кпп", "должность_род", "подписант_род", "подписант", "основание_род"}
FIZ_ONLY = {"паспорт"}


def _fill_from_db(role):
    """Колбэк выпадающего списка: подставляет реквизиты выбранного контрагента."""
    pick = st.session_state.get(f"{role}_pick")
    if pick in labels:
        for k, v in parties[labels.index(pick)].items():
            st.session_state[f"{role}_{k}"] = v


def party_form(role, title, org=True):
    st.subheader(title)
    st.selectbox("📇 Подставить сохранённого контрагента",
                 ["— ввести вручную —"] + labels, key=f"{role}_pick",
                 on_change=_fill_from_db, args=(role,),
                 help="Контрагенты сохраняются в базу при генерации пакета"
                 if labels else "База пока пуста — контрагенты появятся "
                                "после первой генерации")
    data = {}
    if not org:
        тип = st.radio("Статус исполнителя", ["Самозанятый", "ИП"],
                       key=f"{role}_тип", horizontal=True)
        data["тип"] = тип
    for k, label, hint in PARTY_FIELDS:
        if org and k in FIZ_ONLY:
            continue
        if not org and k in ORG_ONLY:
            continue
        if not org and k == "паспорт" and data.get("тип") == "ИП":
            continue  # для ИП паспорт в документах не используется
        data[k] = st.text_input(label, key=f"{role}_{k}", placeholder=hint)
    if org:
        data["тип"] = st.session_state.get(f"{role}_тип") or "ООО"
    return data


col_z, col_i = st.columns(2)
with col_z:
    заказчик = party_form("z", "Заказчик (организация)", org=True)
with col_i:
    исполнитель = party_form("i", "Исполнитель", org=False)

# ---------------------------------------------------------------- договор
st.subheader("Договор / документ")
c1, c2, c3, c4, c5 = st.columns(5)
номер = c1.text_input("№ договора", value=core.next_contract_number())
дата = c2.date_input("Дата договора", value=datetime.date.today(), format="DD.MM.YYYY")
город = c3.text_input("Город", value="Екатеринбург")
начало = c4.date_input("Услуги/работы: начало", value=datetime.date.today(),
                       format="DD.MM.YYYY")
окончание = c5.date_input("Услуги/работы: окончание",
                          value=datetime.date.today() + datetime.timedelta(days=30),
                          format="DD.MM.YYYY")

st.subheader("Услуги / работы")
st.caption("Сумма строки = количество × цена. Итог по документам считается сам. "
           "Даты можно не заполнять — возьмутся сроки договора.")
услуги_df = st.data_editor(
    [{"наименование": "", "колво": 1, "цена": 0.0, "начало": None, "окончание": None,
      "требования": ""}],
    num_rows="dynamic", use_container_width=True,
    column_config={
        "наименование": st.column_config.TextColumn("Наименование", width="large"),
        "колво": st.column_config.NumberColumn("Кол-во", min_value=0, format="%g"),
        "цена": st.column_config.NumberColumn("Цена, руб.", min_value=0.0,
                                              format="%.2f"),
        "начало": st.column_config.DateColumn("Начало", format="DD.MM.YYYY"),
        "окончание": st.column_config.DateColumn("Окончание", format="DD.MM.YYYY"),
        "требования": st.column_config.TextColumn("Доп. требования (для Задания)"),
    })

# ---------------------------------------------------------------- блоки
st.subheader("Блоки и условия")
b1, b2, b3, b4 = st.columns(4)
блок_нда = b1.checkbox("Конфиденциальность (NDA)")
нда_лет = b1.number_input("NDA после договора, лет", 1, 10, 3, disabled=not блок_нда)
блок_фм = b2.checkbox("Форс-мажор (раздел)")
фм_месяцев = b2.number_input("Расторжение после, мес.", 1, 12, 2, disabled=not блок_фм)
блок_ис = b3.checkbox("Пункт об интеллектуальной собственности (Точка)", True,
                      help="Уберите галочку, если объектов ИС в работах не возникает")
предоплата = b4.slider("Предоплата по договору, %", 0, 100, 50)
b4.caption(f"Постоплата: {100 - предоплата}% после Акта")

with st.expander("Сроки и штрафы договора (по умолчанию обычно подходят)"):
    e1, e2, e3, e4, e5 = st.columns(5)
    дни = {
        "справка_дней": e1.number_input("Справка самозанятого, дн.", 1, 30, 3),
        "пояснения_дней": e2.number_input("Пояснения, дн.", 1, 30, 3),
        "акт_дней": e3.number_input("Передача Акта, дн.", 1, 30, 5),
        "приемка_дней": e4.number_input("Приемка Акта, дн.", 1, 30, 5),
        "устранение_дней": e5.number_input("Устранение недостатков, дн.", 1, 30, 5),
        "предоплата_дней": e1.number_input("Оплата аванса, дн.", 1, 30, 5),
        "постоплата_дней": e2.number_input("Постоплата, дн.", 1, 30, 5),
        "расторжение_дней": e3.number_input("Расторжение, дн.", 1, 60, 10),
        "претензия_дней": e4.number_input("Ответ на претензию, дн.", 1, 60, 10),
        "фм_уведомление_дней": e5.number_input("Уведомл. о форс-мажоре, дн.", 1, 30, 10),
    }
    s1, s2 = st.columns(2)
    штраф_чек = s1.number_input("Штраф за невыдачу чека, %", 0, 100, 20)
    штраф_увед = s2.number_input("Штраф за неуведомление об утрате НПД, руб.",
                                 0, 1000000, 10000)

with st.expander("Поля документов «Точка» (приложение, счёт-оферта)"):
    t1, t2 = st.columns(2)
    прил_номер = t1.text_input("№ приложения к договору", value="1")
    прил_дата = t2.date_input("Дата приложения", value=datetime.date.today(),
                              format="DD.MM.YYYY")
    оферта_оплата = st.text_input(
        "Срок оплаты аванса (текстом)",
        value="в течение 5 (пяти) рабочих дней с даты выставления Счета")
    оферта_срок = st.text_input(
        "Срок выполнения работ (текстом)",
        value="в течение 10 (десяти) рабочих дней с даты внесения аванса")
    оферта_результат = st.text_input("Результат работ",
                                     value="результат работ, указанных в Счете")
    оферта_формат = st.text_input("Формат передачи результата",
                                  value="ссылкой на облачное хранилище")
    ндс_строка = st.text_input("Строка НДС (пусто = подставится по статусу исполнителя)",
                               value="")

st.subheader("Какие документы сгенерировать")
st.caption("Пакет «Договор с самозанятым»")
d1, d2, d3, d4, d5, d6 = st.columns(6)
docs = []
if d1.checkbox("Договор", True): docs.append("договор")
if d2.checkbox("Прил.1 Перечень", True): docs.append("прил1")
if d3.checkbox("Прил.2 Задание", True): docs.append("прил2")
if d4.checkbox("Прил.3 Акт", True): docs.append("акт")
if d5.checkbox("Прил.4 Согласие ПД", True): docs.append("согласие")
if d6.checkbox("Счет", False): docs.append("счет")
st.caption("Формы «Точка»")
p1, p2, p3, _, _, _ = st.columns(6)
if p1.checkbox("Приложение к договору", False): docs.append("точка_прил")
if p2.checkbox("Акт сдачи-приёмки", False): docs.append("точка_акт")
if p3.checkbox("Счёт-оферта", False): docs.append("точка_счет")

сохранить = st.checkbox("Сохранить контрагентов в базу и записать в журнал", True)


def resolve_templates(keys):
    """Шаблоны для генерации: свежие с Диска, иначе локальные из комплекта."""
    if not yd_ok:
        return core.TPL_DIR
    tpl_dir = os.path.join(tempfile.gettempdir(), "tpl_cache")
    os.makedirs(tpl_dir, exist_ok=True)
    for key in keys:
        fname = core.DOCS[key][0]
        with open(os.path.join(tpl_dir, fname), "wb") as f:
            f.write(disk.fetch_template(fname, core.TPL_DIR))
    return tpl_dir


# ---------------------------------------------------------------- генерация
if st.button("🚀 Сгенерировать пакет документов", type="primary",
             use_container_width=True):
    услуги = []
    for u in услуги_df:
        if not str(u.get("наименование") or "").strip():
            continue
        услуги.append({k: (None if v is None or str(v) in ("NaT", "nan") else v)
                       for k, v in u.items()})
    if not услуги:
        st.error("Добавьте хотя бы одну услугу/работу с наименованием.")
        st.stop()
    if not заказчик["наименование"] or not исполнитель["наименование"]:
        st.error("Заполните наименования Заказчика и Исполнителя.")
        st.stop()
    if not docs:
        st.error("Отметьте хотя бы один документ.")
        st.stop()

    сумма = sum(float(u.get("цена") or 0) * float(u.get("колво") or 1) for u in услуги)
    data = {
        "договор_номер": номер, "договор_дата": дата, "город": город,
        "услуги_начало": начало, "услуги_окончание": окончание,
        "заказчик": заказчик, "исполнитель": исполнитель, "услуги": услуги,
        "предоплата_процент": предоплата, "постоплата_процент": 100 - предоплата,
        "штраф_чек_процент": штраф_чек, "штраф_уведомление": штраф_увед,
        "блок_нда": блок_нда, "нда_лет": нда_лет,
        "блок_фм": блок_фм, "фм_месяцев": фм_месяцев,
        "блок_согласие": "согласие" in docs, "блок_ис": блок_ис,
        "прил_номер": прил_номер, "прил_дата": прил_дата,
        "оферта_срок_оплаты": оферта_оплата, "оферта_срок_работ": оферта_срок,
        "оферта_результат": оферта_результат, "оферта_формат": оферта_формат,
        "ндс_строка": ндс_строка or None,
        **дни,
    }
    tmp = tempfile.mkdtemp()
    try:
        paths = core.generate_package(data, tmp, docs=docs,
                                      tpl_dir=resolve_templates(docs))
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for p in paths:
                z.write(p, os.path.basename(p))

        if сохранить:
            core.save_party(заказчик)
            core.save_party(исполнитель)
            data["стоимость"] = сумма
            core.append_journal(data, docs)
            push_db()

        st.success(f"Готово! Документов: {len(paths)}. "
                   f"Сумма: {core.money_fmt(сумма)} руб.")

        if yd_ok:  # выгрузка пакета на Диск + публичная ссылка
            пакет = f"№{номер} — {исполнитель['наименование']}"
            for ch in '/\\:*?"<>|':
                пакет = пакет.replace(ch, "-")
            folder = f"{disk.base}/Документы/{пакет}"
            try:
                disk.ensure_path(folder)
                for p in paths:
                    disk.upload_file(p, folder + "/" + os.path.basename(p))
                url = disk.publish(folder)
                st.markdown(f"📂 Пакет на Яндекс Диске: **{пакет}**" +
                            (f" — [открыть / поделиться]({url})" if url else ""))
            except Exception as e:
                st.warning(f"На Диск загрузить не удалось ({e}), скачайте zip ниже.")

        st.download_button("⬇️ Скачать пакет (zip)", buf.getvalue(),
                           file_name=f"Документы_{номер}_{исполнитель['наименование']}.zip",
                           mime="application/zip", use_container_width=True)
    except Exception as e:
        st.error(str(e))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
