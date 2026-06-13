# -*- coding: utf-8 -*-
"""Вариант с ОДНИМ комплектом шаблонов (без выбора).
Шаблоны лежат на Диске плоско в «Шаблоны/», тип исполнителя — Самозанятый.
Запуск:  streamlit run app_single.py"""
import os
import streamlit as st
import core, ydisk
import form_common as fc

st.set_page_config(page_title="Генератор договоров", page_icon="📄", layout="wide")
fc._apply_pending_clear()
fc.password_gate()
st.title("📄 Генератор договоров")

KIND = "Самозанятый"   # тип исполнителя единственного комплекта
WITH_VAT = False        # строка НДС не нужна

disk = fc.get_disk()
yd_ok = fc.connect_disk(disk, multi=False)
fc.sidebar(disk, yd_ok)
fc.show_result_panel()


def load_single():
    """Список файлов единственного комплекта: с Диска (плоско) или локально."""
    if st.session_state.get("sets"):
        return st.session_state["sets"]
    files = []
    if yd_ok:
        try:
            files = disk.list_flat_files()
        except Exception as e:
            st.warning(f"Не удалось прочитать шаблоны с Диска ({e}).")
    if not files:
        files = core.local_sets().get(core.SINGLE_SET, [])
    st.session_state["sets"] = files
    return files


def fetch(_set, files):
    out = []
    for f in files:
        data = disk.fetch_flat(f, core.TPL_DIR) if yd_ok else \
            open(os.path.join(core.TPL_DIR, core.SINGLE_SET, f), "rb").read()
        out.append((f, data))
    return out


files = load_single()
if not files:
    st.error("Не найдено ни одного шаблона в папке «Шаблоны/».")
    st.stop()

_, ctest = st.columns([4, 1])
fc.check_set_button(ctest, None, files, KIND, fetch)

parties = core.load_parties()
labels = [f'{p["наименование"]} · ИНН {p["инн"] or "—"} · {p["тип"]}' for p in parties]
filler = fc.make_filler(parties, labels)

col_z, col_i = st.columns(2)
with col_z:
    заказчик = fc.party_form("z", "Заказчик", "ЮЛ", parties, labels, filler)
with col_i:
    исполнитель = fc.party_form("i", f"Исполнитель ({KIND})", KIND,
                                parties, labels, filler)

f = fc.document_fields(WITH_VAT)

st.subheader("Какие документы сгенерировать")
dcols = st.columns(max(len(files), 1))
выбранные = [fn for i, fn in enumerate(files)
             if dcols[i].checkbox(os.path.splitext(fn)[0], True,
                                  key=f"doc_{fn}")]

st.checkbox("Сохранить контрагентов в базу и записать в журнал", True,
            key="сохранить")


def gen(choices):
    data, услуги = fc.assemble(f, заказчик, исполнитель)
    if not услуги:
        st.error("Добавьте хотя бы одну услугу/работу с наименованием."); return
    if not заказчик["наименование"] or not исполнитель["наименование"]:
        st.error("Заполните наименования Заказчика и Исполнителя."); return
    if not выбранные:
        st.error("Отметьте хотя бы один документ."); return
    try:
        fc.generate_and_store(data, услуги, fetch(None, выбранные), core.SINGLE_SET,
                              disk, yd_ok, choices, заказчик, исполнитель)
    except Exception as e:
        st.error(str(e))


fc.conflict_or_generate(заказчик, исполнитель, gen)
fc.history_expander(disk, yd_ok)
