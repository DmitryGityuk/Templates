# -*- coding: utf-8 -*-
"""Вариант с ВЫБОРОМ комплекта шаблонов (Самозанятые / ИП / ЮЛ / …).
Каждый комплект — папка в «Шаблоны/» с манифест.json.
Запуск:  streamlit run app_multi.py"""
import os
import streamlit as st
import core, ydisk
import form_common as fc

st.set_page_config(page_title="Генератор договоров", page_icon="📄", layout="wide")
fc._apply_pending_clear()
fc.password_gate()
st.title("📄 Генератор договоров")

disk = fc.get_disk()
yd_ok = fc.connect_disk(disk, multi=True)
fc.sidebar(disk, yd_ok)
fc.show_result_panel()


def load_sets():
    if st.session_state.get("sets"):
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
            st.warning(f"Не удалось прочитать комплекты с Диска ({e}).")
    if not sets:
        for name, files in core.local_sets().items():
            mp = os.path.join(core.TPL_DIR, name, core.MANIFEST)
            raw = open(mp, "rb").read() if os.path.exists(mp) else None
            m = core.read_manifest(raw, name)
            m["files"] = [f for f in files if f.endswith(".docx")]
            sets[name] = m
    st.session_state["sets"] = sets
    return sets


def fetch(set_name, files):
    out = []
    for f in files:
        data = disk.fetch(set_name, f, core.TPL_DIR) if yd_ok else \
            open(os.path.join(core.TPL_DIR, set_name, f), "rb").read()
        out.append((f, data))
    return out


sets = load_sets()
if not sets:
    st.error("Не найдено ни одного комплекта шаблонов (папки в «Шаблоны/»).")
    st.stop()

cset, ctest = st.columns([4, 1])
set_name = cset.selectbox("📁 Комплект шаблонов", list(sets.keys()), key="set_pick")
m = sets[set_name]
KIND = m["тип_исполнителя"]
cset.caption(f"Исполнитель: {KIND}" + (" · с НДС" if m["ндс"] else ""))
fc.check_set_button(ctest, set_name, m["files"], KIND, fetch)

parties = core.load_parties()
labels = [f'{p["наименование"]} · ИНН {p["инн"] or "—"} · {p["тип"]}' for p in parties]
filler = fc.make_filler(parties, labels)

col_z, col_i = st.columns(2)
with col_z:
    заказчик = fc.party_form("z", "Заказчик", "ЮЛ", parties, labels, filler)
with col_i:
    исполнитель = fc.party_form("i", f"Исполнитель ({KIND})", KIND,
                                parties, labels, filler)

f = fc.document_fields(m["ндс"])

st.subheader("Какие документы сгенерировать")
dcols = st.columns(max(len(m["files"]), 1))
выбранные = [fn for i, fn in enumerate(m["files"])
             if dcols[i].checkbox(os.path.splitext(fn)[0], True,
                                  key=f"doc_{set_name}_{fn}")]

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
        fc.generate_and_store(data, услуги, fetch(set_name, выбранные), set_name,
                              disk, yd_ok, choices, заказчик, исполнитель)
    except Exception as e:
        st.error(str(e))


fc.conflict_or_generate(заказчик, исполнитель, gen)
fc.history_expander(disk, yd_ok)
