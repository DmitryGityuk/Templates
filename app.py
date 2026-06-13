# -*- coding: utf-8 -*-
"""Точка входа по умолчанию = вариант с выбором комплекта.
Для варианта с одним комплектом запускайте app_single.py."""
import runpy
runpy.run_path("app_multi.py", run_name="__main__")
