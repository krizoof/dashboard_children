import streamlit as st
import yaml
from pathlib import Path
import hashlib
import os

def load_config():
    """Загружает конфигурацию из файла"""
    config_path = Path("config.yaml")
    if not config_path.exists():
        return None
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)

def check_password():
    """Проверяет пароль из конфигурации"""
    config = load_config()
    if not config:
        return False
    
    def password_entered():
        """Проверяет введенный пароль"""
        if st.session_state["password"] == config["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "Пароль", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "Пароль", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Неверный пароль")
        return False
    else:
        return True 