import sys
import os
import time
import asyncio
import aiohttp
import json
import re
import random
import subprocess
import webbrowser
import pyautogui
import ntpath
import threading
import logging
from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QHBoxLayout, QPushButton
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve, QTimer, QPoint, QRect, pyqtSignal
from PyQt6.QtGui import QPixmap, QColor, QIcon
from pystray import Icon, Menu, MenuItem
from PIL import Image
import speech_recognition as sr
import pyttsx3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import ctypes
import requests  # Для синхронного теста API

# Настройка логирования
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("alfred_assistant.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Исправление кодировки консоли
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

# Подавление предупреждений Qt
os.environ["QT_LOGGING_RULES"] = "qt5ct.debug=false;qt.qpa.*=false"

class AlfredAssistant(QMainWindow):
    auto_hide_signal = pyqtSignal()

    def __init__(self, app):
        super().__init__()  # Вызываем __init__ базового класса первым делом
        self.app = app
        logger.info("Инициализация AlfredAssistant")
        print("Инициализация AlfredAssistant")
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Tool)
        self.screen = self.app.primaryScreen().geometry()
        self.window_width = 500
        self.window_height = 150
        self.start_x = self.screen.width()
        self.start_y = 150
        logger.debug(f"Размеры экрана: {self.screen.width()}x{self.screen.height()}")
        logger.debug(f"Начальные координаты окна: x={self.start_x}, y={self.start_y}")
        self.setGeometry(self.start_x, self.start_y, self.window_width, self.window_height)
        self.auto_hide_signal.connect(self.start_auto_hide_timer) 
       
        # Глобальные настройки
        self.MISTRAL_API_KEY = "K7LIwiBBBwJJyEWtHtQpW8QoWiGVZQf4"
        logger.debug(f"API ключ: {self.MISTRAL_API_KEY}")
        self.USER_HISTORY_FILE = "user_histories.json"
        self.USER_PREFS_FILE = "user_preferences.json"
        self.user_histories = {}
        self.user_preferences = {}
        self.command_usage = {}
        self.last_command_time = {}
        self.command_cooldown = 2
        self.is_muted_manually = False
        self.is_command_running = False
        self.CURRENT_QUEUE = 0
        self.current_track_index = -1
        self.track_history = []
        
        
        self.ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")
        logger.debug(f"Путь к assets: {self.ASSETS_DIR}")
        self.json_lock = threading.Lock()
        
        # Флаг для переключения между реальным API и заглушкой
        self.USE_MISTRAL_API = True  # Установите в False для теста без API
        
        # Словарь синонимов
        self.synonyms = {
            "actions": {
                "open": ["открой", "запусти", "включи", "открыть", "запустить", "включить", "покажи", "open", "launch", "start"],
                "close": ["закрой", "выключи", "закрыть", "выключить", "close", "exit"],
                "play": ["воспроизведи", "прогон", "включи", "играй", "поставь", "слушай", "play"],
                "pause": ["поставь на паузу", "пауза", "приостанови", "pause", "stop"],
                "resume": ["возобнови", "продолжи", "resume", "play again"],
                "app_volume_up": ["увеличь громкость в", "громче в", "increase volume in"],
                "app_volume_down": ["уменьши громкость в", "тише в", "decrease volume in"],
                "find": ["найди", "поищи", "ищи", "найти", "поиск", "find", "search"],
                "weather": ["погода", "weather", "какая погода", "what's the weather"],
                "time": ["время", "time", "сколько времени", "what time is it"],
                "screenshot": ["скриншот", "снимок экрана", "screenshot", "take a screenshot"],
                "reminder": ["напомни", "напомнить", "установи напоминание", "set a reminder"],
                "volume_up": ["громче", "увеличь громкость", "volume up", "louder"],
                "volume_down": ["тише", "уменьши громкость", "volume down", "quieter"],
                "mute": ["выключи звук", "mute", "без звука"],
                "unmute": ["включи звук", "unmute", "с звуком"],
                "minimize": ["сверни", "минимизируй", "minimize"],
                "maximize": ["разверни", "максимизируй", "maximize"],
                "input": ["введи", "вводи", "напиши", "вставить", "type", "input"],
                "set_queue": ["очередь", "установи очередь", "set queue"],
                "check_queue": ["какая сейчас очередь", "проверь очередь", "сколько в очереди", "check queue"],
                "next_track": ["следующий трек", "следующий", "дальше", "next track", "next"],
                "previous_track": ["предыдущий трек", "назад", "предыдущий", "previous track", "previous"],
                "set_preference": ["установи", "задай", "запомни", "set preference"],
                "shutdown": ["выключи компьютер", "shutdown", "выключить"],
                "restart": ["перезагрузи компьютер", "restart", "перезагрузить"],
                "lock": ["заблокируй компьютер", "lock", "заблокировать"],
                "new_tab": ["новая вкладка", "открой новую вкладку", "new tab"],
                "close_tab": ["закрой вкладку", "close tab"],
                "switch_tab": ["переключи вкладку", "switch tab"],
                "refresh": ["обнови страницу", "refresh", "перезагрузить страницу"],
                "go_back": ["назад", "вернись назад", "go back"],
                "go_forward": ["вперёд", "давай вперёд", "go forward"],
                "scroll_down": ["пролистни вниз", "прокрути вниз", "листай вниз", "scroll down"],
                "scroll_up": ["пролистни вверх", "прокрути вверх", "листай вверх", "scroll up"],
                "click": ["кликни", "нажми", "click"],
                "double_click": ["двойной клик", "дважды кликни", "double click"],
                "right_click": ["правый клик", "кликни правой", "right click"],
                "type_password": ["введи пароль", "type password"],
                "type_email": ["введи email", "type email"],
                "clear": ["очисти", "clear", "удали текст"],
                "click_button": ["нажми на", "кликни на", "найди и нажми"],

            },
            "objects": {
                "browser": ["браузер", "интернет", "сайт", "веб", "browser", "web"],
                "music": ["музыка", "песня", "песни", "мелодия", "аудио", "трек", "композиция", "music", "song"],
                "youtube": ["ютуб", "youtube", "видео", "ролики"],
                "notepad": ["блокнот", "текст", "ноутпад", "notepad"],
                "explorer": ["проводник", "файлы", "папки", "explorer", "files"],
                "calculator": ["калькулятор", "расчёт", "счёт", "calculator"],
                "soundcloud": ["soundcloud", "саундклауд"],
                "yandex_music": ["яндекс музыка", "yandex music"],
                "spotify": ["спотифай", "spotify"],
                "discord": ["дискорд", "discord"],
                "telegram": ["телеграм", "telegram"],
                "whatsapp": ["ватсап", "whatsapp"],
                "skype": ["скайп", "skype"],
                "vscode": ["визуал студио", "vscode", "visual studio code"],
                "word": ["ворд", "word", "microsoft word"],
                "excel": ["эксель", "excel", "microsoft excel"],
                "powerpoint": ["пауэрпойнт", "powerpoint", "microsoft powerpoint"],
                "paint": ["пейнт", "paint"],
                "cmd": ["командная строка", "cmd", "command prompt"],
                "task_manager": ["диспетчер задач", "task manager"],
                "settings": ["настройки", "параметры", "settings"],
                "control_panel": ["панель управления", "control panel"],
            },
            "additional": {
                "play_first": ["включи первое", "запустить первое", "открой первое", "первый ролик", "первый трек", "play first"],
                "play_by_number": ["включи второй", "включи третий", "включи четвёртый", "второй ролик", "третий ролик", "четвёртый ролик", "play second", "play third"],
                "play_by_name": ["включи с надписью", "воспроизведи с названием", "открой с текстом", "play with title"],
                "scroll_down": ["пролистни вниз", "прокрути вниз", "листай вниз", "scroll down"],
                "scroll_up": ["пролистни вверх", "прокрути вверх", "листай вверх", "scroll up"],
                "default_music": ["по умолчанию", "default music"],
                "default_city": ["мой город", "default city"],
                "on_left": ["слева", "на лево", "left"],
                "on_right": ["справа", "на право", "right"],
                "in_center": ["по центру", "в центре", "center"],
            }
        }
        
        self._init_ui()
        self.engine = pyttsx3.init()
        logger.debug("pyttsx3 инициализирован")
        self._init_tray()

    def _init_ui(self):
        logger.info("Инициализация интерфейса")
        print("Инициализация интерфейса")
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        
        background_path = os.path.join(self.ASSETS_DIR, 'background.png').replace('\\', '/')
        logger.debug(f"Проверка фона: {background_path}")
        if os.path.exists(background_path):
            logger.debug("Фон найден, устанавливаем стиль")
            self.central_widget.setStyleSheet(f"background-image: url('{background_path}'); border-radius: 10px;")
        else:
            logger.error(f"Фон не найден: {background_path}")
            print(f"Ошибка: Фон не найден: {background_path}")
        
        mic_icon_path = os.path.join(self.ASSETS_DIR, "mic_icon.png").replace('\\', '/')
        logger.debug(f"Проверка иконки микрофона: {mic_icon_path}")
        if os.path.exists(mic_icon_path):
            logger.debug("Иконка микрофона найдена")
            self.mic_label = QLabel(self.central_widget)
            self.mic_label.setPixmap(QPixmap(mic_icon_path).scaled(40, 40))
            self.mic_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.main_layout.addWidget(self.mic_label)
        else:
            logger.error(f"Иконка микрофона не найдена: {mic_icon_path}")
            print(f"Ошибка: Иконка микрофона не найдена: {mic_icon_path}")
        
        self.label = QLabel("Ожидаю 'Альфред'...", self.central_widget)
        self.label.setStyleSheet("color: white; font-size: 16px; padding: 10px;")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.main_layout.addWidget(self.label)
        logger.debug("Метка 'Ожидаю Альфред' добавлена")
        
        self.indicator_layout = QHBoxLayout()
        self.indicator_images = {}
        for color in ["white", "green", "blue", "yellow", "red"]:
            path = os.path.join(self.ASSETS_DIR, f"indicator_{color}.png").replace('\\', '/')
            logger.debug(f"Проверка индикатора {color}: {path}")
            if os.path.exists(path):
                self.indicator_images[color] = QPixmap(path).scaled(10, 10)
            else:
                logger.error(f"Индикатор {color} не найден: {path}")
                print(f"Ошибка: Индикатор {color} не найден: {path}")
        # Добавляем новые состояния
        self.indicator_images["processing"] = self.indicator_images.get("blue", QPixmap())  # Синий для обработки
        self.indicator_images["idle"] = self.indicator_images.get("white", QPixmap())      # Белый для ожидания
        self.indicator_label = QLabel(self.central_widget)
        self.indicator_label.setPixmap(self.indicator_images.get("white", QPixmap()))
        self.indicator_layout.addWidget(self.indicator_label)
        self.main_layout.addLayout(self.indicator_layout)
        logger.debug("Индикатор добавлен")
        
        self.mute_button = QPushButton()
        self.stop_button = QPushButton()
        self.volume_button = QPushButton()
        
        self.control_layout = QHBoxLayout()
        for icon_name, button in [("mute_button", self.mute_button), ("stop_button", self.stop_button),]:
            icon_path = os.path.join(self.ASSETS_DIR, f"{icon_name}.png").replace('\\', '/')
            logger.debug(f"Проверка иконки кнопки {icon_name}: {icon_path}")
            if os.path.exists(icon_path):
                button.setIcon(QIcon(icon_path))
                button.setFixedSize(40, 40)
                button.setStyleSheet("background-color: transparent; border: none;")
            else:
                logger.error(f"Иконка {icon_name} не найдена: {icon_path}")
                print(f"Ошибка: Иконка {icon_name} не найдена: {icon_path}")
            self.control_layout.addWidget(button)
        self.main_layout.addLayout(self.control_layout)
        logger.debug("Кнопки управления добавлены")
        
        self.mute_button.clicked.connect(self.toggle_mute)
        self.stop_button.clicked.connect(self.stop_current_command)

    def _init_tray(self):
        logger.info("Инициализация системного трея")
        print("Инициализация системного трея")
        tray_icon_path = os.path.join(self.ASSETS_DIR, "tray_icon.png")
        logger.debug(f"Проверка иконки трея: {tray_icon_path}")
        if os.path.exists(tray_icon_path):
            self.icon = Icon("Alfred", Image.open(tray_icon_path), "Альфред", Menu(
                MenuItem("Показать", self.on_show),
                MenuItem("Выход", self.on_quit)
            ))
            logger.debug("Системный трей инициализирован")
        else:
            logger.error(f"Иконка трея не найдена: {tray_icon_path}")
            print(f"Ошибка: Иконка трея не найдена: {tray_icon_path}")
            self.icon = None

    def is_gui_available(self):
        logger.debug("Проверка доступности GUI")
        return os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY") or os.environ.get("XDG_SESSION_TYPE") == "x11"

    def safe_press(self, key):
        logger.debug(f"Попытка выполнить нажатие клавиши: {key}")
        if self.is_gui_available():
            pyautogui.press(key)
            logger.info(f"Клавиша {key} нажата")
        else:
            logger.warning(f"GUI недоступен, действие '{key}' невозможно выполнить.")
            print(f"Ошибка: GUI недоступен для действия '{key}'")

    def animate_show(self):
        logger.info("Запуск анимации появления окна")
        print("Запуск анимации появления окна")
        logger.debug(f"Текущее состояние окна: visible={self.isVisible()}")
        if not self.isVisible():
            logger.debug("Окно не видимо, показываем")
            print("Попытка показать окно")
            self.show()
            self.raise_()
            logger.debug(f"После show(): visible={self.isVisible()}")
        else:
            logger.debug("Окно уже видимо")
            print("Окно уже видимо")
        
        animation_pos = QPropertyAnimation(self, b"pos")
        end_x = self.screen.width() - self.window_width - 20
        logger.debug(f"Начальная позиция: ({self.start_x}, {self.start_y})")
        logger.debug(f"Конечная позиция: ({end_x}, {self.start_y})")
        animation_pos.setStartValue(QPoint(self.start_x, self.start_y))
        animation_pos.setEndValue(QPoint(end_x, self.start_y))
        animation_pos.setDuration(500)
        animation_pos.setEasingCurve(QEasingCurve.Type.InOutQuad)
        animation_pos.start()
        logger.debug("Анимация запущена")

    def animate_hide(self):
        logger.info("Запуск анимации скрытия окна")
        print("Запуск анимации скрытия окна")
        animation_pos = QPropertyAnimation(self, b"pos")
        end_x = self.screen.width()
        logger.debug(f"Начальная позиция: {self.pos()}")
        logger.debug(f"Конечная позиция: ({end_x}, {self.start_y})")
        animation_pos.setStartValue(self.pos())
        animation_pos.setEndValue(QPoint(end_x, self.start_y))
        animation_pos.setDuration(500)
        animation_pos.setEasingCurve(QEasingCurve.Type.InOutQuad)
        animation_pos.finished.connect(self.hide)
        animation_pos.start()
        logger.debug("Анимация скрытия запущена")

    def update_indicator(self, state):
        logger.debug(f"Обновление индикатора: {state}")
        print(f"Обновление индикатора: {state}")
        if state in self.indicator_images:
            self.indicator_label.setPixmap(self.indicator_images[state])
        else:
            logger.warning(f"Неизвестное состояние индикатора: {state}")
            print(f"Ошибка: Неизвестное состояние индикатора: {state}")
        self.app.processEvents()

    def start_auto_hide_timer(self):
        logger.debug("Запуск таймера автоскрытия")
        print("Запуск таймера автоскрытия")
        QTimer.singleShot(5000, self.animate_hide)

    def toggle_mute(self):
        self.is_muted_manually = not self.is_muted_manually
        mute_icon = os.path.join(self.ASSETS_DIR, "mute_button.png")
        volume_icon = os.path.join(self.ASSETS_DIR, "volume_button.png")
        logger.debug(f"Проверка иконок: mute={mute_icon}, volume={volume_icon}")
        if os.path.exists(mute_icon) and os.path.exists(volume_icon):
            self.mute_button.setIcon(QIcon(mute_icon if self.is_muted_manually else volume_icon))
            logger.info(f"Звук {'выключен' if self.is_muted_manually else 'включён'}")
            print(f"Звук {'выключен' if self.is_muted_manually else 'включён'}")
        else:
            logger.error(f"Одна из иконок не найдена: mute={mute_icon}, volume={volume_icon}")
            print(f"Ошибка: Одна из иконок не найдена")

    def stop_current_command(self):
        self.is_command_running = False
        if self.engine:
            self.engine.stop()
            logger.debug("Остановлен engine")
        self.update_indicator("idle")
        logger.info("Текущая команда остановлена")
        print("Текущая команда остановлена")

    def load_user_prefs(self, user_id):
        logger.debug(f"Загрузка настроек пользователя: {user_id}")
        print(f"Загрузка настроек пользователя: {user_id}")
        with self.json_lock:
            if os.path.exists(self.USER_PREFS_FILE):
                with open(self.USER_PREFS_FILE, "r", encoding="utf-8") as f:
                    self.user_preferences = json.load(f)
            else:
                logger.debug(f"Файл {self.USER_PREFS_FILE} не существует, создаём новый")
                print(f"Файл {self.USER_PREFS_FILE} не существует, создаём новый")
        if user_id not in self.user_preferences:
            self.user_preferences[user_id] = {"language": "ru", "default_music": "youtube", "default_city": "Moscow"}
            with self.json_lock:
                with open(self.USER_PREFS_FILE, "w", encoding="utf-8") as f:
                    json.dump(self.user_preferences, f, ensure_ascii=False, indent=4)
        logger.debug(f"Настройки пользователя: {self.user_preferences[user_id]}")
        print(f"Настройки пользователя: {self.user_preferences[user_id]}")

    def set_preference(self, user_id, key, value):
        logger.info(f"Установка предпочтения для {user_id}: {key} = {value}")
        print(f"Установка предпочтения для {user_id}: {key} = {value}")
        self.load_user_prefs(user_id)
        self.user_preferences[user_id][key] = value
        with self.json_lock:
            with open(self.USER_PREFS_FILE, "w", encoding="utf-8") as f:
                json.dump(self.user_preferences, f, ensure_ascii=False, indent=4)

    async def generate_text_mistral(self, user_id, prompt):
        logger.info(f"Вызов Mistral API с prompt: {prompt}")
        print(f"Вызов Mistral API с prompt: {prompt}")

        # Заглушка для теста
        if not self.USE_MISTRAL_API:
            logger.debug("Используется заглушка вместо Mistral API")
            print("Используется заглушка вместо Mistral API")
            return f"Тестовый ответ: я услышал '{prompt}'. Проверьте API."

        self.load_user_prefs(user_id)
        if user_id not in self.user_histories:
            self.user_histories[user_id] = []
        self.user_histories[user_id].append({"role": "user", "content": prompt})
        if len(self.user_histories[user_id]) > 10:
            self.user_histories[user_id] = self.user_histories[user_id][-10:]
        with self.json_lock:
            with open(self.USER_HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.user_histories, f, ensure_ascii=False, indent=4)
        logger.debug(f"История пользователя сохранена: {self.user_histories[user_id]}")
        print(f"История пользователя сохранена: {self.user_histories[user_id]}")

        api_key = self.MISTRAL_API_KEY
        logger.debug(f"Используемый API ключ: {api_key}")
        print(f"Используемый API ключ: {api_key}")
        headers = {
            "Authorization": f"Bearer {api_key}".strip(),
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        logger.debug(f"Сформированные заголовки: {headers}")
        print(f"Сформированные заголовки: {headers}")
        data = {
            "model": "mistral-small-latest",
            "messages": self.user_histories[user_id],
            "temperature": 0.7,
            "max_tokens": 150
        }
        logger.debug(f"Сформированные данные: {data}")
        print(f"Сформированные данные: {data}")

        # Тестовый синхронный запрос через requests
        logger.debug("Тестируем API через requests (синхронный запрос)")
        print("Тестируем API через requests (синхронный запрос)")
        try:
            sync_response = requests.post(
                "https://api.mistral.ai/v1/chat/completions",
                headers=headers,
                json=data,
                timeout=10
            )
            logger.debug(f"Статус синхронного ответа: {sync_response.status_code}")
            logger.debug(f"Текст синхронного ответа: {sync_response.text}")
            print(f"Синхронный запрос (requests): статус {sync_response.status_code}, ответ: {sync_response.text}")
        except Exception as e:
            logger.error(f"Ошибка синхронного запроса: {e}")
            print(f"Ошибка синхронного запроса: {e}")

        # Асинхронный запрос через aiohttp
        logger.debug("Отправляем асинхронный запрос через aiohttp")
        print("Отправляем асинхронный запрос через aiohttp")
        try:
            async with aiohttp.ClientSession() as session:
                logger.debug("Сессия aiohttp открыта")
                print("Сессия aiohttp открыта")
                async with session.post("https://api.mistral.ai/v1/chat/completions", headers=headers, json=data) as response:
                    logger.debug(f"Статус ответа Mistral API: {response.status}")
                    print(f"Статус ответа Mistral API: {response.status}")
                    response_text = await response.text()
                    logger.debug(f"Полный текст ответа: {response_text}")
                    print(f"Полный текст ответа: {response_text}")
                    if response.status == 200:
                        result = await response.json()
                        assistant_message = result["choices"][0]["message"]["content"]
                        self.user_histories[user_id].append({"role": "assistant", "content": assistant_message})
                        with self.json_lock:
                            with open(self.USER_HISTORY_FILE, "w", encoding="utf-8") as f:
                                json.dump(self.user_histories, f, ensure_ascii=False, indent=4)
                        logger.info(f"Mistral API вернул: {assistant_message}")
                        print(f"Mistral API вернул: {assistant_message}")
                        return assistant_message
                    else:
                        error_message = f"Ошибка API: {response.status} - {response_text}"
                        logger.error(error_message)
                        print(error_message)
                        return f"Ошибка API. Проверьте статус сервиса. {error_message}"
        except aiohttp.ClientConnectionError as e:
            error_message = f"Ошибка соединения с Mistral API: {e}"
            logger.error(error_message)
            print(error_message)
            return f"Нет соединения с сервером. {error_message}"
        except Exception as e:
            error_message = f"Ошибка при вызове Mistral API: {e}"
            logger.error(error_message)
            print(error_message)
            return f"Ошибка при обработке запроса: {error_message}"

    def clean_text_for_speech(self, text):
        logger.debug(f"Очистка текста для речи: {text}")
        cleaned_text = re.sub(r"[^a-zA-Zа-яА-Я0-9\s.,!?]", "", text)
        logger.debug(f"Очищенный текст: {cleaned_text}")
        return cleaned_text

    def get_chrome_driver(self):
        logger.debug("Получение ChromeDriver")
        print("Получение ChromeDriver")
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            logger.info("ChromeDriver успешно запущен")
            print("ChromeDriver успешно запущен")
            return driver
        except Exception as e:
            logger.error(f"Ошибка запуска ChromeDriver: {e}")
            print(f"Ошибка запуска ChromeDriver: {e}")
            return None

    def open_browser(self):
        logger.debug("Попытка открыть браузер")
        print("Попытка открыть браузер")
        try:
            webbrowser.open("https://www.google.com")
            logger.info("Браузер открыт")
            print("Браузер открыт")
            return "Браузер открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия браузера: {e}")
            print(f"Ошибка открытия браузера: {e}")
            return f"Ошибка открытия браузера: {e}"

    def open_notepad(self):
        logger.debug("Попытка открыть блокнот")
        print("Попытка открыть блокнот")
        try:
            subprocess.Popen("notepad.exe")
            logger.info("Блокнот открыт")
            print("Блокнот открыт")
            return "Блокнот открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия блокнота: {e}")
            print(f"Ошибка открытия блокнота: {e}")
            return f"Ошибка открытия блокнота: {e}"

    def open_explorer(self):
        logger.debug("Попытка открыть проводник")
        print("Попытка открыть проводник")
        try:
            subprocess.Popen("explorer.exe")
            logger.info("Проводник открыт")
            print("Проводник открыт")
            return "Проводник открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия проводника: {e}")
            print(f"Ошибка открытия проводника: {e}")
            return f"Ошибка открытия проводника: {e}"

    def open_calculator(self):
        logger.debug("Попытка открыть калькулятор")
        print("Попытка открыть калькулятор")
        try:
            subprocess.Popen("calc.exe")
            logger.info("Калькулятор открыт")
            print("Калькулятор открыт")
            return "Калькулятор открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия калькулятора: {e}")
            print(f"Ошибка открытия калькулятора: {e}")
            return f"Ошибка открытия калькулятора: {e}"

    def open_discord(self):
        logger.debug("Попытка открыть Discord")
        print("Попытка открыть Discord")
        try:
            subprocess.Popen(r"C:\Users\%USERNAME%\AppData\Local\Discord\app-*.exe", shell=True)
            logger.info("Discord открыт")
            print("Discord открыт")
            return "Discord открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия Discord: {e}")
            print(f"Ошибка открытия Discord: {e}")
            return f"Ошибка открытия Discord: {e}"

    def open_telegram(self):
        logger.debug("Попытка открыть Telegram")
        print("Попытка открыть Telegram")
        try:
            subprocess.Popen(r"C:\Users\%USERNAME%\AppData\Roaming\Telegram Desktop\Telegram.exe", shell=True)
            logger.info("Telegram открыт")
            print("Telegram открыт")
            return "Telegram открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия Telegram: {e}")
            print(f"Ошибка открытия Telegram: {e}")
            return f"Ошибка открытия Telegram: {e}"

    def open_whatsapp(self):
        logger.debug("Попытка открыть WhatsApp")
        print("Попытка открыть WhatsApp")
        try:
            subprocess.Popen(r"C:\Program Files\WindowsApps\5319275A.WhatsApp_*\\WhatsApp.exe", shell=True)
            logger.info("WhatsApp открыт")
            print("WhatsApp открыт")
            return "WhatsApp открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия WhatsApp: {e}")
            print(f"Ошибка открытия WhatsApp: {e}")
            return f"Ошибка открытия WhatsApp: {e}"

    def open_skype(self):
        logger.debug("Попытка открыть Skype")
        print("Попытка открыть Skype")
        try:
            subprocess.Popen(r"C:\Program Files (x86)\Microsoft\Skype for Desktop\Skype.exe", shell=True)
            logger.info("Skype открыт")
            print("Skype открыт")
            return "Skype открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия Skype: {e}")
            print(f"Ошибка открытия Skype: {e}")
            return f"Ошибка открытия Skype: {e}"

    def open_vscode(self):
        logger.debug("Попытка открыть VS Code")
        print("Попытка открыть VS Code")
        try:
            subprocess.Popen(r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code\Code.exe", shell=True)
            logger.info("Visual Studio Code открыт")
            print("Visual Studio Code открыт")
            return "Visual Studio Code открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия VS Code: {e}")
            print(f"Ошибка открытия VS Code: {e}")
            return f"Ошибка открытия VS Code: {e}"

    def open_word(self):
        logger.debug("Попытка открыть Word")
        print("Попытка открыть Word")
        try:
            subprocess.Popen(r"C:\Program Files (x86)\Microsoft Office\root\Office*\WINWORD.EXE", shell=True)
            logger.info("Microsoft Word открыт")
            print("Microsoft Word открыт")
            return "Microsoft Word открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия Word: {e}")
            print(f"Ошибка открытия Word: {e}")
            return f"Ошибка открытия Word: {e}"

    def open_excel(self):
        logger.debug("Попытка открыть Excel")
        print("Попытка открыть Excel")
        try:
            subprocess.Popen(r"C:\Program Files (x86)\Microsoft Office\root\Office*\EXCEL.EXE", shell=True)
            logger.info("Microsoft Excel открыт")
            print("Microsoft Excel открыт")
            return "Microsoft Excel открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия Excel: {e}")
            print(f"Ошибка открытия Excel: {e}")
            return f"Ошибка открытия Excel: {e}"

    def open_powerpoint(self):
        logger.debug("Попытка открыть PowerPoint")
        print("Попытка открыть PowerPoint")
        try:
            subprocess.Popen(r"C:\Program Files (x86)\Microsoft Office\root\Office*\POWERPNT.EXE", shell=True)
            logger.info("Microsoft PowerPoint открыт")
            print("Microsoft PowerPoint открыт")
            return "Microsoft PowerPoint открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия PowerPoint: {e}")
            print(f"Ошибка открытия PowerPoint: {e}")
            return f"Ошибка открытия PowerPoint: {e}"

    def open_paint(self):
        logger.debug("Попытка открыть Paint")
        print("Попытка открыть Paint")
        try:
            subprocess.Popen("mspaint.exe")
            logger.info("Paint открыт")
            print("Paint открыт")
            return "Paint открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия Paint: {e}")
            print(f"Ошибка открытия Paint: {e}")
            return f"Ошибка открытия Paint: {e}"

    def open_cmd(self):
        logger.debug("Попытка открыть командную строку")
        print("Попытка открыть командную строку")
        try:
            subprocess.Popen("cmd.exe")
            logger.info("Командная строка открыта")
            print("Командная строка открыта")
            return "Командная строка открыта."
        except Exception as e:
            logger.error(f"Ошибка открытия командной строки: {e}")
            print(f"Ошибка открытия командной строки: {e}")
            return f"Ошибка открытия командной строки: {e}"

    def open_task_manager(self):
        logger.debug("Попытка открыть диспетчер задач")
        print("Попытка открыть диспетчер задач")
        try:
            subprocess.Popen("taskmgr.exe")
            logger.info("Диспетчер задач открыт")
            print("Диспетчер задач открыт")
            return "Диспетчер задач открыт."
        except Exception as e:
            logger.error(f"Ошибка открытия диспетчера задач: {e}")
            print(f"Ошибка открытия диспетчера задач: {e}")
            return f"Ошибка открытия диспетчера задач: {e}"
        
    def on_quit(self, icon, item):
        logger.info("Выход из приложения")
        print("Выход из приложения")
        self.icon.stop()
        self.app.quit()


    def open_settings(self):
        logger.debug("Попытка открыть настройки")
        print("Попытка открыть настройки")
        try:
            subprocess.Popen("ms-settings:")
            logger.info("Настройки открыты")
            print("Настройки открыты")
            return "Настройки открыты."
        except Exception as e:
            logger.error(f"Ошибка открытия настроек: {e}")
            print(f"Ошибка открытия настроек: {e}")
            return f"Ошибка открытия настроек: {e}"
        
    def on_show(self, icon, item):
        logger.info("Отображение окна")
        print("Отображение окна")
        self.show()
        self.raise_()

    def run(self):
         logger.info("Запуск приложения")
         print("Запуск приложения")
         self.animate_show()  # Показать интерфейс с анимацией
         self.start_listening()  # Запустить прослушивание голосовых команд
         if self.icon:  # Если есть системный трей
             threading.Thread(target=self.icon.run, daemon=True).start()  # Запустить трей в отдельном потоке

    def open_control_panel(self):
        logger.debug("Попытка открыть панель управления")
        print("Попытка открыть панель управления")
        try:
            subprocess.Popen("control.exe")
            logger.info("Панель управления открыта")
            print("Панель управления открыта")
            return "Панель управления открыта."
        except Exception as e:
            logger.error(f"Ошибка открытия панели управления: {e}")
            print(f"Ошибка открытия панели управления: {e}")
            return f"Ошибка открытия панели управления: {e}"

    def close_application(self, app_name):
        logger.debug(f"Попытка закрыть приложение: {app_name}")
        print(f"Попытка закрыть приложение: {app_name}")
        try:
            os.system(f"taskkill /IM {app_name}.exe /F")
            logger.info(f"Приложение {app_name} закрыто")
            print(f"Приложение {app_name} закрыто")
            return f"Приложение {app_name} закрыто."
        except Exception as e:
            logger.error(f"Ошибка закрытия приложения {app_name}: {e}")
            print(f"Ошибка закрытия приложения {app_name}: {e}")
            return f"Ошибка закрытия {app_name}: {e}"

    def play_music_on_youtube(self):
        logger.debug("Попытка воспроизвести музыку на YouTube")
        print("Попытка воспроизвести музыку на YouTube")
        driver = self.get_chrome_driver()
        if not driver:
            logger.error("Не удалось запустить ChromeDriver для YouTube")
            print("Не удалось запустить ChromeDriver для YouTube")
            return "Ошибка запуска ChromeDriver для YouTube."
        try:
            driver.get("https://www.youtube.com")
            time.sleep(2)
            search_box = driver.find_element(By.NAME, "search_query")
            search_box.send_keys("music")
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)
            videos = driver.find_elements(By.XPATH, '//ytd-video-renderer')
            if videos:
                videos[0].click()
                self.current_track_index = 0
                self.track_history.append("music")
                logger.info("Воспроизвожу музыку на YouTube")
                print("Воспроизвожу музыку на YouTube")
                return "Воспроизвожу музыку на YouTube."
            else:
                logger.warning("Видео не найдено на YouTube")
                print("Видео не найдено на YouTube")
                return "Не удалось найти музыку."
        except Exception as e:
            logger.error(f"Ошибка воспроизведения на YouTube: {e}")
            print(f"Ошибка воспроизведения на YouTube: {e}")
            return f"Ошибка воспроизведения на YouTube: {e}"
        finally:
            driver.quit()

    def play_music_on_soundcloud(self):
        logger.debug("Попытка воспроизвести музыку на SoundCloud")
        print("Попытка воспроизвести музыку на SoundCloud")
        driver = self.get_chrome_driver()
        if not driver:
            logger.error("Не удалось запустить ChromeDriver для SoundCloud")
            print("Не удалось запустить ChromeDriver для SoundCloud")
            return "Ошибка запуска ChromeDriver для SoundCloud."
        try:
            driver.get("https://soundcloud.com")
            time.sleep(2)
            search_box = driver.find_element(By.XPATH, '//input[@placeholder="Search for artists, tracks, or podcasts"]')
            search_box.send_keys("music")
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)
            tracks = driver.find_elements(By.XPATH, '//li[@class="searchList__item"]')
            if tracks:
                tracks[0].click()
                self.current_track_index = 0
                self.track_history.append("music")
                logger.info("Воспроизвожу музыку на SoundCloud")
                print("Воспроизвожу музыку на SoundCloud")
                return "Воспроизвожу музыку на SoundCloud."
            else:
                logger.warning("Треки не найдены на SoundCloud")
                print("Треки не найдены на SoundCloud")
                return "Не удалось найти музыку."
        except Exception as e:
            logger.error(f"Ошибка воспроизведения на SoundCloud: {e}")
            print(f"Ошибка воспроизведения на SoundCloud: {e}")
            return f"Ошибка воспроизведения на SoundCloud: {e}"
        finally:
            driver.quit()

    def play_music_on_yandex_music(self):
        logger.debug("Попытка воспроизвести музыку на Яндекс.Музыке")
        print("Попытка воспроизвести музыку на Яндекс.Музыке")
        driver = self.get_chrome_driver()
        if not driver:
            logger.error("Не удалось запустить ChromeDriver для Яндекс.Музыки")
            print("Не удалось запустить ChromeDriver для Яндекс.Музыки")
            return "Ошибка запуска ChromeDriver для Яндекс.Музыки."
        try:
            driver.get("https://music.yandex.com")
            time.sleep(2)
            search_box = driver.find_element(By.XPATH, '//input[@placeholder="Поиск"]')
            search_box.send_keys("music")
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)
            tracks = driver.find_elements(By.XPATH, '//div[@class="track__title"]')
            if tracks:
                tracks[0].click()
                self.current_track_index = 0
                self.track_history.append("music")
                logger.info("Воспроизвожу музыку на Яндекс.Музыке")
                print("Воспроизвожу музыку на Яндекс.Музыке")
                return "Воспроизвожу музыку на Яндекс.Музыке."
            else:
                logger.warning("Треки не найдены на Яндекс.Музыке")
                print("Треки не найдены на Яндекс.Музыке")
                return "Не удалось найти музыку."
        except Exception as e:
            logger.error(f"Ошибка воспроизведения на Яндекс.Музыке: {e}")
            print(f"Ошибка воспроизведения на Яндекс.Музыке: {e}")
            return f"Ошибка воспроизведения на Яндекс.Музыке: {e}"
        finally:
            driver.quit()

    def play_music_on_spotify(self):
        logger.debug("Попытка воспроизвести музыку на Spotify")
        print("Попытка воспроизвести музыку на Spotify")
        driver = self.get_chrome_driver()
        if not driver:
            logger.error("Не удалось запустить ChromeDriver для Spotify")
            print("Не удалось запустить ChromeDriver для Spotify")
            return "Ошибка запуска ChromeDriver для Spotify."
        try:
            driver.get("https://open.spotify.com")
            time.sleep(2)
            search_box = driver.find_element(By.XPATH, '//input[@data-testid="search-input"]')
            search_box.send_keys("music")
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)
            tracks = driver.find_elements(By.XPATH, '//div[@data-testid="tracklist-row"]')
            if tracks:
                tracks[0].click()
                self.current_track_index = 0
                self.track_history.append("music")
                logger.info("Воспроизвожу музыку на Spotify")
                print("Воспроизвожу музыку на Spotify")
                return "Воспроизвожу музыку на Spotify."
            else:
                logger.warning("Треки не найдены на Spotify")
                print("Треки не найдены на Spotify")
                return "Не удалось найти музыку."
        except Exception as e:
            logger.error(f"Ошибка воспроизведения на Spotify: {e}")
            print(f"Ошибка воспроизведения на Spotify: {e}")
            return f"Ошибка воспроизведения на Spotify: {e}"
        finally:
            driver.quit()

    def pause_music(self):
        logger.debug("Попытка приостановить музыку")
        print("Попытка приостановить музыку")
        self.safe_press("playpause")
        logger.info("Музыка приостановлена")
        print("Музыка приостановлена")
        return "Музыка приостановлена."

    def resume_music(self):
        logger.debug("Попытка возобновить музыку")
        print("Попытка возобновить музыку")
        self.safe_press("playpause")
        logger.info("Музыка возобновлена")
        print("Музыка возобновлена")
        return "Музыка возобновлена."

    def next_track(self):
        logger.debug("Попытка переключиться на следующий трек")
        print("Попытка переключиться на следующий трек")
        self.safe_press("nexttrack")
        self.current_track_index += 1
        self.track_history.append(f"Track {self.current_track_index}")
        logger.info("Переключён на следующий трек")
        print("Переключён на следующий трек")
        return "Следующий трек."

    def previous_track(self):
        logger.debug("Попытка переключиться на предыдущий трек")
        print("Попытка переключиться на предыдущий трек")
        self.safe_press("prevtrack")
        if self.current_track_index > 0:
            self.current_track_index -= 1
        logger.info("Переключён на предыдущий трек")
        print("Переключён на предыдущий трек")
        return "Предыдущий треk."

    def adjust_app_volume(self, app_name, increase=True):
        logger.debug(f"Попытка {'увеличить' if increase else 'уменьшить'} громкость для {app_name}")
        print(f"Попытка {'увеличить' if increase else 'уменьшить'} громкость для {app_name}")
        try:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.name().lower().find(app_name.lower()) != -1:
                    volume = session.SimpleAudioVolume
                    current_volume = volume.GetMasterVolume()
                    new_volume = min(1.0, current_volume + 0.1) if increase else max(0.0, current_volume - 0.1)
                    volume.SetMasterVolume(new_volume, None)
                    logger.info(f"Громкость для {app_name} {'увеличена' if increase else 'уменьшена'}")
                    print(f"Громкость для {app_name} {'увеличена' if increase else 'уменьшена'}")
                    return f"Громкость для {app_name} {'увеличена' if increase else 'уменьшена'}."
            logger.warning(f"Приложение {app_name} не найдено для регулировки громкости")
            print(f"Ошибка: Приложение {app_name} не найдено для регулировки громкости")
            return f"Приложение {app_name} не найдено."
        except Exception as e:
            logger.error(f"Ошибка регулировки громкости приложения {app_name}: {e}")
            print(f"Ошибка регулировки громкости приложения {app_name}: {e}")
            return f"Ошибка регулировки громкости: {e}"

    def adjust_volume(self, increase=True):
        logger.debug(f"Попытка {'увеличить' if increase else 'уменьшить'} системную громкость")
        print(f"Попытка {'увеличить' if increase else 'уменьшить'} системную громкость")
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            current_volume = volume.GetMasterVolumeLevelScalar()
            new_volume = min(1.0, current_volume + 0.1) if increase else max(0.0, current_volume - 0.1)
            volume.SetMasterVolumeLevelScalar(new_volume, None)
            logger.info(f"Громкость {'увеличена' if increase else 'уменьшена'}")
            print(f"Громкость {'увеличена' if increase else 'уменьшена'}")
            return "Громкость " + ("увеличена" if increase else "уменьшена")
        except Exception as e:
            logger.error(f"Ошибка регулировки громкости: {e}")
            print(f"Ошибка регулировки громкости: {e}")
            return f"Ошибка регулировки громкости: {e}"

    def mute_system(self):
        logger.debug("Попытка выключить звук")
        print("Попытка выключить звук")
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            volume.SetMute(1, None)
            logger.info("Звук выключен")
            print("Звук выключен")
            return "Звук выключен."
        except Exception as e:
            logger.error(f"Ошибка выключения звука: {e}")
            print(f"Ошибка выключения звука: {e}")
            return f"Ошибка выключения звука: {e}"

    def unmute_system(self):
        logger.debug("Попытка включить звук")
        print("Попытка включить звук")
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            volume.SetMute(0, None)
            logger.info("Звук включён")
            print("Звук включён")
            return "Звук включён."
        except Exception as e:
            logger.error(f"Ошибка включения звука: {e}")
            print(f"Ошибка включения звука: {e}")
            return f"Ошибка включения звука: {e}"

    def get_weather(self, city):
        logger.debug(f"Попытка получить погоду для {city}")
        print(f"Попытка получить погоду для {city}")
        url = f"https://wttr.in/{city}?format=%C+%t"
        try:
            response = subprocess.run(["curl", "-s", url], capture_output=True, text=True)
            weather = response.stdout.strip()
            logger.info(f"Погода в {city}: {weather}")
            print(f"Погода в {city}: {weather}")
            return f"Погода в {city}: {weather}"
        except Exception as e:
            logger.error(f"Ошибка получения погоды: {e}")
            print(f"Ошибка получения погоды: {e}")
            return f"Ошибка получения погоды: {e}"

    def get_time(self):
        logger.debug("Попытка получить текущее время")
        print("Попытка получить текущее время")
        current_time = time.strftime('%H:%M:%S')
        logger.info(f"Текущее время: {current_time}")
        print(f"Текущее время: {current_time}")
        return f"Текущее время: {current_time}"

    def take_screenshot(self):
        logger.debug("Попытка сделать скриншот")
        print("Попытка сделать скриншot")
        try:
            screenshot = pyautogui.screenshot()
            screenshot.save("screenshot.png")
            logger.info("Скриншот сохранён")
            print("Скриншот сохранён")
            return "Скриншот сохранён как screenshot.png"
        except Exception as e:
            logger.error(f"Ошибка сохранения скриншота: {e}")
            print(f"Ошибка сохранения скриншота: {e}")
            return f"Ошибка сохранения скриншота: {e}"

    def set_reminder(self, minutes):
        logger.debug(f"Попытка установить напоминание на {minutes} минут")
        print(f"Попытка установить напоминание на {minutes} минут")
        try:
            seconds = int(minutes) * 60
            QTimer.singleShot(seconds * 1000, lambda: self.label.setText("Напоминание!"))
            logger.info(f"Напоминание установлено через {minutes} минут")
            print(f"Напоминание установлено через {minutes} минут")
            return f"Напоминание установлено через {minutes} минут."
        except Exception as e:
            logger.error(f"Ошибка установки напоминания: {e}")
            print(f"Ошибка установки напоминания: {e}")
            return f"Ошибка установки напоминания: {e}"

    def minimize_window(self):
        logger.debug("Попытка свернуть окно")
        print("Попытка свернуть окно")
        try:
            pyautogui.hotkey("win", "down")
            logger.info("Окно свёрнуто")
            print("Окно свёрнуто")
            return "Окно свёрнуто."
        except Exception as e:
            logger.error(f"Ошибка минимизации окна: {e}")
            print(f"Ошибка минимизации окна: {e}")
            return f"Ошибка минимизации окна: {e}"

    def maximize_window(self):
        logger.debug("Попытка развернуть окно")
        print("Попытка развернуть окно")
        try:
            pyautogui.hotkey("win", "up")
            logger.info("Окно развёрнуто")
            print("Окно развёрнуто")
            return "Окно развёрнуто."
        except Exception as e:
            logger.error(f"Ошибка максимизации окна: {e}")
            print(f"Ошибка максимизации окна: {e}")
            return f"Ошибка максимизации окна: {e}"

    def input_text(self, text):
        logger.debug(f"Попытка ввести текст: {text}")
        print(f"Попытка ввести текст: {text}")
        try:
            pyautogui.write(text)
            pyautogui.press("enter")
            logger.info(f"Введён текст: {text}")
            print(f"Введён текст: {text}")
            return f"Введён текст: {text}"
        except Exception as e:
            logger.error(f"Ошибка ввода текста: {e}")
            print(f"Ошибка ввода текста: {e}")
            return f"Ошибка ввода текста: {e}"

    def set_queue(self, queue_number):
        logger.debug(f"Попытка установить очередь: {queue_number}")
        print(f"Попытка установить очередь: {queue_number}")
        try:
            self.CURRENT_QUEUE = int(queue_number)
            logger.info(f"Очередь установлена на {self.CURRENT_QUEUE}")
            print(f"Очередь установлена на {self.CURRENT_QUEUE}")
            return f"Очередь установлена на {self.CURRENT_QUEUE}."
        except ValueError:
            logger.error("Ошибка: указано некорректное значение очереди")
            print("Ошибка: указано некорректное значение очереди")
            return "Пожалуйста, укажите число для очереди."

    def check_queue(self):
        logger.debug(f"Проверка очереди: {self.CURRENT_QUEUE}")
        print(f"Проверка очереди: {self.CURRENT_QUEUE}")
        logger.info(f"Текущая очередь: {self.CURRENT_QUEUE}")
        print(f"Текущая очередь: {self.CURRENT_QUEUE}")
        return f"Текущая очередь: {self.CURRENT_QUEUE}"

    def find_and_play_youtube(self, query, play_type=None, number=None, title=None):
        logger.debug(f"Попытка найти и воспроизвести на YouTube: {query}, тип: {play_type}, номер: {number}, заголовок: {title}")
        print(f"Попытка найти и воспроизвести на YouTube: {query}, тип: {play_type}, номер: {number}, заголовок: {title}")
        driver = self.get_chrome_driver()
        if not driver:
            logger.error("Не удалось запустить ChromeDriver для YouTube")
            print("Не удалось запустить ChromeDriver для YouTube")
            return "Ошибка запуска ChromeDriver для YouTube."
        try:
            driver.get("https://www.youtube.com")
            time.sleep(2)
            search_box = driver.find_element(By.NAME, "search_query")
            search_box.send_keys(query)
            search_box.send_keys(Keys.RETURN)
            time.sleep(2)
            videos = driver.find_elements(By.XPATH, '//ytd-video-renderer')
            if not videos:
                logger.warning(f"Видео по запросу '{query}' не найдено")
                print(f"Видео по запросу '{query}' не найдено")
                return "Видео не найдено."
            if play_type == "first":
                videos[0].click()
                self.current_track_index = 0
                self.track_history.append(query)
            elif play_type == "number" and number is not None:
                if 0 < number <= len(videos):
                    videos[number - 1].click()
                    self.current_track_index = number - 1
                    self.track_history.append(query)
                else:
                    logger.warning(f"Номер видео {number} недоступен")
                    print(f"Номер видео {number} недоступен")
                    return "Указанный номер видео недоступен."
            elif play_type == "title" and title:
                for i, video in enumerate(videos):
                    video_title = video.find_element(By.ID, "video-title").text.lower()
                    if title.lower() in video_title:
                        video.click()
                        self.current_track_index = i
                        self.track_history.append(query)
                        break
                else:
                    logger.warning(f"Видео с названием '{title}' не найдено")
                    print(f"Видео с названием '{title}' не найдено")
                    return f"Видео с названием '{title}' не найдено."
            logger.info(f"Воспроизвожу '{query}' на YouTube")
            print(f"Воспроизвожу '{query}' на YouTube")
            return f"Воспроизвожу {query} на YouTube."
        except Exception as e:
            logger.error(f"Ошибка поиска/воспроизведения видео на YouTube: {e}")
            print(f"Ошибка поиска/воспроизведения видео на YouTube: {e}")
            return f"Ошибка поиска/воспроизведения видео: {e}"
        finally:
            driver.quit()

    def scroll_page(self, direction="down"):
        logger.debug(f"Попытка прокрутки страницы: {direction}")
        print(f"Попытка прокрутки страницы: {direction}")
        try:
            if direction == "down":
                pyautogui.scroll(-500)
                logger.info("Прокручено вниз")
                print("Прокручено вниз")
                return "Прокручено вниз."
            else:
                pyautogui.scroll(500)
                logger.info("Прокручено вверх")
                print("Прокручено вверх")
                return "Прокручено вверх."
        except Exception as e:
            logger.error(f"Ошибка прокрутки страницы: {e}")
            print(f"Ошибка прокрутки страницы: {e}")
            return f"Ошибка прокрутки страницы: {e}"

    def shutdown_computer(self):
        logger.debug("Попытка выключить компьютер")
        print("Попытка выключить компьютер")
        try:
            os.system("shutdown /s /t 1")
            logger.info("Компьютер выключается")
            print("Компьютер выключается")
            return "Компьютер выключается."
        except Exception as e:
            logger.error(f"Ошибка выключения компьютера: {e}")
            print(f"Ошибка выключения компьютера: {e}")
            return f"Ошибка выключения компьютера: {e}"

    def restart_computer(self):
        logger.debug("Попытка перезагрузить компьютер")
        print("Попытка перезагрузить компьютер")
        try:
            os.system("shutdown /r /t 1")
            logger.info("Компьютер перезагружается")
            print("Компьютер перезагружается")
            return "Компьютер перезагружается."
        except Exception as e:
            logger.error(f"Ошибка перезагрузки компьютера: {e}")
            print(f"Ошибка перезагрузки компьютера: {e}")
            return f"Ошибка перезагрузки компьютера: {e}"

    def lock_computer(self):
        logger.debug("Попытка заблокировать компьютер")
        print("Попытка заблокировать компьютер")
        try:
            os.system("rundll32.exe user32.dll,LockWorkStation")
            logger.info("Компьютер заблокирован")
            print("Компьютер заблокирован")
            return "Компьютер заблокирован."
        except Exception as e:
            logger.error(f"Ошибка блокировки компьютера: {e}")
            print(f"Ошибка блокировки компьютера: {e}")
            return f"Ошибка блокировки компьютера: {e}"

    def new_tab(self):
        logger.debug("Попытка открыть новую вкладку")
        print("Попытка открыть новую вкладку")
        try:
            pyautogui.hotkey("ctrl", "t")
            logger.info("Новая вкладка открыта")
            print("Новая вкладка открыта")
            return "Новая вкладка открыта."
        except Exception as e:
            logger.error(f"Ошибка открытия новой вкладки: {e}")
            print(f"Ошибка открытия новой вкладки: {e}")
            return f"Ошибка открытия новой вкладки: {e}"

    def close_tab(self):
        logger.debug("Попытка закрыть вкладку")
        print("Попытка закрыть вкладку")
        try:
            pyautogui.hotkey("ctrl", "w")
            logger.info("Вкладка закрыта")
            print("Вкладка закрыта")
            return "Вкладка закрыта."
        except Exception as e:
            logger.error(f"Ошибка закрытия вкладки: {e}")
            print(f"Ошибка закрытия вкладки: {e}")
            return f"Ошибка закрытия вкладки: {e}"

    def switch_tab(self):
        logger.debug("Попытка переключиться на следующую вкладку")
        print("Попытка переключиться на следующую вкладку")
        try:
            pyautogui.hotkey("ctrl", "tab")
            logger.info("Переключено на следующую вкладку")
            print("Переключено на следующую вкладку")
            return "Переключено на следующую вкладку."
        except Exception as e:
            logger.error(f"Ошибка переключения вкладок: {e}")
            print(f"Ошибка переключения вкладок: {e}")
            return f"Ошибка переключения вкладок: {e}"

    def refresh_page(self):
        logger.debug("Попытка обновить страницу")
        print("Попытка обновить страницу")
        try:
            pyautogui.press("f5")
            logger.info("Страница обновлена")
            print("Страница обновлена")
            return "Страница обновлена."
        except Exception as e:
            logger.error(f"Ошибка обновления страницы: {e}")
            print(f"Ошибка обновления страницы: {e}")
            return f"Ошибка обновления страницы: {e}"

    def go_back(self):
        logger.debug("Попытка вернуться назад")
        print("Попытка вернуться назад")
        try:
            pyautogui.hotkey("alt", "left")
            logger.info("Вернулся назад")
            print("Вернулся назад")
            return "Вернулся назад."
        except Exception as e:
            logger.error(f"Ошибка перехода назад: {e}")
            print(f"Ошибка перехода назад: {e}")
            return f"Ошибка перехода назад: {e}"

    def go_forward(self):
        logger.debug("Попытка перейти вперёд")
        print("Попытка перейти вперёд")
        try:
            pyautogui.hotkey("alt", "right")
            logger.info("Перешёл вперёд")
            print("Перешёл вперёд")
            return "Перешёл вперёд."
        except Exception as e:
            logger.error(f"Ошибка перехода вперёд: {e}")
            print(f"Ошибка перехода вперёд: {e}")
            return f"Ошибка перехода вперёд: {e}"

    def click_mouse(self, position="center"):
        logger.debug(f"Попытка клика мышью в позиции: {position}")
        print(f"Попытка клика мышью в позиции: {position}")
        try:
            if position == "left":
                pyautogui.moveTo(100, pyautogui.position()[1])
            elif position == "right":
                pyautogui.moveTo(self.screen.width() - 100, pyautogui.position()[1])
            elif position == "center":
                pyautogui.moveTo(self.screen.width() // 2, self.screen.height() // 2)
            pyautogui.click()
            logger.info(f"Клик выполнен в позиции: {position}")
            print(f"Клик выполнен в позиции: {position}")
            return f"Клик выполнен {position}."
        except Exception as e:
            logger.error(f"Ошибка клика мышью: {e}")
            print(f"Ошибка клика мышью: {e}")
            return f"Ошибка клика: {e}"

    def double_click_mouse(self, position="center"):
        logger.debug(f"Попытка двойного клика мышью в позиции: {position}")
        print(f"Попытка двойного клика мышью в позиции: {position}")
        try:
            if position == "left":
                pyautogui.moveTo(100, pyautogui.position()[1])
            elif position == "right":
                pyautogui.moveTo(self.screen.width() - 100, pyautogui.position()[1])
            elif position == "center":
                pyautogui.moveTo(self.screen.width() // 2, self.screen.height() // 2)
            pyautogui.doubleClick()
            logger.info(f"Двойной клик выполнен в позиции: {position}")
            print(f"Двойной клик выполнен в позиции: {position}")
            return f"Двойной клик выполнен {position}."
        except Exception as e:
            logger.error(f"Ошибка двойного клика: {e}")
            print(f"Ошибка двойного клика: {e}")
            return f"Ошибка двойного клика: {e}"

    def right_click_mouse(self, position="center"):
        logger.debug(f"Попытка правого клика мышью в позиции: {position}")
        print(f"Попытка правого клика мышью в позиции: {position}")
        try:
            if position == "left":
                pyautogui.moveTo(100, pyautogui.position()[1])
            elif position == "right":
                pyautogui.moveTo(self.screen.width() - 100, pyautogui.position()[1])
            elif position == "center":
                pyautogui.moveTo(self.screen.width() // 2, self.screen.height() // 2)
            pyautogui.rightClick()
            logger.info(f"Правый клик выполнен в позиции: {position}")
            print(f"Правый клик выполнен в позиции: {position}")
            return f"Правый клик выполнен {position}."
        except Exception as e:
            logger.error(f"Ошибка правого клика: {e}")
            print(f"Ошибка правого клика: {e}")
            return f"Ошибка правого клика: {e}"

    def type_password(self):
        logger.debug("Попытка ввести пароль")
        print("Попытка ввести пароль")
        try:
            password = "your_password"  # Замените на ваш пароль
            pyautogui.write(password)
            pyautogui.press("enter")
            logger.info("Пароль введён")
            print("Пароль введён")
            return "Пароль введён."
        except Exception as e:
            logger.error(f"Ошибка ввода пароля: {e}")
            print(f"Ошибка ввода пароля: {e}")
            return f"Ошибка ввода пароля: {e}"

    def type_email(self):
        logger.debug("Попытка ввести email")
        print("Попытка ввести email")
        try:
            email = "your_email@example.com"  # Замените на ваш email
            pyautogui.write(email)
            pyautogui.press("enter")
            logger.info("Email введён")
            print("Email введён")
            return "Email введён."
        except Exception as e:
            logger.error(f"Ошибка ввода email: {e}")
            print(f"Ошибка ввода email: {e}")
            return f"Ошибка ввода email: {e}"

    def clear_text(self):
        logger.debug("Попытка очистить текст")
        print("Попытка очистить текст")
        try:
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            logger.info("Текст очищен")
            print("Текст очищен")
            return "Текст очищен."
        except Exception as e:
            logger.error(f"Ошибка очистки текста: {e}")
            print(f"Ошибка очистки текста: {e}")
            return f"Ошибка очистки текста: {e}"

    async def find_and_click_button(self, button_text):
        """Находит и нажимает кнопку с заданным текстом на экране."""
        logger.info(f"Поиск кнопки с текстом: {button_text}")
        print(f"Поиск кнопки с текстом: {button_text}")

        screenshot_path = "screen_analysis.png"

        try:
            # Делаем скриншот
            screenshot = pyautogui.screenshot()
            screenshot.save(screenshot_path)
            logger.info(f"Скриншот сохранен: {screenshot_path}")

            # Читаем изображение в бинарный формат
            with open(screenshot_path, "rb") as image_file:
                image_data = image_file.read()

            # Отправляем изображение в Pixtral API
            headers = {
                "Authorization": f"Bearer {self.MISTRAL_API_KEY}",
                "Content-Type": "application/octet-stream"
            }
            
            data = {"query": f"Найди кнопку с текстом '{button_text}' и верни её координаты."}

            async with aiohttp.ClientSession() as session:
                async with session.post(
                    "https://api.pixtral.ai/v1/analyze",  # Уточните реальный URL Pixtral
                    headers=headers,
                    data=image_data,
                    params=data
                ) as response:

                    if response.status == 200:
                        result = await response.json()
                        button_coordinates = result.get("coordinates", None)

                        if button_coordinates:
                            x, y = button_coordinates["x"], button_coordinates["y"]
                            logger.info(f"Координаты кнопки: x={x}, y={y}")
                            print(f"Координаты кнопки: x={x}, y={y}")

                            # Кликаем по найденной кнопке
                            pyautogui.moveTo(x, y, duration=0.2)
                            pyautogui.click()
                            return f"Нажал на кнопку '{button_text}'."
                        else:
                            logger.warning("Кнопка не найдена")
                            print("Кнопка не найдена")
                            return "Не удалось найти кнопку."

                    else:
                        error_message = f"Ошибка Pixtral API: {response.status} - {await response.text()}"
                        logger.error(error_message)
                        print(error_message)
                        return "Ошибка анализа экрана."

        except Exception as e:
            logger.error(f"Ошибка при поиске кнопки: {e}")
            print(f"Ошибка при поиске кнопки: {e}")
            return "Ошибка обработки экрана."


    async def handle_command(self, user_id, command):
        logger.info(f"Получена команда: {command}")
        print(f"Получена команда: {command}")
        current_time = time.time()

        if user_id in self.last_command_time and current_time - self.last_command_time[user_id] < self.command_cooldown:
            response = "Слишком частые команды. Пожалуйста, подождите."
            logger.warning("Команда отклонена: слишком частые команды")
            print("Команда отклонена: слишком частые команды")
            self.label.setText(response)
            if not self.is_muted_manually:
                self.engine.say(self.clean_text_for_speech(response))
                self.engine.runAndWait()
            self.update_indicator("error")
            self.start_auto_hide_timer()
            return

        self.last_command_time[user_id] = current_time
        self.command_usage[command] = self.command_usage.get(command, 0) + 1
        self.is_command_running = True
        self.update_indicator("processing")
        command_lower = command.lower()

        found_action = None
        found_object = None
        additional_info = None

        for action, synonyms_list in self.synonyms["actions"].items():
            if any(syn in command_lower for syn in synonyms_list):
                found_action = action
                break

        for obj, synonyms_list in self.synonyms["objects"].items():
            if any(syn in command_lower for syn in synonyms_list):
                found_object = obj
                break

        for add, synonyms_list in self.synonyms["additional"].items():
            if any(syn in command_lower for syn in synonyms_list):
                additional_info = add
                break

        logger.debug(f"Найденное действие: {found_action}, объект: {found_object}, доп. информация: {additional_info}")
        print(f"Найденное действие: {found_action}, объект: {found_object}, доп. информация: {additional_info}")

        if not found_action:
            logger.info(f"Команда не распознана локально, передача в Mistral: {command}")
            print(f"Команда не распознана локально, передача в Mistral: {command}")
            response = await self.generate_text_mistral(user_id, command)
        else:
            try:
                if found_action == "open":
                    response = {
                        "browser": self.open_browser,
                        "notepad": self.open_notepad,
                        "explorer": self.open_explorer,
                        "calculator": self.open_calculator,
                        "discord": self.open_discord,
                        "telegram": self.open_telegram,
                        "whatsapp": self.open_whatsapp,
                        "skype": self.open_skype,
                        "vscode": self.open_vscode,
                        "word": self.open_word,
                        "excel": self.open_excel,
                        "powerpoint": self.open_powerpoint,
                        "paint": self.open_paint,
                        "cmd": self.open_cmd,
                        "task_manager": self.open_task_manager,
                        "settings": self.open_settings,
                        "control_panel": self.open_control_panel
                    }.get(found_object, lambda: f"Не могу открыть {found_object}")()
                
                elif found_action == "close":
                    response = self.close_application(found_object)
                
                elif found_action == "play":
                    if found_object == "music":
                        default_music = self.user_preferences.get(user_id, {}).get("default_music", "youtube")
                        response = {
                            "youtube": self.play_music_on_youtube,
                            "soundcloud": self.play_music_on_soundcloud,
                            "yandex_music": self.play_music_on_yandex_music,
                            "spotify": self.play_music_on_spotify
                        }.get(default_music, self.play_music_on_youtube)()
                    elif found_object == "youtube":
                        query = command_lower.replace("play", "").replace("youtube", "").strip()
                        play_type = "first" if additional_info == "play_first" else ("number" if additional_info == "play_by_number" else "title")
                        number = int(re.search(r"\d+", command_lower).group()) if play_type == "number" and re.search(r"\d+", command_lower) else None
                        title = command_lower.split("с надписью")[-1].strip() if play_type == "title" and "с надписью" in command_lower else None
                        response = self.find_and_play_youtube(query, play_type, number, title)
                
                elif found_action == "pause":
                    response = self.pause_music()
                elif "нажми на" in command_lower:
                    button_name = command_lower.replace("нажми на", "").strip()
                    response = await self.find_and_click_button(button_name)
    
                elif found_action == "resume":
                    response = self.resume_music()
                elif found_action == "next_track":
                    response = self.next_track()
                elif found_action == "previous_track":
                    response = self.previous_track()
                
                elif found_action == "app_volume_up":
                    response = self.adjust_app_volume(found_object, True)
                elif found_action == "app_volume_down":
                    response = self.adjust_app_volume(found_object, False)
                
                elif found_action == "volume_up":
                    response = self.adjust_volume(True)
                elif found_action == "volume_down":
                    response = self.adjust_volume(False)
                elif found_action == "mute":
                    response = self.mute_system()
                elif found_action == "unmute":
                    response = self.unmute_system()
                elif found_action == "weather":
                    city = self.user_preferences.get(user_id, {}).get("default_city", "Moscow")
                    response = self.get_weather(city)
                elif found_action == "time":
                    response = self.get_time()
                elif found_action == "screenshot":
                    response = self.take_screenshot()
                elif found_action == "reminder":
                    minutes = re.search(r"\d+", command_lower)
                    minutes = int(minutes.group()) if minutes else 5
                    response = self.set_reminder(minutes)
                elif found_action == "minimize":
                    response = self.minimize_window()
                elif found_action == "maximize":
                    response = self.maximize_window()
                elif found_action == "input":
                    text = command_lower.replace("введи", "").replace("вводи", "").strip()
                    response = self.input_text(text)
                elif found_action == "set_queue":
                    queue_number = re.search(r"\d+", command_lower)
                    queue_number = int(queue_number.group()) if queue_number else None
                    response = self.set_queue(queue_number) if queue_number else "Укажите число для очереди."
                elif found_action == "check_queue":
                    response = self.check_queue()
                elif found_action == "set_preference":
                    key_value = command_lower.split("установи")[-1].strip().split()
                    if len(key_value) >= 2:
                        key, value = key_value[0], " ".join(key_value[1:])
                        self.set_preference(user_id, key, value)
                        response = f"Предпочтение {key} установлено на {value}."
                    else:
                        response = "Укажите ключ и значение для установки предпочтения."
                elif found_action == "shutdown":
                    response = self.shutdown_computer()
                elif found_action == "restart":
                    response = self.restart_computer()
                elif found_action == "lock":
                    response = self.lock_computer()
                elif found_action == "new_tab":
                    response = self.new_tab()
                elif found_action == "close_tab":
                    response = self.close_tab()
                elif found_action == "switch_tab":
                    response = self.switch_tab()
                elif found_action == "refresh":
                    response = self.refresh_page()
                elif found_action == "go_back":
                    response = self.go_back()
                elif found_action == "go_forward":
                    response = self.go_forward()
                elif found_action == "scroll_down":
                    response = self.scroll_page("down")
                elif found_action == "scroll_up":
                    response = self.scroll_page("up")
                elif found_action == "click":
                    position = "center"
                    if additional_info == "on_left":
                        position = "left"
                    elif additional_info == "on_right":
                        position = "right"
                    response = self.click_mouse(position)
                elif found_action == "double_click":
                    position = "center"
                    if additional_info == "on_left":
                        position = "left"
                    elif additional_info == "on_right":
                        position = "right"
                    response = self.double_click_mouse(position)
                elif found_action == "right_click":
                    position = "center"
                    if additional_info == "on_left":
                        position = "left"
                    elif additional_info == "on_right":
                        position = "right"
                    response = self.right_click_mouse(position)
                elif found_action == "type_password":
                    response = self.type_password()
                elif found_action == "type_email":
                    response = self.type_email()
                elif found_action == "clear":
                    response = self.clear_text()
                else:
                    response = f"Команда '{found_action}' не поддерживается."
            except Exception as e:
                response = f"Ошибка выполнения команды: {e}"
                logger.error(f"Ошибка выполнения команды: {e}")

        self.label.setText(response)
        logger.info(f"Ответ: {response}")
        print(f"Ответ: {response}")
        if not self.is_muted_manually:
            self.engine.say(self.clean_text_for_speech(response))
            self.engine.runAndWait()
        self.update_indicator("idle")
        self.start_auto_hide_timer()
        self.is_command_running = False

    def start_listening(self):
        logger.info("Запуск прослушивания")
        print("Запуск прослушивания")
        threading.Thread(target=self.listen_loop, daemon=True).start()

    def listen_loop(self):
        while True:
            self.listen_to_mic()

    def listen_to_mic(self):
        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            logger.debug("Слушаю микрофон...")
            print("Слушаю микрофон...")
            self.label.setText("Слушаю...")
            self.update_indicator("processing")
            recognizer.adjust_for_ambient_noise(source, duration=1)
            try:
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)  # Исправлено
                logger.debug("Аудио получено")
                print("Аудио получено")
                command = recognizer.recognize_google(audio, language='ru-RU')
                logger.info(f"Распознано: {command}")
                print(f"Распознано: {command}")
                self.label.setText(f"Распознано: {command}")
                if "альфред" in command.lower():
                    QTimer.singleShot(0, lambda: asyncio.ensure_future(self.handle_command("local_user", command)))
            except sr.UnknownValueError:
                logger.warning("Не удалось распознать речь")
                print("Не удалось распознать речь")
                self.label.setText("Не удалось распознать речь")
            except sr.RequestError as e:
                logger.error(f"Ошибка сервиса распознавания: {e}")
                print(f"Ошибка сервиса распознавания: {e}")
                self.label.setText(f"Ошибка сервиса: {e}")
            finally:
                self.update_indicator("idle")
                self.auto_hide_signal.emit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    assistant = AlfredAssistant(app)
    assistant.run()
    sys.exit(app.exec())