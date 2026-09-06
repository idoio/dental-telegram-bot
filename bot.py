import asyncio
import sys
import os
import re
from datetime import datetime, date, timedelta
from typing import List, Dict
import logging

# Настройка event loop для Windows
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# Импортируем библиотеки
from aiogram import Bot, Dispatcher, types, F
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton, Message, \
    CallbackQuery
from aiogram.filters import Command
import yadisk
from dotenv import load_dotenv

# Для работы с Excel
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===== НАСТРОЙКА ЛОГИРОВАНИЯ =====
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("dental_bot.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ===== ЗАГРУЗКА КОНФИГУРАЦИИ =====
load_dotenv()

# Проверяем обязательные переменные
REQUIRED_ENV_VARS = ['BOT_TOKEN', 'YANDEX_TOKEN', 'ADMIN_CHAT_ID']
missing_vars = [var for var in REQUIRED_ENV_VARS if not os.getenv(var)]

if missing_vars:
    logger.error(f"❌ Отсутствуют переменные окружения: {', '.join(missing_vars)}")
    logger.info("Создайте файл .env с содержимым из примера выше")
    exit(1)

# Загружаем конфигурацию
BOT_TOKEN = os.getenv("BOT_TOKEN")
YANDEX_TOKEN = os.getenv("YANDEX_TOKEN")
ADMIN_CHAT_ID = os.getenv("ADMIN_CHAT_ID")

# Настройки клиники
CLINIC_NAME = os.getenv("CLINIC_NAME", "Стоматологическая клиника")
CLINIC_PHONE = os.getenv("CLINIC_PHONE", "+7 (XXX) XXX-XX-XX")
CLINIC_ADDRESS = os.getenv("CLINIC_ADDRESS", "г. Москва, ул. Примерная, 123")
CLINIC_HOURS = os.getenv("CLINIC_HOURS", """Пн-Пт: 9:00-20:00
Сб-Вс: 10:00-18:00""")

# Настройки врачей
DOCTORS = {
    "Иванова Анна Петровна": ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница"],
    "Петров Сергей Иванович": ["Вторник", "Среда", "Четверг", "Пятница", "Суббота"],
    "Сидорова Мария Владимировна": ["Понедельник", "Среда", "Пятница", "Суббота"],
}

# Время приема
WORK_HOURS = {
    "start": 9,
    "end": 20,
    "interval": 60
}


# ===== КЛАСС ДЛЯ РАБОТЫ С EXCEL =====
class ExcelDataManager:
    """Управление данными в Excel файле"""

    def __init__(self, token: str = None):
        """Инициализация"""
        self.excel_file_path = "dental_appointments.xlsx"
        self.disk = None

        if token:
            try:
                self.disk = yadisk.YaDisk(token=token)
                if not self.disk.check_token():
                    logger.warning("⚠️ Токен Яндекс.Диска невалиден")
                    self.disk = None
                else:
                    logger.info("✅ Яндекс.Диск подключен")
            except Exception as e:
                logger.warning(f"⚠️ Ошибка подключения к Яндекс.Диску: {e}")
                self.disk = None

        self._init_excel_file()
        self._convert_old_to_new_if_exists()

    def _init_excel_file(self):
        """Создание или загрузка Excel файла"""
        try:
            if not os.path.exists(self.excel_file_path):
                self._create_new_excel_file()
                logger.info("✅ Создан новый Excel файл для записей")
            else:
                logger.info("✅ Excel файл уже существует")

        except Exception as e:
            logger.error(f"❌ Ошибка инициализации Excel файла: {e}")
            raise

    def _convert_old_to_new_if_exists(self):
        """Конвертация старого файла в новый формат"""
        old_path = "dental_applications.xlsx"
        if os.path.exists(old_path):
            try:
                old_wb = load_workbook(old_path)
                old_ws = old_wb.active

                if not os.path.exists(self.excel_file_path):
                    self._create_new_excel_file()

                new_wb = load_workbook(self.excel_file_path)
                new_ws = new_wb.active

                if old_ws.max_row > 1:
                    for row in old_ws.iter_rows(min_row=2, values_only=True):
                        if len(row) >= 7:
                            new_row = [
                                new_ws.max_row,
                                "",
                                row[1] if len(row) > 1 else "",
                                row[2] if len(row) > 2 else "",
                                row[3] if len(row) > 3 else "",
                                row[4] if len(row) > 4 else "",
                                "",
                                "",
                                "",
                                row[5] if len(row) > 5 else "Новая",
                                row[6] if len(row) > 6 else "telegram",
                                ""
                            ]
                            new_ws.append(new_row)

                    new_wb.save(self.excel_file_path)
                    logger.info(f"✅ Перенесено {old_ws.max_row - 1} записей из старого файла")

            except Exception as e:
                logger.error(f"❌ Ошибка конвертации старого файла: {e}")

    def _create_new_excel_file(self):
        """Создание нового Excel файла"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Записи"

        headers = [
            "ID",
            "ID пользователя",
            "Дата создания",
            "Время создания",
            "ФИО",
            "Телефон",
            "Дата приема",
            "Время приема",
            "Врач",
            "Статус",
            "Источник",
            "Комментарий"
        ]

        ws.append(headers)
        self._apply_styles_to_sheet(ws)
        wb.save(self.excel_file_path)

    def _apply_styles_to_sheet(self, ws):
        """Применение стилей к листу Excel"""
        header_font = Font(bold=True, color="FFFFFF", size=12)
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border

        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            adjusted_width = min(max_length + 2, 30)
            ws.column_dimensions[column_letter].width = adjusted_width

    def save_appointment(self, user_id: int, full_name: str, phone: str,
                         appointment_date: str, appointment_time: str,
                         doctor: str, source: str = "telegram") -> int:
        """Сохранение записи на прием в Excel"""
        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active

            last_row = ws.max_row
            if last_row == 1:
                app_id = 1
            else:
                max_id = 0
                for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                    if row[0] and int(row[0]) > max_id:
                        max_id = int(row[0])
                app_id = max_id + 1

            now = datetime.now()

            new_row = [
                app_id,
                user_id,
                now.strftime("%d.%m.%Y"),
                now.strftime("%H:%M:%S"),
                full_name,
                phone,
                appointment_date,
                appointment_time,
                doctor,
                "Подтверждена",
                source,
                ""
            ]

            ws.append(new_row)
            phone_cell = ws.cell(row=last_row + 1, column=6)
            phone_cell.number_format = '@'

            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            for col in range(1, 13):
                cell = ws.cell(row=last_row + 1, column=col)
                cell.border = thin_border

            wb.save(self.excel_file_path)
            logger.info(f"✅ Запись сохранена: ID={app_id}, {full_name}, {appointment_date} {appointment_time}")

            if self.disk:
                try:
                    remote_path = f"/dental_appointments_{datetime.now().strftime('%Y%m')}.xlsx"
                    self.disk.upload(self.excel_file_path, remote_path, overwrite=True)
                    logger.info(f"✅ Файл загружен на Яндекс.Диск: {remote_path}")
                except Exception as e:
                    logger.warning(f"⚠️ Не удалось загрузить на Яндекс.Диск: {e}")

            return app_id

        except Exception as e:
            logger.error(f"❌ Ошибка сохранения записи в Excel: {e}")
            return -1

    def get_user_appointments(self, user_id: int) -> List[Dict]:
        """Получение всех записей пользователя"""
        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active

            appointments = []

            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[1] == user_id and row[9] in ["Подтверждена", "Новая"]:
                    app_data = {
                        'ID': row[0],
                        'ФИО': row[4],
                        'Телефон': row[5],
                        'Дата_приема': row[6],
                        'Время_приема': row[7],
                        'Врач': row[8],
                        'Статус': row[9],
                        'Дата_создания': row[2],
                        'Время_создания': row[3]
                    }
                    appointments.append(app_data)

            return appointments

        except Exception as e:
            logger.error(f"❌ Ошибка чтения записей пользователя: {e}")
            return []

    def cancel_appointment(self, appointment_id: int, comment: str = "") -> bool:
        """Отмена записи"""
        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active

            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=1).value == appointment_id:
                    ws.cell(row=row, column=10, value="Отменена")
                    ws.cell(row=row, column=12, value=comment)

                    wb.save(self.excel_file_path)
                    logger.info(f"✅ Запись отменена: ID={appointment_id}")
                    return True

            logger.warning(f"⚠️ Запись не найдена: ID={appointment_id}")
            return False

        except Exception as e:
            logger.error(f"❌ Ошибка отмены записи: {e}")
            return False

    def update_appointment(self, appointment_id: int, new_date: str = None,
                           new_time: str = None, new_doctor: str = None,
                           comment: str = "") -> bool:
        """Обновление записи"""
        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active

            for row in range(2, ws.max_row + 1):
                if ws.cell(row=row, column=1).value == appointment_id:
                    if new_date:
                        ws.cell(row=row, column=7, value=new_date)
                    if new_time:
                        ws.cell(row=row, column=8, value=new_time)
                    if new_doctor:
                        ws.cell(row=row, column=9, value=new_doctor)

                    ws.cell(row=row, column=12, value=comment)

                    wb.save(self.excel_file_path)
                    logger.info(f"✅ Запись обновлена: ID={appointment_id}")
                    return True

            logger.warning(f"⚠️ Запись не найдена: ID={appointment_id}")
            return False

        except Exception as e:
            logger.error(f"❌ Ошибка обновления записи: {e}")
            return False

    def check_time_slot(self, date: str, time: str, doctor: str) -> bool:
        """Проверка, свободно ли время у врача"""
        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active

            for row in ws.iter_rows(min_row=2, values_only=True):
                if (row[6] == date and row[7] == time and
                        row[8] == doctor and row[9] in ["Подтверждена", "Новая"]):
                    return False

            return True

        except Exception as e:
            logger.error(f"❌ Ошибка проверки времени: {e}")
            return True

    def get_statistics(self, days: int = 7) -> Dict:
        """Получение статистики"""
        try:
            wb = load_workbook(self.excel_file_path)
            ws = wb.active

            stats = {
                'total': 0,
                'today': 0,
                'active': 0,
                'canceled': 0,
                'by_doctors': {},
                'by_days': {}
            }

            today = date.today().strftime("%d.%m.%Y")

            for row in ws.iter_rows(min_row=2, values_only=True):
                stats['total'] += 1

                if row[6] == today:
                    stats['today'] += 1

                if row[9] == 'Подтверждена':
                    stats['active'] += 1
                elif row[9] == 'Отменена':
                    stats['canceled'] += 1

                doctor = row[8] or 'не указан'
                if doctor not in stats['by_doctors']:
                    stats['by_doctors'][doctor] = 0
                stats['by_doctors'][doctor] += 1

                app_date = row[6] or 'не указана'
                if app_date not in stats['by_days']:
                    stats['by_days'][app_date] = 0
                stats['by_days'][app_date] += 1

            return stats

        except Exception as e:
            logger.error(f"❌ Ошибка получения статистики: {e}")
            return {'total': 0, 'today': 0, 'active': 0, 'canceled': 0, 'by_doctors': {}, 'by_days': {}}

    def create_backup(self) -> bool:
        """Создание резервной копии"""
        try:
            if not os.path.exists(self.excel_file_path):
                logger.warning("❌ Нет файла для создания бэкапа")
                return False

            backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            import shutil
            shutil.copy2(self.excel_file_path, backup_name)

            if self.disk:
                try:
                    remote_path = f"/backups/{backup_name}"
                    self.disk.upload(backup_name, remote_path, overwrite=True)
                    os.remove(backup_name)
                    logger.info(f"✅ Бэкап создан на Яндекс.Диске: {remote_path}")
                except Exception as e:
                    logger.warning(f"⚠️ Не удалось загрузить бэкап на Яндекс.Диск: {e}")
            else:
                logger.info(f"✅ Локальный бэкап создан: {backup_name}")

            return True

        except Exception as e:
            logger.error(f"❌ Ошибка создания бэкапа: {e}")
            return False


# ===== ИНИЦИАЛИЗАЦИЯ МЕНЕДЖЕРА ДАННЫХ =====
try:
    print("=" * 50)
    print("ИНИЦИАЛИЗАЦИЯ СИСТЕМЫ ХРАНЕНИЯ ЗАПИСЕЙ")

    try:
        import openpyxl
    except ImportError:
        print("📦 Устанавливаю openpyxl...")
        import subprocess

        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
        print("✅ openpyxl установлен")

    data_manager = ExcelDataManager(YANDEX_TOKEN)
    logger.info("✅ Система хранения данных инициализирована")

    if os.path.exists("dental_appointments.xlsx"):
        wb = load_workbook("dental_appointments.xlsx")
        ws = wb.active
        print(f"📊 Структура таблицы: {ws.max_column} столбцов, {ws.max_row} строк")

except Exception as e:
    logger.error(f"❌ Не удалось инициализировать систему хранения: {e}")
    print("\n🔄 Использую упрощенное локальное хранилище...")


    class SimpleStorage:
        def save_appointment(self, *args, **kwargs):
            print("⚠️ Используется упрощенное хранилище")
            return 1

        def get_user_appointments(self, user_id):
            return []

        def cancel_appointment(self, *args, **kwargs):
            return True

        def update_appointment(self, *args, **kwargs):
            return True

        def check_time_slot(self, *args, **kwargs):
            return True

        def get_statistics(self, *args, **kwargs):
            return {'total': 0, 'today': 0, 'active': 0, 'canceled': 0}

        def create_backup(self):
            return True


    data_manager = SimpleStorage()

# ===== ИНИЦИАЛИЗАЦИЯ БОТА =====
bot = Bot(token=BOT_TOKEN, timeout=30)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)


# ===== СОСТОЯНИЯ ДЛЯ ЗАПИСИ =====
class AppointmentStates(StatesGroup):
    waiting_for_name = State()
    waiting_for_phone = State()
    waiting_for_date = State()
    waiting_for_time = State()
    waiting_for_doctor = State()
    waiting_for_confirmation = State()


# ===== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====
def get_main_keyboard():
    buttons = [
        [KeyboardButton(text="📅 Записаться на прием")],
        [KeyboardButton(text="📋 Мои записи")],
        [KeyboardButton(text="🕐 Часы работы")],
        [KeyboardButton(text="📍 Адрес и контакты")],
        [KeyboardButton(text="💰 Узнать цены")],
        [KeyboardButton(text="👨‍⚕️ Наши врачи")]
    ]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def get_phone_keyboard():
    buttons = [
        [KeyboardButton(text="📱 Отправить номер", request_contact=True)],
        [KeyboardButton(text="↩️ Назад в меню")]
    ]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def get_cancel_keyboard():
    buttons = [[KeyboardButton(text="❌ Отменить запись")]]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def get_back_keyboard():
    buttons = [[KeyboardButton(text="↩️ Назад")]]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def get_yes_no_keyboard():
    buttons = [
        [KeyboardButton(text="✅ Да, подтверждаю")],
        [KeyboardButton(text="❌ Нет, изменить")],
        [KeyboardButton(text="↩️ Отменить запись")]
    ]
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def generate_dates_keyboard():
    """Генерация клавиатуры с датами"""
    buttons = []
    today = date.today()

    # 5 ближайших дней
    for i in range(5):
        current_date = today + timedelta(days=i)
        date_str = current_date.strftime("%d.%m.%Y")
        weekday = current_date.strftime("%A")

        weekdays_ru = {
            "Monday": "Пн",
            "Tuesday": "Вт",
            "Wednesday": "Ср",
            "Thursday": "Чт",
            "Friday": "Пт",
            "Saturday": "Сб",
            "Sunday": "Вс"
        }

        weekday_ru = weekdays_ru.get(weekday, weekday)

        if i == 0:
            button_text = f"📅 Сегодня {date_str}"
        elif i == 1:
            button_text = f"📅 Завтра {date_str}"
        else:
            button_text = f"📅 {date_str} ({weekday_ru})"

        buttons.append([KeyboardButton(text=button_text)])

    buttons.append([KeyboardButton(text="↩️ Назад")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def generate_time_keyboard():
    """Генерация клавиатуры со временем"""
    buttons = []

    # Основные часы (каждые 2 часа)
    for hour in range(WORK_HOURS["start"], WORK_HOURS["end"], 2):
        time_str = f"{hour:02d}:00"
        buttons.append([KeyboardButton(text=time_str)])

    buttons.append([KeyboardButton(text="↩️ Назад")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


def generate_doctors_keyboard():
    """Генерация клавиатуры с врачами"""
    buttons = []

    for doctor in DOCTORS.keys():
        # Укороченные имена для кнопок
        if "Иванова" in doctor:
            short_name = "Иванова А.П."
        elif "Петров" in doctor:
            short_name = "Петров С.И."
        elif "Сидорова" in doctor:
            short_name = "Сидорова М.В."
        else:
            short_name = doctor[:15] + "..." if len(doctor) > 15 else doctor

        buttons.append([KeyboardButton(text=f"👨‍⚕️ {short_name}")])

    buttons.append([KeyboardButton(text="↩️ Назад")])
    return ReplyKeyboardMarkup(keyboard=buttons, resize_keyboard=True)


# ===== ОСНОВНЫЕ ОБРАБОТЧИКИ =====
@dp.message(Command("start"))
async def cmd_start(message: Message):
    try:
        await message.answer(
            f"👋 Добро пожаловать в <b>{CLINIC_NAME}</b>!\n\n"
            f"Я - ваш виртуальный помощник. Помогу записаться на прием.\n\n"
            f"Выберите действие из меню ниже:",
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка в команде /start: {e}")


@dp.message(F.text == "🕐 Часы работы")
async def show_hours(message: Message):
    try:
        await message.answer(
            f"<b>🕐 Часы работы {CLINIC_NAME}:</b>\n\n"
            f"{CLINIC_HOURS}\n\n"
            f"📞 Телефон: {CLINIC_PHONE}\n"
            f"📍 Адрес: {CLINIC_ADDRESS}",
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при показе часов работы: {e}")


@dp.message(F.text == "📍 Адрес и контакты")
async def show_address(message: Message):
    try:
        await message.answer(
            f"<b>📍 Контакты {CLINIC_NAME}:</b>\n\n"
            f"📞 Телефон: {CLINIC_PHONE}\n"
            f"📍 Адрес: {CLINIC_ADDRESS}\n\n"
            f"<b>Как нас найти:</b>\n"
            f"<i>{CLINIC_ADDRESS}</i>",
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при показе адреса: {e}")


@dp.message(F.text == "💰 Узнать цены")
async def show_prices(message: Message):
    try:
        await message.answer(
            f"<b>💰 Цены на услуги {CLINIC_NAME}:</b>\n\n"
            f"👨‍⚕️ <b>Консультация стоматолога:</b> от 1500 руб.\n"
            f"🦷 <b>Лечение кариеса:</b> от 3500 руб.\n"
            f"🌟 <b>Профессиональная чистка:</b> от 4000 руб.\n"
            f"😁 <b>Отбеливание зубов:</b> от 12000 руб.\n"
            f"👑 <b>Установка коронки:</b> от 15000 руб.\n\n"
            f"📞 <b>Точную стоимость уточняйте по телефону:</b>\n"
            f"{CLINIC_PHONE}\n\n"
            f"<i>*Цены могут меняться в зависимости от сложности случая</i>",
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при показе цен: {e}")


@dp.message(F.text == "👨‍⚕️ Наши врачи")
async def show_doctors(message: Message):
    try:
        doctors_list = "\n\n".join([f"<b>{doctor}</b>\nРабочие дни: {', '.join(days)}"
                                    for doctor, days in DOCTORS.items()])

        await message.answer(
            f"<b>👨‍⚕️ Наши специалисты {CLINIC_NAME}:</b>\n\n"
            f"{doctors_list}\n\n"
            f"📞 <b>Записаться к любому специалисту:</b>\n"
            f"{CLINIC_PHONE}",
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при показе врачей: {e}")


@dp.message(F.text == "📋 Мои записи")
async def show_my_appointments(message: Message):
    try:
        # Получаем записи пользователя
        appointments = data_manager.get_user_appointments(message.from_user.id)

        if not appointments:
            await message.answer(
                "📝 <b>У вас пока нет активных записей.</b>\n\n"
                "Хотите записаться на прием?",
                parse_mode="HTML",
                reply_markup=get_main_keyboard()
            )
            return

        # Формируем список записей
        appointments_text = "<b>📋 Ваши активные записи:</b>\n\n"
        for i, app in enumerate(appointments, 1):
            appointments_text += (
                f"<b>{i}. Запись №{app['ID']}</b>\n"
                f"   👤 Врач: {app['Врач']}\n"
                f"   📅 Дата: {app['Дата_приема']}\n"
                f"   🕐 Время: {app['Время_приема']}\n"
                f"   📞 Телефон: {app['Телефон']}\n"
                f"   📝 Статус: {app['Статус']}\n\n"
            )

        appointments_text += "Для управления записями обратитесь к администратору."

        await message.answer(
            appointments_text,
            parse_mode="HTML",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при показе записей: {e}")
        await message.answer(
            "❌ Ошибка при получении записей.\nПопробуйте позже.",
            reply_markup=get_main_keyboard()
        )


@dp.message(F.text == "📅 Записаться на прием")
async def start_appointment(message: Message, state: FSMContext):
    try:
        await state.set_state(AppointmentStates.waiting_for_name)
        await message.answer(
            "📝 <b>Для записи на прием введите ваше ФИО:</b>\n\n"
            "Например: Иванов Иван Иванович\n\n"
            "<i>Или нажмите 'Отменить запись' для возврата в меню</i>",
            parse_mode="HTML",
            reply_markup=get_cancel_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при начале записи: {e}")
        await message.answer("❌ Произошла ошибка. Попробуйте позже.", reply_markup=get_main_keyboard())


# ===== ПРОЦЕСС ЗАПИСИ =====
@dp.message(AppointmentStates.waiting_for_name)
async def process_name(message: Message, state: FSMContext):
    try:
        if message.text == "❌ Отменить запись":
            await state.clear()
            await message.answer("✅ Запись отменена.", reply_markup=get_main_keyboard())
            return

        if len(message.text.strip().split()) < 2:
            await message.answer("❌ Пожалуйста, введите полное ФИО", parse_mode="HTML")
            return

        await state.update_data(full_name=message.text.strip())
        await state.set_state(AppointmentStates.waiting_for_phone)

        await message.answer(
            "📱 <b>Теперь укажите ваш номер телефона:</b>\n\n"
            "Вы можете:\n"
            "1. Нажать кнопку '📱 Отправить номер'\n"
            "2. Ввести номер вручную: +79991234567\n\n"
            "<i>Или нажмите 'Отменить запись'</i>",
            parse_mode="HTML",
            reply_markup=get_phone_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при обработке имени: {e}")
        await message.answer("❌ Ошибка. Попробуйте снова.", reply_markup=get_main_keyboard())


@dp.message(AppointmentStates.waiting_for_phone)
async def process_phone(message: Message, state: FSMContext):
    try:
        if message.text == "↩️ Назад в меню":
            await state.clear()
            await cmd_start(message)
            return

        user_data = await state.get_data()

        if message.contact:
            phone = message.contact.phone_number
        else:
            phone = message.text.strip()

        if phone.startswith('8'):
            phone = '+7' + phone[1:]
        elif phone.startswith('7') and not phone.startswith('+7'):
            phone = '+' + phone

        if not re.match(r'^\+7\d{10}$', phone):
            await message.answer(
                "❌ Неверный формат номера.\n"
                "Введите номер в формате: +79991234567",
                parse_mode="HTML"
            )
            return

        await state.update_data(phone=phone)
        await state.set_state(AppointmentStates.waiting_for_date)

        await message.answer(
            "📅 <b>Выберите дату приема:</b>\n\n"
            "Доступны ближайшие даты:",
            parse_mode="HTML",
            reply_markup=generate_dates_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при обработке телефона: {e}")
        await message.answer("❌ Ошибка. Попробуйте снова.", reply_markup=get_main_keyboard())


@dp.message(AppointmentStates.waiting_for_date)
async def process_date(message: Message, state: FSMContext):
    try:
        if message.text == "↩️ Назад":
            await state.set_state(AppointmentStates.waiting_for_phone)
            # Отправляем простой текст, а не вызываем process_phone
            await message.answer(
                "📱 <b>Введите номер телефона:</b>",
                parse_mode="HTML",
                reply_markup=get_phone_keyboard()
            )
            return

        date_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', message.text)
        if not date_match:
            await message.answer("❌ Пожалуйста, выберите дату из списка")
            return

        selected_date = date_match.group(1)

        try:
            date_obj = datetime.strptime(selected_date, "%d.%m.%Y").date()
            today = date.today()
            if date_obj < today:
                await message.answer("❌ Нельзя выбрать прошедшую дату")
                return
        except:
            pass

        await state.update_data(appointment_date=selected_date)
        await state.set_state(AppointmentStates.waiting_for_time)

        # Отправляем простой текст сначала
        await message.answer(
            f"🕐 <b>Выберите время приема на {selected_date}:</b>",
            parse_mode="HTML"
        )

        # Затем отправляем клавиатуру отдельно
        await message.answer(
            "Доступное время:",
            reply_markup=generate_time_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при обработке даты: {e}")
        await message.answer("❌ Ошибка. Попробуйте снова.", reply_markup=get_main_keyboard())


@dp.message(AppointmentStates.waiting_for_time)
async def process_time(message: Message, state: FSMContext):
    try:
        if message.text == "↩️ Назад":
            await state.set_state(AppointmentStates.waiting_for_date)
            await message.answer(
                "📅 <b>Выберите дату приема:</b>",
                parse_mode="HTML",
                reply_markup=generate_dates_keyboard()
            )
            return

        time_match = re.match(r'^(\d{1,2}):00$', message.text)
        if not time_match:
            await message.answer("❌ Пожалуйста, выберите время из списка")
            return

        selected_time = message.text
        hour = int(time_match.group(1))

        if hour < WORK_HOURS["start"] or hour >= WORK_HOURS["end"]:
            await message.answer(f"❌ Клиника работает с {WORK_HOURS['start']}:00 до {WORK_HOURS['end']}:00")
            return

        await state.update_data(appointment_time=selected_time)
        await state.set_state(AppointmentStates.waiting_for_doctor)

        await message.answer(
            f"👨‍⚕️ <b>Выберите врача:</b>\n\n"
            f"Дата: {selected_time}",
            parse_mode="HTML",
            reply_markup=generate_doctors_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при обработке времени: {e}")
        await message.answer("❌ Ошибка. Попробуйте снова.", reply_markup=get_main_keyboard())


@dp.message(AppointmentStates.waiting_for_doctor)
async def process_doctor(message: Message, state: FSMContext):
    try:
        if message.text == "↩️ Назад":
            await state.set_state(AppointmentStates.waiting_for_time)
            await message.answer(
                "🕐 <b>Выберите время приема:</b>",
                parse_mode="HTML",
                reply_markup=generate_time_keyboard()
            )
            return

        doctor_text = message.text
        if "👨‍⚕️" in doctor_text:
            doctor_name = doctor_text.split("👨‍⚕️")[1].strip()
            # Преобразуем короткое имя обратно в полное
            if "Иванова А.П." in doctor_name:
                doctor_name = "Иванова Анна Петровна"
            elif "Петров С.И." in doctor_name:
                doctor_name = "Петров Сергей Иванович"
            elif "Сидорова М.В." in doctor_name:
                doctor_name = "Сидорова Мария Владимировна"
        else:
            doctor_name = doctor_text.strip()

        if doctor_name not in DOCTORS:
            await message.answer("❌ Пожалуйста, выберите врача из списка")
            return

        await state.update_data(doctor=doctor_name)
        await state.set_state(AppointmentStates.waiting_for_confirmation)

        user_data = await state.get_data()

        await message.answer(
            f"✅ <b>Проверьте данные записи:</b>\n\n"
            f"👤 <b>Пациент:</b> {user_data['full_name']}\n"
            f"📞 <b>Телефон:</b> {user_data['phone']}\n"
            f"📅 <b>Дата:</b> {user_data['appointment_date']}\n"
            f"🕐 <b>Время:</b> {user_data['appointment_time']}\n"
            f"👨‍⚕️ <b>Врач:</b> {user_data['doctor']}\n\n"
            f"<b>Всё верно?</b>",
            parse_mode="HTML",
            reply_markup=get_yes_no_keyboard()
        )
    except Exception as e:
        logger.error(f"❌ Ошибка при обработке врача: {e}")
        await message.answer("❌ Ошибка. Попробуйте снова.", reply_markup=get_main_keyboard())


@dp.message(AppointmentStates.waiting_for_confirmation)
async def process_confirmation(message: Message, state: FSMContext):
    try:
        if message.text == "↩️ Отменить запись":
            await state.clear()
            await message.answer("✅ Запись отменена.", reply_markup=get_main_keyboard())
            return

        user_data = await state.get_data()

        if message.text == "✅ Да, подтверждаю":
            appointment_id = data_manager.save_appointment(
                user_id=message.from_user.id,
                full_name=user_data['full_name'],
                phone=user_data['phone'],
                appointment_date=user_data['appointment_date'],
                appointment_time=user_data['appointment_time'],
                doctor=user_data['doctor']
            )

            if appointment_id > 0:
                await message.answer(
                    f"🎉 <b>Запись успешно создана!</b>\n\n"
                    f"<b>Номер записи:</b> №{appointment_id}\n"
                    f"👤 <b>Пациент:</b> {user_data['full_name']}\n"
                    f"📅 <b>Дата:</b> {user_data['appointment_date']}\n"
                    f"🕐 <b>Время:</b> {user_data['appointment_time']}\n"
                    f"👨‍⚕️ <b>Врач:</b> {user_data['doctor']}\n\n"
                    f"⏰ <b>Пожалуйста, приходите за 10 минут до назначенного времени.</b>",
                    parse_mode="HTML",
                    reply_markup=get_main_keyboard()
                )

                # Уведомление администратора
                try:
                    await bot.send_message(
                        chat_id=ADMIN_CHAT_ID,
                        text=f"📅 <b>НОВАЯ ЗАПИСЬ №{appointment_id}</b>\n\n"
                             f"👤 <b>Пациент:</b> {user_data['full_name']}\n"
                             f"📞 <b>Телефон:</b> {user_data['phone']}\n"
                             f"📅 <b>Дата:</b> {user_data['appointment_date']}\n"
                             f"🕐 <b>Время:</b> {user_data['appointment_time']}\n"
                             f"👨‍⚕️ <b>Врач:</b> {user_data['doctor']}\n"
                             f"👤 <b>User ID:</b> {message.from_user.id}",
                        parse_mode="HTML"
                    )
                except Exception as e:
                    logger.error(f"❌ Не удалось уведомить администратора: {e}")
            else:
                await message.answer(
                    "❌ Ошибка при сохранении записи.\n"
                    f"Позвоните нам: {CLINIC_PHONE}",
                    reply_markup=get_main_keyboard()
                )

            await state.clear()

        elif message.text == "❌ Нет, изменить":
            await state.set_state(AppointmentStates.waiting_for_date)
            await message.answer(
                "📅 <b>Выберите новую дату приема:</b>",
                parse_mode="HTML",
                reply_markup=generate_dates_keyboard()
            )

        else:
            await message.answer("Пожалуйста, выберите один из вариантов")
    except Exception as e:
        logger.error(f"❌ Ошибка при подтверждении: {e}")
        await message.answer("❌ Ошибка. Попробуйте снова.", reply_markup=get_main_keyboard())


# ===== ОБРАБОТЧИК ДЛЯ КНОПКИ "↩️ Назад в меню" =====
@dp.message(F.text == "↩️ Назад в меню")
async def back_to_main_menu(message: Message, state: FSMContext):
    try:
        await state.clear()
        await cmd_start(message)
    except Exception as e:
        logger.error(f"❌ Ошибка при возврате в меню: {e}")


# ===== ЗАПУСК БОТА =====
async def main():
    print("=" * 50)
    print("🚀 ЗАПУСК СТОМАТОЛОГИЧЕСКОГО БОТА")
    print("=" * 50)
    print(f"🏥 Клиника: {CLINIC_NAME}")
    print(f"📁 Формат хранения: Excel")
    print(f"👨‍⚕️ Врачей в системе: {len(DOCTORS)}")
    print("=" * 50)
    print("✅ Все кнопки активны")
    print("=" * 50)

    try:
        await dp.start_polling(bot, skip_updates=True)
    except Exception as e:
        logger.error(f"❌ Критическая ошибка при запуске бота: {e}")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n🛑 Бот остановлен")
    except Exception as e:
        logger.error(f"❌ Непредвиденная ошибка: {e}")
