import pandas as pd
import random
from datetime import datetime, timedelta

# Список возможных названий событий
event_names = [
    "Запуск процесса",
    "Завершение задачи",
    "Обновление данных",
    "Проверка системы",
    "Отправка уведомления",
    "Обработка запроса",
    "Загрузка файла",
    "Сохранение изменений",
    "Инициализация модуля",
    "Завершение работы"
]


# Функция для генерации случайного времени
def random_time(start, end):
    return start + timedelta(seconds=random.randint(0, int((end - start).total_seconds())))


# Функция для генерации логов
def generate_logs(num_events):
    logs = []
    for i in range(num_events):
        event_id = i + 1
        event_name = random.choice(event_names)  # Случайное название события
        start_time = random_time(datetime.now(), datetime.now() + timedelta(days=15))
        end_time = start_time + timedelta(minutes=random.randint(1, 120))  # Событие длится от 1 до 120 минут
        duration = end_time - start_time  # Разница между временем начала и окончания
        logs.append({
            "ID события": event_id,
            "Название события": event_name,
            "Время начала события": start_time.strftime('%Y-%m-%d %H:%M:%S'),
            "Время окончания события": end_time.strftime('%Y-%m-%d %H:%M:%S'),
            "Продолжительность (мин)": duration.total_seconds() / 60  # Продолжительность в минутах
        })
    return logs


# Генерация 10 логов
num_events = 1000
logs = generate_logs(num_events)

# Создание DataFrame и запись в Excel
df = pd.DataFrame(logs)
excel_file = 'logs.xlsx'
df.to_excel(excel_file, index=False)

print(f"Логи успешно записаны в файл {excel_file}")
