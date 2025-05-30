FROM python:3.12-slim

WORKDIR /app

# Установка зависимостей
COPY requirements.txt .
RUN pip install -r requirements.txt

# Копирование файлов приложения
COPY . .

# Создание директории для данных
RUN mkdir -p /data

# Установка переменных окружения
ENV PYTHONUNBUFFERED=1

# Открытие порта
EXPOSE 8501

# Запуск приложения
CMD ["streamlit", "run", "dashboard.py", "--server.address", "0.0.0.0"] 