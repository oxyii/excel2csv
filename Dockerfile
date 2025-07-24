# Multi-stage build для оптимизации размера образа
FROM golang:1.21-alpine AS builder

# Установка зависимостей для сборки
RUN apk add --no-cache git

# Копирование исходного кода
WORKDIR /app
COPY go.mod go.sum ./
RUN go mod download

COPY . .

# Сборка HTTP сервера
RUN CGO_ENABLED=0 GOOS=linux go build -a -installsuffix cgo -o excel2csv-server ./cmd/excel2csv-server

# Production образ на Debian
FROM debian:12-slim

# Установка LibreOffice и зависимостей
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-calc \
    fonts-liberation \
    fonts-dejavu-core \
    ca-certificates \
    curl \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# Создание пользователя для безопасности
RUN useradd -r -u 1001 -m -d /home/excel2csv excel2csv

# Копирование исполняемого файла
COPY --from=builder /app/excel2csv-server /usr/local/bin/excel2csv-server
RUN chmod +x /usr/local/bin/excel2csv-server

# Создание директории для временных файлов
RUN mkdir -p /tmp/excel2csv && chown excel2csv:excel2csv /tmp/excel2csv

# Переключение на непривилегированного пользователя
USER excel2csv
WORKDIR /home/excel2csv

# Настройка переменных окружения
ENV PORT=8080
ENV TMPDIR=/tmp/excel2csv
ENV HOME=/home/excel2csv

# Проверка здоровья контейнера
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:${PORT}/health || exit 1

# Открытие порта
EXPOSE 8080

# Запуск сервера
CMD ["excel2csv-server"]
