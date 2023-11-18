# Используйте официальный образ .NET для сборки и запуска приложения
FROM mcr.microsoft.com/dotnet/runtime:5.0 AS base
WORKDIR /app

# Скопируйте исполняемый файл приложения в образ
COPY bin/Release/app.publish/ .

# Укажите, что приложение должно быть запущено при старте контейнера
ENTRYPOINT ["KodinMyynti.exe"]