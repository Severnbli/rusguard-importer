#!/usr/bin/env bash
cd "$(dirname "$0")" || exit 1
mkdir -p data
export LOCAL_UID="$(id -u)" LOCAL_GID="$(id -g)"
# Разрешить контейнеру (тот же пользователь) подключаться к X-серверу
command -v xhost >/dev/null && xhost +SI:localuser:"$(id -un)" >/dev/null
exec docker compose run --rm --build importer
