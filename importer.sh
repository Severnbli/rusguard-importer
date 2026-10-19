#!/usr/bin/env bash
cd "$(dirname "$0")" || exit 1
mkdir -p data
export LOCAL_UID="$(id -u)" LOCAL_GID="$(id -g)"
# Разрешить контейнеру (тот же пользователь) подключаться к X-серверу
command -v xhost >/dev/null && xhost +SI:localuser:"$(id -un)" >/dev/null
# по умолчанию берётся опубликованный образ; --build — собрать локально из исходников
if [ "$1" = "--build" ]; then
  exec docker compose run --rm --build --pull never importer
fi
exec docker compose run --rm importer
