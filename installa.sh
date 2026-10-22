#!/usr/bin/env bash
# Installa tutto quello che serve a FileGUY su Linux.
#   bash installa.sh              pacchetti del sistema (con sudo), ambiente Python e librerie
#   bash installa.sh --compila    in più crea gli eseguibili in dist/
# Funziona con apt (Debian/Ubuntu), dnf (Fedora/RHEL), zypper (openSUSE) e pacman (Arch).
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null; then
    echo "Python 3 non trovato: lo installo con il gestore dei pacchetti (serve la password di sudo)"
    if command -v apt-get >/dev/null; then
        sudo apt-get update && sudo apt-get install -y python3
    elif command -v dnf >/dev/null; then
        sudo dnf install -y python3
    elif command -v zypper >/dev/null; then
        sudo zypper install -y python3
    elif command -v pacman >/dev/null; then
        sudo pacman -S --needed --noconfirm python
    fi
fi
if ! command -v python3 >/dev/null; then
    echo "Installa Python 3 e rilancia: bash installa.sh"
    exit 1
fi
exec python3 installa.py "$@"
