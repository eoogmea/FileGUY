#!/usr/bin/env bash
# Prepara Linux ed esegue tutti i test di FileGUY.
# Uso, dalla cartella del programma:   bash test/linux/prepara_e_prova.sh
# Per provare anche una stampa vera:    FILEGUY_STAMPA_VERA=1 bash test/linux/prepara_e_prova.sh
#   (sulla stampante predefinita, oppure su quella indicata con FILEGUY_STAMPANTE=nome)
# Funziona su Debian/Ubuntu (apt), Fedora/RHEL (dnf), openSUSE (zypper) e Arch (pacman).
set -u
cd "$(dirname "$0")/../.."

echo "=== Pacchetti del sistema (serve la password di sudo) ==="
if command -v apt-get >/dev/null; then
    sudo apt-get update
    sudo apt-get install -y python3 python3-venv python3-pip python3-tk cups-client xdg-utils fontconfig xvfb
elif command -v dnf >/dev/null; then
    sudo dnf install -y python3 python3-pip python3-tkinter cups-client xdg-utils fontconfig xorg-x11-server-Xvfb
elif command -v zypper >/dev/null; then
    sudo zypper install -y python3 python3-pip python3-tk cups-client xdg-utils fontconfig xvfb-run
elif command -v pacman >/dev/null; then
    sudo pacman -S --needed --noconfirm python python-pip tk cups xdg-utils fontconfig xorg-server-xvfb
else
    echo "Gestore dei pacchetti non riconosciuto: installa a mano python3, tkinter, cups-client, xdg-utils, fontconfig"
fi

echo "=== Ambiente Python (.venv-linux) ==="
python3 -m venv .venv-linux || exit 1
. .venv-linux/bin/activate
pip install --upgrade pip >/dev/null
pip install -r requirements.txt || exit 1

# Le prove con le finestre hanno bisogno di uno schermo: se non c'è usa quello virtuale (xvfb)
esegui() {
    if [ -n "${DISPLAY:-}${WAYLAND_DISPLAY:-}" ]; then
        python "$@"
    else
        xvfb-run -a python "$@"
    fi
}

esito=0
for prova in test/test_linux.py test/test_moduli.py test/test_editor_modelli.py test/test_editor_excel.py; do
    echo
    echo "=== $prova ==="
    esegui -W ignore "$prova" || esito=1
done
echo
echo "=== test/esegui_test.py (servizio) ==="
esegui -W ignore test/esegui_test.py SingleGUY || esito=1

echo
if [ $esito -eq 0 ]; then
    echo "TUTTO OK. L'immagine con i font da guardare: test/_run/linux/font_linux.png"
else
    echo "CI SONO ERRORI: vedi sopra le righe [ERR]"
fi
exit $esito
