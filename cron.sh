#!/bin/bash
export VIRTUALENVWRAPPER_PYTHON=/usr/bin/python3
. /usr/local/bin/virtualenvwrapper.sh

cd /home/mystikite/public_html/gw2
workon gw2
python3 ./gw2_achievements_xls.py > /dev/null
