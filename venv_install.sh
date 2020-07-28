#!/bin/bash
python3 -m venv venv
source venv/bin/activate
curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py
python get-pip.py
pip install -U pip
pip3 install -r requirements.txt