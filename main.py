import os

from flask import Flask, request
from pynput import keyboard

from keyboard_service import KeyboardService

app = Flask(__name__, instance_relative_config=True)
app.config.from_mapping(
    SECRET_KEY='dev',
    DATABASE=os.path.join(app.instance_path, 'flaskr.sqlite'),
)

keyboard_service = KeyboardService()

try:
    os.makedirs(app.instance_path)
except OSError:
    pass


@app.route('/completions', methods=['POST'])
def completions():
    data = request.get_json()
    print(data)
    return {'data': data}, 200


@app.route('/update-store', methods=['POST'])
def update_store():
    data = request.get_json()
    print(data)
    return {'data': data}, 200


@app.route('/start', methods=['POST'])
def start():
    try:
        keyboard_service.start()
    except:
        pass
    return {'data': 'started'}, 200


@app.route('/stop', methods=['POST'])
def stop():
    try:
        keyboard_service.stop()
    except:
        pass
    return {'data': 'stopped'}, 200
