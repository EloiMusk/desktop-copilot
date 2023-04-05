import os

from flask import Flask, request
from pynput import keyboard

from keyboard_service import on_press, on_release

listener = keyboard.Listener(
    on_press=on_press,
    on_release=on_release)


def create_app(test_config=None):
    # create and configure the app
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_mapping(
        SECRET_KEY='dev',
        DATABASE=os.path.join(app.instance_path, 'flaskr.sqlite'),
    )

    if test_config is None:
        # load the instance config, if it exists, when not testing
        app.config.from_pyfile('config.py', silent=True)
    else:
        # load the test config if passed in
        app.config.from_mapping(test_config)

    # ensure the instance folder exists
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
            listener.start()
        except:
            pass
        return {'data': 'started'}, 200

    @app.route('/stop', methods=['POST'])
    def stop():
        try:
            listener.stop()
        except:
            pass
        return {'data': 'stopped'}, 200

    return app
