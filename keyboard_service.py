from pynput.keyboard import Key, Controller, KeyCode
from pynput import keyboard


class KeyboardService:

    def __init__(self):
        self.controller = Controller()
        self.listener = keyboard.Listener(
            on_press=self.on_press,
        )

    def start(self):
        self.listener.start()

    def stop(self):
        self.listener.stop()
        self.listener = keyboard.Listener(
            on_press=self.on_press,
        )

    def on_press(self, key):
        try:
            self.stop()
            print("Pressing keys")
            with self.controller.pressed(Key.ctrl):
                self.controller.press(Key.f15)
                self.controller.release(Key.f15)

        except AttributeError:
            print('AttributeError')
