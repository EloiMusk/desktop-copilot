from pynput.keyboard import Key, Controller, KeyCode
from pynput import keyboard

controller = Controller()


def on_press(key):
    try:
        controller.press(Key.alt)
        controller.press(Key.f15)
        controller.release(Key.alt)
        controller.release(Key.f15)
        return False
    except AttributeError:
        print('AttributeError')


def on_release(key):
    print('{0} released'.format(
        key))
    if key == keyboard.Key.esc:
        # Stop listener
        return False
