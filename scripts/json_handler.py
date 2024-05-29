import json


def save_settings(settings, filename='settings.json'):
    with open(filename, 'w') as file:
        json.dump(settings, file, indent=4)


def load_settings(filename='settings.json'):
    try:
        with open(filename, 'r') as file:
            settings = json.load(file)
            return settings
    except FileNotFoundError:
        return {}
