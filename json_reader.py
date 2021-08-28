import json


def read_file(json_file):
    with open(json_file, 'r', encoding="utf-8") as load_str:
        load_json = json.load(load_str)
        return load_json


if __name__ == '__main__':
    read_file("format_conf.json")
    print(__name__)
