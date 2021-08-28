
import json_reader


def generate_format(json_file):
    return json_reader.read_file(json_file)


if __name__ == '__main__':
    generate_format("format_conf.json")
    print(__name__)
