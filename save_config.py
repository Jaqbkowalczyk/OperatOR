
import yaml


def save_config(list_doc, config_file):
    with open(config_file, "w") as f:
        yaml.dump(list_doc, f)