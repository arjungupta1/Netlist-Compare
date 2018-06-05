import os
import glob
from bs4 import BeautifulSoup
import shutil


def main():
    file1, file2 = find_files()

    dict1 = parse_to_dict(file1)
    dict2 = parse_to_dict(file2)

    # We are comparing file 1 to file 2 and replacing the packages in dir2 with the ones in dir1.
    file_mapping = find_dirs(file1, file2)
    print(file_mapping)

    create_new_symbols(dict1, dict2, file_mapping)


def parse_to_dict(file):
    extension = os.path.splitext(file)[1]
    file_content = open(file, 'r').read()
    print(extension)
    if extension not in [".htm", ".html"]:
        print("Invalid file selected. Unable to parse a document that is not a htm or html file.")
        raise SystemExit
    soup = BeautifulSoup(file_content, "html.parser")
    title = soup.title.string
    if title != "Component Report":
        print("\nInvalid file selected. Unable to parse a document that is not a Component Report.")
        print("File selected was: {}\n".format(file))
        raise SystemExit
    table = soup.find("table")
    dict = {}
    # first "tr" contains column name definitions
    for row in table.find_all('tr')[1:]:
        cells = row.find_all('td')
        if len(cells) > 0:
            ref_des = cells[0].text.strip()
            package_name = cells[4].text.strip()
            dict[ref_des] = package_name
    return dict


def find_dirs(file1, file2):
    dirs = [d for d in os.listdir(os.getcwd()) if os.path.isdir(d)]
    print("Current directories: ")
    print("-------------------------------------------------------------")
    for i in range(0, len(dirs)):
        print("{}. {}".format(i + 1, dirs[i]))
    print("-------------------------------------------------------------")
    file_mapping = {**validate_dir(file1, dirs), **validate_dir(file2, dirs)}
    return file_mapping


def find_files():
    subdir = "Data"
    path = os.path.normpath(os.path.join(os.getcwd(), subdir))
    os.chdir(path)
    files = [f for f in os.listdir(os.getcwd()) if os.path.isfile(f)]

    print("Current files: ")
    print("-------------------------------------------------------------")
    for i in range(0, len(files)):
        print("{}. {}".format(i + 1, files[i]))
    print("-------------------------------------------------------------")

    file1 = validate_file(files)
    file2 = validate_file(files)
    file1, file2 = check_for_dupes(file1, file2, files)
    print("Selected files are: {} and {}".format(file1, file2))
    return file1, file2


def validate_file(files):
    file_num = -1
    try:
        file_num = int(input("Enter the number representing the file you want to read from: "))
        if file_num > len(files) or file_num < 0:
            file_num = -1
            raise ValueError
    except ValueError:
        print("Entered number is invalid.")
        while 0 < file_num <= len(files):
            file_num = int(input("Enter the number representing the file you want to read from: "))
    return files[file_num - 1]


def validate_dir(file, dirs):
    dir_num = -1
    try:
        dir_num = int(input("Enter the directory that you want to link to {}: ".format(file)))
        if dir_num > len(dirs) or dir_num < 0:
            dir_num = -1
            raise ValueError
    except ValueError:
        print("Entered number is invalid.")
        while 0 < dir_num <= len(dirs):
            dir_num = int(input("Enter the directory that you want to link to {}:".format(file)))
    return {file: dirs[dir_num - 1]}


def check_for_dupes(file1, file2, files):
    valid_files = False

    while not valid_files:
        if file1 == file2:
            print("Duplicate file found!")
            file2 = validate_file(files)
            continue
        else:
            valid_files = True
    return file1, file2


def create_new_symbols(dict1, dict2, file_mapping):
    prev_path = os.getcwd()
    dir_values = list(file_mapping.values())
    first_dir = dir_values[0]
    second_dir = dir_values[1]
    second_path = os.path.normpath(os.path.join(prev_path, second_dir))
    new_path = create_new_path(prev_path)

    # CHECK WHAT NEW SYMBOLS MATCH UP TO OLD SYMBOLS. Returns a dict
    new_mapping = get_old_to_new_mapping(dict1, dict2)

    # go into second path, look for .fsm, .dra, .bsm, .psm files
    old_dir_files = search_for_files(new_mapping, second_path)
    change_and_add_files(old_dir_files, second_path, new_path)


def create_new_path(prev_path):
    new_path = os.path.normpath(os.path.join(prev_path, 'symbols_updated'))
    if not os.path.exists(new_path):
        os.makedirs(new_path)
    return new_path


def get_old_to_new_mapping(dict1, dict2):
    mapping = {}
    dict1.pop("H3") # temp fix... unsure how to handle but leave it for now
    for (refdes1, refdes2) in zip(dict1, dict2):
        if refdes1 == refdes2:
            if dict1[refdes1] == dict2[refdes2]:
                mapping[dict1[refdes1]] = dict1[refdes1]
            else:
                mapping[dict1[refdes1]] = dict2[refdes2]
        else:
            print("Error parsing. Need to quit!")
            raise SystemExit
    return mapping


def search_for_files(symbol_mapping, path):
    os.chdir(path)
    file_types = ['*.dra', '*.fsm', '*.psm', '*.bsm', '*.osm', '*.ssm']
    files_grabbed = []
    for files in file_types:
        files_grabbed.extend(os.path.basename(x) for x in glob.glob(files))
    symbol_to_file_map = {}
    num_files = _check_copied_structure(os.getcwd())
    for file in files_grabbed:
        split_file = file.split('.')[0]
        if split_file not in symbol_mapping:
            symbol_mapping[split_file] = split_file

    for symbol in symbol_mapping.keys():
        value = symbol_mapping[symbol]
        if symbol not in symbol_to_file_map:
            symbol_to_file_map[symbol] = []
        for file in files_grabbed:
            if file in symbol_to_file_map[symbol]:
                continue
            filename, extension = os.path.splitext(file)
            # print("Filename: {}\t Value: {}".format(filename, value))
            if value.lower() == filename.lower():
                symbol_to_file_map[symbol].append(file)

    return symbol_to_file_map


def change_and_add_files(symbol_to_file_map, path, new_path):
    os.chdir(path)
    remove_files(new_path)
    num_files = add_files(symbol_to_file_map, new_path)


def remove_files(path):
    path_files = os.listdir(path)
    count = 0
    for file in path_files:
        file_path = os.path.join(path, file)
        if os.path.isfile(file_path):
            os.unlink(file_path)
            count += 1

    print("Removed {} file{} from directory.".format(count, 's' if count > 1 else '')
          if count > 0 else "No files in directory!")


def add_files(mapping, new_path):
    for key, value in mapping.items():
        changed_file_names = list(map(lambda x: '.'.join([key, x.split('.')[1]]), value))
        for old_file, new_file in zip(value, changed_file_names):
            new_file_path = os.path.join(new_path, new_file)
            shutil.copy2(old_file, new_file_path)
            # print("Copying {} to {}".format(os.path.basename(new_file_path), os.path.dirname(new_file_path)))
    num_files = _check_copied_structure(new_path)
    return num_files


def _check_copied_structure(file_path):
    file_types = ['*.dra', '*.fsm', '*.psm', '*.bsm', '*.osm', '*.ssm']
    lens = {}
    for type in file_types:
        lens[type] = len(glob.glob(type))
    print(lens)
    return lens


if __name__ == '__main__':
    main()
