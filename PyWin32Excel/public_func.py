# functions
def is_jpn(string):
    for ch in string:
        if u'\u3040' <= ch <= u'\u30ff':
            return True

    return False


def write_list(txt_path, lst):
    f = open(txt_path, 'w', encoding='UTF-8')
    for var in lst:
        f.write(var)
    f.close()

