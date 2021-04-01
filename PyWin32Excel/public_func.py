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


def get_list(txt_path):
    res = ()
    f = open(txt_path, 'r', encoding='UTF-8')
    for line in f:
        if line[0] == '[' or line[0] == '\n':
            continue
        res += (line.split(','),)
    return res


def dic_compare_lst(dic, list_to_search):
    list_output = []
    for item in list_to_search:
        if item in dic:
            list_output.append(dic[item])
        else:
            list_output.append('[NONE]')
    return list_output
