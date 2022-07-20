import openpyxl
import collections

cache = None
def list_personalities(sheet_path: str):
    global cache
    if cache is not None:
        return cache
    cache = _list_personalities(sheet_path)
    return cache

def _list_personalities(sheet_path: str):
    srcdata = collections.defaultdict(list)
    dstdata = collections.defaultdict(list)
    wb = openpyxl.load_workbook(sheet_path, data_only=True)
    ws = wb["ペルソナリスト"]
    personalities = []
    for lines in ws:
        if lines[0].value == "No":
            continue
        items = [v.value for v in lines][2:4]
        for item in items:
            personalities.append(item.replace("\n", ""))
    return personalities

if __name__ == "__main__":
    import sys
    print(len(list_personalities(sys.argv[1])))
