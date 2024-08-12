import pandas


def worksheets_dimensions(path: str) -> dict[int | str, tuple[int, int]]:
    file = pandas.ExcelFile(path)
    info = {sheet_name: file.parse(sheet_name).shape for sheet_name in file.sheet_names}
    return info
