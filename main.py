import os.path
import shutil
import xlwings as xw

from urllib.parse import quote

pre_made_mini_url = "./pre-made-mini-URLs.xlsx"
combined_url = "./fully_combined_urls.xlsx"
colors = [
    "#A96E6E",
    "#F3A079",
    "#99B2FF",
    "#D8D8A4",
    "#DB4DFF",
    "#9BF8FF",
    "#A9D6A2",
    "#D1FF36",
    "#86FFDB",
    "#A7DBFB",
    "#F4E0F4",
    "#E7A7BF",
    "#FF9A9E",
    "#BAE0FF",
    "#E5FF8E",
    "#FCC88C",
    "#E7DDDC",
    "#DAD7D7",
    "#FFDD99",
    "#99FFAD",
    "#8780F6",
    "#FAD1EA",
    "#F7C3D4",
]
look_for_items = [
    "CURRENT_COMPANY",
    "REGION",
    "SENIORITY_LEVEL",
    "CURRENT_TITLE",
]


def run_combine():
    if os.path.exists(combined_url):
        os.remove(combined_url)
    shutil.copy(pre_made_mini_url, combined_url)
    book = xw.Book(combined_url)
    sheet = book.sheets["Sheet1"]
    last_line_number = get_last_line_number(sheet, "A4")
    rows = sheet[f"A4:B{last_line_number}"].value
    rows = sorted(rows, key=lambda r: r[1])
    rows = [r for r in rows if r[1] != 0]
    b_sum = 0
    c_index = 0
    urls = []
    new_rows = []
    for index, row in enumerate(rows):
        b_sum += row[1]
        if b_sum <= 2000:
            urls.append(row[0])
        else:
            new_rows.append((urls, b_sum - row[1], c_index))
            urls = [row[0]]
            b_sum = row[1]
            c_index += 1

    sheet.range(f"A4:A{last_line_number}").clear_contents()
    sheet.range(f"B4:B{last_line_number}").clear_contents()

    rows = [(c_quote(combine_link(row[0])), row[1], row[2]) for row in new_rows]
    print([(row[1], row[2]) for row in rows])

    sheet.range(f"A4:A{last_line_number}").value = [[r[0]] for r in rows]
    sheet.range(f"B4:B{last_line_number}").value = [[r[1]] for r in rows]
    sheet.range(f"C4:C{last_line_number}").value = [[r[2]] for r in rows]

    last_line_number = get_last_line_number(sheet, "A4")
    for index in range(4, last_line_number):
        color = colors[int(sheet[f"C{index}"].value) % len(colors)]
        sheet.range(f"A{index}:C{index}").color = color
    book.save()
    book.close()


def combine_link(rng):
    first = rng[0]
    for look_for in look_for_items:
        dat = [a.partition(f"(type:{look_for},values:List(")[2].partition(")))")[0] + ")" for a in rng]
        dat = [a + ")" if a.endswith("List()") else a for a in dat]
        first = first.replace(dat[0], ",".join(dat))
    return first.replace("%2520", "%20")


def c_quote(url):
    link = url.partition("?")
    link = link[0] + "?" + "&".join([l.partition("=")[0] + "=" + quote(l.partition("=")[2]) for l in link[2].split("&")])
    return link.replace("%25", "%")


def get_last_line_number(sheet, start_cell):
    column_range = sheet.range(start_cell)
    last_cell = column_range.end("down")
    last_line_number = last_cell.row
    return last_line_number


if __name__ == "__main__":
    run_combine()
