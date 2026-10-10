"""回讀瀏覽器下載的 Excel / ZIP / PNG，不依賴網站內部匯出函式。"""
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image, ImageOps, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'verification'
NS = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
checks = []
for filename, expected_name, sheet in [('period.xlsx', 'expected_period.json', '期間統計'), ('detail.xlsx', 'expected_detail.json', '月明細'), ('filtered-period.xlsx', 'expected_filtered_period.json', '期間統計'), ('filtered-detail.xlsx', 'expected_filtered_detail.json', '月明細')]:
    with zipfile.ZipFile(OUT / filename) as archive:
        assert archive.testzip() is None
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        sheets = workbook.findall('x:sheets/x:sheet', NS)
        assert len(sheets) == 1 and sheets[0].attrib['name'] == sheet
        actual, types = [], []
        for row in ET.fromstring(archive.read('xl/worksheets/sheet1.xml')).findall('x:sheetData/x:row', NS):
            values, kinds = [], []
            for cell in row:
                kind = cell.attrib.get('t')
                if kind == 'n':
                    values.append(int(cell.find('x:v', NS).text))
                else:
                    values.append(''.join(cell.itertext()))
                kinds.append(kind)
            actual.append(values)
            types.append(kinds)
        expected = json.loads((OUT / expected_name).read_text())
        assert actual == expected
        numeric_columns = [2, 3, 4]
        for row in types[1:]:
            assert all(row[i] == 'n' for i in numeric_columns)
            assert all(row[i] == 'inlineStr' for i in range(len(row)) if i not in numeric_columns)
        if filename == 'detail.xlsx':
            assert len(actual) == 3401
            assert all(len(row[6]) == 2 for row in actual[1:])
            assert sum(row[3] < 0 for row in actual[1:]) == 8
            assert sum(row[4] < 0 for row in actual[1:]) == 74
        elif filename == 'period.xlsx':
            assert len(actual) == 580
    checks.append(f'{filename}：單工作表、欄名、全部數值、列排序及文字／數值型別均正確')

with zipfile.ZipFile(OUT / 'charts.zip') as archive:
    assert archive.testzip() is None
    names = archive.namelist()
    assert names == [f'圖表_11501-11509_全_{i:02}.png' for i in range(1, 7)]
    previews = []
    sizes = []
    for i, name in enumerate(names):
        from io import BytesIO
        image = Image.open(BytesIO(archive.read(name)))
        image.load()
        count = 100 if i < 5 else 79
        assert image.size == (2420, (100 + count * 72 + 135) * 2)
        sizes.append(image.size)
        if i in [0, 5]:
            # 只擷取表頭及頁尾以供人工檢視，原始下載圖檔維持完整。
            width, height = image.size
            head = image.crop((0, 0, width, 680))
            foot = image.crop((0, height - 160, width, height))
            summary = Image.new('RGB', (width, 840), 'white')
            summary.paste(head, (0, 0)); summary.paste(foot, (0, 680))
            summary.thumbnail((1210, 420))
            previews.append(summary)
    contact = Image.new('RGB', (1210, 840), '#eeeeee')
    for i, image in enumerate(previews):
        contact.paste(image, (0, i * 420))
    contact.save(OUT / 'chart-pages-preview.png')
checks.append('charts.zip：CRC 正確，六張圖檔依頁碼命名；前五張 100 項、末張 79 項，PNG 尺寸正確')
filtered_count = json.loads((OUT / 'filtered-count.json').read_text())['count']
if filtered_count > 100:
    with zipfile.ZipFile(OUT / 'filtered-charts.zip') as archive:
        assert archive.testzip() is None
        names = archive.namelist()
        assert len(names) == (filtered_count + 99) // 100
        for i, name in enumerate(names):
            from io import BytesIO
            image = Image.open(BytesIO(archive.read(name)))
            image.load()
            count = min(100, filtered_count - i * 100)
            assert image.size == (2420, (100 + count * 72 + 135) * 2)
else:
    image = Image.open(OUT / 'filtered-chart.png')
    image.load()
    assert image.size == (2420, (100 + filtered_count * 72 + 135) * 2)
checks.append('篩選後圖表：完整品項、每張最多 100 項與 PNG 尺寸正確')
image = Image.open(OUT / 'single.png'); image.verify()
checks.append('single.png：有效 PNG')
(OUT / 'output-results.json').write_text(json.dumps({'checks': checks, 'png_sizes': sizes}, ensure_ascii=False, indent=2))
print(json.dumps(checks, ensure_ascii=False, indent=2))
