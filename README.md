# json2excel
## test
```python
    json_dic = {
        "style": SheetStyle.title_on_column,
        "sheetTitle": "测试卷",
        "titles": [
            "学科",
            "题号",
            "分数",
        ],
        "content": {
            "学科": [1, 3, 5, 8],
            "题号": ["🌸", "", "🐸"],
            "分数": ["", "110"],
        }
    }
    xlsx = 'test.xlsx'
    xls = 'test.xls'
    # json2xlsx json ton xlsx file
    json2xlsx(json_dic, xlsx)
    # json2xls json ton xls file
    json2xls(json_dic, xls)
```

![test](https://github.com/xiaominghe2014/json2excel/raw/master/test.png)