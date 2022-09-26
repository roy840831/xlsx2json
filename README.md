# xlsx2json

## 概述

This a tool for create json from excel(xlsx).

很多時候會需要透過json產出大量相似的定義檔或是測試資料，

剛好excel的行列格式能效率的產出資料，

透過本程式能將excel與json相互轉換。

## 功能介绍
```sh
# 查看使用帮助
運行環境為python3，相依套件記載在requirements.txt，
未安裝python的，有編譯二進制程式在\dist\xlsx2json.exe
xlsx2json.py -h or xlsx2json.exe --help

# 使用说明
usage: xlsx2json.py [-h] [-x XLSX] [-j JSON] [-o OUTPUT]
description: Can create json file from xlsx or create xlsx file from json
options:
  -h, --help                  show this help message and exit
  -x XLSX, --xlsx XLSX        xlsx file path(absolute or relative)
  -j JSON, --json JSON        json file path(absolute or relative)
  -o OUTPUT, --output OUTPUT  output file path(absolute or relative)
  
# example
將example.json輸出到output.xlsx
xlsx2json.exe -j example.json
將example.json輸出到output.xlsx
xlsx2json.exe -j example.json -o example.xlsx
將example.xlsx輸出到output.json
xlsx2json.exe -j example.xlsx

 ```

## excel對應json結構

本程式 透過下列三種轉換方式，將json轉換到excel，
```sh
    pyStructForJson = {
        "dict": "object",
        "list": "array",
        "list-dict": "dictInArray",
    }
```
### 1. object:
  key-value型，excel內為dict表示，左邊放KEY的時候右邊為value。
  
### 2. array:
  連續資料型，excel內為list表示，以index當作KEY，array的value放在index的右邊。
  
### 3. array of object:
  同樣重複key的object出現在同一個array中，為本程式特別定義的資料結構，
  
  excel內為list-dict表示，將鍵值轉置，第一行為key下排為填object的value。
  
  這樣的結構在excel內，可以藉由拖拉快速產出測試資料或是類似結構的定義值，
