# ExcelToJson
Unity ExcelToJson
## 支持常用值类型 包括数组
## 支持自定义类型转化,默认读取第二行，或者采用自动识别根据内容自动转化


示例：
![image](https://user-images.githubusercontent.com/25007281/236627668-75bef3d0-dd65-4507-bf9c-bd101e3315ca.png)
```
[
  {
    "ID": 1001,
    "Name": "技能名1",
    "Damage": [
      10,
      20,
      30
    ],
    "Speed": 1.5,
    "IsTest": true,
    "TestName": [
      "asd",
      "asdaf"
    ]
  },
  {
    "ID": 1002,
    "Name": "技能名2",
    "Damage": [
      10,
      20,
      30
    ],
    "Speed": 2.0,
    "IsTest": false,
    "TestName": null
  }
]
```
