# pdfScanner
## 功能
- 扫描pdf文件中的单词高亮, 按照不同的高亮颜色进行分类, 支持文件夹批量扫描
- 开始扫描: 开始扫描pdf高亮并结合勾选的功能选项进行操作
- 添加配置: 添加百度的key-secret(用于翻译单词),添加有道词典网页版的cookie(用于添加单词本),选择配置加载本地词典进行翻译
- 翻译单词: 本地词典+联网翻译单词
- 添加单词本: 自动添加单词到网易有道词典单词本
- 单词Excel: 将扫描的结果归类整理到Excel, 并按照pdf文件名保存为一个个Excel worksheet, 并进行排序(output.xlsx)
- 单词本xml: 生成网易有道词典单词本xml文件(youdao_wordbook.xml)

## 说明
- 在一切开始之前, 需要先点击添加配置, 填入百度翻译api的key和secret, 以及有道词典的cookie, 并选择加载对应的本地词典
- 功能选项和生成文件的选项之间存在依赖关系
- 生成的产物都保存在pdf同级目录下
- 本地翻译需要加载词典, 可以在"添加配置"中选择同级目录下Dictionary中两个英汉大词典

## 效果图

#### 界面
<img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/54fef90b-a4d7-4061-b44a-d8e1d99f6150">
   <img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/1d68322d-7b2d-4d19-ad69-d41494c3c08a">  <img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/9b4d442f-4159-49ca-b071-c9d4c67a1634">  <img width="340" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/c0d8ff81-0624-4315-b298-5458fd447a71">


#### 电脑端有道单词本
<img width="600" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/59ea605c-55d1-4293-aa89-9dbc73b843c8">


#### 手机端有道单词本
<img width="260" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/b9e42dcb-0c98-4286-8bfc-f863c13962f2"><img width="260" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/041c767a-d391-4a14-9d8d-59d67a4750ae">


## 打赏
<img width="196" alt="image" src="https://github.com/hgzerowzh/pdfScanner/assets/64787489/76a3ab13-23f7-45f3-b978-ec4ac580e140">
