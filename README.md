# spider_wencai
爬取问财网的股票子公司A股数据
xiugai.py 存放 程序代码，直接运行即可
ase.min.js 是根据cookie随机生成请求页面时的v值
.xslx 是放需要爬取的股票名字，命名格式    股票名 子公司     
这个.xslx的文件在程序中 选择指定的文件

运行程序会打印运行情况，在公司下显示有异常，表示此公司在页面可能找不到（如退市，或页面格式不一致），需要手动搜索看一下
运行结束会生成一个 test.xslx的文件，里面存储爬取的数据，也可在程序中修改生成文件的名称
