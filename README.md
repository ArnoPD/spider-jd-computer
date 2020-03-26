# 目录
* [App&Package](#App&Package)
* [运行](#运行)
* [遇到的问题](#遇到的问题)
* [结果示例](#结果示例)

# spider-jd-computer
爬取京东电脑类目并提取定义的电脑硬件参数帮助选购

## App&Package
- Python 3.7（其他未经测试）
- FireFox
- geckodriver（FireFox的web driver）
- xlwings 0.18.0
- selenium 3.141.0
- BeautifulSoup 4.7.1

## 运行
指定爬取链接，运行list.py
```
link = '爬取链接' + 'page={}'
```

ps.win系统未经测试

## 遇到的问题
[在PyCharm使用xlwings提示权限问题，但没有示意授权](https://github.com/xlwings/xlwings/issues/1146)


### 结果示例
id | 产品名 | 链接 | 价格 | 优惠券 | 折扣价 | CPU | 硬盘 | 内存 | 显卡 | 显存 | 重量 | 色域 | 刷新率
---|---|---|---|---|---|---|---|---|---|---|---|---|---
100005638619 | 荣耀笔记本电脑MagicBook 14 | 第三方Linux版 14英寸全面屏轻薄本（AMD锐龙R5 3500U 8G 512G 冰河银 | https://item.jd.com/100005638619.html | 3099 | 无 | 3099 | 锐龙 5 | 512GB SSD | 8GB | AMD R5 共享系统内存（集成） | 2.01kg | 其他 | 无