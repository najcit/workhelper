## RPA(Robotic Process Automation)
1. 全称机器人流程自动化
2. 通过特定的、可模拟人类在计算机界面上进行操作的技术
3. 按指定的规则自动执行相应的计算机操作的流程任务

## 界面元素操作技术
1. 发送模拟指令消息(如发送快捷键消息)
2. 绝对坐标定位界面元素 + 模拟鼠标、键盘
3. 图像识别定位界面元素 + 模拟鼠标、键盘
4. 侵入式技术获取界面元素信息(UIA等) + 模拟鼠标、键盘

## GUI 自动化实现原理
1. automation.py + uiautomation 获取控件信息及操作控件
2. pyautogui 图像模糊匹配获取位置信息，模拟键盘操作和触发快捷键
3. 如果程序不支持uiautomation，则只能使用图片模糊匹配

## WEB 自动化实现原理
1. Selenium + WebDriver 
2. Playwright 

### 参考
1. [RPA界面元素智能自适应定位与操控技术-金克](https://baijiahao.baidu.com/s?id=1750962224735488757&wfr=spider&for=pc)
2. [自动化测试框架结构图](https://mp.weixin.qq.com/s?__biz=MzkxMzI4ODgyOA==&mid=2247511176&idx=1&sn=03281f833b55e467c0f61e8a39e1427a&chksm=c17d179bf60a9e8d12a3febccdc0a122fdaff2756ada2be4d0ea3616c2f4ab34bb65da52e7e2&scene=21#wechat_redirect)