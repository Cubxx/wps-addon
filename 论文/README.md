# 说明
## 介绍
这是一个使用JavaScript开发的WPS加载项，其主要功能是协助论文写作  
该加载项部署在服务器上，使用时需要安装相应的xml文件，见 **使用教程**  
由于加载项文件在服务端，用户无需进行手动更新，但需要在使用时保持联网  
**文件部署在两台服务器上**  
Github服务端会经常更新，使用的是较安全的https协议，但是Github有时候会被墙，不太稳定  
私人服务器因为上传文件比较麻烦，偶尔才会更新，而且是使用不太安全的http协议，但更稳定
>参考指南 [WPS开发者文档](https://qn.cache.wpscdn.cn/encs/doc/office_v19/webhelpframe.htm)
## 功能
- 对参考文献索引进行APA排版，错误部分将会标红
  - 它不能转换索引格式，只是按APA格式进行排版  
  - 需要在使用前确保索引符合APA内容格式

# 使用教程
## 请使用Edge或Chrome浏览器
## 点击[Github部署（推荐）](https://cubxx.github.io/wps-addon/%E8%AE%BA%E6%96%87/publish.html)或[私人部署](http://47.113.221.157:81/wps-addon/publish.html)
安装成功后会在本地生成 `publish.xml` 文件，WPS通过该文件实现加载项功能
>安装、卸载、更新时保持WPS关闭状态  
因为跨域问题，使用Github部署可能状态无效，但这不影响安装
## 检验是否安装成功（可跳过）
`Win`+`E`打开文件资源管理器  
在地址栏输入`%appdata%/kingsoft/wps/jsaddons`  
检查该目录下是否存在 `publish.xml`
## 打开WPS，载入xml
随便打开一个文档，`开发工具 > 加载项 > 添加`  
转到上述 `publish.xml` 文件路径，并选择该文件  
主选项卡面板将出现 `论文` 加载项
>显示所有文件后，才能看到 .xml 文件
## 先选中文字，再点击功能按键

# 问题解决
## 部署网页中没有安装选项（使用私人部署的问题）
将 <chrome://flags/#block-insecure-private-network-requests> 粘贴至浏览器地址栏中  
如果浏览器未显示网页的话，请自行查询该浏览器如何进行flags设置，或更换为Chrome内核浏览器（如：Edge, Goolge, 360, QQ, Sogou, 星愿）  
接着将 `block-insecure-private-network-requests` 设置为 `Disabled`
>**详细**: **[WPS官方解决文档](https://www.kdocs.cn/l/cv7pyp6sqOFC)**  
>**注意**: 此操作将关闭浏览器限制非本地服务器访问[本地服务器](http://localhost)的功能，可能存在一定风险
## 载入xml后没有出现加载项
**自定义功能区**  
在WPS中 `文件 > 选项 > 自定义功能区` 勾选 `论文`  
**加载模式问题**  
`右键桌面WPS打开文件所在位置\11.1.0.***（版本号）\offlce6\cfgs\`  
找到 `oem.ini` 右键编辑，使 `JsApiPlugin=false`，保存退出，重启WPS  
## 其他
请速DD我

# 功能原理
## 参考文献格式化
### 正则匹配
- 作者：  
  - 中文 `[?啊啊, 啊啊, & 啊啊. `  
  - 英文 `A'a, A., B-b, B., C c, C. '`
- 日期 `(1234b). `
- 标题 `*. ` 或 `*? `
- 来源：  
  - 一般文献 `*!(,.), 1(2-3), a4-5, b6. `  
  - 学位论文 `*!(,.1), *!(,.1). ` 或 `*!(,.1) *!(,.1). `  
  - 结尾 ` ?]? `
