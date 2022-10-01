# 说明
## 介绍
这是一个使用JavaScript开发的WPS加载项，其主要功能是协助论文写作  
该加载项部署在服务器上，使用时需要安装相应的xml文件，见 **使用教程**  
由于加载项文件在服务端，用户无需进行手动更新，但需要在使用时保持联网
  
加载项文件部署在两个服务端上
| **部署服务器** | **安全性** | **稳定性** | **更新** |
| :------------: | :--------: | :--------: | :------: |
|     Github     |     高     |     差     |   及时   |
|      私人      |     低     |     高     |   延迟   |

可以选择一种方法进行部署
| **部署模式** | **部署难易度** |
| :----------: | :------------: |
|   publish    |      复杂      |
|  jsplugins   |      简单      |
>指南 [WPS开发者文档](https://qn.cache.wpscdn.cn/encs/doc/office_v19/webhelpframe.htm)
## 功能
### 对参考文献索引进行APA排版，疑似错误部分将会标红
它不能转换索引格式，只是按APA格式进行排版，所以需要在使用前确保索引符合APA内容格式  
> **错误标红仅图一乐，以实际APA格式为准**  

错误标红常见原因：  
- 标点符号错误，比如少了逗号、少了句号、多了空格
- 缺少页码，比如`*, 17(10).`
- 作者名出现罕见字符或同时出现中英文  
- 标题同时出现问号和句号或多个句号，比如 `你摸鱼了吗? 大学生网课情况调研. `
- 文献来源部分无法识别，比如书籍索引、会议文章索引
  - 程序仅支持识别**一般索引**和**学位论文索引**格式
- 程序分段错误，这将导致正确格式也被标红，请查看**调试信息**
  - 小概率事件
- 等等...
> 经过2021年部分心理学报文章的参考文献测试，在内容格式正确前提下，报错率**低于1%**  
> 但如果你的参考文献中包含大量书籍索引，请谨慎使用  
> 查看**调试信息**：选择一段索引运行程序，然后点击js调试器 > Console > 输入 `_debug.txt` 回车

# 使用教程（jsplugins模式）
## 告诉WPS加载项文件在哪
`右键桌面WPS打开文件所在位置\11.1.0.***（版本号）\offlce6\cfgs\`  
找到 `oem.ini` 右键编辑，在 `JSPluginsServer=` 的右边，填入以下任意一个链接：  
- Github部署 https://cubxx.github.io/wps-addon/%E8%AE%BA%E6%96%87/jsplugins.xml
- 私人部署 http://47.113.221.157:81/wps-addon/jsplugins.xml
> 建议先看看链接是否有效，如果你看到 `This XML file does not …` 说明有效
## 确保jsplugins模式打开
在 `oem.ini` 中应该有写 `JsApiPlugin=true` 和 `JsApiShowWebDebugger=true`  
启动WPS，WPS将根据第一步填入的链接自动获取 `jsplugins.xml` 文件
## 检验是否安装成功
`Win`+`E`打开文件资源管理器  
在地址栏输入`%appdata%/kingsoft/wps/jsaddons`  
检查该目录下是否存在 `jsplugins.xml`
## 打开WPS，载入xml
随便打开一个文档，`开发工具 > 加载项 > 添加`  
转到上述 `jsplugins.xml` 文件路径，并选择该文件  
主选项卡面板将出现 `论文` 加载项
>显示所有文件后，才能看到 .xml 文件
## 先选中文字，再点击功能按键
点击 `参考文献` 按钮后，按钮会在程序运行期间**保持灰色**，一般速度大概是100字/s

# 使用教程（publish模式）
## 请使用Edge或Chrome浏览器
## 点击[Github部署](https://cubxx.github.io/wps-addon/%E8%AE%BA%E6%96%87/publish.html)或[私人部署](http://47.113.221.157:81/wps-addon/publish.html)
安装成功后会在本地生成 `publish.xml` 文件，WPS通过该文件实现加载项功能
>安装、卸载、更新时请关闭WPS  
因为跨域问题，使用Github部署可能状态无效，但这不影响安装
## 检验是否安装成功
`Win`+`E`打开文件资源管理器  
在地址栏输入`%appdata%/kingsoft/wps/jsaddons`  
检查该目录下是否存在 `publish.xml`
## 打开WPS，载入xml
随便打开一个文档，`开发工具 > 加载项 > 添加`  
转到上述 `publish.xml` 文件路径，并选择该文件  
主选项卡面板将出现 `论文` 加载项
>显示所有文件后，才能看到 .xml 文件
## 先选中文字，再点击功能按键
点击 `参考文献` 按钮后，按钮会在程序运行期间**保持灰色**，一般速度大概是100字/s

# 问题解决
## 部署网页中没有安装选项（publish模式）
将 <chrome://flags/#block-insecure-private-network-requests> 粘贴至浏览器地址栏中  
如果浏览器未显示网页的话，请自行查询该浏览器如何进行flags设置，或更换为Chrome内核浏览器（如：Edge, Goolge, 360, QQ, Sogou, 星愿）  
接着将 `block-insecure-private-network-requests` 设置为 `Disabled`
>**详细**: **[WPS官方解决文档](https://www.kdocs.cn/l/cv7pyp6sqOFC)**  
>**注意**: 此操作将关闭浏览器限制非本地服务器访问[本地服务器](http://localhost)的功能，可能存在一定风险
## 载入xml文件后没有出现加载项
### 自定义功能区  
在WPS中 `文件 > 选项 > 自定义功能区` 勾选 `论文`  
### 加载模式问题*（仅适用于 `publish.xml`）  
`右键桌面WPS打开文件所在位置\11.1.0.***（版本号）\offlce6\cfgs\`  
找到 `oem.ini` 右键编辑，使 `JsApiPlugin=false`，保存退出，重启WPS  
### 网络问题
可以通过点击以下链接判断能否连接至服务器：  
- Github部署 https://cubxx.github.io/wps-addon/%E8%AE%BA%E6%96%87/wps-addon-build/ribbon.xml
- 私人部署 http://47.113.221.157:81/wps-addon/wps-addon-build/ribbon.xml
### 加载延迟
网络没问题的话，可能是你打开WPS的时候，它还没加载好  
在资源管理器中进入 `%appdata%/kingsoft/wps/jsaddons` 找到相应模式的 `.xml` 文件并打开  
如果文件内含有网址链接，则说明WPS将通过该网址获取加载项  
如果没有 `.xml` 文件或文件中没有网址链接：  
- jsplugins模式
  - 检查 `oem.ini` 文件，重复第一步
  - 保持WPS打开，WPS将通过 `oem.ini` 获取 `jsaddons.xml` 文件
  - 检查 `jsaddons.xml` 文件，WPS将通过该文件获取加载项
- publish模式
  - 进入部署网页，查看安装情况，或者重装
  - 检查 `publish.xml` 文件，WPS将通过该文件获取加载项

# 功能原理
## 参考文献格式化
### 正则匹配
- 作者：  
  - 中文 `[?啊啊, 啊啊, & 啊啊. `  
  - 英文 `A'a, A., B-b, B., & C c, C. `
- 日期 `(1234b). `
- 标题  
  - `*. `
  - `*? `
- 来源：  
  - 一般文献  
    - `*, 1(2-3), 4-5, 6. `
    - `*, (2-3), 4-5, 6. `
    - `*, 1, 4-5, 6. ` 
    - `*, 4-5, 6. `  
  - 学位论文 `*, *. `  
