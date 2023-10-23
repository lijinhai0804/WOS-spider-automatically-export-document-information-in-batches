# WOS-spider-automatically-export-document-information-in-batches
## 简介（Brief introduction）
  The tool simulates user operations to automatically and repeatedly export document information on Web of Science (WOS) for Bibliometrics. Users can freely set the number of documents to be exported, choose different export file formats, and rename the exported files. The original creator of this tool is CSDN blogger: Parzival_  link:(https://blog.csdn.net/Parzival_/article/details/122360528) and I haven’t found his Github account yet. My main work was to optimize the original code and use updated versions of tools and more commonly used Google Chrome browser to implement the operation.Please include this statement and all source links when reproducing.
<br>
<br>   本工具通过模拟用户操作，自动且重复地从Web of Science (WOS)导出文献信息来用于文献计量。用户可以自由设置要导出的文档数量，选择不同的导出文件格式并重命名导出的文件。该工具的原始创建者是CSDN博主：Parzival_
<br>链接：(https://blog.csdn.net/Parzival_/article/details/122360528) 我还没有找到他的Github账户。我的主要工作是优化原始代码并使用更新版本的工具和更常用的Google Chrome浏览器来实现操作。请在转载时包含此声明和所有来源链接。
## 下载及安装操作（Download and Install）
1. main.py是主文件，必须下载。merge.py用来将导出的不同excel表内容融合到一张表内，按需下载。driver install.py不一定会用得到，如果电脑里是最新的谷歌浏览器，且安装了最新的selenium库，不一定会需要driver install.py安装驱动，但是如果主程序运行报错，可以尝试运行driver install.py。
2. 我所使用的软件及库版本：python 3.11;selenium 4.14;chrome 118.0.5993.89
3. 代码中用到的库：selenium（必要）\webdriver_manager\pandas\xlrd\os\glob。Pycharm可以扫描未安装的库然后来安装。
## 如何使用（How to use）
1. 在使用前main.py中代码末尾的主要函数的参数需要进行修改。url是复制已经检索好的WOS网址；record_num设置为下载篇数；download_path是下载文件储存地址（**一定要是空文件夹**）；record_format是下载的文件格式，目前可以填excel、bib；reverse按时间降序排列，默认关闭。
2. 建议不要在翻墙的时候使用。除非是在国外的大学就读。
3. 第一次下载时，预留了10秒的时间手动更改所需的下载字段（即作者、摘要、参考文献这些），之后的自动下载都只会默认点击已经自定义好的选项，所以**要在第一次下载时设定好**。
4. **强烈建议连上校园网直接IP免登录进去WOS**，如果做不到，则需要启用login函数（main.py中已经默认注释掉了）。
5. 每次下载完后，记得把文件转移到别的地方，清空文件夹，再开启新一次的下载。
6. merge.py融合的文件默认储存在代码所处的文件夹。
7. 如果显示文件下载失败，可能是WOS的问题，换一个时间试试。
