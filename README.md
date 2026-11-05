# AutoOfficeMCP
---
## 基于 .NET Framework COM 以及  VSTO (Visual Studio Tools for Office) 两个功能(如果归类起来就是 Framework COM)

要使用本项目 可以去 [Releases](https://github.com/whyxiaoxing/AutoOfficeMCP/releases/latest) 下载二进制文件(不保证最新)或者下载本项目的源码进行编译<br>

目前共有 200+ 工具<br>
下面是突出特点<br>
| 对象 | 重点开放 |
|---|---|
| word | 对文字，排版 表格操作进重点开放 对于其余的页眉部分没有突出 主要是办公日常 不考虑特殊情况 |
| xlsx | 对数据收集，信息收集 跨xlsx 筛选做出重点开放 对函数部分全部开放 其余不做处理 | 
| ppt | 开放比较少，主要是动画与页数排版 |

附加内容有
[SKILL.md](/workspace/SKILL.md)
[Config.md](./Config.json.md)

# 使用步骤

1. 根据 `Config.json.md`修改 `config.json`
2. 运行 `register-urlacl.ps1`
3. 运行 `AutoOfficeMCP.exe` 
4. 修改 `SKILL.md`中的 `MCP_BASE_URL`
5. 你是在你的AI终端 添加MCP 支持 sse和 http
6. 让AI读取 `SKILL.md`
7. 发送要求
