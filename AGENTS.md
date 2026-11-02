# SheetPorter 项目约定

- 默认用中文沟通。保持纯本地处理，不添加联网或上传数据的功能。
- 引擎位于 `fillcore.py`，图形界面位于 `sheetporter.py`。
- 修改填表行为后运行 `python -m unittest discover -s tests -v`。
- 公开内容仅包含合成示例；不提交用户规则、真实表格、输出、凭据或本机路径。
- 示例机构统一使用“示例医院”。
- Windows 程序从公开源码重新构建；不发布来源不明或与源码不同步的旧 exe。
