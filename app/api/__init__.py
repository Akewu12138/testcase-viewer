"""API 路由层

阶段2 迁入：用 Flask Blueprint 组织路由，按功能分文件：
- testcase.py        用例读取相关
- search.py          搜索筛选
- summary.py         汇总统计
- file_management.py 上传/重载

路由层保持「薄」，只做：接收请求 -> 调用 service -> 返回结果。
"""
