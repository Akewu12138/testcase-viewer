"""业务逻辑层

阶段2 迁入：从 testcase_viewer.py 拆分出纯业务逻辑，包括：
- excel_reader.py   Excel 读取（read_testcases 等）
- excel_writer.py   Excel 写回（save_result 等）
- column_detector.py 列检测（detect_columns 等）

这些函数不依赖 Flask，最容易写单元测试。
"""
