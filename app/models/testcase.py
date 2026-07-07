"""领域模型数据类（阶段2 T1）。

定义测试用例相关的三个核心数据结构：
- TestCase: 一条测试用例的领域数据
- ExecutionRecord: 一次执行的结果记录
- SheetMeta: Excel Sheet 的解析元数据

当前阶段代码仍用 dict 传递数据，这些 dataclass 作为类型文档和未来重构目标。
T9 加类型注解后，核心函数可逐步改为返回 dataclass。
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class ExecutionRecord:
    """一次测试执行的结果记录。

    一个 TestCase 可以有多次执行记录（当前实现只存最后一次，这是设计局限）。
    """
    result: str = ''              # 测试结果：通过/失败/阻塞/跳过
    actual_result: str = ''       # 实际结果（原"测试现象备注"）
    tester: str = ''              # 测试人员
    bug_id: str = ''              # BugID
    bug_frequency: str = ''       # Bug频率
    issue_time: str = ''          # 问题时间
    executed_at: str = ''         # 执行时间（自动填充）


@dataclass
class TestCase:
    """一条测试用例的领域数据。

    对应 Excel 的一行，包含用例本身的内容和已保存的执行状态。
    raw_columns 保留原始列值（col_N），用于前端动态展示。
    """
    # 领域数据（通过列映射从 Excel 解析）
    id: str = ''                  # 用例编号
    title: str = ''               # 用例标题
    module: str = ''              # 所属模块
    priority: str = ''            # 优先级
    precondition: str = ''        # 前置条件
    steps: str = ''               # 测试步骤
    expected: str = ''            # 预期结果
    purpose: str = ''             # 测试目的

    # 解析元数据
    row: int = 0                  # Excel 行号（1-based）
    raw_columns: Dict[str, str] = field(default_factory=dict)   # 原始列值 col_N -> value
    raw_headers: Dict[str, str] = field(default_factory=dict)   # 原始列名 _header_N -> name
    search_text: str = ''         # 全文搜索文本（所有字段拼接）

    # 执行状态（从 Excel 读取的已保存结果）
    saved_result: str = ''
    saved_actual_result: str = ''
    saved_tester: str = ''
    saved_bug_id: str = ''
    saved_bug_frequency: str = ''
    saved_issue_time: str = ''

    def to_dict(self) -> Dict:
        """转为 dict（兼容现有前端和 API 代码）。"""
        return {
            'id': self.id,
            'title': self.title,
            'module': self.module,
            'priority': self.priority,
            'precondition': self.precondition,
            'steps': self.steps,
            'expected': self.expected,
            'purpose': self.purpose,
            '_row': self.row,
            '_search_text': self.search_text,
            '_saved_result': self.saved_result,
            '_saved_actual_result': self.saved_actual_result,
            '_saved_tester': self.saved_tester,
            '_saved_bug_id': self.saved_bug_id,
            '_saved_bug_frequency': self.saved_bug_frequency,
            '_saved_issue_time': self.saved_issue_time,
            **{k: v for k, v in self.raw_columns.items()},
            **{k: v for k, v in self.raw_headers.items()},
        }


@dataclass
class SheetMeta:
    """Excel Sheet 的解析元数据。

    记录用例数据来自哪个 Sheet、表头在哪行、列映射关系。
    """
    sheet_name: str = ''          # Sheet 名称
    header_row: int = 1           # 表头行号（1-based）
    headers: List[str] = field(default_factory=list)        # 表头列表
    mapping: Dict[str, int] = field(default_factory=dict)   # 列映射：field_name -> col_index


@dataclass
class LoadedData:
    """read_testcases 返回的完整数据包。"""
    testcases: List[Dict] = field(default_factory=list)     # 当前用 dict，未来改 List[TestCase]
    mapping: Dict[str, int] = field(default_factory=dict)
    headers: List[str] = field(default_factory=list)
    sheet_name: str = ''
    header_row: int = 1
