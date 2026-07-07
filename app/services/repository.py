"""Repository 抽象层（阶段2 T2 核心）。

定义数据访问接口 TestCaseRepository，API 层通过此接口访问数据，
不直接碰 STATE 全局字典和 Excel 文件。

当前 ExcelTestCaseRepository 委托给现有 read_testcases/save_result 函数。
未来换数据库时，只需新增一个实现（如 SqliteTestCaseRepository），API 零改动。

依赖注入：app factory 创建 repository 实例，注入给 API 蓝图。
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any


class TestCaseRepository(ABC):
    """测试用例数据访问接口。

    所有数据读写都通过此接口，API 层不直接碰 STATE 或 Excel 文件。
    切换底层存储（Excel → SQLite → PostgreSQL）时，只需新增实现类。
    """

    @abstractmethod
    def load(self, filepath: str) -> tuple:
        """加载用例文件，返回 (testcases, mapping, headers, sheet_name, header_row)。

        Args:
            filepath: Excel 文件路径

        Returns:
            五元组：用例列表、列映射、表头、Sheet名、表头行号
        """
        ...

    @abstractmethod
    def get(self, index: int) -> Optional[Dict[str, Any]]:
        """按索引获取单条用例（含已保存的执行结果）。

        Args:
            index: 用例索引（0-based）

        Returns:
            用例 dict，索引越界返回 None
        """
        ...

    @abstractmethod
    def list_all(self) -> List[Dict[str, Any]]:
        """获取全部用例（含已保存的执行结果）。

        Returns:
            用例 dict 列表
        """
        ...

    @abstractmethod
    def save(self, filepath: str, row_number: int, header_row_idx: int,
             sheet_name: str, result: str, actual_result: str,
             tester: str, bug_id: str, bug_frequency: str,
             issue_time: str) -> bool:
        """保存测试执行结果到 Excel 指定行。

        Args:
            filepath: Excel 文件路径
            row_number: Excel 行号（1-based）
            header_row_idx: 表头行号（1-based）
            sheet_name: Sheet 名称
            result: 测试结果（通过/失败/阻塞/跳过）
            actual_result: 实际结果
            tester: 测试人员
            bug_id: BugID
            bug_frequency: Bug频率
            issue_time: 问题时间

        Returns:
            True 表示成功
        """
        ...

    @abstractmethod
    def get_summary(self, testcases: List[Dict[str, Any]]) -> Dict[str, Any]:
        """根据用例列表生成汇总统计。

        Args:
            testcases: 用例列表

        Returns:
            汇总 dict：total/pass/fail/block/skip/none/pass_rate 等
        """
        ...


class ExcelTestCaseRepository(TestCaseRepository):
    """Excel 文件实现的 Repository。

    当前委托给 testcase_viewer 模块的现有函数（过渡方案）。
    T3-T5 搬代码后，改为从 app/services/ 导入。
    """

    def load(self, filepath: str) -> tuple:
        # 延迟导入避免循环依赖，T3 完成后改为从 app.services.excel_reader 导入
        from testcase_viewer import read_testcases
        return read_testcases(filepath)

    def get(self, index: int) -> Optional[Dict[str, Any]]:
        from testcase_viewer import STATE
        if index < 0 or index >= len(STATE['testcases']):
            return None
        return STATE['testcases'][index]

    def list_all(self) -> List[Dict[str, Any]]:
        from testcase_viewer import STATE
        return STATE['testcases']

    def save(self, filepath: str, row_number: int, header_row_idx: int,
             sheet_name: str, result: str, actual_result: str,
             tester: str, bug_id: str, bug_frequency: str,
             issue_time: str) -> bool:
        from testcase_viewer import save_result
        save_result(filepath, row_number, header_row_idx, sheet_name,
                    result, actual_result, tester, bug_id, bug_frequency, issue_time)
        return True

    def get_summary(self, testcases: List[Dict[str, Any]]) -> Dict[str, Any]:
        """生成汇总统计（从 api_summary 逻辑抽出）。"""
        total = len(testcases)
        counts = {'通过': 0, '失败': 0, '阻塞': 0, '跳过': 0, '未执行': 0}
        for tc in testcases:
            r = tc.get('_saved_result', '').strip()
            if r in counts:
                counts[r] += 1
            else:
                counts['未执行'] += 1

        executed = total - counts['未执行']
        pass_rate = round(counts['通过'] / total * 100, 1) if total > 0 else 0

        return {
            'total': total,
            'executed': executed,
            'pass': counts['通过'],
            'fail': counts['失败'],
            'block': counts['阻塞'],
            'skip': counts['跳过'],
            'none': counts['未执行'],
            'pass_rate': pass_rate,
        }


# ============================================================
# 工厂函数：按配置创建 repository 实例
# ============================================================

def create_repository(backend: str = 'excel') -> TestCaseRepository:
    """按配置创建 repository 实例。

    Args:
        backend: 存储后端类型，当前支持 'excel'，未来可扩展 'sqlite'/'postgres'

    Returns:
        TestCaseRepository 实例
    """
    if backend == 'excel':
        return ExcelTestCaseRepository()
    raise ValueError(f"不支持的存储后端: {backend}（当前仅支持 'excel'）")
