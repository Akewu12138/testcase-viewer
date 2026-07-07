"""InMemoryTestCaseRepository — 内存实现的 Repository（阶段2 T10）。

用于单元测试，不依赖真实 Excel 文件。
换数据库时，测试用此夹具验证 API 逻辑，不碰真实存储。
"""

from typing import Dict, List, Optional, Any
from app.services.repository import TestCaseRepository


class InMemoryTestCaseRepository(TestCaseRepository):
    """内存 Repository，所有数据存在实例属性中，不读写文件。"""

    def __init__(self, testcases: List[Dict[str, Any]] = None):
        self._testcases: List[Dict[str, Any]] = testcases or []
        self._mapping: Dict[str, int] = {}
        self._headers: List[str] = []
        self._sheet_name: str = 'test_sheet'
        self._header_row: int = 1
        self._save_calls: List[Dict] = []  # 记录 save 调用（测试可断言）

    def load(self, filepath: str) -> tuple:
        """返回内存中预置的数据（忽略 filepath）。"""
        return (
            self._testcases,
            self._mapping,
            self._headers,
            self._sheet_name,
            self._header_row,
        )

    def get(self, index: int) -> Optional[Dict[str, Any]]:
        if index < 0 or index >= len(self._testcases):
            return None
        return self._testcases[index]

    def list_all(self) -> List[Dict[str, Any]]:
        return self._testcases

    def save(self, filepath: str, row_number: int, header_row_idx: int,
             sheet_name: str, result: str, actual_result: str,
             tester: str, bug_id: str, bug_frequency: str,
             issue_time: str) -> bool:
        """只更新内存，不写文件。记录调用参数供测试断言。"""
        self._save_calls.append({
            'filepath': filepath,
            'row_number': row_number,
            'header_row_idx': header_row_idx,
            'sheet_name': sheet_name,
            'result': result,
            'actual_result': actual_result,
            'tester': tester,
            'bug_id': bug_id,
            'bug_frequency': bug_frequency,
            'issue_time': issue_time,
        })
        # 更新内存中对应用例的执行状态
        for tc in self._testcases:
            if tc.get('_row') == row_number:
                tc['_saved_result'] = result
                tc['_saved_actual_result'] = actual_result
                tc['_saved_tester'] = tester
                tc['_saved_bug_id'] = bug_id
                tc['_saved_bug_frequency'] = bug_frequency
                tc['_saved_issue_time'] = issue_time
                break
        return True

    def get_summary(self, testcases: List[Dict[str, Any]]) -> Dict[str, Any]:
        """与 ExcelTestCaseRepository 相同的汇总逻辑。"""
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

    # ---- 测试辅助方法 ----

    def set_testcases(self, testcases: List[Dict[str, Any]]):
        """设置内存用例数据（测试时预置数据用）。"""
        self._testcases = testcases

    @property
    def save_call_count(self) -> int:
        """返回 save 被调用的次数（测试断言用）。"""
        return len(self._save_calls)
