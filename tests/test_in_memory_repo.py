"""InMemoryTestCaseRepository 单元测试（阶段2 T10）。

验证内存 Repository 的接口契约，确保换数据库时测试逻辑不变。
"""

import sys
import os
import webbrowser
webbrowser.open = lambda url: None

project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from app.services.in_memory_repo import InMemoryTestCaseRepository


class TestInMemoryRepo:
    """测试 InMemoryTestCaseRepository 接口契约"""

    def _make_testcases(self):
        """预置 3 条测试用例"""
        return [
            {'id': 'TC001', 'title': '登录测试', '_row': 2, '_saved_result': '通过'},
            {'id': 'TC002', 'title': '登出测试', '_row': 3, '_saved_result': '失败'},
            {'id': 'TC003', 'title': '搜索测试', '_row': 4, '_saved_result': ''},
        ]

    def test_load_returns_preset_data(self):
        """load 返回预置数据（不读文件）"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        testcases, mapping, headers, sheet_name, header_row = repo.load('fake.xlsx')
        assert len(testcases) == 3
        assert sheet_name == 'test_sheet'
        assert header_row == 1

    def test_get_valid_index(self):
        """get 有效索引 → 返回用例"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        tc = repo.get(0)
        assert tc is not None
        assert tc['id'] == 'TC001'

    def test_get_invalid_index(self):
        """get 无效索引 → 返回 None"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        assert repo.get(-1) is None
        assert repo.get(999) is None

    def test_list_all(self):
        """list_all → 返回全部用例"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        all_tcs = repo.list_all()
        assert len(all_tcs) == 3

    def test_save_updates_memory(self):
        """save → 更新内存中对应用例的执行状态"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        repo.save(
            filepath='fake.xlsx', row_number=4, header_row_idx=1,
            sheet_name='test_sheet', result='阻塞', actual_result='网络不通',
            tester='张三', bug_id='BUG001', bug_frequency='偶现', issue_time='2024-01-01'
        )
        tc = repo.get(2)
        assert tc['_saved_result'] == '阻塞'
        assert tc['_saved_actual_result'] == '网络不通'
        assert tc['_saved_tester'] == '张三'

    def test_save_records_call(self):
        """save → 记录调用参数（测试可断言）"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        repo.save('fake.xlsx', 2, 1, 'test_sheet', '通过', 'OK', '李四', '', '', '')
        assert repo.save_call_count == 1

    def test_summary(self):
        """get_summary → 正确统计各状态数量"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        summary = repo.get_summary(repo.list_all())
        assert summary['total'] == 3
        assert summary['pass'] == 1
        assert summary['fail'] == 1
        assert summary['none'] == 1
        assert summary['executed'] == 2

    def test_summary_empty(self):
        """get_summary 空列表 → 不崩溃"""
        repo = InMemoryTestCaseRepository([])
        summary = repo.get_summary([])
        assert summary['total'] == 0
        assert summary['pass_rate'] == 0

    def test_does_not_touch_filesystem(self):
        """save → 不写真实文件（文件不存在也不崩）"""
        repo = InMemoryTestCaseRepository(self._make_testcases())
        # 传入不存在的路径，不应抛异常
        repo.save('/nonexistent/path.xlsx', 2, 1, 'test', '通过', '', '', '', '', '')
        assert repo.save_call_count == 1
