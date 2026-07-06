.PHONY: install run test lint format clean help

help: ## 显示所有可用命令
	@echo "测试用例显示工具 - 可用命令："
	@echo "  make install  安装依赖（运行时 + 开发）"
	@echo "  make run      启动服务"
	@echo "  make test     运行测试（带覆盖率）"
	@echo "  make lint     代码检查（flake8）"
	@echo "  make format   自动格式化（black + isort）"
	@echo "  make clean    清理缓存文件"

install: ## 安装所有依赖
	pip install -r requirements.txt -r requirements-dev.txt

run: ## 启动服务（阶段1仍用旧入口）
	python testcase_viewer.py

test: ## 运行测试
	python -m pytest tests/ -v --tb=short --cov=testcase_viewer --cov-report=term

lint: ## 代码检查
	python -m flake8 app/ tests/ testcase_viewer.py

format: ## 自动格式化
	python -m black app/ tests/ testcase_viewer.py
	python -m isort app/ tests/ testcase_viewer.py

clean: ## 清理缓存
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	rm -rf .pytest_cache .coverage htmlcov
