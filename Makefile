.PHONY: install run test lint lint-py lint-js format clean help

# Python 解释器：优先用系统自带的（有 pytest），可用 make test PYTHON=xxx 覆盖
PYTHON ?= /usr/bin/python3

help: ## 显示所有可用命令
	@echo "测试用例显示工具 - 可用命令："
	@echo "  make install   安装依赖（运行时 + 开发）"
	@echo "  make run       启动服务"
	@echo "  make test      运行测试（带覆盖率）"
	@echo "  make lint      代码检查（Python flake8 + JS 语法）"
	@echo "  make lint-py   仅检查 Python（flake8，需先 pip install flake8）"
	@echo "  make lint-js   仅检查 JS 语法（node --check，零依赖）"
	@echo "  make format    自动格式化（black + isort，需先安装）"
	@echo "  make clean     清理缓存文件"

install: ## 安装所有依赖
	pip3 install -r requirements.txt -r requirements-dev.txt

run: ## 启动服务
	$(PYTHON) testcase_viewer.py

test: ## 运行测试
	$(PYTHON) -m pytest tests/ -v --tb=short --cov=testcase_viewer --cov-report=term

lint: lint-py lint-js ## 代码检查（Python + JS）

lint-py: ## Python 代码检查（flake8）
	@if $(PYTHON) -m flake8 --version >/dev/null 2>&1; then \
	    echo "=== flake8 检查 ==="; \
	    $(PYTHON) -m flake8 app/ tests/ testcase_viewer.py; \
	    echo "flake8 检查通过 ✅"; \
	else \
	    echo "⚠️  flake8 未安装，跳过 Python 检查"; \
	    echo "   安装方法: pip3 install flake8"; \
	    echo "   当前用 $(PYTHON) 语法检查替代:"; \
	    $(PYTHON) -m py_compile testcase_viewer.py && echo "testcase_viewer.py 语法检查通过 ✅" || echo "❌ 语法错误"; \
	    find app/ -name '*.py' -exec $(PYTHON) -m py_compile {} \; && echo "app/ 语法检查通过 ✅" || echo "❌ 语法错误"; \
	fi

lint-js: ## JS 语法检查（node --check，零依赖）
	@if command -v node >/dev/null 2>&1; then \
	    echo "=== JS 语法检查 ==="; \
	    node --check app/static/js/app.js && echo "app.js 语法检查通过 ✅"; \
	else \
	    echo "⚠️  node 未安装，跳过 JS 检查"; \
	fi

format: ## 自动格式化（需先安装 black + isort）
	@if $(PYTHON) -m black --version >/dev/null 2>&1; then \
	    $(PYTHON) -m black app/ tests/ testcase_viewer.py; \
	else \
	    echo "⚠️  black 未安装，跳过格式化"; \
	    echo "   安装方法: pip3 install black isort"; \
	fi
	@if $(PYTHON) -m isort --version >/dev/null 2>&1; then \
	    $(PYTHON) -m isort app/ tests/ testcase_viewer.py; \
	fi

clean: ## 清理缓存
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	rm -rf .pytest_cache .coverage htmlcov
