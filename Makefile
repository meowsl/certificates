# Including commands
install:
	pyenv install 3.10.0
	python -m venv .venv
	poetry install --no-root
	poetry env use 3.10.0

run:
	poetry run python .\main.py