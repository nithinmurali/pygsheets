
TEST_PATH=tests

clean-pyc:
	find . -name '*.pyc' -exec rm -f {} +
	find . -name '*.pyo' -exec rm -f {} +
	find . -name '*~' -exec rm -f  {} +
	find . -name '__pycache__' -exec rm -rf {} +

clean-build:
	rm --force --recursive build/
	rm --force --recursive dist/
	rm --force --recursive *.egg-info
	rm .coverage

clean: clean-pyc clean-build

doc:
	cd docs;make pre;make html;cd ..

lint:
	flake8 --filename = ./pygsheets/*.py

test: clean-pyc
	py.test -vs --cov=pygsheets --cov-config .coveragerc ../pygsheets $(TEST_PATH)

install:
	python setup.py install

publish: clean
	python setup.py publish

.PHONY: clean-pyc clean-build
