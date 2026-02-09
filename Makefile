
.PHONY: test smoke report-biometano

test:
	pytest -q

smoke:
	./scripts/smoke_test.sh

report-biometano:
	dcf biometano report --input case_files/biometano_case.yaml
