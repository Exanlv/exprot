build-xlsx:
	rm -rf ./tmp/test;
	php index.php;
	cd ./tmp/test; zip -r ../../result.xlsx *;
	open result.xlsx;
