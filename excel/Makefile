.PHONY: all clean build test

all:

clean:
	$(MAKE) -C ./cmd/example clean
	rm -f cover.out cover.html ./testdata/output.xlsx

build:
	$(MAKE) -C ./cmd/example build

test:
	go test -v -coverprofile=cover.out
	go tool cover -func=cover.out
	go tool cover -html=./cover.out -o cover.html
