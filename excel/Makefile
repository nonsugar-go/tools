.PHONY: all clean build test

all:

clean:
	$(MAKE) -C ./cmd/example clean
	$(MAKE) -C ./cmd/getinfo clean
	rm -f cover.out cover.html ./testdata/output.xlsx

build:
	go generate
	$(MAKE) -C ./cmd/example build
	$(MAKE) -C ./cmd/getinfo build

test:
	go test -v -coverprofile=cover.out
	go tool cover -func=cover.out
	go tool cover -html=./cover.out -o cover.html
