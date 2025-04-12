.PHONY: all clean build test

all:

clean:
	$(MAKE) -C ./cmd/table clean
	$(MAKE) -C ./cmd/title clean
	$(MAKE) -C ./cmd/file_picker clean
	rm -f cover.out cover.html

build:
	$(MAKE) -C ./cmd/table build
	$(MAKE) -C ./cmd/title build
	$(MAKE) -C ./cmd/file_picker build

test:
	go test -v -coverprofile=cover.out
	go tool cover -func=cover.out
	go tool cover -html=./cover.out -o cover.html
