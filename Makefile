.PHONY: all clean build

all: clean build

clean:
	(cd ./excel/cmd/example && go clean && rm -f output.xlsx)
	(cd ./tui/cmd/list-simple && go clean)
	(cd ./tui/cmd/title && go clean)

build: 
	(cd ./excel/cmd/example && go build)
	(cd ./tui/cmd/list-simple && go build)
	(cd ./tui/cmd/title && go build)
