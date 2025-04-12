.PHONY: all clean build test

all: clean test build

clean:
	$(MAKE) -C ./excel clean
	$(MAKE) -C ./tui clean

build: 
	$(MAKE) -C ./excel build
	$(MAKE) -C ./tui build

test:
	$(MAKE) -C ./excel test 
	$(MAKE) -C ./tui test
