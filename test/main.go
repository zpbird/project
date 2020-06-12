package main

import "fmt"

// main ...
func main() {
	var aaa = []struct {
		name string
		age  int
	}{}
	lll := len(aaa)
	fmt.Println(lll)
	aaa = append(aaa, struct {
		name string
		age  int
	}{name: "aaa",
		age: len(aaa)})

	fmt.Println(len(aaa))

}
