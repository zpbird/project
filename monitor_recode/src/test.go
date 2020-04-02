package test

import "fmt"

func foo() (int, string) {
	return 10, "Q1mi"
}
func test() {
	x, _ := foo()
	_, y := foo()
	fmt.Println("x=", x)
	fmt.Println("y=", y)
}
