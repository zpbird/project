package main

import "fmt"

// main ...
func main() {
	var aaa struct {
		name1    *string
		name2    *string
		fullname string
	}
	aaa.name1 = new(string)
	aaa.name2 = new(string)

	*aaa.name1 = "zp"
	*aaa.name2 = "bird"
	aaa.fullname = *aaa.name1 + *aaa.name2
	for i := 1; i < 10; i++ {
		*aaa.name2 += "y"
		fmt.Println(*aaa.name2)
		fmt.Println(aaa.fullname)
	}

}
