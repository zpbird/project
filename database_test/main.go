package main

import (
	"context"
	"fmt"
	"log"
	"os"

	"github.com/gin-gonic/gin"
	"github.com/jackc/pgx/v4/pgxpool"
)

func db(abbr string) {
	var dburl = "host=192.168.100.11 port=5432 user=zp password=790204 database=ttjt sslmode=disable"
	dbpool, err := pgxpool.Connect(context.Background(), dburl)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Unable to connect to database: %v\n", err)
		os.Exit(1)
	}
	defer dbpool.Close()

	var n string
	var i int
	var t bool
	err = dbpool.QueryRow(context.Background(), "select id,name,exist from base.company where abbr=$1", abbr).Scan(&i, &n, &t)
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println(i, n, t)
}

// main ...
func main() {
	g := gin.Default()
	g.GET("/test", func(c *gin.Context) {
		// fmt.Println("请求路径", c.FullPath())
		abb := c.Query("abbr")
		db(abb)
	})
	if err := g.Run(":8090"); err != nil {
		log.Fatal(err.Error())
	}
}
