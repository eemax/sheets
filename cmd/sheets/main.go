// Command sheets reads, writes and formats spreadsheets for AI agents.
package main

import (
	"os"

	"github.com/eemax/sheets/internal/cli"
)

func main() {
	os.Exit(cli.Main(os.Args[1:], os.Stdin, os.Stdout, os.Stderr))
}
